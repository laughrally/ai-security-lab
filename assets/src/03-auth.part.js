/* ============================================================
   ▼ STEP 1: Supabaseの設定値をここに貼り付けてください
   Supabaseダッシュボード → Settings → API Keys → Publishable key
   ※ publishable キーはブラウザに公開される前提のキーです。
     テーブルへのアクセス制御は RLS ポリシーで行っています。
   ============================================================ */
const SUPABASE_URL  = 'https://zektuwdsmqpdbbslxyal.supabase.co';
const SUPABASE_ANON = 'sb_publishable_MKMYWP8AP0xMcoKlEazytw_lFTTebRn';
const WORKER_URL    = 'https://ai-security-lab-stripe.laughrally.workers.dev';

/* ============================================================
   ▼ 無料で見られるタブ名（これ以外はログイン必須）
   ============================================================ */

/* ========== Supabase初期化 ========== */
let _supabase = null;
let currentUser = null;
let currentUserPlan = null;

function initSupabase(){
  if(SUPABASE_URL === 'YOUR_SUPABASE_URL') return;
  try {
    _supabase = createClient(SUPABASE_URL, SUPABASE_ANON);
    console.log('Supabase initialized:', !!_supabase);

    // 決済完了後のリダイレクト処理
    const urlParams = new URLSearchParams(window.location.search);
    if(urlParams.get('payment') === 'success'){
      // 決済完了メッセージを表示してログイン画面を開く
      history.replaceState(null, '', window.location.pathname);
      setTimeout(()=>{
        openAuthModal();
        setAuthMsg('success','🎉 決済が完了しました！登録したメールアドレスとパスワードでログインしてください。');
      }, 500);
    }
    if(urlParams.get('payment') === 'cancel'){
      history.replaceState(null, '', window.location.pathname);
    }

    // URLハッシュにアクセストークンがある場合（OAuth後のリダイレクト）
    if(window.location.hash && window.location.hash.includes('access_token')){
      const params = new URLSearchParams(window.location.hash.substring(1));
      const access_token = params.get('access_token');
      const refresh_token = params.get('refresh_token');
      const type = params.get('type');
      if(access_token){
        // supabase-js は detectSessionInUrl によりURL内のトークンを自動で消費する。
        // 手動で setSession を呼ぶと二重処理になりセッションが壊れるため、
        // 既存セッションの有無を確認してから必要な場合のみ張り直す。
        (async ()=>{
          let session = null;
          try {
            const { data } = await _supabase.auth.getSession();
            session = data?.session || null;
          } catch(e){ /* noop */ }

          if(!session){
            const { data, error } = await _supabase.auth.setSession({access_token, refresh_token});
            if(error){
              console.error('setSession failed:', error.message);
            } else {
              session = data?.session || null;
            }
          }

          if(session){
            if(type === 'recovery'){
              // パスワードリセットリンクからのアクセス → PW設定画面を表示
              openAuthModal();
              toggleAuthMode('newpw');
            } else {
              currentUser = session.user;
              updateAuthUI();
              closeAuthModal();
            }
          }
          history.replaceState(null, '', window.location.pathname);
        })();
      }
    }

    _supabase.auth.onAuthStateChange((_event, session)=>{
      if(_event === 'PASSWORD_RECOVERY'){
        openAuthModal();
        toggleAuthMode('newpw');
        return;
      }
      // ログイン・ログアウトで人が切り替わったら読み直す（2026-11-23）。以前は前の人の章末ワークの記述、
      // テストの回答済み表示、法人の受講者一覧がメモリに残り、同じタブで次にログインした人に見えた
      const _uid = session?.user?.id || null;
      if(window._aslUid !== undefined && window._aslUid !== _uid && (_event === 'SIGNED_IN' || _event === 'SIGNED_OUT')){
        location.reload();
        return;
      }
      window._aslUid = _uid;
      claimLocalData(_uid);
      currentUser = session?.user ?? null;
      if(!currentUser){ currentUserPlan=null; }
      updateAuthUI();
      if(session){ closeAuthModal(); fetchUserPlan(); }
    });
    _supabase.auth.getSession().then(({data:{session}})=>{
      if(window._aslUid === undefined) window._aslUid = session?.user?.id || null;
      claimLocalData(session?.user?.id || null);
      currentUser = session?.user ?? null;
      updateAuthUI();
      if(currentUser) fetchUserPlan();
    });
  } catch(e) {
    console.error('Supabase init error:', e);
  }
}

// この端末に残る受講データ（受講済み・学習時間）は、最後にログインした人のもの。
// 別の人がログインしたら先に消す（2026-11-23）。ログアウトせずにセッションが切れたあと、
// 次の人の受講済みや学習時間に前の人の分が混ざらないように
const LOCAL_LEARN_KEYS = ['asl_completed','asl_active_seconds','asl_lesson_started','asl_synced_seconds','asl_last_tab','asl_lesson_served'];
// 本文を受け取った時刻の署名（Worker が付ける）。受講済みにするときに送り返し、受講時間の上限に使われる。
// 最初に受け取ったものを残す（開き直しで時刻が新しくなると、それまでの受講時間が短く数えられるため）
const SERVED_KEY = 'asl_lesson_served';
function rememberServed(lessonId, served){
  if(!served || !served.at || !served.sig) return;
  const all = store.getJSON(SERVED_KEY, {}) || {};
  // 最初に受け取ったものを残す。ただし Worker は30日より古い署名を使わないので、29日を過ぎたら新しいものに替える
  const old = all[lessonId];
  if(old && old.at <= served.at && Date.now() - old.at < 29 * 864e5) return;
  all[lessonId] = served;
  store.set(SERVED_KEY, all);
}
function servedFor(lessonId){ return (store.getJSON(SERVED_KEY, {}) || {})[lessonId] || null; }
function claimLocalData(uid){
  if(!uid) return;
  try {
    const owner = store.get('asl_owner');
    if(owner && owner !== uid) LOCAL_LEARN_KEYS.forEach(k => store.remove(k));
    if(owner !== uid) store.set('asl_owner', uid);
  } catch(e){}
}

let currentExamPassed = false;
// レッスン追加（2026-11）より前に同じプランで修了証を受け取った人。Worker の /api/records が返す
let currentCertifiedEarlier = false;
async function fetchUserPlan(){
  if(!_supabase || !currentUser) return;
  if(window._planFetching) return window._planFetching;   // 起動時に2経路から呼ばれるので1回にまとめる
  window._planFetching = _fetchUserPlan().finally(() => { window._planFetching = null; });
  return window._planFetching;
}
async function _fetchUserPlan(){
  // .single() は失敗しても例外を投げず { error } を返す。以前は通信失敗でも 'basic' 扱いになり、
  // 有料の受講者が自分の講座を開けなかった。2回まで取り直し、それでも駄目なら知らせる（2026-11-22）
  let data = null, error = null;
  for(let i = 0; i < 3; i++){
    try {
      ({ data, error } = await _supabase.from('profiles').select('plan,exam_passed').eq('id', currentUser.id).single());
    } catch(e) { error = e; }
    if(data && !error) break;
    await new Promise(r => setTimeout(r, 800 * (i + 1)));
  }
  if(data && !error){
    currentUserPlan = data.plan ?? 'free';
    currentExamPassed = data.exam_passed ?? false;
  } else {
    // 読めなかったときは Worker と同じく「プランなし」に倒す（以前は画面だけ Basic 扱いで、本文の取得で弾かれた）
    currentUserPlan = currentUserPlan || 'free';
    currentExamPassed = false;
    console.warn('プランを取得できませんでした:', error && error.message);
    alert('ご契約プランの情報を読み込めませんでした。通信環境をご確認のうえ、ページを再読み込みしてください。');
  }
  if(!window._corpLoaded){ window._corpLoaded = true; loadCorpProgress(); }
  if(currentUserPlan === 'standard' || currentUserPlan === 'premium'){
    refreshRecords();
  }
  updateCertificateUI();
  applyPlanLock();
  try{ applyAdminVisibility(); }catch(e){}
  // Supabaseから受講済みレッスンを復元
  await syncCompletedFromSupabase();
  try{ applyAdminVisibility(); }catch(e){}
  // 管理者ダッシュボードの表示は applyAdminVisibility() に一本化
}

// Supabaseから受講済みレッスンを取得してlocalStorageとUIに反映
async function syncCompletedFromSupabase(){
  if(!_supabase || !currentUser) return;
  try {
    const { data, error } = await _supabase
      .from('lesson_logs')
      .select('lesson_id')
      .eq('user_id', currentUser.id);
    if(error){ console.warn('sync error:', error.message); return; }

    // 受講済みはサーバーの記録で置き換える（2026-11-23）。以前は端末の記録と合わせており、送信に失敗した
    // レッスンが端末にだけ「受講済み」として残り、修了の条件は満たした表示なのに申請がサーバーで弾かれた
    const remoteCompleted = (data || []).map(r => r.lesson_id);
    saveCompleted(remoteCompleted);

    // ※ lesson_logs への書き込みはWorker経由のみ。
    //   ローカルの申告をそのまま正とすると受講記録を捏造できるため、
    //   ここでは同期しない（サーバー側の記録が正）。

    // UIを更新
    updateCompletionUI();
    updateDashboard();
  } catch(e){
    console.warn('syncCompletedFromSupabase failed:', e);
  }
}

function updateCertificateUI(){
  const formBox   = document.querySelector('.cert-form-box');
  const upgradeBox = document.getElementById('cert-upgrade-notice');
  if(!formBox) return;
  if(currentUserPlan !== 'standard' && currentUserPlan !== 'premium'){
    formBox.style.display = 'none';
    if(!upgradeBox){
      const box = document.createElement('div');
      box.id = 'cert-upgrade-notice';
      box.style.cssText = 'background:var(--bg2);border:1.5px solid var(--amber,#f59e0b);border-radius:12px;padding:24px 28px;text-align:center;margin-top:16px;';
      box.innerHTML = '<div style="font-size:22px;margin-bottom:10px">🎓</div>'
        + '<div style="font-weight:600;font-size:15px;margin-bottom:8px">修了証はStandard・Premiumプランの特典です</div>'
        + '<div style="font-size:13px;color:var(--text3);margin-bottom:18px;line-height:1.7">ご契約のプランは修了証の発行対象外です。<br>Standard（' + yen(COURSE.prices.standard) + '）以上で修了証PDFを取得できます。</div>'
        + '<button type="button" class="form-submit" onclick="showPlans()" style="width:auto;padding:10px 28px">料金プランを見る →</button>';
      formBox.parentNode.insertBefore(box, formBox);
    }
  } else {
    formBox.style.display = '';
    if(upgradeBox) upgradeBox.remove();
  }
}

function updateAuthUI(){
  const bar        = document.getElementById('auth-user-bar');
  const email      = document.getElementById('auth-user-email');
  const navCta     = document.getElementById('nav-cta');
  const navPricing = document.getElementById('nav-pricing');
  const navProg    = document.getElementById('nav-progress');
  const progBar    = document.getElementById('nav-progress-bar');
  const navLogin    = document.getElementById('nav-login');
  const navProblem  = document.getElementById('nav-problem');
  const pricingSection = document.getElementById('pricing');
  if(currentUser){
    bar.classList.add('show');
    email.textContent = currentUser.email ?? 'ログイン済み';
    const sbEmail = document.getElementById('sb-user-email');
    if(sbEmail) sbEmail.textContent = currentUser.email ?? '';
    lockPaidTabs(false);
    if(navCta)        navCta.style.display     = 'none';
    if(navPricing)    navPricing.style.display = 'none';
    if(navLogin)      navLogin.style.display   = 'none';
    if(navProblem)    navProblem.style.display  = 'none';
    if(navProg)       navProg.style.display    = 'inline-flex';
    // FABはCSSのlearn-wrapper.show連動で制御（JS側では触らない）
    if(progBar)       progBar.classList.add('show');
    if(pricingSection) pricingSection.style.display = 'none';
    const hero = document.getElementById('top');
    const problem = document.getElementById('problem');
    const guide = document.getElementById('how-to-use');
    if(hero)    hero.style.display    = 'none';
    if(problem) problem.style.display = 'none';
    if(guide)   guide.classList.add('show');
    document.body.classList.add('logged-in');
    updateCompletionUI();
    updateDashboard();
  } else {
    bar.classList.remove('show');
    lockPaidTabs(true);
    if(navCta)        navCta.style.display     = '';
    if(navPricing)    navPricing.style.display = '';
    if(navLogin)      navLogin.style.display   = '';
    if(navProblem)    navProblem.style.display  = '';
    if(navProg)       navProg.style.display    = 'none';
    // FABはCSSのlearn-wrapper.show連動で制御
    if(progBar)       progBar.classList.remove('show');
    if(pricingSection) pricingSection.style.display = '';
    const hero2 = document.getElementById('top');
    const problem2 = document.getElementById('problem');
    const guide2 = document.getElementById('how-to-use');
    if(hero2)    hero2.style.display    = '';
    if(problem2) problem2.style.display = '';
    if(guide2)   guide2.classList.remove('show');
    document.body.classList.remove('logged-in');
    const dash = document.getElementById('progress-dashboard'); if(dash) dash.style.display='none';
    hideLearnUI();
  }
}

/* ログイン済みかどうかで有料タブのクラスを切り替え */
function lockPaidTabs(lock){
  document.querySelectorAll('.tab').forEach(t=>{
    const m = t.getAttribute('onclick')?.match(/switchTab\('([^']+)'\)/);
    if(!m) return;
    const name = m[1];
    if(!FREE_TABS.includes(name)){
      t.classList.toggle('locked', lock);
    }
  });
}

// プランに応じてL23〜L34のアクセスをロック
function applyPlanLock(){
  document.querySelectorAll('.sb-row, .lrow').forEach(el => {
    const onclick = el.getAttribute('onclick') || '';
    // 章末テストの行（goChapterTest(7)）は、その章のワーク（work7）と同じ判定にする
    const m = onclick.match(/(?:switchTab|goLesson)\('([^']+)'\)/) || onclick.match(/(?:goChapterTest|openWorkTab)\((\d+)\)/);
    // キーボード・読み上げで使えるように（2026-11-23。認定試験・修了証・レポートの行に無かった）
    if(el.classList.contains('sb-row') && onclick){
      if(!el.hasAttribute('tabindex')) el.setAttribute('tabindex', '0');
      if(!el.hasAttribute('role')) el.setAttribute('role', 'button');
      if(el.classList.contains('active')) el.setAttribute('aria-current', 'page'); else el.removeAttribute('aria-current');
      const t = el.querySelector('.sb-row-title');
      if(t && el.classList.contains('done') && !t.querySelector('.sr-only')) t.insertAdjacentHTML('beforeend', '<span class="sr-only">（受講済み）</span>');
    }
    if(!m) return;
    const id = /^\d+$/.test(m[1]) ? 'work' + m[1] : m[1];
    const r = currentUser ? tabDenyReason(id) : null;
    const locked = !!(r && r !== 'login');
    el.classList.toggle('plan-locked', locked);
    el.style.opacity = '';
    el.style.pointerEvents = '';
    el.title = locked ? r.msg : '';
    if(locked) el.setAttribute('aria-disabled', 'true'); else el.removeAttribute('aria-disabled');
  });
}

/* ===== 認証モーダル制御 ===== */
function openAuthModal(){
  document.getElementById('auth-overlay').classList.remove('hidden');
  setAuthMsg('', '');
}
function closeAuthModal(){
  document.getElementById('auth-overlay').classList.add('hidden');
}
function setAuthMsg(type, text){
  const el = document.getElementById('auth-msg');
  el.className = 'auth-msg' + (type ? ' '+type : '');
  el.textContent = text;
}
function toggleAuthMode(mode){
  document.getElementById('auth-form-login' ).style.display = mode==='login'  ? 'block' : 'none';
  document.getElementById('auth-form-reset' ).style.display = mode==='reset'  ? 'block' : 'none';
  document.getElementById('auth-form-newpw' ).style.display = mode==='newpw'  ? 'block' : 'none';
  const titles = {login:'受講者ログイン', reset:'パスワードの再設定', newpw:'新しいパスワードを設定'};
  document.getElementById('auth-modal-title').textContent = titles[mode] || '受講者ログイン';
  setAuthMsg('', '');
}

/* ===== Google ログイン ===== */
/* ===== メール+PW ログイン ===== */
async function signInWithEmail(){
  if(!_supabase){ setAuthMsg('error','Supabaseが未設定です。'); return; }
  const email = document.getElementById('auth-email').value.trim();
  const pw    = document.getElementById('auth-pw').value;
  if(!email||!pw){ setAuthMsg('error','メールアドレスとパスワードを入力してください。'); return; }
  setAuthMsg('', '読み込み中...');
  const {error} = await _supabase.auth.signInWithPassword({email, password:pw});
  if(error){
    setAuthMsg('error', 'ログインに失敗しました。メールアドレスまたはパスワードをご確認ください。');
  } else {
    closeAuthModal();
  }
}

/* ===== パスワードリセット ===== */
async function sendPasswordReset(){
  if(!_supabase){ setAuthMsg('error','Supabaseが未設定です。'); return; }
  const email = document.getElementById('auth-reset-email').value.trim();
  if(!email){ setAuthMsg('error','メールアドレスを入力してください。'); return; }
  setAuthMsg('', '送信中...');
  const {error} = await _supabase.auth.resetPasswordForEmail(email, {
    redirectTo: 'https://ai-security-lab.com/reset-password.html'
  });
  if(error){
    setAuthMsg('error', '送信に失敗しました。メールアドレスをご確認ください。');
  } else {
    setAuthMsg('success', '✅ リセットメールを送信しました。メール内のリンクをクリックしてください。');
  }
}

/* ===== パスワード更新（リセットリンクから） ===== */
async function updatePassword(){
  if(!_supabase){ setAuthMsg('error','Supabaseが未設定です。'); return; }
  const pw  = document.getElementById('auth-newpw').value;
  const pw2 = document.getElementById('auth-newpw2').value;
  if(!pw || !pw2){ setAuthMsg('error','パスワードを入力してください。'); return; }
  if(pw.length < 10){ setAuthMsg('error','パスワードは10文字以上で設定してください（推測されにくくするため）。'); return; }
  if(pw !== pw2){ setAuthMsg('error','パスワードが一致しません。'); return; }
  setAuthMsg('', '更新中...');
  const {error} = await _supabase.auth.updateUser({password: pw});
  if(error){
    console.error('updateUser failed:', error);
    const detail = error.message || '';
    const hint = /session|JWT|token/i.test(detail)
      ? 'リンクの有効期限が切れているか、既に使用済みの可能性があります。お手数ですがパスワード再設定メールを再送してください。'
      : detail;
    setAuthMsg('error', '更新に失敗しました。' + (hint ? '（' + hint + '）' : 'もう一度お試しください。'));
  } else {
    setAuthMsg('success', '✅ パスワードを更新しました！');
    setTimeout(()=>{ toggleAuthMode('login'); }, 2000);
  }
}


async function signOut(){
  if(!_supabase) return;
  // ログアウト時に、この端末に残る受講データを消す（共用PCで次の人の記録に混ざらないように）。
  // 以前は受講済みだけを消しており、学習時間の累計や開始時刻が次の人の送信に乗っていた（2026-11-22）
  LOCAL_LEARN_KEYS.forEach(k => store.remove(k));
  currentCertifiedEarlier = false;
  currentExamPassed = false;
  await _supabase.auth.signOut();
  updateCompletionUI();
  updateDashboard();
}

