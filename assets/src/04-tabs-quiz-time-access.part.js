/* ===== タブ切り替え（認証チェック付き） ===== */
// 未ログインのときのタブ表示の切り替え（トップページのタブ）。開けるかどうかは window.switchTab が判定済み。
// 以前は同じ switchTab という名前で、上書き版と2つあった（2026-11-23 に名前を分けた）
function showTabGuest(name){
  document.querySelectorAll('.tab').forEach(t=>{
    const m = t.getAttribute('onclick')?.match(/switchTab\('([^']+)'\)/);
    if(m) t.classList.toggle('active', m[1]===name);
  });
  document.querySelectorAll('.tcon').forEach(t=>t.classList.remove('active'));
  const el = document.getElementById('tab-'+name);
  if(el) el.classList.add('active');
  markLessonStart(name);
  shuffleQuizOptions(name);
  if(name === 'certificate') updateCertificateUI();
  document.getElementById('lessons').scrollIntoView({behavior:'smooth'});
}

/* ===== その他の既存関数 ===== */
function toggleCh(el){ el.closest('.chapter').classList.toggle('open') }

/* ===== 学習記録API（採点はサーバー側で行う） ===== */

/**
 * 理解度チェックの選択肢の並びをランダム化する。
 * 正解が常に特定の位置・最も長い選択肢になっていると、
 * 内容を理解していなくてもパターンで正答できてしまうため。
 *
 * ※ 各ボタンの onclick には「元の選択肢番号」が埋め込まれており、
 *    見た目の順序を変えてもサーバーへ送る値は変わらない。
 */
function shuffleQuizOptions(tabId){
  const tcon = document.getElementById('tab-' + tabId);
  if(!tcon) return;
  tcon.querySelectorAll('.qopts').forEach(box => {
    // 回答済み・シャッフル済みの設問には触れない
    if(box.dataset.shuffled === '1') return;
    const btns = Array.from(box.querySelectorAll('.qopt'));
    if(btns.length < 2) return;
    if(btns.some(b => b.classList.contains('ok') || b.classList.contains('ng'))) return;
    for(let i = btns.length - 1; i > 0; i--){
      const j = Math.floor(Math.random() * (i + 1));
      [btns[i], btns[j]] = [btns[j], btns[i]];
    }
    btns.forEach(b => box.appendChild(b));
    box.dataset.shuffled = '1';
  });
}
window.shuffleQuizOptions = shuffleQuizOptions;

// ---- 受講時間の実測（2026-08-26 全面改修）----
// 旧実装は「初めて開いた時刻」から「受講済みを押した時刻」までの単純な経過時間で、
// 複数レッスンを開くと同じ時間帯が何重にも計上され、離席分も加算され、
// リロードで消えていた。実績が実稼働の3.2倍に膨らむ例が出ていた。
// 新実装は「表示されていて、かつ操作がある間」だけを積算する。
const ACTIVE_KEY = 'asl_active_seconds';   // { lessonN|workN: 累計秒 }
const START_KEY  = 'asl_lesson_started';   // { lessonN|workN: 最初に開いたISO時刻 }
const SYNC_KEY   = 'asl_synced_seconds';   // { lessonN: 最後にサーバへ送った秒数 }
const IDLE_LESSON_MS = 3 * 60 * 1000;      // レッスンは読むので3分
const IDLE_WORK_MS   = 5 * 60 * 1000;      // ワークは書くので5分（考える時間を削らない）
const RESYNC_STEP_S  = 60;                 // 完了済みレッスンは60秒たまるごとに送り直す
const MAX_TICK_S = 30;                     // 1回の加算上限。スリープ復帰の跳ねを抑える
const MAX_TOTAL_S = 3 * 3600;              // 1レッスンの上限（Worker の LESSON_MAX_SECONDS と同じ）
const TRACK_RE = /^(lesson\d+|work\d+)$/;

let activeItem  = null;   // いま表示している lessonN / workN（それ以外は null）
let tickAt      = null;   // 前回の加算時刻
let lastInputAt = Date.now();
let wasVisible  = (document.visibilityState === 'visible');

function loadActive(){ try { return JSON.parse(localStorage.getItem(ACTIVE_KEY) || '{}'); } catch(e){ return {}; } }
function saveActive(o){ try { localStorage.setItem(ACTIVE_KEY, JSON.stringify(o)); } catch(e){} }
function loadStarts(){ try { return JSON.parse(localStorage.getItem(START_KEY) || '{}'); } catch(e){ return {}; } }
function loadSynced(){ try { return JSON.parse(localStorage.getItem(SYNC_KEY) || '{}'); } catch(e){ return {}; } }
function saveSynced(o){ try { localStorage.setItem(SYNC_KEY, JSON.stringify(o)); } catch(e){} }

// 完了済みレッスンを学習し直したぶんをサーバへ反映する。
// 受講済みボタンは pointer-events:none で二度と押せないため、
// これが無いと復習の時間は localStorage に溜まるだけでサーバに届かない。
// lesson_logs は upsert なので、送り直せば duration_seconds が上書きされる。
function syncStudiedAgain(itemId, secs){
  if(!/^lesson\d+$/.test(itemId)) return;
  if(typeof currentUser === 'undefined' || !currentUser) return;
  let completed = [];
  try { completed = loadCompleted(); } catch(e){ return; }
  if(!completed.includes(itemId)) return;          // 未完了なら markComplete が送る
  const sy = loadSynced();
  if(secs - (sy[itemId] || 0) < RESYNC_STEP_S) return;
  sy[itemId] = secs; saveSynced(sy);
  apiFetch('/api/lesson/complete', {
    method: 'POST',
    body: JSON.stringify({
      lessonId: itemId,
      served: servedFor(itemId),
      startedAt: loadStarts()[itemId] || null,
      durationSeconds: secs,
    }),
  }).catch(e => console.warn('復習時間の反映に失敗:', e.message));
}

// 前回の加算時刻から現在までを確定する。
// 可視判定は「その区間で表示されていたか」を見るため、遷移後の状態ではなく
// 直前の状態（wasVisible）を使う。そうしないと非表示になる瞬間の分が落ちる。
function flushActive(){
  if(activeItem && tickAt !== null){
    const now  = Date.now();
    const limit = activeItem.startsWith('work') ? IDLE_WORK_MS : IDLE_LESSON_MS;
    const idle = (now - lastInputAt) > limit;
    if(wasVisible && !idle){
      const d = Math.min(Math.floor((now - tickAt) / 1000), MAX_TICK_S);
      if(d > 0){
        const o = loadActive();
        const v = Math.min((o[activeItem] || 0) + d, MAX_TOTAL_S);
        o[activeItem] = v;
        saveActive(o);
        syncStudiedAgain(activeItem, v);
      }
    }
    tickAt = now;
  }
  wasVisible = (document.visibilityState === 'visible');
}

// タブを開いたときに呼ばれる。レッスン／ワーク以外へ移ったら計測を止める。
function markLessonStart(tabId){
  flushActive();                       // 直前の項目ぶんを確定してから切り替える
  if(TRACK_RE.test(tabId)){
    activeItem = tabId;
    tickAt = Date.now();
    const st = loadStarts();
    if(!st[tabId]){                    // 「最初に開いた時刻」は初回だけ記録して保持する
      st[tabId] = new Date().toISOString();
      try { localStorage.setItem(START_KEY, JSON.stringify(st)); } catch(e){}
    }
  } else {
    activeItem = null;                 // ダッシュボード等では加算しない
    tickAt = null;
  }
}
window.markLessonStart = markLessonStart;

['mousemove','keydown','scroll','click','touchstart'].forEach(ev =>
  window.addEventListener(ev, () => { lastInputAt = Date.now(); }, { passive: true })
);
document.addEventListener('visibilitychange', flushActive);
window.addEventListener('pagehide', flushActive);
setInterval(flushActive, 15000);       // 15秒ごとに確定。リロードで失う分を最小化する

async function apiFetch(path, options){
  if(!_supabase) throw new Error('Supabaseが未初期化です');
  const { data } = await _supabase.auth.getSession();
  const token = data?.session?.access_token;
  if(!token) throw new Error('ログインが必要です');
  const res = await fetch(WORKER_URL + path, {
    ...options,
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + token,
      ...(options?.headers || {}),
    },
  });
  const body = await res.json().catch(()=>({}));
  if(!res.ok){
    const err = new Error(body.error || ('通信に失敗しました（' + res.status + '）'));
    err.status = res.status;      // 呼び出し側で 429 などを判別できるようにする
    err.body = body;
    throw err;
  }
  return body;
}
window.apiFetch = apiFetch;

/**
 * 理解度テストの回答。
 * 正解はクライアントに存在せず、Workerが判定して返す。
 * 初回回答が評価対象として記録される。
 */
async function doQ(btn, selected, id){
  if(!currentUser){
    const tq = btn.closest('.tcon');
    const tid = tq ? tq.id.replace('tab-','') : '';
    if(FREE_TABS.includes(tid)) return doQGuest(btn, selected, id);
    openAuthModal(); return;
  }

  const qopts = btn.closest('.qopts');
  const quiz  = btn.closest('.quiz');
  const tcon  = quiz ? quiz.closest('.tcon') : null;
  const tabId = tcon ? tcon.id.replace('tab-','') : null;
  if(!tabId) return;

  // 二重送信を防ぐ
  if(qopts.dataset.sending === '1') return;
  qopts.dataset.sending = '1';
  const prevText = btn.textContent;
  btn.style.opacity = '.6';

  try {
    const r = await apiFetch('/api/quiz/answer', {
      method: 'POST',
      body: JSON.stringify({ lessonId: tabId, questionId: id, selected }),
    });

    btn.style.opacity = '';

    // 解説はレッスン本文に無く、回答した後にサーバーから返ってくる。
    // 先読みされないよう、受け取ってから流し込む。
    const fbEl = document.getElementById(id + (r.isCorrect ? '-ok' : '-ng'));
    if(fbEl && r.explanation) fbEl.innerHTML = (r.isCorrect ? '✅ ' : '❌ ') + escHtml(r.explanation);

    if(r.isCorrect){
      qopts.querySelectorAll('.qopt').forEach(b => { b.style.pointerEvents = 'none'; });
      btn.classList.add('ok');
      document.getElementById(id+'-ok').classList.add('show');
      document.getElementById(id+'-ng').classList.remove('show');
      unlockCompleteBtn(tabId);
      // 正解した設問を控えておく。レッスンを開き直したときに再回答させないため。
      // 正解の位置はサーバーにしかないので、doQ が送る「元の選択肢番号」を保存する
      // （表示順はシャッフルされるが、この番号は QUIZ_ANSWER_KEY と対応している）。
      try {
        const done = JSON.parse(localStorage.getItem(quizDoneKey()) || '{}');
        done[id] = { s: selected, e: r.explanation || '' };
        localStorage.setItem(quizDoneKey(), JSON.stringify(done));
      } catch(e){}
    } else {
      // 不正解：この選択肢だけ無効化。他は引き続き選べる。
      btn.style.pointerEvents = 'none';
      btn.classList.add('ng');
      document.getElementById(id+'-ng').classList.add('show');
      if(r.attemptsLeft === 0) quizAttemptsUsedUp(qopts, id, tabId);
    }
  } catch(e) {
    btn.style.opacity = '';
    btn.textContent = prevText;
    if(e.status === 429){ quizAttemptsUsedUp(qopts, id, tabId); return; }
    alert('回答を記録できませんでした。\n' + (e.message || '') + '\n通信環境をご確認のうえ、もう一度お試しください。');
  } finally {
    qopts.dataset.sending = '';
  }
}

// 理解度チェックは1問3回まで（Worker の QUIZ_MAX_ATTEMPTS）。使い切ったら選択肢を閉じ、受講済みにできるようにする
function quizAttemptsUsedUp(qopts, id, tabId){
  qopts.querySelectorAll('.qopt').forEach(b => { b.style.pointerEvents = 'none'; b.setAttribute('aria-disabled', 'true'); });
  const ng = document.getElementById(id + '-ng');
  if(ng){
    ng.classList.add('show');
    const note = document.createElement('div');
    note.className = 'q-used-up';
    note.textContent = '回答できる回数（3回）を使い切りました。レッスンを読み直してから、受講済みにしてください（正誤は受講実績レポートに記録されます）。';
    if(!ng.querySelector('.q-used-up')) ng.appendChild(note);
  }
  unlockCompleteBtn(tabId);
}

// 未ログインで無料公開レッスンの設問に答えたとき（2026-11-22）。答え合わせと解説だけを出し、記録はしない
async function doQGuest(btn, selected, id){
  const qopts = btn.closest('.qopts');
  if(!qopts || qopts.dataset.sending === '1') return;
  qopts.dataset.sending = '1';
  try {
    const tq = btn.closest('.tcon');
    const tabId = tq ? tq.id.replace('tab-','') : '';
    const res = await fetch(WORKER_URL + '/api/quiz/answer', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ lessonId: tabId, questionId: id, selected }),
    });
    const r = await res.json().catch(() => ({}));
    if(!res.ok) throw new Error(r.error || '通信に失敗しました');
    const fb = document.getElementById(id + (r.isCorrect ? '-ok' : '-ng'));
    const note = '<br><span style="font-size:12px;opacity:.8">※ 未ログインのため記録はされません。ご購入後は回答と受講記録が残ります。</span>';
    if(fb) fb.innerHTML = (r.isCorrect ? '✅ ' : '❌ ') + escHtml(r.explanation || '') + note;
    if(r.isCorrect){
      qopts.querySelectorAll('.qopt').forEach(b => { b.style.pointerEvents = 'none'; });
      btn.classList.add('ok');
      document.getElementById(id + '-ok').classList.add('show');
      document.getElementById(id + '-ng').classList.remove('show');
    } else {
      btn.style.pointerEvents = 'none';
      btn.classList.add('ng');
      document.getElementById(id + '-ng').classList.add('show');
    }
  } catch(e) {
    alert('答え合わせができませんでした。\n' + (e.message || ''));
  } finally {
    qopts.dataset.sending = '';
  }
}
window.doQGuest = doQGuest;

// 章末ワークのページを開き、確認テストの位置まで移動する
function goChapterTest(ch){
  window._aslTestFocus = ch;
  window.switchTab('work' + ch);
  // 設問が描かれるのを待ってから移動する（固定の待ち時間だと、遅い回線で先に動いてしまう）
  let tries = 0;
  const go = () => {
    const el = document.getElementById('test-body-' + ch);
    if(el && el.children.length){ el.scrollIntoView({ block:'start', behavior:'smooth' }); return; }
    if(++tries < 30) setTimeout(go, 100);
  };
  go();
}
window.goChapterTest = goChapterTest;

// 受講実績の要約（修了条件・章ごとの提出／合格）をサーバーから取り直し、画面に反映する（2026-11-23）。
// ログイン時の1回だけだと、その場でワーク提出・テスト合格・試験合格をしても修了証の申請ボタンが押せなかった
async function refreshRecords(){
  if(!currentUser || !(currentUserPlan === 'standard' || currentUserPlan === 'premium')) return;
  try {
    const r = await apiFetch('/api/records', { method: 'GET' });
    const sm = r && r.summary;
    if(!sm) return;
    window._recSummary = sm;
    currentCertifiedEarlier = !!sm.certifiedEarlier;
    if(sm.examPassed) currentExamPassed = true;
    (sm.workDoneList || []).forEach(c => { _workSubmittedCh[c] = true; });
    (sm.testPassedList || []).forEach(c => { _testResultCh[c] = Object.assign(_testResultCh[c] || {}, { passed: true }); });
    updateCompletionUI();
    try { buildSideNav(loadCompleted()); } catch(e){}
  } catch(e) { console.warn('受講実績を取得できませんでした:', e.message); }
}
window.refreshRecords = refreshRecords;

function unlockCompleteBtn(tabId){
  const btn = document.getElementById('complete-' + tabId);
  if(btn){
    btn.disabled = false;
    btn.style.opacity = '1';
    btn.style.pointerEvents = 'auto';
    btn.style.borderColor = 'var(--green)';
    btn.style.color = 'var(--green)';
    // 「正解しました！受講済みにできます」メッセージ
    const note = document.getElementById('note-' + tabId);
    if(note && !btn.classList.contains('done')){
      note.innerHTML = '✅ 正解！受講済みボタンが解放されました';
      note.style.color = 'var(--green)';
    }
  }
}

/* モーダルの外側クリックで閉じる */
document.getElementById('auth-overlay').addEventListener('click', function(e){
  if(e.target === this) closeAuthModal();
});

/* Enterキーでログイン */
document.getElementById('auth-pw').addEventListener('keydown', e=>{
  if(e.key==='Enter') signInWithEmail();
});

/* 初期化 */
lockPaidTabs(true);
initSupabase();

/* ESMスコープ内の関数をグローバルに公開（onclick属性から呼べるように） */

// ---- プランによるアクセス判定（全ての遷移経路で必ず通す） ----
// 以前は上書き版 switchTab が showActiveLesson を直接呼んでおり、
// ログイン中はプラン制限が一切効かなかった（Basicでも L23〜L34 が閲覧できた）。
// タブを開けるかどうかの判定は、この1か所だけで行う（2026-11-23）。
// 以前は switchTab・canAccessTab・tabOpenable・applyPlanLock がそれぞれ別の条件を持っており、
// 片方だけ直して食い違うことがあった。
// 戻り値：null＝開ける／'login'＝ログインが必要／{msg, upgrade}＝開けない（理由と、プラン案内を出すか）
function tabDenyReason(name){
  if(FREE_TABS.includes(name)) return null;
  if(!currentUser) return 'login';
  const plan = currentUserPlan || 'free';
  const paid = plan === 'basic' || plan === 'standard' || plan === 'premium';
  if(name === 'records') return null;
  if(name === 'exam') return plan === 'premium' ? null : { msg: '認定試験は Premium プランで受験できます。', upgrade: true };
  if(name === 'certificate') return (plan === 'standard' || plan === 'premium') ? null : { msg: '修了証は Standard 以上のプランで発行できます。', upgrade: true };
  if(!paid) return { msg: '序章と L01 は無料でお読みいただけます。L02 以降はご購入後に受講できます。', upgrade: true };
  if(PREMIUM_TABS.includes(name) && plan !== 'premium') return { msg: 'このレッスンは Premium プランで受講できます。', upgrade: true };
  if(STANDARD_TABS.includes(name) && plan === 'basic') return { msg: 'このレッスンは Standard 以上のプランで受講できます。', upgrade: true };
  if(ALL_LESSONS_LIST.includes(name) && !getPlanLessons().includes(name)) return { msg: 'このレッスンはご購入後に受講できます。', upgrade: true };
  return null;
}
// 開けないときは理由を伝え、プランの案内へ進めるようにする
function canAccessTab(name){
  if(name === 'certificate' && typeof refreshRecords === 'function') setTimeout(refreshRecords, 0);
  const r = tabDenyReason(name);
  if(!r) return true;
  if(r === 'login'){ openAuthModal(); return false; }
  // ブラウザの確認ダイアログではなく、画面内のプラン案内に理由を添えて出す（2026-11-23）
  if(r.upgrade) showPlans(r.msg);
  else alert(r.msg);
  return false;
}
// 料金プランの案内。ログイン中は学習画面がトップページの料金欄を隠しているので、
// 今のプランより上のプランだけを小さな画面に出し、そこから申込フォームを開けるようにする（2026-11-23）
function showPlans(reason){
  if(!currentUser){ const p = document.getElementById('plans'); if(p) p.scrollIntoView({ behavior: 'smooth' }); return; }
  closePlanPick();
  const order = ['basic', 'standard', 'premium'];
  const cur = order.indexOf(currentUserPlan);
  const rows = order.filter((p, i) => i > cur).map(p => `<div class="pp-row">
      <div class="pp-info"><div class="pp-name">${p.toUpperCase()}<span class="pp-price">${yen(COURSE.prices[p])}（税込）</span></div>
      <div class="pp-desc">${escHtml(PLAN_RANGE_TEXT[p])}<br>標準学習時間 ${escHtml(PLAN_HOURS[p])}${p === 'basic' ? '' : '・修了証あり'}</div></div>
      <button type="button" class="form-submit pp-btn" onclick="closePlanPick();openFormModal('${FORM_PLAN_NAME[p]}')">申し込む</button></div>`).join('');
  const ov = document.createElement('div');
  ov.id = 'plan-pick-overlay';
  ov.setAttribute('role', 'dialog'); ov.setAttribute('aria-modal', 'true'); ov.setAttribute('aria-labelledby', 'pp-title');
  ov.innerHTML = `<div class="pp-box"><button type="button" class="pp-close" onclick="closePlanPick()" aria-label="閉じる">×</button>
    <div class="pp-title" id="pp-title">${reason ? '🔒 ' + escHtml(reason) : '受講範囲を広げる'}</div>
    ${reason ? '<p class="pp-lead">上のプランにお申し込みいただくと受講できます。</p>' : ''}
    ${rows || '<p class="pp-desc">いまのプランで全レッスンを受講できます。</p>'}
    <p class="pp-note">お支払いは決済画面（Stripe）で行います。ご不明点はお問い合わせください。</p></div>`;
  ov.addEventListener('click', e => { if(e.target === ov) closePlanPick(); });
  document.body.appendChild(ov);
  const b = ov.querySelector('.pp-btn, .pp-close'); if(b) b.focus();
}
function closePlanPick(){ const ov = document.getElementById('plan-pick-overlay'); if(ov) ov.remove(); }
window.closePlanPick = closePlanPick;
window.showPlans = showPlans;
window.canAccessTab = canAccessTab;
