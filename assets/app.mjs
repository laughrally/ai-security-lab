// AI Security Lab — 受講画面と LP のスクリプト。
// 正本は assets/src/*.part.js（機能ごと）。asl-worker/course/course-sync.py が番号の順につないで assets/app.mjs を作る。
// app.mjs を直接直さないこと（check-all.sh が食い違いを止める）。つないだ全体で1つの ES モジュールになる。
// 冒頭の COURSE ブロックは course.json から書き出す
// <COURSE:BEGIN> course/course-sync.py が course.json から書き出す。手で直さない
const COURSE = {"free": ["intro", "lesson1"], "plans": {"basic": 5, "standard": 6, "premium": 7}, "workChapters": {"basic": 0, "standard": 6, "premium": 7}, "prices": {"basic": 9800, "standard": 59800, "premium": 149800}, "rules": {"quizMaxAttempts": 3, "workQuestions": 4, "testQuestions": 5, "testPass": 4, "testMaxAttempts": 3, "examQuestions": 60, "examPass": 48, "examMinutes": 60, "examCooldownHours": 24, "lessonMaxMinutes": 180, "corpMaxSeats": 50}, "chapters": [{"no": 0, "tag": "序章", "name": "なぜ今、AIセキュリティを学ぶ必要があるのか", "color": "var(--red)", "lessons": ["intro"]}, {"no": 1, "tag": "第1章", "name": "AIと情報漏洩の基礎知識", "color": "var(--red)", "lessons": ["lesson1", "lesson2", "lesson3", "lesson4"], "workGoal": "自社のAI利用実態マップ"}, {"no": 2, "tag": "第2章", "name": "ツール別リスクと設定", "color": "var(--amber)", "lessons": ["lesson5", "lesson6", "lesson7", "lesson8"], "workGoal": "入力禁止情報リスト"}, {"no": 3, "tag": "第3章", "name": "法律・著作権・ハルシネーション", "color": "var(--amber)", "lessons": ["lesson9", "lesson10", "lesson11"], "workGoal": "リスク評価シート"}, {"no": 4, "tag": "第4章", "name": "組織・ガイドライン整備", "color": "var(--blue)", "lessons": ["lesson12", "lesson13", "lesson14", "lesson15"], "workGoal": "社内AI利用ガイドライン"}, {"no": 5, "tag": "第5章", "name": "実践：事例研究とツール別ガイド", "color": "var(--blue)", "lessons": ["lesson16", "lesson17", "lesson28", "lesson29", "lesson30", "lesson31", "lesson32"], "workGoal": "AI導入計画メモ"}, {"no": 6, "tag": "第6章", "name": "AI攻撃手口と防衛策・管理者設定", "color": "#0ea5e9", "lessons": ["lesson18", "lesson19", "lesson20", "lesson21", "lesson22", "lesson33", "lesson34"], "badge": "Standard+", "workGoal": "インシデント対応フロー"}, {"no": 7, "tag": "第7章", "name": "AI技術とセキュリティの深層・開発", "color": "#8b5cf6", "lessons": ["lesson23", "lesson24", "lesson25", "lesson26", "lesson27", "lesson35", "lesson36"], "badge": "Premium", "workGoal": "AI活用ロードマップ"}], "titles": {"intro": "序章：なぜ今学ぶか", "lesson1": "入力情報はどこへ行くのか", "lesson2": "学習オフ設定の意味と限界", "lesson3": "シャドーAIとは何か", "lesson4": "クラウドAIとオンプレAIの違い", "lesson5": "ChatGPT各プランの違い", "lesson6": "Claude・Gemini・Copilot比較", "lesson7": "社外秘・個人情報の判断基準", "lesson8": "業種別リスク事例", "lesson9": "AIと個人情報保護法", "lesson10": "著作権侵害になるケース", "lesson11": "ハルシネーションと法的責任", "lesson12": "ガイドライン作成テンプレート", "lesson13": "インシデント初動対応フロー", "lesson14": "中小企業の最低限の体制", "lesson15": "今日からできる10のアクション", "lesson16": "よくある事例集（20事例）", "lesson17": "主要AIツール完全比較", "lesson18": "AIを使ったフィッシング詐欺", "lesson19": "プロンプトインジェクション攻撃", "lesson20": "ディープフェイクと音声詐欺", "lesson21": "AIエージェント・自動化ツールのリスク", "lesson22": "インシデント対応ロールプレイ", "lesson23": "LLMの仕組み入門", "lesson24": "プロンプトエンジニアリング基礎", "lesson25": "RAG・ファインチューニングのリスク", "lesson26": "業種別AIロードマップ", "lesson27": "MCPとAIエージェント最前線", "lesson28": "ChatGPTを安全に使う", "lesson29": "GeminiとGoogle Workspace", "lesson30": "CopilotとMicrosoft 365", "lesson31": "Claudeを安全に使う", "lesson32": "MetaのMuseとSNSのAI", "lesson33": "AIブラウザとエージェント機能", "lesson34": "管理者のための設定集", "lesson35": "Claude Codeを安全に使う", "lesson36": "AIコーディングと秘密情報"}};
const FREE_TABS = COURSE.free;
// 受講順の内部ID（序章を除く）。内部IDは追加順なので番号順ではない。表示番号はこの並び順
const ALL_LESSONS_LIST = COURSE.chapters.filter(c => c.no >= 1).flatMap(c => c.lessons);
const PLAN_LESSON_COUNT = Object.fromEntries(Object.entries(COURSE.plans).map(([p, n]) => [p, COURSE.chapters.filter(c => c.no >= 1 && c.no <= n).reduce((s, c) => s + c.lessons.length, 0)]));
const PLAN_RANGE_LABEL = Object.fromEntries(Object.entries(COURSE.plans).map(([p, n]) => [p, '第1〜' + n + '章']));
const _chLessons = (lo, hi) => COURSE.chapters.filter(c => c.no > lo && c.no <= hi).flatMap(c => c.lessons);
const _works = (lo, hi) => Array.from({ length: Math.max(0, hi - lo) }, (_, i) => 'work' + (lo + i + 1));
// Basic は章末ワークなし。Standard の章までのワークは Standard 以上、それより後は Premium
const STANDARD_TABS = [..._chLessons(COURSE.plans.basic, COURSE.plans.standard), ..._works(0, COURSE.plans.standard)];
const PREMIUM_TABS = [..._chLessons(COURSE.plans.standard, COURSE.plans.premium), ..._works(COURSE.plans.standard, COURSE.plans.premium)];
// 各章の最後のレッスン → 章番号（章末ワークをその後ろに挟む）
const CH_END = Object.fromEntries(COURSE.chapters.filter(c => c.no >= 1).map(c => [c.lessons[c.lessons.length - 1], c.no]));
const LESSON_LABELS = Object.fromEntries([['intro', COURSE.titles.intro], ...ALL_LESSONS_LIST.map((id, i) => [id, 'L' + String(i + 1).padStart(2, '0') + '：' + COURSE.titles[id]])]);
const CHAPTER_DATA = COURSE.chapters.map(c => Object.assign({ tag: c.tag, name: c.name, lessons: c.lessons, color: c.color }, c.badge ? { plan: c.badge } : {}));
// 章ごとの最初のレッスン（内部番号）と、表示番号での範囲（例：L23〜L29）
const CH_FIRST_LESSON = Object.fromEntries(COURSE.chapters.filter(c => c.no >= 1).map(c => [c.no, parseInt(c.lessons[0].slice(6), 10)]));
const CH_RANGE_LABEL = Object.fromEntries(COURSE.chapters.filter(c => c.no >= 1).map(c => { const a = ALL_LESSONS_LIST.indexOf(c.lessons[0]) + 1, b = ALL_LESSONS_LIST.indexOf(c.lessons[c.lessons.length - 1]) + 1; return [c.no, 'L' + String(a).padStart(2, '0') + '〜L' + String(b).padStart(2, '0')]; }));
const WORK_CHAPTERS = Object.fromEntries(COURSE.chapters.filter(c => c.workGoal).map(c => [c.no, { name: c.name, goal: c.workGoal, lessons: CH_RANGE_LABEL[c.no] }]));
// 標準学習時間と受講範囲の表記（受講実績レポート・修了証の表示に使う）
const PLAN_HOURS = {"basic": "約6時間10分", "standard": "約14時間", "premium": "約18時間10分"};
const PLAN_RANGE_TEXT = {"basic": "序章＋第1〜5章（全22レッスン）", "standard": "序章＋第1〜6章（全29レッスン）＋章末ワーク・確認テスト（全6章）", "premium": "序章＋第1〜7章（全36レッスン）＋章末ワーク・確認テスト（全7章）＋認定試験"};
const PLAN_LABELS_JA = Object.fromEntries(Object.keys(COURSE.plans).map(p => [p, { name: p[0].toUpperCase() + p.slice(1) + 'プラン', range: PLAN_RANGE_TEXT[p], hours: PLAN_HOURS[p] }]));
// <COURSE:END>

// HTML に文字を入れるときのエスケープ
function escHtml(t){ return String(t == null ? '' : t).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
// CSV の1セル：数式として解釈されうる先頭（= + - @ タブ 改行）を無害化し、ダブルクォートで囲む
function csvCell(v){ let t = String(v == null ? '' : v); if(/^[=+\-@\t\r]/.test(t)) t = "'" + t; return '"' + t.replace(/"/g, '""') + '"'; }
window.escHtml = escHtml;
// 端末への保存（localStorage）の読み書き。
const store = {
  get(k, def = null){ try { const v = localStorage.getItem(k); return v == null ? def : v; } catch(e){ return def; } },
  getJSON(k, def){ try { const v = localStorage.getItem(k); return v == null ? def : JSON.parse(v); } catch(e){ return def; } },
  set(k, v){ try { localStorage.setItem(k, typeof v === 'string' ? v : JSON.stringify(v)); } catch(e){} },
  remove(k){ try { localStorage.removeItem(k); } catch(e){} },
};

// バージョンを固定する。@2 のままだとCDN側の更新がそのまま本番に入るため。Supabase のライブラリは自前で配信する。
// 外部の配信元の改ざんや停止の影響を受けないように。取得時に jsDelivr と unpkg の SHA-256 が一致することを確かめた
const { createClient } = window.supabase;



// ---- インラインのイベントハンドラ（onclick="…"）を、CSP で 'unsafe-inline' を外しても動かす----
// ページ・目次・レッスン本文（DB から配信）には onclick="switchTab('lesson3');return false;" のような書き方が
// 多数あり、すべてを書き換えるのは本文の入れ直しも要る。そこで、ブラウザがインラインの実行を止めているときだけ、
// 属性の文字を小さな構文解析器で読み、決まった形の呼び出しだけを実行する。
//   使える形：関数(引数, …) ／ window.関数(…) ／ return 式 ／ if(event.target===this) 文
//   引数：'文字列' "文字列" 数 true false null this event this.value document.querySelector('…')
//   呼べる関数：90-inline-actions の INLINE_ACTIONS にある名前だけ（それ以外は fetch も自前の関数も呼べない）
function scrollToId(id){ const el = document.getElementById(id); if(el) el.scrollIntoView({ behavior: 'smooth' }); }
const inlineShim = (() => {
  // 呼べるのは INLINE_ACTIONS（90-inline-actions）にある名前だけ。一覧は呼ぶときに参照する（このファイルより後で定義されるため）
  const action = (name) => (typeof INLINE_ACTIONS !== 'undefined' && Object.prototype.hasOwnProperty.call(INLINE_ACTIONS, name)) ? INLINE_ACTIONS[name] : null;
  function tokenize(src){
    const t = []; let i = 0;
    while(i < src.length){
      const c = src[i];
      if(/\s/.test(c)){ i++; continue; }
      if(c === "'" || c === '"'){
        let j = i + 1, v = '';
        while(j < src.length && src[j] !== c){ if(src[j] === '\\'){ j++; } v += src[j]; j++; }
        if(j >= src.length) throw new Error('文字列が閉じていない');
        t.push({ k: 's', v }); i = j + 1; continue;
      }
      const num = /^-?\d+(\.\d+)?/.exec(src.slice(i));
      if(num && (c === '-' || /\d/.test(c))){ t.push({ k: 'n', v: Number(num[0]) }); i += num[0].length; continue; }
      const id = /^[A-Za-z_$][\w$]*(\.[A-Za-z_$][\w$]*)*/.exec(src.slice(i));
      if(id){ t.push({ k: 'i', v: id[0] }); i += id[0].length; continue; }
      if(src.startsWith('===', i)){ t.push({ k: 'p', v: '===' }); i += 3; continue; }
      if('(),;'.includes(c)){ t.push({ k: 'p', v: c }); i++; continue; }
      throw new Error('使えない文字: ' + c);
    }
    return t;
  }
  // dry が true なら、構文と呼べる関数かを確かめるだけで実行しない（テスト用）
  function run(code, el, ev, dry){
    const t = tokenize(code); let p = 0;
    const peek = () => t[p], eat = (v) => { const x = t[p]; if(!x || (v !== undefined && x.v !== v)) throw new Error('構文: ' + code); p++; return x; };
    function value(){
      const x = eat();
      if(x.k === 's' || x.k === 'n') return x.v;
      if(x.k !== 'i') throw new Error('値: ' + code);
      if(x.v === 'true') return true;
      if(x.v === 'false') return false;
      if(x.v === 'null') return null;
      if(x.v === 'this') return el;
      if(x.v === 'event') return ev;
      if(x.v === 'this.value') return el && el.value;
      if(x.v === 'event.target') return ev && ev.target;
      if(x.v === 'document.querySelector'){ eat('('); const sel = value(); eat(')'); return dry ? null : document.querySelector(String(sel)); }
      if(peek() && peek().v === '('){ return call(x.v); }
      throw new Error('値: ' + x.v);
    }
    function call(name){
      const fname = name.startsWith('window.') ? name.slice(7) : name;
      if(fname.includes('.')) throw new Error('呼べない: ' + name);
      const fn = action(fname);
      if(typeof fn !== 'function') throw new Error('呼べない: ' + name);
      eat('('); const args = [];
      if(peek() && peek().v !== ')'){ args.push(value()); while(peek() && peek().v === ','){ eat(','); args.push(value()); } }
      eat(')');
      return dry ? undefined : fn.apply(el, args);
    }
    function stmt(){
      const x = peek();
      if(!x) return undefined;
      if(x.k === 'i' && x.v === 'return'){ eat(); return { ret: value() }; }
      if(x.k === 'i' && x.v === 'if'){
        eat(); eat('('); const a = value(); eat('==='); const b = value(); eat(')');
        if(a === b || dry) return stmt(); else { while(peek() && peek().v !== ';') eat(); return undefined; }
      }
      value(); return undefined;
    }
    let r;
    while(p < t.length){
      const s = stmt();
      if(s && 'ret' in s){ r = s.ret; break; }
      if(peek() && peek().v === ';') eat(';');
    }
    return r;
  }
  // ブラウザがインラインのハンドラを実行するかを確かめる（実行するなら何もしない＝二重に動かさない）
  let native = false;
  try {
    const probe = document.createElement('div');
    probe.setAttribute('onclick', 'window.__aslInlineOk=1');
    probe.click();
    native = window.__aslInlineOk === 1;
    delete window.__aslInlineOk;
  } catch(e){}
  if(!native){
    const handle = (type) => (ev) => {
      for(let el = ev.target; el && el !== document; el = el.parentElement){
        if(!(el instanceof Element)) continue;
        const code = el.getAttribute('on' + type);
        if(!code) continue;
        try {
          if(run(code, el, ev) === false) ev.preventDefault();
        } catch(e){ console.error('ハンドラを実行できません:', code, e.message); }
        if(ev.cancelBubble) break;
      }
    };
    ['click', 'change', 'input'].forEach(type => document.addEventListener(type, handle(type)));
  }
  return { native, run };
})();
window.__inlineShim = inlineShim;
/* ===== 申込フォーム ===== */
const PRICE_ID_MAP = {
  'ベーシック 9800円':    'price_1TJYfK2OQszM8w2hEWWM5R36',
  'スタンダード 59800円': 'price_1TP0562OQszM8w2hbUgLZ5if',
  'プレミアム 149800円':   'price_1TP05i2OQszM8w2hhNpHkmfG',
};
let currentPriceId = '';
// 申込フォームのプラン名（決済のメタデータにも入るので文字は変えない）→ course.json のプラン
const FORM_PLAN_KEY = { 'ベーシック 9800円': 'basic', 'スタンダード 59800円': 'standard', 'プレミアム 149800円': 'premium' };
const FORM_PLAN_NAME = Object.fromEntries(Object.entries(FORM_PLAN_KEY).map(([k, v]) => [v, k]));
const yen = n => '¥' + Number(n).toLocaleString('ja-JP');
let currentPlanName = '';

function toggleSubmitBtn(){
  const checked = document.getElementById('form-agree').checked;
  const btn = document.getElementById('form-submit-btn');
  btn.disabled = !checked;
  btn.style.opacity = checked ? '1' : '0.5';
  btn.style.cursor = checked ? 'pointer' : 'not-allowed';
}

function openFormModal(plan){
  document.getElementById('form-overlay').classList.remove('hidden');
  currentPriceId = PRICE_ID_MAP[plan] || '';
  currentPlanName = plan;
  // 表記は course.json の価格と受講範囲から作る
  const key = FORM_PLAN_KEY[plan];
  const range = key ? '第1〜' + COURSE.plans[key] + '章・' + (key === 'premium' ? '全' : '') + PLAN_LESSON_COUNT[key] + 'レッスン' : '';
  document.getElementById('form-plan-label').textContent = key ? key.toUpperCase() + ' · ' + yen(COURSE.prices[key]) : plan;
  document.getElementById('legal-plan-name').textContent = key ? key.toUpperCase() + 'プラン（' + range + '）' : plan;
  document.getElementById('legal-plan-price').textContent = key ? yen(COURSE.prices[key]) + '（税込）' : '';
  // 同意チェックをリセット
  document.getElementById('form-agree').checked = false;
  toggleSubmitBtn();
  setFormMsg('','');
  // ベーシック以外はコンサルフィールドを表示
  const isConsult = key !== 'basic';
  document.getElementById('form-consult-fields').style.display = isConsult ? 'block' : 'none';
  // フィールドリセット
  ['form-name','form-company','form-tel','form-email','form-challenge'].forEach(id=>{
    const el = document.getElementById(id);
    if(el) el.value = '';
  });
  document.getElementById('form-contact-email').checked = true;
}
function closeFormModal(){
  document.getElementById('form-overlay').classList.add('hidden');
}
function setFormMsg(type, text){
  const el = document.getElementById('form-msg');
  el.className = 'form-msg' + (type ? ' '+type : '');
  el.textContent = text;
}

async function submitForm(){
  const name    = document.getElementById('form-name').value.trim();
  const company = document.getElementById('form-company').value.trim();
  const tel     = document.getElementById('form-tel').value.trim();
  const email   = document.getElementById('form-email').value.trim();
  const isConsult = FORM_PLAN_KEY[currentPlanName] !== 'basic';
  const challenge = isConsult ? document.getElementById('form-challenge').value.trim() : '';
  const contactPref = isConsult ? (document.querySelector('input[name="form-contact"]:checked')?.value || 'メール') : '';

  if(!name || !tel || !email){
    setFormMsg('error','お名前・電話番号・メールアドレスは必須です。');
    return;
  }
  if(!/^[^@]+@[^@]+\.[^@]+$/.test(email)){
    setFormMsg('error','メールアドレスの形式が正しくありません。');
    return;
  }
  if(!currentPriceId){
    setFormMsg('error','プランの選択でエラーが発生しました。再度お試しください。');
    return;
  }

  const btn = document.getElementById('form-submit-btn');
  btn.disabled = true;
  btn.textContent = '処理中...';
  setFormMsg('','');

  try {
    const res = await fetch(WORKER_URL + '/create-checkout', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        priceId: currentPriceId,
        email,
        metadata: { name, company, tel, challenge, contactPref, plan: currentPlanName }
      }),
    });
    const data = await res.json();
    if(data.url){
      window.location.href = data.url;
    } else {
      throw new Error(data.error || '決済画面の取得に失敗しました。');
    }
  } catch(e) {
    setFormMsg('error', e.message || '送信に失敗しました。時間をおいて再度お試しください。');
    btn.disabled = false;
    btn.textContent = '決済画面へ進む →';
  }
}

/* フォームの外側クリックで閉じる */
document.getElementById('form-overlay').addEventListener('click', function(e){
  if(e.target === this) closeFormModal();
});

/* ===== ダークモード ===== */
function toggleDark(){
  const isDark = document.body.classList.toggle('dark');
  document.getElementById('dark-toggle').textContent = isDark ? '🌙' : '☀️';
  store.set('theme', isDark ? 'dark' : 'light');
}
// 前回の設定は index.html の先頭で本文の描画前に当てている。ここではボタンの表示だけ合わせる
{ const btn = document.getElementById('dark-toggle'); if(btn && document.body.classList.contains('dark')) btn.textContent = '🌙'; }

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
      // ログイン・ログアウトで人が切り替わったら読み直す。
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

// この端末に残る受講データ（受講済み・学習時間）は、最後にログインした人のもの。別の人がログインしたら先に消す。
// ログアウトせずにセッションが切れたあと、次の人の受講済みや学習時間に前の人の分が混ざらないように
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
  // .single() は失敗しても例外を投げず { error } を返す。2回まで取り直し、それでも駄目なら知らせる
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
    // 読めなかったときは Worker と同じく「プランなし」に倒す
    currentUserPlan = currentUserPlan || 'free';
    currentExamPassed = false;
    console.warn('プランを取得できませんでした:', error && error.message);
    alert('ご契約プランの情報を読み込めませんでした。通信環境をご確認のうえ、ページを再読み込みしてください。');
  }
  if(!window._corpLoaded){ window._corpLoaded = true; loadCorpProgress(); }
  if(currentUserPlan === 'standard' || currentUserPlan === 'premium'){
    window._recordsPromise = refreshRecords();
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

    // 受講済みはサーバーの記録で置き換える。
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
  document.querySelectorAll('.tab[data-tab]').forEach(t=>{
    const name = t.dataset.tab;
    if(!FREE_TABS.includes(name)){
      t.classList.toggle('locked', lock);
    }
  });
}

// プランの範囲外の行に鍵を付ける（判定は tabDenyReason）
function applyPlanLock(){
  document.querySelectorAll('.sb-row, .lrow').forEach(el => {
    // どのタブへの行かは data-tab 属性で持つ（onclick の文字は読まない）。章末テストの行は、その章のワークと同じ判定
    const m = el.dataset.tab ? [null, el.dataset.tab] : null;
    // キーボード・読み上げで使えるように
    if(el.classList.contains('sb-row') && el.hasAttribute('onclick')){
      if(!el.hasAttribute('tabindex')) el.setAttribute('tabindex', '0');
      if(!el.hasAttribute('role')) el.setAttribute('role', 'button');
      if(el.classList.contains('active')) el.setAttribute('aria-current', 'page'); else el.removeAttribute('aria-current');
      const t = el.querySelector('.sb-row-title');
      if(t && el.classList.contains('done') && !t.querySelector('.sr-only')) t.insertAdjacentHTML('beforeend', '<span class="sr-only">（受講済み）</span>');
    }
    if(!m) return;
    const id = m[1];
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
  LOCAL_LEARN_KEYS.forEach(k => store.remove(k));
  currentCertifiedEarlier = false;
  currentExamPassed = false;
  await _supabase.auth.signOut();
  updateCompletionUI();
  updateDashboard();
}

/* ===== タブ切り替え（認証チェック付き） ===== */
// 未ログインのときのタブ表示の切り替え（トップページのタブ）。開けるかどうかは window.switchTab が判定済み。
function showTabGuest(name){
  document.querySelectorAll('.tab[data-tab]').forEach(t => t.classList.toggle('active', t.dataset.tab === name));
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

// ---- 受講時間の実測----
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
const MAX_TOTAL_S = COURSE.rules.lessonMaxMinutes * 60;   // 1レッスンの上限（Worker の LESSON_MAX_SECONDS と同じ値を course.json から）
const TRACK_RE = /^(lesson\d+|work\d+)$/;

let activeItem  = null;   // いま表示している lessonN / workN（それ以外は null）
let tickAt      = null;   // 前回の加算時刻
let lastInputAt = Date.now();
let wasVisible  = (document.visibilityState === 'visible');

function loadActive(){ return store.getJSON(ACTIVE_KEY, {}) || {}; }
function saveActive(o){ store.set(ACTIVE_KEY, o); }
function loadStarts(){ return store.getJSON(START_KEY, {}) || {}; }
function loadSynced(){ return store.getJSON(SYNC_KEY, {}) || {}; }
function saveSynced(o){ store.set(SYNC_KEY, o); }

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
      store.set(START_KEY, st);
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
// 答え合わせの先頭の記号。解説が「✅ 正解です」で始まるときは付けない
function quizMark(r){
  const ex = String(r.explanation || '');
  return /^[✅❌]/.test(ex) ? '' : (r.isCorrect ? '✅ ' : '❌ ');
}
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
    if(fbEl && r.explanation) fbEl.innerHTML = quizMark(r) + escHtml(r.explanation);

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
        const done = store.getJSON(quizDoneKey(), {}) || {};
        done[id] = { s: selected, e: r.explanation || '' };
        store.set(quizDoneKey(), done);
      } catch(e){}
    } else {
      // 不正解：この選択肢だけ無効化。他は引き続き選べる。
      btn.style.pointerEvents = 'none'; btn.disabled = true; btn.setAttribute('aria-disabled', 'true');
      btn.classList.add('ng');
      document.getElementById(id+'-ng').classList.add('show');
      if(r.attemptsLeft === 0) quizAttemptsUsedUp(qopts, id, tabId);
      else if(Number.isInteger(r.attemptsLeft)){
        const ng = document.getElementById(id + '-ng');
        if(ng){ let left = ng.querySelector('.q-left'); if(!left){ left = document.createElement('div'); left.className = 'q-used-up q-left'; ng.appendChild(left); }
          left.textContent = 'この設問は、あと ' + r.attemptsLeft + ' 回答えられます。'; }
      }
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
    note.textContent = '回答できる回数（' + COURSE.rules.quizMaxAttempts + '回）を使い切りました。レッスンを読み直してから、受講済みにしてください（正誤は受講実績レポートに記録されます）。';
    if(!ng.querySelector('.q-used-up')) ng.appendChild(note);
  }
  unlockCompleteBtn(tabId);
}

// 未ログインで無料公開レッスンの設問に答えたとき。答え合わせと解説だけを出し、記録はしない
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
    if(fb) fb.innerHTML = quizMark(r) + escHtml(r.explanation || '') + note;
    if(r.isCorrect){
      qopts.querySelectorAll('.qopt').forEach(b => { b.style.pointerEvents = 'none'; });
      btn.classList.add('ok');
      document.getElementById(id + '-ok').classList.add('show');
      document.getElementById(id + '-ng').classList.remove('show');
    } else {
      btn.style.pointerEvents = 'none'; btn.disabled = true; btn.setAttribute('aria-disabled', 'true');
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
    // 回答済みなら、結果（合否・学び直す）の欄へ。まだなら設問の先頭へ
    const sum = document.getElementById('test-summary-' + ch);
    if(sum && sum.textContent.trim()){ sum.scrollIntoView({ block:'center', behavior:'smooth' }); return; }
    if(el && el.children.length){ el.scrollIntoView({ block:'start', behavior:'smooth' }); return; }
    if(++tries < 30) setTimeout(go, 100);
  };
  go();
}

// 受講実績の要約（修了条件・章ごとの提出／合格）をサーバーから取り直し、画面に反映する。
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

// ---- プランによるアクセス判定（全ての遷移経路で必ず通す） ----
// タブを開けるかどうかの判定は、この1か所（tabDenyReason）だけで行う。
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
  // ブラウザの確認ダイアログではなく、画面内のプラン案内に理由を添えて出す
  if(r.upgrade) showPlans(r.msg);
  else alert(r.msg);
  return false;
}
// 料金プランの案内。
// ログイン中は学習画面がトップページの料金欄を隠しているので、今のプランより上のプランだけを小さな画面に出し、そこから申込フォームを開けるようにする
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
window.canAccessTab = canAccessTab;
// ===================== 法的ページモーダル =====================
const LEGAL_PAGES = {
  privacy: `
    <h1>プライバシーポリシー</h1>
    <p>合同会社LaughRally（以下「当社」）は、AI Security Lab（以下「本サービス」）における利用者の個人情報の取扱いについて、以下のとおり定めます。</p>

    <h2>1. 取得する情報</h2>
    <p>本サービスでは、次の情報を取得します。</p>
    <ul>
      <li>メールアドレス（アカウント登録時）</li>
      <li>メールアドレス（お役立ち情報の配信をお申し込みいただいた場合）</li>
      <li>お名前・会社名（修了証および受講実績レポートの発行を申請された場合）</li>
      <li>受講履歴（レッスンの開始・完了日時、学習時間）</li>
      <li>課題および確認テストの回答内容と評価結果</li>
      <li>決済に関する情報（決済代行事業者を通じて処理され、当社はカード番号を保持しません）</li>
    </ul>

    <h2>2. 利用目的</h2>
    <ul>
      <li>本サービスの提供および受講管理</li>
      <li>修了証・受講実績レポートの発行</li>
      <li>お問い合わせへの対応</li>
      <li>本サービスの改善および教材の更新</li>
      <li>重要なお知らせの連絡</li>
      <li>お申し込みいただいた方への、生成AIの情報管理に関するお役立ち情報の配信（各メールに記載のリンクからいつでも配信を停止できます）</li>
    </ul>

    <h2>3. 第三者提供</h2>
    <p>当社は、法令に基づく場合を除き、あらかじめ本人の同意を得ることなく個人情報を第三者に提供しません。</p>
    <p>ただし、本サービスの運営に必要な範囲で、以下の事業者に業務を委託しています。</p>
    <table>
      <tr><th>委託先</th><th>委託する業務</th></tr>
      <tr><td>Stripe, Inc.</td><td>決済処理</td></tr>
      <tr><td>Supabase, Inc.</td><td>アカウント管理・受講データの保管</td></tr>
      <tr><td>Cloudflare, Inc.</td><td>アプリケーションの実行基盤</td></tr>
      <tr><td>Google LLC</td><td>修了証の生成および送信</td></tr>
    </table>

    <h2>4. 保管期間</h2>
    <p>受講データは、本サービスの提供に必要な期間および法令に定める期間、保管します。アカウントの削除をご希望の場合は、下記お問い合わせ先までご連絡ください。</p>
    <p>お役立ち情報の配信にお申し込みいただいたメールアドレスは、配信停止のお手続きをいただくまで保管します。配信停止後は、誤って再配信することを防ぐ目的でのみ保持し、配信には使用しません。</p>

    <h2>5. 開示・訂正・削除の請求</h2>
    <p>ご本人からの個人情報の開示、訂正、利用停止、削除のご請求には、本人確認のうえ、合理的な期間内に対応します。</p>

    <h2>6. Cookie等の利用</h2>
    <p>本サービスは、ログイン状態の保持および受講進捗の保存のためにブラウザのローカルストレージを使用します。広告目的でのトラッキングは行っていません。</p>

    <h2>7. お問い合わせ窓口</h2>
    <p>合同会社LaughRally<br>
    〒107-0062 東京都港区南青山2-2-15<br>
    <a href="mailto:info@laughrally.tech">info@laughrally.tech</a></p>

    <h2>8. 改定</h2>
    <p>本ポリシーの内容は、法令の改正等に応じて変更されることがあります。重要な変更を行う場合は、本サービス上でお知らせします。</p>

    <p style="margin-top:24px;font-size:13px">制定日：2026年4月22日<br>最終改定日：2026年8月17日</p>
  `,

  terms: `
    <h1>利用規約</h1>
    <p>本規約は、合同会社LaughRally（以下「当社」）が提供するAI Security Lab（以下「本サービス」）の利用条件を定めるものです。</p>

    <h2>第1条（適用）</h2>
    <p>本規約は、本サービスを利用するすべての方（以下「利用者」）に適用されます。</p>

    <h2>第2条（アカウント）</h2>
    <ul>
      <li>アカウントは1名につき1つとし、複数名での共有はできません。</li>
      <li>複数名でご利用の場合は、人数分のお申し込みが必要です。</li>
      <li>ログイン情報の管理は利用者の責任において行うものとします。</li>
    </ul>

    <h2>第3条（サービス内容）</h2>
    <ul>
      <li>本サービスは買い切り型で提供し、受講期限は設けていません。</li>
      <li>提供する教材の範囲は、お申し込みのプランによって異なります。</li>
      <li>法令の改正や技術動向を踏まえ、教材の内容を更新することがあります。この場合、追加のご負担なくご利用いただけます。</li>
    </ul>

    <h2>第4条（修了証）</h2>
    <ul>
      <li>修了証はStandard以上のプランを対象とし、所定の修了条件を満たした場合に発行します。</li>
      <li>修了条件は、ご契約プランの全レッスンの受講完了、全章の章末ワークの提出、全章の章末確認テストの合格とします（Premiumプランは認定試験の合格を含みます）。</li>
      <li>受講状況は当社のシステムが記録に基づいて確認します。</li>
    </ul>

    <h2>第5条（禁止事項）</h2>
    <p>利用者は、次の行為を行ってはなりません。</p>
    <ul>
      <li>教材の全部または一部を複製、転載、頒布、販売する行為</li>
      <li>アカウントを第三者に貸与、譲渡する行為</li>
      <li>本サービスの運営を妨害する行為</li>
      <li>不正な手段により受講記録を作出する行為</li>
      <li>法令または公序良俗に反する行為</li>
    </ul>

    <h2>第6条（知的財産権）</h2>
    <p>本サービスで提供する教材、設問、その他一切のコンテンツに関する著作権その他の知的財産権は、当社または正当な権利者に帰属します。</p>
    <p>なお、レッスン内で配布するガイドラインテンプレートについては、利用者ご自身の組織内での利用に限り、編集・複製を認めます。</p>

    <h2>第7条（免責）</h2>
    <ul>
      <li>本サービスは情報提供を目的とするものであり、特定の法的効果や結果を保証するものではありません。</li>
      <li>教材の内容を業務に適用した結果について、当社は責任を負いません。個別の法的判断が必要な場合は、専門家にご相談ください。</li>
      <li>通信環境や利用機器に起因する不具合について、当社は責任を負いません。</li>
    </ul>

    <h2>第8条（サービスの変更・中断）</h2>
    <p>当社は、システムの保守、障害への対応その他やむを得ない事由により、本サービスの全部または一部を一時的に中断することがあります。</p>

    <h2>第9条（返品・キャンセル）</h2>
    <p>デジタルコンテンツの性質上、購入完了後の返品・返金はお受けできません。詳細は特定商取引法に基づく表記をご確認ください。</p>

    <h2>第10条（準拠法・管轄）</h2>
    <p>本規約は日本法に準拠し、本サービスに関して紛争が生じた場合は、東京地方裁判所を第一審の専属的合意管轄裁判所とします。</p>

    <p style="margin-top:24px;font-size:13px">制定日：2026年4月22日<br>最終改定日：2026年8月17日</p>
  `,

  contact: `
    <h1>お問い合わせ</h1>
    <p>本サービスに関するご質問、受講に関するご相談、法人でのご利用のご検討など、下記までお気軽にご連絡ください。</p>

    <h2>お問い合わせ先</h2>
    <table>
      <tr><th>事業者名</th><td>合同会社LaughRally</td></tr>
      <tr><th>代表者</th><td>佐藤 誓哉</td></tr>
      <tr><th>所在地</th><td>〒107-0062 東京都港区南青山2-2-15</td></tr>
      <tr><th>メール</th><td><a href="mailto:info@laughrally.tech">info@laughrally.tech</a></td></tr>
      <tr><th>受付時間</th><td>平日 10:00〜18:00（土日祝を除く）</td></tr>
    </table>
    <p>3営業日以内にご返信いたします。お急ぎの場合はその旨を件名にご記載ください。</p>

    <h2>法人でのご利用について</h2>
    <p>複数名でのご受講、宛名を指定した見積書・請求書・領収書の発行、社内研修としての導入のご相談も承っています。ご希望の内容とご受講予定人数をお知らせください。</p>

    <h2>受講に関するご相談</h2>
    <p>ログインできない、修了証が発行できない、受講実績レポートが出力できないなどの場合は、ご登録のメールアドレスとあわせてお問い合わせください。</p>
  `
};

function openLegalModal(key){
  const el = document.getElementById('legal-content');
  const ov = document.getElementById('legal-overlay');
  if(!el || !ov || !LEGAL_PAGES[key]) return;
  el.innerHTML = LEGAL_PAGES[key];
  ov.classList.add('show');
  document.body.style.overflow = 'hidden';
  ov.scrollTop = 0;
}

function closeLegalModal(){
  const ov = document.getElementById('legal-overlay');
  if(!ov) return;
  ov.classList.remove('show');
  document.body.style.overflow = '';
}

document.addEventListener('keydown', e => {
  if(e.key === 'Escape') closeLegalModal();
});


// ===================== ガイドラインテンプレートのダウンロード =====================
function downloadGuideline(){
  const txt = [
'生成AI利用ガイドライン（テンプレート）',
'',
'〇〇株式会社',
'制定日：____年__月__日',
'',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'1. 目的',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'本ガイドラインは、当社における生成AIの業務利用にあたり、情報漏洩および',
'法令違反を防止するとともに、業務の効率化を安全に進めることを目的とする。',
'',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'2. 適用範囲',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'雇用形態を問わず、当社の業務に従事するすべての者に適用する。',
'（正社員・契約社員・パート・アルバイト・派遣社員・業務委託先を含む）',
'',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'3. 利用を認めるツール',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'以下のツールに限り、業務での利用を認める。',
'',
'  ・____________________（プラン：__________）',
'  ・____________________（プラン：__________）',
'',
'上記以外のツールを業務で利用する場合は、事前にAI管理責任者の承認を得ること。',
'個人のアカウントで業務データを扱うことは、いかなる場合も認めない。',
'',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'4. 入力してはならない情報',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'次の情報は、いかなるツールにも入力してはならない。',
'',
'  □ 顧客の氏名・住所・電話番号・メールアドレス',
'  □ 取引先名と金額が特定できる見積書・請求書の記載内容',
'  □ 未公開の経営情報（決算数値、事業計画、人事情報）',
'  □ 従業員の個人情報（給与、評価、健康状態）',
'  □ 秘密保持契約の対象となる情報',
'  □ 設計図面、製造ノウハウ、独自の技術情報',
'  □ ____________________（自社で追加）',
'',
'※ 病歴・信条・社会的身分などの要配慮個人情報は特に厳重に扱うこと。',
'',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'5. 条件付きで入力してよい情報',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'次の加工を行った場合に限り、入力を認める。',
'',
'  ・社名を伏せる（例：「A社」「当社の取引先」と置き換える）',
'  ・金額を丸める（例：「約500万円」）',
'  ・個人名を役割に置き換える（例：「担当者」「先方の窓口」）',
'',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'6. 出力の取扱い',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'  ・生成AIの出力は必ず人が確認してから使用する。',
'  ・事実関係、数値、法令の解釈は、一次情報で裏取りを行う。',
'  ・特定の作品や作風を指示して生成させた成果物を、社外向けの',
'    広告・販促に使用しない。',
'  ・出力をそのまま用いて第三者に損害を与えた場合、責任は当社が負う。',
'',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'7. 外部データの読み込み',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'外部から受け取った文書やWebページを生成AIに読み込ませる場合、',
'その内容に指示文が仕込まれている可能性（プロンプトインジェクション）',
'に留意し、出力に不審な挙動がないか確認する。',
'',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'8. 本人確認',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'振込先の変更、緊急の送金指示、機密情報の送付依頼を受けた場合は、',
'相手から伝えられた連絡先ではなく、当社が従来から把握している',
'連絡先に自社から掛け直して確認する。',
'声や映像は偽装されうることを前提とする。',
'',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'9. 管理体制',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'  AI管理責任者：____________________（役職・氏名）',
'  権限：ツールの承認、利用状況の確認、インシデント発生時の初動指揮',
'',
'  報告先（一次）：____________________',
'  報告先（二次）：____________________',
'',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'10. 問題発生時の対応',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'  ・情報漏洩の疑いが生じたとき',
'      → 直ちにAI管理責任者へ報告する。自己判断で解決しない。',
'      → 原因の特定を待たずに、まず報告する。',
'',
'  ・判断に迷うとき',
'      → 「迷ったら入力しない」を原則とし、AI管理責任者に相談する。',
'',
'  ・新しい使い方を試したいとき',
'      → AI管理責任者に事前相談してから実施する。',
'',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'11. 周知と点検',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'  ・本ガイドラインは、全対象者に配布し、周知を行う。',
'  ・年1回以上、運用状況を点検し、必要に応じて見直す。',
'  ・法令の改正があった場合は、速やかに内容を確認する。',
'',
'━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━',
'',
'本テンプレートは AI Security Lab（合同会社LaughRally）が提供するものです。',
'貴社の実情に合わせて編集のうえ、ご利用ください。',
'https://ai-security-lab.com',
''
  ].join('\n');

  const blob = new Blob(['\ufeff' + txt], { type: 'text/plain;charset=utf-8' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url;
  a.download = 'AI利用ガイドライン_テンプレート.txt';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ===================== 管理者：受講ログのエクスポート =====================
// #admin-dashboard は既定で display:none。管理者アカウントでログインしたときのみ表示する。
// 管理者はアカウントIDで判定する（Worker の ADMIN_USER_IDS と同じ）。表示を切り替えるだけで、権限はサーバーが判定する
const ADMIN_USER_IDS = ['0fc82602-6512-41b5-9ef8-6a14b533d7f7'];
function isAdminUser(){ return !!(currentUser && ADMIN_USER_IDS.includes(currentUser.id)); }

function applyAdminVisibility(){
  const sec = document.getElementById('admin-dashboard');
  if(!sec) return;
  sec.style.display = isAdminUser() ? 'block' : 'none';
  if(isAdminUser() && !window._mfaChecked){ window._mfaChecked = true; adminMfaChallengeIfNeeded().then(adminMfaRefresh); }
}
// ===== 管理者の二段階認証 =====
let _mfaEnrollId = null;
async function adminMfaRefresh(){
  const st = document.getElementById('admin-mfa-status');
  const startBtn = document.getElementById('admin-mfa-start');
  if(!st || !_supabase || !isAdminUser()) return;
  try {
    const { data } = await _supabase.auth.mfa.listFactors();
    const verified = (data && data.totp || []).filter(f => f.status === 'verified');
    const { data: aal } = await _supabase.auth.mfa.getAuthenticatorAssuranceLevel();
    if(verified.length){
      st.innerHTML = '✅ 登録済みです。管理者の操作は、ログイン時に認証アプリのコードを入れた場合だけ使えます。' +
        (aal && aal.currentLevel === 'aal2' ? '<br>（いまのログインはコード入力済みです）' : '<br>⚠️ いまのログインはコード未入力です。管理者の操作を使うには、ページを再読み込みしてコードを入れてください。');
      if(startBtn) startBtn.style.display = 'none';
    } else {
      st.textContent = '未登録です。登録すると、パスワードが漏れても管理者の操作（レッスン本文の投入など）は使えなくなります。';
      if(startBtn) startBtn.style.display = '';
    }
  } catch(e) {
    st.textContent = '状態を確認できませんでした：' + (e.message || '');
  }
}
async function adminMfaEnroll(){
  try {
    const { data, error } = await _supabase.auth.mfa.enroll({ factorType: 'totp', friendlyName: 'ASL admin ' + new Date().toISOString().slice(0,19) });
    if(error) throw error;
    _mfaEnrollId = data.id;
    document.getElementById('admin-mfa-qr').src = data.totp.qr_code;
    document.getElementById('admin-mfa-secret').textContent = data.totp.secret;
    document.getElementById('admin-mfa-enroll').style.display = '';
    document.getElementById('admin-mfa-start').style.display = 'none';
  } catch(e) { alert('登録を始められませんでした：' + (e.message || '')); }
}
async function adminMfaVerify(){
  const code = (document.getElementById('admin-mfa-code').value || '').trim();
  if(!/^\d{6}$/.test(code)){ alert('6桁の数字を入れてください'); return; }
  try {
    const { error } = await _supabase.auth.mfa.challengeAndVerify({ factorId: _mfaEnrollId, code });
    if(error) throw error;
    document.getElementById('admin-mfa-enroll').style.display = 'none';
    alert('登録しました。次回から、ログイン後に認証アプリのコードを入れてください。');
    adminMfaRefresh();
  } catch(e) { alert('コードを確認できませんでした：' + (e.message || '')); }
}
// 登録済みの管理者がコード未入力でログインしているときは、コードを求める
async function adminMfaChallengeIfNeeded(){
  if(!_supabase || !isAdminUser()) return;
  try {
    const { data: aal } = await _supabase.auth.mfa.getAuthenticatorAssuranceLevel();
    if(!aal || aal.nextLevel !== 'aal2' || aal.currentLevel === 'aal2') return;
    const { data } = await _supabase.auth.mfa.listFactors();
    const f = (data && data.totp || []).find(x => x.status === 'verified');
    if(!f) return;
    const code = prompt('管理者の操作には、認証アプリの6桁のコードが必要です。コードを入れてください（受講だけなら「キャンセル」で構いません）');
    if(!code) return;
    const { error } = await _supabase.auth.mfa.challengeAndVerify({ factorId: f.id, code: code.trim() });
    if(error) alert('コードを確認できませんでした：' + error.message);
    adminMfaRefresh();
  } catch(e) { console.warn('二段階認証の確認に失敗:', e.message); }
}

// ===== 法人の担当者：受講者の進捗 =====
let _corpApps = null;
async function loadCorpProgress(){
  try {
    const r = await apiFetch('/api/corporate/progress', { method: 'GET' });
    _corpApps = (r && r.applications) || [];
  } catch(e) { _corpApps = []; }
  if(_corpApps.length){
    try { buildProgressPanel(); } catch(e){}
    // 担当者専用のアカウント（受講プランなし）なら、ログイン直後に進捗を開く
    if((currentUserPlan || 'free') === 'free') openCorpProgress();
  }
}
function corpRows(){
  const rows = [];
  (_corpApps || []).forEach(a => (a.participants || []).forEach(p => rows.push({ a, p })));
  return rows;
}
function openCorpProgress(){
  const fmt = d => d ? new Date(d).toLocaleDateString('ja-JP') : '—';
  const hrs = s => s ? (Math.round(s / 360) / 10) + '時間' : '—';
  let html = '<div style="font-family:IBM Plex Mono,monospace;font-size:11px;color:var(--red);letter-spacing:.15em;margin-bottom:8px">// CORPORATE</div>' +
    '<h2 style="font-size:20px;font-weight:700;margin:0 0 6px">受講者の進捗</h2>' +
    '<p style="font-size:13px;color:var(--text3);margin:0 0 16px">お申し込みの担当者として登録されているメールアドレスでログインしたときに表示されます。回答の内容は表示しません。</p>';
  (_corpApps || []).forEach(a => {
    const statusJa = { invoiced: 'お支払い待ち', paid: 'アカウント発行中', provisioned: '受講中' }[a.status] || a.status;
    html += `<div style="margin:18px 0 8px;font-weight:700">${escHtml(a.company)}　<span style="font-weight:400;color:var(--text3);font-size:13px">${escHtml(a.appNo)} ／ ${escHtml(a.plan)} ${a.seats}名 ／ ${statusJa}</span></div>` +
      '<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:13px;min-width:640px">' +
      '<tr style="text-align:left;color:var(--text3)"><th style="padding:6px">氏名</th><th>レッスン</th><th>章末ワーク</th><th>確認テスト</th><th>試験</th><th>学習時間</th><th>最終受講</th><th>修了証</th></tr>';
    (a.participants || []).forEach(p => {
      const td = v => `<td style="padding:6px;border-top:1px solid var(--border)">${v}</td>`;
      // 申込より前からご本人が持っていたアカウントは、Worker が進捗を返さない
      if(p.existingAccount){ html += '<tr>' + td(escHtml(p.name || p.email)) + `<td colspan="7" style="padding:6px;border-top:1px solid var(--border);color:var(--text3)">お申込み前からのアカウントのため、進捗は表示しません（表示にはご本人の同意が必要です。運営までご連絡ください）</td></tr>`; return; }
      if(!p.registered){ html += '<tr>' + td(escHtml(p.name || p.email)) + `<td colspan="7" style="padding:6px;border-top:1px solid var(--border);color:var(--text3)">アカウント未発行（お支払い確認後に発行されます）</td></tr>`; return; }
      html += '<tr>' + td(escHtml(p.name || p.email)) + td(`${p.lessonsDone} / ${a.lessonsTotal}`) +
        td(a.chapters ? `${p.worksDone} / ${a.chapters}章` : '—') + td(a.chapters ? `${p.testsPassed} / ${a.chapters}章` : '—') +
        td(a.plan === 'premium' ? (p.examPassed ? '合格' : '未') : '—') + td(hrs(p.totalSeconds)) + td(fmt(p.lastActivity)) +
        td(p.certificateIssuedAt ? '発行済（' + fmt(p.certificateIssuedAt) + '）' : '—') + '</tr>';
    });
    html += '</table></div>';
  });
  html += '<div style="margin-top:18px;display:flex;gap:10px;flex-wrap:wrap"><button onclick="downloadCorpCsv()" style="background:var(--green-s);color:#fff;border:none;border-radius:8px;padding:11px 18px;font-weight:600;cursor:pointer;min-height:44px">📥 CSVをダウンロード</button>' +
    '<button onclick="closeCorpProgress()" style="background:var(--bg2);color:var(--text2);border:1px solid var(--border2);border-radius:8px;padding:11px 18px;cursor:pointer;min-height:44px">閉じる</button></div>';
  let ov = document.getElementById('corp-progress-overlay');
  if(!ov){
    ov = document.createElement('div');
    ov.id = 'corp-progress-overlay';
    ov.setAttribute('role', 'dialog');
    ov.setAttribute('aria-modal', 'true');
    ov.style.cssText = 'position:fixed;inset:0;z-index:600;background:rgba(15,23,42,.55);display:flex;align-items:flex-start;justify-content:center;padding:70px 16px 24px;overflow-y:auto';
    ov.addEventListener('click', e => { if(e.target === ov) closeCorpProgress(); });
    document.body.appendChild(ov);
  }
  ov.innerHTML = '<div style="background:var(--bg);color:var(--text);border-radius:14px;max-width:980px;width:100%;padding:24px">' + html + '</div>';
  ov.style.display = 'flex';
}
function closeCorpProgress(){ const ov = document.getElementById('corp-progress-overlay'); if(ov) ov.style.display = 'none'; }
function downloadCorpCsv(){
  const q = csvCell;
  const lines = [['会社名','申込番号','プラン','氏名','メールアドレス','レッスン完了','レッスン総数','章末ワーク提出章','確認テスト合格章','章数','認定試験','学習時間(秒)','最終受講','修了証発行日'].map(q).join(',')];
  corpRows().forEach(({a, p}) => lines.push([a.company, a.appNo, a.plan, p.name, p.email, p.registered && !p.existingAccount ? p.lessonsDone : '', a.lessonsTotal,
    p.registered ? p.worksDone : '', p.registered ? p.testsPassed : '', a.chapters, a.plan === 'premium' ? (p.examPassed ? '合格' : '未') : '',
    p.registered ? p.totalSeconds : '', p.lastActivity || '', p.certificateIssuedAt || ''].map(q).join(',')));
  const blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'ASL_受講者進捗_' + new Date().toISOString().slice(0,10) + '.csv';
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

window.applyAdminVisibility = applyAdminVisibility;

async function adminExportLogs(){
  const btn = document.getElementById('admin-export-btn');
  const out = document.getElementById('admin-export-result');
  if(!isAdminUser()){
    if(out) out.textContent = 'この操作を行う権限がありません。';
    return;
  }
  if(btn){ btn.disabled = true; btn.textContent = '取得中...'; }
  if(out) out.textContent = '';
  try{
    const { data } = await _supabase.auth.getSession();
    const token = data?.session?.access_token;
    if(!token) throw new Error('ログインが必要です');
    const r = await fetch(`${WORK_API}/api/admin/logs`, { headers:{ Authorization:'Bearer '+token }});
    if(!r.ok){
      const e = await r.json().catch(()=>({}));
      throw new Error(e.error || 'サーバーから取得できませんでした');
    }
    const d = await r.json();
    const rows = d.rows || [];
    if(!rows.length){ if(out) out.textContent = '対象のログがありません。'; return; }
    // CSVの数式インジェクション対策（受講実績レポート側と同じ扱い）
    const q = csvCell;
    const lines = [['user_id','lesson_id','completed_at','plan'].map(q).join(',')];
    rows.forEach(x => lines.push([x.user_id, x.lesson_id, x.completed_at, x.plan].map(q).join(',')));
    const blob = new Blob(['\ufeff' + lines.join('\r\n')], { type:'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `受講ログ_${new Date().toISOString().slice(0,10)}.csv`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(()=>URL.revokeObjectURL(url), 1000);
    if(out) out.textContent = `${rows.length} 件をダウンロードしました。`;
  }catch(e){
    if(out) out.textContent = 'ダウンロードに失敗しました：' + (e.message || '');
  }finally{
    if(btn){ btn.disabled = false; btn.textContent = '📥 受講ログCSVをダウンロード'; }
  }
}



// ===================== 章末ワーク =====================

const WORK_QUESTIONS = {
  'w1-1': { min:150, ref:'L03', q:'あなたの部署または会社で、業務に生成AIツールが使われている場面をすべて挙げてください。会社が公式に導入したものだけでなく、個人のアカウントで使われているもの、無料版を使っているものも含めて書き出してください。「誰が・どのツールを・どんな業務で」使っているかがわかる形で記述してください。心当たりがない場合は、そう判断した根拠（ヒアリングの実施、ログの確認など）を書いてください。' },
  'w1-2': { min:150, ref:'L01', q:'w1-1で挙げたツールに、これまで実際に入力された情報の種類を書き出してください。顧客名・取引先名・金額・個人情報・社内資料の文面など、具体的に挙げてください。そのうえで、それらの情報が現在どこに保存されている可能性があるかを、L01の内容をふまえて記述してください。' },
  'w1-3': { min:150, ref:'L02', q:'w1-1で挙げたツールのうち、主要なもの1つ以上について、実際に管理画面を開いて学習オフ設定の状態を確認してください。確認したツール名、設定の場所、現在の状態（オン／オフ）を記述してください。そのうえで、学習オフにしても防げないリスクが何かを、L02の内容をふまえて書いてください。' },
  'w1-4': { min:200, ref:'L04・総括', q:'ここまでの3問をふまえて、自社のAI利用実態を整理してください。最もリスクが高いと考えられる箇所を1つ挙げ、その理由と、当面どう対処するかを記述してください。対処が「現状維持」である場合も、なぜそう判断したかを書いてください。' },

  'w2-1': { min:150, ref:'L07', q:'自社で扱う情報を「入力してよい情報」「条件付きで入力してよい情報」「絶対に入力してはいけない情報」の3つに分類してください。それぞれ最低3つずつ、自社の業務で実際に扱っているものを具体的に挙げてください。「顧客情報」のような抽象的な括りではなく、「見積書に記載された取引先名と金額」のように、判断できる粒度で書いてください。' },
  'w2-2': { min:150, ref:'L07', q:'w2-1で「条件付き」に分類した情報について、その条件を具体的に定めてください。どのツールなら許可するのか、どう加工すれば入力してよいのか（社名を伏せる、数値を丸めるなど）、誰の承認が必要かを記述してください。条件付きに分類したものがない場合は、なぜ2分類で足りると判断したかを書いてください。' },
  'w2-3': { min:150, ref:'L05・L06', q:'現在自社で使っている、または導入を検討している生成AIツールについて、L05・L06で学んだセキュリティ設定の観点から評価してください。プランの違い（無料版か有料版か、法人向けプランか）、学習に使われるかどうか、管理者による統制が可能かを整理し、現状のままでよいか、変更すべきかを記述してください。' },
  'w2-4': { min:200, ref:'L08・総括', q:'L08の業種別リスク事例のうち、自社の業種または業務に最も近いものを挙げ、自社で同じことが起きる可能性を検討してください。起こりうると考える場合はその経路を、起こりにくいと考える場合はその根拠を記述してください。そのうえで、w2-1で作成した入力禁止情報リストを社内で共有する方法（朝礼、書面配布、システム上の掲示など）を具体的に書いてください。' },

  'w3-1': { min:150, ref:'L09', q:'自社が業務で扱う個人情報のうち、生成AIに入力される可能性があるものを挙げてください。氏名・連絡先だけでなく、健康情報や信条など要配慮個人情報にあたるものがあれば区別して記述してください。そのうえで、L09の内容をふまえ、それらをAIに入力することが個人情報保護法上どう位置づけられるかを自社の状況に即して書いてください。' },
  'w3-2': { min:150, ref:'L10', q:'自社の業務で、生成AIが出力した文章・画像・コードをそのまま使っている、または使う可能性がある場面を挙げてください。そのうえで、L10で学んだ「著作権侵害になるケース・ならないケース」に照らし、自社の使い方がどちらに近いか、リスクを下げるために何を変えるべきかを記述してください。該当する場面がない場合は、今後発生しうる場面を想定して書いてください。' },
  'w3-3': { min:150, ref:'L11', q:'生成AIの誤った出力（ハルシネーション）が自社の業務で実害につながりうる場面を、具体的に1つ以上挙げてください。誰に対してどのような損害が生じるか、その場合に責任を負うのは誰かを、L11の内容をふまえて記述してください。' },
  'w3-4': { min:200, ref:'総括', q:'w3-1〜w3-3で挙げたリスクを、「発生可能性」と「発生した場合の影響の大きさ」の2軸で評価し、優先的に対処すべきものを1つ選んでください。その選定理由と、具体的な対処方針（チェック体制、利用制限、記録の残し方など）を記述してください。あわせて、自社に法務担当や顧問弁護士がいる場合は相談すべき論点を、いない場合はどこに相談するかを書いてください。' },

  'w4-1': { min:150, ref:'L12', q:'L12のテンプレートをもとに、自社のAI利用ガイドラインの「適用範囲」と「禁止事項」を作成してください。適用範囲は誰に適用されるか（正社員のみか、パート・派遣・業務委託を含むか）を明記してください。禁止事項は第2章で作成した入力禁止情報リストを反映させ、実際に運用できる粒度で記述してください。' },
  'w4-2': { min:150, ref:'L13', q:'自社でAIによる情報漏洩が疑われる事態が起きた場合の初動対応を、L13の内容をふまえて記述してください。誰が最初に報告を受けるか、その人が最初に行うこと、外部（顧客・取引先・監督官庁）への連絡が必要になる判断基準を、自社の実際の体制に即して具体的に書いてください。' },
  'w4-3': { min:150, ref:'L14', q:'L14の内容をふまえ、自社でAI利用を管理する体制について記述してください。専任者を置けない場合、誰が兼任するのが現実的か、その人にどこまでの権限を持たせるかを書いてください。あわせて、ガイドラインを従業員に周知し、遵守を確認する方法（研修、誓約書、定期チェックなど）を具体的に記述してください。' },
  'w4-4': { min:200, ref:'L15・総括', q:'L15の「今日からできる10のアクション」のうち、自社で未実施のものを挙げ、優先順位をつけてください。上位3つについて、いつまでに、誰が、何をするかを記述してください。すでにすべて実施済みの場合は、それぞれが形骸化していないかを点検した結果と、次に取り組むべきことを書いてください。' },

  'w5-1': { min:150, ref:'L16', q:'L16の20事例から、自社の業務に最も近いものを1つ選び、なぜそれを選んだかを記述してください。そのうえで、その事例で問題となった点が自社でも起こりうるか、自社の場合はどこが違うかを具体的に書いてください。' },
  'w5-2': { min:150, ref:'L16', q:'自社の業務のうち、生成AIを使うことで最も効果が見込める作業を1つ挙げてください。現在その作業に月あたりどれくらいの時間がかかっているか、AIを使うことでどの程度短縮できると見込むかを記述してください。数値が把握できていない場合は、概算とその根拠を書いてください。' },
  'w5-3': { min:150, ref:'L17', q:'w5-2で挙げた業務に使うツールを、L17の比較をふまえて選定してください。選んだツール名、プラン、想定する月額費用、選定理由を記述してください。あわせて、そのツールを使ううえで第2章で作成した入力禁止情報リストとの整合が取れているかを確認し、追加で必要な設定や運用ルールがあれば書いてください。' },
  'w5-4': { min:200, ref:'総括', q:'w5-2・w5-3をふまえ、自社のAI導入計画をまとめてください。導入する業務、使用するツール、対象となる従業員、開始時期、効果を測る指標を記述してください。導入しないという結論であっても構いませんが、その場合はなぜ現時点で導入しないのか、どういう条件が整えば導入するのかを書いてください。' },

  'w6-1': { min:150, ref:'L23', q:'自社に届くメールや連絡のうち、AIを使ったフィッシングの標的になりやすいものを挙げてください。取引先を装った請求書、社内の上長を装った指示など、自社で実際にありうる形を具体的に書いてください。そのうえで、L23で学んだ見分け方のうち、自社の従業員が実行できるものを記述してください。' },
  'w6-2': { min:150, ref:'L24', q:'自社で生成AIに外部の文書やWebページを読み込ませる業務があるか、あるいは今後発生しうるかを検討してください。ある場合、L24のプロンプトインジェクションが成立しうる経路を具体的に記述してください。ない場合は、なぜ発生しないと判断したかを書いてください。' },
  'w6-3': { min:150, ref:'L25', q:'自社で「本人確認」が必要になる場面を挙げてください。電話での振込指示、役員からの緊急依頼、取引先からの口座変更連絡などが該当します。L25をふまえ、声や映像が偽装されうる前提で、確実に本人を確認する手段を自社の実情に即して記述してください。' },
  'w6-4': { min:200, ref:'L27・総括', q:'L27のロールプレイをふまえ、自社で情報漏洩が発生した場合の対応フローを作成してください。発見から報告、被害範囲の確認、関係者への連絡、再発防止までの流れを時系列で記述してください。第4章のw4-2で書いた初動対応と矛盾がないか確認し、実際に動かせる内容になっているかを点検したうえで書いてください。' },

  'w7-1': { min:150, ref:'L31', q:'自社の業務で生成AIに指示を出す際、現在どのような書き方をしているかを具体的に記述してください。実際に使っている指示文があれば、そのまま書き出してください。そのうえで、L31の内容をふまえ、より安全で精度の高い指示に書き換えるとどうなるかを記述してください。' },
  'w7-2': { min:150, ref:'L32', q:'自社のデータ（社内マニュアル、過去の見積書、顧客対応履歴など）をAIに読み込ませて活用する構想があるかを検討してください。ある場合、L32をふまえてRAGとファインチューニングのどちらが適切か、その理由と想定されるリスクを記述してください。ない場合は、どういうデータであれば読み込ませてよいと考えるかを書いてください。' },
  'w7-3': { min:150, ref:'L34', q:'L34をふまえ、自社でAIエージェントや自動化ツールを使う場合に許容できる範囲を検討してください。AIに自動で実行させてよい操作と、必ず人間の確認を挟むべき操作を区別して記述してください。現時点で使う予定がない場合も、将来使うとしたらどう線を引くかを書いてください。' },
  'w7-4': { min:200, ref:'L33・総括', q:'L33をふまえ、自社のAI活用ロードマップを作成してください。今後1年間で、どの業務にどの順序でAIを導入していくか、各段階で必要になるセキュリティ対策とあわせて記述してください。第5章のw5-4で作成した導入計画との整合を確認し、より長期の視点で見直した結果を書いてください。' }
};

const WORK_API = WORKER_URL;   // 同じ Worker（以前は URL を2か所に書いていた）
const _workTimers = {};
const _workStarted = {};
let _workLoaded = {};
const _workSubmittedCh = {};   // { 章番号: true } サイドバーの完了表示に使う
const _workCurQ = {};          // { 章番号: 表示中の設問番号 }
const _workReopen = {};        // { 章番号: true } 書き直しモード

async function workToken(){
  const { data } = await _supabase.auth.getSession();
  return data?.session?.access_token || null;
}

function workCount(t){ return (t||'').trim().length; }

function renderWorkChapter(ch){
  const info = WORK_CHAPTERS[ch];
  const host = document.getElementById('work-body-'+ch);
  if(!host) return;
  let steps = '<div class="work-steps" id="work-steps-'+ch+'">';
  for(let i=1;i<=COURSE.rules.workQuestions;i++) steps += `<div class="work-step" id="wst-${ch}-${i}"></div>`;
  steps += '</div>';
  let html = steps;
  for(let i=1;i<=COURSE.rules.workQuestions;i++){
    const qid = `w${ch}-${i}`;
    const d = WORK_QUESTIONS[qid];
    html += `
      <div class="work-q${i===1?' is-active':''}" id="wq-${qid}">
        <div class="work-q-head">
          <span class="work-q-no">設問 ${i} / ${COURSE.rules.workQuestions}</span>
          <span class="work-q-ref">参照：${d.ref}</span>
        </div>
        <p class="work-q-text">${d.q}</p>
        <textarea class="work-input" id="wa-${qid}" data-qid="${qid}"
          placeholder="ここに記入してください（${d.min}文字以上）"
          oninput="onWorkInput('${qid}')"></textarea>
        <div class="work-meta">
          <span class="work-count" id="wc-${qid}">0 / ${d.min} 文字</span>
          <span class="work-saved" id="ws-${qid}"></span>
        </div>
        <div class="work-nav">
          <button class="work-nav-btn" onclick="gotoWorkQ(${ch},${i-1})" ${i===1?'disabled':''}>← 前の設問</button>
          <button class="work-nav-btn" onclick="gotoWorkQ(${ch},${i+1})" ${i===COURSE.rules.workQuestions?'disabled':''}>次の設問 →</button>
        </div>
      </div>`;
  }
  host.innerHTML = html;
  _workCurQ[ch] = 1;
}

function gotoWorkQ(ch, n){
  if(n<1 || n>COURSE.rules.workQuestions) return;
  for(let i=1;i<=COURSE.rules.workQuestions;i++){
    const el = document.getElementById(`wq-w${ch}-${i}`);
    if(el) el.classList.toggle('is-active', i===n);
  }
  _workCurQ[ch] = n;
  updateWorkSteps(ch);
  const host = document.getElementById('work-body-'+ch);
  if(host) host.scrollIntoView({behavior:'smooth', block:'start'});
  const ta = document.getElementById(`wa-w${ch}-${n}`);
  if(ta && !ta.readOnly) setTimeout(()=>ta.focus(), 260);
}

function updateWorkSteps(ch){
  for(let i=1;i<=COURSE.rules.workQuestions;i++){
    const st = document.getElementById(`wst-${ch}-${i}`);
    if(!st) continue;
    const qid = `w${ch}-${i}`;
    const el = document.getElementById('wa-'+qid);
    const ok = el && workCount(el.value) >= WORK_QUESTIONS[qid].min;
    st.classList.toggle('is-ok', !!ok);
    st.classList.toggle('is-cur', !ok && _workCurQ[ch]===i);
  }
}

async function loadWorkChapter(ch){
  const host = document.getElementById('work-body-'+ch);
  const status = document.getElementById('work-status-'+ch);
  if(!host) return;
  if(!_workLoaded[ch]) renderWorkChapter(ch);
  const t = await workToken();
  if(!t){ status.textContent = 'ログインすると記入できます。'; return; }
  try{
    const r = await fetch(`${WORK_API}/api/work/${ch}`, { headers:{ Authorization:'Bearer '+t }});
    const d = await r.json();
    if(!r.ok){ status.textContent = d.error || '読み込みに失敗しました'; return; }
    (d.answers||[]).forEach(a=>{
      const el = document.getElementById('wa-'+a.question_id);
      // 保存前の入力がある欄は上書きしない。入力後1.5秒以内にタブを切り替えて戻ると、古い値で上書きされ、その古い値が保存されて入力が消えていた
      if(el && el.dataset.dirty !== '1'){ el.value = a.answer_text || ''; }
    });
    for(let i=1;i<=COURSE.rules.workQuestions;i++) updateWorkCount(`w${ch}-${i}`);
    _workLoaded[ch] = true;
    if(d.reopened){ _workReopen[ch] = true; }
    setWorkSubmitted(ch, !!d.submitted, d.answers||[]);
    if(d.reopened && status) status.innerHTML = '<span class="work-note">書き直しの途中です。修正後、あらためて提出してください（提出するまで、この章は未提出の扱いです）。</span>';
  }catch(e){
    status.textContent = '読み込みに失敗しました。通信環境をご確認ください。';
  }
}

function setWorkSubmitted(ch, submitted, answers){
  _workSubmittedCh[ch] = !!submitted;
  const status = document.getElementById('work-status-'+ch);
  const btn = document.getElementById('work-submit-'+ch);
  document.querySelectorAll(`#work-body-${ch} .work-input`).forEach(el=>{
    el.readOnly = submitted;
    el.classList.toggle('is-locked', submitted);
  });
  if(submitted){
    const at = (answers.find(a=>a.submitted_at)||{}).submitted_at;
    const d = at ? new Date(at).toLocaleString('ja-JP') : '';
    status.innerHTML = `<span class="work-done">✓ 提出済み</span>${d?`　${d}`:''}<br><span class="work-note">内容は受講実績レポートに記録されます。書き直した場合は最新の内容と提出日時が記録されます。</span><br><button class="work-reopen" onclick="reopenWork(${ch})">書き直す</button>`;
    if(btn){ btn.disabled = true; btn.textContent = '✓ 提出済み'; btn.classList.add('is-done'); }
  }else{
    status.textContent = '';
    if(btn){ btn.disabled = false; btn.classList.remove('is-done'); }
    updateWorkSubmitBtn(ch);
  }
}

function updateWorkCount(qid){
  const el = document.getElementById('wa-'+qid);
  const out = document.getElementById('wc-'+qid);
  if(!el||!out) return false;
  const n = workCount(el.value);
  const min = WORK_QUESTIONS[qid].min;
  out.textContent = `${n} / ${min} 文字`;
  const ok = n >= min;
  out.classList.toggle('is-ok', ok);
  return ok;
}

function updateWorkSubmitBtn(ch){
  const btn = document.getElementById('work-submit-'+ch);
  let ok = 0;
  for(let i=1;i<=COURSE.rules.workQuestions;i++) if(updateWorkCount(`w${ch}-${i}`)) ok++;
  updateWorkSteps(ch);
  if(!btn || btn.classList.contains('is-done')) return;   // 提出済み
  btn.disabled = ok < COURSE.rules.workQuestions;
  const need = COURSE.rules.workQuestions;
  btn.textContent = ok < need ? `提出する（${ok}/${need} 問が必要文字数に到達）` : '提出する';
}

function onWorkInput(qid){
  const ch = parseInt(qid.split('-')[0].slice(1),10);
  const _el = document.getElementById('wa-'+qid);
  if(_el) _el.dataset.dirty = '1';
  updateWorkSubmitBtn(ch);
  if(!_workStarted[qid]) _workStarted[qid] = new Date().toISOString();
  clearTimeout(_workTimers[qid]);
  const s = document.getElementById('ws-'+qid);
  if(s) s.textContent = '';
  _workTimers[qid] = setTimeout(()=>saveWorkDraft(qid), 1500);
}

// 下書きを保存する。保存できたか（または保存の必要が無いか）を true／false で返す
async function saveWorkDraft(qid){
  const el = document.getElementById('wa-'+qid);
  const s  = document.getElementById('ws-'+qid);
  if(!el || el.readOnly) return true;
  const sentValue = el.value;
  const t = await workToken();
  if(!t) return false;
  try{
    const r = await fetch(`${WORK_API}/api/work/draft`, {
      method:'POST',
      headers:{ Authorization:'Bearer '+t, 'Content-Type':'application/json' },
      body: JSON.stringify({ questionId:qid, answerText:sentValue, startedAt:_workStarted[qid], reopen: !!_workReopen[parseInt(qid.split('-')[0].slice(1),10)] })
    });
    const d = await r.json();
    // 送った内容のまま変わっていなければ「保存済み」に戻す（送信中にさらに入力していたら未保存のまま）
    if(r.ok && el.value === sentValue) el.dataset.dirty = '';
    if(s) s.textContent = r.ok ? '保存しました' : (d.error||'保存に失敗しました');
    if(s && r.ok) setTimeout(()=>{ if(s.textContent==='保存しました') s.textContent=''; }, 2500);
    return r.ok;
  }catch(e){
    if(s) s.textContent = '保存に失敗しました';
    return false;
  }
}

async function submitWorkChapter(ch){
  const btn = document.getElementById('work-submit-'+ch);
  const status = document.getElementById('work-status-'+ch);
  if(!confirm(`第${ch}章の章末ワークを提出します。\n提出後は編集できません。よろしいですか？`)) return;
  btn.disabled = true;
  btn.textContent = '提出中...';
  // 未保存の下書きを確定させる。1件でも保存できなければ提出しない（画面と違う古い文が提出済みになるため）
  for(let i=1;i<=COURSE.rules.workQuestions;i++){
    const qid = `w${ch}-${i}`;
    clearTimeout(_workTimers[qid]);
    if(!(await saveWorkDraft(qid))){
      btn.disabled = false; btn.textContent = '提出する';
      if(status) status.textContent = '下書きを保存できなかったため、提出していません。通信環境をご確認のうえ、もう一度「提出する」を押してください。';
      return;
    }
  }
  const t = await workToken();
  try{
    // 章末ワークに費やした実時間を送る。Worker 側はすでに durationSeconds を
    // 受け取れる実装になっており（chapter_works.duration_seconds）、
    // これまではフロントが送っていなかったため常に NULL だった。
    flushActive();
    const workSecs = loadActive()['work' + ch] || null;
    const r = await fetch(`${WORK_API}/api/work/submit`, {
      method:'POST',
      headers:{ Authorization:'Bearer '+t, 'Content-Type':'application/json' },
      body: JSON.stringify({ chapter: ch, durationSeconds: workSecs })
    });
    const d = await r.json();
    if(!r.ok){
      status.textContent = d.error || '提出に失敗しました';
      if(d.shortfall){
        status.textContent += '（' + d.shortfall.map(x=>`${x.questionId}: ${x.length}/${x.minLength}字`).join('、') + '）';
      }
      btn.disabled = false;
      updateWorkSubmitBtn(ch);
      return;
    }
    _workReopen[ch] = false;
    setWorkSubmitted(ch, true, [{submitted_at:d.submittedAt}]);
    try{ buildSideNav(loadCompleted()); }catch(e2){}
    refreshRecords();
  }catch(e){
    status.textContent = '提出に失敗しました。通信環境をご確認ください。';
    btn.disabled = false;
    updateWorkSubmitBtn(ch);
  }
}

window.loadWorkChapter   = loadWorkChapter;
// 提出済みの章を編集可能に戻す（記録は保持され、再提出で上書きされる）
async function reopenWork(ch){
  if(!confirm(`第${ch}章の章末ワークを書き直します。\n再提出するまで提出済みの扱いは解除されます。よろしいですか？`)) return;
  _workReopen[ch] = true;
  document.querySelectorAll(`#work-body-${ch} .work-input`).forEach(el=>{
    el.readOnly = false;
    el.classList.remove('is-locked');
  });
  _workSubmittedCh[ch] = false;
  const status = document.getElementById('work-status-'+ch);
  if(status) status.innerHTML = '<span class="work-note">書き直しモードです。修正後、あらためて提出してください。</span>';
  const btn = document.getElementById('work-submit-'+ch);
  if(btn){ btn.textContent = '提出する'; btn.disabled = false; btn.classList.remove('is-done'); }
  updateWorkSubmitBtn(ch);
  try{ buildSideNav(loadCompleted()); }catch(e){}
}
// ===================== 章末確認テスト =====================
const TEST_QUESTIONS = {
  't1-1': { q:'社内で「ChatGPTは有料版なら学習に使われないので安全だ」という説明がなされていました。この説明の問題点はどれですか。', c:['プランに関わらず入力内容はすべて学習に使われることになるため','学習利用の有無はプランではなく国ごとの法規制によって決まるため','学習利用の有無は個人向けか法人向けかで決まり、別の論点もあるから','有料版であっても入力内容は他の利用者から閲覧できてしまうから'] },
  't1-2': { q:'シャドーAIが組織にとって危険とされる最大の理由はどれですか。', c:['組織が把握できていないため、漏洩が起きても発覚せず対策も打てないから','無料版のAIは有料版より性能が低く、業務の品質が下がってしまうから','個人アカウントでのAI利用は法律で明確に禁止されているものだから','複数のAIツールを併用するとAI同士が情報を共有してしまうから'] },
  't1-3': { q:'自社の顧客名簿を扱う業務でAIを使いたいと相談されました。第1章をふまえた対応はどれですか。', c:['どのAIでも入力内容は暗号化されるため、そのまま使ってよいと伝える','AIの利用は一律に禁止し、従来どおり手作業で行うよう指示を出す','学習オフにすれば問題ないため、その設定方法だけを案内しておく','情報の種類と契約形態を確認し、入力してよい範囲を先に決めておく'] },
  't1-4': { q:'オンプレミス型AIについての説明として正しいものはどれですか。', c:['クラウド型と比べて、どのような規模でも必ず安価に導入できる','データを社外に出さずに済む一方、導入と運用の負担は大きい','インターネット接続が不要なため、モデルの更新も一切必要ない','中小企業向けの設計であり、専門知識がなくても運用ができる'] },
  't1-5': { q:'「入力した情報がどこへ行くのか」を考えるうえで、最初に確認すべきことはどれですか。', c:['そのAIの回答精度が業務で使える水準に達しているかどうか','開発元が日本の企業であり、国内法の適用を受けるかどうか','利用しているのがどの事業者のどの契約形態のサービスか','社内で何人がそのAIを日常的に利用しているかどうか'] },

  't2-1': { q:'社内で使うAIツールを選定する際、セキュリティの観点で最も重視すべき点はどれですか。', c:['回答が生成されるまでの速度が業務に支障しない水準かどうか','管理者が利用状況を把握し、設定を統制できるかどうか','日本語の表現が自然で、そのまま使える文章が出てくるか','対応しているファイル形式が業務の必要を満たすか'] },
  't2-2': { q:'「社外秘」と「個人情報」の扱いについて、正しい理解はどれですか。', c:['どちらも同種の情報であり、社内では同じ基準で扱えばよい','個人情報だけが法規制の対象であり、社外秘は自由に扱える','社外秘だけが企業の責任範囲で、個人情報は本人の責任となる','個人情報は法規制、社外秘は契約や営業秘密の観点から別に判断する'] },
  't2-3': { q:'生成AIに営業秘密を入力した場合のリスクとして、第2章で扱った論点はどれですか。', c:['秘密管理性が否定され、営業秘密として保護されなくなりうる','入力した時点で自動的に特許出願がなされたものとみなされてしまう','AI事業者の側に、その営業秘密の所有権が移転してしまうため','入力した内容が検索エンジンを経由して一般に公開されるため'] },
  't2-4': { q:'入力禁止情報のリストを作る際、実務上最も重要なことはどれですか。', c:['できるだけ多くの項目を列挙し、迷ったらすべて禁止にする','経営者だけで決定し、従業員には結果だけを通知しておく','現場が個別の場面で判断できる具体的な粒度で書くこと','法律の条文をそのまま引用し、根拠を明確にしておくこと'] },
  't2-5': { q:'業種別のリスク事例を自社に当てはめて検討する際、適切な姿勢はどれですか。', c:['自社と業種が異なる事例は前提が違うため参考にしない','事例で紹介された対策をそのまま自社にも導入しておく','過去に事故が起きていない業種はリスクが低いものと考える','問題となった構造が自社にも当てはまるかを業種を超えて確認する'] },

  't3-1': { q:'2026年7月に公布された改正個人情報保護法について、正しい説明はどれですか。', c:['課徴金制度が導入され、施行は公布から2年以内に政令で定められる','公布と同時に全面施行されており、すでに適用が開始されている','生成AIへの個人情報の入力を全面的に禁止する内容である','従業員数が一定以下の中小企業は適用除外とされている'] },
  't3-2': { q:'生成AIの出力をそのまま業務に使う場合、著作権の観点で最もリスクが高いのはどれですか。', c:['AIに生成させた文章を、社内向けのメモとして共有する','自社で作成した資料をAIに要約させて社内で回覧する','特定の作品や作風を指示して作らせた成果物を広告に使う','自分で書いた文章の誤字脱字をAIにまとめて修正させて提出する'] },
  't3-3': { q:'ハルシネーションによる損害が発生した場合の責任について、正しい理解はどれですか。', c:['AI事業者が全面的に責任を負うこととなり、利用した企業に責任は生じない','出力をそのまま用いて損害を与えた場合、利用した側が責任を問われる','AIが生成したものである旨を明示すれば、責任は一切生じない','無料版のAIを使った場合に限って、利用者が責任を負う'] },
  't3-4': { q:'要配慮個人情報について正しいものはどれですか。', c:['氏名と連絡先を組み合わせた情報がこれに該当することとされる','企業が保有する営業秘密を含んだ情報のことである','本人が自ら公開している情報であれば該当しない','病歴や信条などが該当し、より厳格な取扱いが必要となる'] },
  't3-5': { q:'自社の業務で法的リスクを評価する際、優先順位のつけ方として適切なものはどれですか。', c:['発生する可能性と、起きた場合の影響の大きさの両面から見る','発生する可能性が最も高いものから順に対処を進めていくこと','起きた場合の影響が最も大きいものだけに絞って対処する','法律の条文に明記されているものだけを評価の対象とする'] },

  't4-1': { q:'社内AI利用ガイドラインの適用範囲を定める際、適切なものはどれですか。', c:['正社員のみを対象とし、それ以外は各自の判断に委ねておく','役員は経営判断を担うため対象から外し、従業員のみに適用する','管理職だけに適用し、一般の従業員には周知のみ行っておく','雇用形態を問わず、業務でAIを使う者すべてを対象とする'] },
  't4-2': { q:'AIによる情報漏洩が疑われる事態が起きた場合、初動として最も優先すべきことはどれですか。', c:['原因を特定し、事実関係を固めたうえで報告を上げていく','事実を確認しつつ、定められた報告先へ速やかに報告する','社外への影響がないと確認できるまでは社内にも伏せておく','再発防止策を立てたうえで、経営層へまとめて報告を行う'] },
  't4-3': { q:'専任のAI管理担当者を置けない中小企業における現実的な体制はどれですか。', c:['体制が組めない以上は、AIの業務利用そのものを全面的に禁止する','担当を定めず、全従業員が各自の判断で運用していく','既存の情報管理担当が兼任し、権限と判断基準を明文化する','導入したAIベンダーに、すべての判断と管理を任せておく'] },
  't4-4': { q:'AI推進法（2025年6月公布・同年9月全面施行）について正しいものはどれですか。', c:['罰則のない基本法だが、AIを活用する事業者も対象に含まれる','AIを開発する事業者のみが対象であり、活用する側は含まれない','違反した事業者に対しては、課徴金が課される仕組みになっている','生成AIの業務利用を、国への事前の許可制とする内容になっている'] },
  't4-5': { q:'ガイドラインを策定した後、実効性を確保するために最も重要なことはどれですか。', c:['文書として完成させたうえで、共有フォルダに保存し公開しておく','周知と遵守確認の仕組みを設け、運用状況を定期的に点検する','想定される事態をすべて網羅した、可能な限り詳細な文書にする','外部の専門家に監修を依頼し、記載内容の正確性を担保しておく'] },

  't5-1': { q:'AI導入の効果を測る指標として、最も適切なものはどれですか。', c:['社内で導入したAIツールの種類と契約している総件数','AIに生成させた文章や資料の累計の分量と件数','対象とした業務にかかっていた時間がどう変わったか','社内でAIを使ったことのある従業員の割合とその推移'] },
  't5-2': { q:'AI導入を検討する際の順序として適切なものはどれですか。', c:['対象業務と扱う情報を特定し、それに合うツールを選ぶ','話題になっているツールを導入し、使える業務を探していく','最も費用の安いツールを選び、後から業務のほうを合わせる','全社員に自由に選ばせたうえで、後からひとつに統一する'] },
  't5-3': { q:'中小企業がAIツールを選定する際、コスト以外に必ず確認すべき点はどれですか。', c:['提供元の企業が設立から十分な年数を経過しているかどうか','画面の見やすさや操作の分かりやすさが十分であるか','対応している言語の種類が業務の範囲を満たしているか','入力データの学習利用の有無と、管理者が統制できるか'] },
  't5-4': { q:'AI導入を検討した結果「現時点では導入しない」と判断した場合、適切な対応はどれですか。', c:['導入しない以上は、検討の経緯を記録として残しておかなくてよい','判断した理由と、どのような条件が整えば導入するかを残す','今後は検討しない方針であることを、社内に明確に宣言する','競合他社が導入した時点で、自社も自動的に導入へ踏み切る'] },
  't5-5': { q:'業務にAIを使うことで最も効果が見込めるのは、一般にどのような作業ですか。', c:['経営の方向を決めるような、最終的な意思決定を伴う作業','契約や法令の解釈など、法的な責任を直接に伴う判断','顧客との信頼関係に直結する、対面での折衝や交渉の業務','定型的で反復が多く、下書きや情報の整理が中心の作業'] },

  't6-1': { q:'AIを使ったフィッシングメールが従来のものより見分けにくい理由はどれですか。', c:['添付されるファイルが必ず暗号化されるようになったため','日本語が自然になり、文面の不自然さで気づけなくなったため','送信元のメールアドレスの偽装が新たに可能になったため','ウイルス対策ソフトを無効化する機能を備えているため'] },
  't6-2': { q:'プロンプトインジェクションが成立しうる状況はどれですか。', c:['自社で作成した社内文書をAIに読ませて要約させるとき','一般的な知識について、AIに質問して回答を得るとき','外部から受け取った文書やWebページをAIに読ませるとき','数値の計算や表の集計をAIに任せて処理させるとき'] },
  't6-3': { q:'取引先から「振込口座を変更したい」という電話があった場合、適切な対応はどれですか。', c:['従来から把握している連絡先に、自社から掛け直して確認する','電話口の担当者の声を聞き、本人であることを確認して進める','折り返し先として伝えられた番号に掛け直し、内容を確認する','同じ内容のメールが届いていることをもって確認済みとする'] },
  't6-4': { q:'AIエージェントに業務を任せる際、最も注意すべき点はどれですか。', c:['処理の速度が人間より遅く、かえって時間がかかる場合がある','利用料金が従量課金であり、費用が想定を超える場合がある','日本語への対応が不完全で、指示が正しく伝わらない場合がある','人の確認を挟まずに送信や決済などの操作を実行しうること'] },
  't6-5': { q:'インシデント対応フローを整備する目的として最も適切なものはどれですか。', c:['責任の所在を明確にし、関係者への処分を行いやすくするため','監督官庁への報告が不要となるよう、社内で処理を完結するため','事態が起きた際に判断へ迷わず、初動を遅らせずに済むため','従業員にインシデントが起こりうることを日頃から意識させるため'] },

  't7-1': { q:'大規模言語モデルが事実と異なる内容を出力することがあるのは、根本的にはなぜですか。', c:['次に来る語を統計的に予測する仕組みで、真偽を検証していないから','学習に用いられたデータのほとんどが誤った内容のものになっているから','インターネットに接続しておらず、最新情報を持たないから','日本語への対応が英語に比べて十分でない状態にあるから'] },
  't7-2': { q:'社内データをAIに活用させる方法として、RAGが適しているのはどのような場合ですか。', c:['モデルの応答の文体や口調そのものを変えたい場合','学習に用いるデータを完全に非公開のままの状態で扱いたい場合','インターネット接続のない環境で動かしたい場合','頻繁に更新される社内文書を最新の状態で参照させたい場合'] },
  't7-3': { q:'プロンプトエンジニアリングの観点から、業務での指示文として望ましいものはどれですか。', c:['解釈の幅を残すため、できるだけ短く簡潔な指示にしておく','前提と目的、出力の形式、守るべき制約を明示しておく','専門用語を避け、誰にでも分かる平易な言葉だけを使う','関連する複数の依頼を、一度にまとめて指示しておく'] },
  't7-4': { q:'MCPのような仕組みでAIが外部ツールと連携できることで、新たに生じる論点はどれですか。', c:['連携の処理が挟まることで、AIの応答速度が低下すること','接続するツールが増えるほど、利用料金が上がっていくこと','接続先のデータを読み書きでき、権限設計を誤ると被害が広がること','連携先の仕様に合わせる必要があるため、日本語での指示が使えなくなること'] },
  't7-5': { q:'AI活用のロードマップを描く際、各段階で併せて検討すべきことはどれですか。', c:['その段階で扱う情報の範囲と、対応するセキュリティ対策','導入するツールの提供元が持つブランドとしての価値の高さ','同業の競合他社が導入に踏み切る時期とその内容','AI技術の分野で今後流行が移り変わっていく方向'] }
};

const TEST_TOTAL_PER_CHAPTER = COURSE.rules.testQuestions;   // course.json の rules
const TEST_PASS_COUNT_UI = COURSE.rules.testPass;
const _testAnswered = {};   // { questionId: true } 回答済み
const _testAttemptsLeft = {};
const _testRetry = {};   // { 章: true } 学び直しの途中（戻した設問にまだ答えていない）   // { questionId: 残りの回答回数 }（サーバーの値。0 の設問は学び直しの対象にしない）
const _testResultCh = {};   // { 章番号: {answered, correct, passed} }

function renderTestChapter(ch){
  const host = document.getElementById('test-body-'+ch);
  if(!host) return;
  let html = '';
  for(let i=1;i<=TEST_TOTAL_PER_CHAPTER;i++){
    const qid = `t${ch}-${i}`;
    const d = TEST_QUESTIONS[qid];
    if(!d) continue;
    const opts = d.c.map((t,idx)=>
      `<button class="test-opt" id="to-${qid}-${idx}" onclick="answerTest('${qid}',${idx})">${t}</button>`
    ).join('');
    html += `
      <div class="test-q" id="tq-${qid}">
        <div class="test-q-no">問 ${i} / ${TEST_TOTAL_PER_CHAPTER}</div>
        <p class="test-q-text">${d.q}</p>
        <div class="test-opts" id="tos-${qid}">${opts}</div>
        <div class="test-exp" id="te-${qid}"></div>
      </div>`;
  }
  host.innerHTML = html;
}

async function loadTestChapter(ch){
  const host = document.getElementById('test-body-'+ch);
  if(!host) return;
  if(!host.dataset.rendered){ renderTestChapter(ch); host.dataset.rendered = '1'; }
  const t = await workToken();
  if(!t) return;
  try{
    const r = await fetch(`${WORK_API}/api/test/${ch}`, { headers:{ Authorization:'Bearer '+t }});
    if(!r.ok) return;
    const d = await r.json();
    (d.answers||[]).forEach(a=>{
      _testAnswered[a.question_id] = true;
      if(a.attemptsLeft != null) _testAttemptsLeft[a.question_id] = a.attemptsLeft;
      lockTestQuestion(a.question_id, a.selected, a.is_correct, a.explanation);
    });
    _testResultCh[ch] = { answered:d.answered, correct:d.correct, passed:d.passed,
                          firstCorrect:d.firstCorrect, attempts:d.attempts, wrongIds:d.wrongIds||[] };
    updateTestSummary(ch);
  }catch(e){ /* 表示は継続 */ }
}

// 解説は引数で受け取る。クライアントには持たせない（回答前に先読みされるため）。
function lockTestQuestion(qid, selected, isCorrect, explanation){
  const box = document.getElementById('tos-'+qid);
  const exp = document.getElementById('te-'+qid);
  const d = TEST_QUESTIONS[qid];
  if(!box || !d) return;
  box.querySelectorAll('.test-opt').forEach((b,idx)=>{
    b.disabled = true;
    b.classList.add('is-done');
    if(idx === selected) b.classList.add(isCorrect ? 'is-correct' : 'is-wrong');
  });
  if(exp){
    exp.className = 'test-exp is-open ' + (isCorrect ? 'ok' : 'ng');
    exp.innerHTML = `<div class="test-verdict">${isCorrect ? '正解' : '不正解'}</div><div>${escHtml(explanation)}</div>`;
  }
}

async function answerTest(qid, idx){
  if(_testAnswered[qid]) return;
  const ch = parseInt(qid.split('-')[0].slice(1),10);
  const box = document.getElementById('tos-'+qid);
  if(box) box.querySelectorAll('.test-opt').forEach(b=>b.disabled = true);
  const t = await workToken();
  if(!t){
    if(box) box.querySelectorAll('.test-opt').forEach(b=>b.disabled = false);
    return;
  }
  try{
    const r = await fetch(`${WORK_API}/api/test/answer`, {
      method:'POST',
      headers:{ Authorization:'Bearer '+t, 'Content-Type':'application/json' },
      body: JSON.stringify({ questionId:qid, selected:idx })
    });
    const d = await r.json();
    if(!r.ok){
      if(box) box.querySelectorAll('.test-opt').forEach(b=>b.disabled = false);
      alert(d.error || '回答を記録できませんでした。時間をおいてもう一度お試しください。');
      return;
    }
    _testAnswered[qid] = true;
    if(d.attemptsLeft != null) _testAttemptsLeft[qid] = d.attemptsLeft;
    lockTestQuestion(qid, idx, d.isCorrect, d.explanation);
    const cur = _testResultCh[ch] || { answered:0, correct:0 };
    // 学び直し（前に間違えた設問への再回答）では、回答数は増やさず、正解なら正解数だけ増やす
    // （以前の数え方では、学び直しのたびに回答数が5を超えていった）
    const again = (cur.wrongIds || []).includes(qid);
    if(!again) cur.answered += 1;
    if(d.isCorrect) cur.correct += 1;
    cur.passed = cur.answered >= TEST_TOTAL_PER_CHAPTER && cur.correct >= TEST_PASS_COUNT_UI;
    if(!d.isCorrect){ cur.wrongIds = (cur.wrongIds || []).filter(x => x !== qid).concat(qid); }
    else if(cur.wrongIds){ cur.wrongIds = cur.wrongIds.filter(x => x !== qid); }
    if(cur.answered >= TEST_TOTAL_PER_CHAPTER) refreshRecords();
    _testResultCh[ch] = cur;
    updateTestSummary(ch);
  }catch(e){
    if(box) box.querySelectorAll('.test-opt').forEach(b=>b.disabled = false);
    alert('回答を送信できませんでした。通信環境をご確認のうえ、もう一度お試しください。');
  }
}

function updateTestSummary(ch){
  const el = document.getElementById('test-summary-'+ch);
  if(!el) return;
  const r = _testResultCh[ch] || { answered:0, correct:0, passed:false };
  if(r.answered === 0){ el.innerHTML = ''; return; }
  // 学び直しの途中は、戻した設問にすべて答えるまで結果を出さない（1問答えた時点で「不合格」が出ていた）
  if(_testRetry[ch]){
    const left = (r.wrongIds || []).filter(q => !_testAnswered[q]).length;
    if(left > 0){
      el.innerHTML = `<div class="test-retry-box"><p class="test-retry-msg">学び直し中：あと ${left} 問。第${ch}章のレッスンで確かめてから答えてください。</p>
        <div class="test-retry-links">${chapterLessonLinks(ch)}</div></div>`;
      return;
    }
    delete _testRetry[ch];
  }
  const done = r.answered >= TEST_TOTAL_PER_CHAPTER;
  if(!done){
    el.innerHTML = `<span class="test-prog">回答済み ${r.answered} / ${TEST_TOTAL_PER_CHAPTER} 問</span>`;
    return;
  }
  const att = (r.attempts && r.attempts > 1) ? `　<span class="test-att">受験 ${r.attempts} 回目</span>` : '';
  if(r.passed){
    // 次の行き先（次の章の最初のレッスン、最後なら認定試験・修了証）を出す
    const nx = typeof nextStep === 'function' ? nextStep() : null;
    el.innerHTML = `<span class="test-pass">${r.correct} / ${TEST_TOTAL_PER_CHAPTER} 問正解　合格</span>${att}`
      + (nx ? `<div style="margin-top:12px"><button type="button" class="form-submit ob-btn" onclick="${nx.onclick}">次へ：${escHtml(nx.label)} →</button></div>` : '');
    return;
  }
  const wrong = (r.wrongIds||[]).length;
  const retryable = (r.wrongIds||[]).filter(q => _testAttemptsLeft[q] !== 0).length;
  if(!retryable){
    el.innerHTML = `<span class="test-fail">${r.correct} / ${TEST_TOTAL_PER_CHAPTER} 問正解　合格ラインは ${TEST_PASS_COUNT_UI} 問です</span>${att}
      <div class="test-retry-box"><p class="test-retry-msg">間違えた設問は、回答できる回数（各${COURSE.rules.testMaxAttempts}回）を使い切りました。<br>
      章のレッスンを読み直したうえで、お問い合わせから運営までご連絡ください。</p></div>`;
    return;
  }
  el.innerHTML =
    `<span class="test-fail">${r.correct} / ${TEST_TOTAL_PER_CHAPTER} 問正解　合格ラインは ${TEST_PASS_COUNT_UI} 問です</span>${att}
     <div class="test-retry-box">
       <p class="test-retry-msg">間違えた ${wrong} 問について、対応するレッスンをもう一度確認してから再挑戦できます。<br>
       <span class="test-retry-sub">受講記録と進捗は保持されます。初回の回答は理解度の記録として残り、修了判定には最新の回答を用います。</span></p>
       <button class="test-retry-btn" onclick="retryTestChapter(${ch})">間違えた問題を学び直す</button>
     </div>`;
}

// 間違えた設問だけを未回答に戻し、参照レッスンへの導線を出す章のレッスンへのリンク（番号と題名）。
function chapterLessonLinks(ch){
  const c = COURSE.chapters.find(x => x.no === ch);
  if(!c) return '';
  return c.lessons.map(id => `<button type="button" class="work-nav-btn" data-tab="${id}" onclick="switchTab('${id}')">L${lessonNo(id)} ${escHtml(COURSE.titles[id] || '')}</button>`).join('');
}
function retryTestChapter(ch){
  _testRetry[ch] = true;
  const r = _testResultCh[ch] || {};
  // 回答回数を使い切った設問は戻さない
  const wrong = (r.wrongIds || []).filter(q => _testAttemptsLeft[q] !== 0);
  if(!wrong.length) return;
  wrong.forEach(qid => {
    _testAnswered[qid] = false;
    const box = document.getElementById('tos-'+qid);
    const exp = document.getElementById('te-'+qid);
    if(box){
      box.querySelectorAll('.test-opt').forEach(b=>{
        b.disabled = false;
        b.classList.remove('is-done','is-correct','is-wrong');
      });
    }
    if(exp){ exp.className = 'test-exp'; exp.innerHTML = ''; }
    const card = document.getElementById('tq-'+qid);
    if(card) card.classList.add('is-retry');
  });
  const el = document.getElementById('test-summary-'+ch);
  if(el){
    el.innerHTML = `<div class="test-retry-box">
      <p class="test-retry-msg">学び直しモードです。間違えた ${wrong.length} 問の範囲を、第${ch}章のレッスンで確かめてから再挑戦してください。</p>
      <div class="test-retry-links">${chapterLessonLinks(ch)}</div>
    </div>`;
  }
  const first = document.getElementById('tq-'+wrong[0]);
  if(first) first.scrollIntoView({behavior:'smooth', block:'center'});
}

window.loadTestChapter = loadTestChapter;
// ===================== /章末確認テスト =====================

// ===================== /章末ワーク =====================


// switchTab後にサイドバー・学習UIを更新
window.switchTab = function(name){
  if(!canAccessTab(name)) return;
  if(window._aslTestFocus && name !== 'work' + window._aslTestFocus) window._aslTestFocus = 0;
  if(currentUser){
    // 前回の続きから再開できるよう、最後に開いたタブを控える（ユーザーごと）
    store.set('asl_last_tab', { u: currentUser.id, t: name });
    // ログイン後はshowActiveLessonで直接切替
    showActiveLesson(name);
    const completed = loadCompleted();
    buildSideNav(completed);
    const lc = document.getElementById('learn-content');
    if(lc) lc.scrollTop = 0;
  } else {
    showTabGuest(name);
  }
  // 章末ワークタブなら内容を読み込む
  if(/^work\d+$/.test(name)){
    const _ch = parseInt(name.slice(4),10);
    loadWorkChapter(_ch);
    loadTestChapter(_ch);
  }
};
window.openGuideModal  = function(){ document.getElementById('guide-modal-overlay').classList.add('show'); };
window.closeGuideModal = function(){ document.getElementById('guide-modal-overlay').classList.remove('show'); };

/* ===== 受講済み管理・修了証 ===== */
// 受講順・章・プランの範囲は course.json が正本（ファイル先頭の COURSE ブロック）。ここに件数や並びを書かないこと。
// 内部ID（lesson28 など）は追加順で、画面の番号（lessonNo()）は受講順。Worker の LESSON_ORDER も同じ正本から作る。
function lessonNo(id){ const i = ALL_LESSONS_LIST.indexOf(id); return i < 0 ? '' : String(i + 1).padStart(2,'0'); }
function getPlanLessons(){
  const plan = currentUserPlan || 'free';
  // 未知のプラン（無料登録のみの 'free' など）は0本。Worker と同じく拒否側に倒す
  return ALL_LESSONS_LIST.slice(0, PLAN_LESSON_COUNT[plan] || 0);
}

function loadCompleted(){
  try{
    let arr = store.getJSON('asl_completed', []) || [];
    // 序章（intro）が含まれていたら削除（旧データのクリーンアップ）
    arr = arr.filter(id => id !== 'intro');
    return arr;
  }catch(e){ return []; }
}
function saveCompleted(arr){ store.set('asl_completed', arr); }

function markComplete(tabId){
  if(!currentUser){ openAuthModal(); return; }
  const completed = loadCompleted();
  if(!completed.includes(tabId)) completed.push(tabId);
  saveCompleted(completed);
  updateCompletionUI();
  updateDashboard();
  // 受講記録はWorker経由で保存する（受講日時はサーバー時刻で確定）
  if(currentUser){
    flushActive();                              // 押した瞬間までを確定してから送る
    const secs    = loadActive()[tabId] || null;
    const startIso = loadStarts()[tabId] || null;
    apiFetch('/api/lesson/complete', {
      method: 'POST',
      body: JSON.stringify({
        lessonId: tabId,
        served: servedFor(tabId),
        startedAt: startIso,
        durationSeconds: secs,
      }),
    }).catch(e => {
      // 保存に失敗したら画面の「受講済み」を取り消して知らせる。
      console.warn('受講記録の保存に失敗:', e.message);
      saveCompleted(loadCompleted().filter(x => x !== tabId));
      const b = document.getElementById('complete-' + tabId);
      if(b){ b.classList.remove('done'); b.textContent = '✓ 受講済みにする（もう一度）'; }
      updateCompletionUI();
      updateDashboard();
      alert('受講済みの記録を保存できませんでした。\n' + (e.message || '') + '\n通信環境をご確認のうえ、もう一度「受講済み」を押してください。');
    });
  }
}

// 次にやること：受講順で最初の未了（レッスン → その章のワーク → 確認テスト …）。最後は認定試験・修了証。
// 戻り値：{ label, onclick, run } または null
function nextStep(){
  const completed = loadCompleted();
  const sm = window._recSummary || {};
  const mk = (label, onclick, run) => ({ label, onclick, run });
  for(const id of planTabsOrdered()){
    if(ALL_LESSONS_LIST.includes(id)){
      if(!completed.includes(id)) return mk('L' + lessonNo(id) + ' ' + (COURSE.titles[id] || ''), `switchTab('${id}')`, () => window.switchTab(id));
      continue;
    }
    const m = /^work(\d+)$/.exec(id);
    if(!m) continue;
    const c = Number(m[1]);
    if(!_workSubmittedCh[c] && !(sm.workDoneList || []).includes(c)) return mk(`第${c}章の章末ワーク`, `switchTab('${id}')`, () => window.switchTab(id));
    if(!(_testResultCh[c] && _testResultCh[c].passed) && !(sm.testPassedList || []).includes(c)) return mk(`第${c}章の確認テスト`, `goChapterTest(${c})`, () => goChapterTest(c));
  }
  if(currentUserPlan === 'premium' && !(sm.examPassed || currentExamPassed)) return mk('認定試験', `switchTab('exam')`, () => window.switchTab('exam'));
  if(currentUserPlan === 'basic') return mk('受講実績レポート（Basic の全レッスンを受講済み）', `switchTab('records')`, () => window.switchTab('records'));
  if(currentUserPlan === 'standard' || currentUserPlan === 'premium') return mk('修了証の申請', `switchTab('certificate')`, () => window.switchTab('certificate'));
  return null;
}
window.nextStep = nextStep;

function updateCompletionUI(){
  const completed = loadCompleted();
  ALL_LESSONS_LIST.forEach(id => {
    const btn  = document.getElementById('complete-' + id);
    const note = document.getElementById('note-' + id);
    if(btn){
      if(completed.includes(id)){ btn.classList.add('done'); btn.textContent = '✓ 受講済み'; }
      else { btn.classList.remove('done'); }
    }
    if(note){
      // 理解度チェックが未了で押せないあいだは、その理由を残す
      if(btn && btn.classList.contains('quiz-locked')) note.textContent = '理解度チェックに正解すると受講済みにできます';
      else {
        const pl = getPlanLessons();
        note.innerHTML = '<span>' + pl.filter(x => completed.includes(x)).length + '</span> / ' + pl.length + ' レッスン完了';
      }
    }
  });
  const fill = document.getElementById('cert-bar-fill');
  const count = document.getElementById('cert-count');
  const notice = document.getElementById('cert-locked-notice');
  const submitBtn = document.getElementById('cert-submit-btn');
  // プランごとの受講範囲（PLAN_LESSON_COUNT / PLAN_RANGE_LABEL）
  const plan = currentUserPlan || 'free';
  const planLessons = getPlanLessons();
  const planDone = planLessons.filter(id => completed.includes(id));
  const planTotal = planLessons.length;
  const pct = planTotal ? planDone.length / planTotal * 100 : 0;
  const rangeLabel = PLAN_RANGE_LABEL[plan] || '';

  if(fill)  fill.style.width = pct + '%';
  if(count) count.textContent = planDone.length;
  const countTotal = document.getElementById('cert-count-total');
  if(countTotal) countTotal.textContent = planTotal;
  const countRange = document.getElementById('cert-count-range');
  if(countRange) countRange.textContent = rangeLabel;
  const certEmail = document.getElementById('cert-email');
  if(certEmail && currentUser) certEmail.value = currentUser.email || '';

  // ナビ進捗ボタン（全プランで更新する）
  const navProg  = document.getElementById('nav-progress');
  const progFill = document.getElementById('nav-progress-fill');
  if(navProg){ navProg.textContent = '📊 進捗 ' + planDone.length + '/' + planTotal; if(currentUser) navProg.style.display = planTotal ? 'inline-flex' : 'none'; }
  const examBadge = document.getElementById('exam-plan-badge');
  if(examBadge) examBadge.textContent = plan === 'premium' ? '✅ ご契約のプランで受験できます' : '🔒 Premiumプラン限定';
  if(progFill) progFill.style.width = pct + '%';

  // 修了証は Standard・Premium だけ
  if(plan !== 'standard' && plan !== 'premium'){
    if(notice) notice.style.display = 'none';
    if(submitBtn){ submitBtn.disabled = true; submitBtn.style.opacity = '.4'; }
    return;
  }
  // レッスン追加より前に修了証を受け取った人は、追加分が未受講でも申請（再発行）できる
  let allDone = planLessons.every(id => completed.includes(id)) || currentCertifiedEarlier;
  const go = (onclick, text) => ` <button type="button" class="ck-go" onclick="${onclick}">${escHtml(text)} →</button>`;
  const firstLesson = planLessons.find(id => !completed.includes(id));
  // 修了条件のチェックリスト。
  const sm = window._recSummary;
  const ckEl = document.getElementById('cert-checklist');
  if(ckEl && sm){
    // 足りない条件には、その場所へ移るボタンを付ける
    const row = (ok, label, val, btn) => `<div class="ck-row">${ok ? '✅' : '⬜'} ${label}　<span style="color:var(--text3)">${val}</span>${ok ? '' : (btn || '')}</div>`;
    const chs = sm.chaptersTotal || 0;
    const missW = Array.from({length: chs}, (_, i) => i + 1).filter(c => !(sm.workDoneList || []).includes(c));
    const missT = Array.from({length: chs}, (_, i) => i + 1).filter(c => !(sm.testPassedList || []).includes(c));
    let h = row(planDone.length >= planTotal || currentCertifiedEarlier, '全レッスンを受講済みにする', planDone.length + ' / ' + planTotal,
      firstLesson ? go(`switchTab('${firstLesson}')`, 'L' + lessonNo(firstLesson) + ' を開く') : '');
    if(chs){
      h += row((sm.chaptersSubmitted || 0) >= chs, '章末ワークを全章提出する', (sm.chaptersSubmitted || 0) + ' / ' + chs + ' 章',
        missW.length ? go(`switchTab('work${missW[0]}')`, `第${missW[0]}章のワークへ`) : '');
      h += row((sm.testPassedChapters || 0) >= chs, `章末確認テストに全章合格する（各章${COURSE.rules.testQuestions}問中${COURSE.rules.testPass}問以上）`, (sm.testPassedChapters || 0) + ' / ' + chs + ' 章',
        missT.length ? go(`goChapterTest(${missT[0]})`, `第${missT[0]}章のテストへ`) : '');
    }
    if(plan === 'premium') h += row(!!(sm.examPassed || currentExamPassed), '認定試験に合格する', (sm.examPassed || currentExamPassed) ? '合格' : '未合格', go(`switchTab('exam')`, '認定試験へ'));
    ckEl.innerHTML = '<div style="font-weight:700;margin-bottom:4px">修了の条件</div>' + h;
    if(!currentCertifiedEarlier && chs && ((sm.chaptersSubmitted || 0) < chs || (sm.testPassedChapters || 0) < chs)) allDone = false;
  }
  const examPassed = currentExamPassed;
  const canApply = plan === 'premium' ? (allDone && examPassed) : allDone;
  const certDesc = document.querySelector('.cert-desc');
  if(certDesc){
    // 説明は修了の条件（下の一覧）と同じ内容にする
    const chs = COURSE.workChapters[plan] || 0;
    const conds = [`全${planLessons.length}レッスンの受講`].concat(chs ? [`章末ワークの提出（全${chs}章）`, `章末確認テストの合格（全${chs}章）`] : []).concat(plan === 'premium' ? ['認定試験の合格'] : []);
    certDesc.innerHTML = escHtml(conds.join('・')) + 'で、修了証を申請できます。<br>申請後、修了証PDFをメールでお送りします。';
  }

  // まだ満たしていない最初の条件を示す
  let noticeText = `⚠️ ${PLAN_RANGE_LABEL[plan] || ''}の全${planTotal}レッスンを受講済みにしてから申請してください（いま ${planDone.length} / ${planTotal}）。`;
  const _sm = window._recSummary;
  if(planDone.length >= planTotal || currentCertifiedEarlier){
    const chs = _sm ? (_sm.chaptersTotal || 0) : 0;
    if(_sm && chs && (_sm.chaptersSubmitted || 0) < chs && !currentCertifiedEarlier){
      const miss = Array.from({length: chs}, (_, i) => i + 1).filter(c => !(_sm.workDoneList || []).includes(c));
      noticeText = `⚠️ 章末ワークが未提出の章があります（第${miss.join('・')}章）。上の「修了の条件」のボタンから開けます。`;
    } else if(_sm && chs && (_sm.testPassedChapters || 0) < chs && !currentCertifiedEarlier){
      const miss = Array.from({length: chs}, (_, i) => i + 1).filter(c => !(_sm.testPassedList || []).includes(c));
      noticeText = `⚠️ 章末確認テストに未合格の章があります（第${miss.join('・')}章）。上の「修了の条件」のボタンから開けます。`;
    } else if(plan === 'premium' && !examPassed){
      noticeText = '⚠️ 認定試験に合格してから申請してください。';
    }
  }
  // 同じプランの修了証が発行済みなら、申請はできない（Worker は 409 を返す）。最初から「発行済み」と示す
  if(currentCertifiedEarlier){
    window._certSubmitted = true;
    if(notice){ notice.textContent = '✅ このプランの修了証は発行済みです。再送や氏名の訂正は、お問い合わせからご連絡ください。'; notice.style.display = 'block'; }
    if(submitBtn){ submitBtn.disabled = true; submitBtn.textContent = '発行済み'; }
    return;
  }
  if(notice){ notice.textContent = noticeText; notice.style.display = canApply ? 'none' : 'block'; }
  if(submitBtn){ const ok = canApply && !window._certSubmitted; submitBtn.disabled = !ok; submitBtn.style.opacity = ok ? '1' : '.4'; }
  if(canApply){ const tb = document.getElementById('cert-tab-btn'); if(tb) tb.style.display = ''; }
}


// ===== 認定試験データ（60問・4択） =====
// 正解（a）と解説（e）はここには無い。Worker の EXAM_ANSWER_KEY だけが持つ。
// クライアントは選択した「元の選択肢番号」を id とともに送り、採点結果を受け取る。
// id は Worker 側のキーと1対1で対応する。並べ替えても id を変えないこと。
const EXAM_QUESTIONS = [
  {id:"e1", lesson:"lesson6", q:"Claudeの利用規約上、企業利用時にデータ学習を無効化するために必要なものはどれか？", o:["API利用またはTeamsプランへの加入","アカウント設定画面での学習オフの切り替え","利用開始時に日本語での利用目的の申請","無料プランでも初回ログイン時に自動適用"]},
  {id:"e2", lesson:"lesson1", q:"「ChatGPT Business」（旧 Team）プランでデータ学習がオフになる条件として正しいものはどれか？", o:["管理者の設定が必要","毎回の手動解除が必要","自動的にオフになる","日本国内では常にオフ"]},
  {id:"e3", lesson:"lesson26", q:"医療・介護分野でAIを活用する際に最も注意すべき法規制はどれか？", o:["個人情報保護法（要配慮個人情報の取り扱い）","著作権法（学習データの利用と権利処理）","景品表示法（サービス広告の表示規制）","電気通信事業法（通信の秘密の保護義務）"]},
  {id:"e4", lesson:"lesson4", q:"クラウドAIとオンプレAIの最大の違いは何か？", o:["応答が返るまでの処理速度","利用画面の操作性の設計","障害時の問い合わせ体制","データの保管・管理場所"]},
  {id:"e5", lesson:"lesson19", q:"プロンプトインジェクション攻撃とはどのような攻撃か？", o:["悪意のある指示をAIに混入させてAIの行動を乗っ取る","学習済みモデルの重みを外部から複製して持ち出す","外部のAIサーバーに物理的に侵入して機器を奪い取る","AIに大量の要求を送り応答できない状態に追い込む"]},
  {id:"e6", lesson:"lesson24", q:"プロンプトエンジニアリングにおいて「Few-shot」とはどのような手法か？", o:["音声で読み上げてAIに指示を伝える手法","指示文をできるだけ短く区切って渡す手法","少人数のチームでAIを共同利用する手法","具体的な例を複数示してから指示する手法"]},
  {id:"e7", lesson:"lesson19", q:"企業がAIエージェントを導入する際、プロンプトインジェクション対策として最も重要なものはどれか？", o:["社内のすべての業務においてAIエージェントの利用を全面的に禁止する","外部コンテンツをそのままAIに処理させない・サンドボックス化","エージェントの応答速度を落として処理の負荷を下げる","エージェントを利用できる従業員の人数を上限まで絞る"]},
  {id:"e8", lesson:"lesson7", q:"AIに入力してはいけない情報として「個人情報」に該当するものはどれか？", o:["公開済みのプレスリリース","顧客の氏名・住所・電話番号","業界全体の市場統計データ","製品の一般的な仕様説明"]},
  {id:"e9", lesson:"lesson20", q:"ディープフェイクを使った詐欺の代表的な事例はどれか？", o:["AIが本人確認を通過して口座から自動で送金した事例","音声合成AIが電子契約書に本人の署名を代行した事例","偽の経営者映像でCFOに送金を指示した香港の事件","AIが偽の身分証を生成して銀行口座を開設した事例"]},
  {id:"e10", lesson:"lesson8", q:"不動産業でAIを使う際に特に注意すべき情報として正しいものはどれか？", o:["公開中の物件の間取り図","建物の築年数と構造の種別","地域ごとの相場価格の推移","顧客の収入情報や家族構成"]},
  {id:"e11", lesson:"lesson27", q:"AIエージェント時代に企業が備えるべきセキュリティ対策として最も重要なものはどれか？", o:["判断を伴う業務をすべてエージェントに移管する","エージェントが自律的に守るため対策は不要になる","業務でのエージェント利用を全面的に取りやめる","エージェントの行動ログを記録し権限を最小化する"]},
  {id:"e12", lesson:"lesson25", q:"RAG（Retrieval-Augmented Generation）の主なリスクはどれか？", o:["参照する文書が増えるほど回答が短く要約されてしまう","検索対象の社内データベースへの不正アクセスが発生しうる","検索の仕組みが日本語の社内文書に対応していない","検索対象にした社内文書をAIが自動的に学習して記憶してしまう"]},
  {id:"e13", lesson:"lesson11", q:"ハルシネーションとは何か？", o:["AIが人間の感情を模倣して共感的に応答する現象","AIが学習時点より古いデータだけを参照する現象","AIが指示より長い文章を繰り返し生成する現象","AIが事実に基づかない情報を自信を持って出力すること"]},
  {id:"e14", lesson:"lesson27", q:"MCP（Model Context Protocol）とは何か？", o:["AIとの通信内容を暗号化して第三者から保護するための標準規格","AIが利用者の意図と感情を推定するための標準規格","AIと外部ツール・データソースを接続するための標準規格","AIの推論にかかる計算量を削減するための技術"]},
  {id:"e15", lesson:"lesson15", q:"AIセキュリティ教育を効果的に行うための方法として最も適切なものはどれか？", o:["具体的な事例を使った定期的な研修","年1回の全社メールによる通知","外部の専門家に対応をすべて委託する","条文をそのまま読み合わせる勉強会"]},
  {id:"e16", lesson:"lesson9", q:"生成AIに個人情報を入力する行為が個人情報保護法上問題になり得るケースはどれか？", o:["匿名化して個人を特定できない統計値を入力する","公開済みの情報を事実確認のために入力する","本人の同意なく第三者の個人情報を入力する","自分自身の連絡先を業務の下書きに入力する"]},
  {id:"e17", lesson:"lesson16", q:"飲食業でAIを活用する際に注意すべき情報漏洩リスクとして最も適切なものはどれか？", o:["AIの提案による食材の発注数量の誤り","AIが厨房機器を操作して調理を始める","自店の定番レシピが検索結果に出ること","アレルギー情報など顧客の個人情報の漏洩"]},
  {id:"e18", lesson:"lesson22", q:"漏洩インシデント発生後に「やってはいけないこと」はどれか？", o:["事実を隠蔽して様子を見る","関係者に事実を報告する","ログなどの証拠を保全する","使用していたAIツールを停止する"]},
  {id:"e19", lesson:"lesson25", q:"社内データでファインチューニングを行う前に確認すべき最重要事項はどれか？", o:["学習の実行にかかる費用と所要時間の見積もり","利用者に提示する画面の設計と操作手順の整理","学習データに個人情報・機密情報が含まれていないか","利用するAIモデルのバージョンと更新の頻度"]},
  {id:"e20", lesson:"lesson2", q:"ChatGPTで「チャット履歴とトレーニング」をオフにした場合の正しい説明はどれか？", o:["30日後に会話データが削除される","過去の会話履歴が一切残らない","他のユーザーと会話の内容が共有される","入力データがOpenAIに送信されなくなる"]},
  {id:"e21", lesson:"lesson10", q:"AIが生成した画像を商業利用する際に注意すべき著作権上のリスクはどれか？", o:["生成された画像の著作権は自動的にAI開発事業者へ帰属する","生成された画像の商業利用は法令によって一律に禁止されている","学習データに含まれる既存著作物に酷似した画像が生成されるリスク","AI自身が著作者として登録され権利を主張できることになる"]},
  {id:"e22", lesson:"lesson9", q:"個人情報保護法における「匿名加工情報」の説明として正しいものはどれか？", o:["医療分野で取り扱う情報だけを指す区分のこと","氏名や連絡先によって個人を特定できる情報","社内規程によって社外秘に分類されている情報","特定の個人を識別できないよう加工された情報"]},
  {id:"e23", lesson:"lesson21", q:"業務自動化にAIエージェントを使う前に確認すべき最重要事項はどれか？", o:["利用するプランが無料か有料かどうかの確認","利用するAIが最新のモデルであるかどうかの確認","エージェントがアクセスできる範囲と権限の最小化","操作画面が現場の担当者にとって使いやすいか"]},
  {id:"e24", lesson:"lesson24", q:"セキュリティ観点でプロンプト設計時に注意すべき点はどれか？", o:["機密情報や個人情報をプロンプトに含めない","指示は日本語ではなく必ず英語で書くこと","同じ指示を避け毎回異なる文面を使うこと","指示文をできるだけ長く詳細に書き込むこと"]},
  {id:"e25", lesson:"lesson12", q:"AI利用ガイドラインを策定した後に行うべき重要なアクションはどれか？", o:["自社サイトで外部に全文を公開すること","法務部だけが内容を管理して保管すること","定期的な見直しと全社員への周知・教育","策定した時点で完了として保管しておくこと"]},
  {id:"e26", lesson:"lesson21", q:"AIエージェントが自律的に動作することの主なリスクはどれか？", o:["担当者の業務量が減り管理が緩むこと","意図しないデータ送信・操作が発生する","処理の待ち時間が業務全体で伸びること","利用料金が想定より下がりすぎること"]},
  {id:"e27", lesson:"lesson7", q:"「社外秘」に分類すべき情報として最も適切なものはどれか？", o:["公開している営業時間","登記済みの代表者氏名","未発表の新製品情報","登記されている本店所在地"]},
  {id:"e28", lesson:"lesson26", q:"法律・税務士業でAIを使う場合、倫理的に特に問題になりうるケースはどれか？", o:["公開されている判例をAIを使って検索する","依頼人との相談内容を無断でAIに入力する","公開済み文書の誤字脱字をAIで修正する","社内会議の議事録作成をAIに補助させる"]},
  {id:"e29", lesson:"lesson14", q:"中小企業がまず取り組むべき最低限のAIセキュリティ対策として正しいものはどれか？", o:["専任のAI担当者を新たに雇用すること","入力禁止情報のルール化と全社周知","業務でのAI利用を一切禁止すること","高額なセキュリティ製品を先に導入すること"]},
  {id:"e30", lesson:"lesson4", q:"中小企業がオンプレAIを導入する際の主なデメリットとして正しいものはどれか？", o:["初期コストと保守コストが高い","クラウド型より性能が著しく低い","国内の法律で導入が禁止されている","データが必ず外部に送信される"]},
  {id:"e31", lesson:"lesson6", q:"Microsoft Copilot（M365）のセキュリティ上の特徴として正しいものはどれか？", o:["組織の設定によらず無条件で外部に共有される仕組み","国内の個人情報保護法の適用対象から外れる仕組み","入力内容がすべてOpenAI社へ直接送信される","既存のM365セキュリティポリシーが適用される"]},
  {id:"e32", lesson:"lesson14", q:"「5つのことで8割のリスクを下げる」対策のひとつとして正しいものはどれか？", o:["AIの利用ログをすべて定期的に削除する","全社員のパスワードを毎日変更する","業務用と個人用AIアカウントの分離","社内でのAIツール利用を全面的に禁止する"]},
  {id:"e33", lesson:"lesson5", q:"ChatGPT無料版を業務で使う場合のリスクとして最も適切なものはどれか？", o:["有料版に比べて回答の質が著しく低くなる","社員が業務上の判断をAIに依存しすぎてしまう","入力内容がモデル改善に使われる可能性がある","混雑する時間帯には接続速度が大きく低下してしまう"]},
  {id:"e34", lesson:"lesson1", q:"ChatGPTの無料版（Plus未加入）でデータ学習をオフにしていない場合、入力した情報はどうなるか？", o:["OpenAI社によって企業秘密として保護される","OpenAIのモデル改善に使われる可能性がある","日本の個人情報保護法によって自動的に保護される","送信後ただちにサーバーから完全に削除される"]},
  {id:"e35", lesson:"lesson16", q:"不動産業でのAI情報漏洩事例として実際に起きやすいシナリオはどれか？", o:["AIが法務局の登記情報を書き換えてしまう","顧客の査定情報をAIに入力し外部に漏洩","AIが所有者に無断で物件を売却してしまう","AIが市場の不動産価格を操作してしまう"]},
  {id:"e36", lesson:"lesson10", q:"文化庁のガイドラインで「著作権侵害になりうる」とされているAIの使い方はどれか？", o:["AIが既存の著作物に似た表現を大量生成して市場を代替する","AIに社内向けの企画のアイデアを複数出させて検討する","AIに自分で書いた文章の誤字脱字をまとめて修正させる","AIに社内会議で使う資料の要約を作らせて共有する"]},
  {id:"e37", lesson:"lesson25", q:"「ベクトルデータベース」がRAGシステムで使われる目的として正しいものはどれか？", o:["利用者のパスワードを暗号化して管理する","社内で扱う画像ファイルをまとめて保存する","社内文書のファイルを暗号化して保護しておく","文章の意味的な類似度で関連情報を高速検索する"]},
  {id:"e38", lesson:"lesson22", q:"インシデント対応ロールプレイで学ぶ最重要の教訓はどれか？", o:["そもそもAIを業務に使わなければ事故は起きなかった","法律の専門家に任せておけば問題はすべて解決する","十分な手順さえ整えておけばインシデントは必ず防ぐことができる","事前に対応手順を決めておくことで、有事の初動が速くなる"]},
  {id:"e39", lesson:"lesson8", q:"介護事業者がAIを活用する際のリスクとして最も注意すべきものはどれか？", o:["利用者の健康情報・介護記録の漏洩","AIが感情を持ち自律的に動くリスク","介護サービスの品質が低下すること","スタッフがAIに過度に依存すること"]},
  {id:"e40", lesson:"lesson20", q:"音声・映像が本物かどうか確認するための対策として最も有効なものはどれか？", o:["画質や音質が十分に高ければ本物と判断する","AI判定ツールの結果だけで真偽を最終的に判断する","送信者のメールアドレスの表記だけを確認して判断する","別の通信手段（電話・直接確認）で本人確認を取る"]},
  {id:"e41", lesson:"lesson16", q:"士業（弁護士・税理士等）のAI利用で特に注意すべき点はどれか？", o:["AI利用にかかる費用が事務所の経営を圧迫するほど大きすぎる","AIが士業の独占業務をそのまま代替してしまうリスク","依頼人の機密情報をAIに入力することによる守秘義務違反","AI自身が国家資格を取得して独立した業務を行うようになるリスク"]},
  {id:"e42", lesson:"lesson5", q:"ChatGPT Businessプランの主な特徴として正しいものはどれか？", o:["日本語には対応していない","データ学習がデフォルトでオフ","追加の費用なく無料で利用できる","個人での利用だけに限定されている"]},
  {id:"e43", lesson:"lesson13", q:"AI起因の情報漏洩が疑われる場合、最初の1時間に最優先でやるべきことはどれか？", o:["漏洩が起きた原因を詳しく調べて特定する","報道機関へ配布するプレスリリースの作成","経営層および役員への状況の報告","該当AIツールの使用停止と証拠保全"]},
  {id:"e44", lesson:"lesson20", q:"ディープフェイク対策として企業が導入すべき「コードワード」制度とはどれか？", o:["重要な指示の際に事前に決めた合言葉で本人確認を行う仕組み","全社員のパスワードを定期的に複雑なものへ変更させる社内ルール","ビデオ会議の通信内容をすべて暗号化して守る仕組み","AIとの会話内容を暗号化して第三者に読ませない仕組み"]},
  {id:"e45", lesson:"lesson2", q:"「学習オフ」設定をしてもなお残るリスクとして正しいものはどれか？", o:["AIが回答を生成する際に社内システムに接触する","設定が一定期間で自動的に元に戻る","入力データがサーバーに送信される","会話履歴が一時的にサーバーに残る"]},
  {id:"e46", lesson:"lesson11", q:"AIのハルシネーションによる法的リスクとして最も適切なものはどれか？", o:["誤った法的情報に基づいて行動し損害が発生する","AIが業務の途中で接続を拒否して停止してしまう","利用にかかるコストが想定より大きく増加してしまう","AIが過剰に正直になり不都合な事実まで述べてしまう"]},
  {id:"e47", lesson:"lesson9", q:"個人情報保護法において「要配慮個人情報」に該当するものはどれか？", o:["メールアドレス","病歴・診療情報","生年月日と性別","氏名と勤務先名"]},
  {id:"e48", lesson:"lesson13", q:"個人情報漏洩が発生した場合、個人情報保護委員会への報告義務が生じる条件として正しいものはどれか？", o:["要配慮個人情報が含まれるなど一定の条件を満たす場合","社内で漏洩に気づいた段階で直ちに報告するべきとされる場合","報道機関が漏洩の事実を外部へ報道してしまっていた場合","漏洩が1件でも発生していれば例外なく該当することになる"]},
  {id:"e49", lesson:"lesson3", q:"シャドーAIが特に危険な理由として最も適切なものはどれか？", o:["利用コストが想定よりかかりすぎる","利用しているAIの性能が低い","会社が情報漏洩を把握できない","社員のAIスキルが上がりすぎてしまう"]},
  {id:"e50", lesson:"lesson18", q:"AIを使ったフィッシング詐欺の特徴として正しいものはどれか？", o:["送信元のメールアドレスが必ず海外のものになる","文章が不自然なため受信者が容易に見破ることができる","自然な日本語で個人に最適化された偽メールを大量生成できる","実行形式の添付ファイルが必ず含まれているため判別しやすい"]},
  {id:"e51", lesson:"lesson18", q:"AIフィッシングを見破るための最も効果的な方法はどれか？", o:["本文の日本語が自然かどうかだけを見て真偽を判断する","そのメールにそのまま返信して本物かどうかを直接尋ねる","添付ファイルを実際に開いて中身を確認したうえで本物かどうか判断する","URLの確認・送信元ドメインの確認・公式サイトへの直接アクセス"]},
  {id:"e52", lesson:"lesson15", q:"「今日からできる10のアクション」のうち、最も即効性が高い対策はどれか？", o:["新しいAIシステムの導入","社外に向けた情報の公開","AIツールの全面的な撤廃","社内のAI利用実態調査"]},
  {id:"e53", lesson:"lesson17", q:"Gemini Notebook（旧 NotebookLM）が他の生成AIツールと大きく異なる特徴はどれか？", o:["アップロードしたドキュメントだけを情報源として回答する","音声による入力だけですべての操作ができる点","法人契約をした企業だけが利用できる点","回答を作るたびに常にインターネット検索を行って情報を集める点"]},
  {id:"e54", lesson:"lesson10", q:"AIで生成した文章の著作権について正しい説明はどれか？", o:["生成された時点で著作権は一切発生しないとされる","人間の創作的寄与がある場合に利用者に帰属しうる","常にAIを開発した事業者に自動的に帰属することになる","入力の内容にかかわらず常に利用者に帰属する"]},
  {id:"e55", lesson:"lesson3", q:"「シャドーAI」とは何か？", o:["会社が公式に導入しているAIツールの総称","会社の許可なく従業員が業務に使うAI","AIが人の指示なく自律的に行動している状態","サイバー攻撃に悪用されているAIのこと"]},
  {id:"e56", lesson:"intro", q:"生成AIが日本のビジネス現場に急速に普及した主な契機はどれか？", o:["2022年11月のChatGPT公開","製造現場への産業用AIロボットの導入","政府によるAI推進のための新しい法律の施行","スマートフォンの世界的な急速な普及"]},
  {id:"e57", lesson:"lesson17", q:"GensparkやManusなどの「AIエージェント型」ツールのリスクとして正しいものはどれか？", o:["生成される回答が業務で使うには短すぎてしまう","自律的に行動するため予期しないデータアクセスが発生しうる","操作画面も生成される回答も日本語に対応していない","月あたりの利用料が他のツールと比べて著しく高い"]},
  {id:"e58", lesson:"lesson23", q:"LLM（大規模言語モデル）がハルシネーションを起こす根本的な理由はどれか？", o:["日本語の文法や語彙の扱いがそもそも苦手にできているから","確率的なパターン予測で文章を生成するため、事実確認をしていないから","文章を生成するために必要な計算能力が不足しているから","常時インターネットに接続して最新の情報を取得していないから"]},
  {id:"e59", lesson:"lesson12", q:"社内AI利用ガイドラインに必ず含めるべき項目として最も重要なものはどれか？", o:["AI技術のこれまでの発展の歴史","入力禁止情報の具体的なリスト","同業他社における導入事例の紹介","主要なAIツールの利用料金の一覧"]},
  {id:"e60", lesson:"lesson23", q:"トランスフォーマーアーキテクチャの「Attention機構」の役割はどれか？", o:["入力された文章を画像のデータに変換して出力する","回答が返ってくるまでの処理速度を引き上げる","文章中の単語間の関係性・重要度を計算する","AIを外部のインターネットへ常時接続させる仕組み"]}
];

/* ===== 受講実績レポート ===== */


function repMsg(type, text){
  const el = document.getElementById('rep-msg');
  if(!el) return;
  el.style.display = text ? 'block' : 'none';
  el.className = 'form-msg' + (type ? ' ' + type : '');
  el.textContent = text || '';
}

function fmtDateTime(iso){
  if(!iso) return '—';
  const d = new Date(iso);
  if(isNaN(d)) return '—';
  const p = n => String(n).padStart(2,'0');
  return `${d.getFullYear()}/${p(d.getMonth()+1)}/${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}
function fmtDuration(sec){
  if(sec === null || sec === undefined) return '—';
  const h = Math.floor(sec/3600), m = Math.floor((sec%3600)/60), s = sec%60;
  if(h) return `${h}時間${m}分`;
  if(m) return `${m}分${s}秒`;
  return `${s}秒`;
}

// 受講記録をサーバーから取得し、レポート用に整形する受講時間の計測を作り直した時刻。
// これより前の duration_seconds は旧方式（開いた時刻→完了を押した時刻）で記録されており、実態の3.8倍に膨らんだ例がある。
// 数値として信用できないためレポートには出さない。DBの値は書き換えず、表示だけ伏せる。
const MEASURE_FIXED_AT = Date.parse('2026-08-30T19:00:00+09:00');
function reportableDuration(log){
  const d = (log.duration_seconds === undefined || log.duration_seconds === null) ? null : log.duration_seconds;
  if(d === null) return null;
  const c = log.completed_at ? Date.parse(log.completed_at) : null;
  if(c !== null && c < MEASURE_FIXED_AT) return null;   // 旧方式の記録は伏せる
  return d;
}

async function fetchRecordData(){
  const data = await apiFetch('/api/records', { method: 'GET' });
  const meta = PLAN_LABELS_JA[data.plan] || PLAN_LABELS_JA.basic;
  const planLessons = getPlanLessons();

  const logMap = {};
  (data.lessons || []).forEach(l => { logMap[l.lesson_id] = l; });
  const quizMap = {};
  (data.quiz || []).forEach(q => { quizMap[q.lesson_id] = q; });

  const rows = planLessons.map(id => {
    const log = logMap[id] || {};
    const q   = quizMap[id];
    return {
      id,
      title: (LESSON_LABELS[id] || id),
      startedAt:   log.started_at   || null,
      completedAt: log.completed_at || null,
      duration:    reportableDuration(log),
      answered:    !!q,
      selected:    q ? q.selected + 1 : null,       // 表示は1始まり
      correctNo:   null,   // 正答番号はサーバーが返さない（2026-11-22。読めば再回答で正解できたため）
      isCorrect:   q ? q.is_correct : null,
      answeredAt:  q ? q.answered_at : null,
    };
  });

  // 章末ワーク（chapter_works）を章ごとにまとめる
  const chTotal = (data.summary && data.summary.chaptersTotal) || 0;
  const workMap = {};
  (data.works || []).forEach(w => {
    if(!workMap[w.chapter]) workMap[w.chapter] = [];
    workMap[w.chapter].push(w);
  });
  const works = [];
  for(let c=1; c<=chTotal; c++){
    const items = (workMap[c] || []).sort((a,b)=> (a.question_id||'').localeCompare(b.question_id||''));
    const info = (typeof WORK_CHAPTERS !== 'undefined' && WORK_CHAPTERS[c]) || {};
    works.push({
      chapter: c,
      name: info.name || '',
      goal: info.goal || '',
      // 4問すべて提出した章だけを「提出済み」とする
      submitted: ((data.summary && data.summary.workDoneList) || []).includes(c),
      submittedAt: (items.find(i => i.submitted_at) || {}).submitted_at || null,
      // duration_seconds は章内の全行に同じ値が入る仕様なので、合計せず最大を取る
      duration: items.reduce((a,i) => Math.max(a, i.duration_seconds || 0), 0) || null,
      answers: items.map(i => ({
        questionId: i.question_id,
        question: ((typeof WORK_QUESTIONS !== 'undefined' && WORK_QUESTIONS[i.question_id]) || {}).q || '',
        text: i.answer_text || '',
        length: (i.answer_text || '').trim().length,
        min: ((typeof WORK_QUESTIONS !== 'undefined' && WORK_QUESTIONS[i.question_id]) || {}).min || null,
        status: i.status
      }))
    });
  }

  // 章末確認テスト（chapter_tests）を章ごとにまとめる
  const testMap = {};
  (data.tests || []).forEach(t => {
    if(!testMap[t.chapter]) testMap[t.chapter] = [];
    testMap[t.chapter].push(t);
  });
  const tests = [];
  for(let c=1; c<=chTotal; c++){
    const items = (testMap[c] || []).sort((a,b)=> (a.question_id||'').localeCompare(b.question_id||''));
    const correct = items.filter(i => i.is_correct).length;
    const fItems = (data.testsFirst || []).filter(t => t.chapter === c);
    const fCorrect = fItems.filter(i => i.is_correct).length;
    const maxAtt = Math.max(1, ...items.map(i => i.attempt_no || 1));
    tests.push({
      chapter: c,
      answered: items.length,
      total: 5,
      correct,
      firstCorrect: fCorrect,
      firstAnswered: fItems.length,
      attempts: maxAtt,
      passed: items.length >= 5 && correct >= 4,
      items: items.map(i => ({
        questionId: i.question_id,
        selected: i.selected + 1,
        correctNo: null,
        isCorrect: i.is_correct,
        answeredAt: i.answered_at
      }))
    });
  }

  const dates = rows.map(r => r.completedAt).filter(Boolean).sort();
  return {
    meta, rows, works, tests,
    plan: data.plan,
    email: data.email,
    summary: data.summary,
    firstDate: dates[0] || null,
    lastDate: dates[dates.length-1] || null,
  };
}

async function openRecordReport(){
  const name    = (document.getElementById('rep-name')?.value || '').trim();
  const company = (document.getElementById('rep-company')?.value || '').trim();
  if(!name){ repMsg('error','氏名を入力してください。'); return; }

  repMsg('', '');
  // ポップアップブロック対策：先にウィンドウを開いてから中身を書き込む
  const win = window.open('', '_blank');
  if(!win){ repMsg('error','ポップアップがブロックされました。許可してから再度お試しください。'); return; }
  win.document.write('<p style="font-family:sans-serif;padding:24px">読み込み中…</p>');

  try {
    const d = await fetchRecordData();
    const s = d.summary;

    const trs = d.rows.map(r => `
      <tr>
        <td class="c">${escHtml(r.title)}</td>
        <td>${fmtDateTime(r.startedAt)}</td>
        <td>${fmtDateTime(r.completedAt)}</td>
        <td class="n">${fmtDuration(r.duration)}</td>
        <td class="n">${r.answered ? '選択肢' + r.selected : '未回答'}</td>
        <td class="n">${r.answered ? (r.isCorrect ? '正解' : '不正解') : '—'}</td>
      </tr>`).join('');

    const rate = s.quizAnswered ? Math.round(s.quizCorrect / s.quizAnswered * 1000)/10 : 0;

    // 集計は Worker の totalSeconds を使わず、レポートに出せる値だけで積み直す。
    // Worker 側は旧方式の記録も合算してしまうため。
    const lessonSecs = d.rows.reduce((a,r) => a + (r.duration || 0), 0);
    const workSecs   = (d.works || []).reduce((a,w) => a + (w.duration || 0), 0);
    const hiddenN    = d.rows.filter(r => r.completedAt && r.duration === null).length;
    const measureNote = hiddenN > 0
      ? `<p class="mnote">※ ${hiddenN}件のレッスンは、2026年8月30日の計測方式の改修より前に受講されたものです。当時の記録方式では受講時間を正確に計測できていなかったため、受講時間の表示を差し控えています（受講の事実・完了日時・理解度チェックの結果は記録のとおりです）。</p>`
      : '';
    const today = fmtDateTime(new Date().toISOString()).split(' ')[0];

    // 章末ワーク（記述式）の出力
    const chTotal = s.chaptersTotal || 0;
    const chDone  = s.chaptersSubmitted || 0;
    // 章末確認テストの出力
    let testsHTML = '';
    if(chTotal > 0 && (d.tests||[]).some(t => t.answered > 0)){
      const trows = (d.tests||[]).map(t => `
        <tr>
          <td>第${t.chapter}章</td>
          <td class="n">${t.answered} / ${t.total} 問</td>
          <td class="n">${t.firstAnswered ? t.firstCorrect + ' 問' : '—'}</td>
          <td class="n">${t.correct} 問</td>
          <td class="n">${t.answered ? Math.round(t.correct/t.answered*1000)/10 : 0}%</td>
          <td class="n">${t.attempts > 1 ? t.attempts + ' 回' : '1 回'}</td>
          <td class="n">${t.answered >= t.total ? (t.passed ? '合格' : '不合格') : '未完了'}</td>
        </tr>`).join('');
      const tAll = (d.tests||[]).reduce((a,t)=>a+t.answered,0);
      const cAll = (d.tests||[]).reduce((a,t)=>a+t.correct,0);
      const fAll = (d.tests||[]).reduce((a,t)=>a+(t.firstCorrect||0),0);
      testsHTML = `
<h2>章末確認テストの結果</h2>
<table>
  <thead><tr><th>章</th><th>回答数</th><th>初回正答</th><th>最終正答</th><th>正答率</th><th>受験回数</th><th>判定</th></tr></thead>
  <tbody>${trows}</tbody>
  <tfoot><tr>
    <th>合計</th><th>${tAll} / ${chTotal*5} 問</th><th>${fAll} 問</th><th>${cAll} 問</th>
    <th>${tAll ? Math.round(cAll/tAll*1000)/10 : 0}%</th><th>—</th><th>—</th>
  </tr></tfoot>
</table>
<p class="wnote">各章5問・4肢択一。合格ラインは4問以上（80%）。判定は当社が設定した正答に基づき講座システムが自動で行っています。<br>
「初回正答」は最初に回答した時点の正答数、「最終正答」は学び直しを含めた最新の回答による正答数です。理解度の把握には初回正答を、修了判定には最終正答を用いています。</p>`;
    }

    let worksHTML = '';
    if(chTotal > 0){
      const blocks = (d.works || []).map(w => {
        const qs = w.answers.map(a => `
          <div class="wq">
            <div class="wq-h">${escHtml(a.questionId)}　<span class="wq-n">${a.length}文字${a.min?`（必要 ${a.min}文字以上）`:''}</span></div>
            <div class="wq-q">${escHtml(a.question)}</div>
            <div class="wq-a">${escHtml(a.text) || '（未記入）'}</div>
          </div>`).join('');
        return `
        <div class="wch">
          <h3>第${w.chapter}章　${escHtml(w.goal)}
            <span class="wst ${w.submitted?'ok':''}">${w.submitted ? '提出済み' + (w.submittedAt ? '（' + fmtDateTime(w.submittedAt) + '）' : '') : '未提出'}</span>
            ${w.duration ? `<span class="wst">所要 ${fmtDuration(w.duration)}</span>` : ''}
          </h3>
          ${qs || '<div class="wq-a">（未記入）</div>'}
        </div>`;
      }).join('');
      worksHTML = `
<h2>章末ワーク（記述式課題）の提出内容</h2>
<p class="wnote">各章の学習内容を受講者自身の業務・組織に適用して記述する課題です。提出状況：<strong>${chDone} / ${chTotal} 章</strong></p>
${blocks}`;
    }

    win.document.open();
    win.document.write(`<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">
<title>受講実績レポート_${escHtml(name)}</title>
<style>
  *{box-sizing:border-box}
  body{font-family:"Hiragino Sans","Yu Gothic",sans-serif;color:#1a1a1a;margin:0;padding:32px;line-height:1.7}
  .wrap{max-width:900px;margin:0 auto}
  h1{font-size:22px;margin:0 0 4px;border-bottom:3px solid #c0392b;padding-bottom:10px}
  .sub{font-size:12px;color:#666;margin-bottom:24px}
  table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:24px}
  th,td{border:1px solid #ccc;padding:7px 9px;text-align:left;vertical-align:top}
  th{background:#f4f4f4;font-weight:700;white-space:nowrap}
  td.n{white-space:nowrap}
  td.c{min-width:200px}
  .info th{width:150px;background:#fafafa}
  .sum{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:24px}
  .sum div{flex:1;min-width:130px;border:1px solid #ccc;border-radius:6px;padding:12px}
  .sum b{display:block;font-size:19px;margin-top:4px}
  .lbl{font-size:11px;color:#666}
  .foot{font-size:11px;color:#666;border-top:1px solid #ccc;padding-top:14px;margin-top:8px}
  h2{font-size:14px;margin:24px 0 10px;padding-left:8px;border-left:4px solid #c0392b}
  @media print{ body{padding:0} .noprint{display:none} table{font-size:10.5px} }
  .btn{background:#c0392b;color:#fff;border:0;padding:10px 20px;border-radius:6px;font-size:14px;cursor:pointer;margin-bottom:20px}
  .mnote{font-size:12px;color:#555;line-height:1.7;margin:10px 0 0;padding:10px 12px;background:#f6f6f6;border-left:3px solid #bbb}
  .wnote{font-size:12px;color:#555;margin:0 0 14px}
  tfoot th{background:#f4f4f4;text-align:left;border:1px solid #ccc;padding:7px 9px}
  .wch{border:1px solid #ccc;border-radius:6px;padding:14px 16px;margin-bottom:14px;page-break-inside:avoid}
  .wch h3{font-size:13px;margin:0 0 12px;padding-bottom:8px;border-bottom:1px solid #e0e0e0;display:flex;justify-content:space-between;align-items:baseline;gap:10px}
  .wst{font-size:11px;font-weight:400;color:#999;white-space:nowrap}
  .wst.ok{color:#1a7f4b;font-weight:700}
  .wq{margin-bottom:14px}
  .wq:last-child{margin-bottom:0}
  .wq-h{font-size:11px;font-weight:700;color:#c0392b;margin-bottom:3px}
  .wq-n{font-weight:400;color:#888}
  .wq-q{font-size:11px;color:#666;line-height:1.6;margin-bottom:6px}
  .wq-a{font-size:12px;line-height:1.85;white-space:pre-wrap;background:#fafafa;border-left:3px solid #ddd;padding:9px 12px}


</style></head><body><div class="wrap">
<button class="btn noprint" data-print>印刷 / PDFとして保存</button>
<h1>受講実績レポート</h1>
<div class="sub">AI Security Lab 生成AI×セキュリティ eラーニング</div>

<h2>受講者情報</h2>
<table class="info">
  <tr><th>氏名</th><td>${escHtml(name)}</td></tr>
  ${company ? `<tr><th>会社名</th><td>${escHtml(company)}</td></tr>` : ''}
  <tr><th>メールアドレス</th><td>${escHtml(d.email || '')}</td></tr>
  <tr><th>講座名</th><td>AI Security Lab 生成AI×セキュリティ eラーニング（${escHtml(d.meta.name)}）</td></tr>
  <tr><th>受講範囲</th><td>${escHtml(d.meta.range)}</td></tr>
  <tr><th>標準学習時間</th><td>${escHtml(d.meta.hours)}（休憩時間を除く）</td></tr>
  <tr><th>受講期間</th><td>${fmtDateTime(d.firstDate)} 〜 ${fmtDateTime(d.lastDate)}</td></tr>
</table>

<h2>集計</h2>
<div class="sum">
  <div><span class="lbl">受講済みレッスン</span><b>${s.lessonsCompleted} / ${s.lessonsTotal}</b></div>
  <div><span class="lbl">レッスン学習時間（実測）</span><b>${lessonSecs ? fmtDuration(lessonSecs) : '—'}</b></div>
  ${chTotal ? `<div><span class="lbl">章末ワーク所要時間（実測）</span><b>${workSecs ? fmtDuration(workSecs) : '—'}</b></div>` : ''}
  <div><span class="lbl">理解度テスト回答数</span><b>${s.quizAnswered} 問</b></div>
  <div><span class="lbl">正答数 / 正答率</span><b>${s.quizCorrect} 問 / ${rate}%</b></div>
  ${chTotal ? `<div><span class="lbl">章末ワーク提出</span><b>${chDone} / ${chTotal} 章</b></div>` : ''}
  ${s.testTotal ? `<div><span class="lbl">章末確認テスト</span><b>${s.testCorrect||0} / ${s.testTotal} 問</b></div>` : ''}
  <div><span class="lbl">修了状況</span><b>${s.completed ? '修了' : '受講中'}</b></div>
</div>
${measureNote}

<h2>レッスン別 受講履歴および理解度テスト結果</h2>
<table>
  <thead><tr>
    <th>レッスン</th><th>開始日時</th><th>完了日時</th><th>学習時間</th><th>回答</th><th>評価結果</th>
  </tr></thead>
  <tbody>${trs}</tbody>
</table>

${testsHTML}
${worksHTML}

<div class="foot">
  ・理解度テストは各レッスンに1問（4肢択一）を設けています。表中の「回答」は<strong>初回回答</strong>を記載しています。<br>
  ・正誤の判定は、当社が設定した正答に基づき講座システムが自動で行っています（受講者による自己採点ではありません）。<br>
  ・学習時間は受講画面の滞在時間の実測値です。標準学習時間とは一致しません。<br>
  ・「—」は該当データがないことを示します。本機能の導入前に受講されたレッスンは開始日時・学習時間が記録されていない場合があります。<br>
  ・章末ワークは受講者が自社の状況を記述する課題です。提出時に設問ごとの必要文字数を満たしているかを講座システムが自動で判定しています。本レポートには最新の提出内容と提出日時を記載しています。<br>
  ・章末確認テストは各章5問（4肢択一）、合格ラインは4問以上です。不合格の場合は該当レッスンを確認のうえ再挑戦でき、「初回正答」は最初に回答した時点の、「最終正答」は学び直しを含めた最新の回答による正答数です。<br><br>
  発行日：${today}／発行者：合同会社LaughRally（AI Security Lab）
</div>
</div></body></html>`);
    win.document.close();
    // 別ウィンドウも元のページの CSP を引き継ぐので、onclick は動かない。ボタンはここで登録する
    win.document.querySelectorAll('[data-print]').forEach(b => b.addEventListener('click', () => win.print()));
    repMsg('success','レポートを別ウィンドウで開きました。');
  } catch(e) {
    win.close();
    repMsg('error','レポートを取得できませんでした。' + (e.message ? '（' + e.message + '）' : ''));
  }
}

async function downloadRecordCSV(){
  const name    = (document.getElementById('rep-name')?.value || '').trim();
  const company = (document.getElementById('rep-company')?.value || '').trim();
  if(!name){ repMsg('error','氏名を入力してください。'); return; }
  repMsg('', '');

  try {
    const d = await fetchRecordData();
    // CSVの数式インジェクション対策：= + - @ で始まる値はExcelが数式として解釈するため
    // 先頭にアポストロフィを付けて無害化する（表示上は元の文字列のまま）
    const q = csvCell;
    const lines = [
      ['氏名','会社名','メールアドレス','講座名','プラン','受講範囲','標準学習時間'].map(q).join(','),
      [name, company, d.email || '', 'AI Security Lab 生成AI×セキュリティ eラーニング',
       d.meta.name, d.meta.range, d.meta.hours].map(q).join(','),
      '',
      ['レッスン番号','レッスン名','開始日時','完了日時','学習時間(秒)','選択した選択肢','評価結果','回答日時'].map(q).join(','),
    ];
    d.rows.forEach(r => {
      lines.push([
        'L' + lessonNo(r.id), r.title,
        r.startedAt ? fmtDateTime(r.startedAt) : '',
        r.completedAt ? fmtDateTime(r.completedAt) : '',
        r.duration == null ? '' : r.duration,
        r.answered ? r.selected : '',
        r.answered ? (r.isCorrect ? '正解' : '不正解') : '未回答',
        r.answeredAt ? fmtDateTime(r.answeredAt) : '',
      ].map(q).join(','));
    });

    // 章末確認テスト
    if((d.tests || []).some(t => t.answered > 0)){
      lines.push('', ['章','設問ID','選択した選択肢','評価結果','回答日時'].map(q).join(','));
      d.tests.forEach(t => {
        t.items.forEach(i => {
          lines.push(['第'+t.chapter+'章', i.questionId, i.selected,
                      i.isCorrect ? '正解' : '不正解',
                      i.answeredAt ? fmtDateTime(i.answeredAt) : ''].map(q).join(','));
        });
      });
    }

    // 章末ワーク（記述式）
    if((d.works || []).length){
      lines.push('', ['章','章末ワーク','設問ID','提出状況','提出日時','文字数','必要文字数','回答内容'].map(q).join(','));
      d.works.forEach(w => {
        if(!w.answers.length){
          lines.push(['第'+w.chapter+'章', w.goal, '', '未提出', '', '', '', ''].map(q).join(','));
          return;
        }
        w.answers.forEach(a => {
          lines.push([
            '第'+w.chapter+'章', w.goal, a.questionId,
            a.status === 'submitted' ? '提出済み' : '下書き',
            w.submittedAt ? fmtDateTime(w.submittedAt) : '',
            a.length, a.min == null ? '' : a.min,
            a.text
          ].map(q).join(','));
        });
      });
    }

    const s = d.summary;
    lines.push('', ['集計','受講済み','累計学習時間(秒)','理解度チェック回答数','同 正答数','章末確認テスト正答','章末ワーク提出','修了状況'].map(q).join(','));
    lines.push(['', `${s.lessonsCompleted}/${s.lessonsTotal}`, d.rows.reduce((a,r)=>a+(r.duration||0),0),
                s.quizAnswered, s.quizCorrect,
                s.testTotal ? `${s.testCorrect||0}/${s.testTotal}` : '対象外',
                s.chaptersTotal ? `${s.chaptersSubmitted||0}/${s.chaptersTotal}` : '対象外',
                s.completed ? '修了' : '受講中'].map(q).join(','));

    // Excelで文字化けしないよう BOM を付ける
    const blob = new Blob(['\ufeff' + lines.join('\r\n')], { type:'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `受講実績_${name}_${new Date().toISOString().slice(0,10)}.csv`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(()=>URL.revokeObjectURL(a.href), 1000);
    repMsg('success','CSVをダウンロードしました。');
  } catch(e) {
    repMsg('error','CSVを作成できませんでした。' + (e.message ? '（' + e.message + '）' : ''));
  }
}

/* ===== 進捗ダッシュボード ===== */

// 章別レッスン定義

function updateDashboard(){
  const completed = loadCompleted();
  const dashboard = document.getElementById('progress-dashboard');
  if(!currentUser){ if(dashboard) dashboard.style.display='none'; return; }
  if(dashboard) dashboard.style.display='block';

  const bar   = document.getElementById('dash-bar');
  const count = document.getElementById('dash-count');
  const planLessons = getPlanLessons();
  const total = planLessons.length;
  const done  = planLessons.filter(id => completed.includes(id)).length;
  const pct   = total ? done / total * 100 : 0;
  if(bar)   bar.style.width = pct + '%';
  if(count) count.textContent = done + ' / ' + total;

  // レポート用の氏名・会社名は修了証フォームの入力を引き継ぐ
  const repName = document.getElementById('rep-name');
  const repComp = document.getElementById('rep-company');
  if(repName && !repName.value){
    repName.value = (document.getElementById('cert-name')?.value || '').trim();
  }
  if(repComp && !repComp.value){
    repComp.value = (document.getElementById('cert-company')?.value || '').trim();
  }

  // 章別カード生成
  const chapEl = document.getElementById('dash-chapters');
  if(!chapEl) return;
  chapEl.innerHTML = CHAPTER_DATA.map(ch => {
    const doneLessons = ch.lessons.filter(id => completed.includes(id));
    const chPct = ch.lessons.length > 0 ? doneLessons.length / ch.lessons.length * 100 : 0;
    const planBadge = ch.plan ? `<span style="font-size:10px;background:${ch.color};color:#fff;padding:1px 7px;border-radius:3px;margin-left:8px">${ch.plan}</span>` : '';
    const lessonsHtml = ch.lessons.map(id => {
      const done = completed.includes(id);
      const lbl = LESSON_LABELS[id] || id;
      const short = lbl.replace(/^L\d+：/,'').replace(/^序章：/,'');
      return `<div class="dash-lesson-item${done?' done':''}" data-tab="${id}" onclick="switchTab('${id}');scrollToId('lessons')">
        <div class="dlcheck">${done?'✓':''}</div>
        <span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1" title="${lbl}">${short}</span>
      </div>`;
    }).join('');
    return `<div class="dash-chapter-card">
      <div class="dash-ch-head">
        <div class="dash-ch-title"><span style="font-family:'IBM Plex Mono',monospace;font-size:10px;color:${ch.color};margin-right:8px">${ch.tag}</span>${ch.name}${planBadge}</div>
        <div class="dash-ch-count">${doneLessons.length} / ${ch.lessons.length}</div>
      </div>
      <div class="dash-ch-bar"><div class="dash-ch-bar-fill" style="width:${chPct}%;background:${ch.color}"></div></div>
      <div class="dash-lessons-grid">${lessonsHtml}</div>
    </div>`;
  }).join('');

  // サイドバー更新
  if(currentUser) buildSideNav(completed);
}
window.updateDashboard = updateDashboard;

// ========== 学習UI（PEP風） ==========

// 各章の最終レッスンの直後に章末ワークを挟んだ順序を作る
const ALL_TABS_ORDERED = (function(){
  const out = ['intro'];
  ALL_LESSONS_LIST.forEach(id => {
    if(id !== 'intro') out.push(id);
    if(CH_END[id]) out.push('work' + CH_END[id]);
  });
  out.push('exam', 'certificate', 'records');
  return out;
})();

// ---- サイドバータブ切替 ----
function setSbTab(tab, btn){
  document.querySelectorAll('.sb-tab').forEach(t => t.classList.remove('active'));
  // ボタンを渡さないときは、同じ種類のタブ（先頭が目次・2つ目が進捗）を選んだ状態にする
  if(!btn) btn = document.querySelectorAll('.sb-tab')[tab === 'lessons' ? 0 : 1];
  if(btn) btn.classList.add('active');
  const list = document.getElementById('sb-lesson-list');
  const panel = document.getElementById('sb-progress-panel');
  if(tab === 'lessons'){
    if(list) list.style.display = '';
    if(panel){ panel.style.display = 'none'; panel.classList.remove('show'); }
  } else {
    if(list) list.style.display = 'none';
    if(panel){ panel.style.display = ''; panel.classList.add('show'); buildProgressPanel(); }
  }
}

/* ===== 序章のポッドキャスト（自前プレーヤー） =====
 * Google Drive の埋め込みは別ドメインのiframeで見た目を一切制御できず、
 * 再生開始まで5〜10秒かかっていた。音声を audio/intro.mp3 として同梱し、
 * CSSに用意済みだった .pod-* のデザインで自前のプレーヤーを組む。
 */
const POD_KEY = 'asl_pod_intro';   // 再生位置の保存先
function podAudio(){ return document.getElementById('pod-audio'); }

function podFmt(sec){
  if(!isFinite(sec) || sec < 0) return '--:--';
  const s = Math.floor(sec);
  return Math.floor(s/60) + ':' + String(s%60).padStart(2,'0');
}

function podToggle(){
  const a = podAudio();
  if(!a) return;
  if(a.paused){
    a.play().catch(e => {
      const t = document.getElementById('pod-cur');
      if(t) t.textContent = '再生できません';
      console.warn('音声を再生できませんでした:', e.message);
    });
  } else {
    a.pause();
  }
}
function podSeek(v){
  const a = podAudio();
  if(a && isFinite(a.duration)) a.currentTime = a.duration * (v / 1000);
}
function podVol(v){
  const a = podAudio();
  if(!a) return;
  a.volume = Number(v);
  a.muted = Number(v) === 0;
  podSyncMute();
}
function podMute(){
  const a = podAudio();
  if(!a) return;
  a.muted = !a.muted;
  podSyncMute();
}
function podSyncMute(){
  const a = podAudio(), btn = document.getElementById('pod-mute');
  if(!a || !btn) return;
  btn.textContent = (a.muted || a.volume === 0) ? '🔇' : '🔊';
}
function podSpeed(rate, btn){
  const a = podAudio();
  if(a) a.playbackRate = rate;
  document.querySelectorAll('.pod-spd').forEach(b => b.classList.remove('active'));
  if(btn) btn.classList.add('active');
}

function initPodcast(){
  const a = podAudio();
  if(!a || a.dataset.wired === '1') return;
  a.dataset.wired = '1';
  const seek = document.getElementById('pod-seek');
  const cur  = document.getElementById('pod-cur');
  const dur  = document.getElementById('pod-dur');
  const play = document.getElementById('pod-play');

  const onMeta = () => {
    if(dur) dur.textContent = podFmt(a.duration);
    // 前回の続きから（最後まで聴き終えていた場合は先頭に戻す）
    try{
      const saved = parseFloat(store.get(POD_KEY, '0'));
      if(saved > 0 && isFinite(a.duration) && saved < a.duration - 5) a.currentTime = saved;
    }catch(e){ /* localStorageが使えない環境は無視 */ }
  };
  a.addEventListener('loadedmetadata', onMeta);
  // preload="metadata" のため、この関数が動く時点で既に読み込み済みのことがある。
  // その場合 loadedmetadata は二度と発火せず、長さが '--:--' のままになる。
  if(a.readyState >= 1) onMeta();
  a.addEventListener('timeupdate', () => {
    if(cur) cur.textContent = podFmt(a.currentTime);
    if(seek && isFinite(a.duration) && a.duration > 0){
      seek.value = String(Math.round(a.currentTime / a.duration * 1000));
    }
    // 5秒ごとに保存する。毎回書くとtimeupdateの頻度ぶん無駄に書き込む。
    if(Math.floor(a.currentTime) % 5 === 0){
      store.set(POD_KEY, String(a.currentTime));
    }
  });
  a.addEventListener('play',  () => { if(play){ play.textContent = '❚❚'; play.setAttribute('aria-label','一時停止'); } });
  a.addEventListener('pause', () => { if(play){ play.textContent = '▶';  play.setAttribute('aria-label','再生'); } });
  a.addEventListener('ended', () => {
    if(play) play.textContent = '▶';
    if(seek) seek.value = '0';
    store.remove(POD_KEY);
  });
  a.addEventListener('error', () => {
    if(cur) cur.textContent = '読み込み失敗';
    console.warn('音声ファイルを読み込めませんでした: audio/intro.mp3');
  });
  podSyncMute();
}

/**
 * ヘッダーのロゴ・カリキュラム・レッスンの遷移先。
 *
 * 未ログイン時は今までどおり LP のアンカーへ飛ばす（return true）。
 * ログイン後は #learn-wrapper（position:fixed・画面全体・z-index 50）が
 * LPを覆っているため、ページをスクロールさせても見た目が一切変わらない。
 * さらに #top と #lessons は updateAuthUI / showLearnUI が display:none に
 * するので、アンカーとしても機能しない。
 * そこでログイン中は学習UI側の操作へ振り替える。
 */
function navGo(where){
  if(!currentUser) return true;   // 未ログインは通常のアンカー動作

  if(where === 'curriculum'){
    // サイドバーの「📖 学習」タブ＝レッスン一覧がカリキュラムそのもの
    const lessonsTab = document.querySelector('#sb-tabs .sb-tab:nth-child(1)');
    if(lessonsTab) setSbTab('lessons', lessonsTab);
    if(window.innerWidth <= 860) openSbDrawer();
    showLearnUI();
    return false;
  }

  if(where === 'top'){
    // ロゴ＝学習の先頭（序章）へ戻す
    window.switchTab('intro');
    return false;
  }

  // 'lessons'：今開いているレッスンの先頭まで戻す
  showLearnUI();
  const lc = document.getElementById('learn-content');
  if(lc) lc.scrollTo({ top: 0, behavior: 'smooth' });
  if(window.innerWidth <= 860) closeSbDrawer();
  return false;
}

// 進捗ボタン → サイドバーの進捗タブを開く
function openProgressPanel(){
  // PC: サイドバーの進捗タブをアクティブに
  const progressTab = document.querySelector('#sb-tabs .sb-tab:nth-child(2)');
  if(progressTab) setSbTab('progress', progressTab);
  // スマホ: ドロワーを開く
  if(window.innerWidth <= 860){
    openSbDrawer();
  }
  // learn-wrapperが表示されていなければ表示
  showLearnUI();
}

// モバイル：ドロワー開閉
function toggleSbDrawer(){
  const sb = document.getElementById('learn-sidebar');
  const ov = document.getElementById('sb-overlay');
  if(!sb) return;
  const isOpen = sb.classList.contains('open');
  if(isOpen){ closeSbDrawer(); } else { openSbDrawer(); }
}
function openSbDrawer(){
  const sb = document.getElementById('learn-sidebar');
  const ov = document.getElementById('sb-overlay');
  if(sb) sb.classList.add('open');
  if(ov) ov.classList.add('show');
}
function closeSbDrawer(){
  const sb = document.getElementById('learn-sidebar');
  const ov = document.getElementById('sb-overlay');
  if(sb) sb.classList.remove('open');
  if(ov) ov.classList.remove('show');
}

// カリキュラムからレッスンへ飛ぶ共通関数
function goLesson(tabId){
  if(!canAccessTab(tabId)) return;
  if(currentUser){
    // ログイン後：先にUIを表示してからレッスンをマウント
    showLearnUI();
    window.switchTab(tabId);   // 前回のタブの記録・章末ワークの読み込みも同じ経路で行う
    window.scrollTo({top: 0, behavior: 'smooth'});
  } else {
    showTabGuest(tabId);
    const s = document.getElementById('lessons');
    if(s) s.scrollIntoView({behavior:'smooth'});
  }
}

// ---- 学習画面 表示/非表示 ----
function showLearnUI(){
  const w = document.getElementById('learn-wrapper');
  if(w) w.classList.add('show');
  // トップページ用に畳んだ序章を、学習画面では全文に戻す
  document.querySelectorAll('.lb.lb-collapsed').forEach(el => el.classList.remove('lb-collapsed'));
  document.querySelectorAll('.lb-more').forEach(el => el.remove());
  // タブバー非表示
  const tw = document.getElementById('tabs-wrap');
  if(tw) tw.classList.add('hidden');
  // lessonsセクションを非表示
  const sec = document.getElementById('lessons');
  if(sec) sec.style.display = 'none';
  document.body.classList.add('logged-in');

  // FABをモバイルのみ表示
  const fabEl = document.getElementById('sb-fab');
  if(fabEl) fabEl.style.display = window.innerWidth <= 860 ? 'flex' : 'none';

  // 全tconをlearn-content-innerに移動（初回のみ）
  const inner = document.getElementById('learn-content-inner');
  if(inner && inner.dataset.moved !== '1'){
    inner.dataset.moved = '1';
    document.querySelectorAll('.tcon').forEach(t => {
      t.style.display = 'none';
      inner.appendChild(t);
    });
  }
}
function hideLearnUI(){
  const w = document.getElementById('learn-wrapper');
  if(w) w.classList.remove('show');
  const tw = document.getElementById('tabs-wrap');
  if(tw) tw.classList.remove('hidden');
  const sec = document.getElementById('lessons');
  if(sec) sec.style.display = '';
  document.body.classList.remove('logged-in');
  // FAB非表示
  const fabEl2 = document.getElementById('sb-fab');
  if(fabEl2) fabEl2.style.display = 'none';
}

/* ===== 有料レッスン本文のサーバー配信 =====
 * 無料公開（序章・L01）以外の本文は index.html に含めない。
 * data-remote="1" が付いた器だけを置き、開いたときに
 * GET /api/lesson/:id から取得して流し込む。
 * Worker 側でログインとプランを検証しているため、
 * 未ログイン・プラン外のユーザーには本文が一切渡らない。
 * （序章と L01 は無料公開なので従来どおり HTML に直接入っている）
 */
const LESSON_PLACEHOLDER =
  '<div class="lc"><div class="lb"><p class="lesson-loading">レッスンを読み込んでいます…</p></div></div>';

// 同じレッスンを連打されても取得は1回だけにする
const _lessonLoading = {};

async function ensureLessonLoaded(tabId){
  const tcon = document.getElementById('tab-' + tabId);
  if(!tcon) return;
  if(tcon.dataset.remote !== '1') return;   // 本文が最初から入っているタブ
  if(tcon.dataset.loaded === '1') return;
  if(_lessonLoading[tabId]) return _lessonLoading[tabId];

  const p = (async () => {
    try {
      const j = await apiFetch('/api/lesson/' + tabId);
      if(!j || typeof j.html !== 'string' || !j.html) throw new Error('本文が空です');
      tcon.innerHTML = j.html;
      tcon.dataset.loaded = '1';
      rememberServed(tabId, j.served);
      // 流し込んだ後に、選択肢シャッフルと受講済み／クイズロックの状態を復元する
      shuffleQuizOptions(tabId);
      updateCompletionUI();
      initQuizLocks();
      fixLessonNav(tabId);
    } catch(e) {
      const msg = document.createElement('span');
      msg.textContent = e && e.message ? e.message : '';
      tcon.innerHTML =
        '<div class="lc"><div class="lb">'
        + '<p style="color:var(--red);font-weight:600">レッスンを読み込めませんでした。</p>'
        + '<p style="opacity:.75;font-size:14px">' + msg.innerHTML + '</p>'
        + '<p><button class="btn-complete" onclick="reloadLesson(\'' + tabId + '\')">もう一度読み込む</button></p>'
        + '</div></div>';
    } finally {
      delete _lessonLoading[tabId];
    }
  })();
  _lessonLoading[tabId] = p;
  return p;
}
window.ensureLessonLoaded = ensureLessonLoaded;

function reloadLesson(tabId){
  const tcon = document.getElementById('tab-' + tabId);
  if(!tcon) return;
  tcon.dataset.loaded = '';
  tcon.innerHTML = LESSON_PLACEHOLDER;
  ensureLessonLoaded(tabId);
}

// ---- アクティブtconをlearn-content-innerに表示 ----
// 初めて学習画面に入った人への案内。プランごとの進め方と修了の条件、最初のレッスンへのボタン。
// 1本でも受講済みにするか、閉じるボタンを押すと出なくなる
function renderOnboard(){
  const box = document.getElementById('learn-onboard');
  if(!box || !currentUser || !currentUserPlan) return;
  const key = 'asl_onboard_closed:' + currentUser.id;
  if(store.get(key) || loadCompleted().length > 0){ box.hidden = true; box.innerHTML = ''; return; }
  if(!box.hidden && box.innerHTML) return;
  const plan = currentUserPlan;
  const paid = plan === 'basic' || plan === 'standard' || plan === 'premium';
  const chs = COURSE.workChapters[plan] || 0;
  const steps = [];
  if(paid){
    const toc = window.innerWidth <= 860 ? '左下の ≡ から開く目次' : '左の目次';
    steps.push(`${toc}の順にレッスンを読み、最後の理解度チェック（各設問${COURSE.rules.quizMaxAttempts}回まで）に答えて「受講済み」にします（全${PLAN_LESSON_COUNT[plan]}レッスン）。`);
    if(chs) steps.push(`各章の終わりに、章末ワーク（記述${COURSE.rules.workQuestions}問）を提出し、章末確認テスト（${COURSE.rules.testQuestions}問中${COURSE.rules.testPass}問以上で合格）を受けます（全${chs}章）。`);
    if(plan === 'premium') steps.push(`全章を終えたら、認定試験（${COURSE.rules.examQuestions}問・${COURSE.rules.examPass}問以上で合格、${COURSE.rules.examMinutes}分）を受けます。`);
    steps.push(plan === 'basic' ? '全レッスンを終えたら、受講実績レポートを出力できます（修了証は Standard 以上のプランで発行できます）。' : '条件をすべて満たすと、目次の下の「修了証の申請」から申請できます。');
  } else {
    steps.push('序章と L01 は無料でお読みいただけます。');
    steps.push('L02 以降は、ご購入後に受講できます。');
  }
  const first = planTabsOrdered().find(id => ALL_LESSONS_LIST.includes(id)) || 'intro';
  const label = PLAN_LABELS_JA[plan] ? PLAN_LABELS_JA[plan].name : '無料登録';
  box.innerHTML = `<div class="ob-card" role="region" aria-label="最初の一歩">
    <button type="button" class="ob-close" aria-label="案内を閉じる" onclick="closeOnboard()">×</button>
    <div class="ob-title">最初の一歩（${escHtml(label)}）</div>
    <ol class="ob-steps">${steps.map(t => '<li>' + escHtml(t) + '</li>').join('')}</ol>
    <div class="ob-actions">
      <button type="button" class="form-submit ob-btn" onclick="closeOnboard();switchTab('${first}')">${first === 'intro' ? '序章' : 'L' + lessonNo(first) + ' ' + escHtml(COURSE.titles[first] || '')} から始める →</button>
      ${paid ? '' : '<button type="button" class="work-nav-btn ob-btn2" onclick="showPlans()">料金プランを見る</button>'}
    </div></div>`;
  box.hidden = false;
}
function closeOnboard(){
  if(currentUser) store.set('asl_onboard_closed:' + currentUser.id, '1');
  const box = document.getElementById('learn-onboard');
  if(box){ box.hidden = true; box.innerHTML = ''; }
}

// 理解度チェックの答え合わせを読み上げで伝える
function markLiveRegions(root){
  (root || document).querySelectorAll('.qfb:not([role])').forEach(el => { el.setAttribute('role', 'status'); el.setAttribute('aria-live', 'polite'); });
}
function showActiveLesson(tabId){
  // 全tconをhidden
  document.querySelectorAll('.tcon').forEach(t => {
    t.classList.remove('active');
    t.style.display = 'none';
  });

  // 対象tconだけ表示
  const tcon = document.getElementById('tab-' + tabId);
  if(!tcon) return;
  tcon.style.display = 'block';
  tcon.classList.add('active');
  markLessonStart(tabId);
  Promise.resolve(ensureLessonLoaded(tabId)).then(() => markLiveRegions(tcon));
  markLiveRegions(tcon);
  shuffleQuizOptions(tabId);
  // 本文が最初から入っている序章・L01 も、前後ナビをプランに合わせる（無料登録の人に購入の案内を出すため）
  if(tcon.dataset.remote !== '1') fixLessonNav(tabId);
}

// ---- 前後ナビ ----
// プランで開けるタブだけをたどる。
// Basic は章末ワーク（開けない）で、Standard は第7章（開けない）でアラートが出て止まった。
function tabOpenable(id){
  if(id === 'records') return false;                       // 流れの外（進捗パネルから開く）
  return tabDenyReason(id) === null;
}
function planTabsOrdered(){ return ALL_TABS_ORDERED.filter(tabOpenable); }
window.tabOpenable = tabOpenable;
function tabLabel(id){
  if(id === 'certificate') return '🎉 修了証を申請する';
  if(id === 'exam') return '🏆 認定試験を受験する';
  if(/^work\d$/.test(id)) return 'W' + id.slice(4) + '：章末ワーク・確認テスト';
  if(id === 'intro') return '序章';
  return LESSON_LABELS[id] || id;
}
function lessonNavFrom(fromId, dir){
  const order = planTabsOrdered();
  let idx = order.indexOf(fromId);
  if(idx < 0){
    // 開けないタブにいる場合は、全体の並びでいちばん近い開けるタブを探す
    const all = ALL_TABS_ORDERED, at = all.indexOf(fromId);
    for(let i = at + dir; i >= 0 && i < all.length; i += dir){ if(tabOpenable(all[i])) { window.switchTab(all[i]); return; } }
    return;
  }
  const next = order[idx + dir];
  if(next) window.switchTab(next);
}
function lessonNav(dir){
  const inner = document.getElementById('learn-content-inner');
  const activeTcon = inner ? inner.querySelector('.tcon.active') : document.querySelector('.tcon.active');
  const currentId = activeTcon?.id?.replace('tab-','') || 'intro';
  lessonNavFrom(currentId, dir);
}

// レッスン本文の末尾にある前後ナビを、受講者のプランに合わせて作り直す。本文（DB）のナビは章末ワークを飛ばして次のレッスンを指しているため
function fixLessonNav(tabId){
  const tcon = document.getElementById('tab-' + tabId);
  const nav = tcon ? tcon.querySelector('.lnav') : null;
  if(!nav) return;
  const order = planTabsOrdered();
  const idx = order.indexOf(tabId);
  const prev = idx > 0 ? order[idx - 1] : null;
  const next = idx >= 0 ? order[idx + 1] : null;
  let html = '';
  if(prev) html += `<a href="#lessons" onclick="lessonNavFrom('${tabId}',-1);return false;" class="lnbtn"><span class="lnl">← PREV</span><span class="lnt">${escHtml(tabLabel(prev))}</span></a>`;
  if(next) html += `<a href="#lessons" onclick="lessonNavFrom('${tabId}',1);return false;" class="lnbtn" style="text-align:right"><span class="lnl">NEXT →</span><span class="lnt">${escHtml(tabLabel(next))}</span></a>`;
  // 範囲の終わりでは、上のプランの案内を開けるようにする
  else if(currentUserPlan === 'basic') html += `<a href="#lessons" data-tab="records" onclick="switchTab('records');return false;" class="lnbtn" style="text-align:right"><span class="lnl">🎉 Basic の全レッスンはここまで</span><span class="lnt">受講実績レポートを出力する →</span></a>`
    + `<a href="#lessons" onclick="showPlans();return false;" class="lnbtn" style="text-align:right"><span class="lnl">続けて学ぶ</span><span class="lnt">Standard 以上で第6章と修了証へ →</span></a>`;
  else if(currentUserPlan !== 'standard' && currentUserPlan !== 'premium') html += `<a href="#lessons" onclick="showPlans();return false;" class="lnbtn" style="text-align:right"><span class="lnl">無料公開はここまで</span><span class="lnt">L02 以降はご購入後に受講できます →</span></a>`;
  nav.innerHTML = html;
}
window.fixLessonNav = fixLessonNav;

// ---- サイドバーのレッスンリスト生成 ----
// 目次の W 行：章末ワークを開く（確認テストへの位置合わせは解除する）
function openWorkTab(ch){ window._aslTestFocus = 0; window.switchTab('work' + ch); }
function buildSideNav(completed){
  if(!currentUser){ hideLearnUI(); return; }
  showLearnUI();

  // 現在のタブ
  const inner = document.getElementById('learn-content-inner');
  const activeTcon = inner ? inner.querySelector('.tcon.active') : document.querySelector('.tcon.active');
  let currentTab = activeTcon?.id?.replace('tab-','') || 'intro';
  // ログイン直後の最初の表示だけ、前回開いていたタブから再開する
  if(!window._aslResumed && currentUserPlan){
    window._aslResumed = true;
    try {
      const last = store.getJSON('asl_last_tab', null);
      if(last && last.u === currentUser.id && last.t && last.t !== currentTab && tabOpenable(last.t)){
        currentTab = last.t;
        setTimeout(() => window.switchTab(last.t), 0);   // 中身の読み込み（ワーク・テスト・修了証）を通常の経路で行う
      } else if(!last && completed.length){
        // 前回の記録が端末に無い（別の端末・ブラウザ）ときは、受講順で最初の未了から開く受講記録（ワーク・テストの状況）を受け取ってから決める。
        // 待つあいだに本人が別の画面を開いていたら、動かさない
        const startTab = currentTab;
        Promise.resolve(window._recordsPromise).catch(() => {}).then(() => {
          const now = (document.querySelector('#learn-content-inner .tcon.active') || {}).id || '';
          if(now && now !== 'tab-' + startTab) return;
          const nx = nextStep();
          if(nx) nx.run();
        });
      }
    } catch(e){}
  }

  // レッスン表示
  showActiveLesson(currentTab);

  // フッター進捗
  const fill = document.getElementById('sb-footer-fill');
  const cnt = document.getElementById('sb-footer-count');
  const planLessons = getPlanLessons();
  const total = planLessons.length;
  const done  = planLessons.filter(id => completed.includes(id)).length;
  const pct = total ? done / total * 100 : 0;
  if(fill) fill.style.width = pct + '%';
  if(cnt) cnt.textContent = done + ' / ' + total;
  // 修了条件のうちレッスン以外（ワーク・テスト・試験）も一目でわかるように
  const sub = document.getElementById('sb-footer-sub');
  if(sub){
    const sm = window._recSummary, pl = currentUserPlan || '';
    if(pl === 'free') sub.textContent = '受講プランなし（担当者用・無料登録）';
    else if(sm && sm.chaptersTotal){
      sub.textContent = '章末ワーク ' + (sm.chaptersSubmitted || 0) + '/' + sm.chaptersTotal + '章 ・ 確認テスト ' + (sm.testPassedChapters || 0) + '/' + sm.chaptersTotal + '章'
        + (pl === 'premium' ? ' ・ 認定試験 ' + ((sm.examPassed || currentExamPassed) ? '合格' : '未') : '');
    } else sub.textContent = '';
  }

  // レッスンリスト
  const list = document.getElementById('sb-lesson-list');
  if(!list) return;
  let html = '';
  CHAPTER_DATA.forEach((ch, idx) => {
    const doneCnt = ch.lessons.filter(id => completed.includes(id)).length;
    const inPlan = ch.lessons.some(id => getPlanLessons().includes(id));
    html += `<div class="sb-ch-label">${ch.tag} · ${ch.name}${inPlan && ch.tag !== '序章' ? `<span class="sb-ch-cnt">${doneCnt}/${ch.lessons.length}</span>` : ''}</div>`;
    ch.lessons.forEach(id => {
      const done = completed.includes(id);
      const active = id === currentTab;
      const raw = LESSON_LABELS[id] || id;
      const title = raw.replace(/^L\d+[：:]/,'').replace(/^序章[：:]/,'').replace(/^【序章】/,'');
      const num = id === 'intro' ? '00' : lessonNo(id);
      html += `<div class="sb-row${active?' active':''}${done?' done':''}" tabindex="0" role="button" data-tab="${id}" onclick="switchTab('${id}');closeSbDrawer()">
        <div class="sb-row-dot"></div>
        <span class="sb-row-num">${num}</span>
        <span class="sb-row-title">${title}</span>
      </div>`;
    });
    // 章末ワーク行（序章の擬似章には付けない）
    const chNo = parseInt((ch.tag||'').replace(/[^0-9]/g,''), 10);
    // Basic には章末ワーク・テストが無いので行を出さない
    if(chNo >= 1 && WORK_CHAPTERS[chNo] && (COURSE.workChapters[currentUserPlan] || 0) > 0){
      const wid = 'work' + chNo;
      const wactive = wid === currentTab;
      const wdone = !!_workSubmittedCh[chNo];
      html += `<div class="sb-row sb-row-work${wactive && window._aslTestFocus !== chNo?' active':''}${wdone?' done':''}" tabindex="0" role="button" data-tab="work${chNo}" onclick="openWorkTab(${chNo});closeSbDrawer()">
        <div class="sb-row-dot"></div>
        <span class="sb-row-num">W${chNo}</span>
        <span class="sb-row-title">章末ワーク：${WORK_CHAPTERS[chNo].goal}</span>
      </div>`;
      // 確認テストを別の行に。
      const tpass = !!(_testResultCh[chNo] && _testResultCh[chNo].passed);
      // 5問すべて答えて不合格なら「再挑戦」の印（未受験と見分けられるように）
      const tfail = !tpass && !!(_testResultCh[chNo] && _testResultCh[chNo].answered >= TEST_TOTAL_PER_CHAPTER);
      const tactive = wactive && window._aslTestFocus === chNo;
      html += `<div class="sb-row sb-row-work${tactive?' active':''}${tpass?' done':''}${tfail?' test-failed':''}" tabindex="0" role="button" data-tab="work${chNo}" onclick="goChapterTest(${chNo});closeSbDrawer()">
        <div class="sb-row-dot"></div>
        <span class="sb-row-num">T${chNo}</span>
        <span class="sb-row-title">章末確認テスト（5問）</span>
      </div>`;
    }
  });
  list.innerHTML = html;

  // Premiumのみ：認定試験リンクをサイドバーに追加
  if(currentUserPlan === 'premium'){
    const examRow = list.querySelector('[data-exam]');
    if(!examRow){
      const div = document.createElement('div');
      div.setAttribute('data-exam','1');
      div.innerHTML = '<div class="sb-ch-label">認定試験</div><div class="sb-row" data-tab="exam" onclick="switchTab(&#39;exam&#39;);closeSbDrawer()" style="color:var(--purple);font-weight:600"><div class="sb-row-dot" style="background:var(--purple)"></div><span class="sb-row-num">EX</span><span class="sb-row-title">AI Security Lab 認定試験</span></div>';
      list.appendChild(div);
    }
  }

  // 修了証の申請（Standard 以上）。サイドバーの最後に常に出す
  if((currentUserPlan === 'standard' || currentUserPlan === 'premium') && !list.querySelector('[data-cert]')){
    const div = document.createElement('div');
    div.setAttribute('data-cert','1');
    const cactive = currentTab === 'certificate';
    div.innerHTML = '<div class="sb-ch-label">修了</div><div class="sb-row' + (cactive ? ' active' : '') + '" data-tab="certificate" onclick="switchTab(&#39;certificate&#39;);closeSbDrawer()" style="font-weight:600"><div class="sb-row-dot"></div><span class="sb-row-num">🎓</span><span class="sb-row-title">修了証の申請</span></div>';
    list.appendChild(div);
  }

  // プランロック適用
  applyPlanLock();
  renderOnboard();
  try{ applyAdminVisibility(); }catch(e){}

  // アクティブ行をスクロール表示
  setTimeout(()=>{
    const activeRow = list.querySelector('.sb-row.active');
    if(activeRow) activeRow.scrollIntoView({block:'nearest', behavior:'smooth'});
  }, 50);

  // パンくず・ボタン更新
  updateLearnTopBar(currentTab);
}

// ---- 上部バー更新 ----
function updateLearnTopBar(tabId){
  const bc = document.getElementById('learn-breadcrumb');
  if(bc){
    const ch = CHAPTER_DATA.find(c => c.lessons.includes(tabId));
    const raw = LESSON_LABELS[tabId] || '';
    const title = raw.replace(/^L\d+[：:]/,'').replace(/^序章[：:]/,'');
    const num = tabId === 'intro' ? '00' : lessonNo(tabId);
    if(ch){
      bc.innerHTML = `<span class="bc-ch" style="color:${ch.color}">${ch.tag}</span><span class="bc-sep">/</span><span>L${num}</span><span class="bc-sep">—</span><span style="color:var(--text2)">${title}</span>`;
    } else {
      // 章末ワーク・修了証・試験・記録のページでは、前のレッスンの表示を残さない
      const other = { certificate:'修了証の申請', exam:'認定試験', records:'受講実績レポート' };
      const t = /^work\d$/.test(tabId) ? ('第' + tabId.slice(4) + '章 章末ワーク・確認テスト') : (other[tabId] || '');
      bc.textContent = t;
    }
  }
  const allTabs = planTabsOrdered();
  const idx = allTabs.indexOf(tabId);
  const prevBtn = document.getElementById('learn-prev-btn');
  const nextBtn = document.getElementById('learn-next-btn');
  if(prevBtn) prevBtn.disabled = idx <= 0;
  if(nextBtn) nextBtn.disabled = idx >= allTabs.length - 1;
}

// ---- 進捗パネル ----
// 進捗パネルの章ごとの「章末ワーク・確認テスト」の状態
function wtLine(ch){
  const no = parseInt((ch.tag || '').replace(/[^0-9]/g, ''), 10);
  if(!no || !WORK_CHAPTERS[no] || no > (COURSE.workChapters[currentUserPlan] || 0)) return '';
  const sm = window._recSummary || {};
  const w = !!_workSubmittedCh[no] || (sm.workDoneList || []).includes(no);
  const t = !!(_testResultCh[no] && _testResultCh[no].passed) || (sm.testPassedList || []).includes(no);
  return `<div class="sb-dash-wt"><button type="button" class="sb-dash-wt-item${w ? ' done' : ''}" data-tab="work${no}" onclick="switchTab('work${no}')">章末ワーク ${w ? '✓ 提出済み' : '未提出'}</button>`
    + `<button type="button" class="sb-dash-wt-item${t ? ' done' : ''}" data-tab="work${no}" onclick="goChapterTest(${no})">確認テスト ${t ? '✓ 合格' : '未合格'}</button></div>`;
}
function buildProgressPanel(){
  const panel = document.getElementById('sb-progress-panel');
  if(!panel) return;
  const completed = loadCompleted();
  let html = '';
  CHAPTER_DATA.forEach(ch => {
    // プランの外の章は、鍵付きで1行だけ出す
    if(ch.tag !== '序章' && !ch.lessons.some(id => getPlanLessons().includes(id))){
      html += `<div class="sb-dash-ch sb-dash-locked"><div class="sb-dash-ch-head"><div><span class="sb-dash-ch-tag">${ch.tag}</span>${ch.name}</div>
        <button type="button" class="sb-dash-wt-item" onclick="showPlans()">🔒 ${escHtml(ch.plan || '上位プラン')}</button></div></div>`;
      return;
    }
    const done = ch.lessons.filter(id => completed.includes(id));
    const pct = ch.lessons.length > 0 ? done.length / ch.lessons.length * 100 : 0;
    const dots = ch.lessons.map(id => {
      const d = completed.includes(id);
      const num = id === 'intro' ? '00' : lessonNo(id);
      return `<div class="sb-dash-dot${d?' done':''}" tabindex="0" role="button" data-tab="${id}" onclick="switchTab('${id}');setSbTab('lessons')" title="${LESSON_LABELS[id]||id}">${num}</div>`;
    }).join('');
    html += `<div class="sb-dash-ch">
      <div class="sb-dash-ch-head">
        <div><span class="sb-dash-ch-tag">${ch.tag}</span>${ch.name}</div>
        <div class="sb-dash-ch-cnt">${done.length}/${ch.lessons.length}</div>
      </div>
      <div class="sb-dash-bar"><div class="sb-dash-bar-fill" style="width:${pct}%;background:${ch.color}"></div></div>
      <div class="sb-dash-items">${dots}</div>
      ${wtLine(ch)}
    </div>`;
  });
  // 修了証への導線（Standard 以上）。
  if(currentUserPlan === 'standard' || currentUserPlan === 'premium'){
    html += `<div class="sb-ch-label" style="margin-top:14px">修了</div>
    <div class="sb-row" data-tab="certificate" onclick="switchTab('certificate');closeSbDrawer()">
      <div class="sb-row-dot"></div>
      <span class="sb-row-num">🎓</span>
      <span class="sb-row-title">修了証の申請</span>
    </div>`;
  }
  // 法人の担当者：受講者の進捗
  if(_corpApps && _corpApps.length){
    html += `<div class="sb-ch-label" style="margin-top:14px">法人のお申し込み</div>
    <div class="sb-row" tabindex="0" role="button" onclick="openCorpProgress()">
      <div class="sb-row-dot"></div>
      <span class="sb-row-num">👥</span>
      <span class="sb-row-title">受講者の進捗（担当者用）</span>
    </div>`;
  }
  // 受講記録（実績レポート）への導線
  html += `<div class="sb-ch-label" style="margin-top:14px">記録の出力</div>
    <div class="sb-row" data-tab="records" onclick="switchTab('records');closeSbDrawer()">
      <div class="sb-row-dot"></div>
      <span class="sb-row-num">📄</span>
      <span class="sb-row-title">受講実績レポート</span>
    </div>`;
  panel.innerHTML = html;
  applyPlanLock();   // 行のキーボード操作・読み上げ用の属性を付ける
}

// 旧関数互換
function buildSideNavCompat(c){ buildSideNav(c); }
function updateBreadcrumb(tabId){ updateLearnTopBar(tabId); }


/* ===== クイズ未正解時は受講済みボタンをロック ===== */
// { 設問ID: 正解した元の選択肢番号 }。共用PCで前の人の正解が混ざらないよう、ユーザーごとに分ける
function quizDoneKey(){ return 'asl_quiz_done:' + ((currentUser && currentUser.id) || 'anon'); }

/**
 * 正解済みの設問を、開き直したあとも回答済みの見た目に戻す。
 * これが無いと、レッスンを開き直すたびに同じ設問へ再回答でき、
 * quiz_answers に余分な行が積み上がる（採点は attempt_no の初回で行うため
 * 成績には影響しないが、受講記録が読みにくくなる）。
 */
function restoreQuizLocks(){
  let done = {};
  done = store.getJSON(quizDoneKey(), {}) || {};
  Object.keys(done).forEach(id => {
    const okEl = document.getElementById(id + '-ok');
    if(!okEl) return;                      // そのレッスンがまだ読み込まれていない
    // 旧形式は数値、新形式は { s: 選択肢番号, e: 解説 }
    const rec = (done[id] && typeof done[id] === 'object') ? done[id] : { s: done[id], e: '' };
    const quiz = okEl.closest('.quiz');
    const qopts = quiz ? quiz.querySelector('.qopts') : null;
    if(!qopts || qopts.dataset.restored === '1') return;
    qopts.querySelectorAll('.qopt').forEach(b => {
      b.style.pointerEvents = 'none';
      const m = (b.getAttribute('onclick') || '').match(/doQ\(this,(\d+),/);
      if(m && Number(m[1]) === rec.s) b.classList.add('ok');
    });
    if(rec.e) okEl.innerHTML = '✅ ' + escHtml(rec.e);
    okEl.classList.add('show');
    const ngEl = document.getElementById(id + '-ng');
    if(ngEl) ngEl.classList.remove('show');
    qopts.dataset.restored = '1';
    // 正解済みなのに「受講済み」が押せなくなる不具合の修正。
    const tcon = quiz.closest('.tcon');
    const tabId = tcon ? tcon.id.replace('tab-', '') : null;
    if(tabId){
      const cb = document.getElementById('complete-' + tabId);
      if(cb) cb.classList.remove('quiz-locked');
      unlockCompleteBtn(tabId);
    }
  });
}
window.restoreQuizLocks = restoreQuizLocks;

function initQuizLocks(){
  restoreQuizLocks();
  const ALL_IDS = ALL_LESSONS_LIST;
  ALL_IDS.forEach(tabId => {
    const btn = document.getElementById('complete-' + tabId);
    if(!btn) return;
    // すでに受講済みなら何もしない
    const completed = loadCompleted();
    if(completed.includes(tabId)) return;
    // そのレッスンにクイズがあるか確認
    const tcon = document.getElementById('tab-' + tabId);
    if(tcon && tcon.querySelector('.quiz') && !tcon.querySelector('.qopts[data-restored="1"]')){
      btn.classList.add('quiz-locked');
      const note = document.getElementById('note-' + tabId);
      if(note) note.innerHTML = '理解度チェックに正解すると受講済みにできます';
    }
  });
}
window.unlockCompleteBtn = unlockCompleteBtn;

document.addEventListener('DOMContentLoaded', ()=>{ updateCompletionUI(); updateDashboard(); initQuizLocks(); initPodcast(); });
// キーボードで操作できるように：クリックで動く div にフォーカスを与え、Enter / Space で押せるようにする
document.addEventListener('DOMContentLoaded', ()=>{
  document.querySelectorAll('.lrow[onclick], .chd').forEach(el => { el.tabIndex = 0; el.setAttribute('role','button'); });
});
// トップページが長すぎる（序章の全文が LP の中にある）ので、未ログインでは途中で畳む
document.addEventListener('DOMContentLoaded', ()=>{
  const lb = document.querySelector('#tab-intro .lb');
  if(!lb || document.body.classList.contains('logged-in')) return;
  lb.classList.add('lb-collapsed');
  const more = document.createElement('div');
  more.className = 'lb-more';
  more.innerHTML = '<button type="button">序章の続きを読む（無料）</button>';
  more.querySelector('button').addEventListener('click', () => { lb.classList.remove('lb-collapsed'); more.remove(); });
  lb.after(more);
});

// ダイアログを Esc で閉じる・読み上げ用の印を付ける
document.addEventListener('keydown', (e)=>{
  if(e.key !== 'Escape') return;
  const close = [
    ['plan-pick-overlay', () => closePlanPick()],
    ['corp-progress-overlay', () => window.closeCorpProgress && window.closeCorpProgress()],
    ['auth-overlay', () => window.closeAuthModal && window.closeAuthModal()],
    ['guide-modal-overlay', () => window.closeGuideModal && window.closeGuideModal()],
    ['form-overlay', () => window.closeFormModal && window.closeFormModal()],
    ['corp-overlay', () => window.closeCorpModal && window.closeCorpModal()],
  ];
  for(const [id, fn] of close){
    const el = document.getElementById(id);
    if(el && getComputedStyle(el).display !== 'none' && !el.classList.contains('hidden')){ fn(); break; }
  }
});
document.addEventListener('DOMContentLoaded', ()=>{
  ['auth-overlay','guide-modal-overlay','form-overlay','corp-overlay'].forEach(id => {
    const el = document.getElementById(id);
    if(el){ el.setAttribute('role','dialog'); el.setAttribute('aria-modal','true'); }
  });
  // href の無いリンク（ログイン画面の切替など）を、キーボードで選べるようにする
  document.querySelectorAll('a[onclick]:not([href])').forEach(a => { a.setAttribute('href','#'); a.addEventListener('click', ev => ev.preventDefault()); });
});
document.addEventListener('keydown', (e)=>{
  if(e.key !== 'Enter' && e.key !== ' ') return;
  const el = e.target;
  if(el && el.matches && el.matches('.sb-row[role=button], .lrow[role=button], .chd[role=button], .sb-dash-dot')){
    e.preventDefault(); el.click();
  }
});
setTimeout(()=>{ updateCompletionUI(); updateDashboard(); initQuizLocks(); }, 800);



// ===== 認定試験UI =====
let examState = {
  active: false,
  answers: {},   // 出題インデックス: 「元の」選択肢番号（表示順ではない）
  questions: [], // 受験ごとにシャッフルした問題セット
  timer: null,
  remaining: COURSE.rules.examMinutes * 60,
  submitted: false,
  sending: false,  // 提出中の二重送信ガード
  startedAt: null, // 開始時刻（受験時間の実測用）
};

// Fisher-Yatesシャッフル
function shuffled(arr){
  const a = arr.slice();
  for(let i = a.length - 1; i > 0; i--){
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

/**
 * 受験ごとに問題順と選択肢順をランダム化する。
 * 正解の位置が固定されていると、内容を理解していなくても
 * パターンで正答できてしまうため、毎回並びを変える。
 *
 * ※ 正解はクライアントに存在しない。各選択肢は「元の選択肢番号」（orig）を
 *    保持しており、表示順を変えてもサーバーへ送る値は変わらない。
 *    レッスン内クイズの shuffleQuizOptions と同じ考え方。
 */
function buildExamSet(){
  return shuffled(EXAM_QUESTIONS).map(q => ({
    id: q.id,
    // 設問は内部ID（lesson10 など）で持ち、表示のときに受講順の番号（L10）にする。
    lesson: q.lesson === 'intro' ? '序章' : 'L' + lessonNo(q.lesson),
    q: q.q,
    o: shuffled(q.o.map((text, orig) => ({ text, orig }))),
  }));
}

function fmtTime(sec){
  const m = Math.floor(sec/60), s = sec%60;
  return m + ':' + String(s).padStart(2,'0');
}

async function startExam(){
  if(currentUserPlan !== 'premium'){
    alert('認定試験はPremiumプランの受講者のみ受験できます。');
    return;
  }
  // 開始前に、受験できるか（24時間の間隔）を確かめる。提出時にしか確認しないと、60分かけた解答が弾かれる
  const status = async () => {
    try { return await apiFetch('/api/exam/status', { method: 'GET' }); }
    catch(e) { alert('受験の準備ができませんでした。通信環境をご確認のうえ、もう一度お試しください。\n' + (e.message || '')); return null; }
  };
  let st = await status();
  if(!st) return;
  if(st.canStart === false && st.reason === 'cooldown'){
    alert('次の受験まで、あと約' + st.retryAfterHours + '時間お待ちください。\n\n' +
          '2回目以降の受験は、前回から24時間の間隔をいただいています。\n' +
          'これまでの受験回数：' + st.attempts + '回');
    return;
  }
  const R = COURSE.rules;
  const confirmed = confirm(`認定試験を開始しますか？\n\n制限時間: ${R.examMinutes}分\n問題数: ${R.examQuestions}問\n合格ライン: ${R.examPass}問以上（${Math.round(R.examPass / R.examQuestions * 100)}%）\n\n開始後はカウントダウンが始まります。`);
  if(!confirmed) return;
  // 開始の署名は「開始する」を押したあとに受け取る
  st = await status();
  if(!st || !st.start) { if(st) alert('受験を開始できませんでした。時間をおいてもう一度お試しください。'); return; }
  examState.start = st.start;
  if(typeof window.switchTab === 'function') window.switchTab('exam');
  examState.active = true;
  examState.answers = {};
  examState.questions = buildExamSet();  // 受験ごとに問題順・選択肢順を変える
  examState.remaining = COURSE.rules.examMinutes * 60;
  examState.submitted = false;
  examState.sending = false;
  examState.startedAt = Date.now();
  renderExam();
  startExamTimer();
  window.scrollTo({top:0, behavior:'smooth'});
}

function startExamTimer(){
  if(examState.timer) clearInterval(examState.timer);
  // 残り時間は「制限時間 −（今 − 開始時刻）」で毎回計算する。
  examState.timer = setInterval(() => {
    examState.remaining = Math.max(0, COURSE.rules.examMinutes * 60 - Math.floor((Date.now() - examState.startedAt) / 1000));
    const timerEl = document.getElementById('examTimer');
    if(timerEl) timerEl.textContent = fmtTime(examState.remaining);
    if(examState.remaining <= 0){
      clearInterval(examState.timer);
      submitExam();
    }
  }, 1000);
}

function renderExam(){
  const container = document.getElementById('examContainer');
  if(!container) return;
  let html = `
    <div style="position:sticky;top:0;background:var(--bg);padding:16px 0;border-bottom:2px solid var(--purple);z-index:10;margin-bottom:24px">
      <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px">
        <div style="font-weight:700;font-size:18px">🏆 AI Security Lab 認定試験</div>
        <div style="display:flex;gap:16px;align-items:center">
          <div style="font-weight:600;color:var(--text2)">回答済み <span id="examAnswered">${Object.keys(examState.answers).length}</span> / ${examState.questions.length}</div>
          <div style="font-weight:700;color:var(--purple)">⏰ <span id="examTimer">${fmtTime(examState.remaining)}</span></div>
          <button onclick="submitExam()" style="background:var(--red-s);color:#fff;border:none;border-radius:8px;padding:10px 20px;font-weight:700;cursor:pointer">提出する</button>
        </div>
      </div>
    </div>
  `;
  const QS = examState.questions;
  QS.forEach((q, i) => {
    html += `<div id="exq-${i}" style="margin-bottom:32px;padding:20px;background:var(--bg2);border-radius:12px;border:1px solid var(--border)">
      <div style="font-size:12px;color:var(--text3);margin-bottom:8px">問題 ${i+1} / ${QS.length}　[${q.lesson}]</div>
      <div style="font-weight:700;margin-bottom:16px;line-height:1.6">${q.q}</div>
      <div style="display:flex;flex-direction:column;gap:8px">`;
    q.o.forEach((opt, j) => {
      // value / onchange に渡すのは「元の選択肢番号」。表示順(j)ではない。
      const selected = examState.answers[i] === opt.orig;
      html += `<label style="display:flex;align-items:center;gap:12px;padding:12px;border:2px solid ${selected ? 'var(--purple)' : 'var(--border)'};border-radius:8px;cursor:pointer;background:${selected ? 'rgba(124,58,237,0.05)' : 'transparent'}">
        <input type="radio" name="q${i}" value="${opt.orig}" ${selected ? 'checked' : ''} onchange="selectAnswer(${i},${opt.orig})" style="cursor:pointer">
        <span>${opt.text}</span>
      </label>`;
    });
    html += `</div></div>`;
  });
  html += `<div style="text-align:center;margin:40px 0">
    <button onclick="submitExam()" style="background:var(--purple-s);color:#fff;border:none;border-radius:12px;padding:16px 48px;font-size:16px;font-weight:700;cursor:pointer">✅ 試験を提出する</button>
  </div>`;
  container.innerHTML = html;
  container.style.display = 'block';
  const introEl = document.getElementById('examIntro');
  if(introEl) introEl.style.display = 'none';
}

// 選んだ設問の見た目と回答数だけを更新する
function selectAnswer(qIdx, optIdx){
  examState.answers[qIdx] = optIdx;
  document.querySelectorAll(`#exq-${qIdx} label`).forEach(l => {
    const on = Number(l.querySelector('input').value) === optIdx;
    l.style.borderColor = on ? 'var(--purple)' : 'var(--border)';
    l.style.background = on ? 'rgba(124,58,237,0.05)' : 'transparent';
  });
  const cnt = document.getElementById('examAnswered');
  if(cnt) cnt.textContent = Object.keys(examState.answers).length;
}

/**
 * 試験の提出。採点はサーバー側で行う。
 * クライアントは「どの設問でどの選択肢（元の番号）を選んだか」だけを送り、
 * 得点・合否・正解・解説を受け取って表示する。
 * 合格フラグ（profiles.exam_passed）もサーバーが立てる。
 */
async function submitExam(){
  if(examState.sending) return;
  const QS = examState.questions;
  if(!QS.length) return;
  const answered = Object.keys(examState.answers).length;
  if(answered < QS.length && examState.remaining > 0){
    const unanswered = QS.length - answered;
    if(!confirm(`${unanswered}問が未回答です。このまま提出しますか？`)) return;
  }
  // 提出中はタイマーを止める。失敗したら残り時間があるぶんだけ再開する。
  if(examState.timer){ clearInterval(examState.timer); examState.timer = null; }
  examState.sending = true;

  // 設問ID → 元の選択肢番号
  const answers = {};
  QS.forEach((q, i) => {
    const v = examState.answers[i];
    if(Number.isInteger(v)) answers[q.id] = v;
  });

  const container = document.getElementById('examContainer');
  if(container) container.innerHTML = '<div style="text-align:center;padding:80px 20px;font-size:16px;color:var(--text3)">採点しています…</div>';
  window.scrollTo({top:0, behavior:'smooth'});

  try {
    const res = await apiFetch('/api/exam/submit', {
      method: 'POST',
      body: JSON.stringify({
        answers,
        start: examState.start || null,
        startedAt: examState.startedAt ? new Date(examState.startedAt).toISOString() : null,
        durationSeconds: examState.startedAt ? Math.floor((Date.now() - examState.startedAt) / 1000) : null,
      }),
    });
    examState.submitted = true;
    examState.active = false;
    renderExamResult(res);
    if(res.passed){
      currentExamPassed = true;   // 修了証の申請ボタンを即時に解放する
      refreshRecords();
      updateCompletionUI();
      updateCertificateUI();
    }
  } catch(e) {
    if(e && e.status === 409){
      // 制限時間を過ぎた（サーバーの判定）。何度提出しても受け付けられないので、そのことを伝えて受験の画面に戻す
      alert((e.message || '制限時間を過ぎたため、この解答は受け付けられませんでした。') + '\n\nもう一度受験してください。');
      clearInterval(examState.timer);
      retryExam();
    } else if(e && e.status === 429){
      // 再受験の間隔制限。開始前にも確認しているので通常ここには来ないが、
      // 来た場合は解答を保持したまま試験画面に戻す（60分の労力を捨てない）。
      alert((e.message || '次の受験まで時間をおいてください。') +
            '\n\n開始から時間がたつと、この解答は受け付けられなくなります（制限時間 ' + COURSE.rules.examMinutes + '分）。次に受験できるようになってから、あらためて受験してください。');
      renderExam();
    } else {
      alert('採点結果を受け取れませんでした。\n' + (e.message || '') + '\n\n解答は保持しています。通信環境をご確認のうえ、もう一度「提出する」を押してください。');
      renderExam();               // 解答を保持したまま試験画面に戻す
      if(examState.remaining > 0) startExamTimer();
    }
  } finally {
    examState.sending = false;
  }
  window.scrollTo({top:0, behavior:'smooth'});
}

/**
 * 結果画面。表示に使う設問文と選択肢はクライアント、
 * 正誤・正解・解説はサーバーのレスポンス（res.results）から取る。
 */
function renderExamResult(res){
  const container = document.getElementById('examContainer');
  if(!container) return;
  const QS = examState.questions;
  const resById = {};
  (res.results || []).forEach(r => { resById[r.id] = r; });

  const total   = res.total   != null ? res.total   : QS.length;
  const correct = res.correct != null ? res.correct : 0;
  const passed  = !!res.passed;
  const pct     = res.scorePct != null ? res.scorePct : Math.round(correct / (total || 1) * 100);
  const passMark = res.passMark != null ? res.passMark : 48;

  let html = `
    <div style="text-align:center;padding:40px 20px;background:${passed ? 'rgba(34,197,94,0.08)' : 'rgba(239,68,68,0.08)'};border:2px solid ${passed ? 'var(--green)' : 'var(--red)'};border-radius:16px;margin-bottom:32px">
      <div style="font-size:64px;margin-bottom:16px">${passed ? '🎉' : '😢'}</div>
      <div style="font-size:28px;font-weight:700;margin-bottom:8px;color:${passed ? 'var(--green)' : 'var(--red)'}">${passed ? '合格' : '不合格'}</div>
      <div style="font-size:20px;margin-bottom:16px">${correct} / ${total} 問正解（${pct}%）</div>
      <div style="color:var(--text3);font-size:14px">合格ライン: ${passMark}問以上（80%）${res.attemptNo ? '　／　受験回数: ' + res.attemptNo + '回目' : ''}</div>
      <div style="color:var(--text3);font-size:12px;margin-top:8px">採点は当社サーバーで行っており、受講者による自己採点ではありません。</div>
      ${passed ? `<div style="margin-top:24px"><button onclick="requestCertificate()" style="background:var(--purple-s);color:#fff;border:none;border-radius:12px;padding:14px 36px;font-size:15px;font-weight:700;cursor:pointer">📜 修了証を申請する</button></div>` : `<div style="margin-top:24px"><button onclick="retryExam()" style="background:var(--red-s);color:#fff;border:none;border-radius:12px;padding:14px 36px;font-size:15px;font-weight:700;cursor:pointer">🔄 受験の画面に戻る（再受験は ${nextExamAt()} ごろから）</button></div>`}
    </div>
  `;

  if(passed){
    // 合格者にだけ全問の解答と解説を出す。
    html += `<div style="font-weight:700;font-size:18px;margin-bottom:16px">📋 全問題の解答と解説</div>`;
    // 受験者が見た順（出題順）で並べる
    QS.forEach((q, idx) => {
      const r = resById[q.id];
      if(!r) return;
      const optText = n => {
        const hit = q.o.find(x => x.orig === n);
        return hit ? hit.text : '—';
      };
      const bg = r.isCorrect ? 'rgba(34,197,94,0.05)' : 'rgba(239,68,68,0.05)';
      const bc = r.isCorrect ? 'var(--green)' : 'var(--red)';
      html += `<div style="margin-bottom:16px;padding:16px;background:${bg};border-left:4px solid ${bc};border-radius:8px">
        <div style="font-size:12px;color:var(--text3);margin-bottom:4px">問題 ${idx+1} [${q.lesson}] ${r.isCorrect ? '✅ 正解' : '❌ 不正解'}</div>
        <div style="font-weight:700;margin-bottom:8px">${q.q}</div>
        <div style="font-size:14px;margin-bottom:4px"><strong>あなたの解答:</strong> ${r.selected !== null && r.selected !== undefined ? optText(r.selected) : '未回答'}</div>
        <div style="font-size:14px;margin-bottom:8px;color:var(--green)"><strong>正解:</strong> ${optText(r.correctIndex)}</div>
        <div style="font-size:13px;color:var(--text3);line-height:1.6"><strong>解説:</strong> ${r.explanation || ''}</div>
      </div>`;
    });
  } else {
    // 不合格のときは正解も解説も出さない。1回受ければ答えが全部手に入る状態だと2回目は必ず満点になり、修了証が受講を証明しなくなる。
    // サーバーも不合格時は correctIndex / explanation を返していない。
    // 復習の手がかりとして、誤答のあった章と問数だけを示す。
    // 設問ごとの正誤を返すと、受け直しの差分で正解が割り出せるため、サーバーも章単位でしか返さない。
    const wrongLessons = res.wrongByChapter || {};
    const keys = Object.keys(wrongLessons).sort((a, b) => (a === '序章' ? -1 : b === '序章' ? 1 : a.localeCompare(b, 'ja', {numeric:true})));
    html += `<div style="font-weight:700;font-size:18px;margin-bottom:12px">📚 復習が必要な章</div>
      <p style="color:var(--text3);font-size:14px;line-height:1.8;margin:0 0 16px">
        公平な認定のため、正解と解説、どの設問を間違えたかは表示していません。誤答のあった章を読み直してから、あらためて受験してください。
      </p>`;
    if(keys.length){
      html += `<div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:20px">` +
        // 章を押すと、その章の最初のレッスンが開く
        keys.map(L => { const no = L === '序章' ? 0 : parseInt(L.replace(/[^0-9]/g, ''), 10); const first = no === 0 ? 'intro' : ((COURSE.chapters.find(c => c.no === no) || {}).lessons || [])[0];
          return `<button type="button" data-tab="${first}" onclick="switchTab('${first}')" style="background:var(--red-l);border:1px solid var(--red);color:var(--red);border-radius:8px;padding:8px 14px;font-size:14px;font-weight:600;cursor:pointer;font-family:inherit">${L}<span style="font-weight:400">　${wrongLessons[L]}問 →</span></button>`; }).join('') +
        `</div>`;
    }
    html += `<p style="color:var(--text3);font-size:13px;line-height:1.8;margin:0">
      2回目以降の受験は、前回から24時間の間隔をいただいています（次に受験できるのは ${nextExamAt()} ごろから）。
    </p>`;
  }
  container.innerHTML = html;
}

// 次に受験できる日時（提出から24時間後。Worker の EXAM_COOLDOWN_MS と同じ）
function nextExamAt(){
  const d = new Date(Date.now() + 24 * 3600e3);
  return (d.getMonth() + 1) + '月' + d.getDate() + '日 ' + d.getHours() + ':' + String(d.getMinutes()).padStart(2, '0');
}
function retryExam(){
  examState = { active:false, answers:{}, questions:[], timer:null, remaining:COURSE.rules.examMinutes * 60, submitted:false, sending:false, startedAt:null };
  const container = document.getElementById('examContainer');
  if(container) container.style.display = 'none';
  const introEl = document.getElementById('examIntro');
  if(introEl) introEl.style.display = 'block';
  window.scrollTo({top:0, behavior:'smooth'});
}

/**
 * 修了証の発行申請。
 * 以前はGASのURLを直接叩いていたが、認証も受講状況の確認もないため
 * URLさえ知っていれば誰でも任意の氏名で発行できた。
 * 現在はWorkerが受講状況とプランをDBで検証してからGASを呼ぶ。
 */
async function submitCertificate(){
  const company = (document.getElementById('cert-company')?.value || '').trim();
  const name    = (document.getElementById('cert-name')?.value || '').trim();
  const msg     = document.getElementById('cert-msg');
  const btn     = document.getElementById('cert-submit-btn');

  if(!currentUser){ openAuthModal(); return; }
  if(!name){ alert('氏名を入力してください。'); return; }

  btn.disabled = true;
  btn.textContent = '送信中...';
  if(msg){ msg.style.display = 'none'; }

  try {
    const r = await apiFetch('/api/certificate', {
      method: 'POST',
      body: JSON.stringify({ name, company }),
    });
    if(msg){
      msg.style.display = 'block';
      msg.style.color = 'var(--green)';
      // サーバ応答をそのまま連結していたので、エスケープしてから埋める
      msg.innerHTML = '✅ 申請を受け付けました。修了証PDFを <strong>' + escHtml(r.email || 'ご登録のメールアドレス') +
                      '</strong> 宛にお送りします。<br>証明書番号：' + escHtml(r.certNo || '') ;
    }
    btn.textContent = '申請済み';
    window._certSubmitted = true;   // 画面の更新でボタンが押せる状態に戻らないように（二重申請の防止）
  } catch(e) {
    if(e.status === 409 || /発行済み/.test(e.message || '')){
      window._certSubmitted = true;
      btn.textContent = '発行済み';
      btn.style.opacity = '.4';
      if(msg){ msg.style.display = 'block'; msg.style.color = 'var(--text2)'; msg.textContent = e.message; }
      return;
    }
    btn.disabled = false;
    btn.textContent = '修了証を申請する';
    if(msg){
      msg.style.display = 'block';
      msg.style.color = 'var(--red)';
      msg.textContent = e.message || '送信に失敗しました。時間をおいて再試行してください。';
    }
  }
}

// 認定試験の合格画面から：修了証の申請欄を開き、氏名の入力欄へ移る。
function requestCertificate(){
  if(typeof window.switchTab === 'function') window.switchTab('certificate');
  setTimeout(() => {
    const el = document.getElementById('cert-name');
    if(el){ el.scrollIntoView({ block: 'center', behavior: 'smooth' }); el.focus({ preventScroll: true }); }
  }, 300);
}

window.openSbDrawer = openSbDrawer;
window.renderExam = renderExam;
window.renderExamResult = renderExamResult;


/* ===== 法人一括申込 =====
   受講者情報を受け取り、Worker 経由で受講申込確認書PDFを発行する。
   価格と人数の上限は course.json（COURSE.prices・COURSE.rules.corpMaxSeats）から。Worker の CORP_PLAN も同じ値を使う。 */
const CORP_PLANS = Object.fromEntries(Object.entries(COURSE.prices).map(([p, price]) => [p, { label: p[0].toUpperCase() + p.slice(1) + 'プラン', price }]));
const CORP_MAX_SEATS = COURSE.rules.corpMaxSeats;

function corpMsg(type, text){
  const el = document.getElementById('corp-msg');
  if(!el) return;
  el.className = 'form-msg' + (type ? ' ' + type : '');
  el.textContent = text || '';
  if(!text) el.style.display = 'none';
}

function corpAddRow(name, email){
  const wrap = document.getElementById('corp-rows');
  if(!wrap) return;
  if(wrap.children.length >= CORP_MAX_SEATS){
    corpMsg('error', `受講者は${CORP_MAX_SEATS}名まで入力できます。`);
    return;
  }
  const row = document.createElement('div');
  row.className = 'corp-row';
  const n = document.createElement('input');
  n.type = 'text'; n.placeholder = 'お名前'; n.maxLength = 50; n.value = name || '';
  const e = document.createElement('input');
  e.type = 'email'; e.placeholder = 'メールアドレス'; e.maxLength = 100; e.value = email || '';
  const d = document.createElement('button');
  d.type = 'button'; d.className = 'corp-del'; d.textContent = '✕';
  d.title = 'この行を削除';
  d.onclick = () => { row.remove(); corpRecalc(); };
  row.append(n, e, d);
  wrap.appendChild(row);
  corpRecalc();
}

function corpRecalc(){
  const wrap = document.getElementById('corp-rows');
  const seats = wrap ? wrap.children.length : 0;
  const plan = CORP_PLANS[document.getElementById('corp-plan')?.value] || CORP_PLANS.standard;
  const total = plan.price * seats;
  const line = document.getElementById('corp-calc-line');
  const tot  = document.getElementById('corp-total');
  if(line) line.textContent = `${plan.label} ¥${plan.price.toLocaleString('ja-JP')} × ${seats}名`;
  if(tot)  tot.textContent  = '¥' + total.toLocaleString('ja-JP');
  // 1行しかないときは削除ボタンを無効化（0名を防ぐ）
  if(wrap) [...wrap.querySelectorAll('.corp-del')].forEach(b => {
    b.disabled = seats <= 1;
    b.style.opacity = seats <= 1 ? '.3' : '';
    b.style.cursor  = seats <= 1 ? 'default' : 'pointer';
  });
}

function openCorpModal(){
  const ov = document.getElementById('corp-overlay');
  if(!ov) return;
  ov.classList.remove('hidden');
  const wrap = document.getElementById('corp-rows');
  if(wrap && !wrap.children.length){ corpAddRow(); corpAddRow(); corpAddRow(); }
  corpMsg('', '');
  corpRecalc();
}

function closeCorpModal(){
  document.getElementById('corp-overlay')?.classList.add('hidden');
}

async function corpSubmit(){
  const btn = document.getElementById('corp-submit');
  const val = id => (document.getElementById(id)?.value || '').trim();
  const company      = val('corp-company');
  const contactName  = val('corp-contact-name');
  const contactEmail = val('corp-contact-email');
  const plan         = val('corp-plan');
  const note         = val('corp-note');
  const website      = val('corp-website');
  const isEmail = v => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v);

  if(!company)      { corpMsg('error','会社名・団体名をご入力ください。'); return; }
  if(!contactName)  { corpMsg('error','ご担当者名をご入力ください。'); return; }
  if(!isEmail(contactEmail)){ corpMsg('error','ご担当者のメールアドレスをご確認ください。'); return; }

  const rows = [...document.querySelectorAll('#corp-rows .corp-row')];
  const participants = [];
  const seen = new Set();
  for(const r of rows){
    const inputs = r.querySelectorAll('input');
    const n = (inputs[0]?.value || '').trim();
    const e = (inputs[1]?.value || '').trim().toLowerCase();
    if(!n && !e) continue;                       // 空行は無視する
    if(!n){ corpMsg('error','受講者のお名前が未入力の行があります。'); return; }
    if(!isEmail(e)){ corpMsg('error',`受講者のメールアドレスをご確認ください：${n}`); return; }
    if(seen.has(e)){ corpMsg('error',`メールアドレスが重複しています：${e}`); return; }
    seen.add(e);
    participants.push({ name:n, email:e });
  }
  if(!participants.length){ corpMsg('error','受講者を1名以上ご入力ください。'); return; }

  btn.disabled = true;
  const original = btn.textContent;
  btn.textContent = '送信中…';
  corpMsg('', '');
  try{
    const res = await fetch(WORKER_URL + '/api/corporate/apply', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ company, contactName, contactEmail, plan, note, participants, website }),
    });
    const data = await res.json().catch(() => ({}));
    if(!res.ok){ throw new Error(data.error || 'お申し込みを受け付けられませんでした。'); }
    corpMsg('success',
      data.pdfIssued === false
        ? `お申し込みを受け付けました（申込番号 ${data.appNo}）。受講申込確認書は担当者より改めてお送りします。`
        : `お申し込みを受け付けました（申込番号 ${data.appNo}）。受講申込確認書を ${contactEmail} 宛にお送りしました。`
    );
    btn.textContent = '受付完了';
  }catch(err){
    corpMsg('error', err.message || 'お申し込みを受け付けられませんでした。時間をおいてお試しください。');
    btn.disabled = false;
    btn.textContent = original;
  }
}

// https://ai-security-lab.com/#corp で法人申込を直接開く（営業メールのリンク用）
// このモジュールは load のあとに走り終えることがあるので、済んでいればその場で開く
function openCorpFromHash(){ if (location.hash === '#corp') { try { openCorpModal(); } catch (e) {} } }
if (document.readyState === 'complete') openCorpFromHash();
else window.addEventListener('load', openCorpFromHash);


/* ===== 管理者：法人申込の一覧と請求書送信 ===== */
const CORP_STATUS_JA = {
  received:    '受付',
  invoiced:    '請求済み',
  paid:        '入金済み',
  provisioned: 'アカウント発行済み',
  cancelled:   'キャンセル',
};

function corpAdminMsg(type, text){
  const el = document.getElementById('corp-admin-msg');
  if(!el) return;
  el.style.color = type === 'error' ? 'var(--red)' : 'var(--green)';
  el.textContent = text || '';
}

async function loadCorpApplications(){
  const list = document.getElementById('corp-admin-list');
  const btn  = document.getElementById('corp-reload-btn');
  if(!list) return;
  if(btn){ btn.disabled = true; btn.textContent = '読み込み中…'; }
  corpAdminMsg('', '');
  try{
    const d = await apiFetch('/api/admin/applications', { method: 'GET' });
    const rows = d.rows || [];
    if(!rows.length){ list.textContent = '申込はまだありません。'; return; }
    const yen = n => '¥' + Number(n || 0).toLocaleString('ja-JP');
    list.innerHTML = rows.map(r => {
      const people = (r.participants || []).map(p => `${escHtml(p.name)}（${escHtml(p.email)}）`).join('<br>');
      const sent = !!r.invoice_sent_at;
      const action = sent
        ? `<span style="color:var(--green);font-weight:600">送信済み</span>${r.invoice_url ? `<br><a href="${escHtml(r.invoice_url)}" target="_blank" rel="noopener" style="font-size:12px">請求書を開く</a>` : ''}`
        : `<button onclick="sendCorpInvoice('${escHtml(r.app_no)}', this)" style="background:var(--red-s);color:#fff;border:none;border-radius:8px;padding:8px 14px;font-size:12px;font-weight:600;cursor:pointer;white-space:nowrap">請求書を送る</button>`;
      return `<div style="border:1px solid var(--border);border-radius:10px;padding:14px 16px;margin-bottom:10px;background:var(--bg2)">
        <div style="display:flex;justify-content:space-between;gap:14px;flex-wrap:wrap;align-items:flex-start">
          <div style="flex:1;min-width:240px">
            <div style="font-weight:700;color:var(--text);font-size:14px">${escHtml(r.company)}
              <span style="font-size:11px;font-weight:500;color:var(--text4);margin-left:8px">${escHtml(CORP_STATUS_JA[r.status] || r.status)}</span>
            </div>
            <div style="font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--text4);margin:3px 0 8px">${escHtml(r.app_no)}　${escHtml((r.created_at||'').slice(0,16).replace('T',' '))}</div>
            <div style="line-height:1.9">
              ご担当者：${escHtml(r.contact_name)}（${escHtml(r.contact_email)}）<br>
              ${escHtml((CORP_PLANS[r.plan] || {}).label || r.plan)}　${r.seats}名　<strong style="color:var(--text)">${yen(r.total_price)}</strong>（税込）<br>
              <span style="color:var(--text4)">受講者</span><br>${people}
              ${r.note ? `<br><span style="color:var(--text4)">備考</span><br>${escHtml(r.note)}` : ''}
            </div>
          </div>
          <div style="text-align:right">${action}</div>
        </div>
      </div>`;
    }).join('');
  }catch(e){
    corpAdminMsg('error', '申込一覧を取得できませんでした。' + (e.message ? `（${e.message}）` : ''));
    list.textContent = '';
  }finally{
    if(btn){ btn.disabled = false; btn.textContent = '🔄 申込一覧を読み込む'; }
  }
}

async function sendCorpInvoice(appNo, btn){
  const row = (await apiFetch('/api/admin/applications', { method: 'GET' })).rows.find(r => r.app_no === appNo);
  const to = row ? row.contact_email : '';
  if(!confirm(`${appNo} の請求書を確定して送信します。\n\n宛先：${to}\n金額：¥${Number(row?.total_price||0).toLocaleString('ja-JP')}（税込）\n\n送信後は取り消せません。よろしいですか？`)) return;
  if(btn){ btn.disabled = true; btn.textContent = '送信中…'; }
  corpAdminMsg('', '');
  try{
    const d = await apiFetch('/api/admin/invoice/send', {
      method: 'POST',
      body: JSON.stringify({ appNo }),
    });
    // 先に一覧を作り直す。loadCorpApplications がメッセージを消すため、
    // 完了を待ってから結果を出す。
    await loadCorpApplications();
    corpAdminMsg('success', `請求書を送信しました（${d.invoiceNumber || appNo}）。宛先：${d.to}`);
  }catch(e){
    corpAdminMsg('error', '送信できませんでした。' + (e.message ? `（${e.message}）` : ''));
    if(btn){ btn.disabled = false; btn.textContent = '請求書を送る'; }
  }
}


// ---- onclick などから呼べる関数の一覧----
// ページ・目次・レッスン本文（DB）の onclick="…" は、CSP でブラウザが実行しないため、00-core の inlineShim が
// この一覧にある名前だけを呼ぶ。ボタンを足すときは、ここに関数を足すこと（ページのテストが漏れを見つける）。
// あわせて window にも載せる（ほかのスクリプトやテストが window.名前 で呼ぶため）。
const INLINE_ACTIONS = {
  adminExportLogs,
  adminMfaEnroll,
  adminMfaVerify,
  answerTest,
  closeAuthModal,
  closeCorpModal,
  closeCorpProgress,
  closeFormModal,
  closeGuideModal: (...a) => window.closeGuideModal(...a),
  closeLegalModal,
  closeOnboard,
  closePlanPick,
  closeSbDrawer,
  corpAddRow,
  corpRecalc,
  corpSubmit,
  doQ,
  downloadCorpCsv,
  downloadGuideline,
  downloadRecordCSV,
  goChapterTest,
  goLesson,
  gotoWorkQ,
  lessonNav,
  lessonNavFrom,
  loadCorpApplications,
  markComplete,
  navGo,
  onWorkInput,
  openAuthModal,
  openCorpModal,
  openCorpProgress,
  openFormModal,
  openGuideModal: (...a) => window.openGuideModal(...a),
  openLegalModal,
  openProgressPanel,
  openRecordReport,
  openWorkTab,
  podMute,
  podSeek,
  podSpeed,
  podToggle,
  podVol,
  reloadLesson,
  reopenWork,
  requestCertificate,
  retryExam,
  retryTestChapter,
  scrollToId,
  selectAnswer,
  sendCorpInvoice,
  sendPasswordReset,
  setSbTab,
  showPlans,
  signInWithEmail,
  signOut,
  startExam,
  submitCertificate,
  submitExam,
  submitForm,
  submitWorkChapter,
  switchTab: (...a) => window.switchTab(...a),   // プランの判定つきの版（08 で定義）
  toggleAuthMode,
  toggleCh,
  toggleDark,
  toggleSbDrawer,
  toggleSubmitBtn,
  updatePassword,
};
Object.entries(INLINE_ACTIONS).forEach(([name, fn]) => { if(typeof window[name] !== 'function') window[name] = fn; });
