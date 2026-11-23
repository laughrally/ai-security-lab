// AI Security Lab — 受講画面と LP のスクリプト。
// 正本は assets/src/*.part.js（機能ごと）。asl-worker/course/course-sync.py が番号の順につないで assets/app.mjs を作る。
// app.mjs を直接直さないこと（check-all.sh が食い違いを止める）。つないだ全体で1つの ES モジュールになる。
// 冒頭の COURSE ブロックは course.json から書き出す
// このスクリプトより前から window にあるもの（inlineShim が「呼べない関数」として使う）
const WINDOW_BEFORE = new Set(Object.getOwnPropertyNames(window));
// <COURSE:BEGIN> course/course-sync.py が course.json から書き出す。手で直さない
const COURSE = {"free": ["intro", "lesson1"], "plans": {"basic": 5, "standard": 6, "premium": 7}, "workChapters": {"basic": 0, "standard": 6, "premium": 7}, "prices": {"basic": 9800, "standard": 59800, "premium": 149800}, "chapters": [{"no": 0, "tag": "序章", "name": "なぜ今、AIセキュリティを学ぶ必要があるのか", "color": "var(--red)", "lessons": ["intro"]}, {"no": 1, "tag": "第1章", "name": "AIと情報漏洩の基礎知識", "color": "var(--red)", "lessons": ["lesson1", "lesson2", "lesson3", "lesson4"], "workGoal": "自社のAI利用実態マップ"}, {"no": 2, "tag": "第2章", "name": "ツール別リスクと設定", "color": "var(--amber)", "lessons": ["lesson5", "lesson6", "lesson7", "lesson8"], "workGoal": "入力禁止情報リスト"}, {"no": 3, "tag": "第3章", "name": "法律・著作権・ハルシネーション", "color": "var(--amber)", "lessons": ["lesson9", "lesson10", "lesson11"], "workGoal": "リスク評価シート"}, {"no": 4, "tag": "第4章", "name": "組織・ガイドライン整備", "color": "var(--blue)", "lessons": ["lesson12", "lesson13", "lesson14", "lesson15"], "workGoal": "社内AI利用ガイドライン"}, {"no": 5, "tag": "第5章", "name": "実践：事例研究とツール別ガイド", "color": "var(--blue)", "lessons": ["lesson16", "lesson17", "lesson28", "lesson29", "lesson30", "lesson31", "lesson32"], "workGoal": "AI導入計画メモ"}, {"no": 6, "tag": "第6章", "name": "AI攻撃手口と防衛策・管理者設定", "color": "#0ea5e9", "lessons": ["lesson18", "lesson19", "lesson20", "lesson21", "lesson22", "lesson33", "lesson34"], "badge": "Standard+", "workGoal": "インシデント対応フロー"}, {"no": 7, "tag": "第7章", "name": "AI技術とセキュリティの深層・開発", "color": "#8b5cf6", "lessons": ["lesson23", "lesson24", "lesson25", "lesson26", "lesson27", "lesson35", "lesson36"], "badge": "Premium", "workGoal": "AI活用ロードマップ"}], "titles": {"intro": "序章：なぜ今学ぶか", "lesson1": "入力情報はどこへ行くのか", "lesson2": "学習オフ設定の意味と限界", "lesson3": "シャドーAIとは何か", "lesson4": "クラウドAIとオンプレAIの違い", "lesson5": "ChatGPT各プランの違い", "lesson6": "Claude・Gemini・Copilot比較", "lesson7": "社外秘・個人情報の判断基準", "lesson8": "業種別リスク事例", "lesson9": "AIと個人情報保護法", "lesson10": "著作権侵害になるケース", "lesson11": "ハルシネーションと法的責任", "lesson12": "ガイドライン作成テンプレート", "lesson13": "インシデント初動対応フロー", "lesson14": "中小企業の最低限の体制", "lesson15": "今日からできる10のアクション", "lesson16": "よくある事例集（20事例）", "lesson17": "主要AIツール完全比較", "lesson18": "AIを使ったフィッシング詐欺", "lesson19": "プロンプトインジェクション攻撃", "lesson20": "ディープフェイクと音声詐欺", "lesson21": "AIエージェント・自動化ツールのリスク", "lesson22": "インシデント対応ロールプレイ", "lesson23": "LLMの仕組み入門", "lesson24": "プロンプトエンジニアリング基礎", "lesson25": "RAG・ファインチューニングのリスク", "lesson26": "業種別AIロードマップ", "lesson27": "MCPとAIエージェント最前線", "lesson28": "ChatGPTを安全に使う", "lesson29": "GeminiとGoogle Workspace", "lesson30": "CopilotとMicrosoft 365", "lesson31": "Claudeを安全に使う", "lesson32": "MetaのMuseとSNSのAI", "lesson33": "AIブラウザとエージェント機能", "lesson34": "管理者のための設定集", "lesson35": "Claude Codeを安全に使う", "lesson36": "AIコーディングと秘密情報"}};
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

// HTML に文字を入れるときのエスケープ（2026-11-23。以前は同じ関数が9か所に別々に書かれていた）
function escHtml(t){ return String(t == null ? '' : t).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
// CSV の1セル：数式として解釈されうる先頭（= + - @ タブ 改行）を無害化し、ダブルクォートで囲む
function csvCell(v){ let t = String(v == null ? '' : v); if(/^[=+\-@\t\r]/.test(t)) t = "'" + t; return '"' + t.replace(/"/g, '""') + '"'; }
window.escHtml = escHtml;
// 端末への保存（localStorage）の読み書き。サイトデータを止めた環境では読むだけで例外になり、
// 以前はページ全体のスクリプトが止まって、どのボタンも効かなくなった（2026-11-23）
const store = {
  get(k, def = null){ try { const v = localStorage.getItem(k); return v == null ? def : v; } catch(e){ return def; } },
  getJSON(k, def){ try { const v = localStorage.getItem(k); return v == null ? def : JSON.parse(v); } catch(e){ return def; } },
  set(k, v){ try { localStorage.setItem(k, typeof v === 'string' ? v : JSON.stringify(v)); } catch(e){} },
  remove(k){ try { localStorage.removeItem(k); } catch(e){} },
};

// バージョンを固定する。@2 のままだとCDN側の更新がそのまま本番に入るため。
// Supabase のライブラリは自前で配信する（/assets/vendor/supabase-js-2.112.3.js。2026-11-23）。
// 外部の配信元の改ざんや停止の影響を受けないように。取得時に jsDelivr と unpkg の SHA-256 が一致することを確かめた
const { createClient } = window.supabase;



// ---- インラインのイベントハンドラ（onclick="…"）を、CSP で 'unsafe-inline' を外しても動かす（2026-11-23）----
// ページ・目次・レッスン本文（DB から配信）には onclick="switchTab('lesson3');return false;" のような書き方が
// 多数あり、すべてを書き換えるのは本文の入れ直しも要る。そこで、ブラウザがインラインの実行を止めているときだけ、
// 属性の文字を小さな構文解析器で読み、決まった形の呼び出しだけを実行する。
//   使える形：関数(引数, …) ／ window.関数(…) ／ return 式 ／ if(event.target===this) 文 ／ window._名前 = 値
//   引数：'文字列' "文字列" 数 true false null this event this.value document.querySelector('…')
//   呼べる関数：このページが window に載せた関数だけ（fetch・eval などのブラウザ組み込みは呼べない）
function scrollToId(id){ const el = document.getElementById(id); if(el) el.scrollIntoView({ behavior: 'smooth' }); }
window.scrollToId = scrollToId;
const inlineShim = (() => {
  // 呼べるのは、このスクリプトが読み込まれた後に window に載った関数だけ（fetch・eval・gtag など、
  // それより前からある関数は呼べない。差し替えられた fetch も同じ）
  const before = WINDOW_BEFORE;
  const isOwnFn = (f, name) => typeof f === 'function' && !before.has(name) && !/\[native code\]/.test(Function.prototype.toString.call(f));
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
      if('(),;='.includes(c)){ t.push({ k: 'p', v: c }); i++; continue; }
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
      const fn = window[fname];
      if(!isOwnFn(fn, fname)) throw new Error('呼べない: ' + name);
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
      if(x.k === 'i' && /^window\._\w+$/.test(x.v) && t[p + 1] && t[p + 1].v === '='){
        eat(); eat('='); const v = value(); if(!dry) window[x.v.slice(7)] = v; return undefined;
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
