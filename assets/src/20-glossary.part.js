/* ===== 用語集（2026-11-23） =====
   - レッスン本文の中で、用語集にある言葉の「そのレッスンで最初の1回」に点線を付ける。押すと1〜2行の説明が出る（ページは移らない）
   - 用語集のページ（tab-glossary）：初めて出てくる章ごとに並べ、検索と「出てくるレッスン」へのリンクを付ける
   - 点数・記録・修了条件・標準学習時間には関わらない
   - 理解度チェック（設問と選択肢）の中には付けない（答えの手がかりにしないため）
   用語と説明の正本は asl-worker/course/glossary.json。下のブロックは course-sync.py が書き出す */
// <GLOSSARY:BEGIN> course/glossary.json と最新のレッスン本文から course-sync.py が書き出す。手で直さない
const GLOSSARY = [
  {"id": "genai", "term": "生成AI", "aliases": [], "short": "文章・画像・音声などを、指示に応じて新しく作り出すAI。ChatGPT・Claude・Gemini・Copilot など。", "lessons": ["intro", "lesson1", "lesson3", "lesson7", "lesson8", "lesson9", "lesson10", "lesson11", "lesson12", "lesson13", "lesson14", "lesson15", "lesson17", "lesson18", "lesson22", "lesson25", "lesson26", "lesson27"]},
  {"id": "personal-info", "term": "個人情報", "aliases": [], "short": "生存する個人に関する情報で、氏名などから特定の個人を識別できるもの（ほかの情報と容易に照合して識別できるものを含む）。", "lessons": ["intro", "lesson1", "lesson2", "lesson3", "lesson4", "lesson5", "lesson7", "lesson8", "lesson9", "lesson12", "lesson13", "lesson14", "lesson15", "lesson16", "lesson17", "lesson18", "lesson22", "lesson33", "lesson24", "lesson25", "lesson26", "lesson27"]},
  {"id": "third-party", "term": "第三者提供", "aliases": [], "short": "個人データを本人以外の他者に渡すこと。原則として本人の同意が必要（個人情報保護法第27条）。", "lessons": ["intro", "lesson1", "lesson5", "lesson8", "lesson9", "lesson16"]},
  {"id": "confidential", "term": "機密情報", "aliases": ["社外秘"], "short": "外に出ると会社や取引先に損害が出る情報。未公開の計画・原価・顧客リスト・ソースコードなど。", "lessons": ["intro", "lesson1", "lesson2", "lesson6", "lesson7", "lesson8", "lesson12", "lesson14", "lesson16", "lesson17", "lesson29", "lesson18", "lesson20", "lesson23", "lesson24", "lesson25"]},
  {"id": "copyright", "term": "著作権", "aliases": [], "short": "文章・画像・音楽などの作品を作った人の権利。AIの生成物が既存の作品に似ると、侵害になることがある。", "lessons": ["intro", "lesson10", "lesson12", "lesson15", "lesson16", "lesson24", "lesson25", "lesson27"]},
  {"id": "ppc", "term": "個人情報保護委員会", "aliases": [], "short": "個人情報保護法を所管する国の機関。個人データの漏えいなどが起きたときの報告先。", "lessons": ["intro", "lesson1", "lesson2", "lesson7", "lesson8", "lesson9", "lesson12", "lesson13", "lesson15", "lesson16", "lesson22"]},
  {"id": "incident", "term": "インシデント", "aliases": [], "short": "情報漏えいや不正利用など、セキュリティ上の事故や、その疑いのある出来事。", "lessons": ["intro", "lesson2", "lesson5", "lesson13", "lesson14", "lesson15", "lesson19", "lesson22", "lesson27"]},
  {"id": "first-response", "term": "初動対応", "aliases": [], "short": "インシデントに気づいた直後の対応。記録を消さず、広がりを止め、決められた相手に報告する。", "lessons": ["intro", "lesson13", "lesson22", "lesson27"]},
  {"id": "ai-guideline", "term": "AI事業者ガイドライン", "aliases": [], "short": "総務省・経済産業省が示す、AIを開発・提供・利用する事業者向けの指針。", "lessons": ["intro", "lesson3", "lesson12", "lesson15"]},
  {"id": "api", "term": "API", "aliases": [], "short": "プログラムからサービスを呼び出すための窓口。自社のシステムから ChatGPT などを使うときに使う。", "lessons": ["lesson1", "lesson2", "lesson4", "lesson6", "lesson7", "lesson12", "lesson16", "lesson17", "lesson31", "lesson19", "lesson21", "lesson34", "lesson27", "lesson35", "lesson36"]},
  {"id": "foreign-third", "term": "外国にある第三者", "aliases": [], "short": "外国の事業者。個人データを渡すときは、同意の取り方などに追加の決まりがある（個人情報保護法第28条）。", "lessons": ["lesson1", "lesson2"]},
  {"id": "audit-log", "term": "監査ログ", "aliases": [], "short": "誰がいつ何をしたかの記録。AIへの入力を、会社が後から確かめるのに使う。", "lessons": ["lesson1", "lesson17", "lesson31", "lesson34"]},
  {"id": "temp-chat", "term": "一時チャット", "aliases": [], "short": "会話履歴に残さず、学習にも使われない ChatGPT の会話の方式。ただし事業者のサーバーには最大30日残る。", "lessons": ["lesson1", "lesson2", "lesson6", "lesson28", "lesson29"]},
  {"id": "browser-ext", "term": "ブラウザ拡張", "aliases": [], "short": "ブラウザに機能を足す追加のプログラム。AI機能付きのものは、閲覧中のページの内容を外部に送ることがある。", "lessons": ["lesson1", "lesson2", "lesson3", "lesson19"]},
  {"id": "training-off", "term": "学習オフ", "aliases": [], "short": "入力した内容をAIの学習（モデルの改良）に使わせない設定。送信や事業者のサーバーでの保存そのものは止まらない。", "lessons": ["lesson2", "lesson5", "lesson16", "lesson28", "lesson31"]},
  {"id": "opt-out", "term": "オプトアウト", "aliases": [], "short": "「止めたい人が自分でオフにする」方式。何もしなければ、学習などに使われる。", "lessons": ["lesson2", "lesson6", "lesson17", "lesson31", "lesson32"]},
  {"id": "mfa", "term": "多要素認証", "aliases": ["MFA"], "short": "パスワードに加えて、スマホのアプリのコードなど別の手段でも本人を確かめるログインの方法。", "lessons": ["lesson2", "lesson18"]},
  {"id": "shadow-ai", "term": "シャドーAI", "aliases": [], "short": "会社が把握・許可していないAIツールを、社員が業務で使っている状態。", "lessons": ["lesson3", "lesson6", "lesson14", "lesson16", "lesson22", "lesson26"]},
  {"id": "on-prem", "term": "オンプレミス", "aliases": ["オンプレAI", "オンプレ"], "short": "AIなどのシステムを、自社のサーバーや設備に置いて運用する形。事業者のサーバーで動くクラウドの対になる言葉。", "lessons": ["lesson4"]},
  {"id": "dpa", "term": "DPA", "aliases": ["データ処理委託契約", "データ処理契約"], "short": "事業者が預かったデータをどう扱うかを定める契約。法人向けプランで結べることが多い。", "lessons": ["lesson4", "lesson5", "lesson6", "lesson7", "lesson9", "lesson17", "lesson29", "lesson31", "lesson34"]},
  {"id": "entrust", "term": "委託", "aliases": [], "short": "業務を外部に任せ、その範囲で個人データを預けること。第三者提供とは扱いが異なり、委託先を監督する義務がある。", "lessons": ["lesson4", "lesson5", "lesson6", "lesson9", "lesson12", "lesson14", "lesson15", "lesson17", "lesson28", "lesson29", "lesson32", "lesson18", "lesson33", "lesson34"]},
  {"id": "hallucination", "term": "ハルシネーション", "aliases": [], "short": "AIが事実と違う内容を、もっともらしく答えてしまうこと。出典や数字は必ず確かめる。", "lessons": ["lesson4", "lesson11", "lesson16", "lesson17", "lesson23", "lesson26"]},
  {"id": "nda", "term": "秘密保持契約", "aliases": ["NDA"], "short": "取引で知った相手の秘密を漏らさないと約束する契約。預かった資料をAIに入れてよいかも、この契約で決まる。", "lessons": ["lesson5", "lesson6", "lesson8", "lesson13", "lesson28", "lesson21", "lesson25", "lesson26"]},
  {"id": "dlp", "term": "DLP", "aliases": [], "short": "機密情報が社外に出るのを見つけて止める仕組み（Data Loss Prevention）。", "lessons": ["lesson6", "lesson17", "lesson29", "lesson30", "lesson34"]},
  {"id": "retention", "term": "データ保持", "aliases": [], "short": "事業者が入力内容や会話を保存しておくこと。削除しても、決められた期間は残ることがある。", "lessons": ["lesson6", "lesson17", "lesson34", "lesson35"]},
  {"id": "claude-code", "term": "Claude Code", "aliases": [], "short": "Anthropic の、パソコン上でコードを読み書きしたりコマンドを実行したりするAIエージェント。", "lessons": ["lesson6", "lesson31", "lesson21", "lesson34", "lesson27", "lesson35", "lesson36"]},
  {"id": "sharepoint", "term": "SharePoint", "aliases": [], "short": "Microsoft 365 のファイル共有の仕組み。権限が緩いと、Copilot が見えてはいけない資料まで読んでしまう。", "lessons": ["lesson6", "lesson30", "lesson34"]},
  {"id": "api-key", "term": "APIキー", "aliases": [], "short": "API を使うための合言葉にあたる文字列。漏れると他人に使われ、料金の請求や情報の流出につながる。", "lessons": ["lesson7", "lesson12", "lesson21", "lesson27"]},
  {"id": "sensitive-pi", "term": "要配慮個人情報", "aliases": [], "short": "人種・信条・病歴・犯罪の経歴など、差別や偏見につながりうる個人情報。取得には原則として本人の同意が要る。", "lessons": ["lesson7", "lesson8", "lesson9", "lesson12", "lesson13", "lesson16", "lesson25"]},
  {"id": "anonymized", "term": "匿名加工情報", "aliases": [], "short": "特定の個人を識別できず、元に戻せないよう法の基準で加工した情報。名前を伏せただけでは当たらない。", "lessons": ["lesson7"]},
  {"id": "mosaic", "term": "モザイク効果", "aliases": [], "short": "一つひとつは問題のない情報でも、組み合わせると個人や機密が特定できてしまうこと。", "lessons": ["lesson7"]},
  {"id": "trade-secret", "term": "営業秘密", "aliases": [], "short": "不正競争防止法で守られる会社の秘密。秘密として管理していること・事業に役立つこと・公になっていないことの3つが条件。", "lessons": ["lesson7", "lesson15", "lesson26"]},
  {"id": "unlawful-acq", "term": "不正取得", "aliases": [], "short": "人をだます・不正にアクセスするなどの手段で個人情報を手に入れること。改正個人情報保護法で、その行為への罰則が設けられる（施行日はレッスンを参照）。", "lessons": ["lesson9"]},
  {"id": "similarity", "term": "類似性", "aliases": [], "short": "生成物が既存の作品と似ているかどうか。依拠性とあわせて、著作権侵害かを判断する基準。", "lessons": ["lesson10"]},
  {"id": "reliance", "term": "依拠性", "aliases": [], "short": "既存の作品をもとにして作ったかどうか。類似性とあわせて、著作権侵害かを判断する基準。", "lessons": ["lesson10"]},
  {"id": "llm", "term": "大規模言語モデル", "aliases": ["LLM"], "short": "大量の文章から学び、続きの言葉を予測して文章を作るAIのモデル。ChatGPT などの中核。", "lessons": ["lesson11", "lesson19", "lesson21", "lesson23", "lesson24", "lesson25", "lesson27"]},
  {"id": "ai-promotion-act", "term": "AI推進法", "aliases": [], "short": "AIの研究開発と活用を進めるための日本の法律（2025年成立）。", "lessons": ["lesson15"]},
  {"id": "ai-agent", "term": "AIエージェント", "aliases": [], "short": "指示を受けて、調べる・操作する・送るなどの作業を自分で順に進めるAI。", "lessons": ["lesson15", "lesson16", "lesson17", "lesson19", "lesson21", "lesson27"]},
  {"id": "agent-mode", "term": "エージェントモード", "aliases": [], "short": "ChatGPT が、Webの操作などの作業を自分で進めるモード。", "lessons": ["lesson28", "lesson33", "lesson34"]},
  {"id": "prompt-injection", "term": "プロンプトインジェクション", "aliases": [], "short": "AIへの入力に悪意のある命令を紛れ込ませ、本来しない動作をさせる攻撃。", "lessons": ["lesson29", "lesson32", "lesson19", "lesson21", "lesson33", "lesson25", "lesson27", "lesson35", "lesson36"]},
  {"id": "indirect-pi", "term": "間接プロンプトインジェクション", "aliases": [], "short": "AIが読むWebページやメール・文書の中に命令を隠しておき、AIを乗っ取る攻撃。", "lessons": ["lesson29"]},
  {"id": "sandbox", "term": "サンドボックス", "aliases": [], "short": "プログラムを、ほかに影響しない隔離された場所で動かす仕組み。", "lessons": ["lesson31", "lesson21", "lesson33", "lesson35"]},
  {"id": "phishing", "term": "フィッシング", "aliases": [], "short": "本物そっくりのメールやサイトで、パスワードやお金をだまし取る手口。生成AIで文面が自然になり、見破りにくくなった。", "lessons": ["lesson18", "lesson22"]},
  {"id": "bec", "term": "ビジネスメール詐欺", "aliases": ["BEC"], "short": "取引先や経営者になりすましたメールで、送金や情報を求める詐欺。", "lessons": ["lesson18", "lesson20"]},
  {"id": "deepfake", "term": "ディープフェイク", "aliases": [], "short": "AIで作った、本物そっくりの偽の映像や音声。経営者になりすました送金の指示などに使われる。", "lessons": ["lesson18", "lesson20"]},
  {"id": "least-priv", "term": "最小権限", "aliases": ["権限の最小化"], "short": "仕事に必要な最低限の権限だけを与える考え方。AIに任せる操作にも当てはめる。", "lessons": ["lesson19", "lesson21"]},
  {"id": "voice-clone", "term": "音声クローン", "aliases": [], "short": "数秒〜数分の声の素材から、本人そっくりの声を合成する技術。", "lessons": ["lesson20"]},
  {"id": "auto-approve", "term": "自動承認", "aliases": [], "short": "AIエージェントの操作を、人が1つずつ確かめずに進めさせる設定。", "lessons": ["lesson21", "lesson27", "lesson36"]},
  {"id": "mcp", "term": "MCP", "aliases": [], "short": "AIエージェントに外部のツールやデータをつなぐための共通の仕組み（Model Context Protocol）。", "lessons": ["lesson21", "lesson34", "lesson27", "lesson35", "lesson36"]},
  {"id": "context-window", "term": "コンテキストウィンドウ", "aliases": [], "short": "LLM が一度に扱える文章の量の上限。超えると、古い内容から忘れていく。", "lessons": ["lesson23"]},
  {"id": "custom-instr", "term": "カスタム指示", "aliases": [], "short": "毎回の会話に自動で付ける、あらかじめ書いておく指示。書いた内容は全部の会話で送られ続ける。", "lessons": ["lesson24"]},
  {"id": "prompt-eng", "term": "プロンプトエンジニアリング", "aliases": [], "short": "AIから望む答えを引き出すための、指示（プロンプト）の書き方の工夫。", "lessons": ["lesson24", "lesson25"]},
  {"id": "few-shot", "term": "Few-shot", "aliases": [], "short": "入力と期待する出力の例をいくつか示してから、本題を頼む書き方。", "lessons": ["lesson24"]},
  {"id": "rag", "term": "RAG", "aliases": [], "short": "社内文書などを検索し、その内容をAIに渡して答えさせる仕組み（検索拡張生成）。", "lessons": ["lesson25"]},
  {"id": "fine-tuning", "term": "ファインチューニング", "aliases": [], "short": "既存のAIモデルを、自社のデータで追加学習させて調整すること。", "lessons": ["lesson25"]},
  {"id": "ai-browser", "term": "AIブラウザ", "aliases": [], "short": "AIがWebの閲覧やフォームへの入力などを代わりに行うブラウザ。", "lessons": ["lesson27"]},
  {"id": "zdr", "term": "ゼロデータ保持", "aliases": ["ZDR"], "short": "入力内容を事業者が保存しない扱い。条件を満たした法人契約に限られる。", "lessons": ["lesson35"]},
  {"id": "dotenv", "term": ".env", "aliases": [], "short": "APIキーなどの設定を書いておくファイル。AIに読ませない設定にしておく。", "lessons": ["lesson35", "lesson36"]},
  {"id": "secret", "term": "シークレット", "aliases": [], "short": "パスワード・APIキー・接続文字列など、知られると悪用される秘密の値。", "lessons": ["lesson36"]}
];
// <GLOSSARY:END>

const GLOSSARY_BY_ID = Object.fromEntries(GLOSSARY.map(t => [t.id, t]));
// 用語（と別名）→ id。長いものから当てる（「プロンプトインジェクション」より先に「間接プロンプトインジェクション」）
const GL_WORDS = GLOSSARY.flatMap(t => [t.term].concat(t.aliases || []).map(w => [w, t.id])).sort((a, b) => b[0].length - a[0].length);
const GL_RE = GL_WORDS.length ? new RegExp(GL_WORDS.map(([w]) => {
  let e = w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  if(/^[A-Za-z0-9.]/.test(w)) e = '(?<![A-Za-z0-9])' + e;
  if(/[A-Za-z0-9]$/.test(w)) e = e + '(?![A-Za-z0-9])';
  return e;
}).join('|'), 'g') : null;
const GL_ID_OF = Object.fromEntries(GL_WORDS);
// 付けない場所：見出し・設問と選択肢・ボタン・リンク・図の見出し・コード
const GL_SKIP = 'h1,h2,h3,h4,.ltit,.leye,.lmeta,.quiz,.qopts,.qopt,.qfb,button,a,.gl-term,code,pre,.dbtit,.fig-t,.lnav,.completion-area,.toc-inline,.resume-bar,script,style';

// レッスン本文の用語に印を付ける（同じレッスンでは最初の1回だけ）
function annotateTerms(tcon){
  if(!GL_RE || !tcon || tcon.dataset.glDone) return;
  const lb = tcon.querySelector('.lb');
  if(!lb) return;
  tcon.dataset.glDone = '1';
  const used = new Set();
  const walker = document.createTreeWalker(lb, NodeFilter.SHOW_TEXT, {
    acceptNode: n => (n.nodeValue.trim() && !(n.parentElement && n.parentElement.closest(GL_SKIP))) ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT
  });
  const nodes = [];
  for(let n = walker.nextNode(); n; n = walker.nextNode()) nodes.push(n);
  for(const node of nodes){
    const text = node.nodeValue;
    GL_RE.lastIndex = 0;
    let m, last = 0, frag = null;
    while((m = GL_RE.exec(text))){
      const id = GL_ID_OF[m[0]];
      if(used.has(id)) continue;
      used.add(id);
      frag = frag || document.createDocumentFragment();
      frag.appendChild(document.createTextNode(text.slice(last, m.index)));
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'gl-term';
      b.dataset.gl = id;
      b.setAttribute('aria-haspopup', 'dialog');
      b.textContent = m[0];
      frag.appendChild(b);
      last = m.index + m[0].length;
    }
    if(frag){
      frag.appendChild(document.createTextNode(text.slice(last)));
      node.parentNode.replaceChild(frag, node);
    }
  }
}

// 用語の説明（小さな吹き出し）
function showTermPopover(btn){
  const t = GLOSSARY_BY_ID[btn.dataset.gl];
  if(!t) return;
  let pop = document.getElementById('gl-pop');
  if(!pop){
    pop = document.createElement('div');
    pop.id = 'gl-pop';
    pop.setAttribute('role', 'dialog');
    pop.setAttribute('aria-label', '用語の説明');
    document.body.appendChild(pop);
  }
  pop.innerHTML = `<div class="gl-pop-term">${escHtml(t.term)}${(t.aliases || []).length ? `<span class="gl-pop-alias">（${escHtml(t.aliases.join('・'))}）</span>` : ''}</div>
    <p class="gl-pop-short">${escHtml(t.short)}</p>
    ${currentUser ? `<button type="button" class="gl-pop-more" data-gl-open="${t.id}">用語集で見る →</button>` : ''}`;
  pop.hidden = false;
  const r = btn.getBoundingClientRect();
  const w = Math.min(320, window.innerWidth - 24);
  pop.style.width = w + 'px';
  pop.style.left = Math.max(12, Math.min(r.left, window.innerWidth - w - 12)) + 'px';
  const below = r.bottom + 8, h = pop.offsetHeight;
  pop.style.top = (below + h > window.innerHeight - 8 ? Math.max(8, r.top - h - 8) : below) + 'px';
  document.querySelectorAll('.gl-term[aria-expanded="true"]').forEach(x => x.setAttribute('aria-expanded', 'false'));
  btn.setAttribute('aria-expanded', 'true');
}
function hideTermPopover(){
  const pop = document.getElementById('gl-pop');
  if(pop) pop.hidden = true;
  document.querySelectorAll('.gl-term[aria-expanded="true"]').forEach(x => x.setAttribute('aria-expanded', 'false'));
}
document.addEventListener('click', (e) => {
  const term = e.target.closest && e.target.closest('.gl-term');
  if(term){ e.preventDefault(); showTermPopover(term); return; }
  const more = e.target.closest && e.target.closest('[data-gl-open]');
  if(more){ hideTermPopover(); openGlossary(more.dataset.glOpen); return; }
  if(!(e.target.closest && e.target.closest('#gl-pop'))) hideTermPopover();
});
document.addEventListener('keydown', (e) => { if(e.key === 'Escape') hideTermPopover(); });
{ const sc = document.getElementById('learn-content'); if(sc) sc.addEventListener('scroll', hideTermPopover, { passive: true }); }
window.addEventListener('resize', hideTermPopover);

// 用語集のページ
function renderGlossary(){
  const tcon = document.getElementById('tab-glossary');
  if(!tcon) return;
  if(!tcon.dataset.built){
    tcon.dataset.built = '1';
    tcon.innerHTML = `<div class="lc gl-page">
      <div class="leye">GLOSSARY</div>
      <h2 class="ltit">用語集</h2>
      <p class="gl-lead">レッスンに出てくる用語を、初めて出てくる章ごとに並べています（${GLOSSARY.length}語）。レッスンの本文で点線の付いた言葉を押しても、ここと同じ説明が出ます。</p>
      <input type="search" id="gl-search" class="gl-search" placeholder="用語をさがす（例：DPA、プロンプト）" aria-label="用語をさがす">
      <div id="gl-list"></div>
    </div>`;
    tcon.querySelector('#gl-search').addEventListener('input', (e) => drawGlossaryList(e.target.value));
  }
  drawGlossaryList(tcon.querySelector('#gl-search').value || '');
}
function drawGlossaryList(q){
  const box = document.getElementById('gl-list');
  if(!box) return;
  const key = String(q || '').trim().toLowerCase();
  const hit = t => !key || [t.term, t.short].concat(t.aliases || []).some(s => s.toLowerCase().includes(key));
  const chOf = id => COURSE.chapters.find(c => c.lessons.includes(id));
  const groups = [];
  GLOSSARY.filter(hit).forEach(t => {
    const ch = chOf(t.lessons[0]);
    let g = groups.find(x => x.ch === ch);
    if(!g){ g = { ch, items: [] }; groups.push(g); }
    g.items.push(t);
  });
  if(!groups.length){ box.innerHTML = '<p class="gl-empty">見つかりませんでした。別の言葉でさがしてください。</p>'; return; }
  box.innerHTML = groups.map(g => `<section class="gl-group"><h3 class="gl-group-h">${escHtml(g.ch ? g.ch.tag + '・' + g.ch.name : 'その他')}</h3>`
    + g.items.map(t => `<div class="gl-item" id="gl-${t.id}">
        <div class="gl-item-term">${escHtml(t.term)}${(t.aliases || []).length ? `<span class="gl-pop-alias">（${escHtml(t.aliases.join('・'))}）</span>` : ''}</div>
        <p class="gl-item-short">${escHtml(t.short)}</p>
        <div class="gl-item-lessons"><span>出てくるレッスン：</span>${t.lessons.slice(0, 6).map(id => {
          const label = id === 'intro' ? '序章' : 'L' + lessonNo(id);
          return tabOpenable(id)
            ? `<button type="button" class="gl-lesson" onclick="switchTab('${id}')" title="${escHtml(COURSE.titles[id] || '')}">${label}</button>`
            : `<span class="gl-lesson is-locked" title="ご契約のプランの範囲外です">${label} 🔒</span>`;
        }).join('')}${t.lessons.length > 6 ? `<span class="gl-more">ほか${t.lessons.length - 6}本</span>` : ''}</div>
      </div>`).join('') + '</section>').join('');
}
// 用語集を開き、その用語の場所へ移る
function openGlossary(id){
  window.switchTab('glossary');
  setTimeout(() => {
    const el = id && document.getElementById('gl-' + id);
    if(!el) return;
    el.scrollIntoView({ block: 'center', behavior: 'smooth' });
    el.classList.add('is-flash');
    setTimeout(() => el.classList.remove('is-flash'), 1600);
  }, 120);
}

// トップページ（未ログイン）で読める序章・L01 の本文にも印を付ける
document.addEventListener('DOMContentLoaded', () => { ['intro', 'lesson1'].forEach(id => annotateTerms(document.getElementById('tab-' + id))); });
