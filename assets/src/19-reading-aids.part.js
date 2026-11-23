/* ===== 読みやすさ：文字の大きさ・レッスン内の目次・読んだ位置（2026-11-23） =====
   - 文字の大きさ（小・中・大）を上部のボタンで切り替える。端末ごとに覚える
   - 見出しが2つ以上あるレッスンは、目次を出す（広い画面は右の欄、狭い画面は本文の前に畳んで）
   - 読み進めた割合と、残りの目安時間（「読了時間：約N分」から）を上部に出す
   - 途中まで読んだレッスンを開き直すと、「続きから読む」を出す（自動では動かさない） */
const FS_KEY = 'asl_font_size';
const READ_KEY = 'asl_read_pos';
const FS_STEPS = [['s', '小'], ['m', '中'], ['l', '大']];

function applyFontSize(){
  const v = store.get(FS_KEY, 'm');
  const inner = document.getElementById('learn-content-inner');
  if(inner) inner.dataset.fs = v;
  const b = document.getElementById('fs-btn');
  if(b){
    const label = (FS_STEPS.find(x => x[0] === v) || FS_STEPS[1])[1];
    b.textContent = '文字 ' + label;
    b.setAttribute('aria-label', '文字の大きさ：' + label + '（押すと切り替わります）');
  }
}
// 小 → 中 → 大 → 小 …。切り替えても、読んでいた位置（割合）を保つ
function cycleFontSize(){
  const sc = document.getElementById('learn-content');
  const ratio = sc ? sc.scrollTop / Math.max(1, sc.scrollHeight - sc.clientHeight) : 0;
  const i = FS_STEPS.findIndex(x => x[0] === store.get(FS_KEY, 'm'));
  store.set(FS_KEY, FS_STEPS[(i + 1) % FS_STEPS.length][0]);
  applyFontSize();
  if(sc) sc.scrollTop = ratio * (sc.scrollHeight - sc.clientHeight);
}

// レッスンを開いたとき（本文の読み込み後）に呼ぶ
function setupReading(tabId){
  const inner = document.getElementById('learn-content-inner');
  const tcon = document.getElementById('tab-' + tabId);
  const topbar = document.getElementById('learn-topbar');
  const old = document.getElementById('lesson-toc');
  if(old) old.remove();
  if(inner) inner.classList.remove('has-toc');
  const lb = tcon && tcon.querySelector('.lb');
  const isLesson = !!lb && (tabId === 'intro' || ALL_LESSONS_LIST.includes(tabId));
  if(topbar) topbar.classList.toggle('is-reading', isLesson);
  if(!isLesson){ updateReadProgress(); return; }

  const heads = [...lb.querySelectorAll('h2')];
  heads.forEach((h, i) => { if(!h.id) h.id = tabId + '-s' + (i + 1); });
  if(heads.length >= 2 && inner){
    const items = heads.map(h => `<li><a href="#${h.id}" data-toc="${h.id}">${escHtml(h.textContent.trim())}</a></li>`).join('');
    // <nav> はサイトの上部バーの見た目が付くので、aside に role を付ける
    const aside = document.createElement('aside');
    aside.id = 'lesson-toc';
    aside.setAttribute('role', 'navigation');
    aside.setAttribute('aria-label', 'このレッスンの目次');
    aside.innerHTML = `<div class="toc-h">このレッスンの目次</div><ol>${items}</ol>`;
    inner.appendChild(aside);
    inner.classList.add('has-toc');
    if(!tcon.querySelector('.toc-inline')){
      const d = document.createElement('details');
      d.className = 'toc-inline';
      d.innerHTML = `<summary>このレッスンの目次（${heads.length}）</summary><ol>${items}</ol>`;
      const meta = tcon.querySelector('.lmeta');
      if(meta) meta.insertAdjacentElement('afterend', d); else lb.insertAdjacentElement('beforebegin', d);
    }
  }
  offerResume(tabId, tcon);
  updateReadProgress();
}

// 途中まで読んでいたら「続きから読む」を出す
function offerResume(tabId, tcon){
  const old = tcon.querySelector('.resume-bar');
  if(old) old.remove();
  const r = (store.getJSON(READ_KEY, {}) || {})[tabId];
  if(!(r > 0.08 && r < 0.95)) return;
  const bar = document.createElement('div');
  bar.className = 'resume-bar';
  bar.innerHTML = `<span>前回は <strong>${Math.round(r * 100)}%</strong> まで読みました</span>
    <button type="button" class="learn-nav-btn resume-go">続きから読む →</button>
    <button type="button" class="resume-close" aria-label="閉じる">×</button>`;
  bar.querySelector('.resume-go').addEventListener('click', () => {
    const sc = document.getElementById('learn-content');
    if(sc) sc.scrollTo({ top: r * (sc.scrollHeight - sc.clientHeight), behavior: 'smooth' });
    bar.remove();
  });
  bar.querySelector('.resume-close').addEventListener('click', () => bar.remove());
  const title = tcon.querySelector('.ltit');
  (title || tcon.firstElementChild).insertAdjacentElement(title ? 'beforebegin' : 'afterbegin', bar);
}

let _readSavedAt = 0;
function updateReadProgress(){
  const sc = document.getElementById('learn-content');
  const bar = document.getElementById('read-progress');
  const left = document.getElementById('read-left');
  const topbar = document.getElementById('learn-topbar');
  if(!sc || !bar) return;
  if(!(topbar && topbar.classList.contains('is-reading'))){ bar.style.width = '0'; if(left) left.textContent = ''; return; }
  const ratio = Math.min(1, sc.scrollTop / Math.max(1, sc.scrollHeight - sc.clientHeight));
  bar.style.width = (ratio * 100).toFixed(1) + '%';
  const tcon = document.querySelector('#learn-content-inner .tcon.active');
  const m = /約(\d+)分/.exec(((tcon && tcon.querySelector('.lmeta')) || {}).textContent || '');
  if(left) left.textContent = m ? (ratio > 0.97 ? '読み終わりました' : `残り約${Math.max(1, Math.ceil(Number(m[1]) * (1 - ratio)))}分`) : '';
  // 今読んでいる見出しを目次で示す
  const heads = tcon ? [...tcon.querySelectorAll('.lb h2[id]')] : [];
  const line = sc.getBoundingClientRect().top + 96;
  let cur = null;
  for(const h of heads){ if(h.getBoundingClientRect().top <= line) cur = h.id; else break; }
  document.querySelectorAll('[data-toc]').forEach(a => a.classList.toggle('is-current', a.dataset.toc === cur));
  // 読んだ位置を覚える（1秒に1回まで。読み終えたら消す）
  const now = Date.now();
  if(tcon && now - _readSavedAt > 1000){
    _readSavedAt = now;
    const id = tcon.id.replace('tab-', '');
    const all = store.getJSON(READ_KEY, {}) || {};
    if(ratio >= 0.95) delete all[id]; else if(ratio > 0.02) all[id] = Math.round(ratio * 1000) / 1000;
    store.set(READ_KEY, all);
  }
}

// 目次のリンク：本文の中でなめらかに移り、畳んだ目次は閉じる（URL は変えない）
document.addEventListener('click', (e) => {
  const a = e.target.closest && e.target.closest('a[data-toc]');
  if(!a) return;
  e.preventDefault();
  const h = document.getElementById(a.dataset.toc);
  if(h) h.scrollIntoView({ behavior: 'smooth', block: 'start' });
  const d = a.closest('details');
  if(d) d.open = false;
});
{
  const sc = document.getElementById('learn-content');
  let ticking = false;
  if(sc) sc.addEventListener('scroll', () => {
    if(ticking) return;
    ticking = true;
    requestAnimationFrame(() => { ticking = false; updateReadProgress(); });
  }, { passive: true });
  const fb = document.getElementById('fs-btn');
  if(fb) fb.addEventListener('click', cycleFontSize);
  applyFontSize();
}
