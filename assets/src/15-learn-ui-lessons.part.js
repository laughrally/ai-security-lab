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
 * L02〜L34 の本文は index.html に含めない。
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
// 初めて学習画面に入った人への案内（2026-11-23）。プランごとの進め方と修了の条件、最初のレッスンへのボタン。
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

// 理解度チェックの答え合わせを読み上げで伝える（本文は DB から来るので、表示のたびに印を付ける。2026-11-23）
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
// プランで開けるタブだけをたどる（2026-11-22）。以前は ALL_TABS_ORDERED をそのままたどり、
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

// レッスン本文の末尾にある前後ナビを、受講者のプランに合わせて作り直す。
// 本文（DB）のナビは章末ワークを飛ばして次のレッスンを指しているため（2026-11-22）
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
  // 範囲の終わりでは、上のプランの案内を開けるようにする（2026-11-23。以前は文字だけで押せなかった）
  else if(currentUserPlan === 'basic') html += `<a href="#lessons" onclick="switchTab('records');return false;" class="lnbtn" style="text-align:right"><span class="lnl">🎉 Basic の全レッスンはここまで</span><span class="lnt">受講実績レポートを出力する →</span></a>`
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
  // ログイン直後の最初の表示だけ、前回開いていたタブから再開する（以前は毎回序章が開いた。2026-11-22）
  if(!window._aslResumed && currentUserPlan){
    window._aslResumed = true;
    try {
      const last = store.getJSON('asl_last_tab', null);
      if(last && last.u === currentUser.id && last.t && last.t !== currentTab && tabOpenable(last.t)){
        currentTab = last.t;
        setTimeout(() => window.switchTab(last.t), 0);   // 中身の読み込み（ワーク・テスト・修了証）を通常の経路で行う
      } else if(!last && completed.length){
        // 前回の記録が端末に無い（別の端末・ブラウザ）ときは、受講順で最初の未了から開く（2026-11-23。以前は序章が開いた）
        // 受講記録（ワーク・テストの状況）を受け取ってから決める（2026-11-23。以前は受け取る前に決めて、開く場所が毎回変わった）。
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
  // 修了条件のうちレッスン以外（ワーク・テスト・試験）も一目でわかるように（2026-11-23）
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
      html += `<div class="sb-row${active?' active':''}${done?' done':''}" tabindex="0" role="button" onclick="switchTab('${id}');closeSbDrawer()">
        <div class="sb-row-dot"></div>
        <span class="sb-row-num">${num}</span>
        <span class="sb-row-title">${title}</span>
      </div>`;
    });
    // 章末ワーク行（序章の擬似章には付けない）
    const chNo = parseInt((ch.tag||'').replace(/[^0-9]/g,''), 10);
    // Basic には章末ワーク・テストが無いので行を出さない（2026-11-23。以前は押せない行が並んでいた）
    if(chNo >= 1 && WORK_CHAPTERS[chNo] && (COURSE.workChapters[currentUserPlan] || 0) > 0){
      const wid = 'work' + chNo;
      const wactive = wid === currentTab;
      const wdone = !!_workSubmittedCh[chNo];
      html += `<div class="sb-row sb-row-work${wactive && window._aslTestFocus !== chNo?' active':''}${wdone?' done':''}" tabindex="0" role="button" onclick="openWorkTab(${chNo});closeSbDrawer()">
        <div class="sb-row-dot"></div>
        <span class="sb-row-num">W${chNo}</span>
        <span class="sb-row-title">章末ワーク：${WORK_CHAPTERS[chNo].goal}</span>
      </div>`;
      // 確認テストを別の行に（2026-11-22）。以前はワークと同じ1行だけで、テストを見落としやすかった
      const tpass = !!(_testResultCh[chNo] && _testResultCh[chNo].passed);
      // 5問すべて答えて不合格なら「再挑戦」の印（未受験と見分けられるように）
      const tfail = !tpass && !!(_testResultCh[chNo] && _testResultCh[chNo].answered >= TEST_TOTAL_PER_CHAPTER);
      const tactive = wactive && window._aslTestFocus === chNo;
      html += `<div class="sb-row sb-row-work${tactive?' active':''}${tpass?' done':''}${tfail?' test-failed':''}" tabindex="0" role="button" onclick="goChapterTest(${chNo});closeSbDrawer()">
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
      div.innerHTML = '<div class="sb-ch-label">認定試験</div><div class="sb-row" onclick="switchTab(&#39;exam&#39;);closeSbDrawer()" style="color:var(--purple);font-weight:600"><div class="sb-row-dot" style="background:var(--purple)"></div><span class="sb-row-num">EX</span><span class="sb-row-title">AI Security Lab 認定試験</span></div>';
      list.appendChild(div);
    }
  }

  // 修了証の申請（Standard 以上）。サイドバーの最後に常に出す
  if((currentUserPlan === 'standard' || currentUserPlan === 'premium') && !list.querySelector('[data-cert]')){
    const div = document.createElement('div');
    div.setAttribute('data-cert','1');
    const cactive = currentTab === 'certificate';
    div.innerHTML = '<div class="sb-ch-label">修了</div><div class="sb-row' + (cactive ? ' active' : '') + '" onclick="switchTab(&#39;certificate&#39;);closeSbDrawer()" style="font-weight:600"><div class="sb-row-dot"></div><span class="sb-row-num">🎓</span><span class="sb-row-title">修了証の申請</span></div>';
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
      // 章末ワーク・修了証・試験・記録のページでは、前のレッスンの表示を残さない（2026-11-22）
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
// 進捗パネルの章ごとの「章末ワーク・確認テスト」の状態（2026-11-23。以前はレッスンだけだった）
function wtLine(ch){
  const no = parseInt((ch.tag || '').replace(/[^0-9]/g, ''), 10);
  if(!no || !WORK_CHAPTERS[no] || no > (COURSE.workChapters[currentUserPlan] || 0)) return '';
  const sm = window._recSummary || {};
  const w = !!_workSubmittedCh[no] || (sm.workDoneList || []).includes(no);
  const t = !!(_testResultCh[no] && _testResultCh[no].passed) || (sm.testPassedList || []).includes(no);
  return `<div class="sb-dash-wt"><button type="button" class="sb-dash-wt-item${w ? ' done' : ''}" onclick="switchTab('work${no}')">章末ワーク ${w ? '✓ 提出済み' : '未提出'}</button>`
    + `<button type="button" class="sb-dash-wt-item${t ? ' done' : ''}" onclick="goChapterTest(${no})">確認テスト ${t ? '✓ 合格' : '未合格'}</button></div>`;
}
function buildProgressPanel(){
  const panel = document.getElementById('sb-progress-panel');
  if(!panel) return;
  const completed = loadCompleted();
  let html = '';
  CHAPTER_DATA.forEach(ch => {
    // プランの外の章は、鍵付きで1行だけ出す（2026-11-23。以前は Standard でも第7章が「0/7」と鍵なしで出た）
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
      return `<div class="sb-dash-dot${d?' done':''}" tabindex="0" role="button" onclick="switchTab('${id}');setSbTab('lessons')" title="${LESSON_LABELS[id]||id}">${num}</div>`;
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
  // 修了証への導線（Standard 以上）。以前はサイドバーにも進捗にも行が無く、Standard の人がたどり着けなかった
  if(currentUserPlan === 'standard' || currentUserPlan === 'premium'){
    html += `<div class="sb-ch-label" style="margin-top:14px">修了</div>
    <div class="sb-row" onclick="switchTab('certificate');closeSbDrawer()">
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
    <div class="sb-row" onclick="switchTab('records');closeSbDrawer()">
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
// { 設問ID: 正解した元の選択肢番号 }。共用PCで前の人の正解が混ざらないよう、ユーザーごとに分ける（2026-11-22）
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
    // 正解済みなのに「受講済み」が押せなくなる不具合の修正（2026-11-22）。
    // 以前は選択肢だけを閉じ、受講済みボタンは閉じたままだった
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
// キーボードで操作できるように（2026-11-22）：クリックで動く div にフォーカスを与え、Enter / Space で押せるようにする
document.addEventListener('DOMContentLoaded', ()=>{
  document.querySelectorAll('.lrow[onclick], .chd').forEach(el => { el.tabIndex = 0; el.setAttribute('role','button'); });
});
// トップページが長すぎる（序章の全文が LP の中にある）ので、未ログインでは途中で畳む（2026-11-22）
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

// ダイアログを Esc で閉じる・読み上げ用の印を付ける（2026-11-22）
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



