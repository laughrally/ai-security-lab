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


