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
window.setSbTab = setSbTab;

