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

