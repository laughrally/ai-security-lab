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



