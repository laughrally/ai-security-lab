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

window.loadCorpApplications = loadCorpApplications;
window.sendCorpInvoice      = sendCorpInvoice;

