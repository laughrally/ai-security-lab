/* ===== 法人一括申込 =====
   受講者情報を受け取り、Worker 経由で受講申込確認書PDFを発行する。
   金額とプラン表記の定義元は Worker の CORP_PLAN。ここは表示用の写し。 */
const CORP_PLANS = {
  basic:    { label: 'Basicプラン',    price: 9800   },
  standard: { label: 'Standardプラン', price: 59800  },
  premium:  { label: 'Premiumプラン',  price: 149800 },
};
const CORP_MAX_SEATS = 50;

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

window.openCorpModal  = openCorpModal;
// https://ai-security-lab.com/#corp で法人申込を直接開く（営業メールのリンク用）
// このモジュールは load のあとに走り終えることがあるので、済んでいればその場で開く
function openCorpFromHash(){ if (location.hash === '#corp') { try { openCorpModal(); } catch (e) {} } }
if (document.readyState === 'complete') openCorpFromHash();
else window.addEventListener('load', openCorpFromHash);
window.closeCorpModal = closeCorpModal;
window.corpAddRow     = corpAddRow;
window.corpRecalc     = corpRecalc;
window.corpSubmit     = corpSubmit;


