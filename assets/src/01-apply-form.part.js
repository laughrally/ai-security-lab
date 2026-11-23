/* ===== 申込フォーム ===== */
const PRICE_ID_MAP = {
  'ベーシック 9800円':    'price_1TJYfK2OQszM8w2hEWWM5R36',
  'スタンダード 59800円': 'price_1TP0562OQszM8w2hbUgLZ5if',
  'プレミアム 149800円':   'price_1TP05i2OQszM8w2hhNpHkmfG',
};
let currentPriceId = '';
// 申込フォームのプラン名（決済のメタデータにも入るので文字は変えない）→ course.json のプラン
const FORM_PLAN_KEY = { 'ベーシック 9800円': 'basic', 'スタンダード 59800円': 'standard', 'プレミアム 149800円': 'premium' };
const FORM_PLAN_NAME = Object.fromEntries(Object.entries(FORM_PLAN_KEY).map(([k, v]) => [v, k]));
const yen = n => '¥' + Number(n).toLocaleString('ja-JP');
let currentPlanName = '';

function toggleSubmitBtn(){
  const checked = document.getElementById('form-agree').checked;
  const btn = document.getElementById('form-submit-btn');
  btn.disabled = !checked;
  btn.style.opacity = checked ? '1' : '0.5';
  btn.style.cursor = checked ? 'pointer' : 'not-allowed';
}

function openFormModal(plan){
  document.getElementById('form-overlay').classList.remove('hidden');
  currentPriceId = PRICE_ID_MAP[plan] || '';
  currentPlanName = plan;
  // 表記は course.json の価格と受講範囲から作る
  const key = FORM_PLAN_KEY[plan];
  const range = key ? '第1〜' + COURSE.plans[key] + '章・' + (key === 'premium' ? '全' : '') + PLAN_LESSON_COUNT[key] + 'レッスン' : '';
  document.getElementById('form-plan-label').textContent = key ? key.toUpperCase() + ' · ' + yen(COURSE.prices[key]) : plan;
  document.getElementById('legal-plan-name').textContent = key ? key.toUpperCase() + 'プラン（' + range + '）' : plan;
  document.getElementById('legal-plan-price').textContent = key ? yen(COURSE.prices[key]) + '（税込）' : '';
  // 同意チェックをリセット
  document.getElementById('form-agree').checked = false;
  toggleSubmitBtn();
  setFormMsg('','');
  // ベーシック以外はコンサルフィールドを表示
  const isConsult = key !== 'basic';
  document.getElementById('form-consult-fields').style.display = isConsult ? 'block' : 'none';
  // フィールドリセット
  ['form-name','form-company','form-tel','form-email','form-challenge'].forEach(id=>{
    const el = document.getElementById(id);
    if(el) el.value = '';
  });
  document.getElementById('form-contact-email').checked = true;
}
function closeFormModal(){
  document.getElementById('form-overlay').classList.add('hidden');
}
function setFormMsg(type, text){
  const el = document.getElementById('form-msg');
  el.className = 'form-msg' + (type ? ' '+type : '');
  el.textContent = text;
}

async function submitForm(){
  const name    = document.getElementById('form-name').value.trim();
  const company = document.getElementById('form-company').value.trim();
  const tel     = document.getElementById('form-tel').value.trim();
  const email   = document.getElementById('form-email').value.trim();
  const isConsult = FORM_PLAN_KEY[currentPlanName] !== 'basic';
  const challenge = isConsult ? document.getElementById('form-challenge').value.trim() : '';
  const contactPref = isConsult ? (document.querySelector('input[name="form-contact"]:checked')?.value || 'メール') : '';

  if(!name || !tel || !email){
    setFormMsg('error','お名前・電話番号・メールアドレスは必須です。');
    return;
  }
  if(!/^[^@]+@[^@]+\.[^@]+$/.test(email)){
    setFormMsg('error','メールアドレスの形式が正しくありません。');
    return;
  }
  if(!currentPriceId){
    setFormMsg('error','プランの選択でエラーが発生しました。再度お試しください。');
    return;
  }

  const btn = document.getElementById('form-submit-btn');
  btn.disabled = true;
  btn.textContent = '処理中...';
  setFormMsg('','');

  try {
    const res = await fetch(WORKER_URL + '/create-checkout', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        priceId: currentPriceId,
        email,
        metadata: { name, company, tel, challenge, contactPref, plan: currentPlanName }
      }),
    });
    const data = await res.json();
    if(data.url){
      window.location.href = data.url;
    } else {
      throw new Error(data.error || '決済画面の取得に失敗しました。');
    }
  } catch(e) {
    setFormMsg('error', e.message || '送信に失敗しました。時間をおいて再度お試しください。');
    btn.disabled = false;
    btn.textContent = '決済画面へ進む →';
  }
}

/* フォームの外側クリックで閉じる */
document.getElementById('form-overlay').addEventListener('click', function(e){
  if(e.target === this) closeFormModal();
});

