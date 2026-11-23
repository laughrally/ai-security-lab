// ===================== 管理者：受講ログのエクスポート =====================
// #admin-dashboard は既定で display:none。管理者アカウントでログインしたときのみ表示する。
// 管理者はアカウントIDで判定する（Worker の ADMIN_USER_IDS と同じ）。表示を切り替えるだけで、権限はサーバーが判定する
const ADMIN_USER_IDS = ['0fc82602-6512-41b5-9ef8-6a14b533d7f7'];
function isAdminUser(){ return !!(currentUser && ADMIN_USER_IDS.includes(currentUser.id)); }

function applyAdminVisibility(){
  const sec = document.getElementById('admin-dashboard');
  if(!sec) return;
  sec.style.display = isAdminUser() ? 'block' : 'none';
  if(isAdminUser() && !window._mfaChecked){ window._mfaChecked = true; adminMfaChallengeIfNeeded().then(adminMfaRefresh); }
}
// ===== 管理者の二段階認証（2026-11-22） =====
let _mfaEnrollId = null;
async function adminMfaRefresh(){
  const st = document.getElementById('admin-mfa-status');
  const startBtn = document.getElementById('admin-mfa-start');
  if(!st || !_supabase || !isAdminUser()) return;
  try {
    const { data } = await _supabase.auth.mfa.listFactors();
    const verified = (data && data.totp || []).filter(f => f.status === 'verified');
    const { data: aal } = await _supabase.auth.mfa.getAuthenticatorAssuranceLevel();
    if(verified.length){
      st.innerHTML = '✅ 登録済みです。管理者の操作は、ログイン時に認証アプリのコードを入れた場合だけ使えます。' +
        (aal && aal.currentLevel === 'aal2' ? '<br>（いまのログインはコード入力済みです）' : '<br>⚠️ いまのログインはコード未入力です。管理者の操作を使うには、ページを再読み込みしてコードを入れてください。');
      if(startBtn) startBtn.style.display = 'none';
    } else {
      st.textContent = '未登録です。登録すると、パスワードが漏れても管理者の操作（レッスン本文の投入など）は使えなくなります。';
      if(startBtn) startBtn.style.display = '';
    }
  } catch(e) {
    st.textContent = '状態を確認できませんでした：' + (e.message || '');
  }
}
async function adminMfaEnroll(){
  try {
    const { data, error } = await _supabase.auth.mfa.enroll({ factorType: 'totp', friendlyName: 'ASL admin ' + new Date().toISOString().slice(0,19) });
    if(error) throw error;
    _mfaEnrollId = data.id;
    document.getElementById('admin-mfa-qr').src = data.totp.qr_code;
    document.getElementById('admin-mfa-secret').textContent = data.totp.secret;
    document.getElementById('admin-mfa-enroll').style.display = '';
    document.getElementById('admin-mfa-start').style.display = 'none';
  } catch(e) { alert('登録を始められませんでした：' + (e.message || '')); }
}
async function adminMfaVerify(){
  const code = (document.getElementById('admin-mfa-code').value || '').trim();
  if(!/^\d{6}$/.test(code)){ alert('6桁の数字を入れてください'); return; }
  try {
    const { error } = await _supabase.auth.mfa.challengeAndVerify({ factorId: _mfaEnrollId, code });
    if(error) throw error;
    document.getElementById('admin-mfa-enroll').style.display = 'none';
    alert('登録しました。次回から、ログイン後に認証アプリのコードを入れてください。');
    adminMfaRefresh();
  } catch(e) { alert('コードを確認できませんでした：' + (e.message || '')); }
}
// 登録済みの管理者がコード未入力でログインしているときは、コードを求める
async function adminMfaChallengeIfNeeded(){
  if(!_supabase || !isAdminUser()) return;
  try {
    const { data: aal } = await _supabase.auth.mfa.getAuthenticatorAssuranceLevel();
    if(!aal || aal.nextLevel !== 'aal2' || aal.currentLevel === 'aal2') return;
    const { data } = await _supabase.auth.mfa.listFactors();
    const f = (data && data.totp || []).find(x => x.status === 'verified');
    if(!f) return;
    const code = prompt('管理者の操作には、認証アプリの6桁のコードが必要です。コードを入れてください（受講だけなら「キャンセル」で構いません）');
    if(!code) return;
    const { error } = await _supabase.auth.mfa.challengeAndVerify({ factorId: f.id, code: code.trim() });
    if(error) alert('コードを確認できませんでした：' + error.message);
    adminMfaRefresh();
  } catch(e) { console.warn('二段階認証の確認に失敗:', e.message); }
}

