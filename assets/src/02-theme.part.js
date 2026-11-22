/* ===== ダークモード ===== */
function toggleDark(){
  const isDark = document.body.classList.toggle('dark');
  document.getElementById('dark-toggle').textContent = isDark ? '🌙' : '☀️';
  store.set('theme', isDark ? 'dark' : 'light');
}
// 前回の設定は index.html の先頭で本文の描画前に当てている。ここではボタンの表示だけ合わせる
{ const btn = document.getElementById('dark-toggle'); if(btn && document.body.classList.contains('dark')) btn.textContent = '🌙'; }

