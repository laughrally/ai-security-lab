/* ===== 序章のポッドキャスト（自前プレーヤー） =====
 * Google Drive の埋め込みは別ドメインのiframeで見た目を一切制御できず、
 * 再生開始まで5〜10秒かかっていた。音声を audio/intro.mp3 として同梱し、
 * CSSに用意済みだった .pod-* のデザインで自前のプレーヤーを組む。
 */
const POD_KEY = 'asl_pod_intro';   // 再生位置の保存先
function podAudio(){ return document.getElementById('pod-audio'); }

function podFmt(sec){
  if(!isFinite(sec) || sec < 0) return '--:--';
  const s = Math.floor(sec);
  return Math.floor(s/60) + ':' + String(s%60).padStart(2,'0');
}

function podToggle(){
  const a = podAudio();
  if(!a) return;
  if(a.paused){
    a.play().catch(e => {
      const t = document.getElementById('pod-cur');
      if(t) t.textContent = '再生できません';
      console.warn('音声を再生できませんでした:', e.message);
    });
  } else {
    a.pause();
  }
}
function podSeek(v){
  const a = podAudio();
  if(a && isFinite(a.duration)) a.currentTime = a.duration * (v / 1000);
}
function podVol(v){
  const a = podAudio();
  if(!a) return;
  a.volume = Number(v);
  a.muted = Number(v) === 0;
  podSyncMute();
}
function podMute(){
  const a = podAudio();
  if(!a) return;
  a.muted = !a.muted;
  podSyncMute();
}
function podSyncMute(){
  const a = podAudio(), btn = document.getElementById('pod-mute');
  if(!a || !btn) return;
  btn.textContent = (a.muted || a.volume === 0) ? '🔇' : '🔊';
}
function podSpeed(rate, btn){
  const a = podAudio();
  if(a) a.playbackRate = rate;
  document.querySelectorAll('.pod-spd').forEach(b => b.classList.remove('active'));
  if(btn) btn.classList.add('active');
}

function initPodcast(){
  const a = podAudio();
  if(!a || a.dataset.wired === '1') return;
  a.dataset.wired = '1';
  const seek = document.getElementById('pod-seek');
  const cur  = document.getElementById('pod-cur');
  const dur  = document.getElementById('pod-dur');
  const play = document.getElementById('pod-play');

  const onMeta = () => {
    if(dur) dur.textContent = podFmt(a.duration);
    // 前回の続きから（最後まで聴き終えていた場合は先頭に戻す）
    try{
      const saved = parseFloat(localStorage.getItem(POD_KEY) || '0');
      if(saved > 0 && isFinite(a.duration) && saved < a.duration - 5) a.currentTime = saved;
    }catch(e){ /* localStorageが使えない環境は無視 */ }
  };
  a.addEventListener('loadedmetadata', onMeta);
  // preload="metadata" のため、この関数が動く時点で既に読み込み済みのことがある。
  // その場合 loadedmetadata は二度と発火せず、長さが '--:--' のままになる。
  if(a.readyState >= 1) onMeta();
  a.addEventListener('timeupdate', () => {
    if(cur) cur.textContent = podFmt(a.currentTime);
    if(seek && isFinite(a.duration) && a.duration > 0){
      seek.value = String(Math.round(a.currentTime / a.duration * 1000));
    }
    // 5秒ごとに保存する。毎回書くとtimeupdateの頻度ぶん無駄に書き込む。
    if(Math.floor(a.currentTime) % 5 === 0){
      try{ localStorage.setItem(POD_KEY, String(a.currentTime)); }catch(e){ /* noop */ }
    }
  });
  a.addEventListener('play',  () => { if(play){ play.textContent = '❚❚'; play.setAttribute('aria-label','一時停止'); } });
  a.addEventListener('pause', () => { if(play){ play.textContent = '▶';  play.setAttribute('aria-label','再生'); } });
  a.addEventListener('ended', () => {
    if(play) play.textContent = '▶';
    if(seek) seek.value = '0';
    try{ localStorage.removeItem(POD_KEY); }catch(e){ /* noop */ }
  });
  a.addEventListener('error', () => {
    if(cur) cur.textContent = '読み込み失敗';
    console.warn('音声ファイルを読み込めませんでした: audio/intro.mp3');
  });
  podSyncMute();
}
window.podToggle = podToggle;
window.podSeek   = podSeek;
window.podVol    = podVol;
window.podMute   = podMute;
window.podSpeed  = podSpeed;

/**
 * ヘッダーのロゴ・カリキュラム・レッスンの遷移先。
 *
 * 未ログイン時は今までどおり LP のアンカーへ飛ばす（return true）。
 * ログイン後は #learn-wrapper（position:fixed・画面全体・z-index 50）が
 * LPを覆っているため、ページをスクロールさせても見た目が一切変わらない。
 * さらに #top と #lessons は updateAuthUI / showLearnUI が display:none に
 * するので、アンカーとしても機能しない。
 * そこでログイン中は学習UI側の操作へ振り替える。
 */
function navGo(where){
  if(!currentUser) return true;   // 未ログインは通常のアンカー動作

  if(where === 'curriculum'){
    // サイドバーの「📖 学習」タブ＝レッスン一覧がカリキュラムそのもの
    const lessonsTab = document.querySelector('#sb-tabs .sb-tab:nth-child(1)');
    if(lessonsTab) setSbTab('lessons', lessonsTab);
    if(window.innerWidth <= 860) openSbDrawer();
    showLearnUI();
    return false;
  }

  if(where === 'top'){
    // ロゴ＝学習の先頭（序章）へ戻す
    window.switchTab('intro');
    return false;
  }

  // 'lessons'：今開いているレッスンの先頭まで戻す
  showLearnUI();
  const lc = document.getElementById('learn-content');
  if(lc) lc.scrollTo({ top: 0, behavior: 'smooth' });
  if(window.innerWidth <= 860) closeSbDrawer();
  return false;
}
window.navGo = navGo;

// 進捗ボタン → サイドバーの進捗タブを開く
function openProgressPanel(){
  // PC: サイドバーの進捗タブをアクティブに
  const progressTab = document.querySelector('#sb-tabs .sb-tab:nth-child(2)');
  if(progressTab) setSbTab('progress', progressTab);
  // スマホ: ドロワーを開く
  if(window.innerWidth <= 860){
    openSbDrawer();
  }
  // learn-wrapperが表示されていなければ表示
  showLearnUI();
}
window.openProgressPanel = openProgressPanel;

// モバイル：ドロワー開閉
function toggleSbDrawer(){
  const sb = document.getElementById('learn-sidebar');
  const ov = document.getElementById('sb-overlay');
  if(!sb) return;
  const isOpen = sb.classList.contains('open');
  if(isOpen){ closeSbDrawer(); } else { openSbDrawer(); }
}
function openSbDrawer(){
  const sb = document.getElementById('learn-sidebar');
  const ov = document.getElementById('sb-overlay');
  if(sb) sb.classList.add('open');
  if(ov) ov.classList.add('show');
}
function closeSbDrawer(){
  const sb = document.getElementById('learn-sidebar');
  const ov = document.getElementById('sb-overlay');
  if(sb) sb.classList.remove('open');
  if(ov) ov.classList.remove('show');
}

// カリキュラムからレッスンへ飛ぶ共通関数
function goLesson(tabId){
  if(!canAccessTab(tabId)) return;
  if(currentUser){
    // ログイン後：先にUIを表示してからレッスンをマウント
    showLearnUI();
    window.switchTab(tabId);   // 前回のタブの記録・章末ワークの読み込みも同じ経路で行う
    window.scrollTo({top: 0, behavior: 'smooth'});
  } else {
    showTabGuest(tabId);
    const s = document.getElementById('lessons');
    if(s) s.scrollIntoView({behavior:'smooth'});
  }
}
window.goLesson = goLesson;

