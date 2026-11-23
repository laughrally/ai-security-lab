// ===== 認定試験UI =====
let examState = {
  active: false,
  answers: {},   // 出題インデックス: 「元の」選択肢番号（表示順ではない）
  questions: [], // 受験ごとにシャッフルした問題セット
  timer: null,
  remaining: COURSE.rules.examMinutes * 60,
  submitted: false,
  sending: false,  // 提出中の二重送信ガード
  startedAt: null, // 開始時刻（受験時間の実測用）
};

// Fisher-Yatesシャッフル
function shuffled(arr){
  const a = arr.slice();
  for(let i = a.length - 1; i > 0; i--){
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

/**
 * 受験ごとに問題順と選択肢順をランダム化する。
 * 正解の位置が固定されていると、内容を理解していなくても
 * パターンで正答できてしまうため、毎回並びを変える。
 *
 * ※ 正解はクライアントに存在しない。各選択肢は「元の選択肢番号」（orig）を
 *    保持しており、表示順を変えてもサーバーへ送る値は変わらない。
 *    レッスン内クイズの shuffleQuizOptions と同じ考え方。
 */
function buildExamSet(){
  return shuffled(EXAM_QUESTIONS).map(q => ({
    id: q.id,
    // 設問は内部ID（lesson10 など）で持ち、表示のときに受講順の番号（L10）にする（2026-11-23）。
    // 以前は表示番号で持っており、レッスンを並べ替えると古い番号のまま残った
    lesson: q.lesson === 'intro' ? '序章' : 'L' + lessonNo(q.lesson),
    q: q.q,
    o: shuffled(q.o.map((text, orig) => ({ text, orig }))),
  }));
}

function fmtTime(sec){
  const m = Math.floor(sec/60), s = sec%60;
  return m + ':' + String(s).padStart(2,'0');
}

async function startExam(){
  if(currentUserPlan !== 'premium'){
    alert('認定試験はPremiumプランの受講者のみ受験できます。');
    return;
  }
  // 開始前に、受験できるか（24時間の間隔）を確かめる。提出時にしか確認しないと、60分かけた解答が弾かれる
  const status = async () => {
    try { return await apiFetch('/api/exam/status', { method: 'GET' }); }
    catch(e) { alert('受験の準備ができませんでした。通信環境をご確認のうえ、もう一度お試しください。\n' + (e.message || '')); return null; }
  };
  let st = await status();
  if(!st) return;
  if(st.canStart === false && st.reason === 'cooldown'){
    alert('次の受験まで、あと約' + st.retryAfterHours + '時間お待ちください。\n\n' +
          '2回目以降の受験は、前回から24時間の間隔をいただいています。\n' +
          'これまでの受験回数：' + st.attempts + '回');
    return;
  }
  const R = COURSE.rules;
  const confirmed = confirm(`認定試験を開始しますか？\n\n制限時間: ${R.examMinutes}分\n問題数: ${R.examQuestions}問\n合格ライン: ${R.examPass}問以上（${Math.round(R.examPass / R.examQuestions * 100)}%）\n\n開始後はカウントダウンが始まります。`);
  if(!confirmed) return;
  // 開始の署名は「開始する」を押したあとに受け取る（提出のときにサーバーが制限時間を確かめるのに使う。2026-11-23）
  st = await status();
  if(!st || !st.start) { if(st) alert('受験を開始できませんでした。時間をおいてもう一度お試しください。'); return; }
  examState.start = st.start;
  if(typeof window.switchTab === 'function') window.switchTab('exam');
  examState.active = true;
  examState.answers = {};
  examState.questions = buildExamSet();  // 受験ごとに問題順・選択肢順を変える
  examState.remaining = COURSE.rules.examMinutes * 60;
  examState.submitted = false;
  examState.sending = false;
  examState.startedAt = Date.now();
  renderExam();
  startExamTimer();
  window.scrollTo({top:0, behavior:'smooth'});
}

function startExamTimer(){
  if(examState.timer) clearInterval(examState.timer);
  // 残り時間は「制限時間 −（今 − 開始時刻）」で毎回計算する（2026-11-23）。以前は1秒ずつ減らすだけで、
  // タブを裏に回したときのタイマーの間引きやスリープで表示が遅れ、サーバーの判定（開始の署名からの実時間）とずれた
  examState.timer = setInterval(() => {
    examState.remaining = Math.max(0, COURSE.rules.examMinutes * 60 - Math.floor((Date.now() - examState.startedAt) / 1000));
    const timerEl = document.getElementById('examTimer');
    if(timerEl) timerEl.textContent = fmtTime(examState.remaining);
    if(examState.remaining <= 0){
      clearInterval(examState.timer);
      submitExam();
    }
  }, 1000);
}

function renderExam(){
  const container = document.getElementById('examContainer');
  if(!container) return;
  let html = `
    <div style="position:sticky;top:0;background:var(--bg);padding:16px 0;border-bottom:2px solid var(--purple);z-index:10;margin-bottom:24px">
      <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px">
        <div style="font-weight:700;font-size:18px">🏆 AI Security Lab 認定試験</div>
        <div style="display:flex;gap:16px;align-items:center">
          <div style="font-weight:600;color:var(--text2)">回答済み <span id="examAnswered">${Object.keys(examState.answers).length}</span> / ${examState.questions.length}</div>
          <div style="font-weight:700;color:var(--purple)">⏰ <span id="examTimer">${fmtTime(examState.remaining)}</span></div>
          <button onclick="submitExam()" style="background:var(--red-s);color:#fff;border:none;border-radius:8px;padding:10px 20px;font-weight:700;cursor:pointer">提出する</button>
        </div>
      </div>
    </div>
  `;
  const QS = examState.questions;
  QS.forEach((q, i) => {
    html += `<div id="exq-${i}" style="margin-bottom:32px;padding:20px;background:var(--bg2);border-radius:12px;border:1px solid var(--border)">
      <div style="font-size:12px;color:var(--text3);margin-bottom:8px">問題 ${i+1} / ${QS.length}　[${q.lesson}]</div>
      <div style="font-weight:700;margin-bottom:16px;line-height:1.6">${q.q}</div>
      <div style="display:flex;flex-direction:column;gap:8px">`;
    q.o.forEach((opt, j) => {
      // value / onchange に渡すのは「元の選択肢番号」。表示順(j)ではない。
      const selected = examState.answers[i] === opt.orig;
      html += `<label style="display:flex;align-items:center;gap:12px;padding:12px;border:2px solid ${selected ? 'var(--purple)' : 'var(--border)'};border-radius:8px;cursor:pointer;background:${selected ? 'rgba(124,58,237,0.05)' : 'transparent'}">
        <input type="radio" name="q${i}" value="${opt.orig}" ${selected ? 'checked' : ''} onchange="selectAnswer(${i},${opt.orig})" style="cursor:pointer">
        <span>${opt.text}</span>
      </label>`;
    });
    html += `</div></div>`;
  });
  html += `<div style="text-align:center;margin:40px 0">
    <button onclick="submitExam()" style="background:var(--purple-s);color:#fff;border:none;border-radius:12px;padding:16px 48px;font-size:16px;font-weight:700;cursor:pointer">✅ 試験を提出する</button>
  </div>`;
  container.innerHTML = html;
  container.style.display = 'block';
  const introEl = document.getElementById('examIntro');
  if(introEl) introEl.style.display = 'none';
}

// 選んだ設問の見た目と回答数だけを更新する（2026-11-23。以前は1つ選ぶたびに60問すべてを描き直していた）
function selectAnswer(qIdx, optIdx){
  examState.answers[qIdx] = optIdx;
  document.querySelectorAll(`#exq-${qIdx} label`).forEach(l => {
    const on = Number(l.querySelector('input').value) === optIdx;
    l.style.borderColor = on ? 'var(--purple)' : 'var(--border)';
    l.style.background = on ? 'rgba(124,58,237,0.05)' : 'transparent';
  });
  const cnt = document.getElementById('examAnswered');
  if(cnt) cnt.textContent = Object.keys(examState.answers).length;
}

/**
 * 試験の提出。採点はサーバー側で行う。
 * クライアントは「どの設問でどの選択肢（元の番号）を選んだか」だけを送り、
 * 得点・合否・正解・解説を受け取って表示する。
 * 合格フラグ（profiles.exam_passed）もサーバーが立てる。
 */
async function submitExam(){
  if(examState.sending) return;
  const QS = examState.questions;
  if(!QS.length) return;
  const answered = Object.keys(examState.answers).length;
  if(answered < QS.length && examState.remaining > 0){
    const unanswered = QS.length - answered;
    if(!confirm(`${unanswered}問が未回答です。このまま提出しますか？`)) return;
  }
  // 提出中はタイマーを止める。失敗したら残り時間があるぶんだけ再開する。
  if(examState.timer){ clearInterval(examState.timer); examState.timer = null; }
  examState.sending = true;

  // 設問ID → 元の選択肢番号
  const answers = {};
  QS.forEach((q, i) => {
    const v = examState.answers[i];
    if(Number.isInteger(v)) answers[q.id] = v;
  });

  const container = document.getElementById('examContainer');
  if(container) container.innerHTML = '<div style="text-align:center;padding:80px 20px;font-size:16px;color:var(--text3)">採点しています…</div>';
  window.scrollTo({top:0, behavior:'smooth'});

  try {
    const res = await apiFetch('/api/exam/submit', {
      method: 'POST',
      body: JSON.stringify({
        answers,
        start: examState.start || null,
        startedAt: examState.startedAt ? new Date(examState.startedAt).toISOString() : null,
        durationSeconds: examState.startedAt ? Math.floor((Date.now() - examState.startedAt) / 1000) : null,
      }),
    });
    examState.submitted = true;
    examState.active = false;
    renderExamResult(res);
    if(res.passed){
      currentExamPassed = true;   // 修了証の申請ボタンを即時に解放する
      refreshRecords();
      updateCompletionUI();
      updateCertificateUI();
    }
  } catch(e) {
    if(e && e.status === 409){
      // 制限時間を過ぎた（サーバーの判定）。何度提出しても受け付けられないので、そのことを伝えて受験の画面に戻す
      alert((e.message || '制限時間を過ぎたため、この解答は受け付けられませんでした。') + '\n\nもう一度受験してください。');
      clearInterval(examState.timer);
      retryExam();
    } else if(e && e.status === 429){
      // 再受験の間隔制限。開始前にも確認しているので通常ここには来ないが、
      // 来た場合は解答を保持したまま試験画面に戻す（60分の労力を捨てない）。
      alert((e.message || '次の受験まで時間をおいてください。') +
            '\n\n解答はこの画面に保持しています。時間をおいてから、あらためて「提出する」を押してください。');
      renderExam();
    } else {
      alert('採点結果を受け取れませんでした。\n' + (e.message || '') + '\n\n解答は保持しています。通信環境をご確認のうえ、もう一度「提出する」を押してください。');
      renderExam();               // 解答を保持したまま試験画面に戻す
      if(examState.remaining > 0) startExamTimer();
    }
  } finally {
    examState.sending = false;
  }
  window.scrollTo({top:0, behavior:'smooth'});
}

/**
 * 結果画面。表示に使う設問文と選択肢はクライアント、
 * 正誤・正解・解説はサーバーのレスポンス（res.results）から取る。
 */
function renderExamResult(res){
  const container = document.getElementById('examContainer');
  if(!container) return;
  const QS = examState.questions;
  const resById = {};
  (res.results || []).forEach(r => { resById[r.id] = r; });

  const total   = res.total   != null ? res.total   : QS.length;
  const correct = res.correct != null ? res.correct : 0;
  const passed  = !!res.passed;
  const pct     = res.scorePct != null ? res.scorePct : Math.round(correct / (total || 1) * 100);
  const passMark = res.passMark != null ? res.passMark : 48;

  let html = `
    <div style="text-align:center;padding:40px 20px;background:${passed ? 'rgba(34,197,94,0.08)' : 'rgba(239,68,68,0.08)'};border:2px solid ${passed ? 'var(--green)' : 'var(--red)'};border-radius:16px;margin-bottom:32px">
      <div style="font-size:64px;margin-bottom:16px">${passed ? '🎉' : '😢'}</div>
      <div style="font-size:28px;font-weight:700;margin-bottom:8px;color:${passed ? 'var(--green)' : 'var(--red)'}">${passed ? '合格' : '不合格'}</div>
      <div style="font-size:20px;margin-bottom:16px">${correct} / ${total} 問正解（${pct}%）</div>
      <div style="color:var(--text3);font-size:14px">合格ライン: ${passMark}問以上（80%）${res.attemptNo ? '　／　受験回数: ' + res.attemptNo + '回目' : ''}</div>
      <div style="color:var(--text3);font-size:12px;margin-top:8px">採点は当社サーバーで行っており、受講者による自己採点ではありません。</div>
      ${passed ? `<div style="margin-top:24px"><button onclick="requestCertificate()" style="background:var(--purple-s);color:#fff;border:none;border-radius:12px;padding:14px 36px;font-size:15px;font-weight:700;cursor:pointer">📜 修了証を申請する</button></div>` : `<div style="margin-top:24px"><button onclick="retryExam()" style="background:var(--red-s);color:#fff;border:none;border-radius:12px;padding:14px 36px;font-size:15px;font-weight:700;cursor:pointer">🔄 受験の画面に戻る（再受験は ${nextExamAt()} ごろから）</button></div>`}
    </div>
  `;

  if(passed){
    // 合格者にだけ全問の解答と解説を出す。
    html += `<div style="font-weight:700;font-size:18px;margin-bottom:16px">📋 全問題の解答と解説</div>`;
    // 受験者が見た順（出題順）で並べる
    QS.forEach((q, idx) => {
      const r = resById[q.id];
      if(!r) return;
      const optText = n => {
        const hit = q.o.find(x => x.orig === n);
        return hit ? hit.text : '—';
      };
      const bg = r.isCorrect ? 'rgba(34,197,94,0.05)' : 'rgba(239,68,68,0.05)';
      const bc = r.isCorrect ? 'var(--green)' : 'var(--red)';
      html += `<div style="margin-bottom:16px;padding:16px;background:${bg};border-left:4px solid ${bc};border-radius:8px">
        <div style="font-size:12px;color:var(--text3);margin-bottom:4px">問題 ${idx+1} [${q.lesson}] ${r.isCorrect ? '✅ 正解' : '❌ 不正解'}</div>
        <div style="font-weight:700;margin-bottom:8px">${q.q}</div>
        <div style="font-size:14px;margin-bottom:4px"><strong>あなたの解答:</strong> ${r.selected !== null && r.selected !== undefined ? optText(r.selected) : '未回答'}</div>
        <div style="font-size:14px;margin-bottom:8px;color:var(--green)"><strong>正解:</strong> ${optText(r.correctIndex)}</div>
        <div style="font-size:13px;color:var(--text3);line-height:1.6"><strong>解説:</strong> ${r.explanation || ''}</div>
      </div>`;
    });
  } else {
    // 不合格のときは正解も解説も出さない。1回受ければ答えが全部手に入る状態だと
    // 2回目は必ず満点になり、修了証が受講を証明しなくなる。
    // サーバーも不合格時は correctIndex / explanation を返していない。
    // 復習の手がかりとして、誤答のあった章と問数だけを示す（2026-11-22）。
    // 設問ごとの正誤を返すと、受け直しの差分で正解が割り出せるため、サーバーも章単位でしか返さない。
    const wrongLessons = res.wrongByChapter || {};
    const keys = Object.keys(wrongLessons).sort((a, b) => (a === '序章' ? -1 : b === '序章' ? 1 : a.localeCompare(b, 'ja', {numeric:true})));
    html += `<div style="font-weight:700;font-size:18px;margin-bottom:12px">📚 復習が必要な章</div>
      <p style="color:var(--text3);font-size:14px;line-height:1.8;margin:0 0 16px">
        公平な認定のため、正解と解説、どの設問を間違えたかは表示していません。誤答のあった章を読み直してから、あらためて受験してください。
      </p>`;
    if(keys.length){
      html += `<div style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:20px">` +
        // 章を押すと、その章の最初のレッスンが開く（2026-11-23。以前はただの文字だった）
        keys.map(L => { const no = L === '序章' ? 0 : parseInt(L.replace(/[^0-9]/g, ''), 10); const first = no === 0 ? 'intro' : ((COURSE.chapters.find(c => c.no === no) || {}).lessons || [])[0];
          return `<button type="button" onclick="switchTab('${first}')" style="background:var(--red-l);border:1px solid var(--red);color:var(--red);border-radius:8px;padding:8px 14px;font-size:14px;font-weight:600;cursor:pointer;font-family:inherit">${L}<span style="font-weight:400">　${wrongLessons[L]}問 →</span></button>`; }).join('') +
        `</div>`;
    }
    html += `<p style="color:var(--text3);font-size:13px;line-height:1.8;margin:0">
      2回目以降の受験は、前回から24時間の間隔をいただいています（次に受験できるのは ${nextExamAt()} ごろから）。
    </p>`;
  }
  container.innerHTML = html;
}

// 次に受験できる日時（提出から24時間後。Worker の EXAM_COOLDOWN_MS と同じ）
function nextExamAt(){
  const d = new Date(Date.now() + 24 * 3600e3);
  return (d.getMonth() + 1) + '月' + d.getDate() + '日 ' + d.getHours() + ':' + String(d.getMinutes()).padStart(2, '0');
}
function retryExam(){
  examState = { active:false, answers:{}, questions:[], timer:null, remaining:COURSE.rules.examMinutes * 60, submitted:false, sending:false, startedAt:null };
  const container = document.getElementById('examContainer');
  if(container) container.style.display = 'none';
  const introEl = document.getElementById('examIntro');
  if(introEl) introEl.style.display = 'block';
  window.scrollTo({top:0, behavior:'smooth'});
}

/**
 * 修了証の発行申請。
 * 以前はGASのURLを直接叩いていたが、認証も受講状況の確認もないため
 * URLさえ知っていれば誰でも任意の氏名で発行できた。
 * 現在はWorkerが受講状況とプランをDBで検証してからGASを呼ぶ。
 */
async function submitCertificate(){
  const company = (document.getElementById('cert-company')?.value || '').trim();
  const name    = (document.getElementById('cert-name')?.value || '').trim();
  const msg     = document.getElementById('cert-msg');
  const btn     = document.getElementById('cert-submit-btn');

  if(!currentUser){ openAuthModal(); return; }
  if(!name){ alert('氏名を入力してください。'); return; }

  btn.disabled = true;
  btn.textContent = '送信中...';
  if(msg){ msg.style.display = 'none'; }

  try {
    const r = await apiFetch('/api/certificate', {
      method: 'POST',
      body: JSON.stringify({ name, company }),
    });
    if(msg){
      msg.style.display = 'block';
      msg.style.color = 'var(--green)';
      // サーバ応答をそのまま連結していたので、エスケープしてから埋める
      msg.innerHTML = '✅ 申請を受け付けました。修了証PDFを <strong>' + escHtml(r.email || 'ご登録のメールアドレス') +
                      '</strong> 宛にお送りします。<br>証明書番号：' + escHtml(r.certNo || '') ;
    }
    btn.textContent = '申請済み';
    window._certSubmitted = true;   // 画面の更新でボタンが押せる状態に戻らないように（二重申請の防止）
  } catch(e) {
    if(e.status === 409 || /発行済み/.test(e.message || '')){
      window._certSubmitted = true;
      btn.textContent = '発行済み';
      btn.style.opacity = '.4';
      if(msg){ msg.style.display = 'block'; msg.style.color = 'var(--text2)'; msg.textContent = e.message; }
      return;
    }
    btn.disabled = false;
    btn.textContent = '修了証を申請する';
    if(msg){
      msg.style.display = 'block';
      msg.style.color = 'var(--red)';
      msg.textContent = e.message || '送信に失敗しました。時間をおいて再試行してください。';
    }
  }
}

// 認定試験の合格画面から：修了証の申請欄を開き、氏名の入力欄へ移る（2026-11-23）。
// 以前はその場で送信しており、氏名が空のため必ず「氏名を入力してください」で止まった
function requestCertificate(){
  if(typeof window.switchTab === 'function') window.switchTab('certificate');
  setTimeout(() => {
    const el = document.getElementById('cert-name');
    if(el){ el.scrollIntoView({ block: 'center', behavior: 'smooth' }); el.focus({ preventScroll: true }); }
  }, 300);
}

window.openSbDrawer = openSbDrawer;
window.renderExam = renderExam;
window.renderExamResult = renderExamResult;


