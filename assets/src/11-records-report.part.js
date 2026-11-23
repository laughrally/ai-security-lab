/* ===== 受講実績レポート ===== */


function repMsg(type, text){
  const el = document.getElementById('rep-msg');
  if(!el) return;
  el.style.display = text ? 'block' : 'none';
  el.className = 'form-msg' + (type ? ' ' + type : '');
  el.textContent = text || '';
}

function fmtDateTime(iso){
  if(!iso) return '—';
  const d = new Date(iso);
  if(isNaN(d)) return '—';
  const p = n => String(n).padStart(2,'0');
  return `${d.getFullYear()}/${p(d.getMonth()+1)}/${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}
function fmtDuration(sec){
  if(sec === null || sec === undefined) return '—';
  const h = Math.floor(sec/3600), m = Math.floor((sec%3600)/60), s = sec%60;
  if(h) return `${h}時間${m}分`;
  if(m) return `${m}分${s}秒`;
  return `${s}秒`;
}

// 受講記録をサーバーから取得し、レポート用に整形する受講時間の計測を作り直した時刻。
// これより前の duration_seconds は旧方式（開いた時刻→完了を押した時刻）で記録されており、実態の3.8倍に膨らんだ例がある。
// 数値として信用できないためレポートには出さない。DBの値は書き換えず、表示だけ伏せる。
const MEASURE_FIXED_AT = Date.parse('2026-08-30T19:00:00+09:00');
function reportableDuration(log){
  const d = (log.duration_seconds === undefined || log.duration_seconds === null) ? null : log.duration_seconds;
  if(d === null) return null;
  const c = log.completed_at ? Date.parse(log.completed_at) : null;
  if(c !== null && c < MEASURE_FIXED_AT) return null;   // 旧方式の記録は伏せる
  return d;
}

async function fetchRecordData(){
  const data = await apiFetch('/api/records', { method: 'GET' });
  const meta = PLAN_LABELS_JA[data.plan] || PLAN_LABELS_JA.basic;
  const planLessons = getPlanLessons();

  const logMap = {};
  (data.lessons || []).forEach(l => { logMap[l.lesson_id] = l; });
  const quizMap = {};
  (data.quiz || []).forEach(q => { quizMap[q.lesson_id] = q; });

  const rows = planLessons.map(id => {
    const log = logMap[id] || {};
    const q   = quizMap[id];
    return {
      id,
      title: (LESSON_LABELS[id] || id),
      startedAt:   log.started_at   || null,
      completedAt: log.completed_at || null,
      duration:    reportableDuration(log),
      answered:    !!q,
      selected:    q ? q.selected + 1 : null,       // 表示は1始まり
      correctNo:   null,   // 正答番号はサーバーが返さない（2026-11-22。読めば再回答で正解できたため）
      isCorrect:   q ? q.is_correct : null,
      answeredAt:  q ? q.answered_at : null,
    };
  });

  // 章末ワーク（chapter_works）を章ごとにまとめる
  const chTotal = (data.summary && data.summary.chaptersTotal) || 0;
  const workMap = {};
  (data.works || []).forEach(w => {
    if(!workMap[w.chapter]) workMap[w.chapter] = [];
    workMap[w.chapter].push(w);
  });
  const works = [];
  for(let c=1; c<=chTotal; c++){
    const items = (workMap[c] || []).sort((a,b)=> (a.question_id||'').localeCompare(b.question_id||''));
    const info = (typeof WORK_CHAPTERS !== 'undefined' && WORK_CHAPTERS[c]) || {};
    works.push({
      chapter: c,
      name: info.name || '',
      goal: info.goal || '',
      // 4問すべて提出した章だけを「提出済み」とする
      submitted: ((data.summary && data.summary.workDoneList) || []).includes(c),
      submittedAt: (items.find(i => i.submitted_at) || {}).submitted_at || null,
      // duration_seconds は章内の全行に同じ値が入る仕様なので、合計せず最大を取る
      duration: items.reduce((a,i) => Math.max(a, i.duration_seconds || 0), 0) || null,
      answers: items.map(i => ({
        questionId: i.question_id,
        question: ((typeof WORK_QUESTIONS !== 'undefined' && WORK_QUESTIONS[i.question_id]) || {}).q || '',
        text: i.answer_text || '',
        length: (i.answer_text || '').trim().length,
        min: ((typeof WORK_QUESTIONS !== 'undefined' && WORK_QUESTIONS[i.question_id]) || {}).min || null,
        status: i.status
      }))
    });
  }

  // 章末確認テスト（chapter_tests）を章ごとにまとめる
  const testMap = {};
  (data.tests || []).forEach(t => {
    if(!testMap[t.chapter]) testMap[t.chapter] = [];
    testMap[t.chapter].push(t);
  });
  const tests = [];
  for(let c=1; c<=chTotal; c++){
    const items = (testMap[c] || []).sort((a,b)=> (a.question_id||'').localeCompare(b.question_id||''));
    const correct = items.filter(i => i.is_correct).length;
    const fItems = (data.testsFirst || []).filter(t => t.chapter === c);
    const fCorrect = fItems.filter(i => i.is_correct).length;
    const maxAtt = Math.max(1, ...items.map(i => i.attempt_no || 1));
    tests.push({
      chapter: c,
      answered: items.length,
      total: 5,
      correct,
      firstCorrect: fCorrect,
      firstAnswered: fItems.length,
      attempts: maxAtt,
      passed: items.length >= 5 && correct >= 4,
      items: items.map(i => ({
        questionId: i.question_id,
        selected: i.selected + 1,
        correctNo: null,
        isCorrect: i.is_correct,
        answeredAt: i.answered_at
      }))
    });
  }

  const dates = rows.map(r => r.completedAt).filter(Boolean).sort();
  return {
    meta, rows, works, tests,
    plan: data.plan,
    email: data.email,
    summary: data.summary,
    firstDate: dates[0] || null,
    lastDate: dates[dates.length-1] || null,
  };
}

async function openRecordReport(){
  const name    = (document.getElementById('rep-name')?.value || '').trim();
  const company = (document.getElementById('rep-company')?.value || '').trim();
  if(!name){ repMsg('error','氏名を入力してください。'); return; }

  repMsg('', '');
  // ポップアップブロック対策：先にウィンドウを開いてから中身を書き込む
  const win = window.open('', '_blank');
  if(!win){ repMsg('error','ポップアップがブロックされました。許可してから再度お試しください。'); return; }
  win.document.write('<p style="font-family:sans-serif;padding:24px">読み込み中…</p>');

  try {
    const d = await fetchRecordData();
    const s = d.summary;

    const trs = d.rows.map(r => `
      <tr>
        <td class="c">${escHtml(r.title)}</td>
        <td>${fmtDateTime(r.startedAt)}</td>
        <td>${fmtDateTime(r.completedAt)}</td>
        <td class="n">${fmtDuration(r.duration)}</td>
        <td class="n">${r.answered ? '選択肢' + r.selected : '未回答'}</td>
        <td class="n">${r.answered ? (r.isCorrect ? '正解' : '不正解') : '—'}</td>
      </tr>`).join('');

    const rate = s.quizAnswered ? Math.round(s.quizCorrect / s.quizAnswered * 1000)/10 : 0;

    // 集計は Worker の totalSeconds を使わず、レポートに出せる値だけで積み直す。
    // Worker 側は旧方式の記録も合算してしまうため。
    const lessonSecs = d.rows.reduce((a,r) => a + (r.duration || 0), 0);
    const workSecs   = (d.works || []).reduce((a,w) => a + (w.duration || 0), 0);
    const hiddenN    = d.rows.filter(r => r.completedAt && r.duration === null).length;
    const measureNote = hiddenN > 0
      ? `<p class="mnote">※ ${hiddenN}件のレッスンは、2026年8月30日の計測方式の改修より前に受講されたものです。当時の記録方式では受講時間を正確に計測できていなかったため、受講時間の表示を差し控えています（受講の事実・完了日時・理解度チェックの結果は記録のとおりです）。</p>`
      : '';
    const today = fmtDateTime(new Date().toISOString()).split(' ')[0];

    // 章末ワーク（記述式）の出力
    const chTotal = s.chaptersTotal || 0;
    const chDone  = s.chaptersSubmitted || 0;
    // 章末確認テストの出力
    let testsHTML = '';
    if(chTotal > 0 && (d.tests||[]).some(t => t.answered > 0)){
      const trows = (d.tests||[]).map(t => `
        <tr>
          <td>第${t.chapter}章</td>
          <td class="n">${t.answered} / ${t.total} 問</td>
          <td class="n">${t.firstAnswered ? t.firstCorrect + ' 問' : '—'}</td>
          <td class="n">${t.correct} 問</td>
          <td class="n">${t.answered ? Math.round(t.correct/t.answered*1000)/10 : 0}%</td>
          <td class="n">${t.attempts > 1 ? t.attempts + ' 回' : '1 回'}</td>
          <td class="n">${t.answered >= t.total ? (t.passed ? '合格' : '不合格') : '未完了'}</td>
        </tr>`).join('');
      const tAll = (d.tests||[]).reduce((a,t)=>a+t.answered,0);
      const cAll = (d.tests||[]).reduce((a,t)=>a+t.correct,0);
      const fAll = (d.tests||[]).reduce((a,t)=>a+(t.firstCorrect||0),0);
      testsHTML = `
<h2>章末確認テストの結果</h2>
<table>
  <thead><tr><th>章</th><th>回答数</th><th>初回正答</th><th>最終正答</th><th>正答率</th><th>受験回数</th><th>判定</th></tr></thead>
  <tbody>${trows}</tbody>
  <tfoot><tr>
    <th>合計</th><th>${tAll} / ${chTotal*5} 問</th><th>${fAll} 問</th><th>${cAll} 問</th>
    <th>${tAll ? Math.round(cAll/tAll*1000)/10 : 0}%</th><th>—</th><th>—</th>
  </tr></tfoot>
</table>
<p class="wnote">各章5問・4肢択一。合格ラインは4問以上（80%）。判定は当社が設定した正答に基づき講座システムが自動で行っています。<br>
「初回正答」は最初に回答した時点の正答数、「最終正答」は学び直しを含めた最新の回答による正答数です。理解度の把握には初回正答を、修了判定には最終正答を用いています。</p>`;
    }

    let worksHTML = '';
    if(chTotal > 0){
      const blocks = (d.works || []).map(w => {
        const qs = w.answers.map(a => `
          <div class="wq">
            <div class="wq-h">${escHtml(a.questionId)}　<span class="wq-n">${a.length}文字${a.min?`（必要 ${a.min}文字以上）`:''}</span></div>
            <div class="wq-q">${escHtml(a.question)}</div>
            <div class="wq-a">${escHtml(a.text) || '（未記入）'}</div>
          </div>`).join('');
        return `
        <div class="wch">
          <h3>第${w.chapter}章　${escHtml(w.goal)}
            <span class="wst ${w.submitted?'ok':''}">${w.submitted ? '提出済み' + (w.submittedAt ? '（' + fmtDateTime(w.submittedAt) + '）' : '') : '未提出'}</span>
            ${w.duration ? `<span class="wst">所要 ${fmtDuration(w.duration)}</span>` : ''}
          </h3>
          ${qs || '<div class="wq-a">（未記入）</div>'}
        </div>`;
      }).join('');
      worksHTML = `
<h2>章末ワーク（記述式課題）の提出内容</h2>
<p class="wnote">各章の学習内容を受講者自身の業務・組織に適用して記述する課題です。提出状況：<strong>${chDone} / ${chTotal} 章</strong></p>
${blocks}`;
    }

    win.document.open();
    win.document.write(`<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">
<title>受講実績レポート_${escHtml(name)}</title>
<style>
  *{box-sizing:border-box}
  body{font-family:"Hiragino Sans","Yu Gothic",sans-serif;color:#1a1a1a;margin:0;padding:32px;line-height:1.7}
  .wrap{max-width:900px;margin:0 auto}
  h1{font-size:22px;margin:0 0 4px;border-bottom:3px solid #c0392b;padding-bottom:10px}
  .sub{font-size:12px;color:#666;margin-bottom:24px}
  table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:24px}
  th,td{border:1px solid #ccc;padding:7px 9px;text-align:left;vertical-align:top}
  th{background:#f4f4f4;font-weight:700;white-space:nowrap}
  td.n{white-space:nowrap}
  td.c{min-width:200px}
  .info th{width:150px;background:#fafafa}
  .sum{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:24px}
  .sum div{flex:1;min-width:130px;border:1px solid #ccc;border-radius:6px;padding:12px}
  .sum b{display:block;font-size:19px;margin-top:4px}
  .lbl{font-size:11px;color:#666}
  .foot{font-size:11px;color:#666;border-top:1px solid #ccc;padding-top:14px;margin-top:8px}
  h2{font-size:14px;margin:24px 0 10px;padding-left:8px;border-left:4px solid #c0392b}
  @media print{ body{padding:0} .noprint{display:none} table{font-size:10.5px} }
  .btn{background:#c0392b;color:#fff;border:0;padding:10px 20px;border-radius:6px;font-size:14px;cursor:pointer;margin-bottom:20px}
  .mnote{font-size:12px;color:#555;line-height:1.7;margin:10px 0 0;padding:10px 12px;background:#f6f6f6;border-left:3px solid #bbb}
  .wnote{font-size:12px;color:#555;margin:0 0 14px}
  tfoot th{background:#f4f4f4;text-align:left;border:1px solid #ccc;padding:7px 9px}
  .wch{border:1px solid #ccc;border-radius:6px;padding:14px 16px;margin-bottom:14px;page-break-inside:avoid}
  .wch h3{font-size:13px;margin:0 0 12px;padding-bottom:8px;border-bottom:1px solid #e0e0e0;display:flex;justify-content:space-between;align-items:baseline;gap:10px}
  .wst{font-size:11px;font-weight:400;color:#999;white-space:nowrap}
  .wst.ok{color:#1a7f4b;font-weight:700}
  .wq{margin-bottom:14px}
  .wq:last-child{margin-bottom:0}
  .wq-h{font-size:11px;font-weight:700;color:#c0392b;margin-bottom:3px}
  .wq-n{font-weight:400;color:#888}
  .wq-q{font-size:11px;color:#666;line-height:1.6;margin-bottom:6px}
  .wq-a{font-size:12px;line-height:1.85;white-space:pre-wrap;background:#fafafa;border-left:3px solid #ddd;padding:9px 12px}


</style></head><body><div class="wrap">
<button class="btn noprint" data-print>印刷 / PDFとして保存</button>
<h1>受講実績レポート</h1>
<div class="sub">AI Security Lab 生成AI×セキュリティ eラーニング</div>

<h2>受講者情報</h2>
<table class="info">
  <tr><th>氏名</th><td>${escHtml(name)}</td></tr>
  ${company ? `<tr><th>会社名</th><td>${escHtml(company)}</td></tr>` : ''}
  <tr><th>メールアドレス</th><td>${escHtml(d.email || '')}</td></tr>
  <tr><th>講座名</th><td>AI Security Lab 生成AI×セキュリティ eラーニング（${escHtml(d.meta.name)}）</td></tr>
  <tr><th>受講範囲</th><td>${escHtml(d.meta.range)}</td></tr>
  <tr><th>標準学習時間</th><td>${escHtml(d.meta.hours)}（休憩時間を除く）</td></tr>
  <tr><th>受講期間</th><td>${fmtDateTime(d.firstDate)} 〜 ${fmtDateTime(d.lastDate)}</td></tr>
</table>

<h2>集計</h2>
<div class="sum">
  <div><span class="lbl">受講済みレッスン</span><b>${s.lessonsCompleted} / ${s.lessonsTotal}</b></div>
  <div><span class="lbl">レッスン学習時間（実測）</span><b>${lessonSecs ? fmtDuration(lessonSecs) : '—'}</b></div>
  ${chTotal ? `<div><span class="lbl">章末ワーク所要時間（実測）</span><b>${workSecs ? fmtDuration(workSecs) : '—'}</b></div>` : ''}
  <div><span class="lbl">理解度テスト回答数</span><b>${s.quizAnswered} 問</b></div>
  <div><span class="lbl">正答数 / 正答率</span><b>${s.quizCorrect} 問 / ${rate}%</b></div>
  ${chTotal ? `<div><span class="lbl">章末ワーク提出</span><b>${chDone} / ${chTotal} 章</b></div>` : ''}
  ${s.testTotal ? `<div><span class="lbl">章末確認テスト</span><b>${s.testCorrect||0} / ${s.testTotal} 問</b></div>` : ''}
  <div><span class="lbl">修了状況</span><b>${s.completed ? '修了' : '受講中'}</b></div>
</div>
${measureNote}

<h2>レッスン別 受講履歴および理解度テスト結果</h2>
<table>
  <thead><tr>
    <th>レッスン</th><th>開始日時</th><th>完了日時</th><th>学習時間</th><th>回答</th><th>評価結果</th>
  </tr></thead>
  <tbody>${trs}</tbody>
</table>

${testsHTML}
${worksHTML}

<div class="foot">
  ・理解度テストは各レッスンに1問（4肢択一）を設けています。表中の「回答」は<strong>初回回答</strong>を記載しています。<br>
  ・正誤の判定は、当社が設定した正答に基づき講座システムが自動で行っています（受講者による自己採点ではありません）。<br>
  ・学習時間は受講画面の滞在時間の実測値です。標準学習時間とは一致しません。<br>
  ・「—」は該当データがないことを示します。本機能の導入前に受講されたレッスンは開始日時・学習時間が記録されていない場合があります。<br>
  ・章末ワークは受講者が自社の状況を記述する課題です。提出時に設問ごとの必要文字数を満たしているかを講座システムが自動で判定しています。本レポートには最新の提出内容と提出日時を記載しています。<br>
  ・章末確認テストは各章5問（4肢択一）、合格ラインは4問以上です。不合格の場合は該当レッスンを確認のうえ再挑戦でき、「初回正答」は最初に回答した時点の、「最終正答」は学び直しを含めた最新の回答による正答数です。<br><br>
  発行日：${today}／発行者：合同会社LaughRally（AI Security Lab）
</div>
</div></body></html>`);
    win.document.close();
    // 別ウィンドウも元のページの CSP を引き継ぐので、onclick は動かない。ボタンはここで登録する
    win.document.querySelectorAll('[data-print]').forEach(b => b.addEventListener('click', () => win.print()));
    repMsg('success','レポートを別ウィンドウで開きました。');
  } catch(e) {
    win.close();
    repMsg('error','レポートを取得できませんでした。' + (e.message ? '（' + e.message + '）' : ''));
  }
}

async function downloadRecordCSV(){
  const name    = (document.getElementById('rep-name')?.value || '').trim();
  const company = (document.getElementById('rep-company')?.value || '').trim();
  if(!name){ repMsg('error','氏名を入力してください。'); return; }
  repMsg('', '');

  try {
    const d = await fetchRecordData();
    // CSVの数式インジェクション対策：= + - @ で始まる値はExcelが数式として解釈するため
    // 先頭にアポストロフィを付けて無害化する（表示上は元の文字列のまま）
    const q = csvCell;
    const lines = [
      ['氏名','会社名','メールアドレス','講座名','プラン','受講範囲','標準学習時間'].map(q).join(','),
      [name, company, d.email || '', 'AI Security Lab 生成AI×セキュリティ eラーニング',
       d.meta.name, d.meta.range, d.meta.hours].map(q).join(','),
      '',
      ['レッスン番号','レッスン名','開始日時','完了日時','学習時間(秒)','選択した選択肢','評価結果','回答日時'].map(q).join(','),
    ];
    d.rows.forEach(r => {
      lines.push([
        'L' + lessonNo(r.id), r.title,
        r.startedAt ? fmtDateTime(r.startedAt) : '',
        r.completedAt ? fmtDateTime(r.completedAt) : '',
        r.duration == null ? '' : r.duration,
        r.answered ? r.selected : '',
        r.answered ? (r.isCorrect ? '正解' : '不正解') : '未回答',
        r.answeredAt ? fmtDateTime(r.answeredAt) : '',
      ].map(q).join(','));
    });

    // 章末確認テスト
    if((d.tests || []).some(t => t.answered > 0)){
      lines.push('', ['章','設問ID','選択した選択肢','評価結果','回答日時'].map(q).join(','));
      d.tests.forEach(t => {
        t.items.forEach(i => {
          lines.push(['第'+t.chapter+'章', i.questionId, i.selected,
                      i.isCorrect ? '正解' : '不正解',
                      i.answeredAt ? fmtDateTime(i.answeredAt) : ''].map(q).join(','));
        });
      });
    }

    // 章末ワーク（記述式）
    if((d.works || []).length){
      lines.push('', ['章','章末ワーク','設問ID','提出状況','提出日時','文字数','必要文字数','回答内容'].map(q).join(','));
      d.works.forEach(w => {
        if(!w.answers.length){
          lines.push(['第'+w.chapter+'章', w.goal, '', '未提出', '', '', '', ''].map(q).join(','));
          return;
        }
        w.answers.forEach(a => {
          lines.push([
            '第'+w.chapter+'章', w.goal, a.questionId,
            a.status === 'submitted' ? '提出済み' : '下書き',
            w.submittedAt ? fmtDateTime(w.submittedAt) : '',
            a.length, a.min == null ? '' : a.min,
            a.text
          ].map(q).join(','));
        });
      });
    }

    const s = d.summary;
    lines.push('', ['集計','受講済み','累計学習時間(秒)','理解度チェック回答数','同 正答数','章末確認テスト正答','章末ワーク提出','修了状況'].map(q).join(','));
    lines.push(['', `${s.lessonsCompleted}/${s.lessonsTotal}`, d.rows.reduce((a,r)=>a+(r.duration||0),0),
                s.quizAnswered, s.quizCorrect,
                s.testTotal ? `${s.testCorrect||0}/${s.testTotal}` : '対象外',
                s.chaptersTotal ? `${s.chaptersSubmitted||0}/${s.chaptersTotal}` : '対象外',
                s.completed ? '修了' : '受講中'].map(q).join(','));

    // Excelで文字化けしないよう BOM を付ける
    const blob = new Blob(['\ufeff' + lines.join('\r\n')], { type:'text/csv;charset=utf-8;' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `受講実績_${name}_${new Date().toISOString().slice(0,10)}.csv`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(()=>URL.revokeObjectURL(a.href), 1000);
    repMsg('success','CSVをダウンロードしました。');
  } catch(e) {
    repMsg('error','CSVを作成できませんでした。' + (e.message ? '（' + e.message + '）' : ''));
  }
}

