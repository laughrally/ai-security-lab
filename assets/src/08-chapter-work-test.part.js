// ===================== 章末ワーク =====================

const WORK_QUESTIONS = {
  'w1-1': { min:150, ref:'L03', q:'あなたの部署または会社で、業務に生成AIツールが使われている場面をすべて挙げてください。会社が公式に導入したものだけでなく、個人のアカウントで使われているもの、無料版を使っているものも含めて書き出してください。「誰が・どのツールを・どんな業務で」使っているかがわかる形で記述してください。心当たりがない場合は、そう判断した根拠（ヒアリングの実施、ログの確認など）を書いてください。' },
  'w1-2': { min:150, ref:'L01', q:'w1-1で挙げたツールに、これまで実際に入力された情報の種類を書き出してください。顧客名・取引先名・金額・個人情報・社内資料の文面など、具体的に挙げてください。そのうえで、それらの情報が現在どこに保存されている可能性があるかを、L01の内容をふまえて記述してください。' },
  'w1-3': { min:150, ref:'L02', q:'w1-1で挙げたツールのうち、主要なもの1つ以上について、実際に管理画面を開いて学習オフ設定の状態を確認してください。確認したツール名、設定の場所、現在の状態（オン／オフ）を記述してください。そのうえで、学習オフにしても防げないリスクが何かを、L02の内容をふまえて書いてください。' },
  'w1-4': { min:200, ref:'L04・総括', q:'ここまでの3問をふまえて、自社のAI利用実態を整理してください。最もリスクが高いと考えられる箇所を1つ挙げ、その理由と、当面どう対処するかを記述してください。対処が「現状維持」である場合も、なぜそう判断したかを書いてください。' },

  'w2-1': { min:150, ref:'L07', q:'自社で扱う情報を「入力してよい情報」「条件付きで入力してよい情報」「絶対に入力してはいけない情報」の3つに分類してください。それぞれ最低3つずつ、自社の業務で実際に扱っているものを具体的に挙げてください。「顧客情報」のような抽象的な括りではなく、「見積書に記載された取引先名と金額」のように、判断できる粒度で書いてください。' },
  'w2-2': { min:150, ref:'L07', q:'w2-1で「条件付き」に分類した情報について、その条件を具体的に定めてください。どのツールなら許可するのか、どう加工すれば入力してよいのか（社名を伏せる、数値を丸めるなど）、誰の承認が必要かを記述してください。条件付きに分類したものがない場合は、なぜ2分類で足りると判断したかを書いてください。' },
  'w2-3': { min:150, ref:'L05・L06', q:'現在自社で使っている、または導入を検討している生成AIツールについて、L05・L06で学んだセキュリティ設定の観点から評価してください。プランの違い（無料版か有料版か、法人向けプランか）、学習に使われるかどうか、管理者による統制が可能かを整理し、現状のままでよいか、変更すべきかを記述してください。' },
  'w2-4': { min:200, ref:'L08・総括', q:'L08の業種別リスク事例のうち、自社の業種または業務に最も近いものを挙げ、自社で同じことが起きる可能性を検討してください。起こりうると考える場合はその経路を、起こりにくいと考える場合はその根拠を記述してください。そのうえで、w2-1で作成した入力禁止情報リストを社内で共有する方法（朝礼、書面配布、システム上の掲示など）を具体的に書いてください。' },

  'w3-1': { min:150, ref:'L09', q:'自社が業務で扱う個人情報のうち、生成AIに入力される可能性があるものを挙げてください。氏名・連絡先だけでなく、健康情報や信条など要配慮個人情報にあたるものがあれば区別して記述してください。そのうえで、L09の内容をふまえ、それらをAIに入力することが個人情報保護法上どう位置づけられるかを自社の状況に即して書いてください。' },
  'w3-2': { min:150, ref:'L10', q:'自社の業務で、生成AIが出力した文章・画像・コードをそのまま使っている、または使う可能性がある場面を挙げてください。そのうえで、L10で学んだ「著作権侵害になるケース・ならないケース」に照らし、自社の使い方がどちらに近いか、リスクを下げるために何を変えるべきかを記述してください。該当する場面がない場合は、今後発生しうる場面を想定して書いてください。' },
  'w3-3': { min:150, ref:'L11', q:'生成AIの誤った出力（ハルシネーション）が自社の業務で実害につながりうる場面を、具体的に1つ以上挙げてください。誰に対してどのような損害が生じるか、その場合に責任を負うのは誰かを、L11の内容をふまえて記述してください。' },
  'w3-4': { min:200, ref:'総括', q:'w3-1〜w3-3で挙げたリスクを、「発生可能性」と「発生した場合の影響の大きさ」の2軸で評価し、優先的に対処すべきものを1つ選んでください。その選定理由と、具体的な対処方針（チェック体制、利用制限、記録の残し方など）を記述してください。あわせて、自社に法務担当や顧問弁護士がいる場合は相談すべき論点を、いない場合はどこに相談するかを書いてください。' },

  'w4-1': { min:150, ref:'L12', q:'L12のテンプレートをもとに、自社のAI利用ガイドラインの「適用範囲」と「禁止事項」を作成してください。適用範囲は誰に適用されるか（正社員のみか、パート・派遣・業務委託を含むか）を明記してください。禁止事項は第2章で作成した入力禁止情報リストを反映させ、実際に運用できる粒度で記述してください。' },
  'w4-2': { min:150, ref:'L13', q:'自社でAIによる情報漏洩が疑われる事態が起きた場合の初動対応を、L13の内容をふまえて記述してください。誰が最初に報告を受けるか、その人が最初に行うこと、外部（顧客・取引先・監督官庁）への連絡が必要になる判断基準を、自社の実際の体制に即して具体的に書いてください。' },
  'w4-3': { min:150, ref:'L14', q:'L14の内容をふまえ、自社でAI利用を管理する体制について記述してください。専任者を置けない場合、誰が兼任するのが現実的か、その人にどこまでの権限を持たせるかを書いてください。あわせて、ガイドラインを従業員に周知し、遵守を確認する方法（研修、誓約書、定期チェックなど）を具体的に記述してください。' },
  'w4-4': { min:200, ref:'L15・総括', q:'L15の「今日からできる10のアクション」のうち、自社で未実施のものを挙げ、優先順位をつけてください。上位3つについて、いつまでに、誰が、何をするかを記述してください。すでにすべて実施済みの場合は、それぞれが形骸化していないかを点検した結果と、次に取り組むべきことを書いてください。' },

  'w5-1': { min:150, ref:'L16', q:'L16の20事例から、自社の業務に最も近いものを1つ選び、なぜそれを選んだかを記述してください。そのうえで、その事例で問題となった点が自社でも起こりうるか、自社の場合はどこが違うかを具体的に書いてください。' },
  'w5-2': { min:150, ref:'L16', q:'自社の業務のうち、生成AIを使うことで最も効果が見込める作業を1つ挙げてください。現在その作業に月あたりどれくらいの時間がかかっているか、AIを使うことでどの程度短縮できると見込むかを記述してください。数値が把握できていない場合は、概算とその根拠を書いてください。' },
  'w5-3': { min:150, ref:'L17', q:'w5-2で挙げた業務に使うツールを、L17の比較をふまえて選定してください。選んだツール名、プラン、想定する月額費用、選定理由を記述してください。あわせて、そのツールを使ううえで第2章で作成した入力禁止情報リストとの整合が取れているかを確認し、追加で必要な設定や運用ルールがあれば書いてください。' },
  'w5-4': { min:200, ref:'総括', q:'w5-2・w5-3をふまえ、自社のAI導入計画をまとめてください。導入する業務、使用するツール、対象となる従業員、開始時期、効果を測る指標を記述してください。導入しないという結論であっても構いませんが、その場合はなぜ現時点で導入しないのか、どういう条件が整えば導入するのかを書いてください。' },

  'w6-1': { min:150, ref:'L23', q:'自社に届くメールや連絡のうち、AIを使ったフィッシングの標的になりやすいものを挙げてください。取引先を装った請求書、社内の上長を装った指示など、自社で実際にありうる形を具体的に書いてください。そのうえで、L23で学んだ見分け方のうち、自社の従業員が実行できるものを記述してください。' },
  'w6-2': { min:150, ref:'L24', q:'自社で生成AIに外部の文書やWebページを読み込ませる業務があるか、あるいは今後発生しうるかを検討してください。ある場合、L24のプロンプトインジェクションが成立しうる経路を具体的に記述してください。ない場合は、なぜ発生しないと判断したかを書いてください。' },
  'w6-3': { min:150, ref:'L25', q:'自社で「本人確認」が必要になる場面を挙げてください。電話での振込指示、役員からの緊急依頼、取引先からの口座変更連絡などが該当します。L25をふまえ、声や映像が偽装されうる前提で、確実に本人を確認する手段を自社の実情に即して記述してください。' },
  'w6-4': { min:200, ref:'L27・総括', q:'L27のロールプレイをふまえ、自社で情報漏洩が発生した場合の対応フローを作成してください。発見から報告、被害範囲の確認、関係者への連絡、再発防止までの流れを時系列で記述してください。第4章のw4-2で書いた初動対応と矛盾がないか確認し、実際に動かせる内容になっているかを点検したうえで書いてください。' },

  'w7-1': { min:150, ref:'L31', q:'自社の業務で生成AIに指示を出す際、現在どのような書き方をしているかを具体的に記述してください。実際に使っている指示文があれば、そのまま書き出してください。そのうえで、L31の内容をふまえ、より安全で精度の高い指示に書き換えるとどうなるかを記述してください。' },
  'w7-2': { min:150, ref:'L32', q:'自社のデータ（社内マニュアル、過去の見積書、顧客対応履歴など）をAIに読み込ませて活用する構想があるかを検討してください。ある場合、L32をふまえてRAGとファインチューニングのどちらが適切か、その理由と想定されるリスクを記述してください。ない場合は、どういうデータであれば読み込ませてよいと考えるかを書いてください。' },
  'w7-3': { min:150, ref:'L34', q:'L34をふまえ、自社でAIエージェントや自動化ツールを使う場合に許容できる範囲を検討してください。AIに自動で実行させてよい操作と、必ず人間の確認を挟むべき操作を区別して記述してください。現時点で使う予定がない場合も、将来使うとしたらどう線を引くかを書いてください。' },
  'w7-4': { min:200, ref:'L33・総括', q:'L33をふまえ、自社のAI活用ロードマップを作成してください。今後1年間で、どの業務にどの順序でAIを導入していくか、各段階で必要になるセキュリティ対策とあわせて記述してください。第5章のw5-4で作成した導入計画との整合を確認し、より長期の視点で見直した結果を書いてください。' }
};

const WORK_API = WORKER_URL;   // 同じ Worker（以前は URL を2か所に書いていた）
const _workTimers = {};
const _workStarted = {};
let _workLoaded = {};
const _workSubmittedCh = {};   // { 章番号: true } サイドバーの完了表示に使う
const _workCurQ = {};          // { 章番号: 表示中の設問番号 }
const _workReopen = {};        // { 章番号: true } 書き直しモード

async function workToken(){
  const { data } = await _supabase.auth.getSession();
  return data?.session?.access_token || null;
}

function workCount(t){ return (t||'').trim().length; }

function renderWorkChapter(ch){
  const info = WORK_CHAPTERS[ch];
  const host = document.getElementById('work-body-'+ch);
  if(!host) return;
  let steps = '<div class="work-steps" id="work-steps-'+ch+'">';
  for(let i=1;i<=COURSE.rules.workQuestions;i++) steps += `<div class="work-step" id="wst-${ch}-${i}"></div>`;
  steps += '</div>';
  let html = steps;
  for(let i=1;i<=COURSE.rules.workQuestions;i++){
    const qid = `w${ch}-${i}`;
    const d = WORK_QUESTIONS[qid];
    html += `
      <div class="work-q${i===1?' is-active':''}" id="wq-${qid}">
        <div class="work-q-head">
          <span class="work-q-no">設問 ${i} / ${COURSE.rules.workQuestions}</span>
          <span class="work-q-ref">参照：${d.ref}</span>
        </div>
        <p class="work-q-text">${d.q}</p>
        <textarea class="work-input" id="wa-${qid}" data-qid="${qid}"
          placeholder="ここに記入してください（${d.min}文字以上）"
          oninput="onWorkInput('${qid}')"></textarea>
        <div class="work-meta">
          <span class="work-count" id="wc-${qid}">0 / ${d.min} 文字</span>
          <span class="work-saved" id="ws-${qid}"></span>
        </div>
        <div class="work-nav">
          <button class="work-nav-btn" onclick="gotoWorkQ(${ch},${i-1})" ${i===1?'disabled':''}>← 前の設問</button>
          <button class="work-nav-btn" onclick="gotoWorkQ(${ch},${i+1})" ${i===COURSE.rules.workQuestions?'disabled':''}>次の設問 →</button>
        </div>
      </div>`;
  }
  host.innerHTML = html;
  _workCurQ[ch] = 1;
}

function gotoWorkQ(ch, n){
  if(n<1 || n>COURSE.rules.workQuestions) return;
  for(let i=1;i<=COURSE.rules.workQuestions;i++){
    const el = document.getElementById(`wq-w${ch}-${i}`);
    if(el) el.classList.toggle('is-active', i===n);
  }
  _workCurQ[ch] = n;
  updateWorkSteps(ch);
  const host = document.getElementById('work-body-'+ch);
  if(host) host.scrollIntoView({behavior:'smooth', block:'start'});
  const ta = document.getElementById(`wa-w${ch}-${n}`);
  if(ta && !ta.readOnly) setTimeout(()=>ta.focus(), 260);
}

function updateWorkSteps(ch){
  for(let i=1;i<=COURSE.rules.workQuestions;i++){
    const st = document.getElementById(`wst-${ch}-${i}`);
    if(!st) continue;
    const qid = `w${ch}-${i}`;
    const el = document.getElementById('wa-'+qid);
    const ok = el && workCount(el.value) >= WORK_QUESTIONS[qid].min;
    st.classList.toggle('is-ok', !!ok);
    st.classList.toggle('is-cur', !ok && _workCurQ[ch]===i);
  }
}

async function loadWorkChapter(ch){
  const host = document.getElementById('work-body-'+ch);
  const status = document.getElementById('work-status-'+ch);
  if(!host) return;
  if(!_workLoaded[ch]) renderWorkChapter(ch);
  const t = await workToken();
  if(!t){ status.textContent = 'ログインすると記入できます。'; return; }
  try{
    const r = await fetch(`${WORK_API}/api/work/${ch}`, { headers:{ Authorization:'Bearer '+t }});
    const d = await r.json();
    if(!r.ok){ status.textContent = d.error || '読み込みに失敗しました'; return; }
    (d.answers||[]).forEach(a=>{
      const el = document.getElementById('wa-'+a.question_id);
      // 保存前の入力がある欄は上書きしない。入力後1.5秒以内にタブを切り替えて戻ると、古い値で上書きされ、その古い値が保存されて入力が消えていた
      if(el && el.dataset.dirty !== '1'){ el.value = a.answer_text || ''; }
    });
    for(let i=1;i<=COURSE.rules.workQuestions;i++) updateWorkCount(`w${ch}-${i}`);
    _workLoaded[ch] = true;
    if(d.reopened){ _workReopen[ch] = true; }
    setWorkSubmitted(ch, !!d.submitted, d.answers||[]);
    if(d.reopened && status) status.innerHTML = '<span class="work-note">書き直しの途中です。修正後、あらためて提出してください（提出するまで、この章は未提出の扱いです）。</span>';
  }catch(e){
    status.textContent = '読み込みに失敗しました。通信環境をご確認ください。';
  }
}

function setWorkSubmitted(ch, submitted, answers){
  _workSubmittedCh[ch] = !!submitted;
  const status = document.getElementById('work-status-'+ch);
  const btn = document.getElementById('work-submit-'+ch);
  document.querySelectorAll(`#work-body-${ch} .work-input`).forEach(el=>{
    el.readOnly = submitted;
    el.classList.toggle('is-locked', submitted);
  });
  if(submitted){
    const at = (answers.find(a=>a.submitted_at)||{}).submitted_at;
    const d = at ? new Date(at).toLocaleString('ja-JP') : '';
    status.innerHTML = `<span class="work-done">✓ 提出済み</span>${d?`　${d}`:''}<br><span class="work-note">内容は受講実績レポートに記録されます。書き直した場合は最新の内容と提出日時が記録されます。</span><br><button class="work-reopen" onclick="reopenWork(${ch})">書き直す</button>`;
    if(btn){ btn.disabled = true; btn.textContent = '✓ 提出済み'; btn.classList.add('is-done'); }
  }else{
    status.textContent = '';
    if(btn){ btn.disabled = false; btn.classList.remove('is-done'); }
    updateWorkSubmitBtn(ch);
  }
}

function updateWorkCount(qid){
  const el = document.getElementById('wa-'+qid);
  const out = document.getElementById('wc-'+qid);
  if(!el||!out) return false;
  const n = workCount(el.value);
  const min = WORK_QUESTIONS[qid].min;
  out.textContent = `${n} / ${min} 文字`;
  const ok = n >= min;
  out.classList.toggle('is-ok', ok);
  return ok;
}

function updateWorkSubmitBtn(ch){
  const btn = document.getElementById('work-submit-'+ch);
  let ok = 0;
  for(let i=1;i<=COURSE.rules.workQuestions;i++) if(updateWorkCount(`w${ch}-${i}`)) ok++;
  updateWorkSteps(ch);
  if(!btn || btn.classList.contains('is-done')) return;   // 提出済み
  btn.disabled = ok < COURSE.rules.workQuestions;
  const need = COURSE.rules.workQuestions;
  btn.textContent = ok < need ? `提出する（${ok}/${need} 問が必要文字数に到達）` : '提出する';
}

function onWorkInput(qid){
  const ch = parseInt(qid.split('-')[0].slice(1),10);
  const _el = document.getElementById('wa-'+qid);
  if(_el) _el.dataset.dirty = '1';
  updateWorkSubmitBtn(ch);
  if(!_workStarted[qid]) _workStarted[qid] = new Date().toISOString();
  clearTimeout(_workTimers[qid]);
  const s = document.getElementById('ws-'+qid);
  if(s) s.textContent = '';
  _workTimers[qid] = setTimeout(()=>saveWorkDraft(qid), 1500);
}

// 下書きを保存する。保存できたか（または保存の必要が無いか）を true／false で返す
async function saveWorkDraft(qid){
  const el = document.getElementById('wa-'+qid);
  const s  = document.getElementById('ws-'+qid);
  if(!el || el.readOnly) return true;
  const sentValue = el.value;
  const t = await workToken();
  if(!t) return false;
  try{
    const r = await fetch(`${WORK_API}/api/work/draft`, {
      method:'POST',
      headers:{ Authorization:'Bearer '+t, 'Content-Type':'application/json' },
      body: JSON.stringify({ questionId:qid, answerText:sentValue, startedAt:_workStarted[qid], reopen: !!_workReopen[parseInt(qid.split('-')[0].slice(1),10)] })
    });
    const d = await r.json();
    // 送った内容のまま変わっていなければ「保存済み」に戻す（送信中にさらに入力していたら未保存のまま）
    if(r.ok && el.value === sentValue) el.dataset.dirty = '';
    if(s) s.textContent = r.ok ? '保存しました' : (d.error||'保存に失敗しました');
    if(s && r.ok) setTimeout(()=>{ if(s.textContent==='保存しました') s.textContent=''; }, 2500);
    return r.ok;
  }catch(e){
    if(s) s.textContent = '保存に失敗しました';
    return false;
  }
}

async function submitWorkChapter(ch){
  const btn = document.getElementById('work-submit-'+ch);
  const status = document.getElementById('work-status-'+ch);
  if(!confirm(`第${ch}章の章末ワークを提出します。\n提出後は編集できません。よろしいですか？`)) return;
  btn.disabled = true;
  btn.textContent = '提出中...';
  // 未保存の下書きを確定させる。1件でも保存できなければ提出しない（画面と違う古い文が提出済みになるため）
  for(let i=1;i<=COURSE.rules.workQuestions;i++){
    const qid = `w${ch}-${i}`;
    clearTimeout(_workTimers[qid]);
    if(!(await saveWorkDraft(qid))){
      btn.disabled = false; btn.textContent = '提出する';
      if(status) status.textContent = '下書きを保存できなかったため、提出していません。通信環境をご確認のうえ、もう一度「提出する」を押してください。';
      return;
    }
  }
  const t = await workToken();
  try{
    // 章末ワークに費やした実時間を送る。Worker 側はすでに durationSeconds を
    // 受け取れる実装になっており（chapter_works.duration_seconds）、
    // これまではフロントが送っていなかったため常に NULL だった。
    flushActive();
    const workSecs = loadActive()['work' + ch] || null;
    const r = await fetch(`${WORK_API}/api/work/submit`, {
      method:'POST',
      headers:{ Authorization:'Bearer '+t, 'Content-Type':'application/json' },
      body: JSON.stringify({ chapter: ch, durationSeconds: workSecs })
    });
    const d = await r.json();
    if(!r.ok){
      status.textContent = d.error || '提出に失敗しました';
      if(d.shortfall){
        status.textContent += '（' + d.shortfall.map(x=>`${x.questionId}: ${x.length}/${x.minLength}字`).join('、') + '）';
      }
      btn.disabled = false;
      updateWorkSubmitBtn(ch);
      return;
    }
    _workReopen[ch] = false;
    setWorkSubmitted(ch, true, [{submitted_at:d.submittedAt}]);
    try{ buildSideNav(loadCompleted()); }catch(e2){}
    refreshRecords();
  }catch(e){
    status.textContent = '提出に失敗しました。通信環境をご確認ください。';
    btn.disabled = false;
    updateWorkSubmitBtn(ch);
  }
}

window.loadWorkChapter   = loadWorkChapter;
// 提出済みの章を編集可能に戻す（記録は保持され、再提出で上書きされる）
async function reopenWork(ch){
  if(!confirm(`第${ch}章の章末ワークを書き直します。\n再提出するまで提出済みの扱いは解除されます。よろしいですか？`)) return;
  _workReopen[ch] = true;
  document.querySelectorAll(`#work-body-${ch} .work-input`).forEach(el=>{
    el.readOnly = false;
    el.classList.remove('is-locked');
  });
  _workSubmittedCh[ch] = false;
  const status = document.getElementById('work-status-'+ch);
  if(status) status.innerHTML = '<span class="work-note">書き直しモードです。修正後、あらためて提出してください。</span>';
  const btn = document.getElementById('work-submit-'+ch);
  if(btn){ btn.textContent = '提出する'; btn.disabled = false; btn.classList.remove('is-done'); }
  updateWorkSubmitBtn(ch);
  try{ buildSideNav(loadCompleted()); }catch(e){}
}
// ===================== 章末確認テスト =====================
const TEST_QUESTIONS = {
  't1-1': { q:'社内で「ChatGPTは有料版なら学習に使われないので安全だ」という説明がなされていました。この説明の問題点はどれですか。', c:['プランに関わらず入力内容はすべて学習に使われることになるため','学習利用の有無はプランではなく国ごとの法規制によって決まるため','学習利用の有無は個人向けか法人向けかで決まり、別の論点もあるから','有料版であっても入力内容は他の利用者から閲覧できてしまうから'] },
  't1-2': { q:'シャドーAIが組織にとって危険とされる最大の理由はどれですか。', c:['組織が把握できていないため、漏洩が起きても発覚せず対策も打てないから','無料版のAIは有料版より性能が低く、業務の品質が下がってしまうから','個人アカウントでのAI利用は法律で明確に禁止されているものだから','複数のAIツールを併用するとAI同士が情報を共有してしまうから'] },
  't1-3': { q:'自社の顧客名簿を扱う業務でAIを使いたいと相談されました。第1章をふまえた対応はどれですか。', c:['どのAIでも入力内容は暗号化されるため、そのまま使ってよいと伝える','AIの利用は一律に禁止し、従来どおり手作業で行うよう指示を出す','学習オフにすれば問題ないため、その設定方法だけを案内しておく','情報の種類と契約形態を確認し、入力してよい範囲を先に決めておく'] },
  't1-4': { q:'オンプレミス型AIについての説明として正しいものはどれですか。', c:['クラウド型と比べて、どのような規模でも必ず安価に導入できる','データを社外に出さずに済む一方、導入と運用の負担は大きい','インターネット接続が不要なため、モデルの更新も一切必要ない','中小企業向けの設計であり、専門知識がなくても運用ができる'] },
  't1-5': { q:'「入力した情報がどこへ行くのか」を考えるうえで、最初に確認すべきことはどれですか。', c:['そのAIの回答精度が業務で使える水準に達しているかどうか','開発元が日本の企業であり、国内法の適用を受けるかどうか','利用しているのがどの事業者のどの契約形態のサービスか','社内で何人がそのAIを日常的に利用しているかどうか'] },

  't2-1': { q:'社内で使うAIツールを選定する際、セキュリティの観点で最も重視すべき点はどれですか。', c:['回答が生成されるまでの速度が業務に支障しない水準かどうか','管理者が利用状況を把握し、設定を統制できるかどうか','日本語の表現が自然で、そのまま使える文章が出てくるか','対応しているファイル形式が業務の必要を満たすか'] },
  't2-2': { q:'「社外秘」と「個人情報」の扱いについて、正しい理解はどれですか。', c:['どちらも同種の情報であり、社内では同じ基準で扱えばよい','個人情報だけが法規制の対象であり、社外秘は自由に扱える','社外秘だけが企業の責任範囲で、個人情報は本人の責任となる','個人情報は法規制、社外秘は契約や営業秘密の観点から別に判断する'] },
  't2-3': { q:'生成AIに営業秘密を入力した場合のリスクとして、第2章で扱った論点はどれですか。', c:['秘密管理性が否定され、営業秘密として保護されなくなりうる','入力した時点で自動的に特許出願がなされたものとみなされてしまう','AI事業者の側に、その営業秘密の所有権が移転してしまうため','入力した内容が検索エンジンを経由して一般に公開されるため'] },
  't2-4': { q:'入力禁止情報のリストを作る際、実務上最も重要なことはどれですか。', c:['できるだけ多くの項目を列挙し、迷ったらすべて禁止にする','経営者だけで決定し、従業員には結果だけを通知しておく','現場が個別の場面で判断できる具体的な粒度で書くこと','法律の条文をそのまま引用し、根拠を明確にしておくこと'] },
  't2-5': { q:'業種別のリスク事例を自社に当てはめて検討する際、適切な姿勢はどれですか。', c:['自社と業種が異なる事例は前提が違うため参考にしない','事例で紹介された対策をそのまま自社にも導入しておく','過去に事故が起きていない業種はリスクが低いものと考える','問題となった構造が自社にも当てはまるかを業種を超えて確認する'] },

  't3-1': { q:'2026年7月に公布された改正個人情報保護法について、正しい説明はどれですか。', c:['課徴金制度が導入され、施行は公布から2年以内に政令で定められる','公布と同時に全面施行されており、すでに適用が開始されている','生成AIへの個人情報の入力を全面的に禁止する内容である','従業員数が一定以下の中小企業は適用除外とされている'] },
  't3-2': { q:'生成AIの出力をそのまま業務に使う場合、著作権の観点で最もリスクが高いのはどれですか。', c:['AIに生成させた文章を、社内向けのメモとして共有する','自社で作成した資料をAIに要約させて社内で回覧する','特定の作品や作風を指示して作らせた成果物を広告に使う','自分で書いた文章の誤字脱字をAIにまとめて修正させて提出する'] },
  't3-3': { q:'ハルシネーションによる損害が発生した場合の責任について、正しい理解はどれですか。', c:['AI事業者が全面的に責任を負うこととなり、利用した企業に責任は生じない','出力をそのまま用いて損害を与えた場合、利用した側が責任を問われる','AIが生成したものである旨を明示すれば、責任は一切生じない','無料版のAIを使った場合に限って、利用者が責任を負う'] },
  't3-4': { q:'要配慮個人情報について正しいものはどれですか。', c:['氏名と連絡先を組み合わせた情報がこれに該当することとされる','企業が保有する営業秘密を含んだ情報のことである','本人が自ら公開している情報であれば該当しない','病歴や信条などが該当し、より厳格な取扱いが必要となる'] },
  't3-5': { q:'自社の業務で法的リスクを評価する際、優先順位のつけ方として適切なものはどれですか。', c:['発生する可能性と、起きた場合の影響の大きさの両面から見る','発生する可能性が最も高いものから順に対処を進めていくこと','起きた場合の影響が最も大きいものだけに絞って対処する','法律の条文に明記されているものだけを評価の対象とする'] },

  't4-1': { q:'社内AI利用ガイドラインの適用範囲を定める際、適切なものはどれですか。', c:['正社員のみを対象とし、それ以外は各自の判断に委ねておく','役員は経営判断を担うため対象から外し、従業員のみに適用する','管理職だけに適用し、一般の従業員には周知のみ行っておく','雇用形態を問わず、業務でAIを使う者すべてを対象とする'] },
  't4-2': { q:'AIによる情報漏洩が疑われる事態が起きた場合、初動として最も優先すべきことはどれですか。', c:['原因を特定し、事実関係を固めたうえで報告を上げていく','事実を確認しつつ、定められた報告先へ速やかに報告する','社外への影響がないと確認できるまでは社内にも伏せておく','再発防止策を立てたうえで、経営層へまとめて報告を行う'] },
  't4-3': { q:'専任のAI管理担当者を置けない中小企業における現実的な体制はどれですか。', c:['体制が組めない以上は、AIの業務利用そのものを全面的に禁止する','担当を定めず、全従業員が各自の判断で運用していく','既存の情報管理担当が兼任し、権限と判断基準を明文化する','導入したAIベンダーに、すべての判断と管理を任せておく'] },
  't4-4': { q:'AI推進法（2025年6月公布・同年9月全面施行）について正しいものはどれですか。', c:['罰則のない基本法だが、AIを活用する事業者も対象に含まれる','AIを開発する事業者のみが対象であり、活用する側は含まれない','違反した事業者に対しては、課徴金が課される仕組みになっている','生成AIの業務利用を、国への事前の許可制とする内容になっている'] },
  't4-5': { q:'ガイドラインを策定した後、実効性を確保するために最も重要なことはどれですか。', c:['文書として完成させたうえで、共有フォルダに保存し公開しておく','周知と遵守確認の仕組みを設け、運用状況を定期的に点検する','想定される事態をすべて網羅した、可能な限り詳細な文書にする','外部の専門家に監修を依頼し、記載内容の正確性を担保しておく'] },

  't5-1': { q:'AI導入の効果を測る指標として、最も適切なものはどれですか。', c:['社内で導入したAIツールの種類と契約している総件数','AIに生成させた文章や資料の累計の分量と件数','対象とした業務にかかっていた時間がどう変わったか','社内でAIを使ったことのある従業員の割合とその推移'] },
  't5-2': { q:'AI導入を検討する際の順序として適切なものはどれですか。', c:['対象業務と扱う情報を特定し、それに合うツールを選ぶ','話題になっているツールを導入し、使える業務を探していく','最も費用の安いツールを選び、後から業務のほうを合わせる','全社員に自由に選ばせたうえで、後からひとつに統一する'] },
  't5-3': { q:'中小企業がAIツールを選定する際、コスト以外に必ず確認すべき点はどれですか。', c:['提供元の企業が設立から十分な年数を経過しているかどうか','画面の見やすさや操作の分かりやすさが十分であるか','対応している言語の種類が業務の範囲を満たしているか','入力データの学習利用の有無と、管理者が統制できるか'] },
  't5-4': { q:'AI導入を検討した結果「現時点では導入しない」と判断した場合、適切な対応はどれですか。', c:['導入しない以上は、検討の経緯を記録として残しておかなくてよい','判断した理由と、どのような条件が整えば導入するかを残す','今後は検討しない方針であることを、社内に明確に宣言する','競合他社が導入した時点で、自社も自動的に導入へ踏み切る'] },
  't5-5': { q:'業務にAIを使うことで最も効果が見込めるのは、一般にどのような作業ですか。', c:['経営の方向を決めるような、最終的な意思決定を伴う作業','契約や法令の解釈など、法的な責任を直接に伴う判断','顧客との信頼関係に直結する、対面での折衝や交渉の業務','定型的で反復が多く、下書きや情報の整理が中心の作業'] },

  't6-1': { q:'AIを使ったフィッシングメールが従来のものより見分けにくい理由はどれですか。', c:['添付されるファイルが必ず暗号化されるようになったため','日本語が自然になり、文面の不自然さで気づけなくなったため','送信元のメールアドレスの偽装が新たに可能になったため','ウイルス対策ソフトを無効化する機能を備えているため'] },
  't6-2': { q:'プロンプトインジェクションが成立しうる状況はどれですか。', c:['自社で作成した社内文書をAIに読ませて要約させるとき','一般的な知識について、AIに質問して回答を得るとき','外部から受け取った文書やWebページをAIに読ませるとき','数値の計算や表の集計をAIに任せて処理させるとき'] },
  't6-3': { q:'取引先から「振込口座を変更したい」という電話があった場合、適切な対応はどれですか。', c:['従来から把握している連絡先に、自社から掛け直して確認する','電話口の担当者の声を聞き、本人であることを確認して進める','折り返し先として伝えられた番号に掛け直し、内容を確認する','同じ内容のメールが届いていることをもって確認済みとする'] },
  't6-4': { q:'AIエージェントに業務を任せる際、最も注意すべき点はどれですか。', c:['処理の速度が人間より遅く、かえって時間がかかる場合がある','利用料金が従量課金であり、費用が想定を超える場合がある','日本語への対応が不完全で、指示が正しく伝わらない場合がある','人の確認を挟まずに送信や決済などの操作を実行しうること'] },
  't6-5': { q:'インシデント対応フローを整備する目的として最も適切なものはどれですか。', c:['責任の所在を明確にし、関係者への処分を行いやすくするため','監督官庁への報告が不要となるよう、社内で処理を完結するため','事態が起きた際に判断へ迷わず、初動を遅らせずに済むため','従業員にインシデントが起こりうることを日頃から意識させるため'] },

  't7-1': { q:'大規模言語モデルが事実と異なる内容を出力することがあるのは、根本的にはなぜですか。', c:['次に来る語を統計的に予測する仕組みで、真偽を検証していないから','学習に用いられたデータのほとんどが誤った内容のものになっているから','インターネットに接続しておらず、最新情報を持たないから','日本語への対応が英語に比べて十分でない状態にあるから'] },
  't7-2': { q:'社内データをAIに活用させる方法として、RAGが適しているのはどのような場合ですか。', c:['モデルの応答の文体や口調そのものを変えたい場合','学習に用いるデータを完全に非公開のままの状態で扱いたい場合','インターネット接続のない環境で動かしたい場合','頻繁に更新される社内文書を最新の状態で参照させたい場合'] },
  't7-3': { q:'プロンプトエンジニアリングの観点から、業務での指示文として望ましいものはどれですか。', c:['解釈の幅を残すため、できるだけ短く簡潔な指示にしておく','前提と目的、出力の形式、守るべき制約を明示しておく','専門用語を避け、誰にでも分かる平易な言葉だけを使う','関連する複数の依頼を、一度にまとめて指示しておく'] },
  't7-4': { q:'MCPのような仕組みでAIが外部ツールと連携できることで、新たに生じる論点はどれですか。', c:['連携の処理が挟まることで、AIの応答速度が低下すること','接続するツールが増えるほど、利用料金が上がっていくこと','接続先のデータを読み書きでき、権限設計を誤ると被害が広がること','連携先の仕様に合わせる必要があるため、日本語での指示が使えなくなること'] },
  't7-5': { q:'AI活用のロードマップを描く際、各段階で併せて検討すべきことはどれですか。', c:['その段階で扱う情報の範囲と、対応するセキュリティ対策','導入するツールの提供元が持つブランドとしての価値の高さ','同業の競合他社が導入に踏み切る時期とその内容','AI技術の分野で今後流行が移り変わっていく方向'] }
};

const TEST_TOTAL_PER_CHAPTER = COURSE.rules.testQuestions;   // course.json の rules
const TEST_PASS_COUNT_UI = COURSE.rules.testPass;
const _testAnswered = {};   // { questionId: true } 回答済み
const _testAttemptsLeft = {};
const _testRetry = {};   // { 章: true } 学び直しの途中（戻した設問にまだ答えていない）   // { questionId: 残りの回答回数 }（サーバーの値。0 の設問は学び直しの対象にしない）
const _testResultCh = {};   // { 章番号: {answered, correct, passed} }

function renderTestChapter(ch){
  const host = document.getElementById('test-body-'+ch);
  if(!host) return;
  let html = '';
  for(let i=1;i<=TEST_TOTAL_PER_CHAPTER;i++){
    const qid = `t${ch}-${i}`;
    const d = TEST_QUESTIONS[qid];
    if(!d) continue;
    const opts = d.c.map((t,idx)=>
      `<button class="test-opt" id="to-${qid}-${idx}" onclick="answerTest('${qid}',${idx})">${t}</button>`
    ).join('');
    html += `
      <div class="test-q" id="tq-${qid}">
        <div class="test-q-no">問 ${i} / ${TEST_TOTAL_PER_CHAPTER}</div>
        <p class="test-q-text">${d.q}</p>
        <div class="test-opts" id="tos-${qid}">${opts}</div>
        <div class="test-exp" id="te-${qid}"></div>
      </div>`;
  }
  host.innerHTML = html;
}

async function loadTestChapter(ch){
  const host = document.getElementById('test-body-'+ch);
  if(!host) return;
  if(!host.dataset.rendered){ renderTestChapter(ch); host.dataset.rendered = '1'; }
  const t = await workToken();
  if(!t) return;
  try{
    const r = await fetch(`${WORK_API}/api/test/${ch}`, { headers:{ Authorization:'Bearer '+t }});
    if(!r.ok) return;
    const d = await r.json();
    (d.answers||[]).forEach(a=>{
      _testAnswered[a.question_id] = true;
      if(a.attemptsLeft != null) _testAttemptsLeft[a.question_id] = a.attemptsLeft;
      lockTestQuestion(a.question_id, a.selected, a.is_correct, a.explanation);
    });
    _testResultCh[ch] = { answered:d.answered, correct:d.correct, passed:d.passed,
                          firstCorrect:d.firstCorrect, attempts:d.attempts, wrongIds:d.wrongIds||[] };
    updateTestSummary(ch);
  }catch(e){ /* 表示は継続 */ }
}

// 解説は引数で受け取る。クライアントには持たせない（回答前に先読みされるため）。
function lockTestQuestion(qid, selected, isCorrect, explanation){
  const box = document.getElementById('tos-'+qid);
  const exp = document.getElementById('te-'+qid);
  const d = TEST_QUESTIONS[qid];
  if(!box || !d) return;
  box.querySelectorAll('.test-opt').forEach((b,idx)=>{
    b.disabled = true;
    b.classList.add('is-done');
    if(idx === selected) b.classList.add(isCorrect ? 'is-correct' : 'is-wrong');
  });
  if(exp){
    exp.className = 'test-exp is-open ' + (isCorrect ? 'ok' : 'ng');
    exp.innerHTML = `<div class="test-verdict">${isCorrect ? '正解' : '不正解'}</div><div>${escHtml(explanation)}</div>`;
  }
}

async function answerTest(qid, idx){
  if(_testAnswered[qid]) return;
  const ch = parseInt(qid.split('-')[0].slice(1),10);
  const box = document.getElementById('tos-'+qid);
  if(box) box.querySelectorAll('.test-opt').forEach(b=>b.disabled = true);
  const t = await workToken();
  if(!t){
    if(box) box.querySelectorAll('.test-opt').forEach(b=>b.disabled = false);
    return;
  }
  try{
    const r = await fetch(`${WORK_API}/api/test/answer`, {
      method:'POST',
      headers:{ Authorization:'Bearer '+t, 'Content-Type':'application/json' },
      body: JSON.stringify({ questionId:qid, selected:idx })
    });
    const d = await r.json();
    if(!r.ok){
      if(box) box.querySelectorAll('.test-opt').forEach(b=>b.disabled = false);
      alert(d.error || '回答を記録できませんでした。時間をおいてもう一度お試しください。');
      return;
    }
    _testAnswered[qid] = true;
    if(d.attemptsLeft != null) _testAttemptsLeft[qid] = d.attemptsLeft;
    lockTestQuestion(qid, idx, d.isCorrect, d.explanation);
    const cur = _testResultCh[ch] || { answered:0, correct:0 };
    // 学び直し（前に間違えた設問への再回答）では、回答数は増やさず、正解なら正解数だけ増やす
    // （以前の数え方では、学び直しのたびに回答数が5を超えていった）
    const again = (cur.wrongIds || []).includes(qid);
    if(!again) cur.answered += 1;
    if(d.isCorrect) cur.correct += 1;
    cur.passed = cur.answered >= TEST_TOTAL_PER_CHAPTER && cur.correct >= TEST_PASS_COUNT_UI;
    if(!d.isCorrect){ cur.wrongIds = (cur.wrongIds || []).filter(x => x !== qid).concat(qid); }
    else if(cur.wrongIds){ cur.wrongIds = cur.wrongIds.filter(x => x !== qid); }
    if(cur.answered >= TEST_TOTAL_PER_CHAPTER) refreshRecords();
    _testResultCh[ch] = cur;
    updateTestSummary(ch);
  }catch(e){
    if(box) box.querySelectorAll('.test-opt').forEach(b=>b.disabled = false);
    alert('回答を送信できませんでした。通信環境をご確認のうえ、もう一度お試しください。');
  }
}

function updateTestSummary(ch){
  const el = document.getElementById('test-summary-'+ch);
  if(!el) return;
  const r = _testResultCh[ch] || { answered:0, correct:0, passed:false };
  if(r.answered === 0){ el.innerHTML = ''; return; }
  // 学び直しの途中は、戻した設問にすべて答えるまで結果を出さない（1問答えた時点で「不合格」が出ていた）
  if(_testRetry[ch]){
    const left = (r.wrongIds || []).filter(q => !_testAnswered[q]).length;
    if(left > 0){
      el.innerHTML = `<div class="test-retry-box"><p class="test-retry-msg">学び直し中：あと ${left} 問。第${ch}章のレッスンで確かめてから答えてください。</p>
        <div class="test-retry-links">${chapterLessonLinks(ch)}</div></div>`;
      return;
    }
    delete _testRetry[ch];
  }
  const done = r.answered >= TEST_TOTAL_PER_CHAPTER;
  if(!done){
    el.innerHTML = `<span class="test-prog">回答済み ${r.answered} / ${TEST_TOTAL_PER_CHAPTER} 問</span>`;
    return;
  }
  const att = (r.attempts && r.attempts > 1) ? `　<span class="test-att">受験 ${r.attempts} 回目</span>` : '';
  if(r.passed){
    // 次の行き先（次の章の最初のレッスン、最後なら認定試験・修了証）を出す
    const nx = typeof nextStep === 'function' ? nextStep() : null;
    el.innerHTML = `<span class="test-pass">${r.correct} / ${TEST_TOTAL_PER_CHAPTER} 問正解　合格</span>${att}`
      + (nx ? `<div style="margin-top:12px"><button type="button" class="form-submit ob-btn" onclick="${nx.onclick}">次へ：${escHtml(nx.label)} →</button></div>` : '');
    return;
  }
  const wrong = (r.wrongIds||[]).length;
  const retryable = (r.wrongIds||[]).filter(q => _testAttemptsLeft[q] !== 0).length;
  if(!retryable){
    el.innerHTML = `<span class="test-fail">${r.correct} / ${TEST_TOTAL_PER_CHAPTER} 問正解　合格ラインは ${TEST_PASS_COUNT_UI} 問です</span>${att}
      <div class="test-retry-box"><p class="test-retry-msg">間違えた設問は、回答できる回数（各${COURSE.rules.testMaxAttempts}回）を使い切りました。<br>
      章のレッスンを読み直したうえで、お問い合わせから運営までご連絡ください。</p></div>`;
    return;
  }
  el.innerHTML =
    `<span class="test-fail">${r.correct} / ${TEST_TOTAL_PER_CHAPTER} 問正解　合格ラインは ${TEST_PASS_COUNT_UI} 問です</span>${att}
     <div class="test-retry-box">
       <p class="test-retry-msg">間違えた ${wrong} 問について、対応するレッスンをもう一度確認してから再挑戦できます。<br>
       <span class="test-retry-sub">受講記録と進捗は保持されます。初回の回答は理解度の記録として残り、修了判定には最新の回答を用います。</span></p>
       <button class="test-retry-btn" onclick="retryTestChapter(${ch})">間違えた問題を学び直す</button>
     </div>`;
}

// 間違えた設問だけを未回答に戻し、参照レッスンへの導線を出す章のレッスンへのリンク（番号と題名）。
function chapterLessonLinks(ch){
  const c = COURSE.chapters.find(x => x.no === ch);
  if(!c) return '';
  return c.lessons.map(id => `<button type="button" class="work-nav-btn" data-tab="${id}" onclick="switchTab('${id}')">L${lessonNo(id)} ${escHtml(COURSE.titles[id] || '')}</button>`).join('');
}
function retryTestChapter(ch){
  _testRetry[ch] = true;
  const r = _testResultCh[ch] || {};
  // 回答回数を使い切った設問は戻さない
  const wrong = (r.wrongIds || []).filter(q => _testAttemptsLeft[q] !== 0);
  if(!wrong.length) return;
  wrong.forEach(qid => {
    _testAnswered[qid] = false;
    const box = document.getElementById('tos-'+qid);
    const exp = document.getElementById('te-'+qid);
    if(box){
      box.querySelectorAll('.test-opt').forEach(b=>{
        b.disabled = false;
        b.classList.remove('is-done','is-correct','is-wrong');
      });
    }
    if(exp){ exp.className = 'test-exp'; exp.innerHTML = ''; }
    const card = document.getElementById('tq-'+qid);
    if(card) card.classList.add('is-retry');
  });
  const el = document.getElementById('test-summary-'+ch);
  if(el){
    el.innerHTML = `<div class="test-retry-box">
      <p class="test-retry-msg">学び直しモードです。間違えた ${wrong.length} 問の範囲を、第${ch}章のレッスンで確かめてから再挑戦してください。</p>
      <div class="test-retry-links">${chapterLessonLinks(ch)}</div>
    </div>`;
  }
  const first = document.getElementById('tq-'+wrong[0]);
  if(first) first.scrollIntoView({behavior:'smooth', block:'center'});
}

window.loadTestChapter = loadTestChapter;
// ===================== /章末確認テスト =====================

// ===================== /章末ワーク =====================


// switchTab後にサイドバー・学習UIを更新
window.switchTab = function(name){
  if(!canAccessTab(name)) return;
  if(window._aslTestFocus && name !== 'work' + window._aslTestFocus) window._aslTestFocus = 0;
  if(currentUser){
    // 前回の続きから再開できるよう、最後に開いたタブを控える（ユーザーごと）
    store.set('asl_last_tab', { u: currentUser.id, t: name });
    // ログイン後はshowActiveLessonで直接切替
    showActiveLesson(name);
    const completed = loadCompleted();
    buildSideNav(completed);
    const lc = document.getElementById('learn-content');
    if(lc) lc.scrollTop = 0;
  } else {
    showTabGuest(name);
  }
  // 章末ワークタブなら内容を読み込む
  if(/^work\d+$/.test(name)){
    const _ch = parseInt(name.slice(4),10);
    loadWorkChapter(_ch);
    loadTestChapter(_ch);
  }
};
window.openGuideModal  = function(){ document.getElementById('guide-modal-overlay').classList.add('show'); };
window.closeGuideModal = function(){ document.getElementById('guide-modal-overlay').classList.remove('show'); };

