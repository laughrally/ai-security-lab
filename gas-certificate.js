// ===================================================
// AI Security Lab｜修了証発行 Google Apps Script
// 2026-08-13 改訂
// ===================================================
// 【変更点】
//  - doGet を廃止し doPost のみに変更。共有シークレットを検証する。
//    （従来は URL を知っていれば誰でも任意の氏名で修了証を発行できた）
//  - 受講状況の検証は Cloudflare Worker 側で実施済み。ここは発行専任。
//  - 修了証に 講座名／プラン／レッスン数／標準学習時間／証明書番号／修了条件 を記載。
//
// 【設定手順】
//  1. このコードを Apps Script プロジェクトに貼り付け
//  2. プロジェクトの設定 → スクリプトプロパティに以下を追加
//       CERT_SHARED_SECRET : Worker と同じランダム文字列
//       LOG_SPREADSHEET_ID : 記録用スプレッドシートのID（任意）
//  3. デプロイ → 新しいデプロイ → ウェブアプリ
//       実行ユーザー：自分 ／ アクセス：全員
//  4. 発行された /exec URL を Worker の GAS_CERT_URL に設定
//
// 【注意】このファイルは描画専任である。
//  studyHours / lessonsTotal / planLabel は Worker から渡される値であり、
//  ここが定義元ではない。標準学習時間やレッスン数の表記を変更する場合は
//  Cloudflare Worker (ai-security-lab-stripe) 側の定数を修正すること。
//
// 【デプロイ時の注意】
//  コードを保存しただけではウェブアプリURLは旧コードを返し続ける。
//  必ず「デプロイを管理」から新しいバージョンを発行すること。
//  （2026-08-16 の Unauthorized はこの再デプロイ漏れが原因）
// ===================================================

const NOTIFY_EMAIL = 'laughrally@gmail.com';
const FROM_NAME    = 'AI Security Lab';
const FROM_EMAIL   = 'info@laughrally.tech';
const ISSUER_NAME  = '合同会社 LaughRally';
const ISSUER_ADDR  = '〒107-0062 東京都港区南青山2-2-15';

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET は受け付けない（誤って叩かれた場合の副作用を防ぐ）
function doGet() {
  return jsonOut({ status: 'error', message: 'Method not allowed' });
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);

    // --- 共有シークレットの検証 ---
    const expected = PropertiesService.getScriptProperties()
      .getProperty('CERT_SHARED_SECRET');
    if (!expected || p.secret !== expected) {
      return jsonOut({ status: 'error', message: 'Unauthorized' });
    }

    // --- 発行する書類の種類で分岐 ---
    // type 未指定は従来どおり修了証として扱う（旧Workerとの互換のため）
    if (p.type === 'application') {
      return handleApplicationDoc(p);
    }

    const data = {
      certNo:       p.certNo       || '',
      name:         p.name         || '',
      company:      p.company      || '',
      email:        p.email        || '',
      courseName:   p.courseName   || 'AI Security Lab 生成AI×セキュリティ eラーニング',
      planLabel:    p.planLabel    || '',
      lessonsTotal: p.lessonsTotal || 0,
      studyHours:   p.studyHours   || '',
      completedDate: p.completedDate || '',
      examPassed:   !!p.examPassed,
    };
    if (!data.name || !data.email || !data.certNo) {
      return jsonOut({ status: 'error', message: 'Missing required fields' });
    }

    const date = new Date().toLocaleDateString('ja-JP',
      { year: 'numeric', month: 'long', day: 'numeric' });

    const pdfBlob = createCertificatePDF(data, date);

    GmailApp.sendEmail(data.email, '【AI Security Lab】修了証のお届け', '', {
      htmlBody: buildClientEmailBody(data),
      attachments: [pdfBlob],
      name: FROM_NAME,
      replyTo: FROM_EMAIL,
    });

    GmailApp.sendEmail(NOTIFY_EMAIL,
      '【修了証発行】' + data.name + 'さん（' + data.certNo + '）', '', {
      htmlBody: buildNotifyEmailBody(data, date),
      name: 'AI Security Lab 通知',
    });

    logToSheet(data, date);

    return jsonOut({ status: 'ok', certNo: data.certNo });

  } catch (err) {
    return jsonOut({ status: 'error', message: err.message });
  }
}

// ===== 修了証PDF生成 =====
function createCertificatePDF(data, date) {
  const html = buildCertificateHTML(data, date);
  const blob = Utilities.newBlob(html, 'text/html', 'certificate.html');
  const tempFile = DriveApp.createFile(blob);
  const pdfBlob = tempFile.getAs('application/pdf');
  pdfBlob.setName('AI_Security_Lab_修了証_' + data.name + '.pdf');
  tempFile.setTrashed(true);
  return pdfBlob;
}

function esc(t) {
  return String(t == null ? '' : t)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ===== 修了証HTMLテンプレート =====
function buildCertificateHTML(d, date) {
  const companyLine = d.company
    ? '<div class="company">' + esc(d.company) + '</div>' : '';
  const examLine = d.examPassed
    ? '<br>認定試験 合格（全60問中48問以上正解）' : '';

  return '<!DOCTYPE html>\n' +
'<html lang="ja"><head><meta charset="UTF-8"><style>\n' +
"  @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+JP:wght@400;700&display=swap');\n" +
'  * { margin:0; padding:0; box-sizing:border-box; }\n' +
'  body { width:210mm; min-height:280mm; font-family:"Noto Serif JP","MS Mincho",serif;\n' +
'         background:#fff; display:flex; align-items:center; justify-content:center; }\n' +
'  .cert { width:210mm; min-height:280mm; position:relative; display:flex; flex-direction:column;\n' +
'          align-items:center; justify-content:center; padding:20mm 18mm; background:#fff; text-align:center; }\n' +
'  .cert::before { content:""; position:absolute; inset:10mm; border:2px solid #1a1a1a; }\n' +
'  .cert::after  { content:""; position:absolute; inset:12mm; border:0.5px solid #999; }\n' +
'  .logo { font-size:10pt; letter-spacing:0.35em; color:#666; margin-bottom:7mm; }\n' +
'  .logo span { color:#c00; }\n' +
'  .title { font-size:30pt; font-weight:700; letter-spacing:0.6em; color:#1a1a1a;\n' +
'           margin-bottom:5mm; padding-right:0.6em; }\n' +
'  .divider { font-size:14pt; color:#c00; letter-spacing:0.5em; margin-bottom:8mm; }\n' +
'  .company { font-size:12pt; color:#555; letter-spacing:0.1em; margin-bottom:3mm; }\n' +
'  .name { font-size:26pt; font-weight:700; color:#1a1a1a; letter-spacing:0.2em;\n' +
'          border-bottom:1.5px solid #1a1a1a; padding-bottom:3mm; padding-right:0.2em; margin-bottom:3mm; }\n' +
'  .suffix { font-size:12pt; color:#444; letter-spacing:0.3em; margin-bottom:8mm; }\n' +
'  .body-text { font-size:10.5pt; color:#333; line-height:2.1; letter-spacing:0.05em; margin-bottom:7mm; }\n' +
'  .course-name { font-size:12.5pt; font-weight:700; color:#1a1a1a; letter-spacing:0.06em; }\n' +
'  .detail { font-size:9pt; color:#444; line-height:1.9; border-top:0.5px solid #ccc;\n' +
'            border-bottom:0.5px solid #ccc; padding:4mm 0; margin-bottom:6mm; width:130mm; }\n' +
'  .detail td { padding:0.8mm 3mm; text-align:left; font-size:9pt; }\n' +
'  .detail td.k { color:#777; width:38mm; }\n' +
'  .footer { position:absolute; bottom:20mm; left:20mm; right:20mm;\n' +
'            display:flex; justify-content:space-between; align-items:flex-end; }\n' +
'  .issue-date { font-size:9pt; color:#666; letter-spacing:0.1em; text-align:left; }\n' +
'  .cert-no { font-size:8pt; color:#999; margin-top:1mm; letter-spacing:0.08em; }\n' +
'  .issuer { text-align:right; }\n' +
'  .issuer-name { font-size:11pt; font-weight:700; color:#1a1a1a; letter-spacing:0.1em; }\n' +
'  .issuer-sub { font-size:8pt; color:#888; margin-top:1mm; }\n' +
'</style></head><body>\n' +
'<div class="cert">\n' +
'  <div class="logo">AI <span>SECURITY</span> LAB</div>\n' +
'  <div class="title">修　了　証</div>\n' +
'  <div class="divider">― ✦ ―</div>\n' +
   companyLine + '\n' +
'  <div class="name">' + esc(d.name) + '</div>\n' +
'  <div class="suffix">殿</div>\n' +
'  <div class="body-text">\n' +
'    あなたは下記の課程を修了されたことをここに証します。<br><br>\n' +
'    <span class="course-name">' + esc(d.courseName) + '</span>\n' +
'  </div>\n' +
'  <table class="detail">\n' +
'    <tr><td class="k">プラン</td><td>' + esc(d.planLabel) + '</td></tr>\n' +
'    <tr><td class="k">修了レッスン数</td><td>全' + esc(d.lessonsTotal) + 'レッスン</td></tr>\n' +
'    <tr><td class="k">標準学習時間</td><td>' + esc(d.studyHours) + '（休憩時間を除く）</td></tr>\n' +
'    <tr><td class="k">修了条件</td><td>全レッスンの受講完了、全章の章末ワーク提出および章末確認テスト合格' + examLine + '</td></tr>\n' +
   (d.completedDate ? '    <tr><td class="k">修了日</td><td>' + esc(d.completedDate) + '</td></tr>\n' : '') +
'  </table>\n' +
'  <div class="footer">\n' +
'    <div>\n' +
'      <div class="issue-date">発行日　' + esc(date) + '</div>\n' +
'      <div class="cert-no">証明書番号　' + esc(d.certNo) + '</div>\n' +
'    </div>\n' +
'    <div class="issuer">\n' +
'      <div class="issuer-name">' + ISSUER_NAME + '</div>\n' +
'      <div class="issuer-sub">' + ISSUER_ADDR + '</div>\n' +
'    </div>\n' +
'  </div>\n' +
'</div></body></html>';
}

// ===== 受講者向けメール =====
function buildClientEmailBody(d) {
  const companyLine = d.company ? '（' + esc(d.company) + '）' : '';
  return '' +
'<div style="font-family:\'Helvetica Neue\',Arial,sans-serif;max-width:560px;margin:0 auto;padding:40px 24px;color:#222">' +
'  <div style="font-size:13px;color:#c00;letter-spacing:0.15em;margin-bottom:8px">AI SECURITY LAB</div>' +
'  <h1 style="font-size:22px;font-weight:700;margin-bottom:24px;border-bottom:2px solid #eee;padding-bottom:16px">修了証のお届け</h1>' +
'  <p style="font-size:15px;line-height:1.8;margin-bottom:16px">' + esc(d.name) + companyLine + ' 様</p>' +
'  <p style="font-size:15px;line-height:1.8;margin-bottom:16px">' +
'    この度は「' + esc(d.courseName) + '」（' + esc(d.planLabel) + '）の全課程を修了されました。<br>誠におめでとうございます。</p>' +
'  <p style="font-size:15px;line-height:1.8;margin-bottom:24px">修了証PDFを添付にてお送りいたします。</p>' +
'  <table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:24px">' +
'    <tr><td style="padding:7px 10px;background:#f5f5f5;width:130px">証明書番号</td><td style="padding:7px 10px;border-bottom:1px solid #eee">' + esc(d.certNo) + '</td></tr>' +
'    <tr><td style="padding:7px 10px;background:#f5f5f5">修了レッスン数</td><td style="padding:7px 10px;border-bottom:1px solid #eee">全' + esc(d.lessonsTotal) + 'レッスン</td></tr>' +
'    <tr><td style="padding:7px 10px;background:#f5f5f5">標準学習時間</td><td style="padding:7px 10px">' + esc(d.studyHours) + '</td></tr>' +
     (d.completedDate ? '<tr><td style="padding:7px 10px;background:#f5f5f5">修了日</td><td style="padding:7px 10px">' + esc(d.completedDate) + '</td></tr>' : '') +
'  </table>' +
'  <div style="background:#f9f9f9;border-left:3px solid #c00;padding:16px 20px;margin-bottom:24px;font-size:14px;line-height:1.7">' +
'    受講日時・学習時間・理解度テストの結果をまとめた「受講実績レポート」は、受講画面の「進捗」タブからPDF・CSVで出力できます。</div>' +
'  <p style="font-size:13px;color:#888;line-height:1.7">──────────────────<br>' +
     ISSUER_NAME + '<br>' + FROM_EMAIL + '<br>' + ISSUER_ADDR + '</p>' +
'</div>';
}

// ===== 管理者向け通知 =====
function buildNotifyEmailBody(d, date) {
  const row = function(k, v) {
    return '<tr><td style="padding:8px 12px;background:#f5f5f5;font-weight:bold;width:130px">' + k +
           '</td><td style="padding:8px 12px;border-bottom:1px solid #eee">' + esc(v) + '</td></tr>';
  };
  return '' +
'<div style="font-family:sans-serif;max-width:480px;margin:0 auto;padding:32px 24px;color:#222">' +
'  <h2 style="font-size:18px;margin-bottom:16px">📜 修了証を発行しました</h2>' +
'  <table style="width:100%;border-collapse:collapse;font-size:14px">' +
     row('証明書番号', d.certNo) +
     row('会社名', d.company || '（個人）') +
     row('氏名', d.name) +
     row('メール', d.email) +
     row('プラン', d.planLabel) +
     row('レッスン数', '全' + d.lessonsTotal + 'レッスン') +
     row('発行日', date) +
'  </table>' +
'  <p style="margin-top:16px;font-size:13px;color:#888">受講状況はWorker側で検証済みです。PDFは受講者へ送信されました。</p>' +
'</div>';
}

// ===== スプレッドシート記録 =====
function logToSheet(d, date) {
  try {
    const id = PropertiesService.getScriptProperties()
      .getProperty('LOG_SPREADSHEET_ID');
    if (!id) return;
    const ss = SpreadsheetApp.openById(id);
    const sheet = ss.getSheetByName('修了証発行') || ss.insertSheet('修了証発行');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['発行日', '証明書番号', '会社名', '氏名',
                       'メールアドレス', 'プラン', 'レッスン数', '認定試験']);
    }
    sheet.appendRow([date, d.certNo, d.company, d.name, d.email,
                     d.planLabel, d.lessonsTotal, d.examPassed ? '合格' : '—']);
  } catch (e) {
    // 記録に失敗してもメール送信は継続する
  }
}


// ===================================================
// 受講申込確認書（法人一括申込）
// 2026-08-20 追加
// ---------------------------------------------------
// 補助金の実績報告では、見積書のほかに「契約書または請書」および
// 「受講を申し込んだことが分かる書類」の提出を求められる。
// この1枚で両方を兼ねられるよう、申込日・受講者名・講座名・プラン・
// 金額・受講開始予定日を記載する。
//
// 【注意】このファイルは描画専任である。
//  planLabel / range / studyHours / 金額 はすべて Worker から渡される値であり、
//  ここが定義元ではない。表記を変えるときは Worker 側の CORP_PLAN を直すこと。
// ===================================================

function handleApplicationDoc(p) {
  const d = {
    appNo:        p.appNo        || '',
    company:      p.company      || '',
    contactName:  p.contactName  || '',
    contactEmail: p.contactEmail || '',
    courseName:   p.courseName   || 'AI Security Lab 生成AI×セキュリティ eラーニング',
    planLabel:    p.planLabel    || '',
    range:        p.range        || '',
    studyHours:   p.studyHours   || '',
    seats:        Number(p.seats || 0),
    unitPrice:    Number(p.unitPrice || 0),
    totalPrice:   Number(p.totalPrice || 0),
    participants: Array.isArray(p.participants) ? p.participants : [],
    note:         p.note || '',
  };
  if (!d.appNo || !d.company || !d.contactEmail || !d.participants.length) {
    return jsonOut({ status: 'error', message: 'Missing required fields' });
  }

  const date = new Date().toLocaleDateString('ja-JP',
    { year: 'numeric', month: 'long', day: 'numeric' });

  const pdfBlob = createApplicationPDF(d, date);

  GmailApp.sendEmail(d.contactEmail,
    '【AI Security Lab】受講申込確認書のお届け（' + d.appNo + '）', '', {
    htmlBody: buildApplicationEmailBody(d, date),
    attachments: [pdfBlob],
    name: FROM_NAME,
    replyTo: FROM_EMAIL,
  });

  GmailApp.sendEmail(NOTIFY_EMAIL,
    '【法人申込】' + d.company + '（' + d.seats + '名・' + d.appNo + '）', '', {
    htmlBody: buildApplicationNotifyBody(d, date),
    attachments: [pdfBlob],
    name: 'AI Security Lab 通知',
  });

  return jsonOut({ status: 'ok', appNo: d.appNo });
}

function createApplicationPDF(d, date) {
  const html = buildApplicationHTML(d, date);
  const blob = Utilities.newBlob(html, 'text/html', 'application.html');
  const tempFile = DriveApp.createFile(blob);
  const pdfBlob = tempFile.getAs('application/pdf');
  pdfBlob.setName('AI_Security_Lab_受講申込確認書_' + d.company + '.pdf');
  tempFile.setTrashed(true);
  return pdfBlob;
}

function yen(n) {
  return Number(n || 0).toLocaleString('ja-JP');
}

function buildApplicationHTML(d, date) {
  var rows = '';
  for (var i = 0; i < d.participants.length; i++) {
    var pt = d.participants[i];
    rows += '<tr><td class="n">' + (i + 1) + '</td><td>' + esc(pt.name) +
            '</td><td class="m">' + esc(pt.email) + '</td></tr>';
  }
  var noteBlock = d.note
    ? '<div class="sec"><div class="sec-h">備考</div><div class="note">' + esc(d.note) + '</div></div>'
    : '';

  return '<!DOCTYPE html>\n' +
'<html lang="ja"><head><meta charset="UTF-8"><style>\n' +
'  * { margin:0; padding:0; box-sizing:border-box; }\n' +
'  body { width:210mm; min-height:290mm; padding:18mm 16mm;\n' +
'         font-family:"Noto Sans JP","Hiragino Sans","MS Gothic",sans-serif;\n' +
'         color:#111; font-size:10.5pt; line-height:1.7; background:#fff; }\n' +
'  .head { display:flex; justify-content:space-between; align-items:flex-start; }\n' +
'  h1 { font-size:19pt; font-weight:700; letter-spacing:.14em; border-bottom:2px solid #111;\n' +
'       padding-bottom:6px; margin-bottom:18px; }\n' +
'  .meta { text-align:right; font-size:9.5pt; color:#333; line-height:1.9; }\n' +
'  .to { font-size:13pt; font-weight:700; margin:14px 0 4px; }\n' +
'  .to-sub { font-size:10pt; color:#444; margin-bottom:18px; }\n' +
'  .lead { margin:0 0 18px; }\n' +
'  .total-box { border:2px solid #111; padding:12px 16px; margin-bottom:20px;\n' +
'               display:flex; justify-content:space-between; align-items:baseline; }\n' +
'  .total-box .lbl { font-size:11pt; font-weight:700; letter-spacing:.08em; }\n' +
'  .total-box .val { font-size:18pt; font-weight:700; }\n' +
'  .sec { margin-bottom:18px; }\n' +
'  .sec-h { font-size:10pt; font-weight:700; border-left:4px solid #111; padding-left:8px; margin-bottom:8px; }\n' +
'  table { width:100%; border-collapse:collapse; font-size:9.5pt; }\n' +
'  th, td { border:1px solid #bbb; padding:6px 9px; text-align:left; vertical-align:top; }\n' +
'  th { background:#f2f2f2; font-weight:700; white-space:nowrap; }\n' +
'  th.k { width:34mm; }\n' +
'  td.n { width:10mm; text-align:center; }\n' +
'  td.m { font-family:"MS Gothic",monospace; font-size:9pt; }\n' +
'  td.r { text-align:right; }\n' +
'  .note { border:1px solid #bbb; padding:8px 10px; font-size:9.5pt; white-space:pre-wrap; }\n' +
'  .issuer { margin-top:26px; border-top:1px solid #ccc; padding-top:12px; font-size:9.5pt; line-height:1.9; }\n' +
'  .issuer .nm { font-size:11pt; font-weight:700; }\n' +
'  .fine { margin-top:14px; font-size:8.5pt; color:#555; line-height:1.8; }\n' +
'</style></head><body>\n' +
'<h1>受 講 申 込 確 認 書</h1>\n' +
'<div class="meta">申込番号：' + esc(d.appNo) + '<br>発行日：' + esc(date) + '</div>\n' +
'<div class="to">' + esc(d.company) + '　御中</div>\n' +
'<div class="to-sub">ご担当者：' + esc(d.contactName) + ' 様</div>\n' +
'<p class="lead">下記のとおりお申し込みを承りました。</p>\n' +
'<div class="total-box"><span class="lbl">お申込金額（税込）</span><span class="val">¥' + yen(d.totalPrice) + '</span></div>\n' +
'<div class="sec"><div class="sec-h">お申し込み内容</div>\n' +
'<table>\n' +
'<tr><th class="k">講座名</th><td>' + esc(d.courseName) + '</td></tr>\n' +
'<tr><th class="k">プラン</th><td>' + esc(d.planLabel) + '</td></tr>\n' +
'<tr><th class="k">受講範囲</th><td>' + esc(d.range) + '</td></tr>\n' +
'<tr><th class="k">標準学習時間</th><td>' + esc(d.studyHours) + '（休憩時間を除く／1名あたり）</td></tr>\n' +
'<tr><th class="k">受講人数</th><td>' + d.seats + ' 名</td></tr>\n' +
'<tr><th class="k">受講料（税込）</th><td>1名につき ¥' + yen(d.unitPrice) + '　×　' + d.seats + ' 名　＝　¥' + yen(d.totalPrice) + '</td></tr>\n' +
'<tr><th class="k">受講形式</th><td>オンライン（eラーニング）／買い切り・受講期限なし</td></tr>\n' +
'<tr><th class="k">受講開始予定日</th><td>ご入金の確認後、アカウントを発行した日</td></tr>\n' +
'</table></div>\n' +
'<div class="sec"><div class="sec-h">受講者</div>\n' +
'<table><tr><th class="n">#</th><th>お名前</th><th>メールアドレス</th></tr>\n' + rows + '</table></div>\n' +
noteBlock +
'<div class="issuer"><div class="nm">' + esc(ISSUER_NAME) + '</div>' + esc(ISSUER_ADDR) + '<br>' +
'代表社員　佐藤 誓哉　／　' + esc(FROM_EMAIL) + '</div>\n' +
'<div class="fine">・アカウントは受講者1名につき1つ発行いたします。<br>\n' +
'・当社は消費税の免税事業者のため、適格請求書（インボイス）の発行はできません。<br>\n' +
'・本書は上記のお申し込みを承ったことを証するものです。</div>\n' +
'</body></html>';
}

function buildApplicationEmailBody(d, date) {
  var list = '';
  for (var i = 0; i < d.participants.length; i++) {
    list += '<li>' + esc(d.participants[i].name) + '（' + esc(d.participants[i].email) + '）</li>';
  }
  return '<div style="font-family:sans-serif;font-size:14px;line-height:1.9;color:#111">' +
    esc(d.company) + '<br>' + esc(d.contactName) + ' 様<br><br>' +
    'お世話になっております。合同会社LaughRally 佐藤です。<br><br>' +
    'このたびは AI Security Lab へお申し込みいただき、誠にありがとうございます。<br>' +
    '受講申込確認書を添付にてお送りいたします。<br><br>' +
    '<b>申込番号</b>：' + esc(d.appNo) + '<br>' +
    '<b>プラン</b>：' + esc(d.planLabel) + '<br>' +
    '<b>受講人数</b>：' + d.seats + ' 名<br>' +
    '<b>お申込金額（税込）</b>：¥' + yen(d.totalPrice) + '<br><br>' +
    '<b>受講者</b><ul>' + list + '</ul>' +
    'このあと、請求書を別途お送りいたします。<br>' +
    'ご入金の確認後、受講者さまごとにアカウントを発行し、パスワード設定用のメールをお送りします。<br><br>' +
    'ご不明点があればいつでもご連絡ください。<br>' +
    '引き続きどうぞよろしくお願いいたします。<br><br>' +
    '<span style="font-size:12px;color:#666">' + esc(ISSUER_NAME) + '　佐藤 誓哉<br>' +
    esc(ISSUER_ADDR) + '<br>' + esc(FROM_EMAIL) + '</span><br><br>' +
    '<span style="font-size:11px;color:#999">※このメールにお心当たりがない場合は、お手数ですが破棄してください。</span>' +
    '</div>';
}

function buildApplicationNotifyBody(d, date) {
  var list = '';
  for (var i = 0; i < d.participants.length; i++) {
    list += '<li>' + esc(d.participants[i].name) + '（' + esc(d.participants[i].email) + '）</li>';
  }
  return '<div style="font-family:sans-serif;font-size:14px;line-height:1.8">' +
    '<b>法人申込を受け付けました</b><br><br>' +
    '申込番号：' + esc(d.appNo) + '<br>' +
    '発行日：' + esc(date) + '<br>' +
    '会社名：' + esc(d.company) + '<br>' +
    'ご担当者：' + esc(d.contactName) + '（' + esc(d.contactEmail) + '）<br>' +
    'プラン：' + esc(d.planLabel) + '<br>' +
    '人数：' + d.seats + ' 名<br>' +
    '金額（税込）：¥' + yen(d.totalPrice) + '<br><br>' +
    '受講者<ul>' + list + '</ul>' +
    (d.note ? '備考：<br><pre style="white-space:pre-wrap;font-family:inherit">' + esc(d.note) + '</pre>' : '') +
    '<br>次にやること：請求書の発行 → 入金確認 → アカウント発行' +
    '</div>';
}
