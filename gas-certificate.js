// ===================================================
// AI Security Lab｜修了証自動発行 Google Apps Script
// ===================================================
// 設定手順：
// 1. https://script.google.com で新規プロジェクト作成
// 2. このコードを貼り付け
// 3. 「デプロイ」→「新しいデプロイ」→ ウェブアプリ
//    - 実行ユーザー：自分
//    - アクセス：全員
// 4. デプロイURLをindex.htmlのWORKER_URLに設定
// ===================================================

const NOTIFY_EMAIL = 'laughrally@gmail.com';       // 通知先（セイヤさん）
const FROM_NAME    = 'AI Security Lab';             // 送信者名
const FROM_EMAIL   = 'info@laughrally.tech';        // 送信元（Gmailエイリアス設定済み）

function doGet(e) {
  try {
    const company = e.parameter['会社名']    || '';
    const name    = e.parameter['担当者名']  || '';
    const email   = e.parameter['メールアドレス'] || '';
    const date    = new Date().toLocaleDateString('ja-JP', {year:'numeric', month:'long', day:'numeric'});

    // 1. 修了証PDFを生成
    const pdfBlob = createCertificatePDF(company, name, date);

    // 2. クライアントに修了証を送信
    GmailApp.sendEmail(
      email,
      '【AI Security Lab】修了証のお届け',
      '',
      {
        htmlBody: buildClientEmailBody(name, company),
        attachments: [pdfBlob],
        name: FROM_NAME,
        replyTo: FROM_EMAIL,
      }
    );

    // 3. セイヤさんに通知
    GmailApp.sendEmail(
      NOTIFY_EMAIL,
      '【修了証申請】' + name + 'さんから申請がありました',
      '',
      {
        htmlBody: buildNotifyEmailBody(company, name, email, date),
        name: 'AI Security Lab 通知',
      }
    );

    // 4. スプレッドシートに記録
    logToSheet(company, name, email, date);

    return ContentService.createTextOutput(JSON.stringify({status:'ok'}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status:'error', message: err.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ===== 修了証PDF生成 =====
function createCertificatePDF(company, name, date) {
  const html = buildCertificateHTML(company, name, date);
  const blob = Utilities.newBlob(html, 'text/html', 'certificate.html');

  // Google DriveにHTMLを一時保存してPDF変換
  const tempFile = DriveApp.createFile(blob);
  const pdfBlob  = tempFile.getAs('application/pdf');
  pdfBlob.setName('AI_Security_Lab_修了証_' + name + '.pdf');
  tempFile.setTrashed(true); // 一時ファイルを削除

  return pdfBlob;
}

// ===== 修了証HTMLテンプレート =====
function buildCertificateHTML(company, name, date) {
  const companyLine = company ? `<div class="company">${company}</div>` : '';

  return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Serif+JP:wght@400;700&display=swap');
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    width: 210mm;
    min-height: 280mm;
    font-family: 'Noto Serif JP', 'MS Mincho', serif;
    background: #fff;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .cert {
    width: 210mm;
    min-height: 280mm;
    position: relative;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    padding: 20mm 18mm;
    background: #fff;
    text-align: center;
  }
  .cert::before {
    content: '';
    position: absolute;
    inset: 10mm;
    border: 2px solid #1a1a1a;
  }
  .cert::after {
    content: '';
    position: absolute;
    inset: 12mm;
    border: 0.5px solid #999;
  }
  .logo {
    font-size: 10pt;
    letter-spacing: 0.35em;
    color: #666;
    margin-bottom: 8mm;
  }
  .logo span { color: #c00; }
  .title {
    font-size: 30pt;
    font-weight: 700;
    letter-spacing: 0.6em;
    color: #1a1a1a;
    margin-bottom: 6mm;
    padding-right: 0.6em;
  }
  .divider {
    font-size: 14pt;
    color: #c00;
    letter-spacing: 0.5em;
    margin-bottom: 10mm;
  }
  .company {
    font-size: 12pt;
    color: #555;
    letter-spacing: 0.1em;
    margin-bottom: 3mm;
  }
  .name {
    font-size: 28pt;
    font-weight: 700;
    color: #1a1a1a;
    letter-spacing: 0.2em;
    border-bottom: 1.5px solid #1a1a1a;
    padding-bottom: 3mm;
    padding-right: 0.2em;
    margin-bottom: 3mm;
  }
  .suffix {
    font-size: 12pt;
    color: #444;
    letter-spacing: 0.3em;
    margin-bottom: 10mm;
  }
  .body-text {
    font-size: 10.5pt;
    color: #333;
    line-height: 2.2;
    letter-spacing: 0.05em;
    margin-bottom: 10mm;
  }
  .course-name {
    font-size: 12pt;
    font-weight: 700;
    color: #1a1a1a;
    letter-spacing: 0.08em;
  }
  .footer {
    position: absolute;
    bottom: 20mm;
    left: 20mm;
    right: 20mm;
    display: flex;
    justify-content: space-between;
    align-items: flex-end;
  }
  .issue-date { font-size: 9pt; color: #666; letter-spacing: 0.1em; }
  .issuer { text-align: right; }
  .issuer-name { font-size: 11pt; font-weight: 700; color: #1a1a1a; letter-spacing: 0.1em; }
  .issuer-sub { font-size: 8pt; color: #888; margin-top: 1mm; }
</style>
</head>
<body>
<div class="cert">
  <div class="logo">AI <span>SECURITY</span> LAB</div>
  <div class="title">修　了　証</div>
  <div class="divider">― ✦ ―</div>
  ${companyLine}
  <div class="name">${name}</div>
  <div class="suffix">殿</div>
  <div class="body-text">
    あなたは下記の課程を修了されたことをここに証します。<br><br>
    <span class="course-name">AI Security Lab｜AI時代のセキュリティ実践講座</span><br>
    全17レッスン・修了
  </div>
  <div class="footer">
    <div class="issue-date">発行日　${date}</div>
    <div class="issuer">
      <div class="issuer-name">合同会社 LaughRally</div>
      <div class="issuer-sub">〒107-0062 東京都港区南青山2-2-15</div>
    </div>
  </div>
</div>
</body>
</html>`;
}

// ===== クライアント向けメール本文 =====
function buildClientEmailBody(name, company) {
  const companyLine = company ? `（${company}）` : '';
  return `
<div style="font-family:'Helvetica Neue',Arial,sans-serif;max-width:560px;margin:0 auto;padding:40px 24px;color:#222">
  <div style="font-size:13px;color:#c00;letter-spacing:0.15em;margin-bottom:8px">AI SECURITY LAB</div>
  <h1 style="font-size:22px;font-weight:700;margin-bottom:24px;border-bottom:2px solid #eee;padding-bottom:16px">修了証のお届け</h1>
  <p style="font-size:15px;line-height:1.8;margin-bottom:16px">
    ${name}${companyLine} 様
  </p>
  <p style="font-size:15px;line-height:1.8;margin-bottom:16px">
    この度は AI Security Lab の全課程を修了されました。<br>
    誠におめでとうございます。
  </p>
  <p style="font-size:15px;line-height:1.8;margin-bottom:24px">
    修了証PDFを添付にてお送りいたします。<br>
    ご活用いただければ幸いです。
  </p>
  <div style="background:#f9f9f9;border-left:3px solid #c00;padding:16px 20px;margin-bottom:24px;font-size:14px;line-height:1.7">
    引き続きAI活用に関するご相談は、いつでもお気軽にお問い合わせください。
  </div>
  <p style="font-size:13px;color:#888;line-height:1.7">
    ──────────────────<br>
    合同会社 LaughRally<br>
    info@laughrally.tech<br>
    〒107-0062 東京都港区南青山2-2-15
  </p>
</div>`;
}

// ===== セイヤさん向け通知メール =====
function buildNotifyEmailBody(company, name, email, date) {
  return `
<div style="font-family:sans-serif;max-width:480px;margin:0 auto;padding:32px 24px;color:#222">
  <h2 style="font-size:18px;margin-bottom:16px">📜 修了証申請が届きました</h2>
  <table style="width:100%;border-collapse:collapse;font-size:14px">
    <tr><td style="padding:8px 12px;background:#f5f5f5;font-weight:bold;width:120px">会社名</td><td style="padding:8px 12px;border-bottom:1px solid #eee">${company || '（個人）'}</td></tr>
    <tr><td style="padding:8px 12px;background:#f5f5f5;font-weight:bold">氏名</td><td style="padding:8px 12px;border-bottom:1px solid #eee">${name}</td></tr>
    <tr><td style="padding:8px 12px;background:#f5f5f5;font-weight:bold">メール</td><td style="padding:8px 12px;border-bottom:1px solid #eee">${email}</td></tr>
    <tr><td style="padding:8px 12px;background:#f5f5f5;font-weight:bold">申請日</td><td style="padding:8px 12px">${date}</td></tr>
  </table>
  <p style="margin-top:16px;font-size:13px;color:#888">修了証PDFは自動でクライアントに送信済みです。</p>
</div>`;
}

// ===== スプレッドシート記録 =====
function logToSheet(company, name, email, date) {
  try {
    const ss = SpreadsheetApp.openById('YOUR_SPREADSHEET_ID'); // ← スプレッドシートIDに変更
    const sheet = ss.getSheetByName('修了証申請') || ss.insertSheet('修了証申請');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['申請日', '会社名', '氏名', 'メールアドレス']);
    }
    sheet.appendRow([date, company, name, email]);
  } catch(e) {
    // スプレッドシートIDが未設定でもメール送信は継続
  }
}
