// ═══════════════════════════════════════════════════════
// CREAL Contact Form → Google Sheets + Email Notification
// ═══════════════════════════════════════════════════════
//
// SETUP:
// 1. Google Sheets에서 새 스프레드시트 생성
// 2. 첫 번째 행에 헤더 입력: Timestamp | Name | Company | Email | Message
// 3. Extensions > Apps Script 클릭
// 4. 이 코드 전체를 붙여넣기
// 5. NOTIFICATION_EMAIL을 본인 이메일로 변경
// 6. Deploy > New deployment > Web app 선택
//    - Execute as: Me
//    - Who has access: Anyone
// 7. 배포 URL을 복사해서 index.html의 APPS_SCRIPT_URL에 붙여넣기
//
// ═══════════════════════════════════════════════════════

const NOTIFICATION_EMAIL = "hello@crealwork.com";
const SHEET_NAME = "Sheet1"; // ← 시트 탭 이름

function doPost(e) {
  try {
    let data;
    if (e.parameter && e.parameter.payload) {
      data = JSON.parse(e.parameter.payload);
    } else {
      data = JSON.parse(e.postData.contents);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // 시트가 없으면 자동 생성 + 헤더 추가
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(["Timestamp", "Name", "Company", "Email", "Message"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    }

    const timestamp = new Date().toLocaleString("en-US", { timeZone: "America/Vancouver" });
    const row = [timestamp, data.name, data.company, data.email, data.message];

    sheet.appendRow(row);

    // 이메일 알림 발송
    sendNotification(data, timestamp);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function sendNotification(data, timestamp) {
  const subject = `New inquiry from ${data.name || "Unknown"}`;

  const body = [
    `New contact form submission on crealwork.com`,
    ``,
    `Name: ${data.name || "-"}`,
    `Company: ${data.company || "-"}`,
    `Email: ${data.email || "-"}`,
    `Message: ${data.message || "-"}`,
    ``,
    `Submitted: ${timestamp}`,
  ].join("\n");

  const htmlBody = `
    <div style="font-family: -apple-system, sans-serif; max-width: 520px; padding: 24px;">
      <div style="border-bottom: 3px solid #FFCC4E; padding-bottom: 16px; margin-bottom: 24px;">
        <strong style="font-size: 18px; color: #1A1816;">New Inquiry</strong>
        <span style="color: #6E6660; font-size: 14px; margin-left: 8px;">crealwork.com</span>
      </div>
      <table style="width: 100%; border-collapse: collapse;">
        <tr><td style="padding: 8px 0; color: #A39D96; width: 90px; vertical-align: top;">Name</td><td style="padding: 8px 0; color: #1A1816; font-weight: 600;">${data.name || "-"}</td></tr>
        <tr><td style="padding: 8px 0; color: #A39D96; vertical-align: top;">Company</td><td style="padding: 8px 0; color: #1A1816;">${data.company || "-"}</td></tr>
        <tr><td style="padding: 8px 0; color: #A39D96; vertical-align: top;">Email</td><td style="padding: 8px 0;"><a href="mailto:${data.email}" style="color: #1A1816; font-weight: 600;">${data.email || "-"}</a></td></tr>
        <tr><td style="padding: 8px 0; color: #A39D96; vertical-align: top;">Message</td><td style="padding: 8px 0; color: #1A1816;">${data.message || "-"}</td></tr>
      </table>
      <div style="margin-top: 24px; padding-top: 16px; border-top: 1px solid #ECE7DE; color: #A39D96; font-size: 12px;">
        ${timestamp}
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: NOTIFICATION_EMAIL,
    subject: subject,
    body: body,
    htmlBody: htmlBody,
  });
}

// CORS preflight 대응
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "CREAL form endpoint active" }))
    .setMimeType(ContentService.MimeType.JSON);
}
