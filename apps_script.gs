// ============================================================
// Google Apps Script — ChatGPT 问卷数据接收后端
// 使用方法：
// 1. Google Sheet → Extensions → Apps Script
// 2. 粘贴全部代码，保存
// 3. Deploy → New deployment → Web app
// 4. Execute as: Me , Who has access: Anyone
// 5. 点 Deploy，复制 URL 发给我
// ============================================================

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const type = data.type;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = type === 'teacher' ? '教师问卷' : '学生问卷';
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'Sheet not found: ' + sheetName }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (type === 'teacher') {
      sheet.appendRow([
        data.timestamp || new Date().toISOString(),
        data.age || '',
        data.teachYears || '',
        (data.targetStudents || []).join('; '),
        (data.studentLevel || []).join('; '),
        data.latinExp || '',
        (data.teachLang || []).join('; '),
        data.schoolType || '',
        data.useAI || '',
        JSON.stringify(data.likert || {})
      ]);
    } else {
      sheet.appendRow([
        data.timestamp || new Date().toISOString(),
        data.age || '',
        data.gender || '',
        data.studyDuration || '',
        data.hskLevel || '',
        (data.studyPurpose || []).join('; '),
        data.studyPurposeOther || '',
        data.toolFreq || '',
        (data.toolsUsed || []).join('; '),
        data.toolsUsedOther || '',
        JSON.stringify(data.likert || {})
      ]);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('Survey data receiver is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}
