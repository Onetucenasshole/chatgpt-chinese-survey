// ============================================================
// Google Apps Script — ChatGPT 问卷数据接收后端
// 使用方法：
// 1. 在 Google Sheet 中点击 Extensions → Apps Script
// 2. 把下面全部代码粘贴进去
// 3. 点击 Deploy → New deployment → 类型选 Web app
// 4. Execute as: Me, Who has access: Anyone
// 5. 点 Deploy，复制 URL
// 6. 把 URL 贴到问卷 HTML 的 GOOGLE_SCRIPT_URL 占位符处
// ============================================================

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const type = data.type; // "teacher" or "student"

    // 获取对应的 sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = type === 'teacher' ? '教师问卷' : '学生问卷';
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ error: `Sheet "${sheetName}" not found` }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (type === 'teacher') {
      appendTeacherRow(sheet, data);
    } else {
      appendStudentRow(sheet, data);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function appendTeacherRow(sheet, data) {
  const row = [
    data.timestamp || new Date().toISOString(),
    data.age || '',
    data.teachYears || '',
    (data.targetStudents || []).join('; '),
    (data.studentLevel || []).join('; '),
    data.latinExp || '',
    (data.teachLang || []).join('; '),
    data.schoolType || '',
    data.useAI || ''
  ];

  // 46 Likert questions
  for (let i = 1; i <= 46; i++) {
    const key = 'q' + i;
    row.push(data.likert && data.likert[key] !== undefined ? data.likert[key] : '');
  }

  sheet.appendRow(row);
}

function appendStudentRow(sheet, data) {
  const row = [
    data.timestamp || new Date().toISOString(),
    data.age || '',
    data.gender || '',
    data.studyDuration || '',
    data.hskLevel || '',
    (data.studyPurpose || []).join('; '),
    data.studyPurposeOther || '',
    data.toolFreq || '',
    (data.toolsUsed || []).join('; '),
    data.toolsUsedOther || ''
  ];

  // 63 Likert questions
  for (let i = 1; i <= 63; i++) {
    const key = 'q' + i;
    row.push(data.likert && data.likert[key] !== undefined ? data.likert[key] : '');
  }

  sheet.appendRow(row);
}

// 可选：处理 CORS 预检请求
function doGet() {
  return ContentService.createTextOutput('Survey data receiver is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}
