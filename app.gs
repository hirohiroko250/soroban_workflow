const CONFIG = {
  spreadsheetId: '1HZ2OnQ477NQXjT-xf5prkyKqLIe3k1UPE3N6JqI7kLI',
  sheetName: '出退勤ログ',
  timezone: 'Asia/Tokyo'
};

function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  template.campuses = fetchCampusesFromSheet_();
  return template
    .evaluate()
    .setTitle('出退勤ログ');
}

function recordAttendance(teacherId, action, campusId) {
  if (!teacherId) {
    throw new Error('先生IDを入力してください');
  }
  if (['start', 'end'].indexOf(action) === -1) {
    throw new Error('不正な操作が指定されました');
  }
  const campus = getCampus_(campusId);

  const sheet = getSheet_();
  const now = new Date();
  const dateStr = Utilities.formatDate(now, CONFIG.timezone, 'yyyy/MM/dd');
  const timeStr = Utilities.formatDate(now, CONFIG.timezone, 'HH:mm:ss');

  const rowIndex = findRowIndex_(sheet, dateStr, teacherId);
  let targetRow = rowIndex;

  if (rowIndex === -1) {
    const newRow = buildEmptyRow_(dateStr, teacherId, campus);
    sheet.appendRow(newRow);
    targetRow = sheet.getLastRow();
  }

  const column = action === 'start' ? 3 : 4;
  const currentValue = sheet.getRange(targetRow, column).getValue();
  if (currentValue) {
    const label = action === 'start' ? '出勤' : '退勤';
    throw new Error(`${dateStr}の${label}は登録済みです`);
  }

  sheet.getRange(targetRow, column).setValue(timeStr);
  sheet.getRange(targetRow, 5).setValue(campus.id);
  sheet.getRange(targetRow, 6).setValue(campus.name);
  sheet.getRange(targetRow, 7).setValue('手動入力');
  sheet.getRange(targetRow, 10).setValue('manual');

  return { date: dateStr, time: timeStr, action, campus };
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  let sheet = spreadsheet.getSheetByName(CONFIG.sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.sheetName);
  }
  ensureHeaderRow_(sheet);
  return sheet;
}

function ensureHeaderRow_(sheet) {
  const expected = ['日付', '先生ID', '出勤時間', '退勤時間', '校舎ID', '校舎名', '区分', '先生名', '生徒数', 'ソース'];
  const firstRow = sheet.getRange(1, 1, 1, expected.length).getValues()[0];
  const hasHeader = expected.every((value, index) => firstRow[index] === value);
  if (!hasHeader) {
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
  }
}

function getCampus_(campusId) {
  if (!campusId) {
    throw new Error('校舎を選択してください');
  }
  const campus = fetchCampusesFromSheet_().find((c) => c.id === campusId);
  if (!campus) {
    throw new Error('無効な校舎が指定されました');
  }
  return campus;
}

function fetchCampusesFromSheet_() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const sheet = spreadsheet.getSheetByName('校舎マスタ');
  if (!sheet) {
    throw new Error('校舎マスタ シートが見つかりません');
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  return values
    .filter(([id, name]) => id && name)
    .map(([id, name]) => ({ id: String(id), name: String(name) }));
}

function buildEmptyRow_(dateStr, teacherId, campus) {
  return [
    dateStr,
    teacherId,
    '',
    '',
    campus.id,
    campus.name,
    '手動入力',
    '',
    '',
    'manual'
  ];
}

function doPost(e) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GAS_API_KEY');
  let payload = {};
  try {
    payload = JSON.parse(e.postData && e.postData.contents ? e.postData.contents : '{}');
  } catch (err) {
    return ContentService.createTextOutput('Invalid JSON').setMimeType(ContentService.MimeType.TEXT).setResponseCode(400);
  }
  if (!payload.apiKey || payload.apiKey !== apiKey) {
    return ContentService.createTextOutput('Forbidden').setMimeType(ContentService.MimeType.TEXT).setResponseCode(403);
  }

  const items = Array.isArray(payload.rows) ? payload.rows : [];
  if (!items.length) {
    return ContentService.createTextOutput('No rows');
  }

  const sheet = getSheet_();
  const appendValues = [];
  items.forEach((item) => {
    const dateStr = item.date;
    const teacherId = item.teacher_id || item.teacherId;
    const startTime = normalizeTime_(item.start_time || item.startTime);
    if (!dateStr || !teacherId || !startTime) {
      return;
    }
    if (findExistingSlot_(sheet, dateStr, teacherId, startTime)) {
      return;
    }

    const endTime = normalizeTime_(item.end_time || item.endTime) || calculateEndTime_(startTime);
    const campusId = item.school_id || item.schoolId || '';
    const campusName = item.school_name || item.schoolName || '';
    const teacherName = item.teacher_name || item.teacherName || '';
    const studentCount = Number(item.attendance_count || item.student_count || 0) || 0;
    const workType = resolveWorkType_(item.work_type || item.workType);
    appendValues.push([
      dateStr,
      teacherId,
      startTime,
      endTime,
      campusId,
      campusName,
      workType,
      teacherName,
      studentCount,
      'oza_scraper'
    ]);
  });

  if (!appendValues.length) {
    return ContentService.createTextOutput('No appendable rows');
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, appendValues.length, appendValues[0].length).setValues(appendValues);
  return ContentService.createTextOutput(`ok:${appendValues.length}`);
}

function normalizeTime_(value) {
  if (!value) {
    return '';
  }
  const text = String(value).trim();
  const match = text.match(/(\d{1,2}:\d{2})/);
  if (!match) {
    return text;
  }
  const [hour, minute] = match[1].split(':').map(Number);
  return Utilities.formatString('%02d:%02d', hour, minute);
}

function calculateEndTime_(startTime) {
  if (!startTime) {
    return '';
  }
  const [hour, minute] = startTime.split(':').map(Number);
  const total = hour * 60 + minute + 50;
  const endHour = Math.floor(total / 60);
  const endMinute = total % 60;
  return Utilities.formatString('%02d:%02d', endHour, endMinute);
}

function findExistingSlot_(sheet, dateStr, teacherId, startTime) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return false;
  }
  const targetDate = normalizeDateString_(dateStr);
  const targetTeacher = String(teacherId || '').trim();
  const range = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  return range.some((row) => {
    const rowDate = normalizeDateString_(row[0]);
    const rowTeacher = String(row[1] || '').trim();
    const rowStart = normalizeTime_(row[2]);
    return rowDate === targetDate && rowTeacher === targetTeacher && rowStart === startTime;
  });
}

function resolveWorkType_(value) {
  if (value === undefined || value === null) {
    return '授業';
  }
  const text = String(value).trim();
  if (!text) {
    return '授業';
  }
  return text;
}

function normalizeDateString_(value) {
  if (!value) {
    return '';
  }
  if (value instanceof Date) {
    return Utilities.formatDate(value, CONFIG.timezone, 'yyyy-MM-dd');
  }
  return String(value).trim().replace(/\//g, '-');
}

function findRowIndex_(sheet, dateStr, teacherId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return -1;
  }
  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    const [date, id] = data[i];
    const dateValue = date instanceof Date
      ? Utilities.formatDate(date, CONFIG.timezone, 'yyyy/MM/dd')
      : String(date || '');
    const idValue = String(id || '').trim();
    if (dateValue === dateStr && idValue === String(teacherId).trim()) {
      return i + 2;
    }
  }
  return -1;
}
