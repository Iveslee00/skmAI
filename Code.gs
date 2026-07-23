/**
 * SKM AI 調查問卷接收端
 * 綁定試算表：190udwPIiTMr-Em-P5y9CiHSdFyWy6V4Rc2eNNkH6eec
 */
const SPREADSHEET_ID = '190udwPIiTMr-Em-P5y9CiHSdFyWy6V4Rc2eNNkH6eec';
const SHEET_NAME = 'AI 問卷回覆';
const HEADERS = [
  '送出時間', '填表人姓名', '所屬部門/館別', '職稱', '使用工具', '其他工具',
  '資料敏感度', 'AI 使用普及程度', '已見效益', '其他效益', '安控疑慮', '其他疑慮',
  '第一優先痛點', '第二優先痛點', '第三優先痛點', '期待支援', 'AI 種子成員推薦', '工具使用情境'
];

function doGet() {
  return json_({ ok: true, service: 'SKM AI Survey receiver' });
}

function setupSheet() {
  const sheet = getSheet_();
  if (sheet.getLastRow() === 1) {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#7C1F35')
      .setFontColor('#FFFFFF');
    sheet.autoResizeColumns(1, HEADERS.length);
  }
  return `已完成「${SHEET_NAME}」欄位初始化`;
}

function doPost(e) {
  try {
    const raw = e && e.parameter && e.parameter.payload;
    if (!raw) throw new Error('Missing payload');
    const data = JSON.parse(raw);
    validatePayload_(data);
    const lock = LockService.getScriptLock();
    lock.waitLock(10 * 1000);
    try {
      const cache = CacheService.getScriptCache();
      const cacheKey = `submission:${data.submission_id}`;
      if (cache.get(cacheKey)) return json_({ ok: true, duplicate: true });
      const sheet = getSheet_();
      sheet.appendRow([
        new Date(),
        data.name, data.dept, data.title,
        list_(data.tools), data.tools_other || '',
        list_(data.sensitivity), data.maturity,
        list_(data.benefits), data.benefits_other || '',
        list_(data.concerns), data.concerns_other || '',
        data.priority1, data.priority2 || '', data.priority3 || '',
        list_(data.support), data.ambassador || '', list_(data.tools_usage)
      ]);
      cache.put(cacheKey, '1', 21600);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  } catch (error) {
    console.error(error);
    return json_({ ok: false, error: error.message });
  }
}

function getSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
  const existingHeaders = sheet.getLastRow() === 0 ? [] : sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), HEADERS.length)).getValues()[0];
  if (HEADERS.some((header, index) => existingHeaders[index] !== header)) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sheet;
}

function validatePayload_(data) {
  const requiredText = ['submission_id', 'name', 'dept', 'title', 'maturity', 'priority1'];
  requiredText.forEach(key => {
    if (!String(data[key] || '').trim()) throw new Error(`Missing required field: ${key}`);
  });
  ['tools', 'tools_usage', 'sensitivity', 'benefits', 'concerns', 'support'].forEach(key => {
    if (!Array.isArray(data[key]) || data[key].filter(Boolean).length === 0) {
      throw new Error(`Missing required selection: ${key}`);
    }
  });
}

function list_(value) {
  return Array.isArray(value) ? value.filter(Boolean).join('、') : '';
}

function json_(body) {
  return ContentService
    .createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}
