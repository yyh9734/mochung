/**
 * 애프터파티 참석 여부 수집용 Google Apps Script
 * Google Sheet 연동 코드
 */
const SHEET_NAME = 'afterparty_rsvp';

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const params = (e && e.parameter) || {};
  let postData = {};
  
  if (e && e.postData && e.postData.contents) {
    try {
      postData = JSON.parse(e.postData.contents);
    } catch (err) {
      // ignore JSON parse error
    }
  }
  
  const data = Object.assign({}, params, postData);
  const action = data.action || 'submit';
  
  try {
    if (action === 'list') {
      return jsonResponse({ ok: true, data: listRSVPs() }, data.callback);
    } else {
      const entry = addRSVP(data);
      return jsonResponse({ ok: true, entry: entry }, data.callback);
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message }, data.callback);
  }
}

function addRSVP(data) {
  const name = cleanText(data.name, 30);
  const relation = cleanText(data.relation || '신랑측', 20);
  const attendance = cleanText(data.attendance || '참석', 20);
  const guestCount = cleanText(data.guestCount || '1명', 20);
  const phone = cleanText(data.phone || '', 20);
  const note = cleanText(data.note || '', 200);
  const createdAt = new Date().toISOString();

  if (!name) throw new Error('성함을 입력해주세요.');

  const sheet = getRSVPSheet();
  sheet.appendRow([
    createdAt,
    name,
    relation,
    attendance,
    guestCount,
    phone,
    note
  ]);

  return {
    createdAt: createdAt,
    name: name,
    relation: relation,
    attendance: attendance,
    guestCount: guestCount,
    phone: phone,
    note: note
  };
}

function listRSVPs() {
  const sheet = getRSVPSheet();
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  
  return values.slice(1).map(row => ({
    createdAt: row[0],
    name: row[1],
    relation: row[2],
    attendance: row[3],
    guestCount: row[4],
    phone: row[5],
    note: row[6]
  }));
}

function getRSVPSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['등록일시', '성함', '구분', '참석여부', '동반인원', '연락처', '메모']);
  }

  return sheet;
}

function jsonResponse(data, callback) {
  if (callback) {
    const sanitizedCallback = String(callback).replace(/[^a-zA-Z0-9_$]/g, '');
    const body = `${sanitizedCallback}(${JSON.stringify(data)});`;
    return ContentService.createTextOutput(body)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function cleanText(value, maxLength) {
  return String(value || '')
    .replace(/[<>]/g, '')
    .trim()
    .slice(0, maxLength);
}
