/**
 * 애프터파티 참석 여부 수집용 Google Apps Script
 * Google Sheet 연동 코드
 */

// 만약 독립 실행형 스크립트로 만드셨다면 아래 SPREADSHEET_ID에 구글 시트 주소의 ID를 넣어주세요.
// 예: https://docs.google.com/spreadsheets/d/1ABC123.../edit -> '1ABC123...'
const SPREADSHEET_ID = ''; 
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
  const createdAt = new Date().toISOString();

  if (!name) throw new Error('성함을 입력해주세요.');

  const sheet = getRSVPSheet();
  sheet.appendRow([
    createdAt,
    name,
    relation,
    attendance
  ]);

  return {
    createdAt: createdAt,
    name: name,
    relation: relation,
    attendance: attendance
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
    attendance: row[3]
  }));
}

function getRSVPSheet() {
  let spreadsheet = null;

  // 1. SPREADSHEET_ID가 지정된 경우
  if (SPREADSHEET_ID && SPREADSHEET_ID.trim()) {
    try {
      spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID.trim());
    } catch (e) {}
  }

  // 2. 바인딩된 시트가 있는 경우
  if (!spreadsheet) {
    try {
      spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    } catch (e) {}
  }

  // 3. 드라이브에서 시트 자동 탐색
  if (!spreadsheet) {
    try {
      const files = DriveApp.searchFiles('mimeType = "application/vnd.google-apps.spreadsheet"');
      if (files.hasNext()) {
        spreadsheet = SpreadsheetApp.open(files.next());
      }
    } catch (e) {}
  }

  if (!spreadsheet) {
    throw new Error('스프레드시트를 찾을 수 없습니다.');
  }

  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) {
    // 첫 번째 시트 탭 활용 또는 신규 생성
    const sheets = spreadsheet.getSheets();
    if (sheets.length > 0 && sheets[0].getLastRow() === 0) {
      sheet = sheets[0];
      sheet.setName(SHEET_NAME);
    } else {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
    }
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['등록일시', '성함', '구분', '참석여부']);
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
