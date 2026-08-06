/**
 * 그림책 스토리 - 공유 저장소 (Google Apps Script)
 *
 * 스프레드시트 한 장을 서버로 써서, A패드에서 쓰던 원고를 B패드에서 이어서
 * 보고 고칠 수 있게 합니다. 이미지는 구글 드라이브 폴더에 올라가고 시트에는
 * 파일 ID만 저장됩니다.
 *
 * 배포 방법은 같은 폴더의 README.md 참고.
 */

/** 비워두면 이 스크립트에 연결된 스프레드시트를 씁니다. */
var SHEET_ID = '';
var SHEET_NAME = 'Stories';
var IMAGE_FOLDER_NAME = '그림책 이미지';

/** 시트 셀 1칸에 5만자 제한이 있어서 나눠 담습니다. */
var DATA_COL_START = 6; // F열
var DATA_COL_COUNT = 8;
var CHUNK = 45000;

var HEADERS = ['id', 'title', 'updatedAt', 'version', 'editor',
               'data1', 'data2', 'data3', 'data4', 'data5', 'data6', 'data7', 'data8'];

/** 코드에 헷갈리는 글자(0/O, 1/I)는 뺐습니다. */
var ALPHABET = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

/* ────────────────────────── 진입점 ────────────────────────── */

function doGet(e) {
  return handle((e && e.parameter) || {});
}

function doPost(e) {
  var params = {};
  try {
    params = JSON.parse(e.postData.contents);
  } catch (err) {
    params = (e && e.parameter) || {};
  }
  return handle(params);
}

function handle(p) {
  try {
    switch (String(p.action || '')) {
      case 'ping':    return json({ ok: true, service: '그림책 스토리' });
      case 'create':  return actCreate(p);
      case 'load':    return actLoad(p);
      case 'save':    return actSave(p);
      case 'poll':    return actPoll(p);
      case 'upload':  return actUpload(p);
      case 'imgdata': return actImgData(p);
      default:        return json({ ok: false, error: '알 수 없는 요청입니다: ' + p.action });
    }
  } catch (err) {
    return json({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ────────────────────────── 동작 ────────────────────────── */

function actCreate(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getSheet();
    var id = uniqueId(sheet);
    var now = new Date().toISOString();
    var row = [id, String(p.title || '새 그림책'), now, 1, String(p.editor || '')];
    while (row.length < HEADERS.length) row.push('');
    sheet.appendRow(row);
    writeData(sheet, sheet.getLastRow(), String(p.data || '{"pages":[]}'));
    return json({ ok: true, id: id, version: 1, updatedAt: now });
  } finally {
    lock.releaseLock();
  }
}

function actLoad(p) {
  var found = findRow(getSheet(), p.id);
  if (!found) return json({ ok: false, notFound: true, error: '그 코드로 저장된 그림책이 없어요.' });
  return json({ ok: true, story: found.story });
}

function actSave(p) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getSheet();
    var found = findRow(sheet, p.id);
    if (!found) return json({ ok: false, notFound: true, error: '그 코드로 저장된 그림책이 없어요.' });

    var base = Number(p.baseVersion || 0);
    var current = Number(found.story.version || 0);

    // 다른 기기가 먼저 저장한 경우 - 덮어쓰지 않고 상대 원고를 돌려줍니다.
    if (base !== current && !truthy(p.force)) {
      return json({ ok: false, conflict: true, story: found.story });
    }

    var now = new Date().toISOString();
    var next = current + 1;
    sheet.getRange(found.row, 2).setValue(String(p.title || found.story.title || ''));
    sheet.getRange(found.row, 3).setValue(now);
    sheet.getRange(found.row, 4).setValue(next);
    sheet.getRange(found.row, 5).setValue(String(p.editor || ''));
    writeData(sheet, found.row, String(p.data || '{"pages":[]}'));

    return json({ ok: true, version: next, updatedAt: now });
  } finally {
    lock.releaseLock();
  }
}

/** 가벼운 변경 확인용. 버전이 같으면 원고를 실어 보내지 않습니다. */
function actPoll(p) {
  var sheet = getSheet();
  var found = findRow(sheet, p.id, true);
  if (!found) return json({ ok: false, notFound: true });
  if (Number(found.version) === Number(p.version || 0)) {
    return json({ ok: true, changed: false, version: found.version });
  }
  var full = findRow(sheet, p.id);
  return json({ ok: true, changed: true, story: full.story });
}

function actUpload(p) {
  var bytes = Utilities.base64Decode(String(p.data || ''));
  var blob = Utilities.newBlob(bytes, String(p.mime || 'image/jpeg'),
                               String(p.name || 'picture.jpg'));
  var file = getImageFolder().createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    // 조직 정책으로 링크 공개가 막혀 있어도 imgdata 경로로 계속 볼 수 있습니다.
  }
  return json({
    ok: true,
    fileId: file.getId(),
    url: 'https://lh3.googleusercontent.com/d/' + file.getId()
  });
}

/** 드라이브 직링크가 막힌 환경을 위한 대체 경로. */
function actImgData(p) {
  var blob = DriveApp.getFileById(String(p.fileId)).getBlob();
  return json({
    ok: true,
    mime: blob.getContentType(),
    data: Utilities.base64Encode(blob.getBytes())
  });
}

/* ────────────────────────── 시트 도우미 ────────────────────────── */

function getSpreadsheet() {
  if (SHEET_ID) return SpreadsheetApp.openById(SHEET_ID);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('SHEET_ID를 채워 주세요. (스크립트가 스프레드시트에 연결돼 있지 않습니다)');
  return ss;
}

function getSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getImageFolder() {
  var it = DriveApp.getFoldersByName(IMAGE_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(IMAGE_FOLDER_NAME);
}

/**
 * id로 행을 찾습니다. metaOnly면 원고 본문은 읽지 않아 poll이 빠릅니다.
 */
function findRow(sheet, id, metaOnly) {
  id = String(id || '').toUpperCase().trim();
  if (!id) return null;
  var last = sheet.getLastRow();
  if (last < 2) return null;

  var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).toUpperCase().trim() !== id) continue;
    var row = i + 2;
    var meta = sheet.getRange(row, 1, 1, 5).getValues()[0];
    if (metaOnly) return { row: row, version: Number(meta[3] || 0) };
    return {
      row: row,
      story: {
        id: String(meta[0]),
        title: String(meta[1] || ''),
        updatedAt: String(meta[2] || ''),
        version: Number(meta[3] || 0),
        editor: String(meta[4] || ''),
        data: readData(sheet, row)
      }
    };
  }
  return null;
}

function readData(sheet, row) {
  var cells = sheet.getRange(row, DATA_COL_START, 1, DATA_COL_COUNT).getValues()[0];
  return cells.map(function (c) { return c === null || c === undefined ? '' : String(c); }).join('');
}

function writeData(sheet, row, text) {
  var chunks = [];
  for (var i = 0; i < text.length; i += CHUNK) chunks.push(text.substr(i, CHUNK));
  if (chunks.length > DATA_COL_COUNT) {
    throw new Error('원고가 너무 큽니다. 그림을 몇 장 덜어 주세요.');
  }
  var out = [];
  for (var c = 0; c < DATA_COL_COUNT; c++) out.push(chunks[c] || '');
  sheet.getRange(row, DATA_COL_START, 1, DATA_COL_COUNT).setValues([out]);
}

function uniqueId(sheet) {
  for (var tries = 0; tries < 40; tries++) {
    var id = '';
    for (var i = 0; i < 6; i++) {
      id += ALPHABET.charAt(Math.floor(Math.random() * ALPHABET.length));
    }
    if (!findRow(sheet, id, true)) return id;
  }
  throw new Error('코드를 만들지 못했습니다. 다시 시도해 주세요.');
}

function truthy(v) {
  return v === true || v === 'true' || v === 1 || v === '1';
}
