const SHEET_NAME = 'Students';

function doGet(e) {
  const p = e && e.parameter ? e.parameter : {};
  const callback = p.callback || '';
  let result;
  try {
    result = handleRequest_(p);
  } catch (err) {
    result = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  const json = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let p = {};
  try {
    p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (_) {}
  const result = handleRequest_(p);
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  const headers = ['id','grade','studentNumber','phone','prefix','firstname','lastname','status','reason','timestamp','updatedAt'];
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function handleRequest_(p) {
  const sh = setupDatabase();
  const action = String(p.action || 'list').toLowerCase();

  if (action === 'list') return { ok: true, data: readAll_(sh) };

  if (action === 'save') {
    const id = Utilities.getUuid();
    const now = new Date().toISOString();
    sh.appendRow([
      id, p.grade || '', p.studentNumber || '', p.phone || '', p.prefix || '',
      p.firstname || '', p.lastname || '', p.status || 'ไป',
      p.reason || '-', p.timestamp || now, now
    ]);
    return { ok: true, id: id };
  }

  if (action === 'update') {
    const id = String(p.id || '');
    const row = findRow_(sh, id);
    if (!row) return { ok: false, error: 'ไม่พบข้อมูล' };
    const now = new Date().toISOString();
    sh.getRange(row, 1, 1, 11).setValues([[
      id, p.grade || '', p.studentNumber || '', p.phone || '', p.prefix || '',
      p.firstname || '', p.lastname || '', p.status || 'ไป',
      p.reason || '-', p.timestamp || now, now
    ]]);
    return { ok: true, id: id };
  }

  if (action === 'delete') {
    const row = findRow_(sh, String(p.id || ''));
    if (!row) return { ok: false, error: 'ไม่พบข้อมูล' };
    sh.deleteRow(row);
    return { ok: true };
  }

  return { ok: false, error: 'Unknown action: ' + action };
}

function readAll_(sh) {
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];
  const h = values[0];
  return values.slice(1).filter(r => r[0]).map(r => {
    const o = {};
    h.forEach((k, i) => o[k] = r[i]);
    return o;
  });
}

function findRow_(sh, id) {
  if (!id) return null;
  const values = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 0), 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === id) return i + 2;
  }
  return null;
}
