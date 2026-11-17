// SMEAG IELTS Mock Report – Google Sheets backend
// Setup: Project Settings → Script properties → add PASSCODE = your staff passcode
const SHEET = 'Reports';
const COLS = ['rid','saved','type','centre','date','sid','family','first','english','dob','sex','school',
  'origin','nation','lang','L','R','W','S','overall','cefr','comments','sdate','trf','photo'];

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(SHEET);
  if (!s) {
    s = ss.insertSheet(SHEET);
    s.getRange(1, 1, s.getMaxRows(), COLS.length).setNumberFormat('@'); // keep dates/IDs as plain text
    s.appendRow(COLS);
    s.setFrozenRows(1);
  }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() { return out_({ ok: true, msg: 'SMEAG IELTS report API is running' }); }

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    const pass = PropertiesService.getScriptProperties().getProperty('PASSCODE');
    if (pass && b.pass !== pass) return out_({ ok: false, error: 'Wrong passcode' });
    const s = getSheet_();
    const v = s.getDataRange().getValues();
    if (b.action === 'list') {
      const h = v.shift();
      return out_({ ok: true, records: v.filter(r => r[0] !== '').map(r => Object.fromEntries(h.map((k, i) => [k, String(r[i])]))) });
    }
    const find = rid => { for (let i = 1; i < v.length; i++) if (String(v[i][0]) === rid) return i + 1; return -1; };
    if (b.action === 'save') {
      const r = b.record, row = find(r.rid), line = COLS.map(k => r[k] == null ? '' : String(r[k]));
      if (row > 0) s.getRange(row, 1, 1, COLS.length).setValues([line]); else s.appendRow(line);
      return out_({ ok: true });
    }
    if (b.action === 'delete') {
      const row = find(b.rid); if (row > 0) s.deleteRow(row);
      return out_({ ok: true });
    }
    return out_({ ok: false, error: 'Unknown action' });
  } catch (err) { return out_({ ok: false, error: String(err) }); }
}
