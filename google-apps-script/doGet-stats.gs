/**
 * Stroke Quiz — Google Apps Script (วางทั้งไฟล์นี้ในโปรเจกต์เดียว)
 *
 * doPost: บันทึกคะแนน หัวตาราง
 *   ['วันที่/เวลา', 'ชื่อ', 'Ward', 'คะแนน', 'จาก 10', 'เปอร์เซ็นต์']
 * doGet: health + ?action=stats (รองรับ JSONP สำหรับหน้าเว็บ)
 *
 * ต้องมีสองบรรทัดนี้เสมอ — ถ้าลืม STATS_TOKEN จะ error ตอน doGet
 */
var SPREADSHEET_ID = '';
var STATS_TOKEN = 'PRH2024STROKE';

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);

  if (!data.token || data.token !== STATS_TOKEN) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'unauthorized' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['วันที่/เวลา', 'ชื่อ', 'Ward', 'คะแนน', 'จาก 10', 'เปอร์เซ็นต์']);
  }

  sheet.appendRow([
    new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' }),
    data.name,
    data.ward,
    data.score,
    data.total,
    Math.round((data.score / data.total) * 100) + '%'
  ]);

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** แถวข้อมูล 6 คอลัมน์: เวลา, ชื่อ, Ward, คะแนน, เต็ม, % */function columnsNoToken_() {
  return { time: 0, token: -1, name: 1, ward: 2, score: 3, total: 4, pct: 5 };
}

/** แถวข้อมูล 7 คอลัมน์: เวลา, token, ชื่อ, Ward, คะแนน, เต็ม, % */
function columnsWithToken_() {
  return { time: 0, token: 1, name: 2, ward: 3, score: 4, total: 5, pct: 6 };
}

/**
 * เรียกแค่ URL /exec → ได้ { ok, msg }
 * เรียก ?action=stats&token=...&callback=... → สถิติ (JSONP)
 */
function doGet(e) {
  e = e || { parameter: {} };
  var action = String(e.parameter.action || '').trim();

  if (!action) {
    return jsonpOut_({ ok: true, msg: 'Stroke Quiz API' }, e);
  }

  if (action !== 'stats') {
    return jsonpOut_({ ok: false, error: 'bad_action' }, e);
  }

  if (String(e.parameter.token || '') !== STATS_TOKEN) {
    return jsonpOut_({ ok: false, error: 'unauthorized' }, e);
  }

  try {
    var payload = buildStatsPayload_();
    return jsonpOut_(payload, e);
  } catch (err) {
    return jsonpOut_({ ok: false, error: String(err && err.message ? err.message : err) }, e);
  }
}

function jsonpOut_(obj, e) {
  var json = JSON.stringify(obj);
  var cb = String(e.parameter.callback || '');
  if (cb && /^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function getDataSheet_() {
  if (SPREADSHEET_ID) {
    return SpreadsheetApp.openById(SPREADSHEET_ID).getSheets()[0];
  }
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function parsePct_(raw, score, total) {
  var n = Number(raw);
  if (!isNaN(n)) return n;
  var s = String(raw != null ? raw : '')
    .replace(/%/g, '')
    .replace(/,/g, '.')
    .trim();
  n = Number(s);
  if (!isNaN(n)) return n;
  return Math.round((score / total) * 100);
}

function inferColumnLayout_(firstRow) {
  if (!firstRow || firstRow.length < 4) return columnsNoToken_();
  var t1 = String(firstRow[1] != null ? firstRow[1] : '').trim();
  if (t1 === STATS_TOKEN) return columnsWithToken_();
  return columnsNoToken_();
}

function buildStatsPayload_() {
  var sh = getDataSheet_();
  var values = sh.getDataRange().getValues();
  if (!values || values.length === 0) {
    return { ok: true, total: 0, topTwo: [] };
  }

  var col;
  var startRow = 1;
  var header = values[0];
  if (header && looksLikeHeaderRow_(header)) {
    col = mapHeaderToColumns_(header);
  } else {
    startRow = 0;
    col = inferColumnLayout_(values[0]);
  }

  var records = [];
  for (var r = startRow; r < values.length; r++) {
    var row = values[r];
    if (!row || row.length === 0) continue;

    if (col.token >= 0) {
      var tok = String(row[col.token] != null ? row[col.token] : '').trim();
      if (tok && tok !== STATS_TOKEN) continue;
    }

    var name = String(row[col.name] != null ? row[col.name] : '').trim();
    var score = Number(row[col.score]);
    if (!name || isNaN(score)) continue;

    var total = Number(row[col.total]);
    if (isNaN(total)) total = 10;
    var pct = parsePct_(row[col.pct], score, total);

    records.push({
      name: name,
      ward: String(row[col.ward] != null ? row[col.ward] : '').trim(),
      score: score,
      total: total,
      pct: pct,
      rowIndex: r
    });
  }

  records.sort(function (a, b) {
    if (b.score !== a.score) return b.score - a.score;
    return b.rowIndex - a.rowIndex;
  });

  var topTwo = records.slice(0, 2).map(function (x) {
    return { name: x.name, ward: x.ward, score: x.score, total: x.total, pct: x.pct };
  });

  return { ok: true, total: records.length, topTwo: topTwo };
}

function looksLikeHeaderRow_(row) {
  for (var c = 0; c < row.length; c++) {
    var t = String(row[c] || '').toLowerCase();
    if (t.indexOf('name') >= 0 || t.indexOf('ชื่อ') >= 0) return true;
    if (t.indexOf('score') >= 0 || t.indexOf('คะแนน') >= 0) return true;
    if (t.indexOf('ward') >= 0) return true;
    if (t.indexOf('จาก') >= 0) return true;
    if (t.indexOf('เปอร์เซ็น') >= 0) return true;
  }
  var scoreCell = Number(row[4]);
  return isNaN(scoreCell);
}

function mapHeaderToColumns_(header) {
  var col = columnsNoToken_();
  var hasTokenCol = false;
  for (var c = 0; c < header.length; c++) {
    var h = String(header[c] || '').toLowerCase().replace(/\s/g, '');
    if (h.indexOf('token') >= 0) {
      col.token = c;
      hasTokenCol = true;
    } else if (h.indexOf('name') >= 0 || h.indexOf('ชื่อ') >= 0) col.name = c;
    else if (h.indexOf('ward') >= 0 || h.indexOf('หน่วย') >= 0) col.ward = c;
    else if (h.indexOf('score') >= 0 || h.indexOf('คะแนน') >= 0) col.score = c;
    else if (h.indexOf('total') >= 0 || h.indexOf('จาก') >= 0) col.total = c;
    else if (h.indexOf('pct') >= 0 || h.indexOf('percent') >= 0 || h.indexOf('%') >= 0 || h.indexOf('เปอร์เซ็น') >= 0)
      col.pct = c;
    else if (h.indexOf('date') >= 0 || h.indexOf('time') >= 0 || h.indexOf('วันที่') >= 0 || h.indexOf('timestamp') >= 0)
      col.time = c;
  }
  if (!hasTokenCol) col.token = -1;
  return col;
}
