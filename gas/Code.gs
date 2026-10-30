/**
 * 運行記録アプリ（スマホの音声入力）の記録を、このスプレッドシートに貯める。
 * スプレッドシートの「拡張機能 → Apps Script」に貼り付け、ウェブアプリとして公開して使う。
 *
 * シート「記録」：出発・到着・終了を1回ずつ1行
 * シート「運行」：出発から終了までを1行（終了したときに足す）
 * シート「月別集計」：運行を月×区分で合計（式で自動）
 */
const LOG_SHEET = '記録';
const TRIP_SHEET = '運行';
const MONTH_SHEET = '月別集計';
const LOG_HEADERS = ['記録ID', '運行ID', '日付', '時刻', '種類', '出発地', '行き先・場所', '区分', '目的',
  'メーター(km)', '区間km', '運行の累計km', '発言'];
const TRIP_HEADERS = ['運行ID', '日付', '出発時刻', '終了時刻', '出発地', '行き先（経由）', '区分', '目的',
  '開始メーター', '終了メーター', '走行km'];
// アプリから直してよい項目
const EDITABLE = ['日付', '時刻', '出発地', '行き先・場所', '区分', '目的', 'メーター(km)'];

function sheetOf_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    // 日付・時刻は文字のまま持つ（自動で日付型に変わると時刻がずれることがある）
    sheet.getRange('B:D').setNumberFormat('@');
    if (name === TRIP_SHEET) monthSheet_(ss);
  }
  return sheet;
}

function monthSheet_(ss) {
  if (ss.getSheetByName(MONTH_SHEET)) return;
  const sheet = ss.insertSheet(MONTH_SHEET);
  sheet.getRange('A1').setFormula(
    `=QUERY({ARRAYFORMULA(LEFT(${TRIP_SHEET}!B2:B,7)),${TRIP_SHEET}!G2:G,${TRIP_SHEET}!K2:K},` +
    `"select Col1, sum(Col3) where Col3 is not null group by Col1 pivot Col2 label Col1 '月'",0)`);
  sheet.getRange('1:1').setFontWeight('bold');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// 同じIDの行がすでにあれば何行目かを返す（再送で二重に入らないように）
function findRow_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

const TEXT_COLS = /日付|時刻/;

// 日付・時刻は先頭に ' を付けて文字のまま書く（日付型になると時刻が19分ずれることがある）
function cell_(v, h) {
  return TEXT_COLS.test(h) && v !== '' && v != null ? "'" + v : v;
}

// 日付・時刻は表示されている文字で読む（4:53:00 → 04:53）
function text_(v, shown, h) {
  if (!TEXT_COLS.test(h)) return v;
  if (/時刻/.test(h)) {
    const m = String(shown).match(/^(\d{1,2}):(\d{2})/);
    return m ? m[1].padStart(2, '0') + ':' + m[2] : shown;
  }
  const m = String(shown).match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})/);
  return m ? `${m[1]}/${m[2].padStart(2, '0')}/${m[3].padStart(2, '0')}` : shown;
}

// 指定した行範囲を {見出し: 値} の配列で読む
function readRows_(sheet, headers, from, n) {
  const range = sheet.getRange(from, 1, n, headers.length);
  const values = range.getValues();
  const shown = range.getDisplayValues();
  return values.map((r, i) => {
    const o = {};
    headers.forEach((h, j) => { o[h] = text_(r[j], shown[i][j], h); });
    return o;
  });
}

// 運行IDの記録を全部読み、[{row, rec}] を日時の順に返す
function tripLogs_(tripId) {
  const sheet = sheetOf_(LOG_SHEET, LOG_HEADERS);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const out = [];
  readRows_(sheet, LOG_HEADERS, 2, last - 1).forEach((rec, i) => {
    if (String(rec['運行ID']) === String(tripId)) out.push({ row: i + 2, rec, order: i });
  });
  out.sort((a, b) => (String(a.rec['日付']) + a.rec['時刻']).localeCompare(String(b.rec['日付']) + b.rec['時刻']) || a.order - b.order);
  return out;
}

// 直したあと、区間km・累計km と「運行」の行を計算し直す
function recompute_(tripId) {
  const items = tripLogs_(tripId);
  if (!items.length) return;
  const sheet = sheetOf_(LOG_SHEET, LOG_HEADERS);
  const colLeg = LOG_HEADERS.indexOf('区間km') + 1;
  const start = Number(items[0].rec['メーター(km)']);
  let prev = start;
  items.forEach((it, i) => {
    const km = Number(it.rec['メーター(km)']);
    const leg = i === 0 || it.rec['種類'] === '出発' ? '' : km - prev;
    sheet.getRange(it.row, colLeg, 1, 2).setValues([[leg, km - start]]);
    prev = km;
  });

  const tripSheet = sheetOf_(TRIP_SHEET, TRIP_HEADERS);
  const row = findRow_(tripSheet, tripId);
  if (!row) return;  // まだ終了していない運行
  const first = items[0].rec;
  const lastRec = items[items.length - 1].rec;
  // 経由地：到着・終了の場所。到着を言わずに終了した区間は、出発で言った行き先も入れる（アプリの routeOf と同じ）
  const places = [];
  items.forEach((it, i) => {
    const next = items[i + 1];
    const p = it.rec['行き先・場所'];
    if (it.rec['種類'] !== '出発') { if (p) places.push(p); }
    else if (p && next && next.rec['種類'] === '終了' && next.rec['行き先・場所'] !== p) places.push(p);
  });
  tripSheet.getRange(row, 2, 1, TRIP_HEADERS.length - 1).setValues([[
    cell_(first['日付'], '日付'), cell_(first['時刻'], '時刻'), cell_(lastRec['時刻'], '時刻'), first['出発地'], places.join(' → '),
    first['区分'], first['目的'], start, prev, prev - start]]);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const req = JSON.parse(e.postData.contents);

    if (req.action === 'log' || req.action === 'trip') {
      const isLog = req.action === 'log';
      const headers = isLog ? LOG_HEADERS : TRIP_HEADERS;
      const sheet = sheetOf_(isLog ? LOG_SHEET : TRIP_SHEET, headers);
      const id = req.row[headers[0]];
      if (!findRow_(sheet, id)) {
        sheet.appendRow(headers.map(h => (h in req.row ? cell_(req.row[h], h) : '')));
      }
      return json_({ ok: true });
    }

    if (req.action === 'edit') {
      const sheet = sheetOf_(LOG_SHEET, LOG_HEADERS);
      const row = findRow_(sheet, req.logId);
      if (!row) return json_({ ok: true });
      const tripId = sheet.getRange(row, 2).getValue();
      Object.keys(req.changes).forEach(h => {
        if (EDITABLE.indexOf(h) === -1) return;
        sheet.getRange(row, LOG_HEADERS.indexOf(h) + 1).setValue(cell_(req.changes[h], h));
      });
      // 区分・目的は運行全体で同じにする
      ['区分', '目的'].forEach(h => {
        if (!(h in req.changes)) return;
        const col = LOG_HEADERS.indexOf(h) + 1;
        tripLogs_(tripId).forEach(it => sheet.getRange(it.row, col).setValue(req.changes[h]));
      });
      recompute_(tripId);
      return json_({ ok: true });
    }

    if (req.action === 'undo') {
      // 直前の記録を取り消す（記録ID・運行IDで探して消す）
      [[LOG_SHEET, LOG_HEADERS, req.logId], [TRIP_SHEET, TRIP_HEADERS, req.tripId]].forEach(([name, headers, id]) => {
        if (!id) return;
        const sheet = sheetOf_(name, headers);
        const row = findRow_(sheet, id);
        if (row) sheet.deleteRow(row);
      });
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'unknown action' });
  } finally {
    lock.releaseLock();
  }
}

// アプリを開いたときに、最近の記録と運行を返す。?trip=運行ID ならその運行の記録だけ
function doGet(e) {
  if (e && e.parameter && e.parameter.trip) {
    return json_({ ok: true, logs: tripLogs_(e.parameter.trip).map(it => it.rec) });
  }
  const read = (name, headers, n) => {
    const sheet = sheetOf_(name, headers);
    const last = sheet.getLastRow();
    if (last < 2) return [];
    const from = Math.max(2, last - n + 1);
    return readRows_(sheet, headers, from, last - from + 1);
  };
  return json_({ ok: true, logs: read(LOG_SHEET, LOG_HEADERS, 50), trips: read(TRIP_SHEET, TRIP_HEADERS, 2000) });
}
