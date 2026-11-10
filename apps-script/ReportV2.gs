/**
 * ZCRIM — Family Report (new format)
 *
 * Add this file to the Apps Script project bound to the "Project Declared" sheet,
 * together with ReportV2.html. It only READS the SurveyData, MemberData and PhotoLog
 * tabs; it never writes to the sheet and does not touch any other code.
 *
 * Open it from the ZCRIM Admin menu (add one line to your onOpen menu):
 *     .addItem('🖨️ Family Report (new)', 'showFamilyReportV2')
 * or run showFamilyReportV2 once from the editor.
 */

function showFamilyReportV2() {
  var html = HtmlService.createHtmlOutputFromFile('ReportV2')
    .setWidth(1000)
    .setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, 'Family Report');
}

// ── Sheet helpers ───────────────────────────────────────────────

function rv2Table_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return { h: {}, rows: [] };
  var v = sh.getDataRange().getValues();
  var h = {};
  v[0].forEach(function (k, i) { h[String(k).trim()] = i; });
  return { h: h, rows: v.slice(1) };
}

function rv2Get_(t, row, col) {
  var i = t.h[col];
  if (i === undefined) return '';
  var x = row[i];
  return x === null || x === undefined ? '' : x;
}

// Timestamps are written as text "dd/mm/yyyy hh:mm:ss AM" — but a cell may also hold a real Date.
function rv2Ms_(x) {
  if (x instanceof Date) return x.getTime();
  var m = String(x).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})\s*(AM|PM)?/i);
  if (!m) return NaN;
  var h = +m[4];
  if (m[7]) { h = h % 12; if (m[7].toUpperCase() === 'PM') h += 12; }
  return new Date(+m[3], +m[2] - 1, +m[1], h, +m[5], +m[6]).getTime();
}

function rv2Text_(x) {
  if (x instanceof Date) {
    return Utilities.formatDate(x, Session.getScriptTimeZone(), 'dd/MM/yyyy hh:mm:ss a');
  }
  return String(x);
}

function rv2SameTime_(a, b) {
  var x = rv2Ms_(a), y = rv2Ms_(b);
  return !isNaN(x) && !isNaN(y) && Math.abs(x - y) <= 2000;
}

// ── Called from the dialog ──────────────────────────────────────

/** List surveys matching a Family ID or name (newest first). Empty query → latest 40. */
function rv2Search(query) {
  var t = rv2Table_('SurveyData');
  var q = String(query || '').trim().toLowerCase();
  var out = [];
  t.rows.forEach(function (r, i) {
    var id = String(rv2Get_(t, r, 'FamilyID'));
    var name = String(rv2Get_(t, r, 'FamilyName'));
    if (!id) return;
    if (q && id.toLowerCase().indexOf(q) < 0 && name.toLowerCase().indexOf(q) < 0) return;
    out.push({
      row: i + 2,
      id: id,
      name: name,
      ward: String(rv2Get_(t, r, 'Ward')),
      zcrim: String(rv2Get_(t, r, 'ZCRIMName')),
      tag: String(rv2Get_(t, r, 'SurveyTag') || 'Original'),
      ts: rv2Text_(rv2Get_(t, r, 'Timestamp')),
      ms: rv2Ms_(rv2Get_(t, r, 'Timestamp')) || 0
    });
  });
  out.sort(function (a, b) { return b.ms - a.ms; });
  return out.slice(0, 40);
}

var RV2_PHOTO_KEYS = {
  'Surveyor_Selfie': 'selfie',
  'Full House (Front)': 'front',
  'Kitchen': 'kitchen',
  'With Family Members': 'familyPhoto',
  'Surrounding Area': 'surroundings',
  'Surveyor with Family': 'surveyorFamily'
};
var RV2_URL_COLS = {
  selfie: 'SelfieURL', front: 'Photo_Front', kitchen: 'Photo_Kitchen',
  familyPhoto: 'Photo_Family', surroundings: 'Photo_Surroundings', surveyorFamily: 'Photo_SurveyorFamily'
};

/** Everything the report needs for one SurveyData row. */
function rv2Report(rowNumber) {
  var sd = rv2Table_('SurveyData');
  var r = sd.rows[rowNumber - 2];
  if (!r) throw new Error('Survey row ' + rowNumber + ' not found.');

  var s = {};
  Object.keys(sd.h).forEach(function (k) { s[k] = rv2Text_(rv2Get_(sd, r, k)); });
  var id = s.FamilyID, ts = rv2Get_(sd, r, 'Timestamp'), tag = s.SurveyTag || 'Original';

  // Members of this survey: same family + same timestamp (fallback: same family + same survey tag)
  var md = rv2Table_('MemberData');
  function memberRows(strict) {
    return md.rows.filter(function (m) {
      if (String(rv2Get_(md, m, 'FamilyID')) !== id) return false;
      return strict ? rv2SameTime_(rv2Get_(md, m, 'Timestamp'), ts)
                    : String(rv2Get_(md, m, 'SurveyTag') || 'Original') === tag;
    });
  }
  var mrows = memberRows(true);
  if (!mrows.length) mrows = memberRows(false);
  var members = mrows.map(function (m) {
    return {
      no: +rv2Get_(md, m, 'MemberNo') || 0,
      name: String(rv2Get_(md, m, 'MemberName')),
      age: String(rv2Get_(md, m, 'Age')),
      income: String(rv2Get_(md, m, 'IncomeRange')),
      declared: String(rv2Get_(md, m, 'DeclaredIncome')),   // empty until the backend saves it
      occupation: String(rv2Get_(md, m, 'Occupation')),
      remarks: String(rv2Get_(md, m, 'Remarks'))
    };
  }).sort(function (a, b) { return a.no - b.no; });

  // Photos: PhotoLog rows with the same family + timestamp (the last one wins if retaken)
  var pl = rv2Table_('PhotoLog');
  var photos = {};
  pl.rows.forEach(function (p) {
    if (String(rv2Get_(pl, p, 'FamilyID')) !== id) return;
    if (!rv2SameTime_(rv2Get_(pl, p, 'Timestamp'), ts)) return;
    var key = RV2_PHOTO_KEYS[String(rv2Get_(pl, p, 'PhotoLabel'))];
    if (!key) return;
    photos[key] = {
      url: String(rv2Get_(pl, p, 'DriveURL')),
      fileId: String(rv2Get_(pl, p, 'FileID')),
      lat: String(rv2Get_(pl, p, 'GPS_Lat')),
      lng: String(rv2Get_(pl, p, 'GPS_Lng'))
    };
  });
  // Fallback: a real Drive link in the SurveyData column
  Object.keys(RV2_URL_COLS).forEach(function (k) {
    if (photos[k]) return;
    var u = s[RV2_URL_COLS[k]] || '';
    var m = u.match(/\/d\/([\w-]{20,})/);
    if (m) photos[k] = { url: u, fileId: m[1], lat: '', lng: '' };
  });
  // Embed the images so the dialog can print them (photos are ~100–200 KB each)
  Object.keys(photos).forEach(function (k) {
    try {
      var blob = DriveApp.getFileById(photos[k].fileId).getBlob();
      photos[k].src = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
    } catch (e) {
      photos[k].src = '';   // no access / deleted — the report shows a link instead
    }
  });

  return {
    survey: s,
    members: members,
    photos: photos,
    generated: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMM yyyy, h:mm a')
  };
}
