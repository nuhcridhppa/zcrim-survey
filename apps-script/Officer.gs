/**
 * ZCRIM — Officer portal (ADC / DCRIM / Admin)
 *
 * Add this file to the Apps Script project next to Code.gs and ReportV2.gs.
 * Read-only: it never writes to the sheet.
 *
 * Officers are entries in ZCRIMS (Code.gs) with  master:true  (Admin),
 * dcrim:true  (DCRIM) or  adc:true  (ADC). All officers see all blocks.
 * Test accounts' families are left out, like the other dashboards.
 *
 * Needs these names in Code.gs → handleApiRequest → allowed + fnMap:
 *   getOfficerOverview, getVillageFamilies, searchFamilies, getFamilyReport
 */

function isOfficer_(z) { return !!(z && (z.master || z.dcrim || z.adc)); }

function officerAuth_(token) {
  var auth = validateSession(token);
  if (!auth.ok) return { ok:false, msg:auth.msg, relogin:auth.relogin };
  if (!isOfficer_(auth.zcrim)) return { ok:false, msg:"Officer access only." };
  return auth;
}

function famRows_() {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Families");
  var last = sh.getLastRow();
  return last > 1 ? sh.getRange(2, 1, last - 1, 10).getValues().filter(notTestFamily) : [];
}

function famStatus_(r) {
  var s = String(r[9] || "").trim().toLowerCase();
  return s === "done" || s === "absent" || s === "resurvey" ? s : "pending";
}

// FamilyID → {count, last:{row, ts, zcrim, tag}} from SurveyData
function surveyIndex_() {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName("SurveyData");
  var idx = {};
  if (!sh || sh.getLastRow() < 2) return idx;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 35).getValues();
  v.forEach(function (r, i) {
    var id = String(r[3] || "").trim().toUpperCase();
    if (!id) return;
    var when = parseTS(r[0]);
    var e = idx[id] || (idx[id] = { count:0, last:null, lastMs:-1 });
    e.count++;
    var ms = when ? when.getTime() : 0;
    if (ms >= e.lastMs) {
      e.lastMs = ms;
      e.last = { row:i + 2, ts:when ? formatTS(when) : String(r[0]), zcrim:String(r[1] || ""), tag:String(r[34] || "Original") };
    }
  });
  return idx;
}

function zcrimNames_() {
  var m = {};
  ZCRIMS.forEach(function (z) { m[z.mobile] = z.name; });
  return m;
}

// ── Progress: summary, trend, block → village tree, ZCRIM board ──
function getOfficerOverview(token) {
  var auth = officerAuth_(token);
  if (!auth.ok) return auth;

  var fam = famRows_();
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName("SurveyData");
  var sur = sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().filter(notTestSurvey) : [];

  var sum = { total:0, done:0, absent:0, resurvey:0, pending:0 };
  var blocks = {};
  fam.forEach(function (r) {
    var st = famStatus_(r), b = String(r[6] || "").trim() || "Unknown", w = String(r[7] || "").trim() || "Unknown";
    sum.total++; sum[st]++;
    var B = blocks[b] || (blocks[b] = { block:b, total:0, done:0, absent:0, resurvey:0, pending:0, wards:{} });
    B.total++; B[st]++;
    var W = B.wards[w] || (B.wards[w] = { ward:w, total:0, done:0, absent:0, resurvey:0, pending:0 });
    W.total++; W[st]++;
  });

  var today = new Date().toDateString();
  var todayCount = sur.filter(function (r) { return tsDay(r[0]) === today; }).length;
  var trend = [];
  for (var d = 13; d >= 0; d--) {
    var dd = new Date(); dd.setDate(dd.getDate() - d);
    var ds = dd.toDateString();
    trend.push({ date:("0" + dd.getDate()).slice(-2) + "/" + ("0" + (dd.getMonth() + 1)).slice(-2),
                 count:sur.filter(function (r) { return tsDay(r[0]) === ds; }).length });
  }

  var stats = {};
  fam.forEach(function (r) {
    var m = String(r[8] || "").trim(), st = famStatus_(r);
    var s = stats[m] || (stats[m] = { total:0, done:0, absent:0 });
    s.total++; if (st === "done") s.done++; if (st === "absent") s.absent++;
  });
  var board = ZCRIMS.filter(function (z) { return !isOfficer_(z) && !z.test; }).map(function (z) {
    var s = stats[z.mobile] || { total:0, done:0, absent:0 };
    return { name:z.name, block:z.block, total:s.total, done:s.done, absent:s.absent,
             pending:s.total - s.done - s.absent, pct:s.total ? Math.round(s.done / s.total * 100) : 0 };
  }).sort(function (a, b) { return b.done - a.done; });

  sum.pct = sum.total ? Math.round(sum.done / sum.total * 100) : 0;
  sum.todayCount = todayCount;
  return {
    ok:true, summary:sum, trend:trend, board:board, updated:formatTS(new Date()),
    blocks:Object.keys(blocks).sort().map(function (k) {
      var B = blocks[k];
      B.wards = Object.keys(B.wards).sort().map(function (w) { return B.wards[w]; });
      return B;
    })
  };
}

function familyOut_(r, idx, names) {
  var id = String(r[0] || "").trim(), s = idx[id.toUpperCase()];
  return {
    id:id, name:String(r[1] || ""), age:String(r[3] || ""), gender:String(r[4] || ""),
    block:String(r[6] || ""), ward:String(r[7] || ""), status:famStatus_(r),
    zcrim:names[String(r[8] || "").trim()] || String(r[8] || ""),
    surveys:s ? s.count : 0, last:s ? s.last : null
  };
}

// ── Village → every family with its status ──
function getVillageFamilies(token, block, ward) {
  var auth = officerAuth_(token);
  if (!auth.ok) return auth;
  block = String(block || "").trim(); ward = String(ward || "").trim();
  var idx = surveyIndex_(), names = zcrimNames_();
  var list = famRows_().filter(function (r) {
    return String(r[7] || "").trim() === ward && (!block || String(r[6] || "").trim() === block);
  }).map(function (r) { return familyOut_(r, idx, names); });
  return { ok:true, block:block, ward:ward, families:list };
}

// ── PPP ID / name search across all blocks ──
function searchFamilies(token, query) {
  var auth = officerAuth_(token);
  if (!auth.ok) return auth;
  var q = sanitizeStr(query, 60).toUpperCase();
  if (q.length < 3) return { ok:false, msg:"Type at least 3 characters." };
  var idx = surveyIndex_(), names = zcrimNames_(), out = [];
  famRows_().some(function (r) {
    var id = String(r[0] || "").toUpperCase(), nm = String(r[1] || "").toUpperCase();
    if (id.indexOf(q) > -1 || nm.indexOf(q) > -1) out.push(familyOut_(r, idx, names));
    return out.length >= 50;
  });
  // exact PPP ID first
  out.sort(function (a, b) { return (b.id.toUpperCase() === q) - (a.id.toUpperCase() === q); });
  return { ok:true, families:out };
}

// ── Full report for one family (latest survey, or the SurveyData row asked for) ──
function getFamilyReport(token, familyId, surveyRow) {
  var auth = officerAuth_(token);
  if (!auth.ok) return auth;
  var id = sanitizeStr(familyId, 20).toUpperCase();
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName("SurveyData");
  var v = sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 35).getValues() : [];
  var surveys = [];
  v.forEach(function (r, i) {
    if (String(r[3] || "").trim().toUpperCase() !== id) return;
    var when = parseTS(r[0]);
    surveys.push({ row:i + 2, ts:when ? formatTS(when) : String(r[0]), ms:when ? when.getTime() : 0,
                   tag:String(r[34] || "Original"), zcrim:String(r[1] || "") });
  });
  if (!surveys.length) return { ok:false, msg:"No survey submitted yet for " + id + "." };
  surveys.sort(function (a, b) { return b.ms - a.ms; });
  var pick = surveys[0];
  var want = parseInt(surveyRow, 10);
  surveys.forEach(function (s) { if (s.row === want) pick = s; });   // only rows of THIS family
  var rep = rv2Report(pick.row);
  rep.ok = true;
  rep.surveys = surveys;
  rep.row = pick.row;
  return rep;
}
