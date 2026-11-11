/**
 * ZCRIM — Officer portal (ADC / DCRIM / Admin)
 *
 * Add this file to the Apps Script project next to Code.gs and ReportV2.gs.
 *
 * Officers are entries in ZCRIMS (Code.gs) with  master:true  (Admin),
 * dcrim:true  (DCRIM) or  adc:true  (ADC). All officers see all blocks.
 * Test accounts' families are left out, like the other dashboards.
 *
 * Needs these names in Code.gs → handleApiRequest → allowed + fnMap:
 *   getOfficerOverview, getVillageFamilies, searchFamilies, getFamilyReport,
 *   getFamilyForSurvey, pushReverification
 * Writes only in pushReverification (Families status + ResurveyLog); everything else is read-only.
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
  rep.history = resurveyHistory_(id);
  rep.canPush = !!(auth.zcrim.master || auth.zcrim.adc);
  rep.surveys = surveys;
  rep.row = pick.row;
  return rep;
}

// ─────────────────────────────────────────────────────────────────
//  OFFICER SURVEY — open any family in the survey form
// ─────────────────────────────────────────────────────────────────
function getFamilyForSurvey(token, familyId) {
  var auth = officerAuth_(token);
  if (!auth.ok) return auth;
  var id = sanitizeStr(familyId, 20).toUpperCase();
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Families");
  var last = sh.getLastRow();
  var data = last > 1 ? sh.getRange(1, 1, last, 10).getValues() : [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (String(r[0] || "").trim().toUpperCase() !== id) continue;
    var status = famStatus_(r);
    var note = status === "resurvey" ? resurveyNote_(id) : null;
    return { ok:true, rowIndex:i + 1, isResurvey:status === "resurvey" || status === "done", status:status, note:note,
             family:{ id:r[0], name:r[1], mobile:r[2], age:r[3], gender:r[4], address:r[5], block:r[6],
                      ward:String(r[7] || "").trim(), assignedTo:String(r[8] || "").trim() } };
  }
  return { ok:false, msg:"Family not found: " + id };
}

// ─────────────────────────────────────────────────────────────────
//  RE-VERIFICATION (grievance) — ADC / Admin send a family back to its ZCRIM
// ─────────────────────────────────────────────────────────────────
var RESURVEY_HDR = ["Timestamp","AdminName","FamilyID","FamilyName","PrevStatus","Reason",
                    "Category","ComplainantName","ComplainantMobile","GrievanceRef","Source"];
var GRIEVANCE_TYPES = ["Income wrongly verified","Family members incorrect","Assets / house details incorrect",
                       "Wrong family surveyed","Other"];

function pushReverification(token, familyId, info) {
  var auth = officerAuth_(token);
  if (!auth.ok) return auth;
  if (!(auth.zcrim.master || auth.zcrim.adc)) return { ok:false, msg:"Only ADC or Admin can send a family for re-verification." };
  info = info || {};
  var id       = sanitizeStr(familyId, 20).toUpperCase();
  var category = GRIEVANCE_TYPES.indexOf(info.category) > -1 ? info.category : "Other";
  var details  = cellSafe(sanitizeStr(info.details, 500));
  var cName    = cellSafe(sanitizeStr(info.complainantName, 80));
  var cMobile  = String(info.complainantMobile || "").replace(/[^\d]/g, "").slice(0, 10);
  var ref      = cellSafe(sanitizeStr(info.reference, 60));
  if (details.length < 10) return { ok:false, msg:"Please describe the grievance (at least 10 characters)." };
  if (cMobile && !validMobile(cMobile)) return { ok:false, msg:"Complainant mobile must be a valid 10-digit number." };

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var famSh = ss.getSheetByName("Families");
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) return { ok:false, msg:"Busy — please try again in a moment." };
  try {
    var last = famSh.getLastRow();
    var data = last > 1 ? famSh.getRange(1, 1, last, 10).getValues() : [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      if (String(r[0] || "").trim().toUpperCase() !== id) continue;
      var prev = famStatus_(r);
      if (prev === "resurvey") return { ok:false, msg:id + " is already waiting for re-verification." };
      if (prev === "pending")  return { ok:false, msg:id + " has not been surveyed yet — it is already in the ZCRIM's list." };
      famSh.getRange(i + 1, 10).setValue("resurvey");
      famSh.getRange(i + 1, 1, 1, 10).setBackground("#ffe0b2");

      var logSh = ss.getSheetByName("ResurveyLog") || ss.insertSheet("ResurveyLog");
      var hdr = logSh.getRange(1, 1, 1, RESURVEY_HDR.length).getValues()[0];
      for (var c = 0; c < RESURVEY_HDR.length; c++) {
        if (!String(hdr[c] || "").trim()) logSh.getRange(1, c + 1).setValue(RESURVEY_HDR[c]).setFontWeight("bold");
      }
      var ts = formatTS(new Date());
      logSh.getRange(logSh.getLastRow() + 1, 1).setNumberFormat("@");
      logSh.appendRow([ts, auth.zcrim.name, r[0], r[1], prev, details, category, cName, cMobile, ref, "Grievance"]);

      // Tell the assigned ZCRIM (email failure must not undo the request)
      var assigned = ZCRIMS.filter(function (z) { return z.mobile === String(r[8] || "").trim(); })[0];
      var emailed = false;
      if (assigned && assigned.email) {
        try {
          GmailApp.sendEmail(assigned.email, "Re-verification requested: " + r[0] + " (" + r[1] + ")",
            "Dear " + assigned.name + ",\n\n" + auth.zcrim.name + " has sent this family back for re-verification.\n\n" +
            "PPP ID   : " + r[0] + "\nName     : " + r[1] + "\nVillage  : " + r[7] + ", " + r[6] + "\n" +
            "Grievance: " + category + "\nDetails  : " + details + (ref ? "\nRef. no. : " + ref : "") + "\n\n" +
            "It is now at the top of your survey list for " + r[7] + ". Please re-survey it carefully.\n\n— ZCRIM Survey Admin",
            { name:"ZCRIM Survey Admin", replyTo:ADMIN_EMAIL });
          emailed = true;
        } catch (e) { Logger.log("Re-verification email failed: " + e.message); }
      }
      return { ok:true, msg:"Sent to " + (assigned ? assigned.name : "the assigned ZCRIM") + " for re-verification.", emailed:emailed };
    }
    return { ok:false, msg:"Family not found: " + id };
  } finally { lock.releaseLock(); }
}

// Latest re-verification request for a family (shown to the ZCRIM and in the report)
function resurveyNote_(familyId) {
  var list = resurveyHistory_(familyId);
  return list.length ? list[list.length - 1] : null;
}
function resurveyHistory_(familyId) {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName("ResurveyLog");
  if (!sh || sh.getLastRow() < 2) return [];
  var id = String(familyId || "").trim().toUpperCase();
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, RESURVEY_HDR.length).getValues();
  return v.filter(function (r) { return String(r[2] || "").trim().toUpperCase() === id; }).map(function (r) {
    var when = parseTS(r[0]);
    return { at:when ? formatTS(when) : String(r[0]), by:String(r[1] || ""), reason:String(r[5] || ""),
             category:String(r[6] || ""), complainant:String(r[7] || ""), reference:String(r[9] || "") };
  });
}
