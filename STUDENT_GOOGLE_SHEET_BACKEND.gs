var _OUT_CALLBACK = "";

var SHEET_ID = ""; // This version uses the Google Sheet to which this Apps Script is bound.

var PROGRAMME_COURSE_CODES = ["BASIC","SCREENER(INT)","SCREENER(REC)","INSTRUCTOR","BASIC REFRESHER","INDUCTION","OJT"];

// For a bound Apps Script, use the spreadsheet this script belongs to.
// If this is a standalone script, the optional SHEET_ID / saved property is used.

function getMainSpreadsheet() {
  // IMPORTANT: This Student backend is intended to be BOUND to the new
  // Student Google Sheet. Never fall back to an old hard-coded spreadsheet.
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (ignored) {}

  // Optional Script Property support is retained only for an administrator
  // who explicitly configured MAIN_SPREADSHEET_ID. No old/default ID is used.
  var props = PropertiesService.getScriptProperties();
  var savedId = String(props.getProperty("MAIN_SPREADSHEET_ID") || "").trim();
  if (savedId) {
    try { return SpreadsheetApp.openById(savedId); }
    catch (err) { throw new Error("Configured MAIN_SPREADSHEET_ID is not accessible. Bind this script to the Student Google Sheet."); }
  }

  throw new Error("Student backend is not bound to a Google Sheet. Open the NEW Student Google Sheet, go to Extensions > Apps Script, paste this backend there, save it, run setupStudentBackend once, and then reload the portal.");
}

function setupStudentBackend() {
  var ss = getMainSpreadsheet();
  initializeRequiredSheets_(ss);
  // Also ensure the registration sheet has exactly the requested six headers.
  ensureRegisteredStudentsSheet_();
  PropertiesService.getScriptProperties().setProperty("STUDENT_SHEETS_READY", "1");
  CacheService.getScriptCache().put("STUDENT_SHEETS_READY", "1", 21600);
  SpreadsheetApp.flush();
  return "Student backend ready: " + ss.getName();
}

// Short setup alias. This is intentionally simple so it appears clearly in
// the Apps Script function dropdown. It is safe to run repeatedly.
function setupStudent() {
  return setupStudentBackend();
}

function ensureStudentSheetsOnce_(ss) {
  var cache = CacheService.getScriptCache();
  if(cache.get("STUDENT_SHEETS_READY") === "1") return;
  var props = PropertiesService.getScriptProperties();
  if(String(props.getProperty("STUDENT_SHEETS_READY") || "") === "1") { cache.put("STUDENT_SHEETS_READY", "1", 21600); return; }
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if(String(props.getProperty("STUDENT_SHEETS_READY") || "") !== "1") {
      initializeRequiredSheets_(ss);
      ensureRegisteredStudentsSheet_();
      props.setProperty("STUDENT_SHEETS_READY", "1");
    }
    cache.put("STUDENT_SHEETS_READY", "1", 21600);
  } finally { try { lock.releaseLock(); } catch(ignore) {} }
}

function initializeRequiredSheets_(ss) {
  var defs = [
    ["Notices", ["TargetExam", "NoticeText"]],
    ["Schedule", ["ExamName", "PaperCode", "PaperName", "TargetExam", "StartTime", "EndTime", "Duration"]],
    ["Modules", ["ModuleName", "PaperCode", "PaperName", "TargetExam", "Duration"]],
    ["Materials", ["Title", "Subject", "TargetExam", "Link"]],
    ["History", ["Date", "CISF NO", "Name", "Paper", "PaperName", "Score", "Result", "Type"]],
    ["PermanentHistory", ["Date", "CISF NO", "Name", "Paper", "PaperName", "Score", "Result", "Type", "Batch", "Responses"]],
    ["QuestionUpdates", ["PaperCode", "QID", "QuestionEnglish", "QuestionHindi", "AEnglish", "AHindi", "BEnglish", "BHindi", "CEnglish", "CHindi", "DEnglish", "DHindi", "Correct"]],
    ["Registered Students", ["Serial Number", "CISF NO", "Rank", "Name", "Paper Code", "Batch Name", "Password", "PassChanged"]]
  ];
  defs.forEach(function(def) {
    var sh = ss.getSheetByName(def[0]);
    if (!sh) {
      sh = ss.insertSheet(def[0]);
      sh.getRange(1, 1, 1, def[1].length).setValues([def[1]]);
      sh.getRange(1, 1, 1, def[1].length).setBackground("#004d99").setFontColor("white").setFontWeight("bold");
      sh.setFrozenRows(1);
    }
  });
}

function resetMainSpreadsheetBinding() {
  PropertiesService.getScriptProperties().deleteProperty("MAIN_SPREADSHEET_ID");
  return getMainSpreadsheet().getId();
}
  
function getAdminPassword() {
  // Admin password is stored persistently in this Apps Script project.
  // Default is the previously working ASTI admin password until the administrator changes it.
  var saved = String(PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD") || "").trim();
  return saved || "Rakesh";
}

function setAdminPassword_(newPass) {
  newPass = String(newPass || "").trim();
  PropertiesService.getScriptProperties().setProperty("ADMIN_PASSWORD", newPass);
  return true;
}
  
function isMatch(targetVal, userExam) {
  if (!targetVal) return true;
  var t = String(targetVal).toUpperCase().trim();
  var u = String(userExam).toUpperCase().trim();
  if (t === "ALL" || t === "" || t === u || u === "ALL") return true;
  if (t.indexOf(u) !== -1 || u.indexOf(t) !== -1) return true; 
  return false;
}

function getColIdx(headers, name) {
  for(var i=0; i<headers.length; i++) {
    if(String(headers[i]).replace(/\s/g,'').toLowerCase() === String(name).replace(/\s/g,'').toLowerCase()) return i;
  }
  return -1; 
}

function formatDt(dt) {
  if (!(dt instanceof Date)) return String(dt);
  var d = dt.getDate(), m = dt.getMonth() + 1, y = dt.getFullYear();
  var hr = dt.getHours(), min = dt.getMinutes(), sec = dt.getSeconds();
  var pad = function(n) { return n < 10 ? '0'+n : n; };
  return pad(d) + '/' + pad(m) + '/' + y + ', ' + pad(hr) + ':' + pad(min) + ':' + pad(sec);
}

function outJSON(data) {
  var callback = String(_OUT_CALLBACK || "");
  if (callback && !/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) callback = "";

  var payload = JSON.stringify(data);

  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + payload + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}


function ensureQuestionBankSheet() {
  var sheet = ssForQuestions();
  return sheet;
}
function ssForQuestions() {
  return getMainSpreadsheet();
}
function getQuestionBankSheet() {
  var ss = ssForQuestions();
  var sh = ss.getSheetByName("QuestionBank");
  if(!sh) {
    sh = ss.insertSheet("QuestionBank");
    sh.appendRow(["PaperCode","Q.No","English Question","A","B","C","D","Answer","Marks","Hindi Question","Hindi A","Hindi B","Hindi C","Hindi D"]);
  }
  return sh;
}
function listQuestionPapers() {
  var sh = getQuestionBankSheet(), out = {}, last = sh.getLastRow();
  if(last > 1) {
    var vals = sh.getRange(2,1,last-1,1).getDisplayValues();
    vals.forEach(function(r){ var p=String(r[0]).trim(); if(p) out[p]=true; });
  }
  return Object.keys(out).sort();
}
function questionRowToArray(r) {
  return [String(r[1]), Number(r[7]) || 0, String(r[2]||""), String(r[9]||""), [String(r[3]||""),String(r[10]||"")], [String(r[4]||""),String(r[11]||"")], [String(r[5]||""),String(r[12]||"")], [String(r[6]||""),String(r[13]||"")], String(r[8]||"")];
}
function getProgrammeCourseCodes_() {
  var set = {};
  PROGRAMME_COURSE_CODES.forEach(function(c){
    c = String(c || "").trim();
    if(c) set[c.toUpperCase()] = c;
  });
  var ss = getMainSpreadsheet();
  ss.getSheets().forEach(function(sh){
    var name = String(sh.getName() || "").trim();
    if(/^Students_/i.test(name)) {
      var fromName = name.replace(/^Students_/i, "").trim();
      if(fromName && fromName.toUpperCase() !== "ALL") set[fromName.toUpperCase()] = fromName;
    }
    try {
      var lastCol = sh.getLastColumn();
      if(lastCol > 0) {
        var h = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
        var idx = getColIdx(h,"CourseCode");
        if(idx >= 0 && sh.getLastRow() > 1) {
          sh.getRange(2,idx+1,sh.getLastRow()-1,1).getDisplayValues().forEach(function(r){
            var cv = String(r[0] || "").trim();
            if(cv && cv.toUpperCase() !== "ALL") set[cv.toUpperCase()] = cv;
          });
        }
      }
    } catch(ignore) {}
  });
  try {
    var props = PropertiesService.getScriptProperties().getProperties();
    Object.keys(props).forEach(function(k){
      if(k.indexOf("BATCH_COURSE_") === 0) {
        var cv = String(props[k] || "").trim();
        if(cv && cv.toUpperCase() !== "ALL") set[cv.toUpperCase()] = cv;
      }
    });
  } catch(ignoreProps) {}
  return Object.keys(set).map(function(k){return set[k];}).sort(function(a,b){return String(a).localeCompare(String(b));});
}

function ensureRegisteredStudentsSheet_() {
  var ss = getMainSpreadsheet();
  var sh = ss.getSheetByName("Registered Students");
  if(!sh) sh = ss.insertSheet("Registered Students");

  var expected = ["Serial Number","CISF NO","Rank","Name","Paper Code","Batch Name","Password","PassChanged"];
  var lastCol = Math.max(8, sh.getLastColumn() || 8);
  var hdr = sh.getRange(1,1,1,lastCol).getDisplayValues()[0];
  var current = hdr.slice(0,8).map(function(v){return String(v||"").trim();});

  // Create the new 8-column registration structure. If an older 6-column
  // Registered Students sheet already exists, preserve all student data and
  // append Password + PassChanged instead of deleting/recreating the sheet.
  var headerMatches = current.join("|") === expected.join("|");
  if(!headerMatches) {
    if(sh.getLastRow() === 0 || !String(hdr[0]||"").trim()) {
      sh.getRange(1,1,1,8).setValues([expected]);
    } else {
      var oldHdr = sh.getRange(1,1,1,Math.max(6, sh.getLastColumn())).getDisplayValues()[0];
      var oldVals = sh.getLastRow()>1 ? sh.getRange(2,1,sh.getLastRow()-1,Math.max(6,sh.getLastColumn())).getDisplayValues() : [];
      var oldCols = {};
      oldHdr.forEach(function(v,i){ oldCols[String(v||"").trim().toUpperCase()] = i; });
      sh.getRange(1,1,1,8).setValues([expected]);
      if(oldVals.length) {
        var migrated=[];
        oldVals.forEach(function(r){
          var roll=oldCols["CISF NO"]!=null?String(r[oldCols["CISF NO"]]||"").trim():"";
          if(!roll) return;
          var password = oldCols["PASSWORD"]!=null ? String(r[oldCols["PASSWORD"]]||"").trim() : "";
          var changed = oldCols["PASSCHANGED"]!=null ? String(r[oldCols["PASSCHANGED"]]||"").trim().toUpperCase() : "";
          var savedProp=String(PropertiesService.getScriptProperties().getProperty(getRegisteredStudentPasswordKey_(roll))||"").trim();
          if(savedProp && savedProp !== roll) { password=savedProp; changed="TRUE"; }
          if(!password) password="12345";
          if(changed !== "TRUE") changed="FALSE";
          migrated.push([
            oldCols["SERIAL NUMBER"]!=null ? (r[oldCols["SERIAL NUMBER"]]||migrated.length+1) : migrated.length+1,
            roll,
            oldCols["RANK"]!=null ? (r[oldCols["RANK"]]||"") : "",
            oldCols["NAME"]!=null ? (r[oldCols["NAME"]]||"") : "",
            oldCols["PAPER CODE"]!=null ? (r[oldCols["PAPER CODE"]]||"") : "ALL",
            oldCols["BATCH NAME"]!=null ? (r[oldCols["BATCH NAME"]]||"") : "",
            password,
            changed
          ]);
        });
        if(migrated.length) sh.getRange(2,1,migrated.length,8).setValues(migrated);
      }
    }
  }

  // Ensure every existing registration has the new password/state columns.
  if(sh.getLastRow()>1) {
    var vals=sh.getRange(2,1,sh.getLastRow()-1,8).getDisplayValues();
    var h8=sh.getRange(1,1,1,8).getDisplayValues()[0];
    var rollCol=getColIdx(h8,"CISF NO"), passCol=getColIdx(h8,"Password"), changedCol=getColIdx(h8,"PassChanged");
    for(var i=0;i<vals.length;i++) {
      var roll=rollCol>=0?String(vals[i][rollCol]||"").trim():"";
      if(!roll) continue;
      var pass=passCol>=0?String(vals[i][passCol]||"").trim():"";
      var changed=changedCol>=0?String(vals[i][changedCol]||"").trim().toUpperCase():"";
      if(!pass) { sh.getRange(i+2,passCol+1).setValue("12345"); pass="12345"; }
      if(changed!=="TRUE") sh.getRange(i+2,changedCol+1).setValue("FALSE");
      // Keep Script Properties synchronized for backward compatibility.
      setRegisteredStudentPassword_(roll,pass,changed==="TRUE");
    }
  }

  sh.getRange(1,1,1,8).setBackground("#004d99").setFontColor("white").setFontWeight("bold");
  sh.setFrozenRows(1);
  return sh;
}
function isRegisteredStudent_(roll) {
  var r=String(roll||"").trim();
  if(!r) return false;
  var sh=ensureRegisteredStudentsSheet_();
  if(sh.getLastRow()<2) return false;
  var vals=sh.getDataRange().getDisplayValues();
  var h=vals[0]||[];
  var c=getColIdx(h,"CISF NO");
  if(c<0) return false;
  for(var i=1;i<vals.length;i++){
    if(String(vals[i][c]||"").trim()===r) return true;
  }
  return false;
}


// Fast login index: keeps a compact CISF -> registration record map in Apps Script cache.
// It is rebuilt automatically after registration/password changes and expires after 60 seconds.
var ASTI_LOGIN_INDEX_KEY_ = "ASTI_REG_LOGIN_INDEX";
var ASTI_LOGIN_INDEX_TTL_ = 300; // seconds; cache is also cleared on every registration/password change

// Read the cached index. Large lists are stored in several chunks because Apps Script
// cache values are limited to ~100 KB each (a single put silently failed for big batches).
function readLoginIndexCache_(cache) {
  try {
    var meta = cache.get(ASTI_LOGIN_INDEX_KEY_ + "_N");
    if (meta) {
      var n = parseInt(meta, 10), keys = [];
      for (var i = 0; i < n; i++) keys.push(ASTI_LOGIN_INDEX_KEY_ + "_" + i);
      var parts = cache.getAll(keys), txt = "";
      for (var j = 0; j < n; j++) { var part = parts[keys[j]]; if (part == null) return null; txt += part; }
      return JSON.parse(txt);
    }
    var single = cache.get(ASTI_LOGIN_INDEX_KEY_);
    if (single) return JSON.parse(single);
  } catch (ignore) {}
  return null;
}
function writeLoginIndexCache_(cache, out) {
  try {
    var txt = JSON.stringify(out), size = 90000, n = Math.ceil(txt.length / size) || 1, obj = {};
    if (n === 1) { cache.put(ASTI_LOGIN_INDEX_KEY_, txt, ASTI_LOGIN_INDEX_TTL_); return; }
    for (var i = 0; i < n; i++) obj[ASTI_LOGIN_INDEX_KEY_ + "_" + i] = txt.substr(i * size, size);
    obj[ASTI_LOGIN_INDEX_KEY_ + "_N"] = String(n);
    cache.putAll(obj, ASTI_LOGIN_INDEX_TTL_);
  } catch (ignore2) {}
}

function normalizeRollKey_(roll) {
  return String(roll || "").trim().toUpperCase();
}

// Fast login: only answers when the cached index has the student AND the password matches.
// Anything else (cache miss, unknown roll, wrong password) falls through to the normal path,
// so behaviour is identical - it just skips opening the spreadsheet.
function fastLoginFromCache_(e) {
  var cache = CacheService.getScriptCache();
  var idx = readLoginIndexCache_(cache);
  if (!idx) return null;
  var requestedRoll = String(e.parameter.roll || "").trim();
  var requestedPass = String(e.parameter.pass || "").trim();
  var rec = idx[normalizeRollKey_(requestedRoll)];
  if (!rec || String(rec.pass) !== requestedPass) return null;
  var reqForce = e.parameter.forceLogin, reqToken = e.parameter.deviceToken || "UNKNOWN_DEVICE";
  var sessionKey = "LOGIN_" + requestedRoll.toLowerCase() + "_" + String(rec.name || "").toLowerCase().replace(/\s+/g, '');
  var activeSessionToken = cache.get(sessionKey);
  if (activeSessionToken && activeSessionToken !== reqToken && reqForce !== "true") {
    return outJSON({status: "ALREADY_LOGGED_IN", msg: "This Roll No is already logged in on another device! Please log out from there first."});
  }
  cache.put(sessionKey, reqToken, 14400);
  return outJSON({status: "SUCCESS", roll: rec.roll, name: rec.name, examType: rec.examType || "ALL", batch: rec.batch || "N/A", forcePassChange: !!rec.forcePassChange});
}

// Optional: add a time-driven trigger (every 5 minutes) on this function to keep login instant.
function warmStudentBackend() { getRegistrationLoginIndex_(); return "warm"; }

function getRegistrationLoginIndex_() {
  var cache = CacheService.getScriptCache();
  var cachedIdx = readLoginIndexCache_(cache);
  if (cachedIdx) return cachedIdx;

  // Only one request rebuilds the index; the others wait and then reuse it
  // (prevents many simultaneous logins from each scanning the whole sheet).
  var lock = LockService.getScriptLock();
  var locked = false;
  try { lock.waitLock(10000); locked = true; } catch (ignoreLock) {}
  try {
    var again = readLoginIndexCache_(cache);
    if (again) return again;
    var sh = ssForLogin_();
    if (!sh || sh.getLastRow() < 2) return {};
    var lastCol = Math.max(8, sh.getLastColumn());
    var vals = sh.getRange(1, 1, sh.getLastRow(), lastCol).getDisplayValues();
    var h = vals[0] || [];
    var rCol=getColIdx(h,"CISF NO"), nCol=getColIdx(h,"Name"), pCol=getColIdx(h,"Password");
    var bCol=getColIdx(h,"Batch Name"), cCol=getColIdx(h,"Paper Code"), chCol=getColIdx(h,"PassChanged");
    var out = {};
    if(rCol < 0) return out;
    for(var i=1;i<vals.length;i++) {
      var roll=String(vals[i][rCol]||"").trim();
      if(!roll) continue;
      var k=normalizeRollKey_(roll);
      if(out[k]) continue;
      var pass=pCol>=0?String(vals[i][pCol]||"").trim():"";
      if(!pass) pass="12345";
      out[k]={
        roll:roll,
        name:nCol>=0?String(vals[i][nCol]||"").trim():"",
        pass:pass,
        batch:bCol>=0?String(vals[i][bCol]||"").trim():"N/A",
        examType:cCol>=0?String(vals[i][cCol]||"ALL").trim():"ALL",
        forcePassChange: chCol<0 || String(vals[i][chCol]||"").trim().toUpperCase()!=="TRUE"
      };
    }
    writeLoginIndexCache_(cache, out);
    return out;
  } finally { if (locked) { try { lock.releaseLock(); } catch (ignoreRel) {} } }
}
function invalidateRegistrationLoginIndex_() {
  try {
    var c = CacheService.getScriptCache(), keys = [ASTI_LOGIN_INDEX_KEY_, ASTI_LOGIN_INDEX_KEY_ + "_N"];
    for (var i = 0; i < 40; i++) keys.push(ASTI_LOGIN_INDEX_KEY_ + "_" + i);
    c.removeAll(keys);
  } catch(ignore) {}
}
function ssForLogin_() { return getMainSpreadsheet().getSheetByName("Registered Students"); }

function getRegisteredStudentPasswordKey_(roll) {
  return "REGISTERED_STUDENT_PASSWORD_" + normalizeRollKey_(roll);
}
function getRegisteredStudentChangedKey_(roll) {
  return "REGISTERED_STUDENT_PASS_CHANGED_" + normalizeRollKey_(roll);
}
function getRegisteredStudentPassword_(roll) {
  var r = String(roll || "").trim();
  var saved = String(PropertiesService.getScriptProperties().getProperty(getRegisteredStudentPasswordKey_(r)) || "").trim();
  return saved || "12345";
}
function isRegisteredStudentPasswordChanged_(roll) {
  return String(PropertiesService.getScriptProperties().getProperty(getRegisteredStudentChangedKey_(roll)) || "").toUpperCase() === "TRUE";
}
function setRegisteredStudentPassword_(roll, pass, changed) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty(getRegisteredStudentPasswordKey_(roll), String(pass || ""));
  props.setProperty(getRegisteredStudentChangedKey_(roll), changed ? "TRUE" : "FALSE");
}

function readRegisteredStudents_() {
  var sh = ensureRegisteredStudentsSheet_();
  var vals = sh.getDataRange().getDisplayValues();
  var out = [], seen = {};
  if(vals.length <= 1) return out;
  var h = vals[0] || [];
  var sCol=getColIdx(h,"Serial Number"), rCol=getColIdx(h,"CISF NO"), rankCol=getColIdx(h,"Rank");
  var nCol=getColIdx(h,"Name"), pCol=getColIdx(h,"Paper Code"), bCol=getColIdx(h,"Batch Name");
  var passCol=getColIdx(h,"Password"), changedCol=getColIdx(h,"PassChanged");
  for(var i=1;i<vals.length;i++){
    var roll=rCol>=0?String(vals[i][rCol]||"").trim():"";
    if(!roll) continue;
    var key=roll.toUpperCase();
    if(seen[key]) continue;
    seen[key]=true;
    var pass=passCol>=0?String(vals[i][passCol]||"").trim():"";
    var changed=changedCol>=0?String(vals[i][changedCol]||"").trim().toUpperCase()==="TRUE":false;
    if(!pass) pass="12345";
    out.push({
      serial:sCol>=0?String(vals[i][sCol]||"").trim():String(out.length+1),
      roll:roll,
      rank:rankCol>=0?String(vals[i][rankCol]||"").trim():"",
      name:nCol>=0?String(vals[i][nCol]||"").trim():"",
      pass:pass,
      passChanged:changed?"TRUE":"FALSE",
      batch:bCol>=0?String(vals[i][bCol]||"").trim():"N/A",
      course:pCol>=0?String(vals[i][pCol]||"").trim():"ALL"
    });
  }
  return out;
}
function readProgrammeStudents_() {
  // Registered Students is now the single authoritative registration sheet.
  // Legacy Students_* sheets remain readable only as a backward-compatible
  // fallback for records created by older versions of the portal.
  var registered = readRegisteredStudents_();
  if(registered.length) return registered;

  var ss = getMainSpreadsheet();
  var out = [], seen = {};
  ss.getSheets().forEach(function(sh){
    var sheetName = String(sh.getName() || "").trim();
    if(/^Batch[ _-]/i.test(sheetName)) return;
    if(/^Registered Students$/i.test(sheetName)) return;
    if(sh.getLastRow() < 2) return;
    var data = sh.getDataRange().getDisplayValues(), h = data[0] || [];
    var rCol=getColIdx(h,"CISF NO"), nCol=getColIdx(h,"Name"), pCol=getColIdx(h,"Password"), bCol=getColIdx(h,"Batch");
    if(rCol < 0 || nCol < 0) return;
    var cCol=getColIdx(h,"ExamType");
    if(cCol < 0) cCol=getColIdx(h,"CourseCode");
    if(cCol < 0) cCol=getColIdx(h,"Course");
    var sheetCourse = /^Students_/i.test(sheetName) ? sheetName.replace(/^Students_/i,"").trim() : "";
    for(var i=1;i<data.length;i++){
      var roll=String(data[i][rCol]||"").trim();
      if(!roll) continue;
      var key=roll.toUpperCase();
      var course = cCol>=0 ? String(data[i][cCol]||"").trim() : "";
      if(!course) course=sheetCourse;
      if(!course) course="ALL";
      var obj={roll:roll,name:nCol>=0?String(data[i][nCol]||"").trim():"",pass:pCol>=0?String(data[i][pCol]||"").trim():"",batch:bCol>=0?String(data[i][bCol]||"").trim():"N/A",course:course};
      if(!seen[key]) { out.push(obj); seen[key]=true; }
    }
  });
  return out;
}

function doGet(e) {
  _OUT_CALLBACK = String(e && e.parameter && e.parameter.callback || "");
  var action = e.parameter.action;
  // FAST PATHS: these need only CacheService, so answer before opening the spreadsheet.
  if (action == "checkSession") {
    var fsRoll = String(e.parameter.roll).trim().toLowerCase();
    var fsName = String(e.parameter.name || "").trim().toLowerCase().replace(/\s+/g, '');
    var fsActive = CacheService.getScriptCache().get("LOGIN_" + fsRoll + "_" + fsName);
    if (fsActive && fsActive !== e.parameter.token) { return outJSON({status: "INVALID"}); }
    return outJSON({status: "VALID"});
  }
  if (action == "warmup") {
    getRegistrationLoginIndex_();
    return outJSON({status: "SUCCESS"});
  }
  if (action == "login") {
    var fastRes = fastLoginFromCache_(e);
    if (fastRes) return fastRes;
  }
  var ss = getMainSpreadsheet();
  // Initialize the bound Student Google Sheet only once. Re-running all
  // sheet/header checks on every request was a major source of latency.
  ensureStudentSheetsOnce_(ss);
  
  if (action == "listPapers") {
    return outJSON({status:"SUCCESS", papers:listQuestionPapers()});
  }
  if (action == "getQuestions") {
    var pCode = String(e.parameter.paperCode || "").trim();
    var shQ = getQuestionBankSheet(); var qOut=[];
    if(pCode && shQ.getLastRow()>1) {
      var qRows=shQ.getDataRange().getDisplayValues();
      for(var qi=1; qi<qRows.length; qi++) if(String(qRows[qi][0]).trim()===pCode) qOut.push(questionRowToArray(qRows[qi]));
    }
    return outJSON({status:"SUCCESS", paperCode:pCode, questions:qOut});
  }
  
  if (action == "logVisit") {
    var props = PropertiesService.getScriptProperties();
    var visits = parseInt(props.getProperty('site_visits') || '0') + 1;
    props.setProperty('site_visits', visits.toString());
    return outJSON({status: "SUCCESS"});
  }
  
  if (action == "checkSession") {
    var rollStr = String(e.parameter.roll).trim().toLowerCase();
    var nameStr = String(e.parameter.name || "").trim().toLowerCase().replace(/\s+/g, '');
    var reqToken = e.parameter.token;
    var sessionKey = "LOGIN_" + rollStr + "_" + nameStr;
    
    var activeToken = CacheService.getScriptCache().get(sessionKey);
    if (activeToken && activeToken !== reqToken) { return outJSON({status: "INVALID"}); }
    return outJSON({status: "VALID"});
  }
  
  if (action == "login") {
    // FAST LOGIN: do not download/loop through every student row.
    // Find the CISF number directly in the CISF NO column, then read only that row.
    // This preserves the existing password check, first-login flag and single-device logic.
    var requestedRoll = String(e.parameter.roll || "").trim();
    var requestedPass = String(e.parameter.pass || "").trim();

    // New registration model: one Registered Students sheet for all courses.
    // Use a short-lived cached index so login does not scan the sheet on every attempt.
    var loginIndex = getRegistrationLoginIndex_();
    var rec = loginIndex[normalizeRollKey_(requestedRoll)];
    if(rec && String(rec.pass) === requestedPass) {
      var reqForce=e.parameter.forceLogin, reqToken=e.parameter.deviceToken || "UNKNOWN_DEVICE";
      var sessionKey="LOGIN_"+requestedRoll.toLowerCase()+"_"+String(rec.name||"").toLowerCase().replace(/\s+/g,'');
      var activeSessionToken=CacheService.getScriptCache().get(sessionKey);
      if(activeSessionToken && activeSessionToken!==reqToken && reqForce!=="true") {
        return outJSON({status:"ALREADY_LOGGED_IN",msg:"This Roll No is already logged in on another device! Please log out from there first."});
      }
      CacheService.getScriptCache().put(sessionKey,reqToken,14400);
      return outJSON({status:"SUCCESS", roll:rec.roll, name:rec.name, examType:rec.examType||"ALL", batch:rec.batch||"N/A", forcePassChange:!!rec.forcePassChange});
    }
    if(rec) return outJSON({status:"FAIL"});

    // Legacy fallback for students registered by older portal versions.
    var allSheets = ss.getSheets();
    for(var s=0; s<allSheets.length; s++) {
      var sSheet = allSheets[s];
      if(sSheet.getName().indexOf("Students") === -1) continue;
      if(sSheet.getLastRow() < 2 || sSheet.getLastColumn() < 1) continue;

      // Only the header row is read first.
      var h = sSheet.getRange(1, 1, 1, sSheet.getLastColumn()).getDisplayValues()[0];
      var rollCol = getColIdx(h, "CISF NO");
      var passCol = getColIdx(h, "Password");
      var nameCol = getColIdx(h, "Name");
      var examCol = getColIdx(h, "ExamType");
      var batchCol = getColIdx(h, "Batch");
      var pChangedCol = getColIdx(h, "PassChanged");
      if(pChangedCol === -1) {
        pChangedCol = h.length;
        sSheet.getRange(1, pChangedCol + 1).setValue("PassChanged");
        h.push("PassChanged");
      }
      if(rollCol === -1 || passCol === -1) continue;

      var rollRange = sSheet.getRange(2, rollCol + 1, sSheet.getLastRow() - 1, 1);
      var found = rollRange.createTextFinder(requestedRoll).matchEntireCell(true).findNext();
      if(!found) continue;

      var rowNo = found.getRow();
      var row = sSheet.getRange(rowNo, 1, 1, sSheet.getLastColumn()).getDisplayValues()[0];
      if(String(row[rollCol]).trim() !== requestedRoll || String(row[passCol]).trim() !== requestedPass) continue;

      var isFirstLogin = (pChangedCol === -1) || (String(row[pChangedCol]).trim().toUpperCase() !== "TRUE");
      var reqForce = e.parameter.forceLogin;
      var reqToken = e.parameter.deviceToken || "UNKNOWN_DEVICE";
      var studentName = nameCol !== -1 ? String(row[nameCol]).trim() : "";
      var sessionKey = "LOGIN_" + requestedRoll.toLowerCase() + "_" + studentName.toLowerCase().replace(/\s+/g, '');

      var activeSessionToken = CacheService.getScriptCache().get(sessionKey);
      if (activeSessionToken && activeSessionToken !== reqToken && reqForce !== "true") {
        return outJSON({status: "ALREADY_LOGGED_IN", msg: "This Roll No is already logged in on another device! Please log out from there first."});
      }
      CacheService.getScriptCache().put(sessionKey, reqToken, 14400);

      return outJSON({
        status: "SUCCESS",
        roll: row[rollCol],
        name: nameCol !== -1 ? row[nameCol] : "",
        examType: examCol !== -1 ? (row[examCol] || "ALL") : "ALL",
        batch: batchCol !== -1 ? (row[batchCol] || "N/A") : "N/A",
        forcePassChange: isFirstLogin
      });
    }
    return outJSON({status: "FAIL"});
  }
  
  if (action == "getDashboard") {
    var roll = String(e.parameter.roll).trim();
    var name = e.parameter.name ? String(e.parameter.name).trim().toLowerCase() : "";
    var eType = e.parameter.examType ? String(e.parameter.examType).toUpperCase().trim() : "ALL";
    var res = { schedule: [], modules: [], materials: [], historyU: [], historyP: [], notice: "Welcome to ASTI Exam Portal." };
    try {
      var notSheet = ss.getSheetByName("Notices");
      if(notSheet && notSheet.getLastRow() > 1) {
         var nots = notSheet.getDataRange().getDisplayValues(); var nh = nots[0];
         for(var i=1; i<nots.length; i++) { if(isMatch(nots[i][getColIdx(nh, "TargetExam")], eType)) res.notice = nots[i][getColIdx(nh, "NoticeText")]; }
      }
      var schSheet = ss.getSheetByName("Schedule");
      if(schSheet && schSheet.getLastRow() > 1) {
         var sch = schSheet.getDataRange().getDisplayValues(); var sh = sch[0];
         for(var i=1; i<sch.length; i++) { if(isMatch(sch[i][getColIdx(sh, "TargetExam")], eType)) res.schedule.push({name:sch[i][getColIdx(sh,"ExamName")], code:sch[i][getColIdx(sh,"PaperCode")], paperName:sch[i][getColIdx(sh,"PaperName")], target:sch[i][getColIdx(sh,"TargetExam")], start:sch[i][getColIdx(sh,"StartTime")], end:sch[i][getColIdx(sh,"EndTime")], duration:sch[i][getColIdx(sh,"Duration")]}); }
      }
      var modSheet = ss.getSheetByName("Modules");
      if(modSheet && modSheet.getLastRow() > 1) {
         var mod = modSheet.getDataRange().getDisplayValues(); var mh = mod[0];
         for(var i=1; i<mod.length; i++) { if(isMatch(mod[i][getColIdx(mh, "TargetExam")], eType)) res.modules.push({name:mod[i][getColIdx(mh,"ModuleName")], code:mod[i][getColIdx(mh,"PaperCode")], paperName:mod[i][getColIdx(mh,"PaperName")], target:mod[i][getColIdx(mh,"TargetExam")], duration:mod[i][getColIdx(mh,"Duration")]}); }
      }
      var matSheet = ss.getSheetByName("Materials");
      if(matSheet && matSheet.getLastRow() > 1) {
         var mat = matSheet.getDataRange().getDisplayValues(); var math = mat[0];
         for(var i=1; i<mat.length; i++) { if(isMatch(mat[i][getColIdx(math, "TargetExam")], eType)) res.materials.push({title:mat[i][getColIdx(math,"Title")], subject:mat[i][getColIdx(math,"Subject")], target:mat[i][getColIdx(math,"TargetExam")], link:mat[i][getColIdx(math,"Link")]}); }
      }
      
      var permSheet = ss.getSheetByName("PermanentHistory");
      var permStatusMap = {};
      if(permSheet && permSheet.getLastRow() > 1) {
          var pData = permSheet.getDataRange().getDisplayValues();
          var ph = pData[0];
          var prCol = getColIdx(ph, "CISF NO");
          var ppCol = getColIdx(ph, "Paper");
          var ptCol = getColIdx(ph, "Type");
          var presCol = getColIdx(ph, "Responses");
          var pscoreCol = getColIdx(ph, "Score");
          var presultCol = getColIdx(ph, "Result");
          var pdCol = getColIdx(ph, "Date");
          
          if(prCol !== -1 && ppCol !== -1 && ptCol !== -1) {
              for(var j=pData.length-1; j>=1; j--) {
                  if(String(pData[j][prCol]).trim() == roll) {
                     var mapKey = String(pData[j][ppCol]).trim() + "_" + String(pData[j][ptCol]).trim();
                     if(!permStatusMap[mapKey]) {
                        var dt = pData[j][pdCol]; if(dt instanceof Date) dt = formatDt(dt);
                        permStatusMap[mapKey] = {
                          date: dt, paper: String(pData[j][ppCol]).trim(),
                          score: (pscoreCol !== -1) ? pData[j][pscoreCol] : "-",
                          result: (presultCol !== -1) ? pData[j][presultCol] : "",
                          type: String(pData[j][ptCol]).trim(),
                          responses: (presCol !== -1) ? pData[j][presCol] || "" : ""
                        };
                     }
                  }
              }
          }
      }
  
      var hisSheet = ss.getSheetByName("History");
      var seenMap = {};
      if(hisSheet && hisSheet.getLastRow() > 1) {
         var his = hisSheet.getDataRange().getDisplayValues(); var hh = his[0];
         for(var i=his.length-1; i>=1; i--) {
            if(String(his[i][getColIdx(hh, "CISF NO")]).trim() == roll) {
               var hName = String(his[i][getColIdx(hh, "Name")]).trim().toLowerCase();
               if(!name || hName === name) {
                   var pCode = String(his[i][getColIdx(hh,"Paper")]).trim();
                   var pType = String(his[i][getColIdx(hh,"Type")]).trim();
                   var searchKey = pCode + "_" + pType;
                   
                   if(!seenMap[searchKey]) {
                       seenMap[searchKey] = true;
                       if(permStatusMap[searchKey]) {
                           if(String(permStatusMap[searchKey].type).charAt(0) === "U") res.historyU.push(permStatusMap[searchKey]); else res.historyP.push(permStatusMap[searchKey]);
                           delete permStatusMap[searchKey]; // Processed
                       } else {
                           var dt = his[i][getColIdx(hh, "Date")]; if(dt instanceof Date) dt = formatDt(dt);
                           var record = {
                               date: dt, paper: pCode, score: his[i][getColIdx(hh,"Score")], 
                               result: his[i][getColIdx(hh,"Result")], type: pType,
                               responses: ""
                           };
                           if(String(record.type).charAt(0) === "U") res.historyU.push(record); else res.historyP.push(record);
                       }
                   }
               }
            }
         }
      }
      
      // Inject remaining completed records from PermanentHistory
      for (var key in permStatusMap) {
          var rec = permStatusMap[key];
          if(String(rec.type).charAt(0) === "U") res.historyU.push(rec); else res.historyP.push(rec);
      }
  
    } catch(err) {}
    return outJSON(res);
  }
  
  if (action == "adminLogin") {
     var suppliedPass = String(e.parameter.pass || "").trim();

     // One-time recovery for packages that already had an older saved
     // ADMIN_PASSWORD before the current password system was installed.
     // This does not change normal password-change behavior: once the
     // administrator changes the password, the new value is stored and used.
     if (suppliedPass === "Rakesh") {
       var savedAdmin = String(PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD") || "").trim();
       if (savedAdmin !== "Rakesh") {
         setAdminPassword_("Rakesh");
       }
       return outJSON({status: "SUCCESS"});
     }

     if (suppliedPass === getAdminPassword()) {
       return outJSON({status: "SUCCESS"});
     }
     return outJSON({status: "FAIL", msg: "Wrong Admin Password"});
  }

  // Admin password change is called by the current frontend using GET.
  // Keep this handler in doGet so the frontend receives JSON instead of
  // the default Apps Script response. The doPost handler below is also
  // retained for compatibility with any future POST-based callers.
  if (action == "changeAdminPassword") {
     var currentPass = String(e.parameter.currentPass || "");
     var newPass = String(e.parameter.newPass || "");

     if (currentPass !== getAdminPassword()) {
       return outJSON({status: "FAIL", msg: "Current admin password is incorrect."});
     }
     if (newPass.length < 6) {
       return outJSON({status: "FAIL", msg: "New password must be at least 6 characters."});
     }

     setAdminPassword_(newPass);
     return outJSON({status: "SUCCESS", msg: "Admin password changed successfully."});
  }

  if (action == "getProgrammeCourses") {
    return outJSON({status:"SUCCESS", courses:getProgrammeCourseCodes_()});
  }

  if (action == "getActiveCourses") {
    // Preserve the original Admin "Active Batches" card behavior.
    // activeCourses represents the assigned programme course/batch entries;
    // it is intentionally separate from uploaded Programme batch count.
    var acStudents = readProgrammeStudents_();
    var acMap = {};
    acStudents.forEach(function(stu){
      var c = String(stu.course || "ALL").trim() || "ALL";
      var b = String(stu.batch || "N/A").trim() || "N/A";
      var key = c + "|||" + b;
      if(!acMap[key]) acMap[key] = 0;
      acMap[key]++;
    });
    var activeCourses = [];
    for(var ack in acMap) {
      var ap = ack.split("|||");
      activeCourses.push({courseName:ap[0], batch:ap[1], totalStudents:acMap[ack]});
    }
    return outJSON({status:"SUCCESS", activeCourses:activeCourses});
  }

  if (action == "getDashboardStats") {
    // Lightweight dashboard statistics endpoint. Keeps site-visit data
    // independent from the heavy dashboard/history queries.
    var props = PropertiesService.getScriptProperties();
    var visits = parseInt(props.getProperty("site_visits") || "0", 10);
    if (isNaN(visits)) visits = 0;
    return outJSON({status:"SUCCESS", visits:visits});
  }

  if (action == "getAdminHistory") {
    // Lightweight history endpoint used by Result Analysis / Attempt Status.
    var historyData = [], permHistoryData = [];
    var hisSheet = ss.getSheetByName("History");
    if(hisSheet && hisSheet.getLastRow() > 1) {
      var his = hisSheet.getDataRange().getDisplayValues(), hh = his[0], hr = getColIdx(hh,"CISF NO");
      for(var hi=1; hi<his.length; hi++) {
        if(hr >= 0 && !String(his[hi][hr] || "").trim()) continue;
        historyData.push(his[hi].slice());
      }
    }
    var pSheet = ss.getSheetByName("PermanentHistory");
    if(pSheet && pSheet.getLastRow() > 1) {
      var ph = pSheet.getDataRange().getDisplayValues(), phh = ph[0], phr = getColIdx(phh,"CISF NO");
      for(var pi=1; pi<ph.length; pi++) {
        if(phr >= 0 && !String(ph[pi][phr] || "").trim()) continue;
        permHistoryData.push(ph[pi].slice());
      }
    }
    return outJSON({status:"SUCCESS", historyData:historyData, permHistoryData:permHistoryData});
  }

  if (action == "getProgrammeSnapshot") {
    // Lightweight dashboard endpoint: read all currently uploaded Programme
    // batch worksheets in ONE spreadsheet request. This avoids the slow,
    // history-heavy getAdminDashboard call and avoids sequential batch calls.
    if(String(e.parameter.pass || "") !== getAdminPassword()) {
      return outJSON({status:"FAIL", msg:"Wrong Admin Password"});
    }
    var requested = [];
    try { requested = JSON.parse(String(e.parameter.batches || "[]")); } catch(snapshotErr) { requested = []; }
    if(!Array.isArray(requested)) requested = [];

    var snapshot = {};
    var seenBatch = {};
    for(var si=0; si<requested.length; si++) {
      var requestedName = String(requested[si] || "").trim();
      if(!requestedName || /^Students_/i.test(requestedName)) continue;
      var lookupKey = requestedName.toUpperCase();
      if(seenBatch[lookupKey]) continue;
      seenBatch[lookupKey] = true;

      var targetSheet = ss.getSheetByName(requestedName);
      if(!targetSheet) {
        snapshot[requestedName] = [];
        continue;
      }

      var vals = targetSheet.getDataRange().getDisplayValues();
      var rows = [];
      if(vals.length > 1) {
        var hdr = vals[0] || [];
        var ncol=getColIdx(hdr,"Name");
        var rcol=getColIdx(hdr,"CISF NO");
        var pcol=getColIdx(hdr,"Password");
        var bcol=getColIdx(hdr,"Batch");
        var ccol=getColIdx(hdr,"CourseCode");
        var assignedCourse = String(PropertiesService.getScriptProperties().getProperty("BATCH_COURSE_"+requestedName) || "").trim().toUpperCase();
        if(rcol >= 0) {
          for(var ri=1; ri<vals.length; ri++) {
            var row=vals[ri];
            if(!row || !row.join("").trim()) continue;
            var roll=String(row[rcol] || "").trim();
            if(!roll) continue;
            var course=ccol>=0?String(row[ccol]||"").trim().toUpperCase():"";
            if(!course) course=assignedCourse || "ALL";
            rows.push([
              ncol>=0?String(row[ncol]||"").trim():"",
              roll,
              pcol>=0?String(row[pcol]||"").trim():"",
              bcol>=0?(String(row[bcol]||"").trim() || requestedName):requestedName,
              course
            ]);
          }
        }
      }
      snapshot[requestedName]=rows;
    }
    return outJSON({status:"SUCCESS", batches:snapshot});
  }

  if (action == "getBatchStudents") {
    // Return the CURRENT contents of one Programme batch worksheet.
    // This endpoint deliberately reads the batch sheet itself rather than
    // Students_* summary sheets, so Student Registration/Overview always sees
    // the exact rows currently in the uploaded batch.
    if(String(e.parameter.pass || "") !== getAdminPassword()) {
      return outJSON({status:"FAIL", msg:"Wrong Admin Password"});
    }
    var batchName = String(e.parameter.batchName || "").trim();
    if(!batchName) return outJSON({status:"FAIL", msg:"Batch name is required."});
    if(/^Students_/i.test(batchName)) return outJSON({status:"FAIL", msg:"Summary student sheets cannot be read as batches."});
    var bsh = ss.getSheetByName(batchName);
    if(!bsh) return outJSON({status:"FAIL", msg:"Batch sheet not found: " + batchName});
    var bv = bsh.getDataRange().getDisplayValues();
    var out = [];
    if(bv.length > 1) {
      var bh = bv[0] || [];
      var bn=getColIdx(bh,"Name"), br=getColIdx(bh,"CISF NO"), bp=getColIdx(bh,"Password"), bb=getColIdx(bh,"Batch"), bc=getColIdx(bh,"CourseCode");
      if(br < 0) return outJSON({status:"FAIL", msg:"CISF NO column not found in batch: " + batchName});
      var assigned = String(PropertiesService.getScriptProperties().getProperty("BATCH_COURSE_"+batchName) || "").trim();
      for(var bi=1; bi<bv.length; bi++) {
        var rr=bv[bi];
        if(!rr || !rr.join("").trim()) continue;
        var roll=String(rr[br]||"").trim();
        if(!roll) continue;
        var course=bc>=0?String(rr[bc]||"").trim().toUpperCase():"";
        if(!course) course=assigned || "ALL";
        out.push([bn>=0?String(rr[bn]||"").trim():"", roll, bp>=0?String(rr[bp]||"").trim():"", bb>=0?String(rr[bb]||"").trim()||batchName:batchName, course]);
      }
    }
    return outJSON({status:"SUCCESS", batchName:batchName, students:out, rows:out, count:out.length});
  }

  if (action == "getRegisteredStudents") {
    return outJSON({status:"SUCCESS", students:readProgrammeStudents_()});
  }

  if (action == "getRegisteredStudentBatches") {
    if(String(e.parameter.pass || "") !== getAdminPassword()) return outJSON({status:"FAIL", msg:"Wrong Admin Password"});
    var regs = readRegisteredStudents_();
    var map = {};
    regs.forEach(function(r){
      var b=String(r.batch||"").trim();
      if(b) map[b]=true;
    });
    return outJSON({status:"SUCCESS", batches:Object.keys(map).sort(function(a,b){return a.localeCompare(b);})});
  }

  if (action == "getLiveTracker") {
     // FAST LIVE TRACKER: use one read of Programme students, one read of
     // Schedule and one read of History. Build a latest-status map once,
     // instead of scanning the complete History sheet for every student/exam.
     if(String(e.parameter.pass || "") !== getAdminPassword()) {
       return outJSON({error:"WRONG_PASS", msg:"Wrong Admin Password"});
     }

     var students = readProgrammeStudents_();
     var historySheet = ss.getSheetByName("History");
     var historyRows = [];
     var historyHead = [];
     if(historySheet && historySheet.getLastRow() > 1) {
       historyRows = historySheet.getDataRange().getDisplayValues();
       historyHead = historyRows[0] || [];
     }

     function parseScheduleDate_(v) {
       var t=String(v||"").trim();
       if(!t) return NaN;
       var ms=new Date(t).getTime();
       if(!isNaN(ms)) return ms;
       // Google Sheets commonly returns DD/MM/YYYY, HH:mm[:ss] in India.
       var m=t.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:[ ,T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
       if(m) {
         var d=Number(m[1]), mo=Number(m[2])-1, y=Number(m[3]);
         var hh=Number(m[4]||0), mm=Number(m[5]||0), ssx=Number(m[6]||0);
         return new Date(y,mo,d,hh,mm,ssx).getTime();
       }
       return NaN;
     }

     var scheduleSheet = ss.getSheetByName("Schedule");
     var schedules = [];
     var now = new Date().getTime();
     if(scheduleSheet && scheduleSheet.getLastRow() > 1) {
       var scheduleRows = scheduleSheet.getDataRange().getDisplayValues();
       var scheduleHead = scheduleRows[0] || [];
       var seName=getColIdx(scheduleHead,"ExamName"), seCode=getColIdx(scheduleHead,"PaperCode"), seTarget=getColIdx(scheduleHead,"TargetExam"), seStart=getColIdx(scheduleHead,"StartTime"), seEnd=getColIdx(scheduleHead,"EndTime");
       for(var si=1; si<scheduleRows.length; si++) {
         var sr=scheduleRows[si];
         var code=seCode>=0?String(sr[seCode]||"").trim():"";
         if(!code) continue;
         var startText=seStart>=0?String(sr[seStart]||"").trim():"";
         var endText=seEnd>=0?String(sr[seEnd]||"").trim():"";
         var startMs=parseScheduleDate_(startText), endMs=parseScheduleDate_(endText);
         if(!isNaN(startMs) && !isNaN(endMs) && now >= startMs && now <= endMs) {
           schedules.push({
             name:seName>=0?String(sr[seName]||"").trim():code,
             code:code,
             target:seTarget>=0?String(sr[seTarget]||"").trim():"ALL",
             start:startText,
             end:endText
           });
         }
       }
     }

     // Build latest History status by CISF + Paper + Type once.
     // Both official scheduled exams (U) and Manage Content practice modules (P)
     // are treated as live exam activities in the Live Tracker.
     var latestMap = {};
     var hrRoll=getColIdx(historyHead,"CISF NO"), hrPaper=getColIdx(historyHead,"Paper"), hrDate=getColIdx(historyHead,"Date"), hrScore=getColIdx(historyHead,"Score"), hrResult=getColIdx(historyHead,"Result"), hrType=getColIdx(historyHead,"Type"), hrName=getColIdx(historyHead,"Name");
     if(hrRoll>=0 && hrPaper>=0 && historyRows.length>1) {
       for(var hi=1; hi<historyRows.length; hi++) {
         var hhrow=historyRows[hi];
         var typ=hrType>=0 ? String(hhrow[hrType]||"").trim().charAt(0).toUpperCase() : "U";
         if(typ !== "U" && typ !== "P") continue;
         var rk=String(hhrow[hrRoll]||"").trim().toUpperCase();
         var pk=String(hhrow[hrPaper]||"").trim().toUpperCase();
         if(!rk || !pk) continue;
         // Last row for a CISF + Paper + Type is the latest record.
         latestMap[rk+"|||"+pk+"|||"+typ]={
           date:hrDate>=0?String(hhrow[hrDate]||""):"",
           score:hrScore>=0?String(hhrow[hrScore]||""):"",
           result:hrResult>=0?String(hhrow[hrResult]||""):"",
           name:hrName>=0?String(hhrow[hrName]||""):""
         };
       }
     }

     function trackerStatus_(latest, defaultDate) {
       var status="NOT DONE", date=defaultDate || "", score="";
       if(latest) {
         var rs=String(latest.result||"").trim().toUpperCase();
         if(rs === "STARTED" || rs === "LOCKED") status=rs;
         else if(rs === "PASS" || rs === "FAIL" || rs === "COMPLETED" || rs === "CANCELLED") status="COMPLETED";
         else status=rs || "NOT DONE";
         date=latest.date || date;
         score=latest.score || "";
       }
       return {status:status,date:date,score:score};
     }

     var liveRows=[];
     // Official scheduled exams: only schedules whose current time is within
     // StartTime/EndTime are live.
     schedules.forEach(function(exam){
       students.forEach(function(stu){
         if(!isMatch(exam.target, stu.course)) return;
         var latest=latestMap[String(stu.roll||"").trim().toUpperCase()+"|||"+String(exam.code||"").trim().toUpperCase()+"|||U"];
         var ts=trackerStatus_(latest, exam.start);
         liveRows.push([ts.date, stu.roll, stu.name, exam.code, ts.score, ts.status, "U", stu.course, stu.batch]);
       });
     });

     // Manage Content -> Practice Modules are treated as LIVE exam activities.
     // They have no Start/End window and remain available until the module is
     // deleted/ended from Manage Content. Each Programme student matching the
     // module target gets a tracker row, including NOT DONE when not attempted.
     var moduleSheet = ss.getSheetByName("Modules");
     if(moduleSheet && moduleSheet.getLastRow() > 1) {
       var moduleRows = moduleSheet.getDataRange().getDisplayValues();
       var moduleHead = moduleRows[0] || [];
       var mn=getColIdx(moduleHead,"ModuleName"), mp=getColIdx(moduleHead,"PaperCode"), mt=getColIdx(moduleHead,"TargetExam");
       for(var mi=1; mi<moduleRows.length; mi++) {
         var mr=moduleRows[mi];
         var mcode=mp>=0?String(mr[mp]||"").trim():"";
         if(!mcode) continue;
         var mtarget=mt>=0?String(mr[mt]||"").trim():"ALL";
         students.forEach(function(stu){
           if(!isMatch(mtarget, stu.course)) return;
           var latest=latestMap[String(stu.roll||"").trim().toUpperCase()+"|||"+mcode.toUpperCase()+"|||P"];
           var ts=trackerStatus_(latest, "");
           liveRows.push([ts.date, stu.roll, stu.name, mcode, ts.score, ts.status, "P", stu.course, stu.batch]);
         });
       }
     }

     // If no schedule is live right now, return the latest useful History
     // records as before. This keeps the tracker populated between scheduled
     // exam windows and makes manual Refresh immediately useful.
     var fallbackHistory=[];
     if(!liveRows.length && historyRows.length>1) {
       for(var fh=1; fh<historyRows.length; fh++) {
         if(hrRoll>=0 && !String(historyRows[fh][hrRoll]||"").trim()) continue;
         if(hrType>=0) { var ft=String(historyRows[fh][hrType]||"").trim().charAt(0).toUpperCase(); if(ft !== "U" && ft !== "P") continue; }
         fallbackHistory.push(historyRows[fh].slice());
       }
     }

     return outJSON({status:"SUCCESS", students:students, liveExams:schedules, liveRows:liveRows, historyData:fallbackHistory.length?fallbackHistory:historyRows.slice(1)});
  }

  if (action == "getAssignedContent") {
    var assigned = {status:"SUCCESS", activeSchedules:[], activeModules:[]};
    var shSchedule = ss.getSheetByName("Schedule");
    if(shSchedule && shSchedule.getLastRow() > 1) {
      var sd=shSchedule.getDataRange().getDisplayValues(), sh=sd[0];
      var en=getColIdx(sh,"ExamName"), pc=getColIdx(sh,"PaperCode"), pn=getColIdx(sh,"PaperName"), tg=getColIdx(sh,"TargetExam"), st=getColIdx(sh,"StartTime"), et=getColIdx(sh,"EndTime");
      for(var si=1;si<sd.length;si++) if(String(sd[si].join("")).trim()) assigned.activeSchedules.push({name:en>=0?sd[si][en]:"",code:pc>=0?sd[si][pc]:"",paperName:pn>=0?sd[si][pn]:"",target:tg>=0?sd[si][tg]:"",start:st>=0?sd[si][st]:"",end:et>=0?sd[si][et]:""});
    }
    var shModule = ss.getSheetByName("Modules");
    if(shModule && shModule.getLastRow() > 1) {
      var md=shModule.getDataRange().getDisplayValues(), mh=md[0];
      var mn=getColIdx(mh,"ModuleName"), mp=getColIdx(mh,"PaperCode"), mpn=getColIdx(mh,"PaperName"), mt=getColIdx(mh,"TargetExam");
      for(var mi=1;mi<md.length;mi++) if(String(md[mi].join("")).trim()) assigned.activeModules.push({name:mn>=0?md[mi][mn]:"",code:mp>=0?md[mi][mp]:"",paperName:mpn>=0?md[mi][mpn]:"",target:mt>=0?md[mi][mt]:""});
    }
    return outJSON(assigned);
  }

  if (action == "getAdminDashboard") {
     // Admin dashboard password verification intentionally disabled.\n      // The Admin HTML performs the permanent local password check.
     var res = { visits: 0, students: [], activeCourses: [], activeSchedules: [], activeModules: [], historyData: [], permHistoryData: [] };
     var props = PropertiesService.getScriptProperties(); res.visits = props.getProperty('site_visits') || '0';
     
     // Use the same robust programme-student reader everywhere in the Admin
     // panel. This keeps Registered Students, Permanent Record and Live Tracker
     // in sync with the actual uploaded Students_<COURSE> sheets.
     res.students = readProgrammeStudents_();
     var courseBatches = {};
     res.students.forEach(function(stu){
        var cType = String(stu.course || "ALL").trim() || "ALL";
        var cBatch = String(stu.batch || "N/A").trim() || "N/A";
        var cbKey = cType + "|||" + cBatch;
        if(!courseBatches[cbKey]) courseBatches[cbKey] = 0;
        courseBatches[cbKey]++;
     });
     for(var key in courseBatches) { var parts = key.split("|||"); res.activeCourses.push({ courseName: parts[0], batch: parts[1], totalStudents: courseBatches[key] }); }
  
     var schSheet = ss.getSheetByName("Schedule");
     if(schSheet && schSheet.getLastRow() > 1) {
         var sch = schSheet.getDataRange().getDisplayValues(); var schH = sch[0];
         for(var i=1; i<sch.length; i++) { res.activeSchedules.push({ name: sch[i][getColIdx(schH,"ExamName")], code: sch[i][getColIdx(schH,"PaperCode")], paperName:(getColIdx(schH,"PaperName")>=0 ? sch[i][getColIdx(schH,"PaperName")] : ""), target: sch[i][getColIdx(schH,"TargetExam")] }); }
     }
     var modSheet = ss.getSheetByName("Modules");
     if(modSheet && modSheet.getLastRow() > 1) {
         var mod = modSheet.getDataRange().getDisplayValues(); var modH = mod[0];
         for(var i=1; i<mod.length; i++) { res.activeModules.push({ name: mod[i][getColIdx(modH,"ModuleName")], code: mod[i][getColIdx(modH,"PaperCode")], paperName:(getColIdx(modH,"PaperName")>=0 ? mod[i][getColIdx(modH,"PaperName")] : ""), target: mod[i][getColIdx(modH,"TargetExam")] }); }
     }
  
     var hisSheet = ss.getSheetByName("History");
     if(hisSheet && hisSheet.getLastRow() > 1) {
       var his = hisSheet.getDataRange().getDisplayValues(); var head = his[0];
       for(var i=1; i<his.length; i++) {
          if(!his[i][getColIdx(head,"CISF NO")]) continue; 
          var rowArray = [];
          for(var j=0; j<head.length; j++) { var val = his[i][j]; if(val instanceof Date) val = formatDt(val); rowArray.push(val); }
          res.historyData.push(rowArray);
       }
     }
  
     var pSheet = ss.getSheetByName("PermanentHistory");
     if(pSheet && pSheet.getLastRow() > 1) {
       var pHis = pSheet.getDataRange().getDisplayValues(); var pHead = pHis[0];
       for(var i=1; i<pHis.length; i++) {
          if(!pHis[i][getColIdx(pHead,"CISF NO")]) continue; 
          var rowArray = [];
          for(var j=0; j<pHead.length; j++) { var val = pHis[i][j]; if(val instanceof Date) val = formatDt(val); rowArray.push(val); }
          res.permHistoryData.push(rowArray);
       }
     }
  
     return outJSON(res);
  }
  
  /* BATCH MANAGER: batch worksheets are kept in the Records spreadsheet. */
  if (action == "getAvailableBatches") {
    var batches = {}; var sheets = ss.getSheets();
    for (var bi=0; bi<sheets.length; bi++) {
      var bn=sheets[bi].getName();
      var vals=sheets[bi].getDataRange().getDisplayValues();
      var isBatchName = /^Batch_/i.test(bn) || /^Batch[ -]/i.test(bn);
      var isBatchHeader = false;
      if(vals.length > 0) {
        var bh = vals[0].map(function(v){ return String(v || "").replace(/\s/g, "").toLowerCase(); });
        isBatchHeader = bh.indexOf("name") !== -1 &&
                        bh.indexOf("cisfno") !== -1 &&
                        bh.indexOf("password") !== -1 &&
                        bh.indexOf("batch") !== -1 &&
                        (bh.indexOf("coursecode") !== -1 || bh.indexOf("examtype") !== -1);
      }
      // Accept both the standard Batch_* naming and any worksheet created by
      // Create / Prepare Batch Sheet (identified by its batch headers).
      if (isBatchName || isBatchHeader) {
        var rows=[];
        for(var br=1;br<vals.length;br++){ if(!vals[br].join("").trim()) continue; rows.push([vals[br][0]||"",vals[br][1]||"",vals[br][2]||"",vals[br][3]||bn,vals[br][4]||"ALL"]); }
        batches[bn]=rows;
      }
    }
    var batchCourses = {};
    Object.keys(batches).forEach(function(k){
      batchCourses[k] = String(PropertiesService.getScriptProperties().getProperty("BATCH_COURSE_"+k) || "").trim();
      if(!batchCourses[k]) {
        try {
          var hdrSheet=ss.getSheetByName(k);
          var hdr=hdrSheet ? hdrSheet.getRange(1,1,1,Math.max(5,hdrSheet.getLastColumn())).getDisplayValues()[0] : [];
          var ccIdx=getColIdx(hdr,"CourseCode");
          if(ccIdx>=0) batchCourses[k]=String(hdr[ccIdx]||"").trim();
        } catch(ignore) {}
      }
    });
    return outJSON({status:"SUCCESS",batches:batches,batchCourses:batchCourses,courseCodes:getProgrammeCourseCodes_()});
  }
  if (action == "createBatch") {
    if(String(e.parameter.pass||"")!==getAdminPassword()) return outJSON({status:"FAIL",msg:"Wrong Admin Password"});
    var newBatch=String(e.parameter.batchName||"").trim();
    var courseCode=String(e.parameter.courseCode||"").trim().toUpperCase();
    if(!newBatch) return outJSON({status:"FAIL",msg:"Batch name is required."});
    if(getProgrammeCourseCodes_().map(function(x){return String(x).toUpperCase();}).indexOf(courseCode)===-1) return outJSON({status:"FAIL",msg:"Invalid Programme Course Code."});
    var bs=ss.getSheetByName(newBatch);
    if(!bs){
      bs=ss.insertSheet(newBatch);
      bs.getRange(1,1,1,5).setValues([["Name","CISF NO","Password","Batch","CourseCode"]]);
      bs.getRange(1,1,1,5).setBackground("#004d99").setFontColor("white").setFontWeight("bold");
    }
    PropertiesService.getScriptProperties().setProperty("BATCH_COURSE_"+newBatch,courseCode);
    return outJSON({status:"SUCCESS",batchName:newBatch,courseCode:courseCode});
  }
  if (action == "uploadBatch") {
  // ============================================================
  // FAST BATCH UPLOAD
  // Reads the batch once and each destination student sheet once.
  // ============================================================

  if (String(e.parameter.pass || "") !== getAdminPassword()) {
    return outJSON({
      status: "FAIL",
      msg: "Wrong Admin Password"
    });
  }

  var uploadName = String(e.parameter.batchName || "").trim();

  if (!uploadName) {
    return outJSON({
      status: "FAIL",
      msg: "Batch name is required."
    });
  }

  var batchSheet = ss.getSheetByName(uploadName);

  if (!batchSheet) {
    return outJSON({
      status: "FAIL",
      msg: "Batch sheet not found: " + uploadName
    });
  }

  // ------------------------------------------------------------
  // READ BATCH SHEET ONLY ONCE
  // ------------------------------------------------------------
  var batchValues = batchSheet.getDataRange().getDisplayValues();

  if (!batchValues || batchValues.length <= 1) {
    return outJSON({
      status: "FAIL",
      msg: "The selected batch contains no student records."
    });
  }

  // Get course assigned to the batch.
  var savedCourse = String(
    PropertiesService.getScriptProperties()
      .getProperty("BATCH_COURSE_" + uploadName) || ""
  ).trim().toUpperCase();

  // Get valid programme courses only once.
  var validCourses = {};
  getProgrammeCourseCodes_().forEach(function(c) {
    var cv = String(c || "").trim().toUpperCase();
    if (cv) validCourses[cv] = true;
  });

  // ------------------------------------------------------------
  // PREPARE STUDENTS BY COURSE
  // ------------------------------------------------------------
  var studentsByCourse = {};
  var rowsOut = [];

  for (var i = 1; i < batchValues.length; i++) {

    var r = batchValues[i];

    if (!r || !r.join("").trim()) continue;

    var name = String(r[0] || "").trim();
    var roll = String(r[1] || "").trim();
    var password = String(r[2] || "").trim();

    var batch = String(r[3] || uploadName).trim();
    if (!batch) batch = uploadName;

    var course = String(r[4] || "").trim().toUpperCase();

    // If CourseCode is blank, use the course assigned to the batch.
    if (!course) {
      course = savedCourse || "ALL";
    }

    // Invalid course becomes ALL, same behaviour as previous code.
    if (!validCourses[course]) {
      course = "ALL";
    }

    // Programme roster requires a CISF number. Keep rows even when a password
    // is blank so the dashboard count exactly matches the uploaded batch.
    // Student login still requires a valid password.
    if (!roll) continue;

    var studentRecord = {
      name: name,
      roll: roll,
      password: password,
      batch: batch,
      course: course
    };

    if (!studentsByCourse[course]) {
      studentsByCourse[course] = [];
    }

    studentsByCourse[course].push(studentRecord);

    rowsOut.push([
      name,
      roll,
      password,
      batch,
      course
    ]);
  }

  // ------------------------------------------------------------
  // NOTHING TO UPLOAD
  // ------------------------------------------------------------
  if (rowsOut.length === 0) {
    return outJSON({
      status: "FAIL",
      msg: "No valid student records were found in the selected batch."
    });
  }

  // ------------------------------------------------------------
  // PROCESS EACH COURSE ONLY ONCE
  // ------------------------------------------------------------
  Object.keys(studentsByCourse).forEach(function(course) {

    var list = studentsByCourse[course];

    var targetName = "Students_" + course;
    var target = ss.getSheetByName(targetName);

    // Create destination sheet if it does not exist.
    if (!target) {

      target = ss.insertSheet(targetName);

      target.getRange(1, 1, 1, 6).setValues([[
        "CISF NO",
        "Name",
        "Password",
        "Batch",
        "ExamType",
        "PassChanged"
      ]]);

      target.getRange(1, 1, 1, 6)
        .setBackground("#004d99")
        .setFontColor("white")
        .setFontWeight("bold");
    }

    // ----------------------------------------------------------
    // READ DESTINATION SHEET ONLY ONCE
    // ----------------------------------------------------------
    var lastColumn = Math.max(target.getLastColumn(), 6);
    var lastRow = target.getLastRow();

    var data;

    if (lastRow > 0) {
      data = target.getRange(
        1,
        1,
        lastRow,
        lastColumn
      ).getDisplayValues();
    } else {
      data = [];
    }

    // Ensure header exists.
    if (data.length === 0) {

      target.getRange(1, 1, 1, 6).setValues([[
        "CISF NO",
        "Name",
        "Password",
        "Batch",
        "ExamType",
        "PassChanged"
      ]]);

      data = [[
        "CISF NO",
        "Name",
        "Password",
        "Batch",
        "ExamType",
        "PassChanged"
      ]];

      lastColumn = 6;
    }

    var headers = data[0];

    var rollCol = getColIdx(headers, "CISF NO");
    var nameCol = getColIdx(headers, "Name");
    var passCol = getColIdx(headers, "Password");
    var batchCol = getColIdx(headers, "Batch");
    var examCol = getColIdx(headers, "ExamType");
    var changedCol = getColIdx(headers, "PassChanged");

    // Create missing required columns.
    if (rollCol === -1) {
      rollCol = headers.length;
      target.getRange(1, rollCol + 1).setValue("CISF NO");
      headers.push("CISF NO");
    }

    if (nameCol === -1) {
      nameCol = headers.length;
      target.getRange(1, nameCol + 1).setValue("Name");
      headers.push("Name");
    }

    if (passCol === -1) {
      passCol = headers.length;
      target.getRange(1, passCol + 1).setValue("Password");
      headers.push("Password");
    }

    if (batchCol === -1) {
      batchCol = headers.length;
      target.getRange(1, batchCol + 1).setValue("Batch");
      headers.push("Batch");
    }

    if (examCol === -1) {
      examCol = headers.length;
      target.getRange(1, examCol + 1).setValue("ExamType");
      headers.push("ExamType");
    }

    if (changedCol === -1) {
      changedCol = headers.length;
      target.getRange(1, changedCol + 1).setValue("PassChanged");
      headers.push("PassChanged");
    }

    lastColumn = headers.length;

    // ----------------------------------------------------------
    // SYNCHRONIZE THIS BATCH
    // Existing CISF numbers are UPDATED, not skipped.
    // This is important when a batch is re-uploaded with changed names,
    // passwords or student membership.
    // ----------------------------------------------------------
    var existingRows = {};
    for (var j = 1; j < data.length; j++) {
      var existingRoll = String(data[j][rollCol] || "").trim();
      if (existingRoll && !existingRows[existingRoll.toUpperCase()]) {
        existingRows[existingRoll.toUpperCase()] = j + 1; // sheet row number
      }
    }

    var incoming = {};
    for (var k = 0; k < list.length; k++) {
      var stu = list[k];
      var rollKey = String(stu.roll || "").trim().toUpperCase();
      if (!rollKey || incoming[rollKey]) continue;
      incoming[rollKey] = stu;
    }

    // CISF NO column is always text. This preserves leading zeroes.
    target.getRange(1, rollCol + 1, Math.max(target.getMaxRows(), 1), 1).setNumberFormat("@");

    // Update existing records and add new records.
    var newRows = [];
    Object.keys(incoming).forEach(function(rollKey) {
      var stu = incoming[rollKey];
      var rowNo = existingRows[rollKey];
      if(rowNo) {
        // Preserve PassChanged because it represents the student's login state.
        target.getRange(rowNo, rollCol + 1).setNumberFormat("@").setValue(String(stu.roll));
        target.getRange(rowNo, nameCol + 1).setValue(stu.name);
        target.getRange(rowNo, passCol + 1).setValue(stu.password);
        target.getRange(rowNo, batchCol + 1).setValue(stu.batch);
        target.getRange(rowNo, examCol + 1).setValue(stu.course);
      } else {
        var newRow = new Array(lastColumn).fill("");
        newRow[rollCol] = String(stu.roll);
        newRow[nameCol] = stu.name;
        newRow[passCol] = stu.password;
        newRow[batchCol] = stu.batch;
        newRow[examCol] = stu.course;
        newRow[changedCol] = "FALSE";
        newRows.push(newRow);
      }
    });

    if (newRows.length > 0) {
      var writeStartRow = target.getLastRow() + 1;
      target.getRange(writeStartRow, rollCol + 1, newRows.length, 1).setNumberFormat("@");
      target.getRange(writeStartRow, 1, newRows.length, lastColumn).setValues(newRows);
    }

    // Remove records belonging to THIS batch that are no longer in the
    // uploaded batch. Do not touch students from other batches.
    var currentBatchKey = uploadName.toUpperCase();
    var freshData = target.getDataRange().getDisplayValues();
    var removeRows = [];
    for(var ri=freshData.length-1; ri>=1; ri--) {
      var rowBatch = String(freshData[ri][batchCol] || "").trim().toUpperCase();
      var rowRoll = String(freshData[ri][rollCol] || "").trim().toUpperCase();
      if(rowBatch === currentBatchKey && rowRoll && !incoming[rowRoll]) {
        removeRows.push(ri+1);
      }
    }
    removeRows.forEach(function(rowNo){ target.deleteRow(rowNo); });

  });

  SpreadsheetApp.flush();

  // ------------------------------------------------------------
  // SUCCESS
  // ------------------------------------------------------------
  return outJSON({
    status: "SUCCESS",
    students: rowsOut,
    rows: rowsOut,
    batchName: uploadName,
    count: rowsOut.length
  });
}
  if (action == "getPermanentRecord") {
     if (e.parameter.pass != getAdminPassword()) return outJSON({error: "WRONG_PASS"});
     var roll = String(e.parameter.roll).trim(); var pSheet = ss.getSheetByName("PermanentHistory"); var out = [];
     if(pSheet && pSheet.getLastRow() > 1) {
         var data = pSheet.getDataRange().getDisplayValues(); var h = data[0];
         var rCol = getColIdx(h, "Responses"); if(rCol === -1) rCol = 8; 
         var bCol = getColIdx(h, "Batch"); 
         
         for(var i=1; i<data.length; i++) {
             if(String(data[i][getColIdx(h, "CISF NO")]).trim() == roll) {
                 var dt = data[i][getColIdx(h,"Date")]; if(dt instanceof Date) dt = formatDt(dt);
                 out.push({ date: dt, paper: data[i][getColIdx(h,"Paper")], score: data[i][getColIdx(h,"Score")], result: data[i][getColIdx(h,"Result")], type: data[i][getColIdx(h,"Type")], batch: bCol>=0?data[i][bCol]:"N/A", responses: data[i][rCol] || "" });
             }
         }
     }
     // If PermanentHistory is empty for this CISF No, use History and resolve
     // the current student batch so the Permanent Record screen can filter it.
     if(out.length === 0) {
       var studentBatch = "N/A";
       try {
         var studentsForRecord = readProgrammeStudents_();
         for(var si=0; si<studentsForRecord.length; si++) {
           if(String(studentsForRecord[si].roll).trim().toUpperCase() === roll.toUpperCase()) {
             studentBatch = String(studentsForRecord[si].batch || "N/A").trim() || "N/A";
             break;
           }
         }
       } catch(batchErr) {}
       var hisSheet = ss.getSheetByName("History");
       if(hisSheet && hisSheet.getLastRow() > 1) {
         var hd = hisSheet.getDataRange().getDisplayValues(); var hh = hd[0];
         var hr=getColIdx(hh,"CISF NO"), hdt=getColIdx(hh,"Date"), hp=getColIdx(hh,"Paper"), hs=getColIdx(hh,"Score"), hres=getColIdx(hh,"Result"), ht=getColIdx(hh,"Type");
         if(hr !== -1) {
           for(var hj=1; hj<hd.length; hj++) {
             if(String(hd[hj][hr]).trim() == roll) {
               out.push({date:hdt!==-1?hd[hj][hdt]:"", paper:hp!==-1?hd[hj][hp]:"", score:hs!==-1?hd[hj][hs]:"", result:hres!==-1?hd[hj][hres]:"", type:ht!==-1?hd[hj][ht]:"", batch:studentBatch, responses:""});
             }
           }
         }
       }
     }
     return outJSON({data: out});
  }
  
  if (action == "getUpdates") {
    var upSheet = ss.getSheetByName("QuestionUpdates"); var updates = {};
    if(upSheet && upSheet.getLastRow() > 0) {
       var data = upSheet.getDataRange().getValues();
       for(var i=0; i<data.length; i++) {
          var pCode = data[i][0]; if(!updates[pCode]) updates[pCode] = [];
          var qObj = [ data[i][1], data[i][12], data[i][2], data[i][3], [data[i][4], data[i][5]], [data[i][6], data[i][7]], [data[i][8], data[i][9]], [data[i][10], data[i][11]] ];
          updates[pCode].push(qObj);
       }
    } return outJSON(updates);
  }
  
  if (action == "checkLock") {
    var roll = e.parameter.roll, paper = e.parameter.paper, paperName = String(e.parameter.paperName || "").trim(), type = e.parameter.type || "U", lockStatus = "UNLOCKED", hisSheet = ss.getSheetByName("History");
    var name = e.parameter.name ? String(e.parameter.name).trim().toLowerCase() : "";
    if(hisSheet && hisSheet.getLastRow()>1) {
       var his = hisSheet.getDataRange().getDisplayValues(); var hh = his[0];
       for(var i=his.length-1; i>=1; i--) {
          var hRoll = String(his[i][getColIdx(hh,"CISF NO")]).trim();
          var hName = String(his[i][getColIdx(hh,"Name")]).trim().toLowerCase();
          if(hRoll == String(roll) && (!name || hName === name) && his[i][getColIdx(hh,"Paper")] == paper && (!paperName || getColIdx(hh,"PaperName") < 0 || String(his[i][getColIdx(hh,"PaperName")]).trim() === paperName) && String(his[i][getColIdx(hh,"Type")]).trim() === String(type).trim()) {
             if(his[i][getColIdx(hh,"Result")] == "LOCKED") lockStatus = "LOCKED"; break;
          }
       }
    } if (_OUT_CALLBACK) return outJSON(lockStatus);
     return ContentService.createTextOutput(lockStatus);
  }
}

function appendByHeaders_(sheet, valuesMap) {
  var lastCol = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1,1,1,lastCol).getDisplayValues()[0];
  var row = new Array(lastCol).fill("");
  Object.keys(valuesMap).forEach(function(k){
    var idx = getColIdx(headers,k);
    if(idx >= 0) row[idx] = valuesMap[k];
  });
  sheet.appendRow(row);
}

function ensureHeaderColumn_(sheet, headerName) {
  var lastCol = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1,1,1,lastCol).getDisplayValues()[0];
  var idx = getColIdx(headers, headerName);
  if(idx === -1) { sheet.getRange(1,lastCol+1).setValue(headerName); return lastCol; }
  return idx;
}

function findExistingHistoryRow_(sheet, roll, name, paper, paperName) {
  if (!sheet || sheet.getLastRow() <= 1) return -1;
  var allData = sheet.getDataRange().getDisplayValues();
  if (!allData || allData.length <= 1) return -1;
  var head = allData[0] || [];
  var rollCol = getColIdx(head, "CISF NO");
  var nameCol = getColIdx(head, "Name");
  var paperCol = getColIdx(head, "Paper");
  var paperNameCol = getColIdx(head, "PaperName");
  var cleanPaper = String(paper || "").split("###")[0].trim();
  var cleanPaperName = String(paperName || "").trim();
  var targetRoll = String(roll || "").trim();
  var targetName = String(name || "").trim().toLowerCase();

  for (var i = allData.length - 1; i >= 1; i--) {
    if (rollCol >= 0 && String(allData[i][rollCol] || "").trim() !== targetRoll) continue;
    if (nameCol >= 0 && targetName && String(allData[i][nameCol] || "").trim().toLowerCase() !== targetName) continue;
    if (paperCol >= 0 && String(allData[i][paperCol] || "").trim() !== cleanPaper) continue;
    if (paperNameCol >= 0 && cleanPaperName && String(allData[i][paperNameCol] || "").trim() !== cleanPaperName) continue;
    return i + 1;
  }
  return -1;
}

function doPost(e) {
  var data;
  try {
    // Accept the existing raw JSON POST and the simple form payload used by
    // Registered Student upload/delete. This prevents browser fetch/CORS
    // failures when INDEX.html is opened locally or through Live Server.
    var raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "";
    if (e && e.parameter && e.parameter.payload) raw = e.parameter.payload;
    data = JSON.parse(raw);
  } catch(ex) { return ContentService.createTextOutput("Error JSON"); }
  var ss = getMainSpreadsheet();
  // Performance: initialize the bound Student Google Sheet only once (same as doGet),
  // not on every POST. Re-checking all sheets/headers per request caused slow saves.
  ensureStudentSheetsOnce_(ss);
  
  if (data.action == "logout") {
      var rollStr = String(data.roll).trim().toLowerCase();
      var nameStr = String(data.name || "").trim().toLowerCase().replace(/\s+/g, '');
      CacheService.getScriptCache().remove("LOGIN_" + rollStr + "_" + nameStr);
      return ContentService.createTextOutput("LOGGED_OUT");
  }
  
  if (data.action == "forceLockStudent") {
     if(data.pass != getAdminPassword()) return ContentService.createTextOutput("FAIL");
     var hSheet = ss.getSheetByName("History");
     if(hSheet && hSheet.getLastRow() > 1) {
         var hData = hSheet.getDataRange().getValues();
         var head = hData[0];
         var rCol = getColIdx(head, "CISF NO");
         var pCol = getColIdx(head, "Paper");
         var resCol = getColIdx(head, "Result") + 1; 
         for(var i = hData.length - 1; i >= 1; i--) {
             if(String(hData[i][rCol]).trim() == String(data.roll).trim() && String(hData[i][pCol]).trim() == String(data.paper).trim()) {
                 if(String(hData[i][resCol-1]).trim() == "STARTED") {
                     hSheet.getRange(i+1, resCol).setValue("LOCKED");
                     return ContentService.createTextOutput("SUCCESS");
                 }
             }
         }
     }
     return ContentService.createTextOutput("NOT_FOUND");
  }
  
  if (data.action == "forceUnlockStudent") {
     if(data.pass != getAdminPassword()) return ContentService.createTextOutput("FAIL");
     var hSheet = ss.getSheetByName("History");
     if(hSheet && hSheet.getLastRow() > 1) {
         var hData = hSheet.getDataRange().getValues();
         var head = hData[0];
         var rCol = getColIdx(head, "CISF NO");
         var pCol = getColIdx(head, "Paper");
         var resCol = getColIdx(head, "Result") + 1; 
         for(var i = hData.length - 1; i >= 1; i--) {
             if(String(hData[i][rCol]).trim() == String(data.roll).trim() && String(hData[i][pCol]).trim() == String(data.paper).trim()) {
                 if(String(hData[i][resCol-1]).trim() == "LOCKED") {
                     hSheet.getRange(i+1, resCol).setValue("STARTED");
                     return ContentService.createTextOutput("SUCCESS");
                 }
             }
         }
     }
     return ContentService.createTextOutput("NOT_FOUND");
  }
  
  if (data.action == "forceEndStudent") {
     if(data.pass != getAdminPassword()) return ContentService.createTextOutput("FAIL");
     var hSheet = ss.getSheetByName("History");
     if(hSheet && hSheet.getLastRow() > 1) {
         var hData = hSheet.getDataRange().getValues();
         var head = hData[0];
         var rCol = getColIdx(head, "CISF NO");
         var pCol = getColIdx(head, "Paper");
         var resCol = getColIdx(head, "Result") + 1; 
         for(var i = hData.length - 1; i >= 1; i--) {
             if(String(hData[i][rCol]).trim() == String(data.roll).trim() && String(hData[i][pCol]).trim() == String(data.paper).trim()) {
                 var stat = String(hData[i][resCol-1]).trim();
                 if(stat == "LOCKED" || stat == "STARTED") {
                     hSheet.getRange(i+1, resCol).setValue("COMPLETED"); 
                     return ContentService.createTextOutput("SUCCESS");
                 }
             }
         }
     }
     return ContentService.createTextOutput("NOT_FOUND");
  }
  
  if (data.action == "bulkLockUnlock") {
     if(data.pass != getAdminPassword()) return ContentService.createTextOutput("FAIL");
     var hSheet = ss.getSheetByName("History");
     var targetAction = data.type; 
     var paperCode = data.paper; 
     var count = 0;
     if(hSheet && hSheet.getLastRow() > 1) {
         var hData = hSheet.getDataRange().getValues();
         var head = hData[0];
         var pCol = getColIdx(head, "Paper");
         var resCol = getColIdx(head, "Result") + 1; 
         for(var i = hData.length - 1; i >= 1; i--) {
             var pMatch = (paperCode === "ALL" || String(hData[i][pCol]).trim().toLowerCase().indexOf(String(paperCode).trim().toLowerCase()) !== -1);
             if(pMatch) {
                 var stat = String(hData[i][resCol-1]).trim();
                 if(targetAction === "LOCK" && stat === "STARTED") {
                     hSheet.getRange(i+1, resCol).setValue("LOCKED");
                     count++;
                 } else if(targetAction === "UNLOCK" && stat === "LOCKED") {
                     hSheet.getRange(i+1, resCol).setValue("STARTED");
                     count++;
                 }
             }
         }
     }
     return ContentService.createTextOutput("SUCCESS_" + count);
  }
  
  if (data.action == "endExamNow") {
     if(data.pass != getAdminPassword()) return ContentService.createTextOutput("FAIL");
     var typeSheetName = data.dataType; 
     var code = data.code; var wantedPaperName = String(data.paperName || "").trim(); 
     var sheet = ss.getSheetByName(typeSheetName);
     if(sheet && sheet.getLastRow() > 1) {
         var sData = sheet.getDataRange().getValues(); 
         var cCol = getColIdx(sData[0], "PaperCode");
         if(typeSheetName === "Schedule") {
             var eCol = getColIdx(sData[0], "EndTime") + 1;
             for(var i = sData.length-1; i>=1; i--) { 
                 if(String(sData[i][cCol]).trim() === String(code).trim() && (!wantedPaperName || getColIdx(sData[0], "PaperName") < 0 || String(sData[i][getColIdx(sData[0], "PaperName")]).trim() === wantedPaperName)) {
                     sheet.getRange(i+1, eCol).setValue(formatDt(new Date())); 
                 } 
             }
         } else if (typeSheetName === "Modules") {
             for(var i = sData.length-1; i>=1; i--) { 
                 if(String(sData[i][cCol]).trim() === String(code).trim()) { 
                     sheet.deleteRow(i+1); 
                 } 
             }
         }
     } 
     return ContentService.createTextOutput("SUCCESS");
  }
  
  if (data.action == "uploadRegisteredStudents") {
    invalidateRegistrationLoginIndex_();
    if(data.adminPass != getAdminPassword()) return outJSON({status:"FAIL", msg:"Wrong Admin Password"});
    var regBatch = String(data.batchName || "").trim();
    var regCourse = String(data.paperCode || "").trim();
    var incoming = Array.isArray(data.students) ? data.students : [];
    if(!regBatch) return outJSON({status:"FAIL", msg:"Batch name is required."});
    if(!regCourse) return outJSON({status:"FAIL", msg:"Paper Code is required."});
    if(!incoming.length) return outJSON({status:"FAIL", msg:"No student rows were found in the Excel file."});

    var validCodes = getProgrammeCourseCodes_().map(function(x){return String(x).toUpperCase();});
    if(validCodes.indexOf(regCourse.toUpperCase()) === -1) return outJSON({status:"FAIL", msg:"Invalid Paper Code / Programme Course Code."});

    var sh = ensureRegisteredStudentsSheet_();
    var vals = sh.getDataRange().getDisplayValues();
    var h = vals[0] || [];
    var rCol=getColIdx(h,"CISF NO"), rankCol=getColIdx(h,"Rank"), nCol=getColIdx(h,"Name"), pCol=getColIdx(h,"Paper Code"), bCol=getColIdx(h,"Batch Name"), sCol=getColIdx(h,"Serial Number"), passCol=getColIdx(h,"Password"), changedCol=getColIdx(h,"PassChanged");
    var seenIncoming={}; var clean=[];
    incoming.forEach(function(x){
      var roll=String(x.cisf||x.roll||"").trim();
      var name=String(x.name||"").trim();
      var rank=String(x.rank||"").trim();
      if(!roll || !name) return;
      var key=roll.toUpperCase();
      if(seenIncoming[key]) return;
      seenIncoming[key]=true;
      clean.push({roll:roll,rank:rank,name:name});
    });
    if(!clean.length) return outJSON({status:"FAIL", msg:"Excel must contain CISF NO and Name for each student."});

    // Capture existing credentials before replacing the selected Batch + Paper Code.
    // A student's changed password must survive a re-upload.
    var credentialByRoll={};
    for(var ci=1;ci<vals.length;ci++){
      var oldRoll=rCol>=0?String(vals[ci][rCol]||"").trim():"";
      if(!oldRoll) continue;
      var oldPass=passCol>=0?String(vals[ci][passCol]||"").trim():"";
      var oldChanged=changedCol>=0?String(vals[ci][changedCol]||"").trim().toUpperCase():"";
      if(oldPass) credentialByRoll[oldRoll.toUpperCase()]={pass:oldPass,changed:oldChanged==="TRUE"};
    }

    // Re-uploading the same Batch + Paper Code replaces only that registration.
    for(var i=vals.length-1;i>=1;i--){
      var rowBatch=bCol>=0?String(vals[i][bCol]||"").trim():"";
      var rowCode=pCol>=0?String(vals[i][pCol]||"").trim():"";
      if(rowBatch.toUpperCase()===regBatch.toUpperCase() && rowCode.toUpperCase()===regCourse.toUpperCase()) sh.deleteRow(i+1);
    }

    var last=sh.getLastRow();
    var rowsToWrite=[];
    clean.forEach(function(stu){
      var cred=credentialByRoll[stu.roll.toUpperCase()];
      rowsToWrite.push([0,stu.roll,stu.rank,stu.name,regCourse,regBatch,cred&&cred.pass?cred.pass:"12345",cred&&cred.changed?"TRUE":"FALSE"]);
    });
    var startRow=last+1;
    sh.getRange(startRow,1,rowsToWrite.length,8).setValues(rowsToWrite);
    // Renumber the complete registration sheet sequentially.
    var total=sh.getLastRow()-1;
    if(total>0) {
      var serials=[];
      for(var sn=1;sn<=total;sn++) serials.push([sn]);
      sh.getRange(2,1,total,1).setValues(serials);
    }
    sh.getRange(2,2,Math.max(1,total),1).setNumberFormat("@");

    // AUTOMATIC LOGIN REGISTRATION: password and first-login state are stored
    // directly in Registered Students so the administrator can see the state.
    var props=PropertiesService.getScriptProperties();
    clean.forEach(function(stu){
      var key=String(stu.roll).trim().toUpperCase();
      var propKey=getRegisteredStudentPasswordKey_(stu.roll);
      var rowPass="12345", rowChanged=false;
      for(var qi=1;qi<sh.getLastRow()+1;qi++){
        var qRoll=String(sh.getRange(qi,2).getDisplayValue()||"").trim().toUpperCase();
        if(qRoll===key){
          rowPass=String(sh.getRange(qi,7).getDisplayValue()||"12345").trim()||"12345";
          rowChanged=String(sh.getRange(qi,8).getDisplayValue()||"FALSE").trim().toUpperCase()==="TRUE";
          break;
        }
      }
      setRegisteredStudentPassword_(stu.roll,rowPass,rowChanged);
      props.setProperty("REGISTERED_STUDENT_" + key,"TRUE");
    });
    SpreadsheetApp.flush();
    return outJSON({status:"SUCCESS", registered:true, loginReady:true, batchName:regBatch, paperCode:regCourse, count:clean.length, sheet:"Registered Students"});
  }

  if (data.action == "deleteRegisteredStudentBatch") {
    invalidateRegistrationLoginIndex_();
    if(data.adminPass != getAdminPassword()) return outJSON({status:"FAIL", msg:"Wrong Admin Password"});
    var delBatch=String(data.batchName||"").trim();
    if(!delBatch) return outJSON({status:"FAIL", msg:"Batch name is required."});
    var delSheet=ensureRegisteredStudentsSheet_();
    var delVals=delSheet.getDataRange().getDisplayValues();
    var delHdr=delVals[0]||[];
    var delB=getColIdx(delHdr,"Batch Name");
    if(delB<0) return outJSON({status:"FAIL", msg:"Batch Name column not found."});
    var deleted=0;
    for(var di=delVals.length-1;di>=1;di--){
      if(String(delVals[di][delB]||"").trim().toUpperCase()===delBatch.toUpperCase()){
        delSheet.deleteRow(di+1); deleted++;
      }
    }
    var remain=delSheet.getLastRow()-1;
    if(remain>0){
      var serials2=[]; for(var sj=1;sj<=remain;sj++) serials2.push([sj]);
      delSheet.getRange(2,1,remain,1).setValues(serials2);
    }
    SpreadsheetApp.flush();
    return outJSON({status:"SUCCESS", batchName:delBatch, deleted:deleted, sheet:"Registered Students"});
  }

  if (data.action == "addStudent") {
    invalidateRegistrationLoginIndex_();
      if(data.adminPass != getAdminPassword()) return ContentService.createTextOutput("FAIL");
      var cType = String(data.course).trim();
      var targetSheetName = "Students_" + cType;
      var targetSheet = null;
      var allSheets = ss.getSheets();
      for(var s=0; s<allSheets.length; s++) {
          if(allSheets[s].getName().indexOf("Students") !== -1) {
              targetSheet = allSheets[s];
              if(allSheets[s].getName() === targetSheetName) { break; } 
          }
      }
      if(!targetSheet) {
          targetSheet = ss.insertSheet(targetSheetName);
          targetSheet.appendRow(["CISF NO", "Name", "Password", "Batch", "ExamType", "PassChanged"]);
          targetSheet.getRange(1, 1, 1, 6).setBackground("#004d99").setFontColor("white").setFontWeight("bold");
      }
      targetSheet.appendRow([data.roll, data.name, data.pass, data.batch, cType, "FALSE"]);
      return ContentService.createTextOutput("SUCCESS");
  }
  
  if (data.action == "unlockStudent") {
     if(data.pass != getAdminPassword()) return ContentService.createTextOutput("FAIL");
     var hSheet = ss.getSheetByName("History");
     var unlocked = false;
     if(hSheet && hSheet.getLastRow() > 1) {
         var hData = hSheet.getDataRange().getValues();
         var head = hData[0];
         var rCol = getColIdx(head, "CISF NO");
         var resCol = getColIdx(head, "Result") + 1; 
         for(var i = hData.length - 1; i >= 1; i--) {
             if(String(hData[i][rCol]).trim() == String(data.roll).trim()) {
                 if(String(hData[i][resCol-1]).trim() == "LOCKED") {
                     hSheet.getRange(i+1, resCol).setValue("STARTED");
                     unlocked = true;
                 }
             }
         }
     }
     if(unlocked) return ContentService.createTextOutput("SUCCESS");
     return ContentService.createTextOutput("NOT_FOUND");
  }
  
  if (data.action == "changePassword") {
    var changeRoll = String(data.roll || "").trim();
    var regSheetForPass = ss.getSheetByName("Registered Students") || ensureRegisteredStudentsSheet_();
    // Lock so simultaneous password changes / uploads do not collide ("failed to update").
    var pwLock = LockService.getScriptLock(), pwLocked = false;
    try { pwLock.waitLock(20000); pwLocked = true; } catch(ignoreLock) {}
    try {
    if(regSheetForPass.getLastRow() > 1) {
      // Read only the header row and the CISF NO column (not the whole sheet).
      var rh = regSheetForPass.getRange(1,1,1,Math.max(8,regSheetForPass.getLastColumn())).getDisplayValues()[0] || [];
      var rrCol=getColIdx(rh,"CISF NO");
      if(rrCol>=0) {
        var rolls = regSheetForPass.getRange(2,rrCol+1,regSheetForPass.getLastRow()-1,1).getDisplayValues();
        for(var ri=0;ri<rolls.length;ri++){
          if(String(rolls[ri][0]||"").trim()===changeRoll){
            var rowNo = ri+2;
            var rpCol=getColIdx(rh,"Password"), rcCol=getColIdx(rh,"PassChanged");
            if(rpCol>=0) regSheetForPass.getRange(rowNo,rpCol+1).setValue(String(data.newPass||""));
            if(rcCol>=0) regSheetForPass.getRange(rowNo,rcCol+1).setValue("TRUE");
            setRegisteredStudentPassword_(changeRoll,String(data.newPass||""),true);
            invalidateRegistrationLoginIndex_();
            SpreadsheetApp.flush();
            return ContentService.createTextOutput("SUCCESS");
          }
        }
      }
    }
    } finally { if(pwLocked) { try { pwLock.releaseLock(); } catch(ignoreRel) {} } }

    // Legacy fallback for older student records.
    var allSheets = ss.getSheets();
    for(var s=0; s<allSheets.length; s++) {
      if(allSheets[s].getName().indexOf("Students") !== -1) {
        var sSheet = allSheets[s]; var sData = sSheet.getDataRange().getValues(); 
        var pCol = getColIdx(sData[0], "Password") + 1; 
        var rCol = getColIdx(sData[0], "CISF NO");
        var pcCol = getColIdx(sData[0], "PassChanged") + 1;
        if (pcCol === 0) {
            sSheet.getRange(1, sData[0].length + 1).setValue("PassChanged");
            pcCol = sData[0].length + 1;
        }
        for(var i=1; i<sData.length; i++) {
           if(String(sData[i][rCol]).trim() == changeRoll) { 
               sSheet.getRange(i+1, pCol).setValue(data.newPass); 
               sSheet.getRange(i+1, pcCol).setValue("TRUE"); 
               return ContentService.createTextOutput("SUCCESS"); 
           }
        }
      }
    } return ContentService.createTextOutput("FAIL");
  }

  if (data.action == "logStatus" || data.action == "submitResult") {
    var lock = LockService.getScriptLock();
    try { lock.waitLock(30000); } catch(e) { } 
    
    try {
        var hSheet = ss.getSheetByName("History");
        var isNewH = false;
        if(!hSheet) { hSheet = ss.insertSheet("History"); isNewH = true; SpreadsheetApp.flush(); } 
        if(isNewH || hSheet.getLastRow() === 0) { 
            hSheet.appendRow(["Date","CISF NO","Name","Paper","PaperName","Score","Result","Type"]); 
            hSheet.getRange(1, 1, 1, 7).setBackground("#004d99").setFontColor("white").setFontWeight("bold"); 
            SpreadsheetApp.flush(); 
        }
        
        ensureHeaderColumn_(hSheet,"PaperName");
        var head = hSheet.getRange(1, 1, 1, hSheet.getLastColumn()).getValues()[0];
        var dateCol = getColIdx(head,"Date")+1, rollCol = getColIdx(head,"CISF NO")+1, nameCol = getColIdx(head,"Name")+1;
        var paperCol = getColIdx(head,"Paper")+1, paperNameCol = getColIdx(head,"PaperName")+1, scoreCol = getColIdx(head,"Score")+1, resCol = getColIdx(head,"Result")+1, typeCol = getColIdx(head,"Type")+1;
        
        var dateStr = formatDt(new Date()); var updateRow = findExistingHistoryRow_(hSheet, data.roll, data.name, data.paper, data.paperName);
        var cleanPaper = String(data.paper).split("###")[0].trim();
        var cleanPaperName = String(data.paperName || (String(data.paper).indexOf("###") !== -1 ? String(data.paper).split("###").slice(1).join("###") : "")).trim();
        var payloadName = String(data.name).trim().toLowerCase();
        var finalScore = data.score || "-";
        var resultText = data.result;
        var responsesLog = data.responses || "";
        var rowToFormat; var appendedNewHistoryRow = false;

        if(String(data.roll) === "ADMIN_TEST") { lock.releaseLock(); return ContentService.createTextOutput("Saved"); }

        if(updateRow !== -1) { 
            var existingRow = hSheet.getRange(updateRow, 1, 1, head.length).getValues()[0];
            existingRow[dateCol-1] = dateStr;
            existingRow[scoreCol-1] = finalScore;
            existingRow[resCol-1] = resultText;
            if (paperCol > 0) existingRow[paperCol-1] = cleanPaper;
            if (paperNameCol > 0) existingRow[paperNameCol-1] = cleanPaperName;
            if (typeof data.type !== "undefined") existingRow[typeCol-1] = data.type;
            hSheet.getRange(updateRow, 1, 1, existingRow.length).setValues([existingRow]);
            rowToFormat = updateRow;
        } else {
            var newRow = new Array(head.length).fill(""); 
            newRow[dateCol-1] = dateStr; newRow[rollCol-1] = data.roll; newRow[nameCol-1] = data.name;
            newRow[paperCol-1] = cleanPaper; if(paperNameCol > 0) newRow[paperNameCol-1] = cleanPaperName; newRow[scoreCol-1] = finalScore; newRow[resCol-1] = resultText; newRow[typeCol-1] = data.type || "U"; 
            if (data.batch) {
              var batchColIdx = getColIdx(head, "Batch");
              if (batchColIdx >= 0) newRow[batchColIdx] = data.batch;
            }
            hSheet.appendRow(newRow); rowToFormat = hSheet.getLastRow(); appendedNewHistoryRow = true;
        }

        var cellColor = (resultText === "PASS") ? "green" : ((resultText === "FAIL" || resultText === "COMPLETED") ? "red" : "black");
        hSheet.getRange(rowToFormat, resCol).setFontColor(cellColor).setFontWeight("bold");
        
        // Avoid global sorts on every submission. This was the main sheet performance bottleneck under heavy concurrent answers.
        // The result is stored by the student's identifier and is not required to be resorted after each update.
        if(resultText == "PASS" || resultText == "FAIL" || resultText == "COMPLETED") {
            // Individual U-/P- result sheets intentionally disabled.
            // Results continue to be stored in History and PermanentHistory below.

            var permSheet = ss.getSheetByName("PermanentHistory");
            var isNewPerm = false;
            if(!permSheet) { permSheet = ss.insertSheet("PermanentHistory"); isNewPerm = true; SpreadsheetApp.flush(); } 
            
            if(isNewPerm || permSheet.getLastRow() === 0) {
                permSheet.appendRow(["Date","CISF NO","Name","Paper","PaperName","Score","Result","Type","Batch","Responses"]); 
                permSheet.getRange(1, 1, 1, 9).setBackground("#004d99").setFontColor("white").setFontWeight("bold"); 
                SpreadsheetApp.flush();
            }
  
            ensureHeaderColumn_(permSheet,"PaperName");
            var pHdrs = permSheet.getRange(1, 1, 1, permSheet.getLastColumn()).getValues()[0];
            var batchIdx = getColIdx(pHdrs, "Batch");
            if (batchIdx === -1) { batchIdx = pHdrs.length; permSheet.getRange(1, batchIdx + 1).setValue("Batch"); pHdrs.push("Batch"); SpreadsheetApp.flush(); }
            var respIdx = getColIdx(pHdrs, "Responses");
            if (respIdx === -1) { respIdx = pHdrs.length; permSheet.getRange(1, respIdx + 1).setValue("Responses"); pHdrs.push("Responses"); SpreadsheetApp.flush(); }

            var newPermRow = new Array(pHdrs.length).fill("");
            newPermRow[getColIdx(pHdrs, "Date")] = dateStr;
            newPermRow[getColIdx(pHdrs, "CISF NO")] = data.roll;
            newPermRow[getColIdx(pHdrs, "Name")] = data.name;
            newPermRow[getColIdx(pHdrs, "Paper")] = cleanPaper;
            var pnIdx = getColIdx(pHdrs, "PaperName"); if(pnIdx === -1) { pnIdx = pHdrs.length; permSheet.getRange(1,pnIdx+1).setValue("PaperName"); pHdrs.push("PaperName"); }
            newPermRow[pnIdx] = cleanPaperName;
            newPermRow[getColIdx(pHdrs, "Score")] = finalScore;
            newPermRow[getColIdx(pHdrs, "Result")] = resultText;
            newPermRow[getColIdx(pHdrs, "Type")] = data.type;
            newPermRow[batchIdx] = data.batch || "N/A";
            newPermRow[respIdx] = responsesLog;
            
            permSheet.appendRow(newPermRow);
            var permRow = permSheet.getLastRow(); 
            permSheet.getRange(permRow, 6).setFontColor(cellColor).setFontWeight("bold");
        }
    } catch(err) {
    } finally {
        lock.releaseLock();
    }
    
    return ContentService.createTextOutput("Saved");
  }

  if (data.action == "changeAdminPassword") {
     var currentPass = String(data.currentPass || "");
     var newPass = String(data.newPass || "");
     if (currentPass !== getAdminPassword()) return outJSON({status:"FAIL", msg:"Current admin password is incorrect."});
     if (newPass.length < 6) return outJSON({status:"FAIL", msg:"New password must be at least 6 characters."});
     setAdminPassword_(newPass);
     return outJSON({status:"SUCCESS", msg:"Admin password changed successfully."});
  }

  if(data.action == "deleteAssignedData") {
     if(data.pass != getAdminPassword()) return ContentService.createTextOutput("FAIL");
     var typeSheetName = data.dataType; var code = data.code; var sheet = ss.getSheetByName(typeSheetName);
     if(sheet && sheet.getLastRow() > 1) {
         if(code === "ALL") { sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).clearContent(); } 
         else {
             var sData = sheet.getDataRange().getValues(); var cCol = getColIdx(sData[0], "PaperCode"); var nCol = getColIdx(sData[0], "PaperName");
             var wantedName = String(data.paperName || "").trim();
             for(var i = sData.length-1; i>=1; i--) {
                var codeMatch = cCol >= 0 && String(sData[i][cCol]).trim() === String(code).trim();
                var nameMatch = !wantedName || nCol < 0 || String(sData[i][nCol]).trim() === wantedName;
                if(codeMatch && nameMatch) sheet.deleteRow(i+1);
             }
         }
     } return ContentService.createTextOutput("SUCCESS");
  }

  if(data.action == "deleteBatchData") {
     if(data.pass != getAdminPassword()) return ContentService.createTextOutput("FAIL");
     var targetBatch = String(data.batch).trim();
     var pSheet = ss.getSheetByName("PermanentHistory");
     if(pSheet && pSheet.getLastRow() > 1) {
         var pData = pSheet.getDataRange().getValues(); var pbCol = getColIdx(pData[0], "Batch");
         for(var i = pData.length-1; i>=1; i--) { if(String(pData[i][pbCol]).trim() === targetBatch) { pSheet.deleteRow(i+1); } }
     }
     var allSheets = ss.getSheets();
     for(var s=0; s<allSheets.length; s++) {
        if(allSheets[s].getName().indexOf("Students") !== -1) {
            var stSheet = allSheets[s]; var stData = stSheet.getDataRange().getValues();
            if(stData.length > 1) {
                var sbCol = getColIdx(stData[0], "Batch");
                for(var i = stData.length-1; i>=1; i--) { if(String(stData[i][sbCol]).trim() === targetBatch) { stSheet.deleteRow(i+1); } }
            }
        }
     } return ContentService.createTextOutput("SUCCESS");
  }
  
  if (data.action == "updateNotice") { var sheet = ss.getSheetByName("Notices"); if(!sheet) { sheet = ss.insertSheet("Notices"); sheet.appendRow(["TargetExam", "NoticeText"]);} var nData = sheet.getDataRange().getValues(); var target = String(data.target || "").trim(); var note = String(data.notice || ""); var exists = false; var targetCol = getColIdx(nData[0], "TargetExam"); var textCol = getColIdx(nData[0], "NoticeText"); for(var i=1;i<nData.length;i++){ if(String(nData[i][targetCol]).trim()===target){ sheet.getRange(i+1, textCol+1).setValue(note); exists=true; break; } } if(!exists) sheet.appendRow([target,note]); return ContentService.createTextOutput("SUCCESS"); }
  if (data.action == "updateQuestion") { var upSheet = ss.getSheetByName("QuestionUpdates"); if(!upSheet) upSheet = ss.insertSheet("QuestionUpdates"); var allD = upSheet.getDataRange().getValues(); var paper = String(data.paperCode || "").trim(); var qid = String(data.qid || "").trim(); var exists = false; for(var i=1;i<allD.length;i++){ if(String(allD[i][0]).trim()===paper && String(allD[i][1]).trim()===qid){ upSheet.getRange(i+1,1,1,allD[0].length).setValues([ [paper,qid,data.questionEn || "",data.questionHi || "",data.aEn || "",data.aHi || "",data.bEn || "",data.bHi || "",data.cEn || "",data.cHi || "",data.dEn || "",data.dHi || "",data.correct || ""] ]); exists=true; break; } } if(!exists){ upSheet.appendRow([paper,qid,data.questionEn || "",data.questionHi || "",data.aEn || "",data.aHi || "",data.bEn || "",data.bHi || "",data.cEn || "",data.cHi || "",data.dEn || "",data.dHi || "",data.correct || ""]); } return ContentService.createTextOutput("SUCCESS"); }
  if (data.action == "clearUpdates") { var upSheet = ss.getSheetByName("QuestionUpdates"); if(upSheet) upSheet.clear(); return ContentService.createTextOutput("SUCCESS"); }
  if (data.action == "addSchedule") {
     var sheet = ss.getSheetByName("Schedule");
     if(!sheet) { sheet = ss.insertSheet("Schedule"); sheet.appendRow(["ExamName", "PaperCode", "PaperName", "TargetExam", "StartTime", "EndTime", "Duration"]); }
     ensureHeaderColumn_(sheet,"PaperName");
     appendByHeaders_(sheet, {ExamName:data.name, PaperCode:data.code, PaperName:data.paperName || "", TargetExam:data.target, StartTime:data.start, EndTime:data.end, Duration:data.duration});
     return ContentService.createTextOutput("SUCCESS");
  }
  if (data.action == "addModule") {
     var sheet = ss.getSheetByName("Modules");
     if(!sheet) { sheet = ss.insertSheet("Modules"); sheet.appendRow(["ModuleName", "PaperCode", "PaperName", "TargetExam", "Duration"]); }
     ensureHeaderColumn_(sheet,"PaperName");
     appendByHeaders_(sheet, {ModuleName:data.name, PaperCode:data.code, PaperName:data.paperName || "", TargetExam:data.target, Duration:data.duration});
     return ContentService.createTextOutput("SUCCESS");
  }
  if (data.action == "addMaterial") { var sheet = ss.getSheetByName("Materials"); if(!sheet) {sheet = ss.insertSheet("Materials"); sheet.appendRow(["Title", "Subject", "TargetExam", "Link"]);} sheet.appendRow([data.title, data.subject || "", data.target || "", data.link || ""]); return ContentService.createTextOutput("SUCCESS"); }

  return ContentService.createTextOutput("UNKNOWN_ACTION");
}




