var _OUT_CALLBACK = "";

// ============================================================
// PAPER SPREADSHEET TARGET
// ============================================================
// The old package contained an invalid/hard-coded Spreadsheet ID.
// This version stores the Paper Spreadsheet ID in Script Properties.
// Run setupPaperBackend() ONCE from the Apps Script project that is
// bound to the dedicated Paper Google Spreadsheet.
// ============================================================

function setupPaperBackend() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Open this Apps Script from the dedicated Paper Google Spreadsheet and run setupPaperBackend() once.");
  PropertiesService.getScriptProperties().setProperty("PAPER_SPREADSHEET_ID", ss.getId());
  ensurePaperSheets_(ss);
  return "SUCCESS: Paper Spreadsheet configured: " + ss.getName() + " / " + ss.getId();
}

function getQuestionSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty("PAPER_SPREADSHEET_ID");
  if (!id) {
    // Helpful fallback for a bound script during manual setup.
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      id = active.getId();
      PropertiesService.getScriptProperties().setProperty("PAPER_SPREADSHEET_ID", id);
    }
  }
  if (!id) throw new Error("Paper Spreadsheet is not configured. In the Paper Spreadsheet: Extensions > Apps Script, run setupPaperBackend() once, then redeploy the web app.");
  try {
    return SpreadsheetApp.openById(id);
  } catch (err) {
    throw new Error("Cannot open the configured Paper Spreadsheet. Run setupPaperBackend() again from the correct Paper Spreadsheet. Details: " + err.message);
  }
}

// ============================================================
// NORMALIZE HEADER
// ============================================================

function normalizeHeader(v) {

  return String(
    v == null ? "" : v
  )
    .toLowerCase()
    .replace(
      /[\s_\-\.()\/\\:]+/g,
      ""
    )
    .replace(
      /[^a-z0-9]/g,
      ""
    );
}

// ============================================================
// FIND COLUMN
// ============================================================

function findColumn(headers, aliases) {

  for (
    var i = 0;
    i < aliases.length;
    i++
  ) {

    var wanted =
      normalizeHeader(
        aliases[i]
      );

    for (
      var j = 0;
      j < headers.length;
      j++
    ) {

      if (
        headers[j] === wanted
      ) {

        return j;
      }
    }
  }

  return -1;
}

// ============================================================
// CONVERT ANSWER TO NUMBER
// ============================================================

function answerToNumber(value) {

  var s =
    String(
      value == null ? "" : value
    )
      .trim()
      .toUpperCase();

  if (
    s === "A" ||
    s === "1"
  ) {

    return 1;
  }

  if (
    s === "B" ||
    s === "2"
  ) {

    return 2;
  }

  if (
    s === "C" ||
    s === "3"
  ) {

    return 3;
  }

  if (
    s === "D" ||
    s === "4"
  ) {

    return 4;
  }

  var n =
    Number(s);

  return (
    n >= 1 &&
    n <= 4
  )
    ? n
    : 1;
}

// ============================================================
// HEADER ROW CHECK
// ============================================================
// IMPORTANT FIX:
// This function was missing in the version you pasted,
// causing:
//
// "isHeaderRow is not defined"
//
// ============================================================

function isHeaderRow(row) {

  if (
    !row ||
    !row.length
  ) {

    return false;
  }

  var joined =
    row
      .map(function(v) {

        return normalizeHeader(v);

      })
      .join("|");

  return (
    joined.indexOf(
      "questionenglish"
    ) >= 0

    ||

    joined.indexOf(
      "questionen"
    ) >= 0

    ||

    joined.indexOf(
      "correct"
    ) >= 0

    ||

    joined.indexOf(
      "qid"
    ) >= 0
  );
}

// ============================================================
// READ QUESTION PAPER ROWS
// ============================================================

function getQuestionPaperRows(sheet) {

  if (
    !sheet ||
    sheet.getLastRow() < 1
  ) {

    return [];
  }

  var values =
    sheet
      .getDataRange()
      .getValues();

  if (
    !values ||
    !values.length
  ) {

    return [];
  }

  // ----------------------------------------------------------
  // FIND HEADER ROW
  // ----------------------------------------------------------

  var headerRowIndex = -1;

  for (
    var h = 0;
    h < Math.min(values.length, 5);
    h++
  ) {

    if (
      isHeaderRow(
        values[h]
      )
    ) {

      headerRowIndex = h;

      break;
    }
  }

  var startRow = 0;

  var headers = [];

  if (
    headerRowIndex >= 0
  ) {

    headers =
      values[
        headerRowIndex
      ].map(
        normalizeHeader
      );

    startRow =
      headerRowIndex + 1;

  } else {

    // Standard 12-column layout

    headers = [

      "qid",

      "correct",

      "questionenglish",

      "questionhindi",

      "aenglish",

      "ahindi",

      "benglish",

      "bhindi",

      "cenglish",

      "chindi",

      "denglish",

      "dhindi"

    ];

    startRow = 0;
  }

  // ----------------------------------------------------------
  // FIND COLUMNS
  // ----------------------------------------------------------

  var c = {};

  c.qid =
    findColumn(
      headers,
      [
        "qid",
        "questionid",
        "questionno",
        "no",
        "srno",
        "serialno",
        "number"
      ]
    );

  c.correct =
    findColumn(
      headers,
      [
        "correct",
        "correctoption",
        "correctanswer",
        "answer",
        "ans",
        "rightanswer"
      ]
    );

  c.qen =
    findColumn(
      headers,
      [
        "questionenglish",
        "questionen",
        "question",
        "qenglish",
        "englishquestion"
      ]
    );

  c.qhi =
    findColumn(
      headers,
      [
        "questionhindi",
        "questionhi",
        "hindquestion",
        "hindiquestion",
        "qhi"
      ]
    );

  c.aen =
    findColumn(
      headers,
      [
        "aenglish",
        "optionaenglish",
        "optionaen",
        "aen",
        "optiona"
      ]
    );

  c.ahi =
    findColumn(
      headers,
      [
        "ahindi",
        "optionahindi",
        "optionahi",
        "ahi"
      ]
    );

  c.ben =
    findColumn(
      headers,
      [
        "benglish",
        "optionbenglish",
        "optionben",
        "ben",
        "optionb"
      ]
    );

  c.bhi =
    findColumn(
      headers,
      [
        "bhindi",
        "optionbhindi",
        "optionbhi",
        "bhi"
      ]
    );

  c.cen =
    findColumn(
      headers,
      [
        "cenglish",
        "optioncenglish",
        "optioncen",
        "cen",
        "optionc"
      ]
    );

  c.chi =
    findColumn(
      headers,
      [
        "chindi",
        "optionchindi",
        "optionchi",
        "chi"
      ]
    );

  c.den =
    findColumn(
      headers,
      [
        "denglish",
        "optiondenglish",
        "optionden",
        "den",
        "optiond"
      ]
    );

  c.dhi =
    findColumn(
      headers,
      [
        "dhindi",
        "optiondhindi",
        "optiondhi",
        "dhi"
      ]
    );

  // ----------------------------------------------------------
  // DEFAULT COLUMN POSITIONS
  // ----------------------------------------------------------

  var defaults = {

    qid: 0,

    correct: 1,

    qen: 2,

    qhi: 3,

    aen: 4,

    ahi: 5,

    ben: 6,

    bhi: 7,

    cen: 8,

    chi: 9,

    den: 10,

    dhi: 11

  };

  Object.keys(
    defaults
  ).forEach(
    function(k) {

      if (
        c[k] < 0
      ) {

        c[k] =
          defaults[k];
      }

    }
  );

  // ----------------------------------------------------------
  // BUILD QUESTION DATA
  // ----------------------------------------------------------

  var out = [];

  for (
    var r = startRow;
    r < values.length;
    r++
  ) {

    var row =
      values[r] || [];

    var qText =
      String(
        row[c.qen] == null
          ? ""
          : row[c.qen]
      ).trim();

    var aText =
      String(
        row[c.aen] == null
          ? ""
          : row[c.aen]
      ).trim();

    var bText =
      String(
        row[c.ben] == null
          ? ""
          : row[c.ben]
      ).trim();

    var cText =
      String(
        row[c.cen] == null
          ? ""
          : row[c.cen]
      ).trim();

    var dText =
      String(
        row[c.den] == null
          ? ""
          : row[c.den]
      ).trim();

    // Ignore completely blank rows

    if (
      !qText &&
      !aText &&
      !bText &&
      !cText &&
      !dText
    ) {

      continue;
    }

    // Ignore repeated header rows

    if (
      isHeaderRow(row)
    ) {

      continue;
    }

    var qid =
      String(
        row[c.qid] == null
          ? ""
          : row[c.qid]
      ).trim();

    // Automatically create QID

    if (!qid) {

      qid =
        String(
          out.length + 1
        );
    }

    out.push([

      qid,

      answerToNumber(
        row[c.correct]
      ),

      qText,

      String(
        row[c.qhi] == null
          ? ""
          : row[c.qhi]
      ),

      [
        aText,

        String(
          row[c.ahi] == null
            ? ""
            : row[c.ahi]
        )
      ],

      [
        bText,

        String(
          row[c.bhi] == null
            ? ""
            : row[c.bhi]
        )
      ],

      [
        cText,

        String(
          row[c.chi] == null
            ? ""
            : row[c.chi]
        )
      ],

      [
        dText,

        String(
          row[c.dhi] == null
            ? ""
            : row[c.dhi]
        )
      ]

    ]);
  }

  return out;
}

function answerToLetter(value) {
  var n = answerToNumber(value);
  return n === 1 ? "A" : n === 2 ? "B" : n === 3 ? "C" : n === 4 ? "D" : "A";
}

// ============================================================
// PAPER DATABASE CONFIGURATION
// ============================================================
var PAPER_CODES = [
  "BASIC",
  "SCREENER(INT)",
  "SCREENER(REC)",
  "INSTRUCTOR",
  "BASIC REFRESHER",
  "INDUCTION",
  "OJT"
];

var PAPER_HEADERS = [
  "PaperName", "QID", "Correct", "QuestionEnglish", "QuestionHindi",
  "AEnglish", "AHindi", "BEnglish", "BHindi", "CEnglish", "CHindi",
  "DEnglish", "DHindi"
];

function ensurePaperSheetsOnce_(ss) {
  var cache = CacheService.getScriptCache();
  if(cache.get("PAPER_SHEETS_READY") === "1") return;
  var props = PropertiesService.getScriptProperties();
  if(String(props.getProperty("PAPER_SHEETS_READY") || "") === "1") { cache.put("PAPER_SHEETS_READY", "1", 21600); return; }
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if(String(props.getProperty("PAPER_SHEETS_READY") || "") !== "1") { ensurePaperSheets_(ss); props.setProperty("PAPER_SHEETS_READY", "1"); }
    cache.put("PAPER_SHEETS_READY", "1", 21600);
  } finally { try { lock.releaseLock(); } catch(ignore) {} }
}

function paperCacheKey_(prefix, code, name) {
  var raw = prefix + "|" + String(code || "") + "|" + String(name || "");
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw);
  var hex = bytes.map(function(b){ var n=(b<0?b+256:b).toString(16); return n.length===1?"0"+n:n; }).join("");
  return "ASTI_PAPER_" + prefix + "_" + hex;
}

function invalidatePaperCaches_(code, name) {
  var cache = CacheService.getScriptCache();
  cache.remove("ASTI_PAPER_CATALOG");
  if(code && name) cache.remove(paperCacheKey_("Q", code, name));
  if(String(PropertiesService.getScriptProperties().getProperty("PAPER_SHEETS_READY") || "") === "1") cache.put("PAPER_SHEETS_READY", "1", 21600);
}

function ensurePaperSheets_(ss) {
  PAPER_CODES.forEach(function(code) {
    var sh = ss.getSheetByName(code);
    if (!sh) {
      sh = ss.insertSheet(code);
      sh.getRange(1,1,1,PAPER_HEADERS.length).setValues([PAPER_HEADERS]);
    } else {
      var lastCol = Math.max(1, sh.getLastColumn());
      var first = sh.getRange(1,1,1,lastCol).getDisplayValues()[0].map(normalizeHeader);
      var hasPaperName = first.indexOf("papername") !== -1;
      if (!hasPaperName && lastCol >= 12) {
        // Legacy 12-column question sheet: insert the new PaperName column
        // without destroying the existing questions.
        sh.insertColumnBefore(1);
        var lastRow = sh.getLastRow();
        sh.getRange(1,1).setValue("PaperName");
        if (lastRow > 1) {
          sh.getRange(2,1,lastRow-1,1).setValues(new Array(lastRow-1).fill(["LEGACY PAPER"]));
        }
      }
      sh.getRange(1,1,1,PAPER_HEADERS.length).setValues([PAPER_HEADERS]);
    }
    sh.getRange(1,1,1,PAPER_HEADERS.length)
      .setBackground("#004d99").setFontColor("white").setFontWeight("bold");
    sh.setFrozenRows(1);
  });
}

// ============================================================
// PAPER NAME / QUESTION STORAGE
// ============================================================
function cleanPaperCode_(v) {
  var s = String(v || "").trim();
  for (var i = 0; i < PAPER_CODES.length; i++) {
    if (PAPER_CODES[i].toUpperCase() === s.toUpperCase()) return PAPER_CODES[i];
  }
  return "";
}

function getPaperNames_(ss, code) {
  code = cleanPaperCode_(code);
  if (!code) return [];
  var sh = ss.getSheetByName(code);
  if (!sh || sh.getLastRow() < 2) return [];
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
  var seen = {}; var out = [];
  vals.forEach(function(r) { var n = String(r[0] || "").trim(); if (n && !seen[n]) { seen[n] = true; out.push(n); } });
  return out.sort(function(a,b){ return a.localeCompare(b); });
}

function getPaperCatalog_(ss) {
  var cache = CacheService.getScriptCache(); var cached = cache.get("ASTI_PAPER_CATALOG");
  if(cached) { try { return JSON.parse(cached); } catch(ignore) {} }
  var out = {}; PAPER_CODES.forEach(function(code) { out[code] = getPaperNames_(ss, code); });
  try { cache.put("ASTI_PAPER_CATALOG", JSON.stringify(out), 300); } catch(ignore2) {}
  return out;
}

function getStoredQuestionRows_(ss, code, paperName) {
  code = cleanPaperCode_(code); paperName = String(paperName || "").trim();
  if (!code || !paperName) return [];
  var cache = CacheService.getScriptCache(); var key = paperCacheKey_("Q", code, paperName); var cached = cache.get(key);
  if(cached) { try { return JSON.parse(cached); } catch(ignore) {} }
  var sh = ss.getSheetByName(code); if (!sh || sh.getLastRow() < 2) return [];
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, PAPER_HEADERS.length).getDisplayValues(); var out = [];
  rows.forEach(function(r) {
    if (String(r[0] || "").trim() !== paperName) return;
    out.push([String(r[1] || "").trim(), answerToNumber(r[2]), String(r[3] || ""), String(r[4] || ""), [String(r[5] || ""), String(r[6] || "")], [String(r[7] || ""), String(r[8] || "")], [String(r[9] || ""), String(r[10] || "")], [String(r[11] || ""), String(r[12] || "")]]);
  });
  try { var packed=JSON.stringify(out); if(packed.length < 95000) cache.put(key, packed, 300); } catch(ignore2) {}
  return out;
}

function getExamPaper_(ss, code, paperName) {
  code = cleanPaperCode_(code); paperName = String(paperName || "").trim();
  if(!code) return {status:"FAIL",msg:"Invalid Paper Code"};
  if(!paperName) { var names=getPaperNames_(ss,code); paperName=names.length?names[0]:""; }
  if(!paperName) return {status:"FAIL",msg:"No paper uploaded for " + code};
  var qs=getStoredQuestionRows_(ss,code,paperName);
  if(!qs.length) return {status:"FAIL",msg:"Paper not found: " + code + " / " + paperName};
  return {status:"SUCCESS",paperCode:code,paperName:paperName,questions:qs,total:qs.length};
}

function removePaper_(ss, code, paperName) {
  code = cleanPaperCode_(code);
  paperName = String(paperName || "").trim();
  if (!code || !paperName) return false;
  var sh = ss.getSheetByName(code);
  if (!sh || sh.getLastRow() < 2) return false;

  var last = sh.getLastRow();
  var vals = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
  var matches = [];
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || "").trim() === paperName) matches.push(i + 2);
  }
  if (!matches.length) return false;

  // Delete contiguous blocks from bottom to top. This preserves all other
  // papers while reducing dozens/hundreds of deleteRow() calls to a few calls.
  var blockStart = matches[matches.length - 1];
  var blockEnd = blockStart;
  for (var j = matches.length - 2; j >= -1; j--) {
    var row = j >= 0 ? matches[j] : 0;
    if (j >= 0 && row === blockStart - 1) {
      blockStart = row;
    } else {
      sh.deleteRows(blockStart, blockEnd - blockStart + 1);
      if (j >= 0) {
        blockStart = row;
        blockEnd = row;
      }
    }
  }
  invalidatePaperCaches_(code, paperName);
  return true;
}

function textValue_(v) {
  return String(v == null ? "" : v).trim();
}

function optionEnglish_(q, index) {
  return (q && q[index] && Array.isArray(q[index])) ? textValue_(q[index][0]) : "";
}

function optionHindi_(q, index) {
  return (q && q[index] && Array.isArray(q[index])) ? textValue_(q[index][1]) : "";
}

function savePaperRows_(ss, code, paperName, questions) {
  code = cleanPaperCode_(code);
  paperName = String(paperName || "").trim();
  if (!code) throw new Error("Invalid Paper Code");
  if (!paperName) throw new Error("Paper Name is required");
  if (!questions || !questions.length) throw new Error("No valid questions found in Excel file");

  var sh = ss.getSheetByName(code);
  if (!sh) { ensurePaperSheets_(ss); sh = ss.getSheetByName(code); }
  removePaper_(ss, code, paperName);

  var rows = questions.map(function(q) {
    return [paperName, q[0], answerToLetter(q[1]), q[2], q[3], q[4][0], q[4][1], q[5][0], q[5][1], q[6][0], q[6][1], q[7][0], q[7][1]];
  });
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, PAPER_HEADERS.length).setValues(rows);
  invalidatePaperCaches_(code, paperName);
  return rows.length;
}

// ============================================================
// SIMPLE XLSX READER (FIRST WORKSHEET)
// ============================================================
function xmlDecode_(s) {
  s = String(s == null ? "" : s);
  try { return XmlService.parse("<root>" + s + "</root>").getRootElement().getText(); }
  catch (e) {
    return s.replace(/&amp;/g,"&").replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/&quot;/g,'"').replace(/&apos;/g,"'");
  }
}

function parseSharedStrings_(xml) {
  var out = [];
  var re = /<si[\s\S]*?<\/si>/g, m;
  while ((m = re.exec(xml)) !== null) {
    var block = m[0];
    var texts = [], tr = /<t(?:\s[^>]*)?>([\s\\S]*?)<\/t>/g, tm;
    while ((tm = tr.exec(block)) !== null) texts.push(xmlDecode_(tm[1]));
    out.push(texts.join(""));
  }
  return out;
}

function columnNumberFromRef_(ref) {
  var m = String(ref || "").match(/^([A-Z]+)\d+$/i);
  if (!m) return 0;
  var letters = m[1].toUpperCase(), n = 0;
  for (var i=0;i<letters.length;i++) n = n * 26 + letters.charCodeAt(i)-64;
  return n;
}

function parseXlsxFirstSheet_(base64) {
  var bytes = Utilities.base64Decode(String(base64 || ""));
  if (!bytes.length) throw new Error("Excel file is empty");
  var zipBlob = Utilities.newBlob(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "upload.xlsx");
  var files = Utilities.unzip(zipBlob);
  var shared = [];
  var sheetBlob = null;
  files.forEach(function(b) {
    var n = b.getName();
    if (n === "xl/sharedStrings.xml") shared = parseSharedStrings_(b.getDataAsString());
    if (!sheetBlob && /^xl\/worksheets\/sheet\d+\.xml$/i.test(n)) sheetBlob = b;
  });
  if (!sheetBlob) throw new Error("No worksheet was found in the Excel file");

  var xml = sheetBlob.getDataAsString();
  var rows = [], rowRe = /<row(?:\s[^>]*)?>([\s\S]*?)<\/row>/g, rm;
  while ((rm = rowRe.exec(xml)) !== null) {
    var rowBlock = rm[1];
    var cells = [];
    var cellRe = /<c\s+([^>]*\br="([A-Z]+\d+)"[^>]*)>([\s\S]*?)<\/c>/g;
    var cm;
    while ((cm = cellRe.exec(rowBlock)) !== null) {
      var attrs = cm[1], ref = cm[2], block = cm[3], col = columnNumberFromRef_(ref);
      var t = "";
      var openMatch = attrs.match(/\bt="([^"]+)"/);
      if (openMatch) t = openMatch[1];
      var vMatch = block.match(/<v>([\s\S]*?)<\/v>/);
      var inlineMatch = block.match(/<is>([\s\S]*?)<\/is>/);
      var val = "";
      if (inlineMatch) {
        var ts=[], tm, tr2=/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;
        while((tm=tr2.exec(inlineMatch[1]))!==null) ts.push(xmlDecode_(tm[1]));
        val=ts.join("");
      } else if (vMatch) {
        val = xmlDecode_(vMatch[1]);
        if (t === "s") { var si = Number(val); val = isFinite(si) && shared[si] !== undefined ? shared[si] : ""; }
        else if (t === "b") val = val === "1" ? "TRUE" : "FALSE";
      }
      cells.push({col:col, value:val});
    }
    if (cells.length) {
      var maxCol = Math.max.apply(null, cells.map(function(c){return c.col;}));
      var arr = new Array(maxCol).fill("");
      cells.forEach(function(c){ arr[c.col-1] = c.value; });
      rows.push(arr);
    }
  }
  return rows;
}

function excelRowsToQuestions_(rows) {
  if (!rows || !rows.length) return [];
  var headerIndex = -1;
  for (var i=0;i<Math.min(rows.length,8);i++) if (isHeaderRow(rows[i])) { headerIndex=i; break; }
  var headers = headerIndex >= 0 ? rows[headerIndex].map(normalizeHeader) : ["qid","correct","questionenglish","questionhindi","aenglish","ahindi","benglish","bhindi","cenglish","chindi","denglish","dhindi"];
  var start = headerIndex >= 0 ? headerIndex+1 : 0;
  var c = {};
  c.qid=findColumn(headers,["qid","questionid","questionno","no","srno","serialno","number"]);
  c.correct=findColumn(headers,["correct","correctoption","correctanswer","answer","ans","rightanswer"]);
  c.qen=findColumn(headers,["questionenglish","questionen","question","qenglish","englishquestion"]);
  c.qhi=findColumn(headers,["questionhindi","questionhi","hindquestion","hindiquestion","qhi"]);
  c.aen=findColumn(headers,["aenglish","optionaenglish","optionaen","a"]);
  c.ahi=findColumn(headers,["ahindi","optionahindi","optionahi"]);
  c.ben=findColumn(headers,["benglish","optionbenglish","optionben","b"]);
  c.bhi=findColumn(headers,["bhindi","optionbhindi","optionbhi"]);
  c.cen=findColumn(headers,["cenglish","optioncenglish","optioncen","c"]);
  c.chi=findColumn(headers,["chindi","optionchindi","optionchi"]);
  c.den=findColumn(headers,["denglish","optiondenglish","optionden","d"]);
  c.dhi=findColumn(headers,["dhindi","optiondhindi","optiondhi"]);
  var alternatingBilingual = headers.length >= 11 &&
    (c.qen === 1 || headers[1] === "questionenglish") &&
    (c.qhi === 2 || headers[2] === "questionhindi") &&
    (headers[3] === "optionaenglish" || headers[3] === "aenglish");
  var defaults = alternatingBilingual
    ? {qid:0,qen:1,qhi:2,aen:3,ahi:4,ben:5,bhi:6,cen:7,chi:8,den:9,dhi:10}
    : {qid:0,correct:1,qen:2,qhi:3,aen:4,ahi:5,ben:6,bhi:7,cen:8,chi:9,den:10,dhi:11};
  Object.keys(defaults).forEach(function(k){if(c[k]<0)c[k]=defaults[k];});
  var out=[];
  for(var r=start;r<rows.length;r++) {
    var row=rows[r]||[];
    var qText=String(row[c.qen] == null ? "" : row[c.qen]).trim();
    var aText=String(row[c.aen] == null ? "" : row[c.aen]).trim();
    var bText=String(row[c.ben] == null ? "" : row[c.ben]).trim();
    var cText=String(row[c.cen] == null ? "" : row[c.cen]).trim();
    var dText=String(row[c.den] == null ? "" : row[c.den]).trim();
    if(!qText&&!aText&&!bText&&!cText&&!dText) continue;
    if(isHeaderRow(row)) continue;
    var qid=String(row[c.qid] == null ? "" : row[c.qid]).trim() || String(out.length+1);
    out.push([qid,answerToNumber(row[c.correct]),qText,String(row[c.qhi] == null ? "" : row[c.qhi]),[aText,String(row[c.ahi] == null ? "" : row[c.ahi])],[bText,String(row[c.bhi] == null ? "" : row[c.bhi])],[cText,String(row[c.chi] == null ? "" : row[c.chi])],[dText,String(row[c.dhi] == null ? "" : row[c.dhi])]]);
  }
  return out;
}

// ============================================================
// JSON / AUTH
// ============================================================
function outJSON(data) {
  var callback = String(_OUT_CALLBACK || "");
  if (callback && !/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) callback = "";
  var payload = JSON.stringify(data);
  if (callback) return ContentService.createTextOutput(callback + "(" + payload + ");").setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
}
function getAdminPassword() {
  var saved=String(PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD")||"").trim();
  return saved || "Rakesh";
}
function setAdminPassword_(newPass) { PropertiesService.getScriptProperties().setProperty("ADMIN_PASSWORD",String(newPass||"").trim()); return true; }
function verifyAdminPassword(pass) { return String(pass||"").trim() === getAdminPassword(); }

// ============================================================
// PAPER SPREADSHEET DIAGNOSTIC
// ============================================================
function testPaperSpreadsheet() {
  try {
    var ss = getQuestionSpreadsheet_();
    ensurePaperSheets_(ss);
    return outJSON({
      status: "SUCCESS",
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      paperCodes: PAPER_CODES,
      sheets: ss.getSheets().map(function(sh){ return sh.getName(); })
    });
  } catch (err) {
    return outJSON({status:"FAIL", msg:String(err && err.message || err)});
  }
}

// ============================================================
// GET
// ============================================================
function doGet(e) {
  _OUT_CALLBACK=String(e&&e.parameter&&e.parameter.callback||"");
  var action=e&&e.parameter?String(e.parameter.action||""):"";
  try {
    // Opening the Web App root without an action also performs a Paper
    // Spreadsheet health check. Normal portal requests still use action=... .
    if(!action || action === "testPaperSpreadsheet") return testPaperSpreadsheet();

    var ss=getQuestionSpreadsheet_();
    ensurePaperSheetsOnce_(ss);

    if(action === "getAvailableQuestionPapers" || action === "getPaperCatalog") {
      return outJSON({status:"SUCCESS", paperCodes:PAPER_CODES, papers:getPaperCatalog_(ss)});
    }
    if(action === "getUploadedQuestionPapers") {
      var cat=getPaperCatalog_(ss), names=[];
      PAPER_CODES.forEach(function(code){ cat[code].forEach(function(name){ names.push({paperCode:code,paperName:name,questionCount:getStoredQuestionRows_(ss,code,name).length}); }); });
      return outJSON({status:"SUCCESS",papers:names,catalog:cat});
    }
    if(action === "getPaperNames") {
      var code=cleanPaperCode_(e.parameter.paperCode); if(!code) return outJSON({status:"FAIL",msg:"Invalid Paper Code"});
      return outJSON({status:"SUCCESS",paperCode:code,paperNames:getPaperNames_(ss,code)});
    }
    if(action === "getQuestionPaper") {
      var code2=cleanPaperCode_(e.parameter.paperCode), name2=String(e.parameter.paperName||"").trim();
      if(!code2||!name2) return outJSON({status:"FAIL",msg:"Paper Code and Paper Name are required"});
      var qs=getStoredQuestionRows_(ss,code2,name2);
      if(!qs.length) return outJSON({status:"FAIL",msg:"Paper not found: "+code2+" / "+name2});
      return outJSON({status:"SUCCESS",paperCode:code2,paperName:name2,questions:qs,total:qs.length});
    }
    if(action === "getExamPaper") {
      return outJSON(getExamPaper_(ss, e.parameter.paperCode, e.parameter.paperName));
    }
    if(action === "removePaper") {
      if(!verifyAdminPassword(e.parameter.pass)) return outJSON({status:"FAIL",msg:"Wrong admin password"});
      var rc=cleanPaperCode_(e.parameter.paperCode), rn=String(e.parameter.paperName||"").trim();
      if(!rc||!rn) return outJSON({status:"FAIL",msg:"Paper Code and Paper Name are required"});
      return outJSON({status:"SUCCESS",removed:removePaper_(ss,rc,rn),paperCode:rc,paperName:rn});
    }
    if(action === "changeAdminPassword") {
      if(!verifyAdminPassword(e.parameter.currentPass)) return outJSON({status:"FAIL",msg:"Current password is incorrect"});
      var np=String(e.parameter.newPass||"").trim(); if(np.length<6) return outJSON({status:"FAIL",msg:"New password must be at least 6 characters"});
      setAdminPassword_(np); return outJSON({status:"SUCCESS",msg:"Admin password changed successfully"});
    }
    return outJSON({status:"FAIL",msg:"Unknown action"});
  } catch(err) { return outJSON({status:"FAIL",msg:err.message}); }
}

// ============================================================
// POST - EXCEL UPLOAD
// ============================================================
function doPost(e) {
  try {
    // Accept both normal form POSTs (used by the browser Excel uploader)
    // and JSON POSTs for backward compatibility.
    var data = {};
    if(e && e.parameter && e.parameter.action) {
      data = {
        action: e.parameter.action,
        pass: e.parameter.pass || "",
        paperCode: e.parameter.paperCode || "",
        paperName: e.parameter.paperName || "",
        fileName: e.parameter.fileName || "",
        fileBase64: e.parameter.fileBase64 || ""
      };
    } else if(e && e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); }
      catch(parseErr) { throw new Error("Invalid upload request: " + parseErr.message); }
    }
    if(String(data.action||"") === "savePaperQuestions") {
      if(!verifyAdminPassword(data.pass)) return outJSON({status:"FAIL",msg:"Wrong admin password"});
      var ssFast=getQuestionSpreadsheet_(); ensurePaperSheetsOnce_(ssFast);
      var codeFast=cleanPaperCode_(data.paperCode), nameFast=String(data.paperName||"").trim();
      if(!codeFast) return outJSON({status:"FAIL",msg:"Invalid Paper Code"});
      if(!nameFast) return outJSON({status:"FAIL",msg:"Paper Name is required"});
      var questionsFast=data.questions;
      if(typeof questionsFast === "string") {
        try { questionsFast=JSON.parse(questionsFast); } catch(errQ) { return outJSON({status:"FAIL",msg:"Invalid question data"}); }
      }
      if(!Array.isArray(questionsFast) || !questionsFast.length) return outJSON({status:"FAIL",msg:"No valid questions received from the browser."});
      // Browser-parsed question rows are already in the same compact structure
      // used by savePaperRows_. One setValues() writes the complete paper.
      var totalFast=savePaperRows_(ssFast,codeFast,nameFast,questionsFast);
      return outJSON({status:"SUCCESS",paperCode:codeFast,paperName:nameFast,total:totalFast,spreadsheetId:ssFast.getId(),spreadsheetName:ssFast.getName(),msg:"Paper uploaded successfully"});
    }
    if(String(data.action||"") === "uploadPaperExcel") {
      if(!verifyAdminPassword(data.pass)) return outJSON({status:"FAIL",msg:"Wrong admin password"});
      var ss=getQuestionSpreadsheet_(); ensurePaperSheetsOnce_(ss);
      var code=cleanPaperCode_(data.paperCode), name=String(data.paperName||"").trim();
      if(!code) return outJSON({status:"FAIL",msg:"Invalid Paper Code"});
      if(!name) return outJSON({status:"FAIL",msg:"Paper Name is required"});
      var questions=excelRowsToQuestions_(parseXlsxFirstSheet_(data.fileBase64));
      if(!questions.length) return outJSON({status:"FAIL",msg:"No valid questions found. Check Excel headings/data."});
      var total=savePaperRows_(ss,code,name,questions);
      // Verify the actual saved rows in the dedicated Paper Spreadsheet before
      // reporting success. This catches wrong spreadsheet/deployment targets.
      var savedRows = getStoredQuestionRows_(ss, code, name);
      if(savedRows.length !== total) {
        return outJSON({status:"FAIL",msg:"Paper write verification failed. Expected "+total+" questions but found "+savedRows.length+" in the "+code+" sheet."});
      }
      return outJSON({status:"SUCCESS",paperCode:code,paperName:name,total:total,spreadsheetId:ss.getId(),spreadsheetName:ss.getName(),msg:"Paper uploaded successfully"});
    }
    if(String(data.action||"") === "removePaper") {
      if(!verifyAdminPassword(data.pass)) return outJSON({status:"FAIL",msg:"Wrong admin password"});
      var ss2=getQuestionSpreadsheet_(); return outJSON({status:"SUCCESS",removed:removePaper_(ss2,cleanPaperCode_(data.paperCode),String(data.paperName||"").trim())});
    }
    return outJSON({status:"FAIL",msg:"Unknown action"});
  } catch(err) { return outJSON({status:"FAIL",msg:err.message}); }
}


