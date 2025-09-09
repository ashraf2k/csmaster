// === Google Apps Script backend for EduPlan ===
// How to use:
// 1) In your Google Sheet, open Extensions → Apps Script.
// 2) Paste this entire file into Code.gs (replace any default code).
// 3) File → Save. Then Deploy → Manage deployments → New deployment → type: Web app.
//    - Execute as: Me (your account)
//    - Who has access: Anyone
// 4) Copy the Web app URL and put it into WEB_APP_URL in config.js.
// 5) Make sure your Sheet has a sheet/tab named "Plans" (or change SHEET_TAB_NAME in config).
//    Create this header row in row 1 exactly:
//    Week | Day | Lesson | Subject | Class | Materials | Textbook Pages | Quiz | ID | Saved At
//
// Notes:
// - All responses are JSON. For GET requests, use query params. For POST, send JSON body.
// - We filter and update rows by the unique `ID` value saved by the front-end.
// - This script assumes a simple, single-sheet model with a fixed header order.

function doGet(e) {
  try {
    var action = (e.parameter.action || "").toLowerCase();
    var sheetId = e.parameter.sheetId;
    var sheetName = e.parameter.sheetName || "Plans";
    if (!sheetId) throw new Error("sheetId is required");

    var ss = SpreadsheetApp.openById(sheetId);
    var sh = ss.getSheetByName(sheetName);
    if (!sh) throw new Error("Sheet/tab '" + sheetName + "' not found");

    if (action === "health") {
      return jsonOutput({ success: true, message: "OK" });
    }

    if (action === "list") {
      var rows = readAllRows_(sh);
      return jsonOutput({ success: true, rows: rows });
    }

    if (action === "byweekclass") {
      var week = String(e.parameter.week || "");
      var cls  = String(e.parameter.class || "");
      if (!week || !cls) throw new Error("week and class are required");
      var rows = readAllRows_(sh).filter(function(r){
        return String(r.Week) === week && String(r.Class) === cls;
      });
      return jsonOutput({ success: true, rows: rows });
    }

    return jsonOutput({ success: false, error: "Unknown or missing action" });
  } catch (err) {
    return jsonOutput({ success: false, error: String(err) }, 500);
  }
}

function doPost(e) {
  try {
    var body = e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    var action = (body.action || "").toLowerCase();
    var sheetId = body.sheetId;
    var sheetName = body.sheetName || "Plans";
    if (!sheetId) throw new Error("sheetId is required");

    var ss = SpreadsheetApp.openById(sheetId);
    var sh = ss.getSheetByName(sheetName);
    if (!sh) throw new Error("Sheet/tab '" + sheetName + "' not found");

    if (action === "append") {
      var row = normalizeRow_(body.row || {});
      sh.appendRow([row.Week, row.Day, row.Lesson, row.Subject, row.Class, row.Materials, row.TextbookPages, row.Quiz, row.ID, row.SavedAt]);
      return jsonOutput({ success: true });
    }

    if (action === "update") {
      var upd = normalizeRow_(body.row || {});
      var idToFind = String(upd.ID || body.id || "");
      if (!idToFind) throw new Error("ID is required for update");
      var rIndex = findRowIndexById_(sh, idToFind);
      if (rIndex <= 0) throw new Error("ID not found");
      sh.getRange(rIndex, 1, 1, 10).setValues([[upd.Week, upd.Day, upd.Lesson, upd.Subject, upd.Class, upd.Materials, upd.TextbookPages, upd.Quiz, upd.ID, upd.SavedAt]]);
      return jsonOutput({ success: true });
    }

    if (action === "delete") {
      var id = String(body.id || "");
      if (!id) throw new Error("id is required for delete");
      var rowIndex = findRowIndexById_(sh, id);
      if (rowIndex <= 0) throw new Error("ID not found");
      sh.deleteRow(rowIndex);
      return jsonOutput({ success: true });
    }

    return jsonOutput({ success: false, error: "Unknown or missing action" });
  } catch (err) {
    return jsonOutput({ success: false, error: String(err) }, 500);
  }
}

// === Helpers ===

function jsonOutput(obj, status) {
  status = status || 200;
  var out = ContentService.createTextOutput(JSON.stringify(obj));
  out.setMimeType(ContentService.MimeType.JSON);
  // NOTE: Apps Script doesn't let you set arbitrary CORS headers here; Web Apps generally work with fetch when deployed publicly.
  return out;
}

function readAllRows_(sh) {
  var range = sh.getDataRange();
  var values = range.getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var out = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i] || values[i].length === 0) continue;
    var rowObj = {};
    for (var j = 0; j < headers.length; j++) {
      rowObj[String(headers[j])] = values[i][j];
    }
    out.push(rowObj);
  }
  return out;
}

function normalizeRow_(r) {
  return {
    Week: String(r.Week || r.weekNumber || ""),
    Day: String(r.Day || r.lessonDay || ""),
    Lesson: Number(r.Lesson || r.lessonNumber || 0),
    Subject: String(r.Subject || r.subjectName || ""),
    Class: String(r.Class || r.classNumber || ""),
    Materials: String(r.Materials || r.materials || ""),
    TextbookPages: String(r.TextbookPages || r.textbookPages || ""),
    Quiz: String(r.Quiz || r.quiz || ""),
    ID: String(r.ID || r.id || ""),
    SavedAt: String(r.SavedAt || r.savedAt || new Date().toISOString())
  };
}

function findRowIndexById_(sh, id) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sh.getRange(2, 9, lastRow - 1, 1).getValues(); // Column I is 9th (ID)
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      return i + 2; // add header row offset
    }
  }
  return -1;
}
