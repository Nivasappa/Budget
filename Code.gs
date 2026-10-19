/**
 * Dheeran's Fund - Google Sheet backend
 *
 * Deploy this as a Web App (Deploy > New deployment > Web app,
 * Execute as: Me, Who has access: Anyone with the link) and paste
 * the resulting URL into the app's "Sync" tab.
 *
 * Expects three sheet tabs, named exactly:
 *   Income        - one header row, one data row
 *   Expenses      - one header row, one row per expense
 *   Investments   - one header row, one row per instrument
 */

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var payload = {
    income: readIncome(ss),
    expenses: readTable(ss.getSheetByName("Expenses")),
    investments: readTable(ss.getSheetByName("Investments")),
  };
  return jsonResponse(payload);
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = JSON.parse(e.postData.contents);

  if (data.income) writeIncome(ss, data.income);
  if (data.expenses) writeTable(ss.getSheetByName("Expenses"), data.expenses);
  if (data.investments) writeTable(ss.getSheetByName("Investments"), data.investments);

  return jsonResponse({ ok: true });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// Income - single row of named fields
// ---------------------------------------------------------------------------
function readIncome(ss) {
  var sheet = ss.getSheetByName("Income");
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var values = sheet.getRange(2, 1, 1, headers.length).getValues()[0];
  var obj = {};
  headers.forEach(function (h, i) { obj[h] = values[i]; });
  return obj;
}

function writeIncome(ss, income) {
  var sheet = ss.getSheetByName("Income");
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = headers.map(function (h) {
    return income[h] !== undefined ? income[h] : "";
  });
  sheet.getRange(2, 1, 1, headers.length).setValues([row]);
}

// ---------------------------------------------------------------------------
// Expenses / Investments - one row per record, matched by header name
// ---------------------------------------------------------------------------
function readTable(sheet) {
  var values = sheet.getDataRange().getValues();
  var headers = values[0];
  return values.slice(1)
    .filter(function (row) { return row[0] !== "" && row[0] !== null; })
    .map(function (row) {
      var obj = {};
      headers.forEach(function (h, i) {
        var val = row[i];
        var upper = String(val).toUpperCase();
        if (val === true || upper === "TRUE") val = true;
        else if (val === false || upper === "FALSE") val = false;
        obj[h] = val;
      });
      return obj;
    });
}

function writeTable(sheet, rows) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  var out = rows.map(function (r) {
    return headers.map(function (h) {
      return r[h] !== undefined ? r[h] : "";
    });
  });
  if (out.length) {
    sheet.getRange(2, 1, out.length, headers.length).setValues(out);
  }
}
