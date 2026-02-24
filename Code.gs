/**
 * Google Sheets Automation — Clean Working Version
 * Author: Gabriel Villanueva · gvillanuevavega.26@gmail.com
 *
 * HOW TO INSTALL:
 * 1. Open your Google Sheet
 * 2. Extensions → Apps Script → paste this entire file
 * 3. Save (Ctrl+S)
 * 4. Reload your Sheet → a "🔧 Automate" menu appears
 * 5. For email alerts: run "Setup Email Alerts" from the menu once
 */

// ─── CONFIG — edit these values ─────────────────────────────────────────────
var SHEET_NAME      = "Data";        // The sheet tab you want to automate
var SUMMARY_SHEET   = "Summary";     // Where the summary report goes
var ALERT_COLUMN    = 4;             // Column number to watch (D = 4)
var ALERT_THRESHOLD = 1000;          // Send alert when this value is exceeded
var ALERT_EMAIL     = "";            // Leave blank to auto-detect your email
// ────────────────────────────────────────────────────────────────────────────


/**
 * Runs when the sheet is opened.
 * Adds the custom Automate menu.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🔧 Automate")
    .addItem("📊 Generate Summary Report", "generateSummaryReport")
    .addItem("🧹 Remove Duplicate Rows",   "removeDuplicates")
    .addItem("🎨 Format All Rows",         "formatAllRows")
    .addSeparator()
    .addItem("📧 Setup Email Alerts",      "setupEmailAlerts")
    .addItem("❌ Remove Email Alerts",     "removeEmailAlerts")
    .addSeparator()
    .addItem("ℹ️ About",                   "showAbout")
    .addToUi();
}


/**
 * Simple trigger — runs on every cell edit.
 * Only does formatting here (no auth required for simple triggers).
 */
function onEdit(e) {
  if (!e) return;

  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  var row     = e.range.getRow();
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  // Style the edited row
  styleRow(sheet, row);
}


/**
 * Applies alternating row colors to every row in the sheet.
 * Called manually from the menu or after bulk edits.
 */
function formatAllRows() {
  var sheet   = getDataSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 1 || lastCol < 1) {
    showAlert("No data found in the sheet.");
    return;
  }

  // Header row
  sheet.getRange(1, 1, 1, lastCol)
    .setBackground("#4A90D9")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Data rows — alternating colors
  for (var r = 2; r <= lastRow; r++) {
    styleRow(sheet, r);
  }

  showAlert("✅ All rows formatted.");
}


/**
 * Styles a single data row (alternating white / light gray).
 */
function styleRow(sheet, row) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  var bg    = (row % 2 === 0) ? "#F0F4FF" : "#FFFFFF";
  var range = sheet.getRange(row, 1, 1, lastCol);
  range.setBackground(bg).setFontColor("#000000").setFontWeight("normal");
}


/**
 * Generates a summary statistics sheet from all numeric columns.
 */
function generateSummaryReport() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = getDataSheet();
  var data      = dataSheet.getDataRange().getValues();

  if (data.length < 2) {
    showAlert("Not enough data — add at least one row below the header.");
    return;
  }

  var headers = data[0];
  var rows    = data.slice(1);

  // Create or clear the summary sheet
  var summarySheet = ss.getSheetByName(SUMMARY_SHEET);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(SUMMARY_SHEET);
  } else {
    summarySheet.clearContents().clearFormats();
  }

  // Build summary table
  var summary = [["Column", "Count", "Sum", "Average", "Max", "Min"]];

  for (var col = 0; col < headers.length; col++) {
    var nums = [];
    for (var row = 0; row < rows.length; row++) {
      var val = rows[row][col];
      if (typeof val === "number" && !isNaN(val)) {
        nums.push(val);
      }
    }

    if (nums.length > 0) {
      var sum = nums.reduce(function(a, b) { return a + b; }, 0);
      summary.push([
        headers[col],
        nums.length,
        Math.round(sum * 100) / 100,
        Math.round((sum / nums.length) * 100) / 100,
        Math.max.apply(null, nums),
        Math.min.apply(null, nums)
      ]);
    }
  }

  if (summary.length === 1) {
    showAlert("No numeric columns found to summarize.");
    return;
  }

  // Write table
  var range = summarySheet.getRange(1, 1, summary.length, summary[0].length);
  range.setValues(summary);

  // Style header
  summarySheet.getRange(1, 1, 1, summary[0].length)
    .setBackground("#4A90D9")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");

  summarySheet.autoResizeColumns(1, summary[0].length);

  // Switch to the summary tab
  summarySheet.activate();
  showAlert("✅ Summary report generated in the '" + SUMMARY_SHEET + "' tab.");
}


/**
 * Removes exact duplicate rows (keeps the first occurrence).
 */
function removeDuplicates() {
  var sheet   = getDataSheet();
  var data    = sheet.getDataRange().getValues();
  var seen    = {};
  var clean   = [];
  var removed = 0;

  for (var i = 0; i < data.length; i++) {
    var key = JSON.stringify(data[i]);
    if (i === 0 || !seen[key]) {
      seen[key] = true;
      clean.push(data[i]);
    } else {
      removed++;
    }
  }

  if (removed === 0) {
    showAlert("✅ No duplicate rows found.");
    return;
  }

  sheet.clearContents();
  sheet.getRange(1, 1, clean.length, clean[0].length).setValues(clean);
  formatAllRows();
  showAlert("✅ Removed " + removed + " duplicate row(s).");
}


/**
 * Installs an on-edit trigger that CAN send emails (installable trigger).
 * The user must run this once from the menu — it requires authorization.
 */
function setupEmailAlerts() {
  // Remove existing alert triggers to avoid duplicates
  removeEmailAlerts(true);

  ScriptApp.newTrigger("checkThresholdAndAlert")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  showAlert("✅ Email alerts activated. You'll get an email when column " + ALERT_COLUMN + " exceeds " + ALERT_THRESHOLD + ".");
}


/**
 * Removes the email alert trigger.
 */
function removeEmailAlerts(silent) {
  var triggers = ScriptApp.getProjectTriggers();
  var removed  = 0;

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "checkThresholdAndAlert") {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }

  if (!silent) {
    showAlert(removed > 0 ? "✅ Email alerts removed." : "No alert triggers found.");
  }
}


/**
 * Called by the installable trigger (has email permissions).
 */
function checkThresholdAndAlert(e) {
  if (!e) return;

  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  var row = e.range.getRow();
  if (row === 1) return; // skip header

  var val = sheet.getRange(row, ALERT_COLUMN).getValue();
  if (typeof val !== "number" || val <= ALERT_THRESHOLD) return;

  var email   = ALERT_EMAIL || Session.getActiveUser().getEmail();
  var ssUrl   = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var subject = "⚠️ Alert: Value $" + val + " exceeds threshold (Row " + row + ")";
  var body    = "A value of " + val + " was entered in row " + row + ", exceeding your threshold of " + ALERT_THRESHOLD + ".\n\nOpen sheet: " + ssUrl;

  MailApp.sendEmail(email, subject, body);
}


// ─── HELPERS ────────────────────────────────────────────────────────────────

function getDataSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    showAlert("Sheet '" + SHEET_NAME + "' not found. Check the SHEET_NAME variable.");
    throw new Error("Sheet not found: " + SHEET_NAME);
  }
  return sheet;
}

function showAlert(msg) {
  SpreadsheetApp.getUi().alert(msg);
}

function showAbout() {
  SpreadsheetApp.getUi().alert(
    "Google Sheets Automation\n" +
    "Built by Gabriel Villanueva\n\n" +
    "📧 gvillanuevavega.26@gmail.com\n" +
    "🔗 upwork.com/freelancers/~0120763e5542a78494\n\n" +
    "Need a custom automation? Let's talk."
  );
}
