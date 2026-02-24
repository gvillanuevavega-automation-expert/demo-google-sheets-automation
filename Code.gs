/**
 * Google Sheets Automation Demo
 * Author: Gabriel Villanueva — Automation Engineer
 * 
 * This script demonstrates common automation patterns:
 * 1. Auto-format new rows when data is entered
 * 2. Send email alerts when a threshold is hit
 * 3. Auto-generate a summary report on demand
 * 4. Clean and deduplicate data with one click
 */

// ─── CONFIG ─────────────────────────────────────────────────────────────────
const CONFIG = {
  SHEET_NAME: "Data",
  SUMMARY_SHEET: "Summary",
  ALERT_EMAIL: Session.getActiveUser().getEmail(),
  THRESHOLD_COLUMN: 4,    // Column D — e.g. "Amount"
  THRESHOLD_VALUE: 1000,  // Alert when value exceeds this
};

// ─── TRIGGER: Auto-format when a row is edited ───────────────────────────────
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAME) return;

  const row = e.range.getRow();
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(row, 1, 1, lastCol);

  // Alternate row colors for readability
  const color = row % 2 === 0 ? "#F8F9FA" : "#FFFFFF";
  range.setBackground(color);

  // Bold the header row
  if (row === 1) {
    range.setFontWeight("bold").setBackground("#4A90D9").setFontColor("#FFFFFF");
  }

  // Check threshold and send alert
  const val = sheet.getRange(row, CONFIG.THRESHOLD_COLUMN).getValue();
  if (typeof val === "number" && val > CONFIG.THRESHOLD_VALUE) {
    sendThresholdAlert(row, val);
  }
}

// ─── EMAIL ALERT ─────────────────────────────────────────────────────────────
function sendThresholdAlert(row, value) {
  const subject = `⚠️ Alert: Value exceeded threshold (Row ${row})`;
  const body = `
A value of ${value} was entered in row ${row}, exceeding your threshold of ${CONFIG.THRESHOLD_VALUE}.

Open your sheet to review: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}
  `.trim();

  MailApp.sendEmail(CONFIG.ALERT_EMAIL, subject, body);
}

// ─── SUMMARY REPORT ──────────────────────────────────────────────────────────
function generateSummaryReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  // Create or clear the summary sheet
  let summarySheet = ss.getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(CONFIG.SUMMARY_SHEET);
  } else {
    summarySheet.clearContents();
  }

  const data = dataSheet.getDataRange().getValues();
  if (data.length < 2) {
    summarySheet.getRange("A1").setValue("No data to summarize.");
    return;
  }

  const headers = data[0];
  const rows = data.slice(1);

  // Find numeric columns and compute totals
  const summary = [["Column", "Count", "Sum", "Average", "Max", "Min"]];

  headers.forEach((header, colIndex) => {
    const vals = rows
      .map(r => r[colIndex])
      .filter(v => typeof v === "number" && !isNaN(v));

    if (vals.length > 0) {
      const sum = vals.reduce((a, b) => a + b, 0);
      summary.push([
        header,
        vals.length,
        sum.toFixed(2),
        (sum / vals.length).toFixed(2),
        Math.max(...vals),
        Math.min(...vals),
      ]);
    }
  });

  summarySheet.getRange(1, 1, summary.length, summary[0].length).setValues(summary);
  summarySheet.getRange(1, 1, 1, summary[0].length)
    .setFontWeight("bold")
    .setBackground("#4A90D9")
    .setFontColor("#FFFFFF");
  summarySheet.autoResizeColumns(1, summary[0].length);

  SpreadsheetApp.getUi().alert("✅ Summary report generated in the 'Summary' sheet!");
}

// ─── DEDUPLICATION ───────────────────────────────────────────────────────────
function removeDuplicates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const seen = new Set();
  const clean = [];

  data.forEach((row, i) => {
    const key = row.join("|");
    if (i === 0 || !seen.has(key)) {
      seen.add(key);
      clean.push(row);
    }
  });

  sheet.clearContents();
  sheet.getRange(1, 1, clean.length, clean[0].length).setValues(clean);

  const removed = data.length - clean.length;
  SpreadsheetApp.getUi().alert(`✅ Done! Removed ${removed} duplicate row(s).`);
}

// ─── CUSTOM MENU ─────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🔧 Automate")
    .addItem("Generate Summary Report", "generateSummaryReport")
    .addItem("Remove Duplicates", "removeDuplicates")
    .addSeparator()
    .addItem("About this automation", "showAbout")
    .addToUi();
}

function showAbout() {
  SpreadsheetApp.getUi().alert(
    "Built by Gabriel Villanueva\nAutomation Engineer\ngvillanuevavega.26@gmail.com\n\nNeed something custom? Let's talk."
  );
}
