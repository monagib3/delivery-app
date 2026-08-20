// ============================================================
//  BADR EL-DIN FARMS — PRODUCTION LIFECYCLE SYSTEM
//  Lives in its OWN spreadsheet ("Badr El-Din Production Tracker"),
//  separate from the main Order Tracker and Inventory Tracker
//  spreadsheets, so that access can be controlled independently.
//  ─────────────────────────────────────────────────────────
//  This file is part of the same Apps Script project as
//  Code.gs / Router.gs / ...Backend.gs / Inventory.gs, but
//  builds and manages a third spreadsheet.
//  ─────────────────────────────────────────────────────────
//  SHEETS (Sub-task PT1 — schema only, no backend logic yet):
//  • Production_Requests      → header row per request
//  • Production_Request_Lines → one row per non-zero flavor per request
//  • Production_Runs          → header row per production run
//  • Production_Run_Lines     → one row per non-zero flavor per run
//  • QC_Records                → one row per QC check against a run
// ============================================================

// ────────────────────────────────────────────────────────────
//    CREATE OR GET PRODUCTION TRACKER SPREADSHEET
// ────────────────────────────────────────────────────────────
const PRODUCTION_SPREADSHEET_NAME = "Badr El-Din Production Tracker";

function _createOrGetProductionSpreadsheet() {
  const files = DriveApp.getFilesByName(PRODUCTION_SPREADSHEET_NAME);
  if (files.hasNext()) {
    const existing = SpreadsheetApp.open(files.next());
    Logger.log("Found existing Production Tracker spreadsheet.");
    PropertiesService.getScriptProperties().setProperty("PRODUCTION_SPREADSHEET_ID", existing.getId());
    return existing;
  }
  const ss = SpreadsheetApp.create(PRODUCTION_SPREADSHEET_NAME);
  PropertiesService.getScriptProperties().setProperty("PRODUCTION_SPREADSHEET_ID", ss.getId());
  Logger.log("✅ Created new Production Tracker spreadsheet: " + ss.getUrl());
  return ss;
}

// ────────────────────────────────────────────────────────────
//    GET PRODUCTION TRACKER SPREADSHEET
// ────────────────────────────────────────────────────────────
function _getProductionSpreadsheet() {
  const id = PropertiesService.getScriptProperties().getProperty("PRODUCTION_SPREADSHEET_ID");
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (err) {
      Logger.log("⚠️ Saved PRODUCTION_SPREADSHEET_ID invalid, recreating: " + err.toString());
    }
  }
  return _createOrGetProductionSpreadsheet();
}

// ────────────────────────────────────────────────────────────
//  🏭  PRODUCTION_REQUESTS SHEET — header row per request
// ────────────────────────────────────────────────────────────
function _buildProductionRequestsSheet(ss) {
  let ws = _getSheet(ss, "🏭 Production_Requests") || ss.insertSheet("🏭 Production_Requests");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#31859C");

  const DARK_TEAL = "#1B4F5C", WHITE = "#FFFFFF", LIGHT_TEAL = "#D6EAF0";
  const headers = ["Request_ID", "Timestamp", "Requested_By", "Status", "Notes"];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_TEAL)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 35);

  // Status dropdown — rows 2–501
  ws.getRange(2, 4, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Requested", "Fulfilled"], true)
      .build()
  );

  ws.setColumnWidth(1, 120);
  ws.setColumnWidth(2, 150);
  ws.setColumnWidth(3, 160);
  ws.setColumnWidth(4, 130);
  ws.setColumnWidth(5, 200);

  ws.setFrozenRows(1);
  ws.setFrozenColumns(2);

  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A2<>"")`)
      .setBackground(LIGHT_TEAL)
      .setRanges([ws.getRange(2, 1, 500, headers.length)])
      .build()
  ]);

  Logger.log("✅ Production_Requests sheet built.");
}

// ────────────────────────────────────────────────────────────
//  🏭  PRODUCTION_REQUEST_LINES SHEET — one row per flavor per request
// ────────────────────────────────────────────────────────────
function _buildProductionRequestLinesSheet(ss) {
  let ws = _getSheet(ss, "🏭 Production_Request_Lines") || ss.insertSheet("🏭 Production_Request_Lines");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#31859C");

  const DARK_TEAL = "#1B4F5C", WHITE = "#FFFFFF", LIGHT_TEAL = "#D6EAF0";
  const headers = ["Line_ID", "Request_ID", "Flavor_ID", "Flavor_Name", "Qty_Requested"];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_TEAL)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 35);

  ws.setColumnWidth(1, 110);
  ws.setColumnWidth(2, 120);
  ws.setColumnWidth(3, 100);
  ws.setColumnWidth(4, 150);
  ws.setColumnWidth(5, 130);

  ws.setFrozenRows(1);
  ws.setFrozenColumns(2);

  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A2<>"")`)
      .setBackground(LIGHT_TEAL)
      .setRanges([ws.getRange(2, 1, 500, headers.length)])
      .build()
  ]);

  Logger.log("✅ Production_Request_Lines sheet built.");
}

// ────────────────────────────────────────────────────────────
//  🏭  PRODUCTION_RUNS SHEET — header row per production run
// ────────────────────────────────────────────────────────────
function _buildProductionRunsSheet(ss) {
  let ws = _getSheet(ss, "🏭 Production_Runs") || ss.insertSheet("🏭 Production_Runs");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#BF9000");

  const DARK_GOLD = "#7F6000", WHITE = "#FFFFFF", LIGHT_GOLD = "#FFF2CC";
  const headers = ["Run_ID", "Request_ID", "Date", "Lot_Number", "Total_Produced", "QC_Status", "Notes", "Inventory_Pulled_By", "Inventory_Pulled_Timestamp"];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_GOLD)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 35);

  // QC_Status dropdown — rows 2–501, col 6
  ws.getRange(2, 6, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Pending", "Approved", "Rejected"], true)
      .build()
  );

  ws.setColumnWidth(1, 110);
  ws.setColumnWidth(2, 120);
  ws.setColumnWidth(3, 120);
  ws.setColumnWidth(4, 140);
  ws.setColumnWidth(5, 130);
  ws.setColumnWidth(6, 110);
  ws.setColumnWidth(7, 200);
  ws.setColumnWidth(8, 150);
  ws.setColumnWidth(9, 170);

  ws.setFrozenRows(1);
  ws.setFrozenColumns(2);

  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A2<>"")`)
      .setBackground(LIGHT_GOLD)
      .setRanges([ws.getRange(2, 1, 500, headers.length)])
      .build()
  ]);

  Logger.log("✅ Production_Runs sheet built.");
}

// ────────────────────────────────────────────────────────────
//  🏭  PRODUCTION_RUN_LINES SHEET — one row per flavor per run
// ────────────────────────────────────────────────────────────
function _buildProductionRunLinesSheet(ss) {
  let ws = _getSheet(ss, "🏭 Production_Run_Lines") || ss.insertSheet("🏭 Production_Run_Lines");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#BF9000");

  const DARK_GOLD = "#7F6000", WHITE = "#FFFFFF", LIGHT_GOLD = "#FFF2CC";
  const headers = ["Line_ID", "Run_ID", "Flavor_ID", "Flavor_Name", "Qty_Produced"];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_GOLD)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 35);

  ws.setColumnWidth(1, 110);
  ws.setColumnWidth(2, 110);
  ws.setColumnWidth(3, 100);
  ws.setColumnWidth(4, 150);
  ws.setColumnWidth(5, 130);

  ws.setFrozenRows(1);
  ws.setFrozenColumns(2);

  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A2<>"")`)
      .setBackground(LIGHT_GOLD)
      .setRanges([ws.getRange(2, 1, 500, headers.length)])
      .build()
  ]);

  Logger.log("✅ Production_Run_Lines sheet built.");
}

// ────────────────────────────────────────────────────────────
//  🏭  QC_RECORDS SHEET — one row per QC check against a run
// ────────────────────────────────────────────────────────────
function _buildQCRecordsSheet(ss) {
  let ws = _getSheet(ss, "🏭 QC_Records") || ss.insertSheet("🏭 QC_Records");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#674EA7");

  const DARK_PURPLE = "#351C75", WHITE = "#FFFFFF", LIGHT_PURPLE = "#D9D2E9";
  const headers = ["QC_ID", "Run_ID", "Timestamp", "QC_By", "Decision", "Notes"];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_PURPLE)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 35);

  // Decision dropdown — rows 2–501, col 5. Matches Production_Runs.QC_Status
  // vocabulary since QC_Records.Decision is the per-check mirror of it.
  ws.getRange(2, 5, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Pending", "Approved", "Rejected"], true)
      .build()
  );

  ws.setColumnWidth(1, 100);
  ws.setColumnWidth(2, 110);
  ws.setColumnWidth(3, 150);
  ws.setColumnWidth(4, 150);
  ws.setColumnWidth(5, 110);
  ws.setColumnWidth(6, 200);

  ws.setFrozenRows(1);
  ws.setFrozenColumns(2);

  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A2<>"")`)
      .setBackground(LIGHT_PURPLE)
      .setRanges([ws.getRange(2, 1, 500, headers.length)])
      .build()
  ]);

  Logger.log("✅ QC_Records sheet built.");
}

// ────────────────────────────────────────────────────────────
//  🏭  PENDING-QUEUE READS — Sub-task PT6
//  Read-only, self-contained (open both spreadsheets themselves,
//  same style as _notifyQCDecision in QualityBackend.js) so
//  Router.js stays a pure dispatcher: call one of these, spread the
//  result into the JSON response, no logic in Router.js itself —
//  same precedent as DashboardBackend.js's _getDashboardData(ss).
//  Each returns an already-shaped { flavors, <list> } object.
//  Per-flavor qtys are joined from the matching *_Lines table using
//  the same map-then-iterate pattern as _getOpenOrders in
//  DeliveryBackend.js — not a new pattern.
// ────────────────────────────────────────────────────────────

// Production Engineer's queue: requests with no run logged against them yet
function _getPendingProductionRequests() {
  const mainSs = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
  );
  const config = _getConfig(mainSs);

  const prodSs    = _getProductionSpreadsheet();
  const reqSheet  = _getSheet(prodSs, "🏭 Production_Requests");
  const lineSheet = _getSheet(prodSs, "🏭 Production_Request_Lines");
  const reqData   = reqSheet.getDataRange().getValues();
  const lineData  = lineSheet.getDataRange().getValues().slice(1);

  // Request_ID → { Flavor_Name → Qty_Requested }
  const linesMap = {};
  lineData.forEach(r => {
    const [, requestId, , flavorName, qty] = r;
    if (!linesMap[requestId]) linesMap[requestId] = {};
    linesMap[requestId][flavorName] = (linesMap[requestId][flavorName] || 0) + (Number(qty) || 0);
  });

  const pendingRequests = [];
  for (let i = 1; i < reqData.length; i++) {
    const row       = reqData[i];
    const requestId = row[0];
    const status    = row[3];
    if (!requestId) continue;
    if (status !== "Requested") continue;

    const rawTimestamp = row[1];
    const timestamp = rawTimestamp instanceof Date
      ? Utilities.formatDate(rawTimestamp, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")
      : String(rawTimestamp);

    const lines = linesMap[requestId] || {};
    const qtys  = config.flavors.map(f => lines[f] || 0);
    const total = qtys.reduce((s, q) => s + q, 0);

    pendingRequests.push({
      requestId:   String(requestId),
      requestedBy: String(row[2]),
      timestamp:   timestamp,
      notes:       row[4] || "",
      qtys:        qtys,
      total:       total
    });
  }

  return { flavors: config.flavors, pendingRequests: pendingRequests };
}

// Quality Engineer's queue: runs awaiting a QC decision
function _getPendingQCRuns() {
  const mainSs = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
  );
  const config = _getConfig(mainSs);

  const prodSs    = _getProductionSpreadsheet();
  const runSheet  = _getSheet(prodSs, "🏭 Production_Runs");
  const lineSheet = _getSheet(prodSs, "🏭 Production_Run_Lines");
  const runData   = runSheet.getDataRange().getValues();
  const lineData  = lineSheet.getDataRange().getValues().slice(1);

  // Run_ID → { Flavor_Name → Qty_Produced }
  const linesMap = {};
  lineData.forEach(r => {
    const [, runId, , flavorName, qty] = r;
    if (!linesMap[runId]) linesMap[runId] = {};
    linesMap[runId][flavorName] = (linesMap[runId][flavorName] || 0) + (Number(qty) || 0);
  });

  const pendingRuns = [];
  for (let i = 1; i < runData.length; i++) {
    const row      = runData[i];
    const runId    = row[0];
    const qcStatus = row[5];
    if (!runId) continue;
    if (qcStatus !== "Pending") continue;

    const rawDate = row[2];
    const date = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(rawDate).substring(0, 10);

    const lines = linesMap[runId] || {};
    const qtys  = config.flavors.map(f => lines[f] || 0);
    const total = qtys.reduce((s, q) => s + q, 0);

    pendingRuns.push({
      runId:     String(runId),
      requestId: String(row[1]),
      date:      date,
      lotNumber: String(row[3]),
      notes:     row[6] || "",
      qtys:      qtys,
      total:     total
    });
  }

  return { flavors: config.flavors, pendingRuns: pendingRuns };
}

// Inventory team's queue: Approved runs not yet pulled into the warehouse
function _getPendingInventoryPulls() {
  const mainSs = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
  );
  const config = _getConfig(mainSs);

  const prodSs    = _getProductionSpreadsheet();
  const runSheet  = _getSheet(prodSs, "🏭 Production_Runs");
  const lineSheet = _getSheet(prodSs, "🏭 Production_Run_Lines");
  const runData   = runSheet.getDataRange().getValues();
  const lineData  = lineSheet.getDataRange().getValues().slice(1);

  // Run_ID → { Flavor_Name → Qty_Produced }
  const linesMap = {};
  lineData.forEach(r => {
    const [, runId, , flavorName, qty] = r;
    if (!linesMap[runId]) linesMap[runId] = {};
    linesMap[runId][flavorName] = (linesMap[runId][flavorName] || 0) + (Number(qty) || 0);
  });

  const pendingPulls = [];
  for (let i = 1; i < runData.length; i++) {
    const row      = runData[i];
    const runId    = row[0];
    const qcStatus = row[5];
    const pulledBy = row[7];
    if (!runId) continue;
    if (qcStatus !== "Approved" || pulledBy) continue;

    const rawDate = row[2];
    const date = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(rawDate).substring(0, 10);

    const lines = linesMap[runId] || {};
    const qtys  = config.flavors.map(f => lines[f] || 0);
    const total = qtys.reduce((s, q) => s + q, 0);

    pendingPulls.push({
      runId:     String(runId),
      requestId: String(row[1]),
      date:      date,
      lotNumber: String(row[3]),
      notes:     row[6] || "",
      qtys:      qtys,
      total:     total
    });
  }

  return { flavors: config.flavors, pendingPulls: pendingPulls };
}
