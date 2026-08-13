// ============================================================
//  BADR EL-DIN FARMS — INVENTORY SYSTEM
//  Lives in its OWN spreadsheet ("Badr El-Din Inventory Tracker"),
//  separate from the main Order Tracker spreadsheet, so that
//  access can be controlled independently.
//  ─────────────────────────────────────────────────────────
//  This file is part of the same Apps Script project as
//  Code.gs / Router.gs / OrderBackend.gs / DeliveryBackend.gs,
//  but builds and manages a second spreadsheet.
//  ─────────────────────────────────────────────────────────
//  SHEETS:
//  • Lists               → dropdown reference data (Warehouses, Sources/Destinations)
//  • Inventory Movements → append-only ledger, signed quantities
//  • Inventory Balance   → formula-driven, one row per flavor
//
//  LOW-STOCK THRESHOLDS (Current Stock, cases):
//  • Mango, Blueberry, Lemon-Mint, Strawberry → 75
//  • Coconut, Orange                          → 25
//  • All other flavors                        → 50
// ============================================================

// ────────────────────────────────────────────────────────────
//    CREATE OR GET INVENTORY SPREADSHEET
// ────────────────────────────────────────────────────────────
const INVENTORY_SPREADSHEET_NAME = "Badr El-Din Inventory Tracker";

function _createOrGetInventorySpreadsheet() {
  const files = DriveApp.getFilesByName(INVENTORY_SPREADSHEET_NAME);
  if (files.hasNext()) {
    const existing = SpreadsheetApp.open(files.next());
    Logger.log("Found existing Inventory spreadsheet.");
    PropertiesService.getScriptProperties().setProperty("INVENTORY_SPREADSHEET_ID", existing.getId());
    return existing;
  }
  const ss = SpreadsheetApp.create(INVENTORY_SPREADSHEET_NAME);
  PropertiesService.getScriptProperties().setProperty("INVENTORY_SPREADSHEET_ID", ss.getId());
  Logger.log("✅ Created new Inventory spreadsheet: " + ss.getUrl());
  return ss;
}

// ────────────────────────────────────────────────────────────
//    GET INVENTORY SPREADSHEET
// ────────────────────────────────────────────────────────────
function _getInventorySpreadsheet() {
  const id = PropertiesService.getScriptProperties().getProperty("INVENTORY_SPREADSHEET_ID");
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (err) {
      Logger.log("⚠️ Saved INVENTORY_SPREADSHEET_ID invalid, recreating: " + err.toString());
    }
  }
  return _createOrGetInventorySpreadsheet();
}

// ────────────────────────────────────────────────────────────
//    GET MAIN CONFIG — Flavors/Distributors live in the main
//    Order Tracker spreadsheet, not the Inventory spreadsheet.
//    Used only by the four _log*Movements functions below.
// ────────────────────────────────────────────────────────────
function _getMainConfig() {
  const mainSs = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
  );
  return _getConfig(mainSs);
}

// ────────────────────────────────────────────────────────────
//    LISTS SHEET
//  • Reference data for dropdowns (Warehouses, Sources/Destinations)
//  • Manually maintained for now — future form could append rows
// ────────────────────────────────────────────────────────────
function _buildListsSheet(ss) {
  let ws = _getSheet(ss, "Lists") || ss.insertSheet(" Lists");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#A6A6A6");

  const DARK_GREY = "#595959", WHITE = "#FFFFFF";

  const headers = ["Warehouses", "Sources/Destinations"];
  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_GREY)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 30);

  // Column A — Warehouses
  ws.getRange(2, 1).setValue("Main Warehouse");

  // Column B — Sources/Destinations
  const destinations = [
    "Packing Line",
    "Samples",
    "Write-off",
    ...CONFIG.DISTRIBUTORS.map(d => d.name)
  ];
  ws.getRange(2, 2, destinations.length, 1).setValues(destinations.map(d => [d]));

  ws.setColumnWidth(1, 180);
  ws.setColumnWidth(2, 220);
  ws.setFrozenRows(1);

  Logger.log("✅ Lists sheet built.");
}

// ────────────────────────────────────────────────────────────
//    INVENTORY MOVEMENTS SHEET
//  • Append-only ledger — every stock change is one row
//  • No pre-filled formulas — all values written as static data
//  • Quantity is signed: positive = stock IN, negative = stock OUT
//  • "Opening Balance" type exists for one-time manual entry only —
//    not surfaced in any web app / form UI
// ────────────────────────────────────────────────────────────
function _buildInventoryMovementsSheet(ss) {
  let ws = _getSheet(ss, "Inventory Movements") || ss.insertSheet(" Inventory Movements");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#8B5A2B");

  const DARK_BROWN = "#5C3A1E", WHITE = "#FFFFFF", LIGHT_BROWN = "#F2E6D9";

  const headers = ["Date", "Type", "Reference", "Production Date", "Product", "Flavor", "Quantity (cases)", "Warehouse", "Source/Destination", "Notes"];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_BROWN)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 35);

  const NUM_ROWS = 1000;
  const listsSheet = _getSheet(ss, "Lists");

  // Type dropdown (col B)
  ws.getRange(2, 2, NUM_ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Opening Balance", "Production In", "Delivery Out", "Return", "Write-off"], true)
      .build()
  );

  // Product dropdown (col E)
  ws.getRange(2, 5, NUM_ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Dots"], true)
      .build()
  );

  // Flavor dropdown (col F)
  ws.getRange(2, 6, NUM_ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.FLAVORS, true)
      .build()
  );

  // Warehouse dropdown from Lists!A (col H)
  ws.getRange(2, 8, NUM_ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(listsSheet.getRange("A2:A1000"), true)
      .build()
  );

  // Source/Destination dropdown from Lists!B (col I)
  ws.getRange(2, 9, NUM_ROWS, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(listsSheet.getRange("B2:B1000"), true)
      .build()
  );

  // Column widths
  ws.setColumnWidth(1, 110);  // Date
  ws.setColumnWidth(2, 130);  // Type
  ws.setColumnWidth(3, 100);  // Reference
  ws.setColumnWidth(4, 110);  // Production Date
  ws.setColumnWidth(5, 80);   // Product
  ws.setColumnWidth(6, 130);  // Flavor
  ws.setColumnWidth(7, 130);  // Quantity (cases)
  ws.setColumnWidth(8, 140);  // Warehouse
  ws.setColumnWidth(9, 160);  // Source/Destination
  ws.setColumnWidth(10, 220); // Notes

  ws.setFrozenRows(1);

  // Alternating row colors
  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A2<>"")`)
      .setBackground(LIGHT_BROWN)
      .setRanges([ws.getRange(2, 1, NUM_ROWS, headers.length)])
      .build()
  ]);

  Logger.log("✅ Inventory Movements sheet built.");
}

// ────────────────────────────────────────────────────────────
//    INVENTORY BALANCE SHEET
//  • Formula-driven — one row per flavor (13 rows)
//  • Each column SUMIFS the Movements ledger by Type + Flavor
//  • Current Stock = sum of all movement types (signed quantities)
//  • Red flag if Current Stock is below that flavor's threshold
// ────────────────────────────────────────────────────────────
function _buildInventoryBalanceSheet(ss) {
  let ws = _getSheet(ss, "Inventory Balance") || ss.insertSheet(" Inventory Balance");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#ED7D31");

  const DARK_ORANGE = "#9C5700", WHITE = "#FFFFFF", LIGHT_ORANGE = "#FCE4D6";
  const RED_BG = "#FFC7CE", RED_FONT = "#9C0006";

  const headers = ["Flavor", "Opening Balance", "Production In", "Delivered Out", "Returns", "Write-offs", "Current Stock"];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_ORANGE)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10)
    .setWrap(true);
  ws.setRowHeight(1, 40);

  const MOV = "' Inventory Movements'";
  const flavorCol = "$F$2:$F$1000";
  const typeCol   = "$B$2:$B$1000";
  const qtyCol    = "$G$2:$G$1000";

  const typeMap = {
    2: "Opening Balance",
    3: "Production In",
    4: "Delivery Out",
    5: "Return",
    6: "Write-off"
  };

  // Low-stock thresholds
  const thresholds = {
    "Mango": 75, "Blueberry": 75, "Lemon-Mint": 75, "Strawberry": 75,
    "Coconut": 25, "Orange": 25
  };
  const DEFAULT_THRESHOLD = 50;

  CONFIG.FLAVORS.forEach((flavor, i) => {
    const r = i + 2;
    ws.getRange(r, 1).setValue(flavor);

    // Columns 2-6: SUMIFS by Type + Flavor
    for (let c = 2; c <= 6; c++) {
      const type = typeMap[c];
      ws.getRange(r, c).setFormula(
        `=SUMIFS(${MOV}!${qtyCol},${MOV}!${flavorCol},A${r},${MOV}!${typeCol},"${type}")`
      );
    }

    // Current Stock = sum of all movement columns
    ws.getRange(r, 7).setFormula(`=SUM(B${r}:F${r})`);

    // Row styling
    ws.getRange(r, 1, 1, headers.length)
      .setBackground(i % 2 === 0 ? LIGHT_ORANGE : WHITE)
      .setFontFamily("Arial").setFontSize(10).setHorizontalAlignment("center");
    ws.getRange(r, 1).setHorizontalAlignment("left").setFontWeight("bold");
  });

  // Column widths
  ws.setColumnWidth(1, 130);
  for (let c = 2; c <= 7; c++) ws.setColumnWidth(c, 130);
  ws.setFrozenRows(1);

  // Low-stock conditional formatting — per-flavor threshold on Current Stock (col G)
  const rules = [];
  CONFIG.FLAVORS.forEach((flavor, i) => {
    const r = i + 2;
    const threshold = thresholds[flavor] !== undefined ? thresholds[flavor] : DEFAULT_THRESHOLD;
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(threshold)
        .setBackground(RED_BG)
        .setFontColor(RED_FONT)
        .setRanges([ws.getRange(r, 7)])
        .build()
    );
  });
  ws.setConditionalFormatRules(rules);

  Logger.log("✅ Inventory Balance sheet built.");
}

// ────────────────────────────────────────────────────────────
//    MAIN ENTRY POINT — Inventory system setup
//  • Run this once to build the Inventory spreadsheet + all sheets
//  • Safe to re-run: rebuilds all three sheets, does not touch
//    the main Order Tracker spreadsheet
// ────────────────────────────────────────────────────────────
function setupInventorySystem() {
  Logger.log("Starting Inventory system setup...");
  const ss = _getInventorySpreadsheet();
  _buildListsSheet(ss);
  _buildInventoryMovementsSheet(ss);
  _buildInventoryBalanceSheet(ss);
  _removeDefaultSheet(ss);
  Logger.log("✅ Inventory setup complete!");
  Logger.log("Inventory spreadsheet: " + ss.getUrl());
}

// ────────────────────────────────────────────────────────────
//    LOG DELIVERY OUT MOVEMENTS
//  • Called by submitDelivery() after a delivery is logged
//  • Appends one "Delivery Out" row per flavor with qty > 0
//  • Quantity is negative (stock leaving the warehouse)
// ────────────────────────────────────────────────────────────
function _logDeliveryOutMovements(orderId, distributor, qtys) {
  const ss = _getInventorySpreadsheet();
  const ws = _getSheet(ss, "Inventory Movements");
  const now = new Date();
  const config = _getMainConfig();

  const rows = [];
  config.flavors.forEach((flavor, i) => {
    const qty = Number(qtys[i]) || 0;
    if (qty > 0) {
      rows.push([now, "Delivery Out", orderId, "", "Dots", flavor, -qty, "Main Warehouse", distributor, ""]);
    }
  });

  if (rows.length === 0) return;

  const startRow = ws.getLastRow() + 1;
  ws.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  for (let i = 0; i < rows.length; i++) {
    ws.getRange(startRow + i, 1).setNumberFormat("yyyy-mm-dd hh:mm");
  }

  Logger.log(`✅ Logged ${rows.length} Delivery Out movement(s) for order ${orderId}`);
}

// ────────────────────────────────────────────────────────────
//    LOG RETURN IN MOVEMENTS
//  • Called by submitReturn() after a return is logged
//  • Appends one "Return" row per flavor with qty > 0
//  • Quantity is positive (stock coming back into the warehouse)
// ────────────────────────────────────────────────────────────
function _logReturnInMovements(returnId, payload) {
  const ss = _getInventorySpreadsheet();
  const ws = _getSheet(ss, "Inventory Movements");
  const now = new Date();
  const config = _getMainConfig();

  const rows = [];
  config.flavors.forEach((flavor, i) => {
    const qty = Number(payload.qtys[i]) || 0;
    if (qty > 0) {
      rows.push([now, "Return", returnId, "", "Dots", flavor, qty, "Main Warehouse", payload.distributor, ""]);
    }
  });

  if (rows.length === 0) return;

  const startRow = ws.getLastRow() + 1;
  ws.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  for (let i = 0; i < rows.length; i++) {
    ws.getRange(startRow + i, 1).setNumberFormat("yyyy-mm-dd hh:mm");
  }

  Logger.log(`✅ Logged ${rows.length} Return movement(s) for ${returnId}`);
}

// ────────────────────────────────────────────────────────────
//    LOG PRODUCTION IN MOVEMENTS
//  • Called by submitProduction() after a production batch is logged
//  • Appends one "Production In" row per flavor with qty > 0
//  • Quantity is positive (stock entering the warehouse)
//  • Production Date is populated from the payload — the first
//    transaction type to write a real date into that column
// ────────────────────────────────────────────────────────────
function _logProductionInMovements(productionId, payload) {
  const ss = _getInventorySpreadsheet();
  const ws = _getSheet(ss, "Inventory Movements");
  const now = new Date();
  const productionDate = payload.productionDate ? new Date(payload.productionDate) : now;
  const config = _getMainConfig();

  const rows = [];
  config.flavors.forEach((flavor, i) => {
    const qty = Number(payload.qtys[i]) || 0;
    if (qty > 0) {
      rows.push([now, "Production In", productionId, productionDate, "Dots", flavor, qty, "Main Warehouse", "Packing Line", ""]);
    }
  });

  if (rows.length === 0) return;

  const startRow = ws.getLastRow() + 1;
  ws.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  for (let i = 0; i < rows.length; i++) {
    ws.getRange(startRow + i, 1).setNumberFormat("yyyy-mm-dd hh:mm");
    ws.getRange(startRow + i, 4).setNumberFormat("yyyy-mm-dd");
  }

  Logger.log(`✅ Logged ${rows.length} Production In movement(s) for ${productionId}`);
}

// ────────────────────────────────────────────────────────────
//    LOG WRITE-OFF MOVEMENTS
//  • Called directly from Router doPost action=submitWriteOff —
//    no separate Backend file or main-spreadsheet sheet exists for
//    this transaction type, so this function IS the entire write.
//  • Quantity is negative (stock leaving the warehouse)
//  • WOF-XXX generated via Script Properties counter (_nextSequentialId
//    in Code.gs), not by counting existing rows — counting rows would
//    break if Inventory Movements rows are ever manually edited/deleted
//  • reasonOrReference is the team's required free-text reason —
//    written to Notes. The ledger's own Reference column holds the
//    auto-generated WOF-XXX, consistent with how Delivery Out/Return/
//    Production In all use their own ID in that column.
//  • Unlike _logDeliveryOutMovements/_logReturnInMovements/
//    _logProductionInMovements, this is NOT wrapped in a non-blocking
//    try/catch by its caller — there is no separate primary transaction
//    to protect, so failures must propagate (Router's outer catch turns
//    them into success:false), not get silently absorbed via _logSyncError
// ────────────────────────────────────────────────────────────
function _logWriteOffMovements(reasonOrReference, productionDate, qtys) {
  const ss = _getInventorySpreadsheet();
  const ws = _getSheet(ss, "Inventory Movements");
  const now = new Date();
  const writeOffId = _nextSequentialId("WOF_COUNTER", "WOF-");
  const prodDate = productionDate ? new Date(productionDate) : "";
  const config = _getMainConfig();

  const rows = [];
  config.flavors.forEach((flavor, i) => {
    const qty = Number(qtys[i]) || 0;
    if (qty > 0) {
      rows.push([now, "Write-off", writeOffId, prodDate, "Dots", flavor, -qty, "Main Warehouse", "Write-off", reasonOrReference || ""]);
    }
  });

  if (rows.length === 0) {
    throw new Error("Write-off submitted with no flavor quantities.");
  }

  const startRow = ws.getLastRow() + 1;
  ws.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  for (let i = 0; i < rows.length; i++) {
    ws.getRange(startRow + i, 1).setNumberFormat("yyyy-mm-dd hh:mm");
    if (prodDate) ws.getRange(startRow + i, 4).setNumberFormat("yyyy-mm-dd");
  }

  const total = rows.reduce((s, r) => s + Math.abs(r[6]), 0);
  Logger.log(`✅ Logged ${rows.length} Write-off movement(s) — ${writeOffId}`);
  return { writeOffId: writeOffId, total: total };
}