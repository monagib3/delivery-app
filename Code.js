// ============================================================
//  BADR EL-DIN FARMS — ORDER TRACKING SYSTEM
//  Version 2.0 — Final Clean Build
//  ─────────────────────────────────────────────────────────
//  HOW TO USE:
//  1. Fill in CONFIG below with your real phone numbers
//  2. Run setupEntireSystem() once to build everything
//  3. Share the form URL from the execution log with distributors
//  4. When CallMeBot API keys arrive, update CONFIG and re-run
//     only _setupFormTrigger() — do NOT re-run setupEntireSystem()
// ============================================================

// ────────────────────────────────────────────────────────────
//  ⚙️  CONFIGURATION — EDIT THESE BEFORE RUNNING
// ────────────────────────────────────────────────────────────
const CONFIG = {

  // Internal team WhatsApp (CallMeBot registered)
  // Format: international without + e.g. "201012345678"
  TEAM_WHATSAPP: [
    { name: "Operations Manager", phone: "201064532222", apiKey: "2973321" },
    { name: "SC - Coordinator", phone: "201012215954", apiKey: "5286750" },
  ],

  // Distributor WhatsApp (CallMeBot registered)
  DISTRIBUTORS: [
    { name: "Family - Alex", phone: "", apiKey: "" },
    { name: "Cairo Office", phone: "", apiKey: "" },
  ],

  FLAVORS: [
    "Mango", "Blueberry", "Strawberry", "Pineapple", "Lemon-Mint",
    "Cantaloupe", "Coconut", "Watermelon", "Orange", "Peach",
    "Red Apple", "Green Apple", "Pomegranate"
  ],

  SPREADSHEET_NAME: "Badr El-Din Order Tracker",
  FORM_TITLE: "Badr El-Din Farms — Place Your Order",
  FORM_DESCRIPTION: "Please fill in this form to place your order. You will receive a WhatsApp confirmation once your order is received.",
};

// ────────────────────────────────────────────────────────────
//  🔍  SHEET FINDER — Finds sheets by partial name
//  Avoids emoji encoding issues with getSheetByName()
// ────────────────────────────────────────────────────────────
function _getSheet(ss, keyword) {
  return ss.getSheets().find(s => s.getName().includes(keyword)) || null;
}

// ────────────────────────────────────────────────────────────
//  🚀  MAIN ENTRY POINT — Run once to build everything
// ────────────────────────────────────────────────────────────
function setupEntireSystem() {
  Logger.log("🌿 Starting Badr El-Din system setup...");
  const ss = _createOrGetSpreadsheet();
  _buildOrdersSheet(ss);
  _buildDeliveriesSheet(ss);
  _buildInstructionsSheet(ss);
  _removeDefaultSheet(ss);
  const form = _buildOrderForm(ss);
  _setupFormTrigger(ss);
  // Save spreadsheet ID for use by the trigger
  PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", ss.getId());
  Logger.log("✅ Setup complete!");
  Logger.log("📋 Spreadsheet: " + ss.getUrl());
  Logger.log("📝 Form URL (share with distributors): " + form.getPublishedUrl());
}

// ────────────────────────────────────────────────────────────
//  📊  CREATE OR GET SPREADSHEET
// ────────────────────────────────────────────────────────────
function _createOrGetSpreadsheet() {
  const files = DriveApp.getFilesByName(CONFIG.SPREADSHEET_NAME);
  if (files.hasNext()) {
    const existing = SpreadsheetApp.open(files.next());
    Logger.log("Found existing spreadsheet, rebuilding sheets...");
    return existing;
  }
  return SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
}

// ────────────────────────────────────────────────────────────
//  📋  ORDERS SHEET
//  • No pre-filled formulas — all data written by onFormSubmit
//  • Status dropdown applied to 500 rows (handles future growth)
//  • Total Ordered written as a value by the trigger, not formula
// ────────────────────────────────────────────────────────────
function _buildOrdersSheet(ss) {
  let ws = _getSheet(ss, "Orders") || ss.insertSheet("📋 Orders");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#2E75B6");

  const DARK_BLUE = "#1F4E79", MED_BLUE = "#2E75B6", WHITE = "#FFFFFF", LIGHT_BLUE = "#DEEAF1";
  const headers = ["Order ID", "Timestamp", "Distributor", ...CONFIG.FLAVORS, "Total Ordered", "Status", "Notes"];

  // Row 1 — main headers
  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_BLUE)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 35);

  // Row 2 — sub-headers (cases)
  const subHeaders = ["", "", "", ...CONFIG.FLAVORS.map(() => "(cases)"), "", "", ""];
  ws.getRange(2, 1, 1, headers.length)
    .setValues([subHeaders])
    .setBackground(MED_BLUE)
    .setFontColor(WHITE)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setFontFamily("Arial")
    .setFontSize(9);
  ws.setRowHeight(2, 18);

  // Status dropdown — applied to 500 rows, no formulas
  const statusCol = 3 + CONFIG.FLAVORS.length + 2;
  ws.getRange(3, statusCol, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Pending", "Partial", "Fulfilled", "Cancelled"], true)
      .build()
  );

  // Column widths
  const totalCol = 3 + CONFIG.FLAVORS.length + 1;
  ws.setColumnWidth(1, 110);
  ws.setColumnWidth(2, 150);
  ws.setColumnWidth(3, 160);
  for (let i = 4; i <= 3 + CONFIG.FLAVORS.length; i++) ws.setColumnWidth(i, 105);
  ws.setColumnWidth(totalCol, 130);
  ws.setColumnWidth(statusCol, 110);
  ws.setColumnWidth(statusCol + 1, 200);

  ws.setFrozenRows(2);
  ws.setFrozenColumns(3);

  // Alternating row colors via conditional formatting
  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A3<>"")`)
      .setBackground(LIGHT_BLUE)
      .setRanges([ws.getRange(3, 1, 500, headers.length)])
      .build()
  ]);

  Logger.log("✅ Orders sheet built.");
}

// ────────────────────────────────────────────────────────────
//  🚚  DELIVERIES SHEET
//  • No pre-filled formulas — all data written by submitDelivery()
//  • Total Delivered written as a static value by the web app trigger
// ────────────────────────────────────────────────────────────
function _buildDeliveriesSheet(ss) {
  let ws = _getSheet(ss, "Deliveries") || ss.insertSheet("🚚 Deliveries");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#70AD47");

  const DARK_GREEN = "#375623", MED_GREEN = "#70AD47", WHITE = "#FFFFFF", LIGHT_GREEN = "#E2EFDA";
  const headers = ["Delivery ID", "Order ID", "Date Delivered", "Distributor", ...CONFIG.FLAVORS, "Total Delivered", "Notes"];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_GREEN)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 35);

  const subHeaders = ["", "", "", "", ...CONFIG.FLAVORS.map(() => "(cases)"), "", ""];
  ws.getRange(2, 1, 1, headers.length)
    .setValues([subHeaders])
    .setBackground(MED_GREEN)
    .setFontColor(WHITE)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setFontFamily("Arial")
    .setFontSize(9);
  ws.setRowHeight(2, 18);

  // Total Delivered — written as a static value by submitDelivery()
  // No pre-filled formulas to avoid getLastRow() returning wrong row
  const totalCol = 4 + CONFIG.FLAVORS.length + 1;

  // Date validation
  ws.getRange(3, 3, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().build()
  );

  ws.setColumnWidth(1, 120);
  ws.setColumnWidth(2, 110);
  ws.setColumnWidth(3, 140);
  ws.setColumnWidth(4, 160);
  for (let i = 5; i <= 4 + CONFIG.FLAVORS.length; i++) ws.setColumnWidth(i, 105);
  ws.setColumnWidth(totalCol, 140);
  ws.setColumnWidth(totalCol + 1, 200);

  ws.setFrozenRows(2);
  ws.setFrozenColumns(4);

  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A3<>"")`)
      .setBackground(LIGHT_GREEN)
      .setRanges([ws.getRange(3, 1, 500, headers.length)])
      .build()
  ]);

  Logger.log("✅ Deliveries sheet built.");
}

// ────────────────────────────────────────────────────────────
//  📊  KPI DASHBOARD — Phase 3 narrow-schema rebuild
//  • Summary cards only, no per-order row grid.
//  • Hidden helper block (cols X:Z, rows 1–501) computes cycle
//    time per order once; reused by this sheet's own card AND by
//    _buildDistributorSummarySheet's per-distributor average.
// ────────────────────────────────────────────────────────────
function _buildKPIDashboardCards(ss) {
  let ws = _getSheet(ss, "KPI") || ss.insertSheet("📊 KPI Dashboard");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#ED7D31");

  const WHITE = "#FFFFFF", DARK_BLUE = "#1F4E79";

  ws.getRange("A1:T1").merge()
    .setValue("📊 Badr El-Din Farms — Order Fulfillment KPI Dashboard")
    .setBackground(DARK_BLUE).setFontColor(WHITE)
    .setFontWeight("bold").setFontSize(14).setFontFamily("Arial")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.setRowHeight(1, 40);

  ws.getRange("A2:T2").merge()
    .setValue("Auto-calculated from Orders, Deliveries & their line tables")
    .setFontStyle("italic").setFontSize(9).setFontColor("#595959")
    .setHorizontalAlignment("center");
  ws.setRowHeight(2, 18);

  ws.setRowHeight(3, 10);
  ws.setRowHeight(4, 20);
  ws.setRowHeight(5, 50);
  ws.setRowHeight(6, 20);
  ws.setRowHeight(7, 10);

  const cardDefs = [
    { label: "TOTAL ORDERS", formula: `=COUNTA('📋 Orders'!A2:A)`, color: "#1F4E79", pct: false, startCol: 1 },
    { label: "TOTAL DELIVERIES", formula: `=COUNTA('🚚 Deliveries'!A2:A)`, color: "#2E75B6", pct: false, startCol: 5 },
    { label: "OVERALL FILL RATE", formula: `=IFERROR(SUM('🚚 Delivery_Lines'!F2:F)/SUM('📋 Order_Lines'!E2:E),0)`, color: "#375623", pct: true, startCol: 9 },
    { label: "PENDING / PARTIAL", formula: `=COUNTIF('📋 Orders'!E2:E,"Pending")+COUNTIF('📋 Orders'!E2:E,"Partial")`, color: "#7030A0", pct: false, startCol: 13 },
    { label: "AVG FULFILLMENT DAYS", formula: `=IFERROR(AVERAGE(Z2:Z501),"—")`, color: "#843C0C", pct: false, startCol: 17 },
  ];

  cardDefs.forEach(card => {
    const sc = card.startCol;
    ws.getRange(4, sc, 1, 4).merge().setValue(card.label)
      .setBackground(card.color).setFontColor(WHITE)
      .setFontWeight("bold").setFontSize(9)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    ws.getRange(5, sc, 1, 4).merge().setFormula(card.formula)
      .setBackground(card.color).setFontColor(WHITE)
      .setFontWeight("bold").setFontSize(24)
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setNumberFormat(card.pct ? "0.0%" : "0");
    ws.getRange(6, sc, 1, 4).merge().setBackground(card.color);
  });

  // Fill Rate conditional formatting — card 3, cols I:L, row 5
  const fillRateCell = ws.getRange(5, 9, 1, 4);
  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.95)
      .setBackground("#E2EFDA").setFontColor("#375623").setRanges([fillRateCell]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.70, 0.949)
      .setBackground("#FFEB9C").setFontColor("#9C6500").setRanges([fillRateCell]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.70)
      .setBackground("#FFC7CE").setFontColor("#9C0006").setRanges([fillRateCell]).build(),
  ]);

  for (let i = 1; i <= 20; i++) ws.setColumnWidth(i, 90);

  // ── Hidden helper block (cols X:Z = 24:26), rows 1–501 ──────
  // X = Order_ID, Y = Distributor, Z = cycle time in days — blank
  // when the order has zero deliveries yet ("fulfilled" = has ≥1
  // delivery, Partial included). Mirrors Orders rows 2–501
  // directly, no row-offset math needed. Formula arrays are built
  // in JS first, then each column is written in a single batch
  // call — 3 API calls instead of 1,500 individual setFormula()s.
  ws.getRange(1, 24, 1, 3).setValues([["Order_ID", "Distributor", "Cycle_Time_Days"]]);

  const xFormulas = [], yFormulas = [], zFormulas = [];
  for (let r = 2; r <= 501; r++) {
    xFormulas.push([`=IF('📋 Orders'!A${r}="","",'📋 Orders'!A${r})`]);
    yFormulas.push([`=IF(X${r}="","",'📋 Orders'!C${r})`]);
    zFormulas.push([
      `=IF(X${r}="","",IF(COUNTIF('🚚 Deliveries'!$B:$B,X${r})=0,"",MINIFS('🚚 Deliveries'!$C:$C,'🚚 Deliveries'!$B:$B,X${r})-'📋 Orders'!B${r}))`
    ]);
  }
  ws.getRange(2, 24, 500, 1).setFormulas(xFormulas);
  ws.getRange(2, 25, 500, 1).setFormulas(yFormulas);
  ws.getRange(2, 26, 500, 1).setFormulas(zFormulas);
  ws.hideColumns(24, 3);

  ws.setFrozenRows(2);

  Logger.log("✅ KPI Dashboard (cards) built.");
}

// ────────────────────────────────────────────────────────────
//  📈  PER-DISTRIBUTOR SUMMARY — Phase 3 narrow-schema rebuild
//  • Renames the old "📈 Summary" tab in place (fuzzy lookup on
//    "Summary" still matches it pre-rename).
//  • Rows come from _getConfig(ss).distributors (active only,
//    sheet-backed), not CONFIG.DISTRIBUTORS.
//  • Cases Ordered / Cases Delivered use a SUMPRODUCT + array
//    VLOOKUP join, since Order_Lines/Delivery_Lines don't carry
//    Distributor directly — only Order_ID, which has to be
//    resolved to a Distributor via Orders/Deliveries per line.
//  • Avg Cycle Time reuses the hidden Y:Z helper block built by
//    _buildKPIDashboardCards — no duplicate MINIFS computation.
// ────────────────────────────────────────────────────────────
function _buildDistributorSummarySheet(ss) {
  let ws = _getSheet(ss, "Summary") || ss.insertSheet("📈 Per-Distributor Summary");
  ws.setName("📈 Per-Distributor Summary");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#FFC000");

  const DARK_GOLD = "#7F6000", WHITE = "#FFFFFF", LIGHT_GOLD = "#FFF2CC", DARK_BLUE = "#1F4E79";

  ws.getRange("A1:G1").merge()
    .setValue("📈 Fill Rate Summary by Distributor")
    .setBackground(DARK_BLUE).setFontColor(WHITE)
    .setFontWeight("bold").setFontSize(13).setFontFamily("Arial")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.setRowHeight(1, 40);

  ws.getRange("A2:G2").merge()
    .setValue("Auto-calculated from Orders, Deliveries & their line tables")
    .setFontStyle("italic").setFontSize(9).setFontColor("#595959")
    .setHorizontalAlignment("center");

  const headers = ["Distributor", "Orders Received", "Cases Ordered", "Cases Delivered", "Fill Rate %", "Avg Cycle Time (days)", "Pending Orders"];
  ws.getRange(3, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_GOLD).setFontColor(WHITE)
    .setFontWeight("bold").setHorizontalAlignment("center")
    .setFontFamily("Arial").setFontSize(10).setWrap(true);
  ws.setRowHeight(3, 34);

  const distributors = _getConfig(ss).distributors;

  distributors.forEach((dist, i) => {
    const r = i + 4;
    ws.getRange(r, 1).setValue(dist.name);

    ws.getRange(r, 2).setFormula(`=COUNTIF('📋 Orders'!$C$2:$C$2000,A${r})`);

    ws.getRange(r, 3).setFormula(
      `=SUMPRODUCT(('📋 Order_Lines'!$E$2:$E$5000)*(IFERROR(VLOOKUP('📋 Order_Lines'!$B$2:$B$5000,'📋 Orders'!$A$2:$C$2000,3,FALSE),"")=A${r}))`
    );
    ws.getRange(r, 4).setFormula(
      `=SUMPRODUCT(('🚚 Delivery_Lines'!$F$2:$F$5000)*(IFERROR(VLOOKUP('🚚 Delivery_Lines'!$B$2:$B$5000,'🚚 Deliveries'!$A$2:$D$2000,4,FALSE),"")=A${r}))`
    );

    ws.getRange(r, 5).setFormula(`=IFERROR(D${r}/C${r},0)`);
    ws.getRange(r, 5).setNumberFormat("0.0%");

    ws.getRange(r, 6).setFormula(
      `=IFERROR(AVERAGEIFS('📊 KPI Dashboard'!$Z$2:$Z$501,'📊 KPI Dashboard'!$Y$2:$Y$501,A${r}),"—")`
    );
    ws.getRange(r, 6).setNumberFormat("0.0");

    ws.getRange(r, 7).setFormula(
      `=COUNTIFS('📋 Orders'!$C$2:$C$2000,A${r},'📋 Orders'!$E$2:$E$2000,"Pending")` +
      `+COUNTIFS('📋 Orders'!$C$2:$C$2000,A${r},'📋 Orders'!$E$2:$E$2000,"Partial")`
    );

    ws.getRange(r, 1, 1, 7)
      .setBackground(i % 2 === 0 ? LIGHT_GOLD : WHITE)
      .setFontFamily("Arial").setFontSize(10).setHorizontalAlignment("center");
    ws.getRange(r, 1).setHorizontalAlignment("left");
  });

  // Fill Rate % conditional formatting
  const fillRange = ws.getRange(4, 5, distributors.length, 1);
  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.95)
      .setBackground("#E2EFDA").setFontColor("#375623").setRanges([fillRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.70, 0.949)
      .setBackground("#FFEB9C").setFontColor("#9C6500").setRanges([fillRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.70)
      .setBackground("#FFC7CE").setFontColor("#9C0006").setRanges([fillRange]).build(),
  ]);

  [180, 120, 120, 130, 110, 150, 130].forEach((w, i) => ws.setColumnWidth(i + 1, w));
  ws.setFrozenRows(3);

  Logger.log("✅ Per-Distributor Summary sheet built.");
}

// ────────────────────────────────────────────────────────────
//  🌿  PER-FLAVOR BREAKDOWN — Phase 3 narrow-schema build (new)
//  • Rows come from the 🌿 Flavors sheet directly (active only),
//    not _getConfig() — needs Flavor_ID for the Reorder_Points
//    join, which _getConfig() doesn't expose.
//  • Current Stock is pulled from the Inventory Tracker via ONE
//    IMPORTRANGE into a hidden staging block (cols K:Q), then
//    looked up locally per row — not one IMPORTRANGE per flavor.
//  • Throughput Ratio dropped per approved plan — Cases Produced
//    and Cases Delivered are shown as separate columns only.
//  • NOTE: first run requires a one-time manual "Allow access"
//    click in the Sheets UI to authorize the IMPORTRANGE — this
//    cannot be done from Apps Script.
// ────────────────────────────────────────────────────────────
function _buildFlavorBreakdownSheet(ss) {
  let ws = _getSheet(ss, "Per-Flavor Breakdown") || ss.insertSheet("🌿 Per-Flavor Breakdown");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#548235");

  const DARK_GREEN = "#375623", WHITE = "#FFFFFF", LIGHT_GREEN = "#E2EFDA", DARK_BLUE = "#1F4E79";
  const RED_BG = "#FFC7CE", RED_FONT = "#9C0006";

  ws.getRange("A1:H1").merge()
    .setValue("🌿 Per-Flavor Breakdown")
    .setBackground(DARK_BLUE).setFontColor(WHITE)
    .setFontWeight("bold").setFontSize(13).setFontFamily("Arial")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.setRowHeight(1, 40);

  ws.getRange("A2:H2").merge()
    .setValue("Ordered / delivered / returned / produced from line tables — Current Stock imported from the Inventory Tracker")
    .setFontStyle("italic").setFontSize(9).setFontColor("#595959")
    .setHorizontalAlignment("center");

  const headers = ["Flavor", "Cases Ordered", "Cases Delivered", "Cases Returned", "Fill Rate %", "Current Stock", "Below Reorder Threshold", "Cases Produced"];
  ws.getRange(3, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_GREEN).setFontColor(WHITE)
    .setFontWeight("bold").setHorizontalAlignment("center")
    .setFontFamily("Arial").setFontSize(10).setWrap(true);
  ws.setRowHeight(3, 34);

  // Active flavors, read directly — need Flavor_ID (col A) for the
  // Reorder_Points join; _getConfig() only returns Flavor_Name.
  const flavorsSheet = _getSheet(ss, "Flavors");
  const flavorRows = flavorsSheet.getDataRange().getValues().slice(1).filter(row => row[3] === true);
  // row shape: [Flavor_ID, Product_ID, Flavor_Name, Is_Active]

  // One-shot IMPORTRANGE staging of ' Inventory Balance' (cols K:Q, hidden).
  // INVENTORY_SPREADSHEET_ID is read at build time, never hardcoded.
  const invId = PropertiesService.getScriptProperties().getProperty("INVENTORY_SPREADSHEET_ID");
  ws.getRange(1, 11).setFormula(
    `=IFERROR(IMPORTRANGE("${invId}", "' Inventory Balance'!A2:G${flavorRows.length + 1}"), "Authorize IMPORTRANGE")`
  );

  flavorRows.forEach((f, i) => {
    const r = i + 4;
    const flavorId = f[0], flavorName = f[2];

    ws.getRange(r, 1).setValue(flavorName);
    ws.getRange(r, 10).setValue(flavorId); // hidden col J

    ws.getRange(r, 2).setFormula(`=SUMIF('📋 Order_Lines'!$D$2:$D$5000,A${r},'📋 Order_Lines'!$E$2:$E$5000)`);
    ws.getRange(r, 3).setFormula(`=SUMIF('🚚 Delivery_Lines'!$E$2:$E$5000,A${r},'🚚 Delivery_Lines'!$F$2:$F$5000)`);
    ws.getRange(r, 4).setFormula(`=SUMIF('↩️ Return_Lines'!$D$2:$D$5000,A${r},'↩️ Return_Lines'!$E$2:$E$5000)`);

    ws.getRange(r, 5).setFormula(`=IFERROR(C${r}/B${r},0)`);
    ws.getRange(r, 5).setNumberFormat("0.0%");

    ws.getRange(r, 6).setFormula(`=IFERROR(VLOOKUP(A${r},$K$1:$Q$${flavorRows.length},7,FALSE),"—")`);

    // ISNUMBER guard: while IMPORTRANGE is pending authorization (or its
    // VLOOKUP fails), col F holds placeholder text, not a stock figure —
    // without this guard the threshold comparison would silently show "NO".
    ws.getRange(r, 7).setFormula(
      `=IFERROR(IF(ISNUMBER(F${r}),IF(F${r}<VLOOKUP(J${r},'📉 Reorder_Points'!$A$2:$B$100,2,FALSE),"YES","NO"),"—"),"—")`
    );

    ws.getRange(r, 8).setFormula(`=SUMIF('🌱 Production_Lines'!$D$2:$D$5000,A${r},'🌱 Production_Lines'!$E$2:$E$5000)`);

    ws.getRange(r, 1, 1, 8)
      .setBackground(i % 2 === 0 ? LIGHT_GREEN : WHITE)
      .setFontFamily("Arial").setFontSize(10).setHorizontalAlignment("center");
    ws.getRange(r, 1).setHorizontalAlignment("left").setFontWeight("bold");
  });

  // Conditional formatting — Fill Rate % (3-tier) and Below Reorder Threshold flag
  const fillRange = ws.getRange(4, 5, flavorRows.length, 1);
  const flagRange = ws.getRange(4, 7, flavorRows.length, 1);
  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThanOrEqualTo(0.95)
      .setBackground("#E2EFDA").setFontColor("#375623").setRanges([fillRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(0.70, 0.949)
      .setBackground("#FFEB9C").setFontColor("#9C6500").setRanges([fillRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0.70)
      .setBackground("#FFC7CE").setFontColor("#9C0006").setRanges([fillRange]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("YES")
      .setBackground(RED_BG).setFontColor(RED_FONT).setRanges([flagRange]).build(),
  ]);

  [140, 120, 130, 130, 110, 120, 170, 130].forEach((w, i) => ws.setColumnWidth(i + 1, w));
  ws.setFrozenRows(3);
  ws.hideColumns(10, 8); // J:Q — Flavor_ID (J) + IMPORTRANGE staging (K:Q)

  Logger.log("✅ Per-Flavor Breakdown sheet built.");
}

// ────────────────────────────────────────────────────────────
//  📖  INSTRUCTIONS SHEET
// ────────────────────────────────────────────────────────────
function _buildInstructionsSheet(ss) {
  let ws = _getSheet(ss, "Instructions") || ss.insertSheet("📖 Instructions");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#7030A0");

  ws.getRange("B1:D1").merge()
    .setValue("📖 How to Use This System")
    .setBackground("#1F4E79").setFontColor("#FFFFFF")
    .setFontWeight("bold").setFontSize(14).setFontFamily("Arial")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.setRowHeight(1, 40);
  ws.setColumnWidth(1, 30);
  ws.setColumnWidth(2, 200);
  ws.setColumnWidth(3, 120);
  ws.setColumnWidth(4, 500);

  const rows = [
    ["SHEET / STEP", "WHO", "WHAT TO DO"],
    ["📝 Google Form", "Distributor", "Clicks the shared link, fills in quantities per flavor, submits. Receives a WhatsApp confirmation automatically."],
    ["📋 Orders Sheet", "Auto-filled", "DO NOT edit manually. Orders appear here automatically with an auto-generated Order ID. Update Status column as orders progress: Pending → Partial → Fulfilled."],
    ["🚚 Deliveries Sheet", "Your team", "When a delivery goes out, add a row: Delivery ID (DEL-001...), Order ID (from Orders sheet), date, distributor, actual quantities. For partial deliveries, add a row now and another later with the same Order ID."],
    ["📊 KPI Dashboard", "Management", "Read-only. Shows fill rate %, days to deliver, delivered vs ordered per flavor. Colors: Green ≥95%, Yellow 70–94%, Red <70%."],
    ["📈 Summary", "Management", "Per-distributor fill rate, average delivery time, and pending orders. Auto-updates."],
    ["", "", ""],
    ["⚠️ ORDER ID FORMAT", "", "Auto-generated as ORD-001, ORD-002... Never edit these."],
    ["⚠️ DELIVERY ID FORMAT", "", "Assigned by your team: DEL-001, DEL-002... Use the same Order ID for split deliveries."],
    ["⚠️ STATUS COLUMN", "", "Update manually in the Orders sheet. Pending = not yet delivered. Partial = some delivered. Fulfilled = fully delivered. Cancelled = cancelled."],
    ["⚠️ WHATSAPP ALERTS", "", "Requires CallMeBot API keys to be filled in the CONFIG section of the Apps Script. Once filled, run _setupFormTrigger() only — do NOT re-run setupEntireSystem()."],
  ];

  rows.forEach((row, i) => {
    const r = i + 2;
    ws.setRowHeight(r, i === 0 ? 25 : 50);
    row.forEach((val, ci) => {
      const cell = ws.getRange(r, ci + 2);
      cell.setValue(val).setFontFamily("Arial").setFontSize(10)
        .setVerticalAlignment("middle").setWrap(true)
        .setBorder(true, true, true, true, false, false, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
      if (i === 0) {
        cell.setBackground("#1F4E79").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
      } else if (row[0].startsWith("⚠️")) {
        cell.setBackground("#FFF2CC").setFontWeight(ci === 0 ? "bold" : "normal").setFontColor(ci === 0 ? "#7F6000" : "#1A1A1A");
      } else if (val === "") {
        cell.setBackground("#F2F2F2");
      } else {
        cell.setBackground(i % 2 === 0 ? "#EBF3FB" : "#FFFFFF");
        if (ci === 0) cell.setFontWeight("bold").setFontColor("#1F4E79");
      }
    });
  });

  Logger.log("✅ Instructions sheet built.");
}

// ────────────────────────────────────────────────────────────
//  📝  BUILD GOOGLE FORM
// ────────────────────────────────────────────────────────────
function _buildOrderForm(ss) {
  const form = FormApp.create(CONFIG.FORM_TITLE);
  form.setDescription(CONFIG.FORM_DESCRIPTION);
  form.setConfirmationMessage("✅ Your order has been received! You will receive a WhatsApp confirmation shortly. Thank you.");
  form.setCollectEmail(false);
  form.setAllowResponseEdits(false);
  form.setLimitOneResponsePerUser(false);
  form.setShowLinkToRespondAgain(false);

  form.addListItem()
    .setTitle("Your Name / Distributor")
    .setChoiceValues(CONFIG.DISTRIBUTORS.map(d => d.name))
    .setRequired(true);

  form.addSectionHeaderItem()
    .setTitle("📦 Order Quantities")
    .setHelpText("Enter the number of cases needed per flavor. Enter 0 if not needed.");

  CONFIG.FLAVORS.forEach(flavor => {
    form.addTextItem()
      .setTitle(`${flavor} (cases)`)
      .setHelpText("Enter 0 if not needed")
      .setRequired(true)
      .setValidation(
        FormApp.createTextValidation()
          .requireNumber()
          .requireNumberGreaterThanOrEqualTo(0)
          .build()
      );
  });

  form.addParagraphTextItem()
    .setTitle("Additional Notes (optional)")
    .setRequired(false);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  PropertiesService.getScriptProperties().setProperty("FORM_URL", form.getPublishedUrl());
  PropertiesService.getScriptProperties().setProperty("FORM_ID", form.getId());

  Logger.log("✅ Form created: " + form.getPublishedUrl());
  return form;
}

// ────────────────────────────────────────────────────────────
//  ⚡  SETUP FORM TRIGGER
//  Run this alone when updating API keys — does NOT rebuild sheets
// ────────────────────────────────────────────────────────────
function _setupFormTrigger(ss) {
  // Accept ss as parameter or fetch from saved ID
  if (!ss) {
    const ssId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
    if (!ssId) {
      Logger.log("❌ No SPREADSHEET_ID saved. Please run setupEntireSystem() first.");
      return;
    }
    ss = SpreadsheetApp.openById(ssId);
  }
  // Remove existing onFormSubmit triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "onFormSubmit") ScriptApp.deleteTrigger(t);
  });
  // Create trigger bound to the spreadsheet ID directly
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(ss.getId())
    .onFormSubmit()
    .create();
  Logger.log("✅ Form submit trigger set for: " + ss.getName());
}

// ────────────────────────────────────────────────────────────
//  ⚡  ON FORM SUBMIT — Writes order to Orders sheet + WhatsApp
// ────────────────────────────────────────────────────────────
function onFormSubmit(e) {
  try {
    const ssId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
    const ss = SpreadsheetApp.openById(ssId);
    const ordSheet = ss.getSheets().find(s => s.getName().includes("Orders"));
    const responses = e.namedValues;

    // Get distributor name
    const distributorName = responses["Your Name / Distributor"] ?
      responses["Your Name / Distributor"][0] : "Unknown";

    // Collect flavor quantities
    const flavorQtys = CONFIG.FLAVORS.map(flavor =>
      parseInt((responses[`${flavor} (cases)`] || ["0"])[0]) || 0
    );

    // Calculate total
    const totalOrdered = flavorQtys.reduce((sum, q) => sum + q, 0);

    // Get notes
    const notes = responses["Additional Notes (optional)"] ?
      responses["Additional Notes (optional)"][0] : "";

    // Generate Order ID
    const nextRow = ordSheet.getLastRow() + 1;
    const orderCount = nextRow - 2;
    const orderId = "ORD-" + String(orderCount).padStart(3, "0");
    const timestamp = new Date();

    // Write full row: OrderID | Timestamp | Distributor | ...flavors | Total | Status | Notes
    const rowData = [orderId, timestamp, distributorName, ...flavorQtys, totalOrdered, "Pending", notes];
    ordSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    ordSheet.getRange(nextRow, 2).setNumberFormat("yyyy-mm-dd hh:mm");

    // Re-apply status dropdown to new row
    const statusCol = 3 + CONFIG.FLAVORS.length + 2;
    ordSheet.getRange(nextRow, statusCol).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(["Pending", "Partial", "Fulfilled", "Cancelled"], true)
        .build()
    );

    // Build WhatsApp summary
    let orderSummary = "";
    CONFIG.FLAVORS.forEach((flavor, i) => {
      if (flavorQtys[i] > 0) orderSummary += `  • ${flavor}: ${flavorQtys[i]} cases\n`;
    });

    // Notify internal team
    const teamMsg =
      `🌿 *New Order Received*\n` +
      `Order ID: *${orderId}*\n` +
      `From: *${distributorName}*\n` +
      `Date: ${timestamp.toLocaleDateString("en-GB")}\n\n` +
      `*Order Details:*\n${orderSummary}\n` +
      `Total: *${totalOrdered} cases*` +
      (notes ? `\nNotes: ${notes}` : "");

    CONFIG.TEAM_WHATSAPP.forEach(m => _sendWhatsApp(m.phone, m.apiKey, teamMsg));

    // Confirm to distributor
    const distConfig = CONFIG.DISTRIBUTORS.find(d => d.name === distributorName);
    if (distConfig) {
      const distMsg =
        `✅ *Order Confirmed*\n` +
        `Hello ${distributorName},\n\n` +
        `Your order has been received.\n` +
        `Order ID: *${orderId}*\n` +
        `Date: ${timestamp.toLocaleDateString("en-GB")}\n\n` +
        `*Your Order:*\n${orderSummary}\n` +
        `Total: *${totalOrdered} cases*\n\n` +
        `You will be notified when your order is dispatched. Thank you! 🌿`;
      _sendWhatsApp(distConfig.phone, distConfig.apiKey, distMsg);
    }

    Logger.log(`✅ Order ${orderId} written and notifications sent.`);

  } catch (err) {
    Logger.log("❌ onFormSubmit error: " + err.toString());
  }
}

// ────────────────────────────────────────────────────────────
//  📱  SEND WHATSAPP VIA CALLMEBOT
// ────────────────────────────────────────────────────────────
function _sendWhatsApp(phone, apiKey, message) {
  try {
    const url = `https://api.callmebot.com/whatsapp.php?phone=${phone}&text=${encodeURIComponent(message)}&apikey=${apiKey}`;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    Logger.log(`📱 WhatsApp to ${phone}: HTTP ${res.getResponseCode()}`);
  } catch (err) {
    Logger.log(`❌ WhatsApp failed for ${phone}: ${err.toString()}`);
  }
}

// ────────────────────────────────────────────────────────────
//    LOG SYNC ERROR
//  • Appends a row to Sync Errors sheet + sends WhatsApp alert
//  • Called when a cross-system write (e.g., Inventory Movements)
//    fails inside another function's try/catch
// ────────────────────────────────────────────────────────────
function _logSyncError(functionName, err, orderId, distributor, details) {
  try {
    const ss = SpreadsheetApp.openById(
      PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
    );
    const errSheet = ss.getSheets().find(s => s.getName().includes("Sync Errors"));
    const nextRow = errSheet.getLastRow() + 1;
    const rowData = [new Date(), functionName, err.toString(), orderId || "", distributor || "", details || ""];
    errSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    errSheet.getRange(nextRow, 1).setNumberFormat("yyyy-mm-dd hh:mm");

    const msg =
      `⚠️ *Sync Error*\n` +
      `Function: *${functionName}*\n` +
      (orderId ? `Order: *${orderId}*\n` : "") +
      (distributor ? `Distributor: *${distributor}*\n` : "") +
      `Error: ${err.toString()}\n\n` +
      `Please check the Sync Errors sheet and Inventory Movements manually.`;
    CONFIG.TEAM_WHATSAPP.forEach(m => _sendWhatsApp(m.phone, m.apiKey, msg));

  } catch (logErr) {
    Logger.log("❌ _logSyncError itself failed: " + logErr.toString());
  }
}

// ────────────────────────────────────────────────────────────
//    NEXT SEQUENTIAL ID (Script Properties counter)
//  • Atomic via LockService — avoids race conditions from two
//    concurrent submissions reading the same counter value
//  • Used by submitProduction() (PRD-XXX) and
//    _logWriteOffMovements() (WOF-XXX)
// ────────────────────────────────────────────────────────────
function _nextSequentialId(counterProperty, prefix) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const next = (Number(props.getProperty(counterProperty)) || 0) + 1;
    props.setProperty(counterProperty, String(next));
    return prefix + String(next).padStart(3, "0");
  } finally {
    lock.releaseLock();
  }
}

// ────────────────────────────────────────────────────────────
//  📋  SCRIPT PROPERTIES COUNTER REGISTRY (documentation only)
//  • No logic change — just a record of every counterProperty/prefix
//    pair passed to _nextSequentialId() across the codebase, so a
//    new counter is never accidentally reused or collided with.
//
//  Counter Property   Prefix   Used for
//  ORD_COUNTER         "ORD-"   Orders (OrderBackend.gs: submitOrder)
//  DEL_COUNTER         "DEL-"   Deliveries (DeliveryBackend.gs: submitDelivery)
//  PRD_COUNTER         "PRD-"   Production (ProductionBackend.gs: submitProduction)
//  WOF_COUNTER         "WOF-"   Write-off (Inventory.gs: _logWriteOffMovements)
//  OL_COUNTER          "OL-"    Order_Lines (Phase 2 — not yet wired to a caller)
//  DL_COUNTER          "DL-"    Delivery_Lines (Phase 2 — not yet wired to a caller)
//  RL_COUNTER          "RL-"    Return_Lines (Phase 2 — not yet wired to a caller)
//  PL_COUNTER          "PL-"    Production_Lines (Phase 2 — not yet wired to a caller)
//
//  Note: ORD-XXX / DEL-XXX / RET-XXX header IDs are currently generated
//  by counting sheet rows, not via _nextSequentialId(). These will be
//  migrated to _nextSequentialId() in Phase 2 sub-tasks 2.3–2.5.
// ────────────────────────────────────────────────────────────

// ────────────────────────────────────────────────────────────
//  🔔  NOTIFY DISTRIBUTOR ON STATUS CHANGE
//  ⚠️  PLACEHOLDER — Will be connected to an onEdit trigger later
//  When built, this will auto-fire when Status changes in Orders sheet
// ────────────────────────────────────────────────────────────
function onStatusChange(orderId, newStatus) {
  try {
    const ssId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
    const ss = SpreadsheetApp.openById(ssId);
    const ordSheet = ss.getSheets().find(s => s.getName().includes("Orders"));
    const data = ordSheet.getDataRange().getValues();

    for (let i = 2; i < data.length; i++) {
      if (data[i][0] === orderId) {
        const distributorName = data[i][2];
        const distConfig = CONFIG.DISTRIBUTORS.find(d => d.name === distributorName);
        if (distConfig) {
          const msg =
            `📦 *Order Update*\n` +
            `Order ID: *${orderId}*\n` +
            `New Status: *${newStatus}*\n\n` +
            `Please contact us if you have any questions. 🌿`;
          _sendWhatsApp(distConfig.phone, distConfig.apiKey, msg);
          Logger.log(`✅ Status update sent to ${distributorName}`);
        }
        break;
      }
    }
  } catch (err) {
    Logger.log("❌ onStatusChange error: " + err.toString());
  }
}

// ────────────────────────────────────────────────────────────
//  🧹  REMOVE DEFAULT SHEET
// ────────────────────────────────────────────────────────────
function _removeDefaultSheet(ss) {
  ["Sheet1", "Sheet 1"].forEach(name => {
    const s = ss.getSheetByName(name);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
}

// ────────────────────────────────────────────────────────────
//  🔧  COLUMN LETTER HELPER
// ────────────────────────────────────────────────────────────
function columnLetter(n) {
  let result = "";
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

// ────────────────────────────────────────────────────────────
//    SYNC ERRORS SHEET
//  • Append-only log for cross-system sync failures
//    (e.g., Inventory Movements write failing during submitDelivery)
//  • No pre-filled formulas — rows appended by _logSyncError()
// ────────────────────────────────────────────────────────────
function _buildSyncErrorsSheet(ss) {
  let ws = _getSheet(ss, "Sync Errors") || ss.insertSheet(" Sync Errors");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#C00000");

  const DARK_RED = "#C00000", WHITE = "#FFFFFF";

  const headers = ["Timestamp", "Function", "Error", "Order ID", "Distributor", "Details"];
  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_RED)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 30);

  ws.setColumnWidth(1, 140);
  ws.setColumnWidth(2, 220);
  ws.setColumnWidth(3, 300);
  ws.setColumnWidth(4, 100);
  ws.setColumnWidth(5, 140);
  ws.setColumnWidth(6, 250);
  ws.setFrozenRows(1);

  Logger.log("✅ Sync Errors sheet built.");
}

// ────────────────────────────────────────────────────────────
//  ↩️  RETURNS SHEET
//  • No pre-filled formulas — all data written by submitReturn()
//  • Total Returned written as a static value, same as Deliveries
// ────────────────────────────────────────────────────────────
function _buildReturnsSheet(ss) {
  let ws = _getSheet(ss, "Returns") || ss.insertSheet(" Returns");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#E65100");

  const DARK_ORANGE = "#BF360C", MED_ORANGE = "#E65100", WHITE = "#FFFFFF", LIGHT_ORANGE = "#FFE0B2";
  const headers = ["Return ID", "Date", "Distributor", ...CONFIG.FLAVORS, "Total Returned", "Notes"];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_ORANGE)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 35);

  const subHeaders = ["", "", "", ...CONFIG.FLAVORS.map(() => "(cases)"), "", ""];
  ws.getRange(2, 1, 1, headers.length)
    .setValues([subHeaders])
    .setBackground(MED_ORANGE)
    .setFontColor(WHITE)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setFontFamily("Arial")
    .setFontSize(9);
  ws.setRowHeight(2, 18);

  // Total Returned — written as a static value by submitReturn()
  // No pre-filled formulas to avoid getLastRow() returning wrong row
  const totalCol = 3 + CONFIG.FLAVORS.length + 1;

  // Date validation
  ws.getRange(3, 2, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().build()
  );

  ws.setColumnWidth(1, 120);
  ws.setColumnWidth(2, 140);
  ws.setColumnWidth(3, 160);
  for (let i = 4; i <= 3 + CONFIG.FLAVORS.length; i++) ws.setColumnWidth(i, 105);
  ws.setColumnWidth(totalCol, 140);
  ws.setColumnWidth(totalCol + 1, 200);

  ws.setFrozenRows(2);
  ws.setFrozenColumns(3);

  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A3<>"")`)
      .setBackground(LIGHT_ORANGE)
      .setRanges([ws.getRange(3, 1, 500, headers.length)])
      .build()
  ]);

  Logger.log("✅ Returns sheet built.");
}

// ────────────────────────────────────────────────────────────
//  🌱  PRODUCTIONS SHEET
//  • No pre-filled formulas — all data written by submitProduction()
//  • Total written as a static value, same as Deliveries/Returns
// ────────────────────────────────────────────────────────────
function _buildProductionsSheet(ss) {
  let ws = _getSheet(ss, "Productions") || ss.insertSheet(" Productions");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#00695C");

  const DARK_TEAL = "#004D40", MED_TEAL = "#00695C", WHITE = "#FFFFFF", LIGHT_TEAL = "#B2DFDB";
  const headers = ["Production ID", "Date", "Lot Number", ...CONFIG.FLAVORS, "Total", "Notes"];

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

  const subHeaders = ["", "", "", ...CONFIG.FLAVORS.map(() => "(cases)"), "", ""];
  ws.getRange(2, 1, 1, headers.length)
    .setValues([subHeaders])
    .setBackground(MED_TEAL)
    .setFontColor(WHITE)
    .setFontStyle("italic")
    .setHorizontalAlignment("center")
    .setFontFamily("Arial")
    .setFontSize(9);
  ws.setRowHeight(2, 18);

  // Total — written as a static value by submitProduction()
  // No pre-filled formulas to avoid getLastRow() returning the wrong row
  const totalCol = 3 + CONFIG.FLAVORS.length + 1;

  // Date validation
  ws.getRange(3, 2, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().build()
  );

  ws.setColumnWidth(1, 120);
  ws.setColumnWidth(2, 140);
  ws.setColumnWidth(3, 140);
  for (let i = 4; i <= 3 + CONFIG.FLAVORS.length; i++) ws.setColumnWidth(i, 105);
  ws.setColumnWidth(totalCol, 110);
  ws.setColumnWidth(totalCol + 1, 200);

  ws.setFrozenRows(2);
  ws.setFrozenColumns(3);

  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A3<>"")`)
      .setBackground(LIGHT_TEAL)
      .setRanges([ws.getRange(3, 1, 500, headers.length)])
      .build()
  ]);

  Logger.log("✅ Productions sheet built.");
}

// ────────────────────────────────────────────────────────────
//  ⚙️  CONFIG SHEETS — Phase 1 of DB migration
//  • Static reference sheets, seeded once from current CONFIG /
//    Inventory.gs threshold values. Not wired into any backend yet.
//  • Never added to setupEntireSystem() — run only via the temporary
//    _runBuildConfigSheets() wrapper below (delete after running).
// ────────────────────────────────────────────────────────────
function _buildConfigSheets(ss) {
  _buildProductsSheet(ss);
  _buildFlavorsSheet(ss);
  _buildDistributorsSheet(ss);
  _buildDistributorProductsSheet(ss);
  _buildTeamContactsSheet(ss);
  _buildReorderPointsSheet(ss);
  Logger.log("✅ Config sheets built.");
}

function _buildProductsSheet(ss) {
  // Full-name keyword avoids "Products" fuzzy-matching "🔗 Distributor_Products"
  let ws = _getSheet(ss, "⚙️ Products") || ss.insertSheet("⚙️ Products");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#616161");

  const headers = ["Product_ID", "Product_Code", "Product_Name", "Pack_Size", "Is_Active"];
  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#424242")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  ws.setFrozenRows(1);

  ws.getRange(2, 1, 1, headers.length).setValues([["PROD-001", "DOTS", "Dots", 12, true]]);

  ws.setColumnWidth(1, 100);
  ws.setColumnWidth(2, 110);
  ws.setColumnWidth(3, 150);
  ws.setColumnWidth(4, 100);
  ws.setColumnWidth(5, 90);

  Logger.log("✅ Products sheet built.");
}

function _buildFlavorsSheet(ss) {
  let ws = _getSheet(ss, "Flavors") || ss.insertSheet("🌿 Flavors");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#616161");

  const headers = ["Flavor_ID", "Product_ID", "Flavor_Name", "Is_Active"];
  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#424242")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  ws.setFrozenRows(1);

  const rows = CONFIG.FLAVORS.map((flavor, i) => [
    "FLV-" + String(i + 1).padStart(3, "0"),
    "PROD-001",
    flavor,
    true
  ]);
  ws.getRange(2, 1, rows.length, headers.length).setValues(rows);

  ws.setColumnWidth(1, 100);
  ws.setColumnWidth(2, 100);
  ws.setColumnWidth(3, 150);
  ws.setColumnWidth(4, 90);

  Logger.log("✅ Flavors sheet built.");
}

function _buildDistributorsSheet(ss) {
  let ws = _getSheet(ss, "Distributors") || ss.insertSheet("🏢 Distributors");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#616161");

  const headers = ["Distributor_ID", "Name", "Slug", "Phone", "CallMeBot_ApiKey", "Is_Active"];
  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#424242")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  ws.setFrozenRows(1);

  const rows = [
    ["DIST-001", "Family - Alex", "family-alex", "", "", true],
    ["DIST-002", "Cairo Office", "cairo-office", "", "", true],
  ];
  ws.getRange(2, 1, rows.length, headers.length).setValues(rows);

  ws.setColumnWidth(1, 110);
  ws.setColumnWidth(2, 150);
  ws.setColumnWidth(3, 120);
  ws.setColumnWidth(4, 120);
  ws.setColumnWidth(5, 150);
  ws.setColumnWidth(6, 90);

  Logger.log("✅ Distributors sheet built.");
}

function _buildDistributorProductsSheet(ss) {
  // Full-name keyword — "Distributor_Products" would otherwise fuzzy-match
  // as a substring of itself fine, but keeping it explicit for symmetry
  // with the Products lookup above and to stay unambiguous on rebuild.
  let ws = _getSheet(ss, "🔗 Distributor_Products") || ss.insertSheet("🔗 Distributor_Products");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#616161");

  const headers = ["DP_ID", "Distributor_ID", "Product_ID"];
  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#424242")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  ws.setFrozenRows(1);

  const rows = [
    ["DP-001", "DIST-001", "PROD-001"],
    ["DP-002", "DIST-002", "PROD-001"],
  ];
  ws.getRange(2, 1, rows.length, headers.length).setValues(rows);

  ws.setColumnWidth(1, 90);
  ws.setColumnWidth(2, 120);
  ws.setColumnWidth(3, 110);

  Logger.log("✅ Distributor_Products sheet built.");
}

function _buildTeamContactsSheet(ss) {
  let ws = _getSheet(ss, "Team_Contacts") || ss.insertSheet("📱 Team_Contacts");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#616161");

  const headers = ["Contact_ID", "Name", "Phone", "CallMeBot_ApiKey", "Is_Active"];
  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#424242")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  ws.setFrozenRows(1);

  const rows = CONFIG.TEAM_WHATSAPP.map((m, i) => [
    "TC-" + String(i + 1).padStart(3, "0"),
    m.name,
    m.phone,
    m.apiKey,
    true
  ]);
  ws.getRange(2, 1, rows.length, headers.length).setValues(rows);

  ws.setColumnWidth(1, 100);
  ws.setColumnWidth(2, 170);
  ws.setColumnWidth(3, 120);
  ws.setColumnWidth(4, 150);
  ws.setColumnWidth(5, 90);

  Logger.log("✅ Team_Contacts sheet built.");
}

function _buildReorderPointsSheet(ss) {
  let ws = _getSheet(ss, "Reorder_Points") || ss.insertSheet("📉 Reorder_Points");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#616161");

  const headers = ["Flavor_ID", "Min_Stock_Cases"];
  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#424242")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  ws.setFrozenRows(1);

  // Matches CLAUDE.md low-stock thresholds (Section 3)
  const THRESHOLDS = {
    "Mango": 75, "Blueberry": 75, "Strawberry": 75, "Lemon-Mint": 75,
    "Coconut": 25, "Orange": 25,
  };
  const rows = CONFIG.FLAVORS.map((flavor, i) => [
    "FLV-" + String(i + 1).padStart(3, "0"),
    THRESHOLDS[flavor] || 50
  ]);
  ws.getRange(2, 1, rows.length, headers.length).setValues(rows);

  ws.setColumnWidth(1, 100);
  ws.setColumnWidth(2, 140);

  Logger.log("✅ Reorder_Points sheet built.");
}

// ────────────────────────────────────────────────────────────
//  🗄️  DB MIGRATION — Phase 2, Sub-task 2.1
//  • Archives the four wide transactional sheets ahead of building
//    the new header + lines narrow schema.
//  • Copy → rename copy → delete original, so the data is preserved
//    read-only under an _ARCHIVE_ prefix instead of being destroyed.
//  • Never added to setupEntireSystem() — run only via the temporary
//    _runBuildPhase2Sheets() wrapper below (delete after running).
// ────────────────────────────────────────────────────────────
function _archiveWideSheets(ss) {
  const targets = [
    { keyword: "📋 Orders", archiveName: "_ARCHIVE_Orders" },
    { keyword: "🚚 Deliveries", archiveName: "_ARCHIVE_Deliveries" },
    { keyword: "Returns", archiveName: "_ARCHIVE_Returns" },
    { keyword: "Productions", archiveName: "_ARCHIVE_Productions" },
  ];

  targets.forEach(({ keyword, archiveName }) => {
    const original = _getSheet(ss, keyword);
    if (!original) {
      Logger.log(`⚠️ _archiveWideSheets: no sheet found for keyword "${keyword}" — skipped.`);
      return;
    }
    const originalName = original.getName();
    const copy = original.copyTo(ss);
    copy.setName(archiveName);
    copy.setTabColor("#9E9E9E");
    ss.deleteSheet(original);
    Logger.log(`✅ Archived "${originalName}" → "${archiveName}".`);
  });
}

// ────────────────────────────────────────────────────────────
//  📋  ORDERS HEADER SHEET — Phase 2 narrow schema
//  • Full-name keyword avoids fuzzy-matching "_ARCHIVE_Orders"
//  • Single header row — line-level detail lives in Order_Lines
// ────────────────────────────────────────────────────────────
function _buildOrdersHeaderSheet(ss) {
  let ws = _getSheet(ss, "📋 Orders") || ss.insertSheet("📋 Orders");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#2E75B6");

  const DARK_BLUE = "#1F4E79", WHITE = "#FFFFFF", LIGHT_BLUE = "#DEEAF1";
  const headers = ["Order_ID", "Timestamp", "Distributor", "Total_Ordered", "Status", "Notes"];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_BLUE)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 35);

  // Status dropdown — rows 2–501
  ws.getRange(2, 5, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Pending", "Partial", "Fulfilled", "Cancelled"], true)
      .build()
  );

  ws.setColumnWidth(1, 110);
  ws.setColumnWidth(2, 150);
  ws.setColumnWidth(3, 160);
  ws.setColumnWidth(4, 130);
  ws.setColumnWidth(5, 110);
  ws.setColumnWidth(6, 200);

  ws.setFrozenRows(1);
  ws.setFrozenColumns(2);

  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A2<>"")`)
      .setBackground(LIGHT_BLUE)
      .setRanges([ws.getRange(2, 1, 500, headers.length)])
      .build()
  ]);

  Logger.log("✅ Orders header sheet built.");
}

// ────────────────────────────────────────────────────────────
//  🚚  DELIVERIES HEADER SHEET — Phase 2 narrow schema
// ────────────────────────────────────────────────────────────
function _buildDeliveriesHeaderSheet(ss) {
  let ws = _getSheet(ss, "🚚 Deliveries") || ss.insertSheet("🚚 Deliveries");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#70AD47");

  const DARK_GREEN = "#375623", WHITE = "#FFFFFF", LIGHT_GREEN = "#E2EFDA";
  const headers = ["Delivery_ID", "Order_ID", "Date_Delivered", "Distributor", "Total_Delivered", "Notes"];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_GREEN)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 35);

  // Date validation — rows 2–501, col 3 (Date_Delivered)
  ws.getRange(2, 3, 500, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().build()
  );

  ws.setColumnWidth(1, 120);
  ws.setColumnWidth(2, 110);
  ws.setColumnWidth(3, 140);
  ws.setColumnWidth(4, 160);
  ws.setColumnWidth(5, 140);
  ws.setColumnWidth(6, 200);

  ws.setFrozenRows(1);
  ws.setFrozenColumns(2);

  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A2<>"")`)
      .setBackground(LIGHT_GREEN)
      .setRanges([ws.getRange(2, 1, 500, headers.length)])
      .build()
  ]);

  Logger.log("✅ Deliveries header sheet built.");
}

// ────────────────────────────────────────────────────────────
//  ↩️  RETURNS HEADER SHEET — Phase 2 narrow schema
// ────────────────────────────────────────────────────────────
function _buildReturnsHeaderSheet(ss) {
  let ws = _getSheet(ss, "↩️ Returns") || ss.insertSheet("↩️ Returns");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#ED7D31");

  const DARK_ORANGE = "#843C0C", WHITE = "#FFFFFF", LIGHT_ORANGE = "#FCE4D6";
  const headers = ["Return_ID", "Date", "Distributor", "Total_Returned", "Notes"];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_ORANGE)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 35);

  ws.setColumnWidth(1, 110);
  ws.setColumnWidth(2, 130);
  ws.setColumnWidth(3, 160);
  ws.setColumnWidth(4, 130);
  ws.setColumnWidth(5, 200);

  ws.setFrozenRows(1);
  ws.setFrozenColumns(2);

  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A2<>"")`)
      .setBackground(LIGHT_ORANGE)
      .setRanges([ws.getRange(2, 1, 500, headers.length)])
      .build()
  ]);

  Logger.log("✅ Returns header sheet built.");
}

// ────────────────────────────────────────────────────────────
//  🌱  PRODUCTIONS HEADER SHEET — Phase 2 narrow schema
// ────────────────────────────────────────────────────────────
function _buildProductionsHeaderSheet(ss) {
  let ws = _getSheet(ss, "🌱 Productions") || ss.insertSheet("🌱 Productions");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#70AD47");

  const DARK_GREEN = "#375623", WHITE = "#FFFFFF", LIGHT_GREEN = "#E2EFDA";
  const headers = ["Production_ID", "Date", "Lot_Number", "Total_Produced", "Notes"];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_GREEN)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 35);

  ws.setColumnWidth(1, 120);
  ws.setColumnWidth(2, 130);
  ws.setColumnWidth(3, 140);
  ws.setColumnWidth(4, 130);
  ws.setColumnWidth(5, 200);

  ws.setFrozenRows(1);
  ws.setFrozenColumns(2);

  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A2<>"")`)
      .setBackground(LIGHT_GREEN)
      .setRanges([ws.getRange(2, 1, 500, headers.length)])
      .build()
  ]);

  Logger.log("✅ Productions header sheet built.");
}

// ────────────────────────────────────────────────────────────
//  📋  ORDER LINES SHEET — Phase 2 narrow schema
//  • One row per flavor per order. Full-name keyword avoids
//    fuzzy-matching "📋 Orders" (header) or "_ARCHIVE_Orders".
// ────────────────────────────────────────────────────────────
function _buildOrderLinesSheet(ss) {
  let ws = _getSheet(ss, "📋 Order_Lines") || ss.insertSheet("📋 Order_Lines");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#2E75B6");

  const DARK_BLUE = "#1F4E79", WHITE = "#FFFFFF", LIGHT_BLUE = "#DEEAF1";
  const headers = ["Line_ID", "Order_ID", "Flavor_ID", "Flavor_Name", "Qty_Ordered"];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_BLUE)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 35);

  ws.setColumnWidth(1, 100);
  ws.setColumnWidth(2, 110);
  ws.setColumnWidth(3, 100);
  ws.setColumnWidth(4, 150);
  ws.setColumnWidth(5, 120);

  ws.setFrozenRows(1);
  ws.setFrozenColumns(2);

  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A2<>"")`)
      .setBackground(LIGHT_BLUE)
      .setRanges([ws.getRange(2, 1, 500, headers.length)])
      .build()
  ]);

  Logger.log("✅ Order_Lines sheet built.");
}

// ────────────────────────────────────────────────────────────
//  🚚  DELIVERY LINES SHEET — Phase 2 narrow schema
// ────────────────────────────────────────────────────────────
function _buildDeliveryLinesSheet(ss) {
  let ws = _getSheet(ss, "🚚 Delivery_Lines") || ss.insertSheet("🚚 Delivery_Lines");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#70AD47");

  const DARK_GREEN = "#375623", WHITE = "#FFFFFF", LIGHT_GREEN = "#E2EFDA";
  const headers = ["Line_ID", "Delivery_ID", "Order_ID", "Flavor_ID", "Flavor_Name", "Qty_Delivered"];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_GREEN)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 35);

  ws.setColumnWidth(1, 100);
  ws.setColumnWidth(2, 120);
  ws.setColumnWidth(3, 110);
  ws.setColumnWidth(4, 100);
  ws.setColumnWidth(5, 150);
  ws.setColumnWidth(6, 120);

  ws.setFrozenRows(1);
  ws.setFrozenColumns(2);

  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A2<>"")`)
      .setBackground(LIGHT_GREEN)
      .setRanges([ws.getRange(2, 1, 500, headers.length)])
      .build()
  ]);

  Logger.log("✅ Delivery_Lines sheet built.");
}

// ────────────────────────────────────────────────────────────
//  ↩️  RETURN LINES SHEET — Phase 2 narrow schema
// ────────────────────────────────────────────────────────────
function _buildReturnLinesSheet(ss) {
  let ws = _getSheet(ss, "↩️ Return_Lines") || ss.insertSheet("↩️ Return_Lines");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#ED7D31");

  const DARK_ORANGE = "#843C0C", WHITE = "#FFFFFF", LIGHT_ORANGE = "#FCE4D6";
  const headers = ["Line_ID", "Return_ID", "Flavor_ID", "Flavor_Name", "Qty_Returned"];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_ORANGE)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 35);

  ws.setColumnWidth(1, 100);
  ws.setColumnWidth(2, 110);
  ws.setColumnWidth(3, 100);
  ws.setColumnWidth(4, 150);
  ws.setColumnWidth(5, 120);

  ws.setFrozenRows(1);
  ws.setFrozenColumns(2);

  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A2<>"")`)
      .setBackground(LIGHT_ORANGE)
      .setRanges([ws.getRange(2, 1, 500, headers.length)])
      .build()
  ]);

  Logger.log("✅ Return_Lines sheet built.");
}

// ────────────────────────────────────────────────────────────
//  🌱  PRODUCTION LINES SHEET — Phase 2 narrow schema
// ────────────────────────────────────────────────────────────
function _buildProductionLinesSheet(ss) {
  let ws = _getSheet(ss, "🌱 Production_Lines") || ss.insertSheet("🌱 Production_Lines");
  ws.clear();
  ws.clearFormats();
  ws.setTabColor("#70AD47");

  const DARK_GREEN = "#375623", WHITE = "#FFFFFF", LIGHT_GREEN = "#E2EFDA";
  const headers = ["Line_ID", "Production_ID", "Flavor_ID", "Flavor_Name", "Qty_Produced"];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(DARK_GREEN)
    .setFontColor(WHITE)
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Arial")
    .setFontSize(10);
  ws.setRowHeight(1, 35);

  ws.setColumnWidth(1, 100);
  ws.setColumnWidth(2, 120);
  ws.setColumnWidth(3, 100);
  ws.setColumnWidth(4, 150);
  ws.setColumnWidth(5, 120);

  ws.setFrozenRows(1);
  ws.setFrozenColumns(2);

  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(MOD(ROW(),2)=0,A2<>"")`)
      .setBackground(LIGHT_GREEN)
      .setRanges([ws.getRange(2, 1, 500, headers.length)])
      .build()
  ]);

  Logger.log("✅ Production_Lines sheet built.");
}

// ────────────────────────────────────────────────────────────
//  ⚙️  GET CONFIG — reads Flavors + Distributors sheets at runtime
//  • Read-only, no side effects — safe to call from any backend file
//  • Excludes rows where Is_Active is not TRUE
// ────────────────────────────────────────────────────────────
function _getConfig(ss) {
  const flavorsSheet = _getSheet(ss, "Flavors");
  const flavors = flavorsSheet.getDataRange().getValues()
    .slice(1)
    .filter(row => row[3] === true)
    .map(row => row[2]);

  const distributorsSheet = _getSheet(ss, "Distributors");
  const distributors = distributorsSheet.getDataRange().getValues()
    .slice(1)
    .filter(row => row[5] === true)
    .map(row => ({
      id: row[0],
      name: row[1],
      slug: row[2],
      phone: row[3],
      apiKey: row[4]
    }));

  return { flavors, distributors };
}
