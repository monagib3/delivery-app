// ============================================================
//  BADR EL-DIN FARMS — PRODUCTION BACKEND
//  Handles all production-related logic for the internal web app
//  ─────────────────────────────────────────────────────────
//  PROJECT FILE STRUCTURE:
//  • Code.gs               → Sheet/form setup, onFormSubmit, WhatsApp helper
//  • Router.gs             → Single doGet/doPost, routes to backends
//  • OrderBackend.gs       → Order submission logic only
//  • DeliveryBackend.gs    → Delivery logic only
//  • ReturnsBackend.gs     → Returns logic only
//  • ProductionBackend.gs  → This file — production logic only
//  ─────────────────────────────────────────────────────────
//  SCHEMA (Phase 2 — narrow tables):
//  • 🌱 Productions      → header row per production batch: Production_ID,
//                          Date, Lot_Number, Total_Produced, Notes
//  • 🌱 Production_Lines → one row per non-zero flavor per batch:
//                          Line_ID (PL-XXX), Production_ID (FK), Flavor_ID (FK),
//                          Flavor_Name, Qty_Produced
//  • Production_ID generated via _nextSequentialId("PRD_COUNTER", "PRD-")
//  • Line_ID generated via _nextSequentialId("PL_COUNTER", "PL-")
//  ─────────────────────────────────────────────────────────
//  Production In is entered by the internal team, not distributors —
//  no Order ID reference, no distributor-facing WhatsApp message.
// ============================================================

// ────────────────────────────────────────────────────────────
//    SUBMIT PRODUCTION — Called by Router doPost
//    Logs a production row and syncs a positive "Production In" movement
// ────────────────────────────────────────────────────────────
function submitProduction(payload) {
  try {
    const ss            = SpreadsheetApp.openById(
      PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
    );
    const prodSheet      = _getSheet(ss, "🌱 Productions");
    const lineSheet      = _getSheet(ss, "🌱 Production_Lines");
    const flavorsSheet   = _getSheet(ss, "🌿 Flavors");
    const config         = _getConfig(ss);

    // Flavor_Name → Flavor_ID map
    const flavorIdMap = {};
    flavorsSheet.getDataRange().getValues().slice(1).forEach(r => { flavorIdMap[r[2]] = r[0]; });

    const productionId = _nextSequentialId("PRD_COUNTER", "PRD-");
    const now   = new Date();
    const total = payload.qtys.reduce((s, q) => s + (Number(q) || 0), 0);

    // Write header row: Production_ID | Date | Lot_Number | Total_Produced | Notes
    const nextRow = prodSheet.getLastRow() + 1;
    const rowData = [productionId, now, payload.lotNumber, total, payload.notes || ""];
    prodSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    prodSheet.getRange(nextRow, 2).setNumberFormat("yyyy-mm-dd");

    // Write line rows: Line_ID | Production_ID | Flavor_ID | Flavor_Name | Qty_Produced
    const lineRows = [];
    config.flavors.forEach((flavor, i) => {
      const qty = Number(payload.qtys[i]) || 0;
      if (qty > 0) {
        lineRows.push([_nextSequentialId("PL_COUNTER", "PL-"), productionId, flavorIdMap[flavor], flavor, qty]);
      }
    });
    if (lineRows.length > 0) {
      const lineNextRow = lineSheet.getLastRow() + 1;
      lineSheet.getRange(lineNextRow, 1, lineRows.length, lineRows[0].length).setValues(lineRows);
    }

    // Notify internal team
    _notifyProduction(ss, productionId, payload.lotNumber, payload.qtys, total);

    // Log Production In movements to Inventory spreadsheet (non-blocking)
    try {
      _logProductionInMovements(productionId, payload);
    } catch (err) {
      _logSyncError("submitProduction → _logProductionInMovements", err, productionId, payload.lotNumber, JSON.stringify(payload.qtys));
    }

    Logger.log(`✅ Production ${productionId} logged (lot ${payload.lotNumber})`);
    return {
      productionId: productionId,
      lotNumber:    payload.lotNumber,
      total:        total
    };

  } catch (err) {
    Logger.log("❌ submitProduction error: " + err.toString());
    throw new Error(err.toString());
  }
}

// ────────────────────────────────────────────────────────────
//    NOTIFY ON PRODUCTION
//    Team-only WhatsApp — production is team-entered, not distributor-facing
// ────────────────────────────────────────────────────────────
function _notifyProduction(ss, productionId, lotNumber, qtys, total) {
  const config = _getConfig(ss);
  let summary = "";
  config.flavors.forEach((f, i) => {
    if (qtys[i] > 0) summary += `  • ${f}: ${qtys[i]} cases\n`;
  });

  const teamMsg =
    `🌱 *Production Logged*\n` +
    `Production ID: *${productionId}*\n` +
    `Lot Number: *${lotNumber}*\n\n` +
    `*Produced:*\n${summary}\n` +
    `Total: *${total} cases*`;
  CONFIG.TEAM_WHATSAPP.forEach(m => _sendWhatsApp(m.phone, m.apiKey, teamMsg));
}
