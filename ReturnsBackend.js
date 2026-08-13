// ============================================================
//  BADR EL-DIN FARMS — RETURNS BACKEND
//  Handles all return-related logic for the internal web app
//  ─────────────────────────────────────────────────────────
//  PROJECT FILE STRUCTURE:
//  • Code.gs            → Sheet/form setup, onFormSubmit, WhatsApp helper
//  • Router.gs          → Single doGet/doPost, routes to backends
//  • OrderBackend.gs    → Order submission logic only
//  • DeliveryBackend.gs → Delivery logic only
//  • ReturnsBackend.gs  → This file — returns logic only
//  ─────────────────────────────────────────────────────────
//  SCHEMA (Phase 2 — narrow tables):
//  • ↩️ Returns       → header row per return: Return_ID, Date,
//                       Distributor, Total_Returned, Notes
//  • ↩️ Return_Lines  → one row per non-zero flavor per return:
//                       Line_ID (RL-XXX), Return_ID (FK), Flavor_ID (FK),
//                       Flavor_Name, Qty_Returned
//  • Return_ID generated via _nextSequentialId("RET_COUNTER", "RET-")
//  • Line_ID generated via _nextSequentialId("RL_COUNTER", "RL-")
//  ─────────────────────────────────────────────────────────
//  Returns are entered by the internal team, not distributors —
//  no Order ID reference, no distributor-facing WhatsApp message.
// ============================================================

// ────────────────────────────────────────────────────────────
//    SUBMIT RETURN — Called by Router doPost
//    Logs a return row and syncs a positive "Return" movement
// ────────────────────────────────────────────────────────────
function submitReturn(payload) {
  try {
    const ss           = SpreadsheetApp.openById(
      PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
    );
    const retSheet      = _getSheet(ss, "↩️ Returns");
    const lineSheet     = _getSheet(ss, "↩️ Return_Lines");
    const flavorsSheet  = _getSheet(ss, "🌿 Flavors");
    const config        = _getConfig(ss);

    // Flavor_Name → Flavor_ID map
    const flavorIdMap = {};
    flavorsSheet.getDataRange().getValues().slice(1).forEach(r => { flavorIdMap[r[2]] = r[0]; });

    // Generate Return ID
    const returnId = _nextSequentialId("RET_COUNTER", "RET-");
    const now      = new Date();
    const total    = payload.qtys.reduce((s, q) => s + (Number(q) || 0), 0);

    // Write header row: Return_ID | Date | Distributor | Total_Returned | Notes
    const nextRow = retSheet.getLastRow() + 1;
    const rowData = [returnId, now, payload.distributor, total, payload.notes || ""];
    retSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    retSheet.getRange(nextRow, 2).setNumberFormat("yyyy-mm-dd");

    // Write line rows: Line_ID | Return_ID | Flavor_ID | Flavor_Name | Qty_Returned
    const lineRows = [];
    config.flavors.forEach((flavor, i) => {
      const qty = Number(payload.qtys[i]) || 0;
      if (qty > 0) {
        lineRows.push([_nextSequentialId("RL_COUNTER", "RL-"), returnId, flavorIdMap[flavor], flavor, qty]);
      }
    });
    if (lineRows.length > 0) {
      const lineNextRow = lineSheet.getLastRow() + 1;
      lineSheet.getRange(lineNextRow, 1, lineRows.length, lineRows[0].length).setValues(lineRows);
    }

    // Notify internal team
    _notifyReturn(ss, returnId, payload.distributor, payload.qtys, total);

    // Log Return movements to Inventory spreadsheet (non-blocking)
    try {
      _logReturnInMovements(returnId, payload);
    } catch (err) {
      _logSyncError("submitReturn → _logReturnInMovements", err, returnId, payload.distributor, JSON.stringify(payload.qtys));
    }

    Logger.log(`✅ Return ${returnId} logged for ${payload.distributor}`);
    return {
      returnId:    returnId,
      distributor: payload.distributor,
      total:       total
    };

  } catch (err) {
    Logger.log("❌ submitReturn error: " + err.toString());
    throw new Error(err.toString());
  }
}

// ────────────────────────────────────────────────────────────
//    NOTIFY ON RETURN
//    Team-only WhatsApp — returns are team-entered, not distributor-facing
// ────────────────────────────────────────────────────────────
function _notifyReturn(ss, returnId, distributorName, qtys, total) {
  const config = _getConfig(ss);
  let summary = "";
  config.flavors.forEach((f, i) => {
    if (qtys[i] > 0) summary += `  • ${f}: ${qtys[i]} cases\n`;
  });

  const teamMsg =
    `↩️ *Return Logged*\n` +
    `Return ID: *${returnId}*\n` +
    `Distributor: *${distributorName}*\n\n` +
    `*Returned:*\n${summary}\n` +
    `Total: *${total} cases*`;
  CONFIG.TEAM_WHATSAPP.forEach(m => _sendWhatsApp(m.phone, m.apiKey, teamMsg));
}
