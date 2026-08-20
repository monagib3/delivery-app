// ============================================================
//  BADR EL-DIN FARMS — PRODUCTION REQUEST BACKEND
//  Handles Production Lifecycle request submission (PT2)
//  ─────────────────────────────────────────────────────────
//  PROJECT FILE STRUCTURE:
//  • Code.gs                     → Sheet/form setup, onFormSubmit, WhatsApp helper
//  • Router.gs                   → Single doGet/doPost, routes to backends
//  • OrderBackend.gs             → Order submission logic only
//  • ProductionBackend.gs        → Production In / Write-off logic only (unrelated — 🌱 Productions)
//  • ProductionTracker.gs        → Production Tracker spreadsheet creation/access + sheet builders
//  • ProductionRequestBackend.gs → This file — production request submission logic only
//  ─────────────────────────────────────────────────────────
//  ⚠️  SCHEMA ONLY UNTIL NOW (see CLAUDE.md §7I) — this file is what
//  wires PRQ_COUNTER/PRQL_COUNTER to a caller for the first time.
//  NOT yet wired to Router.gs or any frontend — backend function only.
//  ─────────────────────────────────────────────────────────
//  SCHEMA — lives in the separate "Badr El-Din Production Tracker"
//  spreadsheet (PRODUCTION_SPREADSHEET_ID), NOT the main Order Tracker:
//  • 🏭 Production_Requests      → header row per request: Request_ID,
//                                  Timestamp, Requested_By, Status, Notes
//  • 🏭 Production_Request_Lines → one row per non-zero flavor per request:
//                                  Line_ID (PRQL-XXX), Request_ID (FK),
//                                  Flavor_ID (FK), Flavor_Name, Qty_Requested
//  • Request_ID generated via _nextSequentialId("PRQ_COUNTER", "PRQ-")
//  • Line_ID generated via _nextSequentialId("PRQL_COUNTER", "PRQL-")
//  • Status starts at "Requested"
//  ─────────────────────────────────────────────────────────
//  Flavor_ID/Flavor_Name for the lines come from 🌿 Flavors in the
//  MAIN spreadsheet (not the Production Tracker) via the same runtime
//  Flavor_Name → Flavor_ID lookup pattern used by Order_Lines/
//  Delivery_Lines/Return_Lines/Production_Lines.
//  ─────────────────────────────────────────────────────────
//  EXPECTED PAYLOAD:
//  {
//    action:      "submitProductionRequest",
//    requestedBy: "Ops Manager",
//    flavors: {
//      "Mango": 20,
//      "Blueberry": 0,
//      ... (all active flavors, zeros included)
//    },
//    notes: "Need before Thursday delivery run"
//  }
// ============================================================

// ────────────────────────────────────────────────────────────
//    SUBMIT PRODUCTION REQUEST — Called by Router doPost (once wired)
//    Writes a header row to 🏭 Production_Requests and line rows to
//    🏭 Production_Request_Lines in the Production Tracker spreadsheet
// ────────────────────────────────────────────────────────────
function submitProductionRequest(payload) {
  try {
    const mainSs = SpreadsheetApp.openById(
      PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
    );
    const flavorsSheet = _getSheet(mainSs, "🌿 Flavors");
    const config        = _getConfig(mainSs);

    // Flavor_Name → Flavor_ID map
    const flavorIdMap = {};
    flavorsSheet.getDataRange().getValues().slice(1).forEach(r => { flavorIdMap[r[2]] = r[0]; });

    const prodSs    = _getProductionSpreadsheet();
    const reqSheet  = _getSheet(prodSs, "🏭 Production_Requests");
    const lineSheet = _getSheet(prodSs, "🏭 Production_Request_Lines");

    // Flavor quantities in config.flavors order
    const flavorQtys = config.flavors.map(f => Number(payload.flavors[f]) || 0);
    const total       = flavorQtys.reduce((sum, q) => sum + q, 0);

    // Generate Request ID
    const requestId = _nextSequentialId("PRQ_COUNTER", "PRQ-");
    const timestamp = new Date();

    // Write header row: Request_ID | Timestamp | Requested_By | Status | Notes
    const nextRow = reqSheet.getLastRow() + 1;
    const rowData  = [requestId, timestamp, payload.requestedBy, "Requested", payload.notes || ""];
    reqSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    reqSheet.getRange(nextRow, 2).setNumberFormat("yyyy-mm-dd hh:mm");

    // Re-apply Status dropdown to the new row only (bulk range 2–501 already
    // pre-applied by the sheet builder) — Status transitions over time
    // (Requested → Fulfilled), same convention as submitOrder's Orders.Status.
    reqSheet.getRange(nextRow, 4).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(["Requested", "Fulfilled"], true)
        .build()
    );

    // Write line rows: Line_ID | Request_ID | Flavor_ID | Flavor_Name | Qty_Requested
    const lineRows = [];
    config.flavors.forEach((flavor, i) => {
      if (flavorQtys[i] > 0) {
        lineRows.push([_nextSequentialId("PRQL_COUNTER", "PRQL-"), requestId, flavorIdMap[flavor], flavor, flavorQtys[i]]);
      }
    });
    if (lineRows.length > 0) {
      const lineNextRow = lineSheet.getLastRow() + 1;
      lineSheet.getRange(lineNextRow, 1, lineRows.length, lineRows[0].length).setValues(lineRows);
    }

    // Notify Production Engineers (+ unassigned-role contacts, fail-open)
    _notifyProductionRequest(mainSs, requestId, payload.requestedBy, config, flavorQtys, total, payload.notes);

    Logger.log(`✅ Production Request ${requestId} logged by ${payload.requestedBy}`);
    return { requestId, requestedBy: payload.requestedBy, total };

  } catch (err) {
    Logger.log("❌ submitProductionRequest error: " + err.toString());
    throw new Error(err.toString());
  }
}

// ────────────────────────────────────────────────────────────
//    NOTIFY ON PRODUCTION REQUEST
//    Reads 📱 Team_Contacts (main spreadsheet) and sends to every
//    active contact whose Role is "Production Engineer" OR blank.
//    Deliberate fail-open default: an unset Role means "notify on
//    everything" until roles are actually assigned, not "skip."
// ────────────────────────────────────────────────────────────
function _notifyProductionRequest(mainSs, requestId, requestedBy, config, flavorQtys, total, notes) {
  const contactsSheet = _getSheet(mainSs, "Team_Contacts");
  const recipients = contactsSheet.getDataRange().getValues().slice(1)
    .filter(r => r[4] === true && (r[5] === "Production Engineer" || !r[5]))
    .map(r => ({ phone: r[2], apiKey: r[3] }));

  let summary = "";
  config.flavors.forEach((f, i) => {
    if (flavorQtys[i] > 0) summary += `  • ${f}: ${flavorQtys[i]} cases\n`;
  });

  const msg =
    `🏭 *Production Request*\n` +
    `Request ID: *${requestId}*\n` +
    `Requested By: *${requestedBy}*\n\n` +
    `*Requested:*\n${summary}\n` +
    `Total: *${total} cases*` +
    (notes ? `\nNotes: ${notes}` : "");

  recipients.forEach(c => _sendWhatsApp(c.phone, c.apiKey, msg));
}
