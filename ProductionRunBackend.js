// ============================================================
//  BADR EL-DIN FARMS — PRODUCTION RUN BACKEND
//  Handles Production Lifecycle run submission (PT3) — the
//  engineer's actual-output record against a Production Request
//  ─────────────────────────────────────────────────────────
//  PROJECT FILE STRUCTURE:
//  • Code.gs                     → Sheet/form setup, onFormSubmit, WhatsApp helper
//  • Router.gs                   → Single doGet/doPost, routes to backends
//  • ProductionBackend.gs        → Production In / Write-off logic only (unrelated — 🌱 Productions)
//  • ProductionTracker.gs        → Production Tracker spreadsheet creation/access + sheet builders
//  • ProductionRequestBackend.gs → submitProductionRequest() — PT2, request intake only
//  • ProductionRunBackend.gs     → This file — production run submission logic only
//  ─────────────────────────────────────────────────────────
//  This file wires PRUN_COUNTER/PRUNL_COUNTER to a caller for the
//  first time (see CLAUDE.md §7I). NOT yet wired to Router.gs or any
//  frontend — backend function only.
//  ─────────────────────────────────────────────────────────
//  SCHEMA — lives in the separate "Badr El-Din Production Tracker"
//  spreadsheet (PRODUCTION_SPREADSHEET_ID), NOT the main Order Tracker:
//  • 🏭 Production_Runs      → header row per run: Run_ID, Request_ID,
//                              Date, Lot_Number, Total_Produced, QC_Status, Notes
//  • 🏭 Production_Run_Lines → one row per non-zero flavor per run:
//                              Line_ID (PRUNL-XXX), Run_ID (FK),
//                              Flavor_ID (FK), Flavor_Name, Qty_Produced
//  • Run_ID generated via _nextSequentialId("PRUN_COUNTER", "PRUN-")
//  • Line_ID generated via _nextSequentialId("PRUNL_COUNTER", "PRUNL-")
//  • QC_Status starts at "Pending" — flipped to Approved/Rejected in a
//    later sub-task (QC_Records submission, not this one)
//  ─────────────────────────────────────────────────────────
//  Flavor_ID/Flavor_Name for the lines come from 🌿 Flavors in the
//  MAIN spreadsheet (not the Production Tracker), same runtime
//  Flavor_Name → Flavor_ID lookup pattern as every other *_Lines table.
//  ─────────────────────────────────────────────────────────
//  Request_ID is a real foreign key, not free text: if payload.requestId
//  doesn't match a row in 🏭 Production_Requests, the whole submission
//  is rejected before anything is written — a run must always be
//  traceable back to the request that authorized it.
//  ─────────────────────────────────────────────────────────
//  No validation against the original request's per-flavor quantities —
//  actual output is allowed to vary from what was requested (over-run,
//  under-run, substitution). This is deliberate, not an oversight.
//  ─────────────────────────────────────────────────────────
//  EXPECTED PAYLOAD:
//  {
//    action:    "submitProductionRun",
//    requestId: "PRQ-001",
//    date:      "2026-08-21",
//    lotNumber: "LOT-2026-08-21-A",
//    flavors: {
//      "Mango": 18,
//      "Blueberry": 0,
//      ... (all active flavors, zeros included)
//    },
//    notes: "Ran short on Mango stock, made up with extra Lemon-Mint"
//  }
// ============================================================

// ────────────────────────────────────────────────────────────
//    SUBMIT PRODUCTION RUN — Called by Router doPost (once wired)
//    Writes a header row to 🏭 Production_Runs and line rows to
//    🏭 Production_Run_Lines, then flips the matching
//    🏭 Production_Requests row's Status to "Fulfilled"
// ────────────────────────────────────────────────────────────
function submitProductionRun(payload) {
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
    const runSheet  = _getSheet(prodSs, "🏭 Production_Runs");
    const lineSheet = _getSheet(prodSs, "🏭 Production_Run_Lines");

    // Validate Request_ID up front — foreign key, not free text.
    // Reject the whole submission before writing anything if it
    // doesn't match a real request, and keep the matched row index
    // so the later Status flip doesn't need a second scan.
    const reqData = reqSheet.getDataRange().getValues();
    let reqRowIndex = -1;
    for (let i = 1; i < reqData.length; i++) {
      if (String(reqData[i][0]) === String(payload.requestId)) {
        reqRowIndex = i;
        break;
      }
    }
    if (reqRowIndex === -1) {
      throw new Error(`Production Request ${payload.requestId} not found — cannot log a run against an unknown request.`);
    }

    // Flavor quantities in config.flavors order
    const flavorQtys = config.flavors.map(f => Number(payload.flavors[f]) || 0);
    const total       = flavorQtys.reduce((sum, q) => sum + q, 0);

    // Generate Run ID
    const runId = _nextSequentialId("PRUN_COUNTER", "PRUN-");

    // Write header row: Run_ID | Request_ID | Date | Lot_Number | Total_Produced | QC_Status | Notes
    const nextRow = runSheet.getLastRow() + 1;
    const rowData  = [runId, payload.requestId, new Date(payload.date), payload.lotNumber, total, "Pending", payload.notes || ""];
    runSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    runSheet.getRange(nextRow, 3).setNumberFormat("yyyy-mm-dd");

    // Re-apply QC_Status dropdown to the new row only (bulk range 2–501
    // already pre-applied by the sheet builder) — QC_Status transitions
    // over time (Pending → Approved/Rejected), same convention as
    // submitProductionRequest's Status re-apply.
    runSheet.getRange(nextRow, 6).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(["Pending", "Approved", "Rejected"], true)
        .build()
    );

    // Write line rows: Line_ID | Run_ID | Flavor_ID | Flavor_Name | Qty_Produced
    const lineRows = [];
    config.flavors.forEach((flavor, i) => {
      if (flavorQtys[i] > 0) {
        lineRows.push([_nextSequentialId("PRUNL_COUNTER", "PRUNL-"), runId, flavorIdMap[flavor], flavor, flavorQtys[i]]);
      }
    });
    if (lineRows.length > 0) {
      const lineNextRow = lineSheet.getLastRow() + 1;
      lineSheet.getRange(lineNextRow, 1, lineRows.length, lineRows[0].length).setValues(lineRows);
    }

    // Flip the request's Status to Fulfilled, as soon as a run is
    // submitted against it — regardless of QC outcome. Row already
    // located above; same spreadsheet, so no non-blocking try/catch
    // here (that treatment is reserved for genuinely cross-spreadsheet
    // writes, e.g. DeliveryBackend.gs's Inventory sync).
    reqSheet.getRange(reqRowIndex + 1, 4).setValue("Fulfilled");

    // Notify Quality Engineers (+ unassigned-role contacts, fail-open)
    _notifyProductionRun(mainSs, runId, payload.requestId, payload.lotNumber, config, flavorQtys, total, payload.notes);

    Logger.log(`✅ Production Run ${runId} logged against request ${payload.requestId} (lot ${payload.lotNumber})`);
    return { runId, requestId: payload.requestId, total };

  } catch (err) {
    Logger.log("❌ submitProductionRun error: " + err.toString());
    throw new Error(err.toString());
  }
}

// ────────────────────────────────────────────────────────────
//    NOTIFY ON PRODUCTION RUN
//    Reads 📱 Team_Contacts (main spreadsheet) and sends to every
//    active contact whose Role is "Quality Engineer" OR blank.
//    Deliberate fail-open default: an unset Role means "notify on
//    everything" until roles are actually assigned, not "skip."
// ────────────────────────────────────────────────────────────
function _notifyProductionRun(mainSs, runId, requestId, lotNumber, config, flavorQtys, total, notes) {
  const contactsSheet = _getSheet(mainSs, "Team_Contacts");
  const recipients = contactsSheet.getDataRange().getValues().slice(1)
    .filter(r => r[4] === true && (r[5] === "Quality Engineer" || !r[5]))
    .map(r => ({ phone: r[2], apiKey: r[3] }));

  let summary = "";
  config.flavors.forEach((f, i) => {
    if (flavorQtys[i] > 0) summary += `  • ${f}: ${flavorQtys[i]} cases\n`;
  });

  const msg =
    `🏭 *Production Run Logged*\n` +
    `Run ID: *${runId}*\n` +
    `Request ID: *${requestId}*\n` +
    `Lot Number: *${lotNumber}*\n\n` +
    `*Produced:*\n${summary}\n` +
    `Total: *${total} cases*` +
    (notes ? `\nNotes: ${notes}` : "");

  recipients.forEach(c => _sendWhatsApp(c.phone, c.apiKey, msg));
}
