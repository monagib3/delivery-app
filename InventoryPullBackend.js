// ============================================================
//  BADR EL-DIN FARMS — INVENTORY PULL BACKEND
//  Handles Production Lifecycle inventory pull confirmation (PT5) —
//  the handoff from an Approved run into the EXISTING warehouse-in
//  logic (submitProduction / 🌱 Productions / Inventory Movements)
//  ─────────────────────────────────────────────────────────
//  PROJECT FILE STRUCTURE:
//  • Code.gs                     → Sheet/form setup, onFormSubmit, WhatsApp helper
//  • Router.gs                   → Single doGet/doPost, routes to backends
//  • ProductionBackend.gs        → submitProduction() — Production In / Write-off,
//                                  called BY this file, never reimplemented here
//  • ProductionTracker.gs        → Production Tracker spreadsheet creation/access + sheet builders
//  • ProductionRequestBackend.gs → submitProductionRequest() — PT2, request intake only
//  • ProductionRunBackend.gs     → submitProductionRun() — PT3, actual-output logging only
//  • QualityBackend.gs           → submitQCDecision() — PT4, QC decision only
//  • InventoryPullBackend.gs     → This file — inventory pull confirmation only
//  ─────────────────────────────────────────────────────────
//  NOT yet wired to Router.gs or any frontend — backend function only.
//  ─────────────────────────────────────────────────────────
//  This is deliberately a thin handoff, not a parallel write path:
//  it calls the EXISTING submitProduction() (ProductionBackend.gs)
//  as-is to do the actual Productions/Production_Lines write, team
//  notification, and Inventory Movements logging — none of that is
//  reimplemented here. This file only (a) translates a Production
//  Run into the payload shape submitProduction() already expects,
//  and (b) marks the run as pulled afterward.
//  ─────────────────────────────────────────────────────────
//  submitProduction()'s ACTUAL expected payload shape (confirmed by
//  reading ProductionBackend.gs and Inventory.gs's
//  _logProductionInMovements, not assumed):
//    { qtys: [...], lotNumber: "...", notes: "...", productionDate: "..." (optional) }
//  qtys is a plain array, positionally aligned to _getConfig(mainSs)
//  .flavors order — NOT a {flavor: qty} object like the PT2/PT3
//  payload shapes.
//  ─────────────────────────────────────────────────────────
//  Run_ID is a real foreign key against 🏭 Production_Runs — hard
//  error before anything runs if it doesn't match, same pattern as
//  ProductionRunBackend.gs/QualityBackend.gs.
//  Also hard errors if QC_Status isn't "Approved" (can't pull what
//  wasn't approved) or if Inventory_Pulled_By is already set (no
//  double-pulling the same run).
//  ─────────────────────────────────────────────────────────
//  productionDate is deliberately set to the run's own Date column
//  (from 🏭 Production_Runs, itself sourced from PT3's payload.date),
//  not "now" — the Inventory Movements ledger should reflect when
//  the batch was actually produced, not when QC/pull happened to
//  catch up. This is a deliberate deviation from every other current
//  submitProduction() caller, which all omit productionDate and let
//  it default to "now" inside _logProductionInMovements.
//  ─────────────────────────────────────────────────────────
//  The existing 🌱 Productions table has no FK column back to this
//  Production Lifecycle (and none is being added), so traceability
//  is carried in the notes passed to submitProduction() instead:
//  "Pulled from Run <runId> (Request <requestId>)".
//  ─────────────────────────────────────────────────────────
//  EXPECTED PAYLOAD:
//  {
//    action:   "submitInventoryPull",
//    runId:    "PRUN-001",
//    pulledBy: "Inventory Clerk"
//  }
// ============================================================

// ────────────────────────────────────────────────────────────
//    SUBMIT INVENTORY PULL — Called by Router doPost (once wired)
//    Translates an Approved Production Run into submitProduction()'s
//    payload shape, calls it as-is, then marks the run as pulled
// ────────────────────────────────────────────────────────────
function submitInventoryPull(payload) {
  try {
    const prodSs    = _getProductionSpreadsheet();
    const runSheet  = _getSheet(prodSs, "🏭 Production_Runs");
    const lineSheet = _getSheet(prodSs, "🏭 Production_Run_Lines");

    // Validate Run_ID up front — foreign key, not free text.
    // Reject before running anything if it doesn't match a real run,
    // and keep the matched row index for the guard checks below and
    // the later Inventory_Pulled_By/Timestamp write.
    const runData = runSheet.getDataRange().getValues();
    let runRowIndex = -1;
    for (let i = 1; i < runData.length; i++) {
      if (String(runData[i][0]) === String(payload.runId)) {
        runRowIndex = i;
        break;
      }
    }
    if (runRowIndex === -1) {
      throw new Error(`Production Run ${payload.runId} not found — cannot pull inventory for an unknown run.`);
    }

    const runRow = runData[runRowIndex];
    const requestId  = runRow[1];
    const runDate    = runRow[2];
    const lotNumber  = runRow[3];
    const qcStatus   = runRow[5];
    const pulledByExisting = runRow[7];

    if (qcStatus !== "Approved") {
      throw new Error(`Production Run ${payload.runId} has QC_Status "${qcStatus}" — can only pull inventory for an Approved run.`);
    }
    if (pulledByExisting) {
      throw new Error(`Production Run ${payload.runId} was already pulled by "${pulledByExisting}" — no double-pulling the same run.`);
    }

    // Pre-fill, don't retype: read per-flavor quantities off
    // 🏭 Production_Run_Lines instead of asking the caller for them.
    const mainSs = SpreadsheetApp.openById(
      PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
    );
    const config = _getConfig(mainSs);

    const flavorQtyMap = {};
    lineSheet.getDataRange().getValues().slice(1)
      .filter(r => r[1] === payload.runId)
      .forEach(r => { flavorQtyMap[r[3]] = r[4]; });
    const qtys = config.flavors.map(f => Number(flavorQtyMap[f]) || 0);

    // Hand off to the EXISTING warehouse-in logic as-is — it owns
    // the Productions/Production_Lines write, the team notification,
    // and the Inventory Movements logging. Not reimplemented here.
    const result = submitProduction({
      qtys:           qtys,
      lotNumber:      lotNumber,
      notes:          `Pulled from Run ${payload.runId} (Request ${requestId})`,
      productionDate: runDate
    });

    // Mark the run as pulled. Row already located above; same
    // spreadsheet, so no non-blocking try/catch here (that treatment
    // is reserved for genuinely cross-spreadsheet writes).
    const now = new Date();
    runSheet.getRange(runRowIndex + 1, 8, 1, 2).setValues([[payload.pulledBy, now]]);
    runSheet.getRange(runRowIndex + 1, 9).setNumberFormat("yyyy-mm-dd hh:mm");

    Logger.log(`✅ Inventory pulled for run ${payload.runId} → Production ${result.productionId}`);
    return { productionId: result.productionId, runId: payload.runId };

  } catch (err) {
    Logger.log("❌ submitInventoryPull error: " + err.toString());
    throw new Error(err.toString());
  }
}

