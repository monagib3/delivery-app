// ============================================================
//  BADR EL-DIN FARMS — QUALITY BACKEND
//  Handles Production Lifecycle QC decisions (PT4) — the final
//  Approved/Rejected call against a Production Run
//  ─────────────────────────────────────────────────────────
//  PROJECT FILE STRUCTURE:
//  • Code.gs                     → Sheet/form setup, onFormSubmit, WhatsApp helper
//  • Router.gs                   → Single doGet/doPost, routes to backends
//  • ProductionBackend.gs        → Production In / Write-off logic only (unrelated — 🌱 Productions)
//  • ProductionTracker.gs        → Production Tracker spreadsheet creation/access + sheet builders
//  • ProductionRequestBackend.gs → submitProductionRequest() — PT2, request intake only
//  • ProductionRunBackend.gs     → submitProductionRun() — PT3, actual-output logging only
//  • QualityBackend.gs           → This file — QC decision logic only
//  ─────────────────────────────────────────────────────────
//  This file wires QC_COUNTER to a caller for the first time
//  (see CLAUDE.md §7I). NOT yet wired to Router.gs or any
//  frontend — backend function only.
//  ─────────────────────────────────────────────────────────
//  SCHEMA — lives in the separate "Badr El-Din Production Tracker"
//  spreadsheet (PRODUCTION_SPREADSHEET_ID), NOT the main Order Tracker:
//  • 🏭 QC_Records → one row per QC decision: QC_ID, Run_ID, Timestamp,
//                    QC_By, Decision, Notes
//  • QC_ID generated via _nextSequentialId("QC_COUNTER", "QC-")
//  ─────────────────────────────────────────────────────────
//  Run_ID is a real foreign key, not free text: if payload.runId
//  doesn't match a row in 🏭 Production_Runs, the whole submission
//  is rejected before anything is written — same pattern as
//  ProductionRunBackend.gs's Request_ID check.
//  ─────────────────────────────────────────────────────────
//  QC_Status is a ONE-WAY lock: a run can only be decided while its
//  QC_Status is "Pending". Once flipped to Approved OR Rejected, no
//  re-deciding in either direction — hard error, not a silent no-op.
//  This is symmetric on purpose: Approved is expected to trigger a
//  downstream Inventory pull in a later sub-task, so allowing a
//  flip-back afterward would leave that action stranded/inconsistent.
//  ─────────────────────────────────────────────────────────
//  decision must be exactly "Approved" or "Rejected" — hard error on
//  anything else, before any write. ("Pending" is a valid QC_Records
//  dropdown value for sheet-editing purposes, per the PT1 schema, but
//  is never a valid submitted decision — a QC check that hasn't
//  happened yet has nothing to record.)
//  ─────────────────────────────────────────────────────────
//  Both outcomes notify Inventory contacts — Approved says "ready to
//  pull into warehouse", Rejected says "failed QC, needs manual
//  write-off." This function does NOT trigger any Inventory movement
//  or write-off itself — that stays a manual action through the
//  existing inventory-ops.html write-off flow. Notification only.
//  ─────────────────────────────────────────────────────────
//  EXPECTED PAYLOAD:
//  {
//    action:   "submitQCDecision",
//    runId:    "PRUN-001",
//    decision: "Approved",
//    qcBy:     "Quality Engineer",
//    notes:    "Batch within spec, ready for warehouse"
//  }
// ============================================================

// ────────────────────────────────────────────────────────────
//    SUBMIT QC DECISION — Called by Router doPost (once wired)
//    Writes a row to 🏭 QC_Records, then flips the matching
//    🏭 Production_Runs row's QC_Status to match Decision
// ────────────────────────────────────────────────────────────
function submitQCDecision(payload) {
  try {
    if (payload.decision !== "Approved" && payload.decision !== "Rejected") {
      throw new Error(`Invalid decision "${payload.decision}" — must be exactly "Approved" or "Rejected".`);
    }

    const prodSs    = _getProductionSpreadsheet();
    const runSheet  = _getSheet(prodSs, "🏭 Production_Runs");
    const qcSheet   = _getSheet(prodSs, "🏭 QC_Records");

    // Validate Run_ID up front — foreign key, not free text.
    // Reject the whole submission before writing anything if it
    // doesn't match a real run, and keep the matched row index so
    // the lock check and the later QC_Status flip don't need a
    // second scan.
    const runData = runSheet.getDataRange().getValues();
    let runRowIndex = -1;
    for (let i = 1; i < runData.length; i++) {
      if (String(runData[i][0]) === String(payload.runId)) {
        runRowIndex = i;
        break;
      }
    }
    if (runRowIndex === -1) {
      throw new Error(`Production Run ${payload.runId} not found — cannot record a QC decision against an unknown run.`);
    }

    // One-way lock: only a "Pending" run can be decided.
    const currentQcStatus = runData[runRowIndex][5];
    if (currentQcStatus !== "Pending") {
      throw new Error(`Production Run ${payload.runId} already has QC_Status "${currentQcStatus}" — decisions are final, no re-deciding.`);
    }

    const requestId = runData[runRowIndex][1];
    const lotNumber = runData[runRowIndex][3];

    // Generate QC ID
    const qcId = _nextSequentialId("QC_COUNTER", "QC-");
    const timestamp = new Date();

    // Write row: QC_ID | Run_ID | Timestamp | QC_By | Decision | Notes
    const nextRow = qcSheet.getLastRow() + 1;
    const rowData  = [qcId, payload.runId, timestamp, payload.qcBy, payload.decision, payload.notes || ""];
    qcSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    qcSheet.getRange(nextRow, 3).setNumberFormat("yyyy-mm-dd hh:mm");

    // Re-apply Decision dropdown to the new row only (bulk range 2–501
    // already pre-applied by the sheet builder), consistent with the
    // Status/QC_Status re-apply convention in the earlier sub-tasks.
    qcSheet.getRange(nextRow, 5).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(["Pending", "Approved", "Rejected"], true)
        .build()
    );

    // Flip the run's QC_Status to match Decision. Row already located
    // above; same spreadsheet, so no non-blocking try/catch here (that
    // treatment is reserved for genuinely cross-spreadsheet writes).
    runSheet.getRange(runRowIndex + 1, 6).setValue(payload.decision);

    // Notify Inventory (+ unassigned-role contacts, fail-open)
    _notifyQCDecision(payload.runId, requestId, lotNumber, payload.decision, payload.notes);

    Logger.log(`✅ QC Decision ${qcId} recorded for run ${payload.runId}: ${payload.decision}`);
    return { qcId, runId: payload.runId, decision: payload.decision };

  } catch (err) {
    Logger.log("❌ submitQCDecision error: " + err.toString());
    throw new Error(err.toString());
  }
}

// ────────────────────────────────────────────────────────────
//    NOTIFY ON QC DECISION
//    Reads 🏭 Production_Run_Lines (Production Tracker spreadsheet)
//    for the per-flavor produced quantities, and 📱 Team_Contacts
//    (main spreadsheet) for recipients — every active contact whose
//    Role is "Inventory" OR blank. Deliberate fail-open default: an
//    unset Role means "notify on everything" until roles are
//    actually assigned, not "skip."
// ────────────────────────────────────────────────────────────
function _notifyQCDecision(runId, requestId, lotNumber, decision, notes) {
  const mainSs = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
  );
  const config = _getConfig(mainSs);

  const prodSs        = _getProductionSpreadsheet();
  const runLineSheet  = _getSheet(prodSs, "🏭 Production_Run_Lines");
  const flavorQtyMap  = {};
  runLineSheet.getDataRange().getValues().slice(1)
    .filter(r => r[1] === runId)
    .forEach(r => { flavorQtyMap[r[3]] = r[4]; });

  let summary = "";
  config.flavors.forEach(f => {
    const qty = Number(flavorQtyMap[f]) || 0;
    if (qty > 0) summary += `  • ${f}: ${qty} cases\n`;
  });

  const contactsSheet = _getSheet(mainSs, "Team_Contacts");
  const recipients = contactsSheet.getDataRange().getValues().slice(1)
    .filter(r => r[4] === true && (r[5] === "Inventory" || !r[5]))
    .map(r => ({ phone: r[2], apiKey: r[3] }));

  const msg = decision === "Approved"
    ? `✅ *QC Approved — Ready for Warehouse*\n` +
      `Run ID: *${runId}*\n` +
      `Request ID: *${requestId}*\n` +
      `Lot Number: *${lotNumber}*\n\n` +
      `*Produced:*\n${summary}\n` +
      `Ready to pull into warehouse.` +
      (notes ? `\nNotes: ${notes}` : "")
    : `❌ *QC Rejected — Needs Write-off*\n` +
      `Run ID: *${runId}*\n` +
      `Request ID: *${requestId}*\n` +
      `Lot Number: *${lotNumber}*\n\n` +
      `*Affected:*\n${summary}\n` +
      `Failed QC — needs manual write-off or repurposing (samples/employee gifts) via Inventory Ops.` +
      (notes ? `\nNotes: ${notes}` : "");

  recipients.forEach(c => _sendWhatsApp(c.phone, c.apiKey, msg));
}
