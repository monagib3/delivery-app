// ============================================================
//  BADR EL-DIN FARMS — REQUEST ROUTER
//  Single entry point for all web app GET and POST requests
//  ─────────────────────────────────────────────────────────
//  PROJECT FILE STRUCTURE:
//  • Code.gs              → Sheet/form setup, onFormSubmit, WhatsApp helper
//  • Router.gs            → This file — single doGet/doPost, routes to backends
//  • DeliveryBackend.gs   → Delivery logic only (no doGet/doPost)
//  • OrderBackend.gs      → Order submission logic only (no doGet/doPost)
//  • ReturnsBackend.gs    → Return logic only (no doGet/doPost)
//  • ProductionBackend.gs → Production In logic only (no doGet/doPost)
//  • Inventory.gs         → Inventory spreadsheet + movements, incl.
//                           _logWriteOffMovements (no Backend file/sheet
//                           of its own — called directly from this Router)
//  • ProductionTracker.gs → Production Tracker spreadsheet creation/access,
//                           sheet builders, and the three pending-queue
//                           reads below (no doGet/doPost)
//  • ProductionRequestBackend.gs → submitProductionRequest() (no doGet/doPost)
//  • ProductionRunBackend.gs     → submitProductionRun() (no doGet/doPost)
//  • QualityBackend.gs           → submitQCDecision() (no doGet/doPost)
//  • InventoryPullBackend.gs     → submitInventoryPull() (no doGet/doPost)
//  • DeliveryApp.html     → Delivery app UI (Apps Script hosted, desktop fallback)
//  • order.html           → Order app UI (GitHub Pages hosted, distributors)
//  ─────────────────────────────────────────────────────────
//  ROUTING TABLE:
//  GET  ?action=getOrders           → DeliveryBackend: _getOpenOrders()
//  GET  ?action=getReturnContext    → flavors + distributors for Returns app
//  GET  ?action=getProductionContext → flavors for Production/Write-off app
//  GET  ?action=getDistributors     → distributor slug+name list for order.html
//  GET  ?action=getDashboardData    → DashboardBackend: _getDashboardData()
//  GET  ?action=getPendingProductionRequests → ProductionTracker: _getPendingProductionRequests()
//  GET  ?action=getPendingQCRuns    → ProductionTracker: _getPendingQCRuns()
//  GET  ?action=getPendingInventoryPulls → ProductionTracker: _getPendingInventoryPulls()
//  GET  (no action)                 → Serve DeliveryApp.html (desktop fallback)
//  POST action=submitDelivery       → DeliveryBackend:    submitDelivery()
//  POST action=submitOrder          → OrderBackend:       submitOrder()
//  POST action=submitReturn         → ReturnsBackend:     submitReturn()
//  POST action=submitProduction     → ProductionBackend:  submitProduction()
//  POST action=submitWriteOff       → Inventory.gs:       _logWriteOffMovements()
//  POST action=submitProductionRequest → ProductionRequestBackend: submitProductionRequest()
//  POST action=submitProductionRun  → ProductionRunBackend: submitProductionRun()
//  POST action=submitQCDecision     → QualityBackend:      submitQCDecision()
//  POST action=submitInventoryPull  → InventoryPullBackend: submitInventoryPull()
// ============================================================

// ────────────────────────────────────────────────────────────
//    doGet — Routes all GET requests
// ────────────────────────────────────────────────────────────
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  // Delivery app API: fetch open orders for GitHub Pages frontend
  if (action === "getOrders") {
    const ss = SpreadsheetApp.openById(
      PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
    );
    const config = _getConfig(ss);
    const openOrders = _getOpenOrders(ss);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, flavors: config.flavors, openOrders }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Returns app API: fetch flavors + distributor list for GitHub Pages frontend
  if (action === "getReturnContext") {
    const ss = SpreadsheetApp.openById(
      PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
    );
    const config = _getConfig(ss);
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        flavors: config.flavors,
        distributors: config.distributors.map(d => d.name)
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Production/Write-off app API: fetch flavor list for GitHub Pages frontend
  // No distributors — Production In and Write-off aren't distributor-scoped
  if (action === "getProductionContext") {
    const ss = SpreadsheetApp.openById(
      PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
    );
    const config = _getConfig(ss);
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        flavors: config.flavors
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Order app API: fetch distributor slug+name list for GitHub Pages frontend
  // {slug, name} only — never expose phone/apiKey to the public frontend
  if (action === "getDistributors") {
    const ss = SpreadsheetApp.openById(
      PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
    );
    const config = _getConfig(ss);
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        distributors: config.distributors.map(d => ({ slug: d.slug, name: d.name }))
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Dashboard app API: fetch stock, sales, and pending-orders data for
  // GitHub Pages frontend — read-only, no recompute (Business Manager Dashboard)
  if (action === "getDashboardData") {
    const ss = SpreadsheetApp.openById(
      PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
    );
    const dashboardData = _getDashboardData(ss);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, ...dashboardData }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Production Engineer's queue: requests still needing a run logged
  if (action === "getPendingProductionRequests") {
    const data = _getPendingProductionRequests();
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, ...data }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Quality Engineer's queue: runs awaiting a QC decision
  if (action === "getPendingQCRuns") {
    const data = _getPendingQCRuns();
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, ...data }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Inventory team's queue: Approved runs not yet pulled into the warehouse
  if (action === "getPendingInventoryPulls") {
    const data = _getPendingInventoryPulls();
    return ContentService
      .createTextOutput(JSON.stringify({ success: true, ...data }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Fallback: serve the desktop delivery app HTML
  const ss         = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
  );
  const config      = _getConfig(ss);
  const openOrders  = _getOpenOrders(ss);
  const template    = HtmlService.createTemplateFromFile("DeliveryApp");
  template.flavors    = config.flavors;
  template.openOrders = openOrders;
  return template.evaluate()
    .setTitle("تسجيل التسليم | Badr El-Din")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0");
}

// ────────────────────────────────────────────────────────────
//    doPost — Routes all POST requests by action field
// ────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action || "submitDelivery";

    // Route: order submission from distributor order web app
    if (action === "submitOrder") {
      const result = submitOrder(payload);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, ...result }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Route: delivery submission from internal delivery web app
    if (action === "submitDelivery") {
      const result = submitDelivery(payload);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, ...result }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Route: return submission from internal delivery web app
    if (action === "submitReturn") {
      const result = submitReturn(payload);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, ...result }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Route: production batch submission from internal delivery web app
    if (action === "submitProduction") {
      const result = submitProduction(payload);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, ...result }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Route: write-off submission from internal delivery web app
    // No Backend file/sheet for write-off — _logWriteOffMovements (Inventory.gs)
    // IS the entire transaction, called directly. Not wrapped in an inner
    // try/catch here: failures must propagate to this function's outer catch,
    // not be silently absorbed via _logSyncError (there's no primary record to protect).
    if (action === "submitWriteOff") {
      const result = _logWriteOffMovements(payload.reason, payload.productionDate, payload.qtys);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, ...result }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Route: production request submission (Production Lifecycle)
    if (action === "submitProductionRequest") {
      const result = submitProductionRequest(payload);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, ...result }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Route: production run submission (Production Lifecycle)
    if (action === "submitProductionRun") {
      const result = submitProductionRun(payload);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, ...result }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Route: QC decision submission (Production Lifecycle)
    if (action === "submitQCDecision") {
      const result = submitQCDecision(payload);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, ...result }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Route: inventory pull confirmation (Production Lifecycle)
    if (action === "submitInventoryPull") {
      const result = submitInventoryPull(payload);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, ...result }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Unknown action — return a clear error
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: "Unknown action: " + action }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
