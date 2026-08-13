// ============================================================
//  BADR EL-DIN FARMS — ORDER BACKEND
//  Handles all order submission logic from the distributor web app
//  ─────────────────────────────────────────────────────────
//  PROJECT FILE STRUCTURE:
//  • Code.gs            → Sheet/form setup, onFormSubmit, WhatsApp helper
//  • Router.gs          → Single doGet/doPost, routes to backends
//  • DeliveryBackend.gs → Delivery logic only
//  • OrderBackend.gs    → This file — order submission logic only
//  • DeliveryApp.html   → Delivery app UI (Apps Script hosted, desktop fallback)
//  • order.html         → Order app UI (GitHub Pages hosted, distributors)
//  ─────────────────────────────────────────────────────────
//  NOTE: doGet and doPost live in Router.gs
//  This file is called by Router.gs — do not add doGet/doPost here
//  ─────────────────────────────────────────────────────────
//  SCHEMA (Phase 2 — narrow tables):
//  • 📋 Orders       → header row per order: Order_ID, Timestamp,
//                      Distributor, Total_Ordered, Status, Notes
//  • 📋 Order_Lines  → one row per non-zero flavor per order:
//                      Line_ID (OL-XXX), Order_ID (FK), Flavor_ID (FK),
//                      Flavor_Name, Qty_Ordered
//  • Order_ID generated via _nextSequentialId("ORD_COUNTER", "ORD-")
//  • Line_ID generated via _nextSequentialId("OL_COUNTER", "OL-")
//  ─────────────────────────────────────────────────────────
//  PAYLOAD EXPECTED FROM order.html:
//  {
//    action:      "submitOrder",
//    distributor: "Family Alex",
//    flavors: {
//      "Mango": 5,
//      "Blueberry": 0,
//      "Strawberry": 3,
//      ... (all 13 flavors, zeros included)
//    },
//    notes: "Please deliver before Thursday"
//  }
// ============================================================

// ────────────────────────────────────────────────────────────
//    SUBMIT ORDER — Called by Router doPost
//    Writes a header row to 📋 Orders and line rows to 📋 Order_Lines
// ────────────────────────────────────────────────────────────
function submitOrder(payload) {
  try {
    const ssId         = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
    const ss           = SpreadsheetApp.openById(ssId);
    const ordSheet     = _getSheet(ss, "📋 Orders");
    const lineSheet    = _getSheet(ss, "📋 Order_Lines");
    const flavorsSheet = _getSheet(ss, "🌿 Flavors");
    const config       = _getConfig(ss);

    // Flavor_Name → Flavor_ID map
    const flavorIdMap = {};
    flavorsSheet.getDataRange().getValues().slice(1).forEach(r => { flavorIdMap[r[2]] = r[0]; });

    // Flavor quantities in config.flavors order
    const flavorQtys   = config.flavors.map(f => Number(payload.flavors[f]) || 0);
    const totalOrdered = flavorQtys.reduce((sum, q) => sum + q, 0);

    // Generate Order ID
    const orderId   = _nextSequentialId("ORD_COUNTER", "ORD-");
    const timestamp = new Date();

    // Write header row: Order_ID | Timestamp | Distributor | Total_Ordered | Status | Notes
    const nextRow = ordSheet.getLastRow() + 1;
    const rowData = [orderId, timestamp, payload.distributor, totalOrdered, "Pending", payload.notes || ""];
    ordSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    ordSheet.getRange(nextRow, 2).setNumberFormat("yyyy-mm-dd hh:mm");

    // Re-apply status dropdown to the new row only (bulk range 2–501 already pre-applied by the sheet builder)
    ordSheet.getRange(nextRow, 5).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(["Pending", "Partial", "Fulfilled", "Cancelled"], true)
        .build()
    );

    // Write line rows: Line_ID | Order_ID | Flavor_ID | Flavor_Name | Qty_Ordered
    const lineRows = [];
    config.flavors.forEach((flavor, i) => {
      if (flavorQtys[i] > 0) {
        lineRows.push([_nextSequentialId("OL_COUNTER", "OL-"), orderId, flavorIdMap[flavor], flavor, flavorQtys[i]]);
      }
    });
    if (lineRows.length > 0) {
      const lineNextRow = lineSheet.getLastRow() + 1;
      lineSheet.getRange(lineNextRow, 1, lineRows.length, lineRows[0].length).setValues(lineRows);
    }

    // Build WhatsApp summary (non-zero flavors only)
    let orderSummary = "";
    config.flavors.forEach((flavor, i) => {
      if (flavorQtys[i] > 0) orderSummary += `  • ${flavor}: ${flavorQtys[i]} cases\n`;
    });

    // Notify internal team
    const teamMsg =
      `🛒 *New Order Received*\n` +
      `Order ID: *${orderId}*\n` +
      `From: *${payload.distributor}*\n` +
      `Date: ${timestamp.toLocaleDateString("en-GB")}\n\n` +
      `*Order Details:*\n${orderSummary}\n` +
      `Total: *${totalOrdered} cases*` +
      (payload.notes ? `\nNotes: ${payload.notes}` : "");
    CONFIG.TEAM_WHATSAPP.forEach(m => _sendWhatsApp(m.phone, m.apiKey, teamMsg));

    // Confirm to distributor
    const distConfig = config.distributors.find(d => d.name === payload.distributor);
    if (distConfig) {
      const distMsg =
        `✅ *Order Confirmed · تم استلام طلبك*\n\n` +
        `Hello ${payload.distributor},\n\n` +
        `Your order has been received.\n` +
        `Order ID: *${orderId}*\n` +
        `Date: ${timestamp.toLocaleDateString("en-GB")}\n\n` +
        `*Your Order:*\n${orderSummary}\n` +
        `Total: *${totalOrdered} cases*\n\n` +
        `You will be notified when your order is dispatched. Thank you! 🌿`;
      _sendWhatsApp(distConfig.phone, distConfig.apiKey, distMsg);
    }

    Logger.log(`✅ Order ${orderId} written via web app for ${payload.distributor}`);
    return { orderId, total: totalOrdered };

  } catch (err) {
    Logger.log("❌ submitOrder error: " + err.toString());
    throw new Error(err.toString());
  }
}