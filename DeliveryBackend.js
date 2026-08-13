// ============================================================
//  BADR EL-DIN FARMS — DELIVERY BACKEND
//  Handles all delivery-related logic for the internal web app
//  ─────────────────────────────────────────────────────────
//  PROJECT FILE STRUCTURE:
//  • Code.gs            → Sheet/form setup, onFormSubmit, WhatsApp helper
//  • Router.gs          → Single doGet/doPost, routes to backends
//  • DeliveryBackend.gs → This file — delivery logic only
//  • OrderBackend.gs    → Order submission logic only
//  • DeliveryApp.html   → Delivery app UI (Apps Script hosted, desktop fallback)
//  • order.html         → Order app UI (GitHub Pages hosted, distributors)
//  ─────────────────────────────────────────────────────────
//  NOTE: doGet and doPost have been moved to Router.gs
//  This file is called by Router.gs — do not add doGet/doPost here
//  ─────────────────────────────────────────────────────────
//  SCHEMA (Phase 2 — narrow tables):
//  • 🚚 Deliveries      → header row per delivery: Delivery_ID, Order_ID,
//                         Date_Delivered, Distributor, Total_Delivered, Notes
//  • 🚚 Delivery_Lines  → one row per non-zero flavor per delivery:
//                         Line_ID (DL-XXX), Delivery_ID (FK), Order_ID,
//                         Flavor_ID (FK), Flavor_Name, Qty_Delivered
//  • Delivery_ID generated via _nextSequentialId("DEL_COUNTER", "DEL-")
//  • Line_ID generated via _nextSequentialId("DL_COUNTER", "DL-")
// ============================================================

// ────────────────────────────────────────────────────────────
//    GET OPEN ORDERS — Called by Router doGet and on page reset
//    Returns only Pending and Partial orders
// ────────────────────────────────────────────────────────────
function getOpenOrders() {
  const ss = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
  );
  return _getOpenOrders(ss);
}

function _getOpenOrders(ss) {
  const ordSheet   = _getSheet(ss, "📋 Orders");
  const lineSheet  = _getSheet(ss, "📋 Order_Lines");
  const ordData    = ordSheet.getDataRange().getValues();
  const lineData   = lineSheet.getDataRange().getValues().slice(1);
  const openOrders = [];
  const config     = _getConfig(ss);

  // Order_ID → { Flavor_Name → Qty_Ordered }, built once from 📋 Order_Lines
  const orderLinesMap = {};
  lineData.forEach(r => {
    const [, lineOrderId, , flavorName, qty] = r;
    if (!orderLinesMap[lineOrderId]) orderLinesMap[lineOrderId] = {};
    orderLinesMap[lineOrderId][flavorName] = (orderLinesMap[lineOrderId][flavorName] || 0) + (Number(qty) || 0);
  });

  // Single header row — data starts at row 2 (array index 1)
  for (let i = 1; i < ordData.length; i++) {
    const row     = ordData[i];
    const orderId = row[0];
    const status  = row[4]; // Status column

    if (!orderId) continue;
    if (status !== "Pending" && status !== "Partial") continue;

    // Parse date
    const rawDate = row[1];
    const date    = rawDate instanceof Date
      ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : String(rawDate).substring(0, 10);

    // Flavor quantities from 📋 Order_Lines, in config.flavors order
    const lines = orderLinesMap[orderId] || {};
    const qtys  = config.flavors.map(f => lines[f] || 0);
    const total = qtys.reduce((s, q) => s + q, 0);

    openOrders.push({
      orderId:     String(orderId),
      distributor: String(row[2]),
      date:        date,
      qtys:        qtys,
      total:       total
    });
  }

  return openOrders;
}

// ────────────────────────────────────────────────────────────
//    SUBMIT DELIVERY — Called by Router doPost
//    Logs a delivery row and updates the order status
// ────────────────────────────────────────────────────────────
function submitDelivery(payload) {
  try {
    const ss           = SpreadsheetApp.openById(
      PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
    );
    const delSheet     = _getSheet(ss, "🚚 Deliveries");
    const delLineSheet = _getSheet(ss, "🚚 Delivery_Lines");
    const flavorsSheet = _getSheet(ss, "🌿 Flavors");
    const config       = _getConfig(ss);

    // Flavor_Name → Flavor_ID map
    const flavorIdMap = {};
    flavorsSheet.getDataRange().getValues().slice(1).forEach(r => { flavorIdMap[r[2]] = r[0]; });

    // Generate Delivery ID
    const deliveryId = _nextSequentialId("DEL_COUNTER", "DEL-");
    const now        = new Date();
    const total      = payload.qtys.reduce((s, q) => s + (Number(q) || 0), 0);

    // Write header row: Delivery_ID | Order_ID | Date_Delivered | Distributor | Total_Delivered | Notes
    const nextRow = delSheet.getLastRow() + 1;
    const rowData = [deliveryId, payload.orderId, now, payload.distributor, total, payload.notes || ""];
    delSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    delSheet.getRange(nextRow, 3).setNumberFormat("yyyy-mm-dd");

    // Write line rows: Line_ID | Delivery_ID | Order_ID | Flavor_ID | Flavor_Name | Qty_Delivered
    const lineRows = [];
    config.flavors.forEach((flavor, i) => {
      const qty = Number(payload.qtys[i]) || 0;
      if (qty > 0) {
        lineRows.push([_nextSequentialId("DL_COUNTER", "DL-"), deliveryId, payload.orderId, flavorIdMap[flavor], flavor, qty]);
      }
    });
    if (lineRows.length > 0) {
      const lineNextRow = delLineSheet.getLastRow() + 1;
      delLineSheet.getRange(lineNextRow, 1, lineRows.length, lineRows[0].length).setValues(lineRows);
    }

    // Update order status
    _updateOrderStatus(ss, payload.orderId, flavorIdMap);

    // Send WhatsApp notifications
    _notifyDelivery(ss, payload.orderId, payload.distributor, payload.qtys, deliveryId, total);
    // Log Delivery Out movements to Inventory spreadsheet (non-blocking)
    try {
      _logDeliveryOutMovements(payload.orderId, payload.distributor, payload.qtys);
    } catch (err) {
      _logSyncError("submitDelivery → _logDeliveryOutMovements", err, payload.orderId, payload.distributor, JSON.stringify(payload.qtys));
    }

    Logger.log(`✅ Delivery ${deliveryId} logged for order ${payload.orderId}`);
    return {
      deliveryId: deliveryId,
      orderId:    payload.orderId,
      total:      total
    };

  } catch (err) {
    Logger.log("❌ submitDelivery error: " + err.toString());
    throw new Error(err.toString());
  }
}

// ────────────────────────────────────────────────────────────
//    UPDATE ORDER STATUS AFTER DELIVERY
//    Compares total delivered vs total ordered per flavor
//    Sets status to Fulfilled if all qtys delivered, else Partial
// ────────────────────────────────────────────────────────────
function _updateOrderStatus(ss, orderId, flavorIdMap) {
  const ordSheet     = _getSheet(ss, "📋 Orders");
  const lineSheet    = _getSheet(ss, "📋 Order_Lines");
  const delLineSheet = _getSheet(ss, "🚚 Delivery_Lines");
  const ordData      = ordSheet.getDataRange().getValues();
  const lineData     = lineSheet.getDataRange().getValues().slice(1);
  const delLineData  = delLineSheet.getDataRange().getValues().slice(1);
  const config       = _getConfig(ss);

  // Find order row (single header row — data starts at array index 1)
  let orderRowIndex = -1;
  for (let i = 1; i < ordData.length; i++) {
    if (String(ordData[i][0]) === String(orderId)) {
      orderRowIndex = i;
      break;
    }
  }
  if (orderRowIndex === -1) return;

  // Ordered qtys from 📋 Order_Lines, keyed by Flavor_Name then mapped to config.flavors order
  const orderedByName = {};
  lineData.forEach(r => {
    const [, lineOrderId, , flavorName, qty] = r;
    if (String(lineOrderId) === String(orderId)) {
      orderedByName[flavorName] = (orderedByName[flavorName] || 0) + (Number(qty) || 0);
    }
  });
  const orderedQtys = config.flavors.map(f => orderedByName[f] || 0);

  // Delivered qtys summed by Flavor_ID from 🚚 Delivery_Lines (across all deliveries for this order),
  // then mapped back to config.flavors order via flavorIdMap
  const deliveredById = {};
  delLineData.forEach(r => {
    const [, , lineOrderId, flavorId, , qty] = r;
    if (String(lineOrderId) === String(orderId)) {
      deliveredById[flavorId] = (deliveredById[flavorId] || 0) + (Number(qty) || 0);
    }
  });
  const totalDelivered = config.flavors.map(f => deliveredById[flavorIdMap[f]] || 0);

  // Determine new status
  const fullyDelivered = orderedQtys.every((qty, i) => totalDelivered[i] >= qty);
  const anyDelivered   = totalDelivered.some(q => q > 0);
  const newStatus      = fullyDelivered ? "Fulfilled" : (anyDelivered ? "Partial" : "Pending");

  // Write status to 📋 Orders (col 5, fixed — Order_ID|Timestamp|Distributor|Total_Ordered|Status|Notes)
  ordSheet.getRange(orderRowIndex + 1, 5).setValue(newStatus);
  Logger.log(`✅ Order ${orderId} status updated to: ${newStatus}`);
}

// ────────────────────────────────────────────────────────────
//    NOTIFY ON DELIVERY
//    Sends WhatsApp to internal team and distributor
// ────────────────────────────────────────────────────────────
function _notifyDelivery(ss, orderId, distributorName, qtys, deliveryId, total) {
  const config = _getConfig(ss);
  let summary = "";
  config.flavors.forEach((f, i) => {
    if (qtys[i] > 0) summary += `  • ${f}: ${qtys[i]} cases\n`;
  });

  // Notify team
  const teamMsg =
    `🚚 *Delivery Logged*\n` +
    `Delivery ID: *${deliveryId}*\n` +
    `Order ID: *${orderId}*\n` +
    `Distributor: *${distributorName}*\n\n` +
    `*Delivered:*\n${summary}\n` +
    `Total: *${total} cases*`;
  CONFIG.TEAM_WHATSAPP.forEach(m => _sendWhatsApp(m.phone, m.apiKey, teamMsg));

  // Notify distributor
  const distConfig = config.distributors.find(d => d.name === distributorName);
  if (distConfig) {
    const distMsg =
      `🚚 *تم شحن طلبك · Your Order Has Been Shipped*\n\n` +
      `رقم الطلب / Order ID: *${orderId}*\n` +
      `رقم التسليم / Delivery ID: *${deliveryId}*\n\n` +
      `*الكميات المشحونة / Shipped:*\n${summary}\n` +
      `الإجمالي / Total: *${total} cases*\n\n` +
      `الرجاء تأكيد الاستلام. شكراً! 🌿`;
    _sendWhatsApp(distConfig.phone, distConfig.apiKey, distMsg);
  }
}