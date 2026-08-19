// ============================================================
//  BADR EL-DIN FARMS — DASHBOARD BACKEND
//  Read-only Business Manager Dashboard data
//  ─────────────────────────────────────────────────────────
//  PROJECT FILE STRUCTURE:
//  • Code.gs             → Sheet/form setup, onFormSubmit, WhatsApp helper
//  • Router.gs           → Single doGet/doPost, routes to backends
//  • DashboardBackend.gs → This file — dashboard read-only logic only
//  • DeliveryBackend.gs  → Order/delivery logic — _getOpenOrders() reused
//                          here as-is, not duplicated
//  • dashboard.html      → Dashboard UI (GitHub Pages hosted, internal team)
//  ─────────────────────────────────────────────────────────
//  NOTE: doGet has been moved to Router.gs
//  This file is called by Router.gs — do not add doGet/doPost here
//  ─────────────────────────────────────────────────────────
//  DATA SOURCES (read-only — no recompute, no writes):
//  • Stock by flavor       ← 🌿 Per-Flavor Breakdown (already-computed cells)
//  • Sales by distributor  ← 📈 Per-Distributor Summary (already-computed cells)
//  • Pending orders        ← _getOpenOrders(ss) (DeliveryBackend.gs) +
//                            🚚 Delivery_Lines (cases outstanding) +
//                            today's date (days open)
// ============================================================

// ────────────────────────────────────────────────────────────
//    GET DASHBOARD DATA — Called by Router doGet
// ────────────────────────────────────────────────────────────
function getDashboardData() {
  const ss = SpreadsheetApp.openById(
    PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID")
  );
  return _getDashboardData(ss);
}

function _getDashboardData(ss) {
  return {
    stockByFlavor:      _readStockByFlavor(ss),
    salesByDistributor: _readSalesByDistributor(ss),
    pendingOrders:      _readPendingOrders(ss)
  };
}

// ────────────────────────────────────────────────────────────
//    STOCK BY FLAVOR — reads 🌿 Per-Flavor Breakdown as-is
//    Cols: A=Flavor, F=Current Stock, G=Below Reorder Threshold
// ────────────────────────────────────────────────────────────
function _readStockByFlavor(ss) {
  const ws   = _getSheet(ss, "Per-Flavor Breakdown");
  const data = ws.getDataRange().getValues();
  const rows = [];

  // Data starts row 4 (array index 3) — stop at first blank Flavor cell
  for (let i = 3; i < data.length; i++) {
    const flavor = data[i][0];
    if (!flavor) break;
    rows.push({
      flavor:         String(flavor),
      currentStock:   data[i][5],
      belowThreshold: data[i][6]
    });
  }
  return rows;
}

// ────────────────────────────────────────────────────────────
//    SALES BY DISTRIBUTOR — reads 📈 Per-Distributor Summary as-is
//    Cols: A=Distributor, C=Cases Ordered, D=Cases Delivered, E=Fill Rate %
// ────────────────────────────────────────────────────────────
function _readSalesByDistributor(ss) {
  const ws   = _getSheet(ss, "Per-Distributor Summary");
  const data = ws.getDataRange().getValues();
  const rows = [];

  // Data starts row 4 (array index 3) — stop at first blank Distributor cell
  for (let i = 3; i < data.length; i++) {
    const distributor = data[i][0];
    if (!distributor) break;
    rows.push({
      distributor:    String(distributor),
      casesOrdered:   data[i][2],
      casesDelivered: data[i][3],
      fillRate:       data[i][4]
    });
  }
  return rows;
}

// ────────────────────────────────────────────────────────────
//    PENDING ORDERS — extends _getOpenOrders(ss) (DeliveryBackend.gs)
//    with cases outstanding and days open. Sorted newest first.
// ────────────────────────────────────────────────────────────
function _readPendingOrders(ss) {
  const openOrders   = _getOpenOrders(ss);
  const delLineSheet = _getSheet(ss, "🚚 Delivery_Lines");
  const delLineData  = delLineSheet.getDataRange().getValues().slice(1);

  // Order_ID → total delivered across all flavors/deliveries
  const deliveredByOrderId = {};
  delLineData.forEach(r => {
    const [, , orderId, , , qty] = r;
    deliveredByOrderId[orderId] = (deliveredByOrderId[orderId] || 0) + (Number(qty) || 0);
  });

  const pending = openOrders.map(o => {
    const delivered = deliveredByOrderId[o.orderId] || 0;
    return {
      orderId:          o.orderId,
      distributor:      o.distributor,
      date:             o.date,
      total:            o.total,
      casesOutstanding: o.total - delivered,
      daysOpen:         _daysOpen(o.date)
    };
  });

  pending.sort((a, b) => b.date.localeCompare(a.date));
  return pending;
}

// ────────────────────────────────────────────────────────────
//    DAYS OPEN — timezone-consistent date-only diff
//    Formats "today" via Session.getScriptTimeZone() (same tz
//    _getOpenOrders uses for order.date) and compares both dates
//    as UTC-midnight representations of their calendar date, so
//    the diff is never off-by-one regardless of server/script tz.
// ────────────────────────────────────────────────────────────
function _daysOpen(dateStr) {
  const tz       = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const [ty, tm, td] = todayStr.split("-").map(Number);
  const [oy, om, od] = dateStr.split("-").map(Number);
  return Math.floor((Date.UTC(ty, tm - 1, td) - Date.UTC(oy, om - 1, od)) / 86400000);
}
