# Badr El-Din Farms — Supply Chain System Documentation

**Last updated:** July 21, 2026
**Maintained by:** Digital Transformation Manager
**System owner context:** Badr El-Din Farms for Agri-Industries, Egypt — Dots basil seed drink line (13 flavors)

---

## 1. System Overview

A near-zero-cost ERP-style system built on Google Workspace, replacing manual WhatsApp/Excel order tracking. Covers:

- Distributor order intake
- Delivery logging
- Inventory tracking
- KPI reporting
- Automated WhatsApp notifications

Two spreadsheets, one Apps Script project, two GitHub Pages frontends.

---

## 2. Architecture Map

```
┌─────────────────────┐         ┌──────────────────────┐
│   order.html         │         │   index.html          │
│   (GitHub Pages)      │         │   (GitHub Pages)       │
│   Distributor-facing  │         │   Internal delivery team│
└──────────┬───────────┘         └──────────┬────────────┘
           │  fetch (POST/GET)               │  fetch (POST/GET)
           ▼                                  ▼
┌─────────────────────────────────────────────────────────┐
│              Apps Script Project (single project)         │
│  Router.gs  → doGet / doPost — single entry point          │
│  Code.gs    → setup, onFormSubmit, WhatsApp, Sync Errors   │
│  OrderBackend.gs    → submitOrder()                        │
│  DeliveryBackend.gs → getOpenOrders(), submitDelivery()     │
│  Inventory.gs       → Inventory spreadsheet + movements     │
│  ReturnsBackend.gs  → submitReturn()                        │
│  ProductionBackend.gs → submitProduction()                   │
└──────────┬───────────────────────────┬────────────────────┘
           │                            │
           ▼                            ▼
┌─────────────────────────┐   ┌─────────────────────────────┐
│ Badr El-Din Order Tracker│   │ Badr El-Din Inventory Tracker│
│ (main spreadsheet)        │   │ (separate spreadsheet,       │
│ Orders, Deliveries,       │   │  separate access control)    │
│ KPI Dashboard, Summary,   │   │ Lists, Inventory Movements,  │
│ Sync Errors, Instructions │   │ Inventory Balance             │
└───────────────────────────┘   └───────────────────────────────┘
```

**Why two spreadsheets:** Inventory access needs to be controlled independently from Orders/Deliveries/KPI access — different people may need to see stock levels vs. order/fulfillment data. Apps Script can read/write across both from the same project; IMPORTRANGE is used by the 🌿 Per-Flavor Breakdown sheet to pull Current Stock from the Inventory Tracker. Authorization is one-time per spreadsheet pair and persists.

---

## 3. Apps Script Files

| File | Responsibility |
|---|---|
| `Code.gs` | `CONFIG` object, sheet/form setup (`setupEntireSystem`), `onFormSubmit` trigger, `_sendWhatsApp` helper, `_logSyncError`, Sync Errors sheet builder, Config sheet builders (`_buildConfigSheets`), shared helpers (`_getSheet`, `_getConfig`, `columnLetter`, `_removeDefaultSheet`) |
| `Router.gs` | Single `doGet`/`doPost` entry point. Routes by `action` parameter to the correct backend function. No business logic lives here. |
| `OrderBackend.gs` | `submitOrder(payload)` — writes a new row to Orders, mirrors `onFormSubmit` logic, sends WhatsApp confirmations. |
| `DeliveryBackend.gs` | `getOpenOrders()` / `_getOpenOrders(ss)`, `submitDelivery(payload)`, `_updateOrderStatus`, `_notifyDelivery`. Calls `_logDeliveryOutMovements` (Inventory.gs) after logging a delivery. |
| `Inventory.gs` | Inventory spreadsheet creation/access (`_getInventorySpreadsheet`), `_buildListsSheet`, `_buildInventoryMovementsSheet`, `_buildInventoryBalanceSheet`, `setupInventorySystem`, `_logDeliveryOutMovements`, `_logReturnInMovements`, `_logProductionInMovements`, `_logWriteOffMovements`. |
| `ReturnsBackend.gs` | `submitReturn(payload)` — writes a header row to `↩️ Returns` and line rows to `↩️ Return_Lines` (Phase 2 narrow schema), generates sequential `RET-XXX`/`RL-XXX` IDs via `_nextSequentialId` (Code.gs), sends a team-only WhatsApp notification (`_notifyReturn`), then calls `_logReturnInMovements` (Inventory.gs) non-blocking. |
| `ProductionBackend.gs` | `submitProduction(payload)` — writes a header row to `🌱 Productions` and line rows to `🌱 Production_Lines` (Phase 2 narrow schema), generates sequential `PRD-XXX`/`PL-XXX` IDs via `_nextSequentialId` (Code.gs), sends a team-only WhatsApp notification (`_notifyProduction`), then calls `_logProductionInMovements` (Inventory.gs) non-blocking. |

**Shared global scope:** All `.gs` files in one Apps Script project share scope — no imports needed. `CONFIG`, `_getSheet`, `_sendWhatsApp`, `_logSyncError`, `_removeDefaultSheet` are all callable from any file.

---

## 4. Frontends

| File | Hosted | Audience | Calls |
|---|---|---|---|
| `order.html` | GitHub Pages (`https://monagib3.github.io/delivery-app/order.html`) | Distributors (via `?dist=` param) | `doPost action=submitOrder` |
| `index.html` (mirrors `DeliveryApp.html`) | GitHub Pages (`https://monagib3.github.io/delivery-app/`) | Internal delivery team | `doGet action=getOrders`, `doPost action=submitDelivery` |
| `DeliveryApp.html` | Apps Script-hosted (`doGet` fallback) | Desktop fallback only | Same as index.html, via `google.script.run` instead of fetch |
| `inventory-ops.html` | GitHub Pages (`https://monagib3.github.io/delivery-app/inventory-ops.html`) | Internal team only (toggle: Production In / Write-off) | `doGet action=getProductionContext`, `doPost action=submitProduction`, `doPost action=submitWriteOff` |
| `home.html` | GitHub Pages (`https://monagib3.github.io/delivery-app/home.html`) | Internal ops team (launcher/home page) | None — static links only to `order.html`, `index.html`, `returns.html`, `inventory-ops.html`; no backend calls |

**Distributor link format:** `order.html?dist=family-alex` / `order.html?dist=cairo-office` — mapped to full names via a `?action=getDistributors` fetch to `Router.gs` on page load (reads the `🏢 Distributors` sheet via `_getConfig()`), not a hardcoded map in the frontend.

**Why GitHub Pages for mobile apps:** Direct Apps Script URLs fail to load reliably on iOS mobile browsers. Hosting static HTML on GitHub Pages and calling Apps Script as a REST API via `fetch` resolves this.

---

## 5. Main Spreadsheet — "Badr El-Din Order Tracker"

| Sheet | Type | Purpose |
|---|---|---|
| `📋 Orders` | Static, trigger-written | Header row per order. Columns: Order_ID, Timestamp, Distributor, Total_Ordered, Status, Notes. Single header row — flavor-level detail lives in `📋 Order_Lines`. |
| `📋 Order_Lines` | Static, trigger-written | One row per non-zero flavor per order. Columns: Line_ID (OL-XXX), Order_ID (FK), Flavor_ID (FK), Flavor_Name, Qty_Ordered. |
| `🚚 Deliveries` | Static, trigger-written | Header row per delivery. Columns: Delivery_ID, Order_ID, Date_Delivered, Distributor, Total_Delivered, Notes. Single header row — flavor-level detail lives in `🚚 Delivery_Lines`. |
| `🚚 Delivery_Lines` | Static, trigger-written | One row per non-zero flavor per delivery. Columns: Line_ID (DL-XXX), Delivery_ID (FK), Order_ID, Flavor_ID (FK), Flavor_Name, Qty_Delivered. |
| `↩️ Returns` | Static, trigger-written | Header row per return. Columns: Return_ID, Date, Distributor, Total_Returned, Notes. Single header row — flavor-level detail lives in `↩️ Return_Lines`. |
| `↩️ Return_Lines` | Static, trigger-written | One row per non-zero flavor per return. Columns: Line_ID (RL-XXX), Return_ID (FK), Flavor_ID (FK), Flavor_Name, Qty_Returned. |
| `🌱 Productions` | Static, trigger-written | Header row per production batch. Columns: Production_ID, Date, Lot_Number, Total_Produced, Notes. Single header row — flavor-level detail lives in `🌱 Production_Lines`. |
| `🌱 Production_Lines` | Static, trigger-written | One row per non-zero flavor per production batch. Columns: Line_ID (PL-XXX), Production_ID (FK), Flavor_ID (FK), Flavor_Name, Qty_Produced. |
| ` KPI Dashboard` | Formula-driven | Fill rate %, days to deliver, per-flavor delivered vs ordered. Pulls from Orders + Deliveries via row-mirrored formulas (rows 9–508). Color-coded: green ≥95%, yellow 70–94%, red <70%. |
| ` Summary` | Formula-driven | Per-distributor fill rate, avg delivery time, pending orders. Pulls from KPI Dashboard. |
| ` Sync Errors` | Static, append-only | Cross-system failure log (currently: Inventory Movements write failures during `submitDelivery`). Columns: Timestamp, Function, Error, Order ID, Distributor, Details. Triggers a WhatsApp alert to the team on each entry. |
| ` Instructions` | Static reference | Human-readable usage guide for the team. |
| `Form Responses 1` | Auto-created by Google Forms | Passive fallback intake channel. Hidden/ignored — not yet hidden in UI (pending task). |
| `⚙️ Products` | Static, seeded reference | Product master. Columns: Product_ID (PK), Product_Code, Product_Name, Pack_Size, Is_Active. Currently one row: `PROD-001 / DOTS / Dots / 12 / TRUE`. |
| `🌿 Flavors` | Static, seeded reference | Flavor master, one row per flavor (13 rows), linked to Products via Product_ID FK. Columns: Flavor_ID (PK), Product_ID (FK), Flavor_Name, Is_Active. Read at runtime by `_getConfig()` — the live source of truth for flavor names/order. |
| `🏢 Distributors` | Static, seeded reference | Distributor master. Columns: Distributor_ID (PK), Name, Slug, Phone, CallMeBot_ApiKey, Is_Active. Read at runtime by `_getConfig()` — replaced the old hardcoded `CONFIG.DISTRIBUTORS`/`DISTRIBUTOR_MAP` as the live source for distributor names/slugs. |
| `🔗 Distributor_Products` | Static, seeded reference | Many-to-many join between Distributors and Products (every distributor currently sells the one Dots product). Columns: DP_ID (PK), Distributor_ID (FK), Product_ID (FK). Not yet read by any code. |
| `📱 Team_Contacts` | Static, seeded reference | Internal team WhatsApp contacts, seeded from `CONFIG.TEAM_WHATSAPP`. Columns: Contact_ID (PK), Name, Phone, CallMeBot_ApiKey, Is_Active. Not yet read by any code — `CONFIG.TEAM_WHATSAPP` in Code.gs remains the live source for this phase. |
| `📉 Reorder_Points` | Static, seeded reference | Per-flavor low-stock threshold, one row per flavor. Columns: Flavor_ID (FK), Min_Stock_Cases. Not yet read by any code — Inventory Balance's low-stock formatting still uses the hardcoded thresholds in `_buildInventoryBalanceSheet` (see §6). |

`_ARCHIVE_Orders`, `_ARCHIVE_Deliveries`, `_ARCHIVE_Returns`, `_ARCHIVE_Productions` — grey tabs, frozen, read-only wide-format historical data preserved from pre-migration. Not read by any live code path.

---

## 6. Inventory Spreadsheet — "Badr El-Din Inventory Tracker"

Separate file, created/accessed via `_getInventorySpreadsheet()`, ID stored in Script Properties as `INVENTORY_SPREADSHEET_ID`.

| Sheet | Type | Purpose |
|---|---|---|
| ` Lists` | Static reference | Dropdown source data. Column A: Warehouses (currently just "Main Warehouse"). Column B: Sources/Destinations (Packing Line, Samples, Write-off, + one row per `CONFIG.DISTRIBUTORS`). Designed so future warehouse/client growth = adding a row here, not a schema change. |
| ` Inventory Movements` | Static, append-only ledger | One row per stock change. Columns: Date, Type, Reference, Production Date, Product, Flavor, Quantity (cases, signed), Warehouse, Source/Destination, Notes. `Type` dropdown: Opening Balance, Production In, Delivery Out, Return, Write-off. Quantity is **signed**: positive = stock in, negative = stock out. |
| ` Inventory Balance` | Formula-driven | One row per flavor (13 rows). Columns: Flavor, Opening Balance, Production In, Delivered Out, Returns, Write-offs, Current Stock. Each column is a `SUMIFS` against Inventory Movements by Type + Flavor. Current Stock = `SUM` of the row. Red-flagged if below that flavor's low-stock threshold. |

### Low-stock thresholds (Current Stock, cases)
| Flavor(s) | Threshold |
|---|---|
| Mango, Blueberry, Lemon-Mint, Strawberry | 75 |
| Coconut, Orange | 25 |
| All other flavors | 50 |

*(Flat per-group thresholds today. A per-flavor `Reorder Point` column is a planned future addition once real consumption data exists — see Section 9.)*

### Product column
Defaults to `"Dots"` everywhere today (single product line). Included from the start so a second product line can be added later without restructuring Inventory Movements/Balance — Balance would become one row per (Product, Flavor) pair, and `SKU Code` could be derived as `Product-Flavor` (e.g. `DOTS-MANGO`).

### Opening Balance entries
`Type = "Opening Balance"` is a valid dropdown value but is **never surfaced in any web app or form UI** — it exists for one-time manual entry only. Each flavor needs exactly one Opening Balance row entered manually (representing real physical stock at go-live) before relying on the system for accurate Current Stock.

### Production Date column
Column 4 of ` Inventory Movements`, inserted after Reference and before Product. Exists for all transaction types but is only actively populated for Production In and Write-off — `_logDeliveryOutMovements` and `_logReturnInMovements` write an empty string in this position so the row stays the correct width. Batch-splitting UI for Delivery Out / Returns is a future addition, not yet built.

---

## 7. Delivery → Inventory Sync

When a delivery is submitted via `submitDelivery(payload)` in `DeliveryBackend.gs`:

1. Writes the delivery row to ` Deliveries` (main spreadsheet)
2. Updates order status via `_updateOrderStatus`
3. Sends WhatsApp notifications via `_notifyDelivery`
4. **Calls `_logDeliveryOutMovements(orderId, distributor, qtys)`** (Inventory.gs), wrapped in try/catch:
   - For each flavor with `qty > 0`, appends one row to ` Inventory Movements`: `Type="Delivery Out"`, `Quantity = -qty`, `Warehouse="Main Warehouse"`, `Source/Destination=distributor`
   - This is **non-blocking** — a failure here does not prevent the delivery itself from being recorded
5. **On failure:** caught by `_logSyncError`, which:
   - Appends a row to ` Sync Errors` (main spreadsheet)
   - Sends a WhatsApp alert to `CONFIG.TEAM_WHATSAPP` with the function name, order ID, distributor, and error message

**Why non-blocking:** Deliveries are business-critical (distributors are waiting on confirmation) and live in a different spreadsheet than Inventory. A transient Inventory-spreadsheet access issue should never prevent a delivery from being logged — it should be surfaced for manual reconciliation instead.

---

## 8. Key Technical Conventions (apply project-wide)

- **Sheet lookup:** always `ss.getSheets().find(s => s.getName().includes("keyword"))` via the `_getSheet(ss, keyword)` helper — never `getSheetByName()` with exact names, to avoid encoding/leading-space mismatches.
- **Leading space in tab names:** sheets are created via `insertSheet(" SheetName")` (leading space) — an inherited convention from the original emoji-prefix design. `_getSheet`'s fuzzy `.includes()` match hides this when looking up sheets programmatically, but **hardcoded formula strings referencing a sheet by name must include the leading space exactly**, or `SUMIFS`/cross-sheet formulas will return `#N/A`/`#REF!`. (This caused a real bug during Inventory Balance development — see Section 10.)
- **Spreadsheet access:** always `SpreadsheetApp.openById(ssId)` with the ID read from `PropertiesService.getScriptProperties()` — never `getActiveSpreadsheet()` inside triggers (returns `null` in trigger context).
- **No pre-filled formulas in transactional sheets:** Orders, Deliveries, and Inventory Movements are written entirely as static values by trigger/API functions. Pre-filled formulas in unused rows break `getLastRow()`-based row-counting logic used to generate sequential IDs (ORD-XXX, DEL-XXX).
- **Template tags:** use `<?!= ?>` (not `<?= ?>`) in `.html` templates served via `HtmlService` — the `=` variant HTML-escapes JSON, breaking client-side `JSON.parse`.
- **Data validation:** `requireNumber()` does **not** exist on `DataValidationBuilder` in Apps Script — do not use it. No direct generic numeric validator is available; numeric correctness for manually entered fields (e.g. Inventory Movements Quantity) is currently unenforced by validation and relies on convention/code review.
- **`setupEntireSystem()`:** rebuilds all main-spreadsheet sheets from scratch. **Never re-run without explicit instruction** — it will wipe formatting/structure (though not necessarily data, depending on the sheet) of every sheet it touches. New incremental sheet-builders use temporary standalone wrapper functions (created, run once, then deleted) instead of being added to `setupEntireSystem()`, unless explicitly approved.
- **Deployment:** code changes to files used by the live web app (`Router.gs`, `OrderBackend.gs`, `DeliveryBackend.gs`, `Inventory.gs` when called from these) require **Deploy → Manage deployments → edit existing deployment → Version: "New version" → Deploy**. Selecting an existing numbered version (not "New version") in that dropdown will silently leave the deployment on old code with no error — confirm the Version field shows a fresh timestamp after deploying. Editor-run functions (e.g. via the function dropdown) always use the latest saved code regardless of deployment state — this discrepancy is a common source of "it works when I test it manually but not through the real app" bugs.
- **Inventory Balance SUMIFS column constants:** `_buildInventoryBalanceSheet`'s `flavorCol`/`typeCol`/`qtyCol` constants are hardcoded column letters against ` Inventory Movements`, not derived from header position — currently `flavorCol` = col F, `typeCol` = col B, `qtyCol` = col G (shifted after the Production Date column was inserted as col 4). Any future Inventory Movements column change must update these three constants or every SUMIFS in Inventory Balance silently returns wrong totals.
- **`clasp push` invocation:** always run as `cd ~/Downloads/Claude/badr-eldin-system && clasp push`, not bare `clasp push` — a bare run failed once in this sandbox with a DNS resolution error.
- **`CONFIG.FLAVORS`/`CONFIG.DISTRIBUTORS` are no longer the live source of truth (as of DB Migration Phase 1):** `Router.gs`, `OrderBackend.gs`, `DeliveryBackend.gs`, `ReturnsBackend.gs`, `ProductionBackend.gs`, and the four `_log*Movements` functions in `Inventory.gs` all call `_getConfig(ss)` — which reads the `🌿 Flavors`/`🏢 Distributors` sheets at runtime — instead of referencing `CONFIG.FLAVORS`/`CONFIG.DISTRIBUTORS` directly. Those `CONFIG` values still exist in `Code.gs` and are still used, but only by setup-time sheet builders and the legacy Google Form flow (`onFormSubmit`), which never execute during a live request. New live-path code should call `_getConfig()`, not `CONFIG.FLAVORS`/`CONFIG.DISTRIBUTORS`.

---

## 9. Pending / Deferred Work

| Item | Status |
|---|---|
| KPI Dashboard charts and filters | Not started |
| `onEdit` trigger → distributor WhatsApp notification on status change | Placeholder `onStatusChange()` exists in `Code.gs`, not wired to a trigger |
| Per-distributor pre-filled Google Form links | Not started |
| Home screen icon for GitHub Pages delivery app | Not started |
| Hide "Form Responses 1" sheet | Not started |
| **Returns handling** | Complete — all 5 sub-tasks done (Returns sheet builder, `ReturnsBackend.gs` + `_logReturnInMovements`, Router wiring for `getReturnContext`/`submitReturn`, `returns.html` frontend, end-to-end test). See `CLAUDE.md` §7. |
| **Production In and Write-off transactions** | Complete — all 7 sub-tasks done (` Productions` sheet builder, `ProductionBackend.gs` with `submitProduction`, `_logProductionInMovements`/`_logWriteOffMovements` in `Inventory.gs`, Router wiring for `getProductionContext`/`submitProduction`/`submitWriteOff`, the `inventory-ops.html` frontend, and the end-to-end test). See `CLAUDE.md` §7B. |
| **Decimal quantity support across frontends** | Complete — `index.html`, `returns.html`, and `inventory-ops.html` all use `step="any"` on qty stepper inputs and `parseFloat` for qty value parsing. `DeliveryApp.html` (Apps Script-hosted mirror of `index.html`) updated to match. |
| **`home.html` mobile launcher page** | Complete — 4 tappable cards (New Order → `order.html`, Log Delivery → `index.html`, Log Return → `returns.html`, Inventory Ops → `inventory-ops.html`), RTL layout with Arabic primary labels + English subtitles, matches existing color palette (`#1a3a2a`/`#2d6a4f`/`#52b788`) and Cairo + DM Mono fonts. Live at `https://monagib3.github.io/delivery-app/home.html`. |
| **`order.html` distributor picker (no `?dist=` param)** | Complete — when `order.html` is accessed without a `?dist=` param (or with an empty one), it now shows a distributor selection screen (`#pickerScreen`, one `.distributor-card` per `DISTRIBUTOR_MAP` entry) instead of the "Invalid Link" error. Tapping a card redirects to `order.html?dist=<key>`. Frontend-only change, styled with `order.html`'s own existing gold/ink palette. A present-but-unrecognized `dist` value still shows the original "Invalid Link" error screen; a valid `dist` value still goes straight to the order form — both unchanged. See `CLAUDE.md` §7D. |
| **Fixed Home button on internal pages** | Complete — `index.html`, `returns.html`, `inventory-ops.html` each got a fixed top-left 🏠 link to `home.html` (`.home-btn`, 44×44px circle, `top:12px; left:12px; z-index:1000`, styled from each file's existing green palette). `order.html`/`home.html` untouched — out of scope. See `CLAUDE.md` §7F. |
| **Git version control setup** | Complete — local working copy (previously ungit'd) connected to the existing GitHub remote `https://github.com/monagib3/delivery-app.git`; local `main` tracks `origin/main`. Reconciled via `git merge --allow-unrelated-histories`, resolving 3 add/add conflicts in favor of local edits. `.gitignore` added for `.claude/settings.local.json` (already-tracked prior copy not yet untracked — needs `git rm --cached`). Commits pushed directly to `main`, no PR workflow. See `CLAUDE.md` §7G. |
| **Business Manager Dashboard** | Complete — new `DashboardBackend.js` (`getDashboardData()`/`_getDashboardData(ss)`) reads already-computed cells straight from `🌿 Per-Flavor Breakdown` and `📈 Per-Distributor Summary` (no recompute), and extends `_getOpenOrders(ss)` (reused from `DeliveryBackend.js`, not duplicated) with cases outstanding from `🚚 Delivery_Lines` and a timezone-safe `_daysOpen()` day count. `Router.js` gained `?action=getDashboardData` (same shape as the other GET context routes). New `dashboard.html` frontend reuses `returns.html`'s palette/header/home-btn/loading/error blocks with one new read-only `.dash-table` component; `home.html` gained a 5th launcher card. Deployed as a new Apps Script version and verified live end-to-end. See `CLAUDE.md` §7H. |
| Per-flavor `Reorder Point` column (replacing flat thresholds) | Deferred until real consumption data exists to set meaningful per-flavor values |
| Production In automation | Currently fully manual entry into Inventory Movements; no automation hook exists |
| Multi-warehouse support | Schema supports it (Warehouse column, dropdown from Lists) but only "Main Warehouse" exists today; no per-warehouse Balance breakout yet |
| Multi-product support | Schema supports it (Product column, defaults to "Dots") but no second product line exists yet |
| Production/quality system integration | Long-term, not scoped |
| Centralized executive dashboard | Long-term, not scoped |
| **DB Migration — Phase 1: Config Layer** | Complete — six Config sheets built (`⚙️ Products`, `🌿 Flavors`, `🏢 Distributors`, `🔗 Distributor_Products`, `📱 Team_Contacts`, `📉 Reorder_Points`), `_getConfig()` added to `Code.gs`, the live API path (Router.gs/OrderBackend.gs/DeliveryBackend.gs/ReturnsBackend.gs/ProductionBackend.gs + the four `_log*Movements` functions in Inventory.gs) migrated off `CONFIG.FLAVORS`/`CONFIG.DISTRIBUTORS`, `?action=getDistributors` added, hardcoded `DISTRIBUTOR_MAP` removed from `order.html`. See `CLAUDE.md` §7E. |
| **DB Migration — Phase 2: Normalize transactional tables** | In Progress — Sub-task 2.1 complete (sheets built, archives created). Sub-task 2.2 complete (historical data migrated — 2 Orders/22 lines, 2 Deliveries/21 lines, 1 Return/1 line, 2 Productions/12 lines). Sub-task 2.3 complete (`OrderBackend.js`/`DeliveryBackend.js` rewritten to narrow schema, `ORD_COUNTER`/`DEL_COUNTER` seeded, deployed as Version 16). Sub-task 2.4 complete (`ReturnsBackend.js`/`ProductionBackend.js` rewritten to narrow schema, `RET_COUNTER`/`RL_COUNTER`/`PL_COUNTER` seeded, `PRD_COUNTER` left untouched, deployed as Version 17). Complete — Sub-task 2.5 was absorbed into Sub-task 2.4. |
| **Phase 2, Sub-task 2.1 — New Sheet Structure Built** | Complete — Eight narrow sheets created (four header + four lines), four wide sheets archived as `_ARCHIVE_` prefixed grey tabs. Tab name verification performed manually before running — confirmed actual tab names were `📋 Orders`, `🚚 Deliveries`, `Returns`, `Productions` (no leading spaces on Returns/Productions despite earlier convention). Archive keywords corrected accordingly. Runner executed successfully, all 12 operations logged clean. |
| **Phase 2, Sub-task 2.2 — Historical Data Migration Complete** | `_migrateHistoricalData` ran successfully. Four `_ARCHIVE_` sheets exploded into narrow header + lines format. Counts: 2 Orders (22 lines), 2 Deliveries (21 lines), 1 Return (1 line), 2 Productions (12 lines). All eight narrow sheets verified. Both temporary functions deleted per project convention. "Family Alex" naming on historical Orders/Deliveries carried through correctly — pre-existing data issue, not a migration bug. |
| **Phase 2, Sub-task 2.3 — OrderBackend.js / DeliveryBackend.js Rewritten to Narrow Schema** | Complete — `submitOrder()` now writes a header row to `📋 Orders` and line rows to `📋 Order_Lines`; `submitDelivery()` now writes a header row to `🚚 Deliveries` and line rows to `🚚 Delivery_Lines`. `_getOpenOrders()` and `_updateOrderStatus()` rewritten to read qtys from the narrow line tables (joined via a `Flavor_Name → Flavor_ID` map built at runtime from `🌿 Flavors`) instead of wide column offsets. Order_ID/Delivery_ID generation migrated from `getLastRow()` counting to `_nextSequentialId()`, backed by two new Script Properties counters, `ORD_COUNTER`/`DEL_COUNTER`, added to the counter registry in `Code.gs`. Counters seeded to 2 (matching the 2 migrated historical Orders/Deliveries from Sub-task 2.2) via a temporary `_initOrderDelCounters()` wrapper — run once manually via the Apps Script editor, then deleted per project convention. Deployed as Version 16. |
| **Phase 2, Sub-task 2.4 — ReturnsBackend.js / ProductionBackend.js Rewritten to Narrow Schema** | Complete — `submitReturn()` now writes a header row to `↩️ Returns` (`Return_ID, Date, Distributor, Total_Returned, Notes`) and line rows to `↩️ Return_Lines` (`Line_ID, Return_ID, Flavor_ID, Flavor_Name, Qty_Returned`); `submitProduction()` now writes a header row to `🌱 Productions` (`Production_ID, Date, Lot_Number, Total_Produced, Notes`) and line rows to `🌱 Production_Lines` (`Line_ID, Production_ID, Flavor_ID, Flavor_Name, Qty_Produced`). Both use the same `Flavor_Name → Flavor_ID` runtime lookup against `🌿 Flavors` established in Sub-task 2.3. Return_ID/Return line-ID generation migrated from `getLastRow()` counting to `_nextSequentialId()`, backed by two new Script Properties counters, `RET_COUNTER`/`RL_COUNTER`. Production line-ID generation added a third new counter, `PL_COUNTER`; `PRD_COUNTER` (Production header ID) was already live and already correct at 4, so it was deliberately left untouched. New counters seeded via a temporary `_initRetPrdCounters()` wrapper (`RET_COUNTER=1`, `RL_COUNTER=1`, `PL_COUNTER=12`, matching the 1 migrated historical Return and 12 migrated historical Production lines from Sub-task 2.2) — run once manually via the Apps Script editor, then deleted per project convention. Deployed as Version 17, end-to-end tested (test rows cleaned up afterward). Note: `Code.gs`'s counter registry comment still marks `OL_COUNTER`/`DL_COUNTER` as "not yet wired to a caller" — stale as of Sub-task 2.3, not a bug, left uncorrected as out of scope for both 2.3 and 2.4. |
| DB Migration — Phase 3: Fix reporting layer | Complete — Three new sheet builders in Code.js: `_buildKPIDashboardCards` (5 summary cards + hidden X:Z cycle-time helper block, batch-written via `setFormulas()`), `_buildDistributorSummarySheet` (renames `📈 Summary` → `📈 Per-Distributor Summary`, sources distributors from `_getConfig()`, SUMPRODUCT+VLOOKUP joins for Cases Ordered/Delivered, AVERAGEIFS reuses KPI Dashboard hidden Y:Z block), `_buildFlavorBreakdownSheet` (new tab, single IMPORTRANGE staged in hidden K:Q, ISNUMBER guard on reorder flag, reads Reorder_Points via Flavor_ID in hidden col J). Old `_buildKPIDashboard`/`_buildSummarySheet` deleted. `setupEntireSystem()` dead calls removed. Deployed via temporary `_runBuildPhase3Reports()` wrapper (deleted after run). IMPORTRANGE auth note: IFERROR suppresses `#REF!` prompt — must temporarily remove IFERROR from K1 on Per-Flavor Breakdown to trigger Allow Access. |
| **Historical data cleanup — Orders sheet** | Rows submitted before the Sub-task 4 naming fix have distributor name "Family Alex" (no hyphen), not matching the canonical "Family - Alex" now used everywhere. These rows don't roll up correctly in Summary's per-distributor grouping. Needs a one-time manual correction pass on the Orders sheet — separate from further DB migration work. |
| Input validation — enforce non-negative Qty fields across all four backends | Not started |
| Partial order close-out flow — status model redesign (Partial currently treated as open workload) | Not started |
| Production efficiency metric — Cases Produced/Delivered shown as separate columns for now; define and add ratio metric once 3-6 months of history exists | Deferred |
| Unify Reorder_Points sheet with Inventory Balance hardcoded thresholds — two sources of truth for the same concept | Not started |
| BigQuery + Connected Sheets migration — replace IMPORTRANGE with BigQuery as reporting layer foundation; enables Looker Studio executive dashboard | Long-term |

---

## 10. Notable Bugs Resolved During Development

| Bug | Cause | Fix |
|---|---|---|
| `_buildInventoryMovementsSheet` threw `TypeError: ...requireNumber is not a function` | `requireNumber()` is not a real `DataValidationBuilder` method | Removed the validation block entirely; Quantity correctness relies on convention |
| Inventory Balance showed `#N/A` in every cell | `MOV` constant in `_buildInventoryBalanceSheet` referenced `'Inventory Movements'` (no leading space) while the actual tab name was `' Inventory Movements'` (leading space) | Corrected the constant to `"' Inventory Movements'"` |
| KPI Dashboard showed `#REF!` after deleting test rows from Orders | KPI Dashboard formulas reference specific Orders row numbers; deleting rows (vs. clearing values) broke those references | Re-ran `_buildKPIDashboard` standalone via a temporary wrapper to rebuild formulas cleanly |
| Inventory sync silently not working after wiring `submitDelivery()` → `_logDeliveryOutMovements` | Live web app deployment was pinned to an old version (April snapshot) — the Version dropdown in "Manage deployments" had not actually been advanced to "New version" on a prior redeploy attempt, despite the deploy dialog appearing to succeed | Explicitly selected "New version" in the dropdown before clicking Deploy; confirmed via the post-deploy confirmation screen showing a fresh timestamp |

**General lesson from the last bug:** when a feature works correctly when test-run from the Apps Script editor but doesn't work through the real (deployed) web app, check the **deployed version's timestamp** in Manage deployments before debugging the code itself — editor runs always use the latest saved code; the live web app does not, until explicitly redeployed.

---

## 11. Glossary of IDs

| ID format | Generated by | Meaning |
|---|---|---|
| `ORD-XXX` | `onFormSubmit` / `submitOrder` | Sequential order ID, auto-generated |
| `DEL-XXX` | `submitDelivery` | Sequential delivery ID, auto-generated |
| `RET-XXX` | `submitReturn` | Sequential return ID, auto-generated |
| `PRD-XXX` | `submitProduction` | Sequential production ID, generated via `_nextSequentialId` Script Properties counter |
| `WOF-XXX` | `_logWriteOffMovements` | Sequential write-off ID, generated via `_nextSequentialId` Script Properties counter |
| `INVENTORY_SPREADSHEET_ID` | `_createOrGetInventorySpreadsheet` | Script Property holding the Inventory spreadsheet's file ID |
| `SPREADSHEET_ID` | `_createOrGetSpreadsheet` | Script Property holding the main Order Tracker spreadsheet's file ID |
