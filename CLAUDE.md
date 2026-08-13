# Badr El-Din Farms — Supply Chain System

This file is read automatically at the start of every Claude Code session in this repo. It is the source of truth for project context and working rules. Read it fully before making any changes.

---

## 1. What this is

An internally owned supply chain system for Badr El-Din Farms (Egypt), replacing manual WhatsApp/Excel order tracking for the "Dots" basil seed drink line. Built at near-zero cost on Google Workspace (Sheets + Apps Script) with two GitHub Pages frontends.

- **Product:** Dots, 13 flavors — Mango, Blueberry, Strawberry, Pineapple, Lemon-Mint, Cantaloupe, Coconut, Watermelon, Orange, Peach, Red Apple, Green Apple, Pomegranate
- **Pack size:** cases of 12 (one unit = one case)
- **Distributors:** up to ~20, currently includes Family Alex, Cairo Office
- **Owner:** Digital Transformation Manager, Lean Six Sigma Black Belt, applies process-improvement discipline throughout — expect that standard in code review too.

---

## 2. Architecture

```
order.html (GitHub Pages)         index.html (GitHub Pages)
distributor-facing                internal delivery team
        │ fetch (POST/GET)               │ fetch (POST/GET)
        ▼                                  ▼
            Apps Script Project (single project, shared scope)
  Router.gs        → single doGet/doPost entry point, routes by `action`
  Code.gs          → CONFIG, setup, onFormSubmit, WhatsApp helper, Sync Errors
  OrderBackend.gs  → submitOrder() only
  DeliveryBackend.gs → getOpenOrders(), submitDelivery() only
  Inventory.gs     → Inventory spreadsheet + movements
  ReturnsBackend.gs → submitReturn() only
  ProductionBackend.gs → submitProduction() only
        │                                  │
        ▼                                  ▼
Badr El-Din Order Tracker         Badr El-Din Inventory Tracker
(main spreadsheet)                 (separate file — access control)
Orders, Deliveries, Returns,       Lists, Inventory Movements,
Productions, KPI Dashboard,        Inventory Balance
Summary, Sync Errors, Instructions
```

**Strict separation of concerns:** `Router.gs` has zero business logic — it only parses `action` and calls the right backend function. `OrderBackend.gs` only handles order submission. `DeliveryBackend.gs` only handles delivery logic. `Inventory.gs` only handles the Inventory spreadsheet. Do not blur these lines when adding new files/functions — new domains (e.g. Returns) get their own `*Backend.gs` file.

**Why two spreadsheets:** Inventory access needs to be controlled independently from Orders/Deliveries/KPI access.

---

## 3. Sheet structures (current)

**Main spreadsheet ("Badr El-Din Order Tracker"):**
| Sheet | Type | Columns |
|---|---|---|
| `📋 Orders` | static, trigger-written | Order_ID, Timestamp, Distributor, Total_Ordered, Status, Notes. Single header row. |
| `📋 Order_Lines` | static, trigger-written | Line_ID (OL-XXX), Order_ID (FK), Flavor_ID (FK), Flavor_Name, Qty_Ordered. One row per non-zero flavor per order. |
| `🚚 Deliveries` | static, trigger-written | Delivery_ID, Order_ID, Date_Delivered, Distributor, Total_Delivered, Notes. Single header row. |
| `🚚 Delivery_Lines` | static, trigger-written | Line_ID (DL-XXX), Delivery_ID (FK), Order_ID, Flavor_ID (FK), Flavor_Name, Qty_Delivered. One row per non-zero flavor per delivery. |
| `↩️ Returns` | static, trigger-written | Return_ID, Date, Distributor, Total_Returned, Notes. Single header row. |
| `↩️ Return_Lines` | static, trigger-written | Line_ID (RL-XXX), Return_ID (FK), Flavor_ID (FK), Flavor_Name, Qty_Returned. One row per non-zero flavor per return. |
| `🌱 Productions` | static, trigger-written | Production_ID, Date, Lot_Number, Total_Produced, Notes. Single header row. |
| `🌱 Production_Lines` | static, trigger-written | Line_ID (PL-XXX), Production_ID (FK), Flavor_ID (FK), Flavor_Name, Qty_Produced. One row per non-zero flavor per production batch. |
| `📊 KPI Dashboard` | formula-driven | 5 summary cards (Total Orders, Total Deliveries, Overall Fill Rate %, Pending/Partial count, Avg Fulfillment Cycle Time). Hidden helper block cols X:Z computes per-order cycle time — reused by Per-Distributor Summary. |
| `📈 Per-Distributor Summary` | formula-driven | one row per active distributor: Orders Received, Cases Ordered, Cases Delivered, Fill Rate %, Avg Cycle Time, Pending Orders. SUMPRODUCT+VLOOKUP joins Order_Lines/Delivery_Lines to Distributors. |
| `🌿 Per-Flavor Breakdown` | formula-driven | one row per active flavor: Cases Ordered, Delivered, Returned, Fill Rate %, Current Stock (via IMPORTRANGE from Inventory Tracker, staged in hidden cols K:Q), Below Reorder Threshold (reads 📉 Reorder_Points), Cases Produced. |
| ` Sync Errors` | append-only | Timestamp, Function, Error, Order ID, Distributor, Details |
| ` Instructions` | static reference | human-readable team guide |
| `⚙️ Products` | static reference, seeded | Product_ID (PK), Product_Code, Product_Name, Pack_Size, Is_Active |
| `🌿 Flavors` | static reference, seeded | Flavor_ID (PK), Product_ID (FK), Flavor_Name, Is_Active |
| `🏢 Distributors` | static reference, seeded | Distributor_ID (PK), Name, Slug, Phone, CallMeBot_ApiKey, Is_Active |
| `🔗 Distributor_Products` | static reference, seeded | DP_ID (PK), Distributor_ID (FK), Product_ID (FK) |
| `📱 Team_Contacts` | static reference, seeded | Contact_ID (PK), Name, Phone, CallMeBot_ApiKey, Is_Active |
| `📉 Reorder_Points` | static reference, seeded | Flavor_ID (FK), Min_Stock_Cases |

`_ARCHIVE_Orders`, `_ARCHIVE_Deliveries`, `_ARCHIVE_Returns`, `_ARCHIVE_Productions` — grey tabs, frozen, read-only wide-format historical data preserved from pre-migration. Not read by any live code path.

**Inventory spreadsheet ("Badr El-Din Inventory Tracker"), separate file:**
| Sheet | Type | Columns |
|---|---|---|
| ` Lists` | static reference | Warehouses, Sources/Destinations |
| ` Inventory Movements` | append-only ledger | Date, Type, Reference, Production Date, Product, Flavor, Quantity (signed), Warehouse, Source/Destination, Notes. Type ∈ {Opening Balance, Production In, Delivery Out, Return, Write-off}. Production Date (col 4) is populated only for Production In / Write-off — Delivery Out and Returns leave it blank |
| ` Inventory Balance` | formula-driven | one row per flavor, SUMIFS by Type, Current Stock = sum of row, red-flagged below threshold |

**Low-stock thresholds (Current Stock, cases):**
Mango/Blueberry/Lemon-Mint/Strawberry → 75 · Coconut/Orange → 25 · everything else → 50

---

## 4. Hard conventions — do not deviate without explicit approval

- **Sheet lookup:** always `ss.getSheets().find(s => s.getName().includes("keyword"))` via the `_getSheet(ss, keyword)` helper. Never `getSheetByName()` with exact names.
- **Leading space in tab names:** sheets are created as `insertSheet(" SheetName")`. `_getSheet`'s fuzzy match hides this for lookups, but **hardcoded formula strings referencing a sheet by name must include the leading space exactly** or `SUMIFS`/cross-sheet formulas return `#N/A`.
- **Spreadsheet access:** always `SpreadsheetApp.openById(ssId)` with the ID from `PropertiesService.getScriptProperties()`. Never `getActiveSpreadsheet()` inside triggers — returns `null` in trigger context.
- **No pre-filled formulas in transactional sheets** (Orders, Deliveries, Inventory Movements, and any future sheet of this type, e.g. Returns). All values written as static data by trigger/API functions. Pre-filled formulas break `getLastRow()`-based sequential ID generation.
- **Template tags:** `<?!= ?>` not `<?= ?>` in `.html` templates served via `HtmlService` — `=` HTML-escapes JSON and breaks client-side `JSON.parse`.
- **`requireNumber()` does not exist** on `DataValidationBuilder` in Apps Script. Don't use it.
- **Non-blocking cross-spreadsheet writes:** any write from the main spreadsheet's flow into the Inventory spreadsheet (e.g. `_logDeliveryOutMovements`) must be wrapped in try/catch, with failures routed to `_logSyncError(functionName, err, orderId, distributor, details)` — never let an Inventory write failure block the primary transaction (delivery, return, etc.) from being recorded.
- **Production Date placeholder:** `_logDeliveryOutMovements` and `_logReturnInMovements` both write an empty string `""` in the Production Date position (col 4) of their row arrays — required so the row stays 10 columns wide and nothing shifts into the wrong column. `_logProductionInMovements` and `_logWriteOffMovements` write a real date there instead.
- **Inventory Balance SUMIFS column constants:** `_buildInventoryBalanceSheet`'s `flavorCol`/`typeCol`/`qtyCol` constants are hardcoded column letters referencing ` Inventory Movements`, not formula-derived — currently `flavorCol` = col F, `typeCol` = col B, `qtyCol` = col G. If Inventory Movements columns are ever added/reordered again, these three constants must be updated to match, or every SUMIFS in Inventory Balance silently returns wrong totals.
- **`setupEntireSystem()`:** rebuilds all main-spreadsheet sheets from scratch. **Never re-run without explicit instruction.** New sheet builders are written as temporary standalone wrapper functions (created, run once manually via the Apps Script editor, then deleted) instead of being added to `setupEntireSystem()`, unless explicitly approved otherwise.
- **Deployment:** code changes to files used by the live web app require **Deploy → Manage deployments → edit existing deployment → Version: "New version" → Deploy**. `clasp push` only uploads code — it does **not** create a new deployment version or run any function. Selecting an existing numbered version (not "New version") silently leaves the deployment on old code with no error. Always confirm the post-deploy screen shows a fresh timestamp.
- **`clasp push` invocation:** always run as `cd ~/Downloads/Claude/badr-eldin-system && clasp push`, not bare `clasp push` — a bare run failed once in this sandbox with a DNS resolution error.
- **`CONFIG.FLAVORS`/`CONFIG.DISTRIBUTORS` are no longer the live source of truth** (as of DB Migration Phase 1): `Router.gs`, `OrderBackend.gs`, `DeliveryBackend.gs`, `ReturnsBackend.gs`, `ProductionBackend.gs`, and the four `_log*Movements` functions in `Inventory.gs` all call `_getConfig(ss)` instead, which reads the `🌿 Flavors`/`🏢 Distributors` sheets at runtime. `CONFIG.FLAVORS`/`CONFIG.DISTRIBUTORS` still exist in Code.gs and are still used — but only by setup-time sheet builders and the legacy Google Form flow (`onFormSubmit`), neither of which run during a live request. New live-path code must use `_getConfig()`, not `CONFIG`.

---

## 5. Working rules — non-negotiable

1. **Always explain the plan and get explicit approval before changing any code.** Do not write or edit files until the plan is confirmed.
2. **Break work into clearly numbered sub-tasks** and tackle one at a time.
3. **When debugging, add Logger statements and diagnose first** — don't guess and rewrite.
4. **Never re-run `setupEntireSystem()`** unless explicitly instructed.
5. **Always show exact line-level changes**, not paraphrased descriptions of what changed.
6. Maintain strict file separation per Section 2 — don't add business logic to `Router.gs`, don't mix domains across backend files.
7. Direct, technical communication. Honest pushback on ideas when warranted — don't just agree.

---

## 6. Known bug patterns (don't reintroduce these)

| Symptom | Cause | Fix |
|---|---|---|
| `requireNumber is not a function` | Not a real `DataValidationBuilder` method | Don't use it |
| `#N/A` across a formula-driven sheet | Hardcoded sheet-name string missing the leading space | Match the actual tab name exactly, leading space included |
| `#REF!` in KPI Dashboard | Rows were deleted (not cleared) from Orders, breaking row-referenced formulas | Clear rows, never delete them; rebuild via temp wrapper if needed |
| Inventory sync silently failing in production but working in editor test runs | Live deployment pinned to an old version — "New version" was never actually selected on a prior deploy | Explicitly select "New version" in Manage Deployments; confirm fresh timestamp |
| `getActiveSpreadsheet()` returns null | Called inside trigger context | Use `SpreadsheetApp.openById()` from ScriptProperties instead |

---

## 7. Completed: Returns feature

Returns flow — internal team logs returns (not distributor-initiated), separate from the Orders/Deliveries flow since returns don't reference an Order ID. All 5 sub-tasks are complete, including the end-to-end test.

- **ID format:** `RET-XXX` (sequential, like `DEL-XXX`)
- **`Returns` sheet** (main spreadsheet): `Return ID | Date | Distributor | [13 flavor cols] | Total Returned | Notes` — no Reason column for now (may be added later as an extra column without breaking anything).
- **Inventory sync:** logs to ` Inventory Movements` with `Type="Return"`, **positive** quantity (opposite sign of Delivery Out), same non-blocking try/catch → `_logSyncError` pattern as deliveries.
- **Frontend:** `returns.html`, GitHub Pages only, no Apps Script-hosted desktop fallback. Internal team facing — distributor dropdown + flavor steppers + notes, no order lookup needed.
- **Backend file:** `ReturnsBackend.gs` — `submitReturn(payload)`.
- **Router additions:** `doGet ?action=getReturnContext` (flavors + distributors), `doPost action=submitReturn`.

Sub-task order: (1) ✅ Returns sheet builder → (2) ✅ `ReturnsBackend.gs` + `_logReturnInMovements` in `Inventory.gs` → (3) ✅ Router wiring → (4) ✅ frontend (`returns.html`) → (5) ✅ end-to-end test. **Feature complete.**

---

## 7B. Completed: Production In and Write-off transactions

- **ID formats:** `PRD-XXX` and `WOF-XXX`, both generated via a shared `_nextSequentialId(counterProperty, prefix)` helper in `Code.gs` — a `LockService`-guarded Script Properties counter (`PRD_COUNTER` / `WOF_COUNTER`), not a row count. Avoids race conditions on concurrent submissions and survives manual row edits/deletes in the ledger.
- **`Productions` sheet** (main spreadsheet): `Production ID | Date | Lot Number | [13 flavor cols] | Total | Notes`.
- **Production In sync:** `_logProductionInMovements` in `Inventory.gs` logs to ` Inventory Movements` with `Type="Production In"`, **positive** quantity, `Source/Destination="Packing Line"`, real Production Date from payload. Non-blocking try/catch → `_logSyncError`, same pattern as deliveries/returns (the `Productions` sheet row is the protected primary transaction).
- **Write-off:** no separate sheet — `_logWriteOffMovements` in `Inventory.gs` is the entire transaction. `Type="Write-off"`, **negative** quantity, `Source/Destination="Write-off"`, optional Production Date. The team's required free-text reason/reference goes in Notes; the ledger's own Reference column holds the auto-generated `WOF-XXX`. **Not** wrapped in non-blocking try/catch by its caller — there's no separate primary record to protect, so failures must propagate rather than being silently absorbed.
- **Backend file:** `ProductionBackend.gs` — `submitProduction(payload)`.
- **Router additions:** `doGet ?action=getProductionContext` (flavors only, no distributors), `doPost action=submitProduction`, `doPost action=submitWriteOff` (calls `_logWriteOffMovements` directly — no Backend file wrapper exists for write-off).
- **Frontend:** `inventory-ops.html`, GitHub Pages only, live at `https://monagib3.github.io/delivery-app/inventory-ops.html`. Internal team facing — toggle between Production In and Write-off modes, shared flavor stepper grid. Quantity steppers use `step="any"` with `parseFloat` (not `parseInt`) to support decimal case quantities.

Sub-task order: (1) ✅ `Productions` sheet builder → (2) ✅ `ProductionBackend.gs` → (3) ✅ `_logProductionInMovements` in `Inventory.gs` → (4) ✅ `_logWriteOffMovements` in `Inventory.gs` → (5) ✅ Router wiring → (6) ✅ frontend (`inventory-ops.html`) → (7) ✅ end-to-end test. **Feature complete.**

**Decimal quantity support:** `index.html`, `returns.html`, and `inventory-ops.html` all now use `step="any"` on qty stepper inputs and `parseFloat` (not `parseInt`) for qty value parsing — consistent across all three frontends. `DeliveryApp.html` (the Apps Script-hosted mirror of `index.html`) was updated to match.

---

## 7C. Completed: home.html launcher page

Mobile-friendly internal launcher — single entry point for the ops team instead of bookmarking four separate URLs.

- **Frontend only:** `home.html`, GitHub Pages only. No backend calls — static links only, no `fetch`/`doGet`/`doPost`.
- **4 tappable cards**, RTL layout, Arabic primary label + English subtitle on each:
  - طلب جديد / New Order → `order.html`
  - تسجيل تسليم / Log Delivery → `index.html`
  - تسجيل مرتجع / Log Return → `returns.html`
  - عمليات المخزون / Inventory Ops → `inventory-ops.html`
- **Styling:** matches the existing palette exactly (`#1a3a2a` / `#2d6a4f` / `#52b788`), Cairo + DM Mono via the same Google Fonts CDN link used by the other frontends.
- **Live at:** `https://monagib3.github.io/delivery-app/home.html`.

**Feature complete.**

---

## 7D. Completed: order.html distributor picker

When `order.html` is accessed without a `?dist=` query param (e.g. via `home.html`'s "New Order" card), it now shows a distributor selection screen instead of going straight to the "Invalid Link" error screen.

- **Frontend only:** `order.html`, GitHub Pages. No `.gs`/Apps Script changes — purely client-side branching on `URLSearchParams`.
- **Trigger condition:** `dist` param missing (`null`) or present-but-empty (`""`) → picker screen. A present-but-unrecognized `dist` value still falls through to the existing "Invalid Link" error screen, unchanged. A valid `dist` value still goes straight to the order form, unchanged.
- **Picker UI:** new `#pickerScreen` block, title "اختر الموزع · Select Distributor", one full-width tappable `.distributor-card` per entry in `DISTRIBUTOR_MAP`. Styled with `order.html`'s own existing gold/ink palette (`--gold`, `--ink`, `--bg-card`, etc.) rather than the green palette used by `home.html`/`index.html`/`returns.html` — `order.html` has always had its own distinct theme, so the picker matches its host page instead of the other frontends.
- **Selection flow:** tapping a card calls `selectDistributor(key)`, which redirects to `order.html?dist=<key>` — same URL format the rest of the system already expects.
- **Backend file:** none — no Router/Backend changes were needed for this feature.

**Feature complete.**

---

## 7E. Completed: DB Migration — Phase 1 (Config Layer)

Foundational config-sheet layer for the DB migration effort — replaces hardcoded `CONFIG.FLAVORS`/`CONFIG.DISTRIBUTORS` reads in the live API path with sheet-backed, runtime-read config, and removes the hardcoded `DISTRIBUTOR_MAP` from `order.html`.

- **Six new sheets** in the main spreadsheet: `⚙️ Products`, `🌿 Flavors`, `🏢 Distributors`, `🔗 Distributor_Products`, `📱 Team_Contacts`, `📉 Reorder_Points` — see §3 for columns. Seeded once via a temporary `_runBuildConfigSheets()` wrapper (created, run manually, then deleted per convention) calling `_buildConfigSheets(ss)` in `Code.gs`. Never added to `setupEntireSystem()`.
- **`_getConfig(ss)`** in `Code.gs` — reads `🌿 Flavors`/`🏢 Distributors` at runtime, returns `{ flavors: [...active Flavor_Name strings...], distributors: [{id, name, slug, phone, apiKey}, ...active only...] }`. Read-only, no side effects.
- **Live API path migrated**: `Router.gs`, `OrderBackend.gs`, `DeliveryBackend.gs`, `ReturnsBackend.gs`, `ProductionBackend.gs`, and the four `_log*Movements` functions in `Inventory.gs` (`_logDeliveryOutMovements`, `_logReturnInMovements`, `_logProductionInMovements`, `_logWriteOffMovements`) all call `_getConfig()` instead of `CONFIG.FLAVORS`/`CONFIG.DISTRIBUTORS`. `Inventory.gs`'s three setup-only sheet builders (`_buildListsSheet`, `_buildInventoryMovementsSheet`, `_buildInventoryBalanceSheet`) deliberately keep `CONFIG.FLAVORS`/`CONFIG.DISTRIBUTORS` — they only run during one-time `setupInventorySystem()`, never during a live request. `_notifyDelivery`, `_notifyReturn`, `_notifyProduction` each gained an `ss` parameter to reach `_getConfig()`. `Inventory.gs` also gained a local `_getMainConfig()` helper (opens the *main* spreadsheet by `SPREADSHEET_ID`, since Flavors/Distributors live there, not in the Inventory spreadsheet) for use by the four `_log*Movements` functions only. `CONFIG.TEAM_WHATSAPP` was explicitly left alone — out of scope for this phase.
- **`?action=getDistributors`** added to `Router.gs` — returns `{slug, name}` pairs only (never `phone`/`apiKey`) for the public `order.html` frontend.
- **`order.html`**: hardcoded `DISTRIBUTOR_MAP` removed. Fetches `?action=getDistributors` on load; the distributor picker screen and the `?dist=` slug lookup are both now driven by the live sheet. Fetch failure shows the existing `#errorScreen` with distinct text from the "unrecognized distributor" case.
- **Naming fix (side effect, not just cosmetic):** `order.html` previously posted `"Family Alex"` (no hyphen) as `payload.distributor`, which never matched the canonical `"Family - Alex"` (with hyphen) used in `CONFIG.DISTRIBUTORS`/the Distributors sheet — so that distributor never received a WhatsApp confirmation, and their orders didn't roll up correctly in Summary. Fixed going forward now that the name comes straight from the sheet. Historical rows are **not** retroactively fixed — see the new backlog item in §9.
- **Out of scope for this phase**: `Code.gs`'s own `CONFIG.FLAVORS`/`CONFIG.DISTRIBUTORS` references (sheet headers, column-width math, the legacy Google Form builder, `onFormSubmit`) were left untouched — different concern, not migrated. `Products`/`Distributor_Products`/`Team_Contacts`/`Reorder_Points` sheets are seeded but not yet read by any code — reserved for later DB migration phases.

Sub-task order: (1) ✅ Six Config sheets (`_buildConfigSheets` + temp wrapper) → (2) ✅ `_getConfig()` helper → (3) ✅ Replace `CONFIG.FLAVORS`/`CONFIG.DISTRIBUTORS` across the live API path → (4) ✅ `?action=getDistributors` route + `order.html` fetch-based picker → (5) ✅ Documentation update (this section). **Feature complete.**

---

## 7F. Completed: Fixed Home button on internal pages

Adds a persistent way back to `home.html` from the three internal-team frontends, so the team never has to use the browser back button or retype a URL mid-task.

- **Files touched:** `index.html`, `returns.html`, `inventory-ops.html` only. `order.html` (distributor-facing) and `home.html` itself were left untouched — out of scope.
- **Markup:** a single `<a href="home.html" class="home-btn">🏠</a>` placed immediately after `<body>` in each file.
- **Style:** `.home-btn` — `position: fixed; top: 12px; left: 12px; z-index: 1000;`, 44×44px circle, white background / `--green-dark` icon color, using each file's own existing `--white`/`--green-dark`/`--shadow` CSS variables (all three files already share the same palette). z-index 1000 sits above every existing fixed/sticky element in these files (`.header` sticky z-100, `.loading-overlay` fixed z-999, `.modal-overlay` fixed z-200) with no overlap, since none of those occupy the top-left corner.

**Feature complete.**

---

## 8. Glossary

| ID | Generated by | Meaning |
|---|---|---|
| `ORD-XXX` | `onFormSubmit` / `submitOrder` | Order ID |
| `DEL-XXX` | `submitDelivery` | Delivery ID |
| `RET-XXX` | `submitReturn` | Return ID |
| `PRD-XXX` | `submitProduction` | Production ID |
| `WOF-XXX` | `_logWriteOffMovements` | Write-off ID |
| `OL-XXX` | `_nextSequentialId("OL_COUNTER")` | Order Line ID |
| `DL-XXX` | `_nextSequentialId("DL_COUNTER")` | Delivery Line ID |
| `RL-XXX` | `_nextSequentialId("RL_COUNTER")` | Return Line ID |
| `PL-XXX` | `_nextSequentialId("PL_COUNTER")` | Production Line ID |
| `SPREADSHEET_ID` | Script Property | Main Order Tracker spreadsheet file ID |
| `INVENTORY_SPREADSHEET_ID` | Script Property | Inventory spreadsheet file ID |

---

## 9. Backlog

| Item | Notes |
|---|---|
| DB Migration — Phase 2: Normalize transactional tables (IN PROGRESS) | Sub-task 2.1 complete (sheets built, archives created). Sub-task 2.2 complete (historical data migrated — 2 Orders/22 lines, 2 Deliveries/21 lines, 1 Return/1 line, 2 Productions/12 lines). Sub-task 2.3 complete (`OrderBackend.js`/`DeliveryBackend.js` rewritten to narrow schema; `ORD_COUNTER`/`DEL_COUNTER` added to the counter registry, seeded to 2 via one-shot `_initOrderDelCounters()` — run once, then deleted; deployed as Version 16). Sub-task 2.4 complete (`ReturnsBackend.js`/`ProductionBackend.js` rewritten to narrow schema: `submitReturn()` writes a header row to `↩️ Returns` + line rows to `↩️ Return_Lines`; `submitProduction()` writes a header row to `🌱 Productions` + line rows to `🌱 Production_Lines`. `RET_COUNTER`/`RL_COUNTER`/`PL_COUNTER` added to the counter registry; `PRD_COUNTER` was already correct at 4 and was deliberately left untouched. Seeded via one-shot `_initRetPrdCounters()` — run once, then deleted. Final counter state: `RET_COUNTER=1`, `RL_COUNTER=1`, `PL_COUNTER=12`, `PRD_COUNTER=4`. Deployed as Version 17, end-to-end tested, test rows cleaned up. Known stale comment in `Code.js`'s counter registry: `OL_COUNTER`/`DL_COUNTER` are still marked "not yet wired to a caller" even though they've been wired since Sub-task 2.3 — not a bug, left as-is, out of scope for 2.4.) Complete — Sub-task 2.5 was absorbed into Sub-task 2.4. |
| DB Migration — Phase 3: Fix reporting layer | Complete — Three new reporting sheet builders added to Code.js: `_buildKPIDashboardCards` (replaces `_buildKPIDashboard`), `_buildDistributorSummarySheet` (replaces `_buildSummarySheet`, renames tab to `📈 Per-Distributor Summary`), `_buildFlavorBreakdownSheet` (new tab). Old `_buildKPIDashboard` and `_buildSummarySheet` deleted. `setupEntireSystem()` dead calls removed. Built via temporary `_runBuildPhase3Reports()` wrapper (run once, then deleted per convention). Deployment note: IMPORTRANGE IFERROR wrapper suppresses the `#REF!` auth prompt — to authorize, temporarily remove IFERROR from K1, press Enter, hover for Allow Access, then the builder restores IFERROR on next run. Known bug fixed during deployment: `setFrozenColumns(1)` conflicts with merged A1:H1 title cell — resolved by removing the freeze. |
| Historical data cleanup — Orders sheet | Rows submitted before the Sub-task 4 naming fix have distributor name "Family Alex" (no hyphen), not matching the canonical "Family - Alex" now used everywhere. These rows don't roll up correctly in Summary's per-distributor grouping. Needs a one-time manual correction pass on the Orders sheet — separate from further DB migration work. |
| Input validation — enforce non-negative Qty fields across all four backends | Not started |
| Partial order close-out flow — status model redesign (Partial currently treated as open workload) | Not started |
| Production efficiency metric — Cases Produced/Delivered shown as separate columns for now; define and add ratio metric once 3-6 months of history exists | Deferred |
| Unify Reorder_Points sheet with Inventory Balance hardcoded thresholds — two sources of truth for the same concept | Not started |
| BigQuery + Connected Sheets migration — replace IMPORTRANGE with BigQuery as reporting layer foundation; enables Looker Studio executive dashboard | Long-term |
