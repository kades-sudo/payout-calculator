# Payout Calculator — Phase 1–4 (+ Phase 3 refinements, update, and format match)

A single-page, client-side web app that automates OMS transaction enrichment, master-list matching, and Process Raw Details generation for payout processing.

Open `index.html` in a browser (or publish it as-is) — no build step or backend required. All parsing and lookups run in the browser via [SheetJS](https://sheetjs.com/).

## Master List workbook ingestion (Phase 2)

The Master List / Configuration upload recognizes up to six sheets by fuzzy name match (case, spacing and punctuation-insensitive), configured in `SHEET_ALIASES` in `index.html`:

| Sheet | Required this phase | Notes |
|---|---|---|
| `Master_List` | Yes | Primary key `MID` (+ `Merchant Key` for shared-MID merchants). Merged across every uploaded workbook, so separate Daily/Weekly lists can be uploaded together. |
| `Card Type Classification` | Yes | Column A = lookup key, column B (`Card Type 1`) = output **Card Type**, column C (`Card Type 2`) = output **Card Type 2** — verified against a real export; header names are matched first, with A/B/C position as fallback. |
| `Terminal_Mapping` | Recommended | Keyed by MID + Terminal ID → Merchant Key + Account Code. Used from Phase 4 onward for the shared-MID chain (see below). |
| `Bank_Cost` | Recommended | Keyed by Cost ID; one rate column per payment group. Used for the payout report's Bank Rate/Bank Fee. |
| `Special_Rate` | No (parsed, not yet used) | Keyed by Account Code; overrides for Account Code 005/5. **Not applied yet** — see Known gaps below. |
| `Card_Types` | Recommended | One row per payment group, naming which Master_List column holds the merchant rate and which Bank_Cost column holds the bank rate. Drives the payout report's column set dynamically — add a row here to add a payment group, no code change needed. |

## Master List / Configuration connections (Phase 4)

Verified against a real Master_List + Terminal_Mapping + Bank_Cost + Card_Types workbook and a real Daily Payout report (3,707-transaction cross-check, exact to the cent on every formula below):

1. **Shared-MID merchant resolution**: for each transaction, `MID + Terminal ID` is looked up in `Terminal_Mapping` to get a `Merchant Key` + `Account Code`; `MID + Merchant Key` is then looked up in `Master_List` (blank Merchant Key = the ordinary one-merchant-per-MID case) to get the merchant's name, rates and configuration. This is now how **both** Process Raw Details' `Merchant` column and the payout report resolve merchants — a single physical MID can host several distinct merchants, disambiguated only by which terminal took the transaction. Falls back to a plain MID match when a terminal isn't in `Terminal_Mapping`.
2. **Cost ID → Bank_Cost**: the merchant's `Cost ID` selects its Bank_Cost row.
3. **Card_Types "Master List Rate Column" → Master_List**: for a transaction's payment group (its `Card Type 2`), this names which Master_List column holds that merchant's rate.
4. **Card_Types "Cost Table Column" → Bank_Cost**: same, for which Bank_Cost column holds the bank's rate.
5. **Below Amount Rule**: a merchant's transactions split into two report lines — `>= threshold` (normal) and `< threshold` (a `"<Merchant> | BELOW <threshold>"` line). Only the below-threshold line carries a per-transaction charge (`Cleared Txn × Rate Per Trnx`); `Gain For Bank Charge` and `Payout Processing Fee` apply once, on the normal line only.
6. **Report formulas** (per merchant line × payment group): `Bank Fee = Gross × Bank Rate`, `Noqoody Charge = Gross × Merchant Rate`, `Noqoody Profit = Noqoody Charge + Charge/Txn − Bank Fee − Refund`, `Internal Transfer = Gross − Noqoody Charge − Charge/Txn − Refund`. `Gross Collection`/`Cleared Txn` count `Sale`-type transactions only; `Refund` sums the absolute value of every non-`Sale` transaction (Refund, Dispute, ...).

### Known gaps (flagged in the app, not silently guessed)

- **Special_Rate (Account Code 005/5)** isn't applied — no sample data existed to verify which Special_Rate column an Account-Code-5 transaction should use, and guessing would silently misprice real transactions.
- **Rent, Portal Amount** are external reconciliation figures (payout portal) that don't exist in any uploaded sheet. They're editable per report line in the UI, defaulting to 0; everything downstream recalculates live.
- **Wallet Classification / Classification** remain unpopulated ("rule pending") — the reference data shows real values for these, but the signal that decides them isn't present in Card Number, Text Message, or any of the six config sheets. Confirmed *not* load-bearing for the payout report itself: the report groups by `Card Type 2`, which is fully resolved.

## Phase 3: calculation refinements, frequency filtering, and Excel export

- **Actual Bank Charges** is now computed, not manual: `SUM(Commission)` for the report line's `Sale`-type transactions (same cross-payment-group scope as Bank Charges itself), per the OMS Commission field. Cross-checked against a real merchant: computed Bank Fee (0.75) and computed Actual Bank Charges (0.76) landed a cent apart — exactly the kind of small, real-world reconciliation variance this metric exists to surface.
- **Payout Frequency filtering**: the report now only includes merchants whose Master_List `Payout Frequency` matches the selected Daily/Weekly toggle. A blank/unconfigured Payout Frequency matches neither and is excluded from both, same treatment as an unresolved merchant; the report banner reports how many transactions were excluded this way.
- **Grid View horizontal scrolling** was fixed at the root cause: the flex/grid ancestors of the scrollable tables had no `min-width: 0`, so a wide table forced the whole page wide instead of scrolling internally. Fixed with `minmax(0,1fr)` / `min-width:0` on the layout chain — applies to Process Raw Details and the Daily/Weekly Payout Report Grid Views alike.
- **Section-aware sticky headers** (web app): implemented with plain CSS `position:sticky` — the payout report's column headers stay pinned while scrolling, and each Account Code's own label sticks directly below them, swapping to the next Account Code's label as you scroll into that section. The first three columns (Account Code, MID, Merchant) stay frozen while scrolling horizontally. This is **not** in the exported Excel file — see the library note below.
- **Combined Excel export**: unchanged behavior (was already a single workbook, two worksheets, from Phase 4) — `Download Payout Report` produces `Payout_Report_[Frequency]_[PostingDate].xlsx` with a `DAILY`/`WEEKLY PAYOUT` sheet and a `Process Raw Details` sheet. The standalone `Download Process Raw Details` button also stays, for the Phase 1 pre-payout verification checkpoint.
- **Excel toggle (+/-) sections**: implemented via native Excel column/row outline grouping (`outlineLevel`) — each payment group's 11 columns collapse independently (bounded by blank spacer columns), and each Account Code's merchant rows collapse independently (bounded by that code's own SUBTOTAL row). Verified in the raw exported XML, not just assumed.
- **Excel column header colors**: implemented using the real Excel built-in Cell Style colors (Input `FFCC99`, Good `C6EFCE`, Bad `FFC7CE`, Neutral `FFEB9C`) plus this workbook's own **actual** theme tints for "Accent 1, Lighter 40%" (`95B3D7`) and "Accent 3, Lighter 40%" (`C3D69B`) — computed from the reference report's own `theme.xml`, not assumed Office defaults. Applied identically to every repeated Card Type section header row. Verified byte-for-byte in the exported file's `styles.xml`.
- **Library swap required**: the free SheetJS build used through Phase 1–4 (`xlsx.full.min.js`) cannot write cell colors or freeze panes at all — verified empirically, not assumed (its writer hardcodes the sheet view and never serializes a `.s` style). Switched to `xlsx-js-style` (same SheetJS 0.18.5 core, confirmed identical read behavior against the real Master List file, plus style-writing support) to make the header colors possible.
- **Excel freeze panes / section-aware sticky headers**: **not implemented in the exported file** — confirmed via the same source-level check that no SheetJS build (including `xlsx-js-style`) exposes a freeze-pane writer, and Excel itself has no concept of a header that changes based on scroll position within one sheet (only one static freeze split per worksheet exists at all). Per the user's decision, this is built into the web app's Grid View instead (see above); the exported Excel sheet gets each Account Code section's own repeated header row (Account Code label → payment-group header → column header) rather than a frozen/sticky one.

## Phase 3 update: MID display bug, reconciliation re-verification

- **Bug found and fixed**: Merchant ID reached the Excel export as whatever type the OMS Raw upload stored it as. When it's genuine Excel *text*, no problem; when it's a genuine Excel *number* (common in real OMS exports, and confirmed by reproducing it with a numeric-MID test file), a 15-digit MID like `777100432320320` would render in Excel's default General number format as scientific notation (`7.77104E+14`) — reading as "MID not displaying correctly." Fixed by forcing Merchant ID to text on export, in both the Payout Report and Process Raw Details sheets (and the standalone Process Raw Details download). Verified in the raw exported XML: the cell now carries `t="str"` regardless of the source file's cell type.
- **Section 1 (reconciliation formulas) and Section 4 (independent Card Type toggles)** were re-verified against the same real 3,707-transaction dataset and the same direct XML inspection used last phase — both reproduce identically (Total Transfer = SUM of main + below-threshold Merchant Payout to the cent; Rent/Gain/Fee present only on the main row, zero on the below row; column `outlineLevel` correctly bounded per payment group by spacer columns). No regressions found, no changes were needed.
- **Section 3 (freeze MID/Merchant columns in Excel)**: at the time, re-tested specifically against `xlsx-js-style` in case it added freeze-pane support beyond the official SheetJS build — it didn't. See below: this was superseded once a real reference file proved freeze panes were achievable via a different technique.

## Format match against a real Daily Payout export

The user supplied a real Daily Payout Excel file (format only, one sheet). Reviewed it byte-for-byte (raw XML, not just parsed values) before changing anything, per instruction. It corrected several assumptions:

- **Column toggle structure was wrong.** All 11 metric columns per Card Type section were being grouped as one collapsible block. The real structure: **Gross Collection stays permanently visible** (never part of the group); the other 10 columns (Bank Rate → Internal Transfer) are `outlineLevel="1"` **and** `hidden="1"` — collapsed by default, not just collapsible. Fixed: `downloadPayoutReport` now excludes the first `GROUP_SUBCOLS` entry from the hidden/outline range and marks the remaining 10 both `level:1` and `hidden:true`.
- **Leading column is "No.", not "Account Code".** A sequential counter that resets per Account Code section (1, 2, 3, ... within each section) — Account Code only ever appears as the section banner. Fixed in both the web app Grid View and the Excel export; `line.accountCode` as a column is gone.
- **Two header colors were wrong**, corrected by reading the actual fill values out of the reference file's `styles.xml` rather than computing them: Noqoody Charge is `B4C6E7` (not the computed `95B3D7`), Noqoody Profit is `C6E0B4` (not the computed `C3D69B`). The other four (Input `FFCC99`, Good `C6EFCE`, Bad `FFC7CE`, Neutral `FFEB9C`) were already exact. Header text in the reference isn't bold — removed the forced bold too.
- **Freeze panes are real and now implemented.** The reference file freezes exactly the first 3 columns (No./MID/Merchant) — `xSplit="3"`, no row freeze at all (no `ySplit`, so no sticky header row in the Excel file itself). This reverses what was reported as a hard library limitation: neither SheetJS build exposes a freeze-pane *write API*, but the file format itself supports it fine, so `downloadPayoutReport` now patches it in directly — generate the workbook normally, unzip the result in-memory with JSZip, replace the one `<sheetViews>` element in the report sheet's XML with a hardcoded frozen-pane block, rezip. Falls back silently to the unpatched file if JSZip isn't available or the patch fails, so this can never block a download. Verified round-trip: the patched file re-opens cleanly and the pane survives.

Not addressed by this pass (no reference data covers it): the exact sort order of merchant rows within an Account Code section — the reference file's order doesn't obviously match ours (insertion order from the transaction stream) and no sort rule has been specified.

### Follow-up: two export bugs found in the generated file

Diagnosed from an actual generated workbook, by comparing its worksheet XML attribute-by-attribute against the reference:

- **No./MID/Merchant rendered as zero-width columns.** Introducing the `!cols` array for the toggles made the writer emit a `<col>` element for *every* column, and the ungrouped ones were passed `{}` — producing `<col min="1" max="1"/>` with no `width`. With no width on the element and no `<sheetFormatPr defaultColWidth>` to fall back on, Excel collapses those columns to nothing. Before the toggle work no `<cols>` element was written at all, so Excel's own defaults applied and the columns displayed fine — this was a regression, not a pre-existing issue. Fixed by giving every emitted column an explicit width (No. 4, MID 16, Merchant 30, Gross Collection 20, metrics 14, spacers 2).
- **The +/- toggles had no clickable control.** Excel decides whether to draw the outline bar from `outlineLevelRow`/`outlineLevelCol` on `<sheetFormatPr>`, and the writer never emits that element at all. The columns and rows carried correct `outlineLevel="1"` attributes, but with no outline bar there was nothing to click — so the detail columns sat hidden with no way to expand them. Fixed by injecting `<sheetFormatPr defaultRowHeight="15" outlineLevelRow="1" outlineLevelCol="1"/>` in the same JSZip patch step as the freeze pane, positioned between `<sheetViews>` and `<cols>` as the schema requires.
- Two attribute cleanups ride along in that patch: the writer emits a `level` attribute that isn't part of `CT_Col` (only `outlineLevel` is), and writes `hidden="true"` where Excel writes `hidden="1"`. Both are now normalized to match the reference exactly.

### Amount formatting

Decimal precision was taken from the reference export's own number formats rather than assumed — it uses three of Excel's built-in formats, and the app now applies the same three to the same columns:

| Column | Format | Built-in id |
|---|---|---|
| Gross Collection, Bank Fee, Noqoody Charge, Merchant Rate/Txn, Noqoody Charge/Txn, Refund, Noqoody Profit, Internal Transfer, and every Merchant Reconciliation / Transfer column | `#,##0.00` | 4 |
| Bank Rate, Merchant Rate | `0.00%` | 10 |
| Cleared Txn | `#,##0` | 3 |
| No., MID, Merchant | General | 0 |

In the Excel export these are applied as **cell number formats**, so values stay numeric (`t="n"`) and keep calculating — nothing is converted to text. The on-screen Grid View already had comma-separated amounts; three inconsistencies were corrected to match the same rules: rates now always show two decimals (`2.50%`, previously `2.5%`), Cleared Txn now renders as an integer count (`21`, previously `21.00`), and zero values now show as `0.00` rather than a bare `0`.

### Live Excel formulas in the export

The exported workbook carries **live formulas**, not just computed values — matching the reference export, which has 181 of them. Editing an input cell in Excel (Rent, Portal Amount, a rate, a Gross Collection) recalculates everything downstream.

Inputs stay plain values; derived cells carry a formula plus the already-computed value as a cached result, so figures read correctly before Excel recalculates. `<calcPr fullCalcOnLoad="1"/>` is patched into the workbook so Excel recalculates on open regardless.

| Cell | Formula |
|---|---|
| Bank Fee | `ROUND(Gross*BankRate,2)` |
| Noqoody Charge | `ROUND(Gross*MerchantRate,2)` |
| Noqoody Charge/Txn | `ROUND(ClearedTxn*RatePerTxn,2)` |
| Noqoody Profit | `ROUND(Charge+ChargePerTxn-BankFee-Refund,2)` |
| Internal Transfer | `ROUND(Gross-Charge-ChargePerTxn-Refund,2)` |
| Total Noqoody Profit / Bank Charges / Total Internal Transfer | `ROUND(<group1>+<group2>+…,2)` across every Card Type section |
| Bank Difference | `ROUND(BankCharges-ActualBankCharges,2)` |
| Transfer Deducted (Rent, Others) / (Rent) | `ROUND(TotalInternalTransfer-Rent[-Gain],2)` |
| Merchant Payout | `ROUND(TransferDeducted-PayoutProcessingFee,2)` |
| Total Transfer | `ROUND(SUM(MerchantPayout over the merchant's main + below rows),2)` — main row only |
| Difference | `ROUND(TotalTransfer-PortalAmount,2)` — main row only |
| Subtotal row | `SUM(<col><firstDataRow>:<col><lastDataRow>)` per Account Code, every numeric column except the two rate columns |
| Grand total row | `<subtotal1>+<subtotal2>+…` |

Cell references are generated from the report's own column layout rather than copied from the reference, since the number of Card Type sections is driven by the `Card_Types` sheet.

Verified by independently evaluating every formula in the exported file against its cached value: **1,726 of 1,726 match**, zero mismatches.

## Workflow (Phase 1 scope)

1. Upload the Master List / Configuration workbook(s) (see above).
2. Select the payout frequency — Daily or Weekly.
3. Upload the OMS Raw Transaction file (16-column export).
4. Click **Process Payout**. The app automatically:
   - Runs the Card Type lookup
   - Runs the Card Type 2 lookup
   - Matches each transaction's Merchant ID against the Master_List to populate Merchant
   - Flags any transaction whose MID isn't found in the Master_List as an unmatched record
   - Builds the Process Raw Details dataset
5. Review the Process Raw Details table, including the match-status flag per row.
6. Click **Download Process Raw Details** to export the exact displayed dataset as `Process_Raw_Details_[Frequency]_[PostingDate].xlsx`.

No sample files on hand? Use **"Load example data to try it out"** on the left rail — it fills in a small labeled example dataset (including one intentionally unmatched MID) so you can see the full pipeline run immediately.

## Deliberately out of scope for Phase 1

Wallet Classification and Classification are displayed as columns with a **rule pending** marker — the lookup source for these hasn't been defined yet, so no rule is invented. Everything else in the later-phase list (rate lookups, fees, refund/profit calculation, final report layout, etc.) is likewise not implemented yet; see the project spec for the full phase breakdown.
