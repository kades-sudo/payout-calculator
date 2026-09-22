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

The merchant row order within an Account Code section was left open by this pass and resolved later — see "Merchant sorting, numbering and the Below Amount row" below.

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

## Merchant sorting, numbering and the Below Amount row

- **Sorting**: merchants are ordered A–Z by Merchant Name. The sort happens at merchant level before report lines are built, so a merchant's main row and its Below Amount row stay together, and Card Type sections (which are column blocks, not rows) are never reordered independently. With the real dataset this reproduces the reference export's own row order exactly.
- **Numbering**: the `No.` column identifies a *merchant*, not a row. It advances once per merchant and resets per Account Code section; a Below Amount row carries the same number as the main row it belongs to (e.g. `1 MADHURA RESTAURANT` / `1 MADHURA RESTAURANT | BELOW 25` / `2 NEW BALANCE DOHA`).
- **Below Amount row display**: a Below Amount row is created only when at least one transaction actually satisfies `Gross Amount < Below Amount Rule` — never merely because the merchant has a rule configured. The comparison is strict, so a Gross Amount exactly equal to the threshold is *not* below it. The main row always remains visible.
- **Below Amount calculation** (unchanged, re-verified): `Noqoody Charge Per Txn` is `0` on the main row and `Cleared Txn × Merchant Rate Per Txn` on the Below Amount row, and flows into that row's Noqoody Profit and Internal Transfer.

Verified in both the web app and the exported workbook (identical order and numbering in each), with a targeted negative test: a merchant with a Below Amount Rule of 25 whose transactions are all exactly 25 produces no Below Amount row anywhere, while its main row still appears. All 1,726 exported formulas still evaluate to their cached values after the reorder, and the column widths, outline toggles, freeze pane and number formats remain intact.

## Header formatting, duplicate MID/TID validation, and the payout dashboard

**Audit status: PENDING VERIFICATION** — implemented and tested against the real Masterlist and a
3,707-row OMS dataset, but not yet confirmed against a production payout run. Open items are listed
at the end of this section.

### 1. Section column header formatting

Column headers wrap onto as many lines as the label needs and are bold everywhere they appear, in
both the web Grid View and the exported workbook, for every Card Type section, the Final
Reconciliation block and the Transfer block.

- **Web**: `.report-table thead th` sets `white-space:normal` + `font-weight:700`. Because the two
  header rows are `position:sticky`, their offsets can no longer be hard-coded once the text wraps to
  a variable height — `renderPayoutReport` measures both rendered rows and publishes `--grp-h` /
  `--head-h`, which the column-header row and the Account Code section label sticky offsets read.
  Measured on the real dataset: group row 32px, column row 74px, so the section label sticks at
  106.25px, exactly below both.
- **Excel**: header cells carry `font.bold` + `alignment.wrapText/vertical/horizontal`, and the header
  rows get an explicit `hpt: 46` — Excel does not auto-fit a row whose height was never written, so
  without this a wrapped header is clipped. 46pt fits three wrapped lines, covering the longest label
  in the report (`Transfer Deducted (Rent, Others)`).
- Subtotal rows, the Grand Total row and the Account Code label rows are bold across every column.
  `styleBoldCell` **merges** into any existing style rather than replacing it, so the number formats
  (`#,##0.00`, `0.00%`, `#,##0`) survive alongside the bold.

Verified by unzipping the generated workbook: `cellXfs` 4–10 all carry `wrapText="true"`, fonts 2–11
all carry `<b/>`, header rows 3 and 10 carry `ht="46" customHeight="1"`. Note that reading the file
back through SheetJS does **not** round-trip `.s`, so a read-back assertion reports no styles even
when the file is correct — the XML is the source of truth here.

### 2. Duplicate MID / TID validation (Master List, not the OMS file)

Runs during the validation stage, as soon as the Master List workbook is parsed and therefore before
any payout report is generated. **Nothing is removed, merged or rewritten** — rows are reported
exactly as uploaded, and no transaction is excluded from the payout calculation as a result.

Severity reflects what each repetition actually does to the lookups:

| Severity | Check | Why |
|---|---|---|
| Error | Master_List duplicate **MID + Merchant Key** | `buildMasterListIndex` keys on this pair; the last row silently wins and the others are unreachable |
| Error | Terminal_Mapping duplicate **MID + Terminal ID** | `buildTerminalMapIndex` keys on this pair; the first row wins, so a different Merchant Key or Account Code on later rows is ignored |
| Warning | One **Terminal ID** under more than one MID | A terminal's payouts could be attributed to the wrong merchant |
| Expected | Repeated **MID** across different Merchant Keys | The shared-MID case — several merchants trading under one physical MID |
| Expected | Repeated **MID** across different Terminal IDs | One merchant normally operates several terminals |

A plain "duplicate MID" check was deliberately **not** implemented against the OMS raw file: measured
on the real 3,707-row dataset it flags 3,688 rows (99.5%) and 3,655 TID rows (98.6%), because every
merchant naturally has many transactions. Reference Number is unique across all 3,707 rows there.

Reference Number, Card Type 2 and Posting Date are OMS transaction fields and do not exist in the
Master List, so they cannot be shown for a configuration duplicate; the panel states this.

**Finding on the supplied Masterlist**: `Terminal_Mapping` rows 213 and 429 are byte-identical
duplicate rows (MID `100005269900174`, TID `10034075`, Account Code 15). Harmless today because both
rows carry the same values, but it is a genuine duplicate and is reported as an Error. No duplicate
`MID + Merchant Key` exists in `Master_List`.

### 3. Payout dashboard

Icon-led stat tiles rendered above the report table, recomputed on every upload, frequency change and
reprocess (it is built inside `renderPayoutReport`, so it cannot drift from the table below it).

- **Categories are read from the `Card_Types` sheet at run time**, never hardcoded. Adding a row there
  adds a report section *and* a dashboard card with the same formulas, formatting, subtotal logic and
  dashboard integration, with no code change. `Others` collects every Card Type 2 present in the data
  that no `Card_Types` row defines.
- **Reference field**: `DASHBOARD_REFERENCE_FIELD` is set to `cardType2` **temporarily**, because
  `Classification` is not yet populated by the pipeline (`runPipeline` sets it to `""`). It is a
  single switch — point it at `classification` once that lookup rule exists.
- **Transaction Count** = cleared (Sale) transactions only, matching the report's own `Cleared Txn`.
- **Total Payout Amount** = `Internal Transfer` per category; the three headline tiles are Cleared
  Transactions, Total Payout Amount (Total Internal Transfer) and Total Profit (Total Noqoody Profit).
- **No double counting**: every figure is summed from `report.merchantLines` — never from a subtotal
  row, the Grand Total row or the Merchant Reconciliation block. Main and Below Amount rows are both
  merchant lines and are each included exactly once.
- Icons are inline SVG, one distinct glyph per category, with a stable hash-based fallback so a newly
  added Card Type gets its own consistent glyph without a code change.
- Stat-tile values use the font's proportional figures (`tabular-nums` is reserved for the aligned
  columns of the report table).

**`Others` is structurally zero on the payout tiles and that is correct**: a Card Type 2 with no
`Card_Types` row has no merchant rate, no bank cost column and no report section, so it never enters
an Internal Transfer or Profit calculation anywhere in the report. A non-zero Others *count* with a
0.00 Others *amount* is therefore a data-quality signal, and the app says so explicitly, naming the
excluded gross and the offending Card Type 2 values.

### 4. Dashboard reconciliation

Four checks re-derive each dashboard total from the merchant lines and compare it with the figure the
report itself publishes; any failure renders a blocking banner rather than showing a wrong number.

Verified on the real dataset (Daily and Weekly): all eight category cards and all three headline
figures match the independently re-summed table rows *and* the Grand Total row
(553 + 887 + 123 = 1,563 cleared; 66,464.23 total payout; −208.96 total profit).

The checks were also **negative-tested** so they are not vacuous: perturbing the grand-total transfer,
the grand-total profit, the Sale count, and an Account Code subtotal each fires exactly one check and
renders the banner.

### Testing performed

- Daily and Weekly payout data. The supplied Masterlist has no Weekly merchants (51 Daily, 204 blank),
  so a Weekly-tagged variant was generated to exercise the Weekly path with the same real data; both
  reconcile with zero mismatches.
- Empty-frequency case: the dashboard now renders with genuine zeros in every category instead of the
  panel showing nothing, with a banner explaining why there is no table.
- `Others`: a fixture retagging 37 Sale rows to a card type the classification sheet does not define.
  Totals stay at 1,563 (518 + 885 + 123 + 37), proving no double counting, and Others is flagged.
- Merchants with and without Below Amount Rules, including the negative case (a rule configured but no
  transaction below it produces no Below row).
- Export regression: 1,726 / 1,726 formulas still evaluate to their cached values; 114 columns all
  with explicit widths; 80 hidden columns; freeze pane (`xSplit="3"`) and
  `sheetFormatPr outlineLevelRow/outlineLevelCol` (the +/- toggles) intact.
- Layout: no horizontal page overflow at 1500px or at 430px (phone) width.
- A missing `<meta charset="utf-8">` was found and added — without it the em dashes in the app's
  notices render as mojibake depending on how the page is served.

## Transaction Type handling and the all-types transaction count

**Audit status: PENDING VERIFICATION.**

### Rules as implemented

| Transaction Type | Gross Collection | Cleared Txn | Refund column | Gain For Bank Charges |
|---|---|---|---|---|
| Sale | adds | counts | – | – |
| Refund | – | – | `ABS(Gross Amount)` | flat **+5** per transaction |
| Dispute | – | – | `ABS(Gross Amount)` | flat **+5** per transaction |
| Reversal | – | – | `ABS(Gross Amount)` **only if its own Gross Amount is negative** | never |
| anything else | – | – | `ABS(Gross Amount)` (original behaviour) | never — and reported in the UI |

The flat charge is `GAIN_PER_REFUND_DISPUTE = 5`, a business rule supplied directly and not derived
from any uploaded sheet. It is **added to** the Master_List `NOQOODY GAIN FOR BANK CHARGE`, not a
replacement for it.

**The flat charge is counted per merchant, over the merchant's entire transaction set, and applied on
the main (>= threshold) line.** This matters: every Refund and Dispute in the real data has a
*negative* Gross Amount, and the Below Amount split is `gross < threshold`, so all of them land on the
*below* line — where Gain For Bank Charges is always zero. Counting per line would silently drop the
charge for every merchant that has a Below Amount Rule.

A transaction type outside these four has no defined rule, so it keeps the original behaviour
(absolute amount into Refund) and is surfaced in a banner rather than being assumed to follow one of
the rules above. No rule was invented for it.

### Transaction count: two different figures, deliberately

- **Dashboard "Total Transactions"** counts **every** transaction type — Sale, Refund, Dispute,
  Reversal and any unrecognised type. Each category card shows this count, with the Sale subset as
  supporting text.
- **The report's `Cleared Txn` column stays Sale-only**, because it is the base for the
  per-transaction charge (`Cleared Txn x Rate Per Trnx`). Changing it would silently change payout
  money. The spec line "Sale -> adds to Gross Collection and transaction count" is read as defining
  this column; the "count all transactions" instruction is read as applying to the dashboard metric.

Both figures are reconciled separately: `Total Transaction Count` against all reported transactions,
and `Cleared (Sale) transactions` against the Sale subset and against the Account Code subtotals.

### Verification

A controlled fixture (`OMS_TxnTypes.xlsx`) was built because the real dataset contains **no Reversal
rows at all** (3,684 Sale, 1 Refund, 22 Dispute — every Refund and Dispute with negative gross).

*AL MARFAT TAILORING* (no Below rule, base gain 4) — 2 Sale, 1 Refund, 2 Dispute, 1 negative Reversal,
1 positive Reversal, 1 unknown type. Expected and produced: Gross `1,500.00`; Cleared `2`;
transactions `8`; Refund `385.00` (the positive-gross Reversal of 300 correctly contributes nothing);
Gain `19.00` = 4 + 5x3.

*MADHURA RESTAURANT* (Below rule 25, base gain 6) — main row Gain `16.00` = 6 + 5x2, proving the flat
charge survives even though both refunds sit on the below row; below row Refund `55.00`, Gain `0.00`.

On the real dataset: 1,565 transactions of which 1,563 Sale (the 2 non-Sale both Debit Card, which is
why that category reads 889 all-types vs 887 cleared); Gain total `126.00` = base `116.00` + 5x2, with
`AL SULTAN MEDICAL CENTER` moving 0 -> 5 and `NEW BALANCE DOHA` 4 -> 9. All five dashboard
reconciliation checks pass. Export re-verified: 1,726 / 1,726 formulas evaluate to their cached
values, 114 columns all with widths, 80 hidden and outlined, freeze pane and `sheetFormatPr` intact.

## Wallet Classification (Prefix + Wallet_Rules)

**Audit status: PENDING VERIFICATION.**

### The rule

Two new Master List sheets drive it. The prefix table is a **membership list, not a classifier** — it
answers only "is this card's BIN in the table?". `Card Type (On us/Off us)` still decides the card's
rail, and the prefix only chooses between that rule row's two values:

```
prefixSet = { first 6 digits of every Prefix row }
cardHead  = first 6 digits of Card Number
matched   = prefixSet.has(cardHead)
label     = matched ? rule.onMatch : rule.onNoMatch     // rule keyed on Card Type (On us/Off us)
```

Both sides are compared on their **first 6 digits only**, whatever their full length. That is the rule
as specified, and it is also all the feed allows: card numbers arrive masked after 6 digits
(`######XXXXXX####`).

**Why this design sidesteps the co-badging problem.** An earlier analysis found 6 prefix heads
carrying both NAPS and Visa transactions (`491227` alone splits 305 / 135). That would have been fatal
had the prefix table been used to *determine* the card's rail. It isn't — both sets match the prefix
but still receive different labels, because the card type differs. Nothing is mislabeled.

### Sheets

| Sheet | Shape | Notes |
|---|---|---|
| `Prefix` | one column (`PREFIX`) | 3,237 rows → 54 distinct 6-digit heads. Stored as numbers in the supplied workbook, which is safe here (all values start 4 or 5, no leading zeros to lose); `leadingDigits` reads the digits, not the cell type. |
| `Wallet_Rules` | 3 columns, 9 rows | `Card Type (On us/Off us)`, `Wallet Classification (Prefix Match)`, `Wallet Classification (No Match)`. First row per card type wins. |

Both are optional and resolved by the same fuzzy name matching as the other config sheets
(`prefix`/`bin`/`napstab`, `walletrules`/`walletclassification`/...), and both appear in the sheet-check
list on the left rail.

### Behaviour when something is missing — nothing is ever guessed

| Situation | Result |
|---|---|
| Either sheet absent | Column stays blank, keeps its **rule pending** marker, banner names which sheet is missing |
| Card type has no `Wallet_Rules` row | Cell shows **no rule**, banner names the card type and its transaction count |
| Card Number has fewer than 6 leading digits | Treated as no-match, counted and reported separately |

The **rule pending** marker on the Wallet Classification header is now conditional — it disappears once
both sheets are present. `Classification` still carries it, since that rule remains undefined.

### Verification

Against the supplied Masterlist and the 3,707-row NAPS_TAB sample, the app's output matches an
independent projection computed straight from the two sheets — exactly, across all five labels:

| Wallet Classification | Rows |
|---|---|
| DEBIT CARD WALLET | 1,174 |
| CREDIT CARD | 1,048 |
| DEBIT CARD | 958 |
| CREDIT CARD WALLET | 267 |
| HIMYAN | 260 |

1,441 of 3,707 transactions match a prefix (38.9%); every row receives a label. Both the Process Raw
Details export and the consolidated payout workbook carry the same distribution.

Also tested: the sample-data path (3 match / 3 no-match, with 9-digit sample prefixes proving the
first-6 truncation); the unconfigured path (old Masterlist — graceful, marker restored, clear banner);
and a deliberately incomplete `Wallet_Rules` (both Himyan rows removed → 260 rows show **no rule** and
are named in the banner, none guessed).

Regression after the change: sorting and per-merchant numbering intact, transaction-type handling
intact (gain 16.00 / 19.00 fixtures unchanged), dashboard reconciliation still zero mismatches, and
1,775 / 1,775 exported formulas evaluate to their cached values.

## Classification (same Wallet_Rules row)

**Audit status: PENDING VERIFICATION.**

### The rule

Classification is resolved from two further columns on the **same `Wallet_Rules` row** that already
drives Wallet Classification — same key (`Card Type (On us/Off us)`), same prefix-match outcome:

| Column | Meaning |
|---|---|
| `Classification (Match)` | value when the card's first 6 digits are in the Prefix list |
| `Classification (No Match)` | value when they are not |

**A blank cell means "carry the Wallet Classification value through unchanged" — not "empty".** That
is what keeps `DEBIT CARD WALLET` intact (its rows are deliberately left blank) and what makes a newly
added card type work with nothing filled in, as specified. Header spellings
`Classification (Prefix Match)` and `Classification (Match)` are both accepted.

Net effect on the supplied data: the 267 credit-card-wallet rows fold back into `CREDIT CARD`;
everything else is a straight copy of Wallet Classification.

| Wallet Classification | → Classification | Rows |
|---|---|---|
| DEBIT CARD WALLET | DEBIT CARD WALLET | 1,174 |
| CREDIT CARD | CREDIT CARD | 1,048 |
| DEBIT CARD | DEBIT CARD | 958 |
| CREDIT CARD WALLET | **CREDIT CARD** | 267 |
| HIMYAN | HIMYAN | 260 |

Resulting Classification: CREDIT CARD 1,315 · DEBIT CARD WALLET 1,174 · DEBIT CARD 958 · HIMYAN 260
= 3,707 of 3,707, with zero blanks.

### Why it lives on the Wallet_Rules row rather than its own sheet

The rule was first described as a mapping keyed on the *Wallet Classification value* (a two-row
exception list). Keyed that way it depends on the exact text of the wallet labels, so renaming a label
would silently produce a wrong answer. Keying on `Card Type (On us/Off us)` + prefix match instead
costs a few repeated cells but cannot break that way, keeps both columns independently editable, and
avoids one column cascading into the other when configuration is incomplete.

### Behaviour when the columns are absent

The Classification columns are optional and detected independently of the Wallet Classification ones.
Without them the column keeps its **rule pending** marker and stays blank, and the results banner says
which columns to add. Nothing is guessed.

### Verification

Against the supplied Masterlist and the 3,707-row sample, the app's output matches a projection
computed independently from the sheet — **0 mismatches** — and equals the rule as stated in
conversation. Both the Process Raw Details export and the consolidated payout workbook carry the same
distribution with **0 blank** Classification cells.

Also tested: the sample-data path (which exercises the blank carry-through — `JCB-ON US` on a prefix
match keeps `DEBIT CARD WALLET`, while the credit rows collapse to `CREDIT CARD`), and the previous
Masterlist whose `Wallet_Rules` has no Classification columns (marker restored, banner names the
columns to add, nothing invented).

Regression after the change: sorting and per-merchant numbering intact; transaction-type fixtures
unchanged (gain 16.00 / 19.00); dashboard reconciliation zero mismatches with all five checks still
firing independently under a negative test; 1,775 / 1,775 exported formulas evaluate to their cached
values.

## App shell: icon sidebar and three sections

**Audit status: PENDING VERIFICATION.** Presentation only — no calculation, lookup or export logic
was touched.

The single scrolling page is now an app shell: a persistent icon sidebar on the left and a content
pane that swaps between three sections.

| Section | State |
|---|---|
| **Dashboard** | Live. The payment-category cards moved here from inside the payout report. |
| **Payout Calculator** | Live. Uploads, pipeline, Master List validation, Process Raw Details, payout report and Excel export — unchanged. |
| **Reconciliation** | Placeholder. Documents the planned payout-vs-portal matching; nothing is computed. |

### What moved, and what deliberately did not

The dashboard cards render into the Dashboard section, but they are still produced by the **same call
inside `renderPayoutReport`** — `renderDashboardView(report)` is invoked from there with the report
that was just computed. One calculation, rendered in one place; the cards cannot drift from the table.
Switching to the Dashboard re-renders from `state.payoutReport` rather than recomputing.

The payout report keeps its own four stat tiles (Report Lines, Total Noqoody Profit, Total Merchant
Payout, Total Transfer), so the calculator still carries a summary of its own.

`Portal Amount` stays an editable column in the payout report feeding `Difference`, and the Excel
format is untouched. When Reconciliation is built it should become a better way to populate and
review those same values, not a replacement for them.

### Supporting pieces

- **Data-status strip** above the content pane — Master List sheets, frequency, OMS rows, processed
  rows, payout report lines — so Dashboard and Reconciliation are never mysteriously empty. Refreshed
  on upload, on processing and on report generation.
- **Empty states that lead somewhere.** The Dashboard with no processed data, and the Reconciliation
  placeholder, both carry a button that switches to the Payout Calculator.
- The large page header was removed: the sidebar brand and each section's own title already named the
  app and the section, so it only duplicated whichever section was open.

### Verification

Section switching, the in-page section links, and the status strip through a full run were tested in
the browser. Regression after the change: Process Raw Details 22 columns with **0** blank Wallet
Classification or Classification cells across 3,707 rows; payout report table and its four stat tiles
intact; dashboard-vs-report reconciliation still 0 mismatched categories with all five checks firing
independently under the negative test; **1,775 / 1,775** exported formulas evaluate to their cached
values; no horizontal overflow at 1440px or 430px; no console errors.

## Visual pass: fixed rail, sticky columns, reference palette

**Audit status: PENDING VERIFICATION.** CSS and markup only — no calculation, lookup or export code
was touched.

### Scroll behaviour — what was measured and fixed

Before: the sidebar was a 353px card in an 820px viewport. It stuck correctly, but left a large empty
gap below it, and the data-status strip scrolled entirely out of view (measured at `top: -2193` at
scroll 2229) — so deep in a long report there was nothing left saying which batch was on screen.

| | Before | After |
|---|---|---|
| Sidebar | `sticky`, 353px card, dead space below | `fixed`, full viewport height at every scroll position |
| Status strip | Scrolled away (`top: -2193`) | Sticky, `top: 18` at every scroll position |
| Workflow rail (steps 1–3) | Scrolled away, leaving an empty column beside the report | Sticky below the strip, scrolls inside itself when taller than the space |
| Section title | Scrolled away | Still scrolls — no longer a problem, since the fixed rail always shows the active section |

The workflow rail being sticky is a usability gain as well as a layout one: frequency and the uploaded
files stay reachable without scrolling back to the top of a 3,000px page.

### Palette

Taken from the supplied reference design, and **confined to the sidebar**:

    --rail:#252845  --rail-2:#31355A  --rail-line:#3A3E63
    --rail-ink:#EDEFF8  --rail-muted:#A7ADCB
    --lime:#D6F35E  --lime-ink:#1E2A06

A soft dark navy rather than black, with lime for the active nav item and the brand mark.

**Lime is deliberately never used in the content area.** It is close enough to the `--ok` green that a
lime element among the data could be read as "this passed" when it is only decoration — and on a
screen where a failed reconciliation must be unmistakable, that is not an acceptable ambiguity. The
content keeps its existing status colours unchanged, so ok / warning / danger stay meaningful.

The rail tokens are intentionally identical in light and dark mode, as in the reference, where the
sidebar is dark regardless of theme.

### Also

- Content is now full-bleed beside the rail instead of a centred 1320px column, so wide screens are
  actually used.
- Panels 14px → 16px radius, dashboard cards 12px → 14px, slightly more padding.
- Under 900px the rail becomes a horizontal bar and everything reverts to a single column.

### Verification

Sidebar position and strip position were measured at four scroll depths (top, 1200, 3000, bottom) —
`fixed`, full height, and `top: 18` throughout. Regression: sorting and per-merchant numbering intact;
transaction-type fixtures unchanged (gain 16.00 / 19.00); Wallet Classification and Classification
distributions unchanged; dashboard still reconciles with the report on every category with zero
mismatches; **1,775 / 1,775** exported formulas evaluate to their cached values; no horizontal
overflow at 1440px or 430px; no console errors.

## Purple palette, theme toggle, merchant search, category bars

**Audit status: PENDING VERIFICATION.** Presentation and one display-only filter — no calculation,
lookup or export logic was touched.

### Palette — whole app, both themes

The accent moves from green to purple, **specifically so green can mean one thing only: "ok"**.
Previously the brand colour and the success colour were both green, which is exactly the ambiguity you
do not want on a screen where a failed reconciliation must be unmistakable.

| | Light | Dark |
|---|---|---|
| bg / surface | `#F4F5FA` / `#FFFFFF` | `#14151C` / `#1C1E27` |
| ink / muted | `#1A1B2E` / `#6B6E85` | `#ECEDF3` / `#A2A5B8` |
| accent / soft | `#6341E8` / `#EDE9FE` | `#A794FF` / `#2A2450` |
| ok · warn · danger | `#15803D` · `#B45309` · `#B91C1C` | `#5BE38B` · `#FBBF24` · `#FB8A8A` |

Sidebar stays dark in both themes (`#1E2030`), with the active item now the accent purple rather than
lime — one brand colour rather than two.

**Every pair was checked programmatically for WCAG AA (>= 4.5:1)**: body text, muted text, accent text,
the accent button label, and each status chip on its own soft background, in both themes, plus the
sidebar. 26 pairs, **0 failures**. The first attempt failed on accent-on-accent-soft at 4.39:1 and the
accent was darkened from `#6C4DF6` to `#6341E8`, which brings it to 5.12:1.

### Theme toggle

Three states — Light / Dark / Match system — in the header strip, remembered in `localStorage`.
`system` removes the `data-theme` attribute and lets `prefers-color-scheme` decide.

A pre-existing `:root[data-theme="dark"]` block carrying the **old green** palette was sitting after
the new one and winning on source order, so forced dark silently kept the old colours. Removed.

### Merchant search

Filters the payout report rows by merchant name or MID, live. **It is display-only**: section labels,
subtotals, the grand total and every exported figure always cover every merchant. A banner states this
whenever a filter is active, so a filtered view can never be mistaken for a smaller payout run.

### Payout share bars

Nine payment categories as horizontal bars scaled to the largest, above the existing cards. Nine cards
cannot be compared against each other at a glance; bars can, while the cards keep the exact figures.

A donut was considered and rejected: nine categories with several at zero, and amounts that must be
compared exactly (`35,820.41` vs `27,317.66`), is not what a donut communicates.

### Verification

Theme toggle across all three states with the accent and background token values read back from the
DOM; search filtered to 2 rows on "MADHURA", 1 by MID, 0 on no match, 31 cleared, with subtotals and
the grand total visible throughout; bar widths 100% / 76.3% / 9.4% matching the underlying amounts.
Regression: sorting, numbering, transaction types, both classification columns, dashboard
reconciliation (0 mismatched categories), sidebar and strip still fixed at every scroll depth,
**1,775 / 1,775** exported formulas, no overflow at 1440px or 430px, no console errors.

One bug found and fixed during the pass: the bar track and fill are `<span>`s, so `width` and `height`
were being ignored until they were given `display:block`.

## Wallet rows: section grouping moves to Wallet Classification

**Audit status: PENDING VERIFICATION.** This changes payout amounts — see the impact below.

### The rule

A transaction's report section now comes from its **Wallet Classification**, not Card Type 2, because
a wallet transaction can be priced differently from the same card outside a wallet. Classification was
considered and rejected: it folds `CREDIT CARD WALLET` into `CREDIT CARD`, which would bill those
transactions at the credit rate.

**The split is by ROW, not by column.** Three extra column sections would have pushed the report past
130 columns of horizontal scrolling. Instead each section header becomes `CREDIT CARD / WALLET` and
holds both, with the row saying which.

Wallet labels follow the convention `<base> WALLET`. Stripping that suffix gives the `Card_Types`
section the row belongs to, and its presence marks the row as a wallet row.

### Row model

A merchant now produces up to **four** rows — the wallet split happens first (it decides the rate),
then the Below Amount split inside each. Empty combinations are skipped:

    MADHURA RESTAURANT | BELOW 25          <- card, below   (no card transactions >= 25)
    MADHURA RESTAURANT | WALLET
    MADHURA RESTAURANT | WALLET | BELOW 25
    NEW BALANCE DOHA                       <- one row only: no wallet, no below

`No.` and `MID` print **only on the merchant's first row**; continuation rows leave both blank so the
block reads as one merchant. `No.` advances once per merchant (on `isPrimary`, not on `!isBelow` —
otherwise wallet main rows would each consume a number).

**Gain For Bank Charges and Payout Processing Fee apply once per merchant, on its first row** —
whichever that is, so a merchant with only wallet activity still receives them. `Total Transfer` sums
every row of the merchant and sits on that same first row.

### Rates

| Row type | Merchant rate | Bank cost |
|---|---|---|
| Card | `Card_Types.Master List Rate Column` (per scheme) | `Card_Types.Cost Table Column` (per scheme) |
| Wallet | `Wallet MerchantRate` (one shared column) | `Wallet` (one shared column) |

One wallet rate and one wallet cost serve every wallet type, whatever card is underneath. Column names
are resolved by candidate list, so `Wallet Rate` / `Wallet MerchantRate` both work.

### Impact on the supplied data

Grouping on Wallet Classification moves 490 Sale transactions (QAR 21,663 gross) onto the wallet rate.
Total Noqoody Profit moves from **−208.96 to −21.31** on the Daily batch, and merchant Internal
Transfer falls correspondingly. This is a pricing change, not a cosmetic one.

### Verification

**98 of 98 rendered rate pairs** match the Masterlist — card rows on their per-scheme columns, wallet
rows on the shared wallet columns — checked merchant by merchant, section by section.

*(That check first reported 3 mismatches, which turned out to be the test keying on merchant name:
two different merchants are both called `NEW BALANCE DOHA`, with MIDs `100005269900113` and
`777100432320320` and Cost IDs 1 and 7. The app keys on MID + Merchant Key and was correct.)*

Dashboard reconciliation across all 55 rows: every category and all three headline figures agree with
the independently re-summed rows **and** the grand-total row, 0 mismatches. Export: **2,903 / 2,903**
formulas evaluate to their cached values; section banners read `CREDIT CARD / WALLET`; lead columns
blank on continuation rows; `Total Transfer` spans the whole merchant block (`SUM(DF4:DF6)` for a
three-row merchant, `SUM(DF7:DF8)` for a two-row one).

Two bugs were found and fixed during the pass: the merchant filter skipped wallet rows entirely, and
searching by MID matched only a merchant's first row, since the MID cell is now blank on continuation
rows — the MID is carried on every row in a data attribute for that reason.

### Still open after this update

- The dashboard still groups by `Card Type 2`; `DASHBOARD_REFERENCE_FIELD` can be switched to
  `Classification` now that the column is populated, but that has not been requested or tested.
- A *future* non-debit wallet label (e.g. `AMEX WALLET`) has no stated rule. Under the blank default
  it would carry through unchanged rather than collapsing the way credit and Himyan do; set that row's
  `Classification (Match)` cell explicitly when such a card type is added.
- `Special_Rate` (Account Code 5) overrides — still unimplemented.
- **Reconciliation** is a placeholder. Before building it: how do portal amounts arrive (an export
  file, or typed per merchant)? What identifies a merchant in the portal — MID alone cannot separate
  shared-MID merchants? Compare against `Total Transfer`? What match tolerance?
- Production verification of the dashboard totals and the duplicate report against a real payout run.

---

## Fix: comma-formatted amounts were silently parsed as zero

**Audit status: PENDING VERIFICATION.** The root cause of the "report doesn't match the raw sheet"
mismatch, found while verifying an exported `Payout_Report_Daily` against its own
`Process Raw Details` sheet.

### The defect

OMS raw exports store `Gross Amount`, `Commission` and `Net Amount` as **text**, and Excel writes a
thousands separator into that text — so any transaction of 1,000 or more arrives as `"2,500.00"`.
`toNumber()` parsed every amount with `Number(v)`, which returns `NaN` for a string containing a
comma, and the `isNaN` guard then turned it into **0**.

Transactions under 1,000 parsed correctly; every transaction at or above 1,000 was counted as zero.
Measured on `OMS_NapsTab` (3,707 rows):

| column | cells with a comma | value silently zeroed | of column total |
|---|---|---|---|
| Gross Amount | 138 | 1,067,735.70 | 1,245,071.15 |
| Net Amount | 137 | 1,054,778.25 | 1,231,035.16 |

86% of the gross was being discarded. The on-screen Process Raw Details table looked correct
throughout, because `fmtAmount()` fell back to `String(v)` and re-displayed the original text.

### The fix

- `toNumber()` strips thousands separators before parsing, and returns numeric input unchanged.
- `fmtAmount()` does the same, so display and calculation agree.
- `buildProcessRawDetailsExportRows()` writes the three amount columns as real numbers with a
  `#,##0.00` format, via a shared `buildProcessRawDetailsSheet()` used by both export paths.
  Excel's `SUM` previously returned 0 on that sheet.
- `Merchant ID`, `Terminal ID` and `Reference Number` **deliberately stay text** — they are
  identifiers, and a numeric-typed MID renders in scientific notation.

Safety checked before applying: after stripping commas, zero values in the OMS file are unparseable;
negatives use a leading minus (23 rows, no parenthesised accounting negatives, no currency symbols);
and the Masterlist contains no comma-bearing numeric strings, so rates and thresholds are unaffected.

### Verification

Report totals, before vs after, same inputs:

| | before | after |
|---|---|---|
| Credit Card section gross | 27,876.33 | 302,370.43 |
| Grand total gross | 66,359.86 | 469,630.10 |
| Total profit | -4,467.75 | -165.81 |

The exported `Process Raw Details` now types all 3,707 amount cells as numbers: `Gross Amount` sums
to **1,245,071.15**, matching an independent parse of the source file. Both export paths (standalone
and bundled-with-report) were captured and checked.

Cross-footed against the raw sheet for `NEW BALANCE DOHA` / MID `777100432320320` — the case that
surfaced the bug: Credit Card **17 Sale rows, 9,040.50** and Credit Card Wallet **1 Sale row,
337.00**, both exactly as rendered, counts included.

Regression: 107/107 rate pairs match the Masterlist, 1,726/1,726 export formulas evaluate to their
cached values, row shape unchanged at 55 rows across 30 merchants.

### Correction to earlier figures in this document

Every figure previously quoted from a `Process Raw Details` sheet was produced by verification
scripts that used the same `Number()` parse, so they are **understated**. This includes the
"758 unmatched transactions / QAR 79,007" finding and the shared-MID per-merchant breakdown; both
need recomputing before they are relied on.

---

## Fix: shared-MID transactions were credited to the wrong merchant

**Audit status: PENDING VERIFICATION.** Found from a report showing BURQAN TRADING under
`Account Code (unassigned)` with 136 transactions and 10,530.50, when BURQAN has exactly one
terminal and one transaction of 30.00.

### Defect 1 — the MID-only fallback stole other merchants' transactions

`resolveMerchant()` fell back to matching on MID alone when a transaction's `MID + Terminal ID`
was not in Terminal_Mapping:

```js
if (!ml && !terminalMatched){
  ml = masterListIndex.get(norm(row.merchantId) + "|");   // MID alone
}
```

MID `777101639900128` is a shared NOQOODY aggregator MID carrying **461 transactions across 22
terminals belonging to different merchants**. Only BURQAN's terminal (`10048839`) is mapped. For the
other 21 terminals the fallback matched on MID alone and landed on the blank-Merchant-Key
Master_List row — BURQAN — so **460 transactions belonging to other merchants were credited to
BURQAN**, at BURQAN's rates.

The fallback's stated purpose was to resolve single-terminal merchants that have no Terminal_Mapping
row. Measured against the real data, that case accounts for **0 transactions**: every use of the
fallback was the defect. It is removed. Matching is now strictly
`MID + Terminal ID → Terminal_Mapping → Merchant Key → Master_List`, and an unmapped terminal is
reported rather than guessed.

### Defect 2 — the account code was taken from whichever row sorted first

```js
accountCode: row._accountCode && String(row._accountCode).trim() !== "" ? row._accountCode : "(unassigned)",
```

The group key is `MID | Merchant Key`, and BURQAN's Merchant Key is blank — the same as the rows the
fallback attached. Its one correctly mapped transaction and the 460 stolen ones collapsed into a
single group, and the account code came from the first row in file order:

```
file row 340   TID 10017805   unmapped   AC ""   <- first, so the group became "(unassigned)"
BURQAN's mapped row is #333 of 461
```

The tag was read correctly and then discarded. The account code is now taken from a terminal-matched
row, and a merchant whose terminals disagree is labelled `(conflicting: …)` rather than silently
placed under one of them. No group in the current Terminal_Mapping spans more than one account code.

### Unmatched transactions are now flagged by reason

The two gaps need different actions, so they are counted and listed separately — on the summary
tiles, in a banner naming the exact terminals and MIDs to add, on each row's status chip, and in the
export's `Match Status` column:

- `Terminal not in Terminal_Mapping` — the MID is known, the terminal is not
- `MID not in Master_List` — the merchant is not configured at all

### Verification

BURQAN now reports under **Account Code 9**, 1 transaction, **30.00**.

The run reconciles to the source file exactly:

| | txns | gross |
|---|---|---|
| Matched | 2,845 | 1,137,354.15 |
| Terminal not in Terminal_Mapping | 776 | 47,464.50 |
| MID not in Master_List | 86 | 60,252.50 |
| **total** | **3,707** | **1,245,071.15** |

That total is the independently parsed total of the OMS file — nothing is lost or double-counted.
Account Code sections in the export: 4, 9, 12, 13, 15, 16; `(unassigned)` no longer appears.

Regression: 107/107 rate pairs match the Masterlist, row shape and transaction-type handling
unchanged, no page errors.

### Still open

- 41 terminals (776 transactions, QAR 47,464.50) need Terminal_Mapping rows.
- 5 MIDs (86 transactions, QAR 60,252.50) are not in Master_List at all — including the Weekly
  merchants not yet added.

---

## Per-transaction fees: the Txn_Fee tier sheet

**Audit status: PENDING VERIFICATION.** Adds per-card-type per-transaction pricing
(`1.90% + 0.50`, `1.35% + 1.00`, …), which the single merchant-level `Rate Per Trnx` column could
not express.

### Why a lookup sheet rather than columns on Master_List

Per-card-type columns would mean 7 card types x 269 merchants = **1,883 cells**, growing by 269 with
every new card type — which defeats the point of Card_Types being dynamic. Bank_Cost already solves
this shape: 7 rows serve 75 merchants, each named by a real contract. `Txn_Fee` follows it exactly.

### Schema

**`Txn_Fee`** — keyed by `Txn Fee ID`, one row per pricing tier:

| Txn Fee ID | Credit Card | Debit Card | Wallet | Himyan | AMEX | JCB | China UnionPay | Scope | Note |
|---|---|---|---|---|---|---|---|---|---|

- Amounts are **currency, formatted `0.00`** — deliberately not the `0.00%` Bank_Cost uses.
- `Scope` = `All` (every cleared Sale) or `Below` (only Sales under the merchant's Below Amount
  Rule, which is what the legacy column did). **Blank means `All`**: a blank cell should read as the
  plain meaning of a contract, not as a special case.
- Wallet rows read the general `Wallet` column, matching the one general wallet rate and wallet bank
  cost.

**`Master_List`** — one new column, `Txn Fee ID`. **`Card_Types`** — one new column,
`Txn Fee Column`, parallel to `Cost Table Column`, so a new card type stays a data-only change.

### The legacy column stays, with a retirement signal

`Rate Per Trnx` still applies when `Txn Fee ID` is blank. With 269 merchants, dropping it outright
would turn every unfilled row's fee silently into zero — the failure mode this codebase has twice
been bitten by. The report now states how many merchants still depend on it, so it can be removed on
evidence once that count reaches zero rather than living on as a second source of truth.

A `Txn Fee ID` pointing at a row that does not exist is **named in a red banner**, not quietly
dropped to the fallback.

### Verification

Migrating the three legacy merchants to tier 2 (`0.50` across the board, `Scope = Below`) reproduces
their figures exactly. Every section carrying transactions is identical:

```
Credit Card          cleared 19   rate 0.50 -> 0.50   charge 9.50 -> 9.50
Debit Card           cleared 15   rate 0.50 -> 0.50   charge 7.50 -> 7.50
Credit Card Wallet   cleared  4   rate 0.50 -> 0.50   charge 2.00 -> 2.00
```

**One behaviour change**, in sections with no transactions: the legacy column blanket-applied its
rate to *every* payment group, so AMEX/JCB/China UnionPay displayed `0.50`; tier 2 sets them to `0`
and they now display `0.00`. No money moves today because those sections are empty, but a future
AMEX transaction would be charged `0` rather than `0.50` — set the tier's AMEX cell if that is
wrong.

Tier 1 (`Scope = All`) resolves per card type and applies on the **main** line, not only below the
threshold:

```
NEW BALANCE DOHA   Credit Card  17 x 0.50 = 8.50   Debit 6 x 1.00 = 6.00   Himyan 3 x 1.00 = 3.00
                   wallets       1 x 0.50 = 0.50           5 x 0.50 = 2.50
```

Regression: 107/107 rate pairs, 1,726/1,726 export formulas, row shape unchanged, no page errors,
and the control total still balances at 3,707 / 1,245,071.15.

### Still open

- No merchant is assigned `Txn Fee ID 1` yet — the new pricing has a tier but no holder.
- AMEX / JCB / China UnionPay are `0` on both tiers; the stated contracts named only Credit, Debit,
  Wallet and Himyan.

---

## Fix: wallet bank cost follows the card scheme, not a general Wallet column

**Audit status: PENDING VERIFICATION.** Raised from a discrepancy in Actual Bank Charges.

### The defect

`computePaymentGroupCell` read a wallet row's bank rate from a general `Wallet` column in Bank_Cost:

```js
bankRate = toRate(pickByCandidates(bankCostRow, WALLET_COST_CANDIDATES));
```

In all seven Bank_Cost rows that column held **the same value as `Debit Card`**, and differed from
`Credit Card` in every one. Worse, the single value was applied to *every* wallet section — so a
Himyan Wallet or AMEX Wallet was also costed at the debit rate.

The bank does not price that way. It charges by **card scheme**, wallet or not: a Credit Card Wallet
transaction costs the Credit Card rate.

### Evidence — the OMS Commission field is the real charge

Wallet Sales in one Daily batch:

| model | modelled | vs actual (1,631.90) |
|---|---|---|
| app before — Bank_Cost `Wallet` column | 1,359.18 | **−272.72** |
| card scheme column | 1,585.36 | −46.54 |

Per Cost ID, the card-scheme rate reconciles **to the cent**:

```
Cost ID / group          actual   implied   Bank_Cost   gap
2  DEBIT CARD WALLET       5.60    0.700%    0.700%    0.00
3  CREDIT CARD WALLET    282.11    1.100%    1.100%    0.00
3  DEBIT CARD WALLET       7.82    0.600%    0.600%    0.00
4  DEBIT CARD WALLET       1.12    0.602%    0.600%    0.00
7  CREDIT CARD WALLET      7.70    1.851%    1.850%    0.00
```

### The change

A wallet row now takes the bank rate of the card scheme underneath it, via the section's existing
`Cost Table Column` — the same lookup card rows already use. This is automatic for any future card
type and needs no new columns.

The **merchant** rate is unchanged: wallet rows still read one general `Wallet MerchantRate`. That
is Noqoody's own pricing and has nothing to do with the bank's cost.

The general `Wallet` column has been **deleted from Bank_Cost** (Option 1 of two considered; the
alternative was keeping it as an explicit per-wallet override). It was removed because the value is
derived, not data, and a duplicate that can drift is a liability. Removed values, for the record:

```
Cost ID 1 0.0075 · 2 0.0070 · 3 0.0060 · 4 0.0060 · 5 0.0060 · 6 0.0070 · 7 0.0050
```

A workbook that still carries the column is **flagged in the report** as ignored, so nobody edits a
cell that no longer affects anything.

### Verification

`NEW BALANCE DOHA | WALLET`, before and after:

```
                       before   after
CREDIT CARD WALLET      0.50%   1.85%
DEBIT CARD WALLET       0.50%   0.50%
HIMYAN                  0.50%   0.85%
AMEX                    0.50%   2.75%
CHINA UNIONPAY          0.50%   2.50%
JCB                     0.50%   2.50%
```

Regression: 107/107 rate pairs, 1,726/1,726 export formulas, row shape unchanged, control total
still 3,707 / 1,245,071.15. The rate test's expectation was updated — it encoded the old
Wallet-column rule and was stale, not failing.

### Separate finding, not fixed: Cost ID 1 looks stale

The residual sits entirely on Cost ID 1, and it is **not** wallet-specific — card and wallet rows
agree with each other, which confirms the rule, but both sit above the stored rate:

```
1 | CREDIT CARD (card)    432 txns   implied 1.821%   Bank_Cost 1.750%
1 | CREDIT CARD (wallet)   96 txns   implied 1.852%   Bank_Cost 1.750%
1 | DEBIT CARD  (card)    377 txns   implied 0.773%   Bank_Cost 0.750%
1 | DEBIT CARD  (wallet)  492 txns   implied 0.775%   Bank_Cost 0.750%
1 | HIMYAN      (card)    116 txns   implied 0.799%   Bank_Cost 0.800%   OK
```

Cost ID 1's Credit Card looks like it should be **1.85%** and Debit around **0.775%**. Himyan is
correct, so it is not a blanket drift. This is a question for the bank, not a code change, and
nothing was altered.

---

## Bank Difference flagging, and merchant drill-down

**Audit status: PENDING VERIFICATION.**

### 1. Bank Difference is flagged where it happens, and names its cause

Every payment-group cell now carries `actualBankFee` — what the bank actually charged *that section*,
from the OMS Commission field — alongside the modelled `bankFee`. A merchant line whose
`|Bank Difference|` exceeds **1.00** (`BANK_DIFF_FLAG_THRESHOLD`) is marked in the Bank Difference
cell, and its tooltip names the sections responsible. Sections contributing under 0.10
(`BANK_DIFF_SECTION_MIN`) are rounding, not a cause, and are left out.

A summary banner groups the flagged lines **by cause**, because the same section drifting across
several merchants points at a Bank_Cost cell rather than at the merchants:

```
9 merchant lines differ from the bank by more than 1.00 — -262.72 in total.
 • Credit Card        -183.37 across 6 merchants.
   Bank_Cost says 1.75%, the bank charged 1.90% on 122,237.00 gross.
   Same section across several merchants — check the Bank_Cost cell, not the merchants.
 • Debit Card WALLET   -36.25 across 3 merchants.  Bank_Cost 0.75%, bank charged 0.85%
 • Debit Card          -33.80 across 3 merchants.  Bank_Cost 0.75%, bank charged 0.85%
 • Credit Card WALLET   -9.30 across 1 merchant.   Bank_Cost 1.75%, bank charged 1.90%
```

On the first run this independently reproduced the Cost ID 1 staleness recorded in the previous
section — the feature found the known defect without being told about it.

### 2. Drill-down: double-click a merchant row

Opens that merchant's own transactions, the payout report's equivalent of Excel's PivotTable
**Show Details**. Rows are keyed the way the report groups them (**MID + Merchant Key**), so a shared
aggregator MID never pulls in another merchant's transactions — verified on
`777100432320320`, which carries ten merchants: the panel returned NEW BALANCE DOHA's 33 rows and
nothing else.

The panel shows Transactions / Cleared (Sale) / Gross (Sale) / Bank Commission, the full Process Raw
Details columns, closes on Escape or backdrop click, and **Download these rows** writes a workbook
whose sheet is named after the merchant, with amounts as real numbers (identifiers stay text).

### A bug this introduced, and caught

Adding the `drillable` class broke `applyMerchantFilter`, which tested `tr.className !== ""` by
string equality — every data row would have been skipped and the merchant search would have matched
nothing. It now tests `classList.contains("drillable")`. Several Playwright tests had the same
brittleness (`tr.className === 'wallet-row'`) and were updated to `classList.contains`; that was a
test defect, not an app one.

### Verification

Merchant search: 72 rows unfiltered, 4 for "NEW BALANCE", 5 by MID `777100432320320`, 72 when
cleared. Regression: 107/107 rate pairs, 1,726/1,726 export formulas, row shape unchanged, control
total still 3,707 / 1,245,071.15, no page errors.

---

## Excel drill-down: hyperlinked merchant names

**Audit status: PENDING VERIFICATION.** The in-app drill-down panel was not what was asked for —
the drill-down needed to work **inside the exported workbook**, in Excel.

### Why it is a hyperlink and not a double-click

A true double-click "Show Details" is an Excel *application* behaviour. It exists in exactly two
places: on a real PivotTable, or behind a VBA macro. The payout report is a custom cross-tab (merged
section banners, per-account subtotals, a reconciliation block) and cannot be a PivotTable without
losing that layout; a macro would force `.xlsm`, an "Enable Content" prompt on every machine, and
would not run in Excel Online or on mobile.

A hyperlink needs no macro and works in Excel desktop, Mac, Online and mobile alike. Chosen over a
PivotTable sheet and over `.xlsm` deliberately.

### What the export now contains

A third sheet, **Merchant Details** — every matched transaction grouped by the merchant the payout
report groups it under (**MID + Merchant Key**), one block per merchant, each opened by a banner row
naming the merchant, MID, Merchant Key, Account Code and transaction count. Blocks follow the
report's own order.

On the report sheet, each merchant's **Merchant** cell is an internal hyperlink to its banner row,
styled as a link. Only the merchant's first row carries it; continuation rows are the same merchant.

**Process Raw Details** gains an AutoFilter and a frozen header — the flat, sliceable view.
Merchant Details cannot carry an AutoFilter because it has a header row per block, so the two sheets
serve different purposes deliberately.

### Verification

```
sheets                      : DAILY PAYOUT | Merchant Details | Process Raw Details
hyperlinks on report sheet  : 41   (one per merchant, invalid targets: 0)
Merchant Details            : 41 blocks, 41 header rows, 1,620 data rows
Process Raw Details filter  : A1:V3708
```

Every link resolves to the right banner:

```
NEW BALANCE DOHA  -> row 53: NEW BALANCE DOHA — MID 777100432320320 · Merchant Key 2 · Account Code 4 · 33 transactions
```

The shared-MID case holds: that MID carries ten merchants and 355 transactions; the block for
NEW BALANCE DOHA contains `{"NEW BALANCE DOHA": 33}` and nothing else.

OOXML validated directly, since LibreOffice cannot load these exports at all (it fails identically
on a pre-change export, so it is not a regression): all 12 parts well-formed; `sheet3.xml` declared
in `workbook.xml`, `workbook.xml.rels` and `[Content_Types].xml`; element order in the report sheet
is schema-correct (`sheetData -> mergeCells -> hyperlinks -> ignoredErrors`); internal links carry a
`location` and need no relationship part. Amounts numeric, identifiers text.

Regression: 107/107 rate pairs, 1,726/1,726 export formulas, row shape unchanged, control total
3,707 / 1,245,071.15, in-app drill-down still returns 33 rows for the same merchant.

---

## Merchant Details: flat layout with a working AutoFilter, unstyled links

**Audit status: PENDING VERIFICATION.** Revises the previous section.

### Correcting the previous entry

That entry said Merchant Details "cannot carry an AutoFilter." That was misleading. Excel allows one
AutoFilter per sheet and it must sit on a single header row — so the limitation was in the
**banner-grouped layout chosen**, not in Excel. Rebuilt flat, the sheet does both.

### What changed

`buildMerchantDetailsSheet` now emits **one header row and nothing but data beneath it**, ordered by
the merchant the report groups it under (MID + Merchant Key), in the report's own order. Per-merchant
banner rows are gone. The sheet carries an AutoFilter across its full range and a frozen header row.

`Account Code` leads the columns — it is how the report is organised, so it is the first thing worth
filtering on. The remaining columns mirror Process Raw Details exactly.

The report's Merchant cell links to that merchant's **first data row** instead of a banner.

### Links no longer look like links

The hyperlink carries no styling: the merchant name keeps the report's own formatting rather than
turning blue and underlined. The hover tooltip ("Show this merchant's transactions") is what
advertises it. The link itself is unaffected — styling and behaviour are independent in OOXML.

### Frozen header rows needed raw XML

`ws["!freeze"]` is inert in this library — it was written and silently dropped, which the export
check caught. `patchWorksheetForExcel` now takes a list of sheet paths whose top row to freeze and
rewrites their `sheetView` directly, alongside the column freeze it already applied to the report.
`sheetViews` is replaced in place so it stays ahead of `cols`/`sheetData`.

### Verification

```
sheet1  DAILY PAYOUT          pane xSplit="3"  (columns)      autoFilter: none
sheet2  Merchant Details      pane ySplit="1"  (header row)    autoFilter: A1:W1621
sheet3  Process Raw Details   pane ySplit="1"  (header row)    autoFilter: A1:V3708
```

Element order schema-correct on all three (`sheetViews` before `cols`/`sheetData`; `autoFilter`
after `sheetData`), all 12 parts well-formed.

```
hyperlinks                          : 41
styled blue/underline               : 0
links landing on the wrong merchant : 0
flat sheet                          : 1,620 data rows, 41 distinct merchants, 0 banner rows
```

Regression: 107/107 rate pairs, 1,726/1,726 export formulas, row shape unchanged, control total
3,707 / 1,245,071.15.

---

## Drill-down that actually filters: one sheet per merchant

**Audit status: PENDING VERIFICATION.** The hyperlink jumped to a row; it did not show only that
merchant's transactions, which is what a drill-down means.

### Why the previous approach could not filter

A hyperlink navigates. It cannot change a filter's state — OOXML stores one static `<autoFilter>`
per sheet, not a filter that responds to a click. So on a shared sheet, a link can only ever scroll
you to a row with every other merchant still around it.

The way to make a click produce a filtered view, without a macro, is to land on a sheet that
contains nothing else. That is also exactly what Excel's PivotTable **Show Details** produces: a new
sheet holding only the rows behind the figure. `buildMerchantSheets` prepares them up front.

### What the export now contains

```
DAILY PAYOUT | Merchant Details | Process Raw Details | <one sheet per merchant>
```

Each merchant sheet carries a title row naming merchant, MID, Merchant Key, Account Code and
transaction count — itself a **back-link to the payout report** — then a header row, then only that
merchant's transactions. Header frozen at row 2, AutoFilter on the header so the merchant's own rows
can still be sliced further.

Sheet names are sanitised to Excel's rules: 31 characters, none of `: \ / ? * [ ]`, and de-duplicated
with a numeric suffix. 11 merchant names in the sample exceed the cap and are truncated.

### Verification

```
total sheets                        : 44   (3 + 41 merchants)
sheet names unique                  : 44 of 44,  longest 31 chars
report hyperlinks                   : 41,  0 broken, 0 landing on a mixed sheet
hyperlinks workbook-wide            : 82   (41 forward + 41 back-links), 0 unresolvable
sheets with a frozen pane           : 44 of 44
sheets with an AutoFilter           : 43 of 44   (the report itself correctly has none)
```

Panes are right per sheet type:

```
DAILY PAYOUT           <pane xSplit="3" .../>   columns
Merchant Details       <pane ySplit="1" .../>   header row
merchant sheet         <pane ySplit="2" .../>   title + header
```

OOXML validated directly: 53 parts, 0 malformed; 44 sheets declared in `workbook.xml` with 44
worksheet content-type overrides and 0 missing parts; `sheetViews` ahead of `sheetData` on every
sheet.

`NEW BALANCE DOHA` — the shared-MID case, that MID carrying ten merchants — gets a sheet with its
**33 rows and no other merchant**, AutoFilter `A2:V35`.

File size 4.12 MB -> 7.63 MB for the extra sheets.

Regression: 107/107 rate pairs, 1,726/1,726 export formulas, row shape unchanged, control total
3,707 / 1,245,071.15.

---

## Reverted: Excel drill-down removed

**Audit status: PENDING VERIFICATION.** Three attempts at an Excel-side drill-down were removed at
the user's instruction, after establishing the requested behaviour is not expressible in `.xlsx`.

### The question, answered

> Clicking a merchant in the payout report should filter Process Raw Details to that merchant only.

**Not achievable in a plain `.xlsx`.** Not a tooling limit — the file format has no mechanism for it:

- A hyperlink does exactly two things: jump to a location, or open a URL. There is no "on click, run
  something" anywhere in a spreadsheet file.
- AutoFilter state is static XML — one stored state per sheet. It cannot vary by which cell was
  clicked, because nothing records that a cell was clicked.
- Defined names, data validation and conditional formatting do not respond to clicks either.

The only mechanisms that change a sheet's filter on click are a **VBA macro** (`.xlsm`, Enable
Content, blocked by most IT policy, dead in Excel Online and mobile) or an **Office Script /
Add-in**, which lives outside the file.

The per-sheet-per-merchant approach was the closest a plain workbook can get — it produced a
filtered *view* by landing on a sheet containing nothing else — but it cost 41 extra sheets and
grew the file from 4.12 MB to 7.63 MB. Rejected on both counts.

### What was removed

`buildMerchantDetailsSheet`, `buildMerchantSheets`, `safeSheetName`, the `Merchant Details` sheet,
the per-merchant sheets, and every hyperlink on the report's Merchant cells.

### What was kept

- **The in-app drill-down** — double-click a merchant row in the web app. Unaffected, and was never
  the thing in question.
- **Bank Difference flagging** with its per-section cause.
- **Process Raw Details: frozen header and AutoFilter.** Zero structural cost (4.12 MB -> 4.18 MB)
  and it is how a merchant's transactions are isolated by hand — filter the Merchant or Merchant ID
  column. Say so and it can go too.

### Verification

```
sheets                        : DAILY PAYOUT | Process Raw Details
hyperlinks left in the report : 0
file size                     : 7.63 MB -> 4.18 MB  (original 4.12 MB)
DAILY PAYOUT                  : <pane xSplit="3"/>            no filter
Process Raw Details           : <pane ySplit="1"/>            autoFilter A1:V3708
OOXML                         : 11 parts, all well-formed
```

Regression: 107/107 rate pairs, 1,726/1,726 export formulas, row shape unchanged, control total
3,707 / 1,245,071.15.

---

## Refund fee moves to Master_List; Dispute is no longer charged

**Audit status: PENDING VERIFICATION.**

### The rule, as confirmed

| Transaction type | Into the Refund column | Fee |
|---|---|---|
| Sale | no (it is Gross Collection) | no |
| **Refund** | ABS(Gross Amount), always | the merchant's **Refund Fee**, once per refund |
| **Dispute** | ABS(Gross Amount) **only when gross is negative** | **none** |
| **Reversal** | ABS(Gross Amount) only when gross is negative | none |

A **blank or zero** `Refund Fee` both mean the merchant is not charged for refunds. Only an explicit
amount charges.

### Three defects this fixed

1. **Dispute was charged as if it were a Refund.** `txnClass` mapped `refund` and `dispute` to one
   class, and the fee applied to both. In the sample batch that was **90.00 over-charged** across 18
   Disputes, with only 5.00 of the total legitimate.
2. **The fee was hard-coded.** `GAIN_PER_REFUND_DISPUTE = 5` applied to every merchant, so a merchant
   without a refund fee could not be represented at all. It now reads `Refund Fee` from Master_List
   per merchant, and the constant is gone — no hard-coded money value remains in the calculation.
3. **Dispute had no negative-gross guard.** Only Reversal was checked, so a positive-gross Dispute
   would have reduced a payout. Every Dispute in the sample is negative, so nothing moved today, but
   the guard now covers both.

### Verification

Against the uploaded Masterlist (`Refund Fee`: 59 merchants at 5, 4 at 0, 206 blank):

```
merchant                  freq     Ref  Disp   RefundFee  base   expected   exported
AL SULTAN MEDICAL CENTER  Daily      0     1        5.00  0.00       0.00       0.00
NEW BALANCE DOHA          Daily      1     0        5.00  4.00       9.00       9.00
CUP TIME TRADING          (none)     0    13        0.00  0.00          — excluded, blank frequency
ICE Q FOR ICE CREAM       (none)     0     4        0.00  0.00          — excluded, blank frequency
```

AL SULTAN previously showed **5.00** for a single Dispute and now correctly shows **0.00**.
NEW BALANCE DOHA is the one genuine refund: base 4.00 + one refund at 5.00 = **9.00**.

A stale `refundDisputeCount` reference survived the rename and was caught by the export test as a
page error before anything shipped.

Regression: 107/107 rate pairs, 1,726/1,726 export formulas, 30 merchants / 55 rows unchanged,
Refund column unchanged at 1,213.00, control total 3,707 / 1,245,071.15.

### Note on the uploaded Masterlist

Every Bank_Cost row now reconciles to what the bank actually charged — Bank Difference flags went
from 9 merchant lines / -262.72 to **zero**. The earlier "Cost ID 1 looks stale" finding was wrong
in its diagnosis: Cost ID 1 was correct, and the luxury merchants simply belonged on new cost rows
(8, 9, 10), which the user added.

### Business decision on record: Dispute is a NEW RULE

Confirmed by the user: not charging Disputes is a **new rule taking effect from now**, not a
correction of past behaviour. **Reports issued before this change stand as issued** — they are not
overstated, nothing needs regenerating, and no merchant is owed a refund of the fee.

The 90.00 charged across 18 Disputes in the sample batch was therefore correct under the rule in
force at the time. From this version onward, only Refunds carry a fee.

### Still open

- 12 active merchants have a blank `Refund Fee` — the newly onboarded luxury group on Cost IDs 8
  and 9. Under the confirmed rule they are not charged for refunds. Intentional or not yet filled in.
- Terminal_Mapping rows 293 and 295 remain an exact duplicate.
- 194 merchants still have a blank Payout Frequency and are excluded from every report.

---

## Profit no longer absorbs refunds

**Audit status: PENDING VERIFICATION.**

### The defect

```js
profit = noqoodyCharge + noqoodyChargePerTxn - bankFee - refund
```

A Refund is the cardholder's money going back, not a Noqoody expense, so deducting it from profit
made a section report a loss it never made:

```
NEW BALANCE DOHA | WALLET   bankFee 14.62   nqCharge 58.46   refund 329.00   profit -285.16
AL SULTAN MEDICAL CENTER    bankFee 30.69   nqCharge 40.92   refund 185.00   profit -174.77
```

58.46 of revenue against 14.62 of bank cost is not a 285.16 loss.

### The fix

```js
profit = noqoodyCharge + noqoodyChargePerTxn - bankFee
```

`Internal Transfer` is untouched and still deducts the refund, so the merchant is still paid net of
it — the refund leaves the payout exactly once, in the column that represents money paid out.

### A formula that was asked for and not implemented

The requested wording was **`Bank Fee + Noqoody Charge - Noqoody Charge/Txn`**. Implemented
literally that adds a cost as revenue and subtracts revenue as a cost. Measured before changing
anything:

```
sections with activity                 : 129
  of those, NO refund at all           : 127
  that the literal formula changes     : 127

whole-report total profit
  as computed before                   :   2,889.36
  refund removed                       :   3,403.36
  the literal wording                  :  23,270.66
  total Bank Fee across the report     :   9,962.65
  difference between the two options   :  19,867.30   ( = 2 x Bank Fee )
```

The stated problem was refunds, yet the literal formula changes 127 sections that have no refund,
and inflates profit eight-fold by exactly twice the bank fee. It was raised with the user with these
figures rather than shipped; what the words described — removing the refund — was implemented.

### Verification

```
NEW BALANCE DOHA | WALLET   profit -285.16 -> 43.84   Internal Transfer 2,535.54 (refund still deducted)
AL SULTAN MEDICAL CENTER    profit -174.77 -> 10.23   Internal Transfer 3,866.04 (refund still deducted)

sections where profit != nqCharge + chg/Txn - bankFee : 0 of 129
whole-report total profit : 2,889.36 -> 3,403.36   (+514.00, the total refund)
```

Regression: 1,726/1,726 export formulas evaluate to their cached values with the changed profit
formula, 107/107 rate pairs, 30 merchants / 55 rows unchanged, control total 3,707 / 1,245,071.15.

---

## Reconciliation block under every Account Code subtotal

**Audit status: PENDING VERIFICATION.** Built from a hand-made block in a user-supplied export,
reproduced as generated output. One block per Account Code subtotal, as requested.

### What it computes

```
RECONCILIATION — Account Code 4
  Bank charges difference (actual vs bank)      = actualBankCharges - totalBankFee
  ACTUAL PROFIT                                 = totalProfit + bankDifference
  Gain For Bank Charges                         = subtotal gain
  ACTUAL PROFIT + GAIN                          = the two above
  Audit portal profit                           [ input ]
  Difference to consider                        = ACTUAL PROFIT + GAIN - portal profit
  Total internal transfers, all charges deducted = subtotal Transfer Deducted (Rent, Others)
  Audit portal total                            [ input ]
  Difference to consider                        = transfers total - portal total
```

`Bank Difference` is already *(modelled − actual)*, so **adding** it applies the correction. The
hand-built version computed *(actual − modelled)* in one cell and subtracted it in the next; this
reuses a column the report already produces and needs one cell instead of two.

### Fixes to the hand-built version

- The duplicate cell (`CA11`, identical formula to `BY11`) is gone.
- `CE11 = CC10 - CC11` referenced an empty cell; dropped.
- Every figure is `ROUND(...,2)`. The hand-built cells were unrounded and carried float noise
  (`0.12999999999999545`, `0.07999999999810825`), which would defeat any `=IF(x=0,...)` check.
- Labels sit beside their values (label merged across A:C, figure in D) rather than two columns apart.
- The block lives in the lead columns, not underneath the data grid, so adding merchants can never
  collide with it — the structural flaw in the original.
- The two audit-portal cells carry the workbook's `input` styling, so it is obvious which cells are
  typed. They are also editable in the web app and flow through to the export.

### Verification

Six blocks, one per Account Code, in both the web report and the export. Formulas reference the
correct subtotal row and chain within the block:

```
RECONCILIATION — Account Code 4      (subtotal on row 9)
  Bank charges difference   0.14      =ROUND(BZ9-BY9,2)
  ACTUAL PROFIT           121.95      =ROUND(BX9+CA9,2)
  ACTUAL PROFIT + GAIN    136.95      =ROUND(D12+D13,2)
  Audit portal profit       0.00      (value - INPUT)
  Difference              136.95      =ROUND(D14-D15,2)
```

Typing 100 into the web app's portal profit moved the difference 136.95 -> 36.95.

Regression: **3,264/3,264** export formulas evaluate to their cached values, 107/107 rate pairs,
30 merchants / 55 rows unchanged, 11 OOXML parts well-formed, control total 3,707 / 1,245,071.15.

### Two bugs found while doing this

**`renderSheetChecks` always reported `Txn_Fee not found`.** Its `rowsById` map was never given a
`txn_fee` entry when that sheet was added, so the panel contradicted the report, which correctly
said three merchants were priced from it. Display only — no figure was ever affected. Fixed.

**`verify-formulas.js` was validating a stale file.** It defaulted to a hard-coded
`exported_Payout_Report_*.xlsx` from an earlier session rather than the newest capture, so the
"1,726/1,726 formulas" reported in several recent entries was measured against an old export, not
the code being committed. The count on the current export is **3,264**. Re-run against the real
file, everything passes — but the earlier claims overstated what had been checked. The script now
selects the most recently written export, so it cannot silently go stale again.

---

## Reconciliation block: left-to-right across the FINAL RECONCILIATION columns

**Audit status: PENDING VERIFICATION.** Replaces the vertical column-D layout with the horizontal
one the user builds by hand, matching the reference workbook cell for cell.

### Layout, as exported

Three rows under each Account Code subtotal (subtotal on row 9 in the sample):

```
row 10   BX "ACTUAL PROFIT NEEDS TO TAKE"        BY   0.14  =ROUND(BZ9-BY9,2)
         BZ 121.95  =ROUND(BX9+CA9,2)            CB "BANK CHARGES DIFFERENCE | ACTUAL VS. BANK"
         CF 15,882.48  =ROUND(CE9,2)             CG "TOTAL INTERNAL TRANSFERS DEDUCTED ALL THE CHARGES"

row 11   BX "ACTUAL PROFIT + NOQOODY GAIN..."    BZ 136.95  =ROUND(BZ10+CD9,2)
         CE "AUDIT PORTAL TOTAL | MATCHING"      CF   0.00  (input)
         CG 15,882.48  =ROUND(CF10-CF11,2)       CH "DIFFERENCE TO CONSIDER"

row 12   BX "AUDIT PORTAL TOTAL | MATCHING"      BZ   0.00  (input)
         CA 136.95  =ROUND(BZ11-BZ12,2)          CB "DIFFERENCE TO CONSIDER"
```

Same cells, same labels, same yellow highlight as the hand-built original.

### Cleaned

- The cell duplicating the bank-charges difference (`CA11` in the original) is gone.
- `CE11 = CC10 - CC11`, which referenced an empty cell, is gone.
- Every figure is `ROUND(...,2)`; the originals carried float noise (`0.12999999999999545`).
- The two audit-portal cells carry the workbook's `input` styling and a consistent number format —
  in the original they were the only cells without the QAR format, and looked like computed cells.
- `Bank Difference` is already *(modelled - actual)*, so `ACTUAL PROFIT` adds it rather than
  recomputing *(actual - modelled)* in a separate cell and subtracting: one cell instead of two.

### Web view

The same three rows render in the same columns, with the two portal figures editable per Account
Code and flowing through to the export. Two display details were needed:

- The block name sits in one sticky cell spanning the three lead columns, clipped to their width —
  as a plain first cell it overflowed across the scrolled columns and covered its neighbour.
- Label text is absolutely positioned so it spills across the empty cells to its right, the way
  Excel renders an overflowing label. An HTML cell otherwise clips it to a numeric column's width.

### Verification

Six blocks, one per Account Code, in both the web report and the export. Formulas reference the
correct subtotal row and chain within the block. 12 inputs rendered (two per block).

Regression: 3,258/3,258 export formulas evaluate to their cached values, 107/107 rate pairs, 30
merchants / 55 rows unchanged, 11 OOXML parts well-formed, control total 3,707 / 1,245,071.15.

---

## Reconciliation block: header row + figures row, Accent 4

**Audit status: PENDING VERIFICATION.** Rebuilt from a second reference workbook, which restructured
the block into a proper header row above a figures row. Cleaner than the previous arrangement and
adopted as-is, with one formula corrected.

### Layout, as exported (subtotal on row 9; block on rows 10-11)

| Col | Header | Formula |
|---|---|---|
| BX | ACTUAL PROFIT NEEDS TO TAKE | `=ROUND(BX9+CB11,2)` |
| BY | ACTUAL PROFIT + NOQOODY GAIN FROM BANK CHARGES | `=ROUND(BX11+CD9,2)` |
| BZ | AUDIT PORTAL PROFIT \| MATCHING | input |
| CA | DIFFERENCE TO CONSIDER | `=ROUND(BY11-BZ11,2)` |
| CB | BANK CHARGES DIFFERENCE \| ACTUAL VS. BANK | `=ROUND(BY9-BZ9,2)` |
| CC | RENT TOTAL | `=ROUND(CC9,2)` |
| CD | AUDIT PORTAL TOTAL \| MATCHING | input |
| CE | TOTAL INTERNAL TRANSFERS DEDUCTED ALL THE CHARGES | `=ROUND(CE9,2)` |
| CF | DIFFERENCE TO CONSIDER | `=ROUND(CE11-CD11,2)` |

Header row in Accent 4 (`#8064A2`, white bold), figures row in a light tint (`#E4DFEC`), the two
audit-portal cells in the workbook's `input` orange.

### The one correction

The reference had `ACTUAL PROFIT = ROUND(BX10-CB12,2)`, subtracting a difference defined as
*(modelled − actual)*. On the sample data the bank took **0.13 more** than modelled, and that formula
returned **133.24** — profit rising because the bank overcharged. Subtracting a negative adds it.

It now **adds** it: `ROUND(BX9+CB11,2)` → **132.98** on that data, which is what the *first*
reference workbook computed for the same figure. Verified on the live export: subtotal profit
122.09, bank charges difference −0.14, actual profit **121.95**.

### Two label changes

`BZ` and `CD` both read "AUDIT PORTAL TOTAL | MATCHING" in the reference, for portal *profit* and
portal *transfer total* respectively; `BZ` is now "AUDIT PORTAL PROFIT | MATCHING". `CC` was
labelled "RENT AUDIT PORTAL TOTAL | MATCHING" but carried `=SUM(CC10)`, the Rent subtotal, with
nothing comparing against it — it is labelled "RENT TOTAL" for what it actually is. **If `CC` was
meant to be a third portal input, it is not one yet.**

### Verification

Six blocks, one per Account Code, in the web report and the export; 12 inputs. Web and export agree
cell for cell, and every column is pixel-aligned with the column header and the subtotal above it:

```
grid col 69..77   colHeader x = subtotal x = reconHead x = reconValue x   ALIGNED (all nine)
```

Regression: 3,264/3,264 export formulas evaluate to their cached values, 107/107 rate pairs, 30
merchants / 55 rows unchanged, 11 OOXML parts well-formed, control total 3,707 / 1,245,071.15.

A note on method: three screenshots in a row appeared to show the block rendering blank, which was
the capture script leaving the table at `scrollLeft=0`, not a layout fault. Measuring the cells'
bounding boxes settled it — the DOM was correct throughout.
