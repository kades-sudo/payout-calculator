# Payout Calculator — Phase 1–4 (+ Phase 3 refinements + Phase 3 update)

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
- **Section 3 (freeze MID/Merchant columns in Excel)**: re-tested specifically against `xlsx-js-style` (the library adopted in Phase 3 for header colors) in case it added freeze-pane support beyond the official SheetJS build — it does not; writing any pane/freeze configuration is silently dropped, confirmed in the raw XML. This remains a hard library limitation, not a guess. The web app's Grid View already freezes the Account Code, MID and Merchant columns while scrolling horizontally (built in the prior Phase 3 pass) — that may already cover the need; flagging this again since it directly re-raises something already decided against for the Excel file specifically.

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
