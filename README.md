# Payout Calculator — Phase 1

A single-page, client-side web app that automates OMS transaction enrichment, master-list matching, and Process Raw Details generation for payout processing.

Open `index.html` in a browser (or publish it as-is) — no build step or backend required. All parsing and lookups run in the browser via [SheetJS](https://sheetjs.com/).

## Workflow (Phase 1 scope)

1. Upload the Master List / Configuration workbook(s) — must contain a `Master_List` sheet (`MID` → `Merchant Name`) and a `Card Type Classification` sheet (lookup key in column A, `Card Type 2` in column C, `Card Type` in column E, matching the source `VLOOKUP` formulas).
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
