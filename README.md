# Payout Calculator — Phase 1–2

A single-page, client-side web app that automates OMS transaction enrichment, master-list matching, and Process Raw Details generation for payout processing.

Open `index.html` in a browser (or publish it as-is) — no build step or backend required. All parsing and lookups run in the browser via [SheetJS](https://sheetjs.com/).

## Master List workbook ingestion (Phase 2)

The Master List / Configuration upload recognizes up to six sheets by fuzzy name match (case, spacing and punctuation-insensitive), configured in `SHEET_ALIASES` in `index.html`:

| Sheet | Required this phase | Notes |
|---|---|---|
| `Master_List` | Yes | Primary key `MID`. Merged across every uploaded workbook, so separate Daily/Weekly lists can be uploaded together. |
| `Card Type Classification` | Yes | Column A = lookup key, column B (`Card Type 1`) = output **Card Type**, column C (`Card Type 2`) = output **Card Type 2** — verified against a real export; header names are matched first, with A/B/C position as fallback. |
| `Terminal_Mapping` | No (parsed, not yet used) | Keyed by MID + Terminal ID. |
| `Bank_Cost` | No (parsed, not yet used) | Keyed by Cost ID. |
| `Special_Rate` | No (parsed, not yet used) | Keyed by Account Code. |
| `Card_Types` | No (parsed, not yet used) | Maps Payment Group → Master List rate column → Bank Cost column. |

The four optional sheets are detected, parsed and row-counted in the Step 1 checklist so a following phase can wire in merchant-rate lookup, bank-cost lookup, and Account Code / terminal matching without re-touching the ingestion layer.

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
