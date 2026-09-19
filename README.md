# Payout Calculator — Phase 1–4

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
- **Actual Bank Charges, Rent, Portal Amount** are external reconciliation figures (bank statement / payout portal) that don't exist in any uploaded sheet. They're editable per report line in the UI, defaulting to 0; everything downstream recalculates live.
- **Wallet Classification / Classification** remain unpopulated ("rule pending") — the reference data shows real values for these, but the signal that decides them isn't present in Card Number, Text Message, or any of the six config sheets. Confirmed *not* load-bearing for the payout report itself: the report groups by `Card Type 2`, which is fully resolved.

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
