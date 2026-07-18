# Weekly Ingest Runbook

**Schedule:** every Thursday, 11:30 PM America/New_York (handled by a Claude Code scheduled
trigger; see below). Runs after the week's reports land in the Drive folder **Precision Blueprint
Development**.

## What the run does
1. **List** new files in the Drive folder (`parentId = 1B9yEG6zbZSQUqzIk5cVI1-19LzqN4CkW`)
   modified since the last successful run (`Ingestion_Log.run_timestamp`).
2. **Download** each recognized report and stage it:

   | Report | Match on title | Table |
   |---|---|---|
   | General Ledger | `General Ledger` | `GL_Snapshots` (+ `GL_Account_Summary`) |
   | Trial Balance | `Trial Balance` | `TrialBalance_Snapshots` |
   | A/R Aging Detail | `A_R Aging` | `AR_Aging_Snapshots` |
   | A/P Aging Detail | `A_P Aging` | `AP_Aging_Snapshots` |
   | JobTread WIP | `job-export` | `JobTread_WIP_Snapshots` |
   | Payroll Detail / Summary | `Payroll`, `_Reports_` | `Payroll_Snapshots` |

   *Ignore:* `Transaction List by Vendor` (per client — not an ingest source).
3. **Parse & normalize** with `scripts/build_financial_db.py`.
4. **Reconcile** (build asserts): TB debits=credits; A/R and A/P open totals tie to the TB A/R and
   A/P lines; payroll totals tie to the report footer. On failure, stop and flag — do not append.
5. **Append** new rows (dedup `GL_Snapshots` on `row_uid`; snapshot tables keyed by `as_of_date`).
   Re-apply `Sub_Contract_Reference` vendor→job mappings to set `matched` allocation.
6. **Log** one `Ingestion_Log` row per file with rows ingested + control total + reconciled Y/N.
7. **Commit & push** to the working branch; refresh `workbook/*.xlsx`.

## New reports the client is adding
- **A/P Aging Detail** and **Payroll Summary** are now first-class (tables 5 & 6).
- If a dedicated report format changes, update the parser and add a note here.

## Notes / decisions on record
- Ashley Espinoza's labor is **administrative/overhead** → `cost_classification = Overhead/Admin`,
  excluded from job cost.
- Subcontractor allocation tiering (`direct`/`matched`/`unallocated`) is intentionally
  conservative pending client direction; adjust in `build_financial_db.py` (`parse_gl`) and
  `Sub_Contract_Reference`.
- General Ledger export is **Cash** basis; Trial Balance is **Accrual** basis — expected, tracked
  in `gl_basis` / `basis` columns.
