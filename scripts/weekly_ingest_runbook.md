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
5. **Append** with `scripts/merge_snapshots.py <new_tables_dir>` — idempotent: re-running the
   same week adds 0 rows; a new week appends (GL dedup on `row_uid`; TB/AR/AP appended whole per
   new `as_of_date`; Payroll on `pay_date`; WIP on `as_of_date`+job). It also **upserts**
   `Sub_Contract_Reference`, preserving manual `linked_*` / `allocation_confidence` edits while
   refreshing GL totals. Vendors mapped to a job there make matching GL sub lines `matched`.
6. **Log** one `Ingestion_Log` row per file with rows ingested + control total + reconciled Y/N.
7. **Commit & push** to the working branch; rebuild `workbook/*.xlsx`.

> **Google Sheet note:** the Drive `create_file` API can *create* a Sheet from CSV but cannot
> edit one in place (nor write multiple tabs). So the Drive "Financial Performance DB (tables)"
> Sheets are point-in-time exports, not live-updated. The **living store is git** (`data/*.csv`)
> + `workbook/Financial_Performance_DB.xlsx`. To keep a live Google Sheet, open that workbook in
> Google Sheets once (it becomes a native multi-tab Sheet you own); the weekly run keeps git +
> the workbook current. Record vendor→job assignments in git `Sub_Contract_Reference.csv` (the
> merge preserves them) — edits made only in a Drive export will not flow back.

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
