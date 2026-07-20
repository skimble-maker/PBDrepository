# Hands-off weekly ingest — claude.ai Routines setup

Set this up once in the **claude.ai Routines UI** (claude.ai → Routines → New). Unlike a
trigger created from inside a session, a Routine created here carries **your** Google Drive +
GitHub connections into each run, so the weekly pull is fully unattended.

## Settings

| Field | Value |
|---|---|
| **Name** | PBD Weekly Financial Ingest |
| **Schedule** | Weekly · **Thursday** · **11:30 PM** |
| **Time zone** | **America/New_York** (handles EST/EDT automatically — no manual DST shift) |
| **Environment / repo** | the environment holding `skimble-maker/PBDrepository`, branch `claude/financial-performance-db-tables-0u0pik` |
| **Connectors** (must be ON) | **Google Drive**, **GitHub** |
| **New session each run** | Yes |

## Prompt (paste verbatim)

```
You are IronGate BP's weekly bookkeeping automation for client Precision Blueprint
Development LLC. Today is the Thursday-night ingest. Work in the git repo
skimble-maker/PBDrepository on branch claude/financial-performance-db-tables-0u0pik.

1. Read scripts/weekly_ingest_runbook.md and schema/SCHEMA.md, then follow the runbook.
2. Using Google Drive, list files in the folder "Precision Blueprint Development"
   (folder id 1B9yEG6zbZSQUqzIk5cVI1-19LzqN4CkW) modified since the most recent
   run_timestamp in data/Ingestion_Log.csv.
3. Download each recognized report and normalize with scripts/build_financial_db.py:
   - General Ledger        -> GL_Snapshots (+ GL_Account_Summary)
   - Trial Balance         -> TrialBalance_Snapshots
   - A_R Aging Detail      -> AR_Aging_Snapshots
   - A_P Aging Detail      -> AP_Aging_Snapshots
   - job-export (JobTread) -> JobTread_WIP_Snapshots
   - Payroll Detail/Summary-> Payroll_Snapshots
   IGNORE "Transaction List by Vendor" — it is not an ingest source.
4. Reconcile before writing anything: Trial Balance debits must equal credits; A/R and A/P
   open totals must tie to the TB A/R and A/P lines; payroll totals must tie to the report
   footer. If ANY check fails, do not append — stop and report what broke.
5. Append only NEW rows: dedup GL_Snapshots on row_uid; snapshot tables are keyed by
   as_of_date. Re-apply Sub_Contract_Reference vendor->job mappings so matching GL
   subcontractor lines get allocation_confidence = matched.
6. Add one Ingestion_Log row per file (rows ingested, control total, reconciled Y/N).
   Refresh workbook/*.xlsx and update the Google Sheets in the Drive subfolder
   "Financial Performance DB (tables)" (folder id 1t_MEQUzxjGu9zAAdshHWyf-nNhu9Wh7Z).
7. Commit and push to the branch. Send me a short summary: which reports loaded, the
   control totals, anything that failed reconciliation, and any subcontractor spend that is
   still unallocated.

Standing decisions: Ashley Espinoza's payroll is Overhead/Admin and is never allocated to
jobs. Keep subcontractor allocation_confidence conservative — direct = an explicit JobTread
job link/number on the line; matched = vendor mapped in Sub_Contract_Reference; otherwise
unallocated. If reports are missing (not yet uploaded by Thursday night), ingest what is
present and note the gaps rather than failing the whole run.
```

## After it's live
- Once this Routine is running, delete the in-session trigger
  `PBD Weekly Financial Ingest (Thu 11:30pm ET)` (trig id on file) to avoid double-runs.
- First run will only add rows that are newer than the 2026-07-18 baseline, so history
  is never duplicated.
