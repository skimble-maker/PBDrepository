# Precision Blueprint Development — Financial Performance Database

Weekly-tracked financial performance for **Precision Blueprint Development LLC**, built by IronGate BP.

Source reports land in the Google Drive folder **Precision Blueprint Development** (mostly by
Thursday afternoon). Every **Thursday 11:30 PM America/New_York** the new outputs are pulled,
normalized into the tables below, reconciled, and appended here.

## Tables (`/data/*.csv` — authoritative; `/workbook/*.xlsx` — Google Sheet companion)

| Table | Grain | Rows @ baseline | Source report(s) |
|---|---|--:|---|
| `Clients` | customer / job | 25 | A/R Aging + JobTread WIP |
| `GL_Snapshots` | GL transaction line | 2,154 | General Ledger |
| `GL_Account_Summary` | GL account rollup | 56 | General Ledger (derived) |
| `TrialBalance_Snapshots` | account balance | 82 | Trial Balance |
| `AR_Aging_Snapshots` | open invoice/payment | 37 | A/R Aging Detail |
| `AP_Aging_Snapshots` | open bill | 7 | A/P Aging Detail |
| `Payroll_Snapshots` | weekly check | 52 | Payroll Summary |
| `JobTread_WIP_Snapshots` | job | 23 | JobTread WIP export |
| `Sub_Contract_Reference` | subcontractor vendor | 38 | GL (seed) — vendor→job map |
| `Ingestion_Log` | file × run | 6 | audit trail |

Column definitions: **[`schema/SCHEMA.md`](schema/SCHEMA.md)**.
Weekly process: **[`scripts/weekly_ingest_runbook.md`](scripts/weekly_ingest_runbook.md)**.

## Baseline reconciliation (all tie out ✓)

| Check | Value |
|---|---|
| Trial Balance Σdebit = Σcredit | **1,881,728.02** |
| A/R open balance = TB A/R line | **432,296.14** |
| A/P open balance = TB A/P line | **12,155.00** |
| Payroll total expense (52 checks) | **80,007.78** |
| GL transaction lines | **2,154** (169 subcontractor lines) |
| JobTread fixed-fee jobs | **23** |
| Clients matched to a JobTread job | **23 / 25** |

## Subcontractor allocation (the reconciliation focus)

Subcontractors are ~$383.8K of cost (Trial Balance) — the largest job-cost driver. The GL records
these as vendor payments (mostly Zelle) with **no job reference**, so cost can't be tied to a job
from the GL alone. `Sub_Contract_Reference` is the bridge: it seeds all 38 subcontractor vendors,
and each GL subcontractor line carries an **`allocation_confidence`** flag:

- `direct` — line has an explicit JobTread link/job number
- `matched` — vendor mapped to a job in `Sub_Contract_Reference`
- `unallocated` — not yet tied to a job *(all lines at baseline)*

Assigning vendors to jobs in `Sub_Contract_Reference` progressively lights up `matched` allocation.

## Regenerating the Google Sheet

`workbook/Financial_Performance_DB.xlsx` is the full multi-tab workbook (incl. transaction-level
GL). Opening it in Google Sheets (File ▸ Import, or "Open with Google Sheets") produces the live
multi-tab Sheet. `workbook/PBD_Financial_DB_forSheet.xlsx` is a lighter variant (GL rolled up to
the account-summary tab) for a faster Sheet.

## Rebuilding the tables

```bash
pip install openpyxl
PBD_GL_TXT=<general-ledger csv/text> python scripts/build_financial_db.py
```
The build **fails loudly** if any control total (TB balance, A/R, A/P, payroll) doesn't reconcile.
