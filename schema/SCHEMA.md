# Financial Performance Database — Schema

**Client:** Precision Blueprint Development LLC (construction / remodel, AZ)
**Maintained by:** IronGate BP · **Baseline load:** 2026-07-18
**Home:** git repository (`/data/*.csv`, authoritative) + Google Sheet companion (`/workbook/*.xlsx`)

Every snapshot table is **append-only**: each weekly pull adds rows stamped with the report's
`as_of_date`/period and an `ingested_at` timestamp, so history is preserved and never overwritten.
Money fields are stored as signed decimals (no `$`/`,`). Dates are ISO `YYYY-MM-DD`.

Data start points differ by source, so the baseline reconciles what exists today and tracks forward:

| Source report | Grain | Baseline coverage | Basis |
|---|---|---|---|
| General Ledger | transaction line | Jan 1 – Jun 30 2026 | Cash |
| Trial Balance | account balance | as of Jul 11 2026 | Accrual |
| A/R Aging Detail | open invoice/payment | as of Jul 11 2026 | — |
| A/P Aging Detail | open bill | as of Jul 15 2026 | — |
| JobTread WIP (`job-export`) | job | as of Jul 18 2026 | — |
| Payroll Summary | weekly check | Jun 20 2025 – Jun 12 2026 | — |
| Payroll Detail | pay period | Jun 5 – Jul 17 2026 | — |

---

## 1. Clients
Master registry of PBD's customers/jobs. Bridges A/R customer paths ↔ JobTread jobs.

| Column | Type | Notes |
|---|---|---|
| client_id | text (PK) | `CL-<slug>` derived from job/customer name |
| customer_name | text | Top-level QBO customer |
| job_name | text | Job / project (QBO leaf) |
| qbo_full_path | text | Full `Customer:Sub:Job` string from A/R |
| jobtread_job_id | text | JobTread ID (links to `JobTread_WIP_Snapshots`) |
| job_number | text | JobTread job number |
| price_type | text | `fixed` (all current jobs are fixed-fee) |
| sources | text | Which reports the entity appears in |
| active | Y/N | |
| notes | text | e.g. "in A/R but not in current WIP export" |

## 2. GL_Snapshots  *(transaction level)*
Every General Ledger line. Deduplicate on `row_uid` when appending new weeks.

| Column | Type | Notes |
|---|---|---|
| gl_account | text | Account whose ledger the line sits under |
| txn_date | date | |
| txn_type | text | Expense, Check, Journal Entry, Deposit, Bill, … |
| num | text | Reference / check no. |
| name | text | Payee / vendor / customer |
| memo | text | Description / memo |
| split_account | text | Offsetting account |
| amount | decimal | Signed |
| running_balance | decimal | As reported |
| **is_subcontractor** | Y/N | Line sits under the `Sub-Contractors` account |
| **allocation_confidence** | enum | `direct` / `matched` / `unallocated` (subcontractor lines only) — see §7 |
| allocated_jobtread_id | text | Job the cost is tied to, when known |
| row_uid | text | Dedup key (hash) |
| source_file, gl_period, gl_basis, ingested_at | | Provenance |

## 3. TrialBalance_Snapshots
| Column | Type | Notes |
|---|---|---|
| as_of_date | date | |
| basis | text | Accrual |
| account | text | Full account path |
| debit / credit | decimal | One populated per row |
| net_debit_credit | decimal | debit − credit |
| source_file, ingested_at | | |

*Integrity:* Σdebit = Σcredit = **1,881,728.02** at baseline.

## 4. AR_Aging_Snapshots
| Column | Type | Notes |
|---|---|---|
| as_of_date | date | |
| aging_bucket | text | CURRENT / 1-30 / 31-60 / 61-90 / 91+ |
| txn_date, txn_type, num | | Invoice or Payment |
| customer_full_path | text | QBO `Customer:Sub:Job` |
| client_id | text | FK → Clients |
| due_date | date | |
| amount | decimal | |
| open_balance | decimal | |
| source_file, ingested_at | | |

*Integrity:* Σopen_balance = **432,296.14** (= TB A/R line).

## 5. AP_Aging_Snapshots
| Column | Type | Notes |
|---|---|---|
| as_of_date | date | |
| aging_bucket | text | |
| txn_date, txn_type, num | | Bill |
| vendor_name | text | |
| due_date | date | |
| days_past_due | int | |
| amount / open_balance | decimal | |
| is_subcontractor | Y/N | (blank at baseline — flag when vendor mapped) |
| source_file, ingested_at | | |

*Integrity:* Σopen_balance = **12,155.00** (= TB A/P line).

## 6. Payroll_Snapshots
One row per weekly check. **Ashley Espinoza's labor is administrative/overhead** and is
**not** allocated to jobs (see `cost_classification`).

| Column | Type | Notes |
|---|---|---|
| pay_date | date | Check date |
| employee_name | text | |
| **cost_classification** | text | `Overhead/Admin` — excluded from job cost |
| hours | decimal | |
| gross_pay | decimal | Total paid |
| tax_withheld / deductions / net_pay | decimal | |
| check_no | text | `DD` = direct deposit |
| employer_liability | decimal | |
| total_expense | decimal | gross + employer liability |
| source_file, ingested_at | | |

*Integrity:* Σtotal_expense = **80,007.78** across 52 checks.

## 7. JobTread_WIP_Snapshots
Job-level work-in-progress / fixed-fee performance.

| Column | Type | Notes |
|---|---|---|
| as_of_date | date | |
| jobtread_job_id | text (FK → Clients) | |
| job_number, job_name | | |
| client_id | text | FK → Clients |
| price_type | text | fixed |
| projected_price, budget_cost, projected_cost, actual_cost, cost_to_complete | decimal | |
| budget_variance, projected_profit, projected_margin, pct_cost_complete | decimal | |
| earned_revenue, invoiced, over_invoiced, under_invoiced, open_invoices | decimal | |
| original_budget_cost, bills, budgeted_profit | decimal | |
| source_file, ingested_at | | |

## 8. Sub_Contract_Reference  *(fixed-fee matching bridge)*
The vendor→job map that powers subcontractor cost allocation. Seeded with **every distinct
subcontractor vendor** found in the GL (38 vendors, ~$379,880 H1 cash). The GL carries **no**
job reference on subcontractor payments (they are Zelle/expense lines), so this table is where a
vendor is assigned to a job.

| Column | Type | Notes |
|---|---|---|
| vendor_name | text (PK) | As it appears in the GL |
| normalized_vendor | text | Lower/space-normalized for matching |
| linked_client_id | text | FK → Clients (fill in) |
| linked_job_name | text | (fill in) |
| linked_jobtread_job_id | text | (fill in) |
| **allocation_confidence** | enum | `direct` / `matched` / `unallocated` |
| match_method | text | manual / memo-link / reference |
| gl_total_paid_h1_cash | decimal | Amount seen in GL (Jan–Jun) |
| gl_txn_count | int | |
| is_1099 | Y/N | |
| notes | text | |

### Allocation-confidence tiers (conservative default — adjustable)
- **direct** — the GL/bill line itself carries an explicit JobTread job link or job number.
- **matched** — vendor is mapped to a job here in `Sub_Contract_Reference` (fixed-fee match).
- **unallocated** — neither; cost is a subcontractor spend not yet tied to a job.

At baseline every subcontractor line is `unallocated` (no job refs in the GL data). As vendors are
mapped here, the weekly ingest promotes matching GL lines to `matched`.

## 9. Ingestion_Log
Audit trail — one row per file ingested per run, with the reconciliation control total.

| Column | Type | Notes |
|---|---|---|
| ingestion_id | text | |
| run_timestamp | datetime | |
| report_type | text | |
| source_file / source_drive_id | text | Drive provenance |
| as_of_or_period | text | |
| basis | text | |
| rows_ingested | int | |
| control_total | decimal | Reconciliation figure |
| reconciled | Y/N | |
| status | text | |
| notes | text | |
