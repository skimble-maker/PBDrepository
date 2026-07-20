#!/usr/bin/env python3
"""
Idempotent merge of a freshly-built weekly snapshot into the canonical /data store.

`build_financial_db.py` produces this-run's table CSVs. This script APPENDS only genuinely
new rows into data/*.csv, so re-running the same week is a no-op and a new week accumulates
history. Dedup keys per table:

  GL_Snapshots            -> row_uid
  GL_Account_Summary      -> (gl_account, gl_period)
  TrialBalance_Snapshots  -> as_of_date            (snapshot is atomic)
  AR_Aging_Snapshots      -> as_of_date
  AP_Aging_Snapshots      -> as_of_date
  Payroll_Snapshots       -> pay_date
  JobTread_WIP_Snapshots  -> (as_of_date, jobtread_job_id)
  Ingestion_Log           -> append-always
  Clients                 -> upsert by client_id (adds new; never clobbers manual fields)
  Sub_Contract_Reference  -> upsert by vendor_name, PRESERVING manual linked_*/allocation
                             edits; refreshes gl_total_paid_h1_cash + gl_txn_count

Usage:  python scripts/merge_snapshots.py <new_tables_dir> [--dest data] [--dry-run]
"""
import csv, os, sys, argparse

KEYS = {  # row-level dedup: append individual rows whose key is new
    "GL_Snapshots":           lambda r: r["row_uid"],
    "GL_Account_Summary":     lambda r: (r["gl_account"], r.get("gl_period","")),
    "Payroll_Snapshots":      lambda r: r["pay_date"],
    "JobTread_WIP_Snapshots": lambda r: (r["as_of_date"], r["jobtread_job_id"]),
}
# atomic snapshots: a report as-of a date is all-or-nothing — append the WHOLE new
# snapshot only if that as_of_date isn't already present (never merge partial dates).
ATOMIC_SNAPSHOT = {"TrialBalance_Snapshots", "AR_Aging_Snapshots", "AP_Aging_Snapshots"}
APPEND_ALWAYS = {"Ingestion_Log"}
UPSERT = {  # table -> (key col, columns that must NOT be overwritten if already set)
    "Clients":                ("client_id", []),
    "Sub_Contract_Reference": ("vendor_name",
                               ["linked_client_id","linked_job_name","linked_jobtread_job_id",
                                "allocation_confidence","match_method","is_1099","notes"]),
}

def read(path):
    if not os.path.exists(path): return [], []
    with open(path, newline="", encoding="utf-8") as f:
        r = csv.DictReader(f); return list(r), (r.fieldnames or [])

def write(path, cols, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore"); w.writeheader()
        for row in rows: w.writerow(row)

def merge(new_dir, dest, dry):
    report = []
    for fn in sorted(os.listdir(new_dir)):
        if not fn.endswith(".csv"): continue
        table = fn[:-4]
        new_rows, new_cols = read(os.path.join(new_dir, fn))
        cur_rows, cur_cols = read(os.path.join(dest, fn))
        cols = cur_cols or new_cols
        added = 0

        if table in APPEND_ALWAYS:
            merged = cur_rows + new_rows; added = len(new_rows)

        elif table in ATOMIC_SNAPSHOT:
            have = {r["as_of_date"] for r in cur_rows}
            merged = list(cur_rows)
            for nr in new_rows:
                if nr["as_of_date"] not in have:          # new snapshot date -> append all its rows
                    merged.append(nr); added += 1

        elif table in UPSERT:
            keycol, protect = UPSERT[table]
            by_key = {r[keycol]: dict(r) for r in cur_rows}
            for nr in new_rows:
                k = nr[keycol]
                if k in by_key:
                    tgt = by_key[k]
                    for c, v in nr.items():                       # refresh non-protected fields
                        if c in protect and (tgt.get(c) or "").strip():
                            continue                              # keep manual edit
                        tgt[c] = v
                else:
                    by_key[k] = dict(nr); added += 1
            merged = list(by_key.values())

        else:  # snapshot tables — key-based dedup
            keyfn = KEYS[table]
            seen = {keyfn(r) for r in cur_rows}
            merged = list(cur_rows)
            for nr in new_rows:
                if keyfn(nr) not in seen:
                    merged.append(nr); seen.add(keyfn(nr)); added += 1

        report.append((table, len(cur_rows), added, len(merged)))
        if not dry:
            write(os.path.join(dest, fn), cols, merged)
    return report

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("new_dir")
    ap.add_argument("--dest", default="data")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()
    rep = merge(a.new_dir, a.dest, a.dry_run)
    print(f"{'TABLE':28} {'existing':>9} {'added':>7} {'total':>7}")
    for t, c, ad, tot in rep:
        print(f"{t:28} {c:>9} {ad:>7} {tot:>7}")
    print(("[dry-run] " if a.dry_run else "") + f"total new rows: {sum(x[2] for x in rep)}")
