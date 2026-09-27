#!/usr/bin/env python3
"""
MySQL experiment: table WITH primary key vs table WITHOUT primary key.

Steps (as given in the assignment):
  1. Create two tables with the same columns:
       data_pk    -> id is the PRIMARY KEY
       data_nopk  -> id is a normal column (no primary key, no index)
  2. Insert 1,000,000 random rows into each table (same data in both).
  3. SELECT the row with id = 10, 1,000, 10,000, 100,000 and 1,000,000 and time it.
  4. Insert 20,000 more random rows into the same tables and time it.
  5. Insert 40,000, then 60,000, then 100,000 more rows and time each.

Extras collected for the analysis:
  - EXPLAIN plan + how many rows MySQL actually read for each SELECT
  - the same SELECTs with LIMIT 1 (shows the "scan stops early" effect)
  - the SELECTs repeated at the end (1,220,000 rows) to show growth with table size
  - table size on disk (data + index)

Usage:
  pip install mysql-connector-python openpyxl
  python3 pk_experiment.py --user root --password YOUR_PASSWORD
  -> writes results.json, then builds pk_experiment_results.xlsx (via build_excel.py)
"""
import argparse
import json
import platform
import random
import statistics
import string
import subprocess
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

import mysql.connector

INITIAL_ROWS = 1_000_000
SELECT_TARGETS = [10, 1_000, 10_000, 100_000, 1_000_000]
INSERT_BATCHES = [20_000, 40_000, 60_000, 100_000]
SELECT_REPEATS = 5          # each SELECT is executed this many times per run
CHUNK = 1_000               # rows per multi-row INSERT statement
TABLES = {"with_pk": "data_pk", "without_pk": "data_nopk"}

CITIES = ["Yogyakarta", "Jakarta", "Surabaya", "Bandung", "Medan", "Semarang",
          "Makassar", "Denpasar", "Malang", "Palembang", "Solo", "Bogor"]
BASE_DATE = datetime(2020, 1, 1)


def ddl(table, with_pk):
    return f"""
    CREATE TABLE {table} (
        id          INT            NOT NULL,
        name        VARCHAR(50)    NOT NULL,
        email       VARCHAR(100)   NOT NULL,
        age         INT            NOT NULL,
        city        VARCHAR(30)    NOT NULL,
        salary      DECIMAL(12,2)  NOT NULL,
        created_at  DATETIME       NOT NULL{',' if with_pk else ''}
        {'PRIMARY KEY (id)' if with_pk else ''}
    ) ENGINE=InnoDB"""


def random_rows(start_id, count, rng):
    """Generate `count` random rows with ids start_id .. start_id+count-1."""
    letters = string.ascii_lowercase
    rows = []
    for i in range(start_id, start_id + count):
        name = "".join(rng.choices(letters, k=rng.randint(5, 12))).capitalize()
        email = f"{name.lower()}{rng.randint(1, 9999)}@example.com"
        rows.append((
            i,
            name,
            email,
            rng.randint(18, 65),
            rng.choice(CITIES),
            round(rng.uniform(3_000_000, 50_000_000), 2),
            BASE_DATE + timedelta(seconds=rng.randint(0, 6 * 365 * 24 * 3600)),
        ))
    return rows


def timed_insert(conn, table, rows):
    """Insert rows in multi-row INSERT statements of CHUNK rows, one transaction. Returns seconds."""
    cur = conn.cursor()
    sql = f"INSERT INTO {table} (id,name,email,age,city,salary,created_at) VALUES (%s,%s,%s,%s,%s,%s,%s)"
    t0 = time.perf_counter()
    for i in range(0, len(rows), CHUNK):
        cur.executemany(sql, rows[i:i + CHUNK])
    conn.commit()
    elapsed = time.perf_counter() - t0
    cur.close()
    return elapsed


def timed_select(conn, table, target, limit1=False):
    cur = conn.cursor()
    sql = f"SELECT * FROM {table} WHERE id = %s" + (" LIMIT 1" if limit1 else "")
    t0 = time.perf_counter()
    cur.execute(sql, (target,))
    res = cur.fetchall()
    elapsed = time.perf_counter() - t0
    cur.close()
    assert len(res) == 1, f"expected 1 row for id={target} in {table}, got {len(res)}"
    return elapsed * 1000.0  # ms


def rows_examined(conn, table, target, limit1=False):
    """EXPLAIN plan + handler counters = how many rows MySQL really touched."""
    cur = conn.cursor(dictionary=True)
    sql = f"SELECT * FROM {table} WHERE id = {int(target)}" + (" LIMIT 1" if limit1 else "")
    cur.execute("EXPLAIN " + sql)
    plan = cur.fetchall()[0]

    def snap():
        cur.execute("SHOW SESSION STATUS LIKE 'Handler_read%'")
        return {r["Variable_name"]: int(r["Value"]) for r in cur.fetchall()}

    # counter deltas (no FLUSH privilege needed); subtract the cost of SHOW STATUS itself
    a, b = snap(), snap()
    overhead = {k: b[k] - a[k] for k in a}
    before = snap()
    cur.execute(sql)
    cur.fetchall()
    after = snap()
    handlers = {k: after[k] - before[k] - overhead[k] for k in after}
    cur.close()
    return {
        "explain_type": plan.get("type"),
        "explain_key": plan.get("key"),
        "explain_rows_estimate": plan.get("rows"),
        "explain_extra": plan.get("Extra"),
        "rows_read": sum(handlers.values()),
        "handler_read_key": handlers.get("Handler_read_key", 0),
        "handler_read_rnd_next": handlers.get("Handler_read_rnd_next", 0),
    }


def table_size(conn, db, table):
    cur = conn.cursor()
    cur.execute(f"ANALYZE TABLE {table}")
    cur.fetchall()
    cur.execute("SELECT COUNT(*) FROM " + table)
    count = cur.fetchone()[0]
    cur.execute("""SELECT data_length, index_length FROM information_schema.tables
                   WHERE table_schema=%s AND table_name=%s""", (db, table))
    data_len, idx_len = cur.fetchone()
    cur.close()
    return {"row_count": count, "data_mb": data_len / 1048576, "index_mb": idx_len / 1048576}


def select_round(conn, run, stage, row_count, limit1, out):
    for target in SELECT_TARGETS:
        for label, table in TABLES.items():
            for rep in range(1, SELECT_REPEATS + 1):
                ms = timed_select(conn, table, target, limit1)
                out.append({"run": run, "stage": stage, "table_rows": row_count,
                            "variant": "WHERE id = N LIMIT 1" if limit1 else "WHERE id = N",
                            "table": label, "target_id": target, "repeat": rep, "ms": ms})


def environment(conn):
    cur = conn.cursor()
    cur.execute("SELECT VERSION()")
    version = cur.fetchone()[0]
    cur.execute("SHOW VARIABLES WHERE Variable_name IN "
                "('innodb_buffer_pool_size','innodb_flush_log_at_trx_commit','innodb_page_size')")
    vars_ = {k: v for k, v in cur.fetchall()}
    cur.close()
    cpu = platform.processor() or platform.machine()
    try:
        if sys.platform == "darwin":
            cpu = subprocess.check_output(["sysctl", "-n", "machdep.cpu.brand_string"], text=True).strip()
        elif Path("/proc/cpuinfo").exists():
            for line in Path("/proc/cpuinfo").read_text().splitlines():
                if line.startswith("model name"):
                    cpu = line.split(":", 1)[1].strip()
                    break
    except Exception:
        pass
    return {
        "mysql_version": version,
        "innodb_buffer_pool_mb": int(vars_.get("innodb_buffer_pool_size", 0)) // 1048576,
        "innodb_flush_log_at_trx_commit": vars_.get("innodb_flush_log_at_trx_commit"),
        "innodb_page_size": vars_.get("innodb_page_size"),
        "os": f"{platform.system()} {platform.release()}",
        "cpu": cpu,
        "python": platform.python_version(),
        "run_started": datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S UTC%z").replace("UTC+0000", "UTC"),
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=3306)
    ap.add_argument("--user", default="root")
    ap.add_argument("--password", default="")
    ap.add_argument("--database", default="pk_experiment")
    ap.add_argument("--runs", type=int, default=3, help="repeat the whole experiment N times (default 3)")
    ap.add_argument("--seed", type=int, default=2026)
    ap.add_argument("--out", default="results.json")
    ap.add_argument("--machine-label", default="", help="e.g. 'MacBook Pro M2' (shown in the Excel file)")
    ap.add_argument("--no-verify", action="store_true", help="skip the final verification run inside MySQL")
    args = ap.parse_args()

    server = mysql.connector.connect(host=args.host, port=args.port, user=args.user, password=args.password)
    server.cursor().execute(f"CREATE DATABASE IF NOT EXISTS {args.database}")
    server.close()
    conn = mysql.connector.connect(host=args.host, port=args.port, user=args.user,
                                   password=args.password, database=args.database, autocommit=False)

    results = {"env": environment(conn), "config": {
        "initial_rows": INITIAL_ROWS, "select_targets": SELECT_TARGETS, "insert_batches": INSERT_BATCHES,
        "select_repeats": SELECT_REPEATS, "insert_chunk": CHUNK, "runs": args.runs, "seed": args.seed,
        "machine_label": args.machine_label,
        "ddl_with_pk": ddl(TABLES["with_pk"], True).strip(),
        "ddl_without_pk": ddl(TABLES["without_pk"], False).strip()},
        "inserts": [], "selects": [], "rows_examined": [], "sizes": []}

    for run in range(1, args.runs + 1):
        print(f"\n=== Run {run}/{args.runs} ===", flush=True)
        rng = random.Random(args.seed + run)
        cur = conn.cursor()
        for label, table in TABLES.items():
            cur.execute(f"DROP TABLE IF EXISTS {table}")
            cur.execute(ddl(table, label == "with_pk"))
        conn.commit()
        cur.close()
        # alternate which table goes first each run, so neither always gets the "warm" advantage
        order = list(TABLES.items()) if run % 2 == 1 else list(reversed(TABLES.items()))

        # Step 2: initial 1,000,000 rows (same data into both tables)
        rows = random_rows(1, INITIAL_ROWS, rng)
        for label, table in order:
            sec = timed_insert(conn, table, rows)
            results["inserts"].append({"run": run, "table": label, "step": "Initial load",
                                       "rows_inserted": INITIAL_ROWS, "rows_before": 0,
                                       "rows_after": INITIAL_ROWS, "seconds": sec})
            print(f"  initial load {label:11s} {INITIAL_ROWS:>9,} rows  {sec:8.2f} s", flush=True)
        del rows
        for label, table in TABLES.items():
            results["sizes"].append({"run": run, "stage": "after 1,000,000", "table": label,
                                     **table_size(conn, args.database, table)})

        # Step 3: SELECT rows 10 .. 1,000,000
        select_round(conn, run, "after initial load", INITIAL_ROWS, False, results["selects"])
        select_round(conn, run, "after initial load", INITIAL_ROWS, True, results["selects"])
        if run == 1:
            for target in SELECT_TARGETS:
                for label, table in TABLES.items():
                    for limit1 in (False, True):
                        results["rows_examined"].append({
                            "table": label, "target_id": target,
                            "variant": "WHERE id = N LIMIT 1" if limit1 else "WHERE id = N",
                            **rows_examined(conn, table, target, limit1)})
        for target in SELECT_TARGETS:
            med = {l: statistics.median(s["ms"] for s in results["selects"]
                                        if s["run"] == run and s["table"] == l and s["target_id"] == target
                                        and s["variant"] == "WHERE id = N" and s["stage"] == "after initial load")
                   for l in TABLES}
            print(f"  SELECT id={target:>9,}  with PK {med['with_pk']:9.3f} ms   without PK {med['without_pk']:9.3f} ms",
                  flush=True)

        # Steps 4-5: insert 20k, 40k, 60k, 100k more rows (cumulative, same tables)
        next_id = INITIAL_ROWS + 1
        for batch in INSERT_BATCHES:
            rows = random_rows(next_id, batch, rng)
            for label, table in order:
                sec = timed_insert(conn, table, rows)
                results["inserts"].append({"run": run, "table": label, "step": f"Insert {batch:,}",
                                           "rows_inserted": batch, "rows_before": next_id - 1,
                                           "rows_after": next_id - 1 + batch, "seconds": sec})
                print(f"  insert {batch:>7,} {label:11s} {sec:8.3f} s", flush=True)
            next_id += batch
        final_rows = next_id - 1

        # Extra: SELECT again on the bigger table
        select_round(conn, run, "after all inserts", final_rows, False, results["selects"])
        for label, table in TABLES.items():
            results["sizes"].append({"run": run, "stage": f"after {final_rows:,}", "table": label,
                                     **table_size(conn, args.database, table)})

    conn.close()
    Path(args.out).write_text(json.dumps(results, indent=1, default=str))
    print(f"\nSaved raw results -> {args.out}")

    Path(args.out).with_name("verification.json").unlink(missing_ok=True)   # never mix in an old run's file
    if not args.no_verify:
        try:
            import verify_experiment
            verify_experiment.run_verification(args.host, args.port, args.user, args.password, args.database,
                                               out_path=str(Path(args.out).with_name("verification.json")))
        except Exception as exc:  # e.g. no SELECT privilege on performance_schema
            print(f"Verification run skipped: {exc}")

    builder = Path(__file__).with_name("build_excel.py")
    if builder.exists():
        subprocess.run([sys.executable, str(builder), args.out], check=False)


if __name__ == "__main__":
    main()
