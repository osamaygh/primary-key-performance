#!/usr/bin/env python3
"""
Verification run: follows the doctor's steps once more and checks everything
directly inside MySQL, using MySQL's own clock and counters.

  * SHOW CREATE TABLE / SHOW INDEX      -> proves one table has a PK and the other has none
  * COUNT(*), COUNT(DISTINCT id), MIN/MAX -> proves exactly 1,000,000 rows, then 1,220,000
  * checksum over all columns             -> proves both tables hold identical data
  * column statistics + sample rows       -> shows the data is random
  * performance_schema TIMER_WAIT         -> MySQL's own time for every SELECT and INSERT
  * performance_schema ROWS_EXAMINED      -> MySQL's own count of rows read by each SELECT

Needs SELECT privilege on performance_schema (root has it).
Usage: python3 verify_experiment.py --user root --password "" [--out verification.json]
"""
import argparse
import json
import random
import time
from pathlib import Path

import mysql.connector

import pk_experiment as px

PS_LAST = """SELECT EVENT_ID, SQL_TEXT, TIMER_WAIT, ROWS_EXAMINED, ROWS_SENT, NO_INDEX_USED
             FROM performance_schema.events_statements_history
             WHERE THREAD_ID = PS_CURRENT_THREAD_ID() AND SQL_TEXT LIKE 'SELECT * FROM data_%'
             ORDER BY EVENT_ID DESC LIMIT 1"""
PS_SUM = """SELECT EVENT_NAME, SUM_TIMER_WAIT, COUNT_STAR
            FROM performance_schema.events_statements_summary_by_thread_by_event_name
            WHERE THREAD_ID = PS_CURRENT_THREAD_ID()
              AND EVENT_NAME IN ('statement/sql/insert', 'statement/sql/commit')"""


def q(conn, sql, params=None, dict_=False):
    cur = conn.cursor(dictionary=dict_)
    cur.execute(sql, params or ())
    rows = cur.fetchall()
    cur.close()
    return rows


def server_insert_totals(conn):
    return {name: (int(t), int(n)) for name, t, n in q(conn, PS_SUM)}


def insert_with_server_time(conn, table, rows):
    before = server_insert_totals(conn)
    client_s = px.timed_insert(conn, table, rows)
    after = server_insert_totals(conn)
    server_ps = sum(after[k][0] - before.get(k, (0, 0))[0] for k in after)
    statements = sum(after[k][1] - before.get(k, (0, 0))[1] for k in after)
    return client_s, server_ps / 1e12, statements


def data_checks(conn, table):
    cnt, dist, mn, mx = q(conn, f"SELECT COUNT(*), COUNT(DISTINCT id), MIN(id), MAX(id) FROM {table}")[0]
    checksum = q(conn, f"""SELECT BIT_XOR(CRC32(CONCAT_WS('|', id, name, email, age, city, salary, created_at)))
                           FROM {table}""")[0][0]
    stats = q(conn, f"""SELECT MIN(age), MAX(age), ROUND(AVG(age),2), COUNT(DISTINCT city), COUNT(DISTINCT name),
                               ROUND(MIN(salary),2), ROUND(MAX(salary),2), ROUND(AVG(salary),2),
                               MIN(created_at), MAX(created_at) FROM {table}""")[0]
    keys = ["age_min", "age_max", "age_avg", "distinct_cities", "distinct_names",
            "salary_min", "salary_max", "salary_avg", "created_min", "created_max"]
    return {"table": table, "row_count": cnt, "distinct_ids": dist, "min_id": mn, "max_id": mx,
            "checksum": int(checksum), **{k: (float(v) if hasattr(v, "is_finite") else v) for k, v in zip(keys, stats)}}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=3306)
    ap.add_argument("--user", default="root")
    ap.add_argument("--password", default="")
    ap.add_argument("--database", default="pk_experiment")
    ap.add_argument("--seed", type=int, default=4040)
    ap.add_argument("--out", default="verification.json")
    a = ap.parse_args()
    run_verification(a.host, a.port, a.user, a.password, a.database, a.seed, a.out)


def run_verification(host, port, user, password, database, seed=4040, out_path="verification.json"):
    print("\n=== Verification run (checks inside MySQL) ===", flush=True)
    conn = mysql.connector.connect(host=host, port=port, user=user, password=password,
                                   database=database, autocommit=False)
    rng = random.Random(seed)
    out = {"env": px.environment(conn), "show_create": {}, "show_index": {}, "data_checks": [],
           "samples": [], "selects": [], "inserts": []}
    T = px.TABLES

    # STEP 1 - tables with and without primary key
    for label, table in T.items():
        q(conn, f"DROP TABLE IF EXISTS {table}")
        q(conn, px.ddl(table, label == "with_pk"))
    conn.commit()
    for label, table in T.items():
        out["show_create"][label] = q(conn, f"SHOW CREATE TABLE {table}")[0][1]
        out["show_index"][label] = [{"Key_name": r["Key_name"], "Column_name": r["Column_name"],
                                     "Non_unique": r["Non_unique"], "Index_type": r["Index_type"]}
                                    for r in q(conn, f"SHOW INDEX FROM {table}", dict_=True)]
    print("STEP 1  tables created:", {k: [i["Key_name"] for i in v] for k, v in out["show_index"].items()})

    # STEP 2 - 1,000,000 random rows (same rows into both tables)
    rows = px.random_rows(1, px.INITIAL_ROWS, rng)
    for label, table in T.items():
        c, s, n = insert_with_server_time(conn, table, rows)
        out["inserts"].append({"step": "Initial load", "table": label, "rows_inserted": len(rows),
                               "client_s": c, "mysql_s": s, "statements": n})
        print(f"STEP 2  {label:11s} {len(rows):,} rows  client {c:.2f} s  MySQL {s:.2f} s  ({n} statements)")
    del rows
    for label, table in T.items():
        out["data_checks"].append({"stage": "after step 2", **data_checks(conn, table)})
    ids = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 1000, 10000, 100000, 1000000]
    for label, table in T.items():
        for r in q(conn, f"SELECT * FROM {table} WHERE id IN ({','.join(map(str, ids))}) ORDER BY id", dict_=True):
            out["samples"].append({"table": label, **{k: str(v) for k, v in r.items()}})

    # STEP 3 - SELECT rows 10 .. 1,000,000; client time AND MySQL's own time + rows examined
    for target in px.SELECT_TARGETS:
        for label, table in T.items():
            for rep in range(1, px.SELECT_REPEATS + 1):
                ms = px.timed_select(conn, table, target)
                ev = q(conn, PS_LAST, dict_=True)[0]
                out["selects"].append({"table": label, "target_id": target, "repeat": rep, "client_ms": ms,
                                       "mysql_ms": int(ev["TIMER_WAIT"]) / 1e9, "rows_examined": ev["ROWS_EXAMINED"],
                                       "rows_sent": ev["ROWS_SENT"], "no_index_used": ev["NO_INDEX_USED"],
                                       "sql_text": ev["SQL_TEXT"]})
        m = {l: sorted(s["mysql_ms"] for s in out["selects"] if s["table"] == l and s["target_id"] == target)[2]
             for l in T}
        ex = {l: next(s["rows_examined"] for s in out["selects"] if s["table"] == l and s["target_id"] == target)
              for l in T}
        print(f"STEP 3  id={target:>9,}  MySQL time: PK {m['with_pk']:8.3f} ms ({ex['with_pk']:,} rows examined)"
              f"   no-PK {m['without_pk']:8.1f} ms ({ex['without_pk']:,} rows examined)")

    # STEP 4 + 5 - insert 20k, then 40k, 60k, 100k into the same tables
    next_id = px.INITIAL_ROWS + 1
    for batch in px.INSERT_BATCHES:
        rows = px.random_rows(next_id, batch, rng)
        for label, table in T.items():
            c, s, n = insert_with_server_time(conn, table, rows)
            cnt = q(conn, f"SELECT COUNT(*) FROM {table}")[0][0]
            out["inserts"].append({"step": f"Insert {batch:,}", "table": label, "rows_inserted": batch,
                                   "client_s": c, "mysql_s": s, "statements": n, "count_after": cnt})
            print(f"STEP 4/5 insert {batch:>7,} {label:11s} client {c:.3f} s  MySQL {s:.3f} s  -> COUNT(*) = {cnt:,}")
        next_id += batch
    for label, table in T.items():
        out["data_checks"].append({"stage": "after step 5", **data_checks(conn, table)})
    conn.close()

    Path(out_path).write_text(json.dumps(out, indent=1, default=str))
    dc = out["data_checks"]
    print("\nchecksums equal after step 2:", dc[0]["checksum"] == dc[1]["checksum"],
          "| after step 5:", dc[2]["checksum"] == dc[3]["checksum"])
    print("Saved ->", out_path)


if __name__ == "__main__":
    main()
