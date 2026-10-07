import argparse
import os
import sys
import time
import random
from datetime import datetime
from typing import List

from .database import get_connection, init_db
from .fetcher import fetch_and_store_company
from .calculator import run_vm_analysis
from .universe import get_nifty50_symbols, get_nifty500_symbols, populate_universe_master

def get_already_ingested_symbols(conn) -> set:
    """Returns set of symbols that already have complete annual financials."""
    cur = conn.cursor()
    cur.execute("SELECT DISTINCT symbol FROM annual_financials GROUP BY symbol HAVING count(*) >= 5")
    return set(r[0] for r in cur.fetchall())

def bulk_ingest(symbols: List[str], delay: float = 0.8, force: bool = False, max_count: int = None):
    init_db()
    conn = get_connection()
    already_done = set() if force else get_already_ingested_symbols(conn)
    conn.close()

    target_list = [s.strip().upper() for s in symbols if s.strip()]
    if not force:
        to_process = [s for s in target_list if s not in already_done]
    else:
        to_process = target_list

    if max_count and max_count > 0:
        to_process = to_process[:max_count]

    total = len(to_process)
    skipped = len(target_list) - total
    print(f"\n=======================================================")
    print(f"  BULK EQUITY INGESTER & VALUATION ENGINE")
    print(f"=======================================================")
    print(f"Total Target Stocks: {len(target_list)}")
    print(f"Already Synced (Skipped): {skipped}")
    print(f"To Process Now: {total}")
    print(f"Delay Between Requests: {delay}s (+ jitter)")
    print(f"=======================================================\n")

    if total == 0:
        print("All target stocks are already up to date in the database!")
        return

    success_count = 0
    fail_count = 0
    start_time = time.time()

    for idx, sym in enumerate(to_process, 1):
        item_start = time.time()
        try:
            # 1. Fetch & Store
            info = fetch_and_store_company(sym)
            
            # 2. Run 34-parameter analysis
            analysis = run_vm_analysis(sym)
            q_score = analysis.get('quality_score', 0)
            max_q = analysis.get('max_quality', 57)
            v_score = analysis.get('valuation_score', 0)
            max_v = analysis.get('max_valuation', 10)
            tot_pct = analysis.get('score_pct', 0)

            success_count += 1
            elapsed = time.time() - start_time
            rate = idx / elapsed if elapsed > 0 else 1
            remaining = total - idx
            eta_sec = remaining / rate if rate > 0 else 0
            eta_str = f"{int(eta_sec // 60)}m {int(eta_sec % 60)}s"

            print(f"[{idx:3d}/{total:3d}] {sym:<12} | {info['name'][:24]:<24} | Q: {q_score:4.1f}/{max_q:.0f} | V: {v_score:3.1f}/{max_v:.0f} ({tot_pct:4.1f}%) | ETA: {eta_str}")

        except Exception as e:
            fail_count += 1
            print(f"[{idx:3d}/{total:3d}] {sym:<12} | ERROR: {e}")

        # Rate-limiting delay with jitter
        if idx < total:
            sleep_time = delay + random.uniform(0.1, 0.4)
            time.sleep(sleep_time)

    total_time = time.time() - start_time
    print(f"\n=======================================================")
    print(f"  BULK INGESTION COMPLETED in {int(total_time // 60)}m {int(total_time % 60)}s")
    print(f"  Success: {success_count} | Failures: {fail_count} | Total Skipped: {skipped}")
    print(f"=======================================================\n")

def main():
    parser = argparse.ArgumentParser(description="Bulk Ingest Financial Statements for Indian Equities")
    parser.add_argument("--tier", choices=["nifty50", "nifty500", "all"], default="nifty50",
                        help="Universe tier to ingest (default: nifty50)")
    parser.add_argument("--symbols", type=str, help="Comma-separated custom symbols (e.g. RELIANCE,TCS,INFY)")
    parser.add_argument("--limit", type=int, default=None, help="Max number of companies to ingest")
    parser.add_argument("--delay", type=float, default=0.8, help="Delay in seconds between requests (default: 0.8)")
    parser.add_argument("--force", action="store_true", help="Re-fetch even if already present in database")
    args = parser.parse_args()

    # Ensure companies master is seeded
    populate_universe_master()

    if args.symbols:
        symbols = [s.strip().upper() for s in args.symbols.split(",") if s.strip()]
    elif args.tier == "nifty50":
        symbols = get_nifty50_symbols()
    elif args.tier == "nifty500":
        symbols = get_nifty500_symbols()
    elif args.tier == "all":
        conn = get_connection()
        symbols = [r[0] for r in conn.execute("SELECT symbol FROM companies ORDER BY symbol").fetchall()]
        conn.close()

    bulk_ingest(symbols, delay=args.delay, force=args.force, max_count=args.limit)

if __name__ == "__main__":
    main()
