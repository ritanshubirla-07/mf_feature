import sqlite3
import datetime
import os
import sys
import argparse
from equity_engine.database import get_connection, init_db, DB_PATH
from equity_engine.fetcher import fetch_and_store_company
from equity_engine.calculator import run_vm_analysis
from equity_engine.template_generator import generate_identical_excel

LOG_FILE = os.path.join(os.path.dirname(__file__), "daily_equity_update.log")

def log(msg):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted = f"[{ts}] {msg}"
    print(formatted)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(formatted + "\n")

from equity_engine.bhavcopy import sync_latest_bhavcopy

def run_daily_append(symbols=None, generate_excels=False, sync_bhav=True, db_path=DB_PATH):
    init_db(db_path)
    
    # 1. High-speed Bulk NSE Bhavcopy update (All 2,000+ stocks in ~2 seconds)
    if sync_bhav:
        log("=== SYNCING OFFICIAL NSE BHAVCOPY FOR ENTIRE MARKET ===")
        bhav_date, bhav_count = sync_latest_bhavcopy(db_path=db_path)
        if bhav_date:
            log(f"  -> Successfully updated daily quotes for {bhav_count} stocks on {bhav_date}")
        else:
            log("  -> Bhavcopy up to date or market holiday.")

    # 2. Financial statement & ratio updates for tracked fundamental universe
    conn = get_connection(db_path)
    if not symbols:
        # Update companies that have fundamental history tracked
        rows = conn.execute("SELECT DISTINCT symbol FROM annual_financials ORDER BY symbol").fetchall()
        symbols = [r['symbol'] for r in rows]
    conn.close()

    if not symbols:
        log("No fundamental companies found in database to update.")
        return

    today_str = datetime.date.today().strftime("%Y-%m-%d")
    log(f"=== UPDATING FUNDAMENTALS & EVALUATION SCORES FOR {len(symbols)} TRACKED STOCKS ({today_str}) ===")

    success_count = 0
    fail_count = 0

    for sym in symbols:
        try:
            # 1. Fetch latest numbers (market quote + any new quarterly/annual filings)
            comp_info = fetch_and_store_company(sym, db_path=db_path)
            
            # 2. Recalculate 34 parameters
            analysis = run_vm_analysis(sym, db_path=db_path)
            
            log(f"  -> Updated {sym}: CMP=Rs. {comp_info['cmp']} | Score={analysis['total_score']}/{analysis['max_total']} ({analysis['score_pct']}%)")

            # 3. Optionally regenerate identical Excel file
            if generate_excels:
                out_path = os.path.join(os.path.dirname(__file__), f"{sym}_VM_Analysis.xlsx")
                generate_identical_excel(sym, output_path=out_path, db_path=db_path)
                log(f"  -> Generated Excel: {out_path}")

            success_count += 1
        except Exception as e:
            log(f"  [ERROR] Failed to update {sym}: {e}")
            fail_count += 1

    log(f"=== DAILY APPEND COMPLETED: {success_count} succeeded, {fail_count} failed ===")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Daily Equity Market Data Append Script")
    parser.add_argument("--symbol", "-s", help="Specific stock symbol to append (default: all in DB)", default=None)
    parser.add_argument("--generate-excel", "-g", action="store_true", help="Also generate identical Excel files after append")
    parser.add_argument("--no-bhav", action="store_true", help="Skip NSE Bhavcopy market quote sync")
    args = parser.parse_args()

    sym_list = [args.symbol.upper()] if args.symbol else None
    run_daily_append(symbols=sym_list, generate_excels=args.generate_excel, sync_bhav=not args.no_bhav)
