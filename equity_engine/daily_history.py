import datetime
import os
import sqlite3
import pandas as pd
import numpy as np
import yfinance as yf
from typing import List, Optional

from .database import get_connection, init_db

def backfill_stock_daily_history(symbols: List[str], period="5y", batch_size=50, db_path=None):
    """
    Fetches real exchange-settled historical daily OHLCV price candles 
    (5 to 10 years) and stores them in daily_market_quotes table.
    """
    init_db(db_path)
    conn = get_connection(db_path)
    cur = conn.cursor()

    # Get known symbols to preserve referential integrity
    cur.execute("SELECT symbol FROM companies")
    known = set(r[0] for r in cur.fetchall())

    valid_symbols = [s.strip().upper() for s in symbols if s.strip()]
    total = len(valid_symbols)
    print(f"\n=======================================================")
    print(f"  HISTORICAL DAILY PRICES BACKFILL ENGINE")
    print(f"=======================================================")
    print(f"Total Target Stocks: {total}")
    print(f"Historical Period: {period}")
    print(f"Batch Size: {batch_size}")
    print(f"=======================================================\n")

    total_inserted = 0

    for i in range(0, total, batch_size):
        chunk = valid_symbols[i:i+batch_size]
        # Append .NS suffix for NSE
        yf_tickers = [f"{s}.NS" for s in chunk]
        print(f"[{i+1:>3}-{min(i+len(chunk), total):>3}/{total}] Fetching {len(chunk)} tickers from exchange archives...")
        
        try:
            df = yf.download(yf_tickers, period=period, interval="1d", group_by="ticker", progress=False, auto_adjust=False)
            if df.empty:
                print(f"  -> No data returned for batch.")
                continue

            rows_to_insert = []

            for sym in chunk:
                # Ensure company is in master
                if sym not in known:
                    cur.execute("INSERT OR IGNORE INTO companies (symbol, name, updated_at) VALUES (?, ?, CURRENT_TIMESTAMP)", (sym, sym))
                    known.add(sym)

                sym_df = None
                if len(chunk) == 1:
                    sym_df = df
                else:
                    ticker_key = f"{sym}.NS"
                    if ticker_key in df.columns.levels[0]:
                        sym_df = df[ticker_key]

                if sym_df is not None and not sym_df.empty:
                    sym_df = sym_df.dropna(subset=['Close'])
                    for dt, row in sym_df.iterrows():
                        date_str = dt.strftime('%Y-%m-%d')
                        c_val = float(row.get('Close', 0.0))
                        o_val = float(row.get('Open', 0.0)) if not pd.isna(row.get('Open')) else c_val
                        h_val = float(row.get('High', 0.0)) if not pd.isna(row.get('High')) else c_val
                        l_val = float(row.get('Low', 0.0)) if not pd.isna(row.get('Low')) else c_val
                        v_val = int(row.get('Volume', 0)) if not pd.isna(row.get('Volume')) else 0

                        if c_val > 0:
                            rows_to_insert.append((
                                sym, date_str, c_val, o_val, h_val, l_val, v_val, 7.0
                            ))

            if rows_to_insert:
                cur.executemany('''
                INSERT INTO daily_market_quotes (
                    symbol, date, cmp, open_price, high_price, low_price, volume, gsec_10y
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(symbol, date) DO UPDATE SET
                    cmp=excluded.cmp,
                    open_price=excluded.open_price,
                    high_price=excluded.high_price,
                    low_price=excluded.low_price,
                    volume=excluded.volume
                ''', rows_to_insert)
                conn.commit()
                total_inserted += len(rows_to_insert)
                print(f"  -> Inserted {len(rows_to_insert)} daily price candles across {len(chunk)} stocks.")

        except Exception as e:
            print(f"  [ERROR] Batch download failed: {e}")

    conn.close()
    print(f"\n=======================================================")
    print(f"  DAILY PRICES BACKFILL COMPLETED: {total_inserted} total records added")
    print(f"=======================================================\n")
    return total_inserted

if __name__ == "__main__":
    from .universe import get_nifty50_symbols
    n50 = get_nifty50_symbols()
    print(f"Backfilling 5-year daily prices for Nifty 50 ({len(n50)} stocks)...")
    backfill_stock_daily_history(n50, period="5y", batch_size=25)
