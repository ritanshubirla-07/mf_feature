import io
import os
import requests
import pandas as pd
from typing import List, Dict
from .database import get_connection, init_db

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
}

CACHE_DIR = os.path.join(os.path.dirname(__file__), "cache")
os.makedirs(CACHE_DIR, exist_ok=True)

NSE_EQUITY_URL = "https://archives.nseindia.com/content/equities/EQUITY_L.csv"
NIFTY_50_URL = "https://archives.nseindia.com/content/indices/ind_nifty50list.csv"
NIFTY_500_URL = "https://archives.nseindia.com/content/indices/ind_nifty500list.csv"

def get_nifty50_symbols() -> List[str]:
    """Fetch official list of Nifty 50 constituents."""
    try:
        r = requests.get(NIFTY_50_URL, headers=HEADERS, timeout=10)
        if r.status_code == 200:
            df = pd.read_csv(io.StringIO(r.text))
            return df['Symbol'].str.strip().tolist()
    except Exception as e:
        print(f"Warning: Failed to fetch online Nifty 50 ({e}), falling back to cache")
    return []

def get_nifty500_symbols() -> List[str]:
    """Fetch official list of Nifty 500 constituents."""
    try:
        r = requests.get(NIFTY_500_URL, headers=HEADERS, timeout=10)
        if r.status_code == 200:
            df = pd.read_csv(io.StringIO(r.text))
            return df['Symbol'].str.strip().tolist()
    except Exception as e:
        print(f"Warning: Failed to fetch online Nifty 500 ({e}), falling back to cache")
    return []

def get_all_nse_equities() -> pd.DataFrame:
    """Fetch master list of all active NSE listed equities."""
    cache_path = os.path.join(CACHE_DIR, "EQUITY_L.csv")
    try:
        r = requests.get(NSE_EQUITY_URL, headers=HEADERS, timeout=15)
        if r.status_code == 200:
            with open(cache_path, "w", encoding="utf-8") as f:
                f.write(r.text)
            df = pd.read_csv(io.StringIO(r.text))
        else:
            df = pd.read_csv(cache_path)
    except Exception:
        if os.path.exists(cache_path):
            df = pd.read_csv(cache_path)
        else:
            raise RuntimeError("Could not retrieve NSE equity master list")

    df.columns = [c.strip() for c in df.columns]
    # Filter for standard equity series (EQ)
    df = df[df['SERIES'].astype(str).str.strip() == 'EQ'].copy()
    df['SYMBOL'] = df['SYMBOL'].astype(str).str.strip()
    df['NAME OF COMPANY'] = df['NAME OF COMPANY'].astype(str).str.strip()
    return df

def populate_universe_master(db_path=None) -> int:
    """
    Seeds all ~2,000+ active NSE equities into the companies master table.
    Preserves existing metadata if already populated from Screener.
    """
    init_db(db_path) if db_path else init_db()
    conn = get_connection(db_path) if db_path else get_connection()
    cur = conn.cursor()

    df = get_all_nse_equities()
    count = 0
    for _, row in df.iterrows():
        sym = row['SYMBOL']
        name = row['NAME OF COMPANY']
        isin = row.get('ISIN NUMBER', '')
        listing_date = row.get('DATE OF LISTING', '')

        # Insert if not exists, or update isin/listing_date without overwriting Screener industry
        cur.execute('''
        INSERT INTO companies (symbol, name, isin, listing_date, updated_at)
        VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(symbol) DO UPDATE SET
            isin=excluded.isin,
            listing_date=excluded.listing_date
        ''', (sym, name, str(isin).strip(), str(listing_date).strip()))
        count += 1

    conn.commit()
    conn.close()
    return count

if __name__ == "__main__":
    print("Fetching universe and seeding companies master table...")
    c = populate_universe_master()
    print(f"Seeded {c} active NSE companies into database!")
    n50 = get_nifty50_symbols()
    print(f"Nifty 50 count: {len(n50)}")
    n500 = get_nifty500_symbols()
    print(f"Nifty 500 count: {len(n500)}")
