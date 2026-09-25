"""
update_amfi_daily.py
====================
Automated daily Indian Mutual Fund NAV ingestion engine.
Directly ingests the statutory AMFI feed (NAVAll.txt) without any third-party APIs.

Features:
  1. Idempotency: Checks the database latest date; will never duplicate daily rows.
  2. Efficiency: Direct stream-append to the 3.8GB master fact table in O(1) memory.
  3. Master Dimension Sync: Updates amfi_nav_master_latest.csv atomically with backup.
  4. Attribute Integrity: Preserves canonical SEBI categories, scheme plans, options,
     AMCs, and statutory quarterly AAUM (Average AUM) disclosures.
  5. Logging: Logs operational details to console and daily_update.log.
"""

import os
import re
import csv
import ssl
import sys
import shutil
import logging
import datetime
import urllib.request
import pandas as pd
import numpy as np

# Paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def get_fact_table_path() -> str:
    """Finds amfi_nav_master.csv whether placed directly in root or inside a folder."""
    direct_file = os.path.join(BASE_DIR, "amfi_nav_master.csv")
    if os.path.isfile(direct_file):
        return direct_file
    nested_file = os.path.join(BASE_DIR, "amfi_nav_master.csv", "amfi_nav_master.csv")
    if os.path.isfile(nested_file):
        return nested_file
    return direct_file

FACT_TABLE_PATH = get_fact_table_path()
DIM_TABLE_PATH = os.path.join(BASE_DIR, "amfi_nav_master_latest.csv")
DIM_BACKUP_PATH = os.path.join(BASE_DIR, "amfi_nav_master_latest.csv.bak")
LOG_FILE_PATH = os.path.join(BASE_DIR, "daily_update.log")

AMFI_URL = "https://www.amfiindia.com/spages/NAVAll.txt"

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE_PATH, encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)

CLEAN_GRP_MAP = {
    'debt scheme': 'Debt Scheme',
    'income': 'Debt Scheme',
    'income/debt oriented schemes': 'Debt Scheme',
    'gilt': 'Debt Scheme',
    'floating rate': 'Debt Scheme',
    'liquid': 'Debt Scheme',
    'money market': 'Debt Scheme',
    'equity scheme': 'Equity Scheme',
    'equity schemes': 'Equity Scheme',
    'growth': 'Equity Scheme',
    'elss': 'Equity Scheme',
    'hybrid scheme': 'Hybrid Scheme',
    'hybrid schemes': 'Hybrid Scheme',
    'balanced': 'Hybrid Scheme',
    'other scheme': 'Other Scheme',
    'index funds': 'Other Scheme',
    'exchange traded funds (etfs)': 'Other Scheme',
    'fund of funds': 'Other Scheme',
    'fund of funds scheme (domestic)': 'Other Scheme',
    'overseas fund of funds': 'Other Scheme',
    'solution oriented scheme': 'Other Scheme',
    'solution oriented schemes **': 'Other Scheme',
    'children’s fund': 'Other Scheme',
    'children?s fund': 'Other Scheme',
    'life cycle funds': 'Other Scheme',
    'gold etfs': 'Other Scheme',
    'other etfs': 'Other Scheme'
}


def parse_plan(plan_str: str, name_str: str) -> str:
    combined = (plan_str + " " + name_str).lower()
    if "direct" in combined:
        return "Direct"
    elif "regular" in combined:
        return "Regular"
    return "Other"


def parse_option(opt_str: str, name_str: str) -> str:
    combined = (opt_str + " " + name_str).lower()
    if "bonus" in combined:
        return "Bonus"
    elif any(k in combined for k in ["idcw", "dividend", "income distribution", "payout", "reinvest", "re-invest"]):
        return "IDCW"
    elif "growth" in combined:
        return "Growth"
    return "Other"


def fetch_amfi_feed(url: str = AMFI_URL) -> str:
    """Downloads NAVAll.txt from AMFI with robust SSL context."""
    logging.info(f"Downloading daily AMFI statutory feed from {url}...")
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    
    req = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Antigravity-MFPipeline/1.0"}
    )
    with urllib.request.urlopen(req, context=ctx, timeout=45) as resp:
        content = resp.read().decode("utf-8", errors="ignore")
    logging.info(f"Downloaded AMFI feed successfully ({len(content) / 1024 / 1024:.2f} MB).")
    return content


def parse_feed_content(content: str):
    """Parses AMFI text content into structured records and finds the primary NAV date."""
    lines = content.splitlines()
    records = []
    
    curr_type = ""
    curr_cat = ""
    curr_grp = ""
    curr_sub = ""
    curr_amc = ""
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        if ";" in line:
            parts = line.split(";")
            if len(parts) >= 8 and parts[0].isdigit():
                code = int(parts[0])
                isin1 = parts[1].strip() if parts[1] != "-" else ""
                isin2 = parts[2].strip() if parts[2] != "-" else ""
                name = parts[3].strip()
                plan_raw = parts[4].strip()
                opt_raw = parts[5].strip()
                nav_str = parts[6].strip()
                dt_raw = parts[7].strip()
                
                try:
                    dt_iso = datetime.datetime.strptime(dt_raw, "%d-%b-%Y").strftime("%Y-%m-%d")
                except Exception:
                    dt_iso = dt_raw
                    
                try:
                    nav_val = float(nav_str)
                except Exception:
                    nav_val = np.nan
                    
                records.append({
                    "scheme_code": code,
                    "isin": isin1,
                    "isin2": isin2,
                    "scheme_name": name,
                    "amc": curr_amc,
                    "scheme_type": curr_type,
                    "category": curr_cat,
                    "category_group": curr_grp,
                    "category_sub": curr_sub,
                    "category_group_clean": CLEAN_GRP_MAP.get(curr_grp.lower(), "Other Scheme"),
                    "scheme_plan": parse_plan(plan_raw, name),
                    "scheme_option": parse_option(opt_raw, name),
                    "nav": nav_val,
                    "nav_date": dt_iso
                })
        else:
            m = re.match(
                r'^(Open Ended Schemes|Close Ended Schemes|Interval Fund Schemes|Interval Fund)\s*\((.*?)\)',
                line,
                re.IGNORECASE
            )
            if m:
                curr_type = m.group(1).strip()
                if curr_type == "Interval Fund":
                    curr_type = "Interval Fund Schemes"
                cat_body = m.group(2).strip()
                if cat_body.endswith(")"):
                    cat_body = cat_body[:-1].strip()
                curr_cat = cat_body
                if " - " in cat_body:
                    grp, sub = cat_body.split(" - ", 1)
                    curr_grp = grp.strip()
                    curr_sub = sub.strip()
                else:
                    curr_grp = cat_body
                    curr_sub = ""
            elif "Mutual Fund" in line or "Asset Management" in line:
                curr_amc = line.strip()
                
    df_feed = pd.DataFrame(records)
    logging.info(f"Parsed {len(df_feed):,} schemes from AMFI feed.")
    return df_feed


def run_daily_update():
    start_time = datetime.datetime.now()
    logging.info("=" * 60)
    logging.info("STARTING AMFI MUTUAL FUND DAILY INGESTION PIPELINE")
    logging.info("=" * 60)
    
    # 1. Fetch & Parse Feed
    raw_content = fetch_amfi_feed()
    df_feed = parse_feed_content(raw_content)
    
    if df_feed.empty:
        logging.error("AMFI feed parsing returned 0 records. Aborting update.")
        return False
        
    latest_feed_date = df_feed["nav_date"].max()
    logging.info(f"Latest NAV date present in AMFI feed: {latest_feed_date}")
    
    # 2. Inspect Current Master Dimension Table
    if not os.path.exists(DIM_TABLE_PATH):
        logging.error(f"Master dimension file not found at {DIM_TABLE_PATH}")
        return False
        
    df_latest = pd.read_csv(DIM_TABLE_PATH)
    current_db_date = df_latest["nav_date"].dropna().max()
    logging.info(f"Current latest date in master database: {current_db_date}")
    
    # Filter feed to only records belonging to the latest trading date
    df_today = df_feed[df_feed["nav_date"] == latest_feed_date].copy()
    logging.info(f"Schemes with NAV published for {latest_feed_date}: {len(df_today):,}")
    
    # Check if this date has already been appended to fact table
    needs_fact_append = (latest_feed_date > str(current_db_date))
    
    if not needs_fact_append:
        logging.info(f"Notice: Date {latest_feed_date} is already present in master database. Skipping fact table append to guarantee idempotency.")
    else:
        logging.info(f"New trading day detected ({latest_feed_date} > {current_db_date}). Appending to fact table...")
        
        # 3. Stream-Append to Fact Table (amfi_nav_master.csv)
        # Columns: scheme_code,date,nav,scheme_name,isin
        fact_path = get_fact_table_path()
        if os.path.exists(fact_path):
            append_count = 0
            ref_lookup = df_latest.set_index("scheme_code").to_dict("index")
            
            with open(fact_path, "a", newline="", encoding="utf-8") as f_out:
                writer = csv.writer(f_out)
                for idx, r in df_today.iterrows():
                    code = r["scheme_code"]
                    nav_val = r["nav"]
                    if pd.isna(nav_val):
                        continue
                        
                    # Use canonical scheme_name from DB if present
                    if code in ref_lookup:
                        name = ref_lookup[code].get("scheme_name") or r["scheme_name"]
                        isin = ref_lookup[code].get("isin") or r["isin"]
                    else:
                        name = r["scheme_name"]
                    isin = r["isin"]
                        
                    writer.writerow([code, latest_feed_date, f"{nav_val:.4f}", name, isin])
                    append_count += 1
                    
            logging.info(f"[SUCCESS] Appended {append_count:,} records for {latest_feed_date} to {fact_path}.")
        else:
            logging.warning(f"Fact table not found at {fact_path}; skipped append.")

    # 4. Synchronize Master Dimension Table (amfi_nav_master_latest.csv)
    logging.info("Updating master dimension table (amfi_nav_master_latest.csv)...")
    
    # Create safety backup first
    shutil.copy2(DIM_TABLE_PATH, DIM_BACKUP_PATH)
    logging.info(f"Created safety backup at {DIM_BACKUP_PATH}")
    
    ref_lookup = df_latest.set_index("scheme_code").to_dict("index")
    updated_records = []
    feed_codes = set()
    
    # Update existing schemes from feed
    for idx, r in df_feed.iterrows():
        code = r["scheme_code"]
        feed_codes.add(code)
        
        if code in ref_lookup:
            ref = ref_lookup[code]
            # Preserve canonical DB attributes
            scheme_name = ref.get("scheme_name") if pd.notna(ref.get("scheme_name")) else r["scheme_name"]
            isin = ref.get("isin") if (pd.notna(ref.get("isin")) and ref.get("isin") != "") else r["isin"]
            isin2 = ref.get("isin2") if (pd.notna(ref.get("isin2")) and ref.get("isin2") != "") else r["isin2"]
            amc = ref.get("amc") if pd.notna(ref.get("amc")) else r["amc"]
            scheme_type = ref.get("scheme_type") if pd.notna(ref.get("scheme_type")) else r["scheme_type"]
            category = ref.get("category") if pd.notna(ref.get("category")) else r["category"]
            category_sub = ref.get("category_sub") if pd.notna(ref.get("category_sub")) else r["category_sub"]
            category_group_clean = ref.get("category_group_clean") if pd.notna(ref.get("category_group_clean")) else r["category_group_clean"]
            category_group = ref.get("category_group") if pd.notna(ref.get("category_group")) else r["category_group"]
            scheme_plan = ref.get("scheme_plan") if pd.notna(ref.get("scheme_plan")) else r["scheme_plan"]
            scheme_option = ref.get("scheme_option") if pd.notna(ref.get("scheme_option")) else r["scheme_option"]
            first_date = ref.get("first_date")
            aaum_cr = ref.get("aaum_cr_quarterly_avg")
            aaum_qtr = ref.get("aaum_quarter")
            aaum_qtr_end = ref.get("aaum_quarter_end")
            
            nav_date = r["nav_date"]
            nav_val = r["nav"]
            last_date = nav_date if pd.notna(nav_date) else ref.get("last_date")
            is_active = (nav_date >= "2026-01-01") if pd.notna(nav_date) else ref.get("is_active", False)
            is_stale = False
            
            updated_records.append({
                "scheme_code": code,
                "isin": isin,
                "isin2": isin2,
                "scheme_name": scheme_name,
                "amc": amc,
                "fund_manager": ref.get("fund_manager"),
                "scheme_type": scheme_type,
                "category": category,
                "category_sub": category_sub,
                "category_group_clean": category_group_clean,
                "category_group": category_group,
                "scheme_plan": scheme_plan,
                "scheme_option": scheme_option,
                "first_date": first_date,
                "last_date": last_date,
                "is_active": is_active,
                "is_stale": is_stale,
                "aaum_cr_quarterly_avg": aaum_cr,
                "aaum_quarter": aaum_qtr,
                "aaum_quarter_end": aaum_qtr_end,
                "nav_date": nav_date,
                "nav": nav_val
            })
            del ref_lookup[code]
        else:
            # New Scheme / NFO
            updated_records.append({
                "scheme_code": code,
                "isin": r["isin"],
                "isin2": r["isin2"],
                "scheme_name": r["scheme_name"],
                "amc": r["amc"],
                "fund_manager": np.nan,
                "scheme_type": r["scheme_type"],
                "category": r["category"],
                "category_sub": r["category_sub"],
                "category_group_clean": r["category_group_clean"],
                "category_group": r["category_group"],
                "scheme_plan": r["scheme_plan"],
                "scheme_option": r["scheme_option"],
                "first_date": r["nav_date"],
                "last_date": r["nav_date"],
                "is_active": True,
                "is_stale": False,
                "aaum_cr_quarterly_avg": np.nan,
                "aaum_quarter": np.nan,
                "aaum_quarter_end": np.nan,
                "nav_date": r["nav_date"],
                "nav": r["nav"]
            })
            logging.info(f"Added brand new scheme: [{code}] {r['scheme_name']}")

    # Retain remaining inactive schemes that were not in today's feed
    for code, ref in ref_lookup.items():
        updated_records.append({
            "scheme_code": code,
            "isin": ref.get("isin"),
            "isin2": ref.get("isin2"),
            "scheme_name": ref.get("scheme_name"),
            "amc": ref.get("amc"),
            "fund_manager": ref.get("fund_manager"),
            "scheme_type": ref.get("scheme_type"),
            "category": ref.get("category"),
            "category_sub": ref.get("category_sub"),
            "category_group_clean": ref.get("category_group_clean"),
            "category_group": ref.get("category_group"),
            "scheme_plan": ref.get("scheme_plan"),
            "scheme_option": ref.get("scheme_option"),
            "first_date": ref.get("first_date"),
            "last_date": ref.get("last_date"),
            "is_active": False,
            "is_stale": ref.get("is_stale", False),
            "aaum_cr_quarterly_avg": ref.get("aaum_cr_quarterly_avg"),
            "aaum_quarter": ref.get("aaum_quarter"),
            "aaum_quarter_end": ref.get("aaum_quarter_end"),
            "nav_date": ref.get("nav_date"),
            "nav": ref.get("nav")
        })

    df_final = pd.DataFrame(updated_records).sort_values("scheme_code").reset_index(drop=True)
    
    # Atomic write via temp file
    temp_path = DIM_TABLE_PATH + ".tmp"
    df_final.to_csv(temp_path, index=False)
    shutil.move(temp_path, DIM_TABLE_PATH)
    
    elapsed = (datetime.datetime.now() - start_time).total_seconds()
    logging.info(f"[SUCCESS] Master dimension table updated successfully: {len(df_final):,} total schemes ({df_final['is_active'].sum():,} active).")
    logging.info(f"Pipeline completed in {elapsed:.2f} seconds.")
    logging.info("=" * 60)
    return True


if __name__ == "__main__":
    run_daily_update()
