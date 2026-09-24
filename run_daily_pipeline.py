"""
run_daily_pipeline.py
=====================
Automated end-to-end daily execution pipeline for AMFI Mutual Fund Analytics.

Steps:
  1. Ingest latest statutory feed from AMFI (NAVAll.txt) -> updates amfi_nav_master_latest.csv
  2. Recompute inception-aware analytics & rolling returns -> updates scratch/full_master_engine_data.json
  3. Recompile offline interactive dashboard -> updates index.html
"""

import os
import sys
import logging
import datetime
import subprocess

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_FILE = os.path.join(BASE_DIR, "daily_pipeline.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)


def run_step(step_name: str, script_name: str) -> bool:
    script_path = os.path.join(BASE_DIR, script_name)
    logging.info(f"--- Starting: {step_name} ({script_name}) ---")
    start_t = datetime.datetime.now()
    
    try:
        res = subprocess.run(
            [sys.executable, script_path],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            check=True
        )
        elapsed = (datetime.datetime.now() - start_t).total_seconds()
        logging.info(f"[SUCCESS] {step_name} completed in {elapsed:.2f}s.")
        if res.stdout.strip():
            for line in res.stdout.strip().splitlines()[-5:]:
                logging.info(f"  > {line}")
        return True
    except subprocess.CalledProcessError as e:
        elapsed = (datetime.datetime.now() - start_t).total_seconds()
        logging.error(f"[FAILED] {step_name} failed after {elapsed:.2f}s with return code {e.returncode}:")
        if e.stderr:
            logging.error(e.stderr.strip())
        if e.stdout:
            logging.error(e.stdout.strip())
        return False


def main():
    pipeline_start = datetime.datetime.now()
    logging.info("==================================================================")
    logging.info(f"AMFI DAILY AUTOMATION PIPELINE STARTED at {pipeline_start.strftime('%Y-%m-%d %H:%M:%S')}")
    logging.info("==================================================================")

    # Step 1: Ingest daily AMFI NAVs
    ok = run_step("AMFI Statutory Feed Ingestion", "update_amfi_daily.py")
    if not ok:
        logging.warning("Feed ingestion had issues, proceeding with database sync...")

    # Step 2: Recompute rolling analytics engine data across 118 categories (Regular & Direct)
    ok = run_step("Rolling Analytics Engine Computation", "build_full_untruncated_engine.py")
    if not ok:
        logging.error("Pipeline aborted: failed to build engine data.")
        return 1

    # Step 3: Recompile standalone index.html dashboard
    ok = run_step("Dashboard Web Application Compilation", "build_complete_dashboard_with_recency.py")
    if not ok:
        logging.error("Pipeline aborted: failed to compile index.html.")
        return 1

    total_elapsed = (datetime.datetime.now() - pipeline_start).total_seconds()
    logging.info("==================================================================")
    logging.info(f"DAILY AUTOMATION PIPELINE FINISHED SUCCESSFULLY in {total_elapsed:.2f}s!")
    logging.info("==================================================================")
    return 0


if __name__ == "__main__":
    sys.exit(main())
