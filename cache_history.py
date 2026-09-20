import csv
import json
import time
import urllib3
import requests

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

INPUT_CSV = "flexi_cap_growth_funds.csv"
OUTPUT_CACHE = "mf_history_cache.json"

def main():
    with open(INPUT_CSV, "r", encoding="utf-8") as f:
        schemes = list(csv.DictReader(f))

    print(f"Starting fetch for {len(schemes)} schemes...")
    cached_data = {}

    for idx, s in enumerate(schemes, 1):
        code = s["schemeCode"]
        name = s["schemeName"]
        url = f"https://api.mfapi.in/mf/{code}"
        
        try:
            res = requests.get(url, verify=False, timeout=30)
            if res.status_code == 200:
                data = res.json()
                nav_list = data.get("data", [])
                meta = data.get("meta", {})
                cached_data[code] = {
                    "schemeCode": code,
                    "schemeName": name,
                    "meta": meta,
                    "data": nav_list
                }
                print(f"[{idx}/{len(schemes)}] {code} fetched ({len(nav_list)} records) - {name[:40]}")
            else:
                print(f"[{idx}/{len(schemes)}] {code} Failed with status {res.status_code}")
        except Exception as e:
            print(f"[{idx}/{len(schemes)}] {code} Error: {e}")

    with open(OUTPUT_CACHE, "w", encoding="utf-8") as f:
        json.dump(cached_data, f)

    print(f"\nAll data cached successfully to {OUTPUT_CACHE} ({len(cached_data)} schemes).")

if __name__ == "__main__":
    main()
