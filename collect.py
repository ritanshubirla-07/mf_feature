import csv
import json
import re
import urllib3
import requests

# 1. Fetch all mutual fund schemes from mfapi
url = "https://api.mfapi.in/mf"
try:
    response = requests.get(url, timeout=30)
except requests.exceptions.SSLError:
    # Fallback for environments with local SSL certificate issues
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    response = requests.get(url, verify=False, timeout=30)

data = response.json()

# 2. Filter for schemes having 'flexi cap' (or 'flexicap') and 'growth' in schemeName
pattern = re.compile(r"(?=.*flexi\s*cap)(?=.*regular)(?=.*growth)", re.IGNORECASE)
funds = [
    {
        "schemeCode": item["schemeCode"],
        "schemeName": item["schemeName"],
        "isinGrowth": item.get("isinGrowth"),
        "isinDivReinvestment": item.get("isinDivReinvestment"),
    }
    for item in data
    if pattern.search(item.get("schemeName", ""))
]

# 3. Save to CSV
with open("flexi_cap_growth_funds.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(
        f, fieldnames=["schemeCode", "schemeName", "isinGrowth", "isinDivReinvestment"]
    )
    writer.writeheader()
    writer.writerows(funds)

print(f"Collected {len(funds)} Flexi Cap Growth funds.")
print("Saved to flexi_cap_growth_funds.json and flexi_cap_growth_funds.csv")
