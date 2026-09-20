# Flexi Cap Mutual Fund 3-Year Rolling Return Analysis

A Python-based financial analytics pipeline that tracks and analyzes **Flexi Cap Mutual Funds (Regular - Growth)** in India across monthly 3-year rolling windows from **2015 to 2026** (~105 windows) using public mutual fund data from [mfapi.in](https://api.mfapi.in).

---

## 🚀 Features

- **Automated Scheme Collection (`collect.py`):** Fetches ~38,000 schemes from `mfapi.in`, filters regular flexi cap growth schemes, and saves scheme codes.
- **Efficient History Caching:** Caches complete NAV history once, enabling instant offline computation across all ~105 rolling windows in seconds.
- **3-Year Rolling Window Evaluation (`window_analysis.py`):**
  - Evaluates monthly rolling windows (e.g., `2015-01-01` to `2018-01-01` through `2023-09-01` to `2026-09-01`).
  - Strict eligibility check: Inception date $\le$ Window Start Date, and active through Window End Date.
  - Robust holiday/weekend lookback: Automatically resolves non-trading dates to the most recent trading day close.
  - Mathematical 3-Year CAGR calculation:
    $$\text{CAGR} = \left(\frac{\text{NAV}_{\text{end}}}{\text{NAV}_{\text{start}}}\right)^{\frac{1}{3}} - 1$$
  - Identifies top quartile performers (`top_25_pct`) for each window.
- **Professional Excel Output:**
  - Native Excel **Outline Grouping (`+ / -`)**: Starts collapsed as a clean 105-row summary table showing high-level stats, quarter, best fund, and CAGR.
  - Clicking `+` unfolds the exact top 25% funds starting from `#2` with their CAGR.
  - Displays a complete **Consistency Summary Table** at the bottom of the sheet grouping total top 25% occurrences by fund name.

---

## 🛠️ Setup & Installation

### 1. Prerequisites
- Python 3.10+
- Git

### 2. Install Required Dependencies
```bash
pip install requests openpyxl urllib3
```

---

## 📖 Usage

### Step 1: Collect Flexi Cap Schemes
Fetches all matching schemes and stores them in `flexi_cap_growth_funds.csv`:
```bash
python collect.py
```

### Step 2: Run Rolling Window Analysis & Generate Excel Report
Fetches historical NAVs (cached to `mf_history_cache.json` on first run) and creates `flexi_cap_3yr_rolling_analysis.xlsx`:
```bash
python window_analysis.py
```

---

## 📁 Project Structure

```text
mf-feature/
├── collect.py              # Fetches and filters regular flexi cap growth schemes
├── cache_history.py        # Optional standalone utility to cache NAV histories
├── window_analysis.py      # Core rolling window engine and Excel generator
├── .gitignore              # Ignores all raw CSV, JSON, and XLSX data files
└── README.md               # Project documentation
```

---

## 🔒 Data Privacy & .gitignore
All generated datasets (`*.csv`, `*.json`) and Excel workbooks (`*.xlsx`) are excluded from version control via `.gitignore` to maintain a clean codebase.
