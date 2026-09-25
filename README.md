# Indian Mutual Fund 3-Year Rolling Return Analytics & Dashboard

An institutional-grade mutual fund analytics engine and offline interactive dashboard covering all SEBI open-ended categories across 18 half-yearly 3-year rolling windows (2015 to 2026).

---

## 🚀 Key Features

1. **Standalone Interactive Dashboard (`index.html`)**:
   - Zero-dependency client-side web application. Open directly via browser (`file:///`) without CORS or server requirements.
   - **5 Dynamic Filter Pills**: Scheme Type (Open/Close/Interval), Category Group (Equity/Hybrid/Debt/Other), Sub-Category (118 categories), Option (Growth/IDCW/Bonus/Other), and **Plan Type (Regular / Direct / All Plans)**.
   - **Table 1: 3-Year Rolling Return Analysis**: Half-yearly windows with expandable Top 25% quartiles.
   - **Table 2: Consistency Summary & Recency Ranking**: Ranked by Composite Score (65% Consistency + 35% Recency) with an interactive 2D Quartile Tracking Matrix.
   - **Client-Side Excel Export**: Generates styled `.xlsx` reports with native Excel outline dropdowns `[+]`/`[-]` via embedded ExcelJS.

2. **Daily AMFI Statutory Ingestion Engine**:
   - Downloads official closing NAVs directly from AMFI (`NAVAll.txt`).
   - Incrementally appends to the historical fact table in O(1) memory and updates master metadata.

3. **Automated Daily Pipeline (`run_daily_pipeline.py`)**:
   - Single-command master pipeline: Ingestion → Analytics Recomputation → Dashboard Rebuild.
   - Pre-configured Windows Task Scheduler script (`setup_daily_cron.ps1`) for daily night updates (23:00 IST).

---

## 🛠️ Step-by-Step Setup (New Machine)

### Step 1: Clone the Repository
```bash
git clone https://github.com/ritanshubirla-07/mf_feature.git
cd mf_feature
```

### Step 2: Install Python Dependencies
```bash
pip install -r requirements.txt
```

### Step 3: Add Historical Fact Table (from Google Drive)
Download `amfi_nav_master.csv` from the provided Google Drive link and place it in the folder as:
```text
mf-feature/
└── amfi_nav_master.csv/
    └── amfi_nav_master.csv      # 3.8GB full historical tick table
```

### Step 4: Open & View the Dashboard
Simply double-click [`index.html`](index.html) or open it in any modern browser. It loads immediately with all 118 categories and Regular/Direct plans.

---

## 🔄 Daily Updates & Automation

To update the dashboard with fresh closing NAVs:

### Option A: One-Click Manual Update (Anytime)
Double-click **`run_daily_pipeline.bat`** (or run `python run_daily_pipeline.py`).
In ~15 seconds, it will:
1. Fetch today's AMFI feed
2. Update the historical database
3. Recompute rolling return analytics
4. Rebuild `index.html` with the latest data

### Option B: Automatic Daily Schedule (Hands-Free)
To have Windows automatically update the dashboard every night at **23:00 (11:00 PM IST)**:
```powershell
powershell -ExecutionPolicy Bypass -File setup_daily_cron.ps1
```
*(Runs automatically in the background even if missed during sleep)*

---

## 📁 Repository Structure

```text
mf-feature/
├── index.html                                  # Complete offline interactive dashboard
├── exceljs.min.js                              # Local library for client-side Excel generation
├── amfi_nav_master_latest.csv                  # Master dimension table (8,400+ active schemes)
├── mf_history_cache.json                       # 10-year historical NAV cache
├── fund_manager_master.csv                     # Fund manager directory & expense ratios
├── scratch/                                    # Pre-computed analytics engine JSON caches
├── requirements.txt                            # Python dependencies
├── run_daily_pipeline.py                       # Master daily automation pipeline
├── run_daily_pipeline.bat                      # Windows one-click batch launcher
├── setup_daily_cron.ps1                        # Windows scheduled task (cron) installer
├── update_amfi_daily.py                        # AMFI daily feed ingestion
├── build_full_untruncated_engine.py            # Inception-aware rolling returns analytics engine
├── build_complete_dashboard_with_recency.py    # Dashboard compiler
├── generate_all_category_excels.py             # Batch category Excel generator
└── README.md                                   # Documentation
```
