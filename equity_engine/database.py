import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), "equity_master.db")

def get_connection(db_path=DB_PATH):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn

def init_db(db_path=DB_PATH):
    conn = get_connection(db_path)
    cur = conn.cursor()
    cur.executescript('''
    CREATE TABLE IF NOT EXISTS companies (
        symbol TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        industry TEXT,
        bse_code TEXT,
        screener_url TEXT,
        is_financial BOOLEAN DEFAULT 0,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );

    CREATE TABLE IF NOT EXISTS daily_market_quotes (
        symbol TEXT,
        date TEXT,
        cmp REAL,
        market_cap REAL,
        pe REAL,
        book_value REAL,
        div_yield REAL,
        roce REAL,
        roe REAL,
        face_value REAL,
        gsec_10y REAL DEFAULT 7.0,
        PRIMARY KEY(symbol, date),
        FOREIGN KEY(symbol) REFERENCES companies(symbol)
    );

    CREATE TABLE IF NOT EXISTS annual_financials (
        symbol TEXT,
        period TEXT,
        fiscal_year INTEGER,
        sales REAL,
        expenses REAL,
        operating_profit REAL,
        opm_pct REAL,
        other_income REAL,
        interest REAL,
        depreciation REAL,
        pbt REAL,
        tax_amount REAL,
        tax_pct REAL,
        pat REAL,
        eps REAL,
        dividend_amount REAL,
        div_payout_pct REAL,
        equity_capital REAL,
        reserves REAL,
        net_worth REAL,
        borrowings REAL,
        other_liabilities REAL,
        fixed_assets REAL,
        cwip REAL,
        investments REAL,
        other_assets REAL,
        cfo REAL,
        capex REAL,
        fcf REAL,
        inventory REAL,
        debtors REAL,
        PRIMARY KEY(symbol, period),
        FOREIGN KEY(symbol) REFERENCES companies(symbol)
    );

    CREATE TABLE IF NOT EXISTS quarterly_financials (
        symbol TEXT,
        period TEXT,
        quarter_date TEXT,
        sales REAL,
        expenses REAL,
        operating_profit REAL,
        opm_pct REAL,
        other_income REAL,
        interest REAL,
        depreciation REAL,
        pbt REAL,
        tax_pct REAL,
        pat REAL,
        eps REAL,
        PRIMARY KEY(symbol, period),
        FOREIGN KEY(symbol) REFERENCES companies(symbol)
    );

    CREATE TABLE IF NOT EXISTS shareholding (
        symbol TEXT,
        period TEXT,
        promoters_pct REAL,
        fiis_pct REAL,
        diis_pct REAL,
        public_pct REAL,
        others_pct REAL,
        PRIMARY KEY(symbol, period),
        FOREIGN KEY(symbol) REFERENCES companies(symbol)
    );

    CREATE TABLE IF NOT EXISTS analysis_results (
        symbol TEXT PRIMARY KEY,
        calculated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        quality_score REAL,
        max_quality REAL,
        valuation_score REAL,
        max_valuation REAL,
        total_score REAL,
        max_total REAL,
        score_pct REAL,
        parameters_json TEXT,
        FOREIGN KEY(symbol) REFERENCES companies(symbol)
    );
    ''')
    conn.commit()
    conn.close()

if __name__ == "__main__":
    init_db()
    print("Database initialized successfully at:", DB_PATH)
