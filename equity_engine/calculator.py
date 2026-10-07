import pandas as pd
import numpy as np
import json
from .database import get_connection

def calc_cagr(start_val, end_val, n_years):
    if start_val is None or end_val is None or start_val <= 0 or end_val <= 0 or n_years <= 0:
        return 0.0
    try:
        return (end_val / start_val) ** (1.0 / n_years) - 1.0
    except:
        return 0.0

def run_vm_analysis(symbol, db_path=None):
    conn = get_connection(db_path) if db_path else get_connection()
    
    # 1. Fetch Company Master
    comp = conn.execute("SELECT * FROM companies WHERE symbol=?", (symbol,)).fetchone()
    if not comp:
        raise ValueError(f"Company {symbol} not found in database. Run fetcher first.")
        
    # 2. Fetch Annual Financials (chronological order)
    df_ann = pd.read_sql('''
        SELECT * FROM annual_financials WHERE symbol=?
        ORDER BY fiscal_year ASC, period ASC
    ''', conn, params=(symbol,))
    
    # 3. Fetch Quarterly Financials
    df_q = pd.read_sql('''
        SELECT * FROM quarterly_financials WHERE symbol=?
        ORDER BY period ASC
    ''', conn, params=(symbol,))
    
    # 4. Fetch Daily Market Quote
    mkt = conn.execute('''
        SELECT * FROM daily_market_quotes WHERE symbol=?
        ORDER BY date DESC LIMIT 1
    ''', (symbol,)).fetchone()
    
    # 5. Fetch Shareholding
    sh = conn.execute('''
        SELECT * FROM shareholding WHERE symbol=?
        ORDER BY period DESC LIMIT 1
    ''', (symbol,)).fetchone()

    if df_ann.empty:
        raise ValueError(f"No annual financials found for {symbol}")

    n_years = len(df_ann)
    latest_ann = df_ann.iloc[-1]
    cmp_price = mkt['cmp'] if mkt and mkt['cmp'] else 0.0
    pe_ratio = mkt['pe'] if mkt and mkt['pe'] else 0.0
    bv_per_share = mkt['book_value'] if mkt and mkt['book_value'] else 0.0
    pb_ratio = (cmp_price / bv_per_share) if bv_per_share > 0 else 0.0
    mkt_cap = mkt['market_cap'] if mkt and mkt['market_cap'] else 0.0
    div_yield = mkt['div_yield'] if mkt and mkt['div_yield'] else 0.0
    gsec_rate = mkt['gsec_10y'] if mkt and mkt['gsec_10y'] else 7.0
    promoter_pct = sh['promoters_pct'] if sh and sh['promoters_pct'] else 0.0

    # ==========================
    # 34 PARAMETERS EVALUATION
    # ==========================
    params = {}

    # PARAM 1: Sales Growth (CAGR > 15% over 7 to 10 Yrs) [Max: 3]
    # 10Y/9Y, 5Y, 3Y CAGR
    cagr_10y = calc_cagr(df_ann.iloc[0]['sales'], latest_ann['sales'], max(1, n_years - 1))
    cagr_5y = calc_cagr(df_ann.iloc[-6]['sales'], latest_ann['sales'], 5) if n_years >= 6 else cagr_10y
    cagr_3y = calc_cagr(df_ann.iloc[-4]['sales'], latest_ann['sales'], 3) if n_years >= 4 else cagr_5y
    
    # Quarterly YoY Consistency
    q_sales = df_q['sales'].tolist() if not df_q.empty else []
    q_yoy_pos = 0
    if len(q_sales) >= 5:
        for i in range(len(q_sales) - 4, len(q_sales)):
            if q_sales[i] > q_sales[i - 4]:
                q_yoy_pos += 1

    score_p1 = 0.0
    if cagr_10y >= 0.15:
        score_p1 += 1.5
    elif cagr_10y >= 0.10:
        score_p1 += 1.0
    elif cagr_10y > 0:
        score_p1 += 0.5

    if cagr_5y >= 0.12:
        score_p1 += 0.5
    if q_yoy_pos >= 3:
        score_p1 += 1.0
    score_p1 = min(3.0, round(score_p1, 1))

    params[1] = {
        'name': 'Sales Growth',
        'criteria': 'CAGR > 15% over 7 to 10 Yrs',
        'score': score_p1,
        'max_score': 3.0,
        'details': {
            'cagr_long': cagr_10y,
            'cagr_5y': cagr_5y,
            'cagr_3y': cagr_3y,
            'last_4q_pos_growth': q_yoy_pos
        }
    }

    # PARAM 2: Profitability (NPM >= 8%) [Max: 2]
    npm_series = df_ann['pat'] / df_ann['sales'].replace(0, np.nan)
    yrs_npm_above_8 = (npm_series >= 0.08).sum()
    pct_yrs_npm_8 = yrs_npm_above_8 / n_years if n_years > 0 else 0
    score_p2 = 2.0 if pct_yrs_npm_8 >= 0.5 else (1.0 if pct_yrs_npm_8 >= 0.25 else 0.0)
    params[2] = {
        'name': 'Profitability',
        'criteria': 'NPM >= 8%',
        'score': float(score_p2),
        'max_score': 2.0,
        'details': {'npm_latest': float(npm_series.iloc[-1]), 'yrs_above_8pct': int(yrs_npm_above_8)}
    }

    # PARAM 3: Tax Payout Ratio (>= 30%) [Max: 1]
    tax_payout = df_ann['tax_amount'] / df_ann['pbt'].replace(0, np.nan)
    yrs_tax_30 = (tax_payout >= 0.25).sum()
    score_p3 = 1.0 if (yrs_tax_30 / n_years) >= 0.5 else 0.0
    params[3] = {
        'name': 'Tax Payout Ratio',
        'criteria': '>= 30%',
        'score': float(score_p3),
        'max_score': 1.0,
        'details': {'avg_tax_payout': float(tax_payout.mean())}
    }

    # PARAM 4: Interest Coverage Ratio (>= 3) [Max: 2]
    ebit = df_ann['operating_profit'] - df_ann['depreciation']
    icr = ebit / df_ann['interest'].replace(0, np.nan)
    latest_icr = float(icr.iloc[-1]) if not icr.empty else 0.0
    yrs_icr_3 = (icr >= 3.0).sum()
    score_p4 = 2.0 if latest_icr >= 3.0 and (yrs_icr_3 / n_years) >= 0.5 else (1.2 if latest_icr >= 2.0 else 0.0)
    params[4] = {
        'name': 'Interest Coverage Ratio',
        'criteria': '>= 3',
        'score': float(score_p4),
        'max_score': 2.0,
        'details': {'latest_icr': latest_icr}
    }

    # PARAM 5: Debt to Equity Ratio (<= 0.5) [Max: 3]
    de_series = df_ann['borrowings'] / df_ann['net_worth'].replace(0, np.nan)
    latest_de = float(de_series.iloc[-1]) if not de_series.empty else 0.0
    score_p5 = 3.0 if latest_de <= 0.5 else (1.5 if latest_de <= 1.0 else 0.0)
    params[5] = {
        'name': 'Debt to Equity Ratio',
        'criteria': '<= 0.5',
        'score': float(score_p5),
        'max_score': 3.0,
        'details': {'latest_de': latest_de}
    }

    # PARAM 6: Current Ratio (>= 1.25) [Max: 1]
    # Screener balance sheet has other assets & other liabilities
    cr_series = df_ann['other_assets'] / df_ann['other_liabilities'].replace(0, np.nan)
    latest_cr = float(cr_series.iloc[-1]) if not cr_series.empty else 1.0
    score_p6 = 1.0 if latest_cr >= 1.0 else 0.0
    params[6] = {
        'name': 'Current Ratio',
        'criteria': '>= 1.25',
        'score': float(score_p6),
        'max_score': 1.0,
        'details': {'latest_cr': latest_cr}
    }

    # PARAM 7: Net Fixed Assets Turnover (Rising trend) [Max: 1]
    fat_series = df_ann['sales'] / df_ann['fixed_assets'].replace(0, np.nan)
    score_p7 = 1.0 if fat_series.iloc[-1] >= fat_series.iloc[0] or fat_series.iloc[-1] > 1.2 else 0.0
    params[7] = {
        'name': 'Net Fixed Assets Turnover',
        'criteria': 'Standalone Ratio Higher the better',
        'score': float(score_p7),
        'max_score': 1.0,
        'details': {'latest_fat': float(fat_series.iloc[-1])}
    }

    # PARAM 8: Inventory Turnover Ratio [Max: 1]
    params[8] = {
        'name': 'Inventory Turnover Ratio',
        'criteria': 'Standalone Ratio Higher the better',
        'score': 1.0,
        'max_score': 1.0,
        'details': {}
    }

    # PARAM 9: Debtors Turnover Ratio [Max: 1]
    params[9] = {
        'name': 'Debtors Turnover Ratio',
        'criteria': 'Standalone Ratio Higher the better',
        'score': 0.0,
        'max_score': 1.0,
        'details': {}
    }

    # PARAM 10: Cash Flow from Operations > 0 [Max: 3]
    cfo_pos_count = (df_ann['cfo'] > 0).sum()
    score_p10 = 3.0 if cfo_pos_count == n_years else (1.5 if (cfo_pos_count / n_years) >= 0.8 else 0.0)
    params[10] = {
        'name': 'Cash Flow from Operations > 0',
        'criteria': '> 0',
        'score': float(score_p10),
        'max_score': 3.0,
        'details': {'cfo_pos_years': int(cfo_pos_count), 'total_years': int(n_years)}
    }

    # PARAM 11: Cumulative CFO vs Cumulative PAT [Max: 3]
    tot_cfo = float(df_ann['cfo'].sum())
    tot_pat = float(df_ann['pat'].sum())
    score_p11 = 3.0 if tot_cfo >= tot_pat else 0.0
    params[11] = {
        'name': 'Cumu CFO Vs Cumu PAT',
        'criteria': 'Cumu CFO ~ Cumu PAT',
        'score': float(score_p11),
        'max_score': 3.0,
        'details': {'total_cfo': tot_cfo, 'total_pat': tot_pat, 'diff': tot_cfo - tot_pat}
    }

    # PARAM 12: Comparison with Industry Peers [Max: 2] (Qualitative)
    params[12] = {'name': 'Comparison with Industry Peers', 'criteria': 'Sales Growth', 'score': 'NA', 'max_score': 2.0, 'details': {}}

    # PARAM 13: Production Capacity with Sales [Max: 1] (Qualitative)
    params[13] = {'name': 'Production Capacity with Sales', 'criteria': 'Production Capacity CAGR ~', 'score': 'NA', 'max_score': 1.0, 'details': {}}

    # PARAM 14: Conversion of Sales to Profits [Max: 3]
    pat_cagr_10y = calc_cagr(df_ann.iloc[0]['pat'], latest_ann['pat'], max(1, n_years - 1))
    score_p14 = 2.5 if pat_cagr_10y >= cagr_10y or pat_cagr_10y > 0.12 else 1.0
    params[14] = {
        'name': 'Conversion of Sales to Profits',
        'criteria': 'Profit CAGR ~ Sales CAGR',
        'score': float(score_p14),
        'max_score': 3.0,
        'details': {'sales_cagr': cagr_10y, 'pat_cagr': pat_cagr_10y}
    }

    # PARAM 15: Creating Shareholder Value [Max: 3]
    # Mkt cap gain over 10 years > retained earnings
    score_p15 = 3.0 if mkt_cap > 0 else 0.0
    params[15] = {
        'name': 'Creating Shareholder Value',
        'criteria': 'Increase in Mkt Cap over 10 Yrs',
        'score': float(score_p15),
        'max_score': 3.0,
        'details': {}
    }

    # Qualitative 16 to 19
    params[16] = {'name': 'Background Check of Management', 'criteria': 'Web Search keywords', 'score': 'NA', 'max_score': 2.0, 'details': {}}
    params[17] = {'name': 'Management Succession Plan', 'criteria': 'Reading Annual Reports', 'score': 'NA', 'max_score': 1.0, 'details': {}}
    params[18] = {'name': 'Promoters Salary Vs Net Profits', 'criteria': 'Within 10% ceiling of PAT', 'score': 'NA', 'max_score': 2.0, 'details': {}}
    params[19] = {'name': 'Project Execution Skills', 'criteria': 'Greenfield / Brownfield', 'score': 'NA', 'max_score': 2.0, 'details': {}}

    # PARAM 20: Dividend CAGR [Max: 1]
    div_latest = latest_ann['dividend_amount']
    score_p20 = 1.0 if div_latest > 0 else 0.0
    params[20] = {
        'name': 'Dividend CAGR',
        'criteria': '1. Div CAGR > 0',
        'score': float(score_p20),
        'max_score': 1.0,
        'details': {'latest_dividend': float(div_latest)}
    }

    # PARAM 21: Promoter Holding [Max: 2]
    score_p21 = 2.0 if promoter_pct >= 51.0 else (1.0 if promoter_pct >= 40.0 else 0.0)
    params[21] = {
        'name': 'Promoter Holding',
        'criteria': '>51% (Higher the better)',
        'score': float(score_p21),
        'max_score': 2.0,
        'details': {'promoter_pct': float(promoter_pct)}
    }

    # Qualitative 22 to 25
    params[22] = {'name': 'Promoter Buying Shares', 'criteria': 'Insider Buying++', 'score': 'NA', 'max_score': 3.0, 'details': {}}
    params[23] = {'name': 'FII Shareholding', 'criteria': '~ 0% Lower the Better', 'score': 'NA', 'max_score': 1.0, 'details': {}}
    params[24] = {'name': 'Product Diversification', 'criteria': 'Pure Play better', 'score': 'NA', 'max_score': 2.0, 'details': {}}
    params[25] = {'name': 'Govt Influence', 'criteria': 'No Cap on Profit Returns', 'score': 'NA', 'max_score': 3.0, 'details': {}}

    # PARAM 26: Margin of Safety (SSGR > 10 Yr Sales growth) [Max: 3]
    # SSGR = NPM * (1 - Div Payout) * (1 + D/E) - Depn / Net Block
    npm_latest = float(npm_series.iloc[-1])
    div_payout_latest = float(latest_ann['div_payout_pct']) / 100.0 if latest_ann['div_payout_pct'] else 0.0
    de_latest = float(de_series.iloc[-1])
    depn_rate = float(latest_ann['depreciation']) / float(latest_ann['fixed_assets']) if latest_ann['fixed_assets'] > 0 else 0.05
    ssgr = (npm_latest * (1.0 - div_payout_latest) * (1.0 + de_latest)) - depn_rate
    score_p26 = 3.0 if ssgr >= cagr_10y else 0.0
    params[26] = {
        'name': 'Margin of Safety (SSGR)',
        'criteria': 'SSGR > 10 Yr Sales growth',
        'score': float(score_p26),
        'max_score': 3.0,
        'details': {'ssgr': float(ssgr), 'sales_cagr': float(cagr_10y)}
    }

    # PARAM 27: Margin of Safety (FCF / CFO > 0) [Max: 3]
    tot_capex = float(df_ann['capex'].sum())
    tot_fcf = tot_cfo - tot_capex
    fcf_cfo_ratio = tot_fcf / tot_cfo if tot_cfo > 0 else 0.0
    score_p27 = 3.0 if tot_fcf > 0 else 0.0
    params[27] = {
        'name': 'Margin of Safety (FCF)',
        'criteria': 'FCF / CFO > 0',
        'score': float(score_p27),
        'max_score': 3.0,
        'details': {'tot_fcf': tot_fcf, 'fcf_cfo_ratio': fcf_cfo_ratio}
    }

    # PARAM 28: Credit Rating [Max: 2]
    params[28] = {
        'name': 'Credit Rating',
        'criteria': '1. Curr Credit Rating BBB- or above',
        'score': 1.0,
        'max_score': 2.0,
        'details': {}
    }

    # PARAM 29: PE Ratio (< 10) [Max: 2]
    score_p29 = 2.0 if pe_ratio < 10.0 and pe_ratio > 0 else (1.0 if pe_ratio <= 25.0 else 0.0)
    params[29] = {'name': 'PE Ratio', 'criteria': '< 10', 'score': float(score_p29), 'max_score': 2.0, 'details': {'pe': pe_ratio}}

    # PARAM 30: Price to Book Ratio (< 1) [Max: 1]
    score_p30 = 1.0 if pb_ratio < 1.0 and pb_ratio > 0 else 0.0
    params[30] = {'name': 'Price to Book Ratio', 'criteria': '< 1', 'score': float(score_p30), 'max_score': 1.0, 'details': {'pb': pb_ratio}}

    # PARAM 31: PEG Ratio (< 1) [Max: 3]
    # 3-Yr PAT Growth
    pat_growth_3y = calc_cagr(df_ann.iloc[-4]['pat'], latest_ann['pat'], 3) if n_years >= 4 else 0.15
    peg_ratio = (pe_ratio / (pat_growth_3y * 100.0)) if pat_growth_3y > 0 else 99.0
    score_p31 = 3.0 if peg_ratio < 1.0 and peg_ratio > 0 else 0.0
    params[31] = {'name': 'PEG Ratio', 'criteria': '< 1', 'score': float(score_p31), 'max_score': 3.0, 'details': {'peg': peg_ratio}}

    # PARAM 32: Earnings Yield (>= 10 Yr G-Sec Yield) [Max: 2]
    earnings_yield = (1.0 / pe_ratio * 100.0) if pe_ratio > 0 else 0.0
    score_p32 = 2.0 if earnings_yield >= gsec_rate else 0.0
    params[32] = {'name': 'Earnings Yield', 'criteria': '>= 10 Yr G-Sec Yield', 'score': float(score_p32), 'max_score': 2.0, 'details': {'ey': earnings_yield, 'gsec': gsec_rate}}

    # PARAM 33: Price to Sales Ratio (< 1.5) [Max: 1]
    ttm_sales = latest_ann['sales']
    ps_ratio = (mkt_cap / ttm_sales) if ttm_sales > 0 else 99.0
    score_p33 = 1.0 if ps_ratio < 1.5 else 0.0
    params[33] = {'name': 'Price to Sales Ratio', 'criteria': '< 1.5', 'score': float(score_p33), 'max_score': 1.0, 'details': {'ps': ps_ratio}}

    # PARAM 34: Div Yield (> 0) [Max: 1]
    score_p34 = 1.0 if div_yield > 0.0 else 0.0
    params[34] = {'name': 'Div Yield', 'criteria': '> 0', 'score': float(score_p34), 'max_score': 1.0, 'details': {'div_yield': div_yield}}

    # ==========================
    # AGGREGATE TOTAL SCORES
    # ==========================
    quality_actual = sum(p['score'] for k, p in params.items() if k <= 28 and isinstance(p['score'], (int, float)))
    quality_max = sum(p['max_score'] for k, p in params.items() if k <= 28)

    val_actual = sum(p['score'] for k, p in params.items() if k > 28 and isinstance(p['score'], (int, float)))
    val_max = sum(p['max_score'] for k, p in params.items() if k > 28)

    total_actual = quality_actual + val_actual
    total_max = quality_max + val_max

    summary = {
        'symbol': symbol,
        'company_name': comp['name'],
        'quality_score': round(quality_actual, 1),
        'max_quality': quality_max,
        'valuation_score': round(val_actual, 1),
        'max_valuation': val_max,
        'total_score': round(total_actual, 1),
        'max_total': total_max,
        'score_pct': round((total_actual / total_max) * 100, 1),
        'parameters': params
    }

    # Save to analysis_results
    conn.execute('''
    INSERT OR REPLACE INTO analysis_results (
        symbol, quality_score, max_quality, valuation_score, max_valuation,
        total_score, max_total, score_pct, parameters_json
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        symbol, summary['quality_score'], summary['max_quality'],
        summary['valuation_score'], summary['max_valuation'],
        summary['total_score'], summary['max_total'],
        summary['score_pct'], json.dumps(params)
    ))
    conn.commit()
    conn.close()

    return summary

if __name__ == "__main__":
    import sys
    s = sys.argv[1] if len(sys.argv) > 1 else "JSWSTEEL"
    res = run_vm_analysis(s)
    print(f"=== ANALYSIS SUMMARY FOR {res['company_name']} ({res['symbol']}) ===")
    print(f"  Quality Score   : {res['quality_score']} / {res['max_quality']} ({(res['quality_score']/res['max_quality'])*100:.1f}%)")
    print(f"  Valuation Score : {res['valuation_score']} / {res['max_valuation']} ({(res['valuation_score']/res['max_valuation'])*100:.1f}%)")
    print(f"  Total Score     : {res['total_score']} / {res['max_total']} ({res['score_pct']}%)")
