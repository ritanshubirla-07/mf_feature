"""
mf_analytics_engine.py
======================
Reverse-engineered, standalone analytics engine reproducing the exact
mathematical formulas used by TigZig (MFPRO) and SEBI/AMFI performance portals.

Calculates:
  1. Point-to-Point CAGR (Fund and Benchmark)
  2. Rolling Returns Distribution (Calendar-day ASOF Lookback: Avg, Median, Min, Max, % Negative)
  3. Annualized Volatility
  4. Tracking Error (TE)
  5. Information Ratio (IR)
  6. Beta, R-Squared, and t-statistic
  7. Annualized Alpha (Compounded daily intercept)
  8. Sharpe Ratio & Sortino Ratio (Downside deviation)
  9. Treynor Ratio
 10. Upside and Downside Capture Ratios
 11. Historical VaR (95%) and CVaR (Expected Shortfall)
 12. Ulcer Index (Drawdown Depth & Duration)
 13. Win Rate

Requires only standard libraries: numpy, pandas.
No external APIs or authentication required once raw NAV/index data is provided.
"""

from typing import Dict, Any, List, Optional
import numpy as np
import pandas as pd


def compute_cagr(start_val: float, end_val: float, days_between: float) -> float:
    """Computes Compound Annual Growth Rate (CAGR).
    
    Formula:
      CAGR = (end_val / start_val) ** (365.25 / days_between) - 1
    """
    if start_val <= 0 or end_val <= 0 or days_between <= 0:
        return 0.0
    return (end_val / start_val) ** (365.25 / days_between) - 1


def compute_rolling_returns(
    df: pd.DataFrame,
    window_days: int = 1826,      # 5 years = 1,826 days; 3 years = 1,096; 1 year = 365
    min_gap_days: int = 1643      # TigZig minimum threshold (1,643 for 5Y; 986 for 3Y)
) -> Dict[str, Any]:
    """Calculates rolling returns using calendar-day look-back with ASOF matching.
    
    Parameters:
      df: DataFrame with ['date', 'nav'], sorted by date ascending.
      window_days: Calendar days look-back (e.g. 1826 for 5Y).
      min_gap_days: Minimum actual days required between start and end date.
      
    Returns:
      Dict with observation count, average, median, min, max, std dev, and % negative.
    """
    dates = pd.to_datetime(df['date']).values
    navs = df['nav'].astype(float).values
    
    cagrs = []
    end_dates = []
    
    for i in range(len(df)):
        end_d = pd.to_datetime(dates[i])
        target_start_d = end_d - pd.Timedelta(days=window_days)
        
        # ASOF Join: find the latest available date on or before target_start_d
        idx = np.searchsorted(dates, np.datetime64(target_start_d), side='right') - 1
        if idx >= 0:
            start_d = pd.to_datetime(dates[idx])
            days = (end_d - start_d).days
            if days >= min_gap_days:
                cagr = compute_cagr(navs[idx], navs[i], days)
                cagrs.append(cagr)
                end_dates.append(end_d)
                
    if not cagrs:
        return {"observations": 0}
        
    cagr_arr = np.array(cagrs)
    min_idx = int(np.argmin(cagr_arr))
    max_idx = int(np.argmax(cagr_arr))
    
    return {
        "observations": len(cagr_arr),
        "avg_cagr_pct": round(float(np.mean(cagr_arr)) * 100, 2),
        "median_cagr_pct": round(float(np.median(cagr_arr)) * 100, 2),
        "min_cagr_pct": round(float(cagr_arr[min_idx]) * 100, 2),
        "min_end_date": str(end_dates[min_idx].strftime('%Y-%m-%d')),
        "max_cagr_pct": round(float(cagr_arr[max_idx]) * 100, 2),
        "max_end_date": str(end_dates[max_idx].strftime('%Y-%m-%d')),
        "std_dev_pct": round(float(np.std(cagr_arr, ddof=1)) * 100, 2),
        "pct_negative": round(float(np.mean(cagr_arr < 0)) * 100, 2)
    }


def compute_all_risk_metrics(
    fund_df: pd.DataFrame,
    bench_df: pd.DataFrame,
    rf_annual_pct: float = 6.0,   # Default risk-free rate (6.0% p.a. in TigZig / India context)
    trading_days: int = 252       # Standard trading days annualization factor
) -> Dict[str, Any]:
    """Computes all 15+ quantitative risk, return, and benchmark metrics
    reproducing TigZig's /api/mf-data and SEBI disclosures.
    
    Parameters:
      fund_df:  DataFrame with ['date', 'nav']
      bench_df: DataFrame with ['date', 'nav'] (or index price/TRI)
      rf_annual_pct: Annualized risk-free rate in percentage (e.g. 6.0 for 6%)
      trading_days: Annualization scalar (252)
      
    Returns:
      Comprehensive dictionary of analytical metrics.
    """
    f = fund_df.copy()
    b = bench_df.copy()
    
    f['date'] = pd.to_datetime(f['date'])
    b['date'] = pd.to_datetime(b['date'])
    
    f = f.sort_values('date').drop_duplicates('date')
    b = b.sort_values('date').drop_duplicates('date')
    
    # Align dates (intersection of trading days)
    merged = pd.merge(f, b, on='date', suffixes=('_fund', '_bench')).sort_values('date').reset_index(drop=True)
    if len(merged) < 30:
        raise ValueError(f"Insufficient overlapping trading days: {len(merged)} (minimum 30 required)")
        
    # Daily percentage returns
    merged['r_fund'] = merged['nav_fund'].pct_change() * 100
    merged['r_bench'] = merged['nav_bench'].pct_change() * 100
    df_clean = merged.dropna().copy()
    
    n = len(df_clean)
    rf_daily = rf_annual_pct / trading_days
    
    rf_series = df_clean['r_fund'].values
    rb_series = df_clean['r_bench'].values
    
    # 1. Total & Annualized Volatility
    daily_vol = float(np.std(rf_series, ddof=1))
    ann_vol = daily_vol * np.sqrt(trading_days)
    
    # 2. Tracking Error (Annualized sample standard deviation of excess returns)
    excess_daily = rf_series - rb_series
    daily_te = float(np.std(excess_daily, ddof=1))
    ann_te = daily_te * np.sqrt(trading_days)
    
    # 3. Information Ratio
    # Numerator: Mean daily excess annualized (or CAGR excess)
    # Denominator: Annualized Tracking Error
    mean_daily_excess = float(np.mean(excess_daily))
    ir = (mean_daily_excess * trading_days) / ann_te if ann_te > 0 else 0.0
    
    # 4. Beta, R-Squared, and Alpha
    # Regression: r_fund = alpha_daily + beta * r_bench
    cov_matrix = np.cov(rf_series, rb_series, ddof=1)
    cov_fb = cov_matrix[0, 1]
    var_b = cov_matrix[1, 1]
    var_f = cov_matrix[0, 0]
    
    beta = float(cov_fb / var_b) if var_b > 0 else 0.0
    r_squared = float((cov_fb ** 2) / (var_f * var_b)) if (var_f * var_b) > 0 else 0.0
    
    # Daily Intercept (Alpha)
    mean_rf = float(np.mean(rf_series))
    mean_rb = float(np.mean(rb_series))
    alpha_daily = mean_rf - (beta * mean_rb)
    
    # TigZig Compounded Annualized Alpha: ((1 + alpha_daily/100)^252 - 1) * 100
    annualized_alpha = ((1 + (alpha_daily / 100)) ** trading_days - 1) * 100
    
    # t-Statistic for Beta
    # SXX = sum((rb - mean_rb)^2) = var_b * (n - 1)
    # SYY = sum((rf - mean_rf)^2) = var_f * (n - 1)
    # SXY = cov_fb * (n - 1)
    sxx = var_b * (n - 1)
    syy = var_f * (n - 1)
    sxy = cov_fb * (n - 1)
    mse = (syy - (beta * sxy)) / (n - 2) if n > 2 else 0.0
    se_beta = np.sqrt(mse / sxx) if (mse > 0 and sxx > 0) else 1e-9
    t_stat = float(beta / se_beta) if se_beta > 0 else 0.0
    
    # 5. Sharpe Ratio
    # (Mean daily excess return over Rf * 252) / (StdDev * sqrt(252))
    fund_excess_rf = rf_series - rf_daily
    mean_excess_rf = float(np.mean(fund_excess_rf))
    sharpe = (mean_excess_rf * trading_days) / ann_vol if ann_vol > 0 else 0.0
    
    # 6. Sortino Ratio
    # Downside deviation only penalizes returns below risk-free rate
    downside_deviations = np.minimum(0, fund_excess_rf)
    downside_var = float(np.mean(downside_deviations ** 2))
    downside_dev = np.sqrt(downside_var) * np.sqrt(trading_days)
    sortino = (mean_excess_rf * trading_days) / downside_dev if downside_dev > 0 else 0.0
    
    # 7. Treynor Ratio
    # (CAGR_fund - Rf) / Beta
    days_total = (merged['date'].iloc[-1] - merged['date'].iloc[0]).days
    cagr_fund = compute_cagr(merged['nav_fund'].iloc[0], merged['nav_fund'].iloc[-1], days_total) * 100
    cagr_bench = compute_cagr(merged['nav_bench'].iloc[0], merged['nav_bench'].iloc[-1], days_total) * 100
    treynor = (cagr_fund - rf_annual_pct) / beta if beta != 0 else 0.0
    
    # 8. Capture Ratios (Upside and Downside)
    up_days = rb_series > 0
    down_days = rb_series < 0
    
    mean_f_up = float(np.mean(rf_series[up_days])) if np.sum(up_days) > 0 else 0.0
    mean_b_up = float(np.mean(rb_series[up_days])) if np.sum(up_days) > 0 else 0.0
    upside_capture = (mean_f_up / mean_b_up) * 100 if mean_b_up != 0 else 0.0
    
    mean_f_down = float(np.mean(rf_series[down_days])) if np.sum(down_days) > 0 else 0.0
    mean_b_down = float(np.mean(rb_series[down_days])) if np.sum(down_days) > 0 else 0.0
    downside_capture = (mean_f_down / mean_b_down) * 100 if mean_b_down != 0 else 0.0
    
    # 9. Win Rate (% of days with positive return)
    win_rate = float(np.mean(rf_series > 0)) * 100
    
    # 10. Value at Risk (95%) & Conditional VaR (CVaR 95%)
    # Historical method (5th percentile of daily returns)
    var_95 = float(np.percentile(rf_series, 5))
    tail_returns = rf_series[rf_series <= var_95]
    cvar_95 = float(np.mean(tail_returns)) if len(tail_returns) > 0 else var_95
    
    # 11. Ulcer Index (Overall drawdown depth and duration pain)
    # Peak up to day i
    nav_arr = merged['nav_fund'].values
    running_max = np.maximum.accumulate(nav_arr)
    percentage_drawdowns = ((nav_arr - running_max) / running_max) * 100
    ulcer_index = float(np.sqrt(np.mean(percentage_drawdowns ** 2)))
    
    return {
        "observations": n,
        "fund_cagr_pct": round(cagr_fund, 2),
        "bench_cagr_pct": round(cagr_bench, 2),
        "annualized_vol_pct": round(ann_vol, 2),
        "tracking_error_pct": round(ann_te, 2),
        "information_ratio": round(ir, 2),
        "beta": round(beta, 2),
        "r_squared": round(r_squared, 4),
        "alpha_pct": round(annualized_alpha, 2),
        "t_stat": round(t_stat, 2),
        "sharpe_ratio": round(sharpe, 2),
        "sortino_ratio": round(sortino, 2),
        "treynor_ratio": round(treynor, 2),
        "upside_capture_pct": round(upside_capture, 2),
        "downside_capture_pct": round(downside_capture, 2),
        "win_rate_pct": round(win_rate, 2),
        "var_95_pct": round(var_95, 2),
        "cvar_95_pct": round(cvar_95, 2),
        "ulcer_index": round(ulcer_index, 2)
    }
