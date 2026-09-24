import json
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CACHE_FILE = "mf_history_cache.json"
OUTPUT_EXCEL = "flexi_cap_3yr_rolling_analysis_half_yearly.xlsx"
DATA_JSON_OUT = "scratch/half_yearly_data.json"

def load_data():
    with open(CACHE_FILE, "r", encoding="utf-8") as f:
        cache = json.load(f)

    parsed = {}
    for code, item in cache.items():
        hist = []
        for e in item.get("data", []):
            try:
                dt = datetime.strptime(e["date"], "%d-%m-%Y")
                nav = float(e["nav"])
                hist.append((dt, nav))
            except (ValueError, TypeError):
                continue
        hist.sort(key=lambda x: x[0])
        parsed[code] = {
            "name": item.get("schemeName", ""),
            "history": hist,
            "earliest_dt": hist[0][0] if hist else None,
            "latest_dt": hist[-1][0] if hist else None
        }
    return parsed

def get_nav_on_or_before(history, target_date, max_lookback_days=5):
    candidates = [entry for entry in history if entry[0] <= target_date]
    if not candidates:
        return None, None
    last_dt, last_nav = candidates[-1]
    if (target_date - last_dt).days <= max_lookback_days:
        return last_dt, last_nav
    return None, None

def generate_windows():
    windows = []
    for yr in range(2015, 2024):
        for mo in [1, 7]:
            w_start = datetime(yr, mo, 1)
            w_end = datetime(yr + 3, mo, 1)
            if w_end > datetime(2026, 9, 1):
                break
            quarter = f"Q{(mo - 1) // 3 + 1}"
            windows.append((w_start, w_end, quarter))
    return windows

def evaluate_windows(parsed_data, windows):
    evaluations = []
    for w_start, w_end, quarter in windows:
        eligible_results = []
        for code, fund in parsed_data.items():
            if not fund["earliest_dt"] or not fund["latest_dt"]:
                continue
            if fund["earliest_dt"] <= w_start and fund["latest_dt"] >= w_end:
                s_dt, s_nav = get_nav_on_or_before(fund["history"], w_start)
                e_dt, e_nav = get_nav_on_or_before(fund["history"], w_end)
                if s_nav and e_nav and s_nav > 0:
                    cagr = ((e_nav / s_nav) ** (1.0 / 3.0) - 1.0) * 100.0
                    eligible_results.append({
                        "schemeCode": code,
                        "schemeName": fund["name"],
                        "cagr": round(cagr, 2)
                    })

        eligible_results.sort(key=lambda x: x["cagr"], reverse=True)
        n = len(eligible_results)
        top_25_count = max(1, round(n * 0.25)) if n > 0 else 0
        top_25_funds = eligible_results[:top_25_count]

        evaluations.append({
            "start_date": w_start.strftime("%Y-%m-%d"),
            "end_date": w_end.strftime("%Y-%m-%d"),
            "quarter": quarter,
            "n": n,
            "top_25_pct": top_25_count,
            "all_funds": eligible_results,
            "top_funds": top_25_funds
        })
    return evaluations

def compute_consistency(evaluations):
    fund_stats = {}
    for win in evaluations:
        top_names = {f["schemeName"]: f["cagr"] for f in win["top_funds"]}
        for f in win["all_funds"]:
            name = f["schemeName"]
            if name not in fund_stats:
                fund_stats[name] = {
                    "eligible_count": 0,
                    "top_count": 0,
                    "top_cagrs": [],
                    "top_windows": []
                }
            fund_stats[name]["eligible_count"] += 1
            if name in top_names:
                fund_stats[name]["top_count"] += 1
                fund_stats[name]["top_cagrs"].append(top_names[name])
                fund_stats[name]["top_windows"].append(win["start_date"][:7])

    ranked = []
    for name, s in fund_stats.items():
        if s["top_count"] > 0:
            consistency = (s["top_count"] / s["eligible_count"] * 100.0) if s["eligible_count"] > 0 else 0
            avg_cagr = sum(s["top_cagrs"]) / len(s["top_cagrs"]) if s["top_cagrs"] else 0
            win_range = f"{s['top_windows'][0]} to {s['top_windows'][-1]}" if s["top_windows"] else "-"
            ranked.append({
                "name": name,
                "top_count": s["top_count"],
                "eligible": s["eligible_count"],
                "consistency": round(consistency, 1),
                "avg_cagr": round(avg_cagr, 2),
                "range": win_range
            })

    ranked.sort(key=lambda x: (x["top_count"], x["consistency"]), reverse=True)
    return ranked

def create_excel(evaluations, consistency_ranked):
    wb = openpyxl.Workbook()

    font_title = Font(name="Calibri", size=14, bold=True, color="1F4E79")
    font_subtitle = Font(name="Calibri", size=10, italic=True, color="595959")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_win_summary = Font(name="Calibri", size=10, bold=True, color="1F4E79")
    font_detail = Font(name="Calibri", size=10, color="333333")
    font_rank = Font(name="Calibri", size=9, italic=True, color="7F7F7F")
    font_detail_cagr = Font(name="Calibri", size=10, bold=True, color="2E7D32")

    fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    fill_win_summary = PatternFill(start_color="EDF2F8", end_color="EDF2F8", fill_type="solid")
    fill_detail_zebra = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")

    border_thin = Side(border_style="thin", color="E0E0E0")
    border_summary_bottom = Side(border_style="medium", color="1F4E79")
    border_double = Side(border_style="double", color="1F4E79")

    cell_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
    summary_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_summary_bottom)

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    # Sheet 1: 3Y Rolling Windows
    ws1 = wb.active
    ws1.title = "3Y_Rolling_Windows"
    ws1.sheet_properties.outlinePr.summaryBelow = False

    ws1["A1"] = "Flexi Cap Mutual Funds — 3-Year Rolling Return Analysis (Regular Plans)"
    ws1["A1"].font = font_title
    ws1["A2"] = f"Half-Yearly Rolling Windows (2015-01-01 to 2026-09-01) | Click [+] on any row to expand Top 25% funds"
    ws1["A2"].font = font_subtitle

    headers = [
        "Start Date",
        "End Date",
        "Quarter",
        "Total Funds (n)",
        "Top 25% Count",
        "Fund Name",
        "3Y CAGR (%)"
    ]

    for col_num, h_text in enumerate(headers, 1):
        cell = ws1.cell(row=4, column=col_num, value=h_text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center if col_num in [1, 2, 3, 4, 5, 7] else align_left
        cell.border = Border(top=border_thin, bottom=border_double, left=border_thin, right=border_thin)

    current_row = 5
    for win in evaluations:
        summary_row = current_row
        top_funds = win["top_funds"]
        best_fund = top_funds[0] if top_funds else None
        best_name = f"#1: {best_fund['schemeName']}" if best_fund else "N/A"
        best_cagr = best_fund["cagr"] if best_fund else 0.0

        ws1.cell(row=summary_row, column=1, value=win["start_date"]).alignment = align_center
        ws1.cell(row=summary_row, column=2, value=win["end_date"]).alignment = align_center
        ws1.cell(row=summary_row, column=3, value=win["quarter"]).alignment = align_center
        ws1.cell(row=summary_row, column=4, value=win["n"]).alignment = align_center
        ws1.cell(row=summary_row, column=5, value=win["top_25_pct"]).alignment = align_center
        ws1.cell(row=summary_row, column=6, value=best_name).alignment = align_left

        c_top = ws1.cell(row=summary_row, column=7, value=best_cagr / 100.0)
        c_top.alignment = align_right
        c_top.number_format = "0.00%"

        for col in range(1, 8):
            cell = ws1.cell(row=summary_row, column=col)
            cell.font = font_win_summary
            cell.fill = fill_win_summary
            cell.border = summary_border

        current_row += 1

        detail_start_row = current_row
        for rank, fund in enumerate(top_funds[1:], 2):
            detail_row = current_row
            r_cell = ws1.cell(row=detail_row, column=1, value=f"#{rank}")
            r_cell.alignment = align_center
            r_cell.font = font_rank
            r_cell.border = cell_border

            ws1.cell(row=detail_row, column=2, value="").border = cell_border
            ws1.cell(row=detail_row, column=3, value="").border = cell_border
            ws1.cell(row=detail_row, column=4, value="").border = cell_border
            ws1.cell(row=detail_row, column=5, value="").border = cell_border

            f_cell = ws1.cell(row=detail_row, column=6, value=fund["schemeName"])
            f_cell.alignment = align_left
            f_cell.font = font_detail
            f_cell.border = cell_border

            c_cagr = ws1.cell(row=detail_row, column=7, value=fund["cagr"] / 100.0)
            c_cagr.alignment = align_right
            c_cagr.number_format = "0.00%"
            c_cagr.font = font_detail_cagr
            c_cagr.border = cell_border

            if rank % 2 == 0:
                for col in range(1, 8):
                    ws1.cell(row=detail_row, column=col).fill = fill_detail_zebra

            current_row += 1

        detail_end_row = current_row - 1
        if detail_end_row >= detail_start_row:
            ws1.row_dimensions.group(detail_start_row, detail_end_row, outline_level=1, hidden=True)

    col_widths = {1: 14, 2: 14, 3: 10, 4: 16, 5: 16, 6: 62, 7: 16}
    for col_idx, width in col_widths.items():
        ws1.column_dimensions[get_column_letter(col_idx)].width = width
    ws1.freeze_panes = "A5"

    # Bottom Section of ws1: Consistency Summary
    current_row += 3
    ws1.cell(row=current_row, column=1, value="Flexi Cap Funds — Top 25% Consistency Summary (Grouped by Name)").font = font_title
    current_row += 1
    ws1.cell(row=current_row, column=1, value=f"Total occurrences in Top Quartile across all {len(evaluations)} half-yearly 3-year rolling windows (2015-01-01 to 2026-09-01)").font = font_subtitle
    current_row += 2

    headers_summary = [
        "Rank",
        "Top 25% Count",
        "Eligible Windows",
        "Consistency (%)",
        "Active Period",
        "Fund Name",
        "Avg 3Y CAGR (%)"
    ]

    header_row = current_row
    for col_num, h_text in enumerate(headers_summary, 1):
        cell = ws1.cell(row=header_row, column=col_num, value=h_text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center if col_num in [1, 2, 3, 4, 5, 7] else align_left
        cell.border = Border(top=border_thin, bottom=border_double, left=border_thin, right=border_thin)

    current_row += 1
    for rank, item in enumerate(consistency_ranked, 1):
        row_num = current_row
        ws1.cell(row=row_num, column=1, value=f"#{rank}").alignment = align_center
        ws1.cell(row=row_num, column=2, value=item["top_count"]).alignment = align_center
        ws1.cell(row=row_num, column=3, value=item["eligible"]).alignment = align_center

        c_cons = ws1.cell(row=row_num, column=4, value=item["consistency"] / 100.0)
        c_cons.alignment = align_right
        c_cons.number_format = "0.0%"

        ws1.cell(row=row_num, column=5, value=item["range"]).alignment = align_center
        ws1.cell(row=row_num, column=6, value=item["name"]).alignment = align_left

        c_avg = ws1.cell(row=row_num, column=7, value=item["avg_cagr"] / 100.0)
        c_avg.alignment = align_right
        c_avg.number_format = "0.00%"
        c_avg.font = font_detail_cagr

        for col in range(1, 8):
            cell = ws1.cell(row=row_num, column=col)
            if col != 7:
                cell.font = font_detail
            cell.border = cell_border
            if rank % 2 == 0:
                cell.fill = fill_detail_zebra

        ws1.row_dimensions[row_num].outline_level = 0
        ws1.row_dimensions[row_num].hidden = False
        current_row += 1

    wb.save(OUTPUT_EXCEL)
    print(f"Report saved to {OUTPUT_EXCEL}")

if __name__ == "__main__":
    parsed = load_data()
    windows = generate_windows()
    evals = evaluate_windows(parsed, windows)
    consistency = compute_consistency(evals)
    create_excel(evals, consistency)

    # Save JSON for UI
    ui_payload = {
        "scheme_type": "Open Ended",
        "category_group": "Equity",
        "sub_category": "Flexi Cap",
        "total_windows": len(evals),
        "windows": evals,
        "consistency": consistency
    }
    os.makedirs("scratch", exist_ok=True)
    with open("scratch/half_yearly_data.json", "w", encoding="utf-8") as f:
        json.dump(ui_payload, f)
    print("UI payload saved to scratch/half_yearly_data.json")
