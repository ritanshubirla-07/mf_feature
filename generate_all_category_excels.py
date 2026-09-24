import json
import base64
import os
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

with open("scratch/full_master_engine_data.json", "r", encoding="utf-8") as f:
    engine_data = json.load(f)

open_ended_db = engine_data["open_ended"]

# Styles
font_title = Font(name="Calibri", size=14, bold=True, color="1F4E79")
font_subtitle = Font(name="Calibri", size=10, italic=True, color="595959")
font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
font_sub_header = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
font_win_summary = Font(name="Calibri", size=10, bold=True, color="1F4E79")
font_detail = Font(name="Calibri", size=10, color="333333")
font_rank = Font(name="Calibri", size=9, italic=True, color="7F7F7F")
font_detail_cagr = Font(name="Calibri", size=10, bold=True, color="2E7D32")
font_meta_label = Font(name="Calibri", size=9, bold=True, color="475569")
font_meta_val = Font(name="Calibri", size=9, bold=True, color="0F172A")
font_black_box = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
font_empty_box = Font(name="Calibri", size=9, color="94A3B8")
font_q_label = Font(name="Calibri", size=9, bold=True, color="1E293B")
font_banner = Font(name="Calibri", size=9, bold=True, color="FFFFFF")

font_card_section = Font(name="Calibri", size=9.5, bold=True, color="1F4E79")
font_card_label = Font(name="Calibri", size=8.5, bold=True, color="64748B")
font_card_val = Font(name="Calibri", size=11, bold=True, color="0F172A")
font_card_val_fm = Font(name="Calibri", size=9, bold=False, color="1E293B")
font_q_status = Font(name="Calibri", size=9, bold=True, color="1F4E79")

fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
fill_dark_slate = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
fill_section_dark = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
fill_win_summary = PatternFill(start_color="EDF2F8", end_color="EDF2F8", fill_type="solid")
fill_detail_zebra = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")
fill_details_strip = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
fill_card_header = PatternFill(start_color="E8EEF5", end_color="E8EEF5", fill_type="solid")
fill_card_label = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
fill_black = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
fill_q_badge = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
fill_q_summary = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

border_thin = Side(border_style="thin", color="E0E0E0")
border_summary_bottom = Side(border_style="medium", color="1F4E79")
border_double = Side(border_style="double", color="1F4E79")

cell_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
header_border = Border(top=border_thin, bottom=border_double, left=border_thin, right=border_thin)
summary_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_summary_bottom)

align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")
align_right = Alignment(horizontal="right", vertical="center")
align_left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

def build_category_excel(cat_name, cat_funds, output_file, window_dates, last_6_labels):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rolling_Analysis"
    ws.sheet_properties.outlinePr.summaryBelow = False
    ws.sheet_format.outlineLevelRow = 1
    ws.views.sheetView[0].showGridLines = True

    # 1. Title Block
    ws["A1"] = "MUTUAL FUND ROLLING RETURN ANALYSIS — STATUTORY AMFI PERFORMANCE ENGINE"
    ws["A1"].font = font_title
    ws["A2"] = f"Scheme Type: Open Ended | Category Group: Equity | Sub Category: {cat_name} | Option: Growth"
    ws["A2"].font = font_subtitle
    ws["A3"] = "Analysis Period: 10 Years (Jan 2015 to Jul 2026) | Source: Statutory AMFI NAV Feed"
    ws["A3"].font = font_subtitle

    # 2. Table 1: 3-Year Rolling Return Analysis
    ws["A5"] = f"{cat_name} Mutual Funds (Growth) — 3-Year Rolling Return Analysis (Regular Plans)"
    ws["A5"].font = font_title
    ws["A6"] = "Half-Yearly Rolling Windows (Jan 2015 to Jul 2026) | Click [+] on left margin to expand Top 25% funds"
    ws["A6"].font = font_subtitle

    t1_headers = [
        (1, "Start Date", align_center, None),
        (2, "End Date", align_center, None),
        (3, "Quarter", align_center, None),
        (4, "Total Funds (n)", align_center, None),
        (5, "Top 25% Count", align_center, None),
        (6, "Top Quartile Fund Name", align_left, 8),
        (9, "3Y CAGR (%)", align_right, None)
    ]
    for col_idx, h, al, merge_to in t1_headers:
        c = ws.cell(row=7, column=col_idx, value=h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = al
        c.border = header_border
        if merge_to:
            for mc in range(col_idx, merge_to + 1):
                m_cell = ws.cell(row=7, column=mc)
                m_cell.fill = fill_header
                m_cell.border = header_border
            ws.merge_cells(start_row=7, start_column=col_idx, end_row=7, end_column=merge_to)

    num_windows = len(window_dates)
    windows_data = []
    for w_idx, win in enumerate(window_dates):
        valid = []
        for f in cat_funds:
            cagrs = f.get("cagrs", [])
            if w_idx < len(cagrs) and cagrs[w_idx] is not None:
                valid.append({"schemeName": f.get("name", ""), "cagr": cagrs[w_idx]})
        valid.sort(key=lambda x: x["cagr"], reverse=True)
        n = len(valid)
        top_k = max(1, round(n * 0.25)) if n > 0 else 0
        top_funds = valid[:top_k]
        windows_data.append({
            "start_date": win[0],
            "end_date": win[1],
            "quarter": win[2],
            "n": n,
            "top_25_pct": top_k,
            "top_funds": top_funds
        })

    curr_row = 8
    for win in windows_data:
        top_funds = win.get("top_funds", [])
        best_fund = top_funds[0] if top_funds else None
        best_name = f"#1: {best_fund['schemeName']}" if best_fund else "N/A"
        best_cagr = best_fund["cagr"] if best_fund else 0.0

        s_row = curr_row
        ws.cell(row=s_row, column=1, value=win["start_date"]).alignment = align_center
        ws.cell(row=s_row, column=2, value=win["end_date"]).alignment = align_center
        ws.cell(row=s_row, column=3, value=win["quarter"]).alignment = align_center
        ws.cell(row=s_row, column=4, value=win["n"]).alignment = align_center
        ws.cell(row=s_row, column=5, value=win["top_25_pct"]).alignment = align_center
        ws.cell(row=s_row, column=6, value=best_name).alignment = align_left
        
        c_top = ws.cell(row=s_row, column=9, value=best_cagr / 100.0)
        c_top.alignment = align_right
        c_top.number_format = "0.00%"

        for col in range(1, 10):
            cell = ws.cell(row=s_row, column=col)
            cell.font = font_win_summary
            cell.fill = fill_win_summary
            cell.border = summary_border
            if col <= 5:
                cell.alignment = align_center
            elif col == 6:
                cell.alignment = align_left
            elif col == 9:
                cell.alignment = align_right

        ws.merge_cells(start_row=s_row, start_column=6, end_row=s_row, end_column=8)
        curr_row += 1

        detail_start_row = curr_row
        for rank, fund in enumerate(top_funds[1:], 2):
            d_row = curr_row
            r_c = ws.cell(row=d_row, column=1, value=f"#{rank}")
            r_c.alignment = align_center
            r_c.font = font_rank

            f_c = ws.cell(row=d_row, column=6, value=fund["schemeName"])
            f_c.alignment = align_left
            f_c.font = font_detail

            c_cagr = ws.cell(row=d_row, column=9, value=fund["cagr"] / 100.0)
            c_cagr.alignment = align_right
            c_cagr.number_format = "0.00%"
            c_cagr.font = font_detail_cagr

            for col in range(1, 10):
                cell = ws.cell(row=d_row, column=col)
                cell.border = cell_border
                if rank % 2 == 0:
                    cell.fill = fill_detail_zebra

            ws.merge_cells(start_row=d_row, start_column=6, end_row=d_row, end_column=8)
            curr_row += 1

        detail_end_row = curr_row - 1
        if detail_end_row >= detail_start_row:
            ws.row_dimensions.group(detail_start_row, detail_end_row, outline_level=1, hidden=True)
            ws.row_dimensions[s_row].collapsed = True

    ws.freeze_panes = None

    # Build Consistency List
    min_required = round(num_windows * 0.75) # 14
    consistency_funds = []

    for f in cat_funds:
        cagrs = f.get("cagrs", [])
        completed_cagrs = [c for c in cagrs if c is not None]
        eligible_count = len(completed_cagrs)
        if eligible_count < min_required:
            continue

        top_count = 0
        for w_idx in range(num_windows):
            if w_idx >= len(cagrs) or cagrs[w_idx] is None:
                continue
            w_top_names = [x["schemeName"] for x in windows_data[w_idx]["top_funds"]]
            if f.get("name", "") in w_top_names:
                top_count += 1

        consistency_score = round((top_count / eligible_count) * 100.0) if eligible_count > 0 else 0

        last_6_details = []
        recent_top_count = 0
        start_w = max(0, num_windows - 6)
        for w_idx in range(start_w, num_windows):
            c_val = cagrs[w_idx] if w_idx < len(cagrs) else None
            q_num = 4
            if c_val is not None:
                all_cat_cagrs = [x.get("cagrs", [])[w_idx] for x in cat_funds if w_idx < len(x.get("cagrs", [])) and x.get("cagrs", [])[w_idx] is not None]
                all_cat_cagrs.sort(reverse=True)
                n_c = len(all_cat_cagrs)
                q1_cut = round(n_c * 0.25)
                q2_cut = round(n_c * 0.50)
                q3_cut = round(n_c * 0.75)
                try:
                    rank_in_w = all_cat_cagrs.index(c_val) + 1
                    if rank_in_w <= q1_cut:
                        q_num = 1
                    elif rank_in_w <= q2_cut:
                        q_num = 2
                    elif rank_in_w <= q3_cut:
                        q_num = 3
                    else:
                        q_num = 4
                except:
                    q_num = 4
            if q_num == 1:
                recent_top_count += 1
            last_6_details.append({"cagr": c_val, "quartile": q_num})

        recency_score = round((recent_top_count / 6.0) * 100.0)
        composite_score = round(0.65 * consistency_score + 0.35 * recency_score)
        if composite_score <= 0:
            continue

        avg_cagr = sum(completed_cagrs) / eligible_count if eligible_count > 0 else 0.0

        consistency_funds.append({
            "name": f.get("name", ""),
            "composite_score": composite_score,
            "consistency_score": consistency_score,
            "recency_score": recency_score,
            "eligible": eligible_count,
            "avg_cagr": avg_cagr,
            "nav": f.get("nav", None),
            "aum": f.get("aum", None),
            "expense_ratio": f.get("expense_ratio", None),
            "fund_manager": f.get("fund_manager", None),
            "last_6_details": last_6_details
        })

    consistency_funds.sort(key=lambda x: (x["composite_score"], x["consistency_score"], x["avg_cagr"]), reverse=True)
    category_total = len(cat_funds)
    top_quota = max(1, round(category_total * 0.25))
    final_consistency = consistency_funds[:top_quota]

    # 3. Table 2: Top 25% Consistency Summary
    curr_row += 3
    ws.cell(row=curr_row, column=1, value=f"{cat_name} Funds (Growth) — Top 25% Consistency Summary (Grouped by Name)").font = font_title
    curr_row += 1
    ws.cell(row=curr_row, column=1, value=f"Ranked by Composite Score (65% Consistency + 35% Recency) | Top 25% Funds ({len(final_consistency)}) with ≥ 75% Windows (≥{min_required} of {num_windows}) | Click [+] on left margin to view Fund Details & 2D Matrix").font = font_subtitle
    curr_row += 2

    t2_headers = [
        (1, "Rank", align_center, None),
        (2, "Composite Score", align_center, None),
        (3, "Consistency Score", align_center, None),
        (4, "Recency Score", align_center, None),
        (5, "Completed Windows", align_center, None),
        (6, "Fund Name", align_left, 8),
        (9, "Avg 3Y CAGR (%)", align_right, None)
    ]
    for col_idx, h, al, merge_to in t2_headers:
        c = ws.cell(row=curr_row, column=col_idx, value=h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = al
        c.border = header_border
        if merge_to:
            for mc in range(col_idx, merge_to + 1):
                m_cell = ws.cell(row=curr_row, column=mc)
                m_cell.fill = fill_header
                m_cell.border = header_border
            ws.merge_cells(start_row=curr_row, start_column=col_idx, end_row=curr_row, end_column=merge_to)
    curr_row += 1

    for rank, item in enumerate(final_consistency, 1):
        fund_row = curr_row
        ws.cell(row=fund_row, column=1, value=f"#{rank}").alignment = align_center
        ws.cell(row=fund_row, column=2, value=item["composite_score"]).alignment = align_center
        ws.cell(row=fund_row, column=3, value=item["consistency_score"]).alignment = align_center
        ws.cell(row=fund_row, column=4, value=item["recency_score"]).alignment = align_center
        ws.cell(row=fund_row, column=5, value=f"{item.get('eligible', 18)} / 18").alignment = align_center
        ws.cell(row=fund_row, column=6, value=item["name"]).alignment = align_left

        c_avg = ws.cell(row=fund_row, column=9, value=item["avg_cagr"] / 100.0)
        c_avg.alignment = align_right
        c_avg.number_format = "0.00%"
        c_avg.font = font_detail_cagr

        for col in range(1, 10):
            cell = ws.cell(row=fund_row, column=col)
            cell.border = cell_border
            cell.fill = fill_win_summary
            if col == 1:
                cell.font = font_win_summary
            elif col == 2:
                cell.font = Font(name="Calibri", size=11, bold=True, color="0F172A")
            elif col == 3:
                cell.font = Font(name="Calibri", size=11, bold=True, color="005F73")
            elif col == 4:
                cell.font = Font(name="Calibri", size=11, bold=True, color="B45309")
            elif col == 5:
                cell.font = Font(name="Calibri", size=10, bold=True, color="475569")
            elif col == 6:
                cell.font = Font(name="Calibri", size=10, bold=True, color="1E293B")

        ws.merge_cells(start_row=fund_row, start_column=6, end_row=fund_row, end_column=8)
        ws.row_dimensions[fund_row].collapsed = True
        curr_row += 1

        sub_start_row = curr_row

        # 1. SECTION HEADER BANNER: FUND DETAILS & OPERATIONS OVERVIEW (Cols B to I)
        banner_meta_row = curr_row
        ws.row_dimensions[banner_meta_row].height = 20
        ws.cell(row=banner_meta_row, column=2, value="FUND OPERATIONS & PORTFOLIO METRICS").font = font_card_section
        ws.cell(row=banner_meta_row, column=2).alignment = align_left
        for c in range(1, 10):
            cell = ws.cell(row=banner_meta_row, column=c)
            cell.fill = fill_card_header
            cell.border = cell_border
        ws.merge_cells(start_row=banner_meta_row, start_column=2, end_row=banner_meta_row, end_column=9)
        curr_row += 1

        # 2. METRIC TITLES ROW (Cols B-C: NAV, D-E: AUM, F: ER, G-I: FM)
        lbl_row = curr_row
        ws.row_dimensions[lbl_row].height = 18
        for c in range(1, 10):
            cell = ws.cell(row=lbl_row, column=c)
            cell.fill = fill_card_label
            cell.border = cell_border

        ws.cell(row=lbl_row, column=2, value="LATEST NAV").font = font_card_label
        ws.cell(row=lbl_row, column=2).alignment = align_center
        ws.merge_cells(start_row=lbl_row, start_column=2, end_row=lbl_row, end_column=3)

        ws.cell(row=lbl_row, column=4, value="LATEST AUM (₹ CR)").font = font_card_label
        ws.cell(row=lbl_row, column=4).alignment = align_center
        ws.merge_cells(start_row=lbl_row, start_column=4, end_row=lbl_row, end_column=5)

        ws.cell(row=lbl_row, column=6, value="EXPENSE RATIO").font = font_card_label
        ws.cell(row=lbl_row, column=6).alignment = align_center

        ws.cell(row=lbl_row, column=7, value="FUND MANAGER(S)").font = font_card_label
        ws.cell(row=lbl_row, column=7).alignment = align_left
        ws.merge_cells(start_row=lbl_row, start_column=7, end_row=lbl_row, end_column=9)
        curr_row += 1

        # 3. METRIC VALUES ROW (Cols B-C: NAV val, D-E: AUM val, F: ER val, G-I: FM val)
        val_row = curr_row
        ws.row_dimensions[val_row].height = 26
        for c in range(1, 10):
            cell = ws.cell(row=val_row, column=c)
            cell.fill = fill_white
            cell.border = cell_border

        nav_val = f"₹{item['nav']:.2f}" if item['nav'] is not None else "—"
        aum_val = f"₹{float(item['aum']):,.2f} Cr" if item['aum'] is not None else "—"
        er_val = f"{item['expense_ratio']}%" if item['expense_ratio'] is not None else "—"
        fm_val = item['fund_manager'] if item['fund_manager'] and item['fund_manager'] != 'None' else "—"

        ws.cell(row=val_row, column=2, value=nav_val).font = font_card_val
        ws.cell(row=val_row, column=2).alignment = align_center
        ws.merge_cells(start_row=val_row, start_column=2, end_row=val_row, end_column=3)

        ws.cell(row=val_row, column=4, value=aum_val).font = font_card_val
        ws.cell(row=val_row, column=4).alignment = align_center
        ws.merge_cells(start_row=val_row, start_column=4, end_row=val_row, end_column=5)

        ws.cell(row=val_row, column=6, value=er_val).font = font_card_val
        ws.cell(row=val_row, column=6).alignment = align_center

        ws.cell(row=val_row, column=7, value=fm_val).font = font_card_val_fm
        ws.cell(row=val_row, column=7).alignment = align_left_wrap
        ws.merge_cells(start_row=val_row, start_column=7, end_row=val_row, end_column=9)
        curr_row += 1

        # 4. 2D QUARTILE MATRIX SECTION BANNER (Cols B to I)
        banner_row = curr_row
        ws.row_dimensions[banner_row].height = 22
        ws.cell(row=banner_row, column=2, value="3-YEAR ROLLING RETURN QUARTILE TRACKING MATRIX (LAST 6 WINDOWS)").font = font_banner
        ws.cell(row=banner_row, column=2).alignment = align_center
        for c in range(1, 10):
            cell = ws.cell(row=banner_row, column=c)
            cell.fill = fill_header
            cell.border = cell_border
        ws.merge_cells(start_row=banner_row, start_column=2, end_row=banner_row, end_column=9)
        curr_row += 1

        # 5. 2D MATRIX HEADER (Cols B to I)
        grid_head_row = curr_row
        ws.row_dimensions[grid_head_row].height = 20
        for c in range(1, 10):
            cell = ws.cell(row=grid_head_row, column=c)
            cell.fill = fill_dark_slate
            cell.border = cell_border

        q_h = ws.cell(row=grid_head_row, column=2, value="Quartile Standing")
        q_h.font = font_sub_header
        q_h.alignment = align_center

        end_month_years = [w[1] for w in window_dates[-6:]]
        for w_idx, em in enumerate(end_month_years):
            col = 3 + w_idx
            c = ws.cell(row=grid_head_row, column=col, value=em)
            c.font = font_sub_header
            c.alignment = align_center

        c_summary_hdr = ws.cell(row=grid_head_row, column=9, value="Quartile Status")
        c_summary_hdr.font = font_sub_header
        c_summary_hdr.alignment = align_center
        curr_row += 1

        # 6. 4 QUARTILE ROWS (Q1 to Q4) - PERFECTLY UNIFORM COLS C TO H
        q_defs = [
            (1, "Q1 (Top 25%)"),
            (2, "Q2 (25% - 50%)"),
            (3, "Q3 (50% - 75%)"),
            (4, "Q4 (Bottom 25%)")
        ]
        for q_id, q_label in q_defs:
            q_row = curr_row
            ws.row_dimensions[q_row].height = 22

            ws.cell(row=q_row, column=1).border = cell_border

            l_cell = ws.cell(row=q_row, column=2, value=q_label)
            l_cell.font = font_q_label
            l_cell.fill = fill_q_badge
            l_cell.alignment = align_center
            l_cell.border = cell_border

            q_count = 0
            for w_idx, w_info in enumerate(item["last_6_details"]):
                col = 3 + w_idx
                cell = ws.cell(row=q_row, column=col)
                cell.border = cell_border

                if w_info["quartile"] == q_id and w_info["cagr"] is not None:
                    q_count += 1
                    # SOLID BLACK WIDENED BOX!
                    cell.value = w_info["cagr"] / 100.0
                    cell.number_format = "0.00%"
                    cell.fill = fill_black
                    cell.font = font_black_box
                    cell.alignment = align_center
                else:
                    cell.value = "-"
                    cell.fill = fill_detail_zebra
                    cell.font = font_empty_box
                    cell.alignment = align_center

            # Col 9 Summary Status
            status_cell = ws.cell(row=q_row, column=9)
            status_cell.border = cell_border
            status_cell.alignment = align_center
            if q_count > 0:
                status_cell.value = f"{q_count} of 6 in {q_label[:2]}"
                status_cell.font = font_q_status
                status_cell.fill = fill_q_summary
            else:
                status_cell.value = "-"
                status_cell.font = font_empty_box
                status_cell.fill = fill_white

            curr_row += 1

        # 7. Clean Spacer Row
        spacer_row = curr_row
        ws.row_dimensions[spacer_row].height = 6
        for c in range(1, 10):
            ws.cell(row=spacer_row, column=c).border = cell_border
        curr_row += 1

        sub_end_row = curr_row - 1
        ws.row_dimensions.group(sub_start_row, sub_end_row, outline_level=1, hidden=True)

    col_widths = {
        1: 12, 2: 16, 3: 14, 4: 14, 5: 14, 6: 14, 7: 14, 8: 14, 9: 18
    }
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    wb.save(output_file)

    # Post-save: patch sheetFormatPr to include outlineLevelRow="1"
    # OpenPyXL doesn't reliably write this attribute; we patch the XML directly.
    import zipfile, shutil, re as _re
    patched = output_file + ".patching"
    with zipfile.ZipFile(output_file, 'r') as zin, zipfile.ZipFile(patched, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'xl/worksheets/sheet1.xml':
                xml = data.decode('utf-8')
                # Ensure outlineLevelRow="1" in sheetFormatPr
                if 'outlineLevelRow' not in xml:
                    xml = xml.replace(
                        '<sheetFormatPr',
                        '<sheetFormatPr outlineLevelRow="1"',
                        1
                    )
                data = xml.encode('utf-8')
            zout.writestr(item, data)
    shutil.move(patched, output_file)


window_dates = engine_data["window_dates"]
last_6_labels = engine_data["last_6_labels"]

b64_map = {}
for cat_name, cat_funds in open_ended_db.items():
    if not isinstance(cat_funds, list):
        continue

    safe_name = re.sub(r'[^a-zA-Z0-9_]', '_', cat_name.lower())
    out_file = f"scratch/{safe_name}_report.xlsx"
    build_category_excel(cat_name, cat_funds, out_file, window_dates, last_6_labels)
    with open(out_file, "rb") as f:
        b64_val = base64.b64encode(f.read()).decode("utf-8")
        b64_map[cat_name] = b64_val

# Also build root flexi cap excel
build_category_excel("Flexi Cap Fund", open_ended_db["Flexi Cap Fund"], "flexi_cap_3yr_rolling_analysis_half_yearly.xlsx", window_dates, last_6_labels)

with open("scratch/full_b64_map.json", "w", encoding="utf-8") as f:
    json.dump(b64_map, f)

print(f"Generated neat, symmetrical Excel reports for all {len(b64_map)} categories with native outline dropdowns!")
