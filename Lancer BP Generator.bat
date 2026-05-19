from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import WorksheetProperties
from styles import *

def build_output_print(wb, config, years, row_map):
    """
    A4 print-ready consolidated output — calibrated for 1 page per statement.
    Mimics boutique restructuring IM output quality.
    """
    ws = wb.create_sheet("OUTPUT (PRINT)")
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85

    n = len(years)
    ccy = config["currency"]
    unit = f"{ccy}k"
    nm   = config["company_name"]

    # ── Page setup: A4, landscape, fit to width ──────────────────────────────
    ws.page_setup.paperSize     = ws.PAPERSIZE_A4
    ws.page_setup.orientation   = "landscape"
    ws.page_setup.fitToPage     = True
    ws.page_setup.fitToWidth    = 1
    ws.page_setup.fitToHeight   = 0
    ws.page_setup.scale         = 85

    ws.page_margins = PageMargins(
        left=0.5, right=0.5, top=0.75, bottom=0.75,
        header=0.3, footer=0.3
    )
    ws.oddHeader.left.text   = f"&\"Calibri,Bold\"&9{nm}"
    ws.oddHeader.center.text = f"&\"Calibri,Bold\"&9CONSOLIDATED BUSINESS PLAN"
    ws.oddHeader.right.text  = f"&9CONFIDENTIAL"
    ws.oddFooter.left.text   = f"&9Prepared by JRC Corporate Consulting"
    ws.oddFooter.center.text = f"&9Page &P of &N"
    ws.oddFooter.right.text  = f"&9{config.get('proj_start_year','2025')}–{config.get('proj_start_year',2025)+n-1}"

    # Print area will be set after we know how many rows we use
    ws.print_title_rows = "1:6"  # repeat header rows on each page

    # ── Column calibration for A4 landscape ─────────────────────────────────
    # A4 landscape ~279mm usable. With n=7 data cols + 2 label cols
    # Label col B = 28 chars wide, others = 11
    set_col_width(ws, "A", 1.5)
    set_col_width(ws, "B", 32)   # label
    set_col_width(ws, "C", 7)    # units
    for i in range(n):
        set_col_width(ws, get_column_letter(4 + i), 12)

    FIRST = 4   # data starts at column D for print layout (tighter)

    def col(idx):
        return get_column_letter(FIRST + idx)

    def end_col():
        return get_column_letter(FIRST + n - 1)

    def link(sheet, c_idx, r):
        return f"=IFERROR('{sheet}'!{col(c_idx)}{r},\"-\")"

    row = 1

    # ════════════════════════════════════════════════════════════════════════
    # COVER BANNER
    # ════════════════════════════════════════════════════════════════════════
    ws.row_dimensions[row].height = 45
    rng = f"B{row}:{end_col()}{row}"
    ws.merge_cells(rng)
    c = ws[f"B{row}"]
    c.value  = nm.upper()
    c.font   = Font(name="Calibri", size=18, bold=True, color=C["white"])
    c.fill   = fill(C["navy"])
    c.alignment = AL["center"]
    row += 1

    ws.row_dimensions[row].height = 20
    rng2 = f"B{row}:{end_col()}{row}"
    ws.merge_cells(rng2)
    c2 = ws[f"B{row}"]
    c2.value  = f"Consolidated Business Plan — {config.get('proj_start_year','2025')}–{config.get('proj_start_year',2025)+n-1}  |  {ccy}k unless stated"
    c2.font   = Font(name="Calibri", size=10, bold=False, color=C["white"])
    c2.fill   = fill(C["mid_blue"])
    c2.alignment = AL["center"]
    row += 1

    ws.row_dimensions[row].height = 8  # thin separator
    row += 1

    # ════════════════════════════════════════════════════════════════════════
    # YEAR HEADERS (repeated via print_title_rows)
    # ════════════════════════════════════════════════════════════════════════
    ws.row_dimensions[row].height = 18
    c = ws.cell(row=row, column=2, value=f"  {nm} — Consolidated P&L")
    style_cell(c, font=Font(name="Calibri", size=9, bold=True, color=C["white"]),
               fill_=fill(C["pl_header"]), align=AL["left"])
    c2 = ws.cell(row=row, column=3, value=unit)
    style_cell(c2, font=FONTS["header"], fill_=fill(C["pl_header"]), align=AL["center"])
    for i, yr in enumerate(years):
        c = ws.cell(row=row, column=FIRST + i, value=yr["fy"])
        style_cell(c, font=FONTS["header"], fill_=fill(C["pl_header"]), align=AL["center"])
    row += 1

    ws.row_dimensions[row].height = 13
    for i, yr in enumerate(years):
        lbl = "Actuals" if yr["is_actual"] else "BP"
        c = ws.cell(row=row, column=FIRST + i, value=lbl)
        clr = C["mid_grey"] if yr["is_actual"] else C["pale_blue"]
        style_cell(c, font=FONTS["note"], fill_=fill(clr), align=AL["center"])
    row += 1

    # ════════════════════════════════════════════════════════════════════════
    # HELPER: write a print row
    # ════════════════════════════════════════════════════════════════════════
    def prow(label, fmls, fnt, nf="k0", f_=None, brd=None, unit_="$k",
             indent=0, ht=15, show_unit=True):
        nonlocal row
        ws.row_dimensions[row].height = ht
        prefix = "  " * indent
        c = ws.cell(row=row, column=2, value=prefix + label)
        style_cell(c, font=fnt, align=AL["left"])
        if show_unit:
            cu = ws.cell(row=row, column=3, value=unit_)
            style_cell(cu, font=FONTS["italic"], align=AL["center"])
        for i, f in enumerate(fmls):
            c = ws.cell(row=row, column=FIRST + i, value=f)
            kw = dict(font=fnt, align=AL["right"], num_fmt=NF[nf])
            if f_:  kw["fill_"] = f_
            if brd: kw["border"] = brd
            style_cell(c, **kw)
        row += 1

    def section(title, color):
        nonlocal row
        ws.row_dimensions[row].height = 17
        rng = f"B{row}:{end_col()}{row}"
        ws.merge_cells(rng)
        c = ws[f"B{row}"]
        c.value = title
        c.font  = Font(name="Calibri", size=9, bold=True, color=C["white"])
        c.fill  = fill(C[color])
        c.alignment = AL["left"]
        row += 1

    def spacer(h=5):
        nonlocal row
        ws.row_dimensions[row].height = h
        row += 1

    def pl(k):  return row_map.get(k, 5)
    def ds(k):  return row_map.get(k, 5)

    # ════════════════════════════════════════════════════════════════════════
    # P&L SUMMARY
    # ════════════════════════════════════════════════════════════════════════
    section("PROFIT & LOSS", "pl_header")

    prow("Gross revenue",
         [f"=IFERROR('P&L'!{col(i)}{pl('pl_gross_revenue')},\"-\")" for i in range(n)],
         FONTS["label"], unit_=unit, indent=1)

    prow("Net realisation",
         [f"=IFERROR('P&L'!{col(i)}{pl('pl_net_real')},\"-\")" for i in range(n)],
         FONTS["label_bold"], nf="k0", brd=BORDERS["thin"], unit_=unit)

    prow("Total variable COGS",
         [f"=IFERROR(-'P&L'!{col(i)}{pl('pl_cogs_var')},\"-\")" for i in range(n)],
         FONTS["label"], unit_=unit, indent=1)

    prow("Gross profit",
         [f"=IFERROR('P&L'!{col(i)}{pl('pl_gross_profit')},\"-\")" for i in range(n)],
         FONTS["total"], nf="k0", f_=fill(C["total_fill"]), brd=BORDERS["thick"],
         unit_=unit, ht=18)

    prow("  Gross margin %",
         [f"=IFERROR('P&L'!{col(i)}{pl('pl_gross_profit')}/'P&L'!{col(i)}{pl('pl_net_real')},\"-\")"
          for i in range(n)],
         FONTS["italic"], nf="pct1", unit_="%", indent=1)

    spacer()
    prow("Fixed OpEx",
         [f"=IFERROR(-'P&L'!{col(i)}{pl('pl_fixed_opex')},\"-\")" for i in range(n)],
         FONTS["label"], unit_=unit, indent=1)

    prow("EBITDA (reported)",
         [f"=IFERROR('P&L'!{col(i)}{pl('pl_ebitda')},\"-\")" for i in range(n)],
         FONTS["total"], nf="k0", f_=fill(C["ebitda_fill"]), brd=BORDERS["thick"],
         unit_=unit, ht=18)

    prow("Adjusted EBITDA",
         [f"=IFERROR('P&L'!{col(i)}{pl('pl_ebitda_adj')},\"-\")" for i in range(n)],
         FONTS["total"], nf="k0", f_=fill(C["ebitda_fill"]), brd=BORDERS["thick"],
         unit_=unit, ht=18)

    prow("  Adj. EBITDA margin %",
         [f"=IFERROR('P&L'!{col(i)}{pl('pl_ebitda_adj')}/'P&L'!{col(i)}{pl('pl_net_real')},\"-\")"
          for i in range(n)],
         FONTS["italic"], nf="pct1", unit_="%", indent=1)

    spacer()
    prow("D&A",
         [f"=IFERROR(-'P&L'!{col(i)}{pl('pl_da')},\"-\")" for i in range(n)],
         FONTS["label"], unit_=unit, indent=1)

    prow("EBIT",
         [f"=IFERROR('P&L'!{col(i)}{pl('pl_ebit')},\"-\")" for i in range(n)],
         FONTS["total"], nf="k0", f_=fill(C["total_fill"]), brd=BORDERS["thick"],
         unit_=unit, ht=18)

    spacer()
    prow("Net finance costs",
         [f"=IFERROR(-'P&L'!{col(i)}{pl('pl_net_interest')},\"-\")" for i in range(n)],
         FONTS["label"], unit_=unit, indent=1)

    prow("Profit before tax",
         [f"=IFERROR('P&L'!{col(i)}{pl('pl_pbt')},\"-\")" for i in range(n)],
         FONTS["subtotal"], nf="k0", f_=fill(C["subtot_fill"]), brd=BORDERS["thin"],
         unit_=unit)

    prow("Tax",
         [f"=IFERROR(-'P&L'!{col(i)}{pl('pl_tax')},\"-\")" for i in range(n)],
         FONTS["label"], unit_=unit, indent=1)

    prow("Net income / (loss)",
         [f"=IFERROR('P&L'!{col(i)}{pl('pl_net_income')},\"-\")" for i in range(n)],
         FONTS["total"], nf="k0", f_=fill(C["total_fill"]), brd=BORDERS["thick"],
         unit_=unit, ht=20)

    prow("  Net margin %",
         [f"=IFERROR('P&L'!{col(i)}{pl('pl_net_income')}/'P&L'!{col(i)}{pl('pl_net_real')},\"-\")"
          for i in range(n)],
         FONTS["italic"], nf="pct1", unit_="%", indent=1)

    spacer(8)

    # ════════════════════════════════════════════════════════════════════════
    # CASH FLOW SUMMARY
    # ════════════════════════════════════════════════════════════════════════
    section("CASH FLOW", "cf_header")

    prow("Cash flow from operations",
         [f"=IFERROR('CASH FLOW'!{col(i)}{pl('cf_operating')},\"-\")" for i in range(n)],
         FONTS["label_bold"], unit_=unit)

    prow("Cash flow from investing (capex)",
         [f"=IFERROR('CASH FLOW'!{col(i)}{pl('cf_investing')},\"-\")" for i in range(n)],
         FONTS["label"], unit_=unit, indent=1)

    prow("Free cash flow (pre-debt service)",
         [f"=IFERROR('CASH FLOW'!{col(i)}{pl('cf_fcf')},\"-\")" for i in range(n)],
         FONTS["total"], nf="k0", f_=fill(C["ebitda_fill"]), brd=BORDERS["thick"],
         unit_=unit, ht=18)

    prow("Cash flow from financing",
         [f"=IFERROR('CASH FLOW'!{col(i)}{pl('cf_financing')},\"-\")" for i in range(n)],
         FONTS["label"], unit_=unit, indent=1)

    prow("Closing cash",
         [f"=IFERROR('CASH FLOW'!{col(i)}{pl('cf_closing_cash')},\"-\")" for i in range(n)],
         FONTS["total"], nf="k0", f_=fill(C["total_fill"]), brd=BORDERS["thick"],
         unit_=unit, ht=18)

    spacer(8)

    # ════════════════════════════════════════════════════════════════════════
    # BALANCE SHEET SUMMARY
    # ════════════════════════════════════════════════════════════════════════
    section("BALANCE SHEET", "bs_header")

    prow("Net PP&E",
         [f"=IFERROR(CAPEX!{col(i)}{row_map.get('capex_nbv',30)},\"-\")" for i in range(n)],
         FONTS["label"], unit_=unit, indent=1)

    prow("Total assets",
         [f"=IFERROR('BALANCE SHEET'!{col(i)}{row_map.get('bs_total_assets',30)},\"-\")" for i in range(n)],
         FONTS["label_bold"], unit_=unit)

    prow("Gross debt",
         [f"=IFERROR('DEBT SCHEDULE'!{col(i)}{ds('ds_total_debt')},\"-\")" for i in range(n)],
         FONTS["label"], unit_=unit, indent=1)

    prow("Net debt",
         [f"=IFERROR('DEBT SCHEDULE'!{col(i)}{ds('ds_total_debt')}-'CASH FLOW'!{col(i)}{pl('cf_closing_cash')},\"-\")"
          for i in range(n)],
         FONTS["label_bold"], nf="k0", brd=BORDERS["thin"], unit_=unit)

    prow("Total equity",
         [f"=IFERROR('BALANCE SHEET'!{col(i)}{row_map.get('bs_total_equity',45)},\"-\")" for i in range(n)],
         FONTS["label"], unit_=unit, indent=1)

    spacer(8)

    # ════════════════════════════════════════════════════════════════════════
    # KEY RATIOS
    # ════════════════════════════════════════════════════════════════════════
    section("KEY RATIOS & COVENANTS", "kpi_header")

    ratio_rows = [
        ("Net leverage (x)",      f"=IFERROR(('DEBT SCHEDULE'!{{c}}{ds('ds_total_debt')}-'CASH FLOW'!{{c}}{pl('cf_closing_cash')})/'P&L'!{{c}}{pl('pl_ebitda_adj')},\"-\")", "mult"),
        ("ICR (Adj EBITDA / Net int.)", f"=IFERROR('P&L'!{{c}}{pl('pl_ebitda_adj')}/'P&L'!{{c}}{pl('pl_net_interest')},\"-\")", "mult"),
        ("DSCR (x)",              f"=IFERROR('CASH FLOW'!{{c}}{pl('cf_operating')}/('P&L'!{{c}}{pl('pl_net_interest')}+'DEBT SCHEDULE'!{{c}}{ds('ds_total_repay')}),\"-\")", "mult"),
        ("Capacity utilisation",  f"=IFERROR(ASSUMPTIONS!{{c}}8/ASSUMPTIONS!{{c}}7,\"-\")", "pct1"),
        ("Capex / EBITDA",        f"=IFERROR(CAPEX!{{c}}{row_map.get('capex_total',20)}/'P&L'!{{c}}{pl('pl_ebitda_adj')},\"-\")", "pct1"),
        ("EBITDA / MT ($/MT)",    f"=IFERROR('P&L'!{{c}}{pl('pl_ebitda_adj')}*1000/ASSUMPTIONS!{{c}}8,\"-\")", "k1"),
    ]

    for label, fml_tpl, fmt in ratio_rows:
        fmls = [fml_tpl.replace("{c}", col(i)) for i in range(n)]
        u = "x" if fmt == "mult" else ("%" if fmt == "pct1" else "$/MT")
        prow(label, fmls, FONTS["formula"], nf=fmt, unit_=u)

    spacer(8)

    # ════════════════════════════════════════════════════════════════════════
    # DEBT STRUCTURE MEMO
    # ════════════════════════════════════════════════════════════════════════
    section("DEBT STRUCTURE — MEMORANDUM", "debt_header")

    debt_memo = [
        ("Total gross debt",        f"=IFERROR('DEBT SCHEDULE'!{{c}}{ds('ds_total_debt')},\"-\")",  "k0"),
        ("Cash interest charge",    f"=IFERROR('DEBT SCHEDULE'!{{c}}{ds('ds_cash_interest')},\"-\")", "k0"),
        ("Mandatory repayment",     f"=IFERROR('DEBT SCHEDULE'!{{c}}{ds('ds_total_repay')},\"-\")",  "k0"),
        ("Cash sweep",              f"=IFERROR('DEBT SCHEDULE'!{{c}}{row_map.get('ds_cash_sweep',20)},\"-\")", "k0"),
    ]

    for label, fml_tpl, fmt in debt_memo:
        fmls = [fml_tpl.replace("{c}", col(i)) for i in range(n)]
        prow(label, fmls, FONTS["formula"], nf=fmt, unit_=unit, indent=1)

    spacer(10)

    # ════════════════════════════════════════════════════════════════════════
    # DISCLAIMER / FOOTER NOTE
    # ════════════════════════════════════════════════════════════════════════
    ws.row_dimensions[row].height = 24
    rng_dis = f"B{row}:{end_col()}{row}"
    ws.merge_cells(rng_dis)
    c_dis = ws[f"B{row}"]
    c_dis.value = (
        "CONFIDENTIAL — This document has been prepared for information purposes only and does not constitute "
        "financial advice. Projections are based on assumptions which may not be realised. "
        "Prepared by JRC Corporate Consulting, DIFC, Dubai."
    )
    c_dis.font  = Font(name="Calibri", size=7, italic=True, color=C["dark_grey"])
    c_dis.fill  = fill(C["light_grey"])
    c_dis.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    row += 1

    # Set print area
    ws.print_area = f"A1:{end_col()}{row}"

    ws.sheet_properties.tabColor = "0D1B3E"
    return ws
