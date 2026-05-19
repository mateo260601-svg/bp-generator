from openpyxl.utils import get_column_letter
from styles import *

FIRST_DATA_COL = 5
LABEL_COL = 2
UNIT_COL  = 3
SRC_COL   = 4
ASSUMP    = "ASSUMPTIONS"

def _link(sheet, col, row):
    return f"='{sheet}'!{col}{row}"

def _safe_div(num, den):
    return f"=IFERROR({num}/{den},\"-\")"

def _col(idx):
    return get_column_letter(FIRST_DATA_COL + idx)

def _row_label(ws, row, indent, label, font, row_ht=16):
    ws.row_dimensions[row].height = row_ht
    prefix = "    " * indent
    c = ws.cell(row=row, column=LABEL_COL, value=prefix + label)
    style_cell(c, font=font, align=AL["left"])

def _data_row(ws, row, indent, label, formulas, font, num_fmt="k0",
              fill_=None, unit="$k", border=None, row_ht=16):
    ws.row_dimensions[row].height = row_ht
    _row_label(ws, row, indent, label, font, row_ht=row_ht)
    c_unit = ws.cell(row=row, column=UNIT_COL, value=unit)
    style_cell(c_unit, font=FONTS["italic"], align=AL["center"])
    for i, f in enumerate(formulas):
        c = ws.cell(row=row, column=FIRST_DATA_COL + i, value=f)
        kw = dict(font=font, align=AL["right"], num_fmt=NF[num_fmt])
        if fill_:   kw["fill_"] = fill_
        if border:  kw["border"] = border
        style_cell(c, **kw)

def build_pl(wb, config, years, row_map):
    """
    row_map: dict we populate with row numbers so other sheets can reference them.
    Keys: pl_revenue, pl_net_real, pl_gross_profit, pl_ebitda, pl_ebitda_adj,
          pl_da, pl_ebit, pl_net_interest, pl_pbt, pl_tax, pl_net_income
    """
    ws = wb.create_sheet("P&L")
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90

    n = len(years)

    # Column widths
    set_col_width(ws, "A", 2)
    set_col_width(ws, "B", 38)
    set_col_width(ws, "C", 8)
    set_col_width(ws, "D", 10)
    for i in range(n):
        set_col_width(ws, get_column_letter(FIRST_DATA_COL + i), 14)
    set_col_width(ws, get_column_letter(FIRST_DATA_COL + n), 28)

    ccy = config["currency"]
    unit = f"{ccy}k"
    nm = config["company_name"]

    row = 1

    # Title
    ws.row_dimensions[row].height = 34
    end = get_column_letter(FIRST_DATA_COL + n - 1)
    merge_and_style(ws, f"B{row}:{end}{row}",
                    f"{nm} — Profit & Loss Statement",
                    FONTS["title"], fill(C["navy"]), AL["center"])
    row += 2

    # Period headers
    ws.row_dimensions[row].height = 18
    c = ws.cell(row=row, column=LABEL_COL, value="")
    style_cell(c, font=FONTS["header"], fill_=fill(C["pl_header"]), align=AL["center"])
    ws.cell(row=row, column=UNIT_COL, value="Units").font = FONTS["header"]
    ws.cell(row=row, column=UNIT_COL).fill = fill(C["pl_header"])
    ws.cell(row=row, column=UNIT_COL).alignment = AL["center"]
    ws.cell(row=row, column=SRC_COL, value="Source").font = FONTS["header"]
    ws.cell(row=row, column=SRC_COL).fill = fill(C["pl_header"])
    ws.cell(row=row, column=SRC_COL).alignment = AL["center"]
    for i, yr in enumerate(years):
        c = ws.cell(row=row, column=FIRST_DATA_COL + i, value=yr["fy"])
        style_cell(c, font=FONTS["header"], fill_=fill(C["pl_header"]), align=AL["center"])
    row += 1

    ws.row_dimensions[row].height = 14
    for i, yr in enumerate(years):
        lbl = "Actuals" if yr["is_actual"] else "BP"
        c = ws.cell(row=row, column=FIRST_DATA_COL + i, value=lbl)
        clr = C["mid_grey"] if yr["is_actual"] else C["pale_blue"]
        style_cell(c, font=FONTS["note"], fill_=fill(clr), align=AL["center"])
    row += 1

    # ──────────────────────────────────────────────────────────────
    # REVENUE
    # ──────────────────────────────────────────────────────────────
    ws.row_dimensions[row].height = 19
    end = get_column_letter(FIRST_DATA_COL + n - 1)
    merge_and_style(ws, f"B{row}:{end}{row}", "REVENUE",
                    FONTS["section"], fill(C["pl_header"]), AL["left"])
    row += 1

    # Gross revenue — pulled from ASSUMPTIONS via formula
    # In a real build these reference the output of the Assumptions calcs
    # Here we write the formula referencing the Assumptions sheet rows
    assump_gross_row  = 17  # row in ASSUMPTIONS where gross revenue sits (calibrated)
    assump_net_row    = 20  # net realisation row
    assump_freight_row= 18
    assump_comm_row   = 19

    gross_fmls  = [f"=ASSUMPTIONS!{_col(i)}{assump_gross_row}"  for i in range(n)]
    freight_fmls= [f"=ASSUMPTIONS!{_col(i)}{assump_freight_row}" for i in range(n)]
    comm_fmls   = [f"=ASSUMPTIONS!{_col(i)}{assump_comm_row}"   for i in range(n)]
    net_fmls    = [f"=ASSUMPTIONS!{_col(i)}{assump_net_row}"    for i in range(n)]

    _data_row(ws, row, 0, "Gross revenue",           gross_fmls,   FONTS["label"],     unit=unit)
    row_map["pl_gross_revenue"] = row; row += 1

    _data_row(ws, row, 1, "Less: freight & forwarding",freight_fmls,FONTS["label"],    unit=unit)
    row_map["pl_freight"] = row; row += 1

    _data_row(ws, row, 1, "Less: sales commission",  comm_fmls,    FONTS["label"],     unit=unit)
    row_map["pl_commission"] = row; row += 1

    # Net realisation — TOTAL row
    net_rev_r = row
    net_fmls2 = [f"={_col(i)}{row_map['pl_gross_revenue']}-{_col(i)}{row_map['pl_freight']}-{_col(i)}{row_map['pl_commission']}"
                 for i in range(n)]
    _data_row(ws, row, 0, "Net realisation", net_fmls2, FONTS["subtotal"],
              fill_=fill(C["subtot_fill"]), unit=unit, border=BORDERS["thin"])
    row_map["pl_net_real"] = row; row += 1
    row_map["pl_revenue"] = row_map["pl_net_real"]

    row += 1  # spacer

    # ──────────────────────────────────────────────────────────────
    # COST OF GOODS SOLD
    # ──────────────────────────────────────────────────────────────
    ws.row_dimensions[row].height = 19
    merge_and_style(ws, f"B{row}:{end}{row}", "COST OF GOODS SOLD",
                    FONTS["section"], fill(C["pl_header"]), AL["left"])
    row += 1

    # Assumption rows in ASSUMPTIONS sheet (calibrated offsets)
    assum_vol_row  = 8   # sales volume
    assum_dm_row   = 23  # direct materials $/MT
    assum_util_row = 27  # utilities $/MT
    assum_pack_row = 28  # packing $/MT
    assum_var_row  = 29  # variable opex other $/MT

    # Direct materials = volume × $/MT / 1000  (result in $k)
    dm_fmls   = [f"=ASSUMPTIONS!{_col(i)}{assum_vol_row}*ASSUMPTIONS!{_col(i)}{assum_dm_row}/1000"  for i in range(n)]
    util_fmls = [f"=ASSUMPTIONS!{_col(i)}{assum_vol_row}*ASSUMPTIONS!{_col(i)}{assum_util_row}/1000" for i in range(n)]
    pack_fmls = [f"=ASSUMPTIONS!{_col(i)}{assum_vol_row}*ASSUMPTIONS!{_col(i)}{assum_pack_row}/1000" for i in range(n)]
    var_fmls  = [f"=ASSUMPTIONS!{_col(i)}{assum_vol_row}*ASSUMPTIONS!{_col(i)}{assum_var_row}/1000"  for i in range(n)]

    _data_row(ws, row, 1, "Direct materials",           dm_fmls,   FONTS["label"], unit=unit)
    row_map["pl_direct_mat"] = row; row += 1

    _data_row(ws, row, 1, "Utilities",                  util_fmls, FONTS["label"], unit=unit)
    row_map["pl_utilities"] = row; row += 1

    _data_row(ws, row, 1, "Packing costs",              pack_fmls, FONTS["label"], unit=unit)
    row_map["pl_packing"] = row; row += 1

    _data_row(ws, row, 1, "Variable OpEx — other",      var_fmls,  FONTS["label"], unit=unit)
    row_map["pl_var_opex"] = row; row += 1

    # Total COGS (variable)
    cogs_rows = ["pl_direct_mat","pl_utilities","pl_packing","pl_var_opex"]
    cogs_fmls = [f"=SUM({_col(i)}{row_map['pl_direct_mat']}:{_col(i)}{row_map['pl_var_opex']})"
                 for i in range(n)]
    _data_row(ws, row, 0, "Total variable COGS", cogs_fmls, FONTS["subtotal"],
              fill_=fill(C["subtot_fill"]), unit=unit, border=BORDERS["thin"])
    row_map["pl_cogs_var"] = row; row += 1

    # Gross profit
    gp_fmls = [f"={_col(i)}{row_map['pl_net_real']}-{_col(i)}{row_map['pl_cogs_var']}"
               for i in range(n)]
    _data_row(ws, row, 0, "Gross profit", gp_fmls, FONTS["total"],
              fill_=fill(C["total_fill"]), unit=unit, border=BORDERS["thick"])
    row_map["pl_gross_profit"] = row; row += 2

    # ──────────────────────────────────────────────────────────────
    # OPERATING EXPENSES (FIXED)
    # ──────────────────────────────────────────────────────────────
    ws.row_dimensions[row].height = 19
    merge_and_style(ws, f"B{row}:{end}{row}", "FIXED OPERATING EXPENSES",
                    FONTS["section"], fill(C["pl_header"]), AL["left"])
    row += 1

    assum_staff_prod  = 32
    assum_staff_sga   = 33
    assum_maint       = 35
    assum_insur       = 36
    assum_rent        = 37
    assum_it          = 38
    assum_prof        = 39
    assum_other       = 40
    assum_restr       = 41

    def assum_fml(assum_row):
        return [f"=ASSUMPTIONS!{_col(i)}{assum_row}" for i in range(n)]

    _data_row(ws, row, 1, "Personnel — production",   assum_fml(assum_staff_prod), FONTS["label"], unit=unit)
    row_map["pl_staff_prod"] = row; row += 1

    _data_row(ws, row, 1, "Personnel — SG&A",         assum_fml(assum_staff_sga),  FONTS["label"], unit=unit)
    row_map["pl_staff_sga"] = row; row += 1

    _data_row(ws, row, 1, "Maintenance & repairs",    assum_fml(assum_maint),      FONTS["label"], unit=unit)
    row_map["pl_maint"] = row; row += 1

    _data_row(ws, row, 1, "Insurance",                assum_fml(assum_insur),      FONTS["label"], unit=unit)
    row_map["pl_insur"] = row; row += 1

    _data_row(ws, row, 1, "Rent / site costs",        assum_fml(assum_rent),       FONTS["label"], unit=unit)
    row_map["pl_rent"] = row; row += 1

    _data_row(ws, row, 1, "IT & systems",             assum_fml(assum_it),         FONTS["label"], unit=unit)
    row_map["pl_it"] = row; row += 1

    _data_row(ws, row, 1, "Professional fees",        assum_fml(assum_prof),       FONTS["label"], unit=unit)
    row_map["pl_prof"] = row; row += 1

    _data_row(ws, row, 1, "Other SG&A",               assum_fml(assum_other),      FONTS["label"], unit=unit)
    row_map["pl_other_sga"] = row; row += 1

    # Total fixed opex
    fix_start = row_map["pl_staff_prod"]
    fix_end   = row_map["pl_other_sga"]
    fix_fmls  = [f"=SUM({_col(i)}{fix_start}:{_col(i)}{fix_end})" for i in range(n)]
    _data_row(ws, row, 0, "Total fixed OpEx", fix_fmls, FONTS["subtotal"],
              fill_=fill(C["subtot_fill"]), unit=unit, border=BORDERS["thin"])
    row_map["pl_fixed_opex"] = row; row += 1

    # EBITDA (reported)
    ebitda_fmls = [f"={_col(i)}{row_map['pl_gross_profit']}-{_col(i)}{row_map['pl_fixed_opex']}"
                   for i in range(n)]
    _data_row(ws, row, 0, "EBITDA (reported)", ebitda_fmls, FONTS["total"],
              fill_=fill(C["ebitda_fill"]), unit=unit, border=BORDERS["thick"], row_ht=20)
    row_map["pl_ebitda"] = row; row += 1

    # EBITDA margin
    ebitda_mgn = [f"=IFERROR({_col(i)}{row_map['pl_ebitda']}/{_col(i)}{row_map['pl_net_real']},\"-\")"
                  for i in range(n)]
    _data_row(ws, row, 1, "EBITDA margin %", ebitda_mgn, FONTS["italic"],
              num_fmt="pct1", unit="%")
    row += 1

    # Restructuring / one-off (addback)
    restr_fmls = assum_fml(assum_restr)
    _data_row(ws, row, 1, "Add: restructuring / one-off costs", restr_fmls,
              FONTS["label"], unit=unit)
    row_map["pl_restr"] = row; row += 1

    # Adjusted EBITDA
    adj_fmls = [f"={_col(i)}{row_map['pl_ebitda']}+{_col(i)}{row_map['pl_restr']}"
                for i in range(n)]
    _data_row(ws, row, 0, "Adjusted EBITDA", adj_fmls, FONTS["total"],
              fill_=fill(C["ebitda_fill"]), unit=unit, border=BORDERS["thick"], row_ht=20)
    row_map["pl_ebitda_adj"] = row; row += 1

    adj_mgn = [f"=IFERROR({_col(i)}{row_map['pl_ebitda_adj']}/{_col(i)}{row_map['pl_net_real']},\"-\")"
               for i in range(n)]
    _data_row(ws, row, 1, "Adj. EBITDA margin %", adj_mgn, FONTS["italic"],
              num_fmt="pct1", unit="%")
    row += 2

    # ──────────────────────────────────────────────────────────────
    # D&A
    # ──────────────────────────────────────────────────────────────
    ws.row_dimensions[row].height = 19
    merge_and_style(ws, f"B{row}:{end}{row}", "DEPRECIATION & AMORTISATION",
                    FONTS["section"], fill(C["pl_header"]), AL["left"])
    row += 1

    # D&A from CAPEX sheet (cross-sheet link)
    da_fmls = [f"=IFERROR('CAPEX'!{_col(i)}20,0)" for i in range(n)]  # row 20 in CAPEX = total D&A
    _data_row(ws, row, 1, "Depreciation of PP&E", da_fmls, FONTS["link"], unit=unit)
    row_map["pl_depreciation"] = row; row += 1

    amort_fmls = [f"=0" for i in range(n)]  # placeholder for intangible amortisation
    _data_row(ws, row, 1, "Amortisation of intangibles", amort_fmls, FONTS["label"], unit=unit)
    row_map["pl_amortisation"] = row; row += 1

    da_total_fmls = [f"={_col(i)}{row_map['pl_depreciation']}+{_col(i)}{row_map['pl_amortisation']}"
                     for i in range(n)]
    _data_row(ws, row, 0, "Total D&A", da_total_fmls, FONTS["subtotal"],
              fill_=fill(C["subtot_fill"]), unit=unit, border=BORDERS["thin"])
    row_map["pl_da"] = row; row += 1

    # EBIT
    ebit_fmls = [f"={_col(i)}{row_map['pl_ebitda']}-{_col(i)}{row_map['pl_da']}"
                 for i in range(n)]
    _data_row(ws, row, 0, "EBIT (Operating profit)", ebit_fmls, FONTS["total"],
              fill_=fill(C["total_fill"]), unit=unit, border=BORDERS["thick"], row_ht=20)
    row_map["pl_ebit"] = row; row += 2

    # ──────────────────────────────────────────────────────────────
    # NET FINANCE COSTS
    # ──────────────────────────────────────────────────────────────
    ws.row_dimensions[row].height = 19
    merge_and_style(ws, f"B{row}:{end}{row}", "NET FINANCE COSTS",
                    FONTS["section"], fill(C["debt_header"]), AL["left"])
    row += 1

    int_fmls      = [f"=IFERROR('DEBT SCHEDULE'!{_col(i)}50,0)" for i in range(n)]
    pik_fmls      = [f"=IFERROR('DEBT SCHEDULE'!{_col(i)}51,0)" for i in range(n)]
    fee_fmls      = [f"=IFERROR('DEBT SCHEDULE'!{_col(i)}52,0)" for i in range(n)]
    int_inc_fmls  = [f"=0" for _ in range(n)]

    _data_row(ws, row, 1, "Interest expense — cash",   int_fmls,     FONTS["link"], unit=unit)
    row_map["pl_int_cash"] = row; row += 1

    _data_row(ws, row, 1, "Interest expense — PIK",    pik_fmls,     FONTS["link"], unit=unit)
    row_map["pl_int_pik"] = row; row += 1

    _data_row(ws, row, 1, "Debt fees & OID amortised", fee_fmls,     FONTS["link"], unit=unit)
    row_map["pl_debt_fees"] = row; row += 1

    _data_row(ws, row, 1, "Interest income",           int_inc_fmls, FONTS["label"], unit=unit)
    row_map["pl_int_income"] = row; row += 1

    net_int_fmls = [f"={_col(i)}{row_map['pl_int_cash']}+{_col(i)}{row_map['pl_int_pik']}+"
                    f"{_col(i)}{row_map['pl_debt_fees']}-{_col(i)}{row_map['pl_int_income']}"
                    for i in range(n)]
    _data_row(ws, row, 0, "Net finance costs", net_int_fmls, FONTS["subtotal"],
              fill_=fill(C["subtot_fill"]), unit=unit, border=BORDERS["thin"])
    row_map["pl_net_interest"] = row; row += 1

    # PBT
    pbt_fmls = [f"={_col(i)}{row_map['pl_ebit']}-{_col(i)}{row_map['pl_net_interest']}"
                for i in range(n)]
    _data_row(ws, row, 0, "Profit before tax (PBT)", pbt_fmls, FONTS["total"],
              fill_=fill(C["total_fill"]), unit=unit, border=BORDERS["thick"], row_ht=20)
    row_map["pl_pbt"] = row; row += 2

    # ──────────────────────────────────────────────────────────────
    # TAX
    # ──────────────────────────────────────────────────────────────
    ws.row_dimensions[row].height = 19
    merge_and_style(ws, f"B{row}:{end}{row}", "TAXATION",
                    FONTS["section"], fill(C["pl_header"]), AL["left"])
    row += 1

    tax_rate_row = 45  # ASSUMPTIONS row for tax rate
    curr_tax_fmls = [
        f"=IFERROR(IF({_col(i)}{row_map['pl_pbt']}>0,"
        f"ASSUMPTIONS!{_col(i)}{tax_rate_row}*{_col(i)}{row_map['pl_pbt']},0),0)"
        for i in range(n)]
    dt_fmls = [f"=0" for _ in range(n)]  # deferred tax placeholder

    _data_row(ws, row, 1, "Current tax charge",     curr_tax_fmls, FONTS["formula"], unit=unit)
    row_map["pl_curr_tax"] = row; row += 1

    _data_row(ws, row, 1, "Deferred tax (charge)/credit", dt_fmls, FONTS["label"], unit=unit)
    row_map["pl_def_tax"] = row; row += 1

    tax_fmls = [f"={_col(i)}{row_map['pl_curr_tax']}+{_col(i)}{row_map['pl_def_tax']}"
                for i in range(n)]
    _data_row(ws, row, 0, "Total tax charge", tax_fmls, FONTS["subtotal"],
              fill_=fill(C["subtot_fill"]), unit=unit, border=BORDERS["thin"])
    row_map["pl_tax"] = row; row += 1

    # Net income
    ni_fmls = [f"={_col(i)}{row_map['pl_pbt']}-{_col(i)}{row_map['pl_tax']}"
               for i in range(n)]
    _data_row(ws, row, 0, "Net income / (loss)", ni_fmls, FONTS["total"],
              fill_=fill(C["total_fill"]), unit=unit, border=BORDERS["thick"], row_ht=22)
    row_map["pl_net_income"] = row; row += 1

    ni_mgn = [f"=IFERROR({_col(i)}{row_map['pl_net_income']}/{_col(i)}{row_map['pl_net_real']},\"-\")"
              for i in range(n)]
    _data_row(ws, row, 1, "Net margin %", ni_mgn, FONTS["italic"],
              num_fmt="pct1", unit="%")
    row += 2

    # ──────────────────────────────────────────────────────────────
    # MEMO / KPI BRIDGE
    # ──────────────────────────────────────────────────────────────
    ws.row_dimensions[row].height = 19
    merge_and_style(ws, f"B{row}:{end}{row}", "MEMORANDUM — KEY METRICS",
                    FONTS["section"], fill(C["kpi_header"]), AL["left"])
    row += 1

    vol_fmls = [f"=ASSUMPTIONS!{_col(i)}{assum_vol_row}" for i in range(n)]
    _data_row(ws, row, 0, "Sales volume", vol_fmls, FONTS["label"], unit="MT", num_fmt="k0")
    row += 1

    spread_fmls = [
        f"=IFERROR(({_col(i)}{row_map['pl_net_real']}-{_col(i)}{row_map['pl_cogs_var']})"
        f"/(ASSUMPTIONS!{_col(i)}{assum_vol_row}/1000),\"-\")"
        for i in range(n)]
    _data_row(ws, row, 0, "Net spread (gross profit / MT)", spread_fmls,
              FONTS["label"], unit="$/MT", num_fmt="k1")
    row += 1

    ebitda_mt = [f"=IFERROR({_col(i)}{row_map['pl_ebitda_adj']}*1000/"
                 f"ASSUMPTIONS!{_col(i)}{assum_vol_row},\"-\")"
                 for i in range(n)]
    _data_row(ws, row, 0, "Adj. EBITDA per MT", ebitda_mt,
              FONTS["label"], unit="$/MT", num_fmt="k1")
    row += 1

    ws.sheet_properties.tabColor = "1A3560"
    return ws
