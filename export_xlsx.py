"""Builds the branded TGFC Buy/Sell Excel sheet, filled in from a saved Buy/Sell.

Same layout, colors and formulas as the Buy/Sell template (build_bs.py).
full=True  -> the whole sheet, with live formulas.
full=False -> the sale-side copy: vendor, cost, costs and profit are left out,
              so a sale-side person can never open the buy side in Excel.
"""
from __future__ import annotations

import io
import os
from datetime import date, datetime

import xlsxwriter

import calc

NAVY = "#1A1F4B"; TINT = "#EEF0F7"; INPUT = "#EBEBEB"; BORDER = "#C8CBD8"
TEXT = "#22253A"; MUTED = "#5A5F78"; FONT = "Arial"
LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")

CUR = '$#,##0.00;($#,##0.00);"-"'
CUR_HIDE0 = '$#,##0.00;($#,##0.00);;'
RATE = '$#,##0.0000;($#,##0.0000);"-"'
RATE_IN = '$#,##0.0000'
NUM = '#,##0;(#,##0);"-"'
PCT = '0.00%'
DATE = 'm/d/yyyy'


def _date(v):
    if not v:
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    try:
        d = date.fromisoformat(str(v)[:10])
        return datetime(d.year, d.month, d.day)
    except Exception:
        return None


def _num(v):
    try:
        if v is None or v == "":
            return None
        f = float(v)
        return None if f != f else f
    except Exception:
        return None


def build(bs: dict, lines: list[dict], full: bool = True) -> bytes:
    lines = [ln for ln in lines if calc.line_has_data(ln)]
    t = calc.totals(bs, lines if full else [{**ln, "cost": None} for ln in lines])
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    ws = wb.add_worksheet("Buy Sell")

    def F(**k):
        base = {"font_name": FONT, "font_size": 9, "font_color": TEXT, "valign": "vcenter"}
        base.update(k)
        return wb.add_format(base)

    thin = {"border": 1, "border_color": BORDER}
    f_label = F(bold=True, font_color=NAVY, bg_color=TINT, **thin)
    f_in = F(bg_color=INPUT, **thin)
    f_in_c = F(bg_color=INPUT, align="center", **thin)
    f_in_date = F(bg_color=INPUT, align="center", num_format=DATE, **thin)
    f_in_num = F(bg_color=INPUT, num_format='#,##0', **thin)
    f_in_rate = F(bg_color=INPUT, num_format=RATE_IN, **thin)
    f_in_cur = F(bg_color=INPUT, num_format='$#,##0.00', **thin)
    f_in_pct = F(bg_color=INPUT, num_format=PCT, **thin)
    f_hidden = F(font_color=MUTED, align="center", bg_color="#F6F6F8", **thin)
    f_calc_cur = F(num_format=CUR, **thin)
    f_calc_cur0 = F(num_format=CUR_HIDE0, **thin)
    f_th = F(bold=True, font_color="white", bg_color=NAVY, align="center", text_wrap=True, border=1, border_color=NAVY)
    f_tot_lbl = F(bold=True, font_color="white", bg_color=NAVY, align="left", border=1, border_color=NAVY)
    tot_border = dict(top=2, top_color=NAVY, bottom=1, bottom_color=BORDER, left=1, left_color=BORDER, right=1, right_color=BORDER)
    f_tot_num = F(bold=True, num_format=NUM, bg_color=TINT, **tot_border)
    f_tot_cur = F(bold=True, num_format=CUR, bg_color=TINT, **tot_border)
    f_tot_rate = F(bold=True, num_format=RATE, bg_color=TINT, **tot_border)
    f_tot_hidden = F(bold=True, font_color=MUTED, align="center", bg_color=TINT, **tot_border)
    f_sec = F(bold=True, font_size=9, font_color=NAVY, bottom=2, bottom_color=NAVY)
    f_key_lbl = F(bold=True, font_color="white", bg_color=NAVY, border=1, border_color=NAVY)
    f_key_cur = F(bold=True, font_size=10, num_format=CUR, bg_color=TINT, border=2, border_color=NAVY)
    f_key_pct = F(bold=True, font_size=10, num_format=PCT, bg_color=TINT, border=2, border_color=NAVY)
    f_key_hidden = F(bold=True, font_size=10, font_color=MUTED, align="center", bg_color=TINT, border=2, border_color=NAVY)
    HID = "—"

    def put(rng, value, fmt, kind="text"):
        """Write an input value into a cell or merged range."""
        if kind == "date":
            value = _date(value)
        elif kind == "num":
            value = _num(value)
        else:
            value = "" if value is None else str(value)
        if ":" in rng:
            if kind == "date" and value is not None:
                ws.merge_range(rng, "", fmt)
                ws.write_datetime(rng.split(":")[0], value, fmt)
            else:
                ws.merge_range(rng, "" if value is None else value, fmt)
        else:
            if value is None or value == "":
                ws.write_blank(rng, None, fmt)
            elif kind == "date":
                ws.write_datetime(rng, value, fmt)
            elif kind == "num":
                ws.write_number(rng, value, fmt)
            else:
                ws.write_string(rng, value, fmt)

    # ---- columns (A..K) ----
    for i, w in enumerate([15, 13, 12, 10, 10, 11, 11, 11, 13, 11, 14]):
        ws.set_column(i, i, w)
    ws.hide_gridlines(2)

    # ---- header ----
    for r in range(4):
        ws.set_row(r, 17)
    if os.path.exists(LOGO):
        ws.insert_image("A1", LOGO, {"x_scale": 0.5, "y_scale": 0.5, "x_offset": 4, "y_offset": 2, "object_position": 3})
    ws.merge_range("E1:K1", "THIRD GENERATION FOOD COMPANY", F(bold=True, font_size=13, font_color=NAVY, align="right"))
    ws.merge_range("E2:K2", "7686 Addisen Path, Inver Grove Heights, MN 55077", F(font_size=8.5, align="right"))
    ws.merge_range("E3:K3", "www.thetgfc.com", F(font_size=8.5, align="right"))
    ws.set_row(4, 6)
    for c in range(11):
        ws.write_blank(4, c, None, F(bottom=5, bottom_color=NAVY))
    ws.set_row(5, 6)
    ws.set_row(6, 13); ws.set_row(7, 26)
    ws.merge_range("A7:K7", "TGFC  ·  ORDER RECAP" + ("" if full else "  ·  SALE-SIDE COPY"),
                   F(bold=True, font_size=7.5, font_color="#C9CDE6", bg_color=NAVY, indent=1))
    ws.merge_range("A8:K8", "Buy / Sell Sheet", F(bold=True, font_size=18, font_color="white", bg_color=NAVY, indent=1))
    ws.set_row(8, 6)

    # ---- PURCHASE ----
    ws.merge_range("A10:K10", "PURCHASE", f_sec)
    ws.set_row(9, 16)
    for r in (10, 11, 12):
        ws.set_row(r, 17)
    ws.write("A11", "TGFC PO", f_label); put("B11:C11", bs.get("po"), f_in_c)
    ws.merge_range("D11:E11", "Vendor", f_label)
    ws.write("I11", "Contact Info", f_label)
    ws.write("A12", "Order Date", f_label); put("B12:C12", bs.get("order_date"), f_in_date, "date")
    ws.merge_range("D12:E12", "P/U Date", f_label); put("F12:H12", bs.get("pu_date"), f_in_date, "date")
    ws.write("I12", "Release #", f_label)
    ws.write("A13", "P/U Location", f_label)
    if full:
        put("F11:H11", bs.get("vendor"), f_in)
        put("J11:K11", bs.get("contact"), f_in)
        put("J12:K12", bs.get("release_no"), f_in_c)
        put("B13:K13", bs.get("pu_location"), f_in)
    else:
        ws.merge_range("F11:H11", HID, f_hidden)
        ws.merge_range("J11:K11", HID, f_hidden)
        ws.merge_range("J12:K12", HID, f_hidden)
        ws.merge_range("B13:K13", HID, f_hidden)
    ws.set_row(13, 6)

    # ---- LINES ----
    HR = 15
    ws.set_row(HR - 1, 24)
    hdr = {"A": "Item Code", "F": "Cases", "G": "Pounds", "H": "Cost\n($/lb)", "I": "$ Cost",
           "J": "Sale Price\n($/lb)", "K": "$ Sale Price"}
    for col, txt in hdr.items():
        ws.write(f"{col}{HR}", txt, f_th)
    ws.merge_range(f"B{HR}:E{HR}", "Item Description", f_th)
    NLINES = max(22, len(lines))
    FIRST = HR + 1
    LAST = FIRST + NLINES - 1
    TOT = LAST + 1
    for i in range(NLINES):
        r = FIRST + i
        ws.set_row(r - 1, 15)
        d = lines[i] if i < len(lines) else {}
        put(f"A{r}", d.get("item_code"), f_in)
        put(f"B{r}:E{r}", d.get("description"), f_in)
        put(f"F{r}", d.get("cases"), f_in_num, "num")
        put(f"G{r}", d.get("pounds"), f_in_num, "num")
        put(f"J{r}", d.get("sale_price"), f_in_rate, "num")
        lbs = calc.n(d.get("pounds"))
        ws.write_formula(f"K{r}", f"=G{r}*J{r}", f_calc_cur0, lbs * calc.n(d.get("sale_price")))
        if full:
            put(f"H{r}", d.get("cost"), f_in_rate, "num")
            ws.write_formula(f"I{r}", f"=G{r}*H{r}", f_calc_cur0, lbs * calc.n(d.get("cost")))
        else:
            ws.write(f"H{r}", HID if d else "", f_hidden)
            ws.write(f"I{r}", HID if d else "", f_hidden)

    ws.set_row(TOT - 1, 17)
    ws.merge_range(f"A{TOT}:E{TOT}", "Totals", f_tot_lbl)
    ws.write_formula(f"F{TOT}", f"=SUM(F{FIRST}:F{LAST})", f_tot_num, t["cases"])
    ws.write_formula(f"G{TOT}", f"=SUM(G{FIRST}:G{LAST})", f_tot_num, t["pounds"])
    ws.write_formula(f"J{TOT}", f"=IF(G{TOT}=0,0,K{TOT}/G{TOT})", f_tot_rate, t["avg_sale"])
    ws.write_formula(f"K{TOT}", f"=SUM(K{FIRST}:K{LAST})", f_tot_cur, t["sale_total"])
    if full:
        ws.write_formula(f"H{TOT}", f"=IF(G{TOT}=0,0,I{TOT}/G{TOT})", f_tot_rate, t["avg_cost"])
        ws.write_formula(f"I{TOT}", f"=SUM(I{FIRST}:I{LAST})", f_tot_cur, t["cost_total"])
    else:
        ws.write(f"H{TOT}", HID, f_tot_hidden)
        ws.write(f"I{TOT}", HID, f_tot_hidden)
    ws.set_row(TOT, 6)

    # ---- SALE ----
    SR = TOT + 2
    ws.set_row(SR - 1, 16)
    ws.merge_range(f"A{SR}:K{SR}", "SALE", f_sec)
    ws.set_row(SR, 17); ws.set_row(SR + 1, 17)
    ws.write(f"A{SR+1}", "Customer", f_label); put(f"B{SR+1}:C{SR+1}", bs.get("customer"), f_in)
    ws.merge_range(f"D{SR+1}:E{SR+1}", "DEL PO", f_label); put(f"F{SR+1}:H{SR+1}", bs.get("del_po"), f_in_c)
    ws.write(f"I{SR+1}", "Delivery Date", f_label); put(f"J{SR+1}:K{SR+1}", bs.get("delivery_date"), f_in_date, "date")
    ws.write(f"A{SR+2}", "DEL Location", f_label); put(f"B{SR+2}:K{SR+2}", bs.get("del_location"), f_in)
    ws.set_row(SR + 2, 6)

    # ---- ALLOWANCES / COSTS / PROFIT ----
    BR = SR + 4
    ws.set_row(BR - 1, 16)
    ws.merge_range(f"A{BR}:C{BR}", "ALLOWANCES (OFF REVENUE)", f_sec)
    ws.write_blank(f"D{BR}", None, F())
    ws.merge_range(f"E{BR}:G{BR}", "COSTS", f_sec)
    ws.write_blank(f"H{BR}", None, F())
    ws.merge_range(f"I{BR}:K{BR}", "PROFIT", f_sec)
    R0 = BR + 1
    for k in range(7):
        ws.set_row(R0 - 1 + k, 17)

    def lbl2(rng, text, fmt=f_label):
        ws.merge_range(rng, text, fmt)

    r = R0
    pp_val = calc.n(bs.get("pp_discount_pct")) / 100.0 if bs.get("pp_discount_pct") not in (None, "") else None
    lbl2(f"A{r}:B{r}", "PP Discount %"); put(f"C{r}", pp_val, f_in_pct, "num"); pp = f"C{r}"
    lbl2(f"A{r+1}:B{r+1}", "Discount $"); disc = f"C{r+1}"
    ws.write_formula(disc, f"={pp}*K{TOT}", f_calc_cur, t["discount"])
    lbl2(f"A{r+2}:B{r+2}", "Program Fee ($/lb)"); put(f"C{r+2}", bs.get("program_fee"), f_in_rate, "num"); pf = f"C{r+2}"
    lbl2(f"A{r+3}:B{r+3}", "Program Fee $"); pfd = f"C{r+3}"
    ws.write_formula(pfd, f"={pf}*G{TOT}", f_calc_cur, t["program_fee_amt"])
    lbl2(f"A{r+4}:B{r+4}", "PUA ($/lb)"); put(f"C{r+4}", bs.get("pua"), f_in_rate, "num"); pua = f"C{r+4}"
    lbl2(f"A{r+5}:B{r+5}", "PUA $"); puad = f"C{r+5}"
    ws.write_formula(puad, f"={pua}*G{TOT}", f_calc_cur, t["pua_amt"])
    lbl2(f"A{r+6}:B{r+6}", "Total Allow", f_key_lbl); tallow = f"C{r+6}"
    ws.write_formula(tallow, f"={disc}+{pfd}+{puad}", f_key_cur, t["total_allow"])

    lbl2(f"E{r}:F{r}", "$ Cost (from lines)")
    lbl2(f"E{r+1}:F{r+1}", "Freight")
    lbl2(f"E{r+2}:F{r+2}", "Other Exp $")
    lbl2(f"E{r+3}:F{r+3}", "Commissions $")
    lbl2(f"E{r+4}:F{r+4}", "Total COGS", f_key_lbl)
    lbl2(f"I{r}:J{r}", "Total Rev $"); rev = f"K{r}"
    ws.write_formula(rev, f"=K{TOT}-{tallow}", f_calc_cur, t["revenue"])
    lbl2(f"I{r+1}:J{r+1}", "Gross Profit", f_key_lbl)
    lbl2(f"I{r+2}:J{r+2}", "Split % (partner)")
    lbl2(f"I{r+3}:J{r+3}", "Partner Profit")
    lbl2(f"I{r+4}:J{r+4}", "TGFC Profit", f_key_lbl)
    lbl2(f"I{r+5}:J{r+5}", "TGFC Margin", f_key_lbl)
    if full:
        cost = f"G{r}"; ws.write_formula(cost, f"=I{TOT}", f_calc_cur, t["cost_total"])
        put(f"G{r+1}", bs.get("freight"), f_in_cur, "num"); frt = f"G{r+1}"
        put(f"G{r+2}", bs.get("other_exp"), f_in_cur, "num"); oth = f"G{r+2}"
        put(f"G{r+3}", bs.get("commissions"), f_in_cur, "num"); com = f"G{r+3}"
        cogs = f"G{r+4}"; ws.write_formula(cogs, f"={cost}+{frt}+{oth}+{com}", f_key_cur, t["cogs"])
        gp = f"K{r+1}"; ws.write_formula(gp, f"={rev}-{cogs}", f_key_cur, t["gross"])
        sp_val = calc.n(bs.get("split_pct")) / 100.0 if bs.get("split_pct") not in (None, "") else None
        put(f"K{r+2}", sp_val, f_in_pct, "num"); split = f"K{r+2}"
        pp_ = f"K{r+3}"; ws.write_formula(pp_, f"={gp}*{split}", f_calc_cur, t["partner"])
        tp = f"K{r+4}"; ws.write_formula(tp, f"={gp}-{pp_}", f_key_cur, t["tgfc"])
        ws.write_formula(f"K{r+5}", f"=IF({rev}=0,0,{tp}/{rev})", f_key_pct, t["margin"])
    else:
        for k in range(4):
            ws.write(f"G{r+k}", HID, f_hidden)
        ws.write(f"G{r+4}", HID, f_key_hidden)
        ws.write(f"K{r+1}", HID, f_key_hidden)
        ws.write(f"K{r+2}", HID, f_hidden)
        ws.write(f"K{r+3}", HID, f_hidden)
        ws.write(f"K{r+4}", HID, f_key_hidden)
        ws.write(f"K{r+5}", HID, f_key_hidden)

    # ---- NOTES ----
    NR = R0 + 8
    ws.set_row(NR - 2, 6)
    ws.set_row(NR - 1, 16)
    ws.merge_range(f"A{NR}:K{NR}", "NOTES", f_sec)
    for k in range(4):
        ws.set_row(NR + k, 15)
    ws.merge_range(f"A{NR+1}:K{NR+4}", str(bs.get("notes") or ""), F(bg_color=INPUT, valign="top", text_wrap=True, **thin))

    # ---- footer ----
    FR = NR + 6
    ws.set_row(FR - 2, 8)
    for c in range(11):
        ws.write_blank(FR - 2, c, None, F(bottom=2, bottom_color=NAVY))
    ws.set_row(FR - 1, 14)
    ws.merge_range(f"A{FR}:E{FR}", "THIRD GENERATION FOOD COMPANY  ·  www.thetgfc.com",
                   F(bold=True, font_size=7.5, font_color=NAVY))
    ws.merge_range(f"F{FR}:K{FR}",
                   "Grey cells: type here.  Every other figure calculates." if full
                   else "Sale-side copy: vendor, cost and profit are not included.",
                   F(font_size=7.5, font_color=MUTED, align="right"))

    # ---- print setup ----
    ws.set_landscape(); ws.set_paper(1)
    ws.set_margins(left=0.35, right=0.35, top=0.35, bottom=0.35)
    ws.center_horizontally()
    ws.fit_to_pages(1, 1)
    ws.print_area(f"A1:K{FR}")
    ws.set_selection("B11")
    wb.set_properties({"title": f"TGFC Buy/Sell {bs.get('po') or ''}".strip(), "company": "Third Generation Food Company"})
    wb.close()
    return buf.getvalue()
