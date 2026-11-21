"""The Buy/Sell math, the same as the branded Excel sheet.

Line:        $ Cost = Pounds x Cost ($/lb)      $ Sale = Pounds x Sale Price ($/lb)
Allowances:  Discount $ = PP Discount % x $ Sale total
             Program Fee $ = Program Fee ($/lb) x total Pounds
             PUA $ = PUA ($/lb) x total Pounds
             Total Allow = Discount $ + Program Fee $ + PUA $
Costs:       Total COGS = $ Cost + Freight + Other Exp + Commissions
Profit:      Total Rev = $ Sale total - Total Allow
             Gross Profit = Total Rev - Total COGS
             Partner Profit = Gross Profit x Split %
             TGFC Profit = Gross Profit - Partner Profit
             TGFC Margin = TGFC Profit / Total Rev
Percent fields (PP Discount %, Split %) are stored as typed: 1 means 1%.
"""
from __future__ import annotations


def n(v) -> float:
    try:
        if v is None or v == "":
            return 0.0
        f = float(v)
        return 0.0 if f != f else f  # NaN -> 0
    except Exception:
        return 0.0


def line_has_data(ln: dict) -> bool:
    return any(
        str(ln.get(k) or "").strip() not in ("", "nan", "None", "0", "0.0")
        for k in ("item_code", "description", "cases", "pounds", "cost", "sale_price")
    )


def totals(bs: dict, lines: list[dict]) -> dict:
    lines = [ln for ln in lines if line_has_data(ln)]
    cases = sum(n(ln.get("cases")) for ln in lines)
    pounds = sum(n(ln.get("pounds")) for ln in lines)
    cost_total = sum(n(ln.get("pounds")) * n(ln.get("cost")) for ln in lines)
    sale_total = sum(n(ln.get("pounds")) * n(ln.get("sale_price")) for ln in lines)

    discount = n(bs.get("pp_discount_pct")) / 100.0 * sale_total
    program_fee = n(bs.get("program_fee")) * pounds
    pua = n(bs.get("pua")) * pounds
    total_allow = discount + program_fee + pua

    cogs = cost_total + n(bs.get("freight")) + n(bs.get("other_exp")) + n(bs.get("commissions"))
    revenue = sale_total - total_allow
    gross = revenue - cogs
    partner = gross * n(bs.get("split_pct")) / 100.0
    tgfc = gross - partner
    margin = (tgfc / revenue) if revenue else 0.0
    return {
        "cases": cases,
        "pounds": pounds,
        "cost_total": cost_total,
        "avg_cost": (cost_total / pounds) if pounds else 0.0,
        "sale_total": sale_total,
        "avg_sale": (sale_total / pounds) if pounds else 0.0,
        "discount": discount,
        "program_fee_amt": program_fee,
        "pua_amt": pua,
        "total_allow": total_allow,
        "cogs": cogs,
        "revenue": revenue,
        "gross": gross,
        "partner": partner,
        "tgfc": tgfc,
        "margin": margin,
    }
