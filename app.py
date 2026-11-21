"""TGFC Buy/Sell — fill in an order recap, see totals and margin live, save it by PO, download the branded sheet.

Signs in with the same accounts as the Freight Desk (same database, same users table).
Buy-side lock: only admins and people with the "Buy side" switch on (Team page) see
vendor, contact, release #, pick-up location, cost, freight, other expenses,
commissions, COGS, split, profit and margin. Everyone else sees the sale side only;
those figures are never sent to their screen, their Excel or their emails.
"""
from __future__ import annotations

import urllib.parse
from datetime import date, datetime, timezone

import bcrypt
import pandas as pd
import streamlit as st

import calc
import db
import export_xlsx

st.set_page_config(page_title="TGFC Buy/Sell", page_icon="📋", layout="wide")

NAVY = "#1A1F4B"
TINT = "#EEF0F7"
GREEN = "#1F7A4D"
RED = "#B42318"

st.markdown(
    f"""
<style>
.block-container {{ padding-top: 1.2rem; max-width: 1180px; }}
.bs-top {{ background:{NAVY}; color:#fff; border-radius:10px; padding:12px 18px; display:flex; align-items:center; gap:14px; margin-bottom:14px; }}
.bs-top .mark {{ background:#fff; color:{NAVY}; font-weight:800; font-size:11px; border-radius:7px; padding:8px 7px; letter-spacing:.02em; }}
.bs-top h1 {{ font-size:18px; margin:0; color:#fff; font-weight:600; }}
.bs-top small {{ display:block; font-size:12px; opacity:.75; }}
.bs-top .who {{ margin-left:auto; font-size:12px; opacity:.9; text-align:right; }}
.sec {{ font-size:12px; font-weight:700; color:{NAVY}; letter-spacing:.08em; text-transform:uppercase;
        border-bottom:2px solid {NAVY}; padding-bottom:3px; margin:14px 0 8px; }}
.lock {{ font-size:12px; color:#5A5F78; background:#F6F6F8; border:1px dashed #C8CBD8; border-radius:8px; padding:6px 10px; }}
.card {{ background:{TINT}; border:1px solid #C8CBD8; border-radius:10px; padding:9px 12px; height:100%; }}
.card.key {{ border:2px solid {NAVY}; }}
.card .l {{ font-size:11px; text-transform:uppercase; letter-spacing:.07em; color:#5A5F78; font-weight:600; }}
.card .v {{ font-size:21px; font-weight:700; color:#22253A; line-height:1.2; }}
.card .s {{ font-size:11px; color:#8A8FA6; }}
.small {{ font-size:12px; color:#8A8FA6; }}
.pill {{ display:inline-block; font-size:11px; font-weight:700; border-radius:999px; padding:2px 9px; }}
</style>
""",
    unsafe_allow_html=True,
)

BUY_FIELDS = ["vendor", "contact", "release_no", "pu_location", "freight", "other_exp", "commissions", "split_pct"]
SALE_FIELDS = ["po", "order_date", "pu_date", "customer", "del_po", "delivery_date", "del_location",
               "pp_discount_pct", "program_fee", "pua", "notes"]
DATE_FIELDS = {"order_date", "pu_date", "delivery_date"}
NUM_FIELDS = {"pp_discount_pct", "program_fee", "pua", "freight", "other_exp", "commissions", "split_pct"}


# ----------------------------------------------------------------------------- helpers
def esc(s) -> str:
    return str(s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def money(v) -> str:
    v = calc.n(v)
    return ("-$" if v < 0 else "$") + "{:,.2f}".format(abs(v))


def rate(v) -> str:
    return "${:,.4f}".format(calc.n(v))


def num(v) -> str:
    return "{:,.0f}".format(calc.n(v))


def pct(v) -> str:
    return "{:.2f}%".format(calc.n(v) * 100)


def txt(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v != v:
        return ""
    return str(v).strip()


def opt_num(v):
    try:
        if v is None or v == "":
            return None
        f = float(v)
        return None if f != f else f
    except Exception:
        return None


def to_date(v):
    if not v:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return date.fromisoformat(str(v)[:10])
    except Exception:
        return None


def fmt_d(v) -> str:
    d = to_date(v)
    return d.strftime("%-m/%-d/%Y") if d else ""


def mailto(to: list[str], subject: str, body: str) -> str:
    return "mailto:" + urllib.parse.quote(",".join(to)) + "?" + urllib.parse.urlencode(
        {"subject": subject, "body": body}, quote_via=urllib.parse.quote
    )


def app_url() -> str:
    try:
        return st.secrets.get("APP_URL", "https://tgfc-buy-sell.streamlit.app")
    except Exception:
        return "https://tgfc-buy-sell.streamlit.app"


def card(label, value, sub="", key=False, color=None):
    style = f' style="color:{color}"' if color else ""
    return (f'<div class="card{" key" if key else ""}"><div class="l">{esc(label)}</div>'
            f'<div class="v"{style}>{esc(value)}</div><div class="s">{esc(sub)}</div></div>')


# ----------------------------------------------------------------------------- auth (same accounts as the Freight Desk)
def hash_pw(pw: str) -> str:
    return bcrypt.hashpw(pw.encode(), bcrypt.gensalt()).decode()


def check_pw(pw: str, h: str) -> bool:
    try:
        return bcrypt.checkpw(pw.encode(), h.encode())
    except Exception:
        return False


def bootstrap_admin():
    """Only when the users table is empty (a brand-new database): create the admin from secrets."""
    if db.fetchone("SELECT id FROM users LIMIT 1"):
        return
    try:
        email = st.secrets.get("ADMIN_EMAIL")
        pw = st.secrets.get("ADMIN_PASSWORD")
        name = st.secrets.get("ADMIN_NAME", "Admin")
    except Exception:
        email = pw = None
    if email and pw:
        db.run("INSERT INTO users (email, name, role, password_hash, must_change) VALUES (%s,%s,%s,%s,%s)",
               (email.lower().strip(), name, "admin", hash_pw(pw), True))


def current_user():
    u = st.session_state.get("user")
    if not u:
        return None
    fresh = db.fetchone("SELECT id,email,name,role,must_change,active FROM users WHERE id=%s", (u["id"],))
    if not fresh or not db.as_bool(fresh["active"]):
        st.session_state.pop("user", None)
        return None
    st.session_state["user"] = fresh
    return fresh


def has_buy_side(u) -> bool:
    """Checked fresh on every click, so switching someone off takes effect immediately."""
    if u["role"] == "admin":
        return True
    row = db.fetchone("SELECT buy_side FROM bs_access WHERE user_id=%s", (u["id"],))
    return bool(row) and db.as_bool(row["buy_side"])


def buy_side_ids() -> set[int]:
    ids = {r["id"] for r in db.fetchall("SELECT id FROM users WHERE role='admin'")}
    ids |= {r["user_id"] for r in db.fetchall("SELECT user_id, buy_side FROM bs_access") if db.as_bool(r["buy_side"])}
    return ids


def top_bar(sub: str, u=None, buy=False):
    who = ""
    if u:
        access = "Full access" if buy else "Sale side only"
        who = f'<span class="who">{esc(u["name"])}<br>{access}</span>'
    st.markdown(f'<div class="bs-top"><span class="mark">TGFC</span><div><h1>Buy / Sell</h1>'
                f'<small>{esc(sub)}</small></div>{who}</div>', unsafe_allow_html=True)


def login_screen():
    top_bar("Order recaps · same sign-in as the Freight Desk")
    if not db.fetchone("SELECT id FROM users LIMIT 1"):
        st.warning("No users exist yet. Point DATABASE_URL at the Freight Desk database, or add ADMIN_EMAIL, "
                   "ADMIN_PASSWORD and ADMIN_NAME to the app's secrets and reload.")
        return
    _, mid, _ = st.columns([1, 1.2, 1])
    with mid:
        with st.form("login"):
            st.subheader("Sign in")
            st.caption("Use your Freight Desk email and password.")
            email = st.text_input("Email")
            pw = st.text_input("Password", type="password")
            if st.form_submit_button("Sign in", type="primary"):
                u = db.fetchone("SELECT * FROM users WHERE email=%s", (email.lower().strip(),))
                if u and db.as_bool(u["active"]) and check_pw(pw, u["password_hash"]):
                    st.session_state["user"] = u
                    st.rerun()
                else:
                    st.error("That email and password don't match.")


def change_password_screen(u, forced=False):
    if forced:
        st.info("Welcome. Please set your own password before continuing.")
    with st.form("chpw"):
        st.subheader("Set a new password")
        st.caption("This is the same password you use for the Freight Desk.")
        p1 = st.text_input("New password", type="password")
        p2 = st.text_input("Repeat it", type="password")
        if st.form_submit_button("Save password", type="primary"):
            if len(p1) < 8:
                st.error("Use at least 8 characters.")
            elif p1 != p2:
                st.error("The two entries don't match.")
            else:
                db.run("UPDATE users SET password_hash=%s, must_change=%s WHERE id=%s", (hash_pw(p1), False, u["id"]))
                st.success("Password saved.")
                st.rerun()


# ----------------------------------------------------------------------------- loading a Buy/Sell into the editor
def blank_record() -> dict:
    return {f: None for f in SALE_FIELDS + BUY_FIELDS}


def load_lines(bs_id) -> list[dict]:
    if not bs_id:
        return []
    return db.fetchall("SELECT * FROM buy_sell_lines WHERE buy_sell_id=%s ORDER BY line_no, id", (bs_id,))


def lines_frame(lines: list[dict], buy: bool, keep_ids=True) -> pd.DataFrame:
    rows = []
    for ln in lines:
        r = {"id": float(ln["id"]) if keep_ids and ln.get("id") else float("nan"),
             "item_code": txt(ln.get("item_code")), "description": txt(ln.get("description")),
             "cases": opt_num(ln.get("cases")), "pounds": opt_num(ln.get("pounds"))}
        if buy:
            r["cost"] = opt_num(ln.get("cost"))
        r["sale_price"] = opt_num(ln.get("sale_price"))
        rows.append(r)
    for _ in range(max(3, 6 - len(rows))):  # a few empty rows ready to type in
        r = {"id": float("nan"), "item_code": "", "description": "", "cases": None, "pounds": None}
        if buy:
            r["cost"] = None
        r["sale_price"] = None
        rows.append(r)
    df = pd.DataFrame(rows)
    for c in ("id", "cases", "pounds", "cost", "sale_price"):
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce").astype("float64")
    return df


def ensure_loaded(buy: bool):
    """Put the record being edited into session state once per version (ver)."""
    ss = st.session_state
    ss.setdefault("ver", 0)
    ss.setdefault("bs_id", None)
    key = f"loaded_{ss.ver}"
    if key in ss and ss.get(f"loaded_buy_{ss.ver}") == buy:
        return
    if ss.get(f"loaded_buy_{ss.ver}") is not None and ss.get(f"loaded_buy_{ss.ver}") != buy:
        ss.ver += 1  # access changed mid-edit: reload so nothing hidden stays on screen
        key = f"loaded_{ss.ver}"
    copy_from = ss.pop("copy_from", None)
    src_id = copy_from or ss.bs_id
    rec = db.fetchone("SELECT * FROM buy_sells WHERE id=%s", (src_id,)) if src_id else None
    if src_id and not rec:
        ss.bs_id = None
    data = blank_record()
    if rec:
        for f in SALE_FIELDS + (BUY_FIELDS if buy else []):
            data[f] = rec.get(f)
    lines = load_lines(src_id) if rec else []
    if copy_from:
        data["po"] = ""
        data["del_po"] = ""
    ss[key] = data
    ss[f"lines_src_{ss.ver}"] = lines_frame(lines, buy, keep_ids=not copy_from)
    ss[f"loaded_buy_{ss.ver}"] = buy


def go(page: str, bs_id=None, copy_from=None, new=False):
    ss = st.session_state
    if new or bs_id is not None or copy_from is not None:
        ss.bs_id = None if (new or copy_from) else bs_id
        if copy_from:
            ss.copy_from = copy_from
        ss.ver = ss.get("ver", 0) + 1
    ss.page = page


# ----------------------------------------------------------------------------- saving
def edited_lines(df: pd.DataFrame) -> list[dict]:
    out = []
    for r in df.to_dict("records"):
        ln = {"id": int(r["id"]) if opt_num(r.get("id")) is not None else None,
              "item_code": txt(r.get("item_code")), "description": txt(r.get("description")),
              "cases": opt_num(r.get("cases")), "pounds": opt_num(r.get("pounds")),
              "sale_price": opt_num(r.get("sale_price"))}
        if "cost" in r:
            ln["cost"] = opt_num(r.get("cost"))
        out.append(ln)
    return out


def save(u, buy: bool, vals: dict, lines: list[dict]):
    """Returns (ok, message). Sale-side users can only ever write sale-side columns."""
    ss = st.session_state
    po = txt(vals.get("po"))
    if not po:
        return False, "Type the TGFC PO first. Every Buy/Sell is saved by its PO."
    clash = db.fetchone("SELECT id FROM buy_sells WHERE po=%s", (po,))
    if clash and clash["id"] != ss.bs_id:
        return False, f"PO {po} is already saved. Open it from Look up to change it."

    fields = SALE_FIELDS + (BUY_FIELDS if buy else [])
    clean = {}
    for f in fields:
        v = vals.get(f)
        if f in NUM_FIELDS:
            clean[f] = opt_num(v)
        elif f in DATE_FIELDS:
            clean[f] = to_date(v)
        else:
            clean[f] = txt(v)
    clean["po"] = po

    if ss.bs_id:
        sets = ", ".join(f"{f}=%s" for f in clean) + ", updated_by=%s, updated_at=%s"
        db.run(f"UPDATE buy_sells SET {sets} WHERE id=%s", tuple(clean.values()) + (u["id"], db.now(), ss.bs_id))
        bs_id = ss.bs_id
    else:
        cols = list(clean) + ["created_by", "updated_by"]
        ph = ", ".join(["%s"] * len(cols))
        bs_id = db.run(f"INSERT INTO buy_sells ({', '.join(cols)}) VALUES ({ph}) RETURNING id",
                       tuple(clean.values()) + (u["id"], u["id"]))

    keep = [ln for ln in lines if calc.line_has_data(ln)]
    if buy:
        db.run("DELETE FROM buy_sell_lines WHERE buy_sell_id=%s", (bs_id,))
        for i, ln in enumerate(keep):
            db.run("INSERT INTO buy_sell_lines (buy_sell_id, line_no, item_code, description, cases, pounds, cost, sale_price) "
                   "VALUES (%s,%s,%s,%s,%s,%s,%s,%s)",
                   (bs_id, i, ln["item_code"], ln["description"], ln["cases"], ln["pounds"], ln.get("cost"), ln["sale_price"]))
    else:
        # Sale side: never touch the cost column. Existing lines keep their cost.
        existing = {r["id"] for r in db.fetchall("SELECT id FROM buy_sell_lines WHERE buy_sell_id=%s", (bs_id,))}
        seen = set()
        for i, ln in enumerate(keep):
            if ln["id"] in existing:
                seen.add(ln["id"])
                db.run("UPDATE buy_sell_lines SET line_no=%s, item_code=%s, description=%s, cases=%s, pounds=%s, sale_price=%s "
                       "WHERE id=%s AND buy_sell_id=%s",
                       (i, ln["item_code"], ln["description"], ln["cases"], ln["pounds"], ln["sale_price"], ln["id"], bs_id))
            else:
                db.run("INSERT INTO buy_sell_lines (buy_sell_id, line_no, item_code, description, cases, pounds, sale_price) "
                       "VALUES (%s,%s,%s,%s,%s,%s,%s)",
                       (bs_id, i, ln["item_code"], ln["description"], ln["cases"], ln["pounds"], ln["sale_price"]))
        for gone in existing - seen:
            db.run("DELETE FROM buy_sell_lines WHERE id=%s AND buy_sell_id=%s", (gone, bs_id))

    ss.bs_id = bs_id
    ss.ver += 1  # reload from the database so the screen shows exactly what was saved
    return True, f"Saved PO {po}."


# ----------------------------------------------------------------------------- email
def email_text(vals: dict, lines: list[dict], t: dict, include_buy: bool) -> tuple[str, str]:
    po = txt(vals.get("po")) or "(no PO yet)"
    cust = txt(vals.get("customer"))
    subject = f"Buy/Sell PO {po}" + (f" · {cust}" if cust else "")
    out = [f"Buy/Sell recap for TGFC PO {po}", ""]
    if include_buy:
        out += [f"Vendor:        {txt(vals.get('vendor'))}",
                f"P/U:           {fmt_d(vals.get('pu_date'))}  {txt(vals.get('pu_location'))}".rstrip(),
                f"Release #:     {txt(vals.get('release_no'))}"]
    elif vals.get("pu_date"):
        out.append(f"P/U date:      {fmt_d(vals.get('pu_date'))}")
    out += [f"Customer:      {cust}",
            f"DEL PO:        {txt(vals.get('del_po'))}",
            f"Delivery:      {fmt_d(vals.get('delivery_date'))}  {txt(vals.get('del_location'))}".rstrip(), "",
            "Items:"]
    for ln in lines:
        if not calc.line_has_data(ln):
            continue
        s = f"  - {txt(ln.get('item_code'))} {txt(ln.get('description'))}: {num(ln.get('cases'))} cs, {num(ln.get('pounds'))} lb"
        if include_buy:
            s += f", cost {rate(ln.get('cost'))}"
        s += f", sell {rate(ln.get('sale_price'))}"
        out.append(s)
    out += ["",
            f"Totals:        {num(t['cases'])} cs, {num(t['pounds'])} lb",
            f"Sale $:        {money(t['sale_total'])}",
            f"Allowances:    {money(t['total_allow'])}",
            f"Total Rev $:   {money(t['revenue'])}"]
    if include_buy:
        out += [f"Total COGS:    {money(t['cogs'])}",
                f"Gross Profit:  {money(t['gross'])}",
                f"TGFC Profit:   {money(t['tgfc'])}",
                f"TGFC Margin:   {pct(t['margin'])}"]
    if txt(vals.get("notes")):
        out += ["", f"Notes: {txt(vals.get('notes'))}"]
    out += ["", f"Open it in the Buy/Sell app: {app_url()}"]
    return subject, "\n".join(out)


# ----------------------------------------------------------------------------- pages
def page_entry(u, buy: bool):
    ss = st.session_state
    ensure_loaded(buy)
    ver = ss.ver
    L = ss[f"loaded_{ver}"]
    flash = ss.pop("flash", None)
    if flash:
        st.success(flash)

    rec = db.fetchone("SELECT * FROM buy_sells WHERE id=%s", (ss.bs_id,)) if ss.bs_id else None
    h1, h2, h3 = st.columns([3, 1, 1])
    if rec:
        h1.markdown(f"### PO {esc(rec['po'])}")
        h1.markdown(f'<span class="small">Last saved by {esc(user_label(rec.get("updated_by")))} · {fmt_d(rec.get("updated_at"))}</span>',
                    unsafe_allow_html=True)
        h3.button("Start a copy", on_click=go, args=("Buy / Sell",), kwargs={"copy_from": ss.bs_id},
                  help="Same lines and terms in a new Buy/Sell with a blank PO")
    else:
        h1.markdown("### New Buy/Sell")
    h2.button("New Buy/Sell", on_click=go, args=("Buy / Sell",), kwargs={"new": True})
    if not buy:
        st.markdown('<div class="lock">🔒 Sale side only. Vendor, cost, profit and margin are hidden for your account.</div>',
                    unsafe_allow_html=True)

    v = {}
    # PURCHASE
    st.markdown('<div class="sec">Purchase</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    v["po"] = c1.text_input("TGFC PO", value=txt(L["po"]), key=f"po_{ver}")
    v["order_date"] = c2.date_input("Order date", value=to_date(L["order_date"]), format="MM/DD/YYYY", key=f"od_{ver}")
    v["pu_date"] = c3.date_input("P/U date", value=to_date(L["pu_date"]), format="MM/DD/YYYY", key=f"pud_{ver}")
    if buy:
        c1, c2, c3 = st.columns(3)
        v["vendor"] = c1.text_input("Vendor", value=txt(L["vendor"]), key=f"ven_{ver}")
        v["contact"] = c2.text_input("Contact info", value=txt(L["contact"]), key=f"con_{ver}")
        v["release_no"] = c3.text_input("Release #", value=txt(L["release_no"]), key=f"rel_{ver}")
        v["pu_location"] = st.text_input("P/U location", value=txt(L["pu_location"]), key=f"pul_{ver}")

    # LINES
    st.markdown('<div class="sec">Items</div>', unsafe_allow_html=True)
    cfg = {
        "id": None,
        "item_code": st.column_config.TextColumn("Item code", width="small"),
        "description": st.column_config.TextColumn("Item description", width="large"),
        "cases": st.column_config.NumberColumn("Cases", min_value=0, step=1, format="%,d"),
        "pounds": st.column_config.NumberColumn("Pounds", min_value=0, step=1, format="%,d"),
        "cost": st.column_config.NumberColumn("Cost ($/lb)", min_value=0, step=0.0001, format="$%.4f"),
        "sale_price": st.column_config.NumberColumn("Sale price ($/lb)", min_value=0, step=0.0001, format="$%.4f"),
    }
    edited = st.data_editor(ss[f"lines_src_{ver}"], key=f"lines_{ver}", num_rows="dynamic", hide_index=True,
                            column_config=cfg, width="stretch")
    lines = edited_lines(edited)

    # SALE
    st.markdown('<div class="sec">Sale</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    v["customer"] = c1.text_input("Customer", value=txt(L["customer"]), key=f"cus_{ver}")
    v["del_po"] = c2.text_input("DEL PO", value=txt(L["del_po"]), key=f"dpo_{ver}")
    v["delivery_date"] = c3.date_input("Delivery date", value=to_date(L["delivery_date"]), format="MM/DD/YYYY", key=f"dd_{ver}")
    v["del_location"] = st.text_input("DEL location", value=txt(L["del_location"]), key=f"dl_{ver}")

    # ALLOWANCES / COSTS / PROFIT inputs
    st.markdown('<div class="sec">Allowances, costs and split</div>', unsafe_allow_html=True)
    cols = st.columns(6 if buy else 3)
    v["pp_discount_pct"] = cols[0].number_input("PP discount %", value=opt_num(L["pp_discount_pct"]), min_value=0.0,
                                                step=0.25, format="%.2f", key=f"pp_{ver}", help="Type 1 for 1%")
    v["program_fee"] = cols[1].number_input("Program fee ($/lb)", value=opt_num(L["program_fee"]), min_value=0.0,
                                            step=0.0025, format="%.4f", key=f"pf_{ver}")
    v["pua"] = cols[2].number_input("PUA ($/lb)", value=opt_num(L["pua"]), min_value=0.0, step=0.0025,
                                    format="%.4f", key=f"pua_{ver}")
    if buy:
        v["freight"] = cols[3].number_input("Freight $", value=opt_num(L["freight"]), min_value=0.0, step=50.0,
                                            format="%.2f", key=f"frt_{ver}")
        v["other_exp"] = cols[4].number_input("Other exp $", value=opt_num(L["other_exp"]), step=10.0,
                                              format="%.2f", key=f"oth_{ver}")
        v["commissions"] = cols[5].number_input("Commissions $", value=opt_num(L["commissions"]), step=10.0,
                                                format="%.2f", key=f"com_{ver}")
        c1, _ = st.columns([1, 5])
        v["split_pct"] = c1.number_input("Split % (partner)", value=opt_num(L["split_pct"]), min_value=0.0,
                                         max_value=100.0, step=5.0, format="%.2f", key=f"spl_{ver}", help="Type 50 for 50%")

    v["notes"] = st.text_area("Notes", value=txt(L["notes"]), key=f"notes_{ver}", height=80)

    # LIVE TOTALS
    t = calc.totals(v, lines)
    st.markdown('<div class="sec">Totals</div>', unsafe_allow_html=True)
    row1 = [card("Cases", num(t["cases"])), card("Pounds", num(t["pounds"])),
            card("$ Sale", money(t["sale_total"]), f"avg {rate(t['avg_sale'])}/lb"),
            card("Total allow", money(t["total_allow"]),
                 f"disc {money(t['discount'])} · fee {money(t['program_fee_amt'])} · PUA {money(t['pua_amt'])}"),
            card("Total rev $", money(t["revenue"]), "sale less allowances", key=True)]
    for col, html in zip(st.columns(len(row1)), row1):
        col.markdown(html, unsafe_allow_html=True)
    if buy:
        gp_color = GREEN if t["gross"] >= 0 else RED
        row2 = [card("$ Cost", money(t["cost_total"]), f"avg {rate(t['avg_cost'])}/lb"),
                card("Total COGS", money(t["cogs"]), "cost + freight + other + comm."),
                card("Gross profit", money(t["gross"]), key=True, color=gp_color),
                card("Partner profit", money(t["partner"]), f"split {calc.n(v.get('split_pct')):.2f}%"),
                card("TGFC profit", money(t["tgfc"]), key=True, color=gp_color),
                card("TGFC margin", pct(t["margin"]), "of total rev", key=True, color=gp_color)]
        st.write("")
        for col, html in zip(st.columns(len(row2)), row2):
            col.markdown(html, unsafe_allow_html=True)
        no_cost = [ln for ln in lines if calc.line_has_data(ln) and calc.n(ln.get("pounds")) and ln.get("cost") is None]
        if no_cost:
            st.warning(f"{len(no_cost)} item line{'s have' if len(no_cost) > 1 else ' has'} no cost yet, so profit is overstated. "
                       "Lines added by sale-side people come in without a cost.")

    # ACTIONS
    st.markdown('<div class="sec">Save and send</div>', unsafe_allow_html=True)
    a1, a2, a3 = st.columns([1, 1.2, 2.4])
    if a1.button("💾 Save", type="primary"):
        ok, msg = save(u, buy, v, lines)
        if ok:
            ss.flash = msg
            st.rerun()
        else:
            st.error(msg)
    po_name = (txt(v.get("po")) or "draft").replace("/", "-")
    a2.download_button("⬇ Download Excel", data=export_xlsx.build(v, lines, full=buy),
                       file_name=f"Buy Sell {po_name}{'' if buy else ' (sale side)'}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       help="The branded TGFC sheet, filled in" + ("" if buy else ". Sale side only."))

    with a3:
        people = db.fetchall("SELECT id, name, email FROM users WHERE active=%s AND id<>%s ORDER BY name", (True, u["id"]))
        full_ids = buy_side_ids()
        names = {p["id"]: f"{p['name']} <{p['email']}>" for p in people}
        default = [p["id"] for p in people if p["id"] in full_ids]
        pick = st.multiselect("Email to", options=list(names), default=default, format_func=lambda i: names[i],
                              key=f"to_{ver}")
        include_buy = buy and bool(pick) and all(i in full_ids for i in pick)
        if pick:
            subject, body = email_text(v, lines, t, include_buy)
            st.link_button("✉ Email the team", mailto([p["email"] for p in people if p["id"] in pick], subject, body))
            st.caption(("Includes cost and profit: everyone picked has buy-side access." if include_buy else
                        "Sale side only" + (" (someone picked doesn't have buy-side access)." if buy else ".")) +
                       " Attach the Excel you downloaded.")

    if rec and u["role"] == "admin":
        with st.expander("Delete this Buy/Sell"):
            sure = st.checkbox(f"Yes, delete PO {rec['po']} and its lines", key=f"del_ok_{ver}")
            if st.button("Delete", disabled=not sure, key=f"del_{ver}"):
                db.run("DELETE FROM buy_sells WHERE id=%s", (rec["id"],))
                go("Look up", new=True)
                ss.flash = f"Deleted PO {rec['po']}."
                st.rerun()


_names_cache: dict = {}


def user_label(uid) -> str:
    if not uid:
        return "—"
    if uid not in _names_cache:
        r = db.fetchone("SELECT name FROM users WHERE id=%s", (uid,))
        _names_cache[uid] = r["name"] if r else "—"
    return _names_cache[uid]


def page_lookup(u, buy: bool):
    flash = st.session_state.pop("flash", None)
    if flash:
        st.success(flash)
    st.subheader("Look up a Buy/Sell")
    q = st.text_input("Search", placeholder="PO, customer, DEL PO" + (", vendor" if buy else ""), label_visibility="collapsed")
    recs = db.fetchall("SELECT * FROM buy_sells ORDER BY updated_at DESC, id DESC LIMIT 500")
    if q.strip():
        ql = q.strip().lower()
        fields = ["po", "customer", "del_po", "del_location"] + (["vendor", "pu_location"] if buy else [])
        recs = [r for r in recs if any(ql in txt(r.get(f)).lower() for f in fields)]
    if not recs:
        st.info("No Buy/Sells saved yet." if not q.strip() else "Nothing matches that search.")
        return
    all_lines = {}
    for ln in db.fetchall("SELECT * FROM buy_sell_lines"):
        all_lines.setdefault(ln["buy_sell_id"], []).append(ln)
    rows = []
    for r in recs:
        t = calc.totals(r, all_lines.get(r["id"], []))
        row = {"PO": r["po"], "Order date": fmt_d(r.get("order_date")), "Customer": txt(r.get("customer")),
               "Delivery": fmt_d(r.get("delivery_date")), "Pounds": round(t["pounds"]),
               "Sale $": round(t["sale_total"], 2), "Total rev $": round(t["revenue"], 2)}
        if buy:
            row.update({"Vendor": txt(r.get("vendor")), "Gross profit": round(t["gross"], 2),
                        "TGFC profit": round(t["tgfc"], 2), "Margin %": round(t["margin"] * 100, 2)})
        rows.append(row)
    df = pd.DataFrame(rows)
    st.dataframe(df, hide_index=True, width="stretch", column_config={
        "Pounds": st.column_config.NumberColumn(format="%,d"),
        "Sale $": st.column_config.NumberColumn(format="$%,.2f"),
        "Total rev $": st.column_config.NumberColumn(format="$%,.2f"),
        "Gross profit": st.column_config.NumberColumn(format="$%,.2f"),
        "TGFC profit": st.column_config.NumberColumn(format="$%,.2f"),
        "Margin %": st.column_config.NumberColumn(format="%.2f%%"),
    })
    ids = {f"PO {r['po']}" + (f" · {txt(r.get('customer'))}" if r.get("customer") else ""): r["id"] for r in recs}
    c1, c2 = st.columns([3, 1])
    choice = c1.selectbox("Open", list(ids), label_visibility="collapsed")
    c2.button("Open", type="primary", on_click=go, args=("Buy / Sell",), kwargs={"bs_id": ids[choice]})


def set_access(uid: int, me: int):
    on = bool(st.session_state.get(f"acc_{uid}"))
    if db.fetchone("SELECT user_id FROM bs_access WHERE user_id=%s", (uid,)):
        db.run("UPDATE bs_access SET buy_side=%s, updated_by=%s, updated_at=%s WHERE user_id=%s", (on, me, db.now(), uid))
    else:
        db.run("INSERT INTO bs_access (user_id, buy_side, updated_by, updated_at) VALUES (%s,%s,%s,%s)", (uid, on, me, db.now()))
    st.session_state.flash_team = f"Buy side {'on' if on else 'off'} for {user_label(uid)}."


def page_team(u):
    st.subheader("Team")
    st.caption("Turn on Buy side for anyone who should see vendor, cost, profit and margin. Everyone else sees the sale side only. "
               "Admins always see everything. To add someone, add them in the Freight Desk under Users: it's the same sign-in.")
    msg = st.session_state.pop("flash_team", None)
    if msg:
        st.success(msg)
    full_ids = buy_side_ids()
    for x in db.fetchall("SELECT id, name, email, role FROM users WHERE active=%s ORDER BY name", (True,)):
        with st.container(border=True):
            a, b, c, d = st.columns([2, 2.8, 1, 1.6])
            a.markdown(f"**{esc(x['name'])}**")
            b.markdown(esc(x["email"]))
            c.markdown(x["role"])
            if x["role"] == "admin":
                d.markdown('<span class="small">Buy side: always (admin)</span>', unsafe_allow_html=True)
            else:
                st.session_state[f"acc_{x['id']}"] = x["id"] in full_ids
                d.toggle("Buy side", key=f"acc_{x['id']}", on_change=set_access, args=(x["id"], u["id"]))


# ----------------------------------------------------------------------------- main
def main():
    try:
        db.conn()
        bootstrap_admin()
    except Exception as e:  # database not reachable
        st.error("Can't reach the database. Check DATABASE_URL in the app's secrets. "
                 "If it's the free Supabase database, it may be paused: resume it in Supabase.")
        st.caption(str(e))
        st.stop()

    u = current_user()
    if not u:
        login_screen()
        return
    buy = has_buy_side(u)
    top_bar("Order recaps · totals and margin as you type", u, buy)
    if db.as_bool(u["must_change"]):
        change_password_screen(u, forced=True)
        return

    pages = ["Buy / Sell", "Look up"]
    if u["role"] == "admin":
        pages.append("Team")
    pages.append("My account")
    ss = st.session_state
    if ss.get("page") not in pages:
        ss.page = "Buy / Sell"
    with st.sidebar:
        st.markdown(f"**{esc(u['name'])}**  \n<span class='small'>{esc(u['email'])}</span>", unsafe_allow_html=True)
        page = st.radio("Go to", pages, key="page", label_visibility="collapsed")
        if st.button("Sign out"):
            for k in list(ss.keys()):
                del ss[k]
            st.rerun()
        st.caption("Third Generation Food Company")

    if page == "Buy / Sell":
        page_entry(u, buy)
    elif page == "Look up":
        page_lookup(u, buy)
    elif page == "Team":
        page_team(u)
    else:
        change_password_screen(u)


main()
