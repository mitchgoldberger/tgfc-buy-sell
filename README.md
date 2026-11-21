# TGFC Buy/Sell

The branded Buy/Sell sheet as its own web app. Type an order on screen, watch the totals, profit and margin work out as you go, save it by TGFC PO, download the branded Excel, and open a ready-made email to the team.

Same sign-in as the Freight Desk: it uses the Freight Desk's database and its `users` table, so everyone keeps their email and password.

## The buy-side lock

Each person has a **Buy side** switch on the **Team** page (only admins see that page). Admins always have it.

| | Buy side ON | Buy side OFF (sale side only) |
|---|---|---|
| TGFC PO, order date, P/U date | ✓ | ✓ |
| Vendor, contact info, release #, P/U location | ✓ | hidden |
| Item code, description, cases, pounds, sale price | ✓ | ✓ |
| Cost ($/lb), $ Cost | ✓ | hidden |
| Customer, DEL PO, delivery date, DEL location | ✓ | ✓ |
| PP discount, program fee, PUA, total allow, total rev | ✓ | ✓ |
| Freight, other exp, commissions, total COGS | ✓ | hidden |
| Split %, gross profit, partner profit, TGFC profit, margin | ✓ | hidden |
| Excel download | full sheet with formulas | sale-side copy, buy side shows "—" |
| Email | includes cost and profit only if every recipient has buy side | sale side only |

Hidden figures never reach a sale-side person's screen, Excel or email. A sale-side person can still save sale-side changes; that never changes a cost, vendor or profit field. Lines they add come in with no cost, and the app warns buy-side people until a cost is filled in.

## Files

- `app.py` — the app
- `db.py` — database (Postgres via `DATABASE_URL`; a local SQLite file when it isn't set, for testing). Adds three tables of its own: `buy_sells`, `buy_sell_lines`, `bs_access`. Doesn't change any Freight Desk table.
- `calc.py` — the math, same as the Excel sheet. Program Fee and PUA are $/lb × total pounds.
- `export_xlsx.py` — builds the branded Excel
- `logo.png` — TGFC logo for the Excel
- `requirements.txt`, `.gitignore`
- `.streamlit/config.toml` — TGFC navy theme
- `.streamlit/secrets.toml.example` — the secrets it needs (never commit real values)

## Put it online

1. **GitHub** (stores the code): github.com → New repository → `tgfc-buy-sell`, Private → Create. Then **Add file → Upload files** and drag in everything in this folder except the `preview` folder. For `.streamlit/config.toml`, use **Add file → Create new file**, type `.streamlit/config.toml` as the name and paste its contents.
2. **Streamlit Community Cloud** (runs the app and gives it a web link): share.streamlit.io → **Create app** → repo `tgfc-buy-sell`, branch `main`, main file `app.py`, app URL `tgfc-buy-sell`.
3. **Secrets**: in **Advanced settings → Secrets**, paste the Freight Desk's `DATABASE_URL` (copy it from the Freight Desk app → Settings → Secrets) and `APP_URL = "https://tgfc-buy-sell.streamlit.app"`. Deploy.
4. **Supabase** (the database both apps share): nothing to set up. The app makes its tables the first time it runs. The database just has to be awake (not paused).
5. **First sign-in**: sign in with your Freight Desk email and password → **Team** → turn on **Buy side** for whoever should see cost and profit.
6. **Slack**: in #tgfc-orders-bs, click **+** in the bookmark bar under the channel name → **Add a bookmark** → paste the app link, name it "Buy/Sell".

Keep the app **public** on Streamlit's side, like the Freight Desk: the app has its own login.

## Running locally

```
pip install -r requirements.txt
streamlit run app.py
```
