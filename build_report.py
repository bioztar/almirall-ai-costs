"""Build the Copilot credits cost workbook + HTML dashboard from the Usage folder.

Inputs (all in <usage-dir>):
  M365AIConsumptionByUser.<label>.csv  admin-center Consumption > Users exports, one per month
  snapshots/credits_<date>.csv          daily API pulls (pull_credits.py); win over exports
  CoworkUserDetails.xlsx                user -> CR (column CR), maintained by hand
  cr_overrides.csv (optional)           upn,CR corrections that win over the file above

    python build_report.py "<path to Teams channel>/Usage" [--eur-rate 0.86]
"""
import calendar, datetime as dt, glob, html, json, pathlib, re, sys
import pandas as pd
sys.path.insert(0, str(pathlib.Path(__file__).parent))  # sibling modules under python -I
from charts import bars
import usage_report
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

USD_PER_CREDIT = 0.01  # Copilot Credit PAYG list price
MONTH_EXPORT = re.compile(r"M365AIConsumptionByUser\.(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*(\.mtd)?\.csv$", re.I)


def load_exports(d: pathlib.Path) -> pd.DataFrame:
    out = []
    for f in sorted(d.glob("M365AIConsumptionByUser.*.csv")):
        if not MONTH_EXPORT.match(f.name):  # skip all-history / undated exports: they span several months
            continue
        x = pd.read_csv(f)
        month = pd.to_datetime(x["Last activity date"]).max().strftime("%Y-%m")
        out.append(pd.DataFrame({"month": month, "upn": x["User Principal Name"].str.lower(),
                                 "name": x["Display Name"], "credits": x["Monthly credits used"],
                                 "sessions": x["Session Count"], "source": f.name}))
    return pd.concat(out) if out else pd.DataFrame(columns=["month", "upn", "name", "credits", "sessions", "source"])


def load_snapshots(d: pathlib.Path):
    files = sorted(glob.glob(str(d / "snapshots/credits_*.csv")))
    snaps = pd.concat([pd.read_csv(f) for f in files]) if files else pd.DataFrame()
    if snaps.empty:
        return snaps, pd.DataFrame(columns=["date", "creditsMTD"])
    snaps["month"] = snaps["snapshotDate"].str[:7]
    daily = snaps.groupby("snapshotDate", as_index=False)["creditsMTD"].sum().rename(columns={"snapshotDate": "date"})
    latest = snaps[snaps["snapshotDate"] == snaps.groupby("month")["snapshotDate"].transform("max")]
    per_user = (latest.groupby(["month", "upn", "displayName", "snapshotDate"], as_index=False)["creditsMTD"].sum()
                .rename(columns={"displayName": "name", "creditsMTD": "credits"}))
    per_user["source"] = "API " + per_user.pop("snapshotDate")
    return per_user, daily


def load_cr(d: pathlib.Path) -> pd.DataFrame:
    """upn, CR, start, end. Priority: cr_overrides.csv > SharePoint list (cr_mapping.csv) > CoworkUserDetails.xlsx."""
    m = pd.read_excel(d / "CoworkUserDetails.xlsx")
    frames = [pd.DataFrame({"upn": m["UserPrincipalName"], "CR": m["CR"], "prio": 0})]
    if (d / "cr_mapping.csv").exists():
        sp = pd.read_csv(d / "cr_mapping.csv")
        frames.append(pd.DataFrame({"upn": sp["UserEmail"], "CR": sp["CR_code"], "prio": 1,
                                    "start": pd.to_datetime(sp["DateFrom"], utc=True).dt.tz_localize(None),
                                    "end": pd.to_datetime(sp["DateTo"], utc=True).dt.tz_localize(None)}))
    if (d / "cr_overrides.csv").exists():
        o = pd.read_csv(d / "cr_overrides.csv")
        frames.append(pd.DataFrame({"upn": o["upn"], "CR": o["CR"], "prio": 2}))
    cr = pd.concat(frames, ignore_index=True).reindex(columns=["upn", "CR", "prio", "start", "end"])
    cr["upn"] = cr["upn"].str.lower().str.strip()
    cr["CR"] = cr["CR"].astype("string").str.strip()
    cr = cr.dropna(subset=["upn", "CR"])
    cr = cr[cr["CR"] != ""]
    cr["start"] = pd.to_datetime(cr["start"]).fillna(pd.Timestamp.min)
    cr["end"] = pd.to_datetime(cr["end"]).fillna(pd.Timestamp.max)
    return cr


def assign_cr(data: pd.DataFrame, cr: pd.DataFrame) -> pd.Series:
    """CR valid on the 1st of each row's month, highest priority source wins."""
    x = data[["month", "upn"]].reset_index().merge(cr, on="upn", how="left")
    month_start = pd.to_datetime(x["month"] + "-01")
    x = x[x["CR"].isna() | ((x["start"] <= month_start + pd.offsets.MonthEnd(0)) & (x["end"] >= month_start))]
    best = x.sort_values("prio").groupby("index")["CR"].last()
    return best.reindex(data.index).fillna("UNMAPPED")


def build_data(d: pathlib.Path, eur_rate):
    exports = load_exports(d)
    api, daily = load_snapshots(d)
    data = pd.concat([exports[~exports["month"].isin(api["month"])] if len(api) else exports, api])
    # sessions exist only in the admin-center export; attach them to API rows of the same month
    data = (data.drop(columns="sessions").merge(exports[["month", "upn", "sessions"]], on=["month", "upn"], how="left")
            .reset_index(drop=True))
    data["CR"] = assign_cr(data, load_cr(d))
    data["usd"] = (data["credits"] * USD_PER_CREDIT).round(2)
    if eur_rate:
        data["eur"] = (data["usd"] * eur_rate).round(2)
    return data.sort_values(["month", "credits"], ascending=[True, False]), daily


# ---------------------------------------------------------------- Excel
def add_table(ws, name, money_cols=(), int_cols=()):
    ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(t)
    ws.freeze_panes = "B2"
    for i, col in enumerate(ws.iter_cols(min_row=1, max_row=1), 1):
        h = str(col[0].value)
        width = max(len(h), *(len(str(c.value or "")) for c in ws[get_column_letter(i)])) + 2
        ws.column_dimensions[get_column_letter(i)].width = min(width, 42)
        fmt = "#,##0.00" if h in money_cols or h.startswith(("USD", "EUR")) else "#,##0" if h in int_cols or h[:4].isdigit() or h == "Total credits" else None
        if fmt:
            for c in ws[get_column_letter(i)][1:]:
                c.number_format = fmt


def write_xlsx(data, daily, path, eur_rate, usage):
    cur = "eur" if eur_rate else "usd"
    by_cr = data.pivot_table(index="CR", columns="month", values="credits", aggfunc="sum", fill_value=0)
    by_cr["Total credits"] = by_cr.sum(axis=1)
    by_cr[f"{cur.upper()} total"] = data.groupby("CR")[cur].sum()
    by_cr = by_cr.sort_values("Total credits", ascending=False).reset_index()
    users = data.pivot_table(index=["name", "upn", "CR"], columns="month", values="credits", aggfunc="sum", fill_value=0)
    users["Total credits"] = users.sum(axis=1)
    users[f"{cur.upper()} total"] = data.groupby(["name", "upn", "CR"])[cur].sum()
    users = users.sort_values("Total credits", ascending=False).reset_index()
    check = data[data["CR"].isin(["UNMAPPED"]) | data["CR"].str.contains(r"\?", regex=True)]
    check = check.groupby(["name", "upn", "CR"], as_index=False)["credits"].sum().sort_values("credits", ascending=False)
    d = daily.assign(dayDelta=daily["creditsMTD"].diff()) if len(daily) else daily
    readme = pd.DataFrame({"Item": ["Built", "Price", "EUR rate", "Sources", "CR map", "Note"], "Value": [
        dt.datetime.now().strftime("%Y-%m-%d %H:%M"), f"{USD_PER_CREDIT} USD per Copilot Credit (PAYG list price)",
        eur_rate or "not set (USD only)", ", ".join(sorted(data["source"].unique())),
        "cr_overrides.csv > SharePoint CR Mapping (cr_mapping.csv, dated) > CoworkUserDetails.xlsx",
        "Named per-user data: core team only. Microsoft per-user figures are for reference; the Azure invoice is the record."]})
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for sheet, df, name in [("By CR", by_cr, "tCR"), ("By user", users, "tUsers"), ("Check CR", check, "tCheck"),
                                ("Daily MTD", d, "tDaily"), *usage_report.sheets(usage),
                                ("Data", data, "tData"), ("README", readme, "tReadme")]:
            df.to_excel(xw, sheet_name=sheet, index=False)
            if len(df):
                add_table(xw.sheets[sheet], name, money_cols=("usd", "eur"), int_cols=("credits", "creditsMTD", "dayDelta"))


# ---------------------------------------------------------------- HTML
def write_html(data, daily, policies, path, eur_rate, usage):
    cur, sym = ("eur", "€") if eur_rate else ("usd", "$")
    money = lambda v: f"{sym}{v:,.0f}"
    cred = lambda v: f"{v:,.0f}"
    months = sorted(data["month"].unique())
    m_now = months[-1]
    now = data[data["month"] == m_now]
    prev = data[data["month"] == months[-2]] if len(months) > 1 else now.iloc[0:0]
    y, mo = map(int, m_now.split("-"))
    as_of = dt.date.fromisoformat(daily["date"].max()) if len(daily) and daily["date"].max()[:7] == m_now else dt.date.today()
    days_in, day = calendar.monthrange(y, mo)[1], min(as_of.day, calendar.monthrange(y, mo)[1])
    forecast = now["credits"].sum() / day * days_in
    # one cap per tier policy; sum the limited ones (an unlimited policy has no value)
    cap = sum(p["periodLimit"]["value"] for p in policies
              if p.get("status") == "enabled" and (p.get("periodLimit") or {}).get("value")) or None

    kpis = [("Credits this month (MTD)", cred(now["credits"].sum()), m_now + f", day {day}/{days_in}"),
            ("Cost this month (MTD)", money(now[cur].sum()), f"at {USD_PER_CREDIT} USD/credit"),
            ("Month-end forecast", money(forecast * USD_PER_CREDIT * (eur_rate or 1)), f"{cred(forecast)} credits, linear"),
            ("Last month", money(prev[cur].sum()), months[-2] if len(months) > 1 else "–"),
            ("Active users", str(now["upn"].nunique()), f"{prev['upn'].nunique()} last month")]
    if cap:
        kpis.append(("Policy cap used", f"{now['credits'].sum()/cap:.0%}", f"of {cred(cap)} credits / month, all enabled policies"))
    kpi_html = "".join(f'<div class=kpi><div class=k>{k}</div><div class=v>{v}</div><div class=s>{s}</div></div>' for k, v, s in kpis)

    cr_now = now.groupby("CR")[cur].sum().sort_values(ascending=False)
    by_month = data.groupby("month")[cur].sum()
    crt = data.pivot_table(index="CR", columns="month", values=cur, aggfunc="sum", fill_value=0)
    crt = crt.loc[crt.sum(axis=1).sort_values(ascending=False).index]
    cr_rows = "".join(f"<tr><td>{html.escape(c)}</td>" + "".join(f"<td class=n>{money(v)}</td>" for v in r) +
                      f"<td class=n><b>{money(r.sum())}</b></td></tr>" for c, r in crt.iterrows())
    top = now.groupby(["name", "CR"], as_index=False)[["credits", cur]].sum().nlargest(15, "credits")
    user_rows = "".join(f"<tr><td>{html.escape(r.name)}</td><td>{html.escape(r.CR)}</td><td class=n>{cred(r.credits)}</td>"
                        f"<td class=n>{money(getattr(r, cur))}</td></tr>" for r in top.itertuples())
    unm = now[(now["CR"] == "UNMAPPED") | now["CR"].str.contains(r"\?", regex=True)]
    unm_html = (f'<div class="box warn"><b>{len(unm)} users without a confirmed CR this month</b> '
                f'({money(unm[cur].sum())}, {unm[cur].sum()/max(now[cur].sum(),1):.0%} of MTD cost): '
                + ", ".join(html.escape(f"{r.name} ({r.CR})") for r in unm.itertuples())
                + ". Fix in cr_overrides.csv or CoworkUserDetails.xlsx, then rebuild.</div>") if len(unm) else ""
    daily_now = daily[daily["date"].str[:7] == m_now] if len(daily) else daily
    daily_html = bars([(r.date[5:], r.creditsMTD) for r in daily_now.itertuples()], cred, "#1b7f4b") if len(daily_now) > 1 else \
        '<p class=muted>Daily trend starts once the API pull has run on two or more days this month.</p>'
    srcs = ", ".join(sorted(data["source"].unique()))

    page = f"""<!doctype html><html lang=en><head><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1">
<title>Copilot Credits Cost</title><style>
:root{{--ink:#1d2433;--muted:#5b6475;--line:#dde2ea;--zebra:#f5f7fa;--accent:#0b5cad;--warn:#a15c00}}
body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;font-size:16.5px;line-height:1.55;color:var(--ink);background:#fff;margin:0}}
main{{max-width:960px;margin:0 auto;padding:28px 16px 60px}}
h1{{font-size:27px;margin:2px 0}} h3{{font-size:16px;margin:18px 0 6px}} h2{{font-size:20px;margin:34px 0 10px;padding-left:10px;border-left:4px solid var(--accent)}}
.tag,.muted,.s{{color:var(--muted);font-size:14px}}
.kpis{{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:12px;margin-top:18px}}
.kpi{{border:1px solid var(--line);border-radius:8px;padding:12px 14px;box-shadow:0 1px 2px rgba(0,0,0,.05)}}
.k{{font-size:13px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em}} .v{{font-size:26px;font-weight:650}}
table{{border-collapse:collapse;width:100%;font-size:15px;margin:8px 0}} th,td{{border:1px solid var(--line);padding:6px 9px;text-align:left}}
th{{background:#eaf1fa}} tr:nth-child(even) td{{background:var(--zebra)}} td.n,th.n{{text-align:right;font-variant-numeric:tabular-nums}}
.box{{border:1px solid var(--line);border-left:5px solid var(--accent);background:#f4f8fd;padding:10px 14px;margin:14px 0;border-radius:4px}}
.warn{{border-left-color:var(--warn);background:#fff8ee}}
svg{{width:100%;height:auto}} .lbl{{font-size:13px;fill:var(--ink)}} .val{{font-size:12px;fill:var(--muted)}}
@media print{{.kpi{{box-shadow:none}} table,.box,svg{{page-break-inside:avoid}} body{{font-size:12pt}}}}
</style></head><body><main>
<div class=tag>Copilot usage-based billing · core team only (named per-user data)</div>
<h1>Copilot Credits — cost management</h1>
<div class=muted>Data as of {as_of:%d %b %Y} · built {dt.datetime.now():%d %b %Y %H:%M} · {USD_PER_CREDIT} USD per credit{f" · {eur_rate} EUR/USD" if eur_rate else ""}</div>
<div class=kpis>{kpi_html}</div>{unm_html}
<h2>Cost by CR — {m_now} (MTD)</h2>{bars(list(cr_now.items()), money)}
<h2>Cost by CR and month</h2><table><tr><th>CR</th>{"".join(f"<th class=n>{m}</th>" for m in crt.columns)}<th class=n>Total</th></tr>{cr_rows}
<tr><th>Total</th>{"".join(f"<th class=n>{money(v)}</th>" for v in crt.sum())}<th class=n>{money(crt.values.sum())}</th></tr></table>
<h2>Monthly total</h2>{bars(list(by_month.items()), money)}
<h2>Daily credits, month to date</h2>{daily_html}
{usage_report.html_section(usage, cred)}
<h2>Top users by cost — {m_now}</h2><table><tr><th>User</th><th>CR</th><th class=n>Credits</th><th class=n>Cost</th></tr>{user_rows}</table>
<p class=muted>Sources: {html.escape(srcs)}. Per-user figures from Microsoft are for reference; the Azure invoice is the record of truth. Full detail in the Excel workbook next to this file.</p>
</main></body></html>"""
    path.write_text(page, encoding="utf-8")


def main(usage_dir, eur_rate=None):
    d = pathlib.Path(usage_dir)
    data, daily = build_data(d, eur_rate)
    pol_files = sorted(glob.glob(str(d / "snapshots/policies_*.json")))
    policies = json.loads(pathlib.Path(pol_files[-1]).read_text()) if pol_files else []
    usage = usage_report.load(d, data, data["month"].max())
    write_xlsx(data, daily, d / "Copilot_Credits_Cost.xlsx", eur_rate, usage)
    write_html(data, daily, policies, d / "Copilot_Credits_Dashboard.html", eur_rate, usage)
    print(f"{len(data)} rows, months {sorted(data['month'].unique())} -> {d}")


if __name__ == "__main__":
    a = sys.argv[1:]
    main(a[0], float(a[a.index("--eur-rate") + 1]) if "--eur-rate" in a else None)
