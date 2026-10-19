"""Usage metrics next to the cost: licensed M365 Copilot (Graph reports) and Cowork (admin-center export).

Inputs in <usage-dir>:
  snapshots/copilot_usage_<date>.csv, copilot_summary_<date>.json, copilot_trend_<date>.json  (pull_credits.py)
  CoworkUserDetails*.csv  admin center > Copilot > Cowork > Usage > Export (tasks, active days); newest file wins
"""
import glob, html, json, os, pathlib
import pandas as pd
from charts import bars, columns

APPS = {"copilotChatWork": "Copilot Chat (work)", "microsoftTeams": "Teams", "outlook": "Outlook",
        "m365App": "M365 app", "loop": "Loop", "edge": "Edge", "excel": "Excel", "word": "Word",
        "powerPoint": "PowerPoint", "copilotChatWeb": "Copilot Chat (web)", "oneNote": "OneNote"}


def _latest(d: pathlib.Path, pattern: str):
    files = sorted(glob.glob(str(d / pattern)), key=os.path.getmtime)
    return pathlib.Path(files[-1]) if files else None


def load(d: pathlib.Path, data: pd.DataFrame, m_now: str) -> dict:
    u = {}
    f = _latest(d, "snapshots/copilot_usage_*.csv")
    if f:
        c = pd.read_csv(f, encoding="utf-8-sig")
        refresh = pd.to_datetime(c["Report Refresh Date"]).max()
        last = pd.to_datetime(c["Last Activity Date"], errors="coerce")
        u["copilot"] = pd.DataFrame({
            "name": c["Display Name"], "upn": c["User Principal Name"].str.lower(),
            "lastActivity": last.dt.date, "activeDays28": c["Active Usage Days for All Apps"].fillna(0),
            "prompts28": c["Prompts submitted for All Apps"].fillna(0),
            "chatWorkPrompts28": c["Prompts submitted for Copilot Chat (work)"].fillna(0),
            "inactive28": last.isna() | (last < refresh - pd.Timedelta(days=28))}).sort_values("prompts28")
        u["refresh"] = refresh.date()
        u["summary"] = json.loads(f.with_name(f.name.replace("usage", "summary").replace(".csv", ".json")).read_text())[0]["adoptionByProduct"][0]
        trend = json.loads(f.with_name(f.name.replace("usage", "trend").replace(".csv", ".json")).read_text())[0]["adoptionByDate"]
        u["trend"] = pd.DataFrame(trend).sort_values("reportDate")
    f = _latest(d, "CoworkUserDetails*.csv")
    if f:
        cw = pd.read_csv(f, sep=None, engine="python", encoding="utf-8-sig")
        cw = cw.drop(columns=["CR"], errors="ignore")  # some exports carry a hand-added CR column
        cw["upn"] = cw["UserPrincipalName"].str.lower()
        now = data[data["month"] == m_now].groupby("upn").agg(CR=("CR", "first"), creditsMTD=("credits", "sum"),
                                                                sessionsMTD=("sessions", "sum"))
        allm = data.groupby("upn").agg(creditsAll=("credits", "sum"), CRany=("CR", "last"))
        x = cw.merge(now, on="upn", how="left").merge(allm, on="upn", how="left")
        x["CR"] = x["CR"].fillna(x["CRany"]).fillna("UNMAPPED")
        x["creditsPerSession"] = (x["creditsMTD"] / x["sessionsMTD"].where(x["sessionsMTD"] > 0)).round(0)
        u["cowork"] = x[["DisplayName", "upn", "CR", "TotalTasks", "UserInitiatedTasks", "ScheduledTasks", "ActiveDays",
                         "LastActivityDate", "creditsMTD", "sessionsMTD", "creditsPerSession", "creditsAll"]
                        ].sort_values("creditsMTD", ascending=False)
        u["cowork_file"] = f"{f.name} ({pd.Timestamp(os.path.getmtime(f), unit='s'):%d %b})"
    return u


def sheets(u: dict):
    out = []
    if "cowork" in u:
        out.append(("Cowork usage", u["cowork"], "tCowork"))
    if "copilot" in u:
        out.append(("Copilot usage", u["copilot"].sort_values("prompts28", ascending=False), "tCopilot"))
        out.append(("Copilot inactive", u["copilot"][u["copilot"]["inactive28"]].drop(columns="inactive28"), "tInactive"))
    return out


def html_section(u: dict, cred) -> str:
    parts = []
    if "cowork" in u:
        cw = u["cowork"]
        rows = "".join(
            f"<tr><td>{html.escape(str(r.DisplayName))}</td><td>{html.escape(str(r.CR))}</td><td class=n>{r.TotalTasks:,.0f}</td>"
            f"<td class=n>{r.ScheduledTasks:,.0f}</td><td class=n>{r.ActiveDays:,.0f}</td>"
            f"<td class=n>{cred(r.creditsMTD) if pd.notna(r.creditsMTD) else '–'}</td>"
            f"<td class=n>{cred(r.creditsPerSession) if pd.notna(r.creditsPerSession) else '–'}</td></tr>"
            for r in cw.head(20).itertuples())
        tasks = cw["TotalTasks"].sum()
        parts.append(f"""<h2>Cowork usage</h2>
<div class=kpis><div class=kpi><div class=k>Cowork users</div><div class=v>{len(cw)}</div><div class=s>in export</div></div>
<div class=kpi><div class=k>Tasks</div><div class=v>{tasks:,.0f}</div><div class=s>{cw['ScheduledTasks'].sum():,.0f} scheduled</div></div>
<div class=kpi><div class=k>Tasks per user</div><div class=v>{tasks/max(len(cw),1):.1f}</div><div class=s>avg</div></div>
<div class=kpi><div class=k>Active days per user</div><div class=v>{cw['ActiveDays'].mean():.1f}</div><div class=s>avg</div></div></div>
<table><tr><th>User</th><th>CR</th><th class=n>Tasks</th><th class=n>Scheduled</th><th class=n>Active days</th><th class=n>Credits MTD</th><th class=n>Credits / session</th></tr>{rows}</table>
<p class=muted>Tasks and active days: {html.escape(u['cowork_file'])}, for the date range picked at export (admin center has no API for these yet). Credits and sessions: this month.</p>""")
    if "copilot" in u:
        s, c = u["summary"], u["copilot"]
        app_rows = sorted(((label, s.get(f"{k}ActiveUsers", 0)) for k, label in APPS.items()), key=lambda r: -r[1])
        inactive = s['anyAppEnabledUsers'] - s['anyAppActiveUsers']
        trend = [(r.reportDate[5:], r.anyAppActiveUsers) for r in u["trend"].itertuples()]
        parts.append(f"""<h2>Microsoft 365 Copilot usage (licensed, last 28 days)</h2>
<div class=kpis><div class=kpi><div class=k>Licensed users</div><div class=v>{s['anyAppEnabledUsers']:,}</div><div class=s>enabled</div></div>
<div class=kpi><div class=k>Active users</div><div class=v>{s['anyAppActiveUsers']:,}</div><div class=s>{s['anyAppActiveUsers']/max(s['anyAppEnabledUsers'],1):.0%} of licensed</div></div>
<div class=kpi><div class=k>Prompts</div><div class=v>{s['totalPromptsSubmitted']:,}</div><div class=s>{s['averagePromptsSubmitted']} per active user</div></div>
<div class=kpi><div class=k>No use in 28 days</div><div class=v>{inactive}</div><div class=s>licence reclaim candidates</div></div></div>
<h3>Active users by app</h3>{bars(app_rows, cred, "#5b3fa8")}
<h3>Daily active users</h3>{columns(trend, "#5b3fa8")}
<p class=muted>Graph Copilot usage report, refreshed {u['refresh']:%d %b %Y}. Per-user detail and the inactive list are in the Excel workbook.</p>""")
    return "".join(parts)
