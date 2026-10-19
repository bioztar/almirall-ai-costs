# copilot-credits-report

Daily cost and usage report for Microsoft 365 Copilot usage-based billing (Copilot Credits, e.g. Copilot Cowork),
built from the Microsoft Graph **beta Copilot Cost Management API** plus the Copilot usage reports.

Output (written into a folder of your choice, e.g. a synced Teams/SharePoint folder):

- `Copilot_Credits_Cost.xlsx` — cost by CR (cost centre) × month, by user, users without a CR, daily month-to-date,
  Cowork usage, licensed Copilot usage and inactive licences, raw data.
- `Copilot_Credits_Dashboard.html` — one self-contained page (no JS): KPIs, month-end forecast, cost by CR, top users,
  Cowork and Copilot usage.

## How it works

| Script | Does |
|---|---|
| `pull_credits.py` | Signs in (device code, delegated), lists spending policies → their groups → members, reads each user's `userBalances/{id}/serviceBalances`, saves `snapshots/credits_<date>.csv`. Also saves the Copilot usage report and, if `CR_SITE` is set, a SharePoint list `CR Mapping` (UserEmail, CR_code, DateFrom, DateTo) as `cr_mapping.csv`. |
| `build_report.py` | Combines snapshots with monthly admin-center exports (`M365AIConsumptionByUser.<month>.csv`), assigns CRs, writes the workbook and dashboard. |

The per-user cost endpoints are **delegated only** (no app-only permission), so the job runs as a signed-in admin on a
machine that person uses. The refresh token is cached in `~/.config/copilot-credits/`; when Entra expires it, the run
fails and the dashboard's "data as of" date stops moving — rerun once with a device code.

Permissions requested: `CopilotCostManagement.Read.All`, `CopilotCostManagement-Policy.Read.All`,
`CopilotCostManagement-Assignment.Read.All`, `CopilotCostManagement-UserData.Read.All`, `GroupMember.Read.All`,
`Sites.Read.All`, `Reports.Read.All` (via the Microsoft Graph Command Line Tools client).

The API only returns the **current period**; earlier months come from the admin-center Consumption export.
Price is the PAYG list price (0.01 USD/credit); pass `--eur-rate` to `build_report.py` for EUR.

## Windows (Task Scheduler)

```powershell
git clone https://github.com/bioztar/almirall-ai-costs; cd almirall-ai-costs
.\setup-windows.ps1 -UsageDir "<synced folder>\Usage" -CrSite "<tenant>.sharepoint.com:/sites/<site>"
```

Installs Python (winget, per user) if missing, creates a venv under `%LOCALAPPDATA%\copilot-credits`, runs once
interactively for the sign-in, then registers a daily task (runs late if the laptop was off). Log:
`%LOCALAPPDATA%\copilot-credits\run.log`.

## macOS / Linux

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
CR_SITE="<tenant>.sharepoint.com:/sites/<site>" .venv/bin/python pull_credits.py "<Usage dir>"
.venv/bin/python build_report.py "<Usage dir>"
```

The report contains named per-user data: keep the output folder restricted.
