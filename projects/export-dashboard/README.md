# China Auto Export Dashboard

Private, Windows-first operations repository for three workflows:

1. clean and validate raw customs, CAAM, and MarkLines data;
2. build and deploy the dashboard from the validated master workbook;
3. manage dashboard users through the protected Vercel application.

Raw source data, master workbooks, generated dashboard HTML, Redis credentials,
and Vercel project linkage are intentionally excluded from Git.

## Repository map

- `src/data_cleaning/` — master-workbook update and mapping tools.
- `src/dashboard_build/` — three source aggregators plus the HTML generator.
- `web/` — Vercel API, login/admin pages, and deployment configuration.
- `scripts/` — the only supported cross-computer entry points.
- `config/mappings/` — version-controlled matching corrections.
- `archive/excel-analysis/` — retained M15 analysis utilities, not part of the dashboard release path.

See [code-map.md](docs/code-map.md) and [operating-procedure.md](docs/operating-procedure.md) before operating the project.

## New-computer setup

```powershell
git clone <private-repository-url>
Set-Location <cloned-folder>
.\scripts\setup.ps1
.\scripts\verify.ps1
```

For deployment machines, also install and authenticate the Vercel CLI, link the
local `web/` folder to the intended Vercel project, and configure the required
environment variables in Vercel. Do not put credentials in this repository.

## Supported commands

```powershell
# 1. Validate source inputs and create a report without changing the master workbook
.\scripts\refresh-master-data.ps1 -MainWorkbook <master.xlsx> -MklsRaw <mkls.xlsx> -CustomsRaw <customs.xlsx> -Report <report.json>

# 2. After resolving unmatched items, write the validated update into a workbook copy
.\scripts\refresh-master-data.ps1 -MainWorkbook <copy-of-master.xlsx> -MklsRaw <mkls.xlsx> -CustomsRaw <customs.xlsx> -Report <report.json> -Writeback

# 3. Build dashboard.html locally from the approved master workbook
.\scripts\build-dashboard.ps1 -DataWorkbook <approved-master.xlsx>

# 4. Build and deploy to the linked Vercel project
.\scripts\deploy-dashboard.ps1 -DataWorkbook <approved-master.xlsx>
```

User creation is performed only through the protected admin page after the
security requirements in [SECURITY.md](SECURITY.md) have been completed.
