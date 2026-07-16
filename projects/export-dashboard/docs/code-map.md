# Code map

| Business task | Active code | Output | Notes |
|---|---|---|---|
| 1. Clean raw data | `src/data_cleaning/refresh_master_workbook.py`, CAAM/customs/MKLS update tools, `config/mappings/` | validated master workbook and JSON report | Default mode validates only; `--writeback` is explicit. |
| 2. Build and deploy dashboard | `src/dashboard_build/gen_data*.py`, `build.py`, `scripts/build-dashboard.ps1`, `scripts/deploy-dashboard.ps1`, `scripts/build-dashboard-mac.sh`, `scripts/deploy-dashboard-mac.sh` | three aggregates, generated `dashboard.html`, Vercel production deployment | Generated JSON and HTML are ignored by Git. MKLS partial-year caps are controlled by `DASHBOARD_MKLS_MAX_YEAR` and `DASHBOARD_MKLS_MAX_MONTH`. |
| 3. Manage users | `web/api/login.js`, `users.js`, `logs.js`, `send-code.js`, `verify-code.js`, `web/public/admin.html`, `login.html`, `web/middleware.js` | Redis user records, access logs, authenticated dashboard session | Production use is gated by `SECURITY.md`. |

`archive/excel-analysis/` contains M15 workbook utilities retained for reference. They are not invoked by the dashboard operating scripts and should not be modified as part of a routine dashboard update.
