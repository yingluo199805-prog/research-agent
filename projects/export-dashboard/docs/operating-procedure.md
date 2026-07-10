# Standard operating procedure

## 1. Clean and validate source data

1. Save raw MKLS and customs files outside the repository.
2. Copy the master workbook to a dated working file; never write into the only source copy.
3. Run `refresh-master-data.ps1` without `-Writeback`.
4. Review the generated report. Stop if its status is `blocked_unmatched`.
5. Add reviewed corrections to `config/mappings/`, record the reason and effective period, then rerun validation.
6. Only after the report is clean, rerun with `-Writeback` against the working workbook.
7. Open and recalculate the resulting workbook in Excel, then retain the report beside the output.

## 2. Build and deploy the dashboard

1. Use only the approved master workbook from step 1.
2. Run `build-dashboard.ps1`. It writes ignored aggregate JSON files and copies the generated HTML to `web/public/dashboard.html`.
3. Check the generated HTML size and manually sample each of the three dashboard tabs against the workbook.
4. Confirm Vercel environment variables and project linkage, then run `deploy-dashboard.ps1`.
5. Verify the production login page, all three data tabs, the displayed data date, and the refreshed dashboard content.

## 3. Add, disable, or review a user

1. Confirm the organisation, contact name, email, and authorised request.
2. An administrator signs in to `admin.html` and creates or removes the user there; do not use raw Redis or unauthenticated curl commands in normal operations.
3. Verify the new user can log in and that the login is recorded.
4. For password reset or disablement, record the request and use the admin workflow. Do not transmit passwords in a chat, repository, or issue tracker.

Before enabling user management on a new deployment, complete every item in `SECURITY.md`.
