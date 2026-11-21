# Finance-System

Google Apps Script bound to the "3 Star Projects - Unified Finance" spreadsheet, with a single-page web app. The finance back office for 3 Star Projects (Indian financial year April-March, GST-aware): ledger, receivables, payables, payroll and advances, project costing, dashboards and month-end close.

> Part of the 3SP family: Masters-System, IMS, Project-Management, Ticketing, 3SP_DocGen. See "How the systems fit together" at the bottom. A detailed code analysis is in `_analysis.md`.

## What it covers

- **Ledger:** company income / expense (employee-submitted and owner entries) in per-year tabs `Transactions_<FY>`, a separate personal ledger for owners, fund transfers and float balances. A personal expense paid from a company account writes a paired "Owner Drawing" row.
- **Receivables (AR):** orders, invoices, customer payments (applied oldest first, leftover kept as an advance), returns, computed customer ledger.
- **Payables (AP):** vendor bills (split across projects), payments (oldest first, TDS), allocations, advances, ageing, **vendor credit notes** (netted in every dashboard and ledger figure). Menu *Check bill balances* finds bills whose outstanding does not match.
- **Payroll:** employees, salary records, salary advances (request, approve, give, recover), man-hour costs per project (`Man_Hour_Costs`).
- **Project costing:** `Project_Costs` rows, rollups per project, P&L / MIS dashboards, AR / AP reconciliation, monthly close (P&L, cash flow, balance sheet).
- **Dates that are missing:** payments, bills and invoices saved without a date are highlighted red on open with a toast; they are never silently dated "today".

## Web app pages

`?page=home` (dashboard, transactions, sales and receivables, salary, masters, tools), `?page=expense` (employee expense form, including man-hours), `?page=owner` (owner quick entry).

## Connections

- **Masters** (library, pinned version in `appsscript.json`): **employees** come only from Masters (Finance has no employee sheet of its own); **customers** are mastered here and copied to Masters; vendors stay Finance's own list. Projects: every active Masters project becomes a Finance project `PRJ-001 - Name`.
- **IMS -> Finance (every 10 minutes, `IM_importImsCostsTrigger`):** material issued to a `PRJ-###` project becomes a **Material** cost: one `Project_Costs` row per project and item, growing as more is issued. It reads Material Requisitions **and hardware sent on Customer Delivery Challans, net of returns** (a renamed challan version `043A` is the same challan). Each line is priced once at the IMS average rate and frozen in the hidden `IMS_Cost_Log`; a row is only updated while Pending; approved rows never change. Vendor bills marked "Count in project cost" must not also cover stock that is issued this way.
- **Document numbers:** invoices carry the DocGen number `TSP/TI/<FY>/<NNN>`; the form warns (and saves as typed) if a number does not match.

## Files

`Router.gs` entry points and bootstrap · `Masters.gs` sheet helpers, party and category masters · `Transactions.gs` ledger, projects costs, transfers · `Invoices.gs` AR · `Bills.gs`, `BillReturns.gs` AP and credit notes · `Salary.gs`, `SalaryAdvance.gs`, `SalaryDashboard.gs` payroll · `DashboardData.gs`, `PersonalDashboard.gs`, `MonthlyClose.gs` reporting · `Directory.gs` customer / vendor directories and the hourly job · `Formatting.gs` · `_Interlink_Masters.gs` Masters + IMS connections and the Interlink menu · `Home.html`, `ExpenseForm.html`, `OwnerTransaction.html`, `_styles.html`, `_helpers.html` the UI.

## Deploy / development

- `clasp push --force`, commit, push to `main`. The GitHub Action (`.github/workflows/deploy.yml`) repeats the clasp push and needs the repo secret `CLASPRC_JSON` (if it fails with `invalid_rapt`, run `clasp login` and refresh the secret). The web app only serves the **deployed version**: after code changes publish a new version (Deploy > Manage deployments > Edit > New version).
- One-time in the editor: `IM_installImportTrigger()` creates the 10-minute import trigger (opening Finance also self-heals it).
- Rules learnt the hard way: Date objects must become text before `google.script.run` returns them (rows with dates silently vanish otherwise); use `getAttribute`, not `.dataset`, in HtmlService pages; new schemas avoid `?` in property names.
- Offline tests (Node `vm`) are kept in `_tests` next to the repos, not in git.

## How the systems fit together

```
              Masters-System  (library: projects, customers, parties, employees, sites)
             /     |      |      \
     Finance   IMS   Project-Mgmt  Ticketing         3SP_DocGen (Next.js) --web app--> IMS
        ^        ^        |                              (documents, numbers, PDFs)
        |        |--reservations / requests-----|
        |--- 10-minute import of project material costs from IMS
```
