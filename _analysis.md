# Finance-System — Analysis

Script ID `1A49a55R…`, owner modular@3starprojects.com. Runtime V8, TZ Asia/Kolkata.
Container-bound to the "3 Star Projects — Unified Finance" spreadsheet (`_getSS()` falls back to Script Property `SPREADSHEET_ID`, then creates a new sheet).
Web App: `executeAs USER_DEPLOYING`, `access ANYONE_ANONYMOUS`.
Library dependency: `Masters` v12 (id `1zWL2xvo…`), not yet called by any Finance code.

## 1. Purpose
Single finance back-office for 3 Star Projects (Indian FY Apr–Mar, GST-aware):
- Ledger of company income/expense (employee-submitted + owner entries), plus a separate personal ledger for owners.
- AR: orders, invoices, customer receipts (FIFO), returns, customer ledger.
- AP: vendor bills, bill payments (FIFO, TDS), allocations across projects, credit notes.
- Payroll: employee/project masters, salary computation, salary advances (request → approve → give → recover).
- Project costing, P&L/MIS dashboards, AR/AP ageing, fund transfers/float balances, monthly close (P&L + cash flow + balance sheet).
- UI is a single-page web app (Home.html) plus two stand-alone forms (employee expense, owner quick entry).

## 2. Files
| File | Role |
|---|---|
| Router.gs | `doGet` (?page=home/expense/owner), `include()`, `getAppUrl`, `getSheetUrl`, `getBootstrap()` (one-call master lists for forms) |
| Masters.gs | Sheet helpers (`_getSS`, `_buildSheet`, `_validateCol`, `_s`, `_n`), name normaliser (`_normName/_normMatch`), M_Category / M_Category_Personal / M_Party / M_Config / M_Account setup + getters, float accounts + balances, party find/create, dropdown appliers, group repair |
| Transactions.gs (132 KB) | Core ledger: FY tabs `Transactions_<FY>`, Personal_Transactions, Project_Costs, Fund_Transfers, Submission_Log; employee/owner/personal submit, settle, soft-delete, paged read, update/bulk update, bulk paste, project rollups, **onEdit** Category→Group |
| Invoices.gs | AR: Orders, Invoices, Payments, Returns; create invoice, record payment (FIFO + advance), customer ledger, edit/cancel/hard-delete |
| Bills.gs (86 KB) | AP: Bills, Bill_Payments, Bill_Allocations; create/split bills, FIFO payment, advance auto-consume, vendor ledger, AP ageing, project cost toggle |
| BillReturns.gs | Vendor credit notes (`Bill_Returns`), reduce bill outstanding |
| Salary.gs | M_Employee, M_Project, Salary sheet; employee/project CRUD; `saveSalaryRecord` |
| SalaryAdvance.gs | Single `Advances` sheet, request/approve/give/recover workflow, ledger, migration from legacy sheets |
| SalaryDashboard.gs | Salary KPIs/trend/breakdowns, bulk salary import, delete |
| DashboardData.gs (80 KB) | All dashboard endpoints (`dashKpis/Pnl/Monthly/ByCategory/ByProject/…`), `deepTally`, project overview, AR/AP reconciliation, fund-transfer matrix. Returns `{ok,data,error}` envelope via `_jsonSafe` |
| PersonalDashboard.gs | Personal KPIs/P&L/monthly/category/parties/paged txns per owner |
| MonthlyClose.gs | `M_MonthlyClose` + `MonthlyClose_View`; FY overview, load/save, quarterly/yearly rollups, auto-fill from txns |
| Directory.gs | Computed `Customer_Directory`/`Vendor_Directory`, party dedupe/merge, hard-delete cascades, **hourly trigger** installer |
| CategoryMigration.gs | One-off category rename map, rebuild M_Category, migrate txns |
| Formatting.gs | `_beautifySheet`, `beautifyAllSheets`, tab colours, money/date column formats |
| _Interlink_Masters.gs | Cached wrapper (5 min) over Masters library + 🔗 Interlink menu via installable onOpen |
| Home.html (466 KB) | SPA: Dashboard, Transactions, Sales & AR, Salary, Masters, Tools tabs (~46 `google.script.run` call sites) |
| ExpenseForm.html / OwnerTransaction.html | Stand-alone entry forms |
| _styles.html / _helpers.html | Shared CSS / JS partials via `include()` |

## 3. Key functions (by area)
- **Entry:** `submitEmployeeExpense`, `submitOwnerTransaction`, `submitOwnerBulk`, `submitPersonalTransaction`, `submitProjectCosts`, `recordFundTransfer`, `processBulkPaste`.
- **Settlement:** `listPendingSettlements`, `quickSettleTransaction`, `bulkSettleEmployeePending`.
- **Auto-drawing rule:** personal expense paid from company account writes personal row + paired "Owner Drawing" row (Class=Drawing, excluded from P&L, in cash); linked via `Linked_Ref`.
- **AR:** `createInvoice`, `recordPayment` (FIFO, leftover = advance, auto-applied on next invoice via `_applyAdvanceToInvoice`), `recordReturn`, `getCustomerLedger` (computed, never stored).
- **AP:** `createBill`, `recordBillPayment`, `quickPayBill`, `splitExistingBill`, `setBillAllocations`, `getVendorLedger`, `dashApAgeing`, `recordBillReturn`.
- **Payroll:** `saveSalaryRecord` (auto-recovers advances), `advRequest/advApproveRequest/advRecordGiven/advRecordRecovery`.
- **Reporting:** `dashKpis`, `dashPnl`, `getProjectOverview(V2)`, `getProjectRollup`, `getARReconciliation`, `getAPReconciliation`, `getFloatBalances`, `getOwnerFundsPosition`.
- **Safety:** `_writeRowVerified` (setValues + flush + read-back), `LockService.getScriptLock` on bills/returns writers.

## 4. Data flow
```
Browser (Home / ExpenseForm / OwnerTransaction)
   └─ google.script.run → Router.getBootstrap + feature endpoints
        Masters (lists, IDs, helpers)  ← every file depends on _getSS/_s/_n/_buildSheet
        Transactions ← Salary (employee lists), Masters
        Invoices / Bills / BillReturns / SalaryAdvance → write own sheets, some write Transactions
        DashboardData / PersonalDashboard / SalaryDashboard / MonthlyClose → READ ONLY across all sheets
        Directory → computed read, writes Customer_/Vendor_Directory
```
Global-scope sharing (no modules): constants like `TC`, `BC`, `IC`, `PMC`, `AC`, `SC` column maps are used across files; load order matters little because all are `var`/function declarations.
Config reads go through M_Config / M_Category (live), cached nowhere except the Interlink wrapper.

## 5. Sheets and key columns
| Sheet | Notes |
|---|---|
| Transactions_<FY> | 37 cols. Txn_ID, Entry_Type(Employee/Owner), Class(Expense/Income/Drawing/Transfer), Direction(IN/OUT), Txn_Date, Month(text MMM-yy), FY, Category, Group, Submitted_By, On_Behalf_Of, Project, Party_ID/Name, Amount, Taxable, GST_Rate, CGST/SGST/IGST, Payment_Mode, Source_of_Funds, Proof_Link, Settle_*, Recovery_*, Linked_Ref, Deleted, Recovered_Amount |
| Personal_Transactions | 18 cols: Ptxn_ID, dates, Owner, Class(Expense/Fund In/Fund Out/Internal), Direction, Category, Party, Amount, Source, Counter_Account, Linked_Ref, Deleted |
| Project_Costs | 19 cols: Cost_ID, Project, Entry_Type(Simple/Material/ManHours), Subcategory/Category, Amount, Amount_Ex_GST, GST_Rate, Status(Pending/Approved/Rejected), Deleted |
| Fund_Transfers | 19 cols: from/to account, amount, Linked_Ref, Deleted |
| Submission_Log | employee expense email/submission log |
| Orders / Invoices / Payments / Returns | AR. Invoices 27 cols (GST split, Outstanding, Paid, Project, discount fields). Payments 21 cols (TDS, Bank_Amount, Applied_Invoices, audit cols) |
| Bills / Bill_Payments / Bill_Allocations / Bill_Returns | AP. Bills 25 cols (Count_In_Project_Cost, Description, Group). Bill_Payments 20 cols. Allocations 9 cols |
| M_Employee / M_Project / Salary / Advances | Payroll. Salary 23 cols; Advances 18 cols (Request/Given/Recovery rows) |
| M_Category, M_Category_Personal, M_Party, M_Config, M_Account | Masters. M_Party: ID, Type, Name, GSTIN, Email, Phone, Address, Opening, Status (CUST-####/VEND-####) |
| M_MonthlyClose, MonthlyClose_View | 43-col monthly P&L/cash/BS |
| Customer_Directory, Vendor_Directory | Computed rollups |

## 6. Triggers
- **Installable onEdit** → `onEdit` (installed by `installOnEditTrigger()`); fills Group from Category in Transactions_<FY> and Bills.
- **Time-based hourly** → `buildDirectories` (installed by `installHourlyDirectoryTrigger()`, removed by `removeDirectoryTrigger()`).
- **Installable onOpen** → `IM_onOpen_` (🔗 Interlink menu) via `IM_installInterlinkMenu()`.
- No simple `onOpen`/custom menu, no form/other triggers. All setup/migration functions are run manually from the editor.

## 7. Observations / risks (not acted on)
1. `getSheetUrl` is defined twice (Router.gs:40, Transactions.gs:335) — same behaviour, but a global name collision.
2. `onEdit` is both the installable handler and a reserved simple-trigger name; it may fire twice (simple + installable) for edits.
3. Web app is `ANYONE_ANONYMOUS` + `USER_DEPLOYING`: every server function reachable from the URL runs as the owner with no auth check, including hard-delete and bulk-update endpoints.
4. `PNL_GROUPS` comment says 14 groups, array holds 10.
5. Masters library dependency (v12) is declared but unused; `IM_refreshDropdowns` is a stub; file header says "Version 1".
6. Perf: Home.html is 466 KB; `deepTally`/dashboards and `getAllTransactions` read whole FY sheets on each call with no caching; 5000-row formatting/validation ranges pre-applied per FY tab.
7. LockService is used in Bills/BillReturns but not in the Transactions/Invoices/Salary writers.
