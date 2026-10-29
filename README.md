# ITO Dashboard

One dashboard with a department switcher: **Orders & payments**, **Dispatch**, **MIS** and **Support Team for Schools**.

- **Dispatch** – built from the master dispatch workbook (`MAIN_Master_Data_Dispatch_Study_Material_2026-27_CORRECTED.xlsx`): overview and pipeline, courier scorecard and delivery time, regions and study material (packed vs pending), and a follow-up list of everything not delivered plus returns (RTO) and re-routes. Uploading a newer dispatch workbook or CSV on its Upload tab switches every Dispatch tab to that file. Totals are checked against the workbook's own "Summary Dispatch" sheet.
- **MIS** and **Support Team for Schools** – upload any Excel/CSV sheet and the page finds the header row, date, status, category and number columns, and builds the report (entries by month, completed vs open, time to close, top values, number totals, column profile, Excel download). Editors can save a report for the team; the newest saved one opens for everyone. Names, emails and phone numbers are never charted or saved.
- **Date filters** in every department: from/to dates and quick ranges (last 7 days, last 30 days, latest month). Orders compares the range with the period before; Dispatch can filter by order, dispatch or delivery date; MIS and Support re-analyse only the entries in range. Each department also has a **What stands out** card of findings worked out from the numbers in view.
- Saved department reports live in the dashboard's `dept_reports` collection (summaries only).


Orders, payment success and student conversion for ITO Olympiad plans, 1 Jun – 27 Oct 2026, built from the monthly order exports.

- `ITO_Payments_Report_Jun-Oct_2026.xlsx` – full report with live formulas (Summary, Monthly, Product_Plan, Funnel, Channels_Campaigns, Data_Quality, Clean_Data). Personal data is masked.
- `public/` – static dashboard, deployed on Vercel: https://ito-dashboard-rho.vercel.app
  - Tabs: Overview (KPIs with sparklines, highlights, monthly trend with a measure switch, subject × tier heatmap), Trends (orders per day with 7-day average, day-of-week heatmap, payment status by week), Funnel & recovery (student funnel, recovered vs lost unpaid orders, orders per student), Plans & products (sortable/searchable plan table, cart types, Flexi add-ons, top cart combinations), Upload & download reports, How to read this.
  - Filters: month, subject, plan tier. Clicking a month bar or a heatmap cell filters every tab. Links like `#trends` open a tab.
- `public/data/orders.json` – cleaned order rows with no names, emails or phone numbers; `public/data/summary.json` – the dashboard's precomputed figures.

## Rebuilding

The raw CSV exports are **not** in this repo (they hold student names, parent emails and phone numbers; see `.gitignore`). Put them in the repo folder, then:

```
npm install
npm run all      # Excel report + dashboard data + static site
```

| Script | What it does |
|---|---|
| `scripts/analyze.js` | Loads the CSVs, removes test/internal accounts, normalises labels, computes metrics |
| `scripts/build.js` | Writes the Excel report |
| `scripts/export_orders.js` | Writes `public/data/orders.json` |
| `scripts/export_dispatch.js` | Reads the dispatch workbook and writes `public/data/dispatch.json` and `dispatch_material.json` (`npm run dispatch`, workbook in the repo folder) |
| `scripts/test_calcs.js` | Runs the dashboard's calculations and checks totals against the report |
| `scripts/precompute.js` | Runs the dashboard's calculations and writes `public/data/summary.json` |
| `scripts/build_site.js` | Wraps `dashboard/page.html` into `public/index.html` and copies the scripts |

`dashboard/page.html`, `dashboard/reports.js` and `dashboard/departments.js` are the same files used on the claude.ai dashboard; `dashboard/dash-shim.js` provides the small runtime they need as a static site.

Deploy: copy `public/index.html`, `public/dash-shim.js`, `public/reports.js`, `public/departments.js` and `public/data/` (`summary.json`, plus `orders.json`, `dispatch.json` and `dispatch_material.json`, which the date filters load on demand) into a folder named `ito-dashboard` and run `vercel deploy --prod --scope ito-ca88` there.

## Definitions

- **Paid** = payment status `captured`. Pending, attempted, created and failed are unpaid.
- **Payment success** = paid orders ÷ order attempts.
- **Student conversion** = students with ≥1 paid order ÷ students with ≥1 order.
- No revenue figures: the exports have no amount, coupon, refund or channel columns.
