# ITO Dashboard

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
| `scripts/test_calcs.js` | Runs the dashboard's calculations and checks totals against the report |
| `scripts/precompute.js` | Runs the dashboard's calculations and writes `public/data/summary.json` |
| `scripts/build_site.js` | Wraps `dashboard/page.html` into `public/index.html` and copies the scripts |

`dashboard/page.html` and `dashboard/reports.js` are the same files used on the claude.ai dashboard; `dashboard/dash-shim.js` provides the small runtime they need as a static site.

Deploy: copy `public/index.html`, `public/dash-shim.js`, `public/reports.js` and `public/data/summary.json` into a folder named `ito-dashboard` and run `vercel deploy --prod --scope ito-ca88` there.

## Definitions

- **Paid** = payment status `captured`. Pending, attempted, created and failed are unpaid.
- **Payment success** = paid orders ÷ order attempts.
- **Student conversion** = students with ≥1 paid order ÷ students with ≥1 order.
- No revenue figures: the exports have no amount, coupon, refund or channel columns.
