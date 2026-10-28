# ITO Dashboard

Orders, payment success and student conversion for ITO Olympiad plans, 1 Jun – 27 Oct 2026, built from the monthly order exports.

- `ITO_Payments_Report_Jun-Oct_2026.xlsx` – full report with live formulas (Summary, Monthly, Product_Plan, Funnel, Channels_Campaigns, Data_Quality, Clean_Data). Personal data is masked.
- `public/` – static dashboard (filters for month, subject and plan tier), deployed on Vercel.
- `public/data/orders.json` – cleaned order rows with no names, emails or phone numbers.

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
| `scripts/build_site.js` | Wraps `dashboard/page.html` into `public/index.html` |

`dashboard/page.html` is the same page used on the claude.ai dashboard; `dashboard/dash-shim.js` provides the small runtime it needs as a static site.

## Definitions

- **Paid** = payment status `captured`. Pending, attempted, created and failed are unpaid.
- **Payment success** = paid orders ÷ order attempts.
- **Student conversion** = students with ≥1 paid order ÷ students with ≥1 order.
- No revenue figures: the exports have no amount, coupon, refund or channel columns.
