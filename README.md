# LedgerFlow ERP Platform

LedgerFlow is a browser-based ERP MVP for small businesses, offices, utility stores, and firms that need a cleaner alternative to paper records and scattered spreadsheets.

The project now has two tracks:

- Root prototype: dependency-free browser MVP for quick product testing.
- Production foundation: Next.js + NestJS + Supabase PostgreSQL/Auth/Storage scaffold.

## What Is Included

- Dashboard with revenue, expenses, inventory value, alerts, and charts
- Business health summaries and data source health
- Items and inventory records with low-stock alerts
- Company asset tracking with condition and assignment
- Employee records with department, role, salary, and status
- Owner registration, employee demo logins, custom roles, and permission previews
- Business templates and organization-defined fields without changing the core schema
- Finance records for income and expense transactions
- Customers and vendors with balances and status
- Data sources screen with CSV import into canonical modules
- Management reports with print support
- JSON export
- Local browser storage, so the demo persists after refresh
- Sample CSV files in `samples/`

## First Run: Setup Guide And Tutorial

1. **Register** the business and owner login.
2. **Setup guide** (required, saved after every step, resumes if the page is closed):
   business basics (type, country, currency, time zone, fiscal year, date format) →
   document details (phone, address, tax number) → work areas to switch on →
   starting plan (import files, enter records, or explore sample data; add staff now or later) → review.
3. **Visual tutorial** (skippable): a spotlight tour of the dashboard, sidebar, add button,
   imports, attention status, and setup. It adapts to each login's role, is remembered per
   user, and can be replayed any time from the **Tutorial** button in the top bar.
4. LedgerFlow then opens the chosen starting point.

Everything collected in the setup guide can be changed later under **Setup & Roles → Business**.
Each employee sees their own tutorial the first time they sign in.

## How To Run

Open `index.html` in a browser. No server or package install is required.

For a local preview URL:

```bash
npm start
```

Then open `http://127.0.0.1:5173`.

## Deploying to Vercel

The current customer-facing MVP is the static application in the repository root.
In Vercel, set the project Root Directory to the repository root (`.`). The checked-in
`vercel.json` disables framework auto-detection and serves the root directory without
a build step. Do not select `apps/api` as the Root Directory; that folder contains the
future NestJS API and does not serve the browser interface at `/`.

## Production Foundation

- Architecture decision: `docs/adr/0001-platform-architecture.md`
- Domain model: `docs/domain-model.md`
- IDURAR clean-room reference review: `docs/idurar-reference-review.md`
- Implementation plan: `docs/implementation-plan.md`
- Supabase migration: `supabase/migrations/0001_platform_foundation.sql`
- NestJS API scaffold: `apps/api`
- Next.js web scaffold: `apps/web`
- Shared domain types: `packages/shared`

The production dependencies are declared but not installed yet.

## MVP Notes

The root version is intentionally browser-only so the product idea can be tested quickly. Its login and permissions simulate the intended workflow but are not production security. The production track will use Supabase Auth, NestJS permission guards, PostgreSQL RLS, staged CSV/XLSX imports, and proper invoice/payment workflows.

See `PRODUCT_BLUEPRINT.md` for the researched V1 product direction.
