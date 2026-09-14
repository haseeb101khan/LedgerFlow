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

## How To Run

Open `index.html` in a browser. No server or package install is required.

For a local preview URL:

```bash
npm start
```

Then open `http://127.0.0.1:5173`.

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
