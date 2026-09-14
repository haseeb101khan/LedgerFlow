# Architecture Assessment

## Current State

### Frontend

- The root `index.html`, `styles.css`, and `app.js` application is the working product prototype.
- It is dependency-free, persists data in browser storage, and already demonstrates onboarding, role previews, configurable fields, imports, module records, alerts, reports, and sample businesses.
- `apps/web` is a Next.js 16 and React 19 scaffold. Its page is still a placeholder and is not yet connected to authentication, the API, or Supabase.

### Backend

- `apps/api` is a NestJS 11 modular-monolith scaffold.
- Organization, imports, analytics, and health modules exist, but organization context currently comes from untrusted request headers.
- Import mapping and analytics endpoints return prototypes/placeholders rather than database-backed behavior.

### Database, Storage, And Authentication

- Supabase PostgreSQL, Auth, and Storage are the accepted target, but no live project is connected yet.
- The foundation migration contains organization-scoped tables, RLS, audit triggers, canonical records, finance ledger tables, staging imports, custom fields, and analytics snapshots.
- The browser prototype hashes local demo passwords, but that mechanism must not be carried into production. Supabase Auth must own passwords and invitations.

## What Can Be Reused

- The prototype's workflows and information architecture are useful acceptance criteria for the production UI.
- Business templates, custom-field behavior, missing-data review, duplicate detection, role preview, alerts, and deterministic dashboard calculations should be ported incrementally.
- The PostgreSQL relational core, import staging tables, lineage columns, RLS direction, and modular-monolith decision are sound foundations.

## Corrections Required Before Customer Data

1. Replace `x-organization-id`, `x-user-id`, and role headers with verified Supabase JWT authentication and a database membership lookup.
2. Enforce action permissions in NestJS guards/services and RLS. Hiding UI controls is only a user-experience layer.
3. Make enabled modules first-class organization configuration and apply them to navigation, APIs, imports, and dashboards.
4. Complete the import pipeline so files are staged and validated before canonical tables change.
5. Replace placeholder analytics with database queries over canonical records and journal entries.
6. Build the real Next.js interface from the tested root prototype instead of maintaining two competing product designs.
7. Add automated tenant-isolation, permission, import, finance-balance, and stock-movement tests before a pilot customer.

## Final Shape

```text
Organization
  -> Enabled modules and business template
  -> Canonical entities and records
  -> System fields plus configured custom fields
  -> Members, roles, action permissions, field access, and record scope
  -> Deterministic analytics, alerts, reports, and dashboards
```

The backend stays a modular monolith. Supabase owns identity and PostgreSQL; NestJS owns trusted organization context, business rules, imports, and analytics; Next.js owns the permission-aware user experience.
