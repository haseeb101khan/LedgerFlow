# ADR 0001: Platform Architecture

## Status

Accepted for V1 foundation.

## Decision

LedgerFlow will use:

- Next.js + TypeScript for the frontend.
- NestJS + TypeScript for the backend API.
- Supabase managed PostgreSQL for the relational database.
- Supabase Auth for identity.
- Supabase Storage for files and documents.
- Row Level Security plus backend RBAC for tenant isolation and permissions.
- A modular monolith backend instead of microservices.

## Context

LedgerFlow is not only a CRUD ERP. It is becoming a company intelligence platform where a business can enter records manually, import spreadsheets, connect external systems later, normalize all records into a canonical company model, and then view analytics, alerts, reports, and AI-assisted insights.

The product is heavily relational. Organizations own members, employees, assets, accounts, journal entries, customers, inventory, imports, files, and analytics. A single shared PostgreSQL database with `organization_id` on every business table is the right starting point.

## Rationale

Next.js gives the product a modern app shell and dashboard experience. NestJS gives the backend explicit modules, controllers, providers, validation boundaries, and room for background workers without splitting too early into microservices. Supabase keeps PostgreSQL, Auth, Storage, APIs, and RLS close together.

Supabase's current documentation describes RLS as database-level authorization rules and notes that `service_role`/secret keys bypass RLS, so those keys must remain server-side. Supabase Storage also models files through buckets and metadata while keeping actual files outside ordinary application tables. These match LedgerFlow's security and document-management needs.

## Consequences

- We design the database around `organizations`, not individual users.
- Every business row carries `organization_id`.
- The dashboard reads derived backend metrics, not manually maintained KPI numbers.
- Imports go through staging, mapping, validation, and normalization before production tables are changed.
- Finance is modeled around accounts, journal entries, and journal lines rather than a single flat transaction table.
- External integrations can be added later without changing the dashboard contract.

## References

- Supabase Row Level Security: https://supabase.com/docs/guides/database/postgres/row-level-security
- Supabase API Keys: https://supabase.com/docs/guides/getting-started/api-keys
- Supabase Storage: https://supabase.com/docs/guides/storage
- Next.js App Router: https://nextjs.org/docs/app
- NestJS Modules: https://docs.nestjs.com/v7/modules
