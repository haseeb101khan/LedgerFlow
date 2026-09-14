# LedgerFlow Implementation Plan

## Phase 1: Platform Foundation

- Create Supabase project.
- Apply `supabase/migrations/0001_platform_foundation.sql`.
- Configure Supabase Auth providers.
- Create private Supabase Storage bucket named `organization-documents`.
- Install monorepo dependencies.
- Start the NestJS API and Next.js web app.
- Replace temporary header-based organization context with real Supabase JWT verification.

## Phase 2: Organization Onboarding And Modules

- Owner creates organization.
- Owner chooses currency, timezone, fiscal year, industry, and branches.
- Owner enables Finance, Sales, HR, and Supply Chain modules or starts from a business template.
- Trigger creates the first `organization_members` owner row.
- Navigation is derived from enabled modules and effective permissions.
- Web app stores active organization selection.

## Phase 3: Roles, Permissions, And Employee Login

- Owner creates roles and grants action permissions such as `inventory.update` or `finance.approve`.
- Add organization, branch, department, and own-record scopes.
- Add field access overrides for salary, cost price, and other sensitive fields.
- Owner creates an employee record, sends an Auth invitation, then links the accepted user to the employee.
- NestJS guards resolve organization membership and permissions from the verified Supabase JWT.

## Phase 4: Custom Fields And Configuration

- Build field definitions for each supported entity.
- Store values in each record's `custom_data` JSONB.
- Include configured fields in forms, tables, imports, filters, and reports where suitable.
- Templates enable modules and install configuration; they never fork the application schema.

## Phase 5: Manual Modules

- Employees and compensation.
- Assets and maintenance records.
- Inventory items and stock movements.
- Contacts.
- Sales invoices with customer selection, line items, tax, discount, issue/due dates, lifecycle status, payment recording, and printable output.
- Issue invoices through one transaction that checks stock, creates stock movements, and posts the Finance entry.
- Allocate invoice numbers through organization/year document sequences rather than a mutable global setting.
- Finance UI backed by accounts, journal entries, and journal lines.

## Phase 6: Import Engine

- Upload CSV/XLSX to Supabase Storage.
- Create `import_runs`.
- Parse into `staging_import_rows`.
- Detect headers and suggest mappings.
- Validate required fields, numbers, dates, duplicates, departments, and currencies.
- Separate required errors from recommended/optional missing information.
- Allow fixing staged rows, skipping invalid rows, or importing valid rows only.
- Let the user confirm mappings.
- Normalize into canonical production tables.
- Preserve source lineage on every record.

## Phase 7: Executive Analytics

- Backend computes dashboard metrics from canonical tables.
- Create monthly/daily `metric_snapshots` for heavier aggregations.
- Add drill-down links from KPI to supporting records.
- Add printed and exported management reports.

## Phase 8: Reports And Alerts

- Needs-attention alerts for low stock, overdue dues, failed imports, and incomplete records.
- Permission-aware printable and exported reports.
- Notifications remain in-app first; email delivery can follow.

## Phase 9: AI Analyst

- AI reads verified analytics context, not raw unrestricted database access.
- Backend owns the organization context.
- AI tools expose safe methods such as revenue, expense, payroll, asset, inventory, and period-comparison queries.

## Phase 10: Integrations

- Accounting connector.
- POS/CRM connector.
- External database connector with read-only credentials.
- Scheduled sync worker.
- Source-of-truth settings per module.
