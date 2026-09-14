# LedgerFlow Domain Model

LedgerFlow's central object is the organization. Users are members of organizations, and every operational record belongs to an organization.

## Core Tenancy

- `organizations`: company profile, base currency, fiscal year, timezone.
- `organization_modules`: the Finance, Sales, HR, and Supply Chain modules enabled for a business, plus per-module settings and labels.
- `organization_members`: user membership, role, status, invitation state.
- `organization_roles`: owner-defined roles, action permissions, and organization/branch/department/own-record scope.
- `role_field_permissions`: role overrides that make individual fields hidden, visible, or editable.
- `branches`: optional physical or operating locations.
- `departments`: finance, HR, sales, operations, inventory, or custom business units.

Navigation and APIs must require both an enabled module and a granted permission. The owner role remains an explicit full-access bypass.

## People

- `employees`: employee identity, role, department, branch, status, data lineage.
- `employee_compensation`: salary history with currency and effective dates.
- `attendance_records`: later attendance/leave support.

## Assets And Inventory

- `asset_categories`: vehicle, machine, IT equipment, furniture, tool, etc.
- `assets`: register of owned business assets, assignment, location, condition, values.
- `asset_events`: purchase, assignment, maintenance, depreciation, retirement, disposal.
- `maintenance_records`: vendor, cost, date, and maintenance status.
- `inventory_items`: shop/store/warehouse items and reorder levels.
- `stock_movements`: stock in, stock out, adjustment, transfer, or return.

## Finance

Finance must be ledger-ready from the start.

- `accounts`: chart of accounts.
- `journal_entries`: accounting source document.
- `journal_lines`: debit and credit lines.
- `budgets`: future budget comparison.
- `contacts`: customers, vendors, suppliers, and other parties.
- `document_sequences`: organization-scoped, year-aware numbering for invoices and later quotes/orders without a shared counter race.
- `invoices`: sales/purchase document header with issue/due dates, lifecycle status, discounts, tax, paid amount, and balance state.
- `invoice_items`: immutable item-name/price snapshots, optionally linked to catalogue items. Historical invoices must survive later catalogue edits.
- `payments`: customer receipts and vendor payments linked to invoices and contacts.

Creating an issued sales invoice, reserving or removing stock, recording a payment, updating invoice balance, and posting Finance entries must happen through one transactional service boundary. Users never enter the same sale twice.

The owner-facing UI can stay simple, but the backend should create proper accounting records.

## Data Platform

- `data_sources`: manual entry, CSV/Excel, accounting API, CRM/POS API, external database.
- `import_runs`: each upload/sync attempt.
- `staging_import_rows`: raw imported rows before production write.
- `mapping_templates`: saved external-to-canonical field mappings.
- `source_ownership`: which source owns each module and whether local edits are allowed.
- `custom_fields`: organization-specific field definitions. Supported records keep their values in `custom_data` JSONB while stable concepts remain real columns.

This supports the flow:

```text
Source -> Ingest -> Stage -> Validate -> Map -> Normalize -> Canonical tables
```

Manual forms and successful imports write to the same canonical tables. Staging rows are temporary review data, not a second record system.

## Intelligence

- `metric_snapshots`: cached daily/monthly metrics.
- `dashboard_preferences`: per-user dashboard settings.
- `ai_conversations`: user conversations with the future AI analyst.
- `ai_queries`: AI tool calls and verified data context.

## Security And Traceability

- `documents`: metadata for files stored in Supabase Storage.
- `audit_logs`: who changed what, when, and from which organization.

Every imported or synchronized record should preserve:

- `source_type`
- `source_id`
- `external_id`
- `last_synced_at`

That is how the product later answers, "Where did this number come from?"
