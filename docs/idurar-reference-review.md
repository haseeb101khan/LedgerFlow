# IDURAR Reference Review

Reviewed upstream commit: `e935dec13cf8b6082a7e7fc4736669e5151a80a3` (2026-08-15).

## License Boundary

IDURAR is published under AGPL-3.0. LedgerFlow may study its behavior and independently implement the same business concepts, but source files should not be copied into LedgerFlow unless we intentionally accept the AGPL obligations for the combined hosted product.

This review is therefore clean-room guidance: concepts and workflows only, with LedgerFlow code written against our own architecture.

## What IDURAR Does Well

- Treats invoices as documents with customer, dates, line items, tax, discount, status, PDF identity, and payments.
- Prevents payments above the outstanding invoice balance.
- Updates invoice payment state after payments are created, edited, or removed.
- Uses reusable list/search/read/create/edit/delete UI shells for repeated business entities.
- Keeps document-specific controllers for workflows that are more than generic CRUD.
- Provides responsive navigation, configurable money/date formatting, recent-document tables, and concise summaries.

## What We Will Adapt

- A dedicated Sales workspace with invoice list, create/read/edit views, payment recording, status filters, and overdue attention states.
- Configuration-driven tables and forms, while keeping explicit domain services for invoices, payments, stock, and accounting.
- Customer autocomplete and create-customer shortcuts during invoice entry.
- Server-calculated line totals, tax, discounts, invoice balance, and status.
- Immutable invoice-line snapshots so later catalogue changes do not rewrite old invoices.
- Soft cancellation/void workflows and audit history instead of destructive document deletion.

## What We Will Not Copy

- MongoDB/Mongoose models: LedgerFlow is relational and organization-scoped in PostgreSQL.
- The global single-business data model: every LedgerFlow record requires `organization_id` and RLS.
- Generic route generation as the authorization boundary: LedgerFlow permissions are action-specific and enforced in NestJS and PostgreSQL.
- JavaScript `Number` for financial values: LedgerFlow uses PostgreSQL `numeric` and decimal strings at API boundaries.
- A mutable global invoice-number setting: LedgerFlow uses an organization/year document sequence allocated transactionally.
- Frontend-only totals or status decisions: the backend remains authoritative.
- Directly decrementing one quantity field without history: issued sales create stock movements, and insufficient stock blocks the transaction.

## LedgerFlow Invoice Transaction

```text
Draft invoice
  -> validate customer, lines, currency, permissions
  -> calculate totals on the server
  -> issue invoice
  -> create stock-out movements when catalogue items are tracked
  -> create accounting entry/receivable
  -> record payment later
  -> update paid amount and payment status
  -> create cash/bank and receivable journal lines
  -> refresh dashboards and alerts
```

The invoice, stock, payment, and journal writes must commit or fail together. This is the main place where LedgerFlow should be stricter than the reference application.
