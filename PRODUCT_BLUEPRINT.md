# LedgerFlow Product Blueprint

LedgerFlow should become a hybrid ERP and company-intelligence platform for small and medium businesses that currently rely on Excel, paper, or oversized enterprise systems.

## Product Positioning

The product should not try to beat SAP, Oracle, Odoo, or ERPNext feature-for-feature on day one. The sharper V1 is:

> A simple business control center that helps owners enter, import, clean, understand, and act on company records.

This means LedgerFlow needs two sides:

- ERP modules for companies that want to manage records directly.
- Business-intelligence workflows for companies that already keep data in spreadsheets, accounting tools, POS systems, or databases.

## What We Learned From Existing Products

- Odoo shows the value of moving from dashboards into operational records, with dashboards driven by live ERP data and filters.
- ERPNext shows that ERP modules should not be isolated: purchases, stock, accounting, assets, and payments affect one another through reliable source documents and ledgers.
- Metabase shows that dashboards become powerful when users can drill from a metric into the records behind it.
- Power BI shows the importance of a broad "get data" workflow: files, databases, online services, and blank/manual entry.
- Zoho Analytics shows that data source health matters: last sync, next sync, sync failures, re-authentication, and audit history become product features.

## V1 Product Shape

LedgerFlow V1 should focus on:

- Finance: income, expenses, receivables, payables, payroll cost, net balance.
- Inventory: items, stock value, reorder levels, categories, locations.
- Assets: asset register, custodian, location, condition, maintenance status, value.
- Employees: department, role, salary, contact, status.
- Contacts: customers, vendors, balances, status.
- Data Sources: manual entry, CSV import, future accounting/database/POS connectors.
- Analytics: KPIs, charts, owner insights, alerts, printable reports.

## Core Principle

Every number on the dashboard should answer three questions:

1. What is the number?
2. Which records created it?
3. What action should the owner consider?

## Data Architecture

External data should be normalized into canonical records.

```text
Excel / CSV / API / Database / Manual Entry
                  |
                  v
          Connector or Importer
                  |
                  v
        Mapping + Validation Layer
                  |
                  v
           Canonical ERP Tables
                  |
                  v
      Metrics, Reports, Alerts, AI Analysis
```

## Canonical Tables For V1

- Company
- User
- Employee
- Contact
- InventoryItem
- Asset
- FinanceTransaction
- DataSource
- ImportRun
- ActivityEvent

Later V2 tables should include SalesInvoice, PurchaseOrder, Payment, StockMovement, Account, GLPosting, Department, Branch, Attachment, ApprovalRequest, and AuditLog.

## UX Direction

LedgerFlow should feel calmer and easier than Excel or old ERP tools:

- Dashboard first, with modules one click away.
- Summary -> chart -> table -> individual record.
- Alerts that explain what needs attention.
- Import workflows for real company data.
- Plain business language, not accounting jargon unless needed.

## Next Backend Step

The current MVP is intentionally frontend-only. The next real engineering step is a Node/Express or Next.js backend with PostgreSQL and Prisma, using a multi-company schema from day one.

Recommended next stack:

- Frontend: React + TypeScript
- Backend: Next.js API routes or Express
- Database: PostgreSQL
- ORM: Prisma
- Auth: email/password first, then organization roles
- Imports: CSV/XLSX parser with saved field mappings
- Analytics: server-calculated metrics, never AI-calculated raw accounting

## Study References

- Odoo Dashboards: https://www.odoo.com/documentation/19.0/applications/productivity/dashboards.html
- Odoo Inventory Dashboards: https://www.odoo.com/documentation/19.0/applications/inventory_and_mrp/inventory/warehouses_storage/reporting/dashboards.html
- ERPNext Accounting: https://docs.frappe.io/erpnext/accounting-introduction
- ERPNext Assets: https://docs.frappe.io/erpnext/assets/introduction
- Metabase Drill-through: https://www.metabase.com/docs/latest/questions/visualizations/drill-through
- Power BI Data Sources: https://learn.microsoft.com/en-us/power-bi/connect-data/desktop-data-sources
- Zoho Analytics Data Sources: https://www.zoho.com/analytics/help/connectors/manage-data-sources.html
