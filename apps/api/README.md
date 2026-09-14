# LedgerFlow API

NestJS backend for the production LedgerFlow platform.

Current scaffold:

- `HealthModule`: service health endpoint.
- `OrganizationsModule`: current organization context placeholder.
- `ImportsModule`: import preview and field mapping placeholder.
- `AnalyticsModule`: dashboard metrics placeholder.

The temporary organization context reads headers while the database and Supabase Auth integration are being wired:

```text
x-organization-id
x-user-id
x-organization-role
```

Replace this with verified Supabase JWT claims and database-backed membership lookup before any real customer data is used.
