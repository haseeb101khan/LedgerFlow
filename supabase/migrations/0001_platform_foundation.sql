create extension if not exists "pgcrypto";

create schema if not exists app;

do $$
begin
  create type public.member_role as enum (
    'owner',
    'admin',
    'finance_manager',
    'hr_manager',
    'operations_manager',
    'department_manager',
    'accountant',
    'employee',
    'viewer'
  );
exception when duplicate_object then null;
end; $$;

do $$
begin
  create type public.member_status as enum ('invited', 'active', 'suspended');
exception when duplicate_object then null;
end; $$;

do $$
begin
  create type public.employee_status as enum ('active', 'on_leave', 'inactive');
exception when duplicate_object then null;
end; $$;

do $$
begin
  create type public.asset_condition as enum ('active', 'maintenance', 'damaged', 'retired', 'sold');
exception when duplicate_object then null;
end; $$;

do $$
begin
  create type public.account_type as enum ('asset', 'liability', 'equity', 'revenue', 'expense');
exception when duplicate_object then null;
end; $$;

do $$
begin
  create type public.normal_balance as enum ('debit', 'credit');
exception when duplicate_object then null;
end; $$;

do $$
begin
  create type public.journal_status as enum ('draft', 'posted', 'void');
exception when duplicate_object then null;
end; $$;

do $$
begin
  create type public.contact_type as enum ('customer', 'vendor', 'both', 'other');
exception when duplicate_object then null;
end; $$;

do $$
begin
  create type public.data_source_type as enum ('manual', 'csv', 'excel', 'api', 'database');
exception when duplicate_object then null;
end; $$;

do $$
begin
  create type public.data_source_status as enum ('fresh', 'ready', 'needs_review', 'stale', 'failed', 'needs_auth', 'planned');
exception when duplicate_object then null;
end; $$;

do $$
begin
  create type public.import_status as enum (
    'uploaded',
    'mapping_required',
    'validating',
    'needs_review',
    'ready',
    'processing',
    'completed',
    'failed',
    'cancelled'
  );
exception when duplicate_object then null;
end; $$;

do $$
begin
  create type public.stock_movement_type as enum ('stock_in', 'stock_out', 'adjustment', 'transfer', 'return');
exception when duplicate_object then null;
end; $$;

create or replace function app.set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create table if not exists public.organizations (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  legal_name text,
  industry text,
  business_type text,
  currency_code char(3) not null default 'USD',
  timezone text not null default 'UTC',
  fiscal_year_start_month smallint not null default 1 check (fiscal_year_start_month between 1 and 12),
  setup_completed_at timestamptz,
  created_by uuid not null references auth.users(id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.organization_modules (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  module_key text not null check (module_key in ('finance', 'sales', 'hr', 'supply_chain')),
  enabled boolean not null default true,
  display_name text,
  settings jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, module_key)
);

create table if not exists public.organization_roles (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  key text not null,
  name text not null,
  description text,
  system_role public.member_role,
  permissions jsonb not null default '{}'::jsonb,
  record_scope text not null default 'organization' check (record_scope in ('organization', 'branch', 'department', 'own')),
  scope_settings jsonb not null default '{}'::jsonb,
  is_system boolean not null default false,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, key)
);

create table if not exists public.organization_members (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  user_id uuid not null references auth.users(id) on delete cascade,
  organization_role_id uuid references public.organization_roles(id) on delete set null,
  role public.member_role not null default 'viewer',
  status public.member_status not null default 'invited',
  invited_by uuid references auth.users(id) on delete set null,
  joined_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, user_id)
);

create or replace function app.is_org_member(org_id uuid)
returns boolean
language sql
stable
security definer
set search_path = public, auth
as $$
  select exists (
    select 1
    from public.organization_members om
    where om.organization_id = org_id
      and om.user_id = auth.uid()
      and om.status = 'active'
  );
$$;

create or replace function app.has_org_role(org_id uuid, allowed_roles public.member_role[])
returns boolean
language sql
stable
security definer
set search_path = public, auth
as $$
  select exists (
    select 1
    from public.organization_members om
    where om.organization_id = org_id
      and om.user_id = auth.uid()
      and om.status = 'active'
      and om.role = any (allowed_roles)
  );
$$;

create or replace function app.has_permission(org_id uuid, permission_key text)
returns boolean
language sql
stable
security definer
set search_path = public, auth
as $$
  select exists (
    select 1
    from public.organization_members om
    left join public.organization_roles r on r.id = om.organization_role_id
    where om.organization_id = org_id
      and om.user_id = auth.uid()
      and om.status = 'active'
      and (
        om.role in ('owner', 'admin')
        or coalesce(r.permissions ->> permission_key, 'false') = 'true'
      )
  );
$$;

create or replace function app.create_owner_membership()
returns trigger
language plpgsql
security definer
set search_path = public, auth
as $$
declare
  owner_role_id uuid;
begin
  insert into public.organization_roles (
    organization_id,
    key,
    name,
    description,
    system_role,
    permissions,
    is_system,
    created_by
  )
  values
    (
      new.id,
      'owner',
      'Owner',
      'Full business access.',
      'owner',
      jsonb_build_object(
        'dashboard.view', true,
        'inventory.view', true,
        'inventory.create', true,
        'inventory.update', true,
        'inventory.delete', true,
        'assets.view', true,
        'employees.view', true,
        'employees.view_salary', true,
        'sales.view', true,
        'sales.create', true,
        'sales.update', true,
        'sales.delete', true,
        'sales.approve', true,
        'finance.view', true,
        'finance.approve', true,
        'contacts.view', true,
        'imports.execute', true,
        'reports.view', true,
        'reports.export', true,
        'settings.manage_fields', true,
        'settings.manage_roles', true,
        'settings.manage_modules', true
      ),
      true,
      new.created_by
    ),
    (
      new.id,
      'staff',
      'Staff',
      'Daily operational access.',
      'employee',
      jsonb_build_object(
        'dashboard.view', true,
        'sales.view', true,
        'sales.create', true,
        'sales.update', true,
        'inventory.view', true,
        'inventory.create', true,
        'inventory.update', true,
        'contacts.view', true,
        'contacts.create', true,
        'contacts.update', true
      ),
      true,
      new.created_by
    )
  on conflict (organization_id, key) do nothing;

  insert into public.organization_modules (organization_id, module_key, enabled)
  values
    (new.id, 'finance', true),
    (new.id, 'sales', true),
    (new.id, 'hr', true),
    (new.id, 'supply_chain', true)
  on conflict (organization_id, module_key) do nothing;

  select id into owner_role_id
  from public.organization_roles
  where organization_id = new.id
    and key = 'owner';

  insert into public.organization_members (
    organization_id,
    user_id,
    organization_role_id,
    role,
    status,
    joined_at
  )
  values (
    new.id,
    new.created_by,
    owner_role_id,
    'owner',
    'active',
    now()
  )
  on conflict (organization_id, user_id)
  do update set
    organization_role_id = excluded.organization_role_id,
    role = 'owner',
    status = 'active',
    joined_at = coalesce(organization_members.joined_at, now()),
    updated_at = now();

  return new;
end;
$$;

drop trigger if exists create_owner_membership_on_organization on public.organizations;
create trigger create_owner_membership_on_organization
after insert on public.organizations
for each row
execute function app.create_owner_membership();

create table if not exists public.branches (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  name text not null,
  code text,
  address text,
  phone text,
  active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, code)
);

create table if not exists public.departments (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  branch_id uuid references public.branches(id) on delete set null,
  name text not null,
  code text,
  active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, name)
);

create table if not exists public.employees (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  auth_user_id uuid references auth.users(id) on delete set null,
  department_id uuid references public.departments(id) on delete set null,
  branch_id uuid references public.branches(id) on delete set null,
  employee_code text,
  full_name text not null,
  job_title text,
  email text,
  login_email text,
  invite_status text not null default 'draft',
  phone text,
  status public.employee_status not null default 'active',
  joined_on date,
  custom_data jsonb not null default '{}'::jsonb,
  source_type public.data_source_type not null default 'manual',
  source_id text,
  external_id text,
  last_synced_at timestamptz,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, employee_code),
  unique (organization_id, source_type, source_id, external_id)
);

create table if not exists public.employee_compensation (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  employee_id uuid not null references public.employees(id) on delete cascade,
  monthly_salary numeric(18,2) not null check (monthly_salary >= 0),
  currency_code char(3) not null,
  effective_from date not null default current_date,
  effective_to date,
  source_type public.data_source_type not null default 'manual',
  source_id text,
  external_id text,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  check (effective_to is null or effective_to >= effective_from)
);

create table if not exists public.asset_categories (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  name text not null,
  useful_life_months integer check (useful_life_months is null or useful_life_months > 0),
  active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, name)
);

create table if not exists public.assets (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  category_id uuid references public.asset_categories(id) on delete set null,
  department_id uuid references public.departments(id) on delete set null,
  branch_id uuid references public.branches(id) on delete set null,
  assigned_employee_id uuid references public.employees(id) on delete set null,
  asset_code text,
  name text not null,
  purchase_date date,
  purchase_cost numeric(18,2) not null default 0 check (purchase_cost >= 0),
  current_value numeric(18,2) not null default 0 check (current_value >= 0),
  currency_code char(3) not null,
  location text,
  condition public.asset_condition not null default 'active',
  custom_data jsonb not null default '{}'::jsonb,
  source_type public.data_source_type not null default 'manual',
  source_id text,
  external_id text,
  last_synced_at timestamptz,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, asset_code),
  unique (organization_id, source_type, source_id, external_id)
);

create table if not exists public.asset_events (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  asset_id uuid not null references public.assets(id) on delete cascade,
  event_type text not null,
  occurred_on date not null default current_date,
  amount numeric(18,2) check (amount is null or amount >= 0),
  notes text,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now()
);

create table if not exists public.maintenance_records (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  asset_id uuid not null references public.assets(id) on delete cascade,
  vendor_name text,
  service_date date not null default current_date,
  cost numeric(18,2) not null default 0 check (cost >= 0),
  currency_code char(3) not null,
  status text not null default 'completed',
  notes text,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.inventory_items (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  category text,
  sku text,
  name text not null,
  quantity_on_hand numeric(18,3) not null default 0 check (quantity_on_hand >= 0),
  reorder_level numeric(18,3) not null default 0 check (reorder_level >= 0),
  unit_cost numeric(18,2) not null default 0 check (unit_cost >= 0),
  selling_price numeric(18,2) not null default 0 check (selling_price >= 0),
  currency_code char(3) not null,
  location text,
  custom_data jsonb not null default '{}'::jsonb,
  source_type public.data_source_type not null default 'manual',
  source_id text,
  external_id text,
  last_synced_at timestamptz,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, sku),
  unique (organization_id, source_type, source_id, external_id)
);

create table if not exists public.stock_movements (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  inventory_item_id uuid not null references public.inventory_items(id) on delete cascade,
  movement_type public.stock_movement_type not null,
  quantity numeric(18,3) not null check (quantity > 0),
  unit_cost numeric(18,2) check (unit_cost is null or unit_cost >= 0),
  occurred_at timestamptz not null default now(),
  reference text,
  notes text,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now()
);

create table if not exists public.contacts (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  type public.contact_type not null default 'other',
  name text not null,
  email text,
  phone text,
  tax_id text,
  billing_address text,
  status text not null default 'active',
  opening_balance numeric(18,2) not null default 0,
  currency_code char(3) not null,
  custom_data jsonb not null default '{}'::jsonb,
  source_type public.data_source_type not null default 'manual',
  source_id text,
  external_id text,
  last_synced_at timestamptz,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, source_type, source_id, external_id)
);

create table if not exists public.accounts (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  parent_account_id uuid references public.accounts(id) on delete restrict,
  code text not null,
  name text not null,
  type public.account_type not null,
  normal_balance public.normal_balance not null,
  active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, code)
);

create table if not exists public.journal_entries (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  entry_date date not null default current_date,
  description text not null,
  reference text,
  status public.journal_status not null default 'draft',
  custom_data jsonb not null default '{}'::jsonb,
  source_type public.data_source_type not null default 'manual',
  source_id text,
  external_id text,
  last_synced_at timestamptz,
  created_by uuid references auth.users(id) on delete set null,
  posted_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, source_type, source_id, external_id)
);

create table if not exists public.journal_lines (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  journal_entry_id uuid not null references public.journal_entries(id) on delete cascade,
  account_id uuid not null references public.accounts(id) on delete restrict,
  department_id uuid references public.departments(id) on delete set null,
  contact_id uuid references public.contacts(id) on delete set null,
  debit numeric(18,2) not null default 0 check (debit >= 0),
  credit numeric(18,2) not null default 0 check (credit >= 0),
  description text,
  created_at timestamptz not null default now(),
  check (
    (debit > 0 and credit = 0)
    or (credit > 0 and debit = 0)
  )
);

create or replace function app.ensure_posted_journal_is_balanced()
returns trigger
language plpgsql
as $$
declare
  total_debits numeric(18,2);
  total_credits numeric(18,2);
begin
  if new.status = 'posted' and old.status is distinct from 'posted' then
    select
      coalesce(sum(debit), 0),
      coalesce(sum(credit), 0)
    into total_debits, total_credits
    from public.journal_lines
    where journal_entry_id = new.id
      and organization_id = new.organization_id;

    if total_debits = 0 or total_debits <> total_credits then
      raise exception 'Cannot post unbalanced journal entry %. Debits %, credits %', new.id, total_debits, total_credits;
    end if;

    new.posted_at = now();
  end if;

  return new;
end;
$$;

drop trigger if exists ensure_posted_journal_is_balanced_on_entry on public.journal_entries;
create trigger ensure_posted_journal_is_balanced_on_entry
before update of status on public.journal_entries
for each row
execute function app.ensure_posted_journal_is_balanced();

create table if not exists public.budgets (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  account_id uuid references public.accounts(id) on delete set null,
  department_id uuid references public.departments(id) on delete set null,
  period_start date not null,
  period_end date not null,
  amount numeric(18,2) not null check (amount >= 0),
  currency_code char(3) not null,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  check (period_end >= period_start)
);

create table if not exists public.document_sequences (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  document_type text not null check (document_type in ('invoice', 'quote', 'sales_order', 'purchase_order', 'payment')),
  document_year integer not null check (document_year between 2000 and 9999),
  prefix text not null default '',
  next_number bigint not null default 1 check (next_number > 0),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, document_type, document_year)
);

create table if not exists public.invoices (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  contact_id uuid references public.contacts(id) on delete set null,
  invoice_number text not null,
  invoice_type text not null check (invoice_type in ('sales', 'purchase')),
  reference text,
  issue_date date not null default current_date,
  due_date date,
  status text not null default 'draft' check (status in ('draft', 'issued', 'partially_paid', 'paid', 'overdue', 'cancelled', 'void')),
  subtotal numeric(18,2) not null default 0 check (subtotal >= 0),
  discount_total numeric(18,2) not null default 0 check (discount_total >= 0),
  tax_total numeric(18,2) not null default 0 check (tax_total >= 0),
  total numeric(18,2) not null default 0 check (total >= 0),
  amount_paid numeric(18,2) not null default 0 check (amount_paid >= 0),
  balance_due numeric(18,2) generated always as (greatest(total - amount_paid, 0)) stored,
  payment_status text not null default 'unpaid' check (payment_status in ('unpaid', 'partially_paid', 'paid', 'refunded')),
  currency_code char(3) not null,
  notes text,
  custom_data jsonb not null default '{}'::jsonb,
  source_type public.data_source_type not null default 'manual',
  source_id text,
  external_id text,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, invoice_number),
  check (due_date is null or due_date >= issue_date),
  check (discount_total <= subtotal + tax_total),
  check (amount_paid <= total)
);

create table if not exists public.invoice_items (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  invoice_id uuid not null references public.invoices(id) on delete cascade,
  inventory_item_id uuid references public.inventory_items(id) on delete set null,
  line_position integer not null default 1 check (line_position > 0),
  item_name text not null,
  description text not null,
  quantity numeric(18,3) not null default 1 check (quantity > 0),
  unit_price numeric(18,2) not null default 0 check (unit_price >= 0),
  discount_total numeric(18,2) not null default 0 check (discount_total >= 0),
  tax_total numeric(18,2) not null default 0 check (tax_total >= 0),
  line_total numeric(18,2) not null default 0 check (line_total >= 0),
  custom_data jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (invoice_id, line_position)
);

create table if not exists public.payments (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  contact_id uuid references public.contacts(id) on delete set null,
  invoice_id uuid references public.invoices(id) on delete set null,
  payment_date date not null default current_date,
  direction text not null check (direction in ('received', 'paid')),
  amount numeric(18,2) not null check (amount > 0),
  currency_code char(3) not null,
  method text,
  reference text,
  notes text,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.data_sources (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  name text not null,
  type public.data_source_type not null,
  status public.data_source_status not null default 'ready',
  source_of_truth boolean not null default false,
  configuration jsonb not null default '{}'::jsonb,
  last_sync_at timestamptz,
  next_sync_at timestamptz,
  last_error text,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.import_runs (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  data_source_id uuid references public.data_sources(id) on delete set null,
  target_entity text not null,
  original_filename text,
  storage_path text,
  status public.import_status not null default 'uploaded',
  rows_total integer not null default 0 check (rows_total >= 0),
  rows_ready integer not null default 0 check (rows_ready >= 0),
  rows_invalid integer not null default 0 check (rows_invalid >= 0),
  field_mapping jsonb not null default '{}'::jsonb,
  error_summary text,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  completed_at timestamptz
);

create table if not exists public.staging_import_rows (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  import_run_id uuid not null references public.import_runs(id) on delete cascade,
  row_number integer not null check (row_number > 0),
  raw_data jsonb not null,
  mapped_data jsonb not null default '{}'::jsonb,
  validation_errors jsonb not null default '[]'::jsonb,
  is_ready boolean not null default false,
  canonical_record_id uuid,
  created_at timestamptz not null default now(),
  unique (import_run_id, row_number)
);

create table if not exists public.mapping_templates (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  data_source_id uuid references public.data_sources(id) on delete cascade,
  target_entity text not null,
  name text not null,
  mapping jsonb not null,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, target_entity, name)
);

create table if not exists public.source_ownership (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  target_entity text not null,
  data_source_id uuid references public.data_sources(id) on delete set null,
  is_external_source_of_truth boolean not null default false,
  allows_local_edits boolean not null default true,
  notes text,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, target_entity)
);

create table if not exists public.custom_fields (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  module_key text not null check (module_key in ('finance', 'sales', 'hr', 'supply_chain')),
  target_entity text not null,
  field_key text not null,
  label text not null,
  field_type text not null default 'text' check (
    field_type in ('text', 'long_text', 'number', 'currency', 'date', 'datetime', 'boolean', 'select', 'multi_select', 'email', 'phone', 'url')
  ),
  required boolean not null default false,
  default_value jsonb,
  options jsonb not null default '[]'::jsonb,
  is_visible boolean not null default true,
  display_order integer not null default 0,
  active boolean not null default true,
  created_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, target_entity, field_key)
);

create table if not exists public.role_field_permissions (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  organization_role_id uuid not null references public.organization_roles(id) on delete cascade,
  target_entity text not null,
  field_key text not null,
  access_level text not null check (access_level in ('hidden', 'visible', 'editable')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_role_id, target_entity, field_key)
);

create table if not exists public.documents (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  entity_table text,
  entity_id uuid,
  document_type text,
  filename text not null,
  storage_bucket text not null default 'organization-documents',
  storage_path text not null,
  mime_type text,
  size_bytes bigint check (size_bytes is null or size_bytes >= 0),
  uploaded_by uuid references auth.users(id) on delete set null,
  created_at timestamptz not null default now(),
  unique (organization_id, storage_bucket, storage_path)
);

create table if not exists public.metric_snapshots (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  snapshot_date date not null,
  period text not null default 'daily',
  metrics jsonb not null,
  created_at timestamptz not null default now(),
  unique (organization_id, snapshot_date, period)
);

create table if not exists public.dashboard_preferences (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  user_id uuid not null references auth.users(id) on delete cascade,
  preferences jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (organization_id, user_id)
);

create table if not exists public.ai_conversations (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  user_id uuid not null references auth.users(id) on delete cascade,
  title text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.ai_queries (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  conversation_id uuid references public.ai_conversations(id) on delete cascade,
  user_id uuid not null references auth.users(id) on delete cascade,
  question text not null,
  verified_context jsonb not null default '{}'::jsonb,
  answer text,
  created_at timestamptz not null default now()
);

create table if not exists public.audit_logs (
  id uuid primary key default gen_random_uuid(),
  organization_id uuid not null references public.organizations(id) on delete cascade,
  actor_user_id uuid references auth.users(id) on delete set null,
  action text not null,
  entity_table text not null,
  entity_id uuid,
  old_values jsonb,
  new_values jsonb,
  request_id text,
  ip_address inet,
  created_at timestamptz not null default now()
);

create or replace function app.audit_row_change()
returns trigger
language plpgsql
security definer
set search_path = public, auth
as $$
declare
  org_id uuid;
  row_id uuid;
begin
  if tg_op = 'INSERT' then
    org_id = new.organization_id;
    row_id = new.id;
  else
    org_id = old.organization_id;
    row_id = old.id;
  end if;

  insert into public.audit_logs (
    organization_id,
    actor_user_id,
    action,
    entity_table,
    entity_id,
    old_values,
    new_values
  )
  values (
    org_id,
    auth.uid(),
    tg_op,
    tg_table_name,
    row_id,
    case when tg_op in ('UPDATE', 'DELETE') then to_jsonb(old) else null end,
    case when tg_op in ('INSERT', 'UPDATE') then to_jsonb(new) else null end
  );

  if tg_op = 'DELETE' then
    return old;
  end if;

  return new;
end;
$$;

create or replace function app.ensure_same_organization()
returns trigger
language plpgsql
as $$
declare
  argument_index integer = 0;
  parent_table text;
  fk_column text;
  fk_value uuid;
  parent_org_id uuid;
begin
  while argument_index < tg_nargs loop
    parent_table = tg_argv[argument_index];
    fk_column = tg_argv[argument_index + 1];
    fk_value = nullif(to_jsonb(new) ->> fk_column, '')::uuid;

    if fk_value is not null then
      execute format('select organization_id from public.%I where id = $1', parent_table)
      into parent_org_id
      using fk_value;

      if parent_org_id is null then
        raise exception 'Invalid reference %.%', parent_table, fk_value;
      end if;

      if parent_org_id <> new.organization_id then
        raise exception 'Cross-organization reference blocked on %.%', tg_table_name, fk_column;
      end if;
    end if;

    argument_index = argument_index + 2;
  end loop;

  return new;
end;
$$;

drop trigger if exists ensure_departments_same_organization on public.departments;
create trigger ensure_departments_same_organization
before insert or update on public.departments
for each row
execute function app.ensure_same_organization('branches', 'branch_id');

drop trigger if exists ensure_organization_members_same_organization on public.organization_members;
create trigger ensure_organization_members_same_organization
before insert or update on public.organization_members
for each row
execute function app.ensure_same_organization('organization_roles', 'organization_role_id');

drop trigger if exists ensure_role_field_permissions_same_organization on public.role_field_permissions;
create trigger ensure_role_field_permissions_same_organization
before insert or update on public.role_field_permissions
for each row
execute function app.ensure_same_organization('organization_roles', 'organization_role_id');

drop trigger if exists ensure_employees_same_organization on public.employees;
create trigger ensure_employees_same_organization
before insert or update on public.employees
for each row
execute function app.ensure_same_organization('departments', 'department_id', 'branches', 'branch_id');

drop trigger if exists ensure_employee_compensation_same_organization on public.employee_compensation;
create trigger ensure_employee_compensation_same_organization
before insert or update on public.employee_compensation
for each row
execute function app.ensure_same_organization('employees', 'employee_id');

drop trigger if exists ensure_assets_same_organization on public.assets;
create trigger ensure_assets_same_organization
before insert or update on public.assets
for each row
execute function app.ensure_same_organization(
  'asset_categories',
  'category_id',
  'departments',
  'department_id',
  'branches',
  'branch_id',
  'employees',
  'assigned_employee_id'
);

drop trigger if exists ensure_asset_events_same_organization on public.asset_events;
create trigger ensure_asset_events_same_organization
before insert or update on public.asset_events
for each row
execute function app.ensure_same_organization('assets', 'asset_id');

drop trigger if exists ensure_maintenance_records_same_organization on public.maintenance_records;
create trigger ensure_maintenance_records_same_organization
before insert or update on public.maintenance_records
for each row
execute function app.ensure_same_organization('assets', 'asset_id');

drop trigger if exists ensure_stock_movements_same_organization on public.stock_movements;
create trigger ensure_stock_movements_same_organization
before insert or update on public.stock_movements
for each row
execute function app.ensure_same_organization('inventory_items', 'inventory_item_id');

drop trigger if exists ensure_journal_lines_same_organization on public.journal_lines;
create trigger ensure_journal_lines_same_organization
before insert or update on public.journal_lines
for each row
execute function app.ensure_same_organization(
  'journal_entries',
  'journal_entry_id',
  'accounts',
  'account_id',
  'departments',
  'department_id',
  'contacts',
  'contact_id'
);

drop trigger if exists ensure_budgets_same_organization on public.budgets;
create trigger ensure_budgets_same_organization
before insert or update on public.budgets
for each row
execute function app.ensure_same_organization('accounts', 'account_id', 'departments', 'department_id');

drop trigger if exists ensure_invoices_same_organization on public.invoices;
create trigger ensure_invoices_same_organization
before insert or update on public.invoices
for each row
execute function app.ensure_same_organization('contacts', 'contact_id');

drop trigger if exists ensure_invoice_items_same_organization on public.invoice_items;
create trigger ensure_invoice_items_same_organization
before insert or update on public.invoice_items
for each row
execute function app.ensure_same_organization('invoices', 'invoice_id', 'inventory_items', 'inventory_item_id');

drop trigger if exists ensure_payments_same_organization on public.payments;
create trigger ensure_payments_same_organization
before insert or update on public.payments
for each row
execute function app.ensure_same_organization('contacts', 'contact_id', 'invoices', 'invoice_id');

drop trigger if exists ensure_import_runs_same_organization on public.import_runs;
create trigger ensure_import_runs_same_organization
before insert or update on public.import_runs
for each row
execute function app.ensure_same_organization('data_sources', 'data_source_id');

drop trigger if exists ensure_staging_import_rows_same_organization on public.staging_import_rows;
create trigger ensure_staging_import_rows_same_organization
before insert or update on public.staging_import_rows
for each row
execute function app.ensure_same_organization('import_runs', 'import_run_id');

drop trigger if exists ensure_mapping_templates_same_organization on public.mapping_templates;
create trigger ensure_mapping_templates_same_organization
before insert or update on public.mapping_templates
for each row
execute function app.ensure_same_organization('data_sources', 'data_source_id');

drop trigger if exists ensure_source_ownership_same_organization on public.source_ownership;
create trigger ensure_source_ownership_same_organization
before insert or update on public.source_ownership
for each row
execute function app.ensure_same_organization('data_sources', 'data_source_id');

drop trigger if exists ensure_ai_queries_same_organization on public.ai_queries;
create trigger ensure_ai_queries_same_organization
before insert or update on public.ai_queries
for each row
execute function app.ensure_same_organization('ai_conversations', 'conversation_id');

do $$
declare
  table_name text;
begin
  foreach table_name in array array[
    'organizations',
    'organization_modules',
    'organization_roles',
    'organization_members',
    'branches',
    'departments',
    'employees',
    'employee_compensation',
    'asset_categories',
    'assets',
    'maintenance_records',
    'inventory_items',
    'contacts',
    'accounts',
    'journal_entries',
    'budgets',
    'document_sequences',
    'invoices',
    'invoice_items',
    'payments',
    'data_sources',
    'import_runs',
    'mapping_templates',
    'source_ownership',
    'custom_fields',
    'role_field_permissions',
    'metric_snapshots',
    'dashboard_preferences',
    'ai_conversations'
  ] loop
    execute format('drop trigger if exists set_updated_at_on_%I on public.%I', table_name, table_name);
    execute format(
      'create trigger set_updated_at_on_%I before update on public.%I for each row execute function app.set_updated_at()',
      table_name,
      table_name
    );
  end loop;
end; $$;

do $$
declare
  table_name text;
begin
  foreach table_name in array array[
    'employees',
    'organization_modules',
    'organization_roles',
    'employee_compensation',
    'assets',
    'maintenance_records',
    'inventory_items',
    'stock_movements',
    'contacts',
    'accounts',
    'document_sequences',
    'journal_entries',
    'journal_lines',
    'invoices',
    'invoice_items',
    'payments',
    'data_sources',
    'import_runs',
    'staging_import_rows',
    'mapping_templates',
    'source_ownership',
    'custom_fields',
    'role_field_permissions',
    'documents'
  ] loop
    execute format('drop trigger if exists audit_%I_changes on public.%I', table_name, table_name);
    execute format(
      'create trigger audit_%I_changes after insert or update or delete on public.%I for each row execute function app.audit_row_change()',
      table_name,
      table_name
    );
  end loop;
end; $$;

create index if not exists idx_organization_members_user_id on public.organization_members(user_id);
create index if not exists idx_organization_members_org_role on public.organization_members(organization_id, role, status);
create index if not exists idx_organization_roles_org_key on public.organization_roles(organization_id, key);
create index if not exists idx_organization_modules_org_enabled on public.organization_modules(organization_id, enabled);
create index if not exists idx_branches_organization_id on public.branches(organization_id);
create index if not exists idx_departments_organization_id on public.departments(organization_id);
create index if not exists idx_employees_organization_id on public.employees(organization_id);
create index if not exists idx_employees_department_id on public.employees(department_id);
create index if not exists idx_assets_organization_id on public.assets(organization_id);
create index if not exists idx_assets_condition on public.assets(organization_id, condition);
create index if not exists idx_inventory_items_organization_id on public.inventory_items(organization_id);
create index if not exists idx_contacts_organization_id on public.contacts(organization_id);
create index if not exists idx_accounts_organization_id on public.accounts(organization_id);
create index if not exists idx_journal_entries_org_date on public.journal_entries(organization_id, entry_date);
create index if not exists idx_journal_lines_entry_id on public.journal_lines(journal_entry_id);
create index if not exists idx_invoices_org_status on public.invoices(organization_id, status);
create index if not exists idx_invoices_org_due_date on public.invoices(organization_id, due_date) where status in ('issued', 'partially_paid', 'overdue');
create index if not exists idx_invoice_items_invoice_id on public.invoice_items(invoice_id);
create index if not exists idx_payments_org_date on public.payments(organization_id, payment_date);
create index if not exists idx_data_sources_organization_id on public.data_sources(organization_id);
create index if not exists idx_import_runs_org_status on public.import_runs(organization_id, status);
create index if not exists idx_staging_import_rows_run_id on public.staging_import_rows(import_run_id);
create index if not exists idx_source_ownership_org_entity on public.source_ownership(organization_id, target_entity);
create index if not exists idx_custom_fields_org_entity on public.custom_fields(organization_id, target_entity);
create index if not exists idx_role_field_permissions_role_entity on public.role_field_permissions(organization_role_id, target_entity);
create index if not exists idx_metric_snapshots_org_date on public.metric_snapshots(organization_id, snapshot_date);
create index if not exists idx_audit_logs_org_created_at on public.audit_logs(organization_id, created_at desc);

alter table public.organizations enable row level security;
alter table public.organization_modules enable row level security;
alter table public.organization_roles enable row level security;
alter table public.organization_members enable row level security;
alter table public.branches enable row level security;
alter table public.departments enable row level security;
alter table public.employees enable row level security;
alter table public.employee_compensation enable row level security;
alter table public.asset_categories enable row level security;
alter table public.assets enable row level security;
alter table public.asset_events enable row level security;
alter table public.maintenance_records enable row level security;
alter table public.inventory_items enable row level security;
alter table public.stock_movements enable row level security;
alter table public.contacts enable row level security;
alter table public.accounts enable row level security;
alter table public.journal_entries enable row level security;
alter table public.journal_lines enable row level security;
alter table public.budgets enable row level security;
alter table public.document_sequences enable row level security;
alter table public.invoices enable row level security;
alter table public.invoice_items enable row level security;
alter table public.payments enable row level security;
alter table public.data_sources enable row level security;
alter table public.import_runs enable row level security;
alter table public.staging_import_rows enable row level security;
alter table public.mapping_templates enable row level security;
alter table public.source_ownership enable row level security;
alter table public.custom_fields enable row level security;
alter table public.role_field_permissions enable row level security;
alter table public.documents enable row level security;
alter table public.metric_snapshots enable row level security;
alter table public.dashboard_preferences enable row level security;
alter table public.ai_conversations enable row level security;
alter table public.ai_queries enable row level security;
alter table public.audit_logs enable row level security;

create policy "Members can read organizations"
on public.organizations for select
to authenticated
using (app.is_org_member(id));

create policy "Authenticated users can create organizations"
on public.organizations for insert
to authenticated
with check (created_by = auth.uid());

create policy "Owners and admins can update organizations"
on public.organizations for update
to authenticated
using (app.has_org_role(id, array['owner', 'admin']::public.member_role[]))
with check (app.has_org_role(id, array['owner', 'admin']::public.member_role[]));

create policy "Members can read organization memberships"
on public.organization_members for select
to authenticated
using (app.is_org_member(organization_id));

create policy "Owners and admins can invite members"
on public.organization_members for insert
to authenticated
with check (app.has_org_role(organization_id, array['owner', 'admin']::public.member_role[]));

create policy "Owners and admins can update memberships"
on public.organization_members for update
to authenticated
using (app.has_org_role(organization_id, array['owner', 'admin']::public.member_role[]))
with check (app.has_org_role(organization_id, array['owner', 'admin']::public.member_role[]));

do $$
declare
  table_name text;
  manage_permission text;
begin
  foreach table_name in array array[
    'organization_modules',
    'organization_roles',
    'custom_fields',
    'role_field_permissions',
    'branches',
    'departments'
  ] loop
    manage_permission = case
      when table_name = 'organization_modules' then 'settings.manage_modules'
      when table_name in ('organization_roles', 'role_field_permissions') then 'settings.manage_roles'
      else 'settings.manage_fields'
    end;
    execute format(
      'create policy "Members can read %1$I" on public.%1$I for select to authenticated using (app.is_org_member(organization_id))',
      table_name
    );
    execute format(
      'create policy "Authorized members can insert %1$I" on public.%1$I for insert to authenticated with check (app.has_permission(organization_id, %2$L))',
      table_name,
      manage_permission
    );
    execute format(
      'create policy "Authorized members can update %1$I" on public.%1$I for update to authenticated using (app.has_permission(organization_id, %2$L)) with check (app.has_permission(organization_id, %2$L))',
      table_name,
      manage_permission
    );
    execute format(
      'create policy "Authorized members can delete %1$I" on public.%1$I for delete to authenticated using (app.has_permission(organization_id, %2$L))',
      table_name,
      manage_permission
    );
  end loop;
end; $$;

do $$
declare
  table_name text;
begin
  foreach table_name in array array[
    'employees'
  ] loop
    execute format(
      'create policy "Authorized members can read %1$I" on public.%1$I for select to authenticated using (app.has_permission(organization_id, ''employees.view''))',
      table_name
    );
    execute format(
      'create policy "Authorized members can insert %1$I" on public.%1$I for insert to authenticated with check (app.has_permission(organization_id, ''employees.create''))',
      table_name
    );
    execute format(
      'create policy "Authorized members can update %1$I" on public.%1$I for update to authenticated using (app.has_permission(organization_id, ''employees.update'')) with check (app.has_permission(organization_id, ''employees.update''))',
      table_name
    );
    execute format(
      'create policy "Authorized members can delete %1$I" on public.%1$I for delete to authenticated using (app.has_permission(organization_id, ''employees.delete''))',
      table_name
    );
  end loop;
end; $$;

create policy "Sensitive people roles can read employee_compensation"
on public.employee_compensation for select
to authenticated
using (app.has_permission(organization_id, 'employees.view_salary'));

create policy "HR roles can insert employee_compensation"
on public.employee_compensation for insert
to authenticated
with check (app.has_permission(organization_id, 'employees.edit_salary'));

create policy "HR roles can update employee_compensation"
on public.employee_compensation for update
to authenticated
using (app.has_permission(organization_id, 'employees.edit_salary'))
with check (app.has_permission(organization_id, 'employees.edit_salary'));

create policy "HR roles can delete employee_compensation"
on public.employee_compensation for delete
to authenticated
using (app.has_permission(organization_id, 'employees.delete'));

do $$
declare
  table_name text;
  permission_prefix text;
begin
  foreach table_name in array array[
    'asset_categories',
    'assets',
    'asset_events',
    'maintenance_records',
    'inventory_items',
    'stock_movements'
  ] loop
    permission_prefix = case
      when table_name in ('asset_categories', 'assets', 'asset_events', 'maintenance_records') then 'assets'
      else 'inventory'
    end;
    execute format(
      'create policy "Authorized members can read %1$I" on public.%1$I for select to authenticated using (app.has_permission(organization_id, %2$L))',
      table_name,
      permission_prefix || '.view'
    );
    execute format(
      'create policy "Authorized members can insert %1$I" on public.%1$I for insert to authenticated with check (app.has_permission(organization_id, %2$L))',
      table_name,
      permission_prefix || '.create'
    );
    execute format(
      'create policy "Authorized members can update %1$I" on public.%1$I for update to authenticated using (app.has_permission(organization_id, %2$L)) with check (app.has_permission(organization_id, %2$L))',
      table_name,
      permission_prefix || '.update'
    );
    execute format(
      'create policy "Authorized members can delete %1$I" on public.%1$I for delete to authenticated using (app.has_permission(organization_id, %2$L))',
      table_name,
      permission_prefix || '.delete'
    );
  end loop;
end; $$;

do $$
declare
  table_name text;
  permission_prefix text;
begin
  foreach table_name in array array[
    'contacts',
    'accounts',
    'journal_entries',
    'journal_lines',
    'budgets',
    'document_sequences',
    'invoices',
    'invoice_items',
    'payments'
  ] loop
    permission_prefix = case
      when table_name = 'contacts' then 'contacts'
      when table_name in ('document_sequences', 'invoices', 'invoice_items') then 'sales'
      else 'finance'
    end;
    execute format(
      'create policy "Authorized members can read %1$I" on public.%1$I for select to authenticated using (app.has_permission(organization_id, %2$L))',
      table_name,
      permission_prefix || '.view'
    );
    execute format(
      'create policy "Authorized members can insert %1$I" on public.%1$I for insert to authenticated with check (app.has_permission(organization_id, %2$L))',
      table_name,
      permission_prefix || '.create'
    );
    execute format(
      'create policy "Authorized members can update %1$I" on public.%1$I for update to authenticated using (app.has_permission(organization_id, %2$L)) with check (app.has_permission(organization_id, %2$L))',
      table_name,
      permission_prefix || '.update'
    );
    execute format(
      'create policy "Authorized members can delete %1$I" on public.%1$I for delete to authenticated using (app.has_permission(organization_id, %2$L))',
      table_name,
      permission_prefix || '.delete'
    );
  end loop;
end; $$;

do $$
declare
  table_name text;
  permission_key text;
begin
  foreach table_name in array array[
    'data_sources',
    'import_runs',
    'staging_import_rows',
    'mapping_templates',
    'source_ownership',
    'documents',
    'metric_snapshots',
    'dashboard_preferences',
    'ai_conversations',
    'ai_queries'
  ] loop
    permission_key = case
      when table_name in ('data_sources', 'import_runs', 'staging_import_rows', 'mapping_templates', 'source_ownership') then 'imports.execute'
      when table_name = 'documents' then 'reports.view'
      else 'dashboard.view'
    end;
    execute format(
      'create policy "Authorized members can read %1$I" on public.%1$I for select to authenticated using (app.has_permission(organization_id, %2$L))',
      table_name,
      permission_key
    );
    execute format(
      'create policy "Authorized members can insert %1$I" on public.%1$I for insert to authenticated with check (app.has_permission(organization_id, %2$L))',
      table_name,
      permission_key
    );
    execute format(
      'create policy "Authorized members can update %1$I" on public.%1$I for update to authenticated using (app.has_permission(organization_id, %2$L)) with check (app.has_permission(organization_id, %2$L))',
      table_name,
      permission_key
    );
    execute format(
      'create policy "Authorized members can delete %1$I" on public.%1$I for delete to authenticated using (app.has_permission(organization_id, %2$L))',
      table_name,
      permission_key
    );
  end loop;
end; $$;

create policy "Owners and admins can read audit logs"
on public.audit_logs for select
to authenticated
using (app.has_org_role(organization_id, array['owner', 'admin']::public.member_role[]));
