export type MemberRole =
  | "owner"
  | "admin"
  | "finance_manager"
  | "hr_manager"
  | "operations_manager"
  | "department_manager"
  | "accountant"
  | "employee"
  | "viewer";

export type Organization = {
  id: string;
  name: string;
  legalName?: string;
  industry?: string;
  businessType?: string;
  currencyCode: string;
  timezone: string;
  fiscalYearStartMonth: number;
  setupCompletedAt?: string;
};

export type ModuleKey = "finance" | "sales" | "hr" | "supply_chain";

export type OrganizationModule = {
  id: string;
  organizationId: string;
  moduleKey: ModuleKey;
  enabled: boolean;
  displayName?: string;
  settings: Record<string, unknown>;
};

export type RecordScope = "organization" | "branch" | "department" | "own";

export type PermissionKey =
  | `${"finance" | "sales" | "employees" | "inventory" | "assets" | "contacts"}.${
      | "view"
      | "create"
      | "update"
      | "delete"
      | "approve"
      | "export"}`
  | "employees.view_salary"
  | "employees.edit_salary"
  | "inventory.adjust"
  | "inventory.transfer"
  | "imports.execute"
  | "reports.view"
  | "reports.create"
  | "reports.export"
  | "reports.share"
  | "settings.manage_fields"
  | "settings.manage_roles"
  | "settings.manage_modules"
  | "dashboard.view";

export type OrganizationRole = {
  id: string;
  organizationId: string;
  key: string;
  name: string;
  description?: string;
  systemRole?: MemberRole;
  permissions: Partial<Record<PermissionKey, boolean>>;
  recordScope: RecordScope;
  scopeSettings: Record<string, unknown>;
  isSystem: boolean;
};

export type CustomFieldType =
  | "text"
  | "long_text"
  | "number"
  | "currency"
  | "date"
  | "datetime"
  | "boolean"
  | "select"
  | "multi_select"
  | "email"
  | "phone"
  | "url";

export type CustomFieldDefinition = {
  id: string;
  organizationId: string;
  moduleKey: ModuleKey;
  targetEntity: string;
  fieldKey: string;
  label: string;
  fieldType: CustomFieldType;
  required: boolean;
  defaultValue?: unknown;
  options: unknown[];
  isVisible: boolean;
  displayOrder: number;
  active: boolean;
};

export type FieldAccessLevel = "hidden" | "visible" | "editable";

export type InvoiceStatus =
  | "draft"
  | "issued"
  | "partially_paid"
  | "paid"
  | "overdue"
  | "cancelled"
  | "void";

export type PaymentStatus = "unpaid" | "partially_paid" | "paid" | "refunded";

export type InvoiceLine = {
  id: string;
  inventoryItemId?: string;
  linePosition: number;
  itemName: string;
  description: string;
  quantity: string;
  unitPrice: string;
  discountTotal: string;
  taxTotal: string;
  lineTotal: string;
  customData: Record<string, unknown>;
};

export type Invoice = {
  id: string;
  organizationId: string;
  contactId?: string;
  invoiceNumber: string;
  invoiceType: "sales" | "purchase";
  reference?: string;
  issueDate: string;
  dueDate?: string;
  status: InvoiceStatus;
  subtotal: string;
  discountTotal: string;
  taxTotal: string;
  total: string;
  amountPaid: string;
  balanceDue: string;
  paymentStatus: PaymentStatus;
  currencyCode: string;
  notes?: string;
  customData: Record<string, unknown>;
  lines: InvoiceLine[];
};

export const invoiceStatusTransitions: Readonly<Record<InvoiceStatus, readonly InvoiceStatus[]>> = {
  draft: ["issued", "cancelled"],
  issued: ["partially_paid", "paid", "overdue", "void"],
  partially_paid: ["paid", "overdue", "void"],
  paid: ["refunded", "void"],
  overdue: ["partially_paid", "paid", "void"],
  cancelled: [],
  void: [],
};

export function canTransitionInvoice(from: InvoiceStatus, to: InvoiceStatus) {
  return invoiceStatusTransitions[from].includes(to);
}

export type DashboardMetrics = {
  revenue: number;
  expenses: number;
  netProfit: number;
  payroll: number;
  assetValue: number;
  inventoryValue: number;
  openReceivables: number;
  openPayables: number;
};

export type DataSourceType = "manual" | "csv" | "excel" | "api" | "database";

export type DataSourceStatus =
  | "fresh"
  | "ready"
  | "needs_review"
  | "stale"
  | "failed"
  | "needs_auth"
  | "planned";

export type ImportStatus =
  | "uploaded"
  | "mapping_required"
  | "validating"
  | "needs_review"
  | "ready"
  | "processing"
  | "completed"
  | "failed"
  | "cancelled";
