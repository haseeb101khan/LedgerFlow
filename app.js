const STORAGE_KEY = "ledgerflow-erp-state-v3";
const OLD_STORAGE_KEYS = ["ledgerflow-erp-state-v2", "ledgerflow-erp-state-v1"];
const SESSION_KEY = "ledgerflow-session-v1";
const SHEETJS_URL = "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";

const MODULES = ["inventory", "finance", "contacts", "employees", "assets"];

const MONEY_FIELDS = {
  inventory: ["unitCost", "sellPrice"],
  assets: ["value"],
  employees: ["salary"],
  finance: ["amount"],
  contacts: ["balance"],
};

const COMMON_CURRENCIES = ["PKR", "USD", "EUR", "GBP", "AED", "SAR", "INR", "CNY", "JPY", "CAD", "AUD"];

const navTitles = {
  inventory: "Items",
  finance: "Finance",
  contacts: "Customers & Vendors",
  employees: "Employees",
  assets: "Assets",
};

const SAMPLE_BUSINESSES = ["Pharmacy", "Utility Store", "Construction Company", "Vehicle Dealership"];

const RESERVED_FIELD_NAMES = new Set([
  "id",
  "sourceName",
  "createdAt",
  "importedAt",
  "infoDismissed",
  "passwordHash",
  "passwordSalt",
  "needsReview",
  "missingFields",
]);

const titles = {
  setup: "Setup & Roles",
  dashboard: "Dashboard",
  inventory: "Items & Inventory",
  assets: "Assets",
  employees: "Employees",
  finance: "Finance",
  contacts: "Customers & Vendors",
  dataSources: "Import Data",
  reports: "Reports",
};

const moduleActions = [
  { key: "view", label: "View" },
  { key: "add", label: "Add" },
  { key: "edit", label: "Edit" },
  { key: "delete", label: "Delete" },
];

const generalPermissions = [
  { key: "sensitiveNumbers", label: "See costs, profit, asset values & salaries" },
  { key: "reports", label: "View reports" },
  { key: "imports", label: "Import data files" },
  { key: "export", label: "Export all data" },
  { key: "settings", label: "Manage setup, fields, roles & employee logins" },
];

const currencyField = {
  name: "currency",
  label: "Currency",
  type: "select",
  required: true,
  locked: true,
  hint: "The currency these amounts were recorded in. LedgerFlow never converts it.",
};

// `section` groups fields in the record form. Fields a business adds are grouped
// under "<Business type> details"; notes always come last.
const schemas = {
  inventory: {
    title: "Item",
    fields: [
      { name: "name", label: "Item Name", type: "text", required: true, locked: true, section: "Basic info" },
      { name: "sku", label: "SKU / Code", type: "text", required: true, section: "Basic info" },
      { name: "category", label: "Category", type: "text", required: true, section: "Basic info" },
      { name: "quantity", label: "Quantity", type: "number", required: true, min: 0, step: "any", locked: true, section: "Stock" },
      { name: "reorderLevel", label: "Reorder Level", type: "number", min: 0, step: "any", section: "Stock" },
      { name: "location", label: "Location", type: "text", section: "Stock" },
      { name: "unitCost", label: "Unit Cost", type: "number", required: true, min: 0, step: "0.01", sensitive: true, section: "Pricing" },
      { name: "sellPrice", label: "Sell Price", type: "number", required: true, min: 0, step: "0.01", section: "Pricing" },
      { ...currencyField, section: "Pricing" },
      { name: "notes", label: "Notes", type: "textarea", full: true, section: "Notes" },
    ],
  },
  assets: {
    title: "Asset",
    fields: [
      { name: "name", label: "Asset Name", type: "text", required: true, locked: true, section: "Basic info" },
      { name: "code", label: "Asset Code", type: "text", required: true, section: "Basic info" },
      {
        name: "condition",
        label: "Condition",
        type: "select",
        options: ["Active", "Maintenance", "Damaged", "Retired"],
        locked: true,
        section: "Basic info",
      },
      { name: "assignedTo", label: "Assigned To", type: "text", section: "Assignment" },
      { name: "location", label: "Location", type: "text", section: "Assignment" },
      { name: "purchaseDate", label: "Purchase Date", type: "date", section: "Value" },
      { name: "value", label: "Value", type: "number", required: true, min: 0, step: "0.01", sensitive: true, section: "Value" },
      { ...currencyField, section: "Value" },
      { name: "notes", label: "Notes", type: "textarea", full: true, section: "Notes" },
    ],
  },
  employees: {
    title: "Employee",
    fields: [
      { name: "name", label: "Full Name", type: "text", required: true, locked: true, section: "Personal" },
      { name: "phone", label: "Phone", type: "text", section: "Personal" },
      { name: "email", label: "Email", type: "email", section: "Personal" },
      { name: "role", label: "Job Title", type: "text", required: true, section: "Job" },
      { name: "department", label: "Department", type: "text", required: true, section: "Job" },
      { name: "status", label: "Status", type: "select", options: ["Active", "On Leave", "Inactive"], locked: true, section: "Job" },
      { name: "salary", label: "Monthly Salary", type: "number", required: true, min: 0, step: "0.01", sensitive: true, section: "Pay" },
      { ...currencyField, section: "Pay" },
      {
        name: "loginEmail",
        label: "Login Email",
        type: "email",
        settingsOnly: true,
        locked: true,
        hint: "The email this employee signs in with.",
        section: "Login & access",
      },
      {
        name: "loginPassword",
        label: "Login Password",
        type: "password",
        settingsOnly: true,
        virtual: true,
        locked: true,
        hint: "Give this password to the employee. At least 6 characters.",
        section: "Login & access",
      },
      { name: "accessRole", label: "Access Role", type: "select", settingsOnly: true, locked: true, section: "Login & access" },
      { name: "notes", label: "Notes", type: "textarea", full: true, section: "Notes" },
    ],
  },
  finance: {
    title: "Transaction",
    fields: [
      { name: "type", label: "Type", type: "select", options: ["Income", "Expense"], required: true, locked: true, section: "Transaction" },
      { name: "date", label: "Date", type: "date", required: true, locked: true, section: "Transaction" },
      { name: "category", label: "Category", type: "text", required: true, section: "Transaction", list: "categoryOptions" },
      { name: "description", label: "Description", type: "text", required: true, section: "Transaction" },
      { name: "amount", label: "Amount", type: "number", required: true, min: 0, step: "0.01", locked: true, section: "Amount & payment" },
      { ...currencyField, section: "Amount & payment" },
      {
        name: "status",
        label: "Payment Status",
        type: "select",
        options: ["Paid", "Pending", "Overdue"],
        locked: true,
        section: "Amount & payment",
        hint: "Pending and overdue transactions appear under Dues.",
      },
      { name: "dueDate", label: "Due Date", type: "date", locked: true, section: "Amount & payment", hint: "When an unpaid amount should be paid." },
      {
        name: "party",
        label: "Customer / Supplier",
        type: "text",
        locked: true,
        section: "Who",
        list: "partyOptions",
        hint: "Pick an existing contact or type any name (for example a utility company).",
      },
      { name: "notes", label: "Notes", type: "textarea", full: true, section: "Notes" },
    ],
  },
  contacts: {
    title: "Contact",
    fields: [
      { name: "name", label: "Name", type: "text", required: true, locked: true, section: "Contact" },
      { name: "type", label: "Type", type: "select", options: ["Customer", "Supplier"], required: true, locked: true, section: "Contact" },
      { name: "phone", label: "Phone", type: "text", section: "Contact" },
      { name: "email", label: "Email", type: "email", section: "Contact" },
      {
        name: "balance",
        label: "Opening Balance",
        type: "number",
        step: "0.01",
        locked: true,
        section: "Account",
        hint: "What was owed before you started using LedgerFlow. Positive: they owe you. Negative: you owe them. Record new bills in Finance.",
      },
      { ...currencyField, section: "Account" },
      { name: "status", label: "Status", type: "select", options: ["Active", "Pending", "Blocked", "Inactive"], locked: true, section: "Account" },
      { name: "notes", label: "Notes", type: "textarea", full: true, section: "Notes" },
    ],
  },
};

const importAliases = {
  inventory: {
    name: ["item", "itemname", "product", "productname", "particulars", "medicine", "article"],
    sku: ["sku", "code", "itemcode", "productcode", "barcode", "partnumber"],
    category: ["category", "group", "itemgroup", "type", "department"],
    quantity: ["quantity", "qty", "stock", "stockqty", "onhand", "units", "balanceqty"],
    reorderLevel: ["reorderlevel", "minimumstock", "minstock", "reorder", "alertlevel", "minqty"],
    unitCost: ["unitcost", "cost", "purchaseprice", "buyprice", "averagecost", "costprice", "tradeprice"],
    sellPrice: ["sellprice", "price", "sellingprice", "retailprice", "rate", "saleprice", "mrp"],
    location: ["location", "warehouse", "store", "aisle", "shelf", "rack"],
    notes: ["notes", "remarks", "description"],
  },
  assets: {
    name: ["asset", "assetname", "item", "name"],
    code: ["code", "assetcode", "tag", "assettag", "serial", "serialnumber"],
    assignedTo: ["assignedto", "custodian", "employee", "owner", "department", "user"],
    location: ["location", "branch", "site", "room"],
    purchaseDate: ["purchasedate", "date", "buydate", "acquireddate", "acquisitiondate"],
    value: ["value", "cost", "purchasevalue", "amount", "bookvalue", "price"],
    condition: ["condition", "status", "state"],
    notes: ["notes", "remarks", "description"],
  },
  employees: {
    name: ["employee", "employeename", "worker", "workername", "staffname", "name", "fullname"],
    role: ["role", "jobtitle", "designation", "position", "title"],
    department: ["department", "division", "team", "section", "dept"],
    phone: ["phone", "mobile", "contact", "contactnumber", "cell", "phonenumber", "mobilenumber"],
    email: ["email", "emailaddress", "mail"],
    loginEmail: ["login", "loginemail", "useremail", "accountemail", "username"],
    accessRole: ["accessrole", "systemrole", "permissions", "dashboardrole"],
    salary: ["salary", "pay", "monthlysalary", "wage", "compensation", "basicpay"],
    status: ["status", "employmentstatus", "state"],
    notes: ["notes", "remarks"],
  },
  finance: {
    date: ["date", "transactiondate", "postingdate", "voucherdate", "billdate"],
    type: ["type", "transactiontype", "kind", "drcr"],
    category: ["category", "account", "head", "ledger", "expensehead", "accounthead"],
    description: ["description", "details", "memo", "narration", "particulars"],
    amount: ["amount", "value", "total", "debit", "credit", "netamount"],
    status: ["status", "paymentstatus", "state"],
    notes: ["notes", "remarks"],
  },
  contacts: {
    name: ["name", "contact", "customer", "vendor", "supplier", "party", "partyname", "customername", "suppliername"],
    type: ["type", "partytype", "contacttype"],
    phone: ["phone", "mobile", "contactnumber", "cell", "phonenumber"],
    email: ["email", "emailaddress"],
    balance: ["balance", "outstanding", "receivable", "payable", "amount", "due", "dues"],
    status: ["status", "state"],
    notes: ["notes", "remarks", "address"],
  },
};

MODULES.forEach((module) => {
  importAliases[module].currency = ["currency", "curr", "ccy", "currencycode"];
});
importAliases.finance.party = ["party", "customer", "supplier", "vendor", "payee", "paidto", "receivedfrom", "client"];
importAliases.finance.dueDate = ["duedate", "due", "paymentdue", "dueon"];

const recommendedImportFields = {
  inventory: ["name", "sku", "category", "quantity", "unitCost", "sellPrice"],
  assets: ["name", "code", "assignedTo", "location", "value"],
  employees: ["name", "role", "department", "phone", "email", "loginEmail"],
  finance: ["date", "type", "category", "description", "amount"],
  contacts: ["name", "type", "phone", "balance"],
};

/*
 * A business template arranges LedgerFlow for one kind of business:
 * module names, what one record is called, base field labels, base fields
 * to hide, and extra fields ("Label|type|choice,choice").
 */
const businessScopes = {
  "General Business": {
    summary: "Items, finance, customers, employees, and assets with general fields.",
    fields: {
      inventory: ["Category Code"],
      finance: ["Reference"],
    },
  },
  "Utility Store": {
    summary: "Grocery and household items with brands, pack sizes, and shelf codes.",
    labels: { inventory: "Store Items", contacts: "Customers & Suppliers" },
    fieldLabels: { inventory: { location: "Shelf / Aisle" } },
    fields: {
      inventory: ["Brand", "Pack Size", "Barcode"],
      finance: ["Shift|select|Morning,Evening"],
      contacts: ["Area"],
    },
  },
  "Mart / Retail": {
    summary: "Retail items with barcodes and brands, counters and shifts.",
    labels: { inventory: "Products", contacts: "Customers & Suppliers" },
    nouns: { inventory: "Product" },
    fields: {
      inventory: ["Barcode", "Brand"],
      finance: ["Counter"],
      employees: ["Shift|select|Morning,Evening,Night"],
    },
  },
  Pharmacy: {
    summary: "Medicines with batch numbers, expiry dates, manufacturers, and prescriptions.",
    labels: { inventory: "Medicines", contacts: "Distributors & Customers" },
    nouns: { inventory: "Medicine" },
    fieldLabels: {
      inventory: { name: "Medicine Name", sku: "Product Code", category: "Medicine Group", location: "Rack" },
    },
    fields: {
      inventory: ["Batch Number", "Expiry Date|date", "Manufacturer", "Prescription Required|select|No,Yes"],
      contacts: ["Drug License No."],
    },
  },
  "Clothing Store": {
    summary: "Garments with sizes, colors, and seasons.",
    labels: { inventory: "Garments" },
    nouns: { inventory: "Garment" },
    fields: {
      inventory: ["Size|select|XS,S,M,L,XL,XXL", "Color", "Season|select|Summer,Winter,All Season"],
      contacts: ["Preferred Size"],
    },
  },
  "Vehicle Dealership": {
    summary: "Vehicles with make, model year, chassis number, mileage, and condition.",
    labels: { inventory: "Vehicles", contacts: "Buyers & Suppliers", assets: "Workshop & Equipment" },
    nouns: { inventory: "Vehicle" },
    fieldLabels: {
      inventory: {
        name: "Vehicle",
        sku: "Stock No.",
        category: "Body Type",
        quantity: "Units",
        unitCost: "Purchase Cost",
        sellPrice: "Asking Price",
        location: "Showroom / Yard",
      },
    },
    hide: { inventory: ["reorderLevel"] },
    fields: {
      inventory: [
        "Make",
        "Model Year|number",
        "Chassis No.",
        "Engine (cc)|number",
        "Mileage (km)|number",
        "Color",
        "Vehicle Condition|select|New,Used,Reconditioned",
      ],
      contacts: ["CNIC / Company Reg."],
      finance: ["Stock No."],
    },
  },
  "Construction Company": {
    summary: "Building materials, machinery, projects, sites, and trades.",
    labels: { inventory: "Materials", assets: "Machinery", contacts: "Clients & Suppliers" },
    nouns: { inventory: "Material", assets: "Machine" },
    fieldLabels: {
      inventory: {
        name: "Material",
        sku: "Material Code",
        unitCost: "Purchase Rate",
        sellPrice: "Billing Rate",
        location: "Store / Site",
      },
      assets: { assignedTo: "Operator", location: "Project Site" },
    },
    fields: {
      inventory: ["Unit|select|Bags,Tons,Cubic ft,Pieces,Sq ft,Rolls,Meters"],
      assets: ["Hours Used|number"],
      finance: ["Project"],
      employees: ["Trade|select|Engineer,Supervisor,Mason,Steel Fixer,Electrician,Plumber,Carpenter,Labour", "Site"],
      contacts: ["Project"],
    },
  },
};

const LEGACY_BUSINESS_TYPES = {
  "Vehicle / Parts": "Vehicle Dealership",
  Construction: "Construction Company",
};

let state = loadState();
let session = loadSession();
let previewRoleId = null;
let currentView = "dashboard";
let editing = null;
let importSession = null;
let editingField = null;
let sheetJsPromise = null;
let resizeTimer = null;

const byId = (id) => document.getElementById(id);

const els = {
  authScreen: byId("authScreen"),
  appShell: byId("appShell"),
  registerForm: byId("registerForm"),
  registerError: byId("registerError"),
  registerIntro: byId("registerIntro"),
  loginForm: byId("loginForm"),
  loginError: byId("loginError"),
  loginTitle: byId("loginTitle"),
  resetBrowserBtn: byId("resetBrowserBtn"),
  navList: byId("navList"),
  topbarBusiness: byId("topbarBusiness"),
  sidebarBusinessName: byId("sidebarBusinessName"),
  userAvatar: byId("userAvatar"),
  userName: byId("userName"),
  userRole: byId("userRole"),
  logoutBtn: byId("logoutBtn"),
  globalSearch: byId("globalSearch"),
  previewSwitch: byId("previewSwitch"),
  previewRoleSelect: byId("previewRoleSelect"),
  previewBanner: byId("previewBanner"),
  previewBannerText: byId("previewBannerText"),
  exitPreviewBtn: byId("exitPreviewBtn"),
  quickAddBtn: byId("quickAddBtn"),
  exportBtn: byId("exportBtn"),
  clearActivityBtn: byId("clearActivityBtn"),
  hideGettingStartedBtn: byId("hideGettingStartedBtn"),
  printReportBtn: byId("printReportBtn"),
  sampleCsvBtn: byId("sampleCsvBtn"),
  importTarget: byId("importTarget"),
  importFileInput: byId("importFileInput"),
  importFileName: byId("importFileName"),
  importWizard: byId("importWizard"),
  setupForm: byId("setupForm"),
  businessNameInput: byId("businessNameInput"),
  businessTypeSelect: byId("businessTypeSelect"),
  currencyInput: byId("currencyInput"),
  dateFormatSelect: byId("dateFormatSelect"),
  ownerFields: byId("ownerFields"),
  ownerNameInput: byId("ownerNameInput"),
  ownerEmailInput: byId("ownerEmailInput"),
  ownerPasswordInput: byId("ownerPasswordInput"),
  newRoleName: byId("newRoleName"),
  newRoleDescription: byId("newRoleDescription"),
  newRoleCopyFrom: byId("newRoleCopyFrom"),
  addRoleBtn: byId("addRoleBtn"),
  fieldModuleSelect: byId("fieldModuleSelect"),
  fieldList: byId("fieldList"),
  fieldEditor: byId("fieldEditor"),
  fieldEditorTitle: byId("fieldEditorTitle"),
  fieldLabelInput: byId("fieldLabelInput"),
  fieldTypeInput: byId("fieldTypeInput"),
  fieldOptionsControl: byId("fieldOptionsControl"),
  fieldOptionsInput: byId("fieldOptionsInput"),
  fieldRequiredInput: byId("fieldRequiredInput"),
  fieldHiddenControl: byId("fieldHiddenControl"),
  fieldHiddenInput: byId("fieldHiddenInput"),
  fieldEditorNote: byId("fieldEditorNote"),
  saveFieldBtn: byId("saveFieldBtn"),
  cancelFieldEditBtn: byId("cancelFieldEditBtn"),
  dataToolsPanel: byId("dataToolsPanel"),
  clearRecordsBtn: byId("clearRecordsBtn"),
  clearSampleBtn: byId("clearSampleBtn"),
  modalBackdrop: byId("modalBackdrop"),
  closeModalBtn: byId("closeModalBtn"),
  modalTitle: byId("modalTitle"),
  modalKicker: byId("modalKicker"),
  recordForm: byId("recordForm"),
  toast: byId("toast"),
  menuToggle: byId("menuToggle"),
};

/* ---------- State and storage ---------- */

function emptyState() {
  return {
    version: 2,
    organization: null,
    roles: defaultRoles(),
    customFields: Object.fromEntries(MODULES.map((module) => [module, []])),
    fieldSettings: Object.fromEntries(MODULES.map((module) => [module, {}])),
    inventory: [],
    assets: [],
    employees: [],
    finance: [],
    contacts: [],
    dataSources: defaultDataSources(),
    activity: [],
  };
}

function defaultDataSources() {
  return [
    {
      id: "source-manual",
      name: "Manual Entry",
      type: "Manual",
      target: "All modules",
      status: "Ready",
      lastSync: "",
      nextSync: "On save",
      records: 0,
      note: "Records typed directly into LedgerFlow.",
    },
    {
      id: "source-file",
      name: "File Import",
      type: "File",
      target: "Any module",
      status: "Ready",
      lastSync: "",
      nextSync: "On upload",
      records: 0,
      note: "CSV, Excel, JSON, or tab-separated files matched into LedgerFlow fields.",
    },
    {
      id: "source-accounting",
      name: "Accounting Connector",
      type: "API",
      target: "Finance",
      status: "Planned",
      lastSync: "",
      nextSync: "Not connected",
      records: 0,
      note: "Future QuickBooks, Xero, Zoho Books, or local accounting integration.",
    },
    {
      id: "source-database",
      name: "Database Connector",
      type: "Database",
      target: "All modules",
      status: "Planned",
      lastSync: "",
      nextSync: "Not connected",
      records: 0,
      note: "Future PostgreSQL, MySQL, SQL Server, or Oracle read-only connection.",
    },
    {
      id: "source-pos-crm",
      name: "POS / CRM Connector",
      type: "API",
      target: "Sales and contacts",
      status: "Planned",
      lastSync: "",
      nextSync: "Not connected",
      records: 0,
      note: "Future sales, customer, and invoice source for shops and firms.",
    },
  ];
}

function blankPermissions() {
  return {
    modules: Object.fromEntries(
      MODULES.map((module) => [module, { view: false, add: false, edit: false, delete: false }]),
    ),
    ...Object.fromEntries(generalPermissions.map((permission) => [permission.key, false])),
  };
}

function fullPermissions() {
  return {
    modules: Object.fromEntries(
      MODULES.map((module) => [module, { view: true, add: true, edit: true, delete: true }]),
    ),
    ...Object.fromEntries(generalPermissions.map((permission) => [permission.key, true])),
  };
}

function buildPermissions(moduleAccess, general = {}) {
  const permissions = blankPermissions();
  Object.entries(moduleAccess).forEach(([module, letters]) => {
    permissions.modules[module] = {
      view: letters.includes("v"),
      add: letters.includes("a"),
      edit: letters.includes("e"),
      delete: letters.includes("d"),
    };
  });
  return { ...permissions, ...general };
}

function defaultRoles() {
  return [
    {
      id: "owner",
      name: "Owner",
      description: "Full control of the business. The owner role cannot be restricted.",
      system: true,
      permissions: fullPermissions(),
    },
    {
      id: "manager",
      name: "Manager",
      description: "Runs daily operations, sees numbers and reports, cannot change setup or logins.",
      system: false,
      permissions: buildPermissions(
        { inventory: "vaed", finance: "vae", contacts: "vaed", employees: "v", assets: "vae" },
        { sensitiveNumbers: true, reports: true, imports: true },
      ),
    },
    {
      id: "staff",
      name: "Staff",
      description: "Daily work: stock updates, recording sales, and customer dues.",
      system: false,
      permissions: buildPermissions({ inventory: "vae", finance: "va", contacts: "vae" }),
    },
  ];
}

function normalizePermissions(raw = {}) {
  const permissions = blankPermissions();
  const legacy = !raw.modules;
  MODULES.forEach((module) => {
    if (legacy) {
      const canView = Boolean(raw[module]);
      permissions.modules[module] = {
        view: canView,
        add: canView && Boolean(raw.createEditRecords),
        edit: canView && Boolean(raw.createEditRecords),
        delete: canView && Boolean(raw.deleteRecords),
      };
      return;
    }
    const entry = raw.modules[module] || {};
    moduleActions.forEach(({ key }) => {
      permissions.modules[module][key] = Boolean(entry[key]);
    });
    if (entry.add || entry.edit || entry.delete) permissions.modules[module].view = true;
  });
  generalPermissions.forEach(({ key }) => {
    const legacyValue = key === "imports" ? raw.dataSources : raw[key];
    permissions[key] = Boolean(legacy ? legacyValue : raw[key]);
  });
  return permissions;
}

function loadState() {
  const saved = readStorage(STORAGE_KEY);
  if (saved) return normalizeLoadedState(saved);

  // Earlier prototype versions held test uploads and demo records. Keep only the
  // owner's account and roles so the owner signs in to an empty business.
  const previous = OLD_STORAGE_KEYS.map(readStorage).find((entry) => entry?.organization?.ownerPasswordHash);
  const fresh = previous
    ? normalizeLoadedState({ organization: previous.organization, roles: previous.roles })
    : normalizeLoadedState(emptyState());
  if (previous) {
    delete fresh.organization.hideGettingStarted;
    delete fresh.organization.sampleLoaded;
    fresh.pendingTemplate = true;
  }
  return fresh;
}

function readStorage(key) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : null;
  } catch (error) {
    console.warn(`Could not read ${key}`, error);
    return null;
  }
}

function normalizeLoadedState(loaded) {
  const base = emptyState();
  const next = { ...base, ...loaded, version: 3 };
  delete next.activeRole;

  next.organization = loaded.organization
    ? {
        type: "General Business",
        currency: "USD",
        dateFormat: defaultDateFormat(),
        moduleLabels: {},
        recordNouns: {},
        ...loaded.organization,
      }
    : null;
  if (next.organization) {
    const type = LEGACY_BUSINESS_TYPES[next.organization.type] || next.organization.type;
    next.organization.type = businessScopes[type] ? type : "General Business";
  }

  MODULES.forEach((module) => {
    next[module] = Array.isArray(loaded[module]) ? loaded[module] : [];
  });
  next.customFields = Object.fromEntries(
    MODULES.map((module) => [
      module,
      (Array.isArray(loaded.customFields?.[module]) ? loaded.customFields[module] : []).map((field) => ({
        name: field.name,
        label: field.label || field.name,
        type: ["text", "number", "date", "select"].includes(field.type) ? field.type : "text",
        options: Array.isArray(field.options) ? field.options : [],
        required: Boolean(field.required),
      })),
    ]),
  );
  next.fieldSettings = Object.fromEntries(MODULES.map((module) => [module, loaded.fieldSettings?.[module] || {}]));

  const roles = Array.isArray(loaded.roles) && loaded.roles.length ? loaded.roles : base.roles;
  next.roles = roles.map((role) => ({
    id: role.id || makeId("role"),
    name: role.name || "Role",
    description: role.description || "Custom access role.",
    system: role.id === "owner",
    permissions: role.id === "owner" ? fullPermissions() : normalizePermissions(role.permissions),
  }));
  if (!next.roles.some((role) => role.id === "owner")) next.roles.unshift(base.roles[0]);

  next.employees = next.employees.map((employee) => {
    const cleaned = { ...employee };
    delete cleaned.inviteStatus;
    delete cleaned.needsReview;
    delete cleaned.missingFields;
    return {
      ...cleaned,
      loginEmail: String(employee.loginEmail || "").toLowerCase(),
      accessRole: assignableRoles(next.roles).some((role) => role.id === employee.accessRole)
        ? employee.accessRole
        : fallbackRoleId(next.roles),
    };
  });
  MODULES.filter((module) => module !== "employees").forEach((module) => {
    next[module] = next[module].map((record) => {
      const cleaned = { ...record };
      delete cleaned.needsReview;
      delete cleaned.missingFields;
      return cleaned;
    });
  });
  next.contacts = next.contacts.map((contact) => (contact.type === "Vendor" ? { ...contact, type: "Supplier" } : contact));
  // A record's currency is frozen when it is created, so later changing the
  // business's default currency never relabels existing amounts.
  const fallbackCurrency = next.organization?.currency || "USD";
  MODULES.forEach((module) => {
    next[module] = next[module].map((record) =>
      isValidCurrency(record.currency) ? record : { ...record, currency: fallbackCurrency },
    );
  });

  next.dataSources = (Array.isArray(loaded.dataSources) ? loaded.dataSources : base.dataSources).map((source) =>
    source.id === "source-csv"
      ? { ...source, id: "source-file", name: "File Import", status: source.status === "Fresh" ? "Fresh" : "Ready" }
      : source,
  );
  next.activity = Array.isArray(loaded.activity) ? loaded.activity : [];
  return next;
}

function saveState() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  } catch (error) {
    console.error("Could not save ERP state", error);
    showToast("Could not save. Browser storage may be full.");
  }
}

function loadSession() {
  return readStorage(SESSION_KEY);
}

function saveSession(value) {
  try {
    if (value) localStorage.setItem(SESSION_KEY, JSON.stringify(value));
    else localStorage.removeItem(SESSION_KEY);
  } catch (error) {
    console.warn("Could not save session", error);
  }
}

function defaultDateFormat() {
  return (navigator.language || "").toLowerCase() === "en-us" ? "MDY" : "DMY";
}

/* ---------- Accounts and access ---------- */

function toHex(bytes) {
  return [...bytes].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

function makeSalt() {
  const bytes = new Uint8Array(16);
  if (globalThis.crypto?.getRandomValues) {
    globalThis.crypto.getRandomValues(bytes);
  } else {
    bytes.forEach((_, index) => {
      bytes[index] = Math.floor(Math.random() * 256);
    });
  }
  return toHex(bytes);
}

function fallbackHash(data) {
  let hash = 2166136261;
  data.forEach((byte) => {
    hash ^= byte;
    hash = Math.imul(hash, 16777619);
  });
  return `fnv-${(hash >>> 0).toString(16)}`;
}

async function hashPassword(password, salt) {
  const data = new TextEncoder().encode(`${salt}:${password}`);
  if (globalThis.crypto?.subtle) {
    return toHex(new Uint8Array(await globalThis.crypto.subtle.digest("SHA-256", data)));
  }
  return fallbackHash(data);
}

async function verifyPassword(password, salt, expectedHash) {
  if (!salt || !expectedHash) return false;
  if (expectedHash.startsWith("fnv-")) {
    return fallbackHash(new TextEncoder().encode(`${salt}:${password}`)) === expectedHash;
  }
  return (await hashPassword(password, salt)) === expectedHash;
}

function isRegistered() {
  return Boolean(state.organization?.name && state.organization?.ownerPasswordHash);
}

function employeeCanLogin(employee) {
  return Boolean(employee.loginEmail && employee.passwordHash && employee.status !== "Inactive");
}

function loginStatus(employee) {
  if (!employee.loginEmail) return "No login";
  if (!employee.passwordHash) return "Password not set";
  if (employee.status === "Inactive") return "Blocked (inactive)";
  return "Can sign in";
}

function currentUser() {
  if (!session?.userId || !isRegistered()) return null;
  if (session.userId === "owner") {
    return {
      id: "owner",
      name: state.organization.ownerName || "Owner",
      email: state.organization.ownerEmail,
      roleId: "owner",
      isOwner: true,
    };
  }
  const employee = state.employees.find((entry) => entry.id === session.userId);
  if (!employee || !employeeCanLogin(employee)) return null;
  return {
    id: employee.id,
    name: employee.name,
    email: employee.loginEmail,
    roleId: employee.accessRole,
    isOwner: false,
  };
}

function assignableRoles(roles = state.roles) {
  return roles.filter((role) => role.id !== "owner");
}

function fallbackRoleId(roles = state.roles) {
  const assignable = assignableRoles(roles);
  return (assignable.find((role) => role.id === "staff") || assignable[0])?.id || "";
}

function currentRole() {
  const user = currentUser();
  if (!user) return null;
  if (user.isOwner) {
    const preview = previewRoleId && state.roles.find((role) => role.id === previewRoleId);
    return preview || state.roles.find((role) => role.id === "owner");
  }
  return (
    assignableRoles().find((role) => role.id === user.roleId) || {
      id: "none",
      name: "No access",
      permissions: blankPermissions(),
    }
  );
}

function canAccess(permission) {
  return Boolean(currentRole()?.permissions?.[permission]);
}

function can(module, action) {
  return Boolean(currentRole()?.permissions?.modules?.[module]?.[action]);
}

function canOpenView(view) {
  if (view === "dashboard") return true;
  if (view === "setup") return canAccess("settings");
  if (view === "dataSources") return canAccess("imports") && MODULES.some((module) => can(module, "add"));
  if (view === "reports") return canAccess("reports");
  if (MODULES.includes(view)) return can(view, "view");
  return false;
}

function roleName(roleId) {
  return state.roles.find((role) => role.id === roleId)?.name || "No role";
}

function startSession(userId) {
  session = { userId, startedAt: new Date().toISOString() };
  saveSession(session);
  previewRoleId = null;
  currentView = "dashboard";
  els.globalSearch.value = "";
  showApp();
}

function logout() {
  session = null;
  previewRoleId = null;
  importSession = null;
  saveSession(null);
  closeModal();
  showAuth();
}

/* ---------- Fields ---------- */

function getFields(module, options = {}) {
  const overrides = state.fieldSettings?.[module] || {};
  const base = schemas[module].fields.map((field) => {
    const override = overrides[field.name] || {};
    return {
      ...field,
      label: override.label || field.label,
      required: field.locked ? Boolean(field.required) : Boolean(override.required ?? field.required),
      hidden: field.locked ? false : Boolean(override.hidden),
    };
  });
  const custom = (state.customFields?.[module] || []).map((field) => ({ ...field, custom: true }));
  return [...base, ...custom].filter((field) => options.includeHidden || !field.hidden);
}

function canSeeField(field) {
  if (field.sensitive && !canAccess("sensitiveNumbers")) return false;
  if (field.settingsOnly && !canAccess("settings")) return false;
  return true;
}

function formFields(module) {
  return getFields(module).filter(canSeeField);
}

function dataFields(module) {
  return getFields(module).filter((field) => !field.virtual);
}

function importableFields(module) {
  return dataFields(module).filter(canSeeField);
}

function findField(module, name) {
  return getFields(module).find((field) => field.name === name);
}

function fieldVisible(module, name) {
  const field = findField(module, name);
  return Boolean(field && canSeeField(field));
}

function fieldLabel(module, name) {
  return getFields(module, { includeHidden: true }).find((field) => field.name === name)?.label || name;
}

function isBlank(value) {
  return value === undefined || value === null || String(value).trim() === "";
}

function missingFieldsFor(module, record) {
  const fields = dataFields(module);
  const nagRecommended = Boolean(record.importedAt) && !record.infoDismissed;
  const recommended = new Set(nagRecommended ? recommendedImportFields[module] || [] : []);
  const missing = fields.filter(
    (field) => (field.required || recommended.has(field.name)) && isBlank(record[field.name]),
  );
  if (module === "employees") {
    const bothBlank = isBlank(record.phone) && isBlank(record.email);
    return missing.filter((field) => (field.name === "phone" || field.name === "email" ? bothBlank : true));
  }
  return missing;
}

function visibleMissingFields(module, record) {
  return missingFieldsFor(module, record).filter(canSeeField);
}

function needsInfo(module, record) {
  return visibleMissingFields(module, record).length > 0;
}

function missingLabels(module, fields) {
  const names = fields.map((field) => field.name);
  const labels = fields.map((field) => field.label);
  if (module === "employees" && names.includes("phone") && names.includes("email")) {
    return [
      ...fields.filter((field) => field.name !== "phone" && field.name !== "email").map((field) => field.label),
      `${fieldLabel(module, "phone")} or ${fieldLabel(module, "email")}`,
    ];
  }
  return labels;
}

function hasRequiredMissing(module, record) {
  return visibleMissingFields(module, record).some((field) => field.required);
}

function makeFieldName(label) {
  const words = String(label)
    .trim()
    .replace(/[^a-zA-Z0-9 ]/g, " ")
    .split(/\s+/)
    .filter(Boolean);
  const [first = "field", ...rest] = words;
  const name = `${first.toLowerCase()}${rest.map((word) => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()).join("")}`;
  return /^\d/.test(name) ? `field${name}` : name;
}

function uniqueFieldName(module, label) {
  const taken = new Set([
    ...RESERVED_FIELD_NAMES,
    ...getFields(module, { includeHidden: true }).map((field) => field.name),
  ]);
  const base = makeFieldName(label);
  let name = base;
  let counter = 2;
  while (taken.has(name)) {
    name = `${base}${counter}`;
    counter += 1;
  }
  return name;
}

function inferCustomFieldType(label) {
  return /date|expiry|warranty|dob/i.test(label) ? "date" : "text";
}

/* ---------- Formatting ---------- */

function formatNumber(value) {
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: 2 }).format(Number(value) || 0);
}

/*
 * Money rules:
 * - Every amount is shown with the currency its record was entered in.
 * - Amounts in different currencies are never added together or converted.
 *   Totals are kept per currency and shown side by side.
 */
function formatCurrency(value, currency) {
  const code = isValidCurrency(currency) ? currency : orgCurrency();
  try {
    return new Intl.NumberFormat("en-US", {
      style: "currency",
      currency: code,
      currencyDisplay: "code",
      minimumFractionDigits: 0,
      maximumFractionDigits: 2,
    }).format(Number(value) || 0);
  } catch (error) {
    return `${code} ${formatNumber(value)}`;
  }
}

function orgCurrency() {
  return state?.organization?.currency || "USD";
}

function recordCurrency(record) {
  return isValidCurrency(record?.currency) ? record.currency : orgCurrency();
}

function formatRecordMoney(record, value) {
  if (isBlank(value)) return "—";
  return formatCurrency(value, recordCurrency(record));
}

function currencyOptions() {
  const used = MODULES.flatMap((module) => state[module].map((record) => record.currency));
  return [...new Set([orgCurrency(), ...COMMON_CURRENCIES, ...used].filter(isValidCurrency))];
}

function sumByCurrency(records, amountOf) {
  return records.reduce((totals, record) => {
    const amount = amountOf(record);
    if (!amount) return totals;
    const code = recordCurrency(record);
    totals[code] = (totals[code] || 0) + amount;
    return totals;
  }, {});
}

function subtractByCurrency(left, right) {
  const result = { ...left };
  Object.entries(right).forEach(([code, amount]) => {
    result[code] = (result[code] || 0) - amount;
  });
  return result;
}

function currencyEntries(totals) {
  const primary = orgCurrency();
  return Object.entries(totals)
    .filter(([, amount]) => Math.abs(amount) > 0.000001)
    .sort(([a], [b]) => (a === primary ? -1 : b === primary ? 1 : a.localeCompare(b)));
}

function formatMoneyTotals(totals) {
  const entries = currencyEntries(totals);
  if (!entries.length) return formatCurrency(0, orgCurrency());
  return entries.map(([code, amount]) => formatCurrency(amount, code)).join(" · ");
}

function moneyHeadline(totals) {
  const entries = currencyEntries(totals);
  if (!entries.length) return { value: formatCurrency(0, orgCurrency()), others: "" };
  const [[code, amount], ...rest] = entries;
  return {
    value: formatCurrency(amount, code),
    others: rest.length
      ? `Also ${rest.map(([other, value]) => formatCurrency(value, other)).join(" · ")} (separate, not converted)`
      : "",
  };
}

function currencyFromText(text) {
  const value = String(text ?? "").toUpperCase();
  if (!value.trim()) return "";
  const code = value.match(/(?:^|[^A-Z])([A-Z]{3})(?=$|[^A-Z])/g)
    ?.map((match) => match.replace(/[^A-Z]/g, ""))
    .find((candidate) => COMMON_CURRENCIES.includes(candidate) || ["QAR", "KWD", "OMR", "BDT", "LKR", "TRY", "MYR"].includes(candidate));
  if (code) return code;
  if (/₨|(?:^|[^A-Z])RS\.?(?=\s|\d|$)|RUPEE/.test(value)) return ["INR", "LKR", "NPR"].includes(orgCurrency()) ? orgCurrency() : "PKR";
  if (value.includes("₹")) return "INR";
  if (value.includes("€") || value.includes("EURO")) return "EUR";
  if (value.includes("£") || value.includes("POUND")) return "GBP";
  if (value.includes("¥") || value.includes("YEN")) return "JPY";
  if (value.includes("DIRHAM")) return "AED";
  if (value.includes("RIYAL")) return "SAR";
  if (value.includes("$") || value.includes("DOLLAR")) return "USD";
  return "";
}

function viewTitle(key) {
  return state?.organization?.moduleLabels?.[key] || titles[key];
}

function navTitle(key) {
  return state?.organization?.moduleLabels?.[key] || navTitles[key] || titles[key];
}

function recordNoun(module) {
  return state?.organization?.recordNouns?.[module] || schemas[module].title;
}

function totalRecordCount() {
  return MODULES.reduce((count, module) => count + state[module].length, 0);
}

function isValidCurrency(code) {
  try {
    new Intl.NumberFormat("en-US", { style: "currency", currency: code }).format(1);
    return /^[A-Z]{3}$/.test(code);
  } catch (error) {
    return false;
  }
}

function parseMoney(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

function isoFromParts(year, month, day) {
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return "";
  return `${year}-${pad2(month)}-${pad2(day)}`;
}

function todayIso() {
  const now = new Date();
  return isoFromParts(now.getFullYear(), now.getMonth() + 1, now.getDate());
}

function daysAgoIso(days) {
  const date = new Date();
  date.setDate(date.getDate() - days);
  return isoFromParts(date.getFullYear(), date.getMonth() + 1, date.getDate());
}

function formatDate(value) {
  if (isBlank(value)) return "—";
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return String(value);
  return new Intl.DateTimeFormat("en-US", { year: "numeric", month: "short", day: "numeric" }).format(
    new Date(`${value}T12:00:00`),
  );
}

function formatDateTime(value) {
  if (!value) return "—";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "—";
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  }).format(date);
}

function makeId(prefix) {
  if (globalThis.crypto?.randomUUID) return `${prefix}-${globalThis.crypto.randomUUID()}`;
  return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function uniqueValues(collection, key) {
  return [...new Set(collection.map((item) => item[key]).filter((value) => !isBlank(value)))].sort();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function badge(label, tone = "blue") {
  return `<span class="badge ${tone}">${escapeHtml(label)}</span>`;
}

function tracksReorderLevel() {
  return Boolean(findField("inventory", "reorderLevel"));
}

function isLowStock(item) {
  if (!tracksReorderLevel()) return false;
  return parseMoney(item.quantity) <= parseMoney(item.reorderLevel);
}

function stockBadge(item) {
  if (!tracksReorderLevel()) {
    return parseMoney(item.quantity) > 0 ? badge("In stock", "green") : badge("Sold / out", "blue");
  }
  if (isLowStock(item)) return badge("Low", "red");
  if (parseMoney(item.quantity) <= parseMoney(item.reorderLevel) * 1.5) return badge("Watch", "amber");
  return badge("Healthy", "green");
}

function statusBadge(status) {
  const tones = {
    Active: "green",
    Paid: "green",
    Pending: "amber",
    Overdue: "red",
    Blocked: "red",
    Inactive: "blue",
    "On Leave": "amber",
    Maintenance: "amber",
    Damaged: "red",
    Retired: "blue",
  };
  return isBlank(status) ? "" : badge(status, tones[status] || "blue");
}

function sourceStatusBadge(status) {
  const tones = {
    Fresh: "green",
    Ready: "green",
    Stale: "amber",
    "Needs Review": "amber",
    Planned: "blue",
    Failed: "red",
    "Needs Auth": "red",
  };
  return badge(status || "Ready", tones[status] || "blue");
}

function recordLabel(module, record) {
  return record.name || record.description || record.code || record.sku || "Untitled record";
}

function titleCell(title, subtitle) {
  return `<div class="record-title"><strong>${escapeHtml(title || "—")}</strong>${
    subtitle ? `<span>${escapeHtml(subtitle)}</span>` : ""
  }</div>`;
}

function showToast(message) {
  els.toast.textContent = message;
  els.toast.classList.add("is-visible");
  window.clearTimeout(showToast.timer);
  showToast.timer = window.setTimeout(() => els.toast.classList.remove("is-visible"), 2600);
}

/* ---------- Auth screens ---------- */

function populateBusinessTypeSelects() {
  document.querySelectorAll("[data-business-types]").forEach((select) => {
    select.innerHTML = Object.keys(businessScopes)
      .map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`)
      .join("");
  });
}

function showAuth() {
  els.appShell.hidden = true;
  els.authScreen.hidden = false;
  els.modalBackdrop.hidden = true;
  const registered = isRegistered();
  els.authScreen.classList.toggle("is-login", registered);
  els.registerForm.hidden = registered;
  els.loginForm.hidden = !registered;
  els.registerError.hidden = true;
  els.loginError.hidden = true;

  if (registered) {
    els.loginTitle.textContent = `Sign in to ${state.organization.name}`;
    els.loginForm.password.value = "";
    return;
  }

  const org = state.organization;
  byId("registerDateFormat").value = org?.dateFormat || defaultDateFormat();
  if (org) {
    els.registerIntro.textContent =
      "Records from the earlier demo were found. Create the owner login to keep using them.";
    byId("registerBusinessName").value = org.name || "";
    byId("registerBusinessType").value = businessScopes[org.type] ? org.type : "General Business";
    byId("registerCurrency").value = org.currency || "USD";
    byId("registerOwnerName").value = org.ownerName || "";
    byId("registerOwnerEmail").value = org.ownerEmail || "";
  }
}

function showApp() {
  els.authScreen.hidden = true;
  els.appShell.hidden = false;
  setView(canOpenView(currentView) ? currentView : "dashboard");
}

function showAuthError(element, message) {
  element.textContent = message;
  element.hidden = false;
}

async function handleRegister(event) {
  event.preventDefault();
  const form = new FormData(els.registerForm);
  const name = String(form.get("businessName") || "").trim();
  const type = String(form.get("businessType") || "General Business");
  const currency = String(form.get("currency") || "").trim().toUpperCase();
  const dateFormat = String(form.get("dateFormat") || "DMY");
  const ownerName = String(form.get("ownerName") || "").trim();
  const ownerEmail = String(form.get("ownerEmail") || "").trim().toLowerCase();
  const password = String(form.get("password") || "");
  const passwordConfirm = String(form.get("passwordConfirm") || "");

  if (!name || !ownerName || !ownerEmail) return showAuthError(els.registerError, "Fill in all required fields.");
  if (!isValidCurrency(currency)) {
    return showAuthError(els.registerError, "Currency must be a 3-letter code like PKR, USD, or AED.");
  }
  if (password.length < 6) return showAuthError(els.registerError, "Password must be at least 6 characters.");
  if (password !== passwordConfirm) return showAuthError(els.registerError, "The two passwords do not match.");
  if (state.employees.some((employee) => employee.loginEmail === ownerEmail)) {
    return showAuthError(els.registerError, "An employee already uses that login email.");
  }

  const existing = state.organization;
  const salt = makeSalt();
  state.organization = {
    ...(existing || {}),
    name,
    type,
    currency,
    dateFormat,
    ownerName,
    ownerEmail,
    ownerPasswordSalt: salt,
    ownerPasswordHash: await hashPassword(password, salt),
    createdAt: existing?.createdAt || new Date().toISOString(),
  };
  if (!existing) applyBusinessScope(type, { silent: true });
  addActivity(`${name} registered as ${type}`, "setup");
  saveState();
  els.registerForm.reset();
  startSession("owner");
  showToast(`Welcome to LedgerFlow, ${ownerName}`);
}

async function handleLogin(event) {
  event.preventDefault();
  const form = new FormData(els.loginForm);
  const email = String(form.get("email") || "").trim().toLowerCase();
  const password = String(form.get("password") || "");
  const org = state.organization;

  if (email === String(org.ownerEmail || "").toLowerCase()) {
    if (await verifyPassword(password, org.ownerPasswordSalt, org.ownerPasswordHash)) {
      startSession("owner");
      return;
    }
  }

  const employee = state.employees.find((entry) => entry.loginEmail && entry.loginEmail === email);
  if (employee?.passwordHash && (await verifyPassword(password, employee.passwordSalt, employee.passwordHash))) {
    if (employee.status === "Inactive") {
      showAuthError(els.loginError, "This login is blocked because the employee is marked inactive.");
      return;
    }
    if (!assignableRoles().some((role) => role.id === employee.accessRole)) {
      showAuthError(els.loginError, "This employee has no access role yet. Ask the owner to assign one.");
      return;
    }
    startSession(employee.id);
    addActivity(`${employee.name} signed in`, "employees");
    return;
  }

  showAuthError(els.loginError, "Email or password is incorrect.");
}

function resetBrowserData() {
  const confirmed = window.confirm(
    "Erase ALL business data, employees, and logins stored in this browser? This cannot be undone.",
  );
  if (!confirmed) return;
  try {
    localStorage.removeItem(STORAGE_KEY);
    OLD_STORAGE_KEYS.forEach((key) => localStorage.removeItem(key));
    localStorage.removeItem(SESSION_KEY);
  } catch (error) {
    console.warn("Could not clear storage", error);
  }
  window.location.reload();
}

/* ---------- Navigation and chrome ---------- */

function setView(view) {
  if (!canOpenView(view)) view = "dashboard";
  const enteringSetup = view === "setup" && currentView !== "setup";
  currentView = view;
  document.querySelectorAll(".view").forEach((node) => node.classList.remove("is-visible"));
  byId(`${view}View`).classList.add("is-visible");
  document.querySelectorAll(".nav-item").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.view === view);
  });
  document.body.classList.remove("menu-open");
  if (enteringSetup) fillSetupForm();
  window.scrollTo({ top: 0 });
  render();
}

function quickAddTarget() {
  if (MODULES.includes(currentView) && can(currentView, "add")) return currentView;
  if (MODULES.includes(currentView)) return null;
  return MODULES.find((module) => can(module, "add")) || null;
}

function addActivity(message, module) {
  const user = currentUser();
  state.activity.unshift({
    id: makeId("activity"),
    message,
    module,
    detail: viewTitle(module) || "System",
    by: user?.name || "",
    date: new Date().toISOString(),
  });
  state.activity = state.activity.slice(0, 40);
  saveState();
}

function canSeeActivity(entry) {
  const module = entry.module || Object.keys(titles).find((key) => viewTitle(key) === entry.detail);
  if (MODULES.includes(module)) return can(module, "view");
  if (module === "imports" || module === "dataSources") return canAccess("imports");
  return canAccess("settings");
}

function render() {
  if (!currentUser()) {
    logout();
    return;
  }
  if (!canOpenView(currentView)) {
    setView("dashboard");
    return;
  }
  renderChrome();
  populateFilters();
  ["finance", "inventory", "contacts", "employees", "assets", "setup"].forEach((view) => {
    if (canOpenView(view)) renderTabs(view);
  });
  if (canAccess("settings")) renderSetup();
  renderDashboard();
  if (can("finance", "view")) renderFinance();
  ["inventory", "contacts", "assets"].forEach((module) => {
    if (can(module, "view")) renderModule(module);
  });
  if (can("employees", "view")) renderEmployeesView();
  if (canOpenView("dataSources")) renderDataSources();
  if (canAccess("reports")) renderReports();
}

function renderChrome() {
  const user = currentUser();
  const role = currentRole();
  const org = state.organization;

  document.querySelectorAll(".nav-item").forEach((button) => {
    button.hidden = !canOpenView(button.dataset.view);
  });
  document.querySelectorAll("[data-view-link]").forEach((button) => {
    button.hidden = !canOpenView(button.dataset.viewLink);
  });
  document.querySelectorAll("[data-add]").forEach((button) => {
    button.hidden = !can(button.dataset.add, "add");
  });
  document.querySelectorAll("[data-finance-tab]").forEach((button) => {
    button.hidden = !canOpenView("finance");
  });
  MODULES.forEach((module) => {
    const label = document.querySelector(`.nav-item[data-view="${module}"] [data-nav-label]`);
    if (label) label.textContent = navTitle(module);
    byId(`${module}Title`).textContent = viewTitle(module);
    const addButton = document.querySelector(`#${module}View .page-actions [data-add]`);
    if (addButton) addButton.textContent = `Add ${recordNoun(module).toLowerCase()}`;
  });
  document.querySelectorAll("[data-nav-group]").forEach((group) => {
    let next = group.nextElementSibling;
    let anyVisible = false;
    while (next && !next.matches("[data-nav-group]")) {
      if (!next.hidden) anyVisible = true;
      next = next.nextElementSibling;
    }
    group.hidden = !anyVisible;
  });
  document.title = `${viewTitle(currentView)} · ${org.name}`;

  const alerts = alertSummary();
  const empty = totalRecordCount() === 0;
  const health = byId("sidebarHealthBox");
  health.className = `health-pill ${empty ? "empty" : alerts.total > 5 ? "alert" : alerts.total ? "watch" : ""}`;
  byId("sidebarHealth").textContent = empty ? "No records yet" : alerts.total > 5 ? "Needs review" : alerts.total ? "Watch" : "All good";
  byId("sidebarHealthNote").textContent = empty
    ? "Add, import, or load sample data."
    : alerts.total
      ? `${alerts.total} thing${alerts.total === 1 ? "" : "s"} need attention`
      : "Nothing urgent right now";

  const setCount = (view, count) => {
    const node = document.querySelector(`[data-nav-count="${view}"]`);
    if (!node) return;
    node.hidden = !count;
    node.textContent = count > 99 ? "99+" : String(count);
  };
  const dues = can("finance", "view") ? buildDues() : { receivables: [], payables: [] };
  setCount("finance", [...dues.receivables, ...dues.payables].filter((item) => item.overdue && item.kind !== "salary").length);
  setCount("inventory", can("inventory", "view") ? state.inventory.filter(isLowStock).length : 0);
  setCount(
    "employees",
    can("employees", "view") && canAccess("sensitiveNumbers")
      ? state.employees.filter((employee) => salaryMonthsDue(employee).some((month) => month < todayIso().slice(0, 7))).length
      : 0,
  );

  els.sidebarBusinessName.textContent = org.name;
  byId("sidebarBusinessType").textContent = org.type;
  els.topbarBusiness.textContent = "";
  els.userName.textContent = user.name;
  els.userAvatar.textContent = (user.name || "?").trim().charAt(0).toUpperCase();
  els.userRole.textContent = user.isOwner
    ? previewRoleId
      ? `Owner · viewing as ${role.name}`
      : "Owner"
    : role.name;

  els.previewSwitch.hidden = !user.isOwner;
  if (user.isOwner) {
    els.previewRoleSelect.innerHTML = [
      `<option value="">Owner (me)</option>`,
      ...assignableRoles().map((entry) => `<option value="${escapeHtml(entry.id)}">${escapeHtml(entry.name)}</option>`),
    ].join("");
    els.previewRoleSelect.value = previewRoleId || "";
  }
  els.previewBanner.hidden = !previewRoleId;
  if (previewRoleId) {
    els.previewBannerText.textContent = `You are seeing exactly what someone with the "${role.name}" role sees.`;
  }

  els.exportBtn.hidden = !canAccess("export") || totalRecordCount() === 0;
  const target = quickAddTarget();
  // Section pages have their own add button in the page header.
  els.quickAddBtn.hidden = !target || MODULES.includes(currentView);
  if (target) els.quickAddBtn.textContent = `Add ${recordNoun(target)}`;
  els.clearActivityBtn.hidden = !canAccess("settings");
}

function populateFilters() {
  fillSelect("#inventoryCategoryFilter", "All", uniqueValues(state.inventory, "category"));
  byId("inventoryCategoryLabel").textContent = fieldLabel("inventory", "category");
  fillSelect("#employeeDepartmentFilter", "All", uniqueValues(state.employees, "department"));
  const financeTab = tabState.finance;
  const typeForTab = financeTab === "income" ? "Income" : financeTab === "expenses" ? "Expense" : "";
  fillSelect(
    "#financeCategoryFilter",
    "All",
    uniqueValues(
      state.finance.filter((entry) => !typeForTab || entry.type === typeForTab),
      "category",
    ),
  );

  const monthSelect = byId("financeMonthFilter");
  const wanted = financeViewState.monthFilter || monthSelect.value || "all";
  const months = [...new Set(state.finance.map((entry) => monthKeyOf(entry.date)).filter((key) => /^\d{4}-\d{2}$/.test(key)))]
    .sort()
    .reverse();
  const years = [...new Set(months.map((key) => key.slice(0, 4)))];
  const options = [
    `<option value="all">All time</option>`,
    ...years.map((year) => `<option value="${year}">Whole of ${year}</option>`),
    ...months.map((key) => `<option value="${key}">${monthName(key, "long")}</option>`),
  ];
  monthSelect.innerHTML = options.join("");
  monthSelect.value = wanted === "all" || years.includes(wanted) || months.includes(wanted) ? wanted : "all";
  financeViewState.monthFilter = monthSelect.value;
}

function fillSelect(selector, label, values) {
  const select = document.querySelector(selector);
  const current = select.value || "all";
  select.innerHTML = `<option value="all">${escapeHtml(label)}</option>${values
    .map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`)
    .join("")}`;
  select.value = values.includes(current) ? current : "all";
}

/* ---------- Setup ---------- */

function fillSetupForm() {
  const org = state.organization;
  if (!org) return;
  els.businessNameInput.value = org.name || "";
  els.businessTypeSelect.value = businessScopes[org.type] ? org.type : "General Business";
  els.currencyInput.value = org.currency || "USD";
  els.dateFormatSelect.value = org.dateFormat || "DMY";
  els.ownerNameInput.value = org.ownerName || "";
  els.ownerEmailInput.value = org.ownerEmail || "";
  els.ownerPasswordInput.value = "";
}

function renderSetup() {
  const isOwner = Boolean(currentUser()?.isOwner);
  els.ownerFields.hidden = !isOwner;
  els.ownerFields.querySelectorAll("input").forEach((input) => {
    input.disabled = !isOwner;
  });
  els.dataToolsPanel.hidden = !isOwner || Boolean(previewRoleId);
  renderSetupChecklist();
  renderRoleBuilder();
  renderRoleCards();
  renderFieldList();
  renderBusinessScopes();
}

function launchSteps() {
  const org = state.organization;
  const loginCount = state.employees.filter(employeeCanLogin).length;
  const fieldsCustomized =
    MODULES.some((module) => state.customFields[module].length) ||
    MODULES.some((module) => Object.keys(state.fieldSettings[module] || {}).length);
  return [
    {
      label: "Business registered",
      done: true,
      detail: `${org.name} · ${org.type} · ${org.currency}`,
    },
    {
      label: "Fields fit your business",
      done: fieldsCustomized,
      detail: "Rename, hide, or add fields for items, finance, employees, and more.",
      view: "setup",
      action: "Adjust fields",
    },
    {
      label: "Roles decide what staff can see",
      done: assignableRoles().length > 0,
      detail: `${assignableRoles().length} staff role${assignableRoles().length === 1 ? "" : "s"} ready to assign.`,
      view: "setup",
      action: "Review roles",
    },
    {
      label: "Employees can sign in",
      done: loginCount > 0,
      detail: state.employees.length
        ? `${loginCount} of ${state.employees.length} employees have a working login.`
        : "Add employees with a login email and password.",
      add: "employees",
      action: "Add employee",
    },
    {
      label: "Business records added",
      done: ["inventory", "finance", "contacts", "assets"].some((module) => state[module].length),
      detail: "Import your existing files, or start adding records by hand.",
      view: "dataSources",
      action: "Import data",
      recordsStep: true,
    },
  ];
}

function renderSetupChecklist() {
  byId("setupChecklist").innerHTML = launchSteps()
    .map(
      (step) => `
      <div class="checklist-item">
        <span class="check-status ${step.done ? "" : "warn"}">${step.done ? "✓" : "!"}</span>
        <div>
          <strong>${escapeHtml(step.label)}</strong>
          <span>${escapeHtml(step.detail)}</span>
        </div>
      </div>
    `,
    )
    .join("");
}

function renderRoleBuilder() {
  const current = els.newRoleCopyFrom.value;
  els.newRoleCopyFrom.innerHTML = [
    `<option value="">No access (choose below)</option>`,
    ...assignableRoles().map((role) => `<option value="${escapeHtml(role.id)}">Copy ${escapeHtml(role.name)}</option>`),
  ].join("");
  els.newRoleCopyFrom.value = assignableRoles().some((role) => role.id === current) ? current : "";
}

function renderRoleCards() {
  byId("roleCards").innerHTML = state.roles
    .map((role) => {
      const locked = role.id === "owner";
      const assigned = state.employees.filter((employee) => employee.accessRole === role.id).length;
      const matrix = MODULES.map(
        (module) => `
          <tr>
            <td>${escapeHtml(viewTitle(module))}</td>
            ${moduleActions
              .map(
                (action) => `
              <td class="check-cell">
                <input
                  type="checkbox"
                  aria-label="${escapeHtml(`${role.name}: ${action.label} ${viewTitle(module)}`)}"
                  data-role-module="${escapeHtml(role.id)}:${module}:${action.key}"
                  ${role.permissions.modules[module][action.key] ? "checked" : ""}
                  ${locked ? "disabled" : ""}
                />
              </td>
            `,
              )
              .join("")}
          </tr>
        `,
      ).join("");
      return `
      <article class="role-card">
        <div class="role-card-header">
          <div>
            <strong>${escapeHtml(role.name)}</strong>
            <p>${escapeHtml(role.description)}</p>
          </div>
          <div class="role-card-actions">
            ${locked ? badge("Full access", "blue") : badge(`${assigned} employee${assigned === 1 ? "" : "s"}`, "green")}
            ${
              locked
                ? ""
                : `<button class="text-button" data-preview-role="${escapeHtml(role.id)}" type="button">Preview</button>
                   <button class="text-button danger" data-remove-role="${escapeHtml(role.id)}" type="button">Remove</button>`
            }
          </div>
        </div>
        <div class="matrix-wrap">
          <table class="permission-matrix">
            <thead>
              <tr><th>Module</th>${moduleActions.map((action) => `<th>${action.label}</th>`).join("")}</tr>
            </thead>
            <tbody>${matrix}</tbody>
          </table>
        </div>
        <div class="permission-grid">
          ${generalPermissions
            .map(
              (permission) => `
            <label class="permission-toggle">
              <input
                type="checkbox"
                data-role-general="${escapeHtml(role.id)}:${permission.key}"
                ${role.permissions[permission.key] ? "checked" : ""}
                ${locked ? "disabled" : ""}
              />
              <span>${escapeHtml(permission.label)}</span>
            </label>
          `,
            )
            .join("")}
        </div>
      </article>
    `;
    })
    .join("");
}

function addRole() {
  const name = els.newRoleName.value.trim();
  if (!name) {
    showToast("Enter a role name");
    return;
  }
  if (state.roles.some((role) => role.name.toLowerCase() === name.toLowerCase())) {
    showToast("A role with that name already exists");
    return;
  }
  const copyFrom = assignableRoles().find((role) => role.id === els.newRoleCopyFrom.value);
  state.roles.push({
    id: makeId("role"),
    name,
    description: els.newRoleDescription.value.trim() || "Custom access role.",
    system: false,
    permissions: copyFrom ? structuredClone(copyFrom.permissions) : blankPermissions(),
  });
  els.newRoleName.value = "";
  els.newRoleDescription.value = "";
  addActivity(`Role created: ${name}`, "setup");
  saveState();
  render();
  showToast(`${name} role created. Tick what it can access below.`);
}

function removeRole(roleId) {
  const role = assignableRoles().find((entry) => entry.id === roleId);
  if (!role) return;
  const assigned = state.employees.filter((employee) => employee.accessRole === roleId).length;
  if (assigned) {
    showToast(`Move ${assigned} employee${assigned === 1 ? "" : "s"} to another role first`);
    return;
  }
  if (!window.confirm(`Remove the ${role.name} role?`)) return;
  state.roles = state.roles.filter((entry) => entry.id !== roleId);
  if (previewRoleId === roleId) previewRoleId = null;
  addActivity(`Role removed: ${role.name}`, "setup");
  saveState();
  render();
  showToast("Role removed");
}

function toggleModulePermission(input) {
  const [roleId, module, action] = input.dataset.roleModule.split(":");
  const role = assignableRoles().find((entry) => entry.id === roleId);
  if (!role) return;
  const access = role.permissions.modules[module];
  access[action] = input.checked;
  if (action === "view" && !input.checked) {
    access.add = false;
    access.edit = false;
    access.delete = false;
  }
  if (action !== "view" && input.checked) access.view = true;
  saveState();
  render();
  showToast(`${role.name} access updated`);
}

function toggleGeneralPermission(input) {
  const [roleId, key] = input.dataset.roleGeneral.split(":");
  const role = assignableRoles().find((entry) => entry.id === roleId);
  if (!role) return;
  role.permissions[key] = input.checked;
  saveState();
  render();
  showToast(`${role.name} access updated`);
}

function renderFieldList() {
  const module = els.fieldModuleSelect.value;
  const fields = getFields(module, { includeHidden: true });
  els.fieldList.innerHTML = fields
    .map((field) => {
      const meta = [
        field.custom ? "Your field" : "Base field",
        { text: "Text", number: "Number", date: "Date", select: "Choice list", email: "Email", password: "Password", textarea: "Long text" }[
          field.type
        ] || field.type,
        field.required ? "Required" : "",
        field.sensitive ? "Sensitive" : "",
        field.settingsOnly ? "Setup access only" : "",
      ]
        .filter(Boolean)
        .join(" · ");
      const isEditing = editingField?.module === module && editingField?.name === field.name;
      return `
      <div class="custom-field-row ${field.hidden ? "is-hidden-field" : ""} ${isEditing ? "is-editing" : ""}">
        <div>
          <strong>${escapeHtml(field.label)} ${field.hidden ? badge("Hidden", "amber") : ""}</strong>
          <span>${escapeHtml(meta)}${field.type === "select" && field.options?.length ? ` · ${escapeHtml(field.options.join(", "))}` : ""}</span>
        </div>
        <div class="row-actions">
          <button type="button" data-edit-field="${module}:${escapeHtml(field.name)}">Edit</button>
          ${
            !field.custom && !field.locked
              ? `<button type="button" data-toggle-field="${module}:${escapeHtml(field.name)}">${field.hidden ? "Show" : "Hide"}</button>`
              : ""
          }
          ${field.custom ? `<button type="button" data-remove-field="${module}:${escapeHtml(field.name)}">Remove</button>` : ""}
        </div>
      </div>
    `;
    })
    .join("");
}

function resetFieldEditor() {
  editingField = null;
  els.fieldEditor.reset();
  els.fieldEditorTitle.textContent = "Add a field";
  els.saveFieldBtn.textContent = "Add Field";
  els.cancelFieldEditBtn.hidden = true;
  els.fieldTypeInput.disabled = false;
  els.fieldRequiredInput.disabled = false;
  els.fieldHiddenControl.hidden = true;
  els.fieldEditorNote.hidden = true;
  syncFieldOptionsControl();
}

function syncFieldOptionsControl() {
  els.fieldOptionsControl.hidden = els.fieldTypeInput.value !== "select" || els.fieldTypeInput.disabled;
}

function startFieldEdit(module, name) {
  const field = getFields(module, { includeHidden: true }).find((entry) => entry.name === name);
  if (!field) return;
  editingField = { module, name, custom: Boolean(field.custom) };
  els.fieldEditorTitle.textContent = `Edit "${field.label}"`;
  els.saveFieldBtn.textContent = "Save Field";
  els.cancelFieldEditBtn.hidden = false;
  els.fieldLabelInput.value = field.label;
  const editableType = field.custom;
  els.fieldTypeInput.value = ["text", "number", "date", "select"].includes(field.type) ? field.type : "text";
  els.fieldTypeInput.disabled = !editableType;
  els.fieldOptionsInput.value = (field.options || []).join(", ");
  els.fieldRequiredInput.checked = Boolean(field.required);
  els.fieldRequiredInput.disabled = Boolean(field.locked);
  els.fieldHiddenControl.hidden = Boolean(field.custom || field.locked);
  els.fieldHiddenInput.checked = Boolean(field.hidden);
  els.fieldEditorNote.hidden = field.custom;
  els.fieldEditorNote.textContent = field.locked
    ? "LedgerFlow uses this field in its calculations, so you can rename it but not hide it or change whether it's required."
    : "Base field: you can rename it, make it required, or hide it.";
  syncFieldOptionsControl();
  renderFieldList();
  els.fieldLabelInput.focus();
}

function saveFieldFromEditor(event) {
  event.preventDefault();
  const module = editingField?.module || els.fieldModuleSelect.value;
  const label = els.fieldLabelInput.value.trim();
  const type = els.fieldTypeInput.value;
  const options = els.fieldOptionsInput.value
    .split(",")
    .map((option) => option.trim())
    .filter(Boolean);
  const required = els.fieldRequiredInput.checked;

  if (!label) {
    showToast("Enter a field name");
    return;
  }
  const clash = getFields(module, { includeHidden: true }).find(
    (field) => field.label.toLowerCase() === label.toLowerCase() && field.name !== editingField?.name,
  );
  if (clash) {
    showToast("Another field already has that name");
    return;
  }

  if (editingField && !editingField.custom) {
    const base = schemas[module].fields.find((field) => field.name === editingField.name);
    const overrides = state.fieldSettings[module];
    overrides[base.name] = {
      ...(overrides[base.name] || {}),
      label: label === base.label ? undefined : label,
      ...(base.locked ? {} : { required, hidden: els.fieldHiddenInput.checked }),
    };
    addActivity(`Field updated: ${label} (${viewTitle(module)})`, "setup");
  } else {
    if (type === "select" && !options.length) {
      showToast("Add at least one choice, separated by commas");
      return;
    }
    if (editingField) {
      const field = state.customFields[module].find((entry) => entry.name === editingField.name);
      Object.assign(field, { label, type, options: type === "select" ? options : [], required });
      addActivity(`Field updated: ${label} (${viewTitle(module)})`, "setup");
    } else {
      state.customFields[module].push({
        name: uniqueFieldName(module, label),
        label,
        type,
        options: type === "select" ? options : [],
        required,
      });
      addActivity(`Field added: ${label} (${viewTitle(module)})`, "setup");
    }
  }

  const wasEditing = Boolean(editingField);
  resetFieldEditor();
  saveState();
  render();
  showToast(wasEditing ? "Field saved" : "Field added");
}

function toggleFieldHidden(module, name) {
  const base = schemas[module].fields.find((field) => field.name === name);
  if (!base || base.locked) return;
  const overrides = state.fieldSettings[module];
  const hidden = !overrides[name]?.hidden;
  overrides[name] = { ...(overrides[name] || {}), hidden };
  addActivity(`${fieldLabel(module, name)} ${hidden ? "hidden" : "shown"} in ${viewTitle(module)}`, "setup");
  saveState();
  render();
}

function removeCustomField(module, name) {
  const field = state.customFields[module].find((entry) => entry.name === name);
  if (!field) return;
  if (!window.confirm(`Remove the "${field.label}" field? Values already saved stay in the records but won't be shown.`)) {
    return;
  }
  state.customFields[module] = state.customFields[module].filter((entry) => entry.name !== name);
  if (editingField?.name === name) resetFieldEditor();
  addActivity(`Field removed: ${field.label} (${viewTitle(module)})`, "setup");
  saveState();
  render();
  showToast("Field removed");
}

function renderBusinessScopes() {
  byId("industryScopeCards").innerHTML = Object.entries(businessScopes)
    .map(
      ([name, scope]) => `
      <article class="scope-card ${state.organization.type === name ? "is-current" : ""}">
        <div>
          <strong>${escapeHtml(name)} ${state.organization.type === name ? badge("Your type", "green") : ""}</strong>
          <p>${escapeHtml(scope.summary)}</p>
        </div>
        <ul>
          ${Object.entries(scope.fields)
            .map(
              ([module, fields]) =>
                `<li>${escapeHtml(scope.labels?.[module] || titles[module])}: ${escapeHtml(
                  fields.map((spec) => spec.split("|")[0]).join(", "),
                )}</li>`,
            )
            .join("")}
        </ul>
        <button class="button ghost" data-apply-scope="${escapeHtml(name)}" type="button">Apply Template</button>
      </article>
    `,
    )
    .join("");
}

function applyBusinessScope(scopeName, { silent = false, replace = false } = {}) {
  const scope = businessScopes[scopeName];
  if (!scope) return 0;
  const org = state.organization;
  org.type = scopeName;
  if (replace) {
    state.customFields = Object.fromEntries(MODULES.map((module) => [module, []]));
    state.fieldSettings = Object.fromEntries(MODULES.map((module) => [module, {}]));
  } else {
    // Base-field names and hidden fields belong to the previous business type.
    MODULES.forEach((module) => {
      Object.values(state.fieldSettings[module]).forEach((override) => {
        delete override.label;
        delete override.hidden;
      });
    });
  }
  org.moduleLabels = { ...(scope.labels || {}) };
  org.recordNouns = { ...(scope.nouns || {}) };
  Object.entries(scope.fieldLabels || {}).forEach(([module, labels]) => {
    Object.entries(labels).forEach(([name, label]) => {
      state.fieldSettings[module][name] = { ...(state.fieldSettings[module][name] || {}), label };
    });
  });
  Object.entries(scope.hide || {}).forEach(([module, names]) => {
    names.forEach((name) => {
      state.fieldSettings[module][name] = { ...(state.fieldSettings[module][name] || {}), hidden: true };
    });
  });
  let added = 0;
  Object.entries(scope.fields).forEach(([module, specs]) => {
    specs.forEach((spec) => {
      const [label, type, optionText] = spec.split("|");
      const exists = getFields(module, { includeHidden: true }).some(
        (field) => field.label.toLowerCase() === label.toLowerCase(),
      );
      if (exists) return;
      state.customFields[module].push({
        name: uniqueFieldName(module, label),
        label,
        type: type || inferCustomFieldType(label),
        options: optionText ? optionText.split(",") : [],
        required: false,
      });
      added += 1;
    });
  });
  if (silent) return added;
  addActivity(`${scopeName} template applied (${added} field${added === 1 ? "" : "s"} added)`, "setup");
  saveState();
  render();
  showToast(added ? `${added} field${added === 1 ? "" : "s"} added for ${scopeName}` : "You already have these fields");
  return added;
}

async function saveSetup(event) {
  event.preventDefault();
  const org = state.organization;
  const isOwner = Boolean(currentUser()?.isOwner);
  const name = els.businessNameInput.value.trim();
  const currency = els.currencyInput.value.trim().toUpperCase();
  if (!name) {
    showToast("Business name is required");
    return;
  }
  if (!isValidCurrency(currency)) {
    showToast("Currency must be a 3-letter code like PKR or USD");
    return;
  }

  const updates = {
    name,
    type: els.businessTypeSelect.value,
    currency,
    dateFormat: els.dateFormatSelect.value,
  };

  if (isOwner) {
    const ownerEmail = els.ownerEmailInput.value.trim().toLowerCase();
    const password = els.ownerPasswordInput.value;
    if (!ownerEmail || !els.ownerNameInput.value.trim()) {
      showToast("Owner name and email are required");
      return;
    }
    if (state.employees.some((employee) => employee.loginEmail === ownerEmail)) {
      showToast("An employee already uses that login email");
      return;
    }
    if (password && password.length < 6) {
      showToast("Password must be at least 6 characters");
      return;
    }
    updates.ownerName = els.ownerNameInput.value.trim();
    updates.ownerEmail = ownerEmail;
    if (password) {
      updates.ownerPasswordSalt = makeSalt();
      updates.ownerPasswordHash = await hashPassword(password, updates.ownerPasswordSalt);
    }
  }

  const previousType = org.type;
  const previousCurrency = org.currency;
  Object.assign(org, updates);
  if (
    updates.type !== previousType &&
    window.confirm(`Also arrange module and field names for a ${updates.type}? Your records stay.`)
  ) {
    applyBusinessScope(updates.type, { silent: true });
  }
  addActivity(`Business profile updated`, "setup");
  saveState();
  fillSetupForm();
  render();
  showToast(
    previousCurrency !== currency
      ? `Profile saved. New records will default to ${currency}; existing records keep their own currency.`
      : "Profile saved",
  );
}

/* ---------- Dashboard ---------- */

const ICONS = {
  income: '<svg viewBox="0 0 24 24"><path d="M17 7 7 17M7 8v9h9"/></svg>',
  expense: '<svg viewBox="0 0 24 24"><path d="M7 17 17 7M8 7h9v9"/></svg>',
  profit: '<svg viewBox="0 0 24 24"><path d="m3 17 6-6 4 4 8-8M14 7h7v7"/></svg>',
  inbox: '<svg viewBox="0 0 24 24"><path d="M4 13h4l2 3h4l2-3h4M5 5h14l2 8v6H3v-6z"/></svg>',
  clock: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>',
  alert: '<svg viewBox="0 0 24 24"><path d="M12 3 2 20h20z"/><path d="M12 10v4M12 17h.01"/></svg>',
  box: '<svg viewBox="0 0 24 24"><path d="M21 8 12 3 3 8v8l9 5 9-5z"/><path d="m3 8 9 5 9-5M12 13v8"/></svg>',
  users: '<svg viewBox="0 0 24 24"><circle cx="9" cy="8" r="3.5"/><path d="M2.5 20a6.5 6.5 0 0 1 13 0M16 4.5a3.5 3.5 0 0 1 0 7M18 14a5.5 5.5 0 0 1 3.5 6"/></svg>',
  info: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 8h.01"/></svg>',
  wallet: '<svg viewBox="0 0 24 24"><rect x="3" y="6" width="18" height="14" rx="2"/><path d="M3 10h18M16 15h2"/></svg>',
  tool: '<svg viewBox="0 0 24 24"><path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18l3 3 6.3-6.3a4 4 0 0 0 5.4-5.4l-2.5 2.5-2.5-2.5z"/></svg>',
};

function kpiCard({ label, value, icon = "", tone = "", extra = "", note = "", negative = false }) {
  return `
    <article class="kpi-card">
      <div class="kpi-top">
        <span class="kpi-label">${escapeHtml(label)}</span>
        ${icon ? `<span class="kpi-icon ${tone}">${ICONS[icon]}</span>` : ""}
      </div>
      <strong class="kpi-value ${negative ? "negative" : ""}">${escapeHtml(value)}</strong>
      ${extra ? `<span class="kpi-extra">${extra}</span>` : ""}
      ${note ? `<span class="kpi-note">${escapeHtml(note)}</span>` : ""}
    </article>
  `;
}

function monthKeyOf(date) {
  return String(date || "").slice(0, 7);
}

function shiftMonth(key, delta) {
  const [year, month] = key.split("-").map(Number);
  const date = new Date(year, month - 1 + delta, 1);
  return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}`;
}

function monthName(key, style = "short") {
  const [year, month] = key.split("-").map(Number);
  const options = style === "long" ? { month: "long", year: "numeric" } : { month: "short" };
  return new Intl.DateTimeFormat("en-US", options).format(new Date(year, month - 1, 1));
}

function lastDayOfMonth(key) {
  const [year, month] = key.split("-").map(Number);
  return isoFromParts(year, month, new Date(year, month, 0).getDate());
}

function round2(value) {
  return Math.round(value * 100) / 100;
}

function financeTotalsWhere(currency, predicate) {
  return state.finance.reduce(
    (totals, entry) => {
      if (recordCurrency(entry) !== currency || !predicate(entry)) return totals;
      const amount = parseMoney(entry.amount);
      if (entry.type === "Income") totals.income += amount;
      if (entry.type === "Expense") totals.expense += amount;
      totals.count += 1;
      return totals;
    },
    { income: 0, expense: 0, count: 0 },
  );
}

function financeStatus(entry) {
  if (isBlank(entry.status) || entry.status === "Paid") return entry.status || "";
  if (entry.status === "Overdue" || (entry.dueDate && entry.dueDate < todayIso())) return "Overdue";
  return "Pending";
}

function isUnpaid(entry) {
  const status = financeStatus(entry);
  return status === "Pending" || status === "Overdue";
}

function totals() {
  const income = sumByCurrency(
    state.finance.filter((entry) => entry.type === "Income"),
    (entry) => parseMoney(entry.amount),
  );
  const expense = sumByCurrency(
    state.finance.filter((entry) => entry.type === "Expense"),
    (entry) => parseMoney(entry.amount),
  );
  const inventoryValue = sumByCurrency(state.inventory, (item) => parseMoney(item.quantity) * parseMoney(item.unitCost));
  const assetValue = sumByCurrency(state.assets, (asset) => parseMoney(asset.value));
  const payroll = sumByCurrency(
    state.employees.filter((employee) => employee.status !== "Inactive"),
    (employee) => parseMoney(employee.salary),
  );
  const dues = buildDues();
  return {
    income,
    expense,
    net: subtractByCurrency(income, expense),
    inventoryValue,
    assetValue,
    receivables: duesTotals(dues.receivables),
    payables: duesTotals(dues.payables),
    payroll,
  };
}

function alertSummary() {
  const lowStock = can("inventory", "view") ? state.inventory.filter(isLowStock).length : 0;
  const maintenance = can("assets", "view") ? state.assets.filter(needsMaintenance).length : 0;
  const unpaid = can("finance", "view") ? state.finance.filter((entry) => financeStatus(entry) === "Overdue").length : 0;
  const missingInfo = MODULES.filter((module) => can(module, "view")).reduce(
    (count, module) => count + state[module].filter((record) => needsInfo(module, record)).length,
    0,
  );
  return { lowStock, maintenance, unpaid, missingInfo, total: lowStock + maintenance + unpaid + missingInfo };
}

function needsMaintenance(asset) {
  return asset.condition === "Maintenance" || asset.condition === "Damaged";
}

/* Dues: money waiting to come in (receivables) or go out (payables). */

function salaryPayment(employee, month) {
  return state.finance.find((entry) => entry.employeeId === employee.id && entry.payrollMonth === month);
}

function payrollStartMonth(employee) {
  return monthKeyOf(employee.createdAt || employee.importedAt) || todayIso().slice(0, 7);
}

function salaryIsPayable(employee) {
  return employee.status !== "Inactive" && parseMoney(employee.salary) > 0;
}

function salaryMonthsDue(employee) {
  if (!salaryIsPayable(employee)) return [];
  const current = todayIso().slice(0, 7);
  const months = [];
  for (let key = payrollStartMonth(employee); key <= current && months.length < 24; key = shiftMonth(key, 1)) {
    if (!salaryPayment(employee, key)) months.push(key);
  }
  return months;
}

function sortDues(a, b) {
  return Number(b.overdue) - Number(a.overdue) || String(a.dueDate || "9999").localeCompare(String(b.dueDate || "9999"));
}

function buildDues() {
  const receivables = [];
  const payables = [];
  if (can("finance", "view")) {
    state.finance.filter(isUnpaid).forEach((entry) => {
      const item = {
        kind: "transaction",
        id: entry.id,
        party: entry.party || "",
        title: entry.description || entry.category || "Transaction",
        detail: entry.category || "",
        amount: parseMoney(entry.amount),
        currency: recordCurrency(entry),
        dueDate: entry.dueDate || "",
        date: entry.date || "",
        overdue: financeStatus(entry) === "Overdue",
      };
      (entry.type === "Income" ? receivables : payables).push(item);
    });
  }
  if (can("contacts", "view")) {
    state.contacts.forEach((contact) => {
      const balance = parseMoney(contact.balance);
      if (!balance) return;
      const item = {
        kind: "balance",
        id: contact.id,
        party: contact.name,
        title: "Opening balance",
        detail: contact.type || "",
        amount: Math.abs(balance),
        currency: recordCurrency(contact),
        dueDate: "",
        date: "",
        overdue: false,
      };
      (balance > 0 ? receivables : payables).push(item);
    });
  }
  if (can("employees", "view") && canAccess("sensitiveNumbers")) {
    const current = todayIso().slice(0, 7);
    state.employees.forEach((employee) => {
      salaryMonthsDue(employee).forEach((month) => {
        payables.push({
          kind: "salary",
          id: employee.id,
          month,
          party: employee.name,
          title: `Salary for ${monthName(month, "long")}`,
          detail: employee.role || "",
          amount: parseMoney(employee.salary),
          currency: recordCurrency(employee),
          dueDate: lastDayOfMonth(month),
          date: "",
          overdue: month < current,
        });
      });
    });
  }
  receivables.sort(sortDues);
  payables.sort(sortDues);
  return { receivables, payables };
}

function duesTotals(items, onlyOverdue = false) {
  return items
    .filter((item) => !onlyOverdue || item.overdue)
    .reduce((totalsByCurrency, item) => {
      totalsByCurrency[item.currency] = (totalsByCurrency[item.currency] || 0) + item.amount;
      return totalsByCurrency;
    }, {});
}

function dueItemHtml(item, direction) {
  const source =
    item.kind === "salary"
      ? badge("Salary", "violet")
      : item.kind === "balance"
        ? badge("Opening balance", "blue")
        : badge(direction === "in" ? "Unpaid income" : "Unpaid bill", "amber");
  const when = item.overdue
    ? badge(item.dueDate ? `Overdue · was due ${formatDate(item.dueDate)}` : "Overdue", "red")
    : item.dueDate
      ? `<span>Due ${escapeHtml(formatDate(item.dueDate))}</span>`
      : item.date
        ? `<span>Recorded ${escapeHtml(formatDate(item.date))}</span>`
        : "";
  let action = "";
  if (item.kind === "transaction" && can("finance", "edit")) {
    action = `<button type="button" class="accent" data-mark-paid="${escapeHtml(item.id)}">${direction === "in" ? "Mark received" : "Mark paid"}</button>`;
  } else if (item.kind === "balance" && can("finance", "add")) {
    action = `<button type="button" class="accent" data-record-payment="${escapeHtml(item.id)}">Record payment</button>`;
  } else if (item.kind === "salary" && can("finance", "add")) {
    action = `<button type="button" class="accent" data-pay-salary="${escapeHtml(item.id)}:${item.month}">Pay salary</button>`;
  }
  const subtitle = [item.party ? item.title : "", item.detail].filter(Boolean).join(" · ");
  return `
    <div class="due-item ${item.overdue ? "overdue" : ""}">
      <div>
        <strong>${escapeHtml(item.party || item.title)}</strong>
        ${subtitle ? `<span>${escapeHtml(subtitle)}</span>` : ""}
        <div class="due-meta">${source}${when}</div>
      </div>
      <div class="due-side">
        <span class="due-amount">${escapeHtml(formatCurrency(item.amount, item.currency))}</span>
        ${action ? `<div class="row-actions">${action}</div>` : ""}
      </div>
    </div>
  `;
}

function greeting() {
  const hour = new Date().getHours();
  if (hour < 12) return "Good morning";
  if (hour < 17) return "Good afternoon";
  return "Good evening";
}

function deltaHtml(now, before, previousLabel, lowerIsBetter = false) {
  if (!before) return "";
  const change = Math.round(((now - before) / before) * 100);
  const up = change >= 0;
  const good = lowerIsBetter ? !up : up;
  return `<span class="delta ${good ? "up" : "down"}">${up ? "▲" : "▼"} ${Math.abs(change)}%</span> vs ${escapeHtml(previousLabel)}`;
}

function renderDashboard() {
  const user = currentUser();
  const org = state.organization;
  const sensitive = canAccess("sensitiveNumbers");

  byId("welcomeKicker").textContent = previewRoleId
    ? `Previewing the ${currentRole().name} role`
    : new Intl.DateTimeFormat("en-US", { weekday: "long", day: "numeric", month: "long", year: "numeric" }).format(new Date());
  byId("welcomeTitle").textContent = `${greeting()}, ${(user.name || "").trim().split(" ")[0] || "there"}`;
  byId("businessScopeSummary").textContent = `${org.name} · ${org.type} · Signed in as ${currentRole().name}`;

  const empty = totalRecordCount() === 0;
  byId("sampleBanner").hidden = !(org.sampleLoaded && user.isOwner && !previewRoleId);
  if (org.sampleLoaded) {
    byId("sampleBannerText").textContent = `You are exploring sample data for a ${org.sampleLoaded}. None of these records are real.`;
  }
  byId("emptyStatePanel").hidden = !empty;
  byId("dashboardMetrics").hidden = empty;
  byId("dashPanels").hidden = empty;
  renderGettingStarted();
  if (empty) {
    renderEmptyState();
    return;
  }

  const dues = buildDues();
  const cards = [];
  const month = todayIso().slice(0, 7);
  const previous = shiftMonth(month, -1);

  if (can("finance", "view")) {
    const { currency } = chartCurrency(state.finance);
    const now = financeTotalsWhere(currency, (entry) => monthKeyOf(entry.date) === month);
    const before = financeTotalsWhere(currency, (entry) => monthKeyOf(entry.date) === previous);
    if (sensitive) {
      const profit = now.income - now.expense;
      cards.push(
        kpiCard({
          label: `Income · ${monthName(month)}`,
          value: formatCurrency(now.income, currency),
          icon: "income",
          tone: "income",
          extra: deltaHtml(now.income, before.income, monthName(previous)),
        }),
        kpiCard({
          label: `Expenses · ${monthName(month)}`,
          value: formatCurrency(now.expense, currency),
          icon: "expense",
          tone: "expense",
          extra: deltaHtml(now.expense, before.expense, monthName(previous), true),
        }),
        kpiCard({
          label: `Profit · ${monthName(month)}`,
          value: formatCurrency(profit, currency),
          icon: "profit",
          tone: "profit",
          negative: profit < 0,
          note: now.income ? `${Math.round((profit / now.income) * 100)}% margin` : "No income recorded yet this month",
        }),
      );
    } else {
      cards.push(
        kpiCard({
          label: `Transactions · ${monthName(month)}`,
          value: formatNumber(now.count),
          icon: "wallet",
        }),
      );
    }
  }

  if (can("finance", "view") || can("contacts", "view")) {
    const owedIn = moneyHeadline(duesTotals(dues.receivables));
    const owedOut = moneyHeadline(duesTotals(dues.payables));
    const overdueIn = dues.receivables.filter((item) => item.overdue).length;
    const overdueOut = dues.payables.filter((item) => item.overdue).length;
    cards.push(
      kpiCard({
        label: "Owed to you",
        value: owedIn.value,
        icon: "inbox",
        tone: "income",
        note: [owedIn.others, `${dues.receivables.length} open · ${overdueIn} overdue`].filter(Boolean).join(" · "),
      }),
      kpiCard({
        label: "You owe",
        value: owedOut.value,
        icon: "clock",
        tone: overdueOut ? "alert" : "expense",
        note: [owedOut.others, `${dues.payables.length} open · ${overdueOut} overdue`].filter(Boolean).join(" · "),
      }),
    );
  }

  if (!can("finance", "view") && can("inventory", "view")) {
    const low = state.inventory.filter(isLowStock).length;
    cards.push(
      kpiCard({
        label: viewTitle("inventory"),
        value: formatNumber(state.inventory.length),
        icon: "box",
        note: tracksReorderLevel() ? `${low} low on stock` : "",
      }),
    );
  }

  byId("dashboardMetrics").innerHTML = cards.join("");

  byId("cashFlowPanel").hidden = !(can("finance", "view") && sensitive);
  byId("duesPanel").hidden = !(can("finance", "view") || can("contacts", "view"));
  byId("inventoryMixPanel").hidden = !can("inventory", "view");
  byId("insightsPanel").hidden = !sensitive;

  renderTodo(dues);
  if (!byId("duesPanel").hidden) renderDuesSummary(dues);
  if (!byId("cashFlowPanel").hidden) renderCashFlowChart();
  if (!byId("inventoryMixPanel").hidden) renderInventoryChart();
  if (sensitive) renderOwnerInsights();
  renderActivity();
}

function renderEmptyState() {
  const user = currentUser();
  const addable = MODULES.filter((module) => can(module, "add"));
  const options = [];
  if (addable.length) {
    options.push(`
      <article class="start-option">
        <span class="start-number">1</span>
        <h3>Enter records by hand</h3>
        <p>Type in your ${escapeHtml(viewTitle("inventory").toLowerCase())}, transactions, customers, and staff one at a time.</p>
        <div class="start-actions">
          ${addable
            .map(
              (module) =>
                `<button class="button ghost small" data-add="${module}" type="button">Add ${escapeHtml(recordNoun(module).toLowerCase())}</button>`,
            )
            .join("")}
        </div>
      </article>
    `);
  }
  if (canOpenView("dataSources")) {
    options.push(`
      <article class="start-option">
        <span class="start-number">${options.length + 1}</span>
        <h3>Import your existing files</h3>
        <p>Upload CSV, Excel, or JSON. LedgerFlow matches your columns, asks what currency the amounts are in, and asks for anything missing.</p>
        <div class="start-actions">
          <button class="button primary" data-view-link="dataSources" type="button">Import a file</button>
        </div>
      </article>
    `);
  }
  if (user.isOwner && !previewRoleId) {
    options.push(`
      <article class="start-option">
        <span class="start-number">${options.length + 1}</span>
        <h3>Explore with a sample business</h3>
        <p>Fill LedgerFlow with realistic records to see how sections and fields arrange themselves. You can remove it any time.</p>
        <div class="start-actions">
          <select id="emptySampleSelect" data-sample-select aria-label="Sample business"></select>
          <button class="button ghost" data-load-sample="emptySampleSelect" type="button">Load sample data</button>
        </div>
      </article>
    `);
  }
  byId("emptyStateTitle").textContent = user.isOwner ? `${state.organization.name} is ready` : "No records yet";
  byId("emptyStateIntro").textContent = options.length
    ? "There are no records yet. Choose how to start. You can combine these at any time."
    : "The owner hasn't added any records yet.";
  byId("startOptions").innerHTML = options.join("");
  populateSampleSelects();
}

function populateSampleSelects() {
  document.querySelectorAll("[data-sample-select]").forEach((select) => {
    const current = select.value;
    select.innerHTML = SAMPLE_BUSINESSES.map(
      (name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`,
    ).join("");
    select.value = SAMPLE_BUSINESSES.includes(current)
      ? current
      : SAMPLE_BUSINESSES.includes(state.organization?.type)
        ? state.organization.type
        : SAMPLE_BUSINESSES[0];
  });
}

function renderGettingStarted() {
  const panel = byId("gettingStartedPanel");
  const steps = launchSteps().filter((step) => totalRecordCount() > 0 || !step.recordsStep);
  const show = canAccess("settings") && !state.organization.hideGettingStarted && steps.some((step) => !step.done);
  panel.hidden = !show;
  if (!show) return;
  byId("gettingStartedSteps").innerHTML = steps
    .map(
      (step) => `
      <div class="launch-step ${step.done ? "is-done" : ""}">
        <span class="check-status ${step.done ? "" : "todo"}">${step.done ? "✓" : ""}</span>
        <div>
          <strong>${escapeHtml(step.label)}</strong>
          <span>${escapeHtml(step.detail)}</span>
          ${
            !step.done && step.view && canOpenView(step.view)
              ? `<button class="link-button" data-view-link="${step.view}" type="button">${escapeHtml(step.action)} →</button>`
              : ""
          }
          ${
            !step.done && step.add && can(step.add, "add")
              ? `<button class="link-button" data-add="${step.add}" type="button">${escapeHtml(step.action)} →</button>`
              : ""
          }
        </div>
      </div>
    `,
    )
    .join("");
}

function renderTodo(dues) {
  const items = [];
  const overdue = [...dues.receivables, ...dues.payables].filter((item) => item.overdue && item.kind !== "salary");
  if (overdue.length && canOpenView("finance")) {
    items.push({
      icon: "alert",
      tone: "red",
      title: `${overdue.length} overdue payment${overdue.length === 1 ? "" : "s"}`,
      detail: overdue
        .slice(0, 3)
        .map((item) => item.party || item.title)
        .join(", "),
      action: `<button type="button" data-finance-tab="dues">Review</button>`,
    });
  }
  const unpaidSalaries = dues.payables.filter((item) => item.kind === "salary");
  if (unpaidSalaries.length) {
    items.push({
      icon: "users",
      tone: unpaidSalaries.some((item) => item.overdue) ? "red" : "",
      title: `${unpaidSalaries.length} unpaid salar${unpaidSalaries.length === 1 ? "y" : "ies"}`,
      detail: formatMoneyTotals(duesTotals(unpaidSalaries)),
      action: `<button type="button" data-open-tab="employees:payroll">Payroll</button>`,
    });
  }
  if (can("inventory", "view")) {
    const low = state.inventory.filter(isLowStock);
    if (low.length) {
      items.push({
        icon: "box",
        tone: "",
        title: `${low.length} low on stock`,
        detail: low
          .slice(0, 3)
          .map((item) => item.name)
          .join(", "),
        action: `<button type="button" data-open-tab="inventory:low">View</button>`,
      });
    }
  }
  if (can("assets", "view")) {
    const repairs = state.assets.filter(needsMaintenance);
    if (repairs.length) {
      items.push({
        icon: "tool",
        tone: "",
        title: `${repairs.length} ${viewTitle("assets").toLowerCase()} need maintenance`,
        detail: repairs
          .slice(0, 3)
          .map((asset) => asset.name)
          .join(", "),
        action: `<button type="button" data-open-tab="assets:maintenance">View</button>`,
      });
    }
  }
  const missing = MODULES.filter((module) => can(module, "view")).flatMap((module) =>
    state[module]
      .filter((record) => needsInfo(module, record))
      .map((record) => ({ module, record, fields: visibleMissingFields(module, record) })),
  );
  missing.slice(0, 5).forEach(({ module, record, fields }) => {
    const canEdit = can(module, "edit");
    const requiredMissing = fields.some((field) => field.required);
    items.push({
      icon: "info",
      tone: "blue",
      title: recordLabel(module, record),
      detail: `${viewTitle(module)} · missing ${missingLabels(module, fields).join(", ")}`,
      action: [
        canEdit ? `<button type="button" data-edit="${module}:${escapeHtml(record.id)}">Fill in</button>` : "",
        canEdit && !requiredMissing
          ? `<button type="button" data-dismiss-info="${module}:${escapeHtml(record.id)}">Skip</button>`
          : "",
      ].join(""),
    });
  });
  if (missing.length > 5) {
    items.push({
      icon: "info",
      tone: "blue",
      title: `${missing.length - 5} more records missing information`,
      detail: 'Open a section and use its "Needs info" tab',
      action: "",
    });
  }

  byId("needsInputList").innerHTML = items.length
    ? items
        .map(
          (item) => `
        <div class="todo-item">
          <span class="todo-dot ${item.tone}">${ICONS[item.icon]}</span>
          <div>
            <strong>${escapeHtml(item.title)}</strong>
            <span>${escapeHtml(item.detail)}</span>
          </div>
          ${item.action ? `<div class="row-actions">${item.action}</div>` : "<span></span>"}
        </div>
      `,
        )
        .join("")
    : `<div class="empty-state">You're all caught up.</div>`;
}

function renderDuesSummary(dues) {
  const owedIn = moneyHeadline(duesTotals(dues.receivables));
  const owedOut = moneyHeadline(duesTotals(dues.payables));
  const next = [
    ...dues.receivables.map((item) => ({ ...item, direction: "in" })),
    ...dues.payables.map((item) => ({ ...item, direction: "out" })),
  ]
    .sort(sortDues)
    .slice(0, 3);
  byId("duesSummary").innerHTML = `
    <div class="dues-summary">
      <div class="in">
        <span>Owed to you</span>
        <strong>${escapeHtml(owedIn.value)}</strong>
        ${owedIn.others ? `<span>${escapeHtml(owedIn.others)}</span>` : ""}
      </div>
      <div class="out">
        <span>You owe</span>
        <strong>${escapeHtml(owedOut.value)}</strong>
        ${owedOut.others ? `<span>${escapeHtml(owedOut.others)}</span>` : ""}
      </div>
    </div>
    <div class="due-list">
      ${next.length ? next.map((item) => dueItemHtml(item, item.direction)).join("") : `<div class="empty-state">No open dues.</div>`}
    </div>
  `;
}

function dismissInfo(module, id) {
  const record = state[module].find((entry) => entry.id === id);
  if (!record || !can(module, "edit")) return;
  record.infoDismissed = true;
  saveState();
  render();
  showToast("Continuing without that information");
}

function chartCurrency(records) {
  const counts = records.reduce((acc, record) => {
    const code = recordCurrency(record);
    acc[code] = (acc[code] || 0) + 1;
    return acc;
  }, {});
  const primary = orgCurrency();
  const currency = counts[primary]
    ? primary
    : Object.entries(counts).sort((a, b) => b[1] - a[1])[0]?.[0] || primary;
  return { currency, excluded: records.length - (counts[currency] || 0) };
}

function prepareCanvas(canvas) {
  const ratio = window.devicePixelRatio || 1;
  const rect = canvas.getBoundingClientRect();
  canvas.width = Math.max(1, Math.floor(rect.width * ratio));
  canvas.height = Math.max(1, Math.floor(rect.height * ratio));
  const ctx = canvas.getContext("2d");
  ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
  return ctx;
}

function drawRoundedBar(ctx, x, y, width, height, color) {
  if (height <= 0) return;
  const radius = Math.min(4, width / 2, height / 2);
  ctx.beginPath();
  ctx.moveTo(x, y + height);
  ctx.lineTo(x, y + radius);
  ctx.quadraticCurveTo(x, y, x + radius, y);
  ctx.lineTo(x + width - radius, y);
  ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
  ctx.lineTo(x + width, y + height);
  ctx.closePath();
  ctx.fillStyle = color;
  ctx.fill();
}

// Bars for income and expenses, a line for profit. All rows are in one currency.
function drawFinanceChart(canvas, rows) {
  const ctx = prepareCanvas(canvas);
  const width = canvas.clientWidth;
  const height = canvas.clientHeight;
  ctx.clearRect(0, 0, width, height);
  if (!rows.length || width < 10) return;

  const values = rows.flatMap((row) => [row.income, row.expense, row.income - row.expense]);
  const max = Math.max(1, ...values);
  const min = Math.min(0, ...values);
  const pad = { top: 10, right: 8, bottom: 26, left: 58 };
  const innerWidth = width - pad.left - pad.right;
  const innerHeight = height - pad.top - pad.bottom;
  const y = (value) => pad.top + innerHeight - ((value - min) / (max - min)) * innerHeight;
  const compact = new Intl.NumberFormat("en-US", { notation: "compact", maximumFractionDigits: 1 });

  ctx.font = "11px Inter, system-ui, sans-serif";
  ctx.lineWidth = 1;
  for (let step = 0; step <= 4; step += 1) {
    const value = min + ((max - min) * step) / 4;
    const lineY = Math.round(y(value)) + 0.5;
    ctx.strokeStyle = "#eaecf0";
    ctx.beginPath();
    ctx.moveTo(pad.left, lineY);
    ctx.lineTo(width - pad.right, lineY);
    ctx.stroke();
    ctx.fillStyle = "#98a2b3";
    ctx.textAlign = "right";
    ctx.fillText(compact.format(value), pad.left - 8, lineY + 4);
  }
  if (min < 0) {
    ctx.strokeStyle = "#d0d5dd";
    ctx.beginPath();
    ctx.moveTo(pad.left, Math.round(y(0)) + 0.5);
    ctx.lineTo(width - pad.right, Math.round(y(0)) + 0.5);
    ctx.stroke();
  }

  const group = innerWidth / rows.length;
  const barWidth = Math.max(3, Math.min(16, group * 0.26));
  const labelEvery = group < 34 ? 2 : 1;
  const centerOf = (index) => pad.left + group * index + group / 2;
  rows.forEach((row, index) => {
    const center = centerOf(index);
    if (row.highlight) {
      ctx.fillStyle = "rgba(14, 124, 107, 0.06)";
      ctx.fillRect(pad.left + group * index, pad.top, group, innerHeight);
    }
    drawRoundedBar(ctx, center - barWidth - 2, y(row.income), barWidth, y(0) - y(row.income), "#12a17a");
    drawRoundedBar(ctx, center + 2, y(row.expense), barWidth, y(0) - y(row.expense), "#f0772b");
    if (index % labelEvery === 0) {
      ctx.fillStyle = row.highlight ? "#101828" : "#667085";
      ctx.textAlign = "center";
      ctx.fillText(row.label, center, height - 8);
    }
  });

  // Months with no transactions have no profit to plot, so the line skips them.
  const hasData = (row) => row.income || row.expense;
  ctx.strokeStyle = "#2e5bff";
  ctx.lineWidth = 2;
  ctx.beginPath();
  rows.forEach((row, index) => {
    if (!hasData(row)) return;
    const pointY = y(row.income - row.expense);
    if (index && hasData(rows[index - 1])) ctx.lineTo(centerOf(index), pointY);
    else ctx.moveTo(centerOf(index), pointY);
  });
  ctx.stroke();
  rows.forEach((row, index) => {
    if (!hasData(row)) return;
    ctx.beginPath();
    ctx.arc(centerOf(index), y(row.income - row.expense), 3, 0, Math.PI * 2);
    ctx.fillStyle = "#ffffff";
    ctx.fill();
    ctx.stroke();
  });
}

function renderCashFlowChart() {
  const { currency, excluded } = chartCurrency(state.finance);
  const month = todayIso().slice(0, 7);
  const rows = Array.from({ length: 12 }, (_, index) => {
    const key = shiftMonth(month, index - 11);
    const monthTotals = financeTotalsWhere(currency, (entry) => monthKeyOf(entry.date) === key);
    return { label: monthName(key), income: monthTotals.income, expense: monthTotals.expense, highlight: key === month };
  });
  byId("cashFlowNote").textContent = `Last 12 months in ${currency}${
    excluded ? ` · ${excluded} transaction${excluded === 1 ? "" : "s"} in other currencies not included (never converted)` : ""
  }`;
  drawFinanceChart(byId("cashFlowChart"), rows);
}

function renderInventoryChart() {
  const canvas = byId("inventoryChart");
  const ctx = prepareCanvas(canvas);
  const width = canvas.clientWidth;
  const height = canvas.clientHeight;
  const showMoney = canAccess("sensitiveNumbers");
  ctx.clearRect(0, 0, width, height);
  const { currency, excluded } = chartCurrency(state.inventory);
  const groupLabel = fieldLabel("inventory", "category").toLowerCase();
  byId("inventoryMixTitle").textContent = `${viewTitle("inventory")} mix`;
  byId("inventoryMixNote").textContent = showMoney
    ? `Stock value by ${groupLabel}, in ${currency}${excluded ? ` (${excluded} in other currencies not included)` : ""}`
    : `Units by ${groupLabel}`;
  const colors = ["#0e7c6b", "#2e5bff", "#f0772b", "#7a5af8", "#12a17a", "#ee46bc", "#667085"];
  const chartItems = showMoney ? state.inventory.filter((item) => recordCurrency(item) === currency) : state.inventory;
  const byCategory = chartItems.reduce((acc, item) => {
    const value = showMoney ? parseMoney(item.quantity) * parseMoney(item.unitCost) : parseMoney(item.quantity);
    const category = item.category || "Uncategorized";
    acc[category] = (acc[category] || 0) + value;
    return acc;
  }, {});
  let entries = Object.entries(byCategory).sort((a, b) => b[1] - a[1]);
  if (entries.length > 7) {
    const rest = entries.slice(6).reduce((sum, [, value]) => sum + value, 0);
    entries = [...entries.slice(0, 6), ["Other", rest]];
  }
  const total = entries.reduce((sum, [, value]) => sum + value, 0);
  const centerX = width / 2;
  const centerY = height / 2;
  const radius = Math.min(width, height) * 0.44;

  if (!total) {
    ctx.fillStyle = "#667085";
    ctx.font = "500 13px Inter, system-ui, sans-serif";
    ctx.textAlign = "center";
    ctx.fillText("No stock recorded yet", centerX, centerY);
    byId("inventoryLegend").innerHTML = "";
    return;
  }

  let start = -Math.PI / 2;
  entries.forEach(([, value], index) => {
    const slice = (value / total) * Math.PI * 2;
    ctx.beginPath();
    ctx.moveTo(centerX, centerY);
    ctx.arc(centerX, centerY, radius, start, start + slice);
    ctx.closePath();
    ctx.fillStyle = colors[index % colors.length];
    ctx.fill();
    start += slice;
  });
  ctx.beginPath();
  ctx.arc(centerX, centerY, radius * 0.62, 0, Math.PI * 2);
  ctx.fillStyle = "#ffffff";
  ctx.fill();
  // The ring centre is small: show a compact total; the legend has exact amounts.
  const compact = new Intl.NumberFormat("en-US", { notation: "compact", maximumFractionDigits: 1 }).format(total);
  ctx.fillStyle = "#101828";
  ctx.font = "600 13px Inter, system-ui, sans-serif";
  ctx.textAlign = "center";
  ctx.fillText(showMoney ? `${currency} ${compact}` : `${compact} units`, centerX, centerY + 5);

  byId("inventoryLegend").innerHTML = entries
    .map(
      ([category, value], index) => `
        <div class="legend-item">
          <span><i class="swatch" style="background:${colors[index % colors.length]}"></i>${escapeHtml(category)}</span>
          <strong>${escapeHtml(showMoney ? formatCurrency(value, currency) : `${formatNumber(value)} units`)}</strong>
        </div>
      `,
    )
    .join("");
}

function renderOwnerInsights() {
  const insights = buildInsights().slice(0, 4);
  byId("ownerInsights").innerHTML = insights.length
    ? insights
        .map(
          (insight) => `
        <div class="insight-item ${insight.tone}">
          <strong>${escapeHtml(insight.title)}</strong>
          <span>${escapeHtml(insight.detail)}</span>
        </div>
      `,
        )
        .join("")
    : `<div class="empty-state">Health summaries appear as records are added.</div>`;
}

function buildInsights() {
  const summary = totals();
  const insights = [];
  const month = todayIso().slice(0, 7);
  const monthIncome = sumByCurrency(
    state.finance.filter((entry) => entry.type === "Income" && monthKeyOf(entry.date) === month),
    (entry) => parseMoney(entry.amount),
  );

  if (can("finance", "view") && state.finance.length) {
    const netEntries = currencyEntries(summary.net);
    const losing = netEntries.filter(([, amount]) => amount < 0);
    const multiCurrency = netEntries.length > 1;
    insights.push(
      losing.length
        ? {
            tone: "critical",
            title: multiCurrency
              ? `Expenses are above income in ${losing.map(([code]) => code).join(" and ")}`
              : "Expenses are above income",
            detail: `${losing
              .map(([code, amount]) => formatCurrency(Math.abs(amount), code))
              .join(" and ")} more has gone out than came in${multiCurrency ? " (each currency counted separately)" : ""}.`,
          }
        : {
            tone: "good",
            title: "Business is net positive",
            detail: `Recorded income is ahead of expenses by ${formatMoneyTotals(summary.net)}${
              multiCurrency ? " (each currency counted separately)" : ""
            }.`,
          },
    );
    const overdueIncome = sumByCurrency(
      state.finance.filter((entry) => entry.type === "Income" && financeStatus(entry) === "Overdue"),
      (entry) => parseMoney(entry.amount),
    );
    if (currencyEntries(overdueIncome).length) {
      insights.push({
        tone: "watch",
        title: "Chase overdue payments",
        detail: `${formatMoneyTotals(overdueIncome)} from customers is past its due date.`,
      });
    }
  }

  if (can("employees", "view") && can("finance", "view")) {
    // Compare payroll with income only within the same currency.
    currencyEntries(summary.payroll).forEach(([code, payroll]) => {
      const income = monthIncome[code] || 0;
      if (!income) return;
      const payrollShare = Math.round((payroll / income) * 100);
      if (payrollShare > 45) {
        insights.push({
          tone: "watch",
          title: "Payroll is a large cost",
          detail:
            payrollShare > 100
              ? `Monthly payroll (${formatCurrency(payroll, code)}) is more than this month's ${code} income so far.`
              : `Monthly payroll equals about ${payrollShare}% of this month's ${code} income.`,
        });
      }
    });
  }

  if (can("inventory", "view")) {
    const lowStock = state.inventory.filter(isLowStock).length;
    if (lowStock) {
      insights.push({
        tone: "watch",
        title: "Reorder stock soon",
        detail: `${lowStock} record${lowStock === 1 ? " is" : "s are"} at or below reorder level.`,
      });
    }
  }

  return insights;
}

function renderActivity() {
  const entries = state.activity.filter(canSeeActivity).slice(0, 8);
  byId("activityTimeline").innerHTML = entries.length
    ? entries
        .map(
          (entry) => `
        <div class="timeline-item">
          <strong>${escapeHtml(entry.message)}</strong>
          <span>${escapeHtml([entry.by, formatDateTime(entry.date)].filter(Boolean).join(" · "))}</span>
        </div>
      `,
        )
        .join("")
    : `<div class="empty-state">Activity will appear here as records change.</div>`;
}

/* ---------- Tabs ---------- */

const TAB_STORAGE_KEY = "ledgerflow-tabs-v1";
const tabState = (() => {
  try {
    return JSON.parse(sessionStorage.getItem(TAB_STORAGE_KEY)) || {};
  } catch (error) {
    return {};
  }
})();

function missingCount(module) {
  return state[module].filter((record) => needsInfo(module, record)).length;
}

function tabDefinitions(view) {
  if (view === "finance") {
    const dues = buildDues();
    const overdue = [...dues.receivables, ...dues.payables].filter((item) => item.overdue).length;
    return [
      { key: "overview", label: "Overview", show: canAccess("sensitiveNumbers") },
      { key: "income", label: "Income", count: state.finance.filter((entry) => entry.type === "Income").length },
      { key: "expenses", label: "Expenses", count: state.finance.filter((entry) => entry.type === "Expense").length },
      { key: "dues", label: "Dues", count: overdue, alert: true, hideZero: true },
      { key: "all", label: "All transactions", count: state.finance.length },
    ];
  }
  if (view === "inventory") {
    const low = state.inventory.filter(isLowStock).length;
    return [
      { key: "all", label: `All ${viewTitle("inventory").toLowerCase()}`, count: state.inventory.length },
      { key: "low", label: "Low stock", show: tracksReorderLevel(), count: low, alert: true, hideZero: true },
      { key: "missing", label: "Needs info", show: missingCount("inventory") > 0, count: missingCount("inventory"), alert: true },
    ];
  }
  if (view === "contacts") {
    return [
      { key: "customers", label: "Customers", count: state.contacts.filter((contact) => contact.type === "Customer").length },
      { key: "suppliers", label: "Suppliers", count: state.contacts.filter((contact) => contact.type === "Supplier").length },
      { key: "all", label: "All", count: state.contacts.length },
      { key: "missing", label: "Needs info", show: missingCount("contacts") > 0, count: missingCount("contacts"), alert: true },
    ];
  }
  if (view === "employees") {
    const unpaid = canAccess("sensitiveNumbers") ? state.employees.filter((employee) => salaryMonthsDue(employee).length).length : 0;
    return [
      { key: "directory", label: "Directory", count: state.employees.length },
      { key: "payroll", label: "Payroll", show: canAccess("sensitiveNumbers"), count: unpaid, alert: true, hideZero: true },
      { key: "access", label: "Logins & access", show: canAccess("settings") },
      { key: "missing", label: "Needs info", show: missingCount("employees") > 0, count: missingCount("employees"), alert: true },
    ];
  }
  if (view === "assets") {
    const repairs = state.assets.filter(needsMaintenance).length;
    return [
      { key: "all", label: `All ${viewTitle("assets").toLowerCase()}`, count: state.assets.length },
      { key: "maintenance", label: "Needs maintenance", count: repairs, alert: true, hideZero: true },
      { key: "missing", label: "Needs info", show: missingCount("assets") > 0, count: missingCount("assets"), alert: true },
    ];
  }
  if (view === "setup") {
    return [
      { key: "business", label: "Business" },
      { key: "roles", label: "Roles & access" },
      { key: "fields", label: "Fields" },
      { key: "templates", label: "Templates & samples" },
    ];
  }
  return [];
}

function visibleTabs(view) {
  return tabDefinitions(view).filter((tab) => tab.show !== false);
}

function activeTab(view) {
  const tabs = visibleTabs(view);
  if (!tabs.length) return "";
  return tabs.some((tab) => tab.key === tabState[view]) ? tabState[view] : tabs[0].key;
}

function rememberTab(view, key) {
  tabState[view] = key;
  try {
    sessionStorage.setItem(TAB_STORAGE_KEY, JSON.stringify(tabState));
  } catch (error) {
    console.warn("Could not remember tab", error);
  }
}

function renderTabs(view) {
  const nav = byId(`${view}Tabs`);
  if (!nav) return;
  const tabs = visibleTabs(view);
  const active = activeTab(view);
  nav.innerHTML = tabs
    .map((tab) => {
      const showCount = tab.count !== undefined && !(tab.hideZero && !tab.count);
      return `
        <button type="button" class="tab ${tab.key === active ? "is-active" : ""}" data-tab="${view}:${tab.key}">
          ${escapeHtml(tab.label)}
          ${showCount ? `<span class="tab-count ${tab.alert && tab.count ? "alert" : ""}">${formatNumber(tab.count)}</span>` : ""}
        </button>
      `;
    })
    .join("");
  document.querySelectorAll(`[data-panel-view="${view}"]`).forEach((panel) => {
    panel.hidden = !panel.dataset.panelTabs.split(" ").includes(active);
  });
}

function openTab(view, key) {
  rememberTab(view, key);
  if (currentView === view) render();
  else setView(view);
}

/* ---------- Finance ---------- */

const financeViewState = { period: "monthly", year: null, currency: null, monthFilter: "all" };

function financeCurrencies() {
  return [...new Set(state.finance.map(recordCurrency))];
}

function renderFinance() {
  const tab = activeTab("finance");
  if (tab === "overview") renderFinanceOverview();
  else if (tab === "dues") renderDuesTab();
  else renderModule("finance");
}

function renderFinanceOverview() {
  const currencies = financeCurrencies();
  if (!currencies.includes(financeViewState.currency)) financeViewState.currency = chartCurrency(state.finance).currency;
  const currency = financeViewState.currency;
  const currentYear = Number(todayIso().slice(0, 4));
  const years = [
    ...new Set([
      currentYear,
      ...state.finance
        .filter((entry) => recordCurrency(entry) === currency)
        .map((entry) => Number(String(entry.date).slice(0, 4)))
        .filter(Boolean),
    ]),
  ].sort((a, b) => b - a);
  if (!years.includes(financeViewState.year)) financeViewState.year = currentYear;
  const year = financeViewState.year;
  const monthly = financeViewState.period === "monthly";
  const month = todayIso().slice(0, 7);
  const money = (value) => formatCurrency(value, currency);

  document.querySelectorAll("[data-finance-period]").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.financePeriod === financeViewState.period);
  });
  byId("financeYearControl").hidden = !monthly;
  byId("financeYearSelect").innerHTML = years
    .map((value) => `<option value="${value}" ${value === year ? "selected" : ""}>${value}</option>`)
    .join("");
  byId("financeCurrencyControl").hidden = currencies.length < 2;
  byId("financeCurrencySelect").innerHTML = currencies
    .map((code) => `<option value="${code}" ${code === currency ? "selected" : ""}>${code}</option>`)
    .join("");
  const otherCount = state.finance.filter((entry) => recordCurrency(entry) !== currency).length;
  byId("financeOverviewNote").textContent = otherCount
    ? `Showing ${currency} only. ${otherCount} transaction${otherCount === 1 ? " is" : "s are"} in other currencies and never added in.`
    : `All amounts in ${currency}.`;

  const inScope = monthly ? (entry) => String(entry.date).startsWith(`${year}-`) : () => true;
  const scopeTotals = financeTotalsWhere(currency, inScope);
  const profit = scopeTotals.income - scopeTotals.expense;
  const scopeLabel = monthly ? String(year) : "all years";
  const kpis = [
    kpiCard({ label: `Income · ${scopeLabel}`, value: money(scopeTotals.income), icon: "income", tone: "income" }),
    kpiCard({ label: `Expenses · ${scopeLabel}`, value: money(scopeTotals.expense), icon: "expense", tone: "expense" }),
    kpiCard({
      label: `Profit · ${scopeLabel}`,
      value: money(profit),
      icon: "profit",
      tone: "profit",
      negative: profit < 0,
      note: scopeTotals.income ? `${Math.round((profit / scopeTotals.income) * 100)}% margin` : "",
    }),
  ];
  if (monthly && year === currentYear) {
    const thisMonth = financeTotalsWhere(currency, (entry) => monthKeyOf(entry.date) === month);
    kpis.push(
      kpiCard({
        label: `Profit · ${monthName(month, "long")}`,
        value: money(thisMonth.income - thisMonth.expense),
        icon: "wallet",
        tone: "violet",
        negative: thisMonth.income - thisMonth.expense < 0,
        note: `In ${money(thisMonth.income)} · Out ${money(thisMonth.expense)}`,
      }),
    );
  }
  byId("financeKpis").innerHTML = kpis.join("");

  const rows = monthly
    ? Array.from({ length: 12 }, (_, index) => {
        const key = `${year}-${pad2(index + 1)}`;
        return {
          key,
          label: monthName(key),
          longLabel: monthName(key, "long"),
          ...financeTotalsWhere(currency, (entry) => monthKeyOf(entry.date) === key),
          current: key === month,
          future: key > month,
        };
      })
    : years
        .slice()
        .sort((a, b) => a - b)
        .map((value) => ({
          key: String(value),
          label: String(value),
          longLabel: String(value),
          ...financeTotalsWhere(currency, (entry) => String(entry.date).startsWith(`${value}-`)),
          current: value === currentYear,
          future: false,
        }));

  byId("financeChartTitle").textContent = monthly ? `Month by month · ${year}` : "Year by year";
  drawFinanceChart(
    byId("financeChart"),
    rows.map((row) => ({ label: row.label, income: row.income, expense: row.expense, highlight: row.current })),
  );

  byId("financePeriodTableTitle").textContent = monthly ? `${year}, month by month` : "Year by year";
  const blank = (row) => row.future && !row.count;
  const margin = (income, expense) => (income ? `${Math.round(((income - expense) / income) * 100)}%` : "—");
  byId("financePeriodTable").innerHTML = `
    <thead>
      <tr>
        <th>${monthly ? "Month" : "Year"}</th>
        <th class="num">Income</th>
        <th class="num">Expenses</th>
        <th class="num">Profit</th>
        <th class="num">Margin</th>
        <th class="num">Transactions</th>
      </tr>
    </thead>
    <tbody>
      ${rows
        .map(
          (row) => `
        <tr class="is-clickable ${row.current ? "is-current" : ""} ${blank(row) ? "is-future" : ""}" data-period-row="${row.key}">
          <td><strong>${escapeHtml(row.longLabel)}</strong>${row.current ? ` ${badge(monthly ? "This month" : "This year", "green")}` : ""}</td>
          <td class="num">${blank(row) ? "—" : `<span class="money-in">${escapeHtml(money(row.income))}</span>`}</td>
          <td class="num">${blank(row) ? "—" : escapeHtml(money(row.expense))}</td>
          <td class="num">${blank(row) ? "—" : `<strong class="${row.income - row.expense < 0 ? "negative-text" : ""}">${escapeHtml(money(row.income - row.expense))}</strong>`}</td>
          <td class="num">${blank(row) ? "—" : margin(row.income, row.expense)}</td>
          <td class="num">${blank(row) ? "—" : formatNumber(row.count)}</td>
        </tr>
      `,
        )
        .join("")}
    </tbody>
    <tfoot>
      <tr>
        <td>Total</td>
        <td class="num">${escapeHtml(money(scopeTotals.income))}</td>
        <td class="num">${escapeHtml(money(scopeTotals.expense))}</td>
        <td class="num">${escapeHtml(money(profit))}</td>
        <td class="num">${margin(scopeTotals.income, scopeTotals.expense)}</td>
        <td class="num">${formatNumber(scopeTotals.count)}</td>
      </tr>
    </tfoot>
  `;

  const scopeNote = `${monthly ? year : "All years"} · ${currency}`;
  byId("financeIncomeCategoriesNote").textContent = scopeNote;
  byId("financeExpenseCategoriesNote").textContent = scopeNote;
  renderCategoryBars("financeIncomeCategories", "Income", currency, inScope);
  renderCategoryBars("financeExpenseCategories", "Expense", currency, inScope);
}

function renderCategoryBars(targetId, type, currency, inScope) {
  const byCategory = {};
  state.finance.forEach((entry) => {
    if (entry.type !== type || recordCurrency(entry) !== currency || !inScope(entry)) return;
    const category = entry.category || "Uncategorized";
    byCategory[category] = (byCategory[category] || 0) + parseMoney(entry.amount);
  });
  const entries = Object.entries(byCategory).sort((a, b) => b[1] - a[1]);
  const total = entries.reduce((sum, [, value]) => sum + value, 0);
  byId(targetId).innerHTML = entries.length
    ? entries
        .slice(0, 8)
        .map(([category, value]) => {
          const share = total ? Math.round((value / total) * 100) : 0;
          return `
          <div class="bar-row">
            <div class="bar-row-top">
              <span>${escapeHtml(category)}</span>
              <strong>${escapeHtml(formatCurrency(value, currency))} · ${share}%</strong>
            </div>
            <div class="bar-track"><div class="bar-fill ${type === "Expense" ? "expense" : ""}" style="width:${share}%"></div></div>
          </div>
        `;
        })
        .join("")
    : `<div class="empty-state">No ${type.toLowerCase()} recorded for this period.</div>`;
}

function dueGroupsHtml(items, direction) {
  if (!items.length) {
    return `<div class="empty-state">${direction === "in" ? "Nobody owes the business anything right now." : "Nothing to pay right now."}</div>`;
  }
  const overdue = items.filter((item) => item.overdue);
  const open = items.filter((item) => !item.overdue);
  return [
    overdue.length ? `<p class="due-group-title">Overdue</p>${overdue.map((item) => dueItemHtml(item, direction)).join("")}` : "",
    open.length
      ? `<p class="due-group-title">${overdue.length ? "Upcoming and open" : "Open"}</p>${open.map((item) => dueItemHtml(item, direction)).join("")}`
      : "",
  ].join("");
}

function renderDuesTab() {
  const dues = buildDues();
  const owedIn = moneyHeadline(duesTotals(dues.receivables));
  const owedInLate = moneyHeadline(duesTotals(dues.receivables, true));
  const owedOut = moneyHeadline(duesTotals(dues.payables));
  const owedOutLate = moneyHeadline(duesTotals(dues.payables, true));
  const lateIn = dues.receivables.filter((item) => item.overdue).length;
  const lateOut = dues.payables.filter((item) => item.overdue).length;
  byId("duesTotals").innerHTML = [
    kpiCard({
      label: "Owed to you",
      value: owedIn.value,
      icon: "inbox",
      tone: "income",
      note: [owedIn.others, `${dues.receivables.length} open`].filter(Boolean).join(" · "),
    }),
    kpiCard({
      label: "Overdue · owed to you",
      value: owedInLate.value,
      icon: "alert",
      tone: lateIn ? "alert" : "",
      note: [owedInLate.others, `${lateIn} overdue`].filter(Boolean).join(" · "),
    }),
    kpiCard({
      label: "You owe",
      value: owedOut.value,
      icon: "clock",
      tone: "expense",
      note: [owedOut.others, `${dues.payables.length} open`].filter(Boolean).join(" · "),
    }),
    kpiCard({
      label: "Overdue · you owe",
      value: owedOutLate.value,
      icon: "alert",
      tone: lateOut ? "alert" : "",
      note: [owedOutLate.others, `${lateOut} overdue`].filter(Boolean).join(" · "),
    }),
  ].join("");
  byId("receivablesList").innerHTML = dueGroupsHtml(dues.receivables, "in");
  byId("payablesList").innerHTML = dueGroupsHtml(dues.payables, "out");
}

function markPaid(id) {
  const entry = state.finance.find((record) => record.id === id);
  if (!entry || !can("finance", "edit")) return;
  entry.status = "Paid";
  entry.paidOn = todayIso();
  const verb = entry.type === "Income" ? "Received" : "Paid";
  addActivity(`${verb}: ${entry.description || entry.category} (${formatCurrency(entry.amount, recordCurrency(entry))})`, "finance");
  saveState();
  render();
  showToast(`Marked as ${verb.toLowerCase()}`);
}

function recordPayment(contactId) {
  const contact = state.contacts.find((record) => record.id === contactId);
  if (!contact || !can("finance", "add")) return;
  const balance = parseMoney(contact.balance);
  if (!balance) {
    showToast("No opening balance to settle. Mark unpaid transactions as paid in Finance → Dues.");
    return;
  }
  const incoming = balance > 0;
  openModal("finance", null, {
    prefill: {
      date: todayIso(),
      type: incoming ? "Income" : "Expense",
      category: incoming ? "Customer payment" : "Supplier payment",
      description: `Payment ${incoming ? "from" : "to"} ${contact.name}`,
      amount: Math.abs(balance),
      currency: recordCurrency(contact),
      status: "Paid",
      party: contact.name,
    },
    context: { settleContactId: contact.id },
  });
}

/* ---------- Employees: payroll and access ---------- */

let payrollMonth = null;

function renderEmployeesView() {
  const tab = activeTab("employees");
  if (tab === "payroll") renderPayroll();
  else if (tab === "access") renderAccess();
  else renderModule("employees");
}

function payrollRows(month) {
  return state.employees.map((employee) => {
    const payment = salaryPayment(employee, month);
    const onPayroll = month >= payrollStartMonth(employee);
    const payable = salaryIsPayable(employee);
    return { employee, payment, onPayroll, payable, unpaid: payable && onPayroll && !payment };
  });
}

function renderPayroll() {
  const current = todayIso().slice(0, 7);
  payrollMonth ||= current;
  const months = Array.from({ length: 12 }, (_, index) => shiftMonth(current, -index));
  byId("payrollMonthSelect").innerHTML = months
    .map((key) => `<option value="${key}" ${key === payrollMonth ? "selected" : ""}>${monthName(key, "long")}</option>`)
    .join("");

  const rows = payrollRows(payrollMonth);
  const due = rows.filter((row) => row.payment || (row.payable && row.onPayroll));
  const payrollTotal = sumByCurrency(due, (row) => (row.payment ? parseMoney(row.payment.amount) : parseMoney(row.employee.salary)));
  const paidTotal = sumByCurrency(
    rows.filter((row) => row.payment).map((row) => row.payment),
    (payment) => parseMoney(payment.amount),
  );
  const unpaidRows = rows.filter((row) => row.unpaid);
  const unpaidTotal = sumByCurrency(
    unpaidRows.map((row) => row.employee),
    (employee) => parseMoney(employee.salary),
  );
  const late = payrollMonth < current;
  const total = moneyHeadline(payrollTotal);
  const paid = moneyHeadline(paidTotal);
  const unpaid = moneyHeadline(unpaidTotal);
  byId("payrollTotals").innerHTML = [
    kpiCard({ label: `Payroll · ${monthName(payrollMonth, "long")}`, value: total.value, icon: "users", note: [total.others, `${due.length} employees`].filter(Boolean).join(" · ") }),
    kpiCard({ label: "Paid", value: paid.value, icon: "income", tone: "income", note: [paid.others, `${rows.filter((row) => row.payment).length} paid`].filter(Boolean).join(" · ") }),
    kpiCard({
      label: late ? "Unpaid · overdue" : "Still to pay",
      value: unpaid.value,
      icon: "clock",
      tone: unpaidRows.length ? (late ? "alert" : "expense") : "",
      note: [unpaid.others, `${unpaidRows.length} unpaid`].filter(Boolean).join(" · "),
    }),
  ].join("");

  const canPay = can("finance", "add");
  byId("payAllSalariesBtn").hidden = !canPay || !unpaidRows.length;
  byId("payrollTable").innerHTML = `
    <thead>
      <tr>
        <th>Employee</th>
        <th>Job</th>
        <th class="num">Monthly salary</th>
        <th>Status</th>
        <th></th>
      </tr>
    </thead>
    <tbody>
      ${
        rows.length
          ? rows
              .map(({ employee, payment, onPayroll, payable, unpaid: isUnpaidRow }) => {
                let status;
                if (payment) status = badge(`Paid ${formatDate(payment.date)}`, "green");
                else if (employee.status === "Inactive") status = badge("Inactive", "blue");
                else if (!payable) status = badge("Salary not set", "amber");
                else if (!onPayroll) status = badge("Joined later", "blue");
                else status = badge(late ? "Unpaid · overdue" : "Unpaid", late ? "red" : "amber");
                const actions = [];
                if (isUnpaidRow && canPay) {
                  actions.push(`<button type="button" class="accent" data-pay-salary="${escapeHtml(employee.id)}:${payrollMonth}">Pay salary</button>`);
                }
                if (!payable && employee.status !== "Inactive" && can("employees", "edit")) {
                  actions.push(`<button type="button" data-edit="employees:${escapeHtml(employee.id)}">Set salary</button>`);
                }
                return `
                <tr>
                  <td>${titleCell(employee.name, employee.department)}</td>
                  <td>${escapeHtml(employee.role || "—")}</td>
                  <td class="num">${escapeHtml(payment ? formatCurrency(payment.amount, recordCurrency(payment)) : formatRecordMoney(employee, employee.salary))}</td>
                  <td>${status}</td>
                  <td>${actions.length ? `<div class="row-actions">${actions.join("")}</div>` : ""}</td>
                </tr>
              `;
              })
              .join("")
          : `<tr><td colspan="5"><div class="empty-state">No employees yet.</div></td></tr>`
      }
    </tbody>
  `;
}

function paySalary(employeeId, month, { batch = false } = {}) {
  const employee = state.employees.find((record) => record.id === employeeId);
  if (!employee || !can("finance", "add") || !canAccess("sensitiveNumbers")) return false;
  if (salaryPayment(employee, month)) {
    if (!batch) showToast(`${employee.name} is already paid for ${monthName(month, "long")}`);
    return false;
  }
  const amount = parseMoney(employee.salary);
  if (!amount) {
    if (!batch) showToast("Set this employee's salary first");
    return false;
  }
  const currency = recordCurrency(employee);
  if (
    !batch &&
    !window.confirm(
      `Pay ${employee.name} ${formatCurrency(amount, currency)} for ${monthName(month, "long")}? It will be recorded as a paid expense.`,
    )
  ) {
    return false;
  }
  state.finance.unshift({
    id: makeId("finance"),
    date: month === todayIso().slice(0, 7) ? todayIso() : lastDayOfMonth(month),
    type: "Expense",
    category: "Salaries",
    description: `Salary - ${employee.name} - ${monthName(month, "long")}`,
    amount,
    currency,
    status: "Paid",
    party: employee.name,
    employeeId: employee.id,
    payrollMonth: month,
    paidOn: todayIso(),
    sourceName: "Payroll",
    createdAt: new Date().toISOString(),
    infoDismissed: true,
  });
  if (!batch) {
    addActivity(`Salary paid: ${employee.name} (${formatCurrency(amount, currency)}, ${monthName(month, "long")})`, "finance");
    saveState();
    render();
    showToast(`${employee.name}'s salary recorded as paid`);
  }
  return true;
}

function payAllSalaries(month) {
  const unpaid = payrollRows(month)
    .filter((row) => row.unpaid)
    .map((row) => row.employee);
  if (!unpaid.length) return;
  const total = formatMoneyTotals(sumByCurrency(unpaid, (employee) => parseMoney(employee.salary)));
  if (!window.confirm(`Pay ${unpaid.length} salar${unpaid.length === 1 ? "y" : "ies"} for ${monthName(month, "long")} (${total})?`)) {
    return;
  }
  const paid = unpaid.filter((employee) => paySalary(employee.id, month, { batch: true })).length;
  addActivity(`Paid ${paid} salaries for ${monthName(month, "long")} (${total})`, "finance");
  saveState();
  render();
  showToast(`${paid} salar${paid === 1 ? "y" : "ies"} recorded as paid`);
}

function renderAccess() {
  const canEdit = can("employees", "edit");
  byId("accessTable").innerHTML = `
    <thead>
      <tr>
        <th>Employee</th>
        <th>Login email</th>
        <th>Role</th>
        <th>Sign-in</th>
        <th></th>
      </tr>
    </thead>
    <tbody>
      ${
        state.employees.length
          ? state.employees
              .map((employee) => {
                const status = loginStatus(employee);
                const tone = status === "Can sign in" ? "green" : status === "No login" ? "blue" : "amber";
                return `
                <tr>
                  <td>${titleCell(employee.name, employee.role)}</td>
                  <td>${escapeHtml(employee.loginEmail || "—")}</td>
                  <td>${escapeHtml(roleName(employee.accessRole))}</td>
                  <td>${badge(status, tone)}</td>
                  <td>${canEdit ? `<div class="row-actions"><button type="button" data-edit="employees:${escapeHtml(employee.id)}">${employee.loginEmail ? "Change" : "Set up login"}</button></div>` : ""}</td>
                </tr>
              `;
              })
              .join("")
          : `<tr><td colspan="5"><div class="empty-state">No employees yet.</div></td></tr>`
      }
    </tbody>
  `;
}

/* ---------- Module tables ---------- */

function contactDues(contact) {
  const byCurrency = {};
  const balance = parseMoney(contact.balance);
  if (balance) byCurrency[recordCurrency(contact)] = balance;
  const name = String(contact.name || "").trim().toLowerCase();
  if (name) {
    state.finance.forEach((entry) => {
      if (!isUnpaid(entry) || String(entry.party || "").trim().toLowerCase() !== name) return;
      const code = recordCurrency(entry);
      const signed = entry.type === "Income" ? parseMoney(entry.amount) : -parseMoney(entry.amount);
      byCurrency[code] = (byCurrency[code] || 0) + signed;
    });
  }
  return byCurrency;
}

function contactDuesHtml(contact) {
  const entries = currencyEntries(contactDues(contact));
  if (!entries.length) return `<span class="muted-note">Settled</span>`;
  return entries
    .map(([code, amount]) =>
      amount > 0
        ? `<div class="money-in">Owes you ${escapeHtml(formatCurrency(amount, code))}</div>`
        : `<div class="money-out">You owe ${escapeHtml(formatCurrency(Math.abs(amount), code))}</div>`,
    )
    .join("");
}

const tableColumns = {
  inventory: [
    {
      field: "name",
      render: (item) => titleCell(item.name, [item.sku, item.location].filter((part) => !isBlank(part)).join(" · ")),
    },
    { field: "category" },
    { field: "quantity", className: "num", render: (item) => escapeHtml(isBlank(item.quantity) ? "—" : formatNumber(item.quantity)) },
    { field: "unitCost", className: "num", render: (item) => escapeHtml(formatRecordMoney(item, item.unitCost)) },
    { field: "sellPrice", className: "num", render: (item) => escapeHtml(formatRecordMoney(item, item.sellPrice)) },
    {
      title: "Stock value",
      sensitive: true,
      className: "num",
      render: (item) =>
        escapeHtml(isBlank(item.unitCost) ? "—" : formatRecordMoney(item, parseMoney(item.quantity) * parseMoney(item.unitCost))),
    },
  ],
  assets: [
    { field: "name", render: (asset) => titleCell(asset.name, [asset.code, asset.notes].filter(Boolean).join(" · ")) },
    { field: "assignedTo" },
    { field: "location" },
    { field: "purchaseDate", render: (asset) => escapeHtml(formatDate(asset.purchaseDate)) },
    { field: "value", className: "num", render: (asset) => escapeHtml(formatRecordMoney(asset, asset.value)) },
  ],
  employees: [
    { field: "name", render: (employee) => titleCell(employee.name, employee.email || employee.phone) },
    { field: "role" },
    { field: "department" },
    { field: "phone" },
    { field: "salary", className: "num", render: (employee) => escapeHtml(formatRecordMoney(employee, employee.salary)) },
  ],
  finance: [
    { field: "date", render: (entry) => escapeHtml(formatDate(entry.date)) },
    {
      field: "description",
      render: (entry) => titleCell(entry.description, [entry.category, entry.party].filter((part) => !isBlank(part)).join(" · ")),
    },
    {
      field: "type",
      render: (entry) =>
        entry.type === "Income" ? badge("Income", "green") : entry.type === "Expense" ? badge("Expense", "amber") : "—",
    },
    { field: "dueDate", render: (entry) => escapeHtml(isUnpaid(entry) ? formatDate(entry.dueDate) : "—") },
    {
      field: "amount",
      className: "num",
      render: (entry) =>
        `<span class="${entry.type === "Income" ? "money-in" : "money-out"}">${escapeHtml(formatRecordMoney(entry, entry.amount))}</span>`,
    },
  ],
  contacts: [
    { field: "name", render: (contact) => titleCell(contact.name, [contact.phone, contact.email].filter(Boolean).join(" · ")) },
    {
      field: "type",
      render: (contact) => (isBlank(contact.type) ? "—" : badge(contact.type, contact.type === "Customer" ? "green" : "blue")),
    },
    { field: "balance", title: "Balance due", render: (contact) => contactDuesHtml(contact) },
  ],
};

function statusCell(module, record) {
  const parts = [];
  if (module === "inventory") parts.push(stockBadge(record));
  if (module === "assets") parts.push(statusBadge(record.condition));
  if (module === "finance") parts.push(statusBadge(financeStatus(record)));
  if (module === "employees" || module === "contacts") parts.push(statusBadge(record.status));
  if (needsInfo(module, record)) parts.push(badge("Needs info", "red"));
  return `<div class="badge-row">${parts.join("")}</div>`;
}

function formatCustomValue(field, value) {
  if (isBlank(value)) return "—";
  if (field.type === "number") {
    // Years, codes and IDs read wrong with thousands separators ("2,025").
    if (/year|code|no\.|number|id\b/i.test(field.label)) return escapeHtml(String(value));
    return escapeHtml(formatNumber(value));
  }
  if (field.type === "date") return escapeHtml(formatDate(value));
  return escapeHtml(value);
}

function visibleColumns(module) {
  const base = tableColumns[module]
    .filter((column) => {
      if (column.sensitive && !canAccess("sensitiveNumbers")) return false;
      if (column.field && !fieldVisible(module, column.field)) return false;
      return true;
    })
    .map((column) => ({
      title: column.title || fieldLabel(module, column.field),
      className: column.className || "",
      render: column.render || ((record) => escapeHtml(isBlank(record[column.field]) ? "—" : record[column.field])),
    }));
  const custom = (state.customFields[module] || []).map((field) => ({
    title: field.label,
    className: field.type === "number" ? "num" : "",
    render: (record) => formatCustomValue(field, record[field.name]),
  }));
  return [...base, ...custom, { title: "Status", className: "", render: (record) => statusCell(module, record) }];
}

function searchText(module, record) {
  const values = formFields(module)
    .filter((field) => !field.virtual)
    .map((field) => (field.name === "accessRole" ? roleName(record.accessRole) : record[field.name]));
  return values.join(" ").toLowerCase();
}

function filteredRecords(module) {
  const query = els.globalSearch.value.trim().toLowerCase();
  const tab = activeTab(module);
  let rows = state[module].filter((record) => {
    if (query && !searchText(module, record).includes(query)) return false;
    if (tab === "missing" && !needsInfo(module, record)) return false;
    return true;
  });

  if (module === "inventory") {
    const category = byId("inventoryCategoryFilter").value;
    rows = rows.filter((item) => (category === "all" || item.category === category) && (tab !== "low" || isLowStock(item)));
  }
  if (module === "assets" && tab === "maintenance") rows = rows.filter(needsMaintenance);
  if (module === "employees") {
    const department = byId("employeeDepartmentFilter").value;
    rows = rows.filter((employee) => department === "all" || employee.department === department);
  }
  if (module === "contacts") {
    if (tab === "customers") rows = rows.filter((contact) => contact.type === "Customer");
    if (tab === "suppliers") rows = rows.filter((contact) => contact.type === "Supplier");
  }
  if (module === "finance") {
    const category = byId("financeCategoryFilter").value;
    const payment = byId("financeStatusFilter").value;
    const period = byId("financeMonthFilter").value;
    rows = rows
      .filter((entry) => {
        if (tab === "income" && entry.type !== "Income") return false;
        if (tab === "expenses" && entry.type !== "Expense") return false;
        if (category !== "all" && entry.category !== category) return false;
        if (payment === "paid" && isUnpaid(entry)) return false;
        if (payment === "unpaid" && !isUnpaid(entry)) return false;
        if (period !== "all" && !String(entry.date || "").startsWith(period)) return false;
        return true;
      })
      .sort((a, b) => String(b.date).localeCompare(String(a.date)));
  }
  return rows;
}

function tableSummary(module, rows) {
  const count = `<strong>${formatNumber(rows.length)}</strong> of ${formatNumber(state[module].length)}`;
  const sensitive = canAccess("sensitiveNumbers");
  if (module === "finance") {
    const income = sumByCurrency(
      rows.filter((entry) => entry.type === "Income"),
      (entry) => parseMoney(entry.amount),
    );
    const expense = sumByCurrency(
      rows.filter((entry) => entry.type === "Expense"),
      (entry) => parseMoney(entry.amount),
    );
    const tab = activeTab("finance");
    const parts = [`${count} transactions`];
    if (tab !== "expenses") parts.push(`In <strong>${escapeHtml(formatMoneyTotals(income))}</strong>`);
    if (tab !== "income") parts.push(`Out <strong>${escapeHtml(formatMoneyTotals(expense))}</strong>`);
    return parts.join(" · ");
  }
  if (module === "inventory" && sensitive) {
    return `${count} · Stock value <strong>${escapeHtml(
      formatMoneyTotals(sumByCurrency(rows, (item) => parseMoney(item.quantity) * parseMoney(item.unitCost))),
    )}</strong>`;
  }
  if (module === "assets" && sensitive) {
    return `${count} · Value <strong>${escapeHtml(formatMoneyTotals(sumByCurrency(rows, (asset) => parseMoney(asset.value))))}</strong>`;
  }
  if (module === "employees" && sensitive) {
    return `${count} · Monthly payroll <strong>${escapeHtml(
      formatMoneyTotals(
        sumByCurrency(
          rows.filter((employee) => employee.status !== "Inactive"),
          (employee) => parseMoney(employee.salary),
        ),
      ),
    )}</strong>`;
  }
  if (module === "contacts") {
    const owedIn = {};
    const owedOut = {};
    rows.forEach((contact) => {
      Object.entries(contactDues(contact)).forEach(([code, amount]) => {
        if (amount > 0) owedIn[code] = (owedIn[code] || 0) + amount;
        if (amount < 0) owedOut[code] = (owedOut[code] || 0) - amount;
      });
    });
    return `${count} · Owed to you <strong>${escapeHtml(formatMoneyTotals(owedIn))}</strong> · You owe <strong>${escapeHtml(
      formatMoneyTotals(owedOut),
    )}</strong>`;
  }
  return count;
}

function renderModule(module) {
  const columns = visibleColumns(module);
  const rows = filteredRecords(module);
  byId(`${module}Head`).innerHTML = `<tr>${columns
    .map((column) => `<th class="${column.className}">${escapeHtml(column.title)}</th>`)
    .join("")}<th></th></tr>`;
  const emptyMessage = state[module].length
    ? "Nothing matches this view."
    : can(module, "add")
      ? "No records yet. Add one, or import a file."
      : "No records yet.";
  byId(`${module}Table`).innerHTML = rows.length
    ? rows
        .map(
          (record) => `
        <tr>
          ${columns.map((column) => `<td class="${column.className}">${column.render(record)}</td>`).join("")}
          <td>${rowActions(module, record)}</td>
        </tr>
      `,
        )
        .join("")
    : `<tr><td colspan="${columns.length + 1}"><div class="empty-state">${escapeHtml(emptyMessage)}</div></td></tr>`;
  const summary = byId(`${module}TableSummary`);
  if (summary) summary.innerHTML = tableSummary(module, rows);
}

function rowActions(module, record) {
  const actions = [];
  if (module === "finance" && isUnpaid(record) && can("finance", "edit")) {
    actions.push(
      `<button type="button" class="accent" data-mark-paid="${escapeHtml(record.id)}">${record.type === "Income" ? "Mark received" : "Mark paid"}</button>`,
    );
  }
  if (module === "contacts" && parseMoney(record.balance) && can("finance", "add")) {
    actions.push(`<button type="button" class="accent" data-record-payment="${escapeHtml(record.id)}">Record payment</button>`);
  }
  if (can(module, "edit")) {
    actions.push(
      `<button type="button" data-edit="${module}:${escapeHtml(record.id)}">${needsInfo(module, record) ? "Fill in" : "Edit"}</button>`,
    );
  }
  if (can(module, "delete")) {
    actions.push(`<button type="button" class="warn" data-delete="${module}:${escapeHtml(record.id)}">Delete</button>`);
  }
  return actions.length ? `<div class="row-actions">${actions.join("")}</div>` : "";
}

/* ---------- Record form ---------- */

function customSectionName() {
  const type = state.organization?.type;
  return !type || type === "General Business" ? "More details" : `${type} details`;
}

function formSections(module) {
  const sections = new Map();
  formFields(module).forEach((field) => {
    const name = field.custom ? customSectionName() : field.section || "Details";
    if (!sections.has(name)) sections.set(name, []);
    sections.get(name).push(field);
  });
  const ordered = [...sections.entries()].filter(([name]) => name !== "Notes");
  if (sections.has("Notes")) ordered.push(["Notes", sections.get("Notes")]);
  return ordered;
}

function openModal(module, id = null, { prefill = null, context = null } = {}) {
  const action = id ? "edit" : "add";
  if (!schemas[module] || !can(module, action)) {
    showToast(`Your role cannot ${action} ${viewTitle(module) || "records"}`);
    return;
  }
  const record = id ? state[module].find((entry) => entry.id === id) : null;
  if (id && !record) return;
  editing = { module, id, context };
  const source = record || prefill || {};
  const missing = record ? visibleMissingFields(module, record) : [];
  const missingNames = new Set(missing.map((field) => field.name));
  const settling = context?.settleContactId ? state.contacts.find((contact) => contact.id === context.settleContactId) : null;

  els.modalKicker.textContent = viewTitle(module);
  els.modalTitle.textContent = settling
    ? `Record payment · ${settling.name}`
    : id
      ? `Edit ${recordNoun(module).toLowerCase()}`
      : `Add ${recordNoun(module).toLowerCase()}`;
  const datalists =
    module === "finance"
      ? `<datalist id="partyOptions">${state.contacts
          .map((contact) => `<option value="${escapeHtml(contact.name)}"></option>`)
          .join("")}</datalist>
         <datalist id="categoryOptions">${uniqueValues(state.finance, "category")
           .map((category) => `<option value="${escapeHtml(category)}"></option>`)
           .join("")}</datalist>`
      : "";
  els.recordForm.innerHTML = `
    <div class="form-error" id="formError" hidden></div>
    ${
      settling
        ? `<div class="form-notice">
            <strong>This payment reduces ${escapeHtml(settling.name)}'s opening balance</strong>
            <span>Open balance: ${escapeHtml(formatCurrency(Math.abs(parseMoney(settling.balance)), recordCurrency(settling)))}. Enter a smaller amount for a part payment.</span>
          </div>`
        : ""
    }
    ${
      missing.length
        ? `<div class="form-notice">
            <strong>Missing information</strong>
            <span>${escapeHtml(missingLabels(module, missing).join(", "))}. Fill in what you know${
              missing.some((field) => field.required) ? "" : ", or save to continue without it"
            }.</span>
          </div>`
        : ""
    }
    ${formSections(module)
      .map(
        ([name, fields]) => `
        <fieldset class="form-section">
          <legend>${escapeHtml(name)}</legend>
          <div class="form-grid">
            ${fields.map((field) => fieldTemplate(field, source, !record, missingNames.has(field.name))).join("")}
          </div>
        </fieldset>
      `,
      )
      .join("")}
    ${datalists}
    <div class="form-actions">
      <button class="button ghost" id="cancelFormBtn" type="button">Cancel</button>
      <button class="button primary" type="submit">${
        settling ? "Record payment" : id ? "Save changes" : `Add ${recordNoun(module).toLowerCase()}`
      }</button>
    </div>
  `;
  els.modalBackdrop.hidden = false;
  const focusTarget =
    els.recordForm.querySelector(".is-missing input, .is-missing select") ||
    els.recordForm.querySelector("input, select, textarea");
  focusTarget?.focus();
}

function defaultValue(field) {
  if (field.type === "date") return todayIso();
  if (field.type === "select") {
    if (field.name === "accessRole") return fallbackRoleId();
    if (field.name === "currency") return orgCurrency();
    return field.custom ? "" : field.options?.[0] || "";
  }
  return "";
}

function selectOptions(field) {
  if (field.name === "accessRole") {
    return assignableRoles().map((role) => ({ value: role.id, label: role.name }));
  }
  if (field.name === "currency") {
    return currencyOptions().map((code) => ({ value: code, label: code }));
  }
  return (field.options || []).map((option) => ({ value: option, label: option }));
}

function fieldTemplate(field, record, isNew, isMissing) {
  const stored = record[field.name];
  const value = field.virtual ? "" : !isBlank(stored) ? stored : isNew ? defaultValue(field) : "";
  const id = `field-${field.name}`;
  const attrs = [
    field.required ? "required" : "",
    field.min !== undefined ? `min="${field.min}"` : "",
    field.step ? `step="${field.step}"` : "",
  ]
    .filter(Boolean)
    .join(" ");
  const common = `id="${id}" name="${escapeHtml(field.name)}" ${attrs}`;
  const hint = field.hint ? `<small>${escapeHtml(field.hint)}</small>` : "";
  const wrap = (control) => `
    <div class="field ${field.full ? "full" : ""} ${isMissing ? "is-missing" : ""}">
      <label for="${id}">${escapeHtml(field.label)}${field.required ? " *" : ""}</label>
      ${control}
      ${hint}
    </div>
  `;

  if (field.type === "textarea") return wrap(`<textarea ${common}>${escapeHtml(value)}</textarea>`);
  if (field.type === "password") {
    const placeholder = record.passwordHash ? "Leave blank to keep current password" : "Set a password so they can sign in";
    return wrap(
      `<input ${common} type="password" autocomplete="new-password" minlength="6" placeholder="${placeholder}" />`,
    );
  }
  if (field.type === "select") {
    const options = selectOptions(field);
    const includeBlank = !field.required || isBlank(value) || !options.some((option) => option.value === value);
    return wrap(`
      <select ${common}>
        ${includeBlank ? `<option value="">${field.required ? "Choose…" : "—"}</option>` : ""}
        ${options
          .map(
            (option) =>
              `<option value="${escapeHtml(option.value)}" ${String(option.value) === String(value) ? "selected" : ""}>${escapeHtml(
                option.label,
              )}</option>`,
          )
          .join("")}
      </select>
    `);
  }
  const list = field.list ? `list="${field.list}" autocomplete="off"` : "";
  return wrap(`<input ${common} ${list} type="${field.type}" value="${escapeHtml(value)}" />`);
}

function closeModal() {
  els.modalBackdrop.hidden = true;
  editing = null;
}

function showFormError(message) {
  const error = byId("formError");
  if (!error) return;
  error.textContent = message;
  error.hidden = false;
  error.scrollIntoView({ block: "nearest" });
}

async function applyEmployeeLogin(record, formData) {
  if (!canAccess("settings")) return "";
  const loginEmail = String(record.loginEmail || "").trim().toLowerCase();
  record.loginEmail = loginEmail;
  if (loginEmail) {
    const ownerEmail = String(state.organization.ownerEmail || "").toLowerCase();
    const taken =
      loginEmail === ownerEmail ||
      state.employees.some((employee) => employee.id !== record.id && employee.loginEmail === loginEmail);
    if (taken) return "That login email is already used by someone else.";
  }
  const password = String(formData.get("loginPassword") || "");
  if (password) {
    if (password.length < 6) return "Password must be at least 6 characters.";
    if (!loginEmail) return "Add a login email before setting a password.";
    record.passwordSalt = makeSalt();
    record.passwordHash = await hashPassword(password, record.passwordSalt);
  }
  if (!loginEmail) {
    delete record.passwordHash;
    delete record.passwordSalt;
  }
  return "";
}

async function handleSubmit(event) {
  event.preventDefault();
  if (!editing) return;
  const { module, id } = editing;
  if (!can(module, id ? "edit" : "add")) {
    showFormError("Your role no longer allows this change.");
    return;
  }
  const schema = schemas[module];
  const formData = new FormData(els.recordForm);
  const existing = id ? state[module].find((entry) => entry.id === id) : null;
  const record = { ...(existing || {}), id: id || makeId(module) };

  formFields(module).forEach((field) => {
    if (field.virtual) return;
    const raw = String(formData.get(field.name) ?? "").trim();
    record[field.name] = field.type === "number" ? (raw === "" ? "" : parseMoney(raw)) : raw;
  });

  if (module === "employees") {
    const error = await applyEmployeeLogin(record, formData);
    if (error) {
      showFormError(error);
      return;
    }
    if (!assignableRoles().some((role) => role.id === record.accessRole)) record.accessRole = fallbackRoleId();
    if (currentUser()?.id === record.id && record.status === "Inactive") {
      showFormError("You cannot mark yourself inactive while signed in.");
      return;
    }
  }
  const settleId = !existing && editing.context?.settleContactId;
  let settledContact = null;
  if (settleId) {
    settledContact = state.contacts.find((contact) => contact.id === settleId);
    const balance = parseMoney(settledContact?.balance);
    const amount = parseMoney(record.amount);
    const contactCurrency = recordCurrency(settledContact);
    if (!settledContact || !balance) {
      showFormError("This contact no longer has an opening balance to settle.");
      return;
    }
    if (record.currency !== contactCurrency) {
      showFormError(`The payment must be in ${contactCurrency}, the same currency as the balance. LedgerFlow never converts currencies.`);
      return;
    }
    if ((balance > 0 && record.type !== "Income") || (balance < 0 && record.type !== "Expense")) {
      showFormError(balance > 0 ? "A payment from a customer must be Income." : "A payment to a supplier must be an Expense.");
      return;
    }
    if (amount <= 0 || amount > Math.abs(balance) + 0.005) {
      showFormError(`Enter an amount up to ${formatCurrency(Math.abs(balance), contactCurrency)}.`);
      return;
    }
    record.status = "Paid";
    record.party = settledContact.name;
  }
  record.infoDismissed = true;

  if (settledContact) {
    const balance = parseMoney(settledContact.balance);
    settledContact.balance = round2(Math.sign(balance) * (Math.abs(balance) - parseMoney(record.amount)));
    addActivity(
      `Payment ${record.type === "Income" ? "from" : "to"} ${settledContact.name}: ${formatCurrency(record.amount, record.currency)}`,
      "finance",
    );
  }

  if (existing) {
    state[module] = state[module].map((entry) => (entry.id === id ? record : entry));
    addActivity(`${recordNoun(module)} updated: ${recordLabel(module, record)}`, module);
    showToast(`${recordNoun(module)} saved`);
  } else {
    record.sourceName = "Manual Entry";
    record.createdAt = new Date().toISOString();
    state[module].unshift(record);
    touchManualSource();
    addActivity(`${recordNoun(module)} added: ${recordLabel(module, record)}`, module);
    showToast(`${recordNoun(module)} added`);
  }
  saveState();
  closeModal();
  render();
}

function touchManualSource() {
  const source = state.dataSources.find((entry) => entry.id === "source-manual");
  if (!source) return;
  source.status = "Fresh";
  source.lastSync = new Date().toISOString();
  source.records = Number(source.records || 0) + 1;
}

function deleteRecord(module, id) {
  if (!can(module, "delete")) {
    showToast("Your role cannot delete records");
    return;
  }
  const record = state[module].find((entry) => entry.id === id);
  if (!record) return;
  if (module === "employees" && currentUser()?.id === id) {
    showToast("You cannot delete your own employee record while signed in");
    return;
  }
  if (!window.confirm(`Delete "${recordLabel(module, record)}"? This cannot be undone.`)) return;
  state[module] = state[module].filter((entry) => entry.id !== id);
  addActivity(`${recordNoun(module)} deleted: ${recordLabel(module, record)}`, module);
  saveState();
  render();
  showToast(`${recordNoun(module)} deleted`);
}

/* ---------- Import ---------- */

function renderDataSources() {
  const sources = state.dataSources;
  const attention = new Set(["Failed", "Stale", "Needs Auth", "Needs Review"]);
  byId("sourceCount").textContent = formatNumber(sources.length);
  byId("freshSourceCount").textContent = formatNumber(
    sources.filter((source) => source.status === "Fresh" || source.status === "Ready").length,
  );
  byId("sourceAlertCount").textContent = formatNumber(sources.filter((source) => attention.has(source.status)).length);

  const targets = MODULES.filter((module) => can(module, "add"));
  const current = els.importTarget.value;
  els.importTarget.innerHTML = targets
    .map((module) => `<option value="${module}">${escapeHtml(viewTitle(module))}</option>`)
    .join("");
  els.importTarget.value = targets.includes(current) ? current : targets[0] || "";
  els.importTarget.disabled = Boolean(importSession && importSession.step !== "mapping");

  byId("sourceCards").innerHTML = sources
    .map(
      (source) => `
      <article class="source-card">
        <div class="panel-header">
          <h3>${escapeHtml(source.name)}</h3>
          ${sourceStatusBadge(source.status)}
        </div>
        <p>${escapeHtml(source.note)}</p>
        <div class="source-meta">
          <span><strong>Type:</strong> ${escapeHtml(source.type)}</span>
          <span><strong>Target:</strong> ${escapeHtml(source.target)}</span>
          <span><strong>Records:</strong> ${formatNumber(source.records || 0)}</span>
          <span><strong>Last update:</strong> ${escapeHtml(source.lastSync ? formatDateTime(source.lastSync) : "Not yet")}</span>
        </div>
      </article>
    `,
    )
    .join("");

  byId("canonicalModelMap").innerHTML = targets
    .map(
      (module) => `
        <div class="model-row">
          <strong>${escapeHtml(viewTitle(module))}</strong>
          <span>${escapeHtml(importableFields(module).map((field) => field.label).join(", "))}</span>
        </div>
      `,
    )
    .join("");

  renderImportWizard();
}

function renderImportWizard() {
  if (!importSession) {
    els.importWizard.innerHTML = importHelpHtml();
    return;
  }
  if (importSession.step === "mapping") els.importWizard.innerHTML = importMappingHtml();
  else if (importSession.step === "review") els.importWizard.innerHTML = importReviewHtml();
  else els.importWizard.innerHTML = importDoneHtml();
}

function importHelpHtml() {
  const module = els.importTarget.value;
  if (!module) return `<div class="empty-state">Your role cannot add records to any module.</div>`;
  return `
    <div class="import-help">
      <strong>How importing works</strong>
      <span>1. Choose a file. Column names don't need to match exactly.</span>
      <span>2. Match each column to a field in ${escapeHtml(viewTitle(module))}, or add it as a new field.</span>
      <span>3. Fill in any missing details, or continue without them.</span>
    </div>
  `;
}

function stepsHtml(active) {
  return `<ol class="wizard-steps">${["Match columns", "Fill missing info", "Done"]
    .map(
      (label, index) =>
        `<li class="${index === active ? "is-active" : index < active ? "is-done" : ""}">${index + 1}. ${label}</li>`,
    )
    .join("")}</ol>`;
}

function formatSample(value) {
  if (value instanceof Date) return dateToIso(value);
  return String(value ?? "").slice(0, 60);
}

function importMappingHtml() {
  const s = importSession;
  const fields = importableFields(s.module);
  const mapped = new Set(s.mapping.filter(Boolean));
  const unmatchedRequired = fields.filter((field) => field.required && !mapped.has(field.name));
  const rows = s.headers
    .map((header, index) => {
      const sample = s.rows.map((row) => row.cells[index]).find((value) => !isBlank(value));
      const current = s.mapping[index] || "";
      return `
      <tr>
        <td><strong>${escapeHtml(header)}</strong></td>
        <td class="muted-cell">${escapeHtml(formatSample(sample)) || "—"}</td>
        <td>
          <select data-map-index="${index}" aria-label="Field for column ${escapeHtml(header)}">
            <option value="">Skip this column</option>
            ${fields
              .map(
                (field) =>
                  `<option value="${escapeHtml(field.name)}" ${current === field.name ? "selected" : ""}>${escapeHtml(field.label)}</option>`,
              )
              .join("")}
            ${
              canAccess("settings")
                ? `<option value="__new__" ${current === "__new__" ? "selected" : ""}>+ Add as new field "${escapeHtml(header)}"</option>`
                : ""
            }
          </select>
        </td>
      </tr>
    `;
    })
    .join("");

  return `
    ${stepsHtml(0)}
    <p class="wizard-intro">
      Found <strong>${formatNumber(s.rows.length)} rows</strong> and ${formatNumber(s.headers.length)} columns in
      <strong>${escapeHtml(s.fileName)}</strong>${s.note ? ` (${escapeHtml(s.note)})` : ""}.
      Choose where each column should go in ${escapeHtml(viewTitle(s.module))}.
    </p>
    <div class="table-wrap compact-table">
      <table class="mapping-table">
        <thead><tr><th>Column in your file</th><th>First value</th><th>Save into</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>
    ${
      unmatchedRequired.length
        ? `<div class="form-notice">
            <strong>Not matched yet</strong>
            <span>${escapeHtml(unmatchedRequired.map((field) => field.label).join(", "))}. You can still continue and fill these in for each row.</span>
          </div>`
        : ""
    }
    <div class="wizard-actions">
      <button class="button primary" data-import-action="review" type="button">Continue</button>
      <button class="button ghost" data-import-action="cancel" type="button">Cancel</button>
    </div>
  `;
}

function currentMissing(entry) {
  return visibleMissingFields(importSession.module, entry.record).filter((field) => field.name !== "currency");
}

function importCounts() {
  const entries = importSession.entries;
  const complete = entries.filter((entry) => !currentMissing(entry).length).length;
  return {
    total: entries.length,
    complete,
    incomplete: entries.length - complete,
    duplicates: entries.filter((entry) => entry.duplicateId).length,
  };
}

function fillInputHtml(index, field, value) {
  const data = `data-fill="${index}:${escapeHtml(field.name)}"`;
  let control;
  if (field.type === "select") {
    control = `<select ${data}><option value="">—</option>${selectOptions(field)
      .map(
        (option) =>
          `<option value="${escapeHtml(option.value)}" ${String(option.value) === String(value ?? "") ? "selected" : ""}>${escapeHtml(option.label)}</option>`,
      )
      .join("")}</select>`;
  } else {
    const type = { number: "number", date: "date", email: "email" }[field.type] || "text";
    control = `<input ${data} type="${type}" ${type === "number" ? 'step="any"' : ""} value="${escapeHtml(value ?? "")}" placeholder="${field.required ? "Required" : "Optional"}" />`;
  }
  return `<label class="fill-field"><span>${escapeHtml(field.label)}${field.required ? " *" : ""}</span>${control}</label>`;
}

function fillStatusHtml(entry) {
  const missing = currentMissing(entry);
  return missing.length ? badge(`${missing.length} missing`, "amber") : badge("Complete", "green");
}

function importReviewHtml() {
  const s = importSession;
  const counts = importCounts();
  const fieldsByName = Object.fromEntries(importableFields(s.module).map((field) => [field.name, field]));
  const askEntries = s.entries.filter((entry) => entry.askFields.length);
  const shown = askEntries.slice(0, 100);
  const canUpdate = can(s.module, "edit");

  const fillRows = shown
    .map((entry) => {
      const index = s.entries.indexOf(entry);
      const inputs = entry.askFields
        .map((name) => fieldsByName[name])
        .filter(Boolean)
        .map((field) => fillInputHtml(index, field, entry.record[field.name]))
        .join("");
      return `
      <div class="fill-row">
        <div class="fill-row-head">
          <strong>Row ${formatNumber(entry.rowNumber)} · ${escapeHtml(recordLabel(s.module, entry.record))}</strong>
          <span data-fill-status="${index}">${fillStatusHtml(entry)}</span>
        </div>
        <div class="fill-fields">${inputs}</div>
      </div>
    `;
    })
    .join("");

  return `
    ${stepsHtml(1)}
    <div class="import-summary">
      <div><strong>${formatNumber(counts.total)}</strong><span>rows read</span></div>
      <div><strong id="importCompleteCount">${formatNumber(counts.complete)}</strong><span>complete</span></div>
      <div><strong id="importIncompleteCount">${formatNumber(counts.incomplete)}</strong><span>missing info</span></div>
      <div><strong>${formatNumber(counts.duplicates)}</strong><span>match existing records</span></div>
    </div>
    ${importCurrencyHtml()}
    ${
      counts.duplicates
        ? `<label class="filter-control inline-control">
            <span>When a row matches an existing record</span>
            <select id="duplicateModeSelect">
              ${canUpdate ? `<option value="update" ${s.duplicateMode === "update" ? "selected" : ""}>Update the existing record</option>` : ""}
              <option value="skip" ${s.duplicateMode === "skip" ? "selected" : ""}>Skip the row</option>
              <option value="add" ${s.duplicateMode === "add" ? "selected" : ""}>Add it as a new record</option>
            </select>
          </label>`
        : ""
    }
    ${
      askEntries.length
        ? `<p class="wizard-intro">Some rows are missing information. Fill in what you know below and leave the rest blank to continue without it. Anything still missing will be listed on the dashboard under <strong>Needs your input</strong>.</p>
           <div class="fill-list">${fillRows}</div>
           ${askEntries.length > shown.length ? `<div class="empty-state">${formatNumber(askEntries.length - shown.length)} more rows with missing info can be completed after importing.</div>` : ""}`
        : `<p class="wizard-intro">Every row has the information LedgerFlow needs.</p>`
    }
    <div class="wizard-actions">
      <button class="button primary" data-import-action="all" type="button">Import all ${formatNumber(counts.total)} rows</button>
      ${
        askEntries.length
          ? `<button class="button ghost" data-import-action="ready" type="button">Import only complete rows (<span id="importReadyLabel">${formatNumber(counts.complete)}</span>)</button>`
          : ""
      }
      <button class="button ghost" data-import-action="mapping" type="button">Back</button>
      <button class="button ghost" data-import-action="cancel" type="button">Cancel</button>
    </div>
  `;
}

function importDoneHtml() {
  const s = importSession;
  const result = s.result;
  return `
    ${stepsHtml(2)}
    <div class="import-summary">
      <div><strong>${formatNumber(result.added)}</strong><span>added</span></div>
      <div><strong>${formatNumber(result.updated)}</strong><span>updated</span></div>
      <div><strong>${formatNumber(result.skipped)}</strong><span>skipped</span></div>
      <div><strong>${formatNumber(result.stillMissing)}</strong><span>still need info</span></div>
    </div>
    <p class="wizard-intro">
      ${escapeHtml(s.fileName)} was imported into ${escapeHtml(viewTitle(s.module))}.
      ${result.currencies ? `Amounts saved in ${escapeHtml(result.currencies)}.` : ""}
      ${result.stillMissing ? "Records that still need information are listed on the dashboard." : ""}
    </p>
    ${
      result.currencyConflicts
        ? `<div class="form-notice">
            <strong>${formatNumber(result.currencyConflicts)} row${result.currencyConflicts === 1 ? " was" : "s were"} not used to update existing records</strong>
            <span>The existing record uses a different currency, so updating it would mix currencies. Edit those records by hand if the currency really changed.</span>
          </div>`
        : ""
    }
    <div class="wizard-actions">
      <button class="button primary" data-view-link="${s.module}" type="button">View ${escapeHtml(viewTitle(s.module))}</button>
      <button class="button ghost" data-import-action="reset" type="button">Import another file</button>
    </div>
  `;
}

function handleImportAction(action) {
  if (!importSession && action !== "reset") return;
  if (action === "cancel" || action === "reset") {
    importSession = null;
    els.importFileInput.value = "";
    els.importFileName.textContent = "No file selected";
    render();
    if (action === "cancel") showToast("Import cancelled");
    return;
  }
  if (action === "review") {
    goToImportReview();
    return;
  }
  if (action === "mapping") {
    importSession.step = "mapping";
    render();
    return;
  }
  if (action === "all" || action === "ready") completeImport(action);
}

async function handleImportFile() {
  const file = els.importFileInput.files?.[0];
  els.importFileName.textContent = file?.name || "No file selected";
  if (!file) return;
  const module = els.importTarget.value;
  if (!module) {
    showToast("Your role cannot import into any module");
    return;
  }
  try {
    const table = await readImportFile(file);
    if (!table.rows.length) {
      showToast("No rows found in that file");
      return;
    }
    importSession = {
      module,
      fileName: file.name,
      note: table.note,
      headers: table.headers,
      rows: table.rows,
      mapping: suggestMapping(module, table.headers),
      duplicateMode: can(module, "edit") ? "update" : "skip",
      step: "mapping",
      entries: [],
    };
    render();
  } catch (error) {
    console.error(error);
    markImportFailed(error.message);
    importSession = null;
    render();
    showToast(error.message || "Could not read that file");
  }
}

async function readImportFile(file) {
  const extension = file.name.split(".").pop().toLowerCase();
  if (extension === "xlsx" || extension === "xls") return readSpreadsheet(await file.arrayBuffer());
  const text = (await file.text()).replace(/^\uFEFF/, "");
  if (extension === "json") return tableFromRows(jsonToRows(text));
  return tableFromRows(parseDelimited(text, detectDelimiter(text)));
}

function loadSheetJs() {
  if (globalThis.XLSX) return Promise.resolve(globalThis.XLSX);
  sheetJsPromise ||= new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = SHEETJS_URL;
    script.onload = () => resolve(globalThis.XLSX);
    script.onerror = () => {
      sheetJsPromise = null;
      script.remove();
      reject(new Error("The Excel reader needs an internet connection. Connect, or save the sheet as CSV."));
    };
    document.head.append(script);
  });
  return sheetJsPromise;
}

async function readSpreadsheet(buffer) {
  const XLSX = await loadSheetJs();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const sheets = workbook.SheetNames.map((name) => ({
    name,
    rows: XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, raw: true, defval: "" }),
  })).filter((sheet) => sheet.rows.some((row) => row.some((cell) => !isBlank(cell))));
  if (!sheets.length) return { headers: [], rows: [] };
  const table = tableFromRows(sheets[0].rows);
  table.note = sheets.length > 1 ? `sheet "${sheets[0].name}"; ${sheets.length - 1} other sheet(s) not read` : `sheet "${sheets[0].name}"`;
  return table;
}

function jsonToRows(text) {
  let data = JSON.parse(text);
  if (!Array.isArray(data) && data && typeof data === "object") {
    data = Object.values(data).find(Array.isArray);
  }
  if (!Array.isArray(data)) throw new Error("The JSON file must contain a list of records.");
  if (data.every(Array.isArray)) return data;
  const headers = [...new Set(data.flatMap((item) => (item && typeof item === "object" ? Object.keys(item) : [])))];
  return [
    headers,
    ...data.map((item) =>
      headers.map((header) => {
        const value = item?.[header];
        return value && typeof value === "object" ? JSON.stringify(value) : value ?? "";
      }),
    ),
  ];
}

function detectDelimiter(text) {
  const firstLine = text.split(/\r?\n/).find((line) => line.trim()) || "";
  const counts = ["\t", ";", ","].map((delimiter) => [delimiter, firstLine.split(delimiter).length - 1]);
  counts.sort((a, b) => b[1] - a[1]);
  return counts[0][1] > 0 ? counts[0][0] : ",";
}

function parseDelimited(text, delimiter) {
  const rows = [];
  let row = [];
  let value = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (inQuotes) {
      if (char === '"' && next === '"') {
        value += '"';
        index += 1;
      } else if (char === '"') {
        inQuotes = false;
      } else {
        value += char;
      }
      continue;
    }
    if (char === '"') {
      inQuotes = true;
    } else if (char === delimiter) {
      row.push(value);
      value = "";
    } else if (char === "\n") {
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
    } else if (char !== "\r") {
      value += char;
    }
  }
  if (value || row.length) {
    row.push(value);
    rows.push(row);
  }
  return rows;
}

function tableFromRows(rawRows) {
  const lines = rawRows
    .map((cells, index) => ({
      line: index + 1,
      cells: cells.map((cell) => (cell instanceof Date || typeof cell === "number" ? cell : String(cell ?? "").trim())),
    }))
    .filter((row) => row.cells.some((cell) => !isBlank(cell)));
  if (!lines.length) return { headers: [], rows: [] };

  const filled = (row) => row.cells.filter((cell) => !isBlank(cell)).length;
  const widest = Math.max(...lines.slice(0, 30).map(filled));
  const headerIndex = Math.max(
    0,
    lines.slice(0, 10).findIndex((row) => filled(row) >= Math.max(2, Math.ceil(widest * 0.6))),
  );
  const columnCount = Math.max(...lines.map((row) => row.cells.length));
  const headers = Array.from(
    { length: columnCount },
    (_, index) => String(formatSample(lines[headerIndex].cells[index]) || "").trim() || `Column ${index + 1}`,
  );
  return { headers, rows: lines.slice(headerIndex + 1).map((row) => ({ rowNumber: row.line, cells: row.cells })) };
}

function normalizeHeader(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function suggestMapping(module, headers) {
  const fields = importableFields(module);
  const used = new Set();
  return headers.map((header) => {
    const key = normalizeHeader(header);
    const match = fields.find(
      (field) =>
        !used.has(field.name) &&
        [normalizeHeader(field.name), normalizeHeader(field.label), ...(importAliases[module]?.[field.name] || [])].includes(key),
    );
    if (!match) return "";
    used.add(match.name);
    return match.name;
  });
}

function createFieldFromColumn(session, index) {
  const label = session.headers[index];
  const existing = getFields(session.module, { includeHidden: true }).find(
    (field) => field.label.toLowerCase() === label.toLowerCase(),
  );
  if (existing) return existing.name;
  const values = session.rows
    .map((row) => row.cells[index])
    .filter((value) => !isBlank(value))
    .slice(0, 50);
  const looksNumeric = (value) =>
    typeof value === "number" || (/^-?[\d,]+(\.\d+)?$/.test(String(value).trim()) && !/^0\d/.test(String(value).trim()));
  const looksDate = (value) => value instanceof Date || Boolean(parseImportDate(String(value)));
  let type = "text";
  if (values.length && values.every(looksDate) && !values.every(looksNumeric)) type = "date";
  else if (values.length && values.every(looksNumeric)) type = "number";
  const field = { name: uniqueFieldName(session.module, label), label, type, options: [], required: false };
  state.customFields[session.module].push(field);
  addActivity(`Field added from import: ${label} (${viewTitle(session.module)})`, "setup");
  return field.name;
}

function goToImportReview() {
  const s = importSession;
  s.mapping = s.mapping.map((value, index) => (value === "__new__" ? createFieldFromColumn(s, index) : value));
  s.entries = buildImportEntries(s);
  s.step = "review";
  saveState();
  render();
}

function buildImportEntries(s) {
  const fields = importableFields(s.module);
  const now = new Date().toISOString();
  return s.rows.map((row) => {
    const values = {};
    s.mapping.forEach((fieldName, column) => {
      const field = fields.find((entry) => entry.name === fieldName);
      if (!field) return;
      const value = normalizeImportValue(field, row.cells[column]);
      if (values[fieldName] === undefined || (isBlank(values[fieldName]) && !isBlank(value))) {
        values[fieldName] = value;
      }
    });
    if (s.module === "finance") inferFinanceType(values, s, row);
    let dollarSymbolOnly = false;
    if (isBlank(values.currency)) {
      // Look for a currency written in the money cells ("Rs 1,250", "$40") or their headers ("Price (USD)").
      const moneyColumns = s.headers
        .map((header, index) => ({ header, index }))
        .filter(
          ({ header, index }) =>
            MONEY_FIELDS[s.module].includes(s.mapping[index]) ||
            /amount|price|cost|salary|pay|balance|value|debit|credit|total|rate/i.test(header),
        );
      for (const { header, index } of moneyColumns) {
        const cellText = String(row.cells[index] ?? "");
        const detected = currencyFromText(cellText) || currencyFromText(header);
        if (detected) {
          values.currency = detected;
          dollarSymbolOnly = detected === "USD" && !/USD|DOLLAR/i.test(`${cellText} ${header}`);
          break;
        }
      }
    }
    const record = { ...values, importedAt: now, sourceName: `File: ${s.fileName}` };
    applyImportDefaults(s.module, record, fields);
    return {
      rowNumber: row.rowNumber,
      record,
      dollarSymbolOnly,
      askFields: visibleMissingFields(s.module, record)
        .filter((field) => field.name !== "currency")
        .map((field) => field.name),
      duplicateId: findDuplicate(s.module, record)?.id || null,
    };
  });
}

function importHasAmounts(s) {
  return s.entries.some((entry) => MONEY_FIELDS[s.module].some((name) => !isBlank(entry.record[name])));
}

function importNeedsCurrencyChoice(s) {
  return importHasAmounts(s) && s.entries.some((entry) => isBlank(entry.record.currency));
}

function importCurrencyHtml() {
  const s = importSession;
  if (!importHasAmounts(s)) return "";
  const detected = {};
  let undetected = 0;
  s.entries.forEach((entry) => {
    if (isBlank(entry.record.currency)) undetected += 1;
    else detected[entry.record.currency] = (detected[entry.record.currency] || 0) + 1;
  });
  const detectedText = Object.entries(detected)
    .map(([code, count]) => `${formatNumber(count)} row${count === 1 ? "" : "s"} in ${code}`)
    .join(", ");
  const dollarNote = s.entries.some((entry) => entry.dollarSymbolOnly)
    ? ` A "$" sign was read as US dollars. If these are another dollar, add a Currency column to the file.`
    : "";

  if (!undetected) {
    return `
      <div class="currency-confirm is-ok">
        <strong>Currency of amounts</strong>
        <span>Read from the file: ${escapeHtml(detectedText)}. Amounts are saved in these currencies and never converted.${escapeHtml(dollarNote)}</span>
      </div>
    `;
  }
  return `
    <div class="currency-confirm ${s.fileCurrency ? "is-ok" : "needs-choice"}" id="importCurrencyBox">
      <strong>What currency are these amounts in?</strong>
      <span>
        ${detectedText ? `Read from the file: ${escapeHtml(detectedText)}. ` : ""}${formatNumber(undetected)} row${
          undetected === 1 ? " doesn't say" : "s don't say"
        } which currency ${undetected === 1 ? "it uses" : "they use"}. Choose the currency the file was recorded in, even if it differs from your default (${escapeHtml(
          orgCurrency(),
        )}). LedgerFlow never converts amounts.${escapeHtml(dollarNote)}
      </span>
      <label class="filter-control inline-control">
        <span>Amounts without a currency are in</span>
        <select id="importCurrencySelect">
          <option value="">Choose currency…</option>
          ${currencyOptions()
            .map((code) => `<option value="${code}" ${code === s.fileCurrency ? "selected" : ""}>${code}</option>`)
            .join("")}
        </select>
      </label>
    </div>
  `;
}

function inferFinanceType(values, s, row) {
  const valueFor = (names) => {
    const index = s.headers.findIndex((header) => names.includes(normalizeHeader(header)));
    return index >= 0 ? parseImportNumber(String(row.cells[index] ?? "")) : NaN;
  };
  const debit = valueFor(["debit", "dr", "withdrawal", "paid", "payment"]);
  const credit = valueFor(["credit", "cr", "deposit", "received", "receipt"]);
  if (isBlank(values.amount)) {
    const fromColumns = Math.abs(debit) > 0 ? debit : Math.abs(credit) > 0 ? credit : NaN;
    if (Number.isFinite(fromColumns)) values.amount = Math.abs(fromColumns);
  }
  if (!isBlank(values.type)) return;
  const amountIndex = s.mapping.indexOf("amount");
  const signedAmount = amountIndex >= 0 ? parseImportNumber(String(row.cells[amountIndex] ?? "")) : NaN;
  if (Math.abs(debit) > 0) values.type = "Expense";
  else if (Math.abs(credit) > 0) values.type = "Income";
  else if (signedAmount < 0) values.type = "Expense";
}

function applyImportDefaults(module, record, fields) {
  fields.forEach((field) => {
    if (!isBlank(record[field.name]) || field.required || field.custom) return;
    if (field.name === "accessRole") record.accessRole = fallbackRoleId();
    else if (field.type === "select" && field.options?.length) record[field.name] = field.options[0];
  });
  if (module === "employees") {
    if (!assignableRoles().some((role) => role.id === record.accessRole)) record.accessRole = fallbackRoleId();
    record.status ||= "Active";
  }
}

function findDuplicate(module, record) {
  const same = (a, b) => !isBlank(a) && !isBlank(b) && String(a).trim().toLowerCase() === String(b).trim().toLowerCase();
  const keysByModule = {
    inventory: [["sku"], ["name"]],
    assets: [["code"], ["name"]],
    employees: [["loginEmail"], ["email"], ["phone"], ["name"]],
    contacts: [["email"], ["phone"], ["name"]],
    finance: [["date", "amount", "description"]],
  };
  for (const keys of keysByModule[module]) {
    if (keys.some((key) => isBlank(record[key]))) continue;
    const match = state[module].find((existing) => keys.every((key) => same(existing[key], record[key])));
    if (match) return match;
    if (keys.length === 1 && module !== "finance") return null;
  }
  return null;
}

function normalizeImportValue(field, raw) {
  if (raw instanceof Date) return dateToIso(raw);
  const text = String(raw ?? "").trim();
  if (!text) return "";
  if (field.type === "number") {
    const number = parseImportNumber(text);
    if (!Number.isFinite(number)) return "";
    return field.min === 0 ? Math.abs(number) : number;
  }
  if (field.type === "date") return parseImportDate(text);
  if (field.type === "email") return text.toLowerCase();
  if (field.name === "currency") {
    const code = text.toUpperCase();
    return /^[A-Z]{3}$/.test(code) && isValidCurrency(code) ? code : currencyFromText(text);
  }
  if (field.type === "select") {
    if (field.options?.includes("Supplier")) {
      const key = normalizeHeader(text);
      if (["vendor", "supplier", "seller", "distributor", "creditor"].includes(key)) return "Supplier";
      if (["customer", "client", "buyer", "debtor"].includes(key)) return "Customer";
    }
    const option = selectOptions(field).find(
      (entry) => normalizeHeader(entry.value) === normalizeHeader(text) || normalizeHeader(entry.label) === normalizeHeader(text),
    );
    return option ? option.value : "";
  }
  return text;
}

function parseImportNumber(text) {
  const value = String(text).trim();
  const match = value.replace(/,/g, "").match(/\d+(?:\.\d+)?|\.\d+/);
  if (!match) return NaN;
  const negative = /^\(.*\)$/.test(value) || /^[^\d]*-\s*[\d.]/.test(value) || /\d\s*-$/.test(value);
  return negative ? -Number(match[0]) : Number(match[0]);
}

function dateToIso(date) {
  const shifted = new Date(date.getTime() + 12 * 60 * 60 * 1000);
  return isoFromParts(shifted.getUTCFullYear(), shifted.getUTCMonth() + 1, shifted.getUTCDate());
}

function parseImportDate(text) {
  const value = String(text).trim();
  let match = value.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (match) return isoFromParts(Number(match[1]), Number(match[2]), Number(match[3]));

  match = value.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4}|\d{2})\b/);
  if (match) {
    const first = Number(match[1]);
    const second = Number(match[2]);
    let year = Number(match[3]);
    if (year < 100) year += 2000;
    let day = first;
    let month = second;
    if (first > 12) {
      day = first;
      month = second;
    } else if (second > 12 || state.organization?.dateFormat === "MDY") {
      month = first;
      day = second;
    }
    return isoFromParts(year, month, day);
  }

  if (/^\d{5}$/.test(value)) {
    const serial = Number(value);
    if (serial > 20000 && serial < 80000) return dateToIso(new Date(Date.UTC(1899, 11, 30) + serial * 86400000));
  }

  if (/[a-z]/i.test(value)) {
    const parsed = new Date(value);
    if (!Number.isNaN(parsed.getTime())) {
      return isoFromParts(parsed.getFullYear(), parsed.getMonth() + 1, parsed.getDate());
    }
  }
  return "";
}

function handleImportFill(input) {
  if (!importSession) return;
  const [indexText, name] = input.dataset.fill.split(":");
  const entry = importSession.entries[Number(indexText)];
  const field = importableFields(importSession.module).find((entry) => entry.name === name);
  if (!entry || !field) return;
  entry.record[name] = normalizeImportValue(field, input.value);
  entry.duplicateId = findDuplicate(importSession.module, entry.record)?.id || null;
  const status = document.querySelector(`[data-fill-status="${indexText}"]`);
  if (status) status.innerHTML = fillStatusHtml(entry);
  const counts = importCounts();
  const complete = byId("importCompleteCount");
  const incomplete = byId("importIncompleteCount");
  const ready = byId("importReadyLabel");
  if (complete) complete.textContent = formatNumber(counts.complete);
  if (incomplete) incomplete.textContent = formatNumber(counts.incomplete);
  if (ready) ready.textContent = formatNumber(counts.complete);
}

function completeImport(mode) {
  const s = importSession;
  const module = s.module;
  if (importNeedsCurrencyChoice(s) && !isValidCurrency(s.fileCurrency)) {
    const box = byId("importCurrencyBox");
    box?.classList.add("needs-choice", "is-flagged");
    box?.scrollIntoView({ block: "center", behavior: "smooth" });
    byId("importCurrencySelect")?.focus();
    showToast("Choose the currency of the amounts before importing");
    return;
  }
  const now = new Date().toISOString();
  let added = 0;
  let updated = 0;
  let skipped = 0;
  let currencyConflicts = 0;
  const touched = [];

  s.entries.forEach((entry) => {
    if (mode === "ready" && currentMissing(entry).length) {
      skipped += 1;
      return;
    }
    const hasAmounts = MONEY_FIELDS[module].some((name) => !isBlank(entry.record[name]));
    const currency = !isBlank(entry.record.currency)
      ? entry.record.currency
      : hasAmounts || importHasAmounts(s)
        ? s.fileCurrency || orgCurrency()
        : orgCurrency();
    const incoming = { ...entry.record, currency };
    const duplicate = entry.duplicateId ? state[module].find((record) => record.id === entry.duplicateId) : null;
    if (duplicate && s.duplicateMode === "skip") {
      skipped += 1;
      return;
    }
    if (duplicate && s.duplicateMode === "update" && can(module, "edit")) {
      if (hasAmounts && recordCurrency(duplicate) !== currency) {
        currencyConflicts += 1;
        skipped += 1;
        return;
      }
      Object.entries(incoming).forEach(([key, value]) => {
        if (key !== "currency" && !isBlank(value)) duplicate[key] = value;
      });
      duplicate.importedAt = now;
      touched.push(duplicate);
      updated += 1;
      return;
    }
    if (!can(module, "add")) {
      skipped += 1;
      return;
    }
    const record = { ...incoming, id: makeId(module), importedAt: now };
    state[module].unshift(record);
    touched.push(record);
    added += 1;
  });

  s.result = {
    added,
    updated,
    skipped,
    currencyConflicts,
    currencies: importHasAmounts(s) ? [...new Set(touched.map(recordCurrency))].join(", ") : "",
    stillMissing: touched.filter((record) => needsInfo(module, record)).length,
  };
  const source = state.dataSources.find((entry) => entry.id === "source-file");
  if (source) {
    source.status = "Fresh";
    source.lastSync = now;
    source.records = Number(source.records || 0) + added + updated;
    source.target = viewTitle(module);
    source.note = `Last import: ${s.fileName} (${added} added, ${updated} updated) into ${viewTitle(module)}.`;
  }
  addActivity(`Imported ${s.fileName}: ${added} added, ${updated} updated`, module);
  s.step = "done";
  saveState();
  render();
  showToast(`${added + updated} record${added + updated === 1 ? "" : "s"} imported`);
}

function markImportFailed(message) {
  const source = state.dataSources.find((entry) => entry.id === "source-file");
  if (!source) return;
  source.status = "Failed";
  source.note = message || "The last import could not be completed.";
  source.lastSync = new Date().toISOString();
  saveState();
}

function downloadSampleCsv() {
  const target = els.importTarget.value || "inventory";
  const samples = {
    inventory:
      "Item Name,SKU,Category,Qty,Reorder Level,Cost Price,Sale Price,Location,Brand\nLED Tube 18W,ELC-TUBE-18,Electrical,40,20,4.5,6.75,Aisle 1,Philips\nSafety Gloves,,Safety,12,25,,2.45,Aisle 2,",
    assets:
      "Asset Name,Asset Code,Assigned To,Location,Purchase Date,Value,Condition,Notes\nGenerator,AST-GEN-001,Operations,Warehouse,10/01/2026,3200,Active,Backup power",
    employees:
      "Worker Name,Designation,Division,Mobile,Email,Pay,Status\nSara Ahmed,Sales Associate,Sales,+92 300 1234567,sara@example.com,2600,Active\nBilal Khan,Cashier,,,,1900,Active",
    finance:
      "Date,Particulars,Category,Debit,Credit,Status\n14/09/2026,Counter sales,Retail Sales,,1250,Paid\n15/09/2026,Electricity bill,Utilities,320,,Pending",
    contacts:
      "Party Name,Type,Phone,Email,Balance,Status\nBlue Star Trading,Customer,+92 321 5550155,accounts@bluestar.example,2200,Pending\nMetro Wholesale,Vendor,,,(1850),Active",
  };
  const blob = new Blob([samples[target]], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `ledgerflow-${target}-sample.csv`;
  link.click();
  URL.revokeObjectURL(url);
  showToast("Sample CSV downloaded");
}

/* ---------- Reports ---------- */

function renderReports() {
  const summary = totals();
  const alerts = alertSummary();
  const sensitive = canAccess("sensitiveNumbers");
  const cards = [];
  if (sensitive && can("finance", "view")) cards.push(["Net Balance", formatMoneyTotals(summary.net)]);
  if (can("inventory", "view")) {
    cards.push(
      sensitive
        ? [`${viewTitle("inventory")} Value`, formatMoneyTotals(summary.inventoryValue)]
        : [`${viewTitle("inventory")} Tracked`, formatNumber(state.inventory.length)],
    );
  }
  if (sensitive && can("assets", "view")) cards.push([`${viewTitle("assets")} Value`, formatMoneyTotals(summary.assetValue)]);
  if (sensitive && can("employees", "view")) cards.push(["Monthly Payroll", formatMoneyTotals(summary.payroll)]);
  if (can("contacts", "view")) cards.push(["Owed To You", formatMoneyTotals(summary.receivables)]);
  if (!sensitive && can("employees", "view")) cards.push(["Employees", formatNumber(state.employees.length)]);

  byId("reportCards").innerHTML = cards
    .map(
      ([label, value]) => `
      <article class="report-card">
        <span>${escapeHtml(label)}</span>
        <strong>${escapeHtml(value)}</strong>
      </article>
    `,
    )
    .join("");

  const lines = [];
  if (can("inventory", "view")) {
    if (tracksReorderLevel()) {
      lines.push([
        "Stock attention",
        `${alerts.lowStock} record${alerts.lowStock === 1 ? " is" : "s are"} at or below reorder level.`,
      ]);
    }
    if (sensitive) {
      const { currency, excluded } = chartCurrency(state.inventory);
      const topCategory = Object.entries(
        state.inventory
          .filter((item) => recordCurrency(item) === currency)
          .reduce((acc, item) => {
            const category = item.category || "Uncategorized";
            acc[category] = (acc[category] || 0) + parseMoney(item.quantity) * parseMoney(item.unitCost);
            return acc;
          }, {}),
      ).sort((a, b) => b[1] - a[1])[0];
      lines.push([
        `Top ${fieldLabel("inventory", "category").toLowerCase()}`,
        topCategory
          ? `${topCategory[0]} holds ${formatCurrency(topCategory[1], currency)} in stock value${
              excluded ? ` (comparing ${currency} records only)` : ""
            }.`
          : "No stock recorded yet.",
      ]);
    }
  }
  if (can("assets", "view")) {
    lines.push([
      "Asset attention",
      `${alerts.maintenance} asset${alerts.maintenance === 1 ? " needs" : "s need"} maintenance or condition review.`,
    ]);
  }
  if (can("contacts", "view")) {
    lines.push([
      "Dues",
      `Customers owe ${formatMoneyTotals(summary.receivables)}. You owe vendors ${formatMoneyTotals(summary.payables)}.`,
    ]);
  }
  if (can("finance", "view")) {
    lines.push(["Unpaid transactions", `${alerts.unpaid} transaction${alerts.unpaid === 1 ? " is" : "s are"} pending or overdue.`]);
  }
  lines.push([
    "Data completeness",
    alerts.missingInfo
      ? `${alerts.missingInfo} record${alerts.missingInfo === 1 ? " is" : "s are"} missing information.`
      : "No records are missing required information.",
  ]);

  byId("reportBody").innerHTML = lines
    .map(
      ([title, detail]) => `
      <div class="report-line">
        <strong>${escapeHtml(title)}</strong>
        <span>${escapeHtml(detail)}</span>
      </div>
    `,
    )
    .join("");
}

/* ---------- Data tools ---------- */

function exportData() {
  if (!canAccess("export")) return;
  const safe = structuredClone(state);
  delete safe.organization.ownerPasswordHash;
  delete safe.organization.ownerPasswordSalt;
  safe.employees.forEach((employee) => {
    delete employee.passwordHash;
    delete employee.passwordSalt;
  });
  const blob = new Blob([JSON.stringify(safe, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `ledgerflow-export-${todayIso()}.json`;
  link.click();
  URL.revokeObjectURL(url);
  showToast("Data exported (passwords excluded)");
}

function loadSampleBusiness(type) {
  const build = globalThis.LEDGERFLOW_SAMPLES?.[type];
  if (!build || !currentUser()?.isOwner) return;
  const confirmText = totalRecordCount()
    ? `Replace ALL current records with the ${type} sample? Module names and fields will be arranged for a ${type}. Your login and roles stay.`
    : `Load the ${type} sample? Module names and fields will be arranged for a ${type}.`;
  if (!window.confirm(confirmText)) return;

  MODULES.forEach((module) => {
    state[module] = [];
  });
  state.activity = [];
  state.dataSources = defaultDataSources();
  importSession = null;
  if (!state.organization.sampleLoaded) state.organization.typeBeforeSample = state.organization.type;
  applyBusinessScope(type, { silent: true, replace: true });

  const now = new Date().toISOString();
  const sample = build({ daysAgo: daysAgoIso, staffRole: fallbackRoleId() });
  let count = 0;
  MODULES.forEach((module) => {
    const customByLabel = Object.fromEntries(
      state.customFields[module].map((field) => [field.label.toLowerCase(), field.name]),
    );
    state[module] = (sample[module] || []).map((entry) => {
      const { extra = {}, imported = false, salaryPaid = false, ...base } = entry;
      const record = {
        ...base,
        currency: base.currency || sample.currency || orgCurrency(),
        id: makeId(module),
        sourceName: `Sample: ${type}`,
        createdAt: now,
      };
      if (module === "contacts" && record.type === "Vendor") record.type = "Supplier";
      Object.entries(extra).forEach(([label, value]) => {
        const name = customByLabel[label.toLowerCase()];
        if (name) record[name] = value;
      });
      if (imported) record.importedAt = now;
      else record.infoDismissed = true;
      if (salaryPaid) record.salaryPaidInSample = true;
      count += 1;
      return record;
    });
  });

  // Some sample staff are already paid for this month, so payroll shows paid and unpaid rows.
  const month = todayIso().slice(0, 7);
  state.employees
    .filter((employee) => employee.salaryPaidInSample)
    .forEach((employee) => {
      delete employee.salaryPaidInSample;
      if (!salaryIsPayable(employee)) return;
      state.finance.unshift({
        id: makeId("finance"),
        date: todayIso(),
        type: "Expense",
        category: "Salaries",
        description: `Salary - ${employee.name} - ${monthName(month, "long")}`,
        amount: parseMoney(employee.salary),
        currency: recordCurrency(employee),
        status: "Paid",
        party: employee.name,
        employeeId: employee.id,
        payrollMonth: month,
        sourceName: `Sample: ${type}`,
        createdAt: now,
        infoDismissed: true,
      });
      count += 1;
    });

  state.organization.sampleLoaded = type;
  delete state.organization.hideGettingStarted;
  const manual = state.dataSources.find((source) => source.id === "source-manual");
  if (manual) {
    manual.status = "Fresh";
    manual.lastSync = now;
    manual.records = count;
  }
  addActivity(`${type} sample loaded (${count} records)`, "setup");
  saveState();
  fillSetupForm();
  setView("dashboard");
  showToast(`${type} sample loaded`);
}

function clearSampleData() {
  const org = state.organization;
  if (!org.sampleLoaded) return;
  if (!window.confirm("Remove all sample records and return to your empty business?")) return;
  const type = org.typeBeforeSample || org.type;
  MODULES.forEach((module) => {
    state[module] = [];
  });
  state.activity = [];
  state.dataSources = defaultDataSources();
  importSession = null;
  applyBusinessScope(type, { silent: true, replace: true });
  delete org.sampleLoaded;
  delete org.typeBeforeSample;
  saveState();
  fillSetupForm();
  setView("dashboard");
  showToast("Sample data removed");
}

function clearAllRecords({ confirmText } = {}) {
  const message =
    confirmText ||
    "Delete ALL records (items, transactions, contacts, employees, assets)? Roles, fields, and the owner login stay.";
  if (!window.confirm(message)) return;
  MODULES.forEach((module) => {
    state[module] = [];
  });
  state.dataSources = defaultDataSources();
  importSession = null;
  delete state.organization.sampleLoaded;
  addActivity("All records cleared", "setup");
  saveState();
  render();
  showToast("All records cleared");
}

/* ---------- Events ---------- */

els.registerForm.addEventListener("submit", handleRegister);
els.loginForm.addEventListener("submit", handleLogin);
els.resetBrowserBtn.addEventListener("click", resetBrowserData);
els.logoutBtn.addEventListener("click", () => {
  logout();
  showToast("Signed out");
});

els.navList.addEventListener("click", (event) => {
  const button = event.target.closest("[data-view]");
  if (button) setView(button.dataset.view);
});

document.addEventListener("click", (event) => {
  const target = event.target;
  const split = (element, attribute) => element.getAttribute(attribute).split(":");

  const tabButton = target.closest("[data-tab]");
  if (tabButton) {
    const [view, key] = split(tabButton, "data-tab");
    openTab(view, key);
    return;
  }
  const openTabButton = target.closest("[data-open-tab]");
  if (openTabButton) {
    const [view, key] = split(openTabButton, "data-open-tab");
    openTab(view, key);
    return;
  }
  const financeTabButton = target.closest("[data-finance-tab]");
  if (financeTabButton) {
    openTab("finance", financeTabButton.dataset.financeTab);
    return;
  }
  const periodButton = target.closest("[data-finance-period]");
  if (periodButton) {
    financeViewState.period = periodButton.dataset.financePeriod;
    render();
    return;
  }
  const periodRow = target.closest("[data-period-row]");
  if (periodRow) {
    financeViewState.monthFilter = periodRow.dataset.periodRow;
    openTab("finance", "all");
    return;
  }
  const markPaidButton = target.closest("[data-mark-paid]");
  if (markPaidButton) {
    markPaid(markPaidButton.dataset.markPaid);
    return;
  }
  const paymentButton = target.closest("[data-record-payment]");
  if (paymentButton) {
    recordPayment(paymentButton.dataset.recordPayment);
    return;
  }
  const salaryButton = target.closest("[data-pay-salary]");
  if (salaryButton) {
    const [employeeId, month] = split(salaryButton, "data-pay-salary");
    paySalary(employeeId, month);
    return;
  }
  if (target.closest("#payAllSalariesBtn")) {
    payAllSalaries(payrollMonth || todayIso().slice(0, 7));
    return;
  }

  const viewLink = target.closest("[data-view-link]");
  if (viewLink) {
    setView(viewLink.dataset.viewLink);
    return;
  }
  const addButton = target.closest("[data-add]");
  if (addButton) {
    openModal(addButton.dataset.add);
    return;
  }
  const editButton = target.closest("[data-edit]");
  if (editButton) {
    const [module, id] = split(editButton, "data-edit");
    openModal(module, id);
    return;
  }
  const deleteButton = target.closest("[data-delete]");
  if (deleteButton) {
    const [module, id] = split(deleteButton, "data-delete");
    deleteRecord(module, id);
    return;
  }
  const dismissButton = target.closest("[data-dismiss-info]");
  if (dismissButton) {
    const [module, id] = split(dismissButton, "data-dismiss-info");
    dismissInfo(module, id);
    return;
  }
  const importAction = target.closest("[data-import-action]");
  if (importAction) {
    handleImportAction(importAction.dataset.importAction);
    return;
  }
  const previewButton = target.closest("[data-preview-role]");
  if (previewButton) {
    previewRoleId = previewButton.dataset.previewRole;
    setView("dashboard");
    showToast(`Previewing ${roleName(previewRoleId)}`);
    return;
  }
  const removeRoleButton = target.closest("[data-remove-role]");
  if (removeRoleButton) {
    removeRole(removeRoleButton.dataset.removeRole);
    return;
  }
  const editFieldButton = target.closest("[data-edit-field]");
  if (editFieldButton) {
    const [module, name] = split(editFieldButton, "data-edit-field");
    startFieldEdit(module, name);
    return;
  }
  const toggleFieldButton = target.closest("[data-toggle-field]");
  if (toggleFieldButton) {
    const [module, name] = split(toggleFieldButton, "data-toggle-field");
    toggleFieldHidden(module, name);
    return;
  }
  const removeFieldButton = target.closest("[data-remove-field]");
  if (removeFieldButton) {
    const [module, name] = split(removeFieldButton, "data-remove-field");
    removeCustomField(module, name);
    return;
  }
  const scopeButton = target.closest("[data-apply-scope]");
  if (scopeButton) {
    const name = scopeButton.dataset.applyScope;
    if (
      window.confirm(
        `Arrange LedgerFlow for a ${name}? Module and field names change to suit that business. Your records and the fields you added stay.`,
      )
    ) {
      applyBusinessScope(name);
    }
    return;
  }
  const sampleButton = target.closest("[data-load-sample]");
  if (sampleButton) {
    const select = byId(sampleButton.dataset.loadSample);
    if (select) loadSampleBusiness(select.value);
    return;
  }
  if (target.id === "cancelFormBtn") closeModal();
});

document.addEventListener("change", (event) => {
  const target = event.target;
  if (target.id === "financeYearSelect") {
    financeViewState.year = Number(target.value);
    render();
    return;
  }
  if (target.id === "financeCurrencySelect") {
    financeViewState.currency = target.value;
    render();
    return;
  }
  if (target.id === "financeMonthFilter") {
    financeViewState.monthFilter = target.value;
    render();
    return;
  }
  if (target.id === "payrollMonthSelect") {
    payrollMonth = target.value;
    render();
    return;
  }
  if (target.matches("[data-role-module]")) toggleModulePermission(target);
  else if (target.matches("[data-role-general]")) toggleGeneralPermission(target);
  else if (target.matches("[data-map-index]") && importSession) {
    importSession.mapping[Number(target.dataset.mapIndex)] = target.value;
    renderImportWizard();
  } else if (target.id === "duplicateModeSelect" && importSession) {
    importSession.duplicateMode = target.value;
  } else if (target.id === "importCurrencySelect" && importSession) {
    importSession.fileCurrency = target.value;
    const box = byId("importCurrencyBox");
    box?.classList.toggle("needs-choice", !target.value);
    box?.classList.toggle("is-ok", Boolean(target.value));
    box?.classList.remove("is-flagged");
  } else if (target.matches("[data-fill]")) {
    handleImportFill(target);
  }
});

document.addEventListener("input", (event) => {
  if (event.target.matches("[data-fill]")) handleImportFill(event.target);
});

els.globalSearch.addEventListener("input", render);
["inventoryCategoryFilter", "employeeDepartmentFilter", "financeCategoryFilter", "financeStatusFilter"].forEach((id) =>
  byId(id).addEventListener("change", render),
);
els.previewRoleSelect.addEventListener("change", () => {
  previewRoleId = els.previewRoleSelect.value || null;
  setView(canOpenView(currentView) ? currentView : "dashboard");
  showToast(previewRoleId ? `Previewing ${roleName(previewRoleId)}` : "Back to owner view");
});
els.exitPreviewBtn.addEventListener("click", () => {
  previewRoleId = null;
  render();
  showToast("Back to owner view");
});
els.setupForm.addEventListener("submit", saveSetup);
els.addRoleBtn.addEventListener("click", addRole);
els.fieldModuleSelect.addEventListener("change", () => {
  resetFieldEditor();
  renderFieldList();
});
els.fieldTypeInput.addEventListener("change", syncFieldOptionsControl);
els.fieldEditor.addEventListener("submit", saveFieldFromEditor);
els.cancelFieldEditBtn.addEventListener("click", () => {
  resetFieldEditor();
  renderFieldList();
});
els.clearRecordsBtn.addEventListener("click", () => clearAllRecords());
els.clearSampleBtn.addEventListener("click", clearSampleData);
els.importTarget.addEventListener("change", () => {
  if (importSession?.step === "mapping") {
    importSession.module = els.importTarget.value;
    importSession.mapping = suggestMapping(importSession.module, importSession.headers);
    importSession.duplicateMode = can(importSession.module, "edit") ? "update" : "skip";
  }
  renderImportWizard();
});
els.importFileInput.addEventListener("change", handleImportFile);
els.sampleCsvBtn.addEventListener("click", downloadSampleCsv);
els.quickAddBtn.addEventListener("click", () => {
  const target = quickAddTarget();
  if (target) openModal(target);
});
els.exportBtn.addEventListener("click", exportData);
els.closeModalBtn.addEventListener("click", closeModal);
els.modalBackdrop.addEventListener("click", (event) => {
  if (event.target === els.modalBackdrop) closeModal();
});
els.recordForm.addEventListener("submit", handleSubmit);
els.clearActivityBtn.addEventListener("click", () => {
  state.activity = [];
  saveState();
  renderActivity();
  showToast("Activity cleared");
});
els.hideGettingStartedBtn.addEventListener("click", () => {
  state.organization.hideGettingStarted = true;
  saveState();
  render();
});
els.printReportBtn.addEventListener("click", () => window.print());
els.menuToggle.addEventListener("click", () => document.body.classList.toggle("menu-open"));
window.addEventListener("resize", () => {
  if (!currentUser()) return;
  window.clearTimeout(resizeTimer);
  resizeTimer = window.setTimeout(() => {
    if (currentView === "dashboard" && totalRecordCount()) {
      if (!byId("cashFlowPanel").hidden) renderCashFlowChart();
      if (!byId("inventoryMixPanel").hidden) renderInventoryChart();
    }
    if (currentView === "finance" && activeTab("finance") === "overview") renderFinanceOverview();
  }, 120);
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !els.modalBackdrop.hidden) closeModal();
});

if (state.pendingTemplate) {
  delete state.pendingTemplate;
  applyBusinessScope(state.organization.type, { silent: true, replace: true });
  saveState();
  OLD_STORAGE_KEYS.forEach((key) => {
    try {
      localStorage.removeItem(key);
    } catch (error) {
      console.warn(`Could not remove ${key}`, error);
    }
  });
}
populateBusinessTypeSelects();
populateSampleSelects();
resetFieldEditor();
if (currentUser()) showApp();
else showAuth();
