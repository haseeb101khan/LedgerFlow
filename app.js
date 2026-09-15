const STORAGE_KEY = "ledgerflow-erp-state-v3";
const OLD_STORAGE_KEYS = ["ledgerflow-erp-state-v2", "ledgerflow-erp-state-v1"];
const SESSION_KEY = "ledgerflow-session-v1";
const DEMO_STORAGE_KEY = "ledgerflow-erp-demo-v1";
const SHEETJS_URL = "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";

const MODULES = ["inventory", "finance", "contacts", "employees", "assets"];
// Sales and purchases are recorded transactions, not editable field-based records.
// They move stock, create finance entries, and create dues in one step.
const TRANSACTION_AREAS = ["sales", "purchases"];
const PERMISSION_AREAS = [...MODULES, ...TRANSACTION_AREAS];
const PAYMENT_METHODS = ["Cash", "Bank transfer", "Card", "Mobile wallet", "Cheque", "Other"];
const ADJUSTMENT_REASONS = [
  "Stock count correction",
  "Damaged",
  "Expired",
  "Lost or stolen",
  "Used internally",
  "Returned by customer",
  "Found stock",
  "Other",
];

const MODULE_GUIDANCE = {
  inventory: "Stock levels, product details, costs, and low-stock alerts.",
  finance: "Income, expenses, payments, dues, and profitability.",
  contacts: "Customers, suppliers, balances, and contact details.",
  employees: "Employee records, salaries, logins, and access roles.",
  assets: "Equipment, vehicles, assignments, value, and maintenance.",
  sales: "Record every sale: items, quantity, price, customer, and payment.",
  purchases: "Record stock bought from suppliers, what it cost, and what you owe.",
};

const SETUP_STEPS = ["Business", "Documents", "Work areas", "Starting plan", "Ready"];
const FISCAL_MONTHS = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

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
  sales: "Sales",
  purchases: "Purchases",
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
  sales: "Sales",
  purchases: "Purchases",
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
      { name: "sku", label: "SKU / Code", type: "text", autoId: true, section: "Basic info" },
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
      { name: "code", label: "Asset Code", type: "text", autoId: true, section: "Basic info" },
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
      { name: "employeeId", label: "Employee ID", type: "text", autoId: true, section: "Personal" },
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
      { name: "reference", label: "Reference / Voucher No.", type: "text", autoId: true, section: "Transaction" },
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
      { name: "code", label: "Contact ID", type: "text", autoId: true, section: "Contact" },
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
importAliases.employees.employeeId = ["employeeid", "empid", "employeecode", "empcode", "staffid", "staffno", "employeeno"];
importAliases.contacts.code = ["contactid", "customerid", "supplierid", "vendorid", "customercode", "suppliercode", "partycode", "accountno"];
importAliases.finance.reference = ["reference", "ref", "refno", "voucherno", "voucher", "invoiceno", "billno", "receiptno"];

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

let session = loadSession();
let state = session?.userId === "guest" ? loadDemoState() : loadState();
let previewRoleId = null;
let currentView = "dashboard";
let editing = null;
let importSession = null;
let editingField = null;
let sheetJsPromise = null;
let resizeTimer = null;
let setupWizardStep = 0;
let tutorialStep = 0;
let tutorialSteps = [];
let tutorialOriginView = "dashboard";
let tutorialFollowUp = false;
let tutorialPositionTimer = null;
let tutorialMode = { type: "main" };

const byId = (id) => document.getElementById(id);

const els = {
  homeScreen: byId("homeScreen"),
  homeMenuBtn: byId("homeMenuBtn"),
  homeNav: byId("homeNav"),
  trialDialogBackdrop: byId("trialDialogBackdrop"),
  trialDialogCloseBtn: byId("trialDialogCloseBtn"),
  trialBlankBtn: byId("trialBlankBtn"),
  trialBlankTitle: byId("trialBlankTitle"),
  trialSampleSelect: byId("trialSampleSelect"),
  trialSampleBtn: byId("trialSampleBtn"),
  trialResetBtn: byId("trialResetBtn"),
  authScreen: byId("authScreen"),
  authBackHomeBtn: byId("authBackHomeBtn"),
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
  exitTrialBtn: byId("exitTrialBtn"),
  trialBanner: byId("trialBanner"),
  trialBannerExitBtn: byId("trialBannerExitBtn"),
  globalSearch: byId("globalSearch"),
  previewSwitch: byId("previewSwitch"),
  previewRoleSelect: byId("previewRoleSelect"),
  previewBanner: byId("previewBanner"),
  previewBannerText: byId("previewBannerText"),
  exitPreviewBtn: byId("exitPreviewBtn"),
  quickAddBtn: byId("quickAddBtn"),
  quickAddLabel: byId("quickAddLabel"),
  exportBtn: byId("exportBtn"),
  tutorialBtn: byId("tutorialBtn"),
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
  countryInput: byId("countryInput"),
  timezoneSelect: byId("timezoneSelect"),
  fiscalYearStartSelect: byId("fiscalYearStartSelect"),
  businessPhoneInput: byId("businessPhoneInput"),
  businessEmailInput: byId("businessEmailInput"),
  registrationNumberInput: byId("registrationNumberInput"),
  businessAddressInput: byId("businessAddressInput"),
  setupModuleChoices: byId("setupModuleChoices"),
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
  setupWizardBackdrop: byId("setupWizardBackdrop"),
  setupWizardProgress: byId("setupWizardProgress"),
  setupWizardForm: byId("setupWizardForm"),
  setupWizardContent: byId("setupWizardContent"),
  setupWizardError: byId("setupWizardError"),
  setupWizardBackBtn: byId("setupWizardBackBtn"),
  setupWizardNextBtn: byId("setupWizardNextBtn"),
  setupWizardStepLabel: byId("setupWizardStepLabel"),
  tourLayer: byId("tourLayer"),
  tourSpotlight: byId("tourSpotlight"),
  tourCard: byId("tourCard"),
  tourStepCount: byId("tourStepCount"),
  tourIcon: byId("tourIcon"),
  tourTitle: byId("tourTitle"),
  tourText: byId("tourText"),
  tourProgress: byId("tourProgress"),
  tourBackBtn: byId("tourBackBtn"),
  tourNextBtn: byId("tourNextBtn"),
  skipTutorialBtn: byId("skipTutorialBtn"),
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
    version: 3,
    organization: null,
    onboarding: {
      setupCompleted: false,
      setupStep: 0,
      startPath: "import",
      addTeam: false,
      tutorials: {},
      tabGuides: {},
      autoGuidesOff: {},
      fieldHelpHidden: {},
    },
    roles: defaultRoles(),
    customFields: Object.fromEntries(MODULES.map((module) => [module, []])),
    fieldSettings: Object.fromEntries(MODULES.map((module) => [module, {}])),
    inventory: [],
    assets: [],
    employees: [],
    finance: [],
    contacts: [],
    sales: [],
    purchases: [],
    stockMovements: [],
    counters: { sale: 0, purchase: 0 },
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
      PERMISSION_AREAS.map((module) => [module, { view: false, add: false, edit: false, delete: false }]),
    ),
    ...Object.fromEntries(generalPermissions.map((permission) => [permission.key, false])),
  };
}

function fullPermissions() {
  return {
    modules: Object.fromEntries(
      PERMISSION_AREAS.map((module) => [module, { view: true, add: true, edit: true, delete: true }]),
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
        { inventory: "vaed", finance: "vae", contacts: "vaed", employees: "v", assets: "vae", sales: "vaed", purchases: "vae" },
        { sensitiveNumbers: true, reports: true, imports: true },
      ),
    },
    {
      id: "staff",
      name: "Staff",
      description: "Daily work: stock updates, recording sales, and customer dues.",
      system: false,
      permissions: buildPermissions({ inventory: "vae", finance: "va", contacts: "vae", sales: "va" }),
    },
  ];
}

function normalizePermissions(raw = {}) {
  const permissions = blankPermissions();
  const legacy = !raw.modules;
  PERMISSION_AREAS.forEach((module) => {
    if (!legacy && TRANSACTION_AREAS.includes(module) && !raw.modules[module]) {
      // Roles saved before sales and purchases existed: derive sensible access.
      const inventory = raw.modules.inventory || {};
      const finance = raw.modules.finance || {};
      permissions.modules[module] =
        module === "sales"
          ? {
              view: Boolean(inventory.view || finance.view),
              add: Boolean(finance.add || inventory.edit),
              edit: Boolean(finance.edit),
              delete: Boolean(finance.delete),
            }
          : {
              view: Boolean(inventory.view && finance.view),
              add: Boolean(inventory.add && finance.add),
              edit: Boolean(inventory.edit && finance.edit),
              delete: Boolean(inventory.delete && finance.delete),
            };
      return;
    }
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

function newDemoState() {
  const demo = emptyState();
  demo.demoMode = true;
  demo.organization = {
    name: "Trial Business",
    type: "General Business",
    currency: "PKR",
    dateFormat: defaultDateFormat(),
    country: "Pakistan",
    timezone: "Asia/Karachi",
    fiscalYearStart: "January",
    ownerName: "Guest",
    enabledModules: [...PERMISSION_AREAS],
    transactionAreasAdded: true,
    moduleLabels: {},
    recordNouns: {},
    createdAt: new Date().toISOString(),
  };
  demo.onboarding.setupCompleted = true;
  return normalizeLoadedState(demo);
}

function loadDemoState() {
  const saved = readStorage(DEMO_STORAGE_KEY);
  return saved ? normalizeLoadedState({ ...saved, demoMode: true }) : newDemoState();
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
    let enabled = Array.isArray(next.organization.enabledModules)
      ? next.organization.enabledModules.filter((module) => PERMISSION_AREAS.includes(module))
      : [...PERMISSION_AREAS];
    if (!next.organization.transactionAreasAdded) {
      // Businesses set up before sales and purchases existed get them switched on
      // wherever stock or finance is used.
      if (enabled.includes("inventory") || enabled.includes("finance")) {
        enabled = [...new Set([...enabled, ...TRANSACTION_AREAS])];
      }
      next.organization.transactionAreasAdded = true;
    }
    next.organization.enabledModules = enabled.length ? enabled : [...PERMISSION_AREAS];
  }

  const onboarding = loaded.onboarding || {};
  next.onboarding = {
    setupCompleted: Boolean(onboarding.setupCompleted),
    setupStep: Math.max(0, Math.min(SETUP_STEPS.length - 1, Number(onboarding.setupStep) || 0)),
    startPath: ["import", "manual", "sample"].includes(onboarding.startPath) ? onboarding.startPath : "import",
    addTeam: Boolean(onboarding.addTeam),
    tutorials: onboarding.tutorials && typeof onboarding.tutorials === "object" ? onboarding.tutorials : {},
    // Per-login: page guides already seen, guides switched off, and field help hidden.
    tabGuides: onboarding.tabGuides && typeof onboarding.tabGuides === "object" ? onboarding.tabGuides : {},
    autoGuidesOff: onboarding.autoGuidesOff && typeof onboarding.autoGuidesOff === "object" ? onboarding.autoGuidesOff : {},
    fieldHelpHidden: onboarding.fieldHelpHidden && typeof onboarding.fieldHelpHidden === "object" ? onboarding.fieldHelpHidden : {},
  };

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
  next.sales = Array.isArray(loaded.sales) ? loaded.sales : [];
  next.purchases = Array.isArray(loaded.purchases) ? loaded.purchases : [];
  next.stockMovements = Array.isArray(loaded.stockMovements)
    ? loaded.stockMovements
    : // Data saved before stock history existed: start each item's history from its current quantity.
      next.inventory
        .filter((item) => Number(item.quantity) > 0)
        .map((item) => ({
          id: `move-opening-${item.id}`,
          date: String(item.createdAt || item.importedAt || new Date().toISOString()).slice(0, 10),
          createdAt: new Date().toISOString(),
          itemId: item.id,
          itemName: item.name,
          sku: item.sku || "",
          type: "Opening stock",
          change: Number(item.quantity),
          balanceAfter: Number(item.quantity),
          reference: "",
          referenceId: "",
          referenceKind: "",
          party: "",
          unitAmount: item.unitCost ?? "",
          currency: item.currency,
          reason: "Stock on hand when stock history started",
          by: "",
        }));
  next.counters = {
    sale: Math.max(Number(loaded.counters?.sale) || 0, next.sales.length),
    purchase: Math.max(Number(loaded.counters?.purchase) || 0, next.purchases.length),
    // Document numbers for vouchers, payment receipts, and adjustment notes are never reused.
    voucher: Number(loaded.counters?.voucher) || 0,
    receipt: Number(loaded.counters?.receipt) || 0,
    paymentOut: Number(loaded.counters?.paymentOut) || 0,
    salary: Number(loaded.counters?.salary) || 0,
    adjustment: Number(loaded.counters?.adjustment) || 0,
  };
  return next;
}

/* Automatic IDs: any record saved without its ID gets the next free one, like ITEM-0007. */

const AUTO_ID_FIELDS = { inventory: "sku", assets: "code", employees: "employeeId", contacts: "code", finance: "reference" };

function autoIdPrefix(module, record = {}) {
  if (module === "inventory") {
    const noun = String(state?.organization?.recordNouns?.inventory || "Item")
      .replace(/[^A-Za-z]/g, "")
      .toUpperCase();
    return !noun ? "ITEM" : noun.length <= 4 ? noun : noun.slice(0, 3);
  }
  if (module === "assets") return "AST";
  if (module === "employees") return "EMP";
  if (module === "contacts") return record.type === "Supplier" ? "SUP" : "CUS";
  return record.type === "Income" ? "RV" : "PV";
}

function autoIdApplies(module, record) {
  if (!record || !AUTO_ID_FIELDS[module]) return false;
  // Finance entries from a sale, purchase, or salary already carry their own number.
  if (module === "finance" && (record.sourceType || (record.employeeId && record.payrollMonth))) return false;
  return true;
}

function autoIdContext(module, excludeId = "") {
  const field = AUTO_ID_FIELDS[module];
  const used = new Set(
    (state[module] || [])
      .filter((record) => record.id !== excludeId)
      .map((record) => String(record[field] ?? "").trim().toUpperCase())
      .filter(Boolean),
  );
  return { used, highest: {} };
}

function generateAutoId(module, record, context, { preview = false } = {}) {
  const { used, highest } = context;
  if (module === "finance") {
    const prefix = autoIdPrefix(module, record);
    if (record.documentNumber && !used.has(String(record.documentNumber).toUpperCase())) return record.documentNumber;
    let counter = Number(state.counters?.voucher) || 0;
    let value;
    do {
      counter += 1;
      value = `${prefix}-${String(counter).padStart(4, "0")}`;
    } while (used.has(value));
    if (!preview) state.counters.voucher = counter;
    return value;
  }
  const prefix = autoIdPrefix(module, record);
  if (highest[prefix] === undefined) {
    const pattern = new RegExp(`^${prefix}-(\\d+)$`);
    highest[prefix] = 0;
    used.forEach((value) => {
      const match = value.match(pattern);
      if (match) highest[prefix] = Math.max(highest[prefix], Number(match[1]));
    });
  }
  let next = highest[prefix];
  let value;
  do {
    next += 1;
    value = `${prefix}-${String(next).padStart(4, "0")}`;
  } while (used.has(value));
  if (!preview) highest[prefix] = next;
  return value;
}

function assignAutoId(module, record, context = autoIdContext(module, record.id)) {
  const field = AUTO_ID_FIELDS[module];
  if (!autoIdApplies(module, record) || !isBlank(record[field])) return false;
  record[field] = generateAutoId(module, record, context);
  context.used.add(String(record[field]).toUpperCase());
  if (module === "finance" && !record.documentNumber) record.documentNumber = record[field];
  if (module === "inventory") {
    state.stockMovements?.forEach((move) => {
      if (move.itemId === record.id && isBlank(move.sku)) move.sku = record[field];
    });
  }
  return true;
}

function previewAutoId(module, record = {}) {
  return generateAutoId(module, record, autoIdContext(module, record.id), { preview: true });
}

function findIdClash(module, record) {
  const field = AUTO_ID_FIELDS[module];
  if (!field || isBlank(record[field])) return null;
  const value = String(record[field]).trim().toUpperCase();
  return state[module].find((entry) => entry.id !== record.id && String(entry[field] ?? "").trim().toUpperCase() === value) || null;
}

function fillAutoIds() {
  if (!state) return;
  Object.keys(AUTO_ID_FIELDS).forEach((module) => {
    const records = state[module] || [];
    const blanks = records.filter((record) => autoIdApplies(module, record) && isBlank(record[AUTO_ID_FIELDS[module]]));
    if (!blanks.length) return;
    const context = autoIdContext(module);
    // Oldest records get the lowest numbers (newest records are kept first in each list).
    [...blanks].reverse().forEach((record) => assignAutoId(module, record, context));
  });
}

function saveState() {
  fillAutoIds();
  try {
    localStorage.setItem(state.demoMode ? DEMO_STORAGE_KEY : STORAGE_KEY, JSON.stringify(state));
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
  if (session?.userId === "guest" && state.demoMode) {
    return { id: "guest", name: "Guest", email: "", roleId: "owner", isOwner: true, isGuest: true };
  }
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

function moduleEnabled(module) {
  const enabled = state.organization?.enabledModules;
  return !Array.isArray(enabled) || enabled.includes(module);
}

function can(module, action) {
  return moduleEnabled(module) && Boolean(currentRole()?.permissions?.modules?.[module]?.[action]);
}

function canOpenView(view) {
  if (view === "dashboard") return true;
  if (view === "setup") return canAccess("settings");
  if (view === "dataSources") return canAccess("imports") && MODULES.some((module) => can(module, "add"));
  if (view === "reports") return canAccess("reports");
  if (PERMISSION_AREAS.includes(view)) return can(view, "view");
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
  const wasGuest = session?.userId === "guest";
  session = null;
  previewRoleId = null;
  importSession = null;
  tutorialSteps = [];
  tutorialFollowUp = false;
  saveSession(null);
  closeModal();
  state = loadState();
  showHome();
  if (wasGuest) showToast("Trial saved in this browser");
}

/* ---------- Fields ---------- */

function getFields(module, options = {}) {
  const overrides = state.fieldSettings?.[module] || {};
  const base = schemas[module].fields.map((field) => {
    const override = overrides[field.name] || {};
    return {
      ...field,
      label: override.label || field.label,
      // ID fields are never required: a blank one is filled in automatically.
      required: field.autoId ? false : field.locked ? Boolean(field.required) : Boolean(override.required ?? field.required),
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

function showHome() {
  els.appShell.hidden = true;
  els.authScreen.hidden = true;
  els.trialDialogBackdrop.hidden = true;
  els.homeScreen.hidden = false;
  els.modalBackdrop.hidden = true;
  els.setupWizardBackdrop.hidden = true;
  els.tourLayer.hidden = true;
  els.homeNav.classList.remove("is-open");
  els.homeMenuBtn.setAttribute("aria-expanded", "false");
  document.body.classList.remove("has-blocking-overlay", "menu-open");
  document.title = "LedgerFlow | Business records, made clear";
  window.scrollTo(0, 0);
}

function openAuthFromHome() {
  state = loadState();
  session = null;
  saveSession(null);
  showAuth();
}

function openTrialDialog() {
  const savedTrial = Boolean(readStorage(DEMO_STORAGE_KEY));
  els.trialBlankTitle.textContent = savedTrial ? "Continue your trial workspace" : "Start with an empty workspace";
  els.trialResetBtn.hidden = !savedTrial;
  els.trialDialogBackdrop.hidden = false;
  els.trialBlankBtn.focus();
}

function beginTrial({ sample = false, fresh = false } = {}) {
  const hadSavedTrial = Boolean(readStorage(DEMO_STORAGE_KEY));
  if (fresh) localStorage.removeItem(DEMO_STORAGE_KEY);
  state = loadDemoState();
  session = { userId: "guest", startedAt: new Date().toISOString() };
  saveSession(session);
  previewRoleId = null;
  currentView = "dashboard";
  els.globalSearch.value = "";
  showApp();
  if (sample) loadSampleBusiness(els.trialSampleSelect.value, { skipConfirm: !hadSavedTrial || fresh });
  else saveState();
}

function resetTrial() {
  if (!window.confirm("Delete this browser's trial records and start with an empty workspace?")) return;
  beginTrial({ fresh: true });
}

function populateBusinessTypeSelects() {
  document.querySelectorAll("[data-business-types]").forEach((select) => {
    select.innerHTML = Object.keys(businessScopes)
      .map((name) => `<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`)
      .join("");
  });
}

function showAuth() {
  els.homeScreen.hidden = true;
  els.trialDialogBackdrop.hidden = true;
  els.appShell.hidden = true;
  els.authScreen.hidden = false;
  els.modalBackdrop.hidden = true;
  els.setupWizardBackdrop.hidden = true;
  els.tourLayer.hidden = true;
  document.body.classList.remove("has-blocking-overlay", "menu-open");
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
  els.homeScreen.hidden = true;
  els.trialDialogBackdrop.hidden = true;
  els.authScreen.hidden = true;
  els.appShell.hidden = false;
  els.exitTrialBtn.hidden = !currentUser()?.isGuest;
  els.trialBanner.hidden = !currentUser()?.isGuest;
  setView(canOpenView(currentView) ? currentView : "dashboard");
  window.setTimeout(beginNewUserFlow, 120);
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
    country: existing?.country || (currency === "PKR" ? "Pakistan" : ""),
    timezone: existing?.timezone || (currency === "PKR" ? "Asia/Karachi" : Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC"),
    fiscalYearStart: existing?.fiscalYearStart || "January",
    businessEmail: existing?.businessEmail || ownerEmail,
    enabledModules: existing?.enabledModules || [...PERMISSION_AREAS],
    transactionAreasAdded: true,
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

/* ---------- First-time setup and tutorial ---------- */

function guidedSelectOptions(values, selected) {
  return values
    .map((value) => `<option value="${escapeHtml(value)}" ${value === selected ? "selected" : ""}>${escapeHtml(value)}</option>`)
    .join("");
}

function timezoneOptions(selected) {
  const detected = Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
  const values = [
    selected,
    detected,
    "Asia/Karachi",
    "Asia/Dubai",
    "Asia/Riyadh",
    "Asia/Kolkata",
    "Europe/London",
    "America/New_York",
    "America/Los_Angeles",
    "UTC",
  ].filter(Boolean);
  return guidedSelectOptions([...new Set(values)], selected || detected);
}

function moduleChoicesHtml(selectedModules, inputName = "enabledModule") {
  const selected = new Set(selectedModules?.length ? selectedModules : PERMISSION_AREAS);
  return PERMISSION_AREAS.map(
    (module) => `
      <label class="module-choice">
        <input type="checkbox" name="${inputName}" value="${module}" ${selected.has(module) ? "checked" : ""} />
        <span>
          <strong>${escapeHtml(navTitle(module))}</strong>
          <span>${escapeHtml(MODULE_GUIDANCE[module])}</span>
        </span>
      </label>
    `,
  ).join("");
}

function setupChoice({ name, value, selected, title, detail }) {
  return `
    <label class="setup-choice">
      <input type="radio" name="${name}" value="${value}" ${value === selected ? "checked" : ""} />
      <span>
        <strong>${escapeHtml(title)}</strong>
        <span>${escapeHtml(detail)}</span>
      </span>
    </label>
  `;
}

function showSetupWizardError(message) {
  els.setupWizardError.textContent = message;
  els.setupWizardError.hidden = false;
}

function openSetupWizard() {
  if (!currentUser()?.isOwner || state.onboarding.setupCompleted) return;
  setupWizardStep = Math.max(0, Math.min(SETUP_STEPS.length - 1, state.onboarding.setupStep || 0));
  els.setupWizardBackdrop.hidden = false;
  document.body.classList.add("has-blocking-overlay");
  renderSetupWizard();
}

function renderSetupWizard() {
  const org = state.organization;
  const selectedModules = org.enabledModules?.length ? org.enabledModules : PERMISSION_AREAS;
  els.setupWizardError.hidden = true;
  els.setupWizardProgress.innerHTML = SETUP_STEPS.map(
    (label, index) => `
      <li class="${index < setupWizardStep ? "is-done" : index === setupWizardStep ? "is-current" : ""}">
        <span>${index < setupWizardStep ? "✓" : index + 1}</span>
        <strong>${escapeHtml(label)}</strong>
      </li>
    `,
  ).join("");

  if (setupWizardStep === 0) {
    els.setupWizardContent.innerHTML = `
      <div class="setup-wizard-copy">
        <h1>Confirm your business basics</h1>
        <p>These settings keep reports, dates, and money consistent across your workspace.</p>
      </div>
      <div class="form-grid">
        <label class="field full">
          <span class="field-label">Business name *</span>
          <input name="businessName" value="${escapeHtml(org.name || "")}" required />
        </label>
        <label class="field">
          <span class="field-label">Business type *</span>
          <select name="businessType">${guidedSelectOptions(Object.keys(businessScopes), org.type)}</select>
        </label>
        <label class="field">
          <span class="field-label">Country *</span>
          <input name="country" value="${escapeHtml(org.country || (org.currency === "PKR" ? "Pakistan" : ""))}" required />
        </label>
        <label class="field">
          <span class="field-label">Default currency *</span>
          <input name="currency" maxlength="3" value="${escapeHtml(org.currency || "PKR")}" required />
          <small>Use a 3-letter code such as PKR, USD, or AED.</small>
        </label>
        <label class="field">
          <span class="field-label">Time zone *</span>
          <select name="timezone">${timezoneOptions(org.timezone || (org.currency === "PKR" ? "Asia/Karachi" : ""))}</select>
        </label>
        <label class="field">
          <span class="field-label">Fiscal year starts *</span>
          <select name="fiscalYearStart">${guidedSelectOptions(FISCAL_MONTHS, org.fiscalYearStart || "January")}</select>
        </label>
        <label class="field">
          <span class="field-label">Dates in your files *</span>
          <select name="dateFormat">
            <option value="DMY" ${org.dateFormat === "DMY" ? "selected" : ""}>Day/Month/Year</option>
            <option value="MDY" ${org.dateFormat === "MDY" ? "selected" : ""}>Month/Day/Year</option>
          </select>
        </label>
      </div>
    `;
  } else if (setupWizardStep === 1) {
    els.setupWizardContent.innerHTML = `
      <div class="setup-wizard-copy">
        <h1>Add details for reports and documents</h1>
        <p>LedgerFlow will use these details on future invoices, exports, and printable reports.</p>
      </div>
      <div class="form-grid">
        <label class="field">
          <span class="field-label">Business phone *</span>
          <input name="businessPhone" type="tel" value="${escapeHtml(org.businessPhone || "")}" required />
        </label>
        <label class="field">
          <span class="field-label">Business email</span>
          <input name="businessEmail" type="email" value="${escapeHtml(org.businessEmail || org.ownerEmail || "")}" />
        </label>
        <label class="field full">
          <span class="field-label">Business address *</span>
          <textarea name="businessAddress" rows="4" required>${escapeHtml(org.businessAddress || "")}</textarea>
        </label>
        <label class="field full">
          <span class="field-label">Registration / tax number</span>
          <input name="registrationNumber" value="${escapeHtml(org.registrationNumber || "")}" />
          <small>Optional. Add your NTN, STRN, company registration, or local tax number when applicable.</small>
        </label>
      </div>
    `;
  } else if (setupWizardStep === 2) {
    els.setupWizardContent.innerHTML = `
      <div class="setup-wizard-copy">
        <h1>Choose the work areas you need</h1>
        <p>Your sidebar will stay focused on these areas. You can turn any area on or off later.</p>
      </div>
      <div class="module-choice-grid">${moduleChoicesHtml(selectedModules)}</div>
    `;
  } else if (setupWizardStep === 3) {
    const teamChoices = moduleEnabled("employees")
      ? `
        <div class="setup-choice-grid">
          ${setupChoice({ name: "teamPlan", value: "later", selected: state.onboarding.addTeam ? "now" : "later", title: "I will start by myself", detail: "Use the owner account now and add employee logins whenever you are ready." })}
          ${setupChoice({ name: "teamPlan", value: "now", selected: state.onboarding.addTeam ? "now" : "later", title: "Add an employee next", detail: "Create an employee login and assign a role immediately after the tutorial." })}
        </div>
      `
      : `
        <input type="hidden" name="teamPlan" value="later" />
        <div class="setup-ready-note">
          <svg viewBox="0 0 24 24" aria-hidden="true"><path d="m5 12 4 4L19 6" /></svg>
          <span>The Employees work area is off, so you will begin with the owner account. You can enable employee access later.</span>
        </div>
      `;
    els.setupWizardContent.innerHTML = `
      <div class="setup-wizard-copy">
        <h1>Choose how you want to begin</h1>
        <p>This only decides where LedgerFlow takes you after the tutorial.</p>
      </div>
      <section class="setup-choice-section">
        <h2>Your existing records</h2>
        <div class="setup-choice-grid">
          ${setupChoice({ name: "startPath", value: "import", selected: state.onboarding.startPath, title: "Import my files", detail: "Bring in CSV, Excel, JSON, or tab-separated records with guided field matching." })}
          ${setupChoice({ name: "startPath", value: "manual", selected: state.onboarding.startPath, title: "Start entering records", detail: "Open your first active work area and create records one at a time." })}
          ${setupChoice({ name: "startPath", value: "sample", selected: state.onboarding.startPath, title: "Explore sample data", detail: "Load a realistic sample first, then remove it when you are ready for real records." })}
        </div>
      </section>
      <section class="setup-choice-section">
        <h2>Your team</h2>
        ${teamChoices}
      </section>
    `;
  } else {
    const startLabels = {
      import: "Import existing files",
      manual: "Enter the first record manually",
      sample: "Explore with sample data",
    };
    els.setupWizardContent.innerHTML = `
      <div class="setup-wizard-copy">
        <h1>Your workspace is ready</h1>
        <p>Review the setup below. Next, a short visual tour will show you where everything lives.</p>
      </div>
      <div class="setup-review">
        <div><strong>Business</strong><span>${escapeHtml(org.name)} · ${escapeHtml(org.type)}</span></div>
        <div><strong>Location & money</strong><span>${escapeHtml(org.country)} · ${escapeHtml(org.currency)} · ${escapeHtml(org.timezone)}</span></div>
        <div><strong>Active work areas</strong><span>${escapeHtml(selectedModules.map(navTitle).join(", "))}</span></div>
        <div><strong>First task</strong><span>${escapeHtml(startLabels[state.onboarding.startPath])}${state.onboarding.addTeam ? " · then add an employee" : ""}</span></div>
      </div>
      <div class="setup-ready-note">
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="m5 12 4 4L19 6" /></svg>
        <span>You can change all of these settings later under Setup & Roles.</span>
      </div>
    `;
  }

  els.setupWizardBackBtn.hidden = setupWizardStep === 0;
  els.setupWizardStepLabel.textContent = `Step ${setupWizardStep + 1} of ${SETUP_STEPS.length}`;
  els.setupWizardNextBtn.textContent = setupWizardStep === SETUP_STEPS.length - 1 ? "Finish setup" : "Save & continue";
  const firstInput = els.setupWizardContent.querySelector("input, select, textarea");
  window.setTimeout(() => firstInput?.focus(), 40);
}

function captureSetupWizardStep() {
  const form = new FormData(els.setupWizardForm);
  const controls = els.setupWizardForm.elements;
  const org = state.organization;

  if (setupWizardStep === 0) {
    const name = String(form.get("businessName") || "").trim();
    const country = String(form.get("country") || "").trim();
    const currency = String(form.get("currency") || "").trim().toUpperCase();
    if (!name || !country) {
      showSetupWizardError("Business name and country are required.");
      return false;
    }
    if (!isValidCurrency(currency)) {
      showSetupWizardError("Currency must be a 3-letter code such as PKR, USD, or AED.");
      return false;
    }
    const type = String(controls.businessType?.value || "General Business");
    if (type !== org.type) applyBusinessScope(type, { silent: true });
    Object.assign(org, {
      name,
      type,
      country,
      currency,
      timezone: String(controls.timezone?.value || "UTC"),
      fiscalYearStart: String(controls.fiscalYearStart?.value || "January"),
      dateFormat: String(controls.dateFormat?.value || "DMY"),
    });
  } else if (setupWizardStep === 1) {
    const phone = String(form.get("businessPhone") || "").trim();
    const address = String(form.get("businessAddress") || "").trim();
    if (!phone || !address) {
      showSetupWizardError("Business phone and address are required for business documents.");
      return false;
    }
    const emailInput = els.setupWizardForm.elements.businessEmail;
    if (emailInput?.value && !emailInput.checkValidity()) {
      showSetupWizardError("Enter a valid business email or leave it blank.");
      return false;
    }
    Object.assign(org, {
      businessPhone: phone,
      businessEmail: String(form.get("businessEmail") || "").trim().toLowerCase(),
      businessAddress: address,
      registrationNumber: String(form.get("registrationNumber") || "").trim(),
    });
  } else if (setupWizardStep === 2) {
    const enabledModules = form.getAll("enabledModule").filter((module) => PERMISSION_AREAS.includes(module));
    if (!enabledModules.length) {
      showSetupWizardError("Choose at least one work area to continue.");
      return false;
    }
    org.enabledModules = enabledModules;
  } else if (setupWizardStep === 3) {
    const startPath = String(form.get("startPath") || "");
    const teamPlan = String(form.get("teamPlan") || "");
    if (!startPath || !teamPlan) {
      showSetupWizardError("Choose a starting method and a team option.");
      return false;
    }
    state.onboarding.startPath = startPath;
    state.onboarding.addTeam = teamPlan === "now";
  }

  return true;
}

function submitSetupWizard(event) {
  event.preventDefault();
  if (!captureSetupWizardStep()) return;
  if (setupWizardStep < SETUP_STEPS.length - 1) {
    setupWizardStep += 1;
    state.onboarding.setupStep = setupWizardStep;
    saveState();
    render();
    renderSetupWizard();
    return;
  }

  state.onboarding.setupCompleted = true;
  state.onboarding.setupStep = SETUP_STEPS.length - 1;
  addActivity("Initial business setup completed", "setup");
  saveState();
  els.setupWizardBackdrop.hidden = true;
  document.body.classList.remove("has-blocking-overlay");
  setView("dashboard");
  startTutorial({ automatic: true, followUp: true });
}

function previousSetupWizardStep() {
  if (setupWizardStep <= 0) return;
  setupWizardStep -= 1;
  state.onboarding.setupStep = setupWizardStep;
  saveState();
  renderSetupWizard();
}

/* ---------- Field and section help ----------
 * Plain explanations shown under each form field and section. New users see them by default;
 * anyone can hide them, and each field's "?" still shows its own help.
 */

const FIELD_HELP = {
  inventory: {
    name: "The name you and your customers use for this item.",
    sku: "Your own code or barcode for this item. Leave it blank and LedgerFlow gives it the next free code.",
    category: "A group such as Grocery, Tablets, or Cement. Used for filters and the stock mix chart.",
    quantity: "How many you have right now. After saving, stock only changes through sales, purchases, and stock adjustments.",
    reorderLevel: "When stock falls to this number, the item is flagged as low stock so you can reorder in time.",
    location: "Where the item is kept, such as a shelf, rack, store room, or site.",
    unitCost: "What one unit costs you. Purchases update it to the average cost automatically. Only roles allowed to see costs can see it.",
    sellPrice: "Your normal selling price. It is suggested on each sale, and can be changed per sale.",
    currency: "The currency the cost and price are in. LedgerFlow never converts or mixes currencies.",
    notes: "Anything else worth remembering about this item.",
  },
  assets: {
    name: "What the asset is, such as a delivery van, laptop, or generator.",
    code: "Your tag or serial number for this asset. Leave it blank for an automatic code.",
    condition: "Active assets are in use. Maintenance and Damaged assets show up under Needs maintenance.",
    assignedTo: "The person or team responsible for it.",
    location: "Where the asset is kept or used.",
    purchaseDate: "When the business got it.",
    value: "What the asset is worth. Counted in total asset value; hidden from roles that can't see values.",
    currency: "The currency of the value. It is never converted.",
    notes: "Warranty, service history, or anything else.",
  },
  employees: {
    name: "The employee's full name as it should appear on salary slips.",
    employeeId: "Staff number used on salary slips and records. Leave it blank for an automatic ID.",
    phone: "A number to reach them.",
    email: "Their personal or work email (separate from the login email below).",
    role: "Their job title, such as Cashier or Site Engineer.",
    department: "The team they belong to. Used to filter the employee list.",
    status: "Active employees are on payroll. Inactive employees are kept for history but not paid.",
    salary: "Monthly salary. Payroll uses it to show who is unpaid each month.",
    currency: "The currency the salary is paid in.",
    loginEmail: "Add this to let the employee sign in. They only see what their access role allows.",
    loginPassword: "A password for their first sign-in. Share it with them privately.",
    accessRole: "Decides which sections and actions this employee can use. Roles are set in Setup & Roles.",
    notes: "Joining date, emergency contact, or anything else.",
  },
  finance: {
    type: "Income is money coming in; Expense is money going out.",
    date: "The date of the transaction.",
    category: "A heading such as Rent, Utilities, or Services. Totals by category appear in the Finance overview.",
    description: "A short line saying what this was for.",
    reference: "A bill, receipt, or voucher number. Leave it blank to get the next voucher number (RV for income, PV for expenses).",
    amount: "The full amount of the transaction.",
    currency: "The currency the amount was recorded in. It is never converted.",
    status: "Paid if settled. Pending or Overdue if still owed; those appear under Finance → Dues.",
    dueDate: "When an unpaid amount should be paid. After this date it is marked overdue.",
    party: "Who paid you or who you paid. Pick a saved contact or type any name.",
    notes: "Anything else about this transaction.",
  },
  contacts: {
    name: "The person or company name.",
    type: "Customer if they buy from you, Supplier if you buy from them.",
    code: "Your account number for this contact. Leave it blank for an automatic ID (CUS- or SUP-).",
    phone: "A number to reach them. Shown on invoices.",
    email: "An email address. Shown on invoices.",
    balance: "What was already owed before you started LedgerFlow. Positive means they owe you; negative means you owe them.",
    currency: "The currency of the opening balance.",
    status: "Active contacts are in use. Blocked or Inactive contacts are kept for history.",
    notes: "Address, payment terms, or anything else.",
  },
};

const SECTION_HELP = {
  "Basic info": "What this record is and how to find it again.",
  Stock: "How much you have and when to reorder.",
  Pricing: "What it costs you and what you sell it for.",
  Assignment: "Who has it and where it is.",
  Value: "What it is worth.",
  Personal: "Who this employee is.",
  Job: "What they do and where they work.",
  Pay: "What they are paid each month.",
  "Login & access": "Optional. Lets this employee sign in and controls what they can see.",
  Transaction: "What happened and when.",
  "Amount & payment": "How much, and whether it has been paid.",
  Who: "The other side of the transaction.",
  Contact: "Who they are and how to reach them.",
  Account: "Money already owed and account status.",
  Notes: "Free text for anything else.",
};

const TRADE_SECTION_HELP = {
  sale: {
    party: "Choose a saved customer, add a new one, or leave it as a walk-in. A name is needed if they will pay later.",
    lines: "Add each item sold with its quantity and price. The list price is suggested and can be changed. Stock goes down when you save.",
    payment: "Say whether the customer paid in full, paid part, or will pay later. Anything unpaid is added to Finance → Dues.",
    details: "Optional reference and notes. The sale number (S-0001…) is created automatically.",
  },
  purchase: {
    party: "Choose the supplier or dealer the stock came from, or add a new one.",
    lines: "Add each item bought with its quantity and cost. Choose + New item for something not in your list yet. Stock goes up when you save.",
    payment: "Say whether you paid the supplier. Anything unpaid is added to Finance → Dues as money you owe.",
    details: "The supplier's bill number and notes. The purchase number (P-0001…) is created automatically.",
  },
};

function fieldHelpText(module, field) {
  if (field.help) return field.help;
  const help = FIELD_HELP[module]?.[field.name];
  if (help) return help;
  if (field.custom) return `A field your business added for ${viewTitle(module).toLowerCase()}. Fill it in if it applies.`;
  return "";
}

function fieldHelpVisible() {
  const user = currentUser();
  return !(user && state.onboarding?.fieldHelpHidden?.[user.id]);
}

function toggleFieldHelp() {
  const user = currentUser();
  if (!user) return;
  state.onboarding.fieldHelpHidden ||= {};
  state.onboarding.fieldHelpHidden[user.id] = fieldHelpVisible();
  saveState();
  applyFieldHelpVisibility();
}

function applyFieldHelpVisibility() {
  const visible = fieldHelpVisible();
  els.recordForm.classList.toggle("show-help", visible);
  const button = els.recordForm.querySelector("[data-toggle-field-help]");
  if (button) button.textContent = visible ? "Hide help" : "Show help";
}

function formHelpBarHtml(text) {
  return `
    <div class="form-help-bar">
      <span>${escapeHtml(text)}</span>
      <button class="link-button" type="button" data-toggle-field-help>${fieldHelpVisible() ? "Hide help" : "Show help"}</button>
    </div>
  `;
}

function sectionHelpHtml(text) {
  return text ? `<p class="section-help">${escapeHtml(text)}</p>` : "";
}

/* ---------- Page and tab guides ----------
 * Each page and each tab has a short spotlight guide. It runs automatically the first time a
 * person opens that tab (remembered per login) and can be replayed with the Guide button.
 */

function tabButtonSelector(view, tab) {
  return `#${view}Tabs [data-tab="${view}:${tab}"]`;
}

function missingInfoGuide(view) {
  const noun = viewTitle(view).toLowerCase();
  return [
    {
      selector: tabButtonSelector(view, "missing"),
      icon: "alert",
      title: "Needs info",
      text: `${viewTitle(view)} records with required details missing, usually from an imported file. Nothing is lost; they just need completing.`,
    },
    {
      selector: `#${view}Table`,
      icon: "info",
      title: "Complete or continue",
      text: `Use Fill in on a row to add what is missing. If you don't have the information, you can save and continue without it. When the list is empty, all ${noun} are complete.`,
    },
  ];
}

function tradeGuides(kind) {
  const config = TRADE[kind];
  const area = config.area;
  const sale = kind === "sale";
  return {
    all: [
      {
        selector: tabButtonSelector(area, "all"),
        icon: sale ? "income" : "expense",
        title: sale ? "All sales" : "All purchases",
        text: sale
          ? "Every sale ever recorded, newest first: who bought, what they bought, the total, what was paid, and what is still owed."
          : "Every purchase of stock, newest first: which supplier, what came in, what it cost, and what you still owe them.",
      },
      {
        selector: `#${area}View [data-new-trade="${kind}"]`,
        icon: "info",
        title: sale ? "Record a sale" : "Record a purchase",
        text: sale
          ? "Use New sale every time you sell. Choose the items, quantity, and price, pick the customer or walk-in, and say how they paid. Stock, income, and dues update together."
          : "Use New purchase whenever stock arrives. Choose the supplier and items (or + New item), enter quantity and cost, and say whether you paid. Stock, costs, and dues update together.",
      },
      {
        selector: `#${area}PeriodFilter`,
        icon: "clock",
        title: "Choose a period",
        text: "Show this month, last month, this year, or all time. The summary next to it totals only what is shown.",
      },
      {
        selector: `#${area}Table`,
        icon: "box",
        title: "Open, invoice, or get paid",
        text: sale
          ? "Click a row to see the full sale. Invoice prints a professional invoice or shop receipt. Receive records a payment for anything unpaid."
          : "Click a row to see the full purchase. Invoice prints the purchase record. Pay records a payment to the supplier.",
      },
    ],
    unpaid: [
      {
        selector: tabButtonSelector(area, "unpaid"),
        icon: "alert",
        title: sale ? "Unpaid sales" : "Unpaid purchases",
        text: sale
          ? "Sales where the customer still owes money, including part-paid ones. Overdue ones are marked red."
          : "Purchases you have not fully paid for yet. Overdue bills are marked red.",
      },
      {
        selector: `#${area}Table`,
        icon: "wallet",
        title: sale ? "Receive a payment" : "Pay a supplier",
        text: "Use the payment button on a row. Enter less than the balance for a part payment. A receipt is created and the balance drops right away.",
      },
    ],
    payments: [
      {
        selector: tabButtonSelector(area, "payments"),
        icon: "wallet",
        title: sale ? "Payments received" : "Payments made",
        text: sale
          ? "Every payment customers made against sales, including money paid at the time of sale and later part payments."
          : "Every payment you made to suppliers for purchases.",
      },
      {
        selector: `#${area}Table`,
        icon: "info",
        title: "Receipts for every payment",
        text: "Each payment has its own receipt number. Receipt prints it again; Open shows the sale or purchase it paid.",
      },
    ],
    cancelled: [
      {
        selector: tabButtonSelector(area, "cancelled"),
        icon: "alert",
        title: sale ? "Cancelled sales" : "Cancelled purchases",
        text: "Records are never deleted. A cancelled one stays here with who cancelled it and why, and its stock and money were reversed.",
      },
    ],
  };
}

function guideDefinitions() {
  const items = viewTitle("inventory");
  const itemNoun = recordNoun("inventory").toLowerCase();
  const sales = tradeGuides("sale");
  const purchases = tradeGuides("purchase");
  return {
    "dashboard:": [
      {
        selector: "#dashboardView .page-head",
        icon: "info",
        title: "Dashboard",
        text: "The overall picture of the business, built from your records. Nothing here is typed in; it updates as you work.",
      },
      {
        selector: "#dashboardMetrics",
        icon: "wallet",
        title: "Key numbers",
        text: "Income, expenses, profit, stock value, and what is owed. Amounts in different currencies are shown separately.",
      },
      { selector: "#duesPanel", icon: "clock", title: "Dues", text: "Who owes you and who you owe, most urgent first. All dues opens Finance → Dues to settle them." },
      { selector: "#needsInputPanel", icon: "alert", title: "To do", text: "Overdue payments, low stock, unpaid salaries, and incomplete records that need you." },
      { selector: "#activityPanel", icon: "users", title: "Recent activity", text: "The latest sales, purchases, payments, and changes, with who made them." },
    ],
    "finance:overview": [
      {
        selector: tabButtonSelector("finance", "overview"),
        icon: "wallet",
        title: "Finance overview",
        text: "Income, expenses, and profit by month or by year, calculated from every transaction, sale, purchase, and salary.",
      },
      {
        selector: "#financeView [data-panel-tabs='overview'] .control-bar",
        icon: "clock",
        title: "Monthly or annual",
        text: "Switch between months and years, pick the year, and choose the currency when you record in more than one.",
      },
      { selector: "#financeKpis", icon: "income", title: "Totals for the period", text: "Income, expenses, profit, and profit margin for the chosen period." },
      { selector: "#financePeriodTable", icon: "info", title: "Month by month", text: "Click any row to open the transactions behind those numbers." },
      { selector: "#financeIncomeCategories", icon: "box", title: "By category", text: "Where money comes from and where it goes, grouped by category." },
    ],
    "finance:income": [
      {
        selector: tabButtonSelector("finance", "income"),
        icon: "income",
        title: "Income",
        text: "All money coming in: sales, customer payments, and any other income such as services or rent received.",
      },
      {
        selector: "#financeView [data-add='finance']",
        icon: "info",
        title: "Add other income",
        text: "Use Add transaction for income that is not a sale. Sales should be recorded in Sales so stock updates too.",
      },
      {
        selector: "#financeView [data-panel-tabs='income expenses all'] .table-toolbar",
        icon: "clock",
        title: "Filter the list",
        text: "Filter by category, paid or unpaid, and month. The summary totals what is shown.",
      },
      {
        selector: "#financeTable",
        icon: "box",
        title: "Receipts and payments",
        text: "Receipt prints a voucher. Receive payment settles unpaid income. Entries from a sale open that sale.",
      },
    ],
    "finance:expenses": [
      {
        selector: tabButtonSelector("finance", "expenses"),
        icon: "expense",
        title: "Expenses",
        text: "All money going out: stock purchases, salaries, rent, bills, and anything else you pay for.",
      },
      {
        selector: "#financeView [data-add='finance']",
        icon: "info",
        title: "Add an expense",
        text: "Use Add transaction for bills, rent, utilities, and other costs. Stock purchases belong in Purchases; salaries in Employees → Payroll.",
      },
      {
        selector: "#financeTable",
        icon: "box",
        title: "Vouchers and payments",
        text: "Voucher prints a payment voucher. Pay settles an unpaid bill. Salary rows open Payroll and their salary slip.",
      },
    ],
    "finance:dues": [
      {
        selector: tabButtonSelector("finance", "dues"),
        icon: "clock",
        title: "Dues",
        text: "Everyone who owes the business, and everyone the business owes, in one place.",
      },
      { selector: "#duesTotals", icon: "alert", title: "What is owed", text: "Totals owed to you and by you, with how much of each is overdue." },
      {
        selector: "#receivablesList",
        icon: "income",
        title: "Owed to you",
        text: "Grouped by customer. Receive payment settles their oldest items first and creates a receipt. Statement prints everything they owe.",
      },
      {
        selector: "#payablesList",
        icon: "expense",
        title: "You owe",
        text: "Grouped by supplier. Pay settles their bills oldest first. Unpaid salaries are listed here too and paid from Employees → Payroll.",
      },
    ],
    "finance:payments": [
      {
        selector: tabButtonSelector("finance", "payments"),
        icon: "wallet",
        title: "Payments",
        text: "Every payment received and paid, newest first: sales, purchases, dues, opening balances, and salaries.",
      },
      {
        selector: "#paymentsDirectionFilter",
        icon: "clock",
        title: "Money in or out",
        text: "Show only money received, only money paid out, or both, for any month.",
      },
      {
        selector: "#paymentsLedger",
        icon: "info",
        title: "Receipts",
        text: "Each payment has a receipt number. Receipt prints it again; Open jumps to what it paid for.",
      },
    ],
    "finance:all": [
      {
        selector: tabButtonSelector("finance", "all"),
        icon: "box",
        title: "All transactions",
        text: "The complete finance record: income and expenses, paid and unpaid, including cancelled entries kept for history.",
      },
      {
        selector: "#financeView [data-panel-tabs='income expenses all'] .table-toolbar",
        icon: "clock",
        title: "Find anything",
        text: "Combine category, payment status, and month filters with the search box at the top.",
      },
    ],
    "inventory:all": [
      {
        selector: tabButtonSelector("inventory", "all"),
        icon: "box",
        title: `All ${items.toLowerCase()}`,
        text: `Every ${itemNoun} with its stock on hand, cost, price, and stock value.`,
      },
      {
        selector: "#inventoryView [data-add='inventory']",
        icon: "info",
        title: `Add a ${itemNoun}`,
        text: "Enter its name, category, prices, and the stock you have now. The code is filled in automatically if you leave it blank.",
      },
      { selector: "#inventoryCategoryFilter", icon: "clock", title: "Filter by category", text: "Show one category at a time." },
      {
        selector: "#inventoryTable",
        icon: "wallet",
        title: "Sell, adjust, and trace",
        text: "Sell starts a sale with this item. Adjust stock records damage, loss, or count corrections with a reason. History shows every stock change.",
      },
    ],
    "inventory:low": [
      {
        selector: tabButtonSelector("inventory", "low"),
        icon: "alert",
        title: "Low stock",
        text: `${items} at or below their reorder level. Record a purchase when new stock arrives and they leave this list.`,
      },
    ],
    "inventory:history": [
      {
        selector: tabButtonSelector("inventory", "history"),
        icon: "clock",
        title: "Stock history",
        text: "Every change to stock: opening stock, sales, purchases, cancellations, and adjustments, with who made it and the stock left after.",
      },
      { selector: "#stockTypeFilter", icon: "box", title: "Show one kind of change", text: "Filter to only sales, purchases, adjustments, or opening stock." },
      { selector: "#stockHistoryWrap", icon: "info", title: "Trace any number", text: "Click a sale or purchase number to open it. Adjustment notes can be printed." },
    ],
    "inventory:missing": missingInfoGuide("inventory"),
    "sales:all": sales.all,
    "sales:unpaid": sales.unpaid,
    "sales:payments": sales.payments,
    "sales:cancelled": sales.cancelled,
    "purchases:all": purchases.all,
    "purchases:unpaid": purchases.unpaid,
    "purchases:payments": purchases.payments,
    "purchases:cancelled": purchases.cancelled,
    "contacts:customers": [
      {
        selector: tabButtonSelector("contacts", "customers"),
        icon: "users",
        title: "Customers",
        text: "People and companies who buy from you, with what they currently owe.",
      },
      {
        selector: "#contactsView [data-add='contacts']",
        icon: "info",
        title: "Add a contact",
        text: "Add a name, type, and phone. Customers are also saved automatically when you add a new customer during a sale.",
      },
      {
        selector: "#contactsTable",
        icon: "wallet",
        title: "History and payments",
        text: "History shows everything bought, sold, and paid with this contact. Record payment settles what they owe.",
      },
    ],
    "contacts:suppliers": [
      {
        selector: tabButtonSelector("contacts", "suppliers"),
        icon: "users",
        title: "Suppliers",
        text: "Suppliers and dealers you buy stock from, with what you owe each of them.",
      },
      { selector: "#contactsTable", icon: "wallet", title: "What you owe them", text: "History shows every purchase and payment. Record payment pays what you owe." },
    ],
    "contacts:all": [
      { selector: tabButtonSelector("contacts", "all"), icon: "users", title: "All contacts", text: "Customers and suppliers together." },
    ],
    "contacts:missing": missingInfoGuide("contacts"),
    "employees:directory": [
      {
        selector: tabButtonSelector("employees", "directory"),
        icon: "users",
        title: "Employee directory",
        text: "Everyone who works for the business: job, department, contact details, and salary.",
      },
      {
        selector: "#employeesView [data-add='employees']",
        icon: "info",
        title: "Add an employee",
        text: "Add their details and salary. To let them sign in, give them a login email, password, and access role. An employee ID is created automatically.",
      },
      { selector: "#employeeDepartmentFilter", icon: "clock", title: "Filter by department", text: "Show one team at a time." },
      { selector: "#employeesTable", icon: "wallet", title: "Pay history", text: "Pay history lists every salary paid to that employee, with slips." },
    ],
    "employees:payroll": [
      {
        selector: tabButtonSelector("employees", "payroll"),
        icon: "wallet",
        title: "Payroll",
        text: "Who has been paid and who is still unpaid for each month.",
      },
      { selector: "#payrollMonthSelect", icon: "clock", title: "Pick a month", text: "Earlier months stay here, so unpaid salaries from past months are never forgotten." },
      {
        selector: "#payrollTable",
        icon: "income",
        title: "Pay salaries",
        text: "Pay salary lets you add a bonus or deduction and creates a salary slip. The payment is also recorded as a Salaries expense in Finance.",
      },
      { selector: "#payAllSalariesBtn", icon: "users", title: "Pay everyone at once", text: "Pays all unpaid salaries for the month, each with its own slip." },
    ],
    "employees:history": [
      {
        selector: tabButtonSelector("employees", "history"),
        icon: "clock",
        title: "Pay history",
        text: "Every salary payment, newest first, with the slip number, method, and any bonus or deduction.",
      },
      { selector: "#payHistoryEmployeeFilter", icon: "users", title: "One employee", text: "Choose an employee to see only their payments." },
      {
        selector: "#payHistoryTable",
        icon: "info",
        title: "Slips and corrections",
        text: "Salary slip prints the slip again. Undo reverses a mistaken payment with a reason; the record stays marked Reversed.",
      },
    ],
    "employees:access": [
      {
        selector: tabButtonSelector("employees", "access"),
        icon: "tool",
        title: "Logins & access",
        text: "Which employees can sign in, the email they use, and the role that decides what they see.",
      },
      {
        selector: "#accessTable",
        icon: "users",
        title: "Set up a login",
        text: "Set up login adds an email, password, and role. Give the email and password to the employee.",
      },
    ],
    "employees:missing": missingInfoGuide("employees"),
    "assets:all": [
      {
        selector: tabButtonSelector("assets", "all"),
        icon: "box",
        title: `All ${viewTitle("assets").toLowerCase()}`,
        text: "Equipment, vehicles, furniture, and anything else the business owns and uses, with who has it and its value.",
      },
      {
        selector: "#assetsView [data-add='assets']",
        icon: "info",
        title: "Add an asset",
        text: "Record what it is, who it is assigned to, where it is, and its value. The asset code is filled in automatically if blank.",
      },
    ],
    "assets:maintenance": [
      {
        selector: tabButtonSelector("assets", "maintenance"),
        icon: "alert",
        title: "Needs maintenance",
        text: "Assets marked Maintenance or Damaged. Edit an asset and set it back to Active when it is repaired.",
      },
    ],
    "assets:missing": missingInfoGuide("assets"),
    "setup:business": [
      {
        selector: tabButtonSelector("setup", "business"),
        icon: "tool",
        title: "Business",
        text: "Edit your business information. It appears at the top of every invoice, receipt, and report.",
      },
      {
        selector: "#businessNameInput",
        icon: "info",
        title: "Name, type, and currency",
        text: "The business type arranges section names and fields to suit you. The default currency is used for new records only.",
      },
      {
        selector: "#businessPhoneInput",
        icon: "info",
        title: "Details for documents",
        text: "Phone, email, tax number, and address are printed on invoices and receipts, so keep them accurate.",
      },
      {
        selector: "#setupModuleChoices",
        icon: "box",
        title: "Work areas",
        text: "Switch sections on or off. Turning one off hides it from everyone without deleting its records.",
      },
      { selector: "#setupChecklist", icon: "alert", title: "Launch checklist", text: "What is left to set up before your team starts using LedgerFlow." },
    ],
    "setup:roles": [
      {
        selector: tabButtonSelector("setup", "roles"),
        icon: "users",
        title: "Roles & access",
        text: "Create roles for managers, cashiers, accountants, or any job, and choose exactly which data each role can see and change.",
      },
      {
        selector: "#setupView .role-builder",
        icon: "info",
        title: "Create a role",
        text: "Name the role and start from an existing one to copy its permissions, then adjust.",
      },
      {
        selector: "#roleCards",
        icon: "tool",
        title: "Choose permissions",
        text: "For each section, tick View, Add, Edit, or Delete. Anything a role can't view is hidden completely from people with that role.",
      },
      {
        selector: '#employeesTabs, .nav-item[data-view="employees"]',
        icon: "users",
        title: "Give roles to people",
        text: "Assign a role to each employee in Employees → Logins & access, along with their login email and password.",
      },
    ],
    "setup:fields": [
      {
        selector: tabButtonSelector("setup", "fields"),
        icon: "box",
        title: "Fields",
        text: "Shape the forms around your business: rename base fields, hide ones you don't use, or add your own.",
      },
      { selector: "#fieldModuleSelect", icon: "info", title: "Pick a section", text: "Choose which section's fields you want to change." },
      { selector: "#fieldList", icon: "tool", title: "Existing fields", text: "Rename or hide base fields. Fields with a lock are needed for LedgerFlow to work and can only be renamed." },
      {
        selector: "#fieldEditor",
        icon: "income",
        title: "Add your own field",
        text: "For example Batch Number, Expiry Date, or Project Site. Choose text, number, date, or a list of choices, and whether it is required.",
      },
    ],
    "setup:templates": [
      {
        selector: tabButtonSelector("setup", "templates"),
        icon: "box",
        title: "Templates & samples",
        text: "Try LedgerFlow with realistic sample data, or apply a business template.",
      },
      {
        selector: "#dataToolsPanel",
        icon: "alert",
        title: "Sample data",
        text: "Loading a sample replaces current records, so use it to explore before entering real data.",
      },
    ],
    "dataSources:": [
      {
        selector: "#dataSourcesView .page-head",
        icon: "inbox",
        title: "Import data",
        text: "Bring in records you already keep in spreadsheets, instead of typing them again.",
      },
      {
        selector: "#dataSourcesView .import-box",
        icon: "info",
        title: "Choose where and what",
        text: "Pick the section to import into, then choose a CSV, Excel, JSON, or tab-separated file.",
      },
      {
        selector: "#canonicalModelMap",
        icon: "box",
        title: "How columns are matched",
        text: "Your columns are matched to these fields. You can change the matching, and LedgerFlow asks for anything required that is missing.",
      },
      { selector: "#sampleCsvBtn", icon: "inbox", title: "Not sure of the format?", text: "Download a sample CSV to see which columns work best." },
    ],
    "reports:": [
      { selector: "#reportCards", icon: "wallet", title: "Reports", text: "A snapshot of the business: money, stock, dues, and team." },
      {
        selector: "#tradeReportPanel",
        icon: "income",
        title: "Sales, purchases & stock",
        text: "Best-selling items, top customers, spending by supplier, gross profit, and stock movement for the chosen period.",
      },
      { selector: "#reportPeriodSelect", icon: "clock", title: "Choose the period", text: "This month, last month, this year, or all time." },
      { selector: "#printReportBtn", icon: "info", title: "Print or save as PDF", text: "Print the report for meetings, or save it as a PDF from the print window." },
    ],
  };
}

function currentGuideKey(view = currentView) {
  return `${view}:${activeTab(view) || ""}`;
}

function guideStepsFor(key) {
  return (guideDefinitions()[key] || []).filter((step) => {
    const target = document.querySelector(step.selector);
    return target && target.getClientRects().length;
  });
}

function guideSeen(key) {
  const user = currentUser();
  return Boolean(user && state.onboarding?.tabGuides?.[user.id]?.[key]);
}

function markGuideSeen(key, status) {
  const user = currentUser();
  if (!user) return;
  state.onboarding.tabGuides ||= {};
  state.onboarding.tabGuides[user.id] ||= {};
  state.onboarding.tabGuides[user.id][key] = status;
  saveState();
}

function autoGuidesOff() {
  const user = currentUser();
  return Boolean(user && state.onboarding?.autoGuidesOff?.[user.id]);
}

function startPageGuide({ automatic = false } = {}) {
  if (!currentUser() || !els.setupWizardBackdrop.hidden || !els.tourLayer.hidden) return false;
  const key = currentGuideKey();
  const steps = guideStepsFor(key);
  if (!steps.length) {
    if (!automatic) showToast("There is no guide for this view yet");
    return false;
  }
  tutorialSteps = steps;
  tutorialStep = 0;
  tutorialOriginView = currentView;
  tutorialFollowUp = false;
  tutorialMode = { type: "guide", key };
  els.tourLayer.hidden = false;
  renderTutorialStep();
  return true;
}

let guideTimer = null;

function scheduleAutoGuide() {
  window.clearTimeout(guideTimer);
  guideTimer = window.setTimeout(() => {
    const user = currentUser();
    if (!user || previewRoleId || autoGuidesOff()) return;
    if (!state.onboarding?.tutorials?.[user.id]) return;
    if (!els.modalBackdrop.hidden || !els.tourLayer.hidden || !els.setupWizardBackdrop.hidden) return;
    const key = currentGuideKey();
    if (guideSeen(key) || !guideDefinitions()[key]) return;
    startPageGuide({ automatic: true });
  }, 450);
}

function addGuideButtons() {
  document.querySelectorAll(".view > .page-head").forEach((head) => {
    if (head.querySelector("[data-page-guide]")) return;
    let actions = head.querySelector(".page-actions");
    if (!actions) {
      actions = document.createElement("div");
      actions.className = "page-actions";
      head.appendChild(actions);
    }
    const button = document.createElement("button");
    button.type = "button";
    button.className = "button ghost guide-button";
    button.dataset.pageGuide = "";
    button.title = "Explain this page and tab";
    button.innerHTML = `<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9" /><path d="M9.5 9a2.5 2.5 0 1 1 3.5 2.3c-.6.3-1 .9-1 1.6V14M12 17h.01" /></svg><span>Guide</span>`;
    actions.prepend(button);
  });
}

function beginNewUserFlow() {
  const user = currentUser();
  if (!user || !els.modalBackdrop.hidden || !els.tourLayer.hidden || !els.setupWizardBackdrop.hidden) return;
  if (user.isOwner && !state.onboarding.setupCompleted) {
    openSetupWizard();
    return;
  }
  if (!state.onboarding.tutorials[user.id]) startTutorial({ automatic: true });
}

function buildTutorialSteps() {
  const user = currentUser();
  const firstModule = MODULES.find((module) => can(module, "view"));
  const firstAddable = MODULES.find((module) => can(module, "add"));
  const steps = [
    {
      selector: "#dashboardView .page-head",
      view: "dashboard",
      icon: "info",
      title: "Your daily starting point",
      text: "The dashboard turns your records into totals, trends, dues, alerts, and recent activity. Start here whenever you want the overall picture.",
    },
    {
      selector: "#navList",
      view: "dashboard",
      icon: "box",
      title: "Everything is grouped in the sidebar",
      text: "Open finance, items, contacts, employees, assets, imports, reports, and settings from here. Each person only sees areas allowed by their role.",
    },
  ];
  if (firstModule) {
    steps.push({
      selector: `.nav-item[data-view="${firstModule}"]`,
      view: "dashboard",
      icon: firstModule === "finance" ? "wallet" : firstModule === "employees" ? "users" : "box",
      title: `Open ${navTitle(firstModule)} here`,
      text: "Each work area has its own overview, filters, alerts, and records. Red counts point to overdue, low-stock, or incomplete information that needs attention.",
    });
  }
  if (canOpenView("sales")) {
    steps.push({
      selector: '.nav-item[data-view="sales"]',
      view: "dashboard",
      icon: "income",
      title: "Record every sale here",
      text: "A sale records what was sold, to whom, the price, and how it was paid. Stock goes down, income is added, and anything unpaid is tracked as a due automatically.",
    });
  }
  if (canOpenView("purchases")) {
    steps.push({
      selector: '.nav-item[data-view="purchases"]',
      view: "dashboard",
      icon: "expense",
      title: "Record stock that comes in",
      text: "A purchase records which supplier the stock came from and what it cost. Stock goes up, the expense is added, and unpaid supplier bills appear under Dues.",
    });
  }
  if (firstAddable) {
    steps.push({
      selector: "#quickAddBtn",
      view: "dashboard",
      icon: "info",
      title: "Add records from anywhere",
      text: `Use this shortcut to create a ${recordNoun(firstAddable).toLowerCase()}. Inside each work area, its own add button gives you the same guided form.`,
    });
  }
  if (canOpenView("dataSources")) {
    steps.push({
      selector: '.nav-item[data-view="dataSources"]',
      view: "dashboard",
      icon: "inbox",
      title: "Bring in your existing records",
      text: "Import CSV, Excel, JSON, or tab-separated files. LedgerFlow helps match columns and pauses for any required details that are missing.",
    });
  }
  steps.push({
    selector: "#sidebarHealthBox",
    view: "dashboard",
    icon: "alert",
    title: "Important work stays visible",
    text: "This status summarizes items needing attention. Red markers are reserved for overdue, critical, or incomplete records so they are difficult to miss.",
  });
  if (user?.isOwner && canOpenView("setup")) {
    steps.push({
      selector: '.nav-item[data-view="setup"]',
      view: "dashboard",
      icon: "tool",
      title: "Shape LedgerFlow around your business",
      text: "In Setup & Roles you can update business details, turn work areas on or off, create staff roles, control permissions, and customize fields.",
    });
  }
  steps.push({
    selector: "#tutorialBtn",
    view: "dashboard",
    icon: "info",
    title: "Replay this tutorial any time",
    text: "Use the Tutorial button whenever you need a reminder. It will always explain the controls available to your current login and role.",
  });
  return steps;
}

function startTutorial({ automatic = false, followUp = false } = {}) {
  if (!currentUser() || !els.setupWizardBackdrop.hidden) return;
  tutorialSteps = buildTutorialSteps();
  if (!tutorialSteps.length) return;
  tutorialStep = 0;
  tutorialOriginView = currentView;
  tutorialFollowUp = Boolean(automatic && followUp);
  tutorialMode = { type: "main" };
  els.tourLayer.hidden = false;
  renderTutorialStep();
}

function renderTutorialStep() {
  const step = tutorialSteps[tutorialStep];
  if (!step) return finishTutorial("completed");
  if (step.view && currentView !== step.view) setView(step.view);

  const guide = tutorialMode.type === "guide";
  els.tourStepCount.textContent = `${guide ? "Guide · " : ""}${tutorialStep + 1} of ${tutorialSteps.length}`;
  byId("skipTutorialBtn").textContent = guide ? "Skip guide" : "Skip tutorial";
  byId("muteGuidesBtn").hidden = !guide;
  els.tourIcon.innerHTML = ICONS[step.icon] || ICONS.info;
  els.tourTitle.textContent = step.title;
  els.tourText.textContent = step.text;
  els.tourBackBtn.hidden = tutorialStep === 0;
  els.tourNextBtn.textContent = tutorialStep === tutorialSteps.length - 1 ? "Finish" : "Next";
  els.tourProgress.innerHTML = tutorialSteps
    .map((_, index) => `<span class="${index <= tutorialStep ? "is-done" : ""}"></span>`)
    .join("");

  const target = document.querySelector(step.selector) || document.querySelector(".main-area");
  const inSidebar = Boolean(target?.closest(".sidebar"));
  document.body.classList.toggle("menu-open", window.innerWidth <= 900 && inSidebar);
  els.tourCard.style.opacity = "0";
  window.clearTimeout(tutorialPositionTimer);
  tutorialPositionTimer = window.setTimeout(() => positionTutorial(target), inSidebar && window.innerWidth <= 900 ? 220 : 40);
}

function positionTutorial(target) {
  if (els.tourLayer.hidden || !target) return;
  target.scrollIntoView({ block: "nearest", inline: "nearest" });
  const rect = target.getBoundingClientRect();
  const pad = 6;
  const left = Math.max(6, rect.left - pad);
  const top = Math.max(6, rect.top - pad);
  const right = Math.min(window.innerWidth - 6, rect.right + pad);
  const bottom = Math.min(window.innerHeight - 6, rect.bottom + pad);
  Object.assign(els.tourSpotlight.style, {
    left: `${left}px`,
    top: `${top}px`,
    width: `${Math.max(24, right - left)}px`,
    height: `${Math.max(24, bottom - top)}px`,
  });

  const cardWidth = Math.min(360, window.innerWidth - 24);
  const cardHeight = els.tourCard.offsetHeight;
  const gap = 14;
  let cardLeft;
  let cardTop;
  if (right + gap + cardWidth <= window.innerWidth - 12) {
    cardLeft = right + gap;
    cardTop = Math.min(top, window.innerHeight - cardHeight - 12);
  } else if (left - gap - cardWidth >= 12) {
    cardLeft = left - gap - cardWidth;
    cardTop = Math.min(top, window.innerHeight - cardHeight - 12);
  } else if (bottom + gap + cardHeight <= window.innerHeight - 12) {
    cardLeft = Math.min(Math.max(12, left), window.innerWidth - cardWidth - 12);
    cardTop = bottom + gap;
  } else {
    cardLeft = Math.min(Math.max(12, left), window.innerWidth - cardWidth - 12);
    cardTop = Math.max(12, top - gap - cardHeight);
  }
  els.tourCard.style.left = `${Math.max(10, cardLeft)}px`;
  els.tourCard.style.top = `${Math.max(10, cardTop)}px`;
  els.tourCard.style.opacity = "1";
}

function moveTutorial(direction) {
  const next = tutorialStep + direction;
  if (next < 0) return;
  if (next >= tutorialSteps.length) {
    finishTutorial("completed");
    return;
  }
  tutorialStep = next;
  renderTutorialStep();
}

function finishTutorial(status) {
  const user = currentUser();
  if (tutorialMode.type === "guide") {
    markGuideSeen(tutorialMode.key, status);
    tutorialMode = { type: "main" };
    els.tourLayer.hidden = true;
    document.body.classList.remove("menu-open");
    window.clearTimeout(tutorialPositionTimer);
    return;
  }
  if (user) {
    // The main tutorial already explains the dashboard.
    markGuideSeen("dashboard:", status);
    state.onboarding.tutorials[user.id] = { status, completedAt: new Date().toISOString() };
    saveState();
  }
  els.tourLayer.hidden = true;
  document.body.classList.remove("menu-open");
  window.clearTimeout(tutorialPositionTimer);
  if (tutorialFollowUp) {
    tutorialFollowUp = false;
    runStartingPlan();
  } else if (canOpenView(tutorialOriginView)) {
    setView(tutorialOriginView);
  }
}

function runStartingPlan() {
  if (state.onboarding.addTeam && can("employees", "add")) {
    setView("employees");
    window.setTimeout(() => openModal("employees"), 80);
    return;
  }
  if (state.onboarding.startPath === "import" && canOpenView("dataSources")) {
    setView("dataSources");
    return;
  }
  if (state.onboarding.startPath === "sample" && currentUser()?.isOwner) {
    const type = SAMPLE_BUSINESSES.includes(state.organization.type) ? state.organization.type : "Utility Store";
    loadSampleBusiness(type);
    return;
  }
  const module = MODULES.find((entry) => can(entry, "add"));
  if (module) {
    setView(module);
    window.setTimeout(() => openModal(module), 80);
  }
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
  if (PERMISSION_AREAS.includes(module)) return can(module, "view");
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
  ["finance", "sales", "purchases", "inventory", "contacts", "employees", "assets", "setup"].forEach((view) => {
    if (canOpenView(view)) renderTabs(view);
  });
  if (canAccess("settings")) renderSetup();
  renderDashboard();
  if (can("finance", "view")) renderFinance();
  if (can("sales", "view")) renderTrades("sale");
  if (can("purchases", "view")) renderTrades("purchase");
  if (can("inventory", "view")) {
    if (activeTab("inventory") === "history") renderStockHistory();
    else renderModule("inventory");
  }
  ["contacts", "assets"].forEach((module) => {
    if (can(module, "view")) renderModule(module);
  });
  if (can("employees", "view")) renderEmployeesView();
  if (canOpenView("dataSources")) renderDataSources();
  if (canAccess("reports")) {
    renderReports();
    renderTradeReports();
  }
  refreshSearchableSelects();
  scheduleAutoGuide();
}

function renderChrome() {
  const user = currentUser();
  const role = currentRole();
  const org = state.organization;
  els.exitTrialBtn.hidden = !user?.isGuest;
  els.trialBanner.hidden = !user?.isGuest;

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
  TRADE_NAV_COUNTS.forEach(([kind, area]) => {
    setCount(
      area,
      can(area, "view")
        ? state[area].filter((trade) => tradeBalance(trade) > 0 && trade.dueDate && trade.dueDate < todayIso()).length
        : 0,
    );
  });
  document.querySelectorAll("[data-new-trade]").forEach((button) => {
    button.hidden = !can(TRADE[button.dataset.newTrade].area, "add");
  });
  setCount(
    "employees",
    can("employees", "view") && canAccess("sensitiveNumbers")
      ? state.employees.filter((employee) => salaryMonthsDue(employee).some((month) => month < todayIso().slice(0, 7))).length
      : 0,
  );

  els.sidebarBusinessName.textContent = org.name;
  byId("workspaceAvatar").textContent =
    String(org.name || "")
      .split(/\s+/)
      .filter(Boolean)
      .slice(0, 2)
      .map((word) => word.charAt(0).toUpperCase())
      .join("") || "CO";
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
  els.quickAddBtn.hidden = !target || PERMISSION_AREAS.includes(currentView);
  if (target) {
    const label = `Add ${recordNoun(target)}`;
    els.quickAddLabel.textContent = label;
    els.quickAddBtn.title = label;
    els.quickAddBtn.setAttribute("aria-label", label);
  }
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
      liveFinance().filter((entry) => !typeForTab || entry.type === typeForTab),
      "category",
    ),
  );

  const monthSelect = byId("financeMonthFilter");
  const wanted = financeViewState.monthFilter || monthSelect.value || "all";
  const months = [...new Set(liveFinance().map((entry) => monthKeyOf(entry.date)).filter((key) => /^\d{4}-\d{2}$/.test(key)))]
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
  els.countryInput.value = org.country || "";
  if (org.timezone && ![...els.timezoneSelect.options].some((option) => option.value === org.timezone)) {
    els.timezoneSelect.add(new Option(org.timezone, org.timezone));
  }
  els.timezoneSelect.value = org.timezone || "UTC";
  els.fiscalYearStartSelect.value = org.fiscalYearStart || "January";
  els.businessPhoneInput.value = org.businessPhone || "";
  els.businessEmailInput.value = org.businessEmail || "";
  els.registrationNumberInput.value = org.registrationNumber || "";
  els.businessAddressInput.value = org.businessAddress || "";
  els.ownerNameInput.value = org.ownerName || "";
  els.ownerEmailInput.value = org.ownerEmail || "";
  els.ownerPasswordInput.value = "";
  els.setupModuleChoices.innerHTML = moduleChoicesHtml(org.enabledModules || PERMISSION_AREAS,"profileEnabledModule");
}

function renderSetup() {
  const isOwner = Boolean(currentUser()?.isOwner && !state.demoMode);
  els.ownerFields.hidden = !isOwner;
  els.ownerFields.querySelectorAll("input").forEach((input) => {
    input.disabled = !isOwner;
  });
  els.businessPhoneInput.required = !state.demoMode;
  els.businessAddressInput.required = !state.demoMode;
  els.dataToolsPanel.hidden = !currentUser()?.isOwner || Boolean(previewRoleId);
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
    MODULES.some((module) => moduleEnabled(module) && state.customFields[module].length) ||
    MODULES.some((module) => moduleEnabled(module) && Object.keys(state.fieldSettings[module] || {}).length);
  return [
    {
      label: state.demoMode ? "Trial workspace ready" : "Business registered",
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
    moduleEnabled("employees") && !state.demoMode && {
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
      done: ["inventory", "finance", "contacts", "assets"].some((module) => moduleEnabled(module) && state[module].length),
      detail: "Import your existing files, or start adding records by hand.",
      view: "dataSources",
      action: "Import data",
      recordsStep: true,
    },
  ].filter(Boolean);
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
      const matrix = PERMISSION_AREAS.map(
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
  const isOwner = Boolean(currentUser()?.isOwner && !state.demoMode);
  const name = els.businessNameInput.value.trim();
  const currency = els.currencyInput.value.trim().toUpperCase();
  const enabledModules = [...els.setupModuleChoices.querySelectorAll('input[name="profileEnabledModule"]:checked')]
    .map((input) => input.value)
    .filter((module) => PERMISSION_AREAS.includes(module));
  if (!name) {
    showToast("Business name is required");
    return;
  }
  if (!isValidCurrency(currency)) {
    showToast("Currency must be a 3-letter code like PKR or USD");
    return;
  }
  if (!enabledModules.length) {
    showToast("Choose at least one active work area");
    return;
  }

  const updates = {
    name,
    type: els.businessTypeSelect.value,
    currency,
    dateFormat: els.dateFormatSelect.value,
    country: els.countryInput.value.trim(),
    timezone: els.timezoneSelect.value,
    fiscalYearStart: els.fiscalYearStartSelect.value,
    businessPhone: els.businessPhoneInput.value.trim(),
    businessEmail: els.businessEmailInput.value.trim().toLowerCase(),
    registrationNumber: els.registrationNumberInput.value.trim(),
    businessAddress: els.businessAddressInput.value.trim(),
    enabledModules,
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
  return liveFinance().reduce(
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
    liveFinance().filter((entry) => entry.type === "Income"),
    (entry) => parseMoney(entry.amount),
  );
  const expense = sumByCurrency(
    liveFinance().filter((entry) => entry.type === "Expense"),
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
  const unpaid = can("finance", "view") ? liveFinance().filter((entry) => financeStatus(entry) === "Overdue").length : 0;
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
  return state.finance.find((entry) => entry.employeeId === employee.id && entry.payrollMonth === month && !entry.cancelled);
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
    liveFinance().filter(isUnpaid).forEach((entry) => {
      const item = {
        kind: "transaction",
        id: entry.id,
        party: entry.party || "",
        title: entry.description || entry.category || "Transaction",
        detail: [entry.category, parseMoney(entry.amountPaid) ? `${formatCurrency(entry.amountPaid, recordCurrency(entry))} already paid` : ""]
          .filter(Boolean)
          .join(" · "),
        amount: outstandingOf(entry),
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
  if (item.kind === "transaction" && canTakePayment(state.finance.find((entry) => entry.id === item.id))) {
    action = `<button type="button" class="accent" data-mark-paid="${escapeHtml(item.id)}">${direction === "in" ? "Receive payment" : "Pay"}</button>`;
  } else if (item.kind === "balance" && can("finance", "add")) {
    action = `<button type="button" class="accent" data-record-payment="${escapeHtml(item.id)}">Record payment</button>`;
  } else if (item.kind === "salary" && canOpenView("employees")) {
    action = `<button type="button" class="accent" data-open-payroll="${item.month}">Payroll</button>`;
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
    const { currency } = chartCurrency(liveFinance());
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
  const { currency, excluded } = chartCurrency(liveFinance());
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
    liveFinance().filter((entry) => entry.type === "Income" && monthKeyOf(entry.date) === month),
    (entry) => parseMoney(entry.amount),
  );

  if (can("finance", "view") && liveFinance().length) {
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
      liveFinance().filter((entry) => entry.type === "Income" && financeStatus(entry) === "Overdue"),
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

const TRADE_NAV_COUNTS = [
  ["sale", "sales"],
  ["purchase", "purchases"],
];

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
      { key: "income", label: "Income", count: liveFinance().filter((entry) => entry.type === "Income").length },
      { key: "expenses", label: "Expenses", count: liveFinance().filter((entry) => entry.type === "Expense").length },
      { key: "dues", label: "Dues", count: overdue, alert: true, hideZero: true },
      { key: "payments", label: "Payments", count: allPayments().length },
      { key: "all", label: "All transactions", count: state.finance.length },
    ];
  }
  if (view === "inventory") {
    const low = state.inventory.filter(isLowStock).length;
    return [
      { key: "all", label: `All ${viewTitle("inventory").toLowerCase()}`, count: state.inventory.length },
      { key: "low", label: "Low stock", show: tracksReorderLevel(), count: low, alert: true, hideZero: true },
      { key: "history", label: "Stock history", count: state.stockMovements.length },
      { key: "missing", label: "Needs info", show: missingCount("inventory") > 0, count: missingCount("inventory"), alert: true },
    ];
  }
  if (view === "sales" || view === "purchases") {
    const list = state[view];
    const live = list.filter((trade) => trade.status !== "Cancelled");
    const cancelled = list.length - live.length;
    return [
      { key: "all", label: view === "sales" ? "All sales" : "All purchases", count: live.length },
      { key: "unpaid", label: "Unpaid", count: live.filter((trade) => tradeBalance(trade) > 0).length, alert: true, hideZero: true },
      {
        key: "payments",
        label: view === "sales" ? "Payments received" : "Payments made",
        count: allPayments({ filter: (entry) => entry.sourceType === (view === "sales" ? "sale" : "purchase") }).length,
      },
      { key: "cancelled", label: "Cancelled", show: cancelled > 0, count: cancelled },
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
      {
        key: "history",
        label: "Pay history",
        show: canAccess("sensitiveNumbers"),
        count: state.finance.filter((entry) => isSalaryEntry(entry) && !entry.cancelled).length,
      },
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
  return [...new Set(liveFinance().map(recordCurrency))];
}

function renderFinance() {
  const tab = activeTab("finance");
  if (tab === "overview") renderFinanceOverview();
  else if (tab === "dues") renderDuesTab();
  else if (tab === "payments") renderPaymentsTab();
  else renderModule("finance");
}

function renderFinanceOverview() {
  const currencies = financeCurrencies();
  if (!currencies.includes(financeViewState.currency)) financeViewState.currency = chartCurrency(liveFinance()).currency;
  const currency = financeViewState.currency;
  const currentYear = Number(todayIso().slice(0, 4));
  const years = [
    ...new Set([
      currentYear,
      ...liveFinance()
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
  const otherCount = liveFinance().filter((entry) => recordCurrency(entry) !== currency).length;
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
  liveFinance().forEach((entry) => {
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
  byId("receivablesList").innerHTML = duePartyListHtml(dues.receivables, "in");
  byId("payablesList").innerHTML = duePartyListHtml(dues.payables, "out");
}

function markPaid(id) {
  openPaymentModal(id);
}

function recordPayment(contactId) {
  const contact = state.contacts.find((record) => record.id === contactId);
  if (!contact || !can("finance", "add")) return;
  const balance = parseMoney(contact.balance);
  if (!balance) {
    showToast("No opening balance to settle. Mark unpaid transactions as paid in Finance → Dues.");
    return;
  }
  // Settle through the person's dues so opening balances and unpaid invoices are paid oldest first.
  openPartyPayment(balance > 0 ? "in" : "out", `${String(contact.name || "").trim().toLowerCase()}|${recordCurrency(contact)}`);
}

/* ---------- Employees: payroll and access ---------- */

let payrollMonth = null;

function renderEmployeesView() {
  const tab = activeTab("employees");
  if (tab === "payroll") renderPayroll();
  else if (tab === "history") renderPayHistory();
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
                if (payment) {
                  actions.push(`<button type="button" data-document="salary:${escapeHtml(payment.id)}">Salary slip</button>`);
                  if (can("finance", "edit")) {
                    actions.push(`<button type="button" class="warn" data-reverse-salary="${escapeHtml(payment.id)}">Undo</button>`);
                  }
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
    liveFinance().forEach((entry) => {
      if (!isUnpaid(entry) || String(entry.party || "").trim().toLowerCase() !== name) return;
      const code = recordCurrency(entry);
      const signed = entry.type === "Income" ? outstandingOf(entry) : -outstandingOf(entry);
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
    {
      field: "name",
      render: (employee) => titleCell(employee.name, [employee.employeeId, employee.email || employee.phone].filter((part) => !isBlank(part)).join(" · ")),
    },
    { field: "role" },
    { field: "department" },
    { field: "phone" },
    { field: "salary", className: "num", render: (employee) => escapeHtml(formatRecordMoney(employee, employee.salary)) },
  ],
  finance: [
    { field: "date", render: (entry) => escapeHtml(formatDate(entry.date)) },
    {
      field: "description",
      render: (entry) =>
        titleCell(
          entry.description,
          [entry.sourceType ? "" : entry.reference, entry.category, entry.party, entry.sourceType ? `From ${entry.sourceType} ${entry.reference || ""}`.trim() : ""]
            .filter((part) => !isBlank(part))
            .join(" · "),
        ),
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
    { field: "name", render: (contact) => titleCell(contact.name, [contact.code, contact.phone, contact.email].filter(Boolean).join(" · ")) },
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
  if (module === "finance") {
    if (record.cancelled) parts.push(badge("Cancelled", "red"));
    else if (isUnpaid(record) && parseMoney(record.amountPaid)) {
      parts.push(badge(financeStatus(record) === "Overdue" ? "Part paid · overdue" : "Part paid", financeStatus(record) === "Overdue" ? "red" : "amber"));
    } else parts.push(statusBadge(financeStatus(record)));
  }
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
        if (tab !== "all" && entry.cancelled) return false;
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
      rows.filter((entry) => entry.type === "Income" && !entry.cancelled),
      (entry) => parseMoney(entry.amount),
    );
    const expense = sumByCurrency(
      rows.filter((entry) => entry.type === "Expense" && !entry.cancelled),
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
  if (module === "finance" && !record.cancelled && isUnpaid(record) && canTakePayment(record)) {
    actions.push(
      `<button type="button" class="accent" data-mark-paid="${escapeHtml(record.id)}">${record.type === "Income" ? "Receive payment" : "Pay"}</button>`,
    );
  }
  if (module === "finance" && isSalaryEntry(record)) {
    // Salary payments are managed from Employees → Payroll.
    if (canOpenView("employees")) actions.push(`<button type="button" data-open-payroll="${escapeHtml(record.payrollMonth)}">Payroll</button>`);
    actions.push(`<button type="button" data-document="salary:${escapeHtml(record.id)}">Salary slip</button>`);
    return `<div class="row-actions">${actions.join("")}</div>`;
  }
  if (module === "finance" && record.settlesContactId) {
    // A payment against an opening balance; its receipt is the record of it.
    actions.push(`<button type="button" data-document="finance:${escapeHtml(record.id)}">Receipt</button>`);
    return `<div class="row-actions">${actions.join("")}</div>`;
  }
  if (module === "employees" && canAccess("sensitiveNumbers")) {
    actions.push(`<button type="button" data-pay-history="${escapeHtml(record.id)}">Pay history</button>`);
  }
  if (module === "finance" && record.sourceType) {
    // Entries created by a sale or purchase are changed only through that sale or purchase.
    const area = record.sourceType === "sale" ? "sales" : "purchases";
    if (can(area, "view")) {
      actions.push(
        `<button type="button" data-view-trade="${record.sourceType}:${escapeHtml(record.sourceId)}">Open ${escapeHtml(record.reference || record.sourceType)}</button>`,
        `<button type="button" data-document="finance:${escapeHtml(record.id)}">Invoice</button>`,
      );
    }
    return `<div class="row-actions">${actions.join("")}</div>`;
  }
  if (module === "finance") {
    actions.push(`<button type="button" data-document="finance:${escapeHtml(record.id)}">${record.type === "Income" ? "Receipt" : "Voucher"}</button>`);
  }
  if (module === "inventory") {
    if (can("sales", "add") && parseMoney(record.quantity) > 0) {
      actions.push(`<button type="button" class="accent" data-sell-item="${escapeHtml(record.id)}">Sell</button>`);
    }
    if (can("inventory", "edit")) {
      actions.push(`<button type="button" data-adjust-stock="${escapeHtml(record.id)}">Adjust stock</button>`);
    }
    actions.push(`<button type="button" data-stock-history="${escapeHtml(record.id)}">History</button>`);
  }
  if (module === "contacts" && parseMoney(record.balance) && can("finance", "add")) {
    actions.push(`<button type="button" class="accent" data-record-payment="${escapeHtml(record.id)}">Record payment</button>`);
  }
  if (module === "contacts" && (can("sales", "view") || can("purchases", "view"))) {
    actions.push(`<button type="button" data-contact-history="${escapeHtml(record.id)}">History</button>`);
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
  setModalWide(false);
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
      module === "finance" && !record && !settling && (can("sales", "add") || can("purchases", "add"))
        ? `<div class="form-notice">
            <strong>Selling or buying stock?</strong>
            <span>Use ${can("sales", "add") ? `<button class="link-button" type="button" data-new-trade="sale">New sale</button>` : ""}${
              can("sales", "add") && can("purchases", "add") ? " or " : ""
            }${can("purchases", "add") ? `<button class="link-button" type="button" data-new-trade="purchase">New purchase</button>` : ""} instead, so stock, the customer or supplier, and dues are updated together. Use this form for other income and expenses such as rent or bills.</span>
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
    ${formHelpBarHtml(
      id
        ? "Each field explains what it is for. Fields marked * are required."
        : `Fill in what you know. Fields marked * are required; ID fields left blank are numbered automatically.`,
    )}
    ${formSections(module)
      .map(
        ([name, fields]) => `
        <fieldset class="form-section">
          <legend>${escapeHtml(name)}</legend>
          ${sectionHelpHtml(SECTION_HELP[name] || (name === customSectionName() ? "Extra details your business added for this kind of record." : ""))}
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
  applyFieldHelpVisibility();
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
  const module = editing?.module || "";
  const autoId = field.autoId && AUTO_ID_FIELDS[module] === field.name;
  const common = `id="${id}" name="${escapeHtml(field.name)}" ${attrs}${
    autoId && isBlank(value) ? ` placeholder="Automatic: ${escapeHtml(previewAutoId(module, { ...record, id: record.id || "" }))}"` : ""
  }`;
  const helpText = fieldHelpText(module, field);
  const hint = field.hint && !helpText ? `<small>${escapeHtml(field.hint)}</small>` : "";
  const wrap = (control) => `
    <div class="field ${field.full ? "full" : ""} ${isMissing ? "is-missing" : ""}">
      <div class="field-label-row">
        <label for="${id}">${escapeHtml(field.label)}${field.required ? " *" : ""}</label>
        ${helpText ? `<button type="button" class="field-help-toggle" data-field-help aria-label="What is ${escapeHtml(field.label)}?" title="What is this?">?</button>` : ""}
      </div>
      ${control}
      ${hint}
      ${helpText ? `<small class="field-help">${escapeHtml(helpText)}</small>` : ""}
    </div>
  `;

  if (field.name === "quantity" && !isNew && editing?.module === "inventory") {
    return `
      <div class="field">
        <label for="${id}">${escapeHtml(field.label)}</label>
        <input id="${id}" type="text" value="${escapeHtml(formatNumber(value))}" disabled />
        <small>Stock changes only through a sale, a purchase, or Adjust stock, so every change is recorded.</small>
      </div>
    `;
  }
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
  activeDocument = null;
  els.modalBackdrop.hidden = true;
  setModalWide(false);
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
  if (editing.kind) {
    submitTransactionForm();
    return;
  }
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
    // Stock on hand only changes through sales, purchases, and stock adjustments.
    if (module === "inventory" && existing && field.name === "quantity") return;
    const raw = String(formData.get(field.name) ?? "").trim();
    record[field.name] = field.type === "number" ? (raw === "" ? "" : parseMoney(raw)) : raw;
  });

  if (module === "inventory" && existing && record.currency !== recordCurrency(existing) && itemHasTrades(existing.id)) {
    showFormError(
      `This item already has sales or purchases in ${recordCurrency(existing)}. Its currency can't change, or its history would mix currencies.`,
    );
    return;
  }

  const clash = findIdClash(module, record);
  if (clash) {
    showFormError(
      `${fieldLabel(module, AUTO_ID_FIELDS[module])} "${record[AUTO_ID_FIELDS[module]]}" is already used by ${recordLabel(module, clash)}. Use another one, or leave it blank to get one automatically.`,
    );
    return;
  }
  if (AUTO_ID_FIELDS[module] && record[AUTO_ID_FIELDS[module]] !== undefined) {
    record[AUTO_ID_FIELDS[module]] = String(record[AUTO_ID_FIELDS[module]]).trim();
  }
  assignAutoId(module, record);

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
    if (module === "inventory") logOpeningStock(record, "Opening stock");
    touchManualSource();
    addActivity(`${recordNoun(module)} added: ${recordLabel(module, record)}`, module);
    const idField = AUTO_ID_FIELDS[module];
    showToast(`${recordNoun(module)} added${idField && record[idField] ? ` · ${record[idField]}` : ""}`);
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
  if (module === "inventory" && itemHasTrades(id)) {
    showToast("This item appears in sales or purchases, so it can't be deleted. Use Adjust stock to bring it to zero.");
    return;
  }
  if (module === "finance" && record.sourceType) {
    showToast(`This entry belongs to ${record.reference || `a ${record.sourceType}`}. Cancel the ${record.sourceType} instead.`);
    return;
  }
  if (module === "finance" && isSalaryEntry(record)) {
    showToast("This is a salary payment. Undo it from Employees → Pay history instead.");
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
        if (key === "currency" || isBlank(value)) return;
        if (module === "inventory" && key === "quantity") {
          const change = round3(parseMoney(value) - parseMoney(duplicate.quantity));
          if (change) {
            recordStockMovement(duplicate, change, { type: "Import update", reason: `Quantity from ${s.fileName}` });
          }
          return;
        }
        duplicate[key] = value;
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
    if (module === "inventory") logOpeningStock(record, "Opening stock (import)", s.fileName);
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

/* ---------- Sales, purchases, and stock movements ----------
 *
 * Every change to stock or money after records exist goes through a recorded
 * transaction:
 * - A sale lowers stock, creates an income entry, and leaves any unpaid part as a due.
 * - A purchase raises stock, updates the item's average cost, creates an expense
 *   entry, and leaves any unpaid part as a due.
 * - A stock adjustment changes stock with a reason (count correction, damage, ...).
 * Each stock change is logged in state.stockMovements with the balance after it.
 * Sales and purchases are never edited or deleted; they are cancelled, which
 * reverses their stock and removes their finance entry from totals.
 */

const TRADE = {
  sale: {
    area: "sales",
    list: "sales",
    noun: "Sale",
    party: "Customer",
    partyType: "Customer",
    priceField: "sellPrice",
    priceLabel: "Price",
    financeType: "Income",
    category: "Sales",
    movement: "Sale",
    reversal: "Sale cancelled",
    sign: -1,
  },
  purchase: {
    area: "purchases",
    list: "purchases",
    noun: "Purchase",
    party: "Supplier",
    partyType: "Supplier",
    priceField: "unitCost",
    priceLabel: "Cost",
    financeType: "Expense",
    category: "Stock purchases",
    movement: "Purchase",
    reversal: "Purchase cancelled",
    sign: 1,
  },
};

const tradeFilters = { sales: "all", purchases: "all", stockType: "all", reportPeriod: "this-month", reportCurrency: "" };

function liveFinance() {
  return state.finance.filter((entry) => !entry.cancelled);
}

function round3(value) {
  return Math.round(value * 1000) / 1000;
}

function outstandingOf(entry) {
  if (!entry || entry.cancelled || !isUnpaid(entry)) return 0;
  return round2(Math.max(0, parseMoney(entry.amount) - parseMoney(entry.amountPaid)));
}

function canTakePayment(entry) {
  if (!entry || entry.cancelled) return false;
  if (entry.sourceType === "sale") return can("sales", "edit") || can("finance", "edit");
  if (entry.sourceType === "purchase") return can("purchases", "edit") || can("finance", "edit");
  return can("finance", "edit");
}

function itemHasTrades(itemId) {
  return [...state.sales, ...state.purchases].some((trade) => trade.lines.some((line) => line.itemId === itemId));
}

function findTrade(kind, id) {
  return state[TRADE[kind].list].find((trade) => trade.id === id);
}

function tradeBalance(trade) {
  if (trade.status === "Cancelled") return 0;
  return round2(Math.max(0, parseMoney(trade.total) - parseMoney(trade.amountPaid)));
}

function tradeStatus(trade) {
  if (trade.status === "Cancelled") return { label: "Cancelled", tone: "red" };
  if (tradeBalance(trade) <= 0) return { label: "Paid", tone: "green" };
  const overdue = Boolean(trade.dueDate && trade.dueDate < todayIso());
  if (parseMoney(trade.amountPaid) > 0) return { label: overdue ? "Part paid · overdue" : "Part paid", tone: overdue ? "red" : "amber" };
  return { label: overdue ? "Unpaid · overdue" : "Unpaid", tone: overdue ? "red" : "amber" };
}

function nextItemCode() {
  return generateAutoId("inventory", {}, autoIdContext("inventory"));
}

function nextTradeNumber(kind) {
  state.counters[kind] = (Number(state.counters[kind]) || 0) + 1;
  return `${kind === "sale" ? "S" : "P"}-${String(state.counters[kind]).padStart(4, "0")}`;
}

function recordStockMovement(item, change, { type, reference = "", referenceId = "", referenceKind = "", party = "", unitAmount = "", reason = "", date = todayIso(), extra = {} } = {}) {
  const balanceAfter = round3(parseMoney(item.quantity) + change);
  item.quantity = balanceAfter;
  state.stockMovements.unshift({
    id: makeId("move"),
    date,
    createdAt: new Date().toISOString(),
    itemId: item.id,
    itemName: item.name,
    sku: item.sku || "",
    type,
    change: round3(change),
    balanceAfter,
    reference,
    referenceId,
    referenceKind,
    party,
    unitAmount,
    currency: recordCurrency(item),
    reason,
    by: currentUser()?.name || "",
    ...extra,
  });
}

function logOpeningStock(item, type, reason = "", date = "") {
  const quantity = parseMoney(item.quantity);
  if (!quantity) return;
  state.stockMovements.unshift({
    id: makeId("move"),
    date: date || String(item.createdAt || item.importedAt || new Date().toISOString()).slice(0, 10),
    createdAt: new Date().toISOString(),
    itemId: item.id,
    itemName: item.name,
    sku: item.sku || "",
    type,
    change: quantity,
    balanceAfter: quantity,
    reference: "",
    referenceId: "",
    referenceKind: "",
    party: "",
    unitAmount: item.unitCost ?? "",
    currency: recordCurrency(item),
    reason,
    by: currentUser()?.name || "",
  });
}

function movementBadge(type) {
  const tones = {
    Sale: "blue",
    Purchase: "green",
    Adjustment: "amber",
    "Sale cancelled": "red",
    "Purchase cancelled": "red",
    "Import update": "violet",
  };
  return badge(type, tones[type] || "violet");
}

/* ----- New sale / new purchase form ----- */

function tradeItemPrice(kind, item) {
  const value = item?.[TRADE[kind].priceField];
  return isBlank(value) ? "" : parseMoney(value);
}

function newTradeLine(kind, itemId = "") {
  const item = itemId ? state.inventory.find((entry) => entry.id === itemId) : null;
  return {
    key: makeId("line"),
    itemId,
    custom: false,
    isNew: false,
    newName: "",
    newSku: "",
    newCategory: "",
    newSellPrice: "",
    description: "",
    quantity: 1,
    unitPrice: tradeItemPrice(kind, item),
    priceTouched: false,
  };
}

function openTradeForm(kind, { itemId = "" } = {}) {
  const config = TRADE[kind];
  if (!can(config.area, "add")) {
    showToast(`Your role cannot record ${config.area}`);
    return;
  }
  const item = itemId ? state.inventory.find((entry) => entry.id === itemId) : null;
  const defaultDue = new Date();
  defaultDue.setDate(defaultDue.getDate() + 14);
  editing = {
    kind,
    draft: {
      date: todayIso(),
      contactId: "",
      partyName: "",
      partyPhone: "",
      saveContact: true,
      currency: item ? recordCurrency(item) : orgCurrency(),
      lines: [newTradeLine(kind, itemId)],
      discount: "",
      payment: "paid",
      amountPaid: "",
      method: "Cash",
      dueDate: isoFromParts(defaultDue.getFullYear(), defaultDue.getMonth() + 1, defaultDue.getDate()),
      reference: "",
      notes: "",
    },
  };
  els.modalKicker.textContent = viewTitle(config.area);
  els.modalTitle.textContent = kind === "sale" ? "New sale" : "New purchase";
  setModalWide(true);
  renderTradeForm();
  els.modalBackdrop.hidden = false;
}

function setModalWide(wide) {
  els.recordForm.closest(".modal")?.classList.toggle("wide", wide);
}

function tradeTotals(draft) {
  const lines = draft.lines.map((line) => ({
    ...line,
    total: round2(parseMoney(line.quantity) * parseMoney(line.unitPrice)),
  }));
  const subtotal = round2(lines.reduce((sum, line) => sum + line.total, 0));
  const discount = round2(Math.min(Math.max(parseMoney(draft.discount), 0), subtotal));
  const total = round2(subtotal - discount);
  const paid = draft.payment === "paid" ? total : draft.payment === "unpaid" ? 0 : round2(Math.max(0, parseMoney(draft.amountPaid)));
  return { lines, subtotal, discount, total, paid, balance: round2(total - paid) };
}

function tradeContactOptions(kind, selected) {
  const config = TRADE[kind];
  const primary = state.contacts.filter((contact) => contact.type === config.partyType);
  const others = state.contacts.filter((contact) => contact.type !== config.partyType);
  const option = (contact) =>
    `<option value="${escapeHtml(contact.id)}" ${contact.id === selected ? "selected" : ""}>${escapeHtml(contact.name)}${
      contact.phone ? ` · ${escapeHtml(contact.phone)}` : ""
    }</option>`;
  return `
    ${
      kind === "sale"
        ? `<option value="" ${selected === "" ? "selected" : ""}>Walk-in customer (not saved)</option>`
        : `<option value="" ${selected === "" ? "selected" : ""}>Choose a supplier…</option>`
    }
    <option value="__new__" ${selected === "__new__" ? "selected" : ""}>+ New ${config.party.toLowerCase()}</option>
    ${primary.length ? `<optgroup label="${config.partyType}s">${primary.map(option).join("")}</optgroup>` : ""}
    ${others.length ? `<optgroup label="Other contacts">${others.map(option).join("")}</optgroup>` : ""}
  `;
}

function tradeLineHint(kind, line) {
  if (line.custom) return kind === "sale" ? "Not taken from stock" : "Not added to stock";
  if (line.isNew) return `Will be added to ${viewTitle("inventory")} with this quantity and cost`;
  const item = state.inventory.find((entry) => entry.id === line.itemId);
  if (!item) return "";
  const currency = recordCurrency(item);
  const stock = parseMoney(item.quantity);
  const parts = [];
  if (kind === "sale") {
    const list = tradeItemPrice(kind, item);
    if (list !== "") parts.push(`List price ${formatCurrency(list, currency)}`);
    parts.push(`${formatNumber(stock)} in stock`);
    if (list !== "" && !isBlank(line.unitPrice) && parseMoney(line.unitPrice) !== list) {
      const diff = parseMoney(line.unitPrice) - list;
      parts.push(`${diff < 0 ? "below" : "above"} list by ${formatCurrency(Math.abs(diff), currency)}`);
    }
  } else {
    if (!isBlank(item.unitCost)) parts.push(`Current cost ${formatCurrency(item.unitCost, currency)}`);
    parts.push(`${formatNumber(stock)} in stock`);
  }
  return parts.join(" · ");
}

function tradeLineWarning(kind, line, draft) {
  if (kind !== "sale" || line.custom || !line.itemId) return "";
  const item = state.inventory.find((entry) => entry.id === line.itemId);
  if (!item) return "";
  const wanted = draft.lines
    .filter((entry) => entry.itemId === line.itemId && !entry.custom)
    .reduce((sum, entry) => sum + parseMoney(entry.quantity), 0);
  return wanted > parseMoney(item.quantity) ? `Only ${formatNumber(item.quantity)} in stock` : "";
}

function renderTradeForm() {
  const { kind, draft } = editing;
  const config = TRADE[kind];
  const totals = tradeTotals(draft);
  const money = (value) => formatCurrency(value, draft.currency);
  const items = state.inventory
    .filter((item) => recordCurrency(item) === draft.currency)
    .sort((a, b) => String(a.name).localeCompare(String(b.name)));
  const hiddenItems = state.inventory.length - items.length;
  const showPartyInputs = draft.contactId === "__new__" || (kind === "sale" && draft.contactId === "");

  const lineRows = totals.lines
    .map((line, index) => {
      const warning = tradeLineWarning(kind, line, draft);
      return `
      <div class="trade-line ${warning ? "has-warning" : ""}" data-line="${line.key}">
        <div class="trade-item">
          <label class="sr-only" for="line-item-${line.key}">Item ${index + 1}</label>
          <select id="line-item-${line.key}" name="line-item-${line.key}" data-trade-refresh>
            <option value="">Choose ${escapeHtml(recordNoun("inventory").toLowerCase())}…</option>
            ${
              kind === "purchase" && can("inventory", "add")
                ? `<option value="__new__" ${line.isNew ? "selected" : ""}>+ New ${escapeHtml(recordNoun("inventory").toLowerCase())} (add to ${escapeHtml(
                    viewTitle("inventory").toLowerCase(),
                  )})</option>`
                : ""
            }
            ${items
              .map(
                (item) =>
                  `<option value="${escapeHtml(item.id)}" ${item.id === line.itemId && !line.custom ? "selected" : ""}>${escapeHtml(item.name)}${
                    item.sku ? ` (${escapeHtml(item.sku)})` : ""
                  } · ${formatNumber(item.quantity)} in stock</option>`,
              )
              .join("")}
            <option value="__custom__" ${line.custom ? "selected" : ""}>${
              kind === "purchase" ? "Other expense or service (not stock)" : "Other item or service (not in stock)"
            }</option>
          </select>
          ${
            line.custom
              ? `<input name="line-desc-${line.key}" type="text" value="${escapeHtml(line.description)}" placeholder="Describe the item or service" />`
              : ""
          }
          ${
            line.isNew
              ? `<div class="trade-new-item">
                  <input name="line-new-name-${line.key}" type="text" value="${escapeHtml(line.newName)}" placeholder="${escapeHtml(fieldLabel("inventory", "name"))} *" />
                  <input name="line-new-sku-${line.key}" type="text" value="${escapeHtml(line.newSku)}" placeholder="${escapeHtml(fieldLabel("inventory", "sku"))} (auto if blank)" />
                  <input name="line-new-category-${line.key}" type="text" value="${escapeHtml(line.newCategory)}" placeholder="${escapeHtml(fieldLabel("inventory", "category"))}" list="tradeCategoryList" />
                  <input name="line-new-sell-${line.key}" type="number" min="0" step="0.01" value="${escapeHtml(line.newSellPrice)}" placeholder="${escapeHtml(fieldLabel("inventory", "sellPrice"))}" />
                </div>`
              : ""
          }
          <small data-line-hint="${line.key}">${escapeHtml(tradeLineHint(kind, line))}</small>
          <small class="line-warning" data-line-warning="${line.key}" ${warning ? "" : "hidden"}>${escapeHtml(warning)}</small>
        </div>
        <label class="trade-cell">
          <span>Qty</span>
          <input name="line-qty-${line.key}" type="number" min="0" step="any" value="${escapeHtml(line.quantity)}" />
        </label>
        <label class="trade-cell">
          <span>${config.priceLabel} each</span>
          <input name="line-price-${line.key}" type="number" min="0" step="0.01" value="${escapeHtml(line.unitPrice)}" />
        </label>
        <div class="trade-cell trade-line-total">
          <span>Total</span>
          <strong data-line-total="${line.key}">${escapeHtml(money(line.total))}</strong>
        </div>
        <button type="button" class="icon-button" data-trade-remove-line="${line.key}" aria-label="Remove line" ${
          draft.lines.length === 1 ? "disabled" : ""
        }>
          <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 6l12 12M18 6 6 18" /></svg>
        </button>
      </div>
    `;
    })
    .join("");

  const paymentChoice = (value, label) => `
    <label class="choice-pill">
      <input type="radio" name="payment" value="${value}" data-trade-refresh ${draft.payment === value ? "checked" : ""} />
      <span>${label}</span>
    </label>
  `;

  const tradeHelp = TRADE_SECTION_HELP[kind];
  els.recordForm.innerHTML = `
    <div class="form-error" id="formError" hidden></div>
    ${formHelpBarHtml(`Four steps: who, what, payment, and details. The ${config.noun.toLowerCase()} number is created automatically.`)}
    <fieldset class="form-section">
      <legend>${config.party}</legend>
      ${sectionHelpHtml(tradeHelp.party)}
      <div class="form-grid">
        <div class="field full">
          <label for="tradeContact">${kind === "sale" ? "Sold to" : "Bought from"}</label>
          <select id="tradeContact" name="contactId" data-trade-refresh>${tradeContactOptions(kind, draft.contactId)}</select>
          ${
            kind === "sale" && draft.contactId === ""
              ? `<small>For a walk-in you can leave the name blank; the sale is recorded as "Walk-in customer". A name is needed if they will pay later.</small>`
              : ""
          }
        </div>
        ${
          showPartyInputs
            ? `
          <div class="field">
            <label for="tradePartyName">${config.party} name${draft.contactId === "__new__" ? " *" : ""}</label>
            <input id="tradePartyName" name="partyName" type="text" value="${escapeHtml(draft.partyName)}" placeholder="${
              draft.contactId === "__new__" ? "Person or company name" : "Optional"
            }" />
          </div>
          <div class="field">
            <label for="tradePartyPhone">Phone</label>
            <input id="tradePartyPhone" name="partyPhone" type="text" value="${escapeHtml(draft.partyPhone)}" placeholder="Optional" />
          </div>
          ${
            draft.contactId === "__new__" && can("contacts", "add")
              ? `<label class="check-control full"><input type="checkbox" name="saveContact" ${draft.saveContact ? "checked" : ""} /><span>Save to ${escapeHtml(
                  viewTitle("contacts"),
                )} for next time</span></label>`
              : ""
          }
        `
            : ""
        }
      </div>
    </fieldset>

    <fieldset class="form-section">
      <legend>${kind === "sale" ? "What was sold" : "What was bought"}</legend>
      ${sectionHelpHtml(tradeHelp.lines)}
      ${
        hiddenItems
          ? `<p class="muted-note trade-note">${hiddenItems} ${escapeHtml(viewTitle("inventory").toLowerCase())} priced in other currencies are hidden. Change the currency below to use them.</p>`
          : ""
      }
      <div class="trade-lines">${lineRows}</div>
      <datalist id="tradeCategoryList">${[...new Set(state.inventory.map((item) => item.category).filter(Boolean))]
        .map((category) => `<option value="${escapeHtml(category)}"></option>`)
        .join("")}</datalist>
      <button type="button" class="button ghost small" data-trade-add-line>
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 5v14M5 12h14" /></svg>
        Add another line
      </button>
      <div class="trade-totals">
        <div><span>Subtotal</span><strong id="tradeSubtotal">${escapeHtml(money(totals.subtotal))}</strong></div>
        <label><span>Discount</span><input name="discount" type="number" min="0" step="0.01" value="${escapeHtml(draft.discount)}" placeholder="0" /></label>
        <div class="grand"><span>Total</span><strong id="tradeTotal">${escapeHtml(money(totals.total))}</strong></div>
      </div>
    </fieldset>

    <fieldset class="form-section">
      <legend>Payment</legend>
      ${sectionHelpHtml(tradeHelp.payment)}
      <div class="form-grid">
        <div class="field">
          <label for="tradeDate">Date</label>
          <input id="tradeDate" name="date" type="date" value="${escapeHtml(draft.date)}" max="${todayIso()}" />
        </div>
        <div class="field">
          <label for="tradeCurrency">Currency</label>
          <select id="tradeCurrency" name="currency" data-trade-refresh>
            ${currencyOptions()
              .map((code) => `<option value="${code}" ${code === draft.currency ? "selected" : ""}>${code}</option>`)
              .join("")}
          </select>
        </div>
        <div class="field full">
          <span class="field-label">${kind === "sale" ? "Has the customer paid?" : "Have you paid the supplier?"}</span>
          <div class="choice-pills">
            ${paymentChoice("paid", "Paid in full")}
            ${paymentChoice("partial", "Part paid")}
            ${paymentChoice("unpaid", kind === "sale" ? "Not paid yet (on credit)" : "Not paid yet")}
          </div>
        </div>
        ${
          draft.payment === "partial"
            ? `<div class="field"><label for="tradeAmountPaid">Amount paid now *</label><input id="tradeAmountPaid" name="amountPaid" type="number" min="0" step="0.01" value="${escapeHtml(
                draft.amountPaid,
              )}" /></div>`
            : ""
        }
        ${
          draft.payment !== "unpaid"
            ? `<div class="field"><label for="tradeMethod">Paid by</label><select id="tradeMethod" name="method">${PAYMENT_METHODS.map(
                (method) => `<option ${method === draft.method ? "selected" : ""}>${method}</option>`,
              ).join("")}</select></div>`
            : ""
        }
        ${
          draft.payment !== "paid"
            ? `<div class="field"><label for="tradeDueDate">Balance due by</label><input id="tradeDueDate" name="dueDate" type="date" value="${escapeHtml(
                draft.dueDate,
              )}" /></div>`
            : ""
        }
        <div class="field full"><div class="trade-balance" id="tradeBalanceNote">${tradeBalanceNote(kind, totals, draft)}</div></div>
      </div>
    </fieldset>

    <fieldset class="form-section">
      <legend>Details</legend>
      ${sectionHelpHtml(tradeHelp.details)}
      <div class="form-grid">
        <div class="field">
          <label for="tradeReference">${kind === "purchase" ? "Supplier bill / invoice no." : "Reference (optional)"}</label>
          <input id="tradeReference" name="reference" type="text" value="${escapeHtml(draft.reference)}" />
        </div>
        <div class="field full">
          <label for="tradeNotes">Notes</label>
          <textarea id="tradeNotes" name="notes">${escapeHtml(draft.notes)}</textarea>
        </div>
      </div>
    </fieldset>

    <div class="form-actions">
      <button class="button ghost" id="cancelFormBtn" type="button">Cancel</button>
      <button class="button primary" type="submit" id="tradeSubmitBtn">Save ${config.noun.toLowerCase()} · ${escapeHtml(money(totals.total))}</button>
    </div>
  `;
  applyFieldHelpVisibility();
}

function tradeBalanceNote(kind, totals, draft) {
  const money = (value) => formatCurrency(value, draft.currency);
  if (totals.total <= 0) return `<span class="muted-note">Add items to see the total.</span>`;
  if (totals.balance <= 0) {
    return kind === "sale"
      ? `<span class="money-in">Fully paid: ${escapeHtml(money(totals.total))} received.</span>`
      : `<span>Fully paid: ${escapeHtml(money(totals.total))} paid to the supplier.</span>`;
  }
  return kind === "sale"
    ? `<span class="negative-text">${escapeHtml(money(totals.balance))} will be added to Dues as owed to you.</span>`
    : `<span class="negative-text">${escapeHtml(money(totals.balance))} will be added to Dues as money you owe.</span>`;
}

function captureTradeForm() {
  if (!editing || !TRADE[editing.kind]) return;
  const { kind, draft } = editing;
  const form = els.recordForm;
  const value = (name) => form.elements[name]?.value ?? undefined;
  const assign = (key, name = key) => {
    const next = value(name);
    if (next !== undefined) draft[key] = next;
  };
  assign("contactId");
  assign("partyName");
  assign("partyPhone");
  if (form.elements.saveContact) draft.saveContact = form.elements.saveContact.checked;
  assign("date");
  const previousCurrency = draft.currency;
  assign("currency");
  assign("discount");
  const payment = form.querySelector('input[name="payment"]:checked');
  if (payment) draft.payment = payment.value;
  assign("amountPaid");
  assign("method");
  assign("dueDate");
  assign("reference");
  assign("notes");

  draft.lines.forEach((line) => {
    // Read the typed price before handling an item change, so choosing an item
    // can fill in its price unless the user already typed one.
    const price = value(`line-price-${line.key}`);
    if (price !== undefined && String(price) !== String(line.unitPrice)) {
      line.unitPrice = price;
      line.priceTouched = true;
    }
    const selected = value(`line-item-${line.key}`);
    if (selected !== undefined) {
      const custom = selected === "__custom__";
      const isNew = selected === "__new__";
      const itemId = custom || isNew ? "" : selected;
      if (custom !== line.custom || isNew !== Boolean(line.isNew) || itemId !== line.itemId) {
        line.custom = custom;
        line.isNew = isNew;
        line.itemId = itemId;
        if (!line.priceTouched || isBlank(line.unitPrice)) {
          const item = state.inventory.find((entry) => entry.id === itemId);
          line.unitPrice = custom || isNew ? "" : tradeItemPrice(kind, item);
          line.priceTouched = false;
        }
      }
    }
    const description = value(`line-desc-${line.key}`);
    if (description !== undefined) line.description = description;
    [
      ["newName", "name"],
      ["newSku", "sku"],
      ["newCategory", "category"],
      ["newSellPrice", "sell"],
    ].forEach(([key, suffix]) => {
      const typed = value(`line-new-${suffix}-${line.key}`);
      if (typed !== undefined) line[key] = typed;
    });
    const quantity = value(`line-qty-${line.key}`);
    if (quantity !== undefined) line.quantity = quantity;
  });

  if (draft.currency !== previousCurrency) {
    // Items priced in another currency can't be used in this transaction.
    draft.lines.forEach((line) => {
      const item = state.inventory.find((entry) => entry.id === line.itemId);
      if (item && recordCurrency(item) !== draft.currency) {
        line.itemId = "";
        line.unitPrice = "";
        line.priceTouched = false;
      }
    });
  }
}

function updateTradeSummary() {
  const { kind, draft } = editing;
  const totals = tradeTotals(draft);
  const money = (value) => formatCurrency(value, draft.currency);
  totals.lines.forEach((line) => {
    const total = els.recordForm.querySelector(`[data-line-total="${line.key}"]`);
    if (total) total.textContent = money(line.total);
    const hint = els.recordForm.querySelector(`[data-line-hint="${line.key}"]`);
    if (hint) hint.textContent = tradeLineHint(kind, line);
    const warning = tradeLineWarning(kind, line, draft);
    const warningNode = els.recordForm.querySelector(`[data-line-warning="${line.key}"]`);
    if (warningNode) {
      warningNode.textContent = warning;
      warningNode.hidden = !warning;
      warningNode.closest(".trade-line")?.classList.toggle("has-warning", Boolean(warning));
    }
  });
  byId("tradeSubtotal").textContent = money(totals.subtotal);
  byId("tradeTotal").textContent = money(totals.total);
  byId("tradeBalanceNote").innerHTML = tradeBalanceNote(kind, totals, draft);
  byId("tradeSubmitBtn").textContent = `Save ${TRADE[kind].noun.toLowerCase()} · ${money(totals.total)}`;
}

// Checks a draft and returns either { error } or everything needed to save it.
function prepareTrade(kind, draft) {
  const config = TRADE[kind];
  const currency = draft.currency;
  if (!isValidCurrency(currency)) return { error: "Choose a valid currency." };
  if (!draft.date) return { error: "Enter the date." };
  if (draft.date > todayIso()) return { error: "The date can't be in the future." };

  const lines = [];
  const quantityByItem = {};
  for (const [index, line] of draft.lines.entries()) {
    const empty = !line.itemId && !line.custom && !line.isNew && isBlank(line.description);
    if (empty && (isBlank(line.quantity) || parseMoney(line.quantity) === 1) && isBlank(line.unitPrice)) continue;
    const label = `Line ${index + 1}`;
    const quantity = parseMoney(line.quantity);
    if (!line.itemId && !line.custom && !line.isNew) return { error: `${label}: choose what was ${kind === "sale" ? "sold" : "bought"}.` };
    if (quantity <= 0) return { error: `${label}: enter a quantity above zero.` };
    if (isBlank(line.unitPrice)) return { error: `${label}: enter the ${config.priceLabel.toLowerCase()}.` };
    const unitPrice = parseMoney(line.unitPrice);
    if (unitPrice < 0) return { error: `${label}: the ${config.priceLabel.toLowerCase()} can't be negative.` };
    if (line.custom) {
      if (isBlank(line.description)) return { error: `${label}: describe the item or service.` };
      lines.push({ itemId: "", custom: true, name: line.description.trim(), sku: "", quantity, unitPrice, listPrice: "", unitCost: "", lineTotal: round2(quantity * unitPrice) });
      continue;
    }
    if (line.isNew) {
      if (kind !== "purchase" || !can("inventory", "add")) return { error: `${label}: new items can only be added from a purchase.` };
      const name = String(line.newName || "").trim();
      if (!name) return { error: `${label}: enter the name of the new ${recordNoun("inventory").toLowerCase()}.` };
      const existing = state.inventory.find((item) => item.name.trim().toLowerCase() === name.toLowerCase() && recordCurrency(item) === currency);
      if (existing) return { error: `${label}: ${existing.name} is already in ${viewTitle("inventory")}. Choose it from the list instead.` };
      const sku = String(line.newSku || "").trim();
      if (sku && state.inventory.some((item) => String(item.sku || "").trim().toLowerCase() === sku.toLowerCase())) {
        return { error: `${label}: the code ${sku} is already used by another ${recordNoun("inventory").toLowerCase()}.` };
      }
      if (!isBlank(line.newSellPrice) && parseMoney(line.newSellPrice) < 0) return { error: `${label}: the sell price can't be negative.` };
      if (lines.some((entry) => entry.newItem && entry.name.toLowerCase() === name.toLowerCase())) {
        return { error: `${label}: ${name} is entered twice. Put the full quantity on one line.` };
      }
      lines.push({
        itemId: "",
        custom: false,
        name,
        sku,
        quantity,
        unitPrice,
        listPrice: "",
        unitCost: unitPrice,
        lineTotal: round2(quantity * unitPrice),
        newItem: {
          sku,
          category: String(line.newCategory || "").trim(),
          sellPrice: isBlank(line.newSellPrice) ? "" : round2(parseMoney(line.newSellPrice)),
        },
      });
      continue;
    }
    const item = state.inventory.find((entry) => entry.id === line.itemId);
    if (!item) return { error: `${label}: that item no longer exists.` };
    if (recordCurrency(item) !== currency) {
      return { error: `${label}: ${item.name} is priced in ${recordCurrency(item)}, but this ${config.noun.toLowerCase()} is in ${currency}.` };
    }
    quantityByItem[item.id] = (quantityByItem[item.id] || 0) + quantity;
    lines.push({
      itemId: item.id,
      custom: false,
      name: item.name,
      sku: item.sku || "",
      quantity,
      unitPrice,
      listPrice: tradeItemPrice(kind, item),
      unitCost: kind === "sale" ? (isBlank(item.unitCost) ? "" : parseMoney(item.unitCost)) : unitPrice,
      lineTotal: round2(quantity * unitPrice),
    });
  }
  if (!lines.length) return { error: `Add at least one item to the ${config.noun.toLowerCase()}.` };
  if (kind === "sale") {
    for (const [itemId, wanted] of Object.entries(quantityByItem)) {
      const item = state.inventory.find((entry) => entry.id === itemId);
      if (wanted > parseMoney(item.quantity) + 0.0005) {
        return {
          error: `Only ${formatNumber(item.quantity)} of ${item.name} in stock, but this sale has ${formatNumber(wanted)}. Record a purchase or adjust stock first.`,
        };
      }
    }
  }

  const subtotal = round2(lines.reduce((sum, line) => sum + line.lineTotal, 0));
  const discount = round2(parseMoney(draft.discount));
  if (discount < 0) return { error: "The discount can't be negative." };
  if (discount > subtotal) return { error: "The discount can't be more than the subtotal." };
  const total = round2(subtotal - discount);
  let paid = total;
  if (draft.payment === "unpaid") paid = 0;
  if (draft.payment === "partial") {
    paid = round2(parseMoney(draft.amountPaid));
    if (!(paid > 0 && paid < total)) {
      return { error: `For a part payment, enter an amount between 0 and ${formatCurrency(total, currency)}.` };
    }
  }
  const balance = round2(total - paid);
  if (balance > 0 && total <= 0) return { error: "An unpaid transaction needs a total above zero." };

  let contact = draft.contactId && draft.contactId !== "__new__" ? state.contacts.find((entry) => entry.id === draft.contactId) : null;
  if (draft.contactId && draft.contactId !== "__new__" && !contact) return { error: "That contact no longer exists." };
  let partyName = contact ? contact.name : String(draft.partyName || "").trim();
  const partyPhone = contact ? contact.phone || "" : String(draft.partyPhone || "").trim();
  if (kind === "purchase" && !partyName) return { error: "Choose or enter the supplier this stock came from." };
  if (draft.contactId === "__new__" && !partyName) return { error: `Enter the ${config.party.toLowerCase()}'s name.` };
  if (balance > 0 && !partyName) {
    return { error: "Enter the customer's name so you know who owes this money, or mark the sale as paid in full." };
  }
  if (!contact && partyName) {
    contact = state.contacts.find((entry) => entry.name.trim().toLowerCase() === partyName.toLowerCase()) || null;
  }
  return {
    kind,
    lines,
    subtotal,
    discount,
    total,
    paid,
    balance,
    currency,
    contact,
    partyName,
    partyPhone,
    createContact: !contact && draft.contactId === "__new__" && draft.saveContact && can("contacts", "add"),
  };
}

function createTrade(kind, draft) {
  const prepared = prepareTrade(kind, draft);
  if (prepared.error) return prepared;
  const config = TRADE[kind];
  const now = new Date().toISOString();
  const user = currentUser();
  let contact = prepared.contact;
  if (prepared.createContact) {
    contact = {
      id: makeId("contacts"),
      name: prepared.partyName,
      type: config.partyType,
      phone: prepared.partyPhone,
      email: "",
      balance: 0,
      currency: prepared.currency,
      status: "Active",
      sourceName: `Added while recording a ${config.noun.toLowerCase()}`,
      createdAt: now,
      infoDismissed: true,
    };
    state.contacts.unshift(contact);
  }

  const id = makeId(kind);
  const number = nextTradeNumber(kind);
  const partyName = prepared.partyName || (kind === "sale" ? "Walk-in customer" : "");
  const trade = {
    id,
    number,
    date: draft.date,
    contactId: contact?.id || "",
    partyName,
    partyPhone: prepared.partyPhone,
    walkIn: !prepared.partyName,
    currency: prepared.currency,
    lines: prepared.lines,
    subtotal: prepared.subtotal,
    discount: prepared.discount,
    total: prepared.total,
    amountPaid: prepared.paid,
    paymentMethod: prepared.paid > 0 ? draft.method || "Cash" : "",
    dueDate: prepared.balance > 0 ? draft.dueDate || "" : "",
    reference: String(draft.reference || "").trim(),
    notes: String(draft.notes || "").trim(),
    status: "Completed",
    payments:
      prepared.paid > 0
        ? [
            {
              id: makeId("payment"),
              date: draft.date,
              amount: prepared.paid,
              method: draft.method || "Cash",
              note: `Paid when the ${config.noun.toLowerCase()} was recorded`,
              by: user?.name || "",
              at: now,
              receiptNumber: nextReceiptNumber(kind === "sale" ? "in" : "out"),
            },
          ]
        : [],
    createdAt: now,
    createdBy: user?.name || "",
  };

  prepared.lines.forEach((line) => {
    if (!line.newItem) return;
    // A new item bought for the first time is added to the item list with no stock;
    // the purchase below then brings its stock in, so its history starts with this purchase.
    const item = {
      id: makeId("inventory"),
      name: line.name,
      sku: line.newItem.sku || nextItemCode(),
      category: line.newItem.category || "General",
      quantity: 0,
      reorderLevel: "",
      location: "",
      unitCost: "",
      sellPrice: line.newItem.sellPrice,
      currency: prepared.currency,
      notes: "",
      sourceName: `Added from purchase ${number}`,
      createdAt: now,
    };
    state.inventory.unshift(item);
    line.itemId = item.id;
    line.sku = item.sku;
    line.addedToStock = true;
    delete line.newItem;
  });

  prepared.lines.forEach((line) => {
    if (line.custom) return;
    const item = state.inventory.find((entry) => entry.id === line.itemId);
    const extra = line.addedToStock ? { newItem: true } : {};
    if (kind === "purchase") {
      // The item's cost becomes the weighted average of stock on hand and this purchase.
      const onHand = Math.max(0, parseMoney(item.quantity));
      const previousCost = item.unitCost;
      const newCost =
        isBlank(previousCost) || onHand <= 0
          ? line.unitPrice
          : round2((onHand * parseMoney(previousCost) + line.quantity * line.unitPrice) / (onHand + line.quantity));
      item.unitCost = newCost;
      extra.previousCost = previousCost;
      extra.newCost = newCost;
    }
    recordStockMovement(item, config.sign * line.quantity, {
      type: config.movement,
      reference: number,
      referenceId: id,
      referenceKind: kind,
      party: partyName,
      unitAmount: line.unitPrice,
      date: draft.date,
      extra,
    });
  });

  if (prepared.total > 0) {
    const entry = {
      id: makeId("finance"),
      date: draft.date,
      type: config.financeType,
      category: config.category,
      description: `${config.noun} ${number}${partyName ? ` - ${partyName}` : ""}`,
      amount: prepared.total,
      amountPaid: prepared.paid,
      payments: trade.payments.map((payment) => ({ ...payment })),
      currency: prepared.currency,
      status: prepared.balance > 0 ? "Pending" : "Paid",
      dueDate: trade.dueDate,
      paidOn: prepared.balance > 0 ? "" : draft.date,
      party: partyName,
      sourceType: kind,
      sourceId: id,
      reference: number,
      notes: trade.notes,
      sourceName: viewTitle(config.area),
      createdAt: now,
      infoDismissed: true,
    };
    state.finance.unshift(entry);
    trade.financeId = entry.id;
  }

  state[config.list].unshift(trade);
  const added = prepared.lines.filter((line) => line.addedToStock);
  if (added.length) {
    addActivity(`New ${viewTitle("inventory").toLowerCase()} added from purchase ${number}: ${added.map((line) => line.name).join(", ")}`, "inventory");
  }
  addActivity(
    `${config.noun} ${number}: ${partyName || "walk-in"} · ${formatCurrency(prepared.total, prepared.currency)}${
      prepared.balance > 0 ? ` (${formatCurrency(prepared.balance, prepared.currency)} unpaid)` : ""
    }`,
    config.area,
  );
  return { trade };
}

function submitTrade() {
  captureTradeForm();
  const { kind } = editing;
  const result = createTrade(kind, editing.draft);
  if (result.error) {
    showFormError(result.error);
    return;
  }
  saveState();
  closeModal();
  rememberTab(TRADE[kind].area, "all");
  if (currentView !== TRADE[kind].area && canOpenView(TRADE[kind].area)) setView(TRADE[kind].area);
  else render();
  const added = result.trade.lines.filter((line) => line.addedToStock).length;
  showToast(
    `${TRADE[kind].noun} ${result.trade.number} saved${
      added ? ` · ${added} new ${added === 1 ? recordNoun("inventory").toLowerCase() : viewTitle("inventory").toLowerCase()} added` : ""
    }`,
  );
}

/* ----- Payments ----- */

function openPaymentModal(financeId) {
  const entry = state.finance.find((record) => record.id === financeId);
  if (!entry || !canTakePayment(entry)) {
    showToast("Your role cannot record payments for this entry");
    return;
  }
  const due = outstandingOf(entry);
  if (!due) {
    showToast("This is already fully paid");
    return;
  }
  const currency = recordCurrency(entry);
  const incoming = entry.type === "Income";
  editing = { kind: "payment", financeId };
  setModalWide(false);
  els.modalKicker.textContent = entry.reference || "Finance";
  els.modalTitle.textContent = incoming ? "Receive payment" : "Record a payment you made";
  els.recordForm.innerHTML = `
    <div class="form-error" id="formError" hidden></div>
    <div class="detail-grid">
      <div><span>${incoming ? "From" : "To"}</span><strong>${escapeHtml(entry.party || "—")}</strong></div>
      <div><span>For</span><strong>${escapeHtml(entry.description || entry.category)}</strong></div>
      <div><span>Total</span><strong>${escapeHtml(formatCurrency(entry.amount, currency))}</strong></div>
      <div><span>Already paid</span><strong>${escapeHtml(formatCurrency(entry.amountPaid || 0, currency))}</strong></div>
      <div><span>Balance due</span><strong class="negative-text">${escapeHtml(formatCurrency(due, currency))}</strong></div>
      ${entry.dueDate ? `<div><span>Due date</span><strong>${escapeHtml(formatDate(entry.dueDate))}</strong></div>` : ""}
    </div>
    <fieldset class="form-section">
      <legend>Payment</legend>
      <div class="form-grid">
        <div class="field">
          <label for="paymentAmount">Amount (${currency}) *</label>
          <input id="paymentAmount" name="amount" type="number" min="0" step="0.01" max="${due}" value="${due}" required />
          <small>Enter less for a part payment.</small>
        </div>
        <div class="field">
          <label for="paymentDate">Date</label>
          <input id="paymentDate" name="date" type="date" value="${todayIso()}" max="${todayIso()}" />
        </div>
        <div class="field">
          <label for="paymentMethod">Method</label>
          <select id="paymentMethod" name="method">${PAYMENT_METHODS.map((method) => `<option>${method}</option>`).join("")}</select>
        </div>
        <div class="field">
          <label for="paymentNote">Note</label>
          <input id="paymentNote" name="note" type="text" placeholder="Receipt no., cheque no., …" />
        </div>
      </div>
    </fieldset>
    <div class="form-actions">
      <button class="button ghost" id="cancelFormBtn" type="button">Cancel</button>
      <button class="button primary" type="submit">Save payment</button>
    </div>
  `;
  els.modalBackdrop.hidden = false;
  byId("paymentAmount").focus();
}

function nextReceiptNumber(direction) {
  return direction === "in" ? nextDocumentNumber("receipt", "RCPT") : nextDocumentNumber("paymentOut", "PAY");
}

function applyPayment(entry, amount, { date, method, note = "", receiptNumber = "", groupId = "" }) {
  const payment = {
    id: makeId("payment"),
    date,
    amount,
    method,
    note,
    by: currentUser()?.name || "",
    at: new Date().toISOString(),
    receiptNumber: receiptNumber || nextReceiptNumber(entry.type === "Income" ? "in" : "out"),
    ...(groupId ? { groupId } : {}),
  };
  entry.amountPaid = round2(parseMoney(entry.amountPaid) + amount);
  entry.payments = [...(entry.payments || []), payment];
  if (entry.amountPaid >= parseMoney(entry.amount) - 0.005) {
    entry.status = "Paid";
    entry.paidOn = date;
  }
  if (entry.sourceType) {
    const trade = findTrade(entry.sourceType, entry.sourceId);
    if (trade) {
      trade.amountPaid = entry.amountPaid;
      trade.payments = [...(trade.payments || []), { ...payment }];
      if (!trade.paymentMethod) trade.paymentMethod = method;
    }
  }
}

function submitPayment() {
  const entry = state.finance.find((record) => record.id === editing.financeId);
  const form = els.recordForm.elements;
  const due = outstandingOf(entry);
  const amount = round2(parseMoney(form.amount.value));
  const date = form.date.value || todayIso();
  if (!entry || !canTakePayment(entry)) {
    showFormError("This payment can no longer be recorded.");
    return;
  }
  if (!(amount > 0) || amount > due + 0.005) {
    showFormError(`Enter an amount above zero and up to ${formatCurrency(due, recordCurrency(entry))}.`);
    return;
  }
  if (date > todayIso()) {
    showFormError("The payment date can't be in the future.");
    return;
  }
  applyPayment(entry, amount, { date, method: form.method.value, note: form.note.value.trim() });
  const receiptNumber = entry.payments[entry.payments.length - 1].receiptNumber;
  addActivity(
    `${entry.type === "Income" ? "Payment received" : "Payment made"}: ${formatCurrency(amount, recordCurrency(entry))} · ${entry.party || entry.description} (${receiptNumber})`,
    entry.sourceType ? TRADE[entry.sourceType].area : "finance",
  );
  saveState();
  closeModal();
  render();
  showToast(`${outstandingOf(entry) ? "Part payment recorded" : "Fully paid"} · ${receiptNumber}`);
  openDocumentFor(`receipt:${receiptNumber}`);
}

/* ----- Cancelling a sale or purchase ----- */

function openCancelTrade(kind, id) {
  const config = TRADE[kind];
  const trade = findTrade(kind, id);
  if (!trade || trade.status === "Cancelled") return;
  if (!can(config.area, "delete")) {
    showToast(`Your role cannot cancel ${config.area}`);
    return;
  }
  const blocked = cancelBlockers(kind, trade);
  if (blocked) {
    showToast(blocked);
    return;
  }
  editing = { kind: "cancel", tradeKind: kind, tradeId: id };
  setModalWide(false);
  els.modalKicker.textContent = `${config.noun} ${trade.number}`;
  els.modalTitle.textContent = `Cancel ${config.noun.toLowerCase()}`;
  const stockLines = trade.lines.filter((line) => !line.custom);
  const reasons =
    kind === "sale"
      ? ["Customer returned the items", "Entered by mistake", "Order cancelled", "Other"]
      : ["Returned to supplier", "Entered by mistake", "Delivery cancelled", "Other"];
  els.recordForm.innerHTML = `
    <div class="form-error" id="formError" hidden></div>
    <div class="form-notice">
      <strong>What cancelling does</strong>
      <span>
        ${
          stockLines.length
            ? `${kind === "sale" ? "Puts back" : "Removes"} stock: ${escapeHtml(
                stockLines.map((line) => `${line.name} ${kind === "sale" ? "+" : "−"}${formatNumber(line.quantity)}`).join(", "),
              )}. `
            : ""
        }Removes the ${escapeHtml(formatCurrency(trade.total, trade.currency))} ${kind === "sale" ? "income" : "expense"} and any balance still due.
        ${
          parseMoney(trade.amountPaid) > 0
            ? `${escapeHtml(formatCurrency(trade.amountPaid, trade.currency))} was already paid. Remember to ${
                kind === "sale" ? "refund the customer" : "collect it back from the supplier"
              }.`
            : ""
        }
        The ${config.noun.toLowerCase()} stays in the records, marked as cancelled.
      </span>
    </div>
    <fieldset class="form-section">
      <legend>Reason</legend>
      <div class="form-grid">
        <div class="field">
          <label for="cancelReason">Reason *</label>
          <select id="cancelReason" name="reason">${reasons.map((reason) => `<option>${reason}</option>`).join("")}</select>
        </div>
        <div class="field full">
          <label for="cancelNote">Details</label>
          <textarea id="cancelNote" name="note" placeholder="Optional"></textarea>
        </div>
      </div>
    </fieldset>
    <div class="form-actions">
      <button class="button ghost" id="cancelFormBtn" type="button">Keep ${config.noun.toLowerCase()}</button>
      <button class="button danger-ghost" type="submit">Cancel ${config.noun.toLowerCase()}</button>
    </div>
  `;
  els.modalBackdrop.hidden = false;
}

function cancelBlockers(kind, trade) {
  if (kind !== "purchase") return "";
  const short = trade.lines
    .filter((line) => !line.custom)
    .map((line) => ({ line, item: state.inventory.find((entry) => entry.id === line.itemId) }))
    .filter(({ line, item }) => item && parseMoney(item.quantity) < line.quantity);
  if (!short.length) return "";
  return `Can't cancel ${trade.number}: some of this stock has already been sold (${short
    .map(({ line, item }) => `${line.name}: ${formatNumber(item.quantity)} left of ${formatNumber(line.quantity)}`)
    .join(", ")}). Use Adjust stock to record a return instead.`;
}

function submitCancel() {
  const { tradeKind: kind, tradeId } = editing;
  const config = TRADE[kind];
  const trade = findTrade(kind, tradeId);
  if (!trade || trade.status === "Cancelled") {
    closeModal();
    return;
  }
  const blocked = cancelBlockers(kind, trade);
  if (blocked) {
    showFormError(blocked);
    return;
  }
  const form = els.recordForm.elements;
  const reason = [form.reason.value, form.note.value.trim()].filter(Boolean).join(": ");
  const now = new Date().toISOString();
  trade.status = "Cancelled";
  trade.cancelledAt = now;
  trade.cancelledBy = currentUser()?.name || "";
  trade.cancelReason = reason;
  trade.lines.forEach((line) => {
    if (line.custom) return;
    const item = state.inventory.find((entry) => entry.id === line.itemId);
    if (!item) return;
    recordStockMovement(item, -config.sign * line.quantity, {
      type: config.reversal,
      reference: trade.number,
      referenceId: trade.id,
      referenceKind: kind,
      party: trade.partyName,
      unitAmount: line.unitPrice,
      reason,
    });
  });
  const entry = state.finance.find((record) => record.id === trade.financeId);
  if (entry) {
    entry.cancelled = true;
    entry.cancelReason = reason;
    entry.cancelledAt = now;
  }
  addActivity(`${config.noun} ${trade.number} cancelled (${reason})`, config.area);
  saveState();
  closeModal();
  render();
  showToast(`${config.noun} ${trade.number} cancelled`);
}

/* ----- Stock adjustments ----- */

function openAdjustStock(itemId) {
  const item = state.inventory.find((entry) => entry.id === itemId);
  if (!item || !can("inventory", "edit")) return;
  editing = { kind: "adjust", itemId };
  setModalWide(false);
  els.modalKicker.textContent = viewTitle("inventory");
  els.modalTitle.textContent = `Adjust stock · ${item.name}`;
  els.recordForm.innerHTML = `
    <div class="form-error" id="formError" hidden></div>
    <div class="detail-grid">
      <div><span>In stock now</span><strong>${escapeHtml(formatNumber(item.quantity))}</strong></div>
      <div><span>${escapeHtml(fieldLabel("inventory", "sku"))}</span><strong>${escapeHtml(item.sku || "—")}</strong></div>
    </div>
    <p class="muted-note">Use this for changes that are not a sale or purchase: counting errors, damage, expiry, loss, or items used by the business.</p>
    <fieldset class="form-section">
      <legend>Change</legend>
      <div class="form-grid">
        <div class="field full">
          <span class="field-label">What happened?</span>
          <div class="choice-pills">
            <label class="choice-pill"><input type="radio" name="action" value="remove" checked /><span>Remove stock</span></label>
            <label class="choice-pill"><input type="radio" name="action" value="add" /><span>Add stock</span></label>
            <label class="choice-pill"><input type="radio" name="action" value="count" /><span>Set to counted amount</span></label>
          </div>
        </div>
        <div class="field">
          <label for="adjustQuantity">Quantity *</label>
          <input id="adjustQuantity" name="quantity" type="number" min="0" step="any" required />
        </div>
        <div class="field">
          <label for="adjustReason">Reason *</label>
          <select id="adjustReason" name="reason">${ADJUSTMENT_REASONS.map((reason) => `<option>${reason}</option>`).join("")}</select>
        </div>
        <div class="field">
          <label for="adjustDate">Date</label>
          <input id="adjustDate" name="date" type="date" value="${todayIso()}" max="${todayIso()}" />
        </div>
        <div class="field full">
          <label for="adjustNote">Details</label>
          <input id="adjustNote" name="note" type="text" placeholder="Optional, e.g. batch number or who counted" />
        </div>
      </div>
    </fieldset>
    <div class="form-actions">
      <button class="button ghost" id="cancelFormBtn" type="button">Cancel</button>
      <button class="button primary" type="submit">Save adjustment</button>
    </div>
  `;
  els.modalBackdrop.hidden = false;
  byId("adjustQuantity").focus();
}

function submitAdjustment() {
  const item = state.inventory.find((entry) => entry.id === editing.itemId);
  if (!item) {
    closeModal();
    return;
  }
  const form = els.recordForm;
  const action = form.querySelector('input[name="action"]:checked')?.value || "remove";
  const raw = form.elements.quantity.value;
  const quantity = parseMoney(raw);
  const current = parseMoney(item.quantity);
  const date = form.elements.date.value || todayIso();
  if (isBlank(raw) || quantity < 0 || (action !== "count" && quantity <= 0)) {
    showFormError("Enter a quantity above zero.");
    return;
  }
  const change = round3(action === "add" ? quantity : action === "remove" ? -quantity : quantity - current);
  if (!change) {
    showFormError(`The counted amount matches the current stock (${formatNumber(current)}). Nothing to change.`);
    return;
  }
  if (current + change < -0.0005) {
    showFormError(`You can't remove more than the ${formatNumber(current)} in stock.`);
    return;
  }
  if (date > todayIso()) {
    showFormError("The date can't be in the future.");
    return;
  }
  const reason = [form.elements.reason.value, form.elements.note.value.trim()].filter(Boolean).join(": ");
  recordStockMovement(item, change, { type: "Adjustment", reason, date });
  addActivity(`Stock adjusted: ${item.name} ${change > 0 ? "+" : ""}${formatNumber(change)} (${reason})`, "inventory");
  saveState();
  closeModal();
  render();
  showToast(`${item.name}: now ${formatNumber(item.quantity)} in stock`);
}

function submitTransactionForm() {
  if (editing.kind === "sale" || editing.kind === "purchase") submitTrade();
  else if (editing.kind === "payment") submitPayment();
  else if (editing.kind === "cancel") submitCancel();
  else if (editing.kind === "adjust") submitAdjustment();
  else if (editing.kind === "partyPayment") submitPartyPayment();
  else if (editing.kind === "salary") submitSalaryPayment();
  else if (editing.kind === "salaryBatch") submitPayAllSalaries();
}

/* ----- Read-only details: a sale or purchase, an item's history, a contact's history ----- */

function openReadOnlyModal(kicker, title, body, actions = "") {
  editing = { kind: "view" };
  setModalWide(true);
  els.modalKicker.textContent = kicker;
  els.modalTitle.textContent = title;
  els.recordForm.innerHTML = `
    ${body}
    <div class="form-actions">
      ${actions}
      <button class="button ghost" id="cancelFormBtn" type="button">Close</button>
    </div>
  `;
  els.modalBackdrop.hidden = false;
}

function openTradeDetails(kind, id) {
  const config = TRADE[kind];
  const trade = findTrade(kind, id);
  if (!trade || !can(config.area, "view")) return;
  const money = (value) => formatCurrency(value, trade.currency);
  const status = tradeStatus(trade);
  const balance = tradeBalance(trade);
  const entry = state.finance.find((record) => record.id === trade.financeId);
  const sensitive = canAccess("sensitiveNumbers");
  const share = trade.subtotal ? trade.total / trade.subtotal : 1;
  const costKnown = trade.lines.filter((line) => !line.custom && line.unitCost !== "");
  const cost = round2(costKnown.reduce((sum, line) => sum + line.quantity * parseMoney(line.unitCost), 0));
  const body = `
    <div class="detail-grid">
      <div><span>Date</span><strong>${escapeHtml(formatDate(trade.date))}</strong></div>
      <div><span>${config.party}</span><strong>${escapeHtml(trade.partyName || "—")}</strong>${
        trade.partyPhone ? `<small>${escapeHtml(trade.partyPhone)}</small>` : trade.walkIn ? "<small>No details recorded</small>" : ""
      }</div>
      <div><span>Status</span><strong>${badge(status.label, status.tone)}</strong></div>
      <div><span>Recorded by</span><strong>${escapeHtml(trade.createdBy || "—")}</strong><small>${escapeHtml(formatDateTime(trade.createdAt))}</small></div>
      ${trade.reference ? `<div><span>${kind === "purchase" ? "Supplier bill no." : "Reference"}</span><strong>${escapeHtml(trade.reference)}</strong></div>` : ""}
      ${trade.dueDate && balance > 0 ? `<div><span>Balance due by</span><strong>${escapeHtml(formatDate(trade.dueDate))}</strong></div>` : ""}
    </div>
    ${
      trade.status === "Cancelled"
        ? `<div class="form-error">Cancelled ${escapeHtml(formatDateTime(trade.cancelledAt))} by ${escapeHtml(trade.cancelledBy || "—")}: ${escapeHtml(
            trade.cancelReason || "no reason given",
          )}. Stock and ${kind === "sale" ? "income" : "expense"} were reversed.</div>`
        : ""
    }
    <div class="table-wrap bordered">
      <table class="data-table compact">
        <thead><tr><th>Item</th><th class="num">Qty</th><th class="num">${config.priceLabel} each</th><th class="num">Total</th></tr></thead>
        <tbody>
          ${trade.lines
            .map((line) => {
              const note =
                line.custom
                  ? "Not from stock"
                  : kind === "sale" && line.listPrice !== "" && parseMoney(line.listPrice) !== line.unitPrice
                    ? `List ${money(line.listPrice)}`
                    : line.sku;
              return `<tr><td>${titleCell(line.name, note)}</td><td class="num">${formatNumber(line.quantity)}</td><td class="num">${escapeHtml(
                money(line.unitPrice),
              )}</td><td class="num">${escapeHtml(money(line.lineTotal))}</td></tr>`;
            })
            .join("")}
        </tbody>
      </table>
    </div>
    <div class="trade-totals readonly">
      <div><span>Subtotal</span><strong>${escapeHtml(money(trade.subtotal))}</strong></div>
      ${trade.discount ? `<div><span>Discount</span><strong>− ${escapeHtml(money(trade.discount))}</strong></div>` : ""}
      <div class="grand"><span>Total</span><strong>${escapeHtml(money(trade.total))}</strong></div>
      <div><span>Paid</span><strong class="money-in">${escapeHtml(money(trade.amountPaid))}</strong></div>
      <div><span>Balance</span><strong class="${balance > 0 ? "negative-text" : ""}">${escapeHtml(money(balance))}</strong></div>
      ${
        kind === "sale" && sensitive && costKnown.length
          ? `<div><span>Gross profit</span><strong>${escapeHtml(money(round2(trade.total - cost)))}</strong><small>${
              costKnown.length < trade.lines.filter((line) => !line.custom).length ? "Some items had no cost recorded" : `Cost of goods ${money(cost)}`
            }</small></div>`
          : ""
      }
    </div>
    <h3 class="detail-heading">Payments</h3>
    ${
      (trade.payments || []).length
        ? `<div class="table-wrap bordered"><table class="data-table compact"><thead><tr><th>Date</th><th>Method</th><th>Note</th><th class="num">Amount</th><th></th></tr></thead><tbody>${trade.payments
            .map(
              (payment) =>
                `<tr><td>${escapeHtml(formatDate(payment.date))}</td><td>${escapeHtml(payment.method || "—")}</td><td>${escapeHtml(
                  [payment.receiptNumber, payment.note, payment.by].filter(Boolean).join(" · ") || "—",
                )}</td><td class="num">${escapeHtml(money(payment.amount))}</td><td><div class="row-actions">${
                  payment.receiptNumber
                    ? `<button type="button" data-document="receipt:${escapeHtml(payment.receiptNumber)}">Receipt</button>`
                    : entry && entry.payments?.some((item) => item.id === payment.id)
                      ? `<button type="button" data-document="payment:${escapeHtml(entry.id)}:${escapeHtml(payment.id)}">Receipt</button>`
                      : ""
                }</div></td></tr>`,
            )
            .join("")}</tbody></table></div>`
        : `<p class="muted-note">No payments recorded yet.</p>`
    }
    ${trade.notes ? `<h3 class="detail-heading">Notes</h3><p>${escapeHtml(trade.notes)}</p>` : ""}
  `;
  const actions = [
    `<button class="button ghost" type="button" data-document="${kind}:${escapeHtml(trade.id)}">Generate invoice</button>`,
    balance > 0 && canTakePayment(entry)
      ? `<button class="button primary" type="button" data-mark-paid="${escapeHtml(entry.id)}">${kind === "sale" ? "Receive payment" : "Pay supplier"}</button>`
      : "",
    trade.status !== "Cancelled" && can(config.area, "delete")
      ? `<button class="button danger-ghost" type="button" data-cancel-trade="${kind}:${escapeHtml(trade.id)}">Cancel ${config.noun.toLowerCase()}</button>`
      : "",
  ].join("");
  openReadOnlyModal(viewTitle(config.area), `${config.noun} ${trade.number}`, body, actions);
}

function openStockHistory(itemId) {
  const item = state.inventory.find((entry) => entry.id === itemId);
  if (!item) return;
  const moves = state.stockMovements.filter((move) => move.itemId === itemId);
  const units = (type) => moves.filter((move) => move.type === type).reduce((sum, move) => sum + Math.abs(move.change), 0);
  const body = `
    <div class="detail-grid">
      <div><span>In stock now</span><strong>${escapeHtml(formatNumber(item.quantity))}</strong></div>
      <div><span>Sold</span><strong>${escapeHtml(formatNumber(units("Sale") - units("Sale cancelled")))}</strong></div>
      <div><span>Bought</span><strong>${escapeHtml(formatNumber(units("Purchase") - units("Purchase cancelled")))}</strong></div>
      <div><span>Adjustments</span><strong>${escapeHtml(formatNumber(moves.filter((move) => move.type === "Adjustment").length))}</strong></div>
    </div>
    ${movementTableHtml(moves, { showItem: false })}
  `;
  openReadOnlyModal(viewTitle("inventory"), `Stock history · ${item.name}`, body);
}

function contactTrades(contact) {
  const name = String(contact.name || "").trim().toLowerCase();
  const matches = (trade) => trade.contactId === contact.id || (!trade.contactId && String(trade.partyName || "").trim().toLowerCase() === name);
  return {
    sales: state.sales.filter(matches),
    purchases: state.purchases.filter(matches),
  };
}

function openContactHistory(contactId) {
  const contact = state.contacts.find((entry) => entry.id === contactId);
  if (!contact) return;
  const { sales, purchases } = contactTrades(contact);
  const live = (list) => list.filter((trade) => trade.status !== "Cancelled");
  const sum = (list, amountOf) => formatMoneyTotals(sumByCurrency(list, amountOf));
  const rows = [
    ...(can("sales", "view") ? sales.map((trade) => ({ kind: "sale", trade })) : []),
    ...(can("purchases", "view") ? purchases.map((trade) => ({ kind: "purchase", trade })) : []),
  ].sort((a, b) => String(b.trade.date).localeCompare(String(a.trade.date)));
  const body = `
    <div class="detail-grid">
      <div><span>Type</span><strong>${escapeHtml(contact.type || "—")}</strong></div>
      <div><span>Phone</span><strong>${escapeHtml(contact.phone || "—")}</strong></div>
      ${can("sales", "view") ? `<div><span>Bought from you</span><strong>${escapeHtml(sum(live(sales), (trade) => trade.total))}</strong><small>${live(sales).length} sales</small></div>` : ""}
      ${can("purchases", "view") ? `<div><span>You bought from them</span><strong>${escapeHtml(sum(live(purchases), (trade) => trade.total))}</strong><small>${live(purchases).length} purchases</small></div>` : ""}
      <div><span>Balance due</span><strong>${contactDuesHtml(contact)}</strong></div>
    </div>
    ${
      rows.length
        ? `<div class="table-wrap bordered"><table class="data-table compact"><thead><tr><th>No.</th><th>Date</th><th>Items</th><th class="num">Total</th><th class="num">Balance</th><th>Status</th><th></th></tr></thead><tbody>${rows
            .map(({ kind, trade }) => {
              const status = tradeStatus(trade);
              return `<tr>
                <td><strong>${escapeHtml(trade.number)}</strong><br><span class="muted-note">${kind === "sale" ? "Sale" : "Purchase"}</span></td>
                <td>${escapeHtml(formatDate(trade.date))}</td>
                <td>${escapeHtml(tradeItemsSummary(trade))}</td>
                <td class="num">${escapeHtml(formatCurrency(trade.total, trade.currency))}</td>
                <td class="num">${escapeHtml(formatCurrency(tradeBalance(trade), trade.currency))}</td>
                <td>${badge(status.label, status.tone)}</td>
                <td><div class="row-actions"><button type="button" data-view-trade="${kind}:${escapeHtml(trade.id)}">Open</button></div></td>
              </tr>`;
            })
            .join("")}</tbody></table></div>`
        : `<div class="empty-state">No sales or purchases recorded with ${escapeHtml(contact.name)} yet.</div>`
    }
    <h3 class="detail-heading">Payments</h3>
    <div class="table-wrap bordered">${paymentsTableHtml(
      allPayments({ filter: (entry) => String(entry.party || "").trim().toLowerCase() === String(contact.name || "").trim().toLowerCase() }),
      { emptyText: `No payments recorded with ${contact.name} yet.` },
    )}</div>
  `;
  const dueKey = `${String(contact.name || "").trim().toLowerCase()}|${recordCurrency(contact)}`;
  const actions = ["in", "out"]
    .filter((direction) => findDueParty(direction, dueKey))
    .map(
      (direction) =>
        `<button class="button ghost" type="button" data-document="statement:${direction}:${escapeHtml(dueKey)}">Statement</button>${
          findDueParty(direction, dueKey).items.some(dueItemCanBePaid)
            ? `<button class="button primary" type="button" data-party-payment="${direction}:${escapeHtml(dueKey)}">${direction === "in" ? "Receive payment" : "Pay"}</button>`
            : ""
        }`,
    )
    .join("");
  openReadOnlyModal(viewTitle("contacts"), contact.name, body, actions);
}

function movementTableHtml(moves, { showItem = true, limit = 300 } = {}) {
  if (!moves.length) return `<div class="empty-state">No stock changes recorded yet.</div>`;
  const shown = moves.slice(0, limit);
  return `
    <div class="table-wrap ${showItem ? "" : "bordered"}">
      <table class="data-table compact">
        <thead>
          <tr>
            <th>Date</th>
            ${showItem ? "<th>Item</th>" : ""}
            <th>Change type</th>
            <th class="num">Change</th>
            <th class="num">Stock after</th>
            <th>Reference</th>
            <th>Who / why</th>
          </tr>
        </thead>
        <tbody>
          ${shown
            .map(
              (move) => `
            <tr>
              <td>${escapeHtml(formatDate(move.date))}</td>
              ${showItem ? `<td>${titleCell(move.itemName, move.sku)}</td>` : ""}
              <td>${movementBadge(move.type)}</td>
              <td class="num"><strong class="${move.change > 0 ? "money-in" : "negative-text"}">${move.change > 0 ? "+" : "−"}${escapeHtml(
                formatNumber(Math.abs(move.change)),
              )}</strong></td>
              <td class="num">${escapeHtml(formatNumber(move.balanceAfter))}</td>
              <td>${
                move.referenceId && move.referenceKind && can(TRADE[move.referenceKind].area, "view")
                  ? `<button type="button" class="link-button" data-view-trade="${move.referenceKind}:${escapeHtml(move.referenceId)}">${escapeHtml(move.reference)}</button>`
                  : move.type === "Adjustment"
                    ? `<button type="button" class="link-button" data-document="movement:${escapeHtml(move.id)}">${escapeHtml(move.documentNumber || "Adjustment note")}</button>`
                    : escapeHtml(move.reference || "—")
              }</td>
              <td>${titleCell(move.party || move.reason || "—", [move.party ? move.reason : "", move.by].filter(Boolean).join(" · "))}</td>
            </tr>
          `,
            )
            .join("")}
        </tbody>
      </table>
    </div>
    ${moves.length > limit ? `<p class="muted-note">Showing the latest ${limit} of ${formatNumber(moves.length)} changes.</p>` : ""}
  `;
}

/* ---------- Payments ledger, dues by person, and payroll payments ----------
 *
 * Every payment lives inside the finance entry it pays (a sale's income, a purchase's
 * expense, a salary, an opening balance, or any other entry). These helpers list them
 * in one place, in the order they happened, and let dues be settled per person.
 */

const paymentFilters = { direction: "all", month: "all" };
const payHistoryFilters = { employeeId: "all" };

function isSalaryEntry(entry) {
  return Boolean(entry?.employeeId && entry?.payrollMonth);
}

function canSeeFinanceEntry(entry) {
  if (entry.sourceType && TRADE[entry.sourceType]) return can(TRADE[entry.sourceType].area, "view") || can("finance", "view");
  if (isSalaryEntry(entry)) return (can("employees", "view") || can("finance", "view")) && canAccess("sensitiveNumbers");
  return can("finance", "view");
}

// Payments recorded on an entry. Older entries marked paid without payment details
// count as one payment on the day they were paid.
function entryPayments(entry) {
  if (Array.isArray(entry.payments) && entry.payments.length) return entry.payments;
  if (isUnpaid(entry) || !(parseMoney(entry.amount) > 0)) return [];
  return [
    {
      id: `auto-${entry.id}`,
      date: entry.paidOn || entry.date,
      amount: parseMoney(entry.amount),
      method: entry.paymentMethod || "",
      note: "",
      by: "",
      auto: true,
    },
  ];
}

function paymentSource(entry) {
  if (entry.sourceType && TRADE[entry.sourceType]) {
    const trade = findTrade(entry.sourceType, entry.sourceId);
    return {
      label: `${TRADE[entry.sourceType].noun} ${trade?.number || entry.reference || ""}`.trim(),
      tone: entry.sourceType === "sale" ? "blue" : "green",
      open:
        trade && can(TRADE[entry.sourceType].area, "view")
          ? `<button type="button" data-view-trade="${entry.sourceType}:${escapeHtml(trade.id)}">Open</button>`
          : "",
    };
  }
  if (isSalaryEntry(entry)) {
    return {
      label: `Salary · ${monthName(entry.payrollMonth, "long")}`,
      tone: "violet",
      open: canOpenView("employees") ? `<button type="button" data-open-payroll="${escapeHtml(entry.payrollMonth)}">Payroll</button>` : "",
    };
  }
  if (entry.settlesContactId || entry.category === "Customer payment" || entry.category === "Supplier payment") {
    return { label: "Opening balance", tone: "blue", open: "" };
  }
  return { label: entry.category || (entry.type === "Income" ? "Income" : "Expense"), tone: "amber", open: "" };
}

function allPayments({ filter = () => true } = {}) {
  const rows = [];
  state.finance.forEach((entry) => {
    if (!canSeeFinanceEntry(entry) || !filter(entry)) return;
    entryPayments(entry).forEach((payment) => {
      rows.push({
        entry,
        payment,
        direction: entry.type === "Income" ? "in" : "out",
        currency: recordCurrency(entry),
        amount: parseMoney(payment.amount),
        date: payment.date || entry.date || "",
        party: entry.party || "",
        cancelled: Boolean(entry.cancelled),
      });
    });
  });
  return rows.sort(
    (a, b) =>
      String(b.date).localeCompare(String(a.date)) ||
      String(b.payment.at || b.entry.createdAt || "").localeCompare(String(a.payment.at || a.entry.createdAt || "")),
  );
}

function paymentDocumentRef(row) {
  if (isSalaryEntry(row.entry)) return `salary:${row.entry.id}`;
  if (row.payment.auto) return `finance:${row.entry.id}`;
  if (row.payment.receiptNumber) return `receipt:${row.payment.receiptNumber}`;
  return `payment:${row.entry.id}:${row.payment.id}`;
}

function paymentsTableHtml(rows, options = {}) {
  const { head, body } = paymentsTableParts(rows, options);
  return `<table class="data-table"><thead>${head}</thead><tbody>${body}</tbody></table>`;
}

function paymentsTableParts(rows, { showSource = true, emptyText = "No payments recorded yet." } = {}) {
  const columns = showSource ? 7 : 6;
  const head = `
    <tr>
      <th>Date</th>
      <th>Receipt no.</th>
      <th>${showSource ? "Received from / paid to" : "Party"}</th>
      ${showSource ? "<th>For</th>" : ""}
      <th>Method</th>
      <th class="num">Amount</th>
      <th></th>
    </tr>
  `;
  if (!rows.length) return { head, body: `<tr><td colspan="${columns}"><div class="empty-state">${escapeHtml(emptyText)}</div></td></tr>` };
  const body = `
        ${rows
          .map((row) => {
            const source = paymentSource(row.entry);
            const signed = `${row.direction === "in" ? "+" : "−"} ${formatCurrency(row.amount, row.currency)}`;
            return `
            <tr class="${row.cancelled ? "is-muted" : ""}">
              <td>${titleCell(formatDate(row.date), row.payment.by || "")}</td>
              <td class="nowrap">${escapeHtml(row.payment.receiptNumber || (isSalaryEntry(row.entry) ? row.entry.slipNumber || "—" : "—"))}</td>
              <td>${titleCell(row.party || "—", row.direction === "in" ? "Received" : "Paid")}</td>
              ${
                showSource
                  ? `<td>${badge(source.label, source.tone)}${row.cancelled ? ` ${badge(isSalaryEntry(row.entry) ? "Reversed" : "Cancelled", "red")}` : ""}${
                      row.payment.note ? `<div class="muted-note">${escapeHtml(row.payment.note)}</div>` : ""
                    }</td>`
                  : ""
              }
              <td>${escapeHtml(row.payment.method || "—")}</td>
              <td class="num"><strong class="${row.direction === "in" ? "money-in" : "negative-text"}">${escapeHtml(signed)}</strong></td>
              <td><div class="row-actions">
                <button type="button" data-document="${escapeHtml(paymentDocumentRef(row))}">${isSalaryEntry(row.entry) ? "Salary slip" : "Receipt"}</button>
                ${source.open}
              </div></td>
            </tr>
          `;
          })
          .join("")}
  `;
  return { head, body };
}

function paymentTotalsText(rows) {
  const live = rows.filter((row) => !row.cancelled);
  const received = formatMoneyTotals(sumByCurrency(live.filter((row) => row.direction === "in"), (row) => row.amount));
  const paid = formatMoneyTotals(sumByCurrency(live.filter((row) => row.direction === "out"), (row) => row.amount));
  return { received, paid };
}

function renderPaymentsTab() {
  const query = els.globalSearch.value.trim().toLowerCase();
  const all = allPayments();
  const months = [...new Set(all.map((row) => monthKeyOf(row.date)).filter(Boolean))].sort().reverse();
  if (paymentFilters.month !== "all" && !months.includes(paymentFilters.month)) paymentFilters.month = "all";
  byId("paymentsDirectionFilter").value = paymentFilters.direction;
  byId("paymentsMonthFilter").innerHTML = [
    `<option value="all">All months</option>`,
    ...months.map((key) => `<option value="${key}" ${key === paymentFilters.month ? "selected" : ""}>${monthName(key, "long")}</option>`),
  ].join("");
  const rows = all.filter((row) => {
    if (paymentFilters.direction !== "all" && row.direction !== paymentFilters.direction) return false;
    if (paymentFilters.month !== "all" && monthKeyOf(row.date) !== paymentFilters.month) return false;
    if (!query) return true;
    return [row.party, row.payment.receiptNumber, row.payment.method, row.payment.note, paymentSource(row.entry).label, row.entry.description]
      .join(" ")
      .toLowerCase()
      .includes(query);
  });
  const { received, paid } = paymentTotalsText(rows);
  byId("paymentsSummary").innerHTML = `<strong>${formatNumber(rows.length)}</strong> payments · Received <strong class="money-in">${escapeHtml(
    received,
  )}</strong> · Paid out <strong class="negative-text">${escapeHtml(paid)}</strong>`;
  const { head, body } = paymentsTableParts(rows, {
    emptyText: all.length ? "No payments match this view." : "No payments yet. Payments appear here when sales, purchases, dues, and salaries are paid.",
  });
  byId("paymentsHead").innerHTML = head;
  byId("paymentsLedger").innerHTML = body;
}

function renderTradePayments(kind) {
  const config = TRADE[kind];
  const query = els.globalSearch.value.trim().toLowerCase();
  const period = tradeFilters[config.area];
  byId(`${config.area}PeriodFilter`).value = period;
  const rows = allPayments({ filter: (entry) => entry.sourceType === kind }).filter((row) => {
    if (!inPeriod(row.date, period)) return false;
    if (!query) return true;
    return [row.party, row.payment.receiptNumber, row.payment.method, row.payment.note, row.entry.reference].join(" ").toLowerCase().includes(query);
  });
  const { head, body } = paymentsTableParts(rows, {
    emptyText: kind === "sale" ? "No payments received for sales in this period." : "No payments made for purchases in this period.",
  });
  byId(`${config.area}Head`).innerHTML = head;
  byId(`${config.area}Table`).innerHTML = body;
  const { received, paid } = paymentTotalsText(rows);
  byId(`${config.area}TableSummary`).innerHTML = `<strong>${formatNumber(rows.length)}</strong> payments · ${
    kind === "sale" ? `Received <strong>${escapeHtml(received)}</strong>` : `Paid <strong>${escapeHtml(paid)}</strong>`
  }`;
}

/* ----- Dues grouped by person ----- */

function duePartyKey(item) {
  return `${String(item.party || "").trim().toLowerCase()}|${item.currency}`;
}

function dueParties(items) {
  const groups = new Map();
  items
    .filter((item) => item.kind !== "salary")
    .forEach((item) => {
      const key = duePartyKey(item);
      if (!groups.has(key)) groups.set(key, { key, party: item.party || "", currency: item.currency, items: [], total: 0, overdue: 0 });
      const group = groups.get(key);
      group.items.push(item);
      group.total = round2(group.total + item.amount);
      if (item.overdue) group.overdue += 1;
    });
  return [...groups.values()].sort((a, b) => b.overdue - a.overdue || b.total - a.total);
}

// Oldest first: opening balances, then by date.
function allocationOrder(items) {
  return [...items].sort(
    (a, b) =>
      Number(b.kind === "balance") - Number(a.kind === "balance") ||
      String(a.date || a.dueDate || "").localeCompare(String(b.date || b.dueDate || "")),
  );
}

function dueItemCanBePaid(item) {
  if (item.kind === "transaction") return canTakePayment(state.finance.find((entry) => entry.id === item.id));
  if (item.kind === "balance") return can("finance", "add");
  return false;
}

function dueItemLinks(item) {
  if (item.kind !== "transaction") return "";
  const entry = state.finance.find((record) => record.id === item.id);
  if (!entry) return "";
  const source = paymentSource(entry);
  return `${source.open}<button type="button" data-document="finance:${escapeHtml(entry.id)}">Invoice</button>`;
}

function duePartyHtml(group, direction) {
  const payable = group.items.some(dueItemCanBePaid);
  const name = group.party || (direction === "in" ? "No name recorded" : "No name recorded");
  return `
    <div class="due-party ${group.overdue ? "overdue" : ""}">
      <div class="due-party-head">
        <div>
          <strong>${escapeHtml(name)}</strong>
          <span>${group.items.length} open${group.overdue ? ` · ${group.overdue} overdue` : ""}</span>
        </div>
        <div class="due-side">
          <span class="due-amount">${escapeHtml(formatCurrency(group.total, group.currency))}</span>
          <div class="row-actions">
            <button type="button" data-document="statement:${direction}:${escapeHtml(group.key)}">Statement</button>
            ${
              payable
                ? `<button type="button" class="accent" data-party-payment="${direction}:${escapeHtml(group.key)}">${direction === "in" ? "Receive payment" : "Pay"}</button>`
                : ""
            }
          </div>
        </div>
      </div>
      <div class="due-party-items">
        ${allocationOrder(group.items)
          .map((item) => {
            const source =
              item.kind === "balance"
                ? badge("Opening balance", "blue")
                : badge(paymentSource(state.finance.find((entry) => entry.id === item.id) || {}).label || "Unpaid", "amber");
            const when = item.overdue
              ? badge(item.dueDate ? `Overdue · due ${formatDate(item.dueDate)}` : "Overdue", "red")
              : item.dueDate
                ? `<span>Due ${escapeHtml(formatDate(item.dueDate))}</span>`
                : item.date
                  ? `<span>${escapeHtml(formatDate(item.date))}</span>`
                  : "";
            return `
            <div class="due-line">
              <div class="due-meta">${source}${when}${item.detail ? `<span>${escapeHtml(item.detail)}</span>` : ""}</div>
              <div class="due-side">
                <span>${escapeHtml(formatCurrency(item.amount, item.currency))}</span>
                <div class="row-actions">
                  ${dueItemLinks(item)}
                  ${
                    item.kind === "transaction" && dueItemCanBePaid(item) && group.items.length > 1
                      ? `<button type="button" data-mark-paid="${escapeHtml(item.id)}">${direction === "in" ? "Receive" : "Pay"} this</button>`
                      : ""
                  }
                </div>
              </div>
            </div>
          `;
          })
          .join("")}
      </div>
    </div>
  `;
}

function salaryDuesHtml(items) {
  const salaries = items.filter((item) => item.kind === "salary");
  if (!salaries.length) return "";
  const byMonth = {};
  salaries.forEach((item) => {
    (byMonth[item.month] ||= []).push(item);
  });
  return `
    <p class="due-group-title">Staff salaries · paid from Employees → Payroll</p>
    ${Object.keys(byMonth)
      .sort()
      .map((month) => {
        const list = byMonth[month];
        const late = list.some((item) => item.overdue);
        return `
        <div class="due-item ${late ? "overdue" : ""}">
          <div>
            <strong>Salaries for ${escapeHtml(monthName(month, "long"))}</strong>
            <span>${list.length} employee${list.length === 1 ? "" : "s"}: ${escapeHtml(list.slice(0, 3).map((item) => item.party).join(", "))}${
              list.length > 3 ? "…" : ""
            }</span>
            <div class="due-meta">${badge("Salary", "violet")}${late ? badge("Overdue", "red") : ""}</div>
          </div>
          <div class="due-side">
            <span class="due-amount">${escapeHtml(formatMoneyTotals(duesTotals(list)))}</span>
            ${canOpenView("employees") ? `<div class="row-actions"><button type="button" class="accent" data-open-payroll="${month}">Open payroll</button></div>` : ""}
          </div>
        </div>
      `;
      })
      .join("")}
  `;
}

function duePartyListHtml(items, direction) {
  const groups = dueParties(items);
  const salaries = direction === "out" ? salaryDuesHtml(items) : "";
  if (!groups.length && !salaries) {
    return `<div class="empty-state">${direction === "in" ? "Nobody owes the business anything right now." : "Nothing to pay right now."}</div>`;
  }
  return `${groups.map((group) => duePartyHtml(group, direction)).join("")}${salaries}`;
}

function findDueParty(direction, key) {
  const dues = buildDues();
  return dueParties(direction === "in" ? dues.receivables : dues.payables).find((group) => group.key === key) || null;
}

function openPartyPayment(direction, key) {
  const group = findDueParty(direction, key);
  if (!group) {
    showToast("Nothing is owed here any more");
    render();
    return;
  }
  const items = allocationOrder(group.items).filter(dueItemCanBePaid);
  if (!items.length) {
    showToast("Your role cannot record payments for these dues");
    return;
  }
  const payableTotal = round2(items.reduce((sum, item) => sum + item.amount, 0));
  const incoming = direction === "in";
  editing = { kind: "partyPayment", direction, key };
  setModalWide(false);
  els.modalKicker.textContent = "Dues";
  els.modalTitle.textContent = incoming ? `Receive payment from ${group.party || "customer"}` : `Pay ${group.party || "supplier"}`;
  els.recordForm.innerHTML = `
    <div class="form-error" id="formError" hidden></div>
    <div class="detail-grid">
      <div><span>${incoming ? "From" : "To"}</span><strong>${escapeHtml(group.party || "—")}</strong></div>
      <div><span>Open items</span><strong>${items.length}</strong></div>
      <div><span>Total ${incoming ? "owed to you" : "you owe"}</span><strong class="negative-text">${escapeHtml(formatCurrency(payableTotal, group.currency))}</strong></div>
    </div>
    <fieldset class="form-section">
      <legend>Payment</legend>
      <div class="form-grid">
        <div class="field">
          <label for="partyPaymentAmount">Amount (${group.currency}) *</label>
          <input id="partyPaymentAmount" name="amount" type="number" min="0" step="0.01" max="${payableTotal}" value="${payableTotal}" required data-party-allocation />
          <small>Enter less for a part payment. It pays the oldest items first.</small>
        </div>
        <div class="field">
          <label for="partyPaymentDate">Date</label>
          <input id="partyPaymentDate" name="date" type="date" value="${todayIso()}" max="${todayIso()}" />
        </div>
        <div class="field">
          <label for="partyPaymentMethod">Method</label>
          <select id="partyPaymentMethod" name="method">${PAYMENT_METHODS.map((method) => `<option>${method}</option>`).join("")}</select>
        </div>
        <div class="field">
          <label for="partyPaymentNote">Note</label>
          <input id="partyPaymentNote" name="note" type="text" placeholder="Cheque no., transfer ID, …" />
        </div>
      </div>
    </fieldset>
    <fieldset class="form-section">
      <legend>How it will be applied</legend>
      <div id="partyAllocation">${allocationPreviewHtml(items, payableTotal, group.currency)}</div>
    </fieldset>
    <div class="form-actions">
      <button class="button ghost" id="cancelFormBtn" type="button">Cancel</button>
      <button class="button primary" type="submit">Save payment</button>
    </div>
  `;
  els.modalBackdrop.hidden = false;
  byId("partyPaymentAmount").focus();
}

function allocate(items, amount) {
  let remaining = round2(amount);
  return items.map((item) => {
    const take = round2(Math.max(0, Math.min(remaining, item.amount)));
    remaining = round2(remaining - take);
    return { item, take };
  });
}

function dueItemLabel(item) {
  if (item.kind === "balance") return "Opening balance";
  const entry = state.finance.find((record) => record.id === item.id);
  return entry ? paymentSource(entry).label : item.title;
}

function allocationPreviewHtml(items, amount, currency) {
  return `
    <div class="allocation-list">
      ${allocate(items, amount)
        .map(
          ({ item, take }) => `
        <div class="allocation-row ${take ? "" : "is-muted"}">
          <span>${escapeHtml(dueItemLabel(item))}${item.dueDate ? ` · due ${escapeHtml(formatDate(item.dueDate))}` : ""}</span>
          <span>${escapeHtml(formatCurrency(take, currency))} of ${escapeHtml(formatCurrency(item.amount, currency))}${
            take && take < item.amount ? " · part" : take ? " · settled" : ""
          }</span>
        </div>`,
        )
        .join("")}
    </div>
  `;
}

function refreshPartyAllocation() {
  if (editing?.kind !== "partyPayment") return;
  const group = findDueParty(editing.direction, editing.key);
  if (!group) return;
  const items = allocationOrder(group.items).filter(dueItemCanBePaid);
  byId("partyAllocation").innerHTML = allocationPreviewHtml(items, parseMoney(byId("partyPaymentAmount").value), group.currency);
}

function submitPartyPayment() {
  const { direction, key } = editing;
  const group = findDueParty(direction, key);
  const form = els.recordForm.elements;
  if (!group) {
    showFormError("These dues have changed. Close this window and try again.");
    return;
  }
  const items = allocationOrder(group.items).filter(dueItemCanBePaid);
  const payableTotal = round2(items.reduce((sum, item) => sum + item.amount, 0));
  const amount = round2(parseMoney(form.amount.value));
  const date = form.date.value || todayIso();
  const method = form.method.value;
  const note = form.note.value.trim();
  if (!(amount > 0) || amount > payableTotal + 0.005) {
    showFormError(`Enter an amount above zero and up to ${formatCurrency(payableTotal, group.currency)}.`);
    return;
  }
  if (date > todayIso()) {
    showFormError("The payment date can't be in the future.");
    return;
  }
  const receiptNumber = nextReceiptNumber(direction);
  const groupId = makeId("receipt");
  const user = currentUser();
  const now = new Date().toISOString();
  allocate(items, amount).forEach(({ item, take }) => {
    if (!take) return;
    if (item.kind === "transaction") {
      const entry = state.finance.find((record) => record.id === item.id);
      applyPayment(entry, take, { date, method, note, receiptNumber, groupId });
      return;
    }
    // Opening balance: record the money as its own paid entry and reduce the balance.
    const contact = state.contacts.find((record) => record.id === item.id);
    if (!contact) return;
    const balance = parseMoney(contact.balance);
    contact.balance = round2(Math.sign(balance) * Math.max(0, Math.abs(balance) - take));
    state.finance.unshift({
      id: makeId("finance"),
      date,
      type: direction === "in" ? "Income" : "Expense",
      category: direction === "in" ? "Customer payment" : "Supplier payment",
      description: `Opening balance ${direction === "in" ? "received from" : "paid to"} ${contact.name}`,
      amount: take,
      amountPaid: take,
      currency: group.currency,
      status: "Paid",
      paidOn: date,
      party: contact.name,
      settlesContactId: contact.id,
      payments: [{ id: makeId("payment"), date, amount: take, method, note, by: user?.name || "", at: now, receiptNumber, groupId }],
      sourceName: "Dues",
      createdAt: now,
      infoDismissed: true,
    });
  });
  addActivity(
    `${direction === "in" ? "Payment received from" : "Payment made to"} ${group.party || "—"}: ${formatCurrency(amount, group.currency)} (${receiptNumber})`,
    "finance",
  );
  saveState();
  closeModal();
  render();
  showToast(`Payment saved · ${receiptNumber}`);
  openDocumentFor(`receipt:${receiptNumber}`);
}

/* ----- Payroll payments ----- */

function openSalaryPayment(employeeId, month) {
  const employee = state.employees.find((record) => record.id === employeeId);
  if (!employee || !can("finance", "add") || !canAccess("sensitiveNumbers")) {
    showToast("Your role cannot pay salaries");
    return;
  }
  if (salaryPayment(employee, month)) {
    showToast(`${employee.name} is already paid for ${monthName(month, "long")}`);
    return;
  }
  const salary = parseMoney(employee.salary);
  if (!salary) {
    showToast("Set this employee's salary first");
    return;
  }
  const currency = recordCurrency(employee);
  editing = { kind: "salary", employeeId, month };
  setModalWide(false);
  els.modalKicker.textContent = viewTitle("employees");
  els.modalTitle.textContent = `Pay ${employee.name}`;
  els.recordForm.innerHTML = `
    <div class="form-error" id="formError" hidden></div>
    <div class="detail-grid">
      <div><span>Employee</span><strong>${escapeHtml(employee.name)}</strong><small>${escapeHtml([employee.role, employee.department].filter(Boolean).join(" · "))}</small></div>
      <div><span>Salary for</span><strong>${escapeHtml(monthName(month, "long"))}</strong></div>
      <div><span>Monthly salary</span><strong>${escapeHtml(formatCurrency(salary, currency))}</strong></div>
    </div>
    <fieldset class="form-section">
      <legend>Payment</legend>
      <div class="form-grid">
        <div class="field">
          <label for="salaryBonus">Bonus / overtime (${currency})</label>
          <input id="salaryBonus" name="bonus" type="number" min="0" step="0.01" placeholder="0" data-salary-net />
        </div>
        <div class="field">
          <label for="salaryDeduction">Deductions (${currency})</label>
          <input id="salaryDeduction" name="deduction" type="number" min="0" step="0.01" placeholder="0" data-salary-net />
          <small>Advances, absences, fines…</small>
        </div>
        <div class="field">
          <label for="salaryDate">Date paid</label>
          <input id="salaryDate" name="date" type="date" value="${todayIso()}" max="${todayIso()}" />
        </div>
        <div class="field">
          <label for="salaryMethod">Method</label>
          <select id="salaryMethod" name="method">${PAYMENT_METHODS.map((method) => `<option>${method}</option>`).join("")}</select>
        </div>
        <div class="field full">
          <label for="salaryNote">Note</label>
          <input id="salaryNote" name="note" type="text" placeholder="Reason for bonus or deduction" />
        </div>
        <div class="field full"><div class="trade-balance" id="salaryNet">Net pay: <strong>${escapeHtml(formatCurrency(salary, currency))}</strong></div></div>
      </div>
    </fieldset>
    <div class="form-actions">
      <button class="button ghost" id="cancelFormBtn" type="button">Cancel</button>
      <button class="button primary" type="submit">Pay and create salary slip</button>
    </div>
  `;
  els.modalBackdrop.hidden = false;
}

function salaryNetFromForm(employee) {
  const form = els.recordForm.elements;
  const bonus = round2(Math.max(0, parseMoney(form.bonus?.value)));
  const deduction = round2(Math.max(0, parseMoney(form.deduction?.value)));
  return { bonus, deduction, net: round2(parseMoney(employee.salary) + bonus - deduction) };
}

function refreshSalaryNet() {
  if (editing?.kind !== "salary") return;
  const employee = state.employees.find((record) => record.id === editing.employeeId);
  if (!employee) return;
  const { net } = salaryNetFromForm(employee);
  byId("salaryNet").innerHTML = `Net pay: <strong>${escapeHtml(formatCurrency(net, recordCurrency(employee)))}</strong>`;
}

function recordSalaryPayment(employee, month, { date, method, note = "", bonus = 0, deduction = 0 }) {
  const currency = recordCurrency(employee);
  const base = parseMoney(employee.salary);
  const amount = round2(base + bonus - deduction);
  const slipNumber = nextDocumentNumber("salary", "SAL");
  const now = new Date().toISOString();
  const entry = {
    id: makeId("finance"),
    date,
    type: "Expense",
    category: "Salaries",
    description: `Salary - ${employee.name} - ${monthName(month, "long")}`,
    amount,
    amountPaid: amount,
    currency,
    status: "Paid",
    party: employee.name,
    employeeId: employee.id,
    payrollMonth: month,
    baseSalary: base,
    bonus,
    deduction,
    slipNumber,
    paidOn: date,
    notes: note,
    payments: [{ id: makeId("payment"), date, amount, method, note, by: currentUser()?.name || "", at: now, receiptNumber: slipNumber }],
    sourceName: "Payroll",
    createdAt: now,
    infoDismissed: true,
  };
  state.finance.unshift(entry);
  return entry;
}

function submitSalaryPayment() {
  const { employeeId, month } = editing;
  const employee = state.employees.find((record) => record.id === employeeId);
  const form = els.recordForm.elements;
  if (!employee || salaryPayment(employee, month)) {
    showFormError("This salary has already been paid.");
    return;
  }
  const { bonus, deduction, net } = salaryNetFromForm(employee);
  const date = form.date.value || todayIso();
  if (net <= 0) {
    showFormError("Deductions can't be more than the salary and bonus.");
    return;
  }
  if (date > todayIso()) {
    showFormError("The payment date can't be in the future.");
    return;
  }
  const entry = recordSalaryPayment(employee, month, { date, method: form.method.value, note: form.note.value.trim(), bonus, deduction });
  addActivity(`Salary paid: ${employee.name} · ${formatCurrency(net, entry.currency)} for ${monthName(month, "long")} (${entry.slipNumber})`, "employees");
  saveState();
  closeModal();
  render();
  showToast(`${employee.name}'s salary paid · ${entry.slipNumber}`);
  openDocumentFor(`salary:${entry.id}`);
}

function openPayAllSalaries(month) {
  const unpaid = payrollRows(month)
    .filter((row) => row.unpaid)
    .map((row) => row.employee);
  if (!unpaid.length || !can("finance", "add")) return;
  editing = { kind: "salaryBatch", month };
  setModalWide(false);
  els.modalKicker.textContent = viewTitle("employees");
  els.modalTitle.textContent = `Pay salaries for ${monthName(month, "long")}`;
  els.recordForm.innerHTML = `
    <div class="form-error" id="formError" hidden></div>
    <div class="allocation-list">
      ${unpaid
        .map(
          (employee) => `<div class="allocation-row"><span>${escapeHtml(employee.name)}${employee.role ? ` · ${escapeHtml(employee.role)}` : ""}</span><span>${escapeHtml(
            formatRecordMoney(employee, employee.salary),
          )}</span></div>`,
        )
        .join("")}
      <div class="allocation-row total"><span>Total</span><span>${escapeHtml(formatMoneyTotals(sumByCurrency(unpaid, (employee) => parseMoney(employee.salary))))}</span></div>
    </div>
    <fieldset class="form-section">
      <legend>Payment</legend>
      <div class="form-grid">
        <div class="field">
          <label for="batchSalaryDate">Date paid</label>
          <input id="batchSalaryDate" name="date" type="date" value="${todayIso()}" max="${todayIso()}" />
        </div>
        <div class="field">
          <label for="batchSalaryMethod">Method</label>
          <select id="batchSalaryMethod" name="method">${PAYMENT_METHODS.map((method) => `<option>${method}</option>`).join("")}</select>
        </div>
      </div>
      <p class="muted-note">Each employee gets their own salary slip. Use Pay salary on one employee to add a bonus or deduction.</p>
    </fieldset>
    <div class="form-actions">
      <button class="button ghost" id="cancelFormBtn" type="button">Cancel</button>
      <button class="button primary" type="submit">Pay ${unpaid.length} salar${unpaid.length === 1 ? "y" : "ies"}</button>
    </div>
  `;
  els.modalBackdrop.hidden = false;
}

function submitPayAllSalaries() {
  const { month } = editing;
  const form = els.recordForm.elements;
  const date = form.date.value || todayIso();
  if (date > todayIso()) {
    showFormError("The payment date can't be in the future.");
    return;
  }
  const unpaid = payrollRows(month)
    .filter((row) => row.unpaid)
    .map((row) => row.employee);
  const entries = unpaid.map((employee) => recordSalaryPayment(employee, month, { date, method: form.method.value }));
  const total = formatMoneyTotals(sumByCurrency(entries, (entry) => entry.amount));
  addActivity(`Paid ${entries.length} salaries for ${monthName(month, "long")} (${total})`, "employees");
  saveState();
  closeModal();
  render();
  showToast(`${entries.length} salar${entries.length === 1 ? "y" : "ies"} paid · slips in Pay history`);
}

function reverseSalaryPayment(entryId) {
  const entry = state.finance.find((record) => record.id === entryId);
  if (!entry || !isSalaryEntry(entry) || entry.cancelled || !can("finance", "edit")) return;
  const reason = window.prompt(
    `Undo ${entry.party}'s salary for ${monthName(entry.payrollMonth, "long")}? The payment stays in the records marked as reversed, and the month shows as unpaid again.\n\nReason:`,
  );
  if (reason === null) return;
  if (!reason.trim()) {
    showToast("Enter a reason to undo a salary payment");
    return;
  }
  entry.cancelled = true;
  entry.cancelReason = reason.trim();
  entry.cancelledAt = new Date().toISOString();
  entry.cancelledBy = currentUser()?.name || "";
  addActivity(`Salary payment reversed: ${entry.party} · ${monthName(entry.payrollMonth, "long")} (${entry.cancelReason})`, "employees");
  saveState();
  render();
  showToast("Salary payment reversed");
}

function salaryEntries(employeeId = "all") {
  return state.finance
    .filter((entry) => isSalaryEntry(entry) && (employeeId === "all" || entry.employeeId === employeeId))
    .sort(
      (a, b) =>
        String(b.paidOn || b.date).localeCompare(String(a.paidOn || a.date)) || String(b.createdAt || "").localeCompare(String(a.createdAt || "")),
    );
}

function renderPayHistory() {
  const query = els.globalSearch.value.trim().toLowerCase();
  if (payHistoryFilters.employeeId !== "all" && !state.employees.some((employee) => employee.id === payHistoryFilters.employeeId)) {
    payHistoryFilters.employeeId = "all";
  }
  byId("payHistoryEmployeeFilter").innerHTML = [
    `<option value="all">All employees</option>`,
    ...state.employees.map(
      (employee) =>
        `<option value="${escapeHtml(employee.id)}" ${employee.id === payHistoryFilters.employeeId ? "selected" : ""}>${escapeHtml(employee.name)}</option>`,
    ),
  ].join("");
  const rows = salaryEntries(payHistoryFilters.employeeId).filter(
    (entry) => !query || [entry.party, entry.slipNumber, entry.notes, monthName(entry.payrollMonth, "long")].join(" ").toLowerCase().includes(query),
  );
  const live = rows.filter((entry) => !entry.cancelled);
  byId("payHistorySummary").innerHTML = `<strong>${formatNumber(live.length)}</strong> salary payments · Total <strong>${escapeHtml(
    formatMoneyTotals(sumByCurrency(live, (entry) => parseMoney(entry.amount))),
  )}</strong>`;
  const canUndo = can("finance", "edit");
  byId("payHistoryTable").innerHTML = `
    <thead>
      <tr>
        <th>Date paid</th>
        <th>Slip no.</th>
        <th>Employee</th>
        <th>For month</th>
        <th>Method</th>
        <th class="num">Net pay</th>
        <th>Status</th>
        <th></th>
      </tr>
    </thead>
    <tbody>
      ${
        rows.length
          ? rows
              .map((entry) => {
                const payment = entryPayments(entry)[0] || {};
                const adjustments = [entry.bonus ? `+${formatCurrency(entry.bonus, recordCurrency(entry))} bonus` : "", entry.deduction ? `−${formatCurrency(entry.deduction, recordCurrency(entry))} deducted` : ""]
                  .filter(Boolean)
                  .join(" · ");
                return `
              <tr class="${entry.cancelled ? "is-muted" : ""}">
                <td>${titleCell(formatDate(entry.paidOn || entry.date), payment.by || "")}</td>
                <td>${escapeHtml(entry.slipNumber || "—")}</td>
                <td>${escapeHtml(entry.party || "—")}</td>
                <td>${escapeHtml(monthName(entry.payrollMonth, "long"))}</td>
                <td>${escapeHtml(payment.method || "—")}</td>
                <td class="num">${titleCell(formatCurrency(entry.amount, recordCurrency(entry)), adjustments)}</td>
                <td>${entry.cancelled ? badge("Reversed", "red") : badge("Paid", "green")}</td>
                <td><div class="row-actions">
                  <button type="button" data-document="salary:${escapeHtml(entry.id)}">Salary slip</button>
                  ${!entry.cancelled && canUndo ? `<button type="button" class="warn" data-reverse-salary="${escapeHtml(entry.id)}">Undo</button>` : ""}
                </div></td>
              </tr>
            `;
              })
              .join("")
          : `<tr><td colspan="8"><div class="empty-state">No salary payments yet. Pay salaries from the Payroll tab.</div></td></tr>`
      }
    </tbody>
  `;
}

/* ----- Salary slips and statements ----- */

function salaryDocument(entry) {
  const employee = state.employees.find((record) => record.id === entry.employeeId);
  if (!entry.slipNumber) {
    entry.slipNumber = nextDocumentNumber("salary", "SAL");
    saveState();
  }
  const currency = recordCurrency(entry);
  const payment = entryPayments(entry)[0] || {};
  const base = entry.baseSalary !== undefined ? parseMoney(entry.baseSalary) : parseMoney(entry.amount);
  const items = [
    { name: `Basic salary · ${monthName(entry.payrollMonth, "long")}`, detail: employee?.role || "", quantity: 1, unitPrice: base, amount: base },
    ...(entry.bonus ? [{ name: "Bonus / overtime", detail: entry.notes || "", quantity: 1, unitPrice: entry.bonus, amount: entry.bonus }] : []),
    ...(entry.deduction ? [{ name: "Deductions", detail: entry.notes || "", quantity: 1, unitPrice: -entry.deduction, amount: -entry.deduction }] : []),
  ];
  return {
    record: entry,
    title: "Salary Slip",
    receiptTitle: "Salary Slip",
    subtitle: monthName(entry.payrollMonth, "long"),
    number: entry.slipNumber,
    numberLabel: "Slip no.",
    currency,
    meta: [
      ["Pay period", monthName(entry.payrollMonth, "long")],
      ["Date paid", formatDate(entry.paidOn || entry.date)],
      ["Method", payment.method || "—"],
      ["Paid by", payment.by || "—"],
    ],
    party: {
      label: "Employee",
      name: entry.party || employee?.name || "—",
      lines: [employee?.role, employee?.department, employee?.phone].filter((line) => !isBlank(line)),
    },
    showQuantity: false,
    priceLabel: "",
    items,
    totals: [
      ["Gross", round2(base + parseMoney(entry.bonus))],
      ...(entry.deduction ? [["Deductions", -parseMoney(entry.deduction)]] : []),
      ["Net pay", parseMoney(entry.amount), "grand"],
    ],
    amountInWords: entry.amount,
    payments: [],
    stamp: paymentStamp(parseMoney(entry.amount), parseMoney(entry.amount), entry.cancelled),
    cancelled: entry.cancelled ? `Reversed: ${entry.cancelReason || ""}` : "",
    notes: entry.notes,
    terms: "",
    signatures: ["Employer signature", "Employee signature"],
    thanks: "",
  };
}

function statementDocument(direction, key) {
  const group = findDueParty(direction, key);
  if (!group) return null;
  const incoming = direction === "in";
  const items = allocationOrder(group.items).map((item) => {
    const entry = item.kind === "transaction" ? state.finance.find((record) => record.id === item.id) : null;
    return {
      name: dueItemLabel(item),
      detail: [
        item.date ? formatDate(item.date) : "",
        entry ? `Total ${formatCurrency(entry.amount, group.currency)} · paid ${formatCurrency(entry.amountPaid || 0, group.currency)}` : "",
        item.dueDate ? `due ${formatDate(item.dueDate)}` : "",
        item.overdue ? "OVERDUE" : "",
      ]
        .filter(Boolean)
        .join(" · "),
      quantity: 1,
      unitPrice: item.amount,
      amount: item.amount,
    };
  });
  return {
    record: {},
    title: "Statement of Account",
    receiptTitle: "Statement",
    subtitle: incoming ? "Amount owed to us" : "Amount we owe",
    number: formatDate(todayIso()),
    numberLabel: "As of",
    currency: group.currency,
    meta: [
      ["Open items", String(group.items.length)],
      ["Overdue", String(group.overdue)],
    ],
    party: partyBlock(incoming ? "Customer" : "Supplier", group.party || "—", "", ""),
    showQuantity: false,
    priceLabel: "",
    items,
    totals: [["Balance due", group.total, "grand"]],
    amountInWords: group.total,
    payments: [],
    stamp: { label: group.overdue ? "Overdue" : "Open", tone: group.overdue ? "unpaid" : "partial" },
    cancelled: "",
    notes: "",
    terms: incoming ? "Please pay the balance due. Contact us if any item on this statement looks wrong." : "",
    signatures: ["Authorized signature", incoming ? "Customer signature" : "Supplier signature"],
    thanks: "",
  };
}

function groupReceiptDocument(receiptNumber) {
  const matches = [];
  state.finance.forEach((entry) => {
    (entry.payments || []).forEach((payment) => {
      if (payment.receiptNumber === receiptNumber) matches.push({ entry, payment });
    });
  });
  if (!matches.length || !canSeeFinanceEntry(matches[0].entry)) return null;
  if (matches.length === 1 && !isSalaryEntry(matches[0].entry)) return paymentDocument(matches[0].entry, matches[0].payment.id);
  if (isSalaryEntry(matches[0].entry)) return salaryDocument(matches[0].entry);
  const first = matches[0];
  const incoming = first.entry.type === "Income";
  const currency = recordCurrency(first.entry);
  const total = round2(matches.reduce((sum, match) => sum + parseMoney(match.payment.amount), 0));
  return {
    record: first.payment,
    title: incoming ? "Payment Receipt" : "Payment Voucher",
    receiptTitle: incoming ? "Payment Receipt" : "Payment Voucher",
    subtitle: `Applied to ${matches.length} items`,
    number: receiptNumber,
    numberLabel: "Receipt no.",
    currency,
    meta: [
      ["Payment date", formatDate(first.payment.date)],
      ["Method", first.payment.method || "—"],
      ...(first.payment.note ? [["Note", first.payment.note]] : []),
      ["Recorded by", first.payment.by || "—"],
    ],
    party: partyBlock(incoming ? "Received from" : "Paid to", first.entry.party, "", ""),
    showQuantity: false,
    priceLabel: "",
    items: matches.map(({ entry, payment }) => ({
      name: paymentSource(entry).label,
      detail: `Total ${formatCurrency(entry.amount, currency)} · balance now ${formatCurrency(outstandingOf(entry), currency)}`,
      quantity: 1,
      unitPrice: payment.amount,
      amount: payment.amount,
    })),
    totals: [[incoming ? "Amount received" : "Amount paid", total, "grand"]],
    amountInWords: total,
    payments: [],
    stamp: { label: incoming ? "Received" : "Paid", tone: "paid" },
    cancelled: "",
    notes: "",
    terms: "",
    signatures: incoming ? ["Received by", "Customer signature"] : ["Paid by", "Received by"],
    thanks: incoming ? "Thank you for your payment." : "",
  };
}

/* ---------- Invoices, receipts, and vouchers ----------
 *
 * Every recorded transaction can produce a printable document:
 * - Sale -> Sales invoice      - Purchase -> Purchase invoice (goods received)
 * - Payment -> Payment receipt - Other finance entry -> Receipt or payment voucher
 * - Stock adjustment -> Stock adjustment note
 * Each document can be shown as an A4 invoice or an 80 mm shop receipt. Documents
 * are built from the saved record, so a reprint always matches what was recorded.
 */

const DOCUMENT_FORMAT_KEY = "ledgerflow-document-format";
let documentFormat = (() => {
  try {
    return localStorage.getItem(DOCUMENT_FORMAT_KEY) === "receipt" ? "receipt" : "a4";
  } catch (error) {
    return "a4";
  }
})();
let activeDocument = null;

const CURRENCY_WORDS = {
  PKR: ["Pakistani Rupees", "Paisa", true],
  INR: ["Indian Rupees", "Paise", true],
  USD: ["US Dollars", "Cents"],
  EUR: ["Euros", "Cents"],
  GBP: ["Pounds Sterling", "Pence"],
  AED: ["UAE Dirhams", "Fils"],
  SAR: ["Saudi Riyals", "Halalas"],
  CNY: ["Chinese Yuan", "Fen"],
  JPY: ["Japanese Yen", "Sen"],
  CAD: ["Canadian Dollars", "Cents"],
  AUD: ["Australian Dollars", "Cents"],
};

function numberToWords(value, southAsian = false) {
  const ones = [
    "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve",
    "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen",
  ];
  const tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
  const belowThousand = (number) => {
    const parts = [];
    if (number >= 100) {
      parts.push(`${ones[Math.floor(number / 100)]} Hundred`);
      number %= 100;
    }
    if (number >= 20) parts.push(`${tens[Math.floor(number / 10)]}${number % 10 ? `-${ones[number % 10]}` : ""}`);
    else if (number > 0) parts.push(ones[number]);
    return parts.join(" ");
  };
  let number = Math.floor(Math.abs(value));
  if (!number) return "Zero";
  const parts = [];
  const scales = southAsian
    ? [
        [1e7, "Crore"],
        [1e5, "Lakh"],
        [1e3, "Thousand"],
      ]
    : [
        [1e12, "Trillion"],
        [1e9, "Billion"],
        [1e6, "Million"],
        [1e3, "Thousand"],
      ];
  scales.forEach(([size, name]) => {
    if (number >= size) {
      const count = Math.floor(number / size);
      parts.push(`${count >= 1000 ? numberToWords(count, southAsian) : belowThousand(count)} ${name}`);
      number %= size;
    }
  });
  if (number) parts.push(belowThousand(number));
  return parts.join(" ");
}

function amountInWords(amount, currency) {
  const [major, minor, southAsian] = CURRENCY_WORDS[currency] || [currency, "Cents", false];
  const absolute = Math.abs(round2(parseMoney(amount)));
  const whole = Math.floor(absolute);
  const fraction = Math.round((absolute - whole) * 100);
  return `${major} ${numberToWords(whole, southAsian)}${fraction ? ` and ${numberToWords(fraction)} ${minor}` : ""} Only`;
}

function plainAmount(value) {
  return new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(parseMoney(value));
}

function nextDocumentNumber(counter, prefix) {
  state.counters[counter] = (Number(state.counters[counter]) || 0) + 1;
  return `${prefix}-${String(state.counters[counter]).padStart(4, "0")}`;
}

function contactForParty(contactId, name) {
  if (contactId) {
    const byId = state.contacts.find((contact) => contact.id === contactId);
    if (byId) return byId;
  }
  const key = String(name || "").trim().toLowerCase();
  return key ? state.contacts.find((contact) => contact.name.trim().toLowerCase() === key) || null : null;
}

function partyBlock(label, name, phone, contactId) {
  const contact = contactForParty(contactId, name);
  return {
    label,
    name: name || "—",
    lines: [phone || contact?.phone, contact?.email, contact?.billingAddress].filter((line) => !isBlank(line)),
  };
}

function paymentStamp(total, paid, cancelled) {
  if (cancelled) return { label: "Cancelled", tone: "void" };
  if (total <= 0 || paid >= total - 0.005) return { label: "Paid", tone: "paid" };
  if (paid > 0) return { label: "Partially paid", tone: "partial" };
  return { label: "Unpaid", tone: "unpaid" };
}

function markDocumentPrinted(record) {
  record.printCount = (Number(record.printCount) || 0) + 1;
  record.lastPrintedAt = new Date().toISOString();
  saveState();
}

/* ----- Document specs built from saved records ----- */

function tradeDocument(kind, trade) {
  const config = TRADE[kind];
  const balance = tradeBalance(trade);
  const paid = parseMoney(trade.amountPaid);
  const sale = kind === "sale";
  return {
    record: trade,
    title: sale ? "Sales Invoice" : "Purchase Invoice",
    receiptTitle: sale ? (balance <= 0 && trade.status !== "Cancelled" ? "Sales Receipt" : "Sales Invoice") : "Goods Received",
    subtitle: sale ? "" : "Goods received from supplier",
    number: trade.number,
    numberLabel: sale ? "Invoice no." : "Purchase no.",
    currency: trade.currency,
    meta: [
      ["Date", formatDate(trade.date)],
      ...(balance > 0 && trade.dueDate ? [["Due date", formatDate(trade.dueDate)]] : []),
      ...(trade.paymentMethod ? [["Payment", trade.paymentMethod]] : []),
      ...(trade.reference ? [[sale ? "Reference" : "Supplier bill no.", trade.reference]] : []),
      [sale ? "Served by" : "Received by", trade.createdBy || "—"],
    ],
    party: partyBlock(sale ? "Bill to" : "Supplier", trade.partyName, trade.partyPhone, trade.contactId),
    showQuantity: true,
    priceLabel: sale ? "Unit price" : "Unit cost",
    items: trade.lines.map((line) => ({
      name: line.name,
      detail: line.custom ? (sale ? "Service / not from stock" : "Not added to stock") : [line.sku, line.addedToStock ? "New item" : ""].filter(Boolean).join(" · "),
      quantity: line.quantity,
      unitPrice: line.unitPrice,
      amount: line.lineTotal,
    })),
    totals: [
      ["Subtotal", trade.subtotal],
      ...(trade.discount ? [["Discount", -trade.discount]] : []),
      ["Total", trade.total, "grand"],
      [sale ? "Amount received" : "Amount paid", trade.status === "Cancelled" ? 0 : paid],
      ["Balance due", balance, balance > 0 ? "due" : ""],
    ],
    amountInWords: trade.total,
    payments: (trade.payments || []).map((payment) => ({ ...payment })),
    stamp: paymentStamp(trade.total, paid, trade.status === "Cancelled"),
    cancelled: trade.status === "Cancelled" ? `Cancelled on ${formatDate(String(trade.cancelledAt).slice(0, 10))}: ${trade.cancelReason || "no reason given"}` : "",
    notes: trade.notes,
    terms:
      balance > 0 && trade.status !== "Cancelled"
        ? `Balance of ${formatCurrency(balance, trade.currency)} is due${trade.dueDate ? ` by ${formatDate(trade.dueDate)}` : ""}.`
        : "",
    signatures: sale ? ["Authorized signature", "Customer signature"] : ["Received by", "Supplier signature"],
    thanks: sale ? "Thank you for your business." : "",
  };
}

function financeDocument(entry) {
  if (!entry.documentNumber) {
    entry.documentNumber = nextDocumentNumber("voucher", entry.type === "Income" ? "RV" : "PV");
    saveState();
  }
  const income = entry.type === "Income";
  const paid = isUnpaid(entry) ? parseMoney(entry.amountPaid) : parseMoney(entry.amount);
  const balance = outstandingOf(entry);
  return {
    record: entry,
    title: income ? "Receipt Voucher" : "Payment Voucher",
    receiptTitle: income ? "Receipt" : "Payment Voucher",
    subtitle: entry.category || "",
    number: entry.documentNumber,
    numberLabel: "Voucher no.",
    currency: recordCurrency(entry),
    meta: [
      ["Date", formatDate(entry.date)],
      ...(balance > 0 && entry.dueDate ? [["Due date", formatDate(entry.dueDate)]] : []),
      ["Category", entry.category || "—"],
      ...(entry.reference ? [["Reference", entry.reference]] : []),
    ],
    party: entry.party ? partyBlock(income ? "Received from" : "Paid to", entry.party, "", "") : null,
    showQuantity: false,
    priceLabel: "",
    items: [{ name: entry.description || entry.category || "Transaction", detail: entry.category, quantity: 1, unitPrice: entry.amount, amount: entry.amount }],
    totals: [
      ["Total", parseMoney(entry.amount), "grand"],
      [income ? "Amount received" : "Amount paid", paid],
      ["Balance due", balance, balance > 0 ? "due" : ""],
    ],
    amountInWords: entry.amount,
    payments: (entry.payments || []).map((payment) => ({ ...payment })),
    stamp: paymentStamp(parseMoney(entry.amount), paid, entry.cancelled),
    cancelled: entry.cancelled ? `Cancelled: ${entry.cancelReason || ""}` : "",
    notes: entry.notes,
    terms: "",
    signatures: income ? ["Received by", "Paid by"] : ["Approved by", "Received by"],
    thanks: income ? "Thank you." : "",
  };
}

function paymentDocument(entry, paymentId) {
  const payment = (entry.payments || []).find((item) => item.id === paymentId);
  if (!payment) return null;
  const trade = entry.sourceType ? findTrade(entry.sourceType, entry.sourceId) : null;
  if (!payment.receiptNumber) {
    payment.receiptNumber = nextReceiptNumber(entry.type === "Income" ? "in" : "out");
    const tradePayment = trade?.payments?.find((item) => item.id === paymentId);
    if (tradePayment) tradePayment.receiptNumber = payment.receiptNumber;
    saveState();
  }
  const income = entry.type === "Income";
  const paidUpTo = round2(
    (entry.payments || [])
      .slice(0, entry.payments.indexOf(payment) + 1)
      .reduce((sum, item) => sum + parseMoney(item.amount), 0),
  );
  const balanceAfter = round2(Math.max(0, parseMoney(entry.amount) - paidUpTo));
  const againstNumber = trade?.number || entry.documentNumber || entry.reference || "";
  return {
    record: payment,
    title: income ? "Payment Receipt" : "Payment Voucher",
    receiptTitle: income ? "Payment Receipt" : "Payment Voucher",
    subtitle: againstNumber ? `Against ${againstNumber}` : "",
    number: payment.receiptNumber,
    numberLabel: "Receipt no.",
    currency: recordCurrency(entry),
    meta: [
      ["Payment date", formatDate(payment.date)],
      ["Method", payment.method || "—"],
      ...(againstNumber ? [["Against", againstNumber]] : []),
      ...(payment.note ? [["Note", payment.note]] : []),
      ["Recorded by", payment.by || "—"],
    ],
    party: partyBlock(income ? "Received from" : "Paid to", trade?.partyName || entry.party, trade?.partyPhone, trade?.contactId),
    showQuantity: false,
    priceLabel: "",
    items: [
      {
        name: income ? `Payment received for ${entry.description || againstNumber}` : `Payment made for ${entry.description || againstNumber}`,
        detail: `Invoice total ${formatCurrency(entry.amount, recordCurrency(entry))}`,
        quantity: 1,
        unitPrice: payment.amount,
        amount: payment.amount,
      },
    ],
    totals: [
      [income ? "Amount received" : "Amount paid", payment.amount, "grand"],
      ["Paid to date", paidUpTo],
      ["Balance remaining", balanceAfter, balanceAfter > 0 ? "due" : ""],
    ],
    amountInWords: payment.amount,
    payments: [],
    stamp: { label: income ? "Received" : "Paid", tone: "paid" },
    cancelled: entry.cancelled ? "The invoice this payment belongs to was cancelled." : "",
    notes: "",
    terms: "",
    signatures: income ? ["Received by", "Customer signature"] : ["Paid by", "Received by"],
    thanks: income ? "Thank you for your payment." : "",
  };
}

function movementDocument(move) {
  if (!move.documentNumber) {
    move.documentNumber = nextDocumentNumber("adjustment", "ADJ");
    saveState();
  }
  const item = state.inventory.find((entry) => entry.id === move.itemId);
  const before = round3(parseMoney(move.balanceAfter) - parseMoney(move.change));
  return {
    record: move,
    title: move.type === "Adjustment" ? "Stock Adjustment Note" : "Stock Record",
    receiptTitle: move.type === "Adjustment" ? "Stock Adjustment" : "Stock Record",
    subtitle: move.type,
    number: move.documentNumber,
    numberLabel: "Note no.",
    currency: move.currency || orgCurrency(),
    meta: [
      ["Date", formatDate(move.date)],
      ["Item", move.itemName],
      ...(move.sku ? [[fieldLabel("inventory", "sku"), move.sku]] : []),
      ["Recorded by", move.by || "—"],
    ],
    party: null,
    showQuantity: false,
    priceLabel: "",
    stock: {
      rows: [
        ["Stock before", formatNumber(before)],
        ["Change", `${move.change > 0 ? "+" : "−"}${formatNumber(Math.abs(move.change))}`],
        ["Stock after", formatNumber(move.balanceAfter)],
        ["Stock now", item ? formatNumber(item.quantity) : "Item removed"],
      ],
      reason: move.reason || "—",
    },
    items: [],
    totals: [],
    amountInWords: null,
    payments: [],
    stamp: { label: move.change > 0 ? "Stock in" : "Stock out", tone: move.change > 0 ? "paid" : "partial" },
    cancelled: "",
    notes: "",
    terms: "",
    signatures: ["Recorded by", "Approved by"],
    thanks: "",
  };
}

/* ----- Rendering ----- */

function documentBusinessHeader(org) {
  const initials =
    String(org.name || "")
      .split(/\s+/)
      .filter(Boolean)
      .slice(0, 2)
      .map((word) => word.charAt(0).toUpperCase())
      .join("") || "LF";
  const contactLines = [
    org.businessAddress,
    [org.businessPhone ? `Phone: ${org.businessPhone}` : "", org.businessEmail ? `Email: ${org.businessEmail}` : ""].filter(Boolean).join("  ·  "),
    org.registrationNumber ? `Tax / Reg. No: ${org.registrationNumber}` : "",
  ].filter(Boolean);
  return { initials, contactLines };
}

function documentHtml(spec, format) {
  const org = state.organization;
  const { initials, contactLines } = documentBusinessHeader(org);
  const money = (value) => formatCurrency(value, spec.currency);
  const copyLabel = Number(spec.record.printCount) > 0 ? "Copy" : "Original";
  const generated = formatDateTime(new Date().toISOString());
  const words = spec.amountInWords !== null && spec.amountInWords !== undefined ? amountInWords(spec.amountInWords, spec.currency) : "";
  const escape = escapeHtml;

  if (format === "receipt") {
    const rule = `<div class="rule"></div>`;
    return `<!doctype html><html><head><meta charset="utf-8"><title>${escape(spec.receiptTitle)} ${escape(spec.number)}</title>
      <style>
        @page { size: 80mm auto; margin: 3mm; }
        * { box-sizing: border-box; }
        html, body { margin: 0; background: #fff; color: #000; }
        body { width: 74mm; margin: 0 auto; padding: 4mm 1mm 6mm; font: 12px/1.35 "Consolas", "Courier New", monospace; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
        .center { text-align: center; }
        .shop { font: 700 17px/1.2 Arial, sans-serif; letter-spacing: .02em; text-transform: uppercase; }
        .small { font-size: 10.5px; }
        .rule { border-top: 1px dashed #000; margin: 6px 0; }
        .rule.double { border-top: 3px double #000; }
        .title { font: 700 13px/1.3 Arial, sans-serif; letter-spacing: .12em; text-transform: uppercase; }
        .row { display: flex; justify-content: space-between; gap: 8px; }
        .row span:last-child { text-align: right; }
        .item { margin: 4px 0; }
        .item-name { font-weight: 700; }
        .total { font: 700 15px/1.4 "Consolas", monospace; }
        .stamp { display: inline-block; margin: 6px 0; padding: 2px 10px; border: 2px solid #000; font: 700 12px Arial, sans-serif; letter-spacing: .15em; text-transform: uppercase; }
        .void { color: #b42318; border-color: #b42318; }
        .words { font-size: 10.5px; font-style: italic; }
      </style></head><body>
      <div class="center">
        <div class="shop">${escape(org.name)}</div>
        ${contactLines.map((line) => `<div class="small">${escape(line)}</div>`).join("")}
      </div>
      ${rule}
      <div class="center title">${escape(spec.receiptTitle)}</div>
      ${spec.subtitle ? `<div class="center small">${escape(spec.subtitle)}</div>` : ""}
      ${rule}
      <div class="row"><span>${escape(spec.numberLabel)}</span><span>${escape(spec.number)}</span></div>
      ${spec.meta.map(([label, value]) => `<div class="row"><span>${escape(label)}</span><span>${escape(value)}</span></div>`).join("")}
      ${spec.party ? `<div class="row"><span>${escape(spec.party.label)}</span><span>${escape(spec.party.name)}</span></div>` : ""}
      ${spec.party?.lines?.length ? `<div class="row"><span></span><span>${escape(spec.party.lines[0])}</span></div>` : ""}
      ${rule}
      ${
        spec.stock
          ? spec.stock.rows.map(([label, value]) => `<div class="row"><span>${escape(label)}</span><span>${escape(value)}</span></div>`).join("") +
            `<div class="item"><div class="small">Reason: ${escape(spec.stock.reason)}</div></div>`
          : spec.items
              .map(
                (item) => `
          <div class="item">
            <div class="item-name">${escape(item.name)}</div>
            ${
              spec.showQuantity
                ? `<div class="row"><span>${escape(formatNumber(item.quantity))} x ${escape(plainAmount(item.unitPrice))}</span><span>${escape(plainAmount(item.amount))}</span></div>`
                : `<div class="row"><span class="small">${escape(item.detail || "")}</span><span>${escape(plainAmount(item.amount))}</span></div>`
            }
          </div>`,
              )
              .join("")
      }
      ${
        spec.totals.length
          ? `${rule}${spec.totals
              .map(([label, value, style]) =>
                style === "grand"
                  ? `<div class="rule double"></div><div class="row total"><span>${escape(label.toUpperCase())}</span><span>${escape(money(value))}</span></div><div class="rule double"></div>`
                  : `<div class="row"><span>${escape(label)}</span><span>${escape(value < 0 ? `-${plainAmount(-value)}` : plainAmount(value))}</span></div>`,
              )
              .join("")}`
          : ""
      }
      ${words ? `<div class="words">${escape(words)}</div>` : ""}
      <div class="center"><span class="stamp ${spec.stamp.tone === "void" ? "void" : ""}">${escape(spec.stamp.label)}</span></div>
      ${spec.cancelled ? `<div class="center small void">${escape(spec.cancelled)}</div>` : ""}
      ${spec.notes ? `${rule}<div class="small">Note: ${escape(spec.notes)}</div>` : ""}
      ${spec.terms ? `<div class="small">${escape(spec.terms)}</div>` : ""}
      ${rule}
      ${spec.thanks ? `<div class="center"><strong>${escape(spec.thanks)}</strong></div>` : ""}
      <div class="center small">${escape(copyLabel)} · Printed ${escape(generated)}</div>
      <div class="center small">Powered by LedgerFlow</div>
    </body></html>`;
  }

  const stampClass = { paid: "paid", partial: "partial", unpaid: "unpaid", void: "void" }[spec.stamp.tone] || "paid";
  return `<!doctype html><html><head><meta charset="utf-8"><title>${escape(spec.title)} ${escape(spec.number)}</title>
    <style>
      @page { size: A4; margin: 12mm; }
      * { box-sizing: border-box; }
      html, body { margin: 0; background: #fff; }
      body { color: #101828; font: 12.5px/1.5 Inter, "Segoe UI", Arial, sans-serif; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .page { position: relative; max-width: 186mm; margin: 0 auto; padding: 8mm 0; }
      .top { display: flex; justify-content: space-between; gap: 24px; padding-bottom: 16px; border-bottom: 3px solid #0e7c6b; }
      .brand { display: flex; gap: 14px; }
      .mark { width: 52px; height: 52px; flex: none; display: grid; place-items: center; border-radius: 10px; background: #0e7c6b; color: #fff; font: 700 18px Arial, sans-serif; }
      .biz-name { margin: 0; font-size: 22px; font-weight: 700; letter-spacing: -.01em; }
      .biz-type { color: #0e7c6b; font-size: 11px; font-weight: 600; letter-spacing: .08em; text-transform: uppercase; }
      .biz-line { color: #475467; font-size: 11.5px; }
      .doc { text-align: right; }
      .doc-title { margin: 0; font-size: 24px; font-weight: 800; letter-spacing: .12em; text-transform: uppercase; color: #101828; }
      .doc-sub { color: #667085; font-size: 11.5px; }
      .doc-number { margin-top: 6px; font-size: 13px; }
      .doc-number strong { font-size: 15px; }
      .stamp { display: inline-block; margin-top: 10px; padding: 4px 14px; border: 2px solid currentColor; border-radius: 6px; font-weight: 800; font-size: 12px; letter-spacing: .18em; text-transform: uppercase; transform: rotate(-4deg); }
      .stamp.paid { color: #067647; } .stamp.partial { color: #b54708; } .stamp.unpaid { color: #b42318; } .stamp.void { color: #b42318; }
      .meta { display: grid; grid-template-columns: repeat(auto-fit, minmax(120px, 1fr)); gap: 0; margin: 18px 0; border: 1px solid #e4e7ec; border-radius: 8px; overflow: hidden; }
      .meta div { padding: 8px 12px; border-right: 1px solid #e4e7ec; }
      .meta div:last-child { border-right: 0; }
      .label { display: block; color: #667085; font-size: 10px; font-weight: 600; letter-spacing: .08em; text-transform: uppercase; }
      .value { font-weight: 600; }
      .parties { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 18px; }
      .party { padding: 12px 14px; background: #f9fafb; border: 1px solid #e4e7ec; border-radius: 8px; }
      .party-name { margin: 2px 0; font-size: 15px; font-weight: 700; }
      table { width: 100%; border-collapse: collapse; }
      .items th { padding: 9px 10px; background: #101828; color: #fff; font-size: 10.5px; font-weight: 600; letter-spacing: .06em; text-align: left; text-transform: uppercase; }
      .items td { padding: 10px; border-bottom: 1px solid #e4e7ec; vertical-align: top; }
      .items tr:nth-child(even) td { background: #fcfcfd; }
      .num { text-align: right !important; white-space: nowrap; font-variant-numeric: tabular-nums; }
      .item-detail { color: #667085; font-size: 11px; }
      .summary { display: grid; grid-template-columns: 1fr 260px; gap: 20px; margin-top: 16px; align-items: start; }
      .words { padding: 10px 12px; background: #f0f9f7; border-left: 3px solid #0e7c6b; border-radius: 4px; font-size: 12px; }
      .totals td { padding: 6px 10px; }
      .totals td:first-child { color: #475467; }
      .totals tr.grand td { padding: 10px; background: #101828; color: #fff; font-size: 15px; font-weight: 700; }
      .totals tr.due td { color: #b42318; font-weight: 700; }
      .section-title { margin: 22px 0 8px; font-size: 11px; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; color: #475467; }
      .payments td, .payments th { padding: 7px 10px; border-bottom: 1px solid #e4e7ec; text-align: left; font-size: 11.5px; }
      .payments th { color: #667085; font-size: 10px; letter-spacing: .06em; text-transform: uppercase; }
      .stock-table td { padding: 10px 12px; border: 1px solid #e4e7ec; font-size: 14px; }
      .stock-table td:first-child { width: 40%; color: #475467; background: #f9fafb; }
      .notes { padding: 10px 12px; border: 1px dashed #d0d5dd; border-radius: 6px; color: #344054; }
      .cancelled-note { margin: 12px 0; padding: 10px 12px; border: 1px solid #fecdca; background: #fef3f2; color: #b42318; border-radius: 6px; font-weight: 600; }
      .watermark { position: absolute; top: 42%; left: 50%; transform: translate(-50%, -50%) rotate(-24deg); color: rgba(180, 35, 24, .12); font-size: 110px; font-weight: 900; letter-spacing: .1em; pointer-events: none; }
      .signatures { display: grid; grid-template-columns: 1fr 1fr; gap: 60px; margin-top: 56px; }
      .signature { padding-top: 8px; border-top: 1px solid #101828; color: #475467; font-size: 11px; text-align: center; }
      .footer { display: flex; justify-content: space-between; gap: 16px; margin-top: 28px; padding-top: 10px; border-top: 1px solid #e4e7ec; color: #667085; font-size: 10.5px; }
      .thanks { color: #0e7c6b; font-weight: 700; font-size: 13px; }
    </style></head><body>
    <div class="page">
      ${spec.cancelled && spec.stamp.tone === "void" ? `<div class="watermark">CANCELLED</div>` : ""}
      <div class="top">
        <div class="brand">
          <div class="mark">${escape(initials)}</div>
          <div>
            <h1 class="biz-name">${escape(org.name)}</h1>
            <div class="biz-type">${escape(org.type || "")}</div>
            ${contactLines.map((line) => `<div class="biz-line">${escape(line)}</div>`).join("")}
          </div>
        </div>
        <div class="doc">
          <h2 class="doc-title">${escape(spec.title)}</h2>
          ${spec.subtitle ? `<div class="doc-sub">${escape(spec.subtitle)}</div>` : ""}
          <div class="doc-number">${escape(spec.numberLabel)} <strong>${escape(spec.number)}</strong></div>
          <div class="stamp ${stampClass}">${escape(spec.stamp.label)}</div>
        </div>
      </div>

      <div class="meta">
        ${spec.meta.map(([label, value]) => `<div><span class="label">${escape(label)}</span><span class="value">${escape(value)}</span></div>`).join("")}
      </div>

      ${spec.cancelled ? `<div class="cancelled-note">${escape(spec.cancelled)}</div>` : ""}

      ${
        spec.party
          ? `<div class="parties">
              <div class="party">
                <span class="label">${escape(spec.party.label)}</span>
                <div class="party-name">${escape(spec.party.name)}</div>
                ${spec.party.lines.map((line) => `<div class="biz-line">${escape(line)}</div>`).join("")}
              </div>
              <div class="party">
                <span class="label">Issued by</span>
                <div class="party-name">${escape(org.name)}</div>
                ${org.businessPhone ? `<div class="biz-line">${escape(org.businessPhone)}</div>` : ""}
                ${org.registrationNumber ? `<div class="biz-line">${escape(org.registrationNumber)}</div>` : ""}
              </div>
            </div>`
          : ""
      }

      ${
        spec.stock
          ? `<table class="stock-table">${spec.stock.rows
              .map(([label, value]) => `<tr><td>${escape(label)}</td><td><strong>${escape(value)}</strong></td></tr>`)
              .join("")}<tr><td>Reason</td><td>${escape(spec.stock.reason)}</td></tr></table>`
          : `<table class="items">
              <thead><tr>
                <th style="width:36px">#</th>
                <th>Description</th>
                ${spec.showQuantity ? `<th class="num">Qty</th><th class="num">${escape(spec.priceLabel)}</th>` : ""}
                <th class="num">Amount (${escape(spec.currency)})</th>
              </tr></thead>
              <tbody>
                ${spec.items
                  .map(
                    (item, index) => `<tr>
                      <td>${index + 1}</td>
                      <td><strong>${escape(item.name)}</strong>${item.detail ? `<div class="item-detail">${escape(item.detail)}</div>` : ""}</td>
                      ${spec.showQuantity ? `<td class="num">${escape(formatNumber(item.quantity))}</td><td class="num">${escape(plainAmount(item.unitPrice))}</td>` : ""}
                      <td class="num">${escape(plainAmount(item.amount))}</td>
                    </tr>`,
                  )
                  .join("")}
              </tbody>
            </table>`
      }

      ${
        spec.totals.length
          ? `<div class="summary">
              <div>
                ${words ? `<span class="label">Amount in words</span><div class="words">${escape(words)}</div>` : ""}
                ${spec.terms ? `<p class="biz-line">${escape(spec.terms)}</p>` : ""}
              </div>
              <table class="totals">
                ${spec.totals
                  .map(
                    ([label, value, style]) =>
                      `<tr class="${style || ""}"><td>${escape(label)}</td><td class="num">${escape(
                        value < 0 ? `− ${money(-value)}` : money(value),
                      )}</td></tr>`,
                  )
                  .join("")}
              </table>
            </div>`
          : ""
      }

      ${
        spec.payments.length
          ? `<div class="section-title">Payment history</div>
             <table class="payments"><thead><tr><th>Date</th><th>Method</th><th>Receipt / note</th><th class="num">Amount</th></tr></thead><tbody>
             ${spec.payments
               .map(
                 (payment) => `<tr><td>${escape(formatDate(payment.date))}</td><td>${escape(payment.method || "—")}</td><td>${escape(
                   [payment.receiptNumber, payment.note].filter(Boolean).join(" · ") || "—",
                 )}</td><td class="num">${escape(money(payment.amount))}</td></tr>`,
               )
               .join("")}
             </tbody></table>`
          : ""
      }

      ${spec.notes ? `<div class="section-title">Notes</div><div class="notes">${escape(spec.notes)}</div>` : ""}

      <div class="signatures">
        ${spec.signatures.map((label) => `<div class="signature">${escape(label)}</div>`).join("")}
      </div>

      <div class="footer">
        <span>${spec.thanks ? `<span class="thanks">${escape(spec.thanks)}</span><br>` : ""}This is a computer-generated document from ${escape(org.name)}'s records.</span>
        <span class="num">${escape(copyLabel)} · Printed ${escape(generated)}<br>Powered by LedgerFlow</span>
      </div>
    </div>
  </body></html>`;
}

function openDocument(spec) {
  if (!spec) return;
  activeDocument = spec;
  editing = { kind: "view" };
  setModalWide(true);
  els.modalKicker.textContent = spec.numberLabel.replace(/ no\.$/, "");
  els.modalTitle.textContent = `${spec.title} ${spec.number}`;
  els.recordForm.innerHTML = `
    <div class="document-toolbar">
      <div class="segmented" role="group" aria-label="Document style">
        <button type="button" data-document-format="a4" class="${documentFormat === "a4" ? "is-active" : ""}">A4 invoice</button>
        <button type="button" data-document-format="receipt" class="${documentFormat === "receipt" ? "is-active" : ""}">Shop receipt (80 mm)</button>
      </div>
      <span class="muted-note">${
        Number(spec.record.printCount) > 0
          ? `Printed ${spec.record.printCount} time${spec.record.printCount === 1 ? "" : "s"} before. New prints are marked "Copy".`
          : `The first print is marked "Original".`
      }</span>
    </div>
    <div class="document-preview ${documentFormat}">
      <iframe id="documentFrame" title="${escapeHtml(spec.title)} preview"></iframe>
    </div>
    <div class="form-actions">
      <button class="button ghost" id="cancelFormBtn" type="button">Close</button>
      <button class="button primary" type="button" data-print-document>
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 9V3h10v6M7 17H4v-7h16v7h-3M7 14h10v7H7z" /></svg>
        Print
      </button>
    </div>
  `;
  byId("documentFrame").srcdoc = documentHtml(spec, documentFormat);
  els.modalBackdrop.hidden = false;
}

function setDocumentFormat(format) {
  documentFormat = format === "receipt" ? "receipt" : "a4";
  try {
    localStorage.setItem(DOCUMENT_FORMAT_KEY, documentFormat);
  } catch (error) {
    console.warn("Could not remember document format", error);
  }
  if (activeDocument) openDocument(activeDocument);
}

function printActiveDocument() {
  const frame = byId("documentFrame");
  if (!frame || !activeDocument) return;
  markDocumentPrinted(activeDocument.record);
  frame.contentWindow.focus();
  frame.contentWindow.print();
  // The next print of this record is a copy.
  window.setTimeout(() => {
    if (activeDocument && byId("documentFrame")) openDocument(activeDocument);
  }, 300);
}

function openDocumentFor(reference) {
  if (reference.startsWith("statement:")) {
    const rest = reference.slice("statement:".length);
    const direction = rest.slice(0, rest.indexOf(":"));
    if (can("finance", "view") || can("sales", "view") || can("purchases", "view")) openDocument(statementDocument(direction, rest.slice(direction.length + 1)));
    return;
  }
  const [kind, id, extra] = reference.split(":");
  if (kind === "receipt") {
    openDocument(groupReceiptDocument(id));
    return;
  }
  if (kind === "salary") {
    const entry = state.finance.find((record) => record.id === id);
    if (entry && isSalaryEntry(entry) && canSeeFinanceEntry(entry)) openDocument(salaryDocument(entry));
    return;
  }
  if (kind === "sale" || kind === "purchase") {
    const trade = findTrade(kind, id);
    if (trade && can(TRADE[kind].area, "view")) openDocument(tradeDocument(kind, trade));
    return;
  }
  if (kind === "finance") {
    const entry = state.finance.find((record) => record.id === id);
    if (!entry) return;
    if (entry.sourceType) {
      openDocumentFor(`${entry.sourceType}:${entry.sourceId}`);
      return;
    }
    if (isSalaryEntry(entry)) {
      openDocumentFor(`salary:${entry.id}`);
      return;
    }
    if (can("finance", "view")) openDocument(financeDocument(entry));
    return;
  }
  if (kind === "payment") {
    const entry = state.finance.find((record) => record.id === id);
    if (entry && (can("finance", "view") || (entry.sourceType && can(TRADE[entry.sourceType].area, "view")))) {
      openDocument(paymentDocument(entry, extra));
    }
    return;
  }
  if (kind === "movement") {
    const move = state.stockMovements.find((record) => record.id === id);
    if (move && can("inventory", "view")) openDocument(movementDocument(move));
  }
}

/* ----- Sales and purchases lists ----- */

function tradeItemsSummary(trade) {
  if (!trade.lines.length) return "—";
  const [first, ...rest] = trade.lines;
  return `${first.name} × ${formatNumber(first.quantity)}${rest.length ? ` +${rest.length} more` : ""}`;
}

function inPeriod(date, period) {
  const value = String(date || "");
  const month = todayIso().slice(0, 7);
  if (period === "this-month") return value.startsWith(month);
  if (period === "last-month") return value.startsWith(shiftMonth(month, -1));
  if (period === "this-year") return value.startsWith(month.slice(0, 4));
  return true;
}

function renderTrades(kind) {
  const config = TRADE[kind];
  const area = config.area;
  const tab = activeTab(area);
  const query = els.globalSearch.value.trim().toLowerCase();
  const period = tradeFilters[area];
  byId(`${area}PeriodFilter`).value = period;
  if (tab === "payments") {
    renderTradePayments(kind);
    return;
  }

  const rows = state[config.list]
    .filter((trade) => {
      if (tab === "cancelled" ? trade.status !== "Cancelled" : trade.status === "Cancelled") return false;
      if (tab === "unpaid" && tradeBalance(trade) <= 0) return false;
      if (!inPeriod(trade.date, period)) return false;
      if (!query) return true;
      const text = [trade.number, trade.partyName, trade.partyPhone, trade.reference, trade.notes, ...trade.lines.map((line) => `${line.name} ${line.sku}`)]
        .join(" ")
        .toLowerCase();
      return text.includes(query);
    })
    .sort((a, b) => String(b.date).localeCompare(String(a.date)) || String(b.number).localeCompare(String(a.number)));

  byId(`${area}Head`).innerHTML = `
    <tr>
      <th>${config.noun} no.</th>
      <th>${config.party}</th>
      <th>Items</th>
      <th class="num">Total</th>
      <th class="num">Paid</th>
      <th class="num">Balance</th>
      <th>Status</th>
      <th></th>
    </tr>
  `;
  byId(`${area}Table`).innerHTML = rows.length
    ? rows
        .map((trade) => {
          const status = tradeStatus(trade);
          const balance = tradeBalance(trade);
          const entry = state.finance.find((record) => record.id === trade.financeId);
          const actions = [
            `<button type="button" data-view-trade="${kind}:${escapeHtml(trade.id)}">View</button>`,
            `<button type="button" data-document="${kind}:${escapeHtml(trade.id)}">Invoice</button>`,
            balance > 0 && canTakePayment(entry)
              ? `<button type="button" class="accent" data-mark-paid="${escapeHtml(entry.id)}">${kind === "sale" ? "Receive" : "Pay"}</button>`
              : "",
          ].join("");
          return `
          <tr class="is-clickable" data-view-trade-row="${kind}:${escapeHtml(trade.id)}">
            <td>${titleCell(trade.number, formatDate(trade.date))}</td>
            <td>${titleCell(trade.partyName || "—", trade.partyPhone || (trade.walkIn ? "No details recorded" : ""))}</td>
            <td>${titleCell(tradeItemsSummary(trade), `${formatNumber(trade.lines.reduce((sum, line) => sum + line.quantity, 0))} units`)}</td>
            <td class="num"><strong>${escapeHtml(formatCurrency(trade.total, trade.currency))}</strong></td>
            <td class="num">${escapeHtml(formatCurrency(trade.status === "Cancelled" ? 0 : trade.amountPaid, trade.currency))}</td>
            <td class="num"><span class="${balance > 0 ? "negative-text" : ""}">${escapeHtml(formatCurrency(balance, trade.currency))}</span></td>
            <td>${badge(status.label, status.tone)}</td>
            <td><div class="row-actions">${actions}</div></td>
          </tr>
        `;
        })
        .join("")
    : `<tr><td colspan="8"><div class="empty-state">${
        state[config.list].length
          ? "Nothing matches this view."
          : kind === "sale"
            ? "No sales recorded yet. Use New sale each time you sell something."
            : "No purchases recorded yet. Use New purchase whenever stock arrives from a supplier."
      }</div></td></tr>`;

  const live = rows.filter((trade) => trade.status !== "Cancelled");
  byId(`${area}TableSummary`).innerHTML = `<strong>${formatNumber(rows.length)}</strong> ${
    kind === "sale" ? "sales" : "purchases"
  }${
    live.length
      ? ` · Total <strong>${escapeHtml(formatMoneyTotals(sumByCurrency(live, (trade) => trade.total)))}</strong> · Unpaid <strong>${escapeHtml(
          formatMoneyTotals(sumByCurrency(live, tradeBalance)),
        )}</strong>`
      : ""
  }`;
}

function renderStockHistory() {
  const query = els.globalSearch.value.trim().toLowerCase();
  const type = tradeFilters.stockType;
  byId("stockTypeFilter").value = type;
  const moves = state.stockMovements.filter((move) => {
    if (type === "sales" && !["Sale", "Sale cancelled"].includes(move.type)) return false;
    if (type === "purchases" && !["Purchase", "Purchase cancelled"].includes(move.type)) return false;
    if (type === "adjustments" && move.type !== "Adjustment") return false;
    if (type === "opening" && !String(move.type).startsWith("Opening") && move.type !== "Import update") return false;
    if (!query) return true;
    return [move.itemName, move.sku, move.reference, move.party, move.reason, move.by].join(" ").toLowerCase().includes(query);
  });
  byId("stockHistoryWrap").innerHTML = movementTableHtml(moves);
  byId("stockHistorySummary").innerHTML = `<strong>${formatNumber(moves.length)}</strong> stock changes`;
}

/* ----- Sales and stock reports ----- */

function renderTradeReports() {
  const panel = byId("tradeReportPanel");
  const show = can("sales", "view") || can("purchases", "view");
  panel.hidden = !show;
  if (!show) return;
  const period = tradeFilters.reportPeriod;
  byId("reportPeriodSelect").value = period;
  const sales = can("sales", "view") ? state.sales.filter((trade) => trade.status !== "Cancelled" && inPeriod(trade.date, period)) : [];
  const purchases = can("purchases", "view")
    ? state.purchases.filter((trade) => trade.status !== "Cancelled" && inPeriod(trade.date, period))
    : [];
  const currencies = [...new Set([...sales, ...purchases].map((trade) => trade.currency))];
  if (!currencies.includes(tradeFilters.reportCurrency)) {
    tradeFilters.reportCurrency = currencies.includes(orgCurrency()) ? orgCurrency() : currencies[0] || orgCurrency();
  }
  const currency = tradeFilters.reportCurrency;
  byId("reportCurrencyControl").hidden = currencies.length < 2;
  byId("reportCurrencySelect").innerHTML = currencies
    .map((code) => `<option value="${code}" ${code === currency ? "selected" : ""}>${code}</option>`)
    .join("");
  const money = (value) => formatCurrency(value, currency);
  const salesIn = sales.filter((trade) => trade.currency === currency);
  const purchasesIn = purchases.filter((trade) => trade.currency === currency);
  const sensitive = canAccess("sensitiveNumbers");

  const itemStats = {};
  let costOfGoods = 0;
  let linesWithoutCost = 0;
  salesIn.forEach((trade) => {
    const share = trade.subtotal ? trade.total / trade.subtotal : 1;
    trade.lines.forEach((line) => {
      const key = line.itemId || `custom:${line.name}`;
      const stats = (itemStats[key] ||= { name: line.name, quantity: 0, revenue: 0, cost: 0, costKnown: true });
      const revenue = line.lineTotal * share;
      stats.quantity += line.quantity;
      stats.revenue += revenue;
      if (!line.custom) {
        if (line.unitCost === "") {
          stats.costKnown = false;
          linesWithoutCost += 1;
        } else {
          stats.cost += line.quantity * parseMoney(line.unitCost);
          costOfGoods += line.quantity * parseMoney(line.unitCost);
        }
      }
    });
  });
  const revenue = round2(salesIn.reduce((sum, trade) => sum + trade.total, 0));
  const unitsSold = salesIn.reduce((sum, trade) => sum + trade.lines.reduce((lines, line) => lines + line.quantity, 0), 0);
  const kpis = [];
  if (can("sales", "view")) {
    kpis.push(
      kpiCard({ label: "Sales", value: formatNumber(salesIn.length), icon: "income", tone: "income", note: `${formatNumber(unitsSold)} units sold` }),
      kpiCard({
        label: "Sales revenue",
        value: money(revenue),
        icon: "wallet",
        note: salesIn.length ? `Average sale ${money(revenue / salesIn.length)}` : "",
      }),
    );
    if (sensitive) {
      kpis.push(
        kpiCard({
          label: "Gross profit",
          value: money(round2(revenue - costOfGoods)),
          icon: "profit",
          tone: "profit",
          negative: revenue - costOfGoods < 0,
          note: linesWithoutCost ? `${linesWithoutCost} sold lines had no cost recorded` : revenue ? `${Math.round(((revenue - costOfGoods) / revenue) * 100)}% margin` : "",
        }),
      );
    }
    kpis.push(
      kpiCard({
        label: "Still owed by customers",
        value: money(salesIn.reduce((sum, trade) => sum + tradeBalance(trade), 0)),
        icon: "inbox",
        tone: "alert",
      }),
    );
  }
  if (can("purchases", "view")) {
    kpis.push(
      kpiCard({
        label: "Purchases",
        value: money(purchasesIn.reduce((sum, trade) => sum + trade.total, 0)),
        icon: "expense",
        tone: "expense",
        note: `${purchasesIn.length} purchases · ${money(purchasesIn.reduce((sum, trade) => sum + tradeBalance(trade), 0))} unpaid`,
      }),
    );
  }
  byId("tradeReportKpis").innerHTML = kpis.join("");

  const tableOrEmpty = (headers, rows, empty) =>
    rows.length
      ? `<table class="data-table compact"><thead><tr>${headers
          .map((header, index) => `<th class="${index ? "num" : ""}">${header}</th>`)
          .join("")}</tr></thead><tbody>${rows
          .map((row) => `<tr>${row.map((cell, index) => `<td class="${index ? "num" : ""}">${cell}</td>`).join("")}</tr>`)
          .join("")}</tbody></table>`
      : `<div class="empty-state">${empty}</div>`;

  byId("topItemsReport").innerHTML = tableOrEmpty(
    ["Item", "Qty sold", "Revenue", ...(sensitive ? ["Gross profit"] : [])],
    Object.values(itemStats)
      .sort((a, b) => b.revenue - a.revenue)
      .slice(0, 10)
      .map((stats) => [
        escapeHtml(stats.name),
        formatNumber(stats.quantity),
        escapeHtml(money(stats.revenue)),
        ...(sensitive ? [stats.costKnown ? escapeHtml(money(stats.revenue - stats.cost)) : "No cost"] : []),
      ]),
    "No sales in this period.",
  );

  const customerStats = {};
  salesIn.forEach((trade) => {
    const key = trade.contactId || (trade.walkIn ? "walk-in" : trade.partyName.toLowerCase());
    const stats = (customerStats[key] ||= { name: trade.walkIn ? "Walk-in customers" : trade.partyName, count: 0, total: 0, owed: 0 });
    stats.count += 1;
    stats.total += trade.total;
    stats.owed += tradeBalance(trade);
  });
  byId("topCustomersReport").innerHTML = tableOrEmpty(
    ["Customer", "Sales", "Total", "Still owed"],
    Object.values(customerStats)
      .sort((a, b) => b.total - a.total)
      .slice(0, 10)
      .map((stats) => [escapeHtml(stats.name), formatNumber(stats.count), escapeHtml(money(stats.total)), escapeHtml(money(stats.owed))]),
    "No sales in this period.",
  );

  const supplierStats = {};
  purchasesIn.forEach((trade) => {
    const key = trade.contactId || trade.partyName.toLowerCase();
    const stats = (supplierStats[key] ||= { name: trade.partyName, count: 0, total: 0, owed: 0 });
    stats.count += 1;
    stats.total += trade.total;
    stats.owed += tradeBalance(trade);
  });
  byId("suppliersReport").innerHTML = tableOrEmpty(
    ["Supplier", "Purchases", "Total", "You owe"],
    Object.values(supplierStats)
      .sort((a, b) => b.total - a.total)
      .map((stats) => [escapeHtml(stats.name), formatNumber(stats.count), escapeHtml(money(stats.total)), escapeHtml(money(stats.owed))]),
    "No purchases in this period.",
  );

  const movementStats = {};
  state.stockMovements
    .filter((move) => inPeriod(move.date, period))
    .forEach((move) => {
      const stats = (movementStats[move.type] ||= { in: 0, out: 0, count: 0 });
      stats.count += 1;
      if (move.change > 0) stats.in += move.change;
      else stats.out += Math.abs(move.change);
    });
  byId("movementReport").innerHTML = tableOrEmpty(
    ["Change type", "Entries", "Units in", "Units out"],
    Object.entries(movementStats).map(([type, stats]) => [movementBadge(type), formatNumber(stats.count), formatNumber(stats.in), formatNumber(stats.out)]),
    "No stock changes in this period.",
  );
}

/* ----- Sample transactions ----- */

function seedSampleTrades() {
  const items = state.inventory.filter((item) => parseMoney(item.quantity) > 0 && !isBlank(item.sellPrice));
  if (!items.length) return 0;
  const currency = recordCurrency(items[0]);
  const usable = items.filter((item) => recordCurrency(item) === currency);
  const customers = state.contacts.filter((contact) => contact.type === "Customer");
  const suppliers = state.contacts.filter((contact) => contact.type === "Supplier");
  const pick = (index) => usable[index % usable.length];
  const due = (days) => daysAgoIso(-days);
  let created = 0;
  const make = (kind, draft) => {
    const result = createTrade(kind, {
      contactId: "",
      partyName: "",
      partyPhone: "",
      saveContact: false,
      currency,
      discount: "",
      payment: "paid",
      amountPaid: "",
      method: "Cash",
      dueDate: "",
      reference: "",
      notes: "",
      ...draft,
      lines: draft.lines
        .filter(([item]) => item)
        .map(([item, quantity, factor = 1]) => ({
          key: makeId("line"),
          itemId: item.id,
          custom: false,
          description: "",
          quantity: kind === "sale" ? Math.max(1, Math.min(quantity, Math.floor(parseMoney(item.quantity)))) : quantity,
          unitPrice: round2(parseMoney(kind === "sale" ? item.sellPrice : item.unitCost || item.sellPrice * 0.8) * factor),
          priceTouched: true,
        })),
    });
    if (!result.error) created += 1;
  };

  const restockQuantity = (item) => (parseMoney(item.quantity) <= 2 ? 1 : 20);
  const lowItem = usable.find((item) => isLowStock(item)) || pick(1);
  if (suppliers[0]) {
    make("purchase", {
      date: daysAgoIso(12),
      contactId: suppliers[0].id,
      lines: [[lowItem, restockQuantity(lowItem)]],
      payment: "paid",
      method: "Bank transfer",
      reference: "INV-40912",
    });
    make("purchase", {
      date: daysAgoIso(4),
      contactId: (suppliers[1] || suppliers[0]).id,
      lines: [[pick(2), restockQuantity(pick(2))]],
      payment: "unpaid",
      dueDate: due(20),
      reference: "INV-41177",
    });
  }
  make("sale", { date: daysAgoIso(14), contactId: customers[0]?.id || "", lines: [[pick(1), 2]], payment: "paid", method: "Card" });
  make("sale", { date: daysAgoIso(9), partyName: "Ali (walk-in)", lines: [[pick(4), 1, 0.95]], payment: "paid", method: "Mobile wallet" });
  make("sale", {
    date: daysAgoIso(6),
    contactId: (customers[1] || customers[0])?.id || "",
    partyName: customers.length ? "" : "Credit customer",
    lines: [[pick(3), 1], [pick(0), 1]],
    payment: "unpaid",
    dueDate: due(-2),
  });
  make("sale", {
    date: daysAgoIso(3),
    contactId: customers[0]?.id || "",
    partyName: customers.length ? "" : "Regular customer",
    lines: [[pick(2), 1]],
    payment: "partial",
    amountPaid: round2(parseMoney(pick(2).sellPrice) / 2),
    method: "Bank transfer",
    dueDate: due(10),
  });
  make("sale", { date: daysAgoIso(1), lines: [[pick(0), 2], [pick(1), 1]], payment: "paid", method: "Cash" });
  return created;
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
    lines.push(["Overdue payments", `${alerts.unpaid} transaction${alerts.unpaid === 1 ? " is" : "s are"} past the due date.`]);
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

function loadSampleBusiness(type, { skipConfirm = false } = {}) {
  const build = globalThis.LEDGERFLOW_SAMPLES?.[type];
  if (!build || !currentUser()?.isOwner) return;
  const confirmText = totalRecordCount()
    ? `Replace ALL current records with the ${type} sample? Module names and fields will be arranged for a ${type}. Your login and roles stay.`
    : `Load the ${type} sample? Module names and fields will be arranged for a ${type}.`;
  if (!skipConfirm && !window.confirm(confirmText)) return;

  resetTransactionRecords();
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

  // Opening stock a month back, then realistic sales and purchases that move it.
  state.inventory.forEach((item) => logOpeningStock(item, "Opening stock", `Sample: ${type}`, daysAgoIso(30)));
  count += seedSampleTrades();

  state.organization.sampleLoaded = type;
  if (state.demoMode) state.organization.name = `${type} Demo`;
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

function resetTransactionRecords() {
  MODULES.forEach((module) => {
    state[module] = [];
  });
  state.sales = [];
  state.purchases = [];
  state.stockMovements = [];
  state.counters = { ...state.counters, sale: 0, purchase: 0 };
}

function clearSampleData() {
  const org = state.organization;
  if (!org.sampleLoaded) return;
  if (!window.confirm("Remove all sample records and return to your empty business?")) return;
  const type = org.typeBeforeSample || org.type;
  resetTransactionRecords();
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
    "Delete ALL records (items, sales, purchases, stock history, transactions, contacts, employees, assets)? Roles, fields, and the owner login stay.";
  if (!window.confirm(message)) return;
  resetTransactionRecords();
  state.dataSources = defaultDataSources();
  importSession = null;
  delete state.organization.sampleLoaded;
  addActivity("All records cleared", "setup");
  saveState();
  render();
  showToast("All records cleared");
}

/* ---------- Events ---------- */

els.homeScreen.addEventListener("click", (event) => {
  if (event.target.closest("[data-open-trial]")) return openTrialDialog();
  if (event.target.closest("[data-home-auth]")) return openAuthFromHome();
  if (event.target.closest("#homeMenuBtn")) {
    const open = els.homeNav.classList.toggle("is-open");
    els.homeMenuBtn.setAttribute("aria-expanded", String(open));
    return;
  }
  if (event.target.closest(".home-nav a")) {
    els.homeNav.classList.remove("is-open");
    els.homeMenuBtn.setAttribute("aria-expanded", "false");
  }
});
els.trialDialogCloseBtn.addEventListener("click", () => (els.trialDialogBackdrop.hidden = true));
els.trialDialogBackdrop.addEventListener("click", (event) => {
  if (event.target === els.trialDialogBackdrop) els.trialDialogBackdrop.hidden = true;
});
els.trialBlankBtn.addEventListener("click", () => beginTrial());
els.trialSampleBtn.addEventListener("click", () => beginTrial({ sample: true }));
els.trialResetBtn.addEventListener("click", resetTrial);
els.authBackHomeBtn.addEventListener("click", showHome);
els.exitTrialBtn.addEventListener("click", logout);
els.trialBannerExitBtn.addEventListener("click", logout);

els.registerForm.addEventListener("submit", handleRegister);
els.loginForm.addEventListener("submit", handleLogin);
els.resetBrowserBtn.addEventListener("click", resetBrowserData);
els.setupWizardForm.addEventListener("submit", submitSetupWizard);
els.setupWizardBackBtn.addEventListener("click", previousSetupWizardStep);
els.tutorialBtn.addEventListener("click", () => startTutorial());
els.tourBackBtn.addEventListener("click", () => moveTutorial(-1));
els.tourNextBtn.addEventListener("click", () => moveTutorial(1));
els.skipTutorialBtn.addEventListener("click", () => finishTutorial("skipped"));
byId("muteGuidesBtn").addEventListener("click", () => {
  const user = currentUser();
  if (user) {
    state.onboarding.autoGuidesOff ||= {};
    state.onboarding.autoGuidesOff[user.id] = true;
  }
  finishTutorial("skipped");
  showToast("Guides won't open automatically. Use the Guide button on any page.");
});
/* ---------- Searchable dropdowns ----------
 * Every dropdown in a form becomes a type-to-search box: type a few letters and matching
 * options appear (names, codes, anything in the option text). The original <select> stays in
 * the form, hidden, so saving, validation, and change handling work exactly as before.
 */

const COMBO_CONTAINERS = ["#recordForm", "#setupForm", "#setupWizardContent", "#importWizard", "#setupView .role-builder"];
const COMBO_RENDER_LIMIT = 150;

function comboOptions(select) {
  return [...select.options].map((option, index) => ({
    index,
    value: option.value,
    label: option.textContent.trim(),
    group: option.parentElement?.tagName === "OPTGROUP" ? option.parentElement.label : "",
    disabled: option.disabled,
  }));
}

function comboLabel(select) {
  const option = select.options[select.selectedIndex];
  return option ? option.textContent.trim() : "";
}

// An empty first option like "Choose…" or "—" acts as the placeholder, not a value.
function comboPlaceholderOption(select) {
  const first = select.options[0];
  return first && first.value === "" ? first.textContent.trim() : "";
}

function enhanceSelect(select) {
  if (select.dataset.combo || select.multiple || select.size > 1) return;
  select.dataset.combo = "on";

  const wrap = document.createElement("div");
  wrap.className = "combo";
  select.parentNode.insertBefore(wrap, select);
  wrap.appendChild(select);
  select.classList.add("combo-native");
  select.tabIndex = -1;
  select.setAttribute("aria-hidden", "true");

  const listId = `${select.id || select.name || "combo"}-list-${Math.random().toString(36).slice(2, 7)}`;
  const input = document.createElement("input");
  input.type = "text";
  input.className = "combo-input";
  input.autocomplete = "off";
  input.spellcheck = false;
  input.setAttribute("role", "combobox");
  input.setAttribute("aria-autocomplete", "list");
  input.setAttribute("aria-expanded", "false");
  input.setAttribute("aria-controls", listId);
  const label = select.id ? document.querySelector(`label[for="${CSS.escape(select.id)}"]`) : null;
  if (label) input.setAttribute("aria-label", label.textContent.replace("*", "").trim());
  else if (select.getAttribute("aria-label")) input.setAttribute("aria-label", select.getAttribute("aria-label"));

  const list = document.createElement("div");
  list.className = "combo-list";
  list.id = listId;
  list.setAttribute("role", "listbox");
  list.hidden = true;
  wrap.append(input, list);

  let matches = [];
  let active = -1;
  let typed = false;

  const sync = () => {
    const placeholder = comboPlaceholderOption(select);
    const empty = select.value === "" && select.selectedIndex <= 0 && placeholder;
    input.value = empty ? "" : comboLabel(select);
    input.placeholder = empty ? placeholder : "Type to search…";
    input.disabled = select.disabled;
    wrap.classList.toggle("is-disabled", select.disabled);
  };

  const close = () => {
    list.hidden = true;
    input.setAttribute("aria-expanded", "false");
    active = -1;
  };

  const highlight = (index) => {
    active = index;
    list.querySelectorAll(".combo-option").forEach((node) => {
      const on = Number(node.dataset.match) === index;
      node.classList.toggle("is-active", on);
      node.setAttribute("aria-selected", String(on));
      if (on) node.scrollIntoView({ block: "nearest" });
    });
  };

  const render = (query) => {
    const q = query.trim().toLowerCase();
    const all = comboOptions(select).filter((option) => !(option.value === "" && option.label === comboPlaceholderOption(select) && q));
    if (!q) {
      matches = all;
    } else {
      const starts = [];
      const contains = [];
      all.forEach((option) => {
        const text = option.label.toLowerCase();
        if (text.startsWith(q) || text.split(/[\s(·-]+/).some((word) => word.startsWith(q))) starts.push(option);
        else if (text.includes(q) || option.group.toLowerCase().includes(q)) contains.push(option);
      });
      matches = [...starts, ...contains];
    }
    const shown = matches.slice(0, COMBO_RENDER_LIMIT);
    let lastGroup = null;
    const escapeRe = (text) => text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const mark = (text) => (q ? escapeHtml(text).replace(new RegExp(`(${escapeRe(escapeHtml(q))})`, "i"), "<mark>$1</mark>") : escapeHtml(text));
    list.innerHTML = shown.length
      ? shown
          .map((option, index) => {
            const heading = option.group !== lastGroup && option.group ? `<div class="combo-group">${escapeHtml(option.group)}</div>` : "";
            lastGroup = option.group;
            return `${heading}<div class="combo-option ${option.value === select.value ? "is-selected" : ""} ${
              option.disabled ? "is-disabled" : ""
            }" role="option" data-match="${index}" aria-selected="false">${mark(option.label) || "&nbsp;"}</div>`;
          })
          .join("") +
        (matches.length > shown.length
          ? `<div class="combo-more">${matches.length - shown.length} more — keep typing to narrow down</div>`
          : "")
      : `<div class="combo-empty">No match for "${escapeHtml(query.trim())}"</div>`;
    list.hidden = false;
    input.setAttribute("aria-expanded", "true");
    const selectedIndex = shown.findIndex((option) => option.value === select.value);
    highlight(q ? (shown.length ? 0 : -1) : selectedIndex);
  };

  const choose = (option) => {
    if (!option || option.disabled) return;
    const changed = select.value !== option.value || select.selectedIndex !== option.index;
    select.selectedIndex = option.index;
    typed = false;
    sync();
    close();
    if (changed) {
      select.dispatchEvent(new Event("input", { bubbles: true }));
      select.dispatchEvent(new Event("change", { bubbles: true }));
    }
  };

  input.addEventListener("focus", () => {
    typed = false;
    input.select();
  });
  input.addEventListener("click", () => {
    if (list.hidden) render("");
  });
  input.addEventListener("input", (event) => {
    event.stopPropagation();
    typed = true;
    render(input.value);
  });
  input.addEventListener("change", (event) => event.stopPropagation());
  input.addEventListener("keydown", (event) => {
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      event.preventDefault();
      if (list.hidden) render(typed ? input.value : "");
      const count = Math.min(matches.length, COMBO_RENDER_LIMIT);
      if (!count) return;
      const step = event.key === "ArrowDown" ? 1 : -1;
      highlight(active < 0 ? (step > 0 ? 0 : count - 1) : (active + step + count) % count);
    } else if (event.key === "Enter") {
      if (!list.hidden) {
        event.preventDefault();
        if (active >= 0) choose(matches[active]);
        else close();
      }
    } else if (event.key === "Tab") {
      // Typing a few letters and pressing Tab picks the best match.
      if (typed && !list.hidden && active >= 0) choose(matches[active]);
      else close();
    } else if (event.key === "Escape") {
      if (!list.hidden) {
        event.preventDefault();
        event.stopPropagation();
        close();
        sync();
      }
    }
  });
  input.addEventListener("blur", () => {
    window.setTimeout(() => {
      if (document.activeElement === input) return;
      close();
      sync();
    }, 120);
  });
  list.addEventListener("mousedown", (event) => {
    // Keep focus in the input so blur doesn't close the list before the click lands.
    event.preventDefault();
  });
  list.addEventListener("click", (event) => {
    const node = event.target.closest(".combo-option");
    if (node) choose(matches[Number(node.dataset.match)]);
  });
  select.addEventListener("change", sync);
  select.addEventListener("combo-sync", () => {
    if (document.activeElement !== input) sync();
  });
  select.addEventListener("focus", () => input.focus());
  label?.addEventListener("click", (event) => {
    event.preventDefault();
    input.focus();
  });
  select.addEventListener("invalid", () => wrap.classList.add("is-invalid"));
  sync();
}

function enhanceSelectsIn(root) {
  if (!root) return;
  root.querySelectorAll("select").forEach(enhanceSelect);
}

function initSearchableSelects() {
  const observer = new MutationObserver((mutations) => {
    mutations.forEach((mutation) => {
      // Options replaced inside an enhanced dropdown: refresh its text.
      if (mutation.target.tagName === "SELECT" && mutation.target.dataset.combo) {
        mutation.target.dispatchEvent(new Event("combo-sync"));
      }
      mutation.addedNodes.forEach((node) => {
        if (node.nodeType !== 1) return;
        if (node.tagName === "SELECT") enhanceSelect(node);
        else enhanceSelectsIn(node);
      });
    });
  });
  COMBO_CONTAINERS.forEach((selector) => {
    const root = document.querySelector(selector);
    if (!root) return;
    enhanceSelectsIn(root);
    observer.observe(root, { childList: true, subtree: true });
  });
}

// Keep the visible text in step when code sets a dropdown's value directly.
function refreshSearchableSelects(root = document) {
  root.querySelectorAll("select[data-combo]").forEach((select) => select.dispatchEvent(new Event("combo-sync")));
}

initSearchableSelects();
addGuideButtons();
document.addEventListener("click", (event) => {
  if (event.target.closest("[data-page-guide]")) {
    startPageGuide();
    return;
  }
  if (event.target.closest("[data-toggle-field-help]")) {
    toggleFieldHelp();
    return;
  }
  const fieldHelp = event.target.closest("[data-field-help]");
  if (fieldHelp) fieldHelp.closest(".field")?.classList.toggle("help-open");
});
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

  const newTrade = target.closest("[data-new-trade]");
  if (newTrade) {
    openTradeForm(newTrade.dataset.newTrade);
    return;
  }
  const sellItem = target.closest("[data-sell-item]");
  if (sellItem) {
    openTradeForm("sale", { itemId: sellItem.dataset.sellItem });
    return;
  }
  const adjustStock = target.closest("[data-adjust-stock]");
  if (adjustStock) {
    openAdjustStock(adjustStock.dataset.adjustStock);
    return;
  }
  const stockHistory = target.closest("[data-stock-history]");
  if (stockHistory) {
    openStockHistory(stockHistory.dataset.stockHistory);
    return;
  }
  const contactHistory = target.closest("[data-contact-history]");
  if (contactHistory) {
    openContactHistory(contactHistory.dataset.contactHistory);
    return;
  }
  const documentButton = target.closest("[data-document]");
  if (documentButton) {
    openDocumentFor(documentButton.dataset.document);
    return;
  }
  const formatButton = target.closest("[data-document-format]");
  if (formatButton) {
    setDocumentFormat(formatButton.dataset.documentFormat);
    return;
  }
  if (target.closest("[data-print-document]")) {
    printActiveDocument();
    return;
  }
  const cancelTrade = target.closest("[data-cancel-trade]");
  if (cancelTrade) {
    const [kind, id] = split(cancelTrade, "data-cancel-trade");
    openCancelTrade(kind, id);
    return;
  }
  const viewTrade = target.closest("[data-view-trade]");
  if (viewTrade) {
    const [kind, id] = split(viewTrade, "data-view-trade");
    openTradeDetails(kind, id);
    return;
  }
  const addLine = target.closest("[data-trade-add-line]");
  if (addLine && TRADE[editing?.kind]) {
    captureTradeForm();
    editing.draft.lines.push(newTradeLine(editing.kind));
    renderTradeForm();
    els.recordForm.querySelector(".trade-line:last-child select")?.focus();
    return;
  }
  const removeLine = target.closest("[data-trade-remove-line]");
  if (removeLine && TRADE[editing?.kind]) {
    captureTradeForm();
    if (editing.draft.lines.length > 1) {
      editing.draft.lines = editing.draft.lines.filter((line) => line.key !== removeLine.dataset.tradeRemoveLine);
    }
    renderTradeForm();
    return;
  }
  const tradeRow = target.closest("[data-view-trade-row]");
  if (tradeRow && !target.closest("button")) {
    const [kind, id] = split(tradeRow, "data-view-trade-row");
    openTradeDetails(kind, id);
    return;
  }

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
    openSalaryPayment(employeeId, month);
    return;
  }
  if (target.closest("#payAllSalariesBtn")) {
    openPayAllSalaries(payrollMonth || todayIso().slice(0, 7));
    return;
  }
  const payrollLink = target.closest("[data-open-payroll]");
  if (payrollLink) {
    payrollMonth = payrollLink.dataset.openPayroll || null;
    closeModal();
    openTab("employees", "payroll");
    return;
  }
  const payHistoryButton = target.closest("[data-pay-history]");
  if (payHistoryButton) {
    payHistoryFilters.employeeId = payHistoryButton.dataset.payHistory;
    openTab("employees", "history");
    return;
  }
  const reverseSalary = target.closest("[data-reverse-salary]");
  if (reverseSalary) {
    reverseSalaryPayment(reverseSalary.dataset.reverseSalary);
    return;
  }
  const partyPayment = target.closest("[data-party-payment]");
  if (partyPayment) {
    const value = partyPayment.dataset.partyPayment;
    openPartyPayment(value.slice(0, value.indexOf(":")), value.slice(value.indexOf(":") + 1));
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
  if (TRADE[editing?.kind] && target.closest("#recordForm")) {
    captureTradeForm();
    if (target.matches("[data-trade-refresh]")) renderTradeForm();
    else updateTradeSummary();
    return;
  }
  const tradeFilterIds = {
    salesPeriodFilter: "sales",
    purchasesPeriodFilter: "purchases",
    stockTypeFilter: "stockType",
    reportPeriodSelect: "reportPeriod",
    reportCurrencySelect: "reportCurrency",
  };
  if (tradeFilterIds[target.id]) {
    tradeFilters[tradeFilterIds[target.id]] = target.value;
    render();
    return;
  }
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
  if (target.id === "paymentsDirectionFilter" || target.id === "paymentsMonthFilter") {
    paymentFilters[target.id === "paymentsDirectionFilter" ? "direction" : "month"] = target.value;
    render();
    return;
  }
  if (target.id === "payHistoryEmployeeFilter") {
    payHistoryFilters.employeeId = target.value;
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
  if (event.target.matches("[data-party-allocation]")) refreshPartyAllocation();
  if (event.target.matches("[data-salary-net]")) refreshSalaryNet();
  if (TRADE[editing?.kind] && event.target.closest("#recordForm") && !event.target.matches("select, [type=radio], [type=checkbox]")) {
    captureTradeForm();
    updateTradeSummary();
  }
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
  if (!els.tourLayer.hidden) renderTutorialStep();
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
  if (event.key !== "Escape") return;
  if (!els.trialDialogBackdrop.hidden) {
    els.trialDialogBackdrop.hidden = true;
    return;
  }
  if (!els.tourLayer.hidden) {
    finishTutorial("skipped");
    return;
  }
  if (!els.modalBackdrop.hidden) closeModal();
});

document.addEventListener(
  "scroll",
  () => {
    if (els.tourLayer.hidden) return;
    window.clearTimeout(tutorialPositionTimer);
    tutorialPositionTimer = window.setTimeout(() => {
      const step = tutorialSteps[tutorialStep];
      positionTutorial(document.querySelector(step?.selector) || document.querySelector(".main-area"));
    }, 30);
  },
  true,
);

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
else showHome();
