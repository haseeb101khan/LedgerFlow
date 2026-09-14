/*
 * Sample businesses for trying LedgerFlow. Each builder returns records for
 * every module. Fields added by a business template are given by label in
 * `extra`. Records default to the sample's `currency`; a record that was paid
 * or priced in another currency sets its own `currency`. `imported: true`
 * marks a record as if it came from a file, so missing details are flagged.
 */
(function () {
  function monthly(months, build) {
    return Array.from({ length: months }, (_, month) => build(month)).flat();
  }

  const samples = {};

  samples.Pharmacy = ({ daysAgo, staffRole }) => ({
    currency: "PKR",
    inventory: [
      { name: "Panadol 500mg Tablets (10x10)", sku: "MED-1001", category: "Analgesic", quantity: 240, reorderLevel: 60, unitCost: 245, sellPrice: 290, location: "Rack A1", extra: { "Batch Number": "PN2409", "Expiry Date": "2027-06-30", Manufacturer: "GSK", "Prescription Required": "No" } },
      { name: "Brufen 400mg Tablets", sku: "MED-1002", category: "Analgesic", quantity: 85, reorderLevel: 40, unitCost: 310, sellPrice: 365, location: "Rack A1", extra: { "Batch Number": "BF2311", "Expiry Date": "2026-11-30", Manufacturer: "Abbott", "Prescription Required": "No" } },
      { name: "Augmentin 625mg Tablets", sku: "MED-1003", category: "Antibiotic", quantity: 32, reorderLevel: 30, unitCost: 780, sellPrice: 915, location: "Rack B2", extra: { "Batch Number": "AG2402", "Expiry Date": "2026-10-15", Manufacturer: "GSK", "Prescription Required": "Yes" } },
      { name: "Amoxil 500mg Capsules", sku: "MED-1004", category: "Antibiotic", quantity: 18, reorderLevel: 25, unitCost: 420, sellPrice: 495, location: "Rack B2", extra: { "Batch Number": "AX2312", "Expiry Date": "2026-09-28", Manufacturer: "GSK", "Prescription Required": "Yes" } },
      { name: "Glucophage 500mg Tablets", sku: "MED-1005", category: "Diabetes", quantity: 140, reorderLevel: 50, unitCost: 190, sellPrice: 230, location: "Rack C1", extra: { "Batch Number": "GL2405", "Expiry Date": "2027-03-31", Manufacturer: "Merck", "Prescription Required": "Yes" } },
      { name: "Concor 5mg Tablets", sku: "MED-1006", category: "Cardiac", quantity: 60, reorderLevel: 30, unitCost: 520, sellPrice: 610, location: "Rack C1", extra: { "Batch Number": "CN2401", "Expiry Date": "2026-12-31", Manufacturer: "Merck", "Prescription Required": "Yes" } },
      { name: "Risek 20mg Capsules", sku: "MED-1007", category: "Gastro", quantity: 95, reorderLevel: 40, unitCost: 380, sellPrice: 450, location: "Rack C2", extra: { "Batch Number": "RS2406", "Expiry Date": "2027-05-31", Manufacturer: "Getz Pharma", "Prescription Required": "No" } },
      { name: "ORS Sachets (Nimkol)", sku: "MED-1008", category: "Gastro", quantity: 400, reorderLevel: 100, unitCost: 18, sellPrice: 25, location: "Rack D1", extra: { "Batch Number": "OR2407", "Expiry Date": "2027-08-31", Manufacturer: "Searle", "Prescription Required": "No" } },
      { name: "Ventolin Inhaler", sku: "MED-1009", category: "Respiratory", quantity: 22, reorderLevel: 15, unitCost: 650, sellPrice: 760, location: "Rack C3", extra: { "Batch Number": "VT2403", "Expiry Date": "2027-01-31", Manufacturer: "GSK", "Prescription Required": "Yes" } },
      { name: "Surbex-Z Tablets", sku: "MED-1010", category: "Vitamins", quantity: 75, reorderLevel: 30, unitCost: 540, sellPrice: 640, location: "Rack D2", extra: { "Batch Number": "SZ2404", "Expiry Date": "2027-04-30", Manufacturer: "Abbott", "Prescription Required": "No" } },
      { name: "Dettol Antiseptic 250ml", sku: "MED-1011", category: "Personal Care", quantity: 48, reorderLevel: 20, unitCost: 360, sellPrice: 420, location: "Rack E1", extra: { "Batch Number": "DT2408", "Expiry Date": "2028-02-28", Manufacturer: "Reckitt", "Prescription Required": "No" } },
      { name: "Insulin Mixtard 70/30 Vial", sku: "MED-1012", category: "Diabetes", quantity: 12, reorderLevel: 10, unitCost: "", sellPrice: 1650, location: "Fridge", imported: true, extra: { "Batch Number": "MX2405", "Expiry Date": "2026-12-15", Manufacturer: "Novo Nordisk", "Prescription Required": "Yes" } },
    ],
    finance: [
      ...monthly(5, (m) => [
        { date: daysAgo(m * 30 + 2), type: "Income", category: "Counter Sales", description: "Counter sales (week 1-2)", amount: 1480000 - m * 60000, status: "Paid" },
        { date: daysAgo(m * 30 + 16), type: "Income", category: "Counter Sales", description: "Counter sales (week 3-4)", amount: 1390000 - m * 45000, status: "Paid" },
        { date: daysAgo(m * 30 + 5), type: "Expense", category: "Medicine Purchases", description: "Stock purchase - Muller & Phipps", amount: 1650000 - m * 30000, status: m === 0 ? "Pending" : "Paid", party: "Muller & Phipps Distributors", dueDate: m === 0 ? daysAgo(-10) : "" },
        ...(m > 0 ? [{ date: daysAgo(m * 30 + 8), type: "Expense", category: "Salaries", description: "Staff salaries", amount: 395000, status: "Paid" }] : []),
        { date: daysAgo(m * 30 + 10), type: "Expense", category: "Rent", description: "Shop rent - Blue Area", amount: 220000, status: "Paid" },
        { date: daysAgo(m * 30 + 12), type: "Expense", category: "Utilities", description: "Electricity bill (IESCO)", amount: 64000 + m * 3000, status: "Paid" },
      ]),
      { date: daysAgo(6), type: "Income", category: "Credit Sales", description: "Monthly bill - Al-Noor Hospital", amount: 126000, status: "Pending", party: "Al-Noor Hospital", dueDate: daysAgo(-9) },
      { date: daysAgo(40), type: "Income", category: "Credit Sales", description: "Invoice #1182 - Shifa Clinic", amount: 48500, status: "Pending", party: "Shifa Clinic (credit account)", dueDate: daysAgo(10) },
    ],
    contacts: [
      { name: "Al-Noor Hospital", type: "Customer", phone: "051-2287450", email: "accounts@alnoor.example", balance: 0, status: "Active" },
      { name: "Shifa Clinic (credit account)", type: "Customer", phone: "051-4440921", email: "", balance: 0, status: "Active" },
      { name: "Mr. Tariq Mehmood", type: "Customer", phone: "", email: "", balance: 3200, status: "Active", imported: true },
      { name: "Muller & Phipps Distributors", type: "Supplier", phone: "051-5120088", email: "orders@mp.example", balance: -185000, status: "Active", extra: { "Drug License No.": "DSL-ICT-00481" } },
      { name: "IBL Healthcare Distribution", type: "Supplier", phone: "051-5732210", email: "isb@iblhc.example", balance: -92400, status: "Active", extra: { "Drug License No.": "DSL-ICT-00932" } },
      { name: "Premier Agency (Getz Pharma)", type: "Supplier", phone: "0300-5551234", email: "", balance: 0, status: "Active", extra: { "Drug License No.": "DSL-RWP-01177" } },
    ],
    employees: [
      { salaryPaid: true, name: "Ahmed Raza", role: "Head Pharmacist", department: "Pharmacy", phone: "0301-5550142", email: "ahmed.raza@example.com", accessRole: staffRole, salary: 180000, status: "Active" },
      { salaryPaid: true, name: "Sana Iqbal", role: "Pharmacist", department: "Pharmacy", phone: "0333-5550177", email: "sana.iqbal@example.com", accessRole: staffRole, salary: 120000, status: "Active" },
      { name: "Usman Ali", role: "Counter Sales", department: "Sales", phone: "0345-5550190", email: "", accessRole: staffRole, salary: 45000, status: "Active" },
      { name: "Hina Tariq", role: "Cashier", department: "Sales", phone: "0312-5550155", email: "", accessRole: staffRole, salary: 50000, status: "Active" },
      { name: "Imran Shah", role: "Delivery Rider", department: "Delivery", phone: "", email: "", accessRole: staffRole, salary: "", status: "Active", imported: true },
    ],
    assets: [
      { name: "Medicine Refrigerator (Haier)", code: "AST-PH-01", assignedTo: "Pharmacy", location: "Back Room", purchaseDate: "2023-02-14", value: 185000, condition: "Active" },
      { name: "POS & Barcode System", code: "AST-PH-02", assignedTo: "Counter", location: "Front Counter", purchaseDate: "2024-07-01", value: 95000, condition: "Active" },
      { name: "Generator 5 kVA", code: "AST-PH-03", assignedTo: "Store", location: "Rooftop", purchaseDate: "2022-05-20", value: 240000, condition: "Active" },
      { name: "Air Conditioner 1.5 Ton (Gree)", code: "AST-PH-04", assignedTo: "Store", location: "Sales Floor", purchaseDate: "2023-06-02", value: 165000, condition: "Active" },
      { name: "Delivery Motorcycle (Honda CD70)", code: "AST-PH-05", assignedTo: "Imran Shah", location: "Parking", purchaseDate: "2024-01-10", value: 158000, condition: "Maintenance", notes: "Clutch plate replacement due" },
    ],
  });

  samples["Utility Store"] = ({ daysAgo, staffRole }) => ({
    currency: "PKR",
    inventory: [
      { name: "Dalda Cooking Oil", sku: "UT-1001", category: "Cooking Oil", quantity: 64, reorderLevel: 20, unitCost: 2650, sellPrice: 2850, location: "Aisle 1", extra: { Brand: "Dalda", "Pack Size": "5 L", Barcode: "8961008200015" } },
      { name: "Tapal Danedar Tea", sku: "UT-1002", category: "Tea", quantity: 120, reorderLevel: 40, unitCost: 1050, sellPrice: 1160, location: "Aisle 2", extra: { Brand: "Tapal", "Pack Size": "950 g", Barcode: "8964000102345" } },
      { name: "Surf Excel Detergent", sku: "UT-1003", category: "Detergent", quantity: 18, reorderLevel: 25, unitCost: 690, sellPrice: 760, location: "Aisle 5", extra: { Brand: "Unilever", "Pack Size": "1 kg", Barcode: "8961014030017" } },
      { name: "Lux Soap", sku: "UT-1004", category: "Personal Care", quantity: 210, reorderLevel: 60, unitCost: 118, sellPrice: 135, location: "Aisle 5", extra: { Brand: "Lux", "Pack Size": "150 g", Barcode: "8961014012341" } },
      { name: "National Iodized Salt", sku: "UT-1005", category: "Spices", quantity: 150, reorderLevel: 50, unitCost: 55, sellPrice: 70, location: "Aisle 3", extra: { Brand: "National", "Pack Size": "800 g", Barcode: "8964000550012" } },
      { name: "Shan Biryani Masala", sku: "UT-1006", category: "Spices", quantity: 95, reorderLevel: 40, unitCost: 120, sellPrice: 145, location: "Aisle 3", extra: { Brand: "Shan", "Pack Size": "50 g", Barcode: "8964000660018" } },
      { name: "Nestle Milkpak", sku: "UT-1007", category: "Dairy", quantity: 36, reorderLevel: 48, unitCost: 290, sellPrice: 320, location: "Chiller", extra: { Brand: "Nestle", "Pack Size": "1 L", Barcode: "8961012345672" } },
      { name: "Colgate MaxFresh Toothpaste", sku: "UT-1008", category: "Personal Care", quantity: 72, reorderLevel: 30, unitCost: 310, sellPrice: 350, location: "Aisle 5", extra: { Brand: "Colgate", "Pack Size": "150 g", Barcode: "8964000770014" } },
      { name: "Super Kernel Basmati Rice", sku: "UT-1009", category: "Rice & Flour", quantity: 40, reorderLevel: 15, unitCost: 1750, sellPrice: 1950, location: "Aisle 4", extra: { Brand: "Guard", "Pack Size": "5 kg", Barcode: "8964000880011" } },
      { name: "Sugar (loose)", sku: "UT-1010", category: "Rice & Flour", quantity: 300, reorderLevel: 100, unitCost: 150, sellPrice: 165, location: "Store Room", extra: { Brand: "Loose", "Pack Size": "1 kg" } },
      { name: "Sunridge Chakki Atta", sku: "UT-1011", category: "Rice & Flour", quantity: 22, reorderLevel: 20, unitCost: 1280, sellPrice: 1390, location: "Aisle 4", extra: { Brand: "Sunridge", "Pack Size": "10 kg", Barcode: "8964000990018" } },
      { name: "Pepsi Bottle", sku: "UT-1012", category: "Beverages", quantity: 90, reorderLevel: 36, unitCost: 180, sellPrice: "", location: "Chiller", imported: true, extra: { Brand: "Pepsi", "Pack Size": "1.5 L" } },
    ],
    finance: [
      ...monthly(5, (m) => [
        { date: daysAgo(m * 30 + 1), type: "Income", category: "Counter Sales", description: "Counter sales - morning shift", amount: 820000 - m * 25000, status: "Paid", extra: { Shift: "Morning" } },
        { date: daysAgo(m * 30 + 1), type: "Income", category: "Counter Sales", description: "Counter sales - evening shift", amount: 1150000 - m * 30000, status: "Paid", extra: { Shift: "Evening" } },
        { date: daysAgo(m * 30 + 4), type: "Expense", category: "Stock Purchase", description: "Stock purchase - Metro Cash & Carry", amount: 1380000 - m * 20000, status: "Paid" },
        ...(m > 0 ? [{ date: daysAgo(m * 30 + 7), type: "Expense", category: "Salaries", description: "Staff salaries", amount: 215000, status: "Paid" }] : []),
        { date: daysAgo(m * 30 + 9), type: "Expense", category: "Rent", description: "Shop rent", amount: 150000, status: "Paid" },
        { date: daysAgo(m * 30 + 13), type: "Expense", category: "Utilities", description: "Electricity bill (LESCO)", amount: 68000 + m * 4000, status: m === 0 ? "Pending" : "Paid", party: "LESCO", dueDate: m === 0 ? daysAgo(-5) : "" },
      ]),
    ],
    contacts: [
      { name: "Imran Bhai (khata)", type: "Customer", phone: "0321-4450012", email: "", balance: 7850, status: "Active", extra: { Area: "Block C" } },
      { name: "Zahid Tailors", type: "Customer", phone: "0300-4418890", email: "", balance: 12400, status: "Pending", extra: { Area: "Main Market" } },
      { name: "Mrs. Nasreen (monthly account)", type: "Customer", phone: "", email: "", balance: 5600, status: "Active", imported: true, extra: { Area: "Street 14" } },
      { name: "Metro Cash & Carry", type: "Supplier", phone: "042-111-638-769", email: "", balance: -245000, status: "Active", extra: { Area: "Thokar Niaz Baig" } },
      { name: "Shah Traders (Unilever distributor)", type: "Supplier", phone: "0333-4471120", email: "orders@shahtraders.example", balance: -88000, status: "Active", extra: { Area: "Shahalam Market" } },
      { name: "Ahmed & Sons (Tapal dealer)", type: "Supplier", phone: "0345-4402231", email: "", balance: 0, status: "Active", extra: { Area: "Akbari Mandi" } },
    ],
    employees: [
      { salaryPaid: true, name: "Kashif Mehmood", role: "Store Manager", department: "Operations", phone: "0300-4012345", email: "kashif@example.com", accessRole: staffRole, salary: 85000, status: "Active" },
      { name: "Ayesha Siddiqui", role: "Cashier", department: "Counter", phone: "0321-4098765", email: "", accessRole: staffRole, salary: 42000, status: "Active" },
      { name: "Waqas Ahmed", role: "Cashier", department: "Counter", phone: "0333-4087654", email: "", accessRole: staffRole, salary: 42000, status: "Active" },
      { name: "Bilal Hussain", role: "Stock Keeper", department: "Store Room", phone: "0345-4076543", email: "", accessRole: staffRole, salary: 36000, status: "Active" },
      { name: "Faizan", role: "Delivery Boy", department: "Delivery", phone: "", email: "", accessRole: staffRole, salary: 10000, status: "Active", imported: true },
    ],
    assets: [
      { name: "POS System with Printer", code: "AST-US-01", assignedTo: "Counter", location: "Front", purchaseDate: "2024-03-12", value: 120000, condition: "Active" },
      { name: "Beverage Chiller (Waves)", code: "AST-US-02", assignedTo: "Store", location: "Aisle 6", purchaseDate: "2023-04-18", value: 145000, condition: "Active" },
      { name: "Deep Freezer (Dawlance)", code: "AST-US-03", assignedTo: "Store", location: "Aisle 6", purchaseDate: "2022-09-05", value: 110000, condition: "Damaged", notes: "Compressor noise, technician called" },
      { name: "CCTV System (8 cameras)", code: "AST-US-04", assignedTo: "Manager", location: "Whole shop", purchaseDate: "2024-08-20", value: 85000, condition: "Active" },
      { name: "Delivery Bike (Yamaha YB125Z)", code: "AST-US-05", assignedTo: "Faizan", location: "Parking", purchaseDate: "2025-02-01", value: 385000, condition: "Active" },
    ],
  });

  samples["Construction Company"] = ({ daysAgo, staffRole }) => ({
    currency: "PKR",
    inventory: [
      { name: "Cement - DG Khan OPC", sku: "MAT-001", category: "Cement", quantity: 820, reorderLevel: 300, unitCost: 1380, sellPrice: 1450, location: "Site Store - DHA Phase 7", extra: { Unit: "Bags" } },
      { name: "Steel Rebar 12mm (Grade 60)", sku: "MAT-002", category: "Steel", quantity: 14, reorderLevel: 10, unitCost: 265000, sellPrice: 275000, location: "Main Yard", extra: { Unit: "Tons" } },
      { name: "Bricks (A-class)", sku: "MAT-003", category: "Masonry", quantity: 42000, reorderLevel: 20000, unitCost: 17, sellPrice: 18, location: "Site - Bahria Commercial", extra: { Unit: "Pieces" } },
      { name: "Ravi Sand", sku: "MAT-004", category: "Aggregates", quantity: 3500, reorderLevel: 2000, unitCost: 55, sellPrice: 60, location: "Site - DHA Phase 7", extra: { Unit: "Cubic ft" } },
      { name: "Margalla Crush", sku: "MAT-005", category: "Aggregates", quantity: 1200, reorderLevel: 1500, unitCost: 130, sellPrice: 140, location: "Site - Bahria Commercial", extra: { Unit: "Cubic ft" } },
      { name: "Porcelain Tiles 24x24", sku: "MAT-006", category: "Finishing", quantity: 1850, reorderLevel: 500, unitCost: 310, sellPrice: 340, location: "Main Store", extra: { Unit: "Sq ft" } },
      { name: "PVC Pipe 4 inch (Beta)", sku: "MAT-007", category: "Plumbing", quantity: 160, reorderLevel: 50, unitCost: 1850, sellPrice: 2000, location: "Main Store", extra: { Unit: "Pieces" } },
      { name: "Electrical Wire 7/29 (Pakistan Cables)", sku: "MAT-008", category: "Electrical", quantity: 38, reorderLevel: 20, unitCost: 21500, sellPrice: 23000, location: "Main Store", extra: { Unit: "Rolls" } },
      { name: "Waterproofing Membrane (imported)", sku: "MAT-009", category: "Finishing", quantity: 45, reorderLevel: 20, unitCost: 180, sellPrice: 195, location: "Main Store", currency: "USD", extra: { Unit: "Rolls" } },
      { name: "Shuttering Plywood 8x4", sku: "MAT-010", category: "Formwork", quantity: 240, reorderLevel: 100, unitCost: 4200, sellPrice: 4500, location: "Main Yard", extra: { Unit: "Pieces" } },
    ],
    finance: [
      { date: daysAgo(3), type: "Income", category: "Progress Billing", description: "Progress bill #4 - DHA Phase 7 villas", amount: 12500000, status: "Pending", party: "Col. (R) Saleem Akhtar (DHA villa client)", dueDate: daysAgo(-12), extra: { Project: "DHA Phase 7 Villas" } },
      { date: daysAgo(33), type: "Income", category: "Progress Billing", description: "Progress bill #3 - DHA Phase 7 villas", amount: 11800000, status: "Paid", extra: { Project: "DHA Phase 7 Villas" } },
      { date: daysAgo(63), type: "Income", category: "Progress Billing", description: "Progress bill #2 - DHA Phase 7 villas", amount: 10200000, status: "Paid", extra: { Project: "DHA Phase 7 Villas" } },
      { date: daysAgo(18), type: "Income", category: "Progress Billing", description: "Progress bill #1 - Bahria Commercial Plaza", amount: 18000000, status: "Pending", party: "Bahria Commercial Developers", dueDate: daysAgo(3), extra: { Project: "Bahria Commercial Plaza" } },
      { date: daysAgo(95), type: "Income", category: "Advance", description: "Mobilization advance - Bahria Commercial Plaza", amount: 5000000, status: "Paid", extra: { Project: "Bahria Commercial Plaza" } },
      { date: daysAgo(125), type: "Income", category: "Progress Billing", description: "Progress bill #1 - DHA Phase 7 villas", amount: 9600000, status: "Paid", extra: { Project: "DHA Phase 7 Villas" } },
      ...monthly(5, (m) => [
        { date: daysAgo(m * 30 + 6), type: "Expense", category: "Materials", description: "Cement purchase - Al-Madina Traders", amount: 1130000 + m * 40000, status: m === 0 ? "Pending" : "Paid", party: "Al-Madina Traders (cement dealer)", dueDate: m === 0 ? daysAgo(-7) : "", extra: { Project: "DHA Phase 7 Villas" } },
        { date: daysAgo(m * 30 + 9), type: "Expense", category: "Materials", description: "Steel rebar - Amreli Steel", amount: 3710000 - m * 150000, status: "Paid", extra: { Project: m % 2 ? "Bahria Commercial Plaza" : "DHA Phase 7 Villas" } },
        ...(m > 0 ? [{ date: daysAgo(m * 30 + 2), type: "Expense", category: "Labour", description: "Labour wages (site payroll)", amount: 1640000, status: "Paid" }] : []),
        { date: daysAgo(m * 30 + 11), type: "Expense", category: "Machinery", description: "Diesel for machinery", amount: 410000 + m * 15000, status: "Paid" },
        { date: daysAgo(m * 30 + 14), type: "Expense", category: "Overheads", description: "Site office rent", amount: 120000, status: "Paid" },
      ]),
      { date: daysAgo(21), type: "Expense", category: "Materials", description: "Imported waterproofing membrane (Sika)", amount: 8100, currency: "USD", status: "Paid", extra: { Project: "Bahria Commercial Plaza" } },
      { date: daysAgo(40), type: "Expense", category: "Machinery", description: "Dump truck gearbox repair", amount: 185000, status: "Paid" },
    ],
    contacts: [
      { name: "Col. (R) Saleem Akhtar (DHA villa client)", type: "Customer", phone: "0300-8551020", email: "saleem.akhtar@example.com", balance: 0, status: "Active", extra: { Project: "DHA Phase 7 Villas" } },
      { name: "Bahria Commercial Developers", type: "Customer", phone: "051-5733400", email: "projects@bcd.example", balance: 0, status: "Active", extra: { Project: "Bahria Commercial Plaza" } },
      { name: "Al-Madina Traders (cement dealer)", type: "Supplier", phone: "0321-5102233", email: "", balance: -1450000, status: "Active" },
      { name: "Amreli Steel Distributor", type: "Supplier", phone: "042-35761100", email: "sales@amreli.example", balance: -2300000, status: "Active" },
      { name: "Sika Middle East (membrane supplier)", type: "Supplier", phone: "+971-4-4398200", email: "export@sika.example", balance: -3200, currency: "USD", status: "Active" },
      { name: "Crush Supplier - Gul Khan", type: "Supplier", phone: "", email: "", balance: -260000, status: "Active", imported: true },
    ],
    employees: [
      { salaryPaid: true, name: "Engr. Faisal Mehmood", role: "Project Engineer", department: "Engineering", phone: "0333-5104455", email: "faisal@example.com", accessRole: staffRole, salary: 280000, status: "Active", extra: { Trade: "Engineer", Site: "DHA Phase 7 Villas" } },
      { salaryPaid: true, name: "Nadeem Akhtar", role: "Site Supervisor", department: "Site Operations", phone: "0300-5107788", email: "", accessRole: staffRole, salary: 150000, status: "Active", extra: { Trade: "Supervisor", Site: "Bahria Commercial Plaza" } },
      { name: "Allah Ditta", role: "Head Mason", department: "Site Operations", phone: "0345-5109911", email: "", accessRole: staffRole, salary: 65000, status: "Active", extra: { Trade: "Mason", Site: "DHA Phase 7 Villas" } },
      { name: "Javed Iqbal", role: "Steel Fixer", department: "Site Operations", phone: "0312-5103322", email: "", accessRole: staffRole, salary: 60000, status: "Active", extra: { Trade: "Steel Fixer", Site: "Bahria Commercial Plaza" } },
      { name: "Arshad Mahmood", role: "Electrician", department: "Site Operations", phone: "0301-5106677", email: "", accessRole: staffRole, salary: 70000, status: "On Leave", extra: { Trade: "Electrician", Site: "DHA Phase 7 Villas" } },
      { name: "Gul Zaman", role: "Labourer", department: "Site Operations", phone: "", email: "", accessRole: staffRole, salary: 38000, status: "Active", imported: true, extra: { Trade: "Labour", Site: "Bahria Commercial Plaza" } },
      { name: "Sher Ali", role: "Labourer", department: "Site Operations", phone: "", email: "", accessRole: staffRole, salary: 38000, status: "Active", imported: true, extra: { Trade: "Labour" } },
    ],
    assets: [
      { name: "JCB 3CX Backhoe Loader", code: "MCH-001", assignedTo: "Gul Khan", location: "DHA Phase 7 Villas", purchaseDate: "2022-03-10", value: 32000000, condition: "Active", extra: { "Hours Used": 4820 } },
      { name: "Concrete Mixer (1 bag)", code: "MCH-002", assignedTo: "Rashid", location: "Bahria Commercial Plaza", purchaseDate: "2023-06-18", value: 450000, condition: "Active", extra: { "Hours Used": 1210 } },
      { name: "Hino Dump Truck", code: "MCH-003", assignedTo: "Shahid", location: "Main Yard", purchaseDate: "2021-11-02", value: 14500000, condition: "Maintenance", notes: "Gearbox repaired, awaiting road test", extra: { "Hours Used": 9300 } },
      { name: "Generator 100 kVA (Cummins)", code: "MCH-004", assignedTo: "Arshad Mahmood", location: "DHA Phase 7 Villas", purchaseDate: "2020-08-15", value: 3800000, condition: "Active", extra: { "Hours Used": 12650 } },
      { name: "Steel Scaffolding Set", code: "MCH-005", assignedTo: "Nadeem Akhtar", location: "Bahria Commercial Plaza", purchaseDate: "2024-01-20", value: 1250000, condition: "Active" },
      { name: "Rotating Laser Level (Bosch GRL 400)", code: "MCH-006", assignedTo: "Engr. Faisal Mehmood", location: "Main Office", purchaseDate: "2024-09-05", value: 1450, currency: "USD", condition: "Active", notes: "Bought from Dubai supplier in USD" },
    ],
  });

  samples["Vehicle Dealership"] = ({ daysAgo, staffRole }) => ({
    currency: "PKR",
    inventory: [
      { name: "Toyota Corolla Altis 1.6 X", sku: "VEH-2401", category: "Sedan", quantity: 1, unitCost: 7450000, sellPrice: 7899000, location: "Showroom", extra: { Make: "Toyota", "Model Year": 2025, "Chassis No.": "NZE170-8841230", "Engine (cc)": 1598, "Mileage (km)": 0, Color: "Super White", "Vehicle Condition": "New" } },
      { name: "Honda Civic RS Turbo", sku: "VEH-2402", category: "Sedan", quantity: 1, unitCost: 9850000, sellPrice: 10399000, location: "Showroom", extra: { Make: "Honda", "Model Year": 2025, "Chassis No.": "FE1-2031187", "Engine (cc)": 1498, "Mileage (km)": 0, Color: "Crystal Black", "Vehicle Condition": "New" } },
      { name: "Suzuki Alto VXL AGS", sku: "VEH-2403", category: "Hatchback", quantity: 2, unitCost: 3060000, sellPrice: 3245000, location: "Yard", extra: { Make: "Suzuki", "Model Year": 2025, "Chassis No.": "HA36S-551209", "Engine (cc)": 998, "Mileage (km)": 0, Color: "Silky Silver", "Vehicle Condition": "New" } },
      { name: "Toyota Aqua S (Japan import)", sku: "VEH-2404", category: "Hatchback", quantity: 1, unitCost: 5150000, sellPrice: 5650000, location: "Yard", notes: "Landed cost incl. customs duty", extra: { Make: "Toyota", "Model Year": 2019, "Chassis No.": "NHP10-6782211", "Engine (cc)": 1500, "Mileage (km)": 38400, Color: "Pearl White", "Vehicle Condition": "Reconditioned" } },
      { name: "Kia Sportage AWD", sku: "VEH-2405", category: "SUV", quantity: 1, unitCost: 9200000, sellPrice: 9650000, location: "Showroom", notes: "Trade-in", extra: { Make: "Kia", "Model Year": 2024, "Chassis No.": "QL-7730025", "Engine (cc)": 1999, "Mileage (km)": 12500, Color: "Snow White Pearl", "Vehicle Condition": "Used" } },
      { name: "Hyundai Tucson FWD A/T", sku: "VEH-2406", category: "SUV", quantity: 0, unitCost: 8300000, sellPrice: 8799000, location: "Delivered", notes: "Sold", extra: { Make: "Hyundai", "Model Year": 2024, "Chassis No.": "NX4-1102847", "Engine (cc)": 1999, "Mileage (km)": 8900, Color: "Phantom Black", "Vehicle Condition": "Used" } },
      { name: "Honda City 1.5 Aspire", sku: "VEH-2407", category: "Sedan", quantity: 1, unitCost: 4650000, sellPrice: 4899000, location: "Yard", extra: { Make: "Honda", "Model Year": 2023, "Chassis No.": "GM6-4410976", "Engine (cc)": 1497, "Mileage (km)": 26300, Color: "Lunar Silver", "Vehicle Condition": "Used" } },
      { name: "Suzuki Cultus VXL", sku: "VEH-2408", category: "Hatchback", quantity: 1, unitCost: "", sellPrice: 3950000, location: "Yard", imported: true, extra: { Make: "Suzuki", "Model Year": 2024, "Engine (cc)": 998, "Mileage (km)": 14200, Color: "Graphite Grey", "Vehicle Condition": "Used" } },
      { name: "Toyota Yaris ATIV X CVT", sku: "VEH-2409", category: "Sedan", quantity: 1, unitCost: 5600000, sellPrice: 5980000, location: "Showroom", extra: { Make: "Toyota", "Model Year": 2024, "Chassis No.": "NSP151-2021180", "Engine (cc)": 1496, "Mileage (km)": 9700, Color: "Beige Mica", "Vehicle Condition": "Used" } },
      { name: "Nissan Dayz Highway Star (Japan import)", sku: "VEH-2410", category: "Hatchback", quantity: 1, unitCost: 3450000, sellPrice: 3850000, location: "Yard", notes: "Landed cost incl. customs duty", extra: { Make: "Nissan", "Model Year": 2020, "Chassis No.": "B21W-0623345", "Engine (cc)": 659, "Mileage (km)": 41000, Color: "Pearl Black", "Vehicle Condition": "Reconditioned" } },
    ],
    finance: [
      { date: daysAgo(4), type: "Income", category: "Vehicle Sales", description: "Sale - Hyundai Tucson FWD", amount: 8799000, status: "Paid", extra: { "Stock No.": "VEH-2406" } },
      { date: daysAgo(9), type: "Income", category: "Bookings", description: "Booking token - Honda Civic RS", amount: 500000, status: "Pending", party: "Sadia Khan (Civic RS booking)", dueDate: daysAgo(-3), extra: { "Stock No.": "VEH-2402" } },
      { date: daysAgo(26), type: "Income", category: "Vehicle Sales", description: "Sale - Toyota Corolla GLi", amount: 4950000, status: "Paid", extra: { "Stock No.": "VEH-2388" } },
      { date: daysAgo(47), type: "Income", category: "Installments", description: "Installment received - Sportage buyer", amount: 1200000, status: "Paid" },
      { date: daysAgo(58), type: "Income", category: "Vehicle Sales", description: "Sale - Suzuki Swift GLX", amount: 4750000, status: "Paid", extra: { "Stock No.": "VEH-2391" } },
      { date: daysAgo(88), type: "Income", category: "Vehicle Sales", description: "Sale - Honda BR-V i-VTEC S", amount: 6450000, status: "Paid", extra: { "Stock No.": "VEH-2379" } },
      { date: daysAgo(118), type: "Income", category: "Vehicle Sales", description: "Sale - Toyota Fortuner Legender", amount: 17850000, status: "Paid", extra: { "Stock No.": "VEH-2365" } },
      { date: daysAgo(31), type: "Expense", category: "Vehicle Purchases", description: "Auction purchase - Toyota Aqua & Nissan Dayz (Japan)", amount: 14200, currency: "USD", status: "Paid" },
      { date: daysAgo(12), type: "Expense", category: "Import Costs", description: "Ocean freight & clearing agent", amount: 2350, currency: "USD", status: "Pending", party: "SBT Japan (auction exporter)", dueDate: daysAgo(2) },
      { date: daysAgo(10), type: "Expense", category: "Import Costs", description: "Customs duty - 2 imported vehicles", amount: 3900000, status: "Paid", extra: { "Stock No.": "VEH-2404" } },
      { date: daysAgo(52), type: "Expense", category: "Vehicle Purchases", description: "Trade-in purchase - Kia Sportage AWD", amount: 9200000, status: "Paid", extra: { "Stock No.": "VEH-2405" } },
      ...monthly(5, (m) => [
        { date: daysAgo(m * 30 + 5), type: "Expense", category: "Rent", description: "Showroom rent - Main Boulevard", amount: 450000, status: "Paid" },
        ...(m > 0 ? [{ date: daysAgo(m * 30 + 7), type: "Expense", category: "Salaries", description: "Staff salaries", amount: 690000, status: "Paid" }] : []),
        { date: daysAgo(m * 30 + 15), type: "Expense", category: "Preparation", description: "Detailing & pre-delivery inspection", amount: 85000 + m * 5000, status: "Paid" },
      ]),
    ],
    contacts: [
      { name: "Ahmed Raza (Sportage installment plan)", type: "Customer", phone: "0300-4441122", email: "ahmed.r@example.com", balance: 3400000, status: "Pending", extra: { "CNIC / Company Reg.": "35202-1234567-1" } },
      { name: "Bilal Motors (dealer credit)", type: "Customer", phone: "042-35881234", email: "", balance: 950000, status: "Active", extra: { "CNIC / Company Reg.": "NTN 4412098-3" } },
      { name: "Sadia Khan (Civic RS booking)", type: "Customer", phone: "0333-4412345", email: "", balance: 9399000, status: "Pending" },
      { name: "SBT Japan (auction exporter)", type: "Supplier", phone: "+81-45-594-7700", email: "sales@sbtjapan.example", balance: 0, currency: "USD", status: "Active" },
      { name: "Indus Motors Authorized Dealer", type: "Supplier", phone: "042-111-786-786", email: "", balance: -1200000, status: "Active" },
      { name: "Clearing Agent - Karachi Port", type: "Supplier", phone: "", email: "", balance: -180000, status: "Active", imported: true },
    ],
    employees: [
      { salaryPaid: true, name: "Hamza Farooq", role: "Sales Manager", department: "Sales", phone: "0300-4801122", email: "hamza@example.com", accessRole: staffRole, salary: 250000, status: "Active" },
      { name: "Omer Sheikh", role: "Sales Executive", department: "Sales", phone: "0321-4803344", email: "omer@example.com", accessRole: staffRole, salary: 110000, status: "Active" },
      { name: "Zain Abbas", role: "Sales Executive", department: "Sales", phone: "", email: "", accessRole: staffRole, salary: 110000, status: "Active", imported: true },
      { name: "Riaz Ahmed", role: "Mechanic", department: "Workshop", phone: "0345-4805566", email: "", accessRole: staffRole, salary: 95000, status: "Active" },
      { salaryPaid: true, name: "Maria Aslam", role: "Accountant", department: "Accounts", phone: "0333-4807788", email: "maria@example.com", accessRole: staffRole, salary: 130000, status: "Active" },
      { name: "Gul Muhammad", role: "Security Guard", department: "Admin", phone: "0312-4809900", email: "", accessRole: staffRole, salary: 45000, status: "Active" },
    ],
    assets: [
      { name: "Hydraulic Car Lift", code: "EQP-001", assignedTo: "Riaz Ahmed", location: "Workshop", purchaseDate: "2023-05-11", value: 1150000, condition: "Active" },
      { name: "Diagnostic Scanner (Launch X431)", code: "EQP-002", assignedTo: "Riaz Ahmed", location: "Workshop", purchaseDate: "2024-02-02", value: 520000, condition: "Active" },
      { name: "Tow Truck (Suzuki Ravi)", code: "EQP-003", assignedTo: "Workshop", location: "Yard", purchaseDate: "2022-10-20", value: 1900000, condition: "Maintenance", notes: "Hydraulic hose leaking" },
      { name: "Car Wash & Detailing Equipment", code: "EQP-004", assignedTo: "Workshop", location: "Wash Bay", purchaseDate: "2024-06-15", value: 380000, condition: "Active" },
      { name: "Showroom Furniture & Lighting", code: "EQP-005", assignedTo: "Admin", location: "Showroom", purchaseDate: "2023-01-09", value: 1600000, condition: "Active" },
    ],
  });

  window.LEDGERFLOW_SAMPLES = samples;
})();
