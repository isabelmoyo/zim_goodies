// =============================================================================
// ZIMGOODIES — Google Apps Script Business Tracker
// Partners: Isabel, Fra, Rumbi
// Costs: USD (Zimbabwe) | Sales: TZS (Tanzania)
// =============================================================================

// ---------------------------------------------------------------------------
// COLUMN INDEX CONSTANTS — change here if columns ever shift, nowhere else
// ---------------------------------------------------------------------------

// SALES LOG columns (1-indexed)
var COL_SALES_DATE       = 1;
var COL_SALES_CUSTOMER   = 2;
var COL_SALES_PRODUCT    = 3;
var COL_SALES_QTY        = 4;
var COL_SALES_UNIT_PRICE = 5;
var COL_SALES_TOTAL_TZS  = 6;
var COL_SALES_PAID_TZS   = 7;
var COL_SALES_BALANCE    = 8;
var COL_SALES_RECORDED   = 9;
var COL_SALES_NOTES      = 10;

// STOCK columns (1-indexed)
var COL_STOCK_PRODUCT    = 1;
var COL_STOCK_QTY        = 2;
var COL_STOCK_COST_USD   = 3;
var COL_STOCK_SELL_TZS   = 4;
var COL_STOCK_VAL_USD    = 5;
var COL_STOCK_VAL_TZS    = 6;
var COL_STOCK_LOW        = 7;
var COL_STOCK_UPDATED    = 8;

// DEBTS columns (1-indexed)
var COL_DEBT_DATE        = 1;
var COL_DEBT_CUSTOMER    = 2;
var COL_DEBT_ITEMS       = 3;
var COL_DEBT_ORIGINAL    = 4;
var COL_DEBT_PAID        = 5;
var COL_DEBT_OWED        = 6;
var COL_DEBT_STATUS      = 7;
var COL_DEBT_RECORDED    = 8;
var COL_DEBT_NOTES       = 9;

// EXPENSES columns (1-indexed)
var COL_EXP_DATE         = 1;
var COL_EXP_CATEGORY     = 2;
var COL_EXP_DESC         = 3;
var COL_EXP_AMOUNT       = 4;
var COL_EXP_CURRENCY     = 5;
var COL_EXP_AMOUNT_TZS   = 6;
var COL_EXP_PAID_BY      = 7;
var COL_EXP_NOTES        = 8;

// PARTNER ACCOUNT columns (1-indexed)
var COL_PA_DATE          = 1;
var COL_PA_PARTNER       = 2;
var COL_PA_TYPE          = 3;
var COL_PA_AMOUNT        = 4;
var COL_PA_CURRENCY      = 5;
var COL_PA_USD_EQUIV     = 6;
var COL_PA_NOTES         = 7;

// STOCK RECEIVED columns (1-indexed)
var COL_SR_DATE          = 1;
var COL_SR_PRODUCT       = 2;
var COL_SR_QTY           = 3;
var COL_SR_COST_UNIT     = 4;
var COL_SR_TOTAL_COST    = 5;
var COL_SR_SUPPLIER      = 6;
var COL_SR_RECORDED      = 7;

// ---------------------------------------------------------------------------
// DEFAULT DATA
// ---------------------------------------------------------------------------

var DEFAULT_PRODUCTS = [
  ['Mazoe Orange 2L',          3.00, 10000, 5],
  ['Mazoe Orange 500ml',       1.20,  3500, 10],
  ['Mazoe Blackcurrant 2L',    3.00, 10000, 5],
  ['Mazoe Blackcurrant 500ml', 1.20,  3500, 10],
  ['Mazoe Raspberry 2L',       3.00, 10000, 5],
  ['Mazoe Pink Lemonade 2L',   3.00, 10000, 5],
  ['Freezits (pack of 10)',    1.50,  5000, 10],
  ['Water 500ml',              0.50,  1500, 20],
  ['Water 1.5L',               0.80,  3000, 10]
];

var DEFAULT_PARTNERS   = ['Isabel', 'Fra', 'Rumbi'];
var DEFAULT_RATE       = 3200;
var DEFAULT_CATEGORIES = ['Stock Purchase', 'Transport', 'Packaging', 'Airtime', 'Other'];

// ---------------------------------------------------------------------------
// HELPER: get or create a sheet safely
// ---------------------------------------------------------------------------
function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

// ---------------------------------------------------------------------------
// HELPER: apply header style
// ---------------------------------------------------------------------------
function styleHeader(range, bgColor) {
  range.setBackground(bgColor)
       .setFontWeight('bold')
       .setFontSize(11)
       .setFontColor('#FFFFFF')
       .setVerticalAlignment('middle')
       .setHorizontalAlignment('center')
       .setWrap(true);
}

// ---------------------------------------------------------------------------
// HELPER: number format helpers
// ---------------------------------------------------------------------------
function fmtTZS(range) {
  range.setNumberFormat('#,##0 "TZS"');
}
function fmtUSD(range) {
  range.setNumberFormat('"$"#,##0.00');
}
function fmtDate(range) {
  range.setNumberFormat('DD/MM/YYYY');
}

// ---------------------------------------------------------------------------
// onOpen — adds custom menu
// ---------------------------------------------------------------------------
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('🛒 ZimGoodies')
      .addItem('🔄 Refresh Dashboard',        'updateDashboard')
      .addItem('📊 Daily WhatsApp Report',    'showDailyReport')
      .addSeparator()
      .addItem('⚙️ Setup / Reset Sheets',     'confirmSetup')
      .addToUi();
  } catch (e) {
    // silently ignore — onOpen errors are non-critical
  }
}

function confirmSetup() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    '⚙️ Setup ZimGoodies',
    'This will create all sheets and default data.\n\nExisting sheets will NOT be deleted — new sheets will be created safely.\n\nProceed?',
    ui.ButtonSet.YES_NO
  );
  if (result === ui.Button.YES) {
    setupSpreadsheet();
  }
}

// ---------------------------------------------------------------------------
// MAIN SETUP
// ---------------------------------------------------------------------------
function setupSpreadsheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    _setupSettings(ss);
    _setupDashboard(ss);
    _setupSalesLog(ss);
    _setupStock(ss);
    _setupDebts(ss);
    _setupExpenses(ss);
    _setupPartnerAccount(ss);
    _setupStockReceived(ss);

    // Reorder sheets visually
    var order = ['SETTINGS','DASHBOARD','SALES LOG','STOCK','DEBTS','EXPENSES','PARTNER ACCOUNT','STOCK RECEIVED'];
    for (var i = 0; i < order.length; i++) {
      var s = ss.getSheetByName(order[i]);
      if (s) ss.setActiveSheet(s), ss.moveActiveSheet(i + 1);
    }

    // Activate SETTINGS at the end
    ss.setActiveSheet(ss.getSheetByName('SETTINGS'));

    SpreadsheetApp.getActive().toast('✅ ZimGoodies is ready! Start with the SETTINGS sheet.', 'Setup Complete', 8);

  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Setup error: ' + e.message + '\n\n' + e.stack);
  }
}

// ---------------------------------------------------------------------------
// SHEET: SETTINGS
// ---------------------------------------------------------------------------
function _setupSettings(ss) {
  var sh = getOrCreateSheet(ss, 'SETTINGS');
  sh.setTabColor('#808080'); // grey

  // Only write default data if the sheet is essentially empty
  var existingData = sh.getLastRow();
  if (existingData > 5) return; // already set up, skip

  sh.clearContents();
  sh.clearFormats();
  sh.clearNotes();

  // ----- Title note -----
  sh.getRange('A1').setValue('✏️ Edit anything in this sheet — changes flow through the whole system')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground('#e8f0fe')
    .setFontColor('#1a73e8');
  sh.getRange('A1:E1').merge().setHorizontalAlignment('left').setWrap(true);
  sh.setRowHeight(1, 40);

  // ----- Business Name -----
  sh.getRange('A3').setValue('Business Name').setFontWeight('bold');
  sh.getRange('B3').setValue('ZimGoodies');

  // ----- Exchange Rate -----
  sh.getRange('A5').setValue('💱 USD → TZS Exchange Rate').setFontWeight('bold');
  sh.getRange('B5').setValue('1 USD =').setHorizontalAlignment('right');
  sh.getRange('C5').setValue(DEFAULT_RATE).setNumberFormat('#,##0').setBackground('#fff2cc').setFontWeight('bold').setFontSize(13);
  sh.getRange('D5').setValue('TZS');
  sh.getRange('A5:D5').setNote('Update this cell whenever the exchange rate changes. All calculations update automatically.');

  // ----- Partner Names -----
  sh.getRange('A7').setValue('👥 Partner Names').setFontWeight('bold');
  sh.getRange('B7').setValue('Partner 1');
  sh.getRange('C7').setValue('Partner 2');
  sh.getRange('D7').setValue('Partner 3');
  sh.getRange('B7:D7').setBackground('#d9ead3').setFontWeight('bold');
  sh.getRange('B8').setValue(DEFAULT_PARTNERS[0]);
  sh.getRange('C8').setValue(DEFAULT_PARTNERS[1]);
  sh.getRange('D8').setValue(DEFAULT_PARTNERS[2]);

  // ----- Expense Categories -----
  sh.getRange('A10').setValue('📋 Expense Categories').setFontWeight('bold');
  sh.getRange('A10:A10').setNote('Add or edit categories here. They appear in dropdowns across all sheets.');
  for (var i = 0; i < DEFAULT_CATEGORIES.length; i++) {
    sh.getRange(11 + i, 1).setValue(DEFAULT_CATEGORIES[i]);
  }

  // ----- Product Table Header -----
  var productHeaderRow = 11 + DEFAULT_CATEGORIES.length + 2;
  sh.getRange(productHeaderRow, 1).setValue('📦 Product List').setFontWeight('bold').setFontSize(12);
  sh.getRange(productHeaderRow, 1, 1, 5).merge();

  var headerRow = productHeaderRow + 1;
  var productHeaders = ['Product Name', 'Cost Price USD', 'Selling Price TZS', 'Low Stock Threshold', 'Notes'];
  sh.getRange(headerRow, 1, 1, productHeaders.length).setValues([productHeaders]);
  styleHeader(sh.getRange(headerRow, 1, 1, productHeaders.length), '#4a86e8');

  // ----- Products -----
  for (var p = 0; p < DEFAULT_PRODUCTS.length; p++) {
    var row = headerRow + 1 + p;
    sh.getRange(row, 1).setValue(DEFAULT_PRODUCTS[p][0]);
    sh.getRange(row, 2).setValue(DEFAULT_PRODUCTS[p][1]).setNumberFormat('"$"#,##0.00');
    sh.getRange(row, 3).setValue(DEFAULT_PRODUCTS[p][2]).setNumberFormat('#,##0');
    sh.getRange(row, 4).setValue(DEFAULT_PRODUCTS[p][3]);
    if (p % 2 === 0) {
      sh.getRange(row, 1, 1, 5).setBackground('#f8f9fa');
    }
  }

  // ----- Column widths -----
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 140);
  sh.setColumnWidth(3, 140);
  sh.setColumnWidth(4, 140);
  sh.setColumnWidth(5, 200);

  // ----- Named Ranges -----
  // Remove any existing named ranges first
  var namedRanges = ss.getNamedRanges();
  for (var n = 0; n < namedRanges.length; n++) {
    var nm = namedRanges[n].getName();
    if (nm === 'EXCHANGE_RATE' || nm === 'PARTNER_NAMES' || nm === 'PRODUCT_LIST' || nm === 'EXPENSE_CATEGORIES') {
      namedRanges[n].remove();
    }
  }

  ss.setNamedRange('EXCHANGE_RATE',       sh.getRange('C5'));
  ss.setNamedRange('PARTNER_NAMES',       sh.getRange('B8:D8'));
  ss.setNamedRange('EXPENSE_CATEGORIES',  sh.getRange(11, 1, DEFAULT_CATEGORIES.length, 1));
  ss.setNamedRange('PRODUCT_LIST',        sh.getRange(headerRow + 1, 1, DEFAULT_PRODUCTS.length, 4));
}

// ---------------------------------------------------------------------------
// SHEET: DASHBOARD
// ---------------------------------------------------------------------------
function _setupDashboard(ss) {
  var sh = getOrCreateSheet(ss, 'DASHBOARD');
  sh.setTabColor('#1155cc'); // blue

  if (sh.getLastRow() > 3) return; // already has data

  sh.clearContents();
  sh.clearFormats();

  sh.getRange('A1').setValue('📊 ZimGoodies Dashboard')
    .setFontWeight('bold').setFontSize(16).setFontColor('#1155cc');
  sh.getRange('A2').setValue('Last refreshed: (not yet refreshed — click 🔄 Refresh Dashboard)')
    .setFontColor('#888888').setFontSize(10);

  sh.setColumnWidth(1, 280);
  sh.setColumnWidth(2, 180);
  sh.setColumnWidth(3, 160);

  // Row 4 placeholder — updateDashboard() writes real content
  sh.getRange('A4').setValue('👉 Use the ZimGoodies menu → Refresh Dashboard to populate this sheet.')
    .setFontColor('#999999').setFontStyle('italic').setFontSize(11);
}

// ---------------------------------------------------------------------------
// SHEET: SALES LOG
// ---------------------------------------------------------------------------
function _setupSalesLog(ss) {
  var sh = getOrCreateSheet(ss, 'SALES LOG');
  sh.setTabColor('#0f9d58'); // green

  var headers = [
    'Date', 'Customer Name', 'Product', 'Qty',
    'Unit Price TZS', 'Total TZS', 'Amount Paid TZS',
    'Balance Owed TZS', 'Recorded By', 'Notes'
  ];

  if (sh.getLastRow() >= 1 && sh.getRange(1,1).getValue() === 'Date') {
    // Headers already exist — just ensure dropdowns and formatting are current
    _applySalesLogFormatting(ss, sh);
    return;
  }

  sh.clearContents();
  sh.clearFormats();
  sh.clearConditionalFormatRules();

  // Headers
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh.getRange(1, 1, 1, headers.length), '#0f9d58');
  sh.setFrozenRows(1);

  // Notes on header cells
  sh.getRange(1, COL_SALES_TOTAL_TZS).setNote('Formula: Qty × Unit Price. Auto-calculated.');
  sh.getRange(1, COL_SALES_BALANCE).setNote('Formula: Total TZS − Amount Paid. Positive = customer still owes money.');
  sh.getRange(1, COL_SALES_UNIT_PRICE).setNote('Auto-filled from SETTINGS product table. You can override manually.');

  // Example row
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var exRow = [
    today, 'Example Customer', DEFAULT_PRODUCTS[0][0], 2,
    DEFAULT_PRODUCTS[0][2],
    '=D2*E2',
    DEFAULT_PRODUCTS[0][2] * 2,
    '=F2-G2',
    DEFAULT_PARTNERS[0],
    'Example row — safe to delete'
  ];
  sh.getRange(2, 1, 1, exRow.length).setValues([exRow]);
  fmtDate(sh.getRange(2, COL_SALES_DATE));
  sh.getRange(2, COL_SALES_UNIT_PRICE).setNumberFormat('#,##0');
  sh.getRange(2, COL_SALES_TOTAL_TZS).setNumberFormat('#,##0 "TZS"');
  sh.getRange(2, COL_SALES_PAID_TZS).setNumberFormat('#,##0 "TZS"');
  sh.getRange(2, COL_SALES_BALANCE).setNumberFormat('#,##0 "TZS"');

  _applySalesLogFormatting(ss, sh);
}

function _applySalesLogFormatting(ss, sh) {
  // Column widths
  sh.setColumnWidth(COL_SALES_DATE,       100);
  sh.setColumnWidth(COL_SALES_CUSTOMER,   160);
  sh.setColumnWidth(COL_SALES_PRODUCT,    200);
  sh.setColumnWidth(COL_SALES_QTY,         60);
  sh.setColumnWidth(COL_SALES_UNIT_PRICE, 130);
  sh.setColumnWidth(COL_SALES_TOTAL_TZS,  140);
  sh.setColumnWidth(COL_SALES_PAID_TZS,   140);
  sh.setColumnWidth(COL_SALES_BALANCE,    140);
  sh.setColumnWidth(COL_SALES_RECORDED,   120);
  sh.setColumnWidth(COL_SALES_NOTES,      220);

  // Conditional formatting: balance owed > 0 → light red
  var rules = sh.getConditionalFormatRules();
  var redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$H2>0')
    .setBackground('#fce8e6')
    .setRanges([sh.getRange('A2:J1000')])
    .build();
  rules.push(redRule);
  sh.setConditionalFormatRules(rules);

  // Alternating row shading (light grey)
  var altRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(MOD(ROW(),2)=0,$A2<>"")')
    .setBackground('#f3f3f3')
    .setRanges([sh.getRange('A2:J1000')])
    .build();
  // Note: conditional format rules are evaluated in order; red row wins for owed > 0
  sh.setConditionalFormatRules([redRule, altRule]);

  // Data validations
  var productNames = DEFAULT_PRODUCTS.map(function(p){ return p[0]; });
  var productRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(productNames, true)
    .setAllowInvalid(true)
    .build();
  sh.getRange(2, COL_SALES_PRODUCT, 998, 1).setDataValidation(productRule);

  var partnerRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DEFAULT_PARTNERS, true)
    .setAllowInvalid(true)
    .build();
  sh.getRange(2, COL_SALES_RECORDED, 998, 1).setDataValidation(partnerRule);
}

// ---------------------------------------------------------------------------
// SHEET: STOCK
// ---------------------------------------------------------------------------
function _setupStock(ss) {
  var sh = getOrCreateSheet(ss, 'STOCK');
  sh.setTabColor('#ff6d00'); // orange

  var headers = [
    'Product', 'Current Qty', 'Cost Price USD', 'Sell Price TZS',
    'Total Stock Value USD', 'Total Stock Value TZS', 'Low Stock?', 'Last Updated'
  ];

  if (sh.getLastRow() >= 1 && sh.getRange(1,1).getValue() === 'Product') {
    _applyStockFormatting(sh);
    return;
  }

  sh.clearContents();
  sh.clearFormats();
  sh.clearConditionalFormatRules();

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh.getRange(1, 1, 1, headers.length), '#e65100');
  sh.setFrozenRows(1);

  sh.getRange(1, COL_STOCK_QTY).setNote('✏️ Update this column when stock changes.');
  sh.getRange(1, COL_STOCK_UPDATED).setNote('✏️ Enter today\'s date when you update the qty.');
  sh.getRange(1, COL_STOCK_COST_USD).setNote('Auto-pulled from SETTINGS. Edit prices in SETTINGS sheet.');

  // Determine SETTINGS product table location
  var settingsSh = ss.getSheetByName('SETTINGS');
  // Products start after header — find product list range via named range
  var productRange = ss.getRangeByName('PRODUCT_LIST');
  var productStartRow = productRange ? productRange.getRow() : 20;
  var settingsName = 'SETTINGS';

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  for (var p = 0; p < DEFAULT_PRODUCTS.length; p++) {
    var row = 2 + p;
    // Product name (hardcoded — matches SETTINGS)
    sh.getRange(row, COL_STOCK_PRODUCT).setValue(DEFAULT_PRODUCTS[p][0]);
    // Current qty — default 0, partners edit
    sh.getRange(row, COL_STOCK_QTY).setValue(0);
    // Cost and sell pulled via VLOOKUP from SETTINGS PRODUCT_LIST
    sh.getRange(row, COL_STOCK_COST_USD)
      .setFormula('=IFERROR(VLOOKUP(A' + row + ',PRODUCT_LIST,2,FALSE),"")');
    sh.getRange(row, COL_STOCK_SELL_TZS)
      .setFormula('=IFERROR(VLOOKUP(A' + row + ',PRODUCT_LIST,3,FALSE),"")');
    // Total value formulas
    sh.getRange(row, COL_STOCK_VAL_USD)
      .setFormula('=IF(B' + row + '="","",B' + row + '*C' + row + ')');
    sh.getRange(row, COL_STOCK_VAL_TZS)
      .setFormula('=IF(B' + row + '="","",B' + row + '*D' + row + ')');
    // Low Stock? — threshold from VLOOKUP col 4
    sh.getRange(row, COL_STOCK_LOW)
      .setFormula('=IF(B' + row + '="","",IF(B' + row + '<=IFERROR(VLOOKUP(A' + row + ',PRODUCT_LIST,4,FALSE),5),"YES","NO"))');
    sh.getRange(row, COL_STOCK_UPDATED).setValue(today);

    if (p % 2 === 0) sh.getRange(row, 1, 1, headers.length).setBackground('#fff3e0');
  }

  _applyStockFormatting(sh);
}

function _applyStockFormatting(sh) {
  sh.setColumnWidth(COL_STOCK_PRODUCT,   220);
  sh.setColumnWidth(COL_STOCK_QTY,        90);
  sh.setColumnWidth(COL_STOCK_COST_USD,  130);
  sh.setColumnWidth(COL_STOCK_SELL_TZS,  140);
  sh.setColumnWidth(COL_STOCK_VAL_USD,   160);
  sh.setColumnWidth(COL_STOCK_VAL_TZS,   160);
  sh.setColumnWidth(COL_STOCK_LOW,       110);
  sh.setColumnWidth(COL_STOCK_UPDATED,   120);

  var dataRange2 = sh.getRange('A2:H100');
  var rules = [];

  // Red if qty = 0
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$B2=0')
    .setBackground('#fce8e6')
    .setRanges([dataRange2])
    .build());

  // Orange if LOW STOCK = YES
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$G2="YES"')
    .setBackground('#ffe0b2')
    .setRanges([dataRange2])
    .build());

  sh.setConditionalFormatRules(rules);

  // Number formats
  sh.getRange('C2:C100').setNumberFormat('"$"#,##0.00');
  sh.getRange('D2:D100').setNumberFormat('#,##0 "TZS"');
  sh.getRange('E2:E100').setNumberFormat('"$"#,##0.00');
  sh.getRange('F2:F100').setNumberFormat('#,##0 "TZS"');
  sh.getRange('H2:H100').setNumberFormat('DD/MM/YYYY');
}

// ---------------------------------------------------------------------------
// SHEET: DEBTS
// ---------------------------------------------------------------------------
function _setupDebts(ss) {
  var sh = getOrCreateSheet(ss, 'DEBTS');
  sh.setTabColor('#cc0000'); // red

  var headers = [
    'Date', 'Customer Name', 'Items Description',
    'Original Amount TZS', 'Amount Paid TZS', 'Still Owed TZS',
    'Status', 'Recorded By', 'Notes'
  ];

  if (sh.getLastRow() >= 1 && sh.getRange(1,1).getValue() === 'Date') {
    _applyDebtsFormatting(sh);
    return;
  }

  sh.clearContents();
  sh.clearFormats();
  sh.clearConditionalFormatRules();

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh.getRange(1, 1, 1, headers.length), '#cc0000');
  sh.setFrozenRows(1);

  sh.getRange(1, COL_DEBT_OWED).setNote('Formula: Original Amount − Amount Paid. Auto-calculated.');
  sh.getRange(1, COL_DEBT_STATUS).setNote('Auto: CLEARED when fully paid, OUTSTANDING when balance remains.');

  _applyDebtsFormatting(sh);
}

function _applyDebtsFormatting(sh) {
  sh.setColumnWidth(COL_DEBT_DATE,      100);
  sh.setColumnWidth(COL_DEBT_CUSTOMER,  160);
  sh.setColumnWidth(COL_DEBT_ITEMS,     220);
  sh.setColumnWidth(COL_DEBT_ORIGINAL,  160);
  sh.setColumnWidth(COL_DEBT_PAID,      140);
  sh.setColumnWidth(COL_DEBT_OWED,      140);
  sh.setColumnWidth(COL_DEBT_STATUS,    140);
  sh.setColumnWidth(COL_DEBT_RECORDED,  120);
  sh.setColumnWidth(COL_DEBT_NOTES,     220);

  var dataRange = sh.getRange('A2:I1000');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G2="⚠️ OUTSTANDING"')
      .setBackground('#fce8e6')
      .setRanges([dataRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$G2="✅ CLEARED"')
      .setBackground('#d9ead3')
      .setRanges([dataRange]).build()
  ]);

  sh.getRange('D2:F1000').setNumberFormat('#,##0 "TZS"');
  sh.getRange('A2:A1000').setNumberFormat('DD/MM/YYYY');

  // Partner dropdown
  var partnerRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DEFAULT_PARTNERS, true)
    .setAllowInvalid(true).build();
  sh.getRange(2, COL_DEBT_RECORDED, 998, 1).setDataValidation(partnerRule);
}

// ---------------------------------------------------------------------------
// SHEET: EXPENSES
// ---------------------------------------------------------------------------
function _setupExpenses(ss) {
  var sh = getOrCreateSheet(ss, 'EXPENSES');
  sh.setTabColor('#f6b26b'); // yellow-orange

  var headers = [
    'Date', 'Category', 'Description', 'Amount',
    'Currency', 'Amount in TZS', 'Paid By', 'Notes'
  ];

  if (sh.getLastRow() >= 1 && sh.getRange(1,1).getValue() === 'Date') {
    _applyExpensesFormatting(sh);
    return;
  }

  sh.clearContents();
  sh.clearFormats();

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh.getRange(1, 1, 1, headers.length), '#b45f06');
  sh.setFrozenRows(1);

  sh.getRange(1, COL_EXP_AMOUNT_TZS)
    .setNote('Formula: if Currency=USD then Amount × Exchange Rate, else Amount as-is. Used for all totals.');
  sh.getRange(1, COL_EXP_CURRENCY)
    .setNote('Select TZS for local expenses. Select USD for stock purchases paid in Zimbabwe.');

  _applyExpensesFormatting(sh);
}

function _applyExpensesFormatting(sh) {
  sh.setColumnWidth(COL_EXP_DATE,       100);
  sh.setColumnWidth(COL_EXP_CATEGORY,   160);
  sh.setColumnWidth(COL_EXP_DESC,       220);
  sh.setColumnWidth(COL_EXP_AMOUNT,     130);
  sh.setColumnWidth(COL_EXP_CURRENCY,    90);
  sh.setColumnWidth(COL_EXP_AMOUNT_TZS, 150);
  sh.setColumnWidth(COL_EXP_PAID_BY,    120);
  sh.setColumnWidth(COL_EXP_NOTES,      220);

  sh.getRange('A2:A1000').setNumberFormat('DD/MM/YYYY');
  sh.getRange('F2:F1000').setNumberFormat('#,##0 "TZS"');

  var catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DEFAULT_CATEGORIES, true)
    .setAllowInvalid(true).build();
  sh.getRange(2, COL_EXP_CATEGORY, 998, 1).setDataValidation(catRule);

  var currencyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TZS', 'USD'], true)
    .setAllowInvalid(false).build();
  sh.getRange(2, COL_EXP_CURRENCY, 998, 1).setDataValidation(currencyRule);

  var partnerRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DEFAULT_PARTNERS, true)
    .setAllowInvalid(true).build();
  sh.getRange(2, COL_EXP_PAID_BY, 998, 1).setDataValidation(partnerRule);
}

// ---------------------------------------------------------------------------
// SHEET: PARTNER ACCOUNT
// ---------------------------------------------------------------------------
function _setupPartnerAccount(ss) {
  var sh = getOrCreateSheet(ss, 'PARTNER ACCOUNT');
  sh.setTabColor('#7030a0'); // purple

  var headers = [
    'Date', 'Partner', 'Type', 'Amount',
    'Currency', 'USD Equivalent', 'Notes'
  ];

  if (sh.getLastRow() >= 1 && sh.getRange(1,1).getValue() === 'Date') {
    _applyPartnerAccountFormatting(sh);
    return;
  }

  sh.clearContents();
  sh.clearFormats();

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh.getRange(1, 1, 1, headers.length), '#7030a0');
  sh.setFrozenRows(1);

  sh.getRange(1, COL_PA_USD_EQUIV)
    .setNote('Formula: if Currency=USD use Amount directly; if TZS divide by Exchange Rate.');

  _applyPartnerAccountFormatting(sh);
}

function _applyPartnerAccountFormatting(sh) {
  sh.setColumnWidth(COL_PA_DATE,       100);
  sh.setColumnWidth(COL_PA_PARTNER,    130);
  sh.setColumnWidth(COL_PA_TYPE,       140);
  sh.setColumnWidth(COL_PA_AMOUNT,     130);
  sh.setColumnWidth(COL_PA_CURRENCY,    90);
  sh.setColumnWidth(COL_PA_USD_EQUIV,  150);
  sh.setColumnWidth(COL_PA_NOTES,      240);

  sh.getRange('A2:A1000').setNumberFormat('DD/MM/YYYY');
  sh.getRange('F2:F1000').setNumberFormat('"$"#,##0.00');

  var partnerRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DEFAULT_PARTNERS, true)
    .setAllowInvalid(true).build();
  sh.getRange(2, COL_PA_PARTNER, 998, 1).setDataValidation(partnerRule);

  var typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Withdrawal', 'Contribution'], true)
    .setAllowInvalid(false).build();
  sh.getRange(2, COL_PA_TYPE, 998, 1).setDataValidation(typeRule);

  var currencyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['TZS', 'USD'], true)
    .setAllowInvalid(false).build();
  sh.getRange(2, COL_PA_CURRENCY, 998, 1).setDataValidation(currencyRule);
}

// ---------------------------------------------------------------------------
// SHEET: STOCK RECEIVED
// ---------------------------------------------------------------------------
function _setupStockReceived(ss) {
  var sh = getOrCreateSheet(ss, 'STOCK RECEIVED');
  sh.setTabColor('#00897b'); // teal

  var headers = [
    'Date', 'Product', 'Qty Received', 'Cost Per Unit USD',
    'Total Cost USD', 'Supplier / Notes', 'Recorded By'
  ];

  if (sh.getLastRow() >= 1 && sh.getRange(1,1).getValue() === 'Date') {
    _applyStockReceivedFormatting(sh);
    return;
  }

  sh.clearContents();
  sh.clearFormats();

  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader(sh.getRange(1, 1, 1, headers.length), '#00897b');
  sh.setFrozenRows(1);

  sh.getRange(1, COL_SR_TOTAL_COST)
    .setNote('Formula: Qty Received × Cost Per Unit USD. Auto-calculated.');

  _applyStockReceivedFormatting(sh);
}

function _applyStockReceivedFormatting(sh) {
  sh.setColumnWidth(COL_SR_DATE,       100);
  sh.setColumnWidth(COL_SR_PRODUCT,    220);
  sh.setColumnWidth(COL_SR_QTY,         90);
  sh.setColumnWidth(COL_SR_COST_UNIT,  150);
  sh.setColumnWidth(COL_SR_TOTAL_COST, 150);
  sh.setColumnWidth(COL_SR_SUPPLIER,   240);
  sh.setColumnWidth(COL_SR_RECORDED,   130);

  sh.getRange('A2:A1000').setNumberFormat('DD/MM/YYYY');
  sh.getRange('D2:E1000').setNumberFormat('"$"#,##0.00');

  var productNames = DEFAULT_PRODUCTS.map(function(p){ return p[0]; });
  var productRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(productNames, true)
    .setAllowInvalid(true).build();
  sh.getRange(2, COL_SR_PRODUCT, 998, 1).setDataValidation(productRule);

  var partnerRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(DEFAULT_PARTNERS, true)
    .setAllowInvalid(true).build();
  sh.getRange(2, COL_SR_RECORDED, 998, 1).setDataValidation(partnerRule);
}

// ---------------------------------------------------------------------------
// UPDATE DASHBOARD
// ---------------------------------------------------------------------------
function updateDashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dash = ss.getSheetByName('DASHBOARD');
    if (!dash) {
      SpreadsheetApp.getUi().alert('❌ DASHBOARD sheet not found. Please run Setup first.');
      return;
    }

    // Get exchange rate safely
    var rate = 0;
    try {
      var rateRange = ss.getRangeByName('EXCHANGE_RATE');
      if (rateRange) rate = Number(rateRange.getValue()) || 0;
    } catch(e) {}
    if (rate === 0) {
      // Fallback: read directly from SETTINGS C5
      var settingsSh = ss.getSheetByName('SETTINGS');
      if (settingsSh) rate = Number(settingsSh.getRange('C5').getValue()) || DEFAULT_RATE;
    }

    // ---- Aggregate SALES LOG ----
    var totalSalesTZS     = 0;
    var totalCollectedTZS = 0;
    var salesLogSh = ss.getSheetByName('SALES LOG');
    if (salesLogSh && salesLogSh.getLastRow() > 1) {
      var salesData = salesLogSh.getRange(2, 1, salesLogSh.getLastRow() - 1, COL_SALES_NOTES).getValues();
      for (var i = 0; i < salesData.length; i++) {
        var totalTZS = Number(salesData[i][COL_SALES_TOTAL_TZS - 1]) || 0;
        var paidTZS  = Number(salesData[i][COL_SALES_PAID_TZS  - 1]) || 0;
        totalSalesTZS     += totalTZS;
        totalCollectedTZS += paidTZS;
      }
    }

    // ---- Aggregate DEBTS ----
    var totalDebtsTZS = 0;
    var outstandingCount = 0;
    var debtsSh = ss.getSheetByName('DEBTS');
    if (debtsSh && debtsSh.getLastRow() > 1) {
      var debtsData = debtsSh.getRange(2, 1, debtsSh.getLastRow() - 1, COL_DEBT_NOTES).getValues();
      for (var d = 0; d < debtsData.length; d++) {
        var status = String(debtsData[d][COL_DEBT_STATUS - 1]);
        var owed   = Number(debtsData[d][COL_DEBT_OWED   - 1]) || 0;
        if (status.indexOf('OUTSTANDING') !== -1 || (owed > 0 && status.indexOf('CLEARED') === -1)) {
          totalDebtsTZS += owed;
          outstandingCount++;
        }
      }
    }

    // ---- Aggregate EXPENSES ----
    var totalExpensesTZS = 0;
    var expensesSh = ss.getSheetByName('EXPENSES');
    if (expensesSh && expensesSh.getLastRow() > 1) {
      var expData = expensesSh.getRange(2, 1, expensesSh.getLastRow() - 1, COL_EXP_NOTES).getValues();
      for (var ex = 0; ex < expData.length; ex++) {
        var amtRaw  = Number(expData[ex][COL_EXP_AMOUNT     - 1]) || 0;
        var curr    = String(expData[ex][COL_EXP_CURRENCY   - 1]).trim().toUpperCase();
        var amtTZS  = Number(expData[ex][COL_EXP_AMOUNT_TZS - 1]) || 0;
        // Prefer the pre-calculated TZS column; fall back to converting ourselves
        if (amtTZS > 0) {
          totalExpensesTZS += amtTZS;
        } else if (curr === 'USD' && rate > 0) {
          totalExpensesTZS += amtRaw * rate;
        } else {
          totalExpensesTZS += amtRaw;
        }
      }
    }

    // ---- Aggregate STOCK RECEIVED ----
    var totalStockCostUSD = 0;
    var srSh = ss.getSheetByName('STOCK RECEIVED');
    if (srSh && srSh.getLastRow() > 1) {
      var srData = srSh.getRange(2, 1, srSh.getLastRow() - 1, COL_SR_RECORDED).getValues();
      for (var sr = 0; sr < srData.length; sr++) {
        totalStockCostUSD += Number(srData[sr][COL_SR_TOTAL_COST - 1]) || 0;
      }
    }

    // ---- Partner Account adjustments ----
    var partnerAdj = { Isabel: 0, Fra: 0, Rumbi: 0 };
    var paSh = ss.getSheetByName('PARTNER ACCOUNT');
    if (paSh && paSh.getLastRow() > 1) {
      var paData = paSh.getRange(2, 1, paSh.getLastRow() - 1, COL_PA_NOTES).getValues();
      for (var pa = 0; pa < paData.length; pa++) {
        var partner = String(paData[pa][COL_PA_PARTNER   - 1]).trim();
        var type    = String(paData[pa][COL_PA_TYPE      - 1]).trim();
        var usdEq   = Number(paData[pa][COL_PA_USD_EQUIV - 1]) || 0;
        if (partnerAdj.hasOwnProperty(partner)) {
          if (type === 'Withdrawal')   partnerAdj[partner] -= usdEq;
          if (type === 'Contribution') partnerAdj[partner] += usdEq;
        }
      }
    }

    // ---- Derived calculations ----
    var grossRevenueUSD = rate > 0 ? totalCollectedTZS / rate : 0;
    var expensesUSD     = rate > 0 ? totalExpensesTZS / rate : 0;
    var netProfitUSD    = grossRevenueUSD - totalStockCostUSD - expensesUSD;
    var perPartnerUSD   = netProfitUSD / 3;
    var perPartnerTZS   = perPartnerUSD * rate;

    var partnerShares = {
      Isabel: perPartnerUSD + (partnerAdj.Isabel || 0),
      Fra:    perPartnerUSD + (partnerAdj.Fra    || 0),
      Rumbi:  perPartnerUSD + (partnerAdj.Rumbi  || 0)
    };

    // ---- Stock Alerts ----
    var stockAlerts = [];
    var stockSh = ss.getSheetByName('STOCK');
    if (stockSh && stockSh.getLastRow() > 1) {
      var stockData = stockSh.getRange(2, 1, stockSh.getLastRow() - 1, COL_STOCK_LOW).getValues();
      for (var st = 0; st < stockData.length; st++) {
        if (String(stockData[st][COL_STOCK_LOW - 1]).trim() === 'YES') {
          stockAlerts.push({
            product: stockData[st][COL_STOCK_PRODUCT - 1],
            qty:     stockData[st][COL_STOCK_QTY    - 1]
          });
        }
      }
    }

    // ---- Write to DASHBOARD ----
    dash.clearContents();
    dash.clearFormats();

    var tz = Session.getScriptTimeZone();
    var now = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');

    var titleRange = dash.getRange('A1');
    titleRange.setValue('📊 ZimGoodies Dashboard')
      .setFontWeight('bold').setFontSize(16).setFontColor('#1155cc');

    dash.getRange('A2').setValue('Last refreshed: ' + now)
      .setFontColor('#888888').setFontSize(10);

    // Exchange rate notice
    dash.getRange('A3').setValue('Exchange Rate in use: 1 USD = ' + _fmtNum(rate) + ' TZS')
      .setFontColor('#666666').setFontSize(10).setFontStyle('italic');

    var r = 5; // current write row

    // --- Section: Revenue ---
    dash.getRange(r, 1).setValue('💰 REVENUE & SALES').setFontWeight('bold').setFontSize(12)
      .setBackground('#e8f0fe').setFontColor('#1155cc');
    dash.getRange(r, 1, 1, 3).merge();
    r++;

    _dashRow(dash, r++, '  Total Sales Revenue',   _fmtTZSval(totalSalesTZS),     'All sales logged (incl. unpaid)');
    _dashRow(dash, r++, '  Total Cash Collected',  _fmtTZSval(totalCollectedTZS), 'Actual money received');
    _dashRow(dash, r++, '  Total Still Owed',       _fmtTZSval(totalDebtsTZS),     outstandingCount + ' outstanding debt(s)');
    _dashRow(dash, r++, '  Gross Revenue USD',      _fmtUSDval(grossRevenueUSD),   'Cash collected ÷ exchange rate');
    r++;

    // --- Section: Costs ---
    dash.getRange(r, 1).setValue('🧾 COSTS & EXPENSES').setFontWeight('bold').setFontSize(12)
      .setBackground('#fff2cc').setFontColor('#b45f06');
    dash.getRange(r, 1, 1, 3).merge();
    r++;

    _dashRow(dash, r++, '  Total Stock Cost USD',  _fmtUSDval(totalStockCostUSD), 'From STOCK RECEIVED log');
    _dashRow(dash, r++, '  Total Expenses',         _fmtTZSval(totalExpensesTZS),  '(' + _fmtUSDval(expensesUSD) + ' USD equiv.)');
    r++;

    // --- Section: Profit ---
    dash.getRange(r, 1).setValue('📈 PROFIT').setFontWeight('bold').setFontSize(12)
      .setBackground('#d9ead3').setFontColor('#274e13');
    dash.getRange(r, 1, 1, 3).merge();
    r++;

    _dashRow(dash, r++, '  Net Profit (USD)',       _fmtUSDval(netProfitUSD),      'Revenue − Stock Cost − Expenses');
    _dashRow(dash, r++, '  Per Partner (3 equal)',  _fmtUSDval(perPartnerUSD),     '(' + _fmtTZSval(perPartnerTZS) + ')');
    r++;

    // --- Partner breakdown ---
    dash.getRange(r, 1).setValue('👥 PARTNER SHARES (after adjustments)').setFontWeight('bold').setFontSize(12)
      .setBackground('#ead1dc').setFontColor('#4a1942');
    dash.getRange(r, 1, 1, 3).merge();
    r++;

    var partners = ['Isabel', 'Fra', 'Rumbi'];
    for (var pi = 0; pi < partners.length; pi++) {
      var pName = partners[pi];
      var pUSD  = partnerShares[pName];
      var pTZS  = pUSD * rate;
      _dashRow(dash, r++, '  ' + pName, _fmtUSDval(pUSD), '(' + _fmtTZSval(pTZS) + ')');
    }
    r++;

    // --- Stock Alerts ---
    dash.getRange(r, 1).setValue('📦 STOCK ALERTS').setFontWeight('bold').setFontSize(12)
      .setBackground('#ffe0b2').setFontColor('#e65100');
    dash.getRange(r, 1, 1, 3).merge();
    r++;

    if (stockAlerts.length === 0) {
      _dashRow(dash, r++, '  ✅ All products adequately stocked', '', '');
    } else {
      for (var sa = 0; sa < stockAlerts.length; sa++) {
        _dashRow(dash, r++, '  ⚠️ ' + stockAlerts[sa].product, stockAlerts[sa].qty + ' remaining', 'LOW STOCK');
        dash.getRange(r - 1, 1, 1, 3).setBackground('#fff3e0');
      }
    }
    r++;

    // --- Outstanding debts summary ---
    dash.getRange(r, 1).setValue('💳 OUTSTANDING DEBTS').setFontWeight('bold').setFontSize(12)
      .setBackground('#fce8e6').setFontColor('#c0392b');
    dash.getRange(r, 1, 1, 3).merge();
    r++;

    _dashRow(dash, r++, '  Open debts count', outstandingCount + ' customer(s)', '');
    _dashRow(dash, r++, '  Total owed', _fmtTZSval(totalDebtsTZS), '');

    // Final formatting
    dash.setColumnWidth(1, 280);
    dash.setColumnWidth(2, 180);
    dash.setColumnWidth(3, 230);

    SpreadsheetApp.getActive().toast('✅ Dashboard refreshed at ' + now, 'Done', 5);

  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Dashboard update error:\n\n' + e.message + '\n\n' + e.stack);
  }
}

// Write a labelled row to the dashboard
function _dashRow(sh, row, label, value, note) {
  sh.getRange(row, 1).setValue(label).setFontSize(10);
  sh.getRange(row, 2).setValue(value).setFontSize(10).setFontWeight('bold').setHorizontalAlignment('right');
  if (note) sh.getRange(row, 3).setValue(note).setFontSize(9).setFontColor('#888888');
  if (row % 2 === 0) sh.getRange(row, 1, 1, 3).setBackground('#f8f9fa');
}

function _fmtNum(n) {
  return Number(n).toLocaleString('en-US');
}
function _fmtTZSval(n) {
  return Number(n).toLocaleString('en-US', {maximumFractionDigits: 0}) + ' TZS';
}
function _fmtUSDval(n) {
  return '$' + Number(n).toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2});
}

// ---------------------------------------------------------------------------
// SHOW DAILY REPORT (WhatsApp-ready)
// ---------------------------------------------------------------------------
function showDailyReport() {
  try {
    var ss   = SpreadsheetApp.getActiveSpreadsheet();
    var tz   = Session.getScriptTimeZone();
    var today = new Date();
    var todayStr = Utilities.formatDate(today, tz, 'dd/MM/yyyy');
    var displayDate = Utilities.formatDate(today, tz, 'EEEE, dd MMMM yyyy');

    // Exchange rate
    var rate = DEFAULT_RATE;
    try {
      var rateRange = ss.getRangeByName('EXCHANGE_RATE');
      if (rateRange) rate = Number(rateRange.getValue()) || DEFAULT_RATE;
    } catch(e) {}
    if (rate <= 0) rate = DEFAULT_RATE;

    // ---- Today's sales ----
    var todaySales = 0, todayCollected = 0, todayOwed = 0, txCount = 0;
    var salesLogSh = ss.getSheetByName('SALES LOG');
    if (salesLogSh && salesLogSh.getLastRow() > 1) {
      var salesData = salesLogSh.getRange(2, 1, salesLogSh.getLastRow() - 1, COL_SALES_NOTES).getValues();
      for (var i = 0; i < salesData.length; i++) {
        var rowDate = salesData[i][COL_SALES_DATE - 1];
        var rowDateStr = '';
        if (rowDate instanceof Date) {
          rowDateStr = Utilities.formatDate(rowDate, tz, 'dd/MM/yyyy');
        } else {
          rowDateStr = String(rowDate).trim();
        }
        if (rowDateStr === todayStr) {
          txCount++;
          todaySales     += Number(salesData[i][COL_SALES_TOTAL_TZS - 1]) || 0;
          todayCollected += Number(salesData[i][COL_SALES_PAID_TZS  - 1]) || 0;
          todayOwed      += Number(salesData[i][COL_SALES_BALANCE   - 1]) || 0;
        }
      }
    }

    // ---- Today's expenses ----
    var todayExpenses = 0;
    var expSh = ss.getSheetByName('EXPENSES');
    if (expSh && expSh.getLastRow() > 1) {
      var expData = expSh.getRange(2, 1, expSh.getLastRow() - 1, COL_EXP_NOTES).getValues();
      for (var ex = 0; ex < expData.length; ex++) {
        var exDateStr = '';
        var exDate = expData[ex][COL_EXP_DATE - 1];
        if (exDate instanceof Date) {
          exDateStr = Utilities.formatDate(exDate, tz, 'dd/MM/yyyy');
        } else {
          exDateStr = String(exDate).trim();
        }
        if (exDateStr === todayStr) {
          var amtTZS = Number(expData[ex][COL_EXP_AMOUNT_TZS - 1]) || 0;
          if (amtTZS > 0) {
            todayExpenses += amtTZS;
          } else {
            var amt  = Number(expData[ex][COL_EXP_AMOUNT   - 1]) || 0;
            var curr = String(expData[ex][COL_EXP_CURRENCY - 1]).trim().toUpperCase();
            todayExpenses += (curr === 'USD') ? amt * rate : amt;
          }
        }
      }
    }

    // ---- All-time outstanding debts ----
    var allDebts = 0, debtors = 0;
    var debtsSh = ss.getSheetByName('DEBTS');
    if (debtsSh && debtsSh.getLastRow() > 1) {
      var debtsData = debtsSh.getRange(2, 1, debtsSh.getLastRow() - 1, COL_DEBT_NOTES).getValues();
      for (var d = 0; d < debtsData.length; d++) {
        var status = String(debtsData[d][COL_DEBT_STATUS - 1]);
        var owed   = Number(debtsData[d][COL_DEBT_OWED   - 1]) || 0;
        if (status.indexOf('OUTSTANDING') !== -1 || (owed > 0 && status.indexOf('CLEARED') === -1)) {
          allDebts += owed;
          debtors++;
        }
      }
    }

    // ---- Stock alerts ----
    var alertLines = [];
    var stockSh = ss.getSheetByName('STOCK');
    if (stockSh && stockSh.getLastRow() > 1) {
      var stockData = stockSh.getRange(2, 1, stockSh.getLastRow() - 1, COL_STOCK_LOW).getValues();
      for (var st = 0; st < stockData.length; st++) {
        if (String(stockData[st][COL_STOCK_LOW - 1]).trim() === 'YES') {
          alertLines.push('• ' + stockData[st][COL_STOCK_PRODUCT - 1] + ': ' + stockData[st][COL_STOCK_QTY - 1] + ' remaining');
        }
      }
    }

    // ---- Derived ----
    var netTodayTZS  = todayCollected - todayExpenses;
    var netTodayUSD  = rate > 0 ? netTodayTZS / rate : 0;
    var perPartTZS   = netTodayTZS / 3;
    var perPartUSD   = netTodayUSD / 3;

    // ---- Build message ----
    var lines = [
      '🛒 *ZimGoodies Daily Report*',
      '📅 ' + displayDate,
      '─────────────────────',
      '💰 Today\'s Sales: ' + _fmtTZSval(todaySales) + ' (' + txCount + ' transaction' + (txCount !== 1 ? 's' : '') + ')',
      '✅ Cash Collected: ' + _fmtTZSval(todayCollected),
      '⚠️  Still Owed Today: ' + _fmtTZSval(todayOwed),
      '🧾 Expenses Today: ' + _fmtTZSval(todayExpenses),
      '📊 Net Today: ' + _fmtTZSval(netTodayTZS) + ' (~' + _fmtUSDval(netTodayUSD) + ')',
      '',
      '👥 *Partner Shares (net today)*',
      '• Isabel: ' + _fmtTZSval(perPartTZS) + ' (~' + _fmtUSDval(perPartUSD) + ')',
      '• Fra:    ' + _fmtTZSval(perPartTZS) + ' (~' + _fmtUSDval(perPartUSD) + ')',
      '• Rumbi:  ' + _fmtTZSval(perPartTZS) + ' (~' + _fmtUSDval(perPartUSD) + ')',
      '',
      '📦 *Stock Alerts*'
    ];

    if (alertLines.length === 0) {
      lines.push('• All products adequately stocked ✅');
    } else {
      for (var al = 0; al < alertLines.length; al++) lines.push(alertLines[al]);
    }

    lines = lines.concat([
      '',
      '💳 *Outstanding Debts*',
      '• Total: ' + _fmtTZSval(allDebts) + ' from ' + debtors + ' customer' + (debtors !== 1 ? 's' : ''),
      '─────────────────────',
      'Rate: 1 USD = ' + _fmtNum(rate) + ' TZS'
    ]);

    var message = lines.join('\n');

    // Display in HTML dialog for easy copy
    var html = HtmlService.createHtmlOutput(
      '<style>' +
      'body { font-family: monospace; font-size: 13px; padding: 12px; white-space: pre-wrap; }' +
      'textarea { width: 100%; height: 380px; font-family: monospace; font-size: 12px; border: 1px solid #ccc; padding: 8px; }' +
      'p { color: #666; font-size: 11px; margin-top: 8px; }' +
      '</style>' +
      '<textarea id="msg">' + message.replace(/</g, '&lt;').replace(/>/g, '&gt;') + '</textarea>' +
      '<p>👆 Click inside the box, press Ctrl+A to select all, then Ctrl+C to copy. Paste into WhatsApp.</p>' +
      '<button onclick="document.getElementById(\'msg\').select();document.execCommand(\'copy\');this.textContent=\'✅ Copied!\'">📋 Copy to Clipboard</button>'
    ).setWidth(500).setHeight(500);

    SpreadsheetApp.getUi().showModalDialog(html, '📊 Daily WhatsApp Report — ' + todayStr);

  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Report error:\n\n' + e.message + '\n\n' + e.stack);
  }
}

// ---------------------------------------------------------------------------
// HELPER: Add formula rows to Expenses/Debts/Partner Account when new data
// rows are entered — called via installable trigger (optional improvement)
// ---------------------------------------------------------------------------

/**
 * onEdit trigger — auto-fills formulas in key columns when partners enter data.
 * Install as an installable trigger if desired, or leave as-is for simple usage.
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh   = e.range.getSheet();
    var name = sh.getName();
    var row  = e.range.getRow();
    var col  = e.range.getColumn();

    // Skip header row
    if (row < 2) return;

    // SALES LOG: auto-fill formulas if any data column edited
    if (name === 'SALES LOG' && col <= COL_SALES_QTY && col >= COL_SALES_DATE) {
      if (sh.getRange(row, COL_SALES_QTY).getValue() > 0) {
        // Total TZS = Qty * Unit Price
        if (sh.getRange(row, COL_SALES_TOTAL_TZS).getValue() === '') {
          sh.getRange(row, COL_SALES_TOTAL_TZS).setFormula('=D' + row + '*E' + row);
          sh.getRange(row, COL_SALES_TOTAL_TZS).setNumberFormat('#,##0 "TZS"');
        }
        // Balance = Total - Paid
        if (sh.getRange(row, COL_SALES_BALANCE).getValue() === '') {
          sh.getRange(row, COL_SALES_BALANCE).setFormula('=F' + row + '-G' + row);
          sh.getRange(row, COL_SALES_BALANCE).setNumberFormat('#,##0 "TZS"');
        }
      }
    }

    // DEBTS: auto-fill Still Owed and Status formulas
    if (name === 'DEBTS' && col <= COL_DEBT_PAID && col >= COL_DEBT_DATE) {
      if (sh.getRange(row, COL_DEBT_ORIGINAL).getValue() !== '') {
        sh.getRange(row, COL_DEBT_OWED)
          .setFormula('=D' + row + '-E' + row)
          .setNumberFormat('#,##0 "TZS"');
        sh.getRange(row, COL_DEBT_STATUS)
          .setFormula('=IF(F' + row + '<=0,"✅ CLEARED","⚠️ OUTSTANDING")');
      }
    }

    // EXPENSES: auto-fill Amount in TZS
    if (name === 'EXPENSES' && (col === COL_EXP_AMOUNT || col === COL_EXP_CURRENCY)) {
      var amt  = sh.getRange(row, COL_EXP_AMOUNT).getValue();
      var curr = sh.getRange(row, COL_EXP_CURRENCY).getValue();
      if (amt !== '' && curr !== '') {
        sh.getRange(row, COL_EXP_AMOUNT_TZS)
          .setFormula('=IF(E' + row + '="USD",D' + row + '*EXCHANGE_RATE,D' + row + ')')
          .setNumberFormat('#,##0 "TZS"');
      }
    }

    // PARTNER ACCOUNT: auto-fill USD Equivalent
    if (name === 'PARTNER ACCOUNT' && (col === COL_PA_AMOUNT || col === COL_PA_CURRENCY)) {
      var paAmt  = sh.getRange(row, COL_PA_AMOUNT).getValue();
      var paCurr = sh.getRange(row, COL_PA_CURRENCY).getValue();
      if (paAmt !== '' && paCurr !== '') {
        sh.getRange(row, COL_PA_USD_EQUIV)
          .setFormula('=IF(E' + row + '="USD",D' + row + ',D' + row + '/EXCHANGE_RATE)')
          .setNumberFormat('"$"#,##0.00');
      }
    }

    // STOCK RECEIVED: auto-fill Total Cost USD
    if (name === 'STOCK RECEIVED' && (col === COL_SR_QTY || col === COL_SR_COST_UNIT)) {
      var qty  = sh.getRange(row, COL_SR_QTY).getValue();
      var cost = sh.getRange(row, COL_SR_COST_UNIT).getValue();
      if (qty !== '' && cost !== '') {
        sh.getRange(row, COL_SR_TOTAL_COST)
          .setFormula('=C' + row + '*D' + row)
          .setNumberFormat('"$"#,##0.00');
      }
    }

  } catch(err) {
    // onEdit errors should be silent — do not alert
  }
}
