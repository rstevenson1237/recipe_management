/**
 * Code.gs
 * Menu wiring, dialog launchers, and the Dashboard sheet (layout + refresh).
 */

var DASHBOARD_TABLE_HEADER_ROW = 16;
var DASHBOARD_TABLE_COLUMNS = [
  'Name', 'Measure Type', 'Yield', 'Portion', 'Ingredients', 'Steps', 'Date Entered', 'Source Sheet'
];

/**
 * 1. RUN THIS ONCE TO SETUP THE MENU
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('🍔 Recipe Tools')
      .addItem('Open Recipe UI', 'openRecipeApp')
      .addItem('Export All Recipes (Print)', 'openExportPreview')
      .addSeparator()
      .addItem('Setup / Repair Dashboard', 'setupWorkbook')
      .addToUi();
}

/**
 * 2. LAUNCHES THE VUE.JS MODAL
 */
function openRecipeApp() {
  // Defensive: guarantees the Dashboard and Helper Data exist so this never
  // crashes on a blank workbook, without re-running full setup on every open.
  ensureWorkbookInitialized_();

  // Helper Data is inlined into the page here (instead of the dialog fetching it
  // over google.script.run after showing a loading spinner) so the form is
  // interactive the instant the modal renders, with no extra client-server round trip.
  var template = HtmlService.createTemplateFromFile('Index');
  template.helperDataJson = JSON.stringify(getHelperData()).replace(/</g, '\\u003c');

  var html = template.evaluate()
      .setWidth(1100)
      .setHeight(800)
      .setTitle('Recipe Management Interface');

  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * One-time (or re-run-to-repair) workbook setup: builds the Dashboard layout,
 * seeds Helper Data, and refreshes the recipe list.
 */
function setupWorkbook() {
  var dashboard = getOrCreateDashboardSheet_();
  renderDashboardLayout_(dashboard);
  ensureHelperDataSheet_();
  refreshDashboard();
  SpreadsheetApp.getActiveSpreadsheet().toast('Dashboard is ready.', '🍔 Recipe Tools');
}

/**
 * Cheap idempotent check used before showing the entry dialog: creates the
 * Dashboard/Helper Data only if they don't exist yet.
 */
function ensureWorkbookInitialized_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName(DASHBOARD_SHEET_NAME);

  if (!dashboard) {
    dashboard = getOrCreateDashboardSheet_();
    renderDashboardLayout_(dashboard);
    refreshDashboard();
  }

  ensureHelperDataSheet_();
}

function getOrCreateDashboardSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName(DASHBOARD_SHEET_NAME);

  if (!dashboard) {
    dashboard = ss.insertSheet(DASHBOARD_SHEET_NAME, 0);
  } else if (ss.getSheets()[0].getSheetId() !== dashboard.getSheetId()) {
    ss.setActiveSheet(dashboard);
    ss.moveActiveSheet(1);
  }

  return dashboard;
}

/**
 * Writes the static parts of the Dashboard: title, usage instructions, the
 * one-time "insert your button here" box, and the recipe-table header row.
 * Safe to re-run - it never touches the recipe rows written by refreshDashboard().
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function renderDashboardLayout_(sheet) {
  sheet.getRange('A1:H1').merge()
      .setValue('🍔 Recipe Management Dashboard')
      .setFontSize(18).setFontWeight('bold')
      .setBackground('#1e293b').setFontColor('#ffffff')
      .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  sheet.getRange('A3').setValue('How to use this workbook').setFontWeight('bold').setFontSize(12);

  var instructions = [
    '1. Click "New Recipe" (button at right) or use the 🍔 Recipe Tools menu > Open Recipe UI to add a recipe.',
    '2. Fill in every section of the dialog - fields are validated as you type, and the Save button unlocks once everything is valid.',
    '3. Ingredients and Instructions are checked against the "Helper Data" sheet - ask the workbook owner to add missing items there.',
    '4. Every recipe you save appears in the table below automatically.',
    '5. Use 🍔 Recipe Tools > Export All Recipes (Print) to generate a print-ready, two-column export of every recipe on file.',
    '',
    'One-time setup for the button at right (if it is not already there):',
    'Insert > Drawing, draw a "New Recipe" button, click Save and Close, then click the drawing once more, open its ⋮ menu,',
    'choose "Assign script", and enter:  openRecipeApp'
  ].join('\n');

  sheet.getRange('A4:H13').merge()
      .setValue(instructions)
      .setWrap(true)
      .setVerticalAlignment('top')
      .setFontSize(10);

  sheet.getRange('J2:M8').merge()
      .setValue('⬅ Insert your "New Recipe" button drawing here.\n\nSee the setup note at left for how to assign it to openRecipeApp.')
      .setWrap(true)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .setFontStyle('italic')
      .setFontColor('#64748b')
      .setBorder(true, true, true, true, false, false, '#94a3b8', SpreadsheetApp.BorderStyle.DASHED);

  sheet.getRange(DASHBOARD_TABLE_HEADER_ROW, 1, 1, DASHBOARD_TABLE_COLUMNS.length)
      .setValues([DASHBOARD_TABLE_COLUMNS])
      .setFontWeight('bold')
      .setBackground('#e2e8f0');
  sheet.setFrozenRows(DASHBOARD_TABLE_HEADER_ROW);

  sheet.setColumnWidth(1, 200);
  for (var col = 2; col <= 8; col++) sheet.setColumnWidth(col, 110);
}

/**
 * Rebuilds the recipe list on the Dashboard from every DB sheet set on file.
 * Called after every successful save and from the setup menu item.
 */
function refreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!dashboard) return;

  var startRow = DASHBOARD_TABLE_HEADER_ROW + 1;
  var lastRow = dashboard.getLastRow();
  if (lastRow >= startRow) {
    dashboard.getRange(startRow, 1, lastRow - startRow + 1, DASHBOARD_TABLE_COLUMNS.length).clearContent();
  }

  var recipes = exportAllRecipes();
  if (recipes.length === 0) return;

  var rows = recipes.map(function(recipe) {
    return [
      recipe.name,
      recipe.measureType,
      formatQtyUom_(recipe.yieldQty, recipe.yieldUom),
      formatQtyUom_(recipe.portionSize, recipe.portionUom),
      recipe.ingredients.length,
      recipe.steps.length,
      dateKeyToDisplay_(recipe.sourceKey),
      recipe.sourceKey
    ];
  });

  dashboard.getRange(startRow, 1, rows.length, DASHBOARD_TABLE_COLUMNS.length).setValues(rows);
}

/**
 * @param {*} qty
 * @param {*} uom
 * @return {string}
 */
function formatQtyUom_(qty, uom) {
  if (qty === '' || qty === null || qty === undefined) return '';
  return uom ? qty + ' ' + uom : String(qty);
}

/**
 * Turns a sheet-set key like "190826" or "190826 (2)" back into a readable date
 * (MM/DD/YYYY), falling back to the raw key if it doesn't parse as DDMMYY.
 * @param {string} key
 * @return {string}
 */
function dateKeyToDisplay_(key) {
  var match = /^(\d{2})(\d{2})(\d{2})/.exec(key);
  if (!match) return key;

  var day = match[1], month = match[2], year = '20' + match[3];
  return month + '/' + day + '/' + year;
}
