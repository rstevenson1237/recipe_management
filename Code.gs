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
      .createMenu('Recipe Tools')
      .addItem('Open Recipe UI', 'openRecipeApp')
      .addItem('Export All Recipes (Print)', 'openExportPreview')
      .addItem('Import Recipes from PDF...', 'openImportDialog')
      .addItem('Setup / Repair Dashboard', 'setupWorkbook')
      .addItem('Clear ALL Recipes...', 'resetAllRecipes')
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
  template.recipeNamesJson = JSON.stringify(getRecipeNameList()).replace(/</g, '\\u003c');

  var html = template.evaluate()
      .setWidth(1100)
      .setHeight(800)
      .setTitle('Recipe Management Interface');

  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * Shows the PDF import dialog. The PDF is read and parsed entirely in the browser
 * (pdf.js, loaded from a CDN by Import.html) and only the parsed recipes are sent to
 * importRecipesFromPdf() - so no Drive scope, advanced service or manifest is needed,
 * and a large PDF can't spend the server's execution budget on text extraction.
 *
 * Helper Data is inlined the same way openRecipeApp() does it: the review step flags
 * ingredients and units this workbook doesn't know about yet, which needs those lists
 * on hand before the user picks a file.
 */
function openImportDialog() {
  ensureWorkbookInitialized_();

  var template = HtmlService.createTemplateFromFile('Import');
  template.helperDataJson = JSON.stringify(getHelperData()).replace(/</g, '\\u003c');
  template.recipeNamesJson = JSON.stringify(getRecipeNameList()).replace(/</g, '\\u003c');

  var html = template.evaluate()
      .setWidth(900)
      .setHeight(700)
      .setTitle('Import Recipes from PDF');

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
    '6. Use 🍔 Recipe Tools > Import Recipes from PDF... to load every recipe out of a PDF that export produced.',
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
 * Full resync of the Dashboard's recipe list from every DB sheet set on file -
 * clears the table and rewrites it sorted alphabetically. This is the expensive
 * path (it rescans every recipe ever entered), so it's reserved for first-time
 * setup and the "Setup / Repair Dashboard" menu item. A normal save uses
 * appendDashboardRow_() instead, which only writes the one new row.
 */
function refreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!dashboard) return;

  clearDashboardRows_(dashboard);
  var startRow = findDashboardHeaderRow_(dashboard) + 1;

  var recipes = exportAllRecipes();
  if (recipes.length === 0) return;

  var rows = recipes.map(function(recipe) {
    return buildDashboardRow_({
      name: recipe.name,
      measureType: recipe.measureType,
      yieldQty: recipe.yieldQty,
      yieldUom: recipe.yieldUom,
      portionSize: recipe.portionSize,
      portionUom: recipe.portionUom,
      ingredientCount: recipe.ingredients.length,
      stepCount: recipe.steps.length,
      sourceKey: recipe.sourceKey
    });
  });

  dashboard.getRange(startRow, 1, rows.length, DASHBOARD_TABLE_COLUMNS.length).setValues(rows);
}

/**
 * Clears every recipe row from the Dashboard's table, leaving the title, instructions
 * and table header untouched. Shared by the full refresh and by resetAllRecipes(), so
 * "which rows belong to the table" is decided in exactly one place.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboard
 */
function clearDashboardRows_(dashboard) {
  var startRow = findDashboardHeaderRow_(dashboard) + 1;
  var lastRow = dashboard.getLastRow();
  if (lastRow >= startRow) {
    dashboard.getRange(startRow, 1, lastRow - startRow + 1, DASHBOARD_TABLE_COLUMNS.length).clearContent();
  }
}

/**
 * Appends a single recipe to the Dashboard's recipe table, in the next available
 * (first empty) row after the header - the fast path called after every save
 * instead of rebuilding the whole table. New recipes land in save order, not
 * resorted alphabetically; run "Setup / Repair Dashboard" to re-sort if desired.
 * @param {DashboardRowEntry} entry
 */
function appendDashboardRow_(entry) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!dashboard) return;

  var startRow = findDashboardHeaderRow_(dashboard) + 1;
  var nextRow = nextDashboardRow_(dashboard, startRow);
  dashboard.getRange(nextRow, 1, 1, DASHBOARD_TABLE_COLUMNS.length).setValues([buildDashboardRow_(entry)]);
}

/**
 * Appends many recipes to the Dashboard's recipe table in one batched write - the PDF
 * import path, where calling appendDashboardRow_() per recipe would be one round trip
 * to the Sheets service per recipe and could push a large import past the execution
 * limit on its own. No-op for an empty list.
 * @param {DashboardRowEntry[]} entries
 */
function appendDashboardRows_(entries) {
  if (entries.length === 0) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!dashboard) return;

  var startRow = findDashboardHeaderRow_(dashboard) + 1;
  var nextRow = nextDashboardRow_(dashboard, startRow);
  dashboard.getRange(nextRow, 1, entries.length, DASHBOARD_TABLE_COLUMNS.length)
      .setValues(entries.map(function(entry) { return buildDashboardRow_(entry); }));
}

/**
 * Locates the recipe table's actual header row by searching column A for a cell
 * that reads exactly "Name" with "Measure Type" immediately to its right (column
 * B) - i.e. the real header, not just any cell containing the word "Name". Never
 * trust DASHBOARD_TABLE_HEADER_ROW alone at runtime: rows inserted or deleted
 * above the table (by hand, or by an older version of this script) shift the
 * header to a different physical row without updating that constant, which would
 * otherwise make new recipes land far below the real table with a blank gap in
 * between. Falls back to the constant only when no header exists yet (a sheet
 * that hasn't been through renderDashboardLayout_() at all).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboard
 * @return {number}
 */
function findDashboardHeaderRow_(dashboard) {
  var matches = dashboard.createTextFinder(DASHBOARD_TABLE_COLUMNS[0]).matchEntireCell(true).matchCase(true).findAll();
  for (var i = 0; i < matches.length; i++) {
    var cell = matches[i];
    if (cell.getColumn() === 1 && cell.getSheet().getRange(cell.getRow(), 2).getValue() === DASHBOARD_TABLE_COLUMNS[1]) {
      return cell.getRow();
    }
  }
  return DASHBOARD_TABLE_HEADER_ROW;
}

/**
 * Finds the next available row for a new recipe: the first empty cell in the
 * Dashboard's Name column (A) at or after startRow. Reads the whole column in one
 * batched call rather than trusting Sheet.getLastRow(), which would be thrown off
 * by unrelated content elsewhere on the sheet (e.g. the button-placeholder box).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboard
 * @param {number} startRow
 * @return {number}
 */
function nextDashboardRow_(dashboard, startRow) {
  var maxRows = dashboard.getMaxRows();
  if (maxRows < startRow) return startRow;

  var nameColumn = dashboard.getRange(startRow, 1, maxRows - startRow + 1, 1).getValues();
  for (var i = 0; i < nameColumn.length; i++) {
    if (nameColumn[i][0] === '' || nameColumn[i][0] === null) return startRow + i;
  }
  return startRow + nameColumn.length;
}

/**
 * Updates an existing Dashboard row in place for a recipe that was just edited via
 * updateRecipeFromWeb() - matched by its pre-edit name and source sheet-set key
 * (the name itself may have just changed, so it can't be part of the match). Falls
 * back to appending a new row if no matching row is found, e.g. the Dashboard was
 * rebuilt or hand-edited since the recipe was saved.
 * @param {string} oldName
 * @param {string} sourceKey
 * @param {DashboardRowEntry} entry
 */
function updateDashboardRow_(oldName, sourceKey, entry) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!dashboard) return;

  var startRow = findDashboardHeaderRow_(dashboard) + 1;
  var lastRow = dashboard.getLastRow();

  if (lastRow >= startRow) {
    var normalizedOldName = String(oldName).trim().toLowerCase();
    var rows = dashboard.getRange(startRow, 1, lastRow - startRow + 1, DASHBOARD_TABLE_COLUMNS.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === normalizedOldName && String(rows[i][7]) === sourceKey) {
        dashboard.getRange(startRow + i, 1, 1, DASHBOARD_TABLE_COLUMNS.length).setValues([buildDashboardRow_(entry)]);
        return;
      }
    }
  }

  appendDashboardRow_(entry);
}

/**
 * @typedef {Object} DashboardRowEntry
 * @property {string} name
 * @property {*} measureType
 * @property {*} yieldQty
 * @property {*} yieldUom
 * @property {*} portionSize
 * @property {*} portionUom
 * @property {number} ingredientCount
 * @property {number} stepCount
 * @property {string} sourceKey
 */

/**
 * Builds one Dashboard table row (matching DASHBOARD_TABLE_COLUMNS order) from
 * either a full Recipe (refreshDashboard) or the payload just saved (appendDashboardRow_).
 * @param {DashboardRowEntry} entry
 * @return {Array<*>}
 */
function buildDashboardRow_(entry) {
  return [
    entry.name,
    entry.measureType,
    formatQtyUom_(entry.yieldQty, entry.yieldUom),
    formatQtyUom_(entry.portionSize, entry.portionUom),
    entry.ingredientCount,
    entry.stepCount,
    dateKeyToDisplay_(entry.sourceKey),
    entry.sourceKey
  ];
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
