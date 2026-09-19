/**
 * Reset.gs
 * The "Clear ALL Recipes" menu action: deletes every date-keyed DB sheet set and empties
 * the Dashboard's recipe table, returning the workbook to the state it had before any
 * recipe was entered.
 *
 * Helper Data is deliberately NOT touched - it is the workbook owner's curated list of
 * allowed ingredients and units, not recipe data, and re-typing it is the one part of
 * setup that can't be regenerated.
 */

/**
 * Deletes every recipe in the workbook, behind two confirmations. Irreversible: the
 * hidden DB sheets are removed outright, so the first dialog states exactly what will
 * go and the second requires the word DELETE to be typed. Anything else cancels.
 */
function resetAllRecipes() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheetNames = listRecipeSheetNames_(ss);
  if (sheetNames.length === 0) {
    ui.alert('Clear ALL Recipes', 'There are no recipe sheets to clear - this workbook is already empty.', ui.ButtonSet.OK);
    return;
  }

  var recipeCount = getRecipeNameList().length;
  var setCount = sheetNames.filter(function(name) { return name.indexOf(HEADER_PREFIX) === 0; }).length;

  var confirmed = ui.alert(
      'Clear ALL Recipes',
      'This permanently deletes ' + recipeCount + ' recipe(s) across ' + setCount + ' batch(es):\n\n' +
      '  • ' + sheetNames.length + ' hidden "DB - ..." sheets are deleted\n' +
      '  • the Dashboard recipe list is emptied (its layout and instructions stay)\n' +
      '  • "Helper Data" is NOT touched - your ingredient and U of M lists are kept\n\n' +
      'This cannot be undone. Continue?',
      ui.ButtonSet.YES_NO);
  if (confirmed !== ui.Button.YES) return;

  var typed = ui.prompt(
      'Confirm - this cannot be undone',
      'Type DELETE (in capitals) to erase all ' + recipeCount + ' recipe(s):',
      ui.ButtonSet.OK_CANCEL);
  if (typed.getSelectedButton() !== ui.Button.OK || typed.getResponseText().trim() !== 'DELETE') {
    ss.toast('Nothing was deleted.', '🍔 Recipe Tools');
    return;
  }

  // Same lock the save path takes, so a reset can't interleave with a save in progress
  // and leave a half-written recipe pointing at sheets that no longer exist.
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  var deleted = 0;
  try {
    // Re-read the names under the lock: a save may have created a new set for today
    // between the confirmation dialogs and now.
    listRecipeSheetNames_(ss).forEach(function(name) {
      var sheet = ss.getSheetByName(name);
      if (sheet) {
        ss.deleteSheet(sheet);
        deleted++;
      }
    });

    var dashboard = ss.getSheetByName(DASHBOARD_SHEET_NAME);
    if (dashboard) clearDashboardRows_(dashboard);
  } finally {
    lock.releaseLock();
  }

  ss.toast('Deleted ' + recipeCount + ' recipe(s) and ' + deleted + ' sheet(s). Helper Data was kept.',
      '🍔 Recipe Tools');
}

/**
 * Names of every recipe-data sheet in the workbook: the Headers/Ingredients/Instructions
 * sheets of every date-keyed set, and nothing else. Dashboard and Helper Data can never
 * match these prefixes, which is what guarantees the workbook keeps a visible sheet (a
 * spreadsheet must always have one) and its configuration after a reset.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {string[]}
 */
function listRecipeSheetNames_(ss) {
  var prefixes = [HEADER_PREFIX, INGREDIENTS_PREFIX, INSTRUCTIONS_PREFIX];

  return ss.getSheets().map(function(sheet) {
    return sheet.getName();
  }).filter(function(name) {
    return prefixes.some(function(prefix) { return name.indexOf(prefix) === 0; });
  });
}
