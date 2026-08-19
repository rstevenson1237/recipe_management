/**
 * 1. RUN THIS ONCE TO SETUP THE MENU
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🍔 Recipe Tools')
      .addItem('Open Recipe UI', 'openRecipeApp')
      .addToUi();
}

/**
 * 2. LAUNCHES THE VUE.JS MODAL
 */
function openRecipeApp() {
  // Ensure the database tabs exist so it doesn't crash on a blank workbook
  ensureSheetsExist(); 
  
  const html = HtmlService.createHtmlOutputFromFile('Index')
      .setWidth(1000)
      .setHeight(750)
      .setTitle('Recipe Management Interface');
  
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * 3. FETCHES ALLOWED INGREDIENTS AND UOM MATRIX FOR VUE
 */
function getHelperData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const helperSheet = ss.getSheetByName('Helper Data');
  
  if (!helperSheet) return {};
  
  const lastRow = helperSheet.getLastRow();
  if (lastRow < 2) return {};
  
  // Assuming 'Helper Data' looks like our previous POC:
  // Column D (index 4) = Ingredient Name
  // Columns E-N (index 5-14) = Valid UOMs
  const data = helperSheet.getRange(2, 4, lastRow - 1, 11).getValues();
  const helperDict = {};
  
  data.forEach(row => {
    const ingName = row[0];
    if (ingName) {
      // Filter out empty cells to leave only valid UOM strings
      const validUoms = row.slice(1).filter(String); 
      helperDict[ingName] = validUoms;
    }
  });
  
  return helperDict;
}

/**
 * 4. RECEIVES JSON FROM VUE AND SAVES TO SPREADSHEET
 */
function saveRecipeFromWeb(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbHeaders = ss.getSheetByName("DB - Headers");
  const dbIngredients = ss.getSheetByName("DB - Ingredients");
  
  // 1. Append the Headers (Recipe Metadata)
  dbHeaders.appendRow(payload.headerData);
  
  // 2. Append the Ingredients
  const recipeName = payload.headerData[0]; // Recipe Name is the first item in the header array
  
  payload.ingredients.forEach(function(ing, index) {
    if (ing.name.trim() !== "") {
      // Maps to: Name, Ingredient, Qty, U of M, Order, Yield Percent, Blank, Helper1, Helper2
      const ingredientRow = [
        recipeName,
        ing.name,
        ing.qty,
        ing.uom,
        index + 1, // Ingredient Order
        ing.yieldPercent,
        "", "", ""
      ];
      dbIngredients.appendRow(ingredientRow);
    }
  });
  
  return true; // Send success signal back to Vue
}

/**
 * Utility: Checks for hidden DB sheets and makes them if missing
 */
function ensureSheetsExist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const required = ["DB - Headers", "DB - Ingredients", "Helper Data"];
  
  required.forEach(name => {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name).hideSheet();
    }
  });
}
