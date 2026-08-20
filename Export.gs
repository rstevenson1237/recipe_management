/**
 * Export.gs
 * Gathers every recipe ever entered (across all date-keyed sheet sets) and renders
 * the print-ready HTML export.
 */

/**
 * Shows the export preview in a modal dialog with a Print button. The user prints
 * or saves-to-PDF straight from the browser's print dialog - nothing is written
 * to Drive.
 */
function openExportPreview() {
  var template = HtmlService.createTemplateFromFile('Export');
  template.recipes = exportAllRecipes();
  template.logoDataUri = getLogoDataUri_();

  var html = template.evaluate()
      .setWidth(1000)
      .setHeight(800)
      .setTitle('Recipe Export Preview');

  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function getLogoDataUri_() {
  var base64 = HtmlService.createHtmlOutputFromFile('Logo').getContent().trim();
  return 'data:image/png;base64,' + base64;
}

/**
 * @typedef {Object} RecipeIngredient
 * @property {*} name
 * @property {*} qty
 * @property {*} uom
 */

/**
 * @typedef {Object} Recipe
 * @property {string} name
 * @property {*} measureType
 * @property {*} reportingUom
 * @property {*} yieldQty
 * @property {*} yieldUom
 * @property {*} weightQty
 * @property {*} weightUom
 * @property {*} volumeQty
 * @property {*} volumeUom
 * @property {*} eachQty
 * @property {*} eachUom
 * @property {*} portionSize
 * @property {*} portionUom
 * @property {*} availableInInventory
 * @property {*} inventoryUom
 * @property {RecipeIngredient[]} ingredients
 * @property {*[]} steps
 * @property {string} sourceKey
 */

/**
 * Collects every recipe across every DB sheet set, merges its Ingredients and
 * Instructions rows in, and returns them sorted alphabetically by name.
 * @return {Recipe[]}
 */
function exportAllRecipes() {
  var sets = listAllSheetSets_();
  /** @type {Recipe[]} */
  var recipes = [];

  sets.forEach(function(set) {
    var headerRows = getDataRows_(set.headers, HEADER_COLUMNS.length);
    var ingredientRows = getDataRows_(set.ingredients, INGREDIENT_COLUMNS.length);
    var instructionRows = getDataRows_(set.instructions, INSTRUCTION_COLUMNS.length);

    headerRows.forEach(function(row) {
      var name = String(row[0]).trim();
      if (!name) return;

      var ingredients = ingredientRows
          .filter(function(r) { return String(r[0]).trim() === name; })
          .sort(function(a, b) { return Number(a[4]) - Number(b[4]); })
          .map(function(r) { return { name: r[1], qty: r[2], uom: r[3] }; });

      var steps = instructionRows
          .filter(function(r) { return String(r[0]).trim() === name; })
          .sort(function(a, b) { return Number(a[1]) - Number(b[1]); })
          .map(function(r) { return r[2]; });

      recipes.push({
        name: name,
        measureType: row[1],
        reportingUom: row[2],
        yieldQty: row[3],
        yieldUom: row[4],
        weightQty: row[5],
        weightUom: row[6],
        volumeQty: row[7],
        volumeUom: row[8],
        eachQty: row[9],
        eachUom: row[10],
        portionSize: row[11],
        portionUom: row[12],
        availableInInventory: row[13],
        inventoryUom: row[14],
        ingredients: ingredients,
        steps: steps,
        sourceKey: set.key
      });
    });
  });

  recipes.sort(function(a, b) {
    return a.name.toLowerCase().localeCompare(b.name.toLowerCase());
  });

  return recipes;
}
