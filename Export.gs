/**
 * Export.gs
 * Gathers every recipe ever entered (across all date-keyed sheet sets) and renders
 * the print-ready HTML export.
 */

/**
 * Version of the printed export's layout and embedded data block. Stamped in the page
 * footer of every exported recipe so the PDF import can tell what it is looking at:
 * v1 has no Preparation column and no data block, v2 has both. Bump this whenever a
 * field is added to what the export prints.
 */
var EXPORT_FORMAT_VERSION = 2;

/**
 * Shows the export preview in a modal dialog with a Print button. The user prints
 * or saves-to-PDF straight from the browser's print dialog - nothing is written
 * to Drive.
 */
function openExportPreview() {
  var template = HtmlService.createTemplateFromFile('Export');
  template.recipes = exportAllRecipes();
  template.logoDataUri = getLogoDataUri_();
  template.formatVersion = EXPORT_FORMAT_VERSION;

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
 * @property {*} prep - free-text preparation method ('minced'), '' when none.
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
          .map(function(r) { return { name: r[1], qty: r[2], uom: r[3], prep: r[5] }; });

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

/**
 * Builds the machine-readable block printed at the bottom of a recipe's exported page,
 * so that an exported PDF can be imported back losslessly instead of having to be
 * re-parsed out of the printed layout (where "Chives, minced" is ambiguous with an
 * ingredient whose name simply contains a comma).
 *
 * Format: RMDATA1[<base64 of a compact JSON object>]. "RMDATA1" versions the envelope;
 * the JSON's own "v" field carries EXPORT_FORMAT_VERSION. The delimiters are safe
 * because base64's alphabet contains no square brackets, and keys are abbreviated
 * purely to keep the printed block short.
 * @param {Recipe} recipe
 * @return {string}
 */
function buildRecipeDataBlock_(recipe) {
  var payload = {
    v: EXPORT_FORMAT_VERSION,
    n: recipe.name,
    mt: blank_(recipe.measureType),
    ru: blank_(recipe.reportingUom),
    yq: blank_(recipe.yieldQty),
    yu: blank_(recipe.yieldUom),
    wq: blank_(recipe.weightQty),
    wu: blank_(recipe.weightUom),
    vq: blank_(recipe.volumeQty),
    vu: blank_(recipe.volumeUom),
    eq: blank_(recipe.eachQty),
    eu: blank_(recipe.eachUom),
    ps: blank_(recipe.portionSize),
    pu: blank_(recipe.portionUom),
    inv: blank_(recipe.availableInInventory),
    iu: blank_(recipe.inventoryUom),
    // Ingredients as positional arrays rather than objects: [name, qty, uom, prep].
    ing: recipe.ingredients.map(function(ing) {
      return [blank_(ing.name), blank_(ing.qty), blank_(ing.uom), blank_(ing.prep)];
    }),
    st: recipe.steps.map(function(step) { return blank_(step); })
  };

  return 'RMDATA1[' + Utilities.base64Encode(JSON.stringify(payload), Utilities.Charset.UTF_8) + ']';
}

/**
 * Normalizes a cell value for the data block: null/undefined become '', everything
 * else is left as the number or string the sheet gave us so quantities round-trip as
 * numbers rather than as strings.
 * @param {*} value
 * @return {*}
 */
function blank_(value) {
  return (value === null || value === undefined) ? '' : value;
}
