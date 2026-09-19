/**
 * Import.gs
 * Server side of the PDF import: takes the recipes Import.html parsed out of an exported
 * PDF in the browser and writes them into a dedicated sheet set.
 *
 * Deliberately NOT routed through getActiveSheetSet(): that applies resolveActiveKey_()'s
 * 25-recipe rollover, and an import must be able to land a PDF of any size in one place.
 * Imports get their own "IMPORT <DDMMYY>" key instead, holding every recipe in the file
 * however many there are.
 */

var IMPORT_KEY_PREFIX = 'IMPORT ';

/**
 * @typedef {Object} ImportSummary
 * @property {boolean} ok
 * @property {number} imported
 * @property {string} sheetKey
 * @property {Array<{name: string, reason: string}>} skipped
 * @property {{items: string[], uoms: string[]}} helperDataAdded
 */

/**
 * Writes parsed recipes into a new "IMPORT <key>" sheet set and appends them to the
 * Dashboard. A recipe whose name already exists anywhere in the workbook is skipped and
 * reported rather than overwritten, because Name is the join key across all three DB
 * sheets - two recipes sharing one would merge into each other's ingredient lists.
 *
 * Every write is batched (one setValues per sheet, one for the Dashboard) so that a PDF
 * with hundreds of recipes stays inside the Apps Script execution limit.
 * @param {*} recipes - array of recipe objects built by Import.html; shape isn't trusted.
 * @return {ImportSummary}
 */
function importRecipesFromPdf(recipes) {
  if (!recipes || !Array.isArray(recipes) || recipes.length === 0) {
    throw new Error('No recipes were sent to import.');
  }

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    var helperData = getHelperData();

    /** @type {Object<string, boolean>} */
    var takenNames = {};
    getRecipeNameList().forEach(function(name) { takenNames[name.trim().toLowerCase()] = true; });

    /** @type {Array<{name: string, reason: string}>} */
    var skipped = [];
    /** @type {Array<*>} */
    var accepted = [];

    recipes.forEach(function(/** @type {*} */ recipe) {
      var name = recipe && recipe.name ? String(recipe.name).trim() : '';
      if (!name) {
        skipped.push({ name: '(unnamed)', reason: 'No recipe name could be read from the PDF.' });
        return;
      }
      var key = name.toLowerCase();
      if (takenNames[key]) {
        // Covers both "already in the workbook" and "appeared twice in this PDF" - the
        // accepted names are added to the same map as they're taken.
        skipped.push({ name: name, reason: 'A recipe with this name already exists - left untouched.' });
        return;
      }
      takenNames[key] = true;
      accepted.push(recipe);
    });

    if (accepted.length === 0) {
      return { ok: true, imported: 0, sheetKey: '', skipped: skipped, helperDataAdded: { items: [], uoms: [] } };
    }

    var helperDataAdded = addUnknownHelperData_(accepted, helperData);

    var set = ensureSheetSetComplete_(resolveImportKey_());
    var startCounts = {
      headers: set.headers.getLastRow(),
      ingredients: set.ingredients.getLastRow(),
      instructions: set.instructions.getLastRow()
    };

    /** @type {Array<Array<*>>} */
    var headerRows = [];
    /** @type {Array<Array<*>>} */
    var ingredientRows = [];
    /** @type {Array<Array<*>>} */
    var instructionRows = [];
    /** @type {DashboardRowEntry[]} */
    var dashboardEntries = [];

    accepted.forEach(function(/** @type {*} */ recipe) {
      var name = String(recipe.name).trim();
      var yieldUom = importString_(recipe.yieldUom);
      var cleanIngredients = importIngredients_(recipe);
      var cleanSteps = importSteps_(recipe);

      headerRows.push([
        name,
        importString_(recipe.measureType),
        importString_(recipe.reportingUom) || yieldUom,
        importNumber_(recipe.yieldQty),
        yieldUom,
        importNumber_(recipe.weightQty),
        importString_(recipe.weightUom),
        importNumber_(recipe.volumeQty),
        importString_(recipe.volumeUom),
        importNumber_(recipe.eachQty),
        importString_(recipe.eachUom),
        importNumber_(recipe.portionSize),
        importString_(recipe.portionUom),
        importString_(recipe.inventoryYesNo) === 'Yes' ? 'Yes' : 'No',
        importString_(recipe.inventoryUom)
      ]);

      cleanIngredients.forEach(function(ing, index) {
        ingredientRows.push([name, ing.name, ing.qty, ing.uom, index + 1, ing.prep]);
      });
      cleanSteps.forEach(function(text, index) {
        instructionRows.push([name, index + 1, text]);
      });

      dashboardEntries.push({
        name: name,
        measureType: importString_(recipe.measureType),
        yieldQty: importNumber_(recipe.yieldQty),
        yieldUom: yieldUom,
        portionSize: importNumber_(recipe.portionSize),
        portionUom: importString_(recipe.portionUom),
        ingredientCount: cleanIngredients.length,
        stepCount: cleanSteps.length,
        sourceKey: set.key
      });
    });

    try {
      appendRows_(set.ingredients, ingredientRows);
      appendRows_(set.instructions, instructionRows);
      appendRows_(set.headers, headerRows);
    } catch (writeError) {
      // Same rollback as saveRecipeFromWeb: a partial write would leave the trio out of
      // lock step, with ingredient rows pointing at recipes that have no header row.
      truncateSheetToRow_(set.headers, startCounts.headers);
      truncateSheetToRow_(set.ingredients, startCounts.ingredients);
      truncateSheetToRow_(set.instructions, startCounts.instructions);
      throw writeError;
    }

    appendDashboardRows_(dashboardEntries);

    return {
      ok: true,
      imported: accepted.length,
      sheetKey: set.key,
      skipped: skipped,
      helperDataAdded: helperDataAdded
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Key for this import's sheet set: "IMPORT <DDMMYY>", suffixed "(2)", "(3)", ... only if
 * that exact set already exists from an earlier import today. Unlike resolveActiveKey_()
 * this never rolls over on recipe count, which is what lets one import hold more than
 * MAX_RECIPES_PER_SET recipes.
 * @return {string}
 */
function resolveImportKey_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var baseKey = IMPORT_KEY_PREFIX + formatDateKey_(new Date());

  var suffixIndex = 1;
  while (ss.getSheetByName(HEADER_PREFIX + buildSetKey_(baseKey, suffixIndex))) {
    suffixIndex++;
  }
  return buildSetKey_(baseKey, suffixIndex);
}

/**
 * Adds every ingredient name and unit of measure in the incoming recipes that this
 * workbook's Helper Data doesn't list yet. Without this, a recipe imported from another
 * workbook would open in the entry dialog with every ingredient flagged red and couldn't
 * be saved again without hand-editing Helper Data first.
 *
 * An unrecognized unit has no category of its own to go by, so it's filed under the
 * recipe's Measure Type (defaulting to Each), which is the column the dialog will look
 * for it in.
 * @param {Array<*>} recipes
 * @param {HelperData} helperData
 * @return {{items: string[], uoms: string[]}}
 */
function addUnknownHelperData_(recipes, helperData) {
  /** @type {Object<string, boolean>} */
  var knownItems = {};
  helperData.items.forEach(function(item) { knownItems[item.trim().toLowerCase()] = true; });

  /** @type {Object<string, boolean>} */
  var knownUoms = {};
  /** @type {Array<keyof HelperData['uom']>} */
  var categories = ['Weight', 'Volume', 'Each'];
  categories.forEach(function(category) {
    helperData.uom[category].forEach(function(/** @type {string} */ uom) {
      knownUoms[uom.trim().toLowerCase()] = true;
    });
  });

  /** @type {string[]} */
  var newItems = [];
  /** @type {{Weight: string[], Volume: string[], Each: string[]}} */
  var newUoms = { Weight: [], Volume: [], Each: [] };

  recipes.forEach(function(/** @type {*} */ recipe) {
    var measureType = importString_(recipe.measureType);
    /** @type {keyof HelperData['uom']} */
    var fallbackCategory = 'Each';
    if (measureType === 'Weight' || measureType === 'Volume') fallbackCategory = measureType;

    /**
     * @param {string} uom
     */
    function noteUom(uom) {
      if (!uom || knownUoms[uom.toLowerCase()]) return;
      knownUoms[uom.toLowerCase()] = true;
      newUoms[fallbackCategory].push(uom);
    }

    noteUom(importString_(recipe.yieldUom));
    noteUom(importString_(recipe.portionUom));
    if (importString_(recipe.inventoryYesNo) === 'Yes') noteUom(importString_(recipe.inventoryUom));

    importIngredients_(recipe).forEach(function(ing) {
      if (ing.name && !knownItems[ing.name.toLowerCase()]) {
        knownItems[ing.name.toLowerCase()] = true;
        newItems.push(ing.name);
      }
      noteUom(ing.uom);
    });
  });

  if (newItems.length === 0 && newUoms.Weight.length === 0 && newUoms.Volume.length === 0 && newUoms.Each.length === 0) {
    return { items: [], uoms: [] };
  }
  return appendHelperDataItems_(newItems, newUoms);
}

/**
 * @typedef {Object} ImportedIngredient
 * @property {string} name
 * @property {*} qty
 * @property {string} uom
 * @property {string} prep
 */

/**
 * The ingredient rows of an incoming recipe, trimmed and with unnamed rows dropped -
 * the same "a row without a name isn't an ingredient" rule the save path applies.
 * @param {*} recipe
 * @return {ImportedIngredient[]}
 */
function importIngredients_(recipe) {
  var ingredients = Array.isArray(recipe.ingredients) ? recipe.ingredients : [];

  return ingredients.map(function(/** @type {*} */ ing) {
    return {
      name: importString_(ing && ing.name),
      qty: importNumber_(ing && ing.qty),
      uom: importString_(ing && ing.uom),
      prep: cleanPreparation_(ing && ing.prep)
    };
  }).filter(function(/** @type {ImportedIngredient} */ ing) {
    return ing.name !== '';
  });
}

/**
 * The instruction texts of an incoming recipe, in order, with blanks dropped. Accepts
 * either plain strings (what the export's data block holds) or {text} objects (the shape
 * the entry dialog uses), so both parse paths in Import.html can hand over their steps
 * without converting first.
 * @param {*} recipe
 * @return {string[]}
 */
function importSteps_(recipe) {
  var steps = Array.isArray(recipe.steps) ? recipe.steps
      : (Array.isArray(recipe.instructions) ? recipe.instructions : []);

  return steps.map(function(/** @type {*} */ step) {
    return importString_(step && step.text !== undefined ? step.text : step);
  }).filter(function(/** @type {string} */ text) {
    return text !== '';
  });
}

/**
 * @param {*} value
 * @return {string}
 */
function importString_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

/**
 * Numeric fields from a PDF arrive as either numbers (from the data block) or strings
 * (from the layout parser). Returns a real number where possible so the sheet stores a
 * number, the trimmed string where it isn't one (a yield the parser couldn't read cleanly
 * is better shown as-is than silently zeroed), and '' for nothing at all.
 * @param {*} value
 * @return {*}
 */
function importNumber_(value) {
  var text = importString_(value);
  if (text === '') return '';

  var numeric = Number(text);
  return isNaN(numeric) ? text : numeric;
}
