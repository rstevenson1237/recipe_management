/**
 * Save.gs
 * Server-side validation and the three-sheet (Headers/Ingredients/Instructions) write
 * path for a new recipe. Client-side validation in Index.html is for responsiveness
 * only - this is the real boundary, since the client's copy of Helper Data can be stale.
 */

/**
 * Receives the payload built by Index.html's saveData() and writes it to the active
 * date-keyed sheet set. Returns { ok: true } on success, or
 * { ok: false, errors: {field: message}, message: string } on validation failure.
 * Throws only for genuine infrastructure errors (the client's withFailureHandler
 * covers that case).
 * @param {*} payload - JSON payload built by Index.html's saveData(); shape isn't
 *     trusted (that's what validateRecipePayload_ is for), so it's typed loosely.
 */
function saveRecipeFromWeb(payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    var helperData = getHelperData();
    var errors = validateRecipePayload_(payload, helperData);
    if (Object.keys(errors).length > 0) {
      return { ok: false, errors: errors, message: 'Please fix the highlighted fields.' };
    }

    var name = String(payload.name).trim();
    if (recipeNameExists_(name)) {
      return {
        ok: false,
        errors: { name: 'A recipe named "' + name + '" already exists.' },
        message: 'Recipe name must be unique.'
      };
    }

    var set = getActiveSheetSet();
    var startCounts = {
      headers: set.headers.getLastRow(),
      ingredients: set.ingredients.getLastRow(),
      instructions: set.instructions.getLastRow()
    };

    try {
      var cleanIngredients = payload.ingredients.filter(function(/** @type {*} */ ing) {
        return ing.name && String(ing.name).trim() !== '';
      });
      var ingredientRows = cleanIngredients.map(function(/** @type {*} */ ing, /** @type {number} */ index) {
        return [name, String(ing.name).trim(), ing.qty, ing.uom, index + 1];
      });
      appendRows_(set.ingredients, ingredientRows);

      var cleanSteps = payload.instructions.filter(function(/** @type {*} */ step) {
        return step.text && String(step.text).trim() !== '';
      });
      var instructionRows = cleanSteps.map(function(/** @type {*} */ step, /** @type {number} */ index) {
        return [name, index + 1, String(step.text).trim()];
      });
      appendRows_(set.instructions, instructionRows);

      appendRows_(set.headers, [[
        name,
        payload.measureType,
        payload.yieldUom, // Reporting U of M always mirrors Yield U of M - not prompted for separately.
        payload.yieldQty,
        payload.yieldUom,
        payload.weightQty || '',
        payload.weightUom || '',
        payload.volumeQty || '',
        payload.volumeUom || '',
        payload.eachQty || '',
        payload.eachUom || '',
        payload.portionSize || '',
        payload.portionUom || '',
        payload.inventoryYesNo,
        payload.inventoryUom || ''
      ]]);
    } catch (writeError) {
      // Roll every sheet back to its pre-write row count so the trio never drifts
      // out of lock step because of a partial write.
      truncateSheetToRow_(set.headers, startCounts.headers);
      truncateSheetToRow_(set.ingredients, startCounts.ingredients);
      truncateSheetToRow_(set.instructions, startCounts.instructions);
      throw writeError;
    }

    refreshDashboard();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Appends every row in a single batched write instead of one appendRow() call per
 * row - each appendRow() is its own round trip to the Sheets service, so this
 * turns an O(rows) sequence of calls into one for the multi-row ingredient and
 * instruction writes. No-op for an empty rows array.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<Array<*>>} rows
 */
function appendRows_(sheet, rows) {
  if (rows.length === 0) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * Deletes any rows appended beyond targetLastRow, restoring a sheet to the row
 * count it had before a failed write.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} targetLastRow
 */
function truncateSheetToRow_(sheet, targetLastRow) {
  var currentLastRow = sheet.getLastRow();
  if (currentLastRow > targetLastRow) {
    sheet.deleteRows(targetLastRow + 1, currentLastRow - targetLastRow);
  }
}

/**
 * Validates the payload against live Helper Data. Returns a map of field name ->
 * error message; an empty object means the payload is valid. Keys match the
 * field-level error keys the dialog displays inline.
 * @param {*} payload
 * @param {HelperData} helperData
 * @return {Object<string, string>}
 */
function validateRecipePayload_(payload, helperData) {
  /** @type {Object<string, string>} */
  var errors = {};
  var measureCategories = ['Weight', 'Volume', 'Each'];

  var name = payload.name ? String(payload.name).trim() : '';
  if (!name) errors.name = 'Recipe Name is required.';

  if (measureCategories.indexOf(payload.measureType) === -1) {
    errors.measureType = 'Select a Measure Type.';
  }

  if (!isPositiveNumber_(payload.yieldQty)) {
    errors.yieldQty = 'Yield Qty must be a number greater than 0.';
  }

  if (payload.measureType && measureCategories.indexOf(payload.measureType) !== -1) {
    var measureTypeUoms = getUomListForCategory_(helperData, payload.measureType);
    if (!payload.yieldUom || measureTypeUoms.indexOf(payload.yieldUom) === -1) {
      errors.yieldUom = 'Select a valid Yield U of M for ' + payload.measureType + '.';
    }
  }

  var portionCategory = findUomCategory_(helperData, payload.portionUom);
  if (payload.portionUom && !portionCategory) {
    errors.portionUom = 'Unrecognized Portion U of M.';
  }

  if (payload.inventoryYesNo === 'Yes') {
    var inventoryCategory = findUomCategory_(helperData, payload.inventoryUom);
    if (!payload.inventoryUom || !inventoryCategory) {
      errors.inventoryUom = 'Select a valid Inventory U of M.';
    }
  }

  // Only the category that actually conflicts with the recipe's Measure Type needs
  // an equivalency entry - the base category is filled from Yield automatically.
  var inventoryEquivCategory = payload.inventoryYesNo === 'Yes' ? findUomCategory_(helperData, payload.inventoryUom) : null;
  /** @type {Object<string, boolean>} */
  var conflictCategories = {};
  if (portionCategory && portionCategory !== payload.measureType) conflictCategories[portionCategory] = true;
  if (inventoryEquivCategory && inventoryEquivCategory !== payload.measureType) conflictCategories[inventoryEquivCategory] = true;

  if (conflictCategories.Weight) {
    if (!isPositiveNumber_(payload.weightQty) || helperData.uom.Weight.indexOf(payload.weightUom) === -1) {
      errors.weightEquivalency = 'Provide a valid Weight equivalency (Qty + U of M).';
    }
  }
  if (conflictCategories.Volume) {
    if (!isPositiveNumber_(payload.volumeQty) || helperData.uom.Volume.indexOf(payload.volumeUom) === -1) {
      errors.volumeEquivalency = 'Provide a valid Volume equivalency (Qty + U of M).';
    }
  }
  if (conflictCategories.Each) {
    if (!isPositiveNumber_(payload.eachQty) || helperData.uom.Each.indexOf(payload.eachUom) === -1) {
      errors.eachEquivalency = 'Provide a valid Each equivalency (Qty + U of M).';
    }
  }

  var allUoms = helperData.uom.Weight.concat(helperData.uom.Volume, helperData.uom.Each);
  var ingredients = Array.isArray(payload.ingredients) ? payload.ingredients : [];
  var meaningfulIngredients = ingredients.filter(function(/** @type {*} */ ing) {
    return ing && ing.name && String(ing.name).trim() !== '';
  });

  if (meaningfulIngredients.length === 0) {
    errors.ingredients = 'Add at least one ingredient.';
  } else {
    var ingredientProblem = meaningfulIngredients.some(function(/** @type {*} */ ing) {
      var validName = helperData.items.indexOf(String(ing.name).trim()) !== -1;
      var validQty = isPositiveNumber_(ing.qty);
      var validUom = allUoms.indexOf(ing.uom) !== -1;
      return !validName || !validQty || !validUom;
    });
    if (ingredientProblem) {
      errors.ingredients = 'Every ingredient needs a recognized name, a Qty > 0, and a valid U of M.';
    }
  }

  return errors;
}

/**
 * @param {*} value
 * @return {boolean}
 */
function isPositiveNumber_(value) {
  var num = Number(value);
  return value !== '' && value !== null && value !== undefined && !isNaN(num) && num > 0;
}

/**
 * @param {HelperData} helperData
 * @param {*} category
 * @return {string[]}
 */
function getUomListForCategory_(helperData, category) {
  if (category === 'Weight') return helperData.uom.Weight;
  if (category === 'Volume') return helperData.uom.Volume;
  if (category === 'Each') return helperData.uom.Each;
  return [];
}

/**
 * @param {HelperData} helperData
 * @param {*} uom
 * @return {?string}
 */
function findUomCategory_(helperData, uom) {
  if (!uom) return null;
  /** @type {Array<keyof HelperData['uom']>} */
  var categories = ['Weight', 'Volume', 'Each'];
  for (var i = 0; i < categories.length; i++) {
    if (helperData.uom[categories[i]].indexOf(uom) !== -1) return categories[i];
  }
  return null;
}
