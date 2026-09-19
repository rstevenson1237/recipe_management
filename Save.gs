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
        return [name, String(ing.name).trim(), ing.qty, ing.uom, index + 1, cleanPreparation_(ing.prep)];
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

    // Append-only: avoids rescanning every recipe on file just to add one row.
    // Full re-sort is available via the "Setup / Repair Dashboard" menu item.
    appendDashboardRow_({
      name: name,
      measureType: payload.measureType,
      yieldQty: payload.yieldQty,
      yieldUom: payload.yieldUom,
      portionSize: payload.portionSize,
      portionUom: payload.portionUom,
      ingredientCount: cleanIngredients.length,
      stepCount: cleanSteps.length,
      sourceKey: set.key
    });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Loads one recipe's full data for the "Edit Existing Recipe" dropdown: the same
 * form-shape fields Index.html's saveData() sends, plus its ingredients and
 * instructions in saved order. Throws if the name can't be found (the dropdown
 * only ever offers names read moments earlier, but the sheet could change in
 * between - e.g. someone else deletes or renames it - so this is a real
 * possibility, not defensive filler).
 * @param {string} name
 * @return {*}
 */
function getRecipeForEdit(name) {
  var location = findRecipeLocation_(name);
  if (!location) {
    throw new Error('Recipe "' + name + '" was not found - it may have been renamed or deleted. Reopen the dialog and try again.');
  }

  var set = location.set;
  var headerRow = set.headers.getRange(location.headerRow, 1, 1, HEADER_COLUMNS.length).getValues()[0];
  var normalized = String(headerRow[0]).trim().toLowerCase();

  var ingredients = getDataRows_(set.ingredients, INGREDIENT_COLUMNS.length)
      .filter(function(r) { return String(r[0]).trim().toLowerCase() === normalized; })
      .sort(function(a, b) { return Number(a[4]) - Number(b[4]); })
      .map(function(r) { return { name: r[1], qty: r[2], uom: r[3], prep: r[5] }; });

  var instructions = getDataRows_(set.instructions, INSTRUCTION_COLUMNS.length)
      .filter(function(r) { return String(r[0]).trim().toLowerCase() === normalized; })
      .sort(function(a, b) { return Number(a[1]) - Number(b[1]); })
      .map(function(r) { return { text: r[2] }; });

  return {
    originalName: String(headerRow[0]).trim(),
    name: String(headerRow[0]).trim(),
    measureType: headerRow[1],
    yieldQty: headerRow[3],
    yieldUom: headerRow[4],
    weightQty: headerRow[5],
    weightUom: headerRow[6],
    volumeQty: headerRow[7],
    volumeUom: headerRow[8],
    eachQty: headerRow[9],
    eachUom: headerRow[10],
    portionSize: headerRow[11],
    portionUom: headerRow[12],
    inventoryYesNo: headerRow[13],
    inventoryUom: headerRow[14],
    ingredients: ingredients,
    instructions: instructions
  };
}

/**
 * Receives the payload built by Index.html's saveData() while editing an existing
 * recipe, and overwrites that recipe's Headers row plus its Ingredients/Instructions
 * rows in place - a recipe never moves to a different sheet set on edit, only its
 * own set's rows are touched. Returns the same { ok, errors, message } shape as
 * saveRecipeFromWeb.
 * @param {string} originalName - the recipe's name before this edit; used to find
 *     its existing rows and excluded from the uniqueness check.
 * @param {*} payload - JSON payload built by Index.html's saveData(); shape isn't
 *     trusted (that's what validateRecipePayload_ is for), so it's typed loosely.
 */
function updateRecipeFromWeb(originalName, payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    var helperData = getHelperData();
    var errors = validateRecipePayload_(payload, helperData);
    if (Object.keys(errors).length > 0) {
      return { ok: false, errors: errors, message: 'Please fix the highlighted fields.' };
    }

    var name = String(payload.name).trim();
    if (recipeNameExists_(name, originalName)) {
      return {
        ok: false,
        errors: { name: 'A recipe named "' + name + '" already exists.' },
        message: 'Recipe name must be unique.'
      };
    }

    var location = findRecipeLocation_(originalName);
    if (!location) {
      return {
        ok: false,
        errors: {},
        message: 'Recipe "' + originalName + '" was not found - it may have been renamed or deleted elsewhere. Reopen the dialog and try again.'
      };
    }

    var set = location.set;

    set.headers.getRange(location.headerRow, 1, 1, HEADER_COLUMNS.length).setValues([[
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

    var cleanIngredients = payload.ingredients.filter(function(/** @type {*} */ ing) {
      return ing.name && String(ing.name).trim() !== '';
    });
    var ingredientRows = cleanIngredients.map(function(/** @type {*} */ ing, /** @type {number} */ index) {
      return [name, String(ing.name).trim(), ing.qty, ing.uom, index + 1, cleanPreparation_(ing.prep)];
    });
    replaceRowsForRecipe_(set.ingredients, originalName, ingredientRows);

    var cleanSteps = payload.instructions.filter(function(/** @type {*} */ step) {
      return step.text && String(step.text).trim() !== '';
    });
    var instructionRows = cleanSteps.map(function(/** @type {*} */ step, /** @type {number} */ index) {
      return [name, index + 1, String(step.text).trim()];
    });
    replaceRowsForRecipe_(set.instructions, originalName, instructionRows);

    updateDashboardRow_(originalName, set.key, {
      name: name,
      measureType: payload.measureType,
      yieldQty: payload.yieldQty,
      yieldUom: payload.yieldUom,
      portionSize: payload.portionSize,
      portionUom: payload.portionUom,
      ingredientCount: cleanIngredients.length,
      stepCount: cleanSteps.length,
      sourceKey: set.key
    });

    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Removes every existing row for a recipe name from an Ingredients/Instructions
 * sheet, then appends newRows in their place - the delete-then-append used to
 * rewrite a recipe's rows on edit, since the new row count rarely matches the old
 * one. Deletion walks bottom-to-top so earlier deletes never shift the row numbers
 * of matches still pending.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string} name
 * @param {Array<Array<*>>} newRows
 */
function replaceRowsForRecipe_(sheet, name, newRows) {
  var normalized = String(name).trim().toLowerCase();
  var lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    var names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = names.length - 1; i >= 0; i--) {
      if (String(names[i][0]).trim().toLowerCase() === normalized) {
        sheet.deleteRow(2 + i);
      }
    }
  }

  appendRows_(sheet, newRows);
}

var MAX_PREPARATION_LENGTH = 200;

/**
 * Normalizes an ingredient's free-text Preparation ("minced", "diced small", ...):
 * trimmed and length-capped, never rejected. It's a note to whoever cooks the recipe
 * rather than a validated field - there is no list of allowed preparations to check it
 * against - so an over-long value is truncated instead of blocking the save.
 * @param {*} value
 * @return {string}
 */
function cleanPreparation_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim().substring(0, MAX_PREPARATION_LENGTH);
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
