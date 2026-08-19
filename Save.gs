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
      var cleanIngredients = payload.ingredients.filter(function(ing) {
        return ing.name && String(ing.name).trim() !== '';
      });
      cleanIngredients.forEach(function(ing, index) {
        set.ingredients.appendRow([name, String(ing.name).trim(), ing.qty, ing.uom, index + 1]);
      });

      var cleanSteps = payload.instructions.filter(function(step) {
        return step.text && String(step.text).trim() !== '';
      });
      cleanSteps.forEach(function(step, index) {
        set.instructions.appendRow([name, index + 1, String(step.text).trim()]);
      });

      set.headers.appendRow([
        name,
        payload.measureType,
        payload.reportingUom,
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
      ]);
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
 * Deletes any rows appended beyond targetLastRow, restoring a sheet to the row
 * count it had before a failed write.
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
 */
function validateRecipePayload_(payload, helperData) {
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
    if (!payload.yieldUom || helperData.uom[payload.measureType].indexOf(payload.yieldUom) === -1) {
      errors.yieldUom = 'Select a valid Yield U of M for ' + payload.measureType + '.';
    }
    if (!payload.reportingUom || helperData.uom[payload.measureType].indexOf(payload.reportingUom) === -1) {
      errors.reportingUom = 'Select a valid Reporting U of M for ' + payload.measureType + '.';
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

  var needsEquivalency =
    (portionCategory && portionCategory !== payload.measureType) ||
    (payload.inventoryYesNo === 'Yes' &&
      findUomCategory_(helperData, payload.inventoryUom) &&
      findUomCategory_(helperData, payload.inventoryUom) !== payload.measureType);

  if (needsEquivalency) {
    if (payload.measureType === 'Weight' || portionCategory === 'Weight' ||
        (payload.inventoryYesNo === 'Yes' && findUomCategory_(helperData, payload.inventoryUom) === 'Weight')) {
      if (!isPositiveNumber_(payload.weightQty) || helperData.uom.Weight.indexOf(payload.weightUom) === -1) {
        errors.weightEquivalency = 'Provide a valid Weight equivalency (Qty + U of M).';
      }
    }
    if (payload.measureType === 'Volume' || portionCategory === 'Volume' ||
        (payload.inventoryYesNo === 'Yes' && findUomCategory_(helperData, payload.inventoryUom) === 'Volume')) {
      if (!isPositiveNumber_(payload.volumeQty) || helperData.uom.Volume.indexOf(payload.volumeUom) === -1) {
        errors.volumeEquivalency = 'Provide a valid Volume equivalency (Qty + U of M).';
      }
    }
    if (payload.measureType === 'Each' || portionCategory === 'Each' ||
        (payload.inventoryYesNo === 'Yes' && findUomCategory_(helperData, payload.inventoryUom) === 'Each')) {
      if (!isPositiveNumber_(payload.eachQty) || helperData.uom.Each.indexOf(payload.eachUom) === -1) {
        errors.eachEquivalency = 'Provide a valid Each equivalency (Qty + U of M).';
      }
    }
  }

  var allUoms = helperData.uom.Weight.concat(helperData.uom.Volume, helperData.uom.Each);
  var ingredients = Array.isArray(payload.ingredients) ? payload.ingredients : [];
  var meaningfulIngredients = ingredients.filter(function(ing) {
    return ing && ing.name && String(ing.name).trim() !== '';
  });

  if (meaningfulIngredients.length === 0) {
    errors.ingredients = 'Add at least one ingredient.';
  } else {
    var ingredientProblem = meaningfulIngredients.some(function(ing) {
      var validName = helperData.items.indexOf(String(ing.name).trim()) !== -1;
      var validQty = isPositiveNumber_(ing.qty);
      var validUom = allUoms.indexOf(ing.uom) !== -1;
      return !validName || !validQty || !validUom;
    });
    if (ingredientProblem) {
      errors.ingredients = 'Every ingredient needs a recognized name, a Qty > 0, and a valid U of M.';
    }
  }

  var instructions = Array.isArray(payload.instructions) ? payload.instructions : [];
  var meaningfulSteps = instructions.filter(function(step) {
    return step && step.text && String(step.text).trim() !== '';
  });
  if (meaningfulSteps.length === 0) {
    errors.instructions = 'Add at least one instruction step.';
  }

  return errors;
}

function isPositiveNumber_(value) {
  var num = Number(value);
  return value !== '' && value !== null && value !== undefined && !isNaN(num) && num > 0;
}

function findUomCategory_(helperData, uom) {
  if (!uom) return null;
  for (var category in helperData.uom) {
    if (helperData.uom[category].indexOf(uom) !== -1) return category;
  }
  return null;
}
