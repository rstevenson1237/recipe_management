/**
 * Sheets.gs
 * Schema constants, date-keyed DB sheet-set resolution/rollover, and Helper Data access.
 *
 * The three "DB - *" sheets always come in matched trios that share one date key
 * (DDMMYY, or "DDMMYY (2)", "DDMMYY (3)", ... once a date's first set fills up).
 * getActiveSheetSet() is the ONLY way callers should obtain those sheets - it is what
 * keeps Headers / Ingredients / Instructions in lock step with each other.
 */

var HEADER_PREFIX = 'DB - Headers ';
var INGREDIENTS_PREFIX = 'DB - Ingredients ';
var INSTRUCTIONS_PREFIX = 'DB - Instructions ';
var HELPER_SHEET_NAME = 'Helper Data';
var DASHBOARD_SHEET_NAME = 'Dashboard';
var MAX_RECIPES_PER_SET = 25;

var HEADER_COLUMNS = [
  'Name', 'Measure Type', 'Reporting U of M', 'Yield Qty', 'Yield U of M',
  'Weight Qty', 'Weight U of M', 'Volume Qty', 'Volume U of M',
  'Each Qty', 'Each U of M', 'Portion Size', 'Portion U of M',
  'Available in Inventory', 'Inventory U of M'
];

// 'Preparation' is appended last on purpose: every existing read indexes into these
// rows by position (Ingredient Order is row[4] all over Save.gs/Export.gs), so a new
// column is only safe at the end. repairSheetColumns_() backfills the header on
// ingredient sheets created before it existed.
var INGREDIENT_COLUMNS = ['Name', 'Ingredient', 'Qty', 'U of M', 'Ingredient Order', 'Preparation'];

var INSTRUCTION_COLUMNS = ['Name', 'Step Order', 'Instruction Text'];

var HELPER_COLUMNS = ['Weight', 'Volume', 'Each', 'Item Data'];

var DEFAULT_UOM_SEED = {
  Weight: ['GRM', 'OZ-wt', 'LB'],
  Volume: ['ML', 'tsp', 'Tbsp', 'OZ-fl', 'Pint', 'Quart', 'Gallon', '1/2 gallon', 'Cup', 'Dash', 'Drop'],
  Each: ['Each']
};

/**
 * Formats a date as DDMMYY using the spreadsheet's own timezone.
 * @param {Date} date
 * @return {string}
 */
function formatDateKey_(date) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), 'ddMMyy');
}

/**
 * Builds the suffixed key for the Nth set of a given base date key.
 * suffixIndex 1 -> "190826", 2 -> "190826 (2)", 3 -> "190826 (3)", ...
 * @param {string} baseKey
 * @param {number} suffixIndex
 * @return {string}
 */
function buildSetKey_(baseKey, suffixIndex) {
  return suffixIndex === 1 ? baseKey : baseKey + ' (' + suffixIndex + ')';
}

/**
 * @typedef {Object} SheetSet
 * @property {string} key
 * @property {GoogleAppsScript.Spreadsheet.Sheet} headers
 * @property {GoogleAppsScript.Spreadsheet.Sheet} ingredients
 * @property {GoogleAppsScript.Spreadsheet.Sheet} instructions
 */

/**
 * Returns {headers, ingredients, instructions} sheet objects for a key, creating
 * whichever of the three are missing (hidden, with a frozen, labeled header row).
 * Creating/repairing all three together is what keeps them in lock step.
 * @param {string} key
 * @return {SheetSet}
 */
function ensureSheetSetComplete_(key) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    key: key,
    headers: getOrCreateDbSheet_(ss, HEADER_PREFIX + key, HEADER_COLUMNS),
    ingredients: getOrCreateDbSheet_(ss, INGREDIENTS_PREFIX + key, INGREDIENT_COLUMNS),
    instructions: getOrCreateDbSheet_(ss, INSTRUCTIONS_PREFIX + key, INSTRUCTION_COLUMNS)
  };
}

/**
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} name
 * @param {string[]} columns
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateDbSheet_(ss, name, columns) {
  var sheet = ss.getSheetByName(name);
  if (sheet) {
    repairSheetColumns_(sheet, columns);
    return sheet;
  }

  sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, columns.length).setValues([columns]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.hideSheet();
  return sheet;
}

/**
 * Brings an existing sheet's header row into line with the current schema: widens the
 * sheet if it holds fewer columns than the schema needs, then rewrites any header cell
 * that doesn't match. This is how a column added to one of the *_COLUMNS constants
 * reaches sheets created before it existed (the 'Preparation' column added to
 * INGREDIENT_COLUMNS after recipes were already on file) - there is no separate
 * migration step to run, every getOrCreateDbSheet_ call is one.
 *
 * Only row 1 is ever touched. Existing data rows are left exactly as they were, so the
 * new column simply reads back as '' for rows written before it existed.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {string[]} columns
 */
function repairSheetColumns_(sheet, columns) {
  var maxColumns = sheet.getMaxColumns();
  if (maxColumns < columns.length) {
    sheet.insertColumnsAfter(maxColumns, columns.length - maxColumns);
  }

  var headerRange = sheet.getRange(1, 1, 1, columns.length);
  var current = headerRange.getValues()[0];
  var stale = columns.some(function(label, i) { return String(current[i]).trim() !== label; });
  if (stale) headerRange.setValues([columns]).setFontWeight('bold');
}

/**
 * Given a base key (today's DDMMYY in production, or a synthetic key in tests),
 * returns the key of the set a new recipe should be written into right now: the
 * base key itself, or the next-numbered suffix once the current one holds
 * MAX_RECIPES_PER_SET recipes. Pure lookup - does not create or modify sheets -
 * which is what makes it safe to exercise directly from runSelfTest().
 * @param {string} baseKey
 * @return {string}
 */
function resolveActiveKey_(baseKey) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var suffixIndex = 1;

  while (true) {
    var key = buildSetKey_(baseKey, suffixIndex);
    var headerSheet = ss.getSheetByName(HEADER_PREFIX + key);

    if (!headerSheet) return key;

    var recipeCount = Math.max(headerSheet.getLastRow() - 1, 0);
    if (recipeCount < MAX_RECIPES_PER_SET) return key;

    suffixIndex++;
  }
}

/**
 * Resolves the sheet set that a new recipe should be written into: today's set,
 * or the next-numbered set for today once the current one holds MAX_RECIPES_PER_SET
 * recipes. Creates sheets as needed.
 * @return {SheetSet}
 */
function getActiveSheetSet() {
  var baseKey = formatDateKey_(new Date());
  var key = resolveActiveKey_(baseKey);
  return ensureSheetSetComplete_(key);
}

/**
 * Enumerates every existing sheet set (across all dates), repairing any that are
 * missing one of the three sheets. Used by export and by the duplicate-name check.
 * @return {SheetSet[]}
 */
function listAllSheetSets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  /** @type {string[]} */
  var keys = [];

  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    if (name.indexOf(HEADER_PREFIX) === 0) {
      keys.push(name.substring(HEADER_PREFIX.length));
    }
  });

  return keys.map(function(key) {
    return ensureSheetSetComplete_(key);
  });
}

/**
 * Reads all data rows (excluding the header row) from a sheet, trimmed to the
 * given column count. Returns [] for an empty/header-only sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} columnCount
 * @return {Array<Array<*>>}
 */
function getDataRows_(sheet, columnCount) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, columnCount).getValues();
}

/**
 * Checks whether a recipe name already exists in ANY header sheet, across all
 * date-keyed sets. Comparison is case-insensitive and trims whitespace, since
 * Name is the join key across all three DB sheets and must be unique.
 * @param {string} name
 * @param {string=} excludeName - when editing a recipe, its own (pre-edit) name -
 *     so renaming a recipe to itself, or leaving the name unchanged, never
 *     reports a collision with itself.
 * @return {boolean}
 */
function recipeNameExists_(name, excludeName) {
  var normalized = String(name).trim().toLowerCase();
  var excludeNormalized = excludeName ? String(excludeName).trim().toLowerCase() : null;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  return ss.getSheets().some(function(sheet) {
    var sheetName = sheet.getName();
    if (sheetName.indexOf(HEADER_PREFIX) !== 0) return false;

    var rows = getDataRows_(sheet, 1);
    return rows.some(function(row) {
      var rowName = String(row[0]).trim().toLowerCase();
      if (excludeNormalized && rowName === excludeNormalized) return false;
      return rowName === normalized;
    });
  });
}

/**
 * @typedef {Object} RecipeLocation
 * @property {SheetSet} set
 * @property {number} headerRow - the recipe's actual row number in set.headers.
 */

/**
 * Finds which sheet set a recipe lives in and its Headers row number, by name
 * (case-insensitive, trimmed) across every date-keyed set on file. Used by the
 * edit flow to locate the rows to overwrite - a recipe's Ingredients and
 * Instructions rows always live in this same sheet set as its Headers row.
 * @param {string} name
 * @return {?RecipeLocation}
 */
function findRecipeLocation_(name) {
  var normalized = String(name).trim().toLowerCase();
  var sets = listAllSheetSets_();

  for (var i = 0; i < sets.length; i++) {
    var rows = getDataRows_(sets[i].headers, 1);
    for (var r = 0; r < rows.length; r++) {
      if (String(rows[r][0]).trim().toLowerCase() === normalized) {
        return { set: sets[i], headerRow: r + 2 }; // +1 for the header row, +1 for the 0-index
      }
    }
  }
  return null;
}

/**
 * Returns every recipe name across all sheet sets, sorted alphabetically - used to
 * populate the "Edit Existing Recipe" dropdown. Only reads the Name column, so it's
 * far cheaper than exportAllRecipes() when ingredients/instructions aren't needed.
 * @return {string[]}
 */
function getRecipeNameList() {
  var sets = listAllSheetSets_();
  /** @type {string[]} */
  var names = [];

  sets.forEach(function(set) {
    getDataRows_(set.headers, 1).forEach(function(row) {
      var name = String(row[0]).trim();
      if (name) names.push(name);
    });
  });

  names.sort(function(a, b) { return a.toLowerCase().localeCompare(b.toLowerCase()); });
  return names;
}

/**
 * Ensures the Helper Data sheet exists, hidden, with headers and (if empty of
 * data) the default UOM lists seeded in. Item Data (column D) is left for the
 * workbook owner to populate with the allowed ingredient list.
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function ensureHelperDataSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(HELPER_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(HELPER_SHEET_NAME);
    sheet.getRange(1, 1, 1, HELPER_COLUMNS.length).setValues([HELPER_COLUMNS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.hideSheet();
  }

  if (sheet.getLastRow() < 2) {
    var weight = DEFAULT_UOM_SEED.Weight;
    var volume = DEFAULT_UOM_SEED.Volume;
    var each = DEFAULT_UOM_SEED.Each;
    var rowCount = Math.max(weight.length, volume.length, each.length);
    var rows = [];
    for (var i = 0; i < rowCount; i++) {
      rows.push([weight[i] || '', volume[i] || '', each[i] || '', '']);
    }
    if (rows.length) {
      sheet.getRange(2, 1, rows.length, HELPER_COLUMNS.length).setValues(rows);
    }
  }

  return sheet;
}

/**
 * @typedef {Object} HelperData
 * @property {{Weight: string[], Volume: string[], Each: string[]}} uom
 * @property {string[]} items
 */

/**
 * Reads Helper Data into the shape the dialog and server-side validation both use:
 * { uom: { Weight: [...], Volume: [...], Each: [...] }, items: [...] }
 * @return {HelperData}
 */
function getHelperData() {
  var sheet = ensureHelperDataSheet_();
  var rows = getDataRows_(sheet, HELPER_COLUMNS.length);

  /** @type {HelperData} */
  var result = { uom: { Weight: [], Volume: [], Each: [] }, items: [] };
  /** @type {Object<string, boolean>} */
  var seenItems = {};

  rows.forEach(function(row) {
    var weight = String(row[0]).trim();
    var volume = String(row[1]).trim();
    var each = String(row[2]).trim();
    var item = String(row[3]).trim();

    if (weight) result.uom.Weight.push(weight);
    if (volume) result.uom.Volume.push(volume);
    if (each) result.uom.Each.push(each);
    if (item && !seenItems[item.toLowerCase()]) {
      seenItems[item.toLowerCase()] = true;
      result.items.push(item);
    }
  });

  return result;
}

/**
 * Appends ingredient names and units of measure to Helper Data that aren't already
 * listed there, and returns what was actually added. Used by the PDF import: a recipe
 * exported from another workbook can reference items this workbook's owner never
 * listed, and leaving them out would make every imported recipe fail validation the
 * next time someone opened it in the entry dialog.
 *
 * Helper Data's four columns are independent single-column lists that merely share a
 * sheet, so each one is filled from its own first free row rather than by appending
 * whole rows - appending rows would push new Weight units far below the existing ones
 * just because the Item Data column happens to be longer.
 * @param {string[]} itemNames
 * @param {{Weight: string[], Volume: string[], Each: string[]}} uomsByCategory
 * @return {{items: string[], uoms: string[]}}
 */
function appendHelperDataItems_(itemNames, uomsByCategory) {
  var sheet = ensureHelperDataSheet_();
  var rows = getDataRows_(sheet, HELPER_COLUMNS.length);
  /** @type {{items: string[], uoms: string[]}} */
  var added = { items: [], uoms: [] };

  var plans = [
    { column: 1, label: 'Weight', values: (uomsByCategory && uomsByCategory.Weight) || [] },
    { column: 2, label: 'Volume', values: (uomsByCategory && uomsByCategory.Volume) || [] },
    { column: 3, label: 'Each', values: (uomsByCategory && uomsByCategory.Each) || [] },
    { column: 4, label: 'Item Data', values: itemNames || [] }
  ];

  plans.forEach(function(plan) {
    /** @type {Object<string, boolean>} */
    var existing = {};
    var firstFreeRow = 2;

    rows.forEach(function(row, index) {
      var value = String(row[plan.column - 1]).trim();
      if (!value) return;
      existing[value.toLowerCase()] = true;
      firstFreeRow = index + 3; // one past this row: +2 for the header row, +1 for the 0-index
    });

    /** @type {string[]} */
    var pending = [];
    plan.values.forEach(function(/** @type {*} */ value) {
      var clean = String(value).trim();
      if (!clean || existing[clean.toLowerCase()]) return;
      existing[clean.toLowerCase()] = true;
      pending.push(clean);
    });
    if (pending.length === 0) return;

    sheet.getRange(firstFreeRow, plan.column, pending.length, 1).setValues(pending.map(function(value) {
      return [value];
    }));

    pending.forEach(function(value) {
      if (plan.column === 4) added.items.push(value);
      else added.uoms.push(value + ' (' + plan.label + ')');
    });
  });

  return added;
}

/**
 * In-workbook self-test for the sheet-set rollover and lock-step invariants. Run
 * manually from the Apps Script editor (select runSelfTest > Run) after changing
 * anything in Sheets.gs or Save.gs.
 *
 * Uses a synthetic "TEST<timestamp>" key namespace so it can never collide with a
 * real DDMMYY date and never touches production data, and deletes every sheet it
 * creates before returning (even on failure). Throws on the first failed
 * assertion; logs one line per passed assertion via Logger.log (View > Logs).
 */
function runSelfTest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var testBaseKey = 'TEST' + new Date().getTime();
  /** @type {string[]} */
  var createdSheetNames = [];

  /**
   * @param {boolean} condition
   * @param {string} message
   */
  function assert(condition, message) {
    if (!condition) throw new Error('runSelfTest FAILED: ' + message);
    Logger.log('OK: ' + message);
  }

  /**
   * @param {SheetSet} set
   * @return {SheetSet}
   */
  function trackSet(set) {
    [set.headers, set.ingredients, set.instructions].forEach(function(sheet) {
      createdSheetNames.push(sheet.getName());
    });
    return set;
  }

  try {
    // --- Rollover: fill a set to MAX_RECIPES_PER_SET, expect a second set to appear ---
    var firstKey = resolveActiveKey_(testBaseKey);
    assert(firstKey === testBaseKey, 'first set uses the base key with no suffix');

    var set = trackSet(ensureSheetSetComplete_(firstKey));
    for (var i = 1; i <= MAX_RECIPES_PER_SET; i++) {
      set.headers.appendRow(['Test Recipe ' + i, 'Weight', 'GRM', 1, 'GRM', '', '', '', '', '', '', '', '', 'No', '']);
      set.ingredients.appendRow(['Test Recipe ' + i, 'Flour', 1, 'GRM', 1, 'sifted']);
      set.instructions.appendRow(['Test Recipe ' + i, 1, 'Mix']);
    }
    assert(Math.max(set.headers.getLastRow() - 1, 0) === MAX_RECIPES_PER_SET,
        'first set holds exactly ' + MAX_RECIPES_PER_SET + ' recipes after filling it');

    var secondKey = resolveActiveKey_(testBaseKey);
    assert(secondKey === buildSetKey_(testBaseKey, 2), 'a full set rolls over to suffix (2), got: ' + secondKey);

    var secondSet = trackSet(ensureSheetSetComplete_(secondKey));
    assert(ss.getSheetByName(HEADER_PREFIX + secondKey) !== null, 'second Headers sheet exists');
    assert(ss.getSheetByName(INGREDIENTS_PREFIX + secondKey) !== null, 'second Ingredients sheet exists');
    assert(ss.getSheetByName(INSTRUCTIONS_PREFIX + secondKey) !== null, 'second Instructions sheet exists');

    secondSet.headers.appendRow(['Test Recipe 27', 'Weight', 'GRM', 1, 'GRM', '', '', '', '', '', '', '', '', 'No', '']);
    secondSet.ingredients.appendRow(['Test Recipe 27', 'Flour', 1, 'GRM', 1, '']);
    secondSet.instructions.appendRow(['Test Recipe 27', 1, 'Mix']);
    assert(Math.max(secondSet.headers.getLastRow() - 1, 0) < MAX_RECIPES_PER_SET, 'second set is not yet full');

    // --- Lock-step: every header name has matching ingredient/instruction rows, no orphans ---
    [set, secondSet].forEach(function(s) {
      var headerNames = getDataRows_(s.headers, 1).map(function(r) { return r[0]; });
      var ingredientNames = getDataRows_(s.ingredients, 1).map(function(r) { return r[0]; });
      var instructionNames = getDataRows_(s.instructions, 1).map(function(r) { return r[0]; });

      headerNames.forEach(function(name) {
        assert(ingredientNames.indexOf(name) !== -1, 'recipe "' + name + '" has ingredient rows');
        assert(instructionNames.indexOf(name) !== -1, 'recipe "' + name + '" has instruction rows');
      });
      ingredientNames.forEach(function(name) {
        assert(headerNames.indexOf(name) !== -1, 'no orphan ingredient rows for "' + name + '"');
      });
      instructionNames.forEach(function(name) {
        assert(headerNames.indexOf(name) !== -1, 'no orphan instruction rows for "' + name + '"');
      });
    });

    // --- Rollback: a write that fails partway through must leave all three sheets unchanged ---
    var rollbackKey = resolveActiveKey_(testBaseKey + '-RB');
    var rollbackSet = trackSet(ensureSheetSetComplete_(rollbackKey));
    var startCounts = {
      headers: rollbackSet.headers.getLastRow(),
      ingredients: rollbackSet.ingredients.getLastRow(),
      instructions: rollbackSet.instructions.getLastRow()
    };

    try {
      rollbackSet.ingredients.appendRow(['Broken Recipe', 'Flour', 1, 'GRM', 1, '']);
      rollbackSet.instructions.appendRow(['Broken Recipe', 1, 'Mix']);
      throw new Error('simulated failure before the header row is written');
    } catch (writeError) {
      truncateSheetToRow_(rollbackSet.headers, startCounts.headers);
      truncateSheetToRow_(rollbackSet.ingredients, startCounts.ingredients);
      truncateSheetToRow_(rollbackSet.instructions, startCounts.instructions);
    }

    assert(rollbackSet.headers.getLastRow() === startCounts.headers, 'rollback restores Headers row count');
    assert(rollbackSet.ingredients.getLastRow() === startCounts.ingredients, 'rollback restores Ingredients row count');
    assert(rollbackSet.instructions.getLastRow() === startCounts.instructions, 'rollback restores Instructions row count');

    // --- Schema repair: an ingredient sheet created before the Preparation column
    // --- existed must gain its header (and keep its data) the next time it's opened ---
    var legacyName = INGREDIENTS_PREFIX + testBaseKey + '-LEGACY';
    var legacySheet = ss.insertSheet(legacyName);
    createdSheetNames.push(legacyName);
    legacySheet.getRange(1, 1, 1, 5).setValues([['Name', 'Ingredient', 'Qty', 'U of M', 'Ingredient Order']]);
    legacySheet.appendRow(['Legacy Recipe', 'Flour', 1, 'GRM', 1]);
    if (legacySheet.getMaxColumns() > 5) legacySheet.deleteColumns(6, legacySheet.getMaxColumns() - 5);
    assert(legacySheet.getMaxColumns() === 5, 'legacy ingredient sheet starts with exactly 5 columns');

    repairSheetColumns_(legacySheet, INGREDIENT_COLUMNS);
    assert(legacySheet.getRange(1, INGREDIENT_COLUMNS.length).getValue() === 'Preparation',
        'repair adds the Preparation header to a legacy ingredient sheet');
    var legacyRow = legacySheet.getRange(2, 1, 1, INGREDIENT_COLUMNS.length).getValues()[0];
    assert(String(legacyRow[1]) === 'Flour' && Number(legacyRow[4]) === 1 && String(legacyRow[5]) === '',
        'repair leaves legacy ingredient data intact, with an empty Preparation');

    Logger.log('runSelfTest: ALL ASSERTIONS PASSED');
  } finally {
    createdSheetNames.forEach(function(name) {
      var sheet = ss.getSheetByName(name);
      if (sheet) ss.deleteSheet(sheet);
    });
  }
}
