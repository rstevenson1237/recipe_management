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

var INGREDIENT_COLUMNS = ['Name', 'Ingredient', 'Qty', 'U of M', 'Ingredient Order'];

var INSTRUCTION_COLUMNS = ['Name', 'Step Order', 'Instruction Text'];

var HELPER_COLUMNS = ['Weight', 'Volume', 'Each', 'Item Data'];

var DEFAULT_UOM_SEED = {
  Weight: ['GRM', 'OZ-wt', 'LB'],
  Volume: ['ML', 'tsp', 'Tbsp', 'OZ-fl', 'Pint', 'Quart', 'Gallon', '1/2 gallon', 'Cup', 'Dash', 'Drop'],
  Each: ['Each']
};

/**
 * Formats a date as DDMMYY using the spreadsheet's own timezone.
 */
function formatDateKey_(date) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return Utilities.formatDate(date, ss.getSpreadsheetTimeZone(), 'ddMMyy');
}

/**
 * Builds the suffixed key for the Nth set of a given base date key.
 * suffixIndex 1 -> "190826", 2 -> "190826 (2)", 3 -> "190826 (3)", ...
 */
function buildSetKey_(baseKey, suffixIndex) {
  return suffixIndex === 1 ? baseKey : baseKey + ' (' + suffixIndex + ')';
}

/**
 * Returns {headers, ingredients, instructions} sheet objects for a key, creating
 * whichever of the three are missing (hidden, with a frozen, labeled header row).
 * Creating/repairing all three together is what keeps them in lock step.
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

function getOrCreateDbSheet_(ss, name, columns) {
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, columns.length).setValues([columns]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.hideSheet();
  return sheet;
}

/**
 * Given a base key (today's DDMMYY in production, or a synthetic key in tests),
 * returns the key of the set a new recipe should be written into right now: the
 * base key itself, or the next-numbered suffix once the current one holds
 * MAX_RECIPES_PER_SET recipes. Pure lookup - does not create or modify sheets -
 * which is what makes it safe to exercise directly from runSelfTest().
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
 */
function getActiveSheetSet() {
  var baseKey = formatDateKey_(new Date());
  var key = resolveActiveKey_(baseKey);
  return ensureSheetSetComplete_(key);
}

/**
 * Enumerates every existing sheet set (across all dates), repairing any that are
 * missing one of the three sheets. Used by export and by the duplicate-name check.
 */
function listAllSheetSets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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
 */
function recipeNameExists_(name) {
  var normalized = String(name).trim().toLowerCase();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  return ss.getSheets().some(function(sheet) {
    var sheetName = sheet.getName();
    if (sheetName.indexOf(HEADER_PREFIX) !== 0) return false;

    var rows = getDataRows_(sheet, 1);
    return rows.some(function(row) {
      return String(row[0]).trim().toLowerCase() === normalized;
    });
  });
}

/**
 * Ensures the Helper Data sheet exists, hidden, with headers and (if empty of
 * data) the default UOM lists seeded in. Item Data (column D) is left for the
 * workbook owner to populate with the allowed ingredient list.
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
 * Reads Helper Data into the shape the dialog and server-side validation both use:
 * { uom: { Weight: [...], Volume: [...], Each: [...] }, items: [...] }
 */
function getHelperData() {
  var sheet = ensureHelperDataSheet_();
  var rows = getDataRows_(sheet, HELPER_COLUMNS.length);

  var result = { uom: { Weight: [], Volume: [], Each: [] }, items: [] };
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
  var createdSheetNames = [];

  function assert(condition, message) {
    if (!condition) throw new Error('runSelfTest FAILED: ' + message);
    Logger.log('OK: ' + message);
  }

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
      set.ingredients.appendRow(['Test Recipe ' + i, 'Flour', 1, 'GRM', 1]);
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
    secondSet.ingredients.appendRow(['Test Recipe 27', 'Flour', 1, 'GRM', 1]);
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
      rollbackSet.ingredients.appendRow(['Broken Recipe', 'Flour', 1, 'GRM', 1]);
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

    Logger.log('runSelfTest: ALL ASSERTIONS PASSED');
  } finally {
    createdSheetNames.forEach(function(name) {
      var sheet = ss.getSheetByName(name);
      if (sheet) ss.deleteSheet(sheet);
    });
  }
}
