# Recipe Management

A Google Apps Script add-on for a recipe spreadsheet: a Dashboard sheet, hidden date-keyed
database sheets, a validated entry dialog, and a print-ready recipe export.

## One-time setup (new workbook, or after pulling this update)

1. Open the spreadsheet, reload it, and use the **🍔 Recipe Tools** menu that appears
   (Apps Script menus are added by `onOpen`, which only runs on a fresh load).
2. Run **🍔 Recipe Tools > Setup / Repair Dashboard**. This creates/repairs the `Dashboard`
   and `Helper Data` sheets. It's safe to re-run any time.
3. **Populate Helper Data** (hidden sheet): column D ("Item Data") should list every allowed
   ingredient name; columns A/B/C hold the allowed Weight/Volume/Each units of measure (seeded
   with a starter set on first run). The entry dialog and server both validate against these
   lists, so nothing outside them will be accepted.
4. **Add the "New Recipe" button to the Dashboard** (one-time, per workbook):
   - `Insert > Drawing`, draw a button, click **Save and Close**.
   - Click the drawing once more, open its **⋮** menu, choose **Assign script**, and enter
     `openRecipeApp`.
   - Drag the drawing into the marked box on the Dashboard (to the right of the instructions).
   - If you'd rather skip this, the **🍔 Recipe Tools > Open Recipe UI** menu item always works.

## How the data is organized

- **Dashboard** — usage instructions, the New Recipe button, and a live list of every recipe
  on file.
- **`DB - Headers <key>` / `DB - Ingredients <key>` / `DB - Instructions <key>`** (hidden) — one
  matched trio per up-to-25-recipe batch, keyed by entry date as `DDMMYY` (`DDMMYY (2)`,
  `DDMMYY (3)`, ... if a date fills up). A PDF import gets its own `IMPORT <DDMMYY>` key instead,
  which holds every recipe in that file however many there are. All three sheets in a trio are
  always created, repaired, and written together, joined on the `Name` column, which is unique
  across every recipe ever entered.
  Ingredient rows are `Name | Ingredient | Qty | U of M | Ingredient Order | Preparation`.
  **Preparation** is optional free text for the prep method (`minced`), printed on the export as a
  single `Ingredient Name, Preparation` field. It was added after the first release, so
  `repairSheetColumns_()` adds the header to ingredient sheets that predate it the next time the
  workbook is opened — existing rows keep their data and read back with an empty Preparation.
- **Helper Data** (hidden) — the allowed ingredient list and units of measure that both the
  dialog and the server validate against.

## Renaming a recipe

Pick it from **Edit an Existing Recipe**, change the Recipe Name, and Update. The recipe's header,
ingredient and instruction rows and its Dashboard row are all re-keyed onto the new name; the new
name still has to be unique across the workbook.

## Export

**🍔 Recipe Tools > Export All Recipes (Print)** opens a preview of every recipe on file
(across every batch), one per page, ingredients on the left and instructions on the right, with
the logo at the top right of each page. Use the Print button in the preview to print or
save as PDF from the browser — nothing is written to Drive.

Each page is stamped `Export v<n>` (`EXPORT_FORMAT_VERSION` in `Export.gs`) and carries a tiny
machine-readable block, `RMDATA1[<base64 JSON>]`, holding that recipe's fields exactly. It is what
lets an exported PDF be imported back losslessly, so it must stay printable text — pale and 5px,
not hidden, since `display:none` text never reaches the PDF. **Bump `EXPORT_FORMAT_VERSION`
whenever a field is added to what the export prints**, so the import can tell what it's reading.

## Import

**🍔 Recipe Tools > Import Recipes from PDF...** reads a PDF produced by the export and writes every
recipe it contains into one new `IMPORT <DDMMYY>` batch — an import is not limited to 25 recipes.

The PDF is parsed in the dialog by pdf.js (loaded from a CDN), so nothing is uploaded to Drive and
no extra OAuth scope, advanced service or manifest is needed. The embedded data block is used when
present; a PDF printed before that existed (v1) is rebuilt from the printed layout instead, which
degrades field by field — anything it can't read is left blank and flagged in the review step rather
than failing the file. A comma after an ingredient name is only read as a Preparation on a v2+ PDF.

Before anything is written, the dialog lists what it found with per-recipe warnings and lets you
deselect rows. On import:

- a recipe whose name already exists is **skipped** and reported (`Name` is the join key across the
  three DB sheets, so two recipes can never share one);
- ingredient names and units of measure the workbook doesn't list yet are **added to Helper Data**,
  so imported recipes open cleanly in the entry dialog afterwards.

## Clearing the workbook

**🍔 Recipe Tools > Clear ALL Recipes...** deletes every `DB - ...` sheet and empties the Dashboard's
recipe list, behind a summary dialog and a prompt that requires typing `DELETE`. It cannot be undone.
`Helper Data` is deliberately kept — it's the curated ingredient/U of M configuration, not recipe
data.

## Regenerating the embedded logo

`Export.gs` reads the logo from `Logo.html`, which holds the image as base64 text (Apps Script
can't ship a `.png` file directly). To swap the logo:

```sh
base64 -w0 AtlasLogo_black_60x24.png > Logo.html
```

Keep `AtlasLogo_black_60x24.png` in the repo as the source of truth.

## Pasting a change in: copy every file it touches

The Apps Script project must hold all eleven deployed files, under these exact names:

`Code.gs`, `Save.gs`, `Sheets.gs`, `Export.gs`, `Import.gs`, `Reset.gs`, and
`Index.html`, `Export.html`, `Import.html`, `Logo.html`.

Apps Script puts every `.gs` file into one shared global scope, so a function defined in one
file is callable from all the others, and features here are deliberately split that way — the
import path in `Import.gs` calls `cleanPreparation_()` from `Save.gs`, the sheet helpers from
`Sheets.gs`, and so on. The cost of that is that a *partial* paste breaks at runtime rather
than at paste time: the editor happily saves a file calling a function no other file defines,
and the call only fails when someone clicks the thing that runs it.

That is what

```
Import failed: ReferenceError: cleanPreparation_ is not defined
```

means — `Import.gs` is present and current, but `Save.gs` in the project is an older copy from
before that helper existed. The fix is to re-paste the stale file, not to change the import
code. When a change spans several files (adding the `Preparation` column touched `Save.gs`,
`Sheets.gs`, `Export.gs`, `Code.gs`, `Index.html` and `Export.html` at once), paste all of
them, in any order, before running anything — and add any new file first, since a missing file
fails the same silent way.

## Checking changes before pasting them into Apps Script

None of `package.json`, `tsconfig.json`, `.gitignore`, or `scripts/` deploy anywhere — the Apps
Script editor only ever sees the `.gs`/`.html` files, pasted in directly (there's no clasp setup
in this repo; add a `.claspignore` for these files if that changes). They exist purely to catch
bugs before that paste, because Apps Script's errors are otherwise runtime-only: a call to a
method that doesn't exist on `Range`/`Sheet`/`Spreadsheet` compiles fine and only fails the moment
a user clicks something.

```sh
npm install
npm run check   # every .gs file against the real Apps Script API
npm test        # the PDF-parsing functions inside Import.html
```

This typechecks every `.gs` file against the real API definitions (`@types/google-apps-script`),
so a bad method name is a build-time error instead of a `TypeError` in front of a user. New `.gs`
files are picked up automatically. A function whose parameters aren't annotated with `@param`
JSDoc silently loses this checking for that function (TypeScript can't check what it can't infer
a type for) — so any new function that takes a `Sheet`/`Spreadsheet`/`Range` should have it
JSDoc-annotated with the real `GoogleAppsScript.Spreadsheet.*` type, the same way the existing
functions in `Sheets.gs`/`Save.gs`/`Export.gs`/`Code.gs` are. Loosely-shaped data (the payload
from the dialog, row values from `getValues()`) is intentionally typed `{*}` rather than modeled
in full — the point of this check is the Apps Script API surface, not our own data shapes.

`npm test` covers the one part of this project with logic that can run outside Apps Script: the PDF
parsers in `Import.html`. It extracts that file's `<script>` block into `.typecheck/` and runs it
against synthetic pdf.js output — a two-column page where an ingredient and an instruction share a
baseline, a v1 PDF whose commas must not become preparations, an overflow page, a recipe with
missing fields. If you change those parsers, run it; if you add a parsing rule, add a case. It has
no dependencies beyond Node, so it stays runnable without installing anything.

There is also `runSelfTest()` in `Sheets.gs`, run manually from the Apps Script editor: it covers the
sheet-set rollover, the lock-step invariant between the three DB sheets, write rollback, and the
`Preparation` column repair, using a throwaway `TEST<timestamp>` key namespace it deletes afterwards.
