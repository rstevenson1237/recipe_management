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
  `DDMMYY (3)`, ... if a date fills up). All three sheets in a trio are always created, repaired,
  and written together, joined on the `Name` column, which is unique across every recipe ever
  entered.
- **Helper Data** (hidden) — the allowed ingredient list and units of measure that both the
  dialog and the server validate against.

## Export

**🍔 Recipe Tools > Export All Recipes (Print)** opens a preview of every recipe on file
(across every batch), one per page, ingredients on the left and instructions on the right, with
the logo at the bottom left of each page. Use the Print button in the preview to print or
save as PDF from the browser — nothing is written to Drive.

## Regenerating the embedded logo

`Export.gs` reads the logo from `Logo.html`, which holds the image as base64 text (Apps Script
can't ship a `.png` file directly). To swap the logo:

```sh
base64 -w0 AtlasLogo_black_60x24.png > Logo.html
```

Keep `AtlasLogo_black_60x24.png` in the repo as the source of truth.

## Checking changes before pasting them into Apps Script

None of `package.json`, `tsconfig.json`, `.gitignore`, or `scripts/` deploy anywhere — the Apps
Script editor only ever sees the `.gs`/`.html` files, pasted in directly (there's no clasp setup
in this repo; add a `.claspignore` for these files if that changes). They exist purely to catch
bugs before that paste, because Apps Script's errors are otherwise runtime-only: a call to a
method that doesn't exist on `Range`/`Sheet`/`Spreadsheet` compiles fine and only fails the moment
a user clicks something.

```sh
npm install
npm run check
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
