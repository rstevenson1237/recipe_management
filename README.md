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
