# BOM Studio v2 — Modify Mode Prototype

Static HTML/CSS/JavaScript website. Upload OLD and NEW Excel BOM files in your browser; processing does not upload BOM contents to a backend. For local testing, run `python -m http.server 8000` in this folder and open `http://localhost:8000`, or publish all extracted files in the repository root using GitHub Pages.

## New system-oriented workflow

- Adds **Modify Mode** immediately before **Main/Sub**, located by column header, not fixed position.
- Allowed values: `Add`, `Add Main`, `Add Sub`, `Delete`, `Change to Sub`, `Change to Main`, `Location, QTY change`.
- Matching OLD and NEW MPN/manufacturer stays the default identity rule. Uses NEW Main/Sub role for matches. Multiple independent changes on one component produce a Delete of the old record plus a new Add record (subject to manual validation).
- Main/Sub groups are grouped together and ordered Main first, then Sub (group matching ambiguities listed on Worksheet 2).
- No change-status fill colours. Highlights only retained OLD PCB rows absent from NEW and missing historical group alternatives, plus any independently blocked/uncertain group relationships in Section E.
- Sections **A–D** (summary, Modify Mode reference, special component checklist, critical IC review) are displayed on the **website only**.
- **Worksheet 2** has only **E. Main/Sub & Import Validation**, with a minimal exception table. Worksheets 3/4 remain Change Log and Needs Review.

## Important validation limitations

This remains a prototype. The browser report writer reconstructs Worksheet 1 from original cell values and styles; complex Excel merges, drawings, macros, formulas, named ranges, SBR-system import rules, and some style details are **not guaranteed to be preserved**. `.xlsm` inputs produce `.xlsx` output without VBA macros. Confirm with approved nonconfidential test data before attempting import. Do not use as an automatic BOM release approval.

For an absent OLD PCB, the retained row is assigned `Add` provisionally and highlighted for manual verification. For an unmatched historical group alternative, it is output once as a new `Add Main`/`Add Sub` record with a blank SBR ID, highlighted for review; the historical source row is not duplicated. Confirm these import conventions with your organization's SBR system.

ExcelJS is loaded from jsDelivr CDN, so the browser requires access to that domain. BOM workbook contents are processed in-browser by the application; don't publish proprietary code or test BOMs without company permission.


## Version 2.1 — automatic download

- Hero title simplified to **BOM Comparison for SBR BOM**.
- After selecting files and clicking **Compare BOMs**, the generated `.xlsx` report is downloaded automatically. No second download button is needed.
- Your browser may ask for download permission or use your configured Downloads folder. If the download is blocked, allow downloads for this site and compare again.

To publish via GitHub Pages, upload the **contents** of this folder to your repository root (with `index.html` at the root).
