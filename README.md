# BOM Studio — browser-only JavaScript prototype

## How to run

1. Keep `index.html`, `app.js`, `bom-engine.js`, `excel-report.js`, `style.css` together.
2. On your Windows PC, open PowerShell in this folder and run `python -m http.server 8000` (developer machine only). Visit http://localhost:8000.
3. Open the link, drag two .xlsx/.xlsm Excel files or Browse, click Compare BOMs, and download the report.
4. To share with coworkers without Python, host these static files using a **company-permitted static site host**. Coworkers require only a browser; hosting must be approved for your company's internal source code.

## Important constraints

- This prototype uses ExcelJS 4.4.0 from the jsDelivr CDN. Internet access and permission to load that URL are necessary; if it's blocked, have the library included locally by an approved software workflow before hosting. Never add a BOM upload server endpoint.
- Files are read in the browser with File/ArrayBuffer and are not uploaded by the app. The external library is third-party JavaScript running on the same page; review/pin/vendor dependencies under your organization's software-security policy before internal distribution. The app itself has no network upload or telemetry calls.
- Input: FIRST worksheet in OLD/NEW workbook. Both .xlsx and .xlsm are read as Excel workbooks; macros are neither run nor copied to the .xlsx output.
- Output: download via browser to browser-configured downloads folder (a website cannot silently save on the Desktop).
- Comparison is a JavaScript port of the Python v1.6 conservative matching rules. Not yet parity-validated on real Excel workbooks in a browser: inspect outputs and reconcile Python differences before engineering production usage.
- The original Excel workbook is read and rewritten by ExcelJS, which may not preserve all Excel features, images, shapes, macro content, exotic styles, formulas, validations or conditional formatting. Verify worksheet fidelity before production.
- PCB/PCBA output rows: no fill, manual review; yellow new/modified/replacements, grey removed old rows, red changed text.
- Summary & Review is worksheet 2 and has no frozen panes. Change Log and Needs Review follow.
- For confidential company BOMs, obtain any required approval before placing proprietary tool code onto a public hosting service. This is not a security bypass for endpoint or web filters.
