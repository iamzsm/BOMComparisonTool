# SBR BOM Comparison Tool v2
Thanks my $20 per month ChatGPT for building this all things.

## How to run

1. Open the link https://iamzsm.github.io/BOMComparisonTool/
2. drag two .xlsx/.xlsm Excel files or Browse, click Compare BOMs, and download the report.
3. Review Worksheet 2 before approving or releasing the BOM.

## Important constraints and Key Features

1. Drag & Drop / Browse: Import OLD and NEW BOM files in .xlsx or .xlsm format.

2. Automatic BOM Comparison: Identify unchanged, modified, added, removed and replaced components.

3. Component Identification: Compare manufacturer names and part numbers alongside quantities and locations.

4. Main/Sub Group Handling: Preserve historical alternative components and synchronize shared quantity/location changes.

5. Colour-Coded Results: Yellow for additions and modifications, grey for removals, and red text for changed values.

6. Special Component Review: Identify PCB, connectors, terminals, holders, batteries and related components for manual verification.

7. Critical IC Review: Flag MCU, Flash, DC-DC, LDO and module-related components for additional engineering review.

8. Engineering Summary: Generate a comparison overview, colour legend, critical-component checklist and manual-review notices.

9. Excel Report Export: Download a structured workbook with four worksheets.

## Version History
v2: Because some useless requirements were mentioned to achieve useless automation on this fucking old and useless system, I had to edit it so my kind, helpful, best-in-the-world mentor wouldn't be disturbed again by that fucking annoying colleague from the dusty Shenzhen office...
to be serious i just add the modify mode column inside and let the main and sub in order. cancel the colour highlighted (remain the highlight if need eng double check)
