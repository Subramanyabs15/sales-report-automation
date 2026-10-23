# Upwork portfolio draft

## Title
Python Sales Automation: REST API to Validated Excel Reports

## Role
Python developer — data import, API integration, validation and reporting

## Description
Built a Python automation tool that imports sales orders from paginated REST APIs, CSV and Excel, validates records, prevents repeat imports from duplicating orders, and generates an Excel report with a clear review trail.

The included demo handles simulated rate limits and temporary server failures. Conflicting order IDs and invalid records are flagged rather than silently changed. Outputs include monthly sales totals, clean orders, rejected records, operational logs and a browser-readable summary.

Demonstrated with synthetic data: 18 input records, 12 accepted orders, two duplicates skipped and four records flagged. Repeating the import adds zero orders and keeps total sales at USD 5,645.20.

Personal portfolio project. All example data is synthetic.

## Skills to select if available
Python; API Integration; Data Processing; Automation; Microsoft Excel

## Media
- `report-preview.png`: a visual summary generated from verified demo results.
- `before-after.png`: generated from the verified first-run and repeat-run results.
- `sample-output/first-run/sales-report.xlsx`: example workbook.
- `DEMO.md`: recording script and expected results.

## Services this can support
- Connect a documented REST API to CSV or Excel reporting.
- Clean and validate recurring order exports.
- Add duplicate checks, retries and logging to Python scripts.
- Set up a scheduled import with clear handover instructions.

## Publication checklist
- [x] Working source code and synthetic sample data.
- [x] Local tests and repeat-import demonstration.
- [x] Setup, API contract and business rules documented.
- [x] Portfolio description drafted without invented client or performance claims.
- [ ] User runs the tool locally and reviews the outputs.
- [ ] Publish source to a chosen GitHub repository.
- [ ] Record a short demonstration video if desired.
- [ ] Publish this entry on Upwork with the generated media.
