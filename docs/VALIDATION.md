# Validation record

Validated locally on Windows on 23 October 2026 with Python 3.12.10 and openpyxl 3.1.5.

- Editable package installation succeeded.
- 20 unittest tests passed, including aggregate-total rollback and rejected control-character export regressions.
- The complete local HTTP demo ran twice and reconciled to the expected counts and USD 5,645.20 total.
- The first demo recovered from simulated HTTP 429 and 503 responses.
- Reimporting the same source inserted zero orders and preserved the total.
- An actual `.xlsx` input with native date and numeric cells imported successfully.
- A simulated API authorization failure returned exit code 1 without creating a database.
- Exported workbook values, date types, text handling, sheet names and total reconciliation were checked programmatically.
- Sample workbook sheets were converted with LibreOffice for visual inspection. Portfolio media is generated from the verified summary data and labeled as a synthetic demo.

Not claimed: production API integration, GitHub-hosted CI execution, live scheduled operation, production-scale performance, or deployment to a client system. The HTML report is included for local viewing; it is not a hosted web application.
