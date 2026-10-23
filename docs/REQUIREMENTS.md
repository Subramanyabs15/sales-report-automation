# Requirements selected from the job sample

Reviewed Upwork listings on 23 October 2026. These are examples of requested capabilities, not evidence of completed client work, guaranteed demand or obtainable rates.

| Listing | Relevant requirements | Demonstrated here |
|---|---|---|
| [AI lead qualification and CRM integration](https://www.upwork.com/jobs/~022113421236273390699/) | REST integrations, retries, no duplicates, documentation and demo | Paginated API adapter, bounded retries, persistent duplicate protection, demo and tests. AI extraction and CRM connectors are not included. |
| [AI-powered CRM automation](https://www.upwork.com/jobs/~022112957305788424217/) | Validation, duplicate detection, API workflows | Deterministic validation and duplicate/conflict decisions. No AI, AWS or CRM deployment is claimed. |
| [Python data pipeline deployment](https://www.upwork.com/jobs/~022113347765012951452/) | Python/API workflow, GitHub Actions, secrets and run instructions | Python CLI, environment-token interface, CI definition and scheduling notes. No hosting deployment is claimed. |
| [Clean hotel booking spreadsheets](https://www.upwork.com/jobs/~022113027544652278345/) | Clean records, uncertainty flags, summaries and provenance | Input-record references, explicit date rules, rejected-row review, monthly summaries. This is a sales schema, not a hotel-booking adapter. |
| [Data pipeline operator](https://www.upwork.com/jobs/~022113548294242595228/) | Reproducible runs, spreadsheet validation, output checks | Local synthetic demo, documented inputs and reconciled outputs. |

## Version 1 acceptance criteria

1. Local API demo imports three pages despite a 429 and a 503 response.
2. The 18-record fixture produces 12 new orders, two duplicate skips, four rejections and sales of 5645.20 USD.
3. The same input run again adds zero orders, reports 14 duplicate skips and leaves sales unchanged.
4. Conflicting IDs never overwrite an accepted record.
5. Source failure returns a nonzero exit status without importing partial API results.
6. A workbook, HTML report, summary JSON and review JSON are generated per successful export.
7. Installation and reproduction require no external account or API key for the demo.

## Future extensions

Prioritize an extension only when a target job needs it: configurable field mapping, cursor pagination, CRM connector, OAuth token refresh, corrected-order updates, line-item schemas or a hosted upload interface.
