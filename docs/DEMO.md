# 90-second demonstration walkthrough

Use only the included synthetic dataset. Start with a fresh demo destination:

```bash
python -m sales_report.demo --output output/recording-demo
```

If that destination already contains data, choose another empty directory for the first-run counts below.

## 0–15 seconds: business problem

Open `examples/orders.csv`. Show repeated IDs, an ambiguous date, a negative quantity, an unexpected currency and the conflicting ORD-003 row.

Say: “This tool imports order exports, validates them, and creates a report without counting the same order twice.”

## 15–40 seconds: run the API import

Run the command. Show the first-run console output: 18 records received, 12 orders inserted, two duplicates skipped, four rejected, two API retries.

Explain that the local API simulates a rate limit and temporary service failure. The retry behavior is real; the failures and records are synthetic.

## 40–65 seconds: show the output

Open the first printed report directory's `report.html`, then `sales-report.xlsx`. Show Summary, Clean orders and Import review.

Monthly totals: August 1095.50, September 1966.20, October 2583.50 USD. Total: 5645.20 USD.

## 65–80 seconds: prove safe reruns

Show the second-run output: zero new orders, 14 duplicates, four rejected records, and the same 5645.20 total. An incoming conflicting ID is flagged and does not overwrite accepted data.

## 80–90 seconds: handover

Show the README, test command and API adapter contract.

Say: “The same import can run on a schedule. A client-specific connector would map your API or spreadsheet into this documented order schema.”

Do not claim production deployment, actual client savings or tested performance at a scale not measured.
