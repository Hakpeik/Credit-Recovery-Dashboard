# Credit Recovery dashboard automation

The dashboard is a static page (`index.html` + `styles.css` + `app.js`) that reads
`data.js`. Everything in `data.js` is generated from the Excel workbooks by
`build_data.py`, so the page also works when opened directly from disk
(`file://`).

## One-time / manual rebuild

```bat
cd "C:\Users\AIO\Desktop\Credit Admin\Loan Recovery Section"
python automation\build_data.py
```

Then open `index.html` in a browser (or press the reload button in the header).

## Automatic updates (recommended)

Double-click **`automation\start_watch.bat`**.

It starts `build_data.py --watch`, which re-scans the folders below and rewrites
`data.js` whenever a workbook is added or changed. The open dashboard reloads
`data.js` by itself every 45 seconds, so new files appear without a manual
refresh.

Optional flags:

```bat
python automation\build_data.py --watch --interval 20
```

`--interval` is the scan interval in seconds (default 30).

## Source folders

| Dashboard section  | Source |
| ------------------ | ------ |
| Overview           | `Data Dashboard.xlsx` → Loan Listing Report, columns BE:CQ, as-of date from cell CO7 |
| Daily Overdue Data | `Credit Recovery Workflow\Daily Overdue Data\<Loan Type>\*.xlsx` (Term Loan, Overdraft, Credit Card) |
| Daily Consolidate  | `Credit Recovery Workflow\Daily Consolidate for Overdue Report\*.xlsx` |
| NPL Movement       | `NPL Movement\<Year>\Recovery MMMYYYY.xlsx` → `NPL Ratio` sheet |

Drop new dated workbooks into the relevant folder; the file name's date is used
for the timeline. Empty workbooks are skipped for detail tables, but their
summary rows are still tracked.

## Cache

Parsed workbooks are cached in `automation\.cache.json` keyed by path, size and
modified time. A warm rebuild takes a few seconds; delete the cache file to force
a full re-parse.
