# Excel Job Search Tracker Generator

A Python script that builds an Excel workbook for tracking job applications. It sets up the sheets, the dropdowns, the status coloring, and a small dashboard, so all you do is fill in rows. Works with Sheets/Libre.

![Dashboard](assets/dashboard.png)

![Job Applications Table](assets/table.png)

## Why a script?

Everyone's job search is different, so instead of one fixed template this is a small script you shape to yourself. Change a few lists at the top (your statuses, platforms, colors) and the whole file rebuilds to match: dropdowns, dashboard, colors, and chart all stay in sync, with nothing to fix by hand. Every run produces a fresh, clean workbook, and the code is open (MIT) so you can see exactly what it does.

## What you get

- A **Dashboard** with a few KPIs (total applications, in progress, interviews, offers, response rate), a status breakdown, a bar chart, and a Next Actions summary (overdue, due today, next 7 days). Offers and acceptances are highlighted once you have any.
- A **Job Applications** table (a real Excel Table) with 18 columns, ordered so the ones you act on show first: company, title, source, work model, date applied, days elapsed, status, next action and its date, then link, resume version, currency, salary, employment type, rating, contacts, and notes. It has dropdowns, an AutoFilter, and colored statuses.
- A **Lookup Lists** sheet holding the dropdown options. Edit these in Excel to change what the dropdowns offer.

The table grows on its own: type in the first empty row and Excel extends the formatting, dropdowns, and the Days Elapsed formula down. The dashboard formulas read the whole Status column, so the KPIs stay correct no matter how many rows you add.

## Requirements

- Python 3.7 or newer
- xlsxwriter

## Usage

Install the dependency (just xlsxwriter):

```bash
pip install -r requirements.txt
```

Run the script:

```bash
python generate_tracker.py
```

Open the resulting `Job_Search_Tracker.xlsx` in Excel, Google Sheets, or LibreOffice; every formula recalculates when the file opens. You don't need Excel installed at all: upload the file to Google Drive and use it entirely in Google Sheets from every device for best experience.

## Configuration

Edit the constants at the top of `generate_tracker.py`:

- `INITIAL_ROWS`: how many empty rows to pre-format (default 200; the table still grows past this).
- `STATUS_STYLES`: your statuses and their colors in one place. It drives the Status dropdown, the dashboard breakdown, and the row coloring.
- The other dropdown lists (`PLATFORMS`, `WORK_MODELS`, `CURRENCIES`, `EMP_TYPES`, `NEXT_ACTIONS`, `RATINGS`): change them to match your own process.
- `IN_PROGRESS`, `INTERVIEWS`, `NO_RESPONSE`: the status groups the dashboard rolls up. Response rate is the share of applications whose status is not Applied, Withdrawn, or Ghosted.
- `SUCCESS`: the statuses treated as wins. Their breakdown row and chart column are highlighted once the count is above zero.
- `COLOR_PRIMARY`: the theme color.
- `DATE_FORMAT`: how dates display (default `dd/mm/yyyy`).
- `DATE_OVERDUE`, `DATE_TODAY`, `DATE_UPCOMING`: the fill and text colors for a Next Action Date that is overdue, due today, or coming up this week.

## Notes

- The example row uses real dates (relative to today), so the demo always looks current. Dates display as `dd/mm/yyyy`; change `DATE_FORMAT` for another style.
- All formulas use plain A1-style references, so they behave the same in Excel, Google Sheets, and LibreOffice.
- Next Action Date is highlighted automatically: amber when it is today, red when it is overdue. The dashboard also counts how many are overdue, due today, or coming up in the next 7 days.
- Running the script won't overwrite an existing `Job_Search_Tracker.xlsx`; it stops and asks you to move or rename the old file first, so you never lose data.
- The script is safe to `import` (it builds the file only when run directly). To generate from your own code, call `build_tracker(path)`.

## Contributing

This is a small personal project, shared in case it is useful to others. Bug reports and pull requests are welcome through GitHub Issues.

## License

MIT License. Copyright (c) 2026 Alex Rabinovich. See the [LICENSE](LICENSE) file.

You are free to use, change, and share this. The only condition is that the copyright notice (my name) stays in the LICENSE file that ships with the code.
