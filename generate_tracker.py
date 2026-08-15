"""
Generate an Excel job-search tracker (dashboard + application table).

Copyright (c) 2026 Alex Rabinovich. MIT License.

Safe to import: the workbook is only built when the file is run directly, or
when you call build_tracker(path) yourself.
"""

import os
from datetime import date, timedelta

import xlsxwriter as xlw
from xlsxwriter.utility import xl_col_to_name
from xlsxwriter.exceptions import FileCreateError

# --- Configuration -----------------------------------------------------------
OUTPUT_FILENAME = 'Job_Search_Tracker.xlsx'
INITIAL_ROWS = 200  # empty rows pre-formatted in the table (it still grows past this in Excel)
TABLE_NAME = 'Applications'
DATE_FORMAT = 'dd/mm/yyyy'  # how dates display in the sheet

# Dropdown lists
PLATFORMS = ['LinkedIn', 'Company Website', 'Telegram', 'WhatsApp', 'Referral', 'Indeed', 'Glassdoor', 'Other']
WORK_MODELS = ['Remote', 'Hybrid', 'On-site']
# Status -> (fill color, text color). One place controls the dropdown order,
# the dashboard breakdown order, and the row coloring. Add or edit a status here
# and it flows everywhere; no need to touch colors further down.
STATUS_STYLES = {
    'Applied': ('#FFF2CC', '#7F6000'),
    'Under Review': ('#FFE699', '#7F6000'),
    'HR Screen': ('#D9E1F2', '#1F4E78'),
    'Technical Interview': ('#BDD7EE', '#1F4E78'),
    'Take-Home Task': ('#E2EFDA', '#375623'),
    'Final Interview': ('#FCE4D6', '#C65911'),
    'Offer Received': ('#C6EFCE', '#006100'),
    'Accepted': ('#A9D08E', '#375623'),
    'Rejected': ('#F8CBAD', '#C65911'),
    'Withdrawn': ('#E7E6E6', '#595959'),
    'Ghosted': ('#D9D9D9', '#595959'),
}
STATUSES = list(STATUS_STYLES)
CURRENCIES = ['ILS (\u20aa)', 'USD ($)', 'EUR (\u20ac)', 'GBP (\u00a3)', 'CAD ($)', 'Other']
EMP_TYPES = ['Full-time', 'Part-time', 'Contract / B2B']
NEXT_ACTIONS = ['Follow-up Email', 'HR Call', 'Technical Test Submission',
                'Interview Prep', 'Wait for Feedback', 'Offer Review']
RATINGS = [1, 2, 3, 4, 5]

# Which statuses roll up into each dashboard KPI.
# If you add a status to STATUS_STYLES above, decide whether it also belongs in
# one of these groups; otherwise it shows in the table but not in these KPIs.
# NO_RESPONSE is business logic (who counts as "never replied"), so it stays manual.
IN_PROGRESS = ['Applied', 'Under Review', 'HR Screen', 'Technical Interview', 'Take-Home Task', 'Final Interview']
INTERVIEWS = ['HR Screen', 'Technical Interview', 'Take-Home Task', 'Final Interview']
NO_RESPONSE = ['Applied', 'Withdrawn', 'Ghosted']

# Statuses treated as wins: highlighted in the breakdown and the chart, but only
# once their count is above zero (no color while you have none of them yet).
SUCCESS = ('Offer Received', 'Accepted')

# Theme colors
COLOR_PRIMARY = '#1F4E78'
COLOR_SUBHEADER_BG = '#D9E1F2'

# Next Action Date highlights (fill, text): overdue is red, today is amber, upcoming is blue.
DATE_OVERDUE = ('#FFC7CE', '#9C0006')
DATE_TODAY = ('#FFEB9C', '#9C6500')
DATE_UPCOMING = ('#DDEBF7', '#1F4E78')

LAST_ROW = INITIAL_ROWS + 1  # header + body rows, 1-based


def build_tracker(filename):
    wb = xlw.Workbook(filename)
    wb.set_properties({'author': 'Alex Rabinovich'})
    ws_dash = wb.add_worksheet('Dashboard')
    ws_app = wb.add_worksheet('Job Applications')
    ws_lookups = wb.add_worksheet('Lookup Lists')

    # --- Formats -------------------------------------------------------------
    hdr_fmt = wb.add_format({'bold': True, 'font_color': '#FFFFFF', 'bg_color': COLOR_PRIMARY,
                             'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    sub_hdr_fmt = wb.add_format({'bold': True, 'font_color': COLOR_PRIMARY, 'bg_color': COLOR_SUBHEADER_BG,
                                 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    label_fmt = wb.add_format({'border': 1, 'valign': 'vcenter'})
    count_fmt = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
    kpi_fmt = wb.add_format({'bold': True, 'font_size': 14, 'font_color': COLOR_PRIMARY,
                             'border': 1, 'align': 'center', 'valign': 'vcenter'})
    pct_fmt = wb.add_format({'bold': True, 'font_size': 14, 'font_color': COLOR_PRIMARY, 'num_format': '0.0%',
                             'border': 1, 'align': 'center', 'valign': 'vcenter'})

    # Body formats. All centered so typed data lines up; notes stay left for readability.
    body_center = wb.add_format({'align': 'center', 'valign': 'vcenter'})
    body_date = wb.add_format({'num_format': DATE_FORMAT, 'align': 'center', 'valign': 'vcenter'})
    body_num = wb.add_format({'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter'})
    body_url = wb.add_format({'font_color': '#0563C1', 'underline': 1, 'align': 'center', 'valign': 'vcenter'})
    body_notes = wb.add_format({'align': 'left', 'valign': 'vcenter'})

    # --- Lookup lists --------------------------------------------------------
    # Written to the Lookup Lists sheet; source ranges are derived, not hand-counted.
    lookups = [
        ('Platform', PLATFORMS),
        ('Work Model', WORK_MODELS),
        ('Status', STATUSES),
        ('Currency', CURRENCIES),
        ('Employment Type', EMP_TYPES),
        ('Next Action', NEXT_ACTIONS),
        ('Rating', RATINGS),
    ]
    source_range = {}
    for col, (title, options) in enumerate(lookups):
        ws_lookups.write(0, col, title, sub_hdr_fmt)
        for row, value in enumerate(options, start=1):
            ws_lookups.write(row, col, value)
        letter = xl_col_to_name(col)
        source_range[title] = f"'Lookup Lists'!${letter}$2:${letter}${len(options) + 1}"
    ws_lookups.set_column(0, len(lookups) - 1, 18)

    # --- Application table ---------------------------------------------------
    # (header, body format, dropdown source or None, column width)
    # Order matters: the first columns are the ones you want visible without scrolling.
    app_columns = [
        ('Company Name', body_center, None, 24),
        ('Job Title', body_center, None, 24),
        ('Source / Platform', body_center, 'Platform', 18),
        ('Location / Work Model', body_center, 'Work Model', 18),
        ('Date Applied', body_date, None, 13),
        ('Days Elapsed', body_center, None, 13),
        ('Status', body_center, 'Status', 20),
        ('Next Action', body_center, 'Next Action', 18),
        ('Next Action Date', body_date, None, 16),
        ('Job Posting URL', body_url, None, 26),
        ('Resume Version', body_center, None, 22),
        ('Currency', body_center, 'Currency', 14),
        ('Expected/Offered Salary', body_num, None, 18),
        ('Employment Type', body_center, 'Employment Type', 20),
        ('Personal Match (1-5)', body_center, 'Rating', 15),
        ('Contact Name & Role', body_center, None, 22),
        ('Contact Info', body_center, None, 24),
        ('Notes & Interview Questions', body_notes, None, 45),
    ]
    headers = [c[0] for c in app_columns]
    date_col = xl_col_to_name(headers.index('Date Applied'))
    status_col = xl_col_to_name(headers.index('Status'))
    company_col = xl_col_to_name(headers.index('Company Name'))
    next_col = xl_col_to_name(headers.index('Next Action Date'))
    days_idx = headers.index('Days Elapsed')

    ws_app.set_row(0, 28)
    ws_app.add_table(0, 0, LAST_ROW - 1, len(app_columns) - 1, {
        'name': TABLE_NAME,
        'style': 'Table Style Medium 2',
        'banded_rows': True,
        'columns': [{'header': header, 'header_format': hdr_fmt, 'format': fmt} for header, fmt, _, _ in app_columns],
    })
    ws_app.freeze_panes(1, 0)

    for idx, (header, fmt, rule, width) in enumerate(app_columns):
        ws_app.set_column(idx, idx, width, fmt)  # column-level format so data typed later is centered too
        if rule:
            rng = f'{xl_col_to_name(idx)}2:{xl_col_to_name(idx)}{LAST_ROW}'
            ws_app.data_validation(rng, {'validate': 'list', 'source': source_range[rule]})

    # Days elapsed since the application date (blank until a date is entered)
    for row in range(2, LAST_ROW + 1):
        formula = f'=IF(ISBLANK({date_col}{row}),"",TODAY()-{date_col}{row})'
        ws_app.write_formula(row - 1, days_idx, formula, body_center)

    # Color each status (colors come from STATUS_STYLES at the top of the file)
    status_range = f'{status_col}2:{status_col}{LAST_ROW}'
    for status, (bg, font) in STATUS_STYLES.items():
        fmt = wb.add_format({'bg_color': bg, 'font_color': font, 'bold': True})
        ws_app.conditional_format(status_range, {'type': 'cell', 'criteria': 'equal to',
                                                 'value': f'"{status}"', 'format': fmt})

    # Highlight the Next Action Date: red if it is already past, amber if it is today.
    # Blank cells stay untouched; the two rules never overlap (a date can't be both).
    next_range = f'{next_col}2:{next_col}{LAST_ROW}'
    overdue_fmt = wb.add_format({'bg_color': DATE_OVERDUE[0], 'font_color': DATE_OVERDUE[1]})
    today_fmt = wb.add_format({'bg_color': DATE_TODAY[0], 'font_color': DATE_TODAY[1]})
    ws_app.conditional_format(next_range, {'type': 'formula',
                                           'criteria': f'=AND(${next_col}2<>"",${next_col}2<TODAY())',
                                           'format': overdue_fmt})
    ws_app.conditional_format(next_range, {'type': 'formula',
                                           'criteria': f'=${next_col}2=TODAY()',
                                           'format': today_fmt})

    # Example row. Dates are real dates relative to today, so the demo always looks current.
    example = {
        'Company Name': 'Wiz',
        'Job Title': 'Senior Product Manager',
        'Job Posting URL': 'https://www.wiz.io/careers/spm-tel-aviv',
        'Source / Platform': 'LinkedIn',
        'Location / Work Model': 'Hybrid',
        'Date Applied': date.today() - timedelta(days=5),
        'Status': 'HR Screen',
        'Next Action': 'HR Call',
        'Next Action Date': date.today() + timedelta(days=3),
        'Resume Version': 'CV_ProductManager_Tech_v3.pdf',
        'Currency': 'ILS (\u20aa)',
        'Expected/Offered Salary': 32000,
        'Employment Type': 'Full-time',
        'Personal Match (1-5)': 5,
        'Contact Name & Role': 'Rachel Cohen (Lead HR)',
        'Contact Info': 'rachel.cohen@wiz.io / @rachel_wiz',
        'Notes & Interview Questions': 'Initial 30-min HR screen scheduled. Focus on cloud security background.',
    }
    for idx, (header, fmt, rule, width) in enumerate(app_columns):
        if header not in example:
            continue
        if header == 'Job Posting URL':
            ws_app.write_url(1, idx, example[header], body_url)
        else:
            ws_app.write(1, idx, example[header], fmt)

    # --- Dashboard -----------------------------------------------------------
    # Whole-column references so the KPIs cover the table at any size. The column
    # letters come from the table layout above, so reordering columns can't break this.
    status_column = f"'Job Applications'!{status_col}:{status_col}"
    # Total = filled Company Name cells; the -1 drops the header row, which a
    # whole-column COUNTA would otherwise count as one.
    total = f"COUNTA('Job Applications'!{company_col}:{company_col})-1"

    def count_of(names):
        return '+'.join(f'COUNTIF({status_column},"{name}")' for name in names)

    ws_dash.set_column(0, 4, 22)
    ws_dash.write('A1', 'Job Search Analytics', sub_hdr_fmt)

    # KPI cards: each label sits on one row with its value on the row directly below.
    kpis = [
        ('Total Applications', f'={total}', kpi_fmt),
        ('In Progress', f'={count_of(IN_PROGRESS)}', kpi_fmt),
        ('Interviews & Tests', f'={count_of(INTERVIEWS)}', kpi_fmt),
        ('Offers Received', f'={count_of(["Offer Received", "Accepted"])}', kpi_fmt),
        ('Response Rate', f'=IF({total}<=0,0,(({total})-({count_of(NO_RESPONSE)}))/({total}))', pct_fmt),
    ]
    for col, (label, formula, fmt) in enumerate(kpis):
        ws_dash.write(2, col, label, sub_hdr_fmt)
        ws_dash.write_formula(3, col, formula, fmt)

    # Status breakdown table. The header sits on this row; counts start right below it.
    # Success rows are tinted only when their count is above zero (conditional format).
    breakdown_header_row = 6
    first_data_row = breakdown_header_row + 1
    ws_dash.write(f'A{breakdown_header_row}', 'Status', sub_hdr_fmt)
    ws_dash.write(f'B{breakdown_header_row}', 'Count', sub_hdr_fmt)
    for i, status in enumerate(STATUSES):
        row = first_data_row + i
        ws_dash.write(f'A{row}', status, label_fmt)
        ws_dash.write_formula(f'B{row}', f'=COUNTIF({status_column}, A{row})', count_fmt)
        if status in SUCCESS:
            bg, font = STATUS_STYLES[status]
            win_fmt = wb.add_format({'bg_color': bg, 'font_color': font, 'bold': True})
            ws_dash.conditional_format(f'A{row}:B{row}', {'type': 'formula',
                                                          'criteria': f'=$B${row}>0', 'format': win_fmt})

    last_data_row = first_data_row + len(STATUSES) - 1

    # Chart. Success columns get their status color; every other column stays the theme
    # color. A zero-count column has no height, so a success color only shows once count > 0.
    points = [{'fill': {'color': STATUS_STYLES[s][0]}} if s in SUCCESS else {} for s in STATUSES]
    chart = wb.add_chart({'type': 'column'})
    chart.add_series({
        'categories': f"='Dashboard'!$A${first_data_row}:$A${last_data_row}",
        'values': f"='Dashboard'!$B${first_data_row}:$B${last_data_row}",
        'fill': {'color': COLOR_PRIMARY},
        'points': points,
        'data_labels': {'value': True},
    })
    chart.set_title({'name': 'Applications by Status'})
    chart.set_legend({'none': True})
    chart.set_size({'width': 560, 'height': 320})
    ws_dash.insert_chart('D6', chart)

    # Next Actions summary: follow-ups that are overdue, due today, or coming up this week.
    # Counts read the Next Action Date column; each row is tinted only when its count > 0.
    na_range = f"'Job Applications'!{next_col}:{next_col}"
    na_header_row = 19
    ws_dash.write(f'A{na_header_row}', 'Next Actions', sub_hdr_fmt)
    ws_dash.write(f'B{na_header_row}', 'Count', sub_hdr_fmt)
    na_buckets = [
        ('Overdue', f'=COUNTIF({na_range},"<"&TODAY())', DATE_OVERDUE),
        ('Due today', f'=COUNTIF({na_range},TODAY())', DATE_TODAY),
        ('Next 7 days', f'=COUNTIFS({na_range},">"&TODAY(),{na_range},"<="&(TODAY()+7))', DATE_UPCOMING),
    ]
    for i, (label, formula, colors) in enumerate(na_buckets):
        r = na_header_row + 1 + i
        ws_dash.write(f'A{r}', label, label_fmt)
        ws_dash.write_formula(f'B{r}', formula, count_fmt)
        bg, font = colors
        cf_fmt = wb.add_format({'bg_color': bg, 'font_color': font, 'bold': True})
        ws_dash.conditional_format(f'A{r}:B{r}', {'type': 'formula', 'criteria': f'=$B${r}>0', 'format': cf_fmt})

    wb.close()


def main():
    if os.path.exists(OUTPUT_FILENAME):
        print(f"'{OUTPUT_FILENAME}' already exists. Rename or move it first so its data isn't overwritten.")
        raise SystemExit(1)
    try:
        build_tracker(OUTPUT_FILENAME)
    except FileCreateError:
        print(f"Could not write '{OUTPUT_FILENAME}'. If it is open in Excel, close it and run again.")
        raise SystemExit(1)
    print(f"Success: '{OUTPUT_FILENAME}' generated.")


if __name__ == '__main__':
    main()
