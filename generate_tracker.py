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

# Pipeline stages (the Stage column) -> (fill, text), in blue shades that deepen
# along the funnel. Order here drives the Stage dropdown, the pipeline breakdown,
# and the drop-off analysis. Add or rename a stage here and it flows everywhere.
STAGE_STYLES = {
    'Applied':                 ('#EAF0F9', '#1F4E78'),
    'Application Review':      ('#DBE6F4', '#1F4E78'),
    'Recruiter / HR Screen':   ('#CBDDEF', '#1F4E78'),
    'Hiring Manager Screen':   ('#BDD7EE', '#1F4E78'),
    'Technical Interview':     ('#A6C8E8', '#1F4E78'),
    'Take-Home Task':          ('#8FB9E0', '#123553'),
    'Panel / Team Interview':  ('#6FA6D6', '#FFFFFF'),
    'Final Interview':         ('#4A8AC4', '#FFFFFF'),
    'Offer / Negotiation':     ('#2E6CA6', '#FFFFFF'),
}
STAGES = list(STAGE_STYLES)

# Outcomes (the Status column) -> (fill, text). In Progress is amber; wins are
# green; Rejected/Ghosted are red; Withdrawn is grey (a voluntary exit, not a loss).
OUTCOME_STYLES = {
    'In Progress':    ('#FFF2CC', '#7F6000'),
    'Offer Received': ('#C6EFCE', '#006100'),
    'Accepted':       ('#A9D08E', '#375623'),
    'Rejected':       ('#FFC7CE', '#9C0006'),
    'Ghosted':        ('#F4B7A6', '#843C0C'),
    'Withdrawn':      ('#E7E6E6', '#595959'),
}
STATUSES = list(OUTCOME_STYLES)

CURRENCIES = ['ILS (\u20aa)', 'USD ($)', 'EUR (\u20ac)', 'GBP (\u00a3)', 'CAD ($)', 'Other']
EMP_TYPES = ['Full-time', 'Part-time', 'Contract / B2B']
NEXT_ACTIONS = ['Follow-up Email', 'HR Call', 'Technical Test Submission',
                'Interview Prep', 'Wait for Feedback', 'Offer Review']
RATINGS = [1, 2, 3, 4, 5]

# Dashboard rollups. These reference the Stage and/or Status values above.
# EARLY_STAGES: still In Progress here = no response yet (used for Response Rate).
# INTERVIEW_STAGES: counted as "Interviews & Tests" when still In Progress.
# BAD_OUTCOMES: what the drop-off funnel counts (Withdrawn stays out on purpose).
EARLY_STAGES = ['Applied', 'Application Review']
INTERVIEW_STAGES = ['Recruiter / HR Screen', 'Hiring Manager Screen', 'Technical Interview',
                    'Take-Home Task', 'Panel / Team Interview', 'Final Interview']
SUCCESS = ('Offer Received', 'Accepted')
BAD_OUTCOMES = ('Rejected', 'Ghosted')

# Drop-off analysis. The donut fills in from the very first bad outcome. The focus
# hint below it appears once there are at least DROP_OFF_MIN bad outcomes in total
# (across all stages), and then points at the most frequent stage. These are gentle
# hypotheses to think about, not verdicts. Edit the text and the threshold freely.
DROP_OFF_MIN = 5
STAGE_HINTS = {
    'Applied': "Most drop-offs happen right after applying. This is often about how closely the resume matches each posting, or how relevant the roles are. Reviewing targeting and keywords is a good first step.",
    'Application Review': "Most drop-offs happen during application review. This often points to resume clarity and relevance; tailored, quantified bullets tend to get further.",
    'Recruiter / HR Screen': "Most drop-offs happen at the recruiter screen. Tightening your intro pitch, aligning on salary early, and showing clear interest in the role can help.",
    'Hiring Manager Screen': "Most drop-offs happen at the hiring-manager stage. Connecting your experience to the team's needs, and showing you researched the role, often makes the difference.",
    'Technical Interview': "Most drop-offs happen at the technical stage. Reviewing core skills and practicing explaining your thinking out loud can help here.",
    'Take-Home Task': "Most drop-offs happen at the take-home stage. Reading the brief closely, documenting your approach, and polishing the presentation matter as much as the result.",
    'Panel / Team Interview': "Most drop-offs happen at the team round. Preparing varied examples and showing how you collaborate can strengthen this stage.",
    'Final Interview': "Most drop-offs happen at the final round. By here it is usually less about skill and more about motivation and fit; prepare your 'why this team' story.",
    'Offer / Negotiation': "Most drop-offs happen around the offer, which is often outside your control. Responding promptly and keeping expectations aligned helps, and reaching this stage at all is a strong signal.",
}
NOT_ENOUGH_DATA = "Not enough closed applications yet for a reliable read. Keep tracking and this will sharpen over time."

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
    hint_fmt = wb.add_format({'text_wrap': True, 'valign': 'top', 'border': 1, 'font_color': COLOR_PRIMARY})

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
        ('Stage', STAGES),
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
    # Stage = where in the process; Status = the outcome. They live side by side so
    # marking an outcome never erases the stage it was reached at.
    app_columns = [
        ('Company Name', body_center, None, 24),
        ('Job Title', body_center, None, 24),
        ('Source / Platform', body_center, 'Platform', 18),
        ('Location / Work Model', body_center, 'Work Model', 18),
        ('Date Applied', body_date, None, 13),
        ('Days Elapsed', body_center, None, 13),
        ('Stage', body_center, 'Stage', 22),
        ('Status', body_center, 'Status', 16),
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
    stage_col = xl_col_to_name(headers.index('Stage'))
    status_col = xl_col_to_name(headers.index('Status'))
    company_col = xl_col_to_name(headers.index('Company Name'))
    next_col = xl_col_to_name(headers.index('Next Action Date'))
    days_idx = headers.index('Days Elapsed')

    # Whole-column references reused by the dashboard and the drop-off helper.
    app_stage = f"'Job Applications'!{stage_col}:{stage_col}"
    app_status = f"'Job Applications'!{status_col}:{status_col}"

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

    # Color the Stage (blue shades) and the Status (outcome colors) columns.
    def color_column(col_letter, styles):
        rng = f'{col_letter}2:{col_letter}{LAST_ROW}'
        for name, (bg, font) in styles.items():
            fmt = wb.add_format({'bg_color': bg, 'font_color': font, 'bold': True})
            ws_app.conditional_format(rng, {'type': 'cell', 'criteria': 'equal to',
                                            'value': f'"{name}"', 'format': fmt})

    color_column(stage_col, STAGE_STYLES)
    color_column(status_col, OUTCOME_STYLES)

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
        'Stage': 'Recruiter / HR Screen',
        'Status': 'In Progress',
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

    # --- Drop-off helper (backstage, on the Lookup Lists sheet) --------------
    # Per stage: count bad outcomes, add a tiny rank key to break ties deterministically,
    # and store the matching hint. The dashboard donut and focus text read from here.
    h = len(lookups) + 1  # one blank column after the dropdown lists
    c_stage, c_cnt, c_key, c_hint = (xl_col_to_name(h + i) for i in range(4))
    c_top, c_topcnt, c_focus = xl_col_to_name(h + 5), xl_col_to_name(h + 6), xl_col_to_name(h + 8)
    last = 1 + len(STAGES)
    rng_stage = f'${c_stage}$2:${c_stage}${last}'
    rng_cnt = f'${c_cnt}$2:${c_cnt}${last}'
    rng_key = f'${c_key}$2:${c_key}${last}'
    rng_hint = f'${c_hint}$2:${c_hint}${last}'

    ws_lookups.write(f'{c_stage}1', 'Stage', sub_hdr_fmt)
    ws_lookups.write(f'{c_cnt}1', 'Drop-offs', sub_hdr_fmt)
    ws_lookups.write(f'{c_key}1', 'Rank', sub_hdr_fmt)
    ws_lookups.write(f'{c_hint}1', 'Hint', sub_hdr_fmt)
    for i, stage in enumerate(STAGES):
        r = 2 + i
        ws_lookups.write(f'{c_stage}{r}', stage)
        bad = '+'.join(f'COUNTIFS({app_status},"{o}",{app_stage},"{stage}")' for o in BAD_OUTCOMES)
        ws_lookups.write_formula(f'{c_cnt}{r}', f'={bad}')
        ws_lookups.write_formula(f'{c_key}{r}', f'={c_cnt}{r}+({len(STAGES) + 1}-ROW())/1000')
        ws_lookups.write(f'{c_hint}{r}', STAGE_HINTS[stage])

    # Top 3 drop-off stages (labels + real counts) for the donut.
    ws_lookups.write(f'{c_top}1', 'Top stage', sub_hdr_fmt)
    ws_lookups.write(f'{c_topcnt}1', 'Drop-offs', sub_hdr_fmt)
    for k in range(1, 4):
        r = 1 + k
        ws_lookups.write_formula(f'{c_top}{r}', f'=INDEX({rng_stage},MATCH(LARGE({rng_key},{k}),{rng_key},0))')
        ws_lookups.write_formula(f'{c_topcnt}{r}', f'=INDEX({rng_cnt},MATCH(LARGE({rng_key},{k}),{rng_key},0))')

    # Focus hint: the hypothesis for the top stage, or NOT_ENOUGH_DATA below the threshold.
    ws_lookups.write(f'{c_focus}1', 'Focus hint', sub_hdr_fmt)
    ws_lookups.write_formula(
        f'{c_focus}2',
        f'=IF(SUM({rng_cnt})<{DROP_OFF_MIN},"{NOT_ENOUGH_DATA}",INDEX({rng_hint},MATCH(MAX({rng_key}),{rng_key},0)))')

    # --- Dashboard -----------------------------------------------------------
    total = f"COUNTA('Job Applications'!{company_col}:{company_col})-1"

    def in_progress_at(stages):
        return '+'.join(f'COUNTIFS({app_status},"In Progress",{app_stage},"{s}")' for s in stages)

    ws_dash.set_column(0, 4, 22)
    ws_dash.write('A1', 'Job Search Analytics', sub_hdr_fmt)

    # KPI cards: each label sits on one row with its value on the row directly below.
    kpis = [
        ('Total Applications', f'={total}', kpi_fmt),
        ('In Progress', f'=COUNTIF({app_status},"In Progress")', kpi_fmt),
        ('Interviews & Tests', f'={in_progress_at(INTERVIEW_STAGES)}', kpi_fmt),
        ('Offers Received', f'=COUNTIF({app_status},"Offer Received")+COUNTIF({app_status},"Accepted")', kpi_fmt),
        ('Response Rate',
         f'=IF({total}<=0,0,(({total})-(({in_progress_at(EARLY_STAGES)})+COUNTIF({app_status},"Ghosted")))/({total}))',
         pct_fmt),
    ]
    for col, (label, formula, fmt) in enumerate(kpis):
        ws_dash.write(2, col, label, sub_hdr_fmt)
        ws_dash.write_formula(3, col, formula, fmt)

    # Pipeline & Outcomes breakdown. In Progress rows are counted by Stage; finished
    # rows by Status. Each row is tinted with its own color only when its count > 0.
    breakdown = [(s, f'=COUNTIFS({app_status},"In Progress",{app_stage},"{s}")', STAGE_STYLES[s]) for s in STAGES]
    breakdown += [(o, f'=COUNTIF({app_status},"{o}")', OUTCOME_STYLES[o]) for o in STATUSES if o != 'In Progress']

    breakdown_header_row = 6
    first_data_row = breakdown_header_row + 1
    ws_dash.write(f'A{breakdown_header_row}', 'Pipeline & Outcomes', sub_hdr_fmt)
    ws_dash.write(f'B{breakdown_header_row}', 'Count', sub_hdr_fmt)
    for i, (label, formula, (bg, font)) in enumerate(breakdown):
        row = first_data_row + i
        ws_dash.write(f'A{row}', label, label_fmt)
        ws_dash.write_formula(f'B{row}', formula, count_fmt)
        tint = wb.add_format({'bg_color': bg, 'font_color': font, 'bold': True})
        ws_dash.conditional_format(f'A{row}:B{row}', {'type': 'formula', 'criteria': f'=$B${row}>0', 'format': tint})
    last_data_row = first_data_row + len(breakdown) - 1

    # Wide bar chart so the 14 category labels stay readable. Each bar keeps its
    # stage/outcome color via per-point fills.
    points = [{'fill': {'color': bg}} for (_, _, (bg, font)) in breakdown]
    chart = wb.add_chart({'type': 'column'})
    chart.add_series({
        'categories': f"='Dashboard'!$A${first_data_row}:$A${last_data_row}",
        'values': f"='Dashboard'!$B${first_data_row}:$B${last_data_row}",
        'points': points,
        'data_labels': {'value': True},
    })
    chart.set_title({'name': 'Pipeline & Outcomes'})
    chart.set_legend({'none': True})
    chart.set_x_axis({'num_font': {'rotation': -45}})
    chart.set_size({'width': 960, 'height': 340})
    ws_dash.insert_chart('D6', chart)

    # Next Actions summary: follow-ups that are overdue, due today, or coming up this week.
    na_range = f"'Job Applications'!{next_col}:{next_col}"
    na_header_row = last_data_row + 2
    ws_dash.write(f'A{na_header_row}', 'Next Actions', sub_hdr_fmt)
    ws_dash.write(f'B{na_header_row}', 'Count', sub_hdr_fmt)
    na_buckets = [
        ('Overdue', f'=COUNTIF({na_range},"<"&TODAY())', DATE_OVERDUE),
        ('Due today', f'=COUNTIF({na_range},TODAY())', DATE_TODAY),
        ('Next 7 days', f'=COUNTIFS({na_range},">"&TODAY(),{na_range},"<="&(TODAY()+7))', DATE_UPCOMING),
    ]
    for i, (label, formula, (bg, font)) in enumerate(na_buckets):
        r = na_header_row + 1 + i
        ws_dash.write(f'A{r}', label, label_fmt)
        ws_dash.write_formula(f'B{r}', formula, count_fmt)
        cf_fmt = wb.add_format({'bg_color': bg, 'font_color': font, 'bold': True})
        ws_dash.conditional_format(f'A{r}:B{r}', {'type': 'formula', 'criteria': f'=$B${r}>0', 'format': cf_fmt})

    # Drop-off donut: top 3 stages where applications are rejected or ghosted.
    donut = wb.add_chart({'type': 'doughnut'})
    donut.add_series({
        'name': 'Drop-offs by stage',
        'categories': f"='Lookup Lists'!${c_top}$2:${c_top}$4",
        'values': f"='Lookup Lists'!${c_topcnt}$2:${c_topcnt}$4",
        'data_labels': {'value': True},
    })
    donut.set_title({'name': 'Top Drop-off Stages'})
    donut.set_size({'width': 440, 'height': 300})
    ws_dash.insert_chart('D25', donut)

    # Focus text: a gentle hypothesis for the top drop-off stage (or a "keep tracking" note).
    ws_dash.write('J26', 'Where to focus', sub_hdr_fmt)
    ws_dash.merge_range('J27:Q34', '', hint_fmt)
    ws_dash.write_formula('J27', f"='Lookup Lists'!${c_focus}$2", hint_fmt)

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
