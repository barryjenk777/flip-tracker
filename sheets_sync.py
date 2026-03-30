"""
sheets_sync.py — Daily Google Sheets backup for Flip Tracker
Writes a fully-formatted, standalone spreadsheet with 5 tabs:
  0. Dashboard     — KPI summary, goal tracker, property flags
  1. Properties    — One row per property, all metrics
  2. Expenses      — All expenses across all properties (CPA-ready)
  3. P&L Summary   — Revenue / COGS / Selling / Net Profit per property
  4. Deal Pipeline — All deal analyzer prospects with metrics
"""

import os
import json
import base64
from datetime import datetime

# ---------------------------------------------------------------------------
# Color palette (RGB 0–1 floats for the gspread / Sheets API)
# ---------------------------------------------------------------------------
C = {
    'navy':        {"red": 0.11, "green": 0.16, "blue": 0.27},   # header bg
    'white':       {"red": 1.00, "green": 1.00, "blue": 1.00},
    'section_bg':  {"red": 0.15, "green": 0.22, "blue": 0.35},   # section divider
    'section_txt': {"red": 0.55, "green": 0.76, "blue": 0.97},   # light blue text
    'alt_row':     {"red": 0.95, "green": 0.96, "blue": 0.98},
    'white_row':   {"red": 1.00, "green": 1.00, "blue": 1.00},
    'sold_bg':     {"red": 0.82, "green": 0.93, "blue": 0.85},   # green tint
    'active_bg':   {"red": 1.00, "green": 0.97, "blue": 0.84},   # amber tint
    'reno_bg':     {"red": 0.84, "green": 0.90, "blue": 0.98},   # blue tint
    'pass_bg':     {"red": 0.82, "green": 0.93, "blue": 0.85},
    'fail_bg':     {"red": 0.98, "green": 0.86, "blue": 0.86},
    'border_bg':   {"red": 1.00, "green": 0.95, "blue": 0.82},
    'pos_txt':     {"red": 0.10, "green": 0.53, "blue": 0.25},   # dark green
    'neg_txt':     {"red": 0.72, "green": 0.11, "blue": 0.11},   # dark red
    'muted':       {"red": 0.45, "green": 0.55, "blue": 0.65},
    'dark_txt':    {"red": 0.10, "green": 0.14, "blue": 0.22},
    'flag_danger': {"red": 1.00, "green": 0.87, "blue": 0.87},
    'flag_warn':   {"red": 1.00, "green": 0.96, "blue": 0.83},
    'flag_good':   {"red": 0.88, "green": 0.96, "blue": 0.89},
}

STAGE_LABELS = {
    'new_lead': 'New Lead', 'analyzing': 'Analyzing',
    'offer_sent': 'Offer Sent', 'under_contract': 'Under Contract',
    'passed': 'Passed', 'converted': 'Converted',
}


# ---------------------------------------------------------------------------
# Auth & connection
# ---------------------------------------------------------------------------
def _get_client():
    import gspread
    from google.oauth2.service_account import Credentials
    raw = os.environ.get('GOOGLE_CREDENTIALS_JSON', '')
    if not raw:
        raise RuntimeError('GOOGLE_CREDENTIALS_JSON env var not set')
    try:
        creds_dict = json.loads(raw)
    except json.JSONDecodeError:
        creds_dict = json.loads(base64.b64decode(raw).decode())
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive.file',
    ]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)


def _get_or_create_tab(spreadsheet, title, rows=1000, cols=40):
    import gspread
    try:
        return spreadsheet.worksheet(title)
    except gspread.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------
def _col_letter(n):
    """Convert 0-based column index to A, B, ..., Z, AA, ..."""
    result = ''
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def _range(r1, c1, r2, c2):
    return f"{_col_letter(c1)}{r1}:{_col_letter(c2)}{r2}"


def _cell_fmt(bg=None, txt=None, bold=False, size=10, halign=None,
              number_format=None, italic=False, wrap=False):
    fmt = {}
    if bg:
        fmt['backgroundColor'] = bg
    tf = {}
    if txt:
        tf['foregroundColor'] = txt
    if bold:
        tf['bold'] = True
    if italic:
        tf['italic'] = True
    tf['fontSize'] = size
    fmt['textFormat'] = tf
    if halign:
        fmt['horizontalAlignment'] = halign
    if number_format:
        fmt['numberFormat'] = {'type': 'NUMBER', 'pattern': number_format}
    if wrap:
        fmt['wrapStrategy'] = 'WRAP'
    return fmt


def _apply_formats(ws, formats):
    """formats = list of (range_str, fmt_dict)"""
    if not formats:
        return
    ws.batch_format(formats)


def _set_col_widths(spreadsheet, ws, widths_px):
    """widths_px = list of pixel widths, one per column (0-indexed)."""
    sheet_id = ws._properties['sheetId']
    reqs = []
    for i, px in enumerate(widths_px):
        reqs.append({
            'updateDimensionProperties': {
                'range': {'sheetId': sheet_id, 'dimension': 'COLUMNS',
                          'startIndex': i, 'endIndex': i + 1},
                'properties': {'pixelSize': px},
                'fields': 'pixelSize'
            }
        })
    if reqs:
        spreadsheet.batch_update({'requests': reqs})


def _fmt_usd(v):
    if v is None or v == '':
        return ''
    try:
        return v if isinstance(v, (int, float)) else float(v)
    except (ValueError, TypeError):
        return ''


def _fmt_pct(v):
    try:
        return round(float(v), 1) if v is not None else ''
    except (ValueError, TypeError):
        return ''


def _fmt_date(v):
    if not v:
        return ''
    try:
        dt = datetime.strptime(str(v)[:10], '%Y-%m-%d')
        return dt.strftime('%b') + ' ' + str(dt.day) + ', ' + str(dt.year)
    except (ValueError, AttributeError):
        return str(v)


# ---------------------------------------------------------------------------
# Tab 0 — Dashboard
# ---------------------------------------------------------------------------
def _write_dashboard(spreadsheet, ws, enriched, prospects, biz_settings,
                     prospect_settings, sync_time):
    from app import calc_prospect_metrics
    ws.clear()

    annual_goal = biz_settings.get('annual_profit_goal', 0)
    total_props = len(enriched)
    statuses = [e['metrics']['status'] for e in enriched]
    n_active = sum(1 for s in statuses if 'active' in s.lower() or 'renovation' in s.lower() or s.lower() in ('acquired', 'renovation', 'listed', 'under contract'))
    n_sold = sum(1 for s in statuses if 'sold' in s.lower())
    n_listed = sum(1 for s in statuses if 'listed' in s.lower())

    sold_profit = sum(e['pnl']['net_profit'] for e in enriched if (e['prop'].get('sale_price') or 0) > 0)
    all_profit = sum(max(0, e['metrics']['gross_profit']) for e in enriched)
    total_invested = sum((e['prop'].get('purchase_price') or 0) + (e['metrics'].get('total_rehab') or 0) for e in enriched)
    avg_margin = sum(e['metrics'].get('profit_margin', 0) for e in enriched if (e['prop'].get('sale_price') or 0) > 0)
    closed_count = sum(1 for e in enriched if (e['prop'].get('sale_price') or 0) > 0)
    avg_margin = avg_margin / closed_count if closed_count else 0
    avg_roi = sum(e['metrics'].get('roi', 0) for e in enriched if (e['prop'].get('sale_price') or 0) > 0)
    avg_roi = avg_roi / closed_count if closed_count else 0

    # Pipeline deal metrics
    n_pass = sum(1 for p in prospects if (p.get('metrics') or {}).get('flip_verdict') == 'PASS')
    n_fail = sum(1 for p in prospects if (p.get('metrics') or {}).get('flip_verdict') == 'FAIL')
    n_border = sum(1 for p in prospects if (p.get('metrics') or {}).get('flip_verdict') == 'BORDERLINE')
    profits = [calc_prospect_metrics(p, prospect_settings).get('gross_profit', 0) for p in prospects]
    avg_prospect_profit = sum(profits) / len(profits) if profits else 0

    on_track = 'YES ✓' if annual_goal > 0 and sold_profit >= annual_goal * 0.5 else ('NO' if annual_goal > 0 else 'N/A')

    # Collect all flags
    all_flags = []
    for e in enriched:
        for flag in (e['metrics'].get('flags') or []):
            all_flags.append((e['prop'].get('address', ''), flag.get('type', ''), flag.get('msg', '')))

    rows = []

    # Title
    rows.append(['FLIP TRACKER — BUSINESS DASHBOARD', '', '', '', '', ''])
    rows.append([f'Last synced: {sync_time}', '', '', '', '', ''])
    rows.append(['', '', '', '', '', ''])

    # Section: Portfolio
    rows.append(['PORTFOLIO OVERVIEW', '', '', '', '', ''])
    rows.append(['Properties', 'Active/Reno', 'Listed', 'Sold', 'Total Invested', ''])
    rows.append([total_props, n_active, n_listed, n_sold, total_invested, ''])
    rows.append(['', '', '', '', '', ''])

    # Section: Profitability
    rows.append(['PROFITABILITY', '', '', '', '', ''])
    rows.append(['Total Projected Profit', 'Realized Profit', 'Avg Margin (Closed)', 'Avg ROI (Closed)', 'Closed Deals', ''])
    rows.append([all_profit, sold_profit, avg_margin, avg_roi, closed_count, ''])
    rows.append(['', '', '', '', '', ''])

    # Section: Annual Goal
    rows.append(['ANNUAL GOAL TRACKER', '', '', '', '', ''])
    rows.append(['Annual Profit Goal', 'Realized Profit', 'All Projected Profit', 'On Track?', '', ''])
    rows.append([annual_goal, sold_profit, all_profit, on_track, '', ''])
    rows.append(['', '', '', '', '', ''])

    # Section: Deal Pipeline
    rows.append(['DEAL PIPELINE', '', '', '', '', ''])
    rows.append(['Total Prospects', 'PASS', 'FAIL', 'BORDERLINE', 'Avg Gross Profit', ''])
    rows.append([len(prospects), n_pass, n_fail, n_border, avg_prospect_profit, ''])
    rows.append(['', '', '', '', '', ''])

    # Section: Property Flags
    rows.append(['ACTIVE FLAGS', '', '', '', '', ''])
    rows.append(['Property', 'Type', 'Message', '', '', ''])
    if all_flags:
        for addr, ftype, msg in all_flags:
            rows.append([addr, ftype.upper(), msg, '', '', ''])
    else:
        rows.append(['No active flags ✓', '', '', '', '', ''])

    ws.update('A1', rows)

    fmts = []
    # Title
    fmts.append(('A1:F1', _cell_fmt(bg=C['navy'], txt=C['white'], bold=True, size=14)))
    fmts.append(('A2:F2', _cell_fmt(bg=C['navy'], txt=C['muted'], italic=True, size=9)))

    # Section headers (rows 4, 8, 12, 16, 20 — 1-indexed)
    section_rows = [4, 8, 12, 16, 20]
    for r in section_rows:
        fmts.append((f'A{r}:F{r}', _cell_fmt(bg=C['section_bg'], txt=C['section_txt'], bold=True, size=10)))

    # Column label rows
    label_rows = [5, 9, 13, 17, 21]
    for r in label_rows:
        fmts.append((f'A{r}:F{r}', _cell_fmt(bg=C['alt_row'], bold=True, size=9)))

    # Value rows — currency formatting for known cols
    for r, cols_fmt in [(6, [(4, '$#,##0')]), (10, [(0, '$#,##0'), (1, '$#,##0'), (2, '0.0"%"'), (3, '0.0"%"')]),
                        (14, [(0, '$#,##0'), (1, '$#,##0'), (2, '$#,##0')]),
                        (18, [(0, '$#,##0'), (1, '$#,##0'), (2, '$#,##0'), (4, '$#,##0')])]:
        for col_i, nfmt in cols_fmt:
            cell = f'{_col_letter(col_i)}{r}'
            fmts.append((cell, _cell_fmt(number_format=nfmt, size=11, bold=True)))

    # Flags section
    flag_row_start = 22
    for i, (addr, ftype, msg) in enumerate(all_flags):
        r = flag_row_start + i
        color = C['flag_danger'] if ftype == 'danger' else C['flag_warn'] if ftype == 'warning' else C['flag_good']
        fmts.append((f'A{r}:F{r}', _cell_fmt(bg=color, size=9)))

    _apply_formats(ws, fmts)
    _set_col_widths(spreadsheet, ws, [200, 130, 130, 130, 160, 80])


# ---------------------------------------------------------------------------
# Tab 1 — Properties
# ---------------------------------------------------------------------------
def _write_properties(spreadsheet, ws, enriched):
    ws.clear()

    headers = [
        'Address', 'City', 'St', 'Status',
        'Purchase Date', 'Purchase Price', 'ARV', 'Sale Price',
        'Rehab Budget', 'Rehab Spent', 'Budget Variance %',
        'Holding Costs', 'Acq. Closing', 'Sale Commission', 'Sale Closing',
        'Total Costs', 'Gross Profit', 'Profit Margin %', 'ROI %', 'Ann. ROI %',
        'Cash In Deal', 'Partner Share', 'Days Held',
        'Passes 70% Rule', 'Est. Sale Date', 'Sale Date', 'Notes'
    ]

    rows = [headers]
    status_rows = []  # (row_index_1based, status)

    for i, e in enumerate(enriched):
        prop, m = e['prop'], e['metrics']
        sale_price = prop.get('sale_price') or 0
        rows.append([
            prop.get('address', ''),
            prop.get('city', ''),
            prop.get('state', ''),
            m.get('status', ''),
            _fmt_date(prop.get('purchase_date')),
            _fmt_usd(prop.get('purchase_price', 0)),
            _fmt_usd(prop.get('arv', 0)),
            _fmt_usd(sale_price) if sale_price else '',
            _fmt_usd(prop.get('rehab_budget', 0)),
            _fmt_usd(m.get('total_rehab', 0)),
            _fmt_pct(m.get('budget_variance', 0)),
            _fmt_usd(m.get('total_holding_cost', 0)),
            _fmt_usd(prop.get('acq_closing_cost', 0)),
            _fmt_usd(m.get('sale_commission', 0)),
            _fmt_usd(m.get('sale_closing', 0)),
            _fmt_usd(m.get('total_costs', 0)),
            _fmt_usd(m.get('gross_profit', 0)),
            _fmt_pct(m.get('profit_margin', 0)),
            _fmt_pct(m.get('roi', 0)),
            _fmt_pct(m.get('annualized_roi', 0)),
            _fmt_usd(m.get('cash_in_deal', 0)),
            _fmt_usd(m.get('partner_share', 0)),
            m.get('days_held', ''),
            'YES' if m.get('passes_70_rule') else 'NO' if m.get('passes_70_rule') is False else 'N/A',
            _fmt_date(prop.get('estimated_sale_date')),
            _fmt_date(prop.get('sale_date')),
            (prop.get('notes') or '')[:120],
        ])
        status_rows.append((i + 2, m.get('status', '')))

    ws.update('A1', rows)
    ws.freeze(rows=1)

    fmts = []
    # Header row
    fmts.append(('A1:AA1', _cell_fmt(bg=C['navy'], txt=C['white'], bold=True, size=9)))

    # Currency & percent columns (0-indexed col → format)
    COL_FMTS = {
        5: '$#,##0', 6: '$#,##0', 7: '$#,##0', 8: '$#,##0', 9: '$#,##0',
        10: '0.0"%"', 11: '$#,##0', 12: '$#,##0', 13: '$#,##0', 14: '$#,##0',
        15: '$#,##0', 16: '$#,##0', 17: '0.0"%"', 18: '0.0"%"', 19: '0.0"%"',
        20: '$#,##0', 21: '$#,##0', 22: '#,##0',
    }
    n_data = len(rows) - 1
    if n_data > 0:
        for col_i, nfmt in COL_FMTS.items():
            col_l = _col_letter(col_i)
            fmts.append((f'{col_l}2:{col_l}{n_data + 1}', _cell_fmt(number_format=nfmt, size=9)))

        # Row colors by status
        for row_1, status in status_rows:
            s = status.lower()
            bg = C['sold_bg'] if 'sold' in s else C['reno_bg'] if any(x in s for x in ('renovation', 'acquired', 'active')) else C['active_bg']
            fmts.append((f'A{row_1}:AA{row_1}', _cell_fmt(bg=bg, size=9)))

        # Profit column — color positive/negative
        for i, e in enumerate(enriched):
            profit = e['metrics'].get('gross_profit', 0) or 0
            row_1 = i + 2
            txt_color = C['pos_txt'] if profit >= 0 else C['neg_txt']
            fmts.append((f'Q{row_1}', _cell_fmt(txt=txt_color, bold=True, size=9)))

    _apply_formats(ws, fmts)
    _set_col_widths(spreadsheet, ws,
        [190, 120, 40, 90, 100, 110, 110, 100, 110, 110, 100,
         110, 110, 110, 110, 110, 110, 90, 80, 90, 110, 110,
         80, 90, 100, 100, 180])


# ---------------------------------------------------------------------------
# Tab 2 — Expenses
# ---------------------------------------------------------------------------
def _write_expenses(spreadsheet, ws, enriched, expense_tax_map):
    ws.clear()

    headers = [
        'Property', 'City', 'Date', 'Category', 'Vendor',
        'Description', 'Amount', 'Credit?', 'Tax Classification', 'Tax Type'
    ]

    expense_rows = []
    for e in enriched:
        prop = e['prop']
        addr = prop.get('address', '')
        city = prop.get('city', '')
        for exp in (prop.get('expenses') or []):
            cat = exp.get('category', 'Other')
            tax_class, tax_type = expense_tax_map.get(cat, ('Renovation - Other', 'cogs'))
            expense_rows.append([
                addr, city,
                _fmt_date(exp.get('date')),
                cat,
                exp.get('vendor', ''),
                exp.get('description', ''),
                _fmt_usd(exp.get('amount', 0)),
                'YES' if exp.get('is_credit') else 'NO',
                tax_class,
                'COGS' if tax_type == 'cogs' else 'Selling',
            ])

    # Sort by property then date
    expense_rows.sort(key=lambda r: (r[0], r[2]))

    # Property subtotals
    all_rows = [headers]
    last_prop = None
    prop_total = 0
    for r in expense_rows:
        if last_prop and r[0] != last_prop:
            all_rows.append([f'  Subtotal: {last_prop}', '', '', '', '', '', prop_total, '', '', ''])
            all_rows.append(['', '', '', '', '', '', '', '', '', ''])
            prop_total = 0
        last_prop = r[0]
        is_credit = r[7] == 'YES'
        prop_total += (-r[6] if is_credit else r[6]) if isinstance(r[6], (int, float)) else 0
        all_rows.append(r)
    if last_prop:
        all_rows.append([f'  Subtotal: {last_prop}', '', '', '', '', '', prop_total, '', '', ''])

    # Grand total
    grand = sum(
        (-r[6] if r[7] == 'YES' else r[6])
        for r in expense_rows
        if isinstance(r[6], (int, float))
    )
    all_rows.append(['', '', '', '', '', '', '', '', '', ''])
    all_rows.append(['GRAND TOTAL', '', '', '', '', '', grand, '', '', ''])

    ws.update('A1', all_rows)
    ws.freeze(rows=1)

    fmts = []
    fmts.append(('A1:J1', _cell_fmt(bg=C['navy'], txt=C['white'], bold=True, size=9)))

    n_total = len(all_rows)
    fmts.append((f'G2:G{n_total}', _cell_fmt(number_format='$#,##0.00', size=9)))

    # Alternate row colors + subtotal/total formatting
    data_row = 2
    last_p = None
    for r in all_rows[1:]:
        addr = r[0]
        if addr.startswith('  Subtotal:') or addr == 'GRAND TOTAL':
            fmts.append((f'A{data_row}:J{data_row}', _cell_fmt(bg=C['section_bg'], txt=C['white'], bold=True, size=9)))
        elif addr == '':
            pass  # blank separator
        else:
            bg = C['alt_row'] if last_p and addr == last_p and data_row % 2 == 0 else C['white_row']
            last_p = addr
            fmts.append((f'A{data_row}:J{data_row}', _cell_fmt(bg=bg, size=9)))
        data_row += 1

    _apply_formats(ws, fmts)
    _set_col_widths(spreadsheet, ws, [190, 110, 100, 160, 140, 200, 110, 60, 170, 80])


# ---------------------------------------------------------------------------
# Tab 3 — P&L Summary
# ---------------------------------------------------------------------------
def _write_pnl(spreadsheet, ws, enriched):
    ws.clear()

    headers = [
        'Property', 'City / State', 'Status', 'Purchase Date', 'Sale Date', 'Days Held',
        '— REVENUE —', 'Gross Sale Price', 'Seller Concessions', 'Net Sale Proceeds',
        '— COGS —', 'Purchase Price', 'Acq. Closing Costs',
        'Reno - Labor', 'Reno - Materials', 'Reno - Permits/Fees', 'Reno - Other',
        'Holding Costs', 'Total COGS',
        '— SELLING —', 'Commission', 'Sale Closing Costs', 'Marketing / Staging',
        'Total Selling Costs',
        '— NET —', 'Net Profit', 'Profit Margin %',
        'Partner A Share', 'Partner B Share',
    ]

    rows = [headers]
    profit_rows = []  # (row_1based, net_profit)

    for i, e in enumerate(enriched):
        prop, m, pnl = e['prop'], e['metrics'], e['pnl']
        sale_price = prop.get('sale_price') or 0

        # Renovation subtotals from P&L
        reno_labor = pnl.get('renovation_subtotals', {}).get('Renovation - Labor', 0)
        reno_mats = pnl.get('renovation_subtotals', {}).get('Renovation - Materials', 0)
        reno_permits = pnl.get('renovation_subtotals', {}).get('Renovation - Permits/Fees', 0)
        reno_other = pnl.get('renovation_subtotals', {}).get('Renovation - Other', 0)

        net_profit = pnl.get('net_profit', 0)
        split = prop.get('partner_split_pct', 50) / 100
        partner_a = net_profit * split
        partner_b = net_profit * (1 - split)

        rows.append([
            prop.get('address', ''),
            f"{prop.get('city', '')}, {prop.get('state', '')}",
            m.get('status', ''),
            _fmt_date(prop.get('purchase_date')),
            _fmt_date(prop.get('sale_date')),
            m.get('days_held', ''),
            '',  # section label col
            _fmt_usd(sale_price) if sale_price else _fmt_usd(prop.get('arv', 0)),
            _fmt_usd(pnl.get('seller_concessions', 0)),
            _fmt_usd(pnl.get('net_sale_proceeds', 0)),
            '',  # section label col
            _fmt_usd(pnl.get('purchase_price', 0)),
            _fmt_usd(pnl.get('acq_closing_total', 0)),
            _fmt_usd(reno_labor),
            _fmt_usd(reno_mats),
            _fmt_usd(reno_permits),
            _fmt_usd(reno_other),
            _fmt_usd(pnl.get('holding_total', pnl.get('total_holding_cost', 0))),
            _fmt_usd(pnl.get('total_cogs', 0)),
            '',  # section label col
            _fmt_usd(pnl.get('commission', 0)),
            _fmt_usd(pnl.get('sale_closing', 0)),
            _fmt_usd(pnl.get('selling_expense_total', 0)),
            _fmt_usd(pnl.get('total_selling', 0)),
            '',  # section label col
            _fmt_usd(net_profit),
            _fmt_pct(m.get('profit_margin', 0)),
            _fmt_usd(partner_a),
            _fmt_usd(partner_b),
        ])
        profit_rows.append((i + 2, net_profit))

    # Totals row
    def col_sum(col_i):
        vals = []
        for r in rows[1:]:
            v = r[col_i] if col_i < len(r) else ''
            if isinstance(v, (int, float)):
                vals.append(v)
        return sum(vals)

    totals = ['TOTALS', '', '', '', '', '']
    for ci in range(6, 29):
        totals.append(col_sum(ci))
    rows.append(totals)

    ws.update('A1', rows)
    ws.freeze(rows=1)

    fmts = []
    fmts.append(('A1:AC1', _cell_fmt(bg=C['navy'], txt=C['white'], bold=True, size=9)))

    # Section divider columns in header
    for col_ltr in ['G', 'K', 'T', 'Y']:
        fmts.append((f'{col_ltr}1', _cell_fmt(bg=C['section_bg'], txt=C['section_txt'], bold=True, size=8)))

    # Number formats for data rows
    n_data = len(rows) - 1
    if n_data > 0:
        USD_COLS = [7, 8, 9, 11, 12, 13, 14, 15, 16, 17, 18, 20, 21, 22, 23, 25, 27, 28]
        for ci in USD_COLS:
            cl = _col_letter(ci)
            fmts.append((f'{cl}2:{cl}{n_data + 1}', _cell_fmt(number_format='$#,##0', size=9)))
        fmts.append((f'R2:R{n_data + 1}', _cell_fmt(number_format='#,##0', size=9)))
        fmts.append((f'AA2:AA{n_data + 1}', _cell_fmt(number_format='0.0"%"', size=9)))

        # Totals row
        total_row = n_data + 1
        fmts.append((f'A{total_row}:AC{total_row}', _cell_fmt(bg=C['section_bg'], txt=C['white'], bold=True, size=9)))

        # Profit column color
        for row_1, profit in profit_rows:
            txt = C['pos_txt'] if profit >= 0 else C['neg_txt']
            fmts.append((f'Z{row_1}', _cell_fmt(txt=txt, bold=True, size=9)))

        # Alternate row backgrounds
        for i in range(len(enriched)):
            row_1 = i + 2
            bg = C['alt_row'] if i % 2 == 0 else C['white_row']
            fmts.append((f'A{row_1}:AC{row_1}', _cell_fmt(bg=bg, size=9)))

    _apply_formats(ws, fmts)
    widths = [190, 130, 90, 100, 100, 70, 30, 100, 110, 110,
              30, 110, 110, 100, 110, 110, 100, 110, 110,
              30, 110, 110, 120, 110, 30, 110, 90, 110, 110]
    _set_col_widths(spreadsheet, ws, widths)


# ---------------------------------------------------------------------------
# Tab 4 — Deal Pipeline
# ---------------------------------------------------------------------------
def _write_pipeline(spreadsheet, ws, prospects, prospect_settings):
    from app import calc_prospect_metrics
    ws.clear()

    headers = [
        'Address', 'City', 'State', 'Stage', 'Source', 'Date Added',
        'Asking Price', 'ARV', 'Est. Rehab', 'Acq. Closing', 'Hard Money/Mo',
        'Hold Months', 'Gross Profit', 'ROI %', 'Profit Margin %',
        'Total Cost / ARV %', 'MAO (70% Rule)', 'Flip Verdict', 'Rental Verdict', 'Notes'
    ]

    rows = [headers]
    verdict_rows = []

    for i, p in enumerate(prospects):
        m = calc_prospect_metrics(p, prospect_settings)
        rows.append([
            p.get('address', ''),
            p.get('city', ''),
            p.get('state', 'VA'),
            STAGE_LABELS.get(p.get('stage', 'new_lead'), p.get('stage', '')),
            p.get('source', ''),
            _fmt_date(p.get('date_added')),
            _fmt_usd(p.get('asking_price', 0)),
            _fmt_usd(p.get('arv', 0)),
            _fmt_usd(p.get('estimated_rehab', 0)),
            _fmt_usd(p.get('acq_closing_costs', 0)),
            _fmt_usd(p.get('monthly_hard_money', 0)),
            p.get('holding_months', ''),
            _fmt_usd(m.get('gross_profit', 0)),
            _fmt_pct(m.get('roi', 0)),
            _fmt_pct(m.get('profit_margin', 0)),
            _fmt_pct(m.get('total_cost_to_arv', 0)),
            _fmt_usd(m.get('mao', 0)),
            m.get('flip_verdict', ''),
            m.get('rental_verdict', ''),
            (p.get('notes') or '')[:100],
        ])
        verdict_rows.append((i + 2, m.get('flip_verdict', 'FAIL')))

    ws.update('A1', rows)
    ws.freeze(rows=1)

    fmts = []
    fmts.append(('A1:T1', _cell_fmt(bg=C['navy'], txt=C['white'], bold=True, size=9)))

    n_data = len(rows) - 1
    if n_data > 0:
        USD_COLS = [6, 7, 8, 9, 10, 12, 16]
        for ci in USD_COLS:
            cl = _col_letter(ci)
            fmts.append((f'{cl}2:{cl}{n_data + 1}', _cell_fmt(number_format='$#,##0', size=9)))
        for ci in [13, 14, 15]:
            cl = _col_letter(ci)
            fmts.append((f'{cl}2:{cl}{n_data + 1}', _cell_fmt(number_format='0.0"%"', size=9)))

        for row_1, verdict in verdict_rows:
            bg = C['pass_bg'] if verdict == 'PASS' else C['fail_bg'] if verdict == 'FAIL' else C['border_bg']
            fmts.append((f'A{row_1}:T{row_1}', _cell_fmt(bg=bg, size=9)))

    _apply_formats(ws, fmts)
    _set_col_widths(spreadsheet, ws,
        [190, 110, 50, 110, 100, 100, 110, 110, 110, 100, 110, 80,
         110, 80, 90, 100, 110, 90, 90, 180])


# ---------------------------------------------------------------------------
# Main sync entry point
# ---------------------------------------------------------------------------
def sync_to_sheets():
    """Called by APScheduler daily and by the manual /api/sync-sheets route."""
    try:
        from app import load_data, calc_property_metrics, calc_pnl, EXPENSE_TAX_MAP, calc_prospect_metrics
        data = load_data()

        sheet_id = os.environ.get('GOOGLE_SHEET_ID', '')
        if not sheet_id:
            print('[sheets_sync] GOOGLE_SHEET_ID not set — skipping sync')
            return

        client = _get_client()
        spreadsheet = client.open_by_key(sheet_id)

        props = data.get('properties', [])
        prospects_raw = data.get('prospects', [])
        prospect_settings = data.get('prospect_settings', {})
        biz_settings = data.get('business_settings', {})

        # Enrich all properties
        enriched = []
        for prop in props:
            try:
                m = calc_property_metrics(prop)
                pnl = calc_pnl(prop, m)
                enriched.append({'prop': prop, 'metrics': m, 'pnl': pnl})
            except Exception as e:
                print(f'[sheets_sync] Could not compute metrics for {prop.get("address")}: {e}')

        # Enrich prospects
        prospects = []
        for p in prospects_raw:
            try:
                m = calc_prospect_metrics(p, prospect_settings)
                prospects.append({**p, 'metrics': m})
            except Exception:
                prospects.append(p)

        now = datetime.now()
        sync_time = now.strftime('%B') + ' ' + str(now.day) + ', ' + str(now.year) + ' at ' + now.strftime('%I:%M %p').lstrip('0')

        ws_dashboard = _get_or_create_tab(spreadsheet, 'Dashboard')
        ws_properties = _get_or_create_tab(spreadsheet, 'Properties')
        ws_expenses   = _get_or_create_tab(spreadsheet, 'Expenses')
        ws_pnl        = _get_or_create_tab(spreadsheet, 'P&L Summary')
        ws_pipeline   = _get_or_create_tab(spreadsheet, 'Deal Pipeline')

        # Each tab is independent — one failure won't block the rest
        tab_results = {}

        for tab_name, writer in [
            ('Dashboard',    lambda: _write_dashboard(spreadsheet, ws_dashboard, enriched, prospects, biz_settings, prospect_settings, sync_time)),
            ('Properties',   lambda: _write_properties(spreadsheet, ws_properties, enriched)),
            ('Expenses',     lambda: _write_expenses(spreadsheet, ws_expenses, enriched, EXPENSE_TAX_MAP)),
            ('P&L Summary',  lambda: _write_pnl(spreadsheet, ws_pnl, enriched)),
            ('Deal Pipeline',lambda: _write_pipeline(spreadsheet, ws_pipeline, prospects, prospect_settings)),
        ]:
            try:
                writer()
                tab_results[tab_name] = 'ok'
                print(f'[sheets_sync] ✓ {tab_name}')
            except Exception as tab_err:
                tab_results[tab_name] = str(tab_err)
                print(f'[sheets_sync] ✗ {tab_name}: {tab_err}')

        # Enforce tab order
        try:
            spreadsheet.reorder_worksheets([ws_dashboard, ws_properties, ws_expenses, ws_pnl, ws_pipeline])
        except Exception:
            pass

        all_ok = all(v == 'ok' for v in tab_results.values())
        print(f'[sheets_sync] Sync {"complete" if all_ok else "partial"} — {len(enriched)} properties, {len(prospects)} prospects — {sync_time}')
        return {'ok': all_ok, 'synced_at': sync_time, 'tabs': tab_results,
                'properties': len(enriched), 'prospects': len(prospects)}

    except Exception as e:
        print(f'[sheets_sync] ERROR: {e}')
        return {'ok': False, 'error': str(e)}
