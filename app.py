#!/usr/bin/env python3
"""
Flip Tracker - Professional Real Estate Flip Investment Dashboard
Standalone Flask application for tracking renovation flip investments.
"""

from flask import Flask, render_template, request, jsonify, Response, send_file
import json
import os
import base64
import io
import re
import csv
from datetime import datetime

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB upload limit

# ---------------------------------------------------------------------------
# Tax categorization for IRS reporting (flips = dealer property = inventory)
# All costs capitalized to basis under IRC Section 263A
# ---------------------------------------------------------------------------
EXPENSE_TAX_MAP = {
    'Labor - Plumbing': ('Renovation - Labor', 'cogs'),
    'Labor - Electrical': ('Renovation - Labor', 'cogs'),
    'Labor - HVAC': ('Renovation - Labor', 'cogs'),
    'Labor - Kitchen': ('Renovation - Labor', 'cogs'),
    'Labor - General': ('Renovation - Labor', 'cogs'),
    'Building Materials': ('Renovation - Materials', 'cogs'),
    'Flooring': ('Renovation - Materials', 'cogs'),
    'Paint': ('Renovation - Materials', 'cogs'),
    'Roofing': ('Renovation - Materials', 'cogs'),
    'Windows & Doors': ('Renovation - Materials', 'cogs'),
    'Appliances': ('Renovation - Materials', 'cogs'),
    'Landscaping': ('Renovation - Other', 'cogs'),
    'Permits': ('Renovation - Permits/Fees', 'cogs'),
    'Dumpster': ('Renovation - Other', 'cogs'),
    'Repairs - Pest': ('Renovation - Other', 'cogs'),
    'Repairs - Foundation': ('Renovation - Other', 'cogs'),
    'Utilities': ('Holding Costs', 'cogs'),
    'Marketing': ('Selling Costs', 'selling'),
    'Staging': ('Selling Costs', 'selling'),
    'Other': ('Renovation - Other', 'cogs'),
}

# Closing Disclosure line item tax classification keywords
CD_TAX_KEYWORDS = {
    'originaon': 'Loan Costs - Capitalized',  # handles stripped ligature "ti"
    'origination': 'Loan Costs - Capitalized',
    'discount': 'Loan Costs - Capitalized',
    'processing': 'Loan Costs - Capitalized',
    'underwriting': 'Loan Costs - Capitalized',
    'application': 'Loan Costs - Capitalized',
    'appraisal': 'Loan Costs - Capitalized',
    'credit report': 'Loan Costs - Capitalized',
    'flood': 'Loan Costs - Capitalized',
    'title': 'Title & Settlement - Capitalized',
    'settlement': 'Title & Settlement - Capitalized',
    'closing fee': 'Title & Settlement - Capitalized',
    'survey': 'Title & Settlement - Capitalized',
    'pest': 'Inspection - Capitalized',
    'inspection': 'Inspection - Capitalized',
    'recording': 'Gov Fees - Capitalized',
    'transfer tax': 'Gov Fees - Capitalized',
    'prepaid interest': 'Prepaid Interest - Capitalized',
    'homeowner': 'Insurance - Capitalized',
    'hazard': 'Insurance - Capitalized',
    'insurance': 'Insurance - Capitalized',
    'property tax': 'Property Tax - Capitalized',
    'tax': 'Property Tax - Capitalized',
    'escrow': 'Escrow - Capitalized',
    'hoa': 'HOA - Capitalized',
    'water': 'Property Tax - Capitalized',
    'sewer': 'Property Tax - Capitalized',
    'admin fee': 'Title & Settlement - Capitalized',
    'document prep': 'Title & Settlement - Capitalized',
    'notary': 'Title & Settlement - Capitalized',
    'wire': 'Title & Settlement - Capitalized',
    'courier': 'Title & Settlement - Capitalized',
    'interest': 'Prepaid Interest - Capitalized',
    'daily interest': 'Prepaid Interest - Capitalized',
    'selement': 'Title & Settlement - Capitalized',  # handles stripped ligature "ttl"
    'preparaon': 'Title & Settlement - Capitalized',  # handles stripped ligature "ti"
}

# ---------------------------------------------------------------------------
# Data persistence (JSON file, with in-memory fallback for Railway)
# ---------------------------------------------------------------------------
DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'flip_data.json')
_memory_store = None


def _default_prospect_settings():
    return {
        'min_profit': 25000,
        'min_roi': 15,
        'arv_multiplier': 0.70,
        'commission_pct': 6.0,
        'closing_cost_pct': 3.0,
        'monthly_holding_cost': 2500,
        'holding_months': 6,
        'rental_expense_ratio': 0.50,
        'min_cashflow_per_door': 200,
        'min_cap_rate': 5.0,
        'min_cash_on_cash': 8.0,
        'down_payment_pct': 20,
        'interest_rate': 7.5,
        'loan_term_years': 30,
    }


def _default_data():
    return {
        'properties': [],
        'prospects': [],
        'prospect_settings': _default_prospect_settings(),
        'settings': {
            'default_commission_pct': 4.0,
            'default_closing_cost_pct': 1.5,
            'default_contingency_pct': 15.0,
            'partner_split_pct': 50.0,
        }
    }


def load_data():
    global _memory_store
    if _memory_store is not None:
        return _memory_store
    try:
        with open(DATA_FILE, 'r') as f:
            _memory_store = json.load(f)
            _memory_store.setdefault('prospects', [])
            _memory_store.setdefault('prospect_settings', _default_prospect_settings())
            return _memory_store
    except (FileNotFoundError, json.JSONDecodeError):
        _memory_store = _default_data()
        return _memory_store


def save_data(data):
    global _memory_store
    _memory_store = data
    try:
        with open(DATA_FILE, 'w') as f:
            json.dump(data, f, indent=2)
    except OSError:
        pass  # Railway read-only FS — memory store is still updated


# ---------------------------------------------------------------------------
# Calculation engine
# ---------------------------------------------------------------------------
def calc_property_metrics(prop):
    """Calculate all derived metrics for a property."""
    purchase_price = prop.get('purchase_price', 0) or 0
    arv = prop.get('arv', 0) or 0
    sale_price = prop.get('sale_price', 0) or 0
    acq_closing_cost = prop.get('acq_closing_cost', 0) or 0
    sale_commission_pct = prop.get('sale_commission_pct', 4.0) or 4.0
    sale_closing_cost_pct = prop.get('sale_closing_cost_pct', 1.5) or 1.5
    contingency_pct = prop.get('contingency_pct', 15.0) or 15.0
    partner_split_pct = prop.get('partner_split_pct', 50.0) or 50.0
    sqft = prop.get('sqft', 0) or 0

    # Dates
    purchase_date = prop.get('purchase_date')
    sale_date = prop.get('sale_date')
    listing_date = prop.get('listing_date')

    # Expenses
    expenses = prop.get('expenses', [])
    draws = prop.get('draws', [])
    mortgage_payments = prop.get('mortgage_payments', [])
    holding_costs = prop.get('holding_costs', {
        'monthly_mortgage': 0, 'monthly_insurance': 0, 'monthly_taxes': 0,
        'monthly_utilities': 0, 'monthly_hoa': 0, 'monthly_lawn': 0, 'monthly_other': 0,
    })

    # ---- Rehab costs ----
    total_rehab = sum(e.get('amount', 0) for e in expenses if not e.get('is_credit'))
    total_credits = sum(e.get('amount', 0) for e in expenses if e.get('is_credit'))
    total_rehab -= total_credits

    rehab_by_category = {}
    for e in expenses:
        if e.get('is_credit'):
            continue
        cat = e.get('category', 'Other')
        rehab_by_category[cat] = rehab_by_category.get(cat, 0) + e.get('amount', 0)

    # Budget tracking — actual budget vs lender budget
    budget = prop.get('rehab_budget', 0) or 0  # what you expect to actually spend
    lender_budget = prop.get('lender_rehab_budget', 0) or 0  # what the lender approved for draws
    if lender_budget == 0:
        lender_budget = budget  # fallback: if no lender budget set, treat same as actual
    budget_variance = ((total_rehab - budget) / budget * 100) if budget > 0 else 0
    budget_remaining = budget - total_rehab
    contingency_amount = budget * (contingency_pct / 100) if budget > 0 else total_rehab * (contingency_pct / 100)

    # ---- Draws ----
    total_draws = sum(d.get('amount', 0) for d in draws)
    draw_credit = total_draws - total_rehab

    # Capital recapture: difference between lender budget and actual spend
    lender_budget_spread = lender_budget - budget  # planned capital recapture
    actual_capital_recapture = lender_budget - total_rehab  # actual capital recapture so far
    lender_budget_remaining = lender_budget - total_draws  # how much more draw capacity left
    draw_utilization = (total_draws / lender_budget * 100) if lender_budget > 0 else 0

    # ---- Holding costs ----
    total_mortgage_payments = sum(m.get('amount', 0) for m in mortgage_payments)
    monthly_hold = sum([
        holding_costs.get('monthly_mortgage', 0),
        holding_costs.get('monthly_insurance', 0),
        holding_costs.get('monthly_taxes', 0),
        holding_costs.get('monthly_utilities', 0),
        holding_costs.get('monthly_hoa', 0),
        holding_costs.get('monthly_lawn', 0),
        holding_costs.get('monthly_other', 0),
    ])
    daily_hold = monthly_hold / 30 if monthly_hold > 0 else 0

    # Days held
    if purchase_date:
        pd = datetime.strptime(purchase_date, '%Y-%m-%d')
        end = datetime.strptime(sale_date, '%Y-%m-%d') if sale_date else datetime.now()
        days_held = (end - pd).days
        months_held = days_held / 30
    else:
        days_held = 0
        months_held = 0

    total_holding_cost = total_mortgage_payments + (monthly_hold - holding_costs.get('monthly_mortgage', 0)) * months_held
    if total_holding_cost == 0 and monthly_hold > 0:
        total_holding_cost = monthly_hold * months_held

    # ---- Sale costs ----
    effective_sale = sale_price if sale_price > 0 else arv
    sale_commission = effective_sale * (sale_commission_pct / 100)
    sale_closing = effective_sale * (sale_closing_cost_pct / 100)

    # ---- Total investment ----
    purchase_settlement = prop.get('purchase_settlement', 0) or 0
    emd = prop.get('emd', 0) or 0
    appraisal_fee = prop.get('appraisal_fee', 0) or 0
    commitment_fee = prop.get('commitment_fee', 0) or 0
    if purchase_settlement > 0:
        total_cash_oop = purchase_settlement + emd + commitment_fee + appraisal_fee + total_rehab + total_holding_cost
    else:
        total_cash_oop = acq_closing_cost + total_rehab + total_holding_cost
    # Draws reimburse rehab — any surplus reduces cash in deal
    draw_surplus = max(total_draws - total_rehab, 0)
    # Also subtract draws that covered rehab (not just surplus)
    draws_applied = min(total_draws, total_rehab)
    cash_in_deal = total_cash_oop - draws_applied - draw_surplus

    # ---- Profit ----
    total_costs = purchase_price + acq_closing_cost + total_rehab + sale_commission + sale_closing + total_holding_cost
    gross_profit = effective_sale - total_costs
    profit_margin = (gross_profit / effective_sale * 100) if effective_sale > 0 else 0
    partner_share = gross_profit * (partner_split_pct / 100)

    # ---- ROI ----
    roi = (gross_profit / cash_in_deal * 100) if cash_in_deal > 0 else 0
    annualized_roi = (roi / (months_held / 12)) if months_held > 0 else 0
    cash_on_cash = (gross_profit / cash_in_deal * 100) if cash_in_deal > 0 else 0

    # ---- 70% Rule ----
    mao = (arv * 0.70) - total_rehab if arv > 0 else 0
    mao_with_holding = (arv * 0.70) - total_rehab - total_holding_cost if arv > 0 else 0
    passes_70_rule = purchase_price <= mao if mao > 0 else None
    total_cost_to_arv = (total_costs / arv * 100) if arv > 0 else 0

    # ---- Cost per sqft ----
    rehab_per_sqft = (total_rehab / sqft) if sqft > 0 else 0
    total_cost_per_sqft = (total_costs / sqft) if sqft > 0 else 0

    # ---- Days on market ----
    dom = 0
    if listing_date:
        ld = datetime.strptime(listing_date, '%Y-%m-%d')
        end_d = datetime.strptime(sale_date, '%Y-%m-%d') if sale_date else datetime.now()
        dom = (end_d - ld).days

    # ---- Holding cost burn ----
    profit_erosion_per_day = daily_hold
    days_until_zero_profit = int(gross_profit / daily_hold) if daily_hold > 0 else 999

    # ---- Status ----
    status = prop.get('status', 'active')
    if sale_date:
        status = 'sold'
    elif listing_date:
        status = 'listed'

    # ---- Risk flags ----
    flags = []
    if budget > 0 and budget_variance > 10:
        flags.append({'type': 'danger', 'msg': f'Budget overrun: {budget_variance:+.1f}%'})
    elif budget > 0 and budget_variance > 5:
        flags.append({'type': 'warning', 'msg': f'Budget variance: {budget_variance:+.1f}%'})
    if roi > 0 and roi < 15:
        flags.append({'type': 'warning', 'msg': f'ROI below 15%: {roi:.1f}%'})
    if roi > 0 and roi < 10:
        flags.append({'type': 'danger', 'msg': f'ROI critically low: {roi:.1f}%'})
    if gross_profit < 15000 and effective_sale > 0:
        flags.append({'type': 'danger', 'msg': 'Profit below $15K minimum floor'})
    if total_cost_to_arv > 85 and arv > 0:
        flags.append({'type': 'danger', 'msg': f'Total cost at {total_cost_to_arv:.0f}% of ARV (>85%)'})
    if passes_70_rule is False and gross_profit < 30000:
        flags.append({'type': 'warning', 'msg': f'Purchase exceeds 70% rule MAO by ${purchase_price - mao:,.0f} — thin margins'})
    elif passes_70_rule is False and gross_profit >= 30000:
        flags.append({'type': 'good', 'msg': f'Bought ${purchase_price - mao:,.0f} over 70% MAO but still profitable ({profit_margin:.0f}% margin)'})
    if contingency_pct < 10:
        flags.append({'type': 'warning', 'msg': 'Contingency below 10% — risky'})
    if dom > 60:
        flags.append({'type': 'warning', 'msg': f'{dom} days on market (>60)'})
    if days_held > 180:
        flags.append({'type': 'warning', 'msg': f'{days_held} days held (>180 benchmark)'})

    return {
        'purchase_price': purchase_price, 'arv': arv, 'sale_price': sale_price,
        'effective_sale': effective_sale, 'sqft': sqft, 'status': status,
        'total_rehab': total_rehab, 'net_rehab': total_rehab,
        'rehab_by_category': rehab_by_category, 'budget': budget,
        'lender_budget': lender_budget,
        'lender_budget_spread': lender_budget_spread,
        'actual_capital_recapture': actual_capital_recapture,
        'lender_budget_remaining': lender_budget_remaining,
        'draw_utilization': draw_utilization,
        'budget_variance': budget_variance, 'budget_remaining': budget_remaining,
        'contingency_pct': contingency_pct, 'contingency_amount': contingency_amount,
        'total_draws': total_draws, 'draw_credit': draw_credit,
        'total_mortgage_payments': total_mortgage_payments, 'monthly_hold': monthly_hold,
        'daily_hold': daily_hold, 'total_holding_cost': total_holding_cost,
        'acq_closing_cost': acq_closing_cost, 'total_cash_oop': total_cash_oop,
        'cash_in_deal': cash_in_deal,
        'sale_commission': sale_commission, 'sale_commission_pct': sale_commission_pct,
        'sale_closing': sale_closing, 'sale_closing_cost_pct': sale_closing_cost_pct,
        'total_costs': total_costs, 'gross_profit': gross_profit,
        'profit_margin': profit_margin, 'partner_split_pct': partner_split_pct,
        'partner_share': partner_share,
        'roi': roi, 'annualized_roi': annualized_roi, 'cash_on_cash': cash_on_cash,
        'mao': mao, 'mao_with_holding': mao_with_holding,
        'passes_70_rule': passes_70_rule, 'total_cost_to_arv': total_cost_to_arv,
        'rehab_per_sqft': rehab_per_sqft, 'total_cost_per_sqft': total_cost_per_sqft,
        'days_held': days_held, 'months_held': months_held, 'dom': dom,
        'profit_erosion_per_day': profit_erosion_per_day,
        'days_until_zero_profit': days_until_zero_profit,
        'flags': flags,
    }


# ---------------------------------------------------------------------------
# P&L Calculator for CPA/Bookkeeper tax reporting
# ---------------------------------------------------------------------------
def calc_pnl(prop, metrics):
    """Generate tax-ready P&L structure from property data.
    For flips (dealer property), all costs are capitalized to basis (COGS).
    Profits are ordinary income subject to self-employment tax.
    """
    cd_purchase = prop.get('closing_disclosure_purchase', {})
    cd_sale = prop.get('closing_disclosure_sale', {})

    # ---- GROSS INCOME ----
    sale_price = prop.get('sale_price', 0) or metrics['effective_sale']
    seller_concessions = 0
    if cd_sale and cd_sale.get('line_items'):
        for item in cd_sale['line_items']:
            if 'concession' in item.get('description', '').lower():
                seller_concessions += item.get('amount', 0)
    net_sale_proceeds = sale_price - seller_concessions

    # ---- COGS: Acquisition ----
    purchase_price = metrics['purchase_price']

    # Acquisition closing costs — itemized from CD or lump sum
    if cd_purchase and cd_purchase.get('line_items'):
        acq_closing_items = cd_purchase['line_items']
        acq_closing_total = sum(item.get('amount', 0) for item in acq_closing_items)
    else:
        acq_closing_items = []
        acq_closing_total = metrics['acq_closing_cost']

    # ---- COGS: Renovation — grouped by tax category ----
    expenses = prop.get('expenses', [])
    renovation_groups = {}  # {pnl_label: [expenses]}
    selling_expenses = []
    for e in expenses:
        if e.get('is_credit'):
            continue
        cat = e.get('category', 'Other')
        pnl_label, tax_type = EXPENSE_TAX_MAP.get(cat, ('Renovation - Other', 'cogs'))
        if tax_type == 'selling':
            selling_expenses.append(e)
        else:
            renovation_groups.setdefault(pnl_label, []).append(e)

    renovation_subtotals = {}
    for label, exps in renovation_groups.items():
        renovation_subtotals[label] = sum(e.get('amount', 0) for e in exps)
    renovation_total = sum(renovation_subtotals.values())

    # Credits
    total_credits = sum(e.get('amount', 0) for e in expenses if e.get('is_credit'))
    renovation_total -= total_credits

    # ---- COGS: Holding Costs (capitalized for flips) ----
    holding = prop.get('holding_costs', {})
    months_held = metrics['months_held']
    holding_breakdown = {
        'Mortgage Interest': metrics['total_mortgage_payments'],
        'Property Taxes': (holding.get('monthly_taxes', 0) or 0) * months_held,
        'Insurance': (holding.get('monthly_insurance', 0) or 0) * months_held,
        'Utilities': (holding.get('monthly_utilities', 0) or 0) * months_held,
        'HOA': (holding.get('monthly_hoa', 0) or 0) * months_held,
        'Lawn/Landscape': (holding.get('monthly_lawn', 0) or 0) * months_held,
        'Other': (holding.get('monthly_other', 0) or 0) * months_held,
    }
    # Remove zero items
    holding_breakdown = {k: v for k, v in holding_breakdown.items() if v > 0}
    holding_total = metrics['total_holding_cost']

    total_cogs = purchase_price + acq_closing_total + renovation_total + holding_total

    # ---- SELLING COSTS ----
    commission = metrics['sale_commission']
    commission_pct = metrics['sale_commission_pct']
    sale_closing = metrics['sale_closing']
    sale_closing_pct = metrics['sale_closing_cost_pct']
    selling_expense_total = sum(e.get('amount', 0) for e in selling_expenses)

    # Sale closing items from CD if available
    sale_closing_items = []
    if cd_sale and cd_sale.get('line_items'):
        sale_closing_items = cd_sale['line_items']
        sale_closing = sum(item.get('amount', 0) for item in sale_closing_items)

    total_selling = commission + sale_closing + selling_expense_total

    # ---- TOTALS ----
    total_costs = total_cogs + total_selling
    net_profit = net_sale_proceeds - total_costs

    # Self-employment tax estimate (15.3% — 12.4% SS + 2.9% Medicare)
    se_tax = net_profit * 0.153 if net_profit > 0 else 0
    net_after_se = net_profit - se_tax

    # Partnership split
    split_pct = prop.get('partner_split_pct', 50) / 100

    return {
        # Income
        'sale_price': sale_price,
        'seller_concessions': seller_concessions,
        'net_sale_proceeds': net_sale_proceeds,
        # COGS - Acquisition
        'purchase_price': purchase_price,
        'acq_closing_items': acq_closing_items,
        'acq_closing_total': acq_closing_total,
        'has_closing_disclosure_purchase': bool(cd_purchase.get('line_items')),
        'has_closing_disclosure_sale': bool(cd_sale.get('line_items')),
        # COGS - Renovation
        'renovation_subtotals': renovation_subtotals,
        'renovation_total': renovation_total,
        'total_credits': total_credits,
        # COGS - Holding
        'holding_breakdown': holding_breakdown,
        'holding_total': holding_total,
        'total_cogs': total_cogs,
        # Selling
        'commission': commission,
        'commission_pct': commission_pct,
        'sale_closing': sale_closing,
        'sale_closing_pct': sale_closing_pct,
        'sale_closing_items': sale_closing_items,
        'selling_expenses': [{'vendor': e.get('vendor', ''), 'description': e.get('description', ''), 'amount': e.get('amount', 0)} for e in selling_expenses],
        'selling_expense_total': selling_expense_total,
        'total_selling': total_selling,
        # Totals
        'total_costs': total_costs,
        'net_profit': net_profit,
        'se_tax_rate': 15.3,
        'se_tax': se_tax,
        'net_after_se': net_after_se,
        'partner_a_share': net_profit * split_pct,
        'partner_b_share': net_profit * (1 - split_pct),
        'partner_split_pct': prop.get('partner_split_pct', 50),
        # Context
        'address': prop.get('address', ''),
        'city': prop.get('city', ''),
        'state': prop.get('state', ''),
        'purchase_date': prop.get('purchase_date'),
        'sale_date': prop.get('sale_date'),
        'days_held': metrics['days_held'],
        'months_held': metrics['months_held'],
        'status': metrics['status'],
    }


def parse_closing_disclosure(pdf_bytes):
    """Parse a CFPB Closing Disclosure PDF and extract financial data."""
    try:
        import pdfplumber
    except ImportError:
        return {'error': 'pdfplumber not installed', 'line_items': [], 'raw_text': ''}

    result = {
        'loan_amount': 0,
        'interest_rate': 0,
        'closing_costs_total': 0,
        'cash_to_close': 0,
        'line_items': [],
        'raw_text': '',
    }

    try:
        pdf = pdfplumber.open(io.BytesIO(pdf_bytes))
        all_text = ''
        for page in pdf.pages:
            text = page.extract_text() or ''
            all_text += text + '\n\n'
        pdf.close()

        # Clean up encoding issues (ligatures, special chars)
        all_text = all_text.replace('\ufb01', 'fi').replace('\ufb02', 'fl')
        all_text = all_text.replace('\ufb00', 'ff').replace('\ufb03', 'ffi').replace('\ufb04', 'ffl')
        # Replace common PDF extraction artifacts
        all_text = re.sub(r'[^\x00-\x7F]', '', all_text)  # strip non-ASCII
        result['raw_text'] = all_text[:20000]

        # Extract loan amount — try multiple patterns
        for pattern in [
            r'Loan\s*Amount\s*\$?\s*([\d,]+\.?\d*)',
            r'Amount\s*\$?\s*([\d,]{4,}\.?\d*)',
            r'Principal.*?\$?\s*([\d,]{4,}\.\d{2})',
        ]:
            m = re.search(pattern, all_text, re.IGNORECASE)
            if m:
                val = float(m.group(1).replace(',', ''))
                if val > 10000:  # must be a meaningful loan amount
                    result['loan_amount'] = val
                    break

        # Extract interest rate — try multiple patterns
        for pattern in [
            r'Interest\s*Rate\s*[:\s]*([\d.]+)\s*%',
            r'Rate\s*[:\s]*([\d.]+)\s*%',
            r'([\d]{1,2}\.\d{1,4})\s*%',
        ]:
            m = re.search(pattern, all_text, re.IGNORECASE)
            if m:
                val = float(m.group(1))
                if 1 < val < 30:  # reasonable interest rate range
                    result['interest_rate'] = val
                    break

        # Extract cash to close — try multiple patterns
        # HUD-1 format: "Cash From Borrower $XX,XXX.XX" or "Cash From X To Seller $XX,XXX.XX"
        # CD format: "Cash to Close $XX,XXX.XX"
        for pattern in [
            r'Cash\s*to\s*Close\s*\$?\s*([\d,]+\.?\d*)',
            r'Cash\s*From\s*(?:X\s*To\s*)?Seller\s*\$?\s*([\d,]+\.?\d*)',
            r'Cash\s*[Ff]rom\s*Borrower\s*\$?\s*([\d,]+\.?\d*)',
            r'Cash\s*(?:from|to)\s*(?:Borrower|Seller)\s*\$?\s*([\d,]+\.?\d*)',
            r'Due\s*from\s*Borrower\s*at\s*Closing\s*\$?\s*([\d,]+\.?\d*)',
            r'303\.?\s*Cash\s*X?\s*[Ff]rom\s*(?:To\s*)?Borrower\s*\$?\s*([\d,]+\.?\d*)',
            r'TOTAL\s*CLOSING\s*COSTS?\s*\$?\s*([\d,]+\.?\d*)',
            r'Se\s*lement\s*charges\s*to\s*borrower.*?\$?\s*([\d,]+\.?\d*)',
        ]:
            m = re.search(pattern, all_text, re.IGNORECASE)
            if m:
                val = float(m.group(1).replace(',', ''))
                if val > 100:
                    result['cash_to_close'] = val
                    break

        # Extract line items — look for lines with dollar amounts
        # IMPORTANT: Filter out non-closing-cost items (sale price, loan amount, deposits, etc.)
        # Closing Disclosures have sections A-H for actual closing costs on page 2
        line_pattern = re.compile(r'^(.+?)\s+\$?([\d,]+\.\d{2})\s*$', re.MULTILINE)

        # These are NOT closing costs — they're summary/header items from page 1 and 3
        exclude_keywords = [
            'total', 'page', 'closing disclosure', 'loan estimate',
            'projected', 'annual', 'monthly', 'sale price', 'sales price',
            'contract price', 'purchase price', 'property value',
            'loan amount', 'principal', 'balance', 'payoff',
            'deposit', 'earnest', 'down payment', 'cash to close',
            'cash from', 'cash to', 'closing costs', 'paid already',
            'adjustments', 'aggregate', 'excess', 'net',
            'amount due', 'amount from', 'amount to',
            'before closing', 'at closing', 'summaries',
            'seller credit', 'seller-credit', 'final',
            'existing loan', 'first mortgage', 'second mortgage',
            'payoff amount', 'amount owed', 'proration',
            'gross amount', 'summary', 'subtotal',
            'contract sales', 'personal property',
            'constuc on draw', 'construction draw',  # lender draws, not closing costs
            'commission to', 'commission paid',  # commissions tracked separately
            'debt paydown', 'american express',  # personal payoffs
            'payoff of', 'loan payoff',
        ]

        # Only include items that look like actual fees/charges (typically under $50K)
        for match in line_pattern.finditer(all_text):
            desc = match.group(1).strip()
            amount = float(match.group(2).replace(',', ''))
            # Skip tiny amounts and anything over $50K (those are loan/price amounts, not fees)
            if amount < 5 or amount > 50000:
                continue
            desc_lower = desc.lower()
            # Skip excluded items
            if any(kw in desc_lower for kw in exclude_keywords):
                continue
            # Skip lines that are just numbers or very short
            if len(desc) < 3:
                continue
            # Classify for tax purposes
            tax_cat = 'Other - Review Required'
            for keyword, category in CD_TAX_KEYWORDS.items():
                if keyword in desc_lower:
                    tax_cat = category
                    break

            result['line_items'].append({
                'description': desc,
                'amount': amount,
                'tax_category': tax_cat,
            })

        # Deduplicate: remove items with same amount + similar description
        seen = set()
        deduped = []
        for item in result['line_items']:
            # Create a key from amount + first 20 chars of description
            key = f"{item['amount']:.2f}_{item['description'][:20].lower()}"
            if key not in seen:
                seen.add(key)
                deduped.append(item)
        result['line_items'] = deduped

        result['closing_costs_total'] = sum(item['amount'] for item in result['line_items'])

    except Exception as e:
        result['error'] = str(e)

    return result


def generate_pnl_csv_rows(pnl, prop):
    """Generate CSV rows for a property P&L report."""
    rows = []
    addr = f"{prop.get('address', '')} {prop.get('city', '')} {prop.get('state', '')}"
    rows.append(['PROFIT & LOSS STATEMENT', addr])
    rows.append(['Purchase Date', pnl.get('purchase_date', 'N/A')])
    rows.append(['Sale Date', pnl.get('sale_date', 'N/A')])
    rows.append(['Days Held', pnl.get('days_held', 0)])
    rows.append([])
    rows.append(['GROSS INCOME', '', 'Amount'])
    rows.append(['Sale Price', '', f"{pnl['sale_price']:.2f}"])
    if pnl['seller_concessions'] > 0:
        rows.append(['Less: Seller Concessions', '', f"-{pnl['seller_concessions']:.2f}"])
    rows.append(['Net Sale Proceeds', '', f"{pnl['net_sale_proceeds']:.2f}"])
    rows.append([])
    rows.append(['COST OF GOODS SOLD (Capitalized to Basis)', '', ''])
    rows.append(['Purchase Price', '', f"{pnl['purchase_price']:.2f}"])
    rows.append(['Acquisition Closing Costs', '', f"{pnl['acq_closing_total']:.2f}"])
    if pnl['acq_closing_items']:
        for item in pnl['acq_closing_items']:
            rows.append(['', f"  {item['description']}", f"{item['amount']:.2f}"])
    rows.append(['Renovation Costs', '', f"{pnl['renovation_total']:.2f}"])
    for label, amount in pnl['renovation_subtotals'].items():
        rows.append(['', f"  {label}", f"{amount:.2f}"])
    if pnl['total_credits'] > 0:
        rows.append(['', '  Less: Credits/Returns', f"-{pnl['total_credits']:.2f}"])
    rows.append(['Holding Costs (Capitalized)', '', f"{pnl['holding_total']:.2f}"])
    for label, amount in pnl['holding_breakdown'].items():
        rows.append(['', f"  {label}", f"{amount:.2f}"])
    rows.append(['TOTAL COGS', '', f"{pnl['total_cogs']:.2f}"])
    rows.append([])
    rows.append(['SELLING COSTS', '', ''])
    rows.append([f"RE Commissions ({pnl['commission_pct']}%)", '', f"{pnl['commission']:.2f}"])
    rows.append([f"Sale Settlement ({pnl['sale_closing_pct']}%)", '', f"{pnl['sale_closing']:.2f}"])
    if pnl['selling_expenses']:
        for se in pnl['selling_expenses']:
            rows.append([f"  {se['description']}", se['vendor'], f"{se['amount']:.2f}"])
    rows.append(['TOTAL SELLING COSTS', '', f"{pnl['total_selling']:.2f}"])
    rows.append([])
    rows.append(['TOTAL ALL COSTS', '', f"{pnl['total_costs']:.2f}"])
    rows.append([])
    rows.append(['NET PROFIT (LOSS)', '', f"{pnl['net_profit']:.2f}"])
    rows.append([f"Est. Self-Employment Tax ({pnl['se_tax_rate']}%)", '', f"{pnl['se_tax']:.2f}"])
    rows.append(['Net After SE Tax', '', f"{pnl['net_after_se']:.2f}"])
    rows.append([])
    rows.append([f"Partner A ({pnl['partner_split_pct']}%)", '', f"{pnl['partner_a_share']:.2f}"])
    rows.append([f"Partner B ({100 - pnl['partner_split_pct']}%)", '', f"{pnl['partner_b_share']:.2f}"])
    return rows


# ---------------------------------------------------------------------------
# Prospect calculation engine
# ---------------------------------------------------------------------------
def calc_prospect_metrics(prospect, settings):
    """Calculate flip + rental metrics for a prospect.
    Mirrors the partner's 'Profit & Loss Property Worksheet' layout/formulas.
    """
    # --- PROPERTY VALUE ---
    mls_list_price = prospect.get('mls_list_price', 0) or 0
    as_is_value = prospect.get('as_is_value', 0) or 0
    arv = prospect.get('arv', 0) or 0
    asking = prospect.get('asking_price', 0) or 0  # = Purchase Price

    # --- REHAB ---
    rehab = prospect.get('estimated_rehab', 0) or 0
    initial_prep = prospect.get('initial_prep', 0) or 0
    rehab_total = rehab + initial_prep

    # --- PURCHASE OFFER ASSUMPTIONS ---
    market_value = as_is_value if as_is_value > 0 else arv
    market_discount_pct = prospect.get('market_discount_pct', 0) or 0  # e.g. 0.40 = 40%
    # Target purchase = market value * (1 - discount)
    target_purchase = market_value * (1 - market_discount_pct) if market_discount_pct > 0 else asking
    discount_from_market = ((asking - market_value) / market_value) if market_value > 0 else 0
    discount_from_list = ((asking - mls_list_price) / mls_list_price) if mls_list_price > 0 else 0

    # --- HOLDING COSTS (worksheet layout) ---
    hold_months = prospect.get('holding_months', 0) or settings.get('holding_months', 6)
    monthly_utilities = prospect.get('monthly_utilities', 0) or 0
    monthly_landscape = prospect.get('monthly_landscape', 0) or 0
    monthly_insurance = prospect.get('monthly_insurance', 0) or 0
    monthly_taxes = prospect.get('monthly_taxes', 0) or 0
    monthly_hard_money = prospect.get('monthly_hard_money', 0) or 0
    monthly_hold_total = monthly_utilities + monthly_landscape + monthly_insurance + monthly_taxes + monthly_hard_money
    holding_total = monthly_hold_total * hold_months
    # If no detailed holding entered, fall back to settings
    if holding_total == 0:
        monthly_hold_total = settings.get('monthly_holding_cost', 2500)
        holding_total = monthly_hold_total * hold_months

    # --- SELLING COST DETAIL (worksheet layout) ---
    est_sales_price = arv if arv > 0 else asking
    settlement_pct = prospect.get('seller_settlement_pct', 0) or settings.get('closing_cost_pct', 1.5) / 100
    if settlement_pct > 1:
        settlement_pct = settlement_pct / 100  # handle if entered as 1.5 vs 0.015
    seller_settlement = est_sales_price * settlement_pct
    seller_concessions = prospect.get('seller_concessions', 0) or 0
    commission_pct = prospect.get('commission_pct', 0) or settings.get('commission_pct', 6.0) / 100
    if commission_pct > 1:
        commission_pct = commission_pct / 100
    re_commission = est_sales_price * commission_pct
    price_reduction = prospect.get('price_reduction', 0) or 0
    total_selling_costs = seller_settlement + seller_concessions + re_commission + price_reduction

    # --- ACQ CLOSING COSTS ---
    acq_closing = prospect.get('acq_closing_costs', 0) or 0

    # --- TOTAL COST SUMMARY (matches worksheet) ---
    total_cost = asking + acq_closing + rehab_total + holding_total + total_selling_costs

    # --- PROFITABILITY ANALYSIS ---
    gross_profit = est_sales_price - total_cost
    cash_on_cash_roi = (gross_profit / total_cost * 100) if total_cost > 0 else 0
    annualized_roi = cash_on_cash_roi * (12 / hold_months) if hold_months > 0 else 0

    # --- MAO / 70% RULE (additional) ---
    mult = settings.get('arv_multiplier', 0.70)
    mao_70 = (arv * 0.70) - rehab_total if arv else 0
    mao_custom = (arv * mult) - rehab_total if arv else 0
    spread = mao_custom - asking

    min_profit = settings.get('min_profit', 25000)
    min_roi = settings.get('min_roi', 15)
    flip_pass = gross_profit >= min_profit and cash_on_cash_roi >= min_roi
    flip_borderline = (not flip_pass and gross_profit >= min_profit * 0.7
                       and cash_on_cash_roi >= min_roi * 0.7)

    # --- RENTAL METRICS ---
    monthly_rent = prospect.get('monthly_rent_estimate', 0) or 0
    one_pct_rule = (monthly_rent / asking * 100) if asking > 0 else 0
    one_pct_pass = one_pct_rule >= 1.0

    expense_ratio = settings.get('rental_expense_ratio', 0.50)
    annual_rent = monthly_rent * 12
    estimated_expenses = annual_rent * expense_ratio
    noi = annual_rent - estimated_expenses
    cap_rate = (noi / asking * 100) if asking > 0 else 0
    grm = (asking / annual_rent) if annual_rent > 0 else 0

    down_pct = settings.get('down_payment_pct', 20) / 100
    rate = settings.get('interest_rate', 7.5) / 100
    term = settings.get('loan_term_years', 30)
    loan_amount = asking * (1 - down_pct)
    monthly_rate = rate / 12
    if monthly_rate > 0 and term > 0:
        n_payments = term * 12
        monthly_payment = loan_amount * (monthly_rate * (1 + monthly_rate)**n_payments) / ((1 + monthly_rate)**n_payments - 1)
    else:
        monthly_payment = 0
    annual_debt = monthly_payment * 12

    monthly_cashflow = monthly_rent - (estimated_expenses / 12) - monthly_payment
    annual_cashflow = monthly_cashflow * 12
    cash_invested = (asking * down_pct) + (asking * settlement_pct)
    cash_on_cash_rental = (annual_cashflow / cash_invested * 100) if cash_invested > 0 else 0
    dscr = (noi / annual_debt) if annual_debt > 0 else 0

    min_cashflow = settings.get('min_cashflow_per_door', 200)
    min_cap = settings.get('min_cap_rate', 5.0)
    min_coc = settings.get('min_cash_on_cash', 8.0)
    rental_pass = (cap_rate >= min_cap and cash_on_cash_rental >= min_coc
                   and monthly_cashflow >= min_cashflow)
    rental_borderline = (not rental_pass and cap_rate >= min_cap * 0.7
                         and monthly_cashflow >= min_cashflow * 0.5)

    return {
        # Property Value
        'mls_list_price': mls_list_price, 'as_is_value': as_is_value,
        # Purchase Offer
        'target_purchase': round(target_purchase, 0),
        'discount_from_market': round(discount_from_market * 100, 1),
        'discount_from_list': round(discount_from_list * 100, 1),
        # Rehab
        'rehab_total': round(rehab_total, 0), 'initial_prep': initial_prep,
        # Holding
        'holding_total': round(holding_total, 0),
        'monthly_hold_total': round(monthly_hold_total, 0),
        'monthly_utilities': monthly_utilities, 'monthly_landscape': monthly_landscape,
        'monthly_insurance': monthly_insurance, 'monthly_taxes': monthly_taxes,
        'monthly_hard_money': monthly_hard_money,
        # Selling
        'est_sales_price': round(est_sales_price, 0),
        'seller_settlement': round(seller_settlement, 0),
        'seller_concessions': seller_concessions,
        're_commission': round(re_commission, 0),
        'commission_pct_used': round(commission_pct * 100, 1),
        'settlement_pct_used': round(settlement_pct * 100, 1),
        'price_reduction': price_reduction,
        'total_selling_costs': round(total_selling_costs, 0),
        # Total Cost Summary
        'acq_closing': acq_closing,
        'total_cost': round(total_cost, 0),
        # Profitability (matches worksheet)
        'gross_profit': round(gross_profit, 0),
        'roi': round(cash_on_cash_roi, 1),
        'cash_on_cash_roi': round(cash_on_cash_roi, 1),
        'annualized_roi': round(annualized_roi, 1),
        'profit_margin': round((gross_profit / est_sales_price * 100) if est_sales_price > 0 else 0, 1),
        # MAO / 70% Rule
        'mao_70': round(mao_70, 0), 'mao_custom': round(mao_custom, 0),
        'spread': round(spread, 0),
        'total_cost_to_arv': round((total_cost / arv * 100) if arv > 0 else 0, 1),
        'flip_pass': flip_pass, 'flip_borderline': flip_borderline,
        'flip_verdict': 'PASS' if flip_pass else ('BORDERLINE' if flip_borderline else 'FAIL'),
        # Rental
        'one_pct_rule': round(one_pct_rule, 2), 'one_pct_pass': one_pct_pass,
        'cap_rate': round(cap_rate, 2), 'cash_on_cash': round(cash_on_cash_rental, 2),
        'dscr': round(dscr, 2), 'grm': round(grm, 1),
        'monthly_cashflow': round(monthly_cashflow, 0),
        'annual_cashflow': round(annual_cashflow, 0),
        'noi': round(noi, 0), 'monthly_payment': round(monthly_payment, 0),
        'loan_amount': round(loan_amount, 0), 'cash_invested': round(cash_invested, 0),
        'rental_pass': rental_pass, 'rental_borderline': rental_borderline,
        'rental_verdict': 'PASS' if rental_pass else ('BORDERLINE' if rental_borderline else 'FAIL'),
        # Thresholds
        'min_profit': min_profit, 'min_roi': min_roi,
        'arv_multiplier': mult,
    }


# ---------------------------------------------------------------------------
# Seed the Willowbrook data
# ---------------------------------------------------------------------------
def seed_willowbrook():
    """Pre-load the 740 Willowbrook Rd data from the spreadsheet."""
    data = load_data()
    for p in data['properties']:
        if 'willowbrook' in p.get('address', '').lower():
            return
    prop = {
        'id': 'willowbrook-740',
        'address': '740 Willowbrook Rd',
        'city': 'Chesapeake',
        'state': 'VA',
        'zip': '23320',
        'sqft': 0,
        'purchase_price': 430000,
        'arv': 645000,
        'sale_price': 0,
        'acq_closing_cost': 17782.57,
        'purchase_settlement': 61532.57,
        'emd': 10000,
        'appraisal_fee': 350,
        'commitment_fee': 999,
        'purchase_date': '2025-12-01',
        'estimated_sale_date': '2026-06-01',
        'sale_date': None,
        'listing_date': None,
        'rehab_budget': 52000,
        'sale_commission_pct': 4.0,
        'sale_closing_cost_pct': 1.5,
        'contingency_pct': 15.0,
        'partner_split_pct': 50.0,
        'status': 'renovation',
        'notes': 'Insurance paid at closing ($3,435.42) — partial reimbursement when we sell.',
        'holding_costs': {
            'monthly_mortgage': 2591.94,
            'monthly_insurance': 0,
            'monthly_taxes': 0,
            'monthly_utilities': 0,
            'monthly_hoa': 0,
            'monthly_lawn': 0,
            'monthly_other': 0,
        },
        'expenses': [
            {'date': '2026-01-15', 'vendor': 'Echols Plumbing', 'description': 'Draw 1 Paypal', 'amount': 3060, 'category': 'Labor - Plumbing', 'is_credit': False},
            {'date': '2026-01-22', 'vendor': 'Echols Plumbing', 'description': 'Draw 2 Paypal', 'amount': 1000, 'category': 'Labor - Plumbing', 'is_credit': False},
            {'date': '2026-02-01', 'vendor': 'Echols Plumbing', 'description': 'Draw 3 Paypal', 'amount': 3000, 'category': 'Labor - Plumbing', 'is_credit': False},
            {'date': '2026-01-20', 'vendor': 'Amazon', 'description': 'Building Materials', 'amount': 3107.57, 'category': 'Building Materials', 'is_credit': False},
            {'date': '2026-01-25', 'vendor': 'Lowes', 'description': 'Building Materials', 'amount': 2976, 'category': 'Building Materials', 'is_credit': False},
            {'date': '2026-02-01', 'vendor': 'Home Depot', 'description': 'Building Materials', 'amount': 3832.96, 'category': 'Building Materials', 'is_credit': False},
            {'date': '2026-01-28', 'vendor': 'Floor Trader', 'description': 'Flooring', 'amount': 2535.29, 'category': 'Flooring', 'is_credit': False},
            {'date': '2026-02-10', 'vendor': 'Echols Plumbing', 'description': 'Draw 4 Paypal', 'amount': 3500, 'category': 'Labor - Plumbing', 'is_credit': False},
            {'date': '2026-02-18', 'vendor': 'Echols Plumbing', 'description': 'Draw 5 Paypal', 'amount': 5000, 'category': 'Labor - Plumbing', 'is_credit': False},
            {'date': '2026-03-01', 'vendor': 'Echols Plumbing', 'description': 'Draw 6 Kitchen', 'amount': 7000, 'category': 'Labor - Kitchen', 'is_credit': False},
            {'date': '2026-03-10', 'vendor': 'Echols Plumbing', 'description': 'Draw 7 Kitchen Final', 'amount': 4180, 'category': 'Labor - Kitchen', 'is_credit': False},
            {'date': '2026-03-12', 'vendor': 'Virtual Tidewater', 'description': 'Marketing Pics', 'amount': 145, 'category': 'Marketing', 'is_credit': False},
            {'date': '2026-03-15', 'vendor': 'Echols Plumbing', 'description': 'Final Payment', 'amount': 5640, 'category': 'Labor - Plumbing', 'is_credit': False},
            {'date': '2026-03-05', 'vendor': 'I2G Source', 'description': 'Termite Repair', 'amount': 550, 'category': 'Repairs - Pest', 'is_credit': False},
            {'date': '2026-03-08', 'vendor': 'TJ Landscaping', 'description': 'Venmo', 'amount': 1000, 'category': 'Landscaping', 'is_credit': False},
        ],
        'draws': [
            {'date': '2026-01-20', 'description': 'Bank Draw 1', 'amount': 48400},
            {'date': '2026-02-15', 'description': 'Bank Draw 2', 'amount': 38650},
            {'date': '2026-03-10', 'description': 'Bank Draw Final', 'amount': 12050},
        ],
        'mortgage_payments': [
            {'date': '2026-01-07', 'amount': 2591.94},
            {'date': '2026-02-05', 'amount': 2591.94},
        ],
    }
    data['properties'].append(prop)
    save_data(data)


def seed_second_property():
    """Pre-load the second property data from the spreadsheet."""
    data = load_data()
    for p in data['properties']:
        if p.get('id') == 'property-2':
            return
    prop = {
        'id': 'property-2',
        'address': '4420 Mallard Cres',
        'city': 'Portsmouth',
        'state': 'VA',
        'zip': '',
        'sqft': 0,
        'purchase_price': 199954,
        'arv': 354000,
        'sale_price': 354000,
        'acq_closing_cost': 5259.16,
        'purchase_settlement': 5259.16,
        'emd': 0,
        'appraisal_fee': 0,
        'commitment_fee': 0,
        'purchase_date': None,
        'estimated_sale_date': None,
        'sale_date': None,
        'listing_date': None,
        'rehab_budget': 30000,
        'sale_commission_pct': 6.0,
        'sale_closing_cost_pct': 1.5,
        'contingency_pct': 15.0,
        'partner_split_pct': 50.0,
        'status': 'active',
        'notes': 'Insurance: $843. No mortgage payments, no draws received, no utilities.',
        'holding_costs': {
            'monthly_mortgage': 0,
            'monthly_insurance': 0,
            'monthly_taxes': 0,
            'monthly_utilities': 0,
            'monthly_hoa': 0,
            'monthly_lawn': 0,
            'monthly_other': 0,
        },
        'expenses': [
            {'date': '', 'vendor': 'Echols Plumbing', 'description': 'Draw 1 Paypal', 'amount': 3500, 'category': 'Labor - Plumbing', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Plumbing', 'description': 'Draw 2 Paypal', 'amount': 3500, 'category': 'Labor - Plumbing', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Plumbing', 'description': 'Draw 3 Paypal', 'amount': 5000, 'category': 'Labor - Plumbing', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Plumbing', 'description': 'Draw 4 Paypal', 'amount': 5000, 'category': 'Labor - Plumbing', 'is_credit': False},
            {'date': '', 'vendor': 'Lowes', 'description': 'Appliances', 'amount': 2490.86, 'category': 'Appliances', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Plumbing', 'description': 'Draw 5 Paypal', 'amount': 4000, 'category': 'Labor - General', 'is_credit': False},
            {'date': '', 'vendor': 'Virtual Tidewater', 'description': 'Photos', 'amount': 145, 'category': 'Marketing', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Plumbing', 'description': 'Draw 6', 'amount': 4500, 'category': 'Labor - General', 'is_credit': False},
            {'date': '', 'vendor': 'Tim', 'description': 'Appraised PayPal - Listing Price Advice', 'amount': 200, 'category': 'Labor - General', 'is_credit': False},
            {'date': '', 'vendor': 'Patrick Murns', 'description': 'Venmo', 'amount': 150, 'category': 'Labor - General', 'is_credit': False},
        ],
        'draws': [],
        'mortgage_payments': [],
    }
    data['properties'].append(prop)
    save_data(data)


def seed_third_property():
    """Pre-load the third property data from the spreadsheet."""
    data = load_data()
    for p in data['properties']:
        if p.get('id') == 'property-3':
            return
    prop = {
        'id': 'property-3',
        'address': '737 Gemstone Ln',
        'city': 'Virginia Beach',
        'state': 'VA',
        'zip': '',
        'sqft': 0,
        'purchase_price': 175000,
        'arv': 238900,
        'sale_price': 238900,
        'acq_closing_cost': 0,
        'purchase_settlement': 0,
        'emd': 0,
        'appraisal_fee': 0,
        'commitment_fee': 0,
        'purchase_date': None,
        'estimated_sale_date': None,
        'sale_date': None,
        'listing_date': None,
        'rehab_budget': 22000,
        'sale_commission_pct': 6.0,
        'sale_closing_cost_pct': 1.5,
        'contingency_pct': 15.0,
        'partner_split_pct': 50.0,
        'status': 'active',
        'notes': 'Condo property. No settlement charges, no draws, no mortgage.',
        'holding_costs': {
            'monthly_mortgage': 0,
            'monthly_insurance': 0,
            'monthly_taxes': 0,
            'monthly_utilities': 0,
            'monthly_hoa': 0,
            'monthly_lawn': 0,
            'monthly_other': 0,
        },
        'expenses': [
            {'date': '', 'vendor': 'Lowes', 'description': 'Building Materials/Appliances', 'amount': 1712.84, 'category': 'Building Materials', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Plumbing', 'description': 'Draw 1 Paypal', 'amount': 3500, 'category': 'Labor - Plumbing', 'is_credit': False},
            {'date': '', 'vendor': 'Super Steamer', 'description': 'Carpet Cleaning', 'amount': 75, 'category': 'Labor - General', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Plumbing', 'description': 'Draw 2 Paypal', 'amount': 3500, 'category': 'Labor - Plumbing', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Plumbing', 'description': 'Draw 3 Paypal', 'amount': 5000, 'category': 'Labor - Plumbing', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Plumbing', 'description': 'Draw 4 Paypal', 'amount': 3500, 'category': 'Labor - Plumbing', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Plumbing', 'description': 'Draw 5 Final Payment', 'amount': 3175, 'category': 'Labor - Plumbing', 'is_credit': False},
            {'date': '', 'vendor': 'Homewisedocs', 'description': 'Condo Docs', 'amount': 326.95, 'category': 'Permits', 'is_credit': False},
        ],
        'draws': [],
        'mortgage_payments': [],
    }
    data['properties'].append(prop)
    save_data(data)


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.route('/')
def flip_dashboard():
    return render_template('flip_tracker.html')


@app.route('/api/flips', methods=['GET'])
def get_flips():
    data = load_data()
    result = []
    for prop in data['properties']:
        metrics = calc_property_metrics(prop)
        result.append({**prop, 'metrics': metrics})
    return jsonify({'properties': result, 'settings': data.get('settings', {})})


@app.route('/api/flips', methods=['POST'])
def add_flip():
    data = load_data()
    prop = request.json
    if not prop.get('id'):
        prop['id'] = prop.get('address', 'property').lower().replace(' ', '-') + '-' + str(len(data['properties']))
    prop.setdefault('expenses', [])
    prop.setdefault('draws', [])
    prop.setdefault('mortgage_payments', [])
    prop.setdefault('holding_costs', {
        'monthly_mortgage': 0, 'monthly_insurance': 0, 'monthly_taxes': 0,
        'monthly_utilities': 0, 'monthly_hoa': 0, 'monthly_lawn': 0, 'monthly_other': 0,
    })
    data['properties'].append(prop)
    save_data(data)
    metrics = calc_property_metrics(prop)
    return jsonify({**prop, 'metrics': metrics})


@app.route('/api/flips/<prop_id>', methods=['PUT'])
def update_flip(prop_id):
    data = load_data()
    for i, prop in enumerate(data['properties']):
        if prop.get('id') == prop_id:
            data['properties'][i].update(request.json)
            save_data(data)
            metrics = calc_property_metrics(data['properties'][i])
            return jsonify({**data['properties'][i], 'metrics': metrics})
    return jsonify({'error': 'Not found'}), 404


@app.route('/api/flips/<prop_id>/expense', methods=['POST'])
def add_expense(prop_id):
    data = load_data()
    for prop in data['properties']:
        if prop.get('id') == prop_id:
            prop.setdefault('expenses', []).append(request.json)
            save_data(data)
            metrics = calc_property_metrics(prop)
            return jsonify({**prop, 'metrics': metrics})
    return jsonify({'error': 'Not found'}), 404


@app.route('/api/flips/<prop_id>/draw', methods=['POST'])
def add_draw(prop_id):
    data = load_data()
    for prop in data['properties']:
        if prop.get('id') == prop_id:
            prop.setdefault('draws', []).append(request.json)
            save_data(data)
            metrics = calc_property_metrics(prop)
            return jsonify({**prop, 'metrics': metrics})
    return jsonify({'error': 'Not found'}), 404


@app.route('/api/flips/<prop_id>/mortgage', methods=['POST'])
def add_mortgage(prop_id):
    data = load_data()
    for prop in data['properties']:
        if prop.get('id') == prop_id:
            prop.setdefault('mortgage_payments', []).append(request.json)
            save_data(data)
            metrics = calc_property_metrics(prop)
            return jsonify({**prop, 'metrics': metrics})
    return jsonify({'error': 'Not found'}), 404


@app.route('/api/flips/<prop_id>', methods=['DELETE'])
def delete_flip(prop_id):
    data = load_data()
    data['properties'] = [p for p in data['properties'] if p.get('id') != prop_id]
    save_data(data)
    return jsonify({'ok': True})


@app.route('/api/flips/settings', methods=['GET'])
def get_flip_settings():
    data = load_data()
    return jsonify(data.get('settings', {}))


@app.route('/api/flips/settings', methods=['POST'])
def update_flip_settings():
    data = load_data()
    data['settings'] = request.json
    save_data(data)
    return jsonify(data['settings'])


# ---------------------------------------------------------------------------
# Prospect / Deal Analyzer routes
# ---------------------------------------------------------------------------
PROSPECT_STAGES = ['new_lead', 'analyzing', 'offer_sent', 'under_contract', 'passed', 'converted']


@app.route('/api/prospects', methods=['GET'])
def get_prospects():
    data = load_data()
    settings = data.get('prospect_settings', _default_prospect_settings())
    result = []
    for p in data.get('prospects', []):
        # Backward-compat: ensure stage_history exists
        p.setdefault('stage_history', [{'stage': p.get('stage', 'new_lead'), 'date': p.get('date_added', '')}])
        metrics = calc_prospect_metrics(p, settings)
        result.append({**p, 'metrics': metrics})
    return jsonify({'prospects': result, 'settings': settings})


@app.route('/api/prospects', methods=['POST'])
def add_prospect():
    data = load_data()
    p = request.json
    address = (p.get('address', '') or '').strip().lower()

    # If same address already exists, UPDATE instead of creating duplicate
    # Match on address alone — ignore city since user might type it differently
    existing = None
    if address:
        for i, existing_p in enumerate(data.get('prospects', [])):
            ex_addr = (existing_p.get('address', '') or '').strip().lower()
            if ex_addr == address:
                existing = i
                break

    if existing is not None:
        # Update existing prospect with new scenario numbers
        data['prospects'][existing].update(p)
        data['prospects'][existing]['last_analyzed'] = datetime.now().strftime('%Y-%m-%d %H:%M')
        save_data(data)
        settings = data.get('prospect_settings', _default_prospect_settings())
        metrics = calc_prospect_metrics(data['prospects'][existing], settings)
        return jsonify({**data['prospects'][existing], 'metrics': metrics, 'updated': True})

    # New prospect
    if not p.get('id'):
        slug = (p.get('address', 'prospect') or 'prospect').lower().replace(' ', '-')[:30]
        p['id'] = slug + '-' + datetime.now().strftime('%Y%m%d%H%M%S')
    p.setdefault('stage', 'new_lead')
    p.setdefault('date_added', datetime.now().strftime('%Y-%m-%d'))
    p['last_analyzed'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    p.setdefault('verdict', None)
    p.setdefault('notes', '')
    p.setdefault('source', '')
    p.setdefault('beds', 0)
    p.setdefault('baths', 0)
    p.setdefault('sqft', 0)
    p.setdefault('year_built', 0)
    p.setdefault('monthly_rent_estimate', 0)
    p.setdefault('stage_history', [{'stage': 'new_lead', 'date': datetime.now().strftime('%Y-%m-%d %H:%M')}])
    data.setdefault('prospects', []).append(p)
    save_data(data)
    settings = data.get('prospect_settings', _default_prospect_settings())
    metrics = calc_prospect_metrics(p, settings)
    return jsonify({**p, 'metrics': metrics})


@app.route('/api/prospects/<prospect_id>', methods=['PUT'])
def update_prospect(prospect_id):
    data = load_data()
    for i, p in enumerate(data.get('prospects', [])):
        if p.get('id') == prospect_id:
            data['prospects'][i].update(request.json)
            save_data(data)
            settings = data.get('prospect_settings', _default_prospect_settings())
            metrics = calc_prospect_metrics(data['prospects'][i], settings)
            return jsonify({**data['prospects'][i], 'metrics': metrics})
    return jsonify({'error': 'Not found'}), 404


@app.route('/api/prospects/<prospect_id>', methods=['DELETE'])
def delete_prospect(prospect_id):
    data = load_data()
    data['prospects'] = [p for p in data.get('prospects', []) if p.get('id') != prospect_id]
    save_data(data)
    return jsonify({'ok': True})


@app.route('/api/prospects/<prospect_id>/stage', methods=['PUT'])
def update_prospect_stage(prospect_id):
    data = load_data()
    stage = request.json.get('stage')
    if stage not in PROSPECT_STAGES:
        return jsonify({'error': f'Invalid stage. Must be one of: {PROSPECT_STAGES}'}), 400
    for i, p in enumerate(data.get('prospects', [])):
        if p.get('id') == prospect_id:
            data['prospects'][i]['stage'] = stage
            data['prospects'][i].setdefault('stage_history', []).append({
                'stage': stage,
                'date': datetime.now().strftime('%Y-%m-%d %H:%M')
            })
            save_data(data)
            return jsonify(data['prospects'][i])
    return jsonify({'error': 'Not found'}), 404


@app.route('/api/prospects/settings', methods=['GET'])
def get_prospect_settings():
    data = load_data()
    return jsonify(data.get('prospect_settings', _default_prospect_settings()))


@app.route('/api/prospects/settings', methods=['POST'])
def update_prospect_settings():
    data = load_data()
    data['prospect_settings'] = request.json
    save_data(data)
    return jsonify(data['prospect_settings'])


@app.route('/api/prospects/<prospect_id>/convert', methods=['POST'])
def convert_prospect(prospect_id):
    data = load_data()
    prospect = None
    for i, p in enumerate(data.get('prospects', [])):
        if p.get('id') == prospect_id:
            prospect = p
            prospect_idx = i
            break
    if not prospect:
        return jsonify({'error': 'Not found'}), 404

    # Create new property from prospect
    new_prop = {
        'id': prospect['id'] + '-prop',
        'address': prospect.get('address', ''),
        'city': prospect.get('city', ''),
        'state': prospect.get('state', 'VA'),
        'zip': prospect.get('zip', ''),
        'sqft': prospect.get('sqft', 0),
        'purchase_price': prospect.get('asking_price', 0),
        'arv': prospect.get('arv', 0),
        'sale_price': 0,
        'acq_closing_cost': 0,
        'purchase_settlement': 0,
        'emd': 0,
        'appraisal_fee': 0,
        'commitment_fee': 0,
        'purchase_date': datetime.now().strftime('%Y-%m-%d'),
        'estimated_sale_date': None,
        'sale_date': None,
        'listing_date': None,
        'rehab_budget': prospect.get('estimated_rehab', 0),
        'sale_commission_pct': 4.0,
        'sale_closing_cost_pct': 1.5,
        'contingency_pct': 15.0,
        'partner_split_pct': 50.0,
        'status': 'active',
        'notes': f'Converted from prospect. Original asking: ${prospect.get("asking_price", 0):,.0f}. {prospect.get("notes", "")}',
        'holding_costs': {
            'monthly_mortgage': 0, 'monthly_insurance': 0, 'monthly_taxes': 0,
            'monthly_utilities': 0, 'monthly_hoa': 0, 'monthly_lawn': 0, 'monthly_other': 0,
        },
        'expenses': [],
        'draws': [],
        'mortgage_payments': [],
    }
    data['properties'].append(new_prop)
    data['prospects'][prospect_idx]['stage'] = 'converted'
    save_data(data)
    return jsonify({'property': new_prop, 'prospect_stage': 'converted'})


@app.route('/api/export/csv')
def export_csv():
    """Export all flip data as CSV for backup."""
    import csv
    import io
    data = load_data()
    output = io.StringIO()
    writer = csv.writer(output)

    # Properties sheet
    writer.writerow(['=== PROPERTIES ==='])
    writer.writerow(['Address', 'City', 'State', 'ZIP', 'Purchase Price', 'ARV', 'Sale Price',
                     'Acq Closing', 'Settlement', 'EMD', 'Rehab Budget', 'Purchase Date',
                     'Sale Date', 'Listing Date', 'Commission %', 'Closing %', 'Split %',
                     'Monthly Mortgage', 'Monthly Insurance', 'Monthly Taxes', 'Monthly Utilities',
                     'Status', 'Notes'])
    for prop in data['properties']:
        hc = prop.get('holding_costs', {})
        m = calc_property_metrics(prop)
        writer.writerow([
            prop.get('address', ''), prop.get('city', ''), prop.get('state', ''), prop.get('zip', ''),
            prop.get('purchase_price', 0), prop.get('arv', 0), prop.get('sale_price', 0),
            prop.get('acq_closing_cost', 0), prop.get('purchase_settlement', 0), prop.get('emd', 0),
            prop.get('rehab_budget', 0), prop.get('purchase_date', ''), prop.get('sale_date', ''),
            prop.get('listing_date', ''), prop.get('sale_commission_pct', 4),
            prop.get('sale_closing_cost_pct', 1.5), prop.get('partner_split_pct', 50),
            hc.get('monthly_mortgage', 0), hc.get('monthly_insurance', 0),
            hc.get('monthly_taxes', 0), hc.get('monthly_utilities', 0),
            m['status'], prop.get('notes', ''),
        ])

    writer.writerow([])
    writer.writerow(['=== EXPENSES ==='])
    writer.writerow(['Property', 'Date', 'Vendor', 'Description', 'Category', 'Amount', 'Is Credit'])
    for prop in data['properties']:
        for e in prop.get('expenses', []):
            writer.writerow([prop.get('address', ''), e.get('date', ''), e.get('vendor', ''),
                           e.get('description', ''), e.get('category', ''), e.get('amount', 0),
                           e.get('is_credit', False)])

    writer.writerow([])
    writer.writerow(['=== DRAWS ==='])
    writer.writerow(['Property', 'Date', 'Description', 'Amount'])
    for prop in data['properties']:
        for d in prop.get('draws', []):
            writer.writerow([prop.get('address', ''), d.get('date', ''), d.get('description', ''), d.get('amount', 0)])

    writer.writerow([])
    writer.writerow(['=== MORTGAGE PAYMENTS ==='])
    writer.writerow(['Property', 'Date', 'Amount'])
    for prop in data['properties']:
        for mp in prop.get('mortgage_payments', []):
            writer.writerow([prop.get('address', ''), mp.get('date', ''), mp.get('amount', 0)])

    writer.writerow([])
    writer.writerow(['=== CALCULATED METRICS ==='])
    writer.writerow(['Property', 'Total Rehab', 'Total Costs', 'Gross Profit', 'ROI %', 'Profit Margin %', 'Cash In Deal', 'Partner Share', 'Days Held'])
    for prop in data['properties']:
        m = calc_property_metrics(prop)
        writer.writerow([prop.get('address', ''), m['total_rehab'], m['total_costs'], m['gross_profit'],
                        round(m['roi'], 2), round(m['profit_margin'], 2), m['cash_in_deal'],
                        m['partner_share'], m['days_held']])

    from flask import Response
    return Response(
        output.getvalue(),
        mimetype='text/csv',
        headers={'Content-Disposition': f'attachment; filename=flip-tracker-backup-{datetime.now().strftime("%Y%m%d")}.csv'}
    )


@app.route('/api/export/json')
def export_json():
    """Export raw JSON data for complete backup."""
    data = load_data()
    from flask import Response
    return Response(
        json.dumps(data, indent=2),
        mimetype='application/json',
        headers={'Content-Disposition': f'attachment; filename=flip-tracker-backup-{datetime.now().strftime("%Y%m%d")}.json'}
    )


@app.route('/api/flips/portfolio', methods=['GET'])
def portfolio_summary():
    data = load_data()
    props = data['properties']
    if not props:
        return jsonify({})
    total_invested = 0
    total_profit = 0
    total_rehab = 0
    active_count = 0
    sold_count = 0
    all_flags = []
    for prop in props:
        m = calc_property_metrics(prop)
        total_invested += m['cash_in_deal']
        total_profit += m['gross_profit']
        total_rehab += m['total_rehab']
        all_flags.extend([{**f, 'property': prop.get('address', 'Unknown')} for f in m['flags']])
        if m['status'] == 'sold':
            sold_count += 1
        else:
            active_count += 1
    avg_roi = (total_profit / total_invested * 100) if total_invested > 0 else 0
    return jsonify({
        'total_properties': len(props), 'active': active_count, 'sold': sold_count,
        'total_invested': total_invested, 'total_profit': total_profit,
        'total_rehab': total_rehab, 'avg_roi': avg_roi, 'all_flags': all_flags,
    })


# ---------------------------------------------------------------------------
# P&L API Routes
# ---------------------------------------------------------------------------
@app.route('/api/flips/<prop_id>/pnl', methods=['GET'])
def get_property_pnl(prop_id):
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404
    metrics = calc_property_metrics(prop)
    pnl = calc_pnl(prop, metrics)
    return jsonify(pnl)


@app.route('/api/flips/pnl/annual', methods=['GET'])
def get_annual_pnl():
    year = request.args.get('year', '2026')
    data = load_data()
    all_pnls = []
    totals = {
        'sale_price': 0, 'net_sale_proceeds': 0,
        'purchase_price': 0, 'acq_closing_total': 0,
        'renovation_total': 0, 'holding_total': 0, 'total_cogs': 0,
        'commission': 0, 'sale_closing': 0, 'selling_expense_total': 0,
        'total_selling': 0, 'total_costs': 0,
        'net_profit': 0, 'se_tax': 0, 'net_after_se': 0,
        'partner_a_share': 0, 'partner_b_share': 0,
    }
    for prop in data['properties']:
        metrics = calc_property_metrics(prop)
        pnl = calc_pnl(prop, metrics)
        all_pnls.append(pnl)
        for key in totals:
            totals[key] += pnl.get(key, 0)

    return jsonify({
        'year': year,
        'property_count': len(all_pnls),
        'properties': all_pnls,
        'totals': totals,
    })


@app.route('/api/flips/<prop_id>/pnl/csv', methods=['GET'])
def export_property_pnl_csv(prop_id):
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404
    metrics = calc_property_metrics(prop)
    pnl = calc_pnl(prop, metrics)
    rows = generate_pnl_csv_rows(pnl, prop)

    output = io.StringIO()
    writer = csv.writer(output)
    for row in rows:
        writer.writerow(row)
    resp = Response(output.getvalue(), mimetype='text/csv')
    addr = prop.get('address', 'property').replace(' ', '_')
    resp.headers['Content-Disposition'] = f'attachment; filename=PnL_{addr}.csv'
    return resp


@app.route('/api/flips/pnl/annual/csv', methods=['GET'])
def export_annual_pnl_csv():
    year = request.args.get('year', '2026')
    data = load_data()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([f'{year} ANNUAL PROFIT & LOSS SUMMARY'])
    writer.writerow([])

    grand_totals = {'sale': 0, 'cogs': 0, 'selling': 0, 'profit': 0, 'se_tax': 0}

    for prop in data['properties']:
        metrics = calc_property_metrics(prop)
        pnl = calc_pnl(prop, metrics)
        rows = generate_pnl_csv_rows(pnl, prop)
        for row in rows:
            writer.writerow(row)
        writer.writerow([])
        writer.writerow(['=' * 50])
        writer.writerow([])
        grand_totals['sale'] += pnl['net_sale_proceeds']
        grand_totals['cogs'] += pnl['total_cogs']
        grand_totals['selling'] += pnl['total_selling']
        grand_totals['profit'] += pnl['net_profit']
        grand_totals['se_tax'] += pnl['se_tax']

    writer.writerow(['ANNUAL TOTALS'])
    writer.writerow(['Total Net Sale Proceeds', '', f"{grand_totals['sale']:.2f}"])
    writer.writerow(['Total COGS', '', f"{grand_totals['cogs']:.2f}"])
    writer.writerow(['Total Selling Costs', '', f"{grand_totals['selling']:.2f}"])
    writer.writerow(['Total Net Profit', '', f"{grand_totals['profit']:.2f}"])
    writer.writerow(['Total Est. SE Tax', '', f"{grand_totals['se_tax']:.2f}"])

    resp = Response(output.getvalue(), mimetype='text/csv')
    resp.headers['Content-Disposition'] = f'attachment; filename=Annual_PnL_{year}.csv'
    return resp


@app.route('/api/flips/<prop_id>/pnl/pdf', methods=['GET'])
def export_property_pnl_pdf(prop_id):
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404
    metrics = calc_property_metrics(prop)
    pnl = calc_pnl(prop, metrics)
    pdf_bytes = generate_pnl_pdf(pnl, prop)
    return send_file(io.BytesIO(pdf_bytes), mimetype='application/pdf',
                     download_name=f"PnL_{prop.get('address', 'property').replace(' ', '_')}.pdf")


@app.route('/api/flips/pnl/annual/pdf', methods=['GET'])
def export_annual_pnl_pdf():
    year = request.args.get('year', '2026')
    data = load_data()
    all_pnls = []
    for prop in data['properties']:
        metrics = calc_property_metrics(prop)
        pnl = calc_pnl(prop, metrics)
        all_pnls.append((pnl, prop))
    pdf_bytes = generate_annual_pnl_pdf(year, all_pnls)
    return send_file(io.BytesIO(pdf_bytes), mimetype='application/pdf',
                     download_name=f"Annual_PnL_{year}.pdf")


def generate_pnl_pdf(pnl, prop):
    """Generate a professional P&L PDF using reportlab."""
    try:
        from reportlab.lib.pagesizes import letter
        from reportlab.lib import colors
        from reportlab.lib.units import inch
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    except ImportError:
        # Fallback: return a simple text PDF
        return b'%PDF-1.0 reportlab not installed'

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=0.5 * inch, bottomMargin=0.5 * inch,
                            leftMargin=0.6 * inch, rightMargin=0.6 * inch)
    styles = getSampleStyleSheet()
    elements = []

    # Paragraph styles for wrapping text in table cells
    cell_style = ParagraphStyle('Cell', fontName='Helvetica', fontSize=9, leading=11)
    cell_bold = ParagraphStyle('CellBold', fontName='Helvetica-Bold', fontSize=9, leading=11)
    cell_sub = ParagraphStyle('CellSub', fontName='Helvetica', fontSize=8, leading=10, textColor=colors.Color(0.4, 0.4, 0.4))
    cell_indent = ParagraphStyle('CellIndent', fontName='Helvetica', fontSize=9, leading=11, leftIndent=16)
    cell_indent2 = ParagraphStyle('CellIndent2', fontName='Helvetica', fontSize=8, leading=10, leftIndent=32, textColor=colors.Color(0.4, 0.4, 0.4))
    amt_style = ParagraphStyle('Amt', fontName='Courier', fontSize=9, leading=11, alignment=2)  # right-aligned
    amt_bold = ParagraphStyle('AmtBold', fontName='Courier-Bold', fontSize=9, leading=11, alignment=2)

    # Header
    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=16, spaceAfter=4)
    sub_style = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=10, textColor=colors.grey)
    elements.append(Paragraph("Profit &amp; Loss Statement", title_style))
    addr = f"{prop.get('address', '')} &mdash; {prop.get('city', '')}, {prop.get('state', '')}"
    elements.append(Paragraph(addr, sub_style))
    dates = f"Purchased: {pnl.get('purchase_date', 'N/A')} | Sold: {pnl.get('sale_date', 'N/A')} | Days Held: {pnl.get('days_held', 0)}"
    elements.append(Paragraph(dates, sub_style))
    elements.append(Spacer(1, 14))

    def f(v):
        return f"${v:,.2f}" if v >= 0 else f"-${abs(v):,.2f}"

    def P(text, style=cell_style):
        return Paragraph(str(text), style)

    def A(val, bold=False):
        return Paragraph(f(val), amt_bold if bold else amt_style)

    # Build table with Paragraph objects for proper text wrapping
    data = []
    data.append([P('GROSS INCOME', cell_bold), '', ''])
    data.append([P('Sale Price', cell_indent), '', A(pnl['sale_price'])])
    if pnl['seller_concessions'] > 0:
        data.append([P('Less: Seller Concessions', cell_indent), '', A(-pnl['seller_concessions'])])
    data.append([P('Net Sale Proceeds', cell_indent), '', A(pnl['net_sale_proceeds'], True)])
    data.append(['', '', ''])
    data.append([P('COST OF GOODS SOLD (Capitalized to Basis)', cell_bold), '', ''])
    data.append([P('Purchase Price', cell_indent), '', A(pnl['purchase_price'])])
    data.append([P('Acquisition Closing Costs', cell_indent), '', A(pnl['acq_closing_total'])])
    if pnl['acq_closing_items']:
        for item in pnl['acq_closing_items'][:15]:
            data.append([P(item['description'], cell_indent2), P(item.get('tax_category', ''), cell_sub), A(item['amount'])])
    data.append([P('Renovation Costs', cell_indent), '', A(pnl['renovation_total'])])
    for label, amount in pnl['renovation_subtotals'].items():
        data.append([P(label, cell_indent2), '', A(amount)])
    if pnl['total_credits'] > 0:
        data.append([P('Less: Credits/Returns', cell_indent2), '', A(-pnl['total_credits'])])
    data.append([P('Holding Costs (Capitalized)', cell_indent), P('IRC §263A', cell_sub), A(pnl['holding_total'])])
    for label, amount in pnl['holding_breakdown'].items():
        data.append([P(label, cell_indent2), '', A(amount)])
    data.append([P('TOTAL COGS', cell_indent), '', A(pnl['total_cogs'], True)])
    data.append(['', '', ''])
    data.append([P('SELLING COSTS', cell_bold), '', ''])
    data.append([P(f"RE Commissions ({pnl['commission_pct']}%)", cell_indent), '', A(pnl['commission'])])
    data.append([P(f"Sale Settlement ({pnl['sale_closing_pct']}%)", cell_indent), '', A(pnl['sale_closing'])])
    for se in pnl.get('selling_expenses', []):
        data.append([P(se['description'], cell_indent2), P(se['vendor'], cell_sub), A(se['amount'])])
    data.append([P('TOTAL SELLING COSTS', cell_indent), '', A(pnl['total_selling'], True)])
    data.append(['', '', ''])
    data.append([P('TOTAL ALL COSTS', cell_bold), '', A(pnl['total_costs'], True)])
    data.append(['', '', ''])
    data.append([P('NET PROFIT (LOSS)', cell_bold), '', A(pnl['net_profit'], True)])
    data.append([P(f"Est. Self-Employment Tax ({pnl['se_tax_rate']}%)", cell_indent), P('12.4% SS + 2.9% Medicare', cell_sub), A(pnl['se_tax'])])
    data.append([P('Net After SE Tax', cell_indent), '', A(pnl['net_after_se'], True)])
    data.append(['', '', ''])
    data.append([P('PARTNERSHIP SPLIT', cell_bold), '', ''])
    data.append([P(f"Partner A ({pnl['partner_split_pct']}%)", cell_indent), '', A(pnl['partner_a_share'])])
    data.append([P(f"Partner B ({100 - pnl['partner_split_pct']}%)", cell_indent), '', A(pnl['partner_b_share'])])

    page_width = letter[0] - 1.2 * inch  # total usable width
    col_widths = [page_width * 0.45, page_width * 0.30, page_width * 0.25]
    table = Table(data, colWidths=col_widths)

    # Style
    style_cmds = [
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]
    # Highlight section headers and totals
    for i, row in enumerate(data):
        cell0 = row[0]
        cell0_text = cell0.text if hasattr(cell0, 'text') else str(cell0)
        if any(cell0_text.startswith(s) for s in ['GROSS', 'COST OF', 'SELLING', 'TOTAL ALL', 'NET PROFIT', 'PARTNERSHIP']):
            style_cmds.append(('BACKGROUND', (0, i), (-1, i), colors.Color(0.94, 0.94, 0.94)))
        if 'TOTAL' in cell0_text or 'NET PROFIT' in cell0_text:
            style_cmds.append(('LINEABOVE', (0, i), (-1, i), 0.75, colors.black))

    table.setStyle(TableStyle(style_cmds))
    elements.append(table)

    # Footer
    elements.append(Spacer(1, 24))
    footer = ParagraphStyle('Footer', parent=styles['Normal'], fontSize=7, textColor=colors.grey)
    elements.append(Paragraph(f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M')} — For tax preparation purposes. Consult your CPA.", footer))

    doc.build(elements)
    return buffer.getvalue()


def generate_annual_pnl_pdf(year, all_pnls):
    """Generate annual summary P&L PDF."""
    try:
        from reportlab.lib.pagesizes import letter
        from reportlab.lib import colors
        from reportlab.lib.units import inch
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    except ImportError:
        return b'%PDF-1.0 reportlab not installed'

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=0.5 * inch)
    styles = getSampleStyleSheet()
    elements = []

    def fmt(v):
        return f"${v:,.2f}" if v >= 0 else f"-${abs(v):,.2f}"

    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=16, spaceAfter=4)
    sub_style = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=10, textColor=colors.grey)
    elements.append(Paragraph(f"{year} Annual Profit & Loss Summary", title_style))
    elements.append(Paragraph(f"{len(all_pnls)} properties", sub_style))
    elements.append(Spacer(1, 16))

    # Summary table
    header = ['Property', 'Sale Price', 'COGS', 'Selling', 'Net Profit', 'SE Tax']
    data = [header]
    totals = [0, 0, 0, 0, 0]
    for pnl, prop in all_pnls:
        data.append([
            prop.get('address', ''),
            fmt(pnl['net_sale_proceeds']),
            fmt(pnl['total_cogs']),
            fmt(pnl['total_selling']),
            fmt(pnl['net_profit']),
            fmt(pnl['se_tax']),
        ])
        totals[0] += pnl['net_sale_proceeds']
        totals[1] += pnl['total_cogs']
        totals[2] += pnl['total_selling']
        totals[3] += pnl['net_profit']
        totals[4] += pnl['se_tax']
    data.append(['TOTALS', fmt(totals[0]), fmt(totals[1]), fmt(totals[2]), fmt(totals[3]), fmt(totals[4])])

    table = Table(data, colWidths=[2.0 * inch, 1.1 * inch, 1.1 * inch, 1.0 * inch, 1.1 * inch, 0.9 * inch])
    style_cmds = [
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
        ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.9, 0.9, 0.9)),
        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
        ('LINEABOVE', (0, -1), (-1, -1), 1, colors.black),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.Color(0.8, 0.8, 0.8)),
    ]
    table.setStyle(TableStyle(style_cmds))
    elements.append(table)

    elements.append(Spacer(1, 24))
    footer = ParagraphStyle('Footer', parent=styles['Normal'], fontSize=7, textColor=colors.grey)
    elements.append(Paragraph(f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M')} — For tax preparation purposes. Consult your CPA.", footer))

    doc.build(elements)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Closing Disclosure Upload Routes
# ---------------------------------------------------------------------------
@app.route('/api/flips/<prop_id>/closing-disclosure', methods=['POST'])
def upload_closing_disclosure(prop_id):
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404

    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    cd_type = request.form.get('type', 'purchase')  # 'purchase' or 'sale'
    pdf_bytes = file.read()

    # Parse the PDF
    parsed = parse_closing_disclosure(pdf_bytes)

    # Store base64-encoded PDF + parsed data
    cd_data = {
        'upload_date': datetime.now().strftime('%Y-%m-%d'),
        'filename': file.filename,
        'pdf_base64': base64.b64encode(pdf_bytes).decode('utf-8'),
        'loan_amount': parsed.get('loan_amount', 0),
        'interest_rate': parsed.get('interest_rate', 0),
        'closing_costs_total': parsed.get('closing_costs_total', 0),
        'cash_to_close': parsed.get('cash_to_close', 0),
        'line_items': parsed.get('line_items', []),
        'raw_text': parsed.get('raw_text', ''),
    }

    key = f'closing_disclosure_{cd_type}'
    prop[key] = cd_data
    save_data(data)

    return jsonify({
        'success': True,
        'type': cd_type,
        'filename': file.filename,
        'line_items': parsed.get('line_items', []),
        'loan_amount': parsed.get('loan_amount', 0),
        'interest_rate': parsed.get('interest_rate', 0),
        'closing_costs_total': parsed.get('closing_costs_total', 0),
        'cash_to_close': parsed.get('cash_to_close', 0),
        'error': parsed.get('error'),
    })


@app.route('/api/flips/<prop_id>/closing-disclosure', methods=['PUT'])
def update_closing_disclosure(prop_id):
    """Save user-corrected parsed data after review."""
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404

    body = request.json
    cd_type = body.get('type', 'purchase')
    key = f'closing_disclosure_{cd_type}'

    if key not in prop:
        return jsonify({'error': f'No {cd_type} closing disclosure uploaded'}), 400

    # Update line items with user corrections
    prop[key]['line_items'] = body.get('line_items', prop[key].get('line_items', []))
    prop[key]['closing_costs_total'] = sum(item.get('amount', 0) for item in prop[key]['line_items'])

    save_data(data)
    return jsonify({'success': True})


def seed_22nd_street():
    """Pre-load 1324 22nd St Chesapeake data from the spreadsheet."""
    data = load_data()
    for p in data['properties']:
        if p.get('id') == '1324-22nd-st':
            return
    prop = {
        'id': '1324-22nd-st',
        'address': '1324 22nd St',
        'city': 'Chesapeake',
        'state': 'VA',
        'zip': '',
        'sqft': 0,
        'purchase_price': 92000,
        'arv': 280000,
        'sale_price': 280000,
        'acq_closing_cost': 9546.30,
        'purchase_settlement': 27127.16,
        'emd': 5000,
        'appraisal_fee': 250,
        'commitment_fee': 0,
        'purchase_date': None,
        'estimated_sale_date': None,
        'sale_date': None,
        'listing_date': None,
        'rehab_budget': 38048,
        'lender_rehab_budget': 83000,
        'sale_commission_pct': 4.0,
        'sale_closing_cost_pct': 1.5,
        'contingency_pct': 15.0,
        'partner_split_pct': 50.0,
        'status': 'active',
        'notes': 'Insurance paid at closing: $2,376.87 (reimbursable). Budget with Devin: $83,000. No bank draws received yet. No mortgage or utility payments.',
        'holding_costs': {
            'monthly_mortgage': 0,
            'monthly_insurance': 0,
            'monthly_taxes': 0,
            'monthly_utilities': 0,
            'monthly_hoa': 0,
            'monthly_lawn': 0,
            'monthly_other': 0,
        },
        'expenses': [
            {'date': '', 'vendor': 'Echols Contracting', 'description': 'Draw 1', 'amount': 3500, 'category': 'Labor - General', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Contracting', 'description': 'Draw 2', 'amount': 3500, 'category': 'Labor - General', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Contracting', 'description': 'Draw 3', 'amount': 3500, 'category': 'Labor - General', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Contracting', 'description': 'Draw 4', 'amount': 3500, 'category': 'Labor - General', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Contracting', 'description': 'Draw 5', 'amount': 3500, 'category': 'Labor - General', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Contracting', 'description': 'Draw 6', 'amount': 3500, 'category': 'Labor - General', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Contracting', 'description': 'Draw 7', 'amount': 3500, 'category': 'Labor - General', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Contracting', 'description': 'Draw 8', 'amount': 3500, 'category': 'Labor - General', 'is_credit': False},
            {'date': '', 'vendor': 'Echols Contracting', 'description': 'Draw 9', 'amount': 3500, 'category': 'Labor - General', 'is_credit': False},
            {'date': '', 'vendor': 'SRS Building Products', 'description': 'Building Materials', 'amount': 4845.38, 'category': 'Building Materials', 'is_credit': False},
            {'date': '', 'vendor': 'City of Chesapeake', 'description': 'Permits', 'amount': 1202.58, 'category': 'Permits', 'is_credit': False},
        ],
        'draws': [],
        'mortgage_payments': [],
    }
    data['properties'].append(prop)
    save_data(data)


# ---------------------------------------------------------------------------
# Startup
# ---------------------------------------------------------------------------
seed_willowbrook()
seed_second_property()
seed_third_property()
seed_22nd_street()

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5002))
    app.run(debug=True, port=port, host='0.0.0.0')
