#!/usr/bin/env python3
"""
Flip Tracker - Professional Real Estate Flip Investment Dashboard
Standalone Flask application for tracking renovation flip investments.
"""

from flask import Flask, render_template, request, jsonify, Response, send_file, session
import json
import os
import uuid
import base64
import io
import re
import csv
import tempfile
import shutil
import signal
import atexit
import threading
import urllib.request as _urllib_req
from collections import defaultdict
from datetime import datetime, timedelta
from functools import wraps

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB upload limit
app.secret_key = os.environ.get('SECRET_KEY', 'flip-tracker-secret-key-change-in-prod')

# Simple password protection — set APP_PASSWORD env var in Railway
APP_PASSWORD = os.environ.get('APP_PASSWORD', '')

# Postmark — set these in Railway env vars
POSTMARK_SERVER_TOKEN = os.environ.get('POSTMARK_SERVER_TOKEN', '')
POSTMARK_FROM_EMAIL   = os.environ.get('POSTMARK_FROM_EMAIL', 'noreply@yourfriendlyagent.net')
INSPECTION_NOTIFY_EMAIL = os.environ.get('INSPECTION_NOTIFY_EMAIL', 'barry@yourfriendlyagent.net')

# Inspection photo categories — mirror of inspector.html CATEGORIES array
INSPECTION_CATEGORIES = [
    {'id': 'kitchen',    'label': 'Kitchen',       'icon': '🍳', 'group': 'Interior'},
    {'id': 'bathrooms',  'label': 'Bathrooms',     'icon': '🚿', 'group': 'Interior'},
    {'id': 'bedrooms',   'label': 'Bedrooms',      'icon': '🛏',  'group': 'Interior'},
    {'id': 'living',     'label': 'Living Areas',  'icon': '🪑', 'group': 'Interior'},
    {'id': 'flooring',   'label': 'Flooring',      'icon': '🪵', 'group': 'Interior'},
    {'id': 'paint',      'label': 'Paint',         'icon': '🎨', 'group': 'Interior'},
    {'id': 'drywall',    'label': 'Drywall',       'icon': '🔲', 'group': 'Interior'},
    {'id': 'electrical', 'label': 'Electrical',    'icon': '⚡', 'group': 'Mechanical'},
    {'id': 'plumbing',   'label': 'Plumbing',      'icon': '🔧', 'group': 'Mechanical'},
    {'id': 'hvac',       'label': 'HVAC',          'icon': '🌡', 'group': 'Mechanical'},
    {'id': 'windows',    'label': 'Windows/Doors', 'icon': '🪟', 'group': 'Exterior'},
    {'id': 'exterior',   'label': 'Exterior',      'icon': '🏠', 'group': 'Exterior'},
    {'id': 'roof',       'label': 'Roof',          'icon': '🏗',  'group': 'Exterior'},
    {'id': 'foundation', 'label': 'Foundation',    'icon': '🧱', 'group': 'Structural'},
    {'id': 'demo',       'label': 'Demo/Cleanup',  'icon': '🚧', 'group': 'Site'},
    {'id': 'other',      'label': 'Other',         'icon': '📸', 'group': 'Other'},
]
_CAT_LABEL = {c['id']: c['label'] for c in INSPECTION_CATEGORIES}
_CAT_ICON  = {c['id']: c['icon']  for c in INSPECTION_CATEGORIES}
_CAT_GROUP = {c['id']: c['group'] for c in INSPECTION_CATEGORIES}


class DataSaveError(Exception):
    """Raised when save_data() cannot persist data to disk."""
    pass


@app.errorhandler(DataSaveError)
def handle_save_error(e):
    return jsonify({
        'error': 'Your changes could not be saved to disk. The data is still in memory for this session, '
                 'but may be lost if the server restarts. Check that the /data volume is mounted and writable.'
    }), 500

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
    'General Contractor': ('Renovation - Labor', 'cogs'),
    'Marketing': ('Selling Costs', 'selling'),
    'Staging': ('Selling Costs', 'selling'),
    'Other': ('Renovation - Other', 'cogs'),
}

# Vendor → default category mappings (case-insensitive vendor match)
VENDOR_CATEGORY_DEFAULTS = {
    'echols plumbing': 'General Contractor',
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
DATA_FILE = os.environ.get('DATA_FILE', '/data/flip_data.json')
PHOTOS_DIR = os.path.join(os.path.dirname(os.path.abspath(DATA_FILE)), 'photos')
_memory_store = None
_data_lock = threading.Lock()  # Serializes all load/save operations


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
        'overhead_expenses': [],
        'overhead_settings': {'monthly_rate': 0, 'start_date': None},
        'settings': {
            'default_commission_pct': 4.0,
            'default_closing_cost_pct': 1.5,
            'default_contingency_pct': 15.0,
            'partner_split_pct': 50.0,
        }
    }


def load_data():
    global _memory_store
    with _data_lock:
        if _memory_store is not None:
            return _memory_store
        try:
            with open(DATA_FILE, 'r') as f:
                _memory_store = json.load(f)
                _memory_store.setdefault('prospects', [])
                _memory_store.setdefault('prospect_settings', _default_prospect_settings())
                _memory_store.setdefault('overhead_expenses', [])
                _memory_store.setdefault('overhead_settings', {'monthly_rate': 0, 'start_date': None})
                return _memory_store
        except (FileNotFoundError, json.JSONDecodeError):
            # Try to seed from bundled flip_data.json in the app directory
            bundled = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'flip_data.json')
            try:
                with open(bundled, 'r') as f:
                    _memory_store = json.load(f)
                    _memory_store.setdefault('prospects', [])
                    _memory_store.setdefault('prospect_settings', _default_prospect_settings())
                    _memory_store.setdefault('overhead_expenses', [])
                    _memory_store.setdefault('overhead_settings', {'monthly_rate': 0, 'start_date': None})
                    # Write to volume without re-acquiring lock (already held)
                    _save_to_disk(_memory_store)
                    return _memory_store
            except (FileNotFoundError, json.JSONDecodeError):
                _memory_store = _default_data()
                return _memory_store


def _save_to_disk(data):
    """Write data atomically to disk. Caller must hold _data_lock. Raises DataSaveError on failure."""
    try:
        dir_name = os.path.dirname(DATA_FILE) or '.'
        os.makedirs(dir_name, exist_ok=True)
        fd, tmp_path = tempfile.mkstemp(dir=dir_name, suffix='.tmp')
        try:
            with os.fdopen(fd, 'w') as f:
                json.dump(data, f, indent=2)
            shutil.move(tmp_path, DATA_FILE)  # atomic on same filesystem
        except Exception:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
            raise
    except OSError as e:
        print(f"ERROR: could not write data file: {e}")
        raise DataSaveError(str(e))


def save_data(data):
    """Update in-memory store and persist to disk. Raises DataSaveError if disk write fails."""
    global _memory_store
    with _data_lock:
        _memory_store = data
        _save_to_disk(data)


# ---------------------------------------------------------------------------
# Graceful shutdown — flush memory store to disk before Railway kills container
# ---------------------------------------------------------------------------
def _flush_on_shutdown(*args):
    if _memory_store is not None:
        try:
            # Snapshot first in case another thread is modifying
            snapshot = json.loads(json.dumps(_memory_store))
            dir_name = os.path.dirname(DATA_FILE) or '.'
            os.makedirs(dir_name, exist_ok=True)
            fd, tmp = tempfile.mkstemp(dir=dir_name, suffix='.tmp')
            with os.fdopen(fd, 'w') as f:
                json.dump(snapshot, f, indent=2)
            shutil.move(tmp, DATA_FILE)
            print('[shutdown] Data flushed to volume successfully')
        except Exception as e:
            print(f'[shutdown] WARNING: could not flush data on shutdown: {e}')

atexit.register(_flush_on_shutdown)
signal.signal(signal.SIGTERM, _flush_on_shutdown)  # Railway sends SIGTERM before killing


# ---------------------------------------------------------------------------
# Calculation engine
# ---------------------------------------------------------------------------
def calc_property_metrics(prop):
    """Calculate all derived metrics for a property."""
    deal_type = prop.get('deal_type', 'flip')          # 'flip' | 'novation'
    assignment_fee = prop.get('assignment_fee', 0) or 0  # novation only: spread received at closing
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

    # ---- Rehab costs — renovation expenses only ----
    # Utilities, Marketing, Staging are holding/selling costs, not rehab.
    HOLDING_EXPENSE_CATS  = {'Utilities'}
    SELLING_EXPENSE_CATS  = {'Marketing', 'Staging'}

    total_rehab = 0
    total_holding_from_expenses = 0   # utility bills entered as expenses
    total_selling_from_expenses = 0   # marketing/staging entered as expenses
    total_credits = 0
    rehab_by_category    = {}
    holding_by_category  = {}

    for e in expenses:
        amt = e.get('amount', 0)
        cat = e.get('category', 'Other')
        if e.get('is_credit'):
            total_credits += amt
            total_rehab -= amt          # credits offset rehab
            continue
        if cat in HOLDING_EXPENSE_CATS:
            total_holding_from_expenses += amt
            holding_by_category[cat] = holding_by_category.get(cat, 0) + amt
        elif cat in SELLING_EXPENSE_CATS:
            total_selling_from_expenses += amt
        else:
            total_rehab += amt
            rehab_by_category[cat] = rehab_by_category.get(cat, 0) + amt

    # Budget tracking — renovation expenses only vs rehab budget
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

    # ---- Pre-purchase / out-of-pocket capital ----
    purchase_settlement = prop.get('purchase_settlement', 0) or 0
    emd = prop.get('emd', 0) or 0
    appraisal_fee = prop.get('appraisal_fee', 0) or 0
    commitment_fee = prop.get('commitment_fee', 0) or 0
    down_payment = prop.get('down_payment', 0) or 0

    # Lender cash back — companion loan proceeds returned to borrower at/after closing
    # (e.g. WCP 0%-down structure: buy CD requires $X, second loan gives back $Y net)
    # Reduces total cash out of pocket. Stored on the property or auto-populated from
    # the lender_cashback closing disclosure cash_to_close.
    lender_cashback = prop.get('lender_cashback', 0) or 0
    cd_lcb = prop.get('closing_disclosure_lender_cashback', {}) or {}
    if cd_lcb.get('cash_to_close', 0) and not lender_cashback:
        lender_cashback = cd_lcb['cash_to_close']

    # Compute cash_in_deal first — needs to happen before cash_invested fallback
    # When purchase_settlement is the CD cash-to-close, down_payment is already embedded in it.
    # EMD, commitment, and appraisal are typically paid OUTSIDE/BEFORE closing (not in CD cash-to-close),
    # so they are added separately. down_payment is NOT added to avoid double-counting.
    # lender_cashback is subtracted — it is money received BACK from the lender, reducing net OOP.
    if deal_type == 'novation':
        # Novation: Barry never closes on a purchase — no purchase settlement, no acq costs.
        # Cash in deal = only what Barry spent out of pocket (rehab + holding + utilities + staging).
        total_cash_oop = total_rehab + total_holding_cost + total_holding_from_expenses + total_selling_from_expenses - lender_cashback
    elif purchase_settlement > 0:
        total_cash_oop = purchase_settlement + emd + commitment_fee + appraisal_fee + total_rehab + total_holding_cost + total_holding_from_expenses + total_selling_from_expenses - lender_cashback
    else:
        total_cash_oop = acq_closing_cost + total_rehab + total_holding_cost + total_holding_from_expenses + total_selling_from_expenses - lender_cashback
    draw_surplus = max(total_draws - total_rehab, 0)
    draws_applied = min(total_draws, total_rehab)
    cash_in_deal = total_cash_oop - draws_applied - draw_surplus

    # cash_invested = out-of-pocket capital returned to Barry first before profit split
    #
    # Priority order:
    #   1. Manual override (user explicitly typed a value)
    #   2. purchase_settlement → cash_in_deal (full draw-adjusted OOP calculation)
    #   3. Purchase CD cash_to_close (uploaded CD has the real number)
    #   4. Sub-field sum (emd + appraisal + commitment + down) — last resort only
    #
    # IMPORTANT: sub-fields are treated as DISPLAY-ONLY breakdown rows in the
    # Cash Invested card. When a CD is uploaded or purchase_settlement is set,
    # sub-fields do NOT change the total — they just label what's inside it.
    # This lets Barry enter EMD, appraisal, commitment as breakdown components
    # without accidentally inflating the total.
    cd_pur = prop.get('closing_disclosure_purchase', {})
    cd_cash_to_close = cd_pur.get('cash_to_close', 0) or 0

    cash_invested_manual = prop.get('cash_invested', 0) or 0
    cash_invested_source = 'manual'
    if cash_invested_manual > 0:
        cash_invested = cash_invested_manual
    elif purchase_settlement > 0:
        # cash_in_deal accounts for draws and is the most accurate measure
        cash_invested = max(cash_in_deal, 0)
        cash_invested_source = 'calculated'
    elif cd_cash_to_close > 0:
        # CD uploaded — use real number, skip sub-field sum so breakdown rows
        # can be safely entered without changing the total
        cash_invested = cd_cash_to_close
        cash_invested_source = 'closing_disclosure'
    elif (emd + appraisal_fee + commitment_fee + down_payment) > 0:
        cash_invested = emd + appraisal_fee + commitment_fee + down_payment
        cash_invested_source = 'subfields'
    else:
        cash_invested = 0
        cash_invested_source = 'manual'

    # ---- Profit ----
    # For active/listed deals, use the full rehab budget as the projected rehab cost
    # (actual spend is almost always lower mid-renovation, inflating profit artificially).
    # For sold/closed deals, use actual spend — the work is done.
    _status_for_profit = prop.get('status', 'active')
    if _status_for_profit not in ('sold', 'closed') and sale_date is None:
        rehab_for_profit = max(total_rehab, budget) if budget > 0 else total_rehab
    else:
        rehab_for_profit = total_rehab

    if deal_type == 'novation':
        # Novation: Barry never purchases the property.
        # Revenue = assignment fee received at closing (the spread between sale price and seller guarantee).
        # Costs = only Barry's out-of-pocket spend: rehab, holding, utilities, staging.
        # Commissions and sale closing costs come out of the seller's side — NOT Barry's costs.
        effective_sale = assignment_fee if assignment_fee > 0 else effective_sale
        sale_commission = 0
        sale_closing = 0
        purchase_price_eff = 0
        acq_closing_eff = 0
        total_costs = rehab_for_profit + total_holding_from_expenses + total_selling_from_expenses + total_holding_cost
        gross_profit = (assignment_fee if assignment_fee > 0 else 0) - total_costs
    else:
        purchase_price_eff = purchase_price
        acq_closing_eff = acq_closing_cost
        total_costs = purchase_price + acq_closing_cost + rehab_for_profit + total_holding_from_expenses + total_selling_from_expenses + sale_commission + sale_closing + total_holding_cost
        gross_profit = effective_sale - total_costs
    profit_margin = (gross_profit / effective_sale * 100) if effective_sale > 0 else 0

    # Distribution base: use actual net proceeds from sale CD if available (most accurate),
    # otherwise fall back to accounting gross_profit.
    # Sale CD cash_to_close = what the seller actually received at closing.
    cd_sale_data = prop.get('closing_disclosure_sale', {})
    net_proceeds_at_close = cd_sale_data.get('cash_to_close', 0) or 0
    distribution_base = net_proceeds_at_close if net_proceeds_at_close > 0 else gross_profit

    # Overhead reimbursement — Barry fronts business overhead (payroll, staging, misc) and
    # gets reimbursed from deal proceeds before the profit split. Set per-deal at closeout.
    overhead_allocation = prop.get('overhead_allocation', 0) or 0

    # Distributable profit = distribution base after returning cash capital invested + overhead
    distributable_profit = distribution_base - cash_invested - overhead_allocation
    partner_share = distributable_profit * (partner_split_pct / 100)  # each partner's share of distributable
    partner_total = cash_invested + overhead_allocation + partner_share  # what Barry actually takes home

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
    # 'closed' is set manually via the closeout wizard and must not be overwritten.
    status = prop.get('status', 'active')
    if status != 'closed':
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

    # Two-box breakdown for Cash Invested card
    # Box 1: what was paid at/before the closing table (net of lender cashback)
    if purchase_settlement > 0:
        pre_closing_cash = purchase_settlement + emd + commitment_fee + appraisal_fee - lender_cashback
    elif cd_cash_to_close > 0:
        pre_closing_cash = cd_cash_to_close - lender_cashback
    else:
        pre_closing_cash = emd + commitment_fee + appraisal_fee + down_payment - lender_cashback
    # Box 2: rehab + holding costs minus all lender draws received (net post-acquisition spend)
    post_acq_net = total_rehab + total_holding_cost - total_draws

    return {
        'deal_type': deal_type, 'assignment_fee': assignment_fee,
        'purchase_price': purchase_price, 'arv': arv, 'sale_price': sale_price,
        'effective_sale': effective_sale, 'sqft': sqft, 'status': status,
        'total_rehab': total_rehab, 'net_rehab': total_rehab,
        'rehab_for_profit': rehab_for_profit,
        'profit_uses_budget': rehab_for_profit > total_rehab,
        'total_holding_from_expenses': total_holding_from_expenses,
        'total_selling_from_expenses': total_selling_from_expenses,
        'holding_by_category': holding_by_category,
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
        'emd': emd, 'appraisal_fee': appraisal_fee, 'down_payment': down_payment,
        'commitment_fee': commitment_fee, 'cash_invested': cash_invested,
        'cash_invested_source': cash_invested_source,
        'pre_closing_cash': pre_closing_cash, 'post_acq_net': post_acq_net,
        'net_proceeds_at_close': net_proceeds_at_close,
        'distribution_base': distribution_base,
        'overhead_allocation': overhead_allocation,
        'sale_commission': sale_commission, 'sale_commission_pct': sale_commission_pct,
        'sale_closing': sale_closing, 'sale_closing_cost_pct': sale_closing_cost_pct,
        'total_costs': total_costs, 'gross_profit': gross_profit,
        'profit_margin': profit_margin, 'partner_split_pct': partner_split_pct,
        'distributable_profit': distributable_profit,
        'partner_share': partner_share, 'partner_total': partner_total,
        'draw_surplus': draw_surplus,
        'lender_cashback': lender_cashback,
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
    deal_type = prop.get('deal_type', 'flip')
    cd_purchase = prop.get('closing_disclosure_purchase', {})
    cd_sale = prop.get('closing_disclosure_sale', {})
    cd_lender_cashback = prop.get('closing_disclosure_lender_cashback', {})

    # ---- GROSS INCOME ----
    if deal_type == 'novation':
        # Novation: revenue = assignment fee (the spread received at closing).
        # No sale price, no concessions — the CD shows assignment fee as a line item.
        assignment_fee = prop.get('assignment_fee', 0) or 0
        sale_price = assignment_fee
        seller_concessions = 0
        net_sale_proceeds = assignment_fee
    else:
        sale_price = prop.get('sale_price', 0) or metrics['effective_sale']
        seller_concessions = 0
        if cd_sale and cd_sale.get('line_items'):
            for item in cd_sale['line_items']:
                if 'concession' in item.get('description', '').lower():
                    seller_concessions += item.get('amount', 0)
        net_sale_proceeds = sale_price - seller_concessions

    # ---- COGS: Acquisition ----
    # Novation: Barry never closes on a purchase — no purchase price or acq costs.
    if deal_type == 'novation':
        purchase_price = 0
        acq_closing_items = []
        acq_closing_total = 0
    else:
        purchase_price = metrics['purchase_price']

    # Acquisition closing costs — itemized from CD or lump sum
    if cd_purchase and cd_purchase.get('line_items'):
        acq_closing_items = list(cd_purchase['line_items'])
        acq_closing_total = sum(item.get('amount', 0) for item in acq_closing_items)
        # Also append lender cashback CD costs (2nd loan fees) to acquisition line items
        if cd_lender_cashback and cd_lender_cashback.get('line_items'):
            lcb_items = cd_lender_cashback['line_items']
            acq_closing_items = acq_closing_items + lcb_items
            acq_closing_total += sum(item.get('amount', 0) for item in lcb_items)
    else:
        # No CD uploaded — build explicit named items from manual sub-fields.
        # EMD, appraisal, and commitment are pre-closing costs entered separately
        # from acq_closing_cost (the settlement/closing table costs).
        acq_closing_items = []
        emd_val       = prop.get('emd', 0) or 0
        appraisal_val = prop.get('appraisal_fee', 0) or 0
        commitment_val = prop.get('commitment_fee', 0) or 0
        base_closing  = metrics['acq_closing_cost']
        if emd_val > 0:
            acq_closing_items.append({'description': 'Earnest Money Deposit (EMD)',
                                      'amount': emd_val,
                                      'tax_category': 'Acquisition - Capitalized'})
        if appraisal_val > 0:
            acq_closing_items.append({'description': 'Appraisal Fee',
                                      'amount': appraisal_val,
                                      'tax_category': 'Loan Costs - Capitalized'})
        if commitment_val > 0:
            acq_closing_items.append({'description': 'Loan Commitment Fee',
                                      'amount': commitment_val,
                                      'tax_category': 'Loan Costs - Capitalized'})
        if base_closing > 0:
            acq_closing_items.append({'description': 'Settlement / Closing Costs',
                                      'amount': base_closing,
                                      'tax_category': 'Title & Settlement - Capitalized'})
        acq_closing_total = emd_val + appraisal_val + commitment_val + base_closing

    # ---- COGS: Renovation — grouped by tax category ----
    # (same for both flip and novation)
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
    # Novation: commissions and sale closing costs come from the seller's CD, not Barry's pocket.
    # Barry's only "selling" costs are staging/marketing he paid out of pocket.
    if deal_type == 'novation':
        commission = 0
        commission_pct = 0
        sale_closing = 0
        sale_closing_pct = 0
        sale_closing_items = []
        selling_expense_total = sum(e.get('amount', 0) for e in selling_expenses)
        total_selling = selling_expense_total
    else:
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

    # S-Corp: SE tax does not apply to distributions — profits flow to K-1 as ordinary income
    se_tax = 0
    net_after_se = net_profit

    # Partnership waterfall: return capital first, then split remainder.
    # Distribution base = actual sale CD net proceeds when available (what was actually
    # received at the closing table), otherwise fall back to accounting net_profit.
    # These differ because accounting net_profit deducts the full purchase price including
    # financed portions, while CD net proceeds are already net of loan payoff.
    split_pct = prop.get('partner_split_pct', 50) / 100
    cash_invested = metrics.get('cash_invested', 0)
    overhead_allocation = prop.get('overhead_allocation', 0) or 0
    net_proceeds_at_close = metrics.get('net_proceeds_at_close', 0)
    using_cd_proceeds = net_proceeds_at_close > 0
    distribution_base = net_proceeds_at_close if using_cd_proceeds else net_profit
    distributable_profit = distribution_base - cash_invested - overhead_allocation
    partner_a_share = distributable_profit * split_pct
    partner_b_share = distributable_profit * (1 - split_pct)
    partner_a_total = cash_invested + overhead_allocation + partner_a_share

    return {
        # Deal type
        'deal_type': deal_type,
        'assignment_fee': prop.get('assignment_fee', 0) or 0,
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
        'se_tax_rate': 0,
        'se_tax': se_tax,
        'net_after_se': net_after_se,
        'cash_invested': cash_invested,
        'overhead_allocation': overhead_allocation,
        'net_proceeds_at_close': net_proceeds_at_close,
        'distribution_base': distribution_base,
        'using_cd_proceeds': using_cd_proceeds,
        'distributable_profit': distributable_profit,
        'partner_a_share': partner_a_share,
        'partner_b_share': partner_b_share,
        'partner_a_total': partner_a_total,
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
        'mortgage_payment_count': len(prop.get('mortgage_payments', [])),
    }


def parse_closing_disclosure(pdf_bytes):
    """Parse settlement statements in multiple formats:
       - CFPB Closing Disclosure (post-2015, standard residential)
       - HUD-1 Settlement Statement (pre-2015 or commercial/investment)
       - ALTA Settlement Statement (cash/subject-to deals)
    """
    try:
        import pdfplumber
    except ImportError:
        return {'error': 'pdfplumber not installed', 'line_items': [], 'raw_text': ''}

    result = {
        'loan_amount': 0,
        'interest_rate': 0,
        'closing_costs_total': 0,
        'cash_to_close': 0,
        'sale_price': 0,
        'closing_date': '',
        'line_items': [],
        'raw_text': '',
        'form_type': 'unknown',
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
        all_text = all_text.replace('\x00', '')  # strip null bytes — PDF ligature artifacts (ti/tt replace with nothing)
        all_text = re.sub(r'[^\x00-\x7F]', '', all_text)  # strip remaining non-ASCII
        result['raw_text'] = all_text[:20000]

        # ---------------------------------------------------------------
        # Detect form type — drives which extraction logic to use
        # ---------------------------------------------------------------
        text_lower = all_text.lower()
        is_hud1 = (
            # 'settlement' may appear as 'selement' after null-byte ligature stripping
            ('settlement statement' in text_lower or 'selement statement' in text_lower) and
            ('u.s. department of housing' in text_lower or 'hud-1' in text_lower or
             '100. gross amount due from borrower' in text_lower or
             '200. amount paid by or in behalf' in text_lower or
             '300. cash at settlement' in text_lower or
             '300. cash at selement' in text_lower)
        )
        is_alta = (
            'alta settlement statement' in text_lower or
            ('american land title association' in text_lower and 'settlement statement' in text_lower)
        )
        is_cfpb = 'closing disclosure' in text_lower

        if is_hud1:
            result['form_type'] = 'HUD-1'
        elif is_alta:
            result['form_type'] = 'ALTA'
        elif is_cfpb:
            result['form_type'] = 'CFPB-CD'
        else:
            result['form_type'] = 'unknown'

        # ---------------------------------------------------------------
        # HUD-1 specific extraction
        # ---------------------------------------------------------------
        if is_hud1:
            def clean_hud_amount(s):
                """Handle PDF artifacts: $418.500.00 (period as thousands sep) → 418500.00"""
                s = s.replace(',', '')
                # Fix period-as-thousands-separator: X.XXX.XX → XXXXX.XX
                # Pattern: digit(s) . three-digits . two-digits at end
                s = re.sub(r'(\d+)\.(\d{3})\.(\d{2})$', r'\1\2.\3', s)
                return float(s)

            # Cash to close: Line 303
            # Actual text: "303. Cash [X} From D To Borrower $54,010.26 603. ..."
            # Brackets may be [X], [X}, {X], (X) due to PDF artifacts
            hud_cash_patterns = [
                # Bracketed X patterns — [X], {X}, (X) variants
                r'303\.?\s*Cash\s*[\[{(][Xx][\]})]\s*[Ff]rom\s*(?:[A-Z]\s*[Tt]o\s*)?Borrower\s*\$?\s*([\d,]+\.?\d*)',
                r'303\.?\s*Cash\s*[\[{(][Xx][\]})]\s*(?:\w+\s+)?[Ff]rom\s*(?:\w+\s*)?Borrower\s*\$?\s*([\d,]+\.?\d*)',
                # Bare X — "303. Cash X From To Borrower $42,996.91" (FROM borrower)
                r'303\.?\s*Cash\s+X\s+[Ff]rom\s+(?:[Tt]o\s+)?Borrower\s*\$?\s*([\d,]+\.?\d*)',
                # Bare X — "303. Cash From X To Borrower $25,106.88" (TO borrower / cashback)
                r'303\.?\s*Cash\s+[Ff]rom\s+X\s+[Tt]o\s+Borrower\s*\$?\s*([\d,]+\.?\d*)',
                # No checkbox text
                r'303\.?\s*Cash\s+[Ff]rom\s+Borrower\s*\$?\s*([\d,]+\.?\d*)',
                r'Cash\s*[\[{(][Xx][\]})]\s*[Ff]rom\s*\w*\s*[Tt]o\s*Borrower\s*\$?\s*([\d,]+\.?\d*)',
                r'Cash\s*[\[{(][Xx][\]})]\s*[Ff]rom\s*Borrower\s*\$?\s*([\d,]+\.?\d*)',
                r'Cash\s*[Ff]rom\s*Borrower\s*\$\s*([\d,]+\.\d{2})',
                # Catch-all: first dollar amount on any line 303 (handles TO borrower / cashback CDs)
                r'303\.?\s*Cash[^\n]*?\$\s*([\d,]+\.\d{2})',
            ]
            for pattern in hud_cash_patterns:
                m = re.search(pattern, all_text, re.IGNORECASE)
                if m:
                    try:
                        val = clean_hud_amount(m.group(1))
                        if val > 100:
                            result['cash_to_close'] = val
                            break
                    except ValueError:
                        continue

            # Loan amount: Line 202 — may have period as thousands separator
            hud_loan_patterns = [
                r'202\.?\s*Principal\s*amount\s*of\s*new\s*loan[s(]?\)?\s*[{]?\s*\$?\s*([\d,.]+)',
                r'202\.?\s*Principal\s*amount.*?\$?\s*([\d,.]+)',
                r'Principal\s*amount\s*of\s*new\s*loan\s*\$?\s*([\d,.]+)',
            ]
            for pattern in hud_loan_patterns:
                m = re.search(pattern, all_text, re.IGNORECASE)
                if m:
                    try:
                        val = clean_hud_amount(m.group(1))
                        if val > 10000:
                            result['loan_amount'] = val
                            break
                    except ValueError:
                        continue

            # Settlement date — on the line after "I. Settlement Date:" in some PDFs
            # 'Settlement' may appear as 'Selement' after null-byte ligature stripping
            # Try inline first, then next-line
            date_found = False
            for pattern in [
                r'(?:Settlement|Selement)\s*Date[:\s]+(\d{1,2}/\d{1,2}/\d{2,4})',
                r'I\.?\s*(?:Settlement|Selement)\s*Date[:\s]+(\d{1,2}/\d{1,2}/\d{2,4})',
            ]:
                m = re.search(pattern, all_text, re.IGNORECASE)
                if m:
                    result['closing_date'] = m.group(1).strip()
                    date_found = True
                    break
            if not date_found:
                # HUD-1 header layout: "I. Selement Date:" on one line, date on next line
                # after the agent name: "New World Title Company, LLC  04/01/2026"
                for pat in [
                    r'(?:Settlement|Selement)\s*Date[:\s]*\n[^\n]*?(\d{1,2}/\d{1,2}/\d{4})',
                    r'(?:Settlement|Selement)\s*Date[:\s]*\n(\d{1,2}/\d{1,2}/\d{2,4})',
                    r'I\.?\s*(?:Settlement|Selement)\s*Date[:\s]*\n[^\n]*?(\d{1,2}/\d{1,2}/\d{4})',
                ]:
                    m = re.search(pat, all_text, re.IGNORECASE)
                    if m:
                        result['closing_date'] = m.group(1).strip()
                        date_found = True
                        break
                if not date_found:
                    # Fallback: grab any date near Settlement Agent line
                    m = re.search(r'(?:Settlement|Selement)\s*Agent.*?(\d{1,2}/\d{1,2}/\d{4})', all_text, re.IGNORECASE | re.DOTALL)
                    if m:
                        result['closing_date'] = m.group(1).strip()

            # Sale price: Line 101
            for pattern in [
                r'101\.?\s*Contract\s*sales?\s*price\s*\$?\s*([\d,.]+)',
                r'Contract\s*sales?\s*price\s*\$?\s*([\d,.]+)',
            ]:
                m = re.search(pattern, all_text, re.IGNORECASE)
                if m:
                    try:
                        val = clean_hud_amount(m.group(1))
                        if val > 50000:
                            result['sale_price'] = val
                            break
                    except ValueError:
                        continue

            # HUD-1 line items: Lines 700-1399 are actual settlement fees
            # The HUD-1 lays out borrower column and seller column on the same line,
            # so we parse by line number range
            hud_line_pattern = re.compile(
                r'(\d{3,4})\.?\s+(.+?)\s+\$?\s*([\d,]+\.\d{2})(?:\s|$)',
                re.MULTILINE
            )
            hud_fee_sections = set(range(700, 1400))
            hud_exclude = [
                'commission', 'total settlement', 'gross amount',
                'less amounts paid', 'cash at settlement', 'cash from', 'cash to',
                'reduction in amount', 'payoff', 'existing loan',
                'settlement charges to seller',
                'policy limit',        # lender/owner title policy limit — amount not a fee
                'tle policy',          # stripped-ligature variant of 'title policy'
                'total selement',      # line 1400 total (stripped ligature of 'settlement')
                'poc by seller',       # paid outside closing by seller — not buyer's cost
                'seller selement',     # seller settlement fee
                'seller settlement',   # seller settlement fee
            ]

            seen_hud = {}
            for match in hud_line_pattern.finditer(all_text):
                line_num = int(match.group(1))
                desc = match.group(2).strip()
                try:
                    amount = float(match.group(3).replace(',', ''))
                except ValueError:
                    continue

                if line_num not in hud_fee_sections:
                    continue

                # Daily interest lines (901-904): regex captures the per-day rate
                # e.g. "@ $62.64 /day $1,879.20" — grab the total after "/day" instead
                if 901 <= line_num <= 904 and '@' in desc:
                    remaining = all_text[match.end():]
                    day_total = re.search(r'/day\s+\$?\s*([\d,]+\.\d{2})', remaining[:40])
                    if day_total:
                        try:
                            amount = float(day_total.group(1).replace(',', ''))
                        except ValueError:
                            pass

                # Some recording/tax lines show "$0.00" for one component but a non-zero
                # total at the end: "Deed $0.00 Mortgage $15.78 ... $15.78"
                # Scan ahead in the line for a non-zero amount when first capture is near-zero
                if amount < 5:
                    lookahead = all_text[match.start():match.start() + 300].split('\n')[0]
                    all_amounts = re.findall(r'\$?\s*([\d,]+\.\d{2})', lookahead)
                    for la in reversed(all_amounts):
                        try:
                            la_val = float(la.replace(',', ''))
                            if la_val >= 5:
                                amount = la_val
                                break
                        except ValueError:
                            continue

                if amount < 5 or amount > 50000:
                    continue
                desc_lower = desc.lower()
                if any(kw in desc_lower for kw in hud_exclude):
                    continue
                if len(desc) < 3:
                    continue

                tax_cat = 'Other - Review Required'
                for keyword, category in CD_TAX_KEYWORDS.items():
                    if keyword in desc_lower:
                        tax_cat = category
                        break

                amt_key = f"{amount:.2f}"
                if amt_key not in seen_hud or len(desc) > len(seen_hud[amt_key]['description']):
                    seen_hud[amt_key] = {'description': desc, 'amount': amount, 'tax_category': tax_cat}

            result['line_items'] = list(seen_hud.values())
            result['closing_costs_total'] = sum(i['amount'] for i in result['line_items'])
            return result

        # ---------------------------------------------------------------
        # ALTA and CFPB CD extraction (shared logic, slight differences)
        # ---------------------------------------------------------------

        # Extract loan amount
        for pattern in [
            r'Loan\s*Amount\s*\$?\s*([\d,]+\.?\d*)',
            r'202\.?\s*Principal\s*amount\s*of\s*new\s*loan[s]?\s*\$?\s*([\d,]+\.?\d*)',
            r'Amount\s*\$?\s*([\d,]{4,}\.?\d*)',
            r'Principal.*?\$?\s*([\d,]{4,}\.\d{2})',
        ]:
            m = re.search(pattern, all_text, re.IGNORECASE)
            if m:
                val = float(m.group(1).replace(',', ''))
                if val > 10000:
                    result['loan_amount'] = val
                    break

        # Extract interest rate
        for pattern in [
            r'Interest\s*Rate\s*[:\s]*([\d.]+)\s*%',
            r'Rate\s*[:\s]*([\d.]+)\s*%',
            r'([\d]{1,2}\.\d{1,4})\s*%',
        ]:
            m = re.search(pattern, all_text, re.IGNORECASE)
            if m:
                val = float(m.group(1))
                if 1 < val < 30:
                    result['interest_rate'] = val
                    break

        # Extract cash to close — CFPB CD and ALTA formats
        for pattern in [
            r'Cash\s*to\s*Close\s*\$?\s*([\d,]+\.?\d*)',
            r'Cash\s*From\s*(?:X\s*To\s*)?Seller\s*\$?\s*([\d,]+\.?\d*)',
            r'Cash\s*[Ff]rom\s*Borrower\s*\$?\s*([\d,]+\.?\d*)',
            r'Cash\s*(?:from|to)\s*(?:Borrower|Buyer|Seller)\s*\$?\s*([\d,]+\.?\d*)',
            r'Due\s*[Ff]rom\s*Borrower\s+\$?\s*([\d,]+\.?\d*)',
            r'Due\s*[Ff]rom\s*Buyer\s+\$?\s*([\d,]+\.?\d*)',
            r'Amount\s*Due\s*[Ff]rom\s*(?:Borrower|Buyer)\s+\$?\s*([\d,]+\.?\d*)',
            r'Due\s*from\s*Borrower\s*at\s*Closing\s*\$?\s*([\d,]+\.?\d*)',
            r'TOTAL\s*CLOSING\s*COSTS?\s*\$?\s*([\d,]+\.?\d*)',
        ]:
            m = re.search(pattern, all_text, re.IGNORECASE)
            if m:
                val = float(m.group(1).replace(',', ''))
                if val > 100:
                    result['cash_to_close'] = val
                    break

        # Extract sale price
        for pattern in [
            r'(?:Contract\s*)?Sales?\s*Price\s*\$?\s*([\d,]+\.?\d*)',
            r'Contract\s*Sales\s*Price\s*\$?\s*([\d,]+\.?\d*)',
            r'101\.?\s*Contract\s*sales\s*price\s*\$?\s*([\d,]+\.?\d*)',
            r'Purchase\s*Price\s*\$?\s*([\d,]+\.?\d*)',
        ]:
            m = re.search(pattern, all_text, re.IGNORECASE)
            if m:
                val = float(m.group(1).replace(',', ''))
                if val > 50000:
                    result['sale_price'] = val
                    break

        # Extract closing / settlement date
        for pattern in [
            r'Closing\s*Date[:\s]+(\d{1,2}/\d{1,2}/\d{2,4})',
            r'Settlement\s*Date[:\s]+(\d{1,2}/\d{1,2}/\d{2,4})',
            r'Disbursement\s*Date[:\s]+(\d{1,2}/\d{1,2}/\d{2,4})',
            r'Closing\s*Date[:\s]+(\w+\s+\d{1,2},?\s+\d{4})',
        ]:
            m = re.search(pattern, all_text, re.IGNORECASE)
            if m:
                result['closing_date'] = m.group(1).strip()
                break

        # Extract line items
        line_pattern = re.compile(r'^(.+?)\s+\$?([\d,]+\.\d{2})\s*$', re.MULTILINE)

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
            'constuc on draw', 'construction draw',
            'commission to', 'commission paid', 'commission',
            'better homes', 'keller williams', 'realty', 'remax', 're/max',
            'debt paydown', 'american express',
            'payoff of', 'loan payoff',
        ]

        for match in line_pattern.finditer(all_text):
            desc = match.group(1).strip()
            amount = float(match.group(2).replace(',', ''))
            if amount < 5 or amount > 50000:
                continue
            desc_lower = desc.lower()
            if any(kw in desc_lower for kw in exclude_keywords):
                continue
            if len(desc) < 3:
                continue
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

        # Deduplicate by amount
        seen_amounts = {}
        for item in result['line_items']:
            amt_key = f"{item['amount']:.2f}"
            if amt_key in seen_amounts:
                if len(item['description']) > len(seen_amounts[amt_key]['description']):
                    seen_amounts[amt_key] = item
            else:
                seen_amounts[amt_key] = item
        result['line_items'] = list(seen_amounts.values())

        # Clean junk lines
        cleaned = []
        for item in result['line_items']:
            desc = item['description']
            desc_lower = desc.lower().strip()
            if desc_lower.startswith('premium'):
                continue
            if desc.startswith('$ '):
                continue
            if desc_lower in ('poc', 'poc $') or (desc_lower.startswith('poc') and len(desc_lower) < 10):
                continue
            desc_cleaned = re.sub(r'^\d{2}', '', desc).strip()
            if desc_cleaned:
                item['description'] = desc_cleaned
            cleaned.append(item)
        result['line_items'] = cleaned

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
    rows.append(['S-Corp Distribution (K-1 ordinary income — SE tax does not apply)', '', ''])
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
# Auth
# ---------------------------------------------------------------------------
@app.before_request
def check_auth():
    """Block all routes except /login if password is set and user not authenticated."""
    if not APP_PASSWORD:
        return  # no password configured — open access
    if request.endpoint in ('login', 'static'):
        return  # always allow login page
    if not session.get('authenticated'):
        if request.is_json or request.path.startswith('/api/'):
            return jsonify({'error': 'Unauthorized'}), 401
        return render_template('login.html'), 401


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if APP_PASSWORD and not session.get('authenticated'):
            if request.is_json:
                return jsonify({'error': 'Unauthorized'}), 401
            return render_template('login.html'), 401
        return f(*args, **kwargs)
    return decorated

@app.route('/login', methods=['GET', 'POST'])
def login():
    if not APP_PASSWORD:
        return jsonify({'ok': True})  # no password set
    if request.method == 'POST':
        pwd = (request.json or {}).get('password', '') if request.is_json else request.form.get('password', '')
        if pwd == APP_PASSWORD:
            session['authenticated'] = True
            return jsonify({'ok': True}) if request.is_json else (__import__('flask').redirect('/'))
        return (jsonify({'error': 'Wrong password'}), 403) if request.is_json else (render_template('login.html', error='Wrong password'), 403)
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return __import__('flask').redirect('/login')


# ---------------------------------------------------------------------------
# Backup / restore
# ---------------------------------------------------------------------------
@app.route('/api/backup/download')
@login_required
def backup_download():
    data = load_data()
    backup = json.dumps(data, indent=2)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    resp = Response(backup, mimetype='application/json')
    resp.headers['Content-Disposition'] = f'attachment; filename=flip_tracker_backup_{ts}.json'
    return resp

@app.route('/api/backup/restore', methods=['POST'])
@login_required
def backup_restore():
    try:
        uploaded = request.json
        if not isinstance(uploaded, dict) or 'properties' not in uploaded:
            return jsonify({'error': 'Invalid backup file'}), 400
        global _memory_store
        _memory_store = None  # force reload after save
        save_data(uploaded)
        return jsonify({'ok': True, 'properties': len(uploaded.get('properties', []))})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.route('/')
@login_required
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


@app.route('/api/flips/<prop_id>/expense/<int:idx>', methods=['PUT'])
def update_expense(prop_id, idx):
    data = load_data()
    for prop in data['properties']:
        if prop.get('id') == prop_id:
            expenses = prop.get('expenses', [])
            if idx < 0 or idx >= len(expenses):
                return jsonify({'error': 'Index out of range'}), 400
            expenses[idx].update(request.json)
            save_data(data)
            metrics = calc_property_metrics(prop)
            return jsonify({**prop, 'metrics': metrics})
    return jsonify({'error': 'Not found'}), 404


@app.route('/api/flips/<prop_id>/expense/<int:idx>', methods=['DELETE'])
def delete_expense(prop_id, idx):
    data = load_data()
    for prop in data['properties']:
        if prop.get('id') == prop_id:
            expenses = prop.get('expenses', [])
            if idx < 0 or idx >= len(expenses):
                return jsonify({'error': 'Index out of range'}), 400
            removed = expenses.pop(idx)
            save_data(data)
            metrics = calc_property_metrics(prop)
            return jsonify({**prop, 'metrics': metrics, 'removed': removed})
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


@app.route('/api/flips/<prop_id>/closeout', methods=['POST'])
@login_required
def closeout_property(prop_id):
    """Lock a sold deal as closed, freeze a snapshot of final metrics."""
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404

    # Apply overhead allocation from the closeout wizard before recalculating metrics.
    # This saves the partner's decision about how much overhead to reimburse from this deal.
    body = request.get_json(silent=True) or {}
    overhead_allocation = body.get('overhead_allocation', 0) or 0
    prop['overhead_allocation'] = overhead_allocation

    metrics = calc_property_metrics(prop)
    partner_b_check = metrics['distributable_profit'] - metrics['partner_share']

    prop['status'] = 'closed'
    prop['closeout_date'] = datetime.now().strftime('%Y-%m-%d')
    prop['closeout_snapshot'] = {
        'purchase_price':              prop.get('purchase_price', 0),
        'sale_price':                  metrics['effective_sale'],
        'total_rehab':                 metrics['total_rehab'],
        'rehab_budget':                prop.get('rehab_budget', 0),
        'total_costs':                 metrics['total_costs'],
        'gross_profit':                metrics['gross_profit'],
        'cash_invested':               metrics['cash_invested'],
        'overhead_allocation':         overhead_allocation,
        'cash_in_deal':                metrics['cash_in_deal'],
        'total_draws':                 metrics['total_draws'],
        'distribution_base':           metrics['distribution_base'],
        'net_proceeds_at_close':       metrics['net_proceeds_at_close'],
        'distributable_profit':        metrics['distributable_profit'],
        'cash_invested_partner_check': metrics['partner_total'],
        'non_cash_partner_check':      round(partner_b_check, 2),
        'partner_split_pct':           prop.get('partner_split_pct', 50),
        'roi':                         round(metrics['roi'], 2),
        'days_held':                   metrics['days_held'],
        'purchase_date':               prop.get('purchase_date', ''),
        'closeout_date':               datetime.now().strftime('%Y-%m-%d'),
    }

    save_data(data)
    return jsonify({'success': True, 'snapshot': prop['closeout_snapshot']})


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
# Business Overhead routes
# ---------------------------------------------------------------------------
def _calc_overhead_totals(data):
    """Calculate overhead totals: auto-accrued monthly salary + manual one-off entries."""
    expenses = data.get('overhead_expenses', [])
    settings = data.get('overhead_settings', {}) or {}
    monthly_rate = settings.get('monthly_rate', 0) or 0
    start_date_str = settings.get('start_date') or ''

    monthly_accrued = 0.0
    months_elapsed = 0.0
    if monthly_rate > 0 and start_date_str:
        try:
            start_dt = datetime.strptime(start_date_str, '%Y-%m-%d')
            days = (datetime.now() - start_dt).days
            months_elapsed = max(days / 30.4375, 0)
            monthly_accrued = monthly_rate * months_elapsed
        except ValueError:
            pass

    manual_total = sum(e.get('amount', 0) for e in expenses)
    total_logged = monthly_accrued + manual_total
    total_allocated = sum(
        p.get('overhead_allocation', 0) or 0
        for p in data.get('properties', [])
        if p.get('status') == 'closed'
    )
    return {
        'expenses': expenses,
        'overhead_settings': settings,
        'monthly_rate': monthly_rate,
        'months_elapsed': round(months_elapsed, 1),
        'monthly_accrued': round(monthly_accrued, 2),
        'manual_total': round(manual_total, 2),
        'total_logged': round(total_logged, 2),
        'total_allocated': round(total_allocated, 2),
        'outstanding': round(max(total_logged - total_allocated, 0), 2),
    }


@app.route('/api/overhead', methods=['GET'])
@login_required
def get_overhead():
    data = load_data()
    return jsonify(_calc_overhead_totals(data))


@app.route('/api/overhead/settings', methods=['POST'])
@login_required
def update_overhead_settings():
    data = load_data()
    body = request.json or {}
    data['overhead_settings'] = {
        'monthly_rate': body.get('monthly_rate', 0) or 0,
        'start_date': body.get('start_date') or None,
    }
    save_data(data)
    return jsonify(_calc_overhead_totals(data))


@app.route('/api/overhead', methods=['POST'])
@login_required
def add_overhead_expense():
    data = load_data()
    expense = request.json or {}
    expense['id'] = f"oh-{datetime.now().strftime('%Y%m%d%H%M%S%f')}"
    data.setdefault('overhead_expenses', []).append(expense)
    save_data(data)
    return jsonify(_calc_overhead_totals(data))


@app.route('/api/overhead/<expense_id>', methods=['DELETE'])
@login_required
def delete_overhead_expense(expense_id):
    data = load_data()
    before = len(data.get('overhead_expenses', []))
    data['overhead_expenses'] = [
        e for e in data.get('overhead_expenses', []) if e.get('id') != expense_id
    ]
    if len(data['overhead_expenses']) == before:
        return jsonify({'error': 'Not found'}), 404
    save_data(data)
    return jsonify(_calc_overhead_totals(data))


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
        # Bump date_added to today so re-analyzed deal floats to #1 in pipeline
        data['prospects'][existing]['date_added'] = datetime.now().strftime('%Y-%m-%d')
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
        'acq_closing_cost': prospect.get('acq_closing_costs', 0),
        'purchase_settlement': 0,
        'emd': prospect.get('emd', 0),
        'appraisal_fee': prospect.get('appraisal_fee', 0),
        'down_payment': prospect.get('down_payment', 0),
        'commitment_fee': 0,
        'cash_invested': prospect.get('cash_invested', 0),
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


# ---------------------------------------------------------------------------
# Business Command Center
# ---------------------------------------------------------------------------

@app.route('/api/business/settings', methods=['GET', 'POST'])
def business_settings_route():
    data = load_data()
    if request.method == 'POST':
        data['business_settings'] = request.json
        save_data(data)
    return jsonify(data.get('business_settings', {'annual_profit_goal': 0, 'year': datetime.now().year}))


@app.route('/api/business/summary')
def business_summary():
    data = load_data()
    year = request.args.get('year', datetime.now().year, type=int)
    biz_settings = data.get('business_settings', {})
    annual_goal = biz_settings.get('annual_profit_goal', 0)
    prospect_settings = data.get('prospect_settings', _default_prospect_settings())

    props = data.get('properties', [])
    prospects = data.get('prospects', [])

    # Split properties into closed (sold) and active
    closed_deals = []
    active_deals = []
    for prop in props:
        m = calc_property_metrics(prop)
        pnl = calc_pnl(prop, m)
        if (prop.get('sale_price') or 0) > 0:
            closed_deals.append({'prop': prop, 'metrics': m, 'pnl': pnl})
        else:
            active_deals.append({'prop': prop, 'metrics': m, 'pnl': pnl})

    # Closed profit this year (use net_profit from P&L)
    year_closed = [d for d in closed_deals if (d['prop'].get('sale_date') or '').startswith(str(year))]
    closed_profit = sum(d['pnl']['net_profit'] for d in year_closed)
    all_time_profits = [d['pnl']['net_profit'] for d in closed_deals]
    avg_profit_per_deal = sum(all_time_profits) / len(all_time_profits) if all_time_profits else 0

    # Hold times (closed deals)
    hold_times = []
    for d in closed_deals:
        pd_str = d['prop'].get('purchase_date')
        sd_str = d['prop'].get('sale_date')
        if pd_str and sd_str:
            try:
                hold_times.append((datetime.strptime(sd_str, '%Y-%m-%d') - datetime.strptime(pd_str, '%Y-%m-%d')).days)
            except ValueError:
                pass
    avg_hold_days = round(sum(hold_times) / len(hold_times)) if hold_times else 0

    # Budget variance across all deals
    variances = []
    for d in closed_deals + active_deals:
        budget = d['prop'].get('rehab_budget', 0) or 0
        actual = sum(e.get('amount', 0) for e in d['prop'].get('expenses', []) if not e.get('is_credit'))
        if budget > 0:
            variances.append((actual - budget) / budget * 100)
    avg_budget_variance = round(sum(variances) / len(variances), 1) if variances else 0

    # Average ROI
    rois = [d['metrics']['roi'] for d in closed_deals if (d['metrics'].get('roi') or 0) != 0]
    avg_roi = round(sum(rois) / len(rois), 1) if rois else 0

    # Capital deployed in active deals (purchase + rehab spent so far)
    capital_deployed = 0
    for d in active_deals:
        purchase = d['prop'].get('purchase_price', 0) or 0
        spent = sum(e.get('amount', 0) for e in d['prop'].get('expenses', []) if not e.get('is_credit'))
        capital_deployed += purchase + spent

    # Pipeline profit from prospects under contract / offer sent
    pipeline_profit = 0
    for p in prospects:
        if p.get('stage') in ('under_contract', 'offer_sent'):
            m = calc_prospect_metrics(p, prospect_settings)
            pipeline_profit += max(0, m.get('gross_profit', 0))

    # --- Conversion funnel from prospects ---
    STAGE_ORDER = ['new_lead', 'analyzing', 'offer_sent', 'under_contract', 'passed', 'converted']

    def highest_stage(p):
        reached = set(h['stage'] for h in (p.get('stage_history') or []))
        reached.add(p.get('stage', 'new_lead'))
        for s in reversed(STAGE_ORDER):
            if s in reached:
                return s
        return 'new_lead'

    analyzed = len(prospects)
    offered = sum(1 for p in prospects if highest_stage(p) in ('offer_sent', 'under_contract', 'converted'))
    contracted = sum(1 for p in prospects if highest_stage(p) in ('under_contract', 'converted'))
    # Closed = actual sold properties + converted prospects that closed
    closed_count = len(year_closed)

    analyze_to_offer = (offered / analyzed) if analyzed > 0 else 0
    offer_to_contract = (contracted / offered) if offered > 0 else 0
    contract_to_close = (closed_count / contracted) if contracted > 0 else 0
    overall_close_rate = (closed_count / analyzed) if analyzed > 0 else 0

    # --- Projection ---
    today = datetime.now()
    year_end = datetime(year, 12, 31)
    days_remaining = max(0, (year_end - today).days)
    months_remaining = round(days_remaining / 30.44, 1)

    gap = max(0, annual_goal - closed_profit)
    deals_needed = round(gap / avg_profit_per_deal, 1) if avg_profit_per_deal > 0 else 0
    leads_needed = round(deals_needed / overall_close_rate) if overall_close_rate > 0 else 0
    leads_per_month_needed = round(leads_needed / months_remaining, 1) if months_remaining > 0 else 0

    days_elapsed = max(1, (today - datetime(year, 1, 1)).days)
    months_elapsed = max(1, days_elapsed / 30.44)
    current_monthly_pace = round(analyzed / months_elapsed, 1) if analyzed > 0 else 0
    on_pace = (current_monthly_pace >= leads_per_month_needed) if leads_per_month_needed > 0 else True

    # --- Capital flow calendar (next 12 months from active deals) ---
    capital_flow = {}
    for d in active_deals:
        close_date = d['prop'].get('estimated_sale_date', '')
        if close_date:
            month_key = close_date[:7]
            if month_key not in capital_flow:
                capital_flow[month_key] = {'month': month_key, 'properties': [], 'expected_proceeds': 0, 'projected_profit': 0}
            arv = d['prop'].get('arv', 0) or 0
            capital_flow[month_key]['properties'].append(d['prop'].get('address', ''))
            capital_flow[month_key]['expected_proceeds'] += arv
            capital_flow[month_key]['projected_profit'] += max(0, d['metrics'].get('gross_profit', 0))

    # Deal source breakdown
    source_counts = {}
    for p in prospects:
        src = (p.get('source') or 'Unknown').strip() or 'Unknown'
        source_counts[src] = source_counts.get(src, 0) + 1

    return jsonify({
        'year': year,
        'goal': {
            'annual_profit_goal': annual_goal,
            'closed_profit': round(closed_profit),
            'pipeline_profit': round(pipeline_profit),
            'gap': round(gap),
            'pct_complete': round(min(100, (closed_profit / annual_goal * 100)) if annual_goal > 0 else 0, 1),
            'on_pace': on_pace,
            'months_remaining': months_remaining,
        },
        'funnel': {
            'analyzed': analyzed,
            'offered': offered,
            'contracted': contracted,
            'closed': closed_count,
            'analyze_to_offer_rate': round(analyze_to_offer * 100, 1),
            'offer_to_contract_rate': round(offer_to_contract * 100, 1),
            'contract_to_close_rate': round(contract_to_close * 100, 1),
            'overall_close_rate': round(overall_close_rate * 100, 1),
        },
        'scorecard': {
            'avg_profit_per_deal': round(avg_profit_per_deal),
            'avg_hold_days': avg_hold_days,
            'avg_budget_variance_pct': avg_budget_variance,
            'avg_roi': avg_roi,
            'capital_deployed': round(capital_deployed),
            'total_closed_deals': len(closed_deals),
            'source_counts': source_counts,
        },
        'projection': {
            'avg_profit_per_deal': round(avg_profit_per_deal),
            'deals_needed': deals_needed,
            'leads_needed': leads_needed,
            'months_remaining': months_remaining,
            'leads_per_month_needed': leads_per_month_needed,
            'current_monthly_pace': current_monthly_pace,
            'on_pace': on_pace,
        },
        'capital_flow': sorted(capital_flow.values(), key=lambda x: x['month']),
        'active_deals': [{
            'address': d['prop'].get('address', ''),
            'city': d['prop'].get('city', ''),
            'capital_in': round((d['prop'].get('purchase_price', 0) or 0) + sum(e.get('amount', 0) for e in d['prop'].get('expenses', []) if not e.get('is_credit'))),
            'estimated_close': d['prop'].get('estimated_sale_date', ''),
            'arv': d['prop'].get('arv', 0) or 0,
            'projected_profit': round(max(0, d['metrics'].get('gross_profit', 0))),
        } for d in active_deals],
    })


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


@app.route('/api/export/lender-csv')
@login_required
def export_lender_csv():
    """Export closed deals as a lender-ready track record CSV."""
    data = load_data()
    closed = [p for p in data['properties'] if p.get('status') == 'closed' and p.get('closeout_snapshot')]

    output = io.StringIO()
    writer = csv.writer(output)

    writer.writerow(['INVESTMENT TRACK RECORD'])
    writer.writerow(['Generated', datetime.now().strftime('%Y-%m-%d')])
    writer.writerow([])
    writer.writerow(['Address', 'City', 'State', 'Purchase Date', 'Close Date', 'Days Held',
                     'Purchase Price', 'Sale Price', 'Total Rehab', 'Rehab Budget', 'Budget Variance %',
                     'Total Costs', 'Net Profit', 'ROI %', 'Capital Invested',
                     'Cash-Invested Partner Check', 'Non-Cash-Invested Partner Check', 'MOIC'])

    totals = {'profit': 0, 'capital': 0, 'cash_inv_check': 0, 'non_cash_check': 0, 'roi_sum': 0}
    for p in closed:
        snap = p.get('closeout_snapshot', {})
        budget = snap.get('rehab_budget', 0)
        actual = snap.get('total_rehab', 0)
        bvar = round((actual - budget) / budget * 100, 1) if budget > 0 else ''
        capital = snap.get('cash_invested', 0)
        ci_check = snap.get('cash_invested_partner_check', 0)
        moic = round(ci_check / capital, 2) if capital > 0 else ''
        writer.writerow([
            p.get('address', ''), p.get('city', ''), p.get('state', ''),
            snap.get('purchase_date', ''), snap.get('closeout_date', ''), snap.get('days_held', ''),
            snap.get('purchase_price', 0), snap.get('sale_price', 0),
            actual, budget, bvar,
            snap.get('total_costs', 0), snap.get('gross_profit', 0), snap.get('roi', 0),
            capital, ci_check, snap.get('non_cash_partner_check', 0), moic,
        ])
        totals['profit']        += snap.get('gross_profit', 0)
        totals['capital']       += capital
        totals['cash_inv_check'] += ci_check
        totals['non_cash_check'] += snap.get('non_cash_partner_check', 0)
        totals['roi_sum']       += snap.get('roi', 0)

    avg_roi  = round(totals['roi_sum'] / len(closed), 1) if closed else 0
    moic_avg = round(totals['cash_inv_check'] / totals['capital'], 2) if totals['capital'] > 0 else ''
    writer.writerow([])
    writer.writerow(['TOTALS', '', '', '', '', '',
                     '', '', '', '', '',
                     '', totals['profit'], f'{avg_roi}% avg',
                     totals['capital'], totals['cash_inv_check'], totals['non_cash_check'], f'{moic_avg}x avg'])
    writer.writerow([])
    writer.writerow(['SUMMARY STATS'])
    writer.writerow(['Total Deals Closed', len(closed)])
    writer.writerow(['Total Net Profit',   totals['profit']])
    writer.writerow(['Total Capital Deployed', totals['capital']])
    writer.writerow(['Average ROI %', avg_roi])
    writer.writerow(['Average MOIC',  moic_avg])

    return Response(
        output.getvalue(),
        mimetype='text/csv',
        headers={'Content-Disposition': f'attachment; filename=lender-track-record-{datetime.now().strftime("%Y%m%d")}.csv'}
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
    writer.writerow(['S-Corp K-1 ordinary income — SE tax does not apply', '', ''])

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
    data.append([P('S-Corp Distribution', cell_indent), P('K-1 ordinary income — SE tax does not apply', cell_sub), P('See CPA', cell_sub)])
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
    header = ['Property', 'Sale Price', 'COGS', 'Selling', 'Net Profit (K-1)']
    data = [header]
    totals = [0, 0, 0, 0]
    for pnl, prop in all_pnls:
        data.append([
            prop.get('address', ''),
            fmt(pnl['net_sale_proceeds']),
            fmt(pnl['total_cogs']),
            fmt(pnl['total_selling']),
            fmt(pnl['net_profit']),
        ])
        totals[0] += pnl['net_sale_proceeds']
        totals[1] += pnl['total_cogs']
        totals[2] += pnl['total_selling']
        totals[3] += pnl['net_profit']
    data.append(['TOTALS', fmt(totals[0]), fmt(totals[1]), fmt(totals[2]), fmt(totals[3])])

    table = Table(data, colWidths=[2.5 * inch, 1.1 * inch, 1.1 * inch, 1.0 * inch, 1.2 * inch])
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
    cd_type = request.form.get('type', 'purchase')  # 'purchase', 'sale', or 'lender_cashback'
    pdf_bytes = file.read()

    # Parse the PDF
    parsed = parse_closing_disclosure(pdf_bytes)

    # For sale CDs: filter out buyer-side items (you're the seller)
    if cd_type == 'sale' and parsed.get('line_items'):
        buyer_keywords = [
            'credit report', 'prepaid interest', 'processing fee',
            'lender', 'mortgage insurance', 'flood', 'tax service',
            'appraisal', 'impound', 'escrow',
            'homeowner', 'hazard',  # buyer's insurance escrow
            'mortgage city', 'mortgage state',  # buyer's mortgage stamps
            'title - icl', 'title - title commitment',
            'title - title update',
        ]
        seller_items = []
        for item in parsed['line_items']:
            desc_lower = item['description'].lower()
            is_buyer = any(kw in desc_lower for kw in buyer_keywords)
            if not is_buyer:
                seller_items.append(item)
        parsed['line_items'] = seller_items
        parsed['closing_costs_total'] = sum(i['amount'] for i in seller_items)

    # Store base64-encoded PDF + parsed data
    cd_data = {
        'upload_date': datetime.now().strftime('%Y-%m-%d'),
        'filename': file.filename,
        'pdf_base64': base64.b64encode(pdf_bytes).decode('utf-8'),
        'loan_amount': parsed.get('loan_amount', 0),
        'interest_rate': parsed.get('interest_rate', 0),
        'closing_costs_total': parsed.get('closing_costs_total', 0),
        'cash_to_close': parsed.get('cash_to_close', 0),
        'sale_price': parsed.get('sale_price', 0),
        'closing_date': parsed.get('closing_date', ''),
        'line_items': parsed.get('line_items', []),
        'raw_text': parsed.get('raw_text', ''),
    }

    key = f'closing_disclosure_{cd_type}'
    prop[key] = cd_data

    def parse_closing_date(raw):
        """Convert various date formats to YYYY-MM-DD."""
        for fmt in ['%m/%d/%Y', '%m/%d/%y', '%B %d, %Y', '%B %d %Y']:
            try:
                return datetime.strptime(raw, fmt).strftime('%Y-%m-%d')
            except ValueError:
                continue
        return None

    if cd_type == 'purchase':
        # Auto-populate purchase_settlement from cash_to_close if not set
        if cd_data['cash_to_close'] > 0 and not prop.get('purchase_settlement', 0):
            prop['purchase_settlement'] = cd_data['cash_to_close']
        # Auto-populate purchase_date from closing date if not set
        if cd_data['closing_date'] and not prop.get('purchase_date'):
            parsed_date = parse_closing_date(cd_data['closing_date'])
            if parsed_date:
                prop['purchase_date'] = parsed_date

    if cd_type == 'sale':
        # Auto-populate sale_price from CD if not set
        if cd_data['sale_price'] > 0 and not prop.get('sale_price', 0):
            prop['sale_price'] = cd_data['sale_price']
        # Auto-populate sale_date from closing date if not set
        if cd_data['closing_date'] and not prop.get('sale_date'):
            parsed_date = parse_closing_date(cd_data['closing_date'])
            if parsed_date:
                prop['sale_date'] = parsed_date

    if cd_type == 'lender_cashback':
        # Auto-populate lender_cashback amount from cash_to_close (money back to borrower)
        if cd_data['cash_to_close'] > 0:
            prop['lender_cashback'] = cd_data['cash_to_close']

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
        'sale_price': parsed.get('sale_price', 0),
        'closing_date': parsed.get('closing_date', ''),
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

    # Update line items and all user-editable header fields
    prop[key]['line_items'] = body.get('line_items', prop[key].get('line_items', []))
    prop[key]['closing_costs_total'] = sum(item.get('amount', 0) for item in prop[key]['line_items'])
    if 'interest_rate' in body:
        prop[key]['interest_rate'] = body['interest_rate']
    if 'cash_to_close' in body and body['cash_to_close'] is not None:
        prop[key]['cash_to_close'] = float(body['cash_to_close'])
        # Sync purchase_settlement for purchase CDs so cash_in_deal recalculates
        if cd_type == 'purchase':
            prop['purchase_settlement'] = float(body['cash_to_close'])
        # Sync lender_cashback for lender cashback CDs
        if cd_type == 'lender_cashback':
            prop['lender_cashback'] = float(body['cash_to_close'])
    if 'loan_amount' in body and body['loan_amount'] is not None:
        prop[key]['loan_amount'] = float(body['loan_amount'])

    save_data(data)
    return jsonify({'success': True})


@app.route('/api/flips/<prop_id>/closing-disclosure/<cd_type>', methods=['DELETE'])
def delete_closing_disclosure(prop_id, cd_type):
    """Remove an uploaded closing disclosure (e.g. erroneous purchase CD on a subject-to deal)."""
    if cd_type not in ('purchase', 'sale', 'lender_cashback'):
        return jsonify({'error': 'Invalid type — must be purchase, sale, or lender_cashback'}), 400
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404
    key = f'closing_disclosure_{cd_type}'
    if key in prop:
        del prop[key]
        # If this was a purchase CD that auto-populated purchase_settlement, leave
        # purchase_settlement as-is — the user will correct it manually if needed.
        # If this was a lender cashback CD, clear the lender_cashback field too.
        if cd_type == 'lender_cashback':
            prop['lender_cashback'] = 0
        save_data(data)
    return jsonify({'success': True})


@app.route('/api/flips/<prop_id>/closing-disclosure/reprocess', methods=['POST'])
def reprocess_closing_disclosure(prop_id):
    """Re-parse already-uploaded CDs and auto-fill sale_price, sale_date, purchase_date, purchase_settlement."""
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404

    def parse_closing_date(raw):
        for fmt in ['%m/%d/%Y', '%m/%d/%y', '%B %d, %Y', '%B %d %Y']:
            try:
                return datetime.strptime(raw, fmt).strftime('%Y-%m-%d')
            except ValueError:
                continue
        return None

    updates = {}

    # Re-parse purchase CD
    cd_purchase = prop.get('closing_disclosure_purchase', {})
    if cd_purchase.get('pdf_base64'):
        pdf_bytes = base64.b64decode(cd_purchase['pdf_base64'])
        parsed = parse_closing_disclosure(pdf_bytes)
        if parsed.get('cash_to_close', 0) > 0:
            cd_purchase['cash_to_close'] = parsed['cash_to_close']
            updates['cd_purchase_cash_to_close'] = parsed['cash_to_close']
            if not prop.get('purchase_settlement', 0):
                # Set purchase_settlement = CD cash_to_close.
                # Sub-fields (emd, appraisal, commitment) are kept as-is — they are
                # display-only breakdown rows in the Cash Invested card and show what
                # makes up the total. The 'Other Closing Costs' auto-calc fills the gap.
                # Do NOT clear sub-fields here — user needs them for the breakdown display.
                prop['purchase_settlement'] = parsed['cash_to_close']
                updates['purchase_settlement'] = parsed['cash_to_close']
            # Clear stale manual cash_invested so backend recalculates from CD/settlement
            prop['cash_invested'] = 0
            updates['cash_invested_recalculated'] = True
        if parsed.get('closing_date') and not prop.get('purchase_date'):
            d = parse_closing_date(parsed['closing_date'])
            if d:
                prop['purchase_date'] = d
                updates['purchase_date'] = d
        if parsed.get('sale_price', 0) > 0:
            cd_purchase['sale_price'] = parsed['sale_price']

    # Re-parse sale CD
    cd_sale = prop.get('closing_disclosure_sale', {})
    if cd_sale.get('pdf_base64'):
        pdf_bytes = base64.b64decode(cd_sale['pdf_base64'])
        parsed = parse_closing_disclosure(pdf_bytes)
        if parsed.get('sale_price', 0) > 0:
            cd_sale['sale_price'] = parsed['sale_price']
            if not prop.get('sale_price', 0):
                prop['sale_price'] = parsed['sale_price']
                updates['sale_price'] = parsed['sale_price']
        if parsed.get('cash_to_close', 0) > 0:
            cd_sale['cash_to_close'] = parsed['cash_to_close']
        if parsed.get('closing_date') and not prop.get('sale_date'):
            d = parse_closing_date(parsed['closing_date'])
            if d:
                prop['sale_date'] = d
                updates['sale_date'] = d

    save_data(data)
    metrics = calc_property_metrics(prop)
    return jsonify({'success': True, 'updates': updates, 'metrics': metrics})


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

# ---------------------------------------------------------------------------
# Vendor defaults — auto-category mapping
# ---------------------------------------------------------------------------
@app.route('/api/vendor-defaults')
@login_required
def get_vendor_defaults():
    """Return vendor→category defaults so the frontend can auto-fill."""
    return jsonify(VENDOR_CATEGORY_DEFAULTS)


@app.route('/api/admin/backfill-vendor-category', methods=['POST'])
@login_required
def backfill_vendor_category():
    """
    Update all existing expenses whose vendor matches a key in VENDOR_CATEGORY_DEFAULTS
    (case-insensitive) to the mapped category. Safe to call repeatedly — only changes
    expenses that don't already have the correct category.
    """
    data = load_data()
    updated = 0
    for prop in data.get('properties', []):
        for exp in prop.get('expenses', []):
            vendor_key = exp.get('vendor', '').lower().strip()
            mapped_cat = VENDOR_CATEGORY_DEFAULTS.get(vendor_key)
            if mapped_cat and exp.get('category') != mapped_cat:
                exp['category'] = mapped_cat
                updated += 1
    if updated > 0:
        save_data(data)
    return jsonify({'ok': True, 'updated': updated})


# ---------------------------------------------------------------------------
# Google Sheets sync — manual trigger route
# ---------------------------------------------------------------------------
@app.route('/api/sheets-status')
def sheets_status():
    configured = bool(os.environ.get('GOOGLE_SHEET_ID') and os.environ.get('GOOGLE_CREDENTIALS_JSON'))
    return jsonify({'configured': configured})


@app.route('/api/sync-sheets', methods=['POST'])
def trigger_sheets_sync():
    from sheets_sync import sync_to_sheets
    # Run synchronously so we can return the real result (fast enough < 15s)
    result = sync_to_sheets()
    if result and result.get('ok'):
        return jsonify({'status': f"✓ All 5 tabs synced — {result.get('synced_at', '')}", 'detail': result})
    elif result:
        failed = [k for k, v in result.get('tabs', {}).items() if v != 'ok']
        ok_tabs = [k for k, v in result.get('tabs', {}).items() if v == 'ok']
        return jsonify({'status': f"⚠ {len(ok_tabs)}/5 tabs synced. Failed: {', '.join(failed)}", 'detail': result})
    return jsonify({'status': 'Sync failed — check Railway logs', 'detail': {}})


# ---------------------------------------------------------------------------
# Background scheduler — daily Google Sheets sync at 6 AM UTC
# ---------------------------------------------------------------------------
def _start_sheets_scheduler():
    try:
        from apscheduler.schedulers.background import BackgroundScheduler
        from sheets_sync import sync_to_sheets
        scheduler = BackgroundScheduler(daemon=True)
        scheduler.add_job(
            sync_to_sheets,
            trigger='cron',
            hour=6, minute=0,
            id='daily_sheets_sync',
            replace_existing=True,
            misfire_grace_time=3600,
        )
        scheduler.start()
        print('[scheduler] Google Sheets daily sync scheduled — 06:00 UTC')
    except Exception as e:
        print(f'[scheduler] Could not start sheets scheduler: {e}')

# Only start if Google credentials are configured
if os.environ.get('GOOGLE_SHEET_ID') and os.environ.get('GOOGLE_CREDENTIALS_JSON'):
    _start_sheets_scheduler()


# ---------------------------------------------------------------------------
# Construction Project Management
# ---------------------------------------------------------------------------

WCP_SCHEMA = [
    ("PRE-CONSTRUCTION & DESIGN", 1, [
        "Permits", "Plans (Architectural, Engineering, Design)",
    ]),
    ("SITE PREPARATION", 2, [
        "Clearing & Grading", "Silt Fence & Erosion Control",
        "Interior Demolition", "Exterior Demolition", "Dumpsters & Trash Removal",
    ]),
    ("FOUNDATION WORK", 3, [
        "Excavation", "Underpinning & Footings", "Foundation & Concrete", "Waterproofing",
    ]),
    ("FRAMING", 4, [
        "Wood Framing", "Steel Framing", "Insulation", "Roofing & Flashing",
    ]),
    ("UTILITY ROUGH-INS", 5, [
        "Electrical Rough-In", "Plumbing Rough-in", "Mechanical Rough-in",
    ]),
    ("EXTERIOR FINISHES", 6, [
        "Siding & Trim", "Masonry & Stonework", "Windows", "Exterior Doors & Frames",
        "Gutters & Downspouts", "Porch, Patio, & Deck", "Exterior Painting",
    ]),
    ("INTERIOR WORK", 7, [
        "Drywall", "Interior Doors & Frames", "Interior Trim & Millwork",
        "Flooring", "Staircase", "Interior Painting",
    ]),
    ("UTILITY FINISHES", 8, [
        "Electrical Fixtures", "Plumbing Fixtures", "Mechanical Finishes",
    ]),
    ("INTERIOR FINISHES", 9, [
        "Tile & Backsplash", "Bathtubs & Showers", "Toilets", "Vanities",
        "Kitchen Cabinets", "Appliances", "Countertops", "Washer/Dryer",
    ]),
    ("EXTERIOR MISCELLANEOUS", 10, [
        "Landscaping", "Tree Removal", "Garage & Door", "Driveway",
        "Sidewalk", "Fence", "Powerwash", "Septic", "Water Well",
    ]),
    ("INTERIOR MISCELLANEOUS", 11, [
        "Fire Alarms", "Fire-Rated Doors", "Sprinklers", "Fireplace",
        "Chimney", "Final Cleaning", "Staging",
    ]),
    ("OTHER", 12, [
        "House numbers & Mailbox",
        "interior & exterior Railings, metal work and balconies",
        "Reglazing of tubs and showers",
    ]),
    ("CONTINGENCY", 13, ["Contingency"]),
]

_WCP_PHASE_LOOKUP = {ph.upper(): (ph, o) for ph, o, _ in WCP_SCHEMA}
_WCP_EXPENSE_LOOKUP = {}
for _wcp_ph, _wcp_ord, _wcp_exps in WCP_SCHEMA:
    for _wcp_exp in _wcp_exps:
        _WCP_EXPENSE_LOOKUP[_wcp_exp.lower().strip()] = (_wcp_ph, _wcp_ord, _wcp_exp)


# ---------------------------------------------------------------------------
# Construction critical-path blocking
# ---------------------------------------------------------------------------
# Threshold: a blocking phase must reach this completion % before the next opens.
# Lower (e.g. to 75) to allow parallel-track tolerance between adjacent phases.
CRITICAL_PATH_THRESHOLD = 100

# Inter-phase blocking chain: phase_order → predecessor phase_order.
# Forms a recursive chain — if a predecessor has no scope items on this property
# the check walks further back until it finds one that does (or reaches the start).
# Exterior phases (6, 10) are intentionally absent; exterior work runs in parallel
# with interior and does not gate anything.
PHASE_CRITICAL_PREDECESSORS = {
    4:  2,   # Framing requires Site Prep (clear site before structural)
    5:  4,   # Utility Rough-Ins require Framing (or Site Prep if no framing)
    7:  5,   # Interior Work requires Utility Rough-Ins (drywall after rough-in inspection)
    8:  7,   # Utility Finishes require Interior Work (fixtures after walls done)
    9:  7,   # Interior Finishes require Interior Work (cabinets after drywall/paint)
    11: 9,   # Interior Miscellaneous requires Interior Finishes (final clean is last)
}

# Intra-phase item blocking: item name → item that must complete first within the phase.
#
# Design note — item-level sub-ordering vs. named-predecessor map:
#   Sub-ordering (adding an item_order int to each scope item) is more flexible but
#   requires maintaining a parallel ordering structure alongside WCP_SCHEMA.
#   The named-predecessor map is simpler and explicit. The one real-world intra-phase
#   exception is Paint → Flooring inside INTERIOR WORK: WCP_SCHEMA lists Flooring
#   before Interior Painting (matching the lender's locked form layout), but
#   construction requires paint to cure before flooring goes down. Rather than
#   reordering the lender-locked schema, we declare the dependency here.
ITEM_CRITICAL_PREDECESSORS = {
    'Flooring': 'Interior Painting',
}

# Maps expense category → list of (phase_name, phase_order, wcp_item_name, fraction).
# Fraction splits a category across two WCP items where one line covers both rough-in
# and finish work (e.g. Electrical labor covers Rough-In AND Fixtures).
# "General Contractor" is intentionally mapped to Drywall as the closest catch-all
# for bulk labor invoices that span interior work. Windows & Doors splits 65/35.
EXPENSE_TO_WCP_SCOPE = {
    'Permits':            [('PRE-CONSTRUCTION & DESIGN', 1, 'Permits', 1.0)],
    'Dumpster':           [('SITE PREPARATION', 2, 'Dumpsters & Trash Removal', 1.0)],
    'Labor - General':    [('SITE PREPARATION', 2, 'Interior Demolition', 1.0)],
    'Repairs - Foundation': [('FOUNDATION WORK', 3, 'Foundation & Concrete', 1.0)],
    'Roofing':            [('FRAMING', 4, 'Roofing & Flashing', 1.0)],
    'Labor - Electrical': [
        ('UTILITY ROUGH-INS', 5, 'Electrical Rough-In', 0.60),
        ('UTILITY FINISHES', 8, 'Electrical Fixtures', 0.40),
    ],
    'Labor - Plumbing': [
        ('UTILITY ROUGH-INS', 5, 'Plumbing Rough-in', 0.60),
        ('UTILITY FINISHES', 8, 'Plumbing Fixtures', 0.40),
    ],
    'Labor - HVAC': [
        ('UTILITY ROUGH-INS', 5, 'Mechanical Rough-in', 0.60),
        ('UTILITY FINISHES', 8, 'Mechanical Finishes', 0.40),
    ],
    'Windows & Doors': [
        ('EXTERIOR FINISHES', 6, 'Windows', 0.65),
        ('EXTERIOR FINISHES', 6, 'Exterior Doors & Frames', 0.35),
    ],
    'Building Materials': [('INTERIOR WORK', 7, 'Drywall', 1.0)],
    'Paint':              [('INTERIOR WORK', 7, 'Interior Painting', 1.0)],
    'Flooring':           [('INTERIOR WORK', 7, 'Flooring', 1.0)],
    'Labor - Kitchen':    [('INTERIOR FINISHES', 9, 'Kitchen Cabinets', 1.0)],
    'Appliances':         [('INTERIOR FINISHES', 9, 'Appliances', 1.0)],
    'Landscaping':        [('EXTERIOR MISCELLANEOUS', 10, 'Landscaping', 1.0)],
    'Staging':            [('INTERIOR MISCELLANEOUS', 11, 'Staging', 1.0)],
}


def _is_phase_blocked(ph_order, phase_pct_map, phase_name_map, depth=0):
    """Walk the predecessor chain, skipping phases with no scope items on this property."""
    if depth > 12:
        return False, ''
    pred_order = PHASE_CRITICAL_PREDECESSORS.get(ph_order)
    if pred_order is None:
        return False, ''
    if pred_order not in phase_pct_map:
        return _is_phase_blocked(pred_order, phase_pct_map, phase_name_map, depth + 1)
    pred_pct = phase_pct_map[pred_order]
    if pred_pct < CRITICAL_PATH_THRESHOLD:
        pred_name = phase_name_map.get(pred_order, f'Phase {pred_order}')
        return True, f'Blocked: finish {pred_name.title()} first ({int(pred_pct)}% done)'
    return False, ''


def _compute_scope_blocking(prop):
    """
    Returns {item_id: {'blocked': bool, 'reason': str}} for all scope items.
    Computed on the fly — never stored. Combines:
      1. Inter-phase chain via PHASE_CRITICAL_PREDECESSORS (recursive predecessor walk)
      2. Intra-phase named-item blocking via ITEM_CRITICAL_PREDECESSORS
    """
    scope = prop.get('scope_items', [])
    if not scope:
        return {}

    phase_pct_map = {}
    phase_name_map = {}
    for _, ph_order, _ in WCP_SCHEMA:
        items_in_phase = [i for i in scope if i.get('phase_order') == ph_order]
        if not items_in_phase:
            continue
        phase_name_map[ph_order] = items_in_phase[0]['phase']
        total_budget = sum(i['budget'] for i in items_in_phase)
        if total_budget > 0:
            weighted = sum(i['budget'] * i['completion_pct'] / 100.0 for i in items_in_phase)
            phase_pct_map[ph_order] = round(weighted / total_budget * 100, 1)
        else:
            # No dollar data — use item count so $0 phases aren't auto-marked 100%
            done = sum(1 for i in items_in_phase if i['completion_pct'] >= 100)
            phase_pct_map[ph_order] = round(done / len(items_in_phase) * 100, 1)

    item_name_pct = {i['name']: i['completion_pct'] for i in scope}

    result = {}
    for item in scope:
        blocked, reason = _is_phase_blocked(item['phase_order'], phase_pct_map, phase_name_map)

        if not blocked:
            pred_item_name = ITEM_CRITICAL_PREDECESSORS.get(item['name'])
            if pred_item_name and pred_item_name in item_name_pct:
                pred_pct = item_name_pct[pred_item_name]
                if pred_pct < CRITICAL_PATH_THRESHOLD:
                    blocked = True
                    reason = f'Blocked: finish {pred_item_name} first ({int(pred_pct)}% done)'

        result[item['id']] = {'blocked': blocked, 'reason': reason}

    return result


def _parse_dollar_v(s):
    if not s:
        return 0.0
    try:
        return max(float(str(s).replace('$', '').replace(',', '').strip()), 0.0)
    except (ValueError, AttributeError):
        return 0.0


def _match_wcp_expense(text):
    if not text:
        return None
    # Some lender PDFs render the "ti" ligature as a stray capital E mid-word
    # (e.g. "Vanities" -> "VaniEes", "Demolition" -> "DemoliEon"). Restore it.
    # Only fires on a capital E flanked by lowercase letters, which normal
    # title-case text never produces, so legitimate names are untouched.
    text = re.sub(r'([a-z])E([a-z])', r'\1ti\2', text)
    t = text.strip().lower()
    if t in _WCP_EXPENSE_LOOKUP:
        return _WCP_EXPENSE_LOOKUP[t]
    for key, val in _WCP_EXPENSE_LOOKUP.items():
        if t.startswith(key) or key.startswith(t[:max(len(t)-3, 5)]):
            return val
    return None


def _make_scope_item(phase, phase_order, name, budget):
    return {
        'id': str(uuid.uuid4()),
        'phase': phase,
        'phase_order': phase_order,
        'name': name,
        'budget': round(float(budget), 2),
        'completion_pct': 0,
        'drawn_pct': 0,
        'notes': '',
        'photos': [],
        'last_updated': None,
        'updated_by': None,
    }


def parse_wcp_pdf(file_bytes):
    """Parse WCP Construction Budget PDF using pdfplumber table extraction."""
    import pdfplumber
    items = []
    seen = set()

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if not row:
                        continue
                    cells = [str(c or '').strip() for c in row]

                    for i, cell in enumerate(cells):
                        m = _match_wcp_expense(cell)
                        if m and m[2] not in seen:
                            ph, o, canonical = m
                            for j in range(i + 1, min(i + 6, len(cells))):
                                amt = _parse_dollar_v(cells[j])
                                if amt > 0:
                                    seen.add(canonical)
                                    items.append(_make_scope_item(ph, o, canonical, amt))
                                    break
                            break

    return sorted(items, key=lambda x: x['phase_order'])


def parse_wcp_xlsx(file_bytes):
    """Parse WCP Construction Budget XLSX."""
    import openpyxl
    items = []
    seen = set()

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    for row in ws.iter_rows(values_only=True):
        cells = [str(c or '').strip() for c in row]
        for i, cell in enumerate(cells):
            m = _match_wcp_expense(cell)
            if m and m[2] not in seen:
                ph, o, canonical = m
                for j in range(i + 1, min(i + 6, len(cells))):
                    amt = _parse_dollar_v(cells[j])
                    if amt > 0:
                        seen.add(canonical)
                        items.append(_make_scope_item(ph, o, canonical, amt))
                        break
                break

    return sorted(items, key=lambda x: x['phase_order'])


def calc_project_metrics(prop):
    scope = prop.get('scope_items', [])
    plan = prop.get('project_plan', {}) or {}
    if not scope:
        return {'has_project': False}

    total_budget = sum(i['budget'] for i in scope)
    weighted_done = sum(i['budget'] * i['completion_pct'] / 100.0 for i in scope)
    if total_budget:
        overall_pct = round(weighted_done / total_budget * 100, 1)
    else:
        done_count = sum(1 for i in scope if i['completion_pct'] == 100)
        overall_pct = round(done_count / len(scope) * 100, 1) if scope else 0

    start_str = plan.get('start_date')
    proj_days = int(plan.get('projected_days') or 0)
    daily_interest = float(plan.get('daily_interest') or 0)
    days_elapsed = days_over = 0
    expected_pct = 0.0
    proj_end = None

    if start_str:
        try:
            start_dt = datetime.strptime(start_str, '%Y-%m-%d')
            days_elapsed = (datetime.now() - start_dt).days
            if proj_days:
                proj_end = (start_dt + timedelta(days=proj_days)).strftime('%Y-%m-%d')
                days_over = max(days_elapsed - proj_days, 0)
                expected_pct = min(round(days_elapsed / proj_days * 100, 1), 100)
        except ValueError:
            pass

    pace_delta = round(overall_pct - expected_pct, 1)
    interest_overrun = round(days_over * daily_interest, 2)

    phase_map = defaultdict(lambda: {'budget': 0.0, 'done_budget': 0.0, 'items': 0, 'complete': 0})
    for item in scope:
        ph = item['phase']
        phase_map[ph]['budget'] += item['budget']
        phase_map[ph]['done_budget'] += item['budget'] * item['completion_pct'] / 100.0
        phase_map[ph]['items'] += 1
        if item['completion_pct'] == 100:
            phase_map[ph]['complete'] += 1

    phases = []
    for ph, o, _ in WCP_SCHEMA:
        if ph in phase_map:
            d = phase_map[ph]
            pct = round(d['done_budget'] / d['budget'] * 100, 1) if d['budget'] else round(d['complete'] / d['items'] * 100, 1)
            phases.append({'phase': ph, 'phase_order': o, 'budget': round(d['budget'], 2),
                           'done_budget': round(d['done_budget'], 2), 'pct': pct,
                           'items': d['items'], 'complete': d['complete']})

    draws = plan.get('draws', [])
    total_drawn = sum(float(d.get('amount_received') or d.get('total_requested', 0))
                      for d in draws if d.get('status') == 'received')

    last_inspection = max((i['last_updated'] for i in scope if i.get('last_updated')), default=None)

    return {
        'has_project': True,
        'total_budget': round(total_budget, 2),
        'overall_pct': overall_pct,
        'drawable_now': round(weighted_done, 2),
        'days_elapsed': days_elapsed,
        'projected_days': proj_days,
        'proj_end': proj_end,
        'days_over': days_over,
        'interest_overrun': interest_overrun,
        'daily_interest': daily_interest,
        'expected_pct': expected_pct,
        'pace_delta': pace_delta,
        'phases': sorted(phases, key=lambda x: x['phase_order']),
        'draws': draws,
        'total_drawn': round(total_drawn, 2),
        'last_inspection': last_inspection,
        'scope_count': len(scope),
    }


def _ensure_photos_dir(prop_id, item_id):
    path = os.path.join(PHOTOS_DIR, prop_id, item_id)
    try:
        os.makedirs(path, exist_ok=True)
    except OSError:
        path = os.path.join(tempfile.gettempdir(), 'flip_photos', prop_id, item_id)
        os.makedirs(path, exist_ok=True)
    return path


def _session_photos_dir(prop_id):
    return _ensure_photos_dir(prop_id, 'session')


def _send_inspection_email(prop, changes, photos, site_notes=''):
    """Send a Postmark inspection-report email with photos grouped by category."""
    if not POSTMARK_SERVER_TOKEN:
        return
    addr = f"{prop.get('address', '')} {prop.get('city', '')}".strip()
    today = datetime.now().strftime('%B %d, %Y')
    subject = f"Inspection: {addr} | {today}"

    html = (
        '<div style="font-family:Arial,sans-serif;max-width:640px;margin:0 auto;">'
        f'<h2 style="color:#1a2740;border-bottom:2px solid #1a2740;padding-bottom:8px;">Weekly Inspection Report</h2>'
        f'<p style="color:#6b7280;font-size:15px;margin-bottom:20px;">{addr}&nbsp;&nbsp;|&nbsp;&nbsp;{today}</p>'
    )

    if site_notes:
        html += (
            f'<div style="background:#f0f9ff;border:1px solid #bae6fd;border-radius:8px;padding:12px 16px;margin-bottom:20px;">'
            f'<div style="font-size:11px;font-weight:700;color:#0369a1;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">Site Notes</div>'
            f'<div style="font-size:14px;color:#0f172a;">{site_notes}</div>'
            f'</div>'
        )

    # Group photos by category bucket, preserving category order
    GROUP_ORDER = ['Interior', 'Mechanical', 'Exterior', 'Structural', 'Site', 'Other']
    cat_photos = defaultdict(list)  # catId -> [photo]
    uncategorized = []
    for ph in photos:
        cat = ph.get('category', '')
        if cat:
            cat_photos[cat].append(ph)
        else:
            uncategorized.append(ph)

    # Build group buckets: {group: [(catId, label, icon, [photos])]}
    group_buckets = defaultdict(list)
    seen_cats = set()
    for cat_def in INSPECTION_CATEGORIES:
        cid = cat_def['id']
        if cid in cat_photos:
            group_buckets[cat_def['group']].append((cid, cat_def['label'], cat_def['icon'], cat_photos[cid]))
            seen_cats.add(cid)
    if uncategorized:
        group_buckets['Other'].append(('other', 'Other', '📸', uncategorized))

    attachments = []
    attach_idx = 0

    if group_buckets:
        html += f'<h3 style="color:#374151;margin-bottom:16px;">Site Photos ({len(photos)})</h3>'
        for group in GROUP_ORDER:
            entries = group_buckets.get(group, [])
            if not entries:
                continue
            html += (
                f'<div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.6px;'
                f'color:#6b7280;margin:16px 0 8px;">{group}</div>'
            )
            for cid, label, icon, cat_ph_list in entries:
                html += (
                    f'<div style="margin-bottom:14px;">'
                    f'<div style="font-size:13px;font-weight:700;color:#374151;margin-bottom:8px;">'
                    f'{icon} {label} ({len(cat_ph_list)})</div>'
                    f'<div style="display:flex;flex-wrap:wrap;gap:6px;">'
                )
                for ph in cat_ph_list[:8]:
                    try:
                        with open(ph['path'], 'rb') as fh:
                            b64_data = base64.b64encode(fh.read()).decode()
                        ext = os.path.splitext(ph['filename'])[1].lower().lstrip('.')
                        ctype = 'image/jpeg' if ext in ('jpg', 'jpeg', 'heic') else f'image/{ext}'
                        cid_str = f'ph{attach_idx}'
                        attachments.append({
                            'Name': ph['filename'],
                            'Content': b64_data,
                            'ContentType': ctype,
                            'ContentID': f'cid:{cid_str}',
                        })
                        note = ph.get('note', '').strip()
                        html += f'<div style="flex-shrink:0;"><img src="cid:{cid_str}" style="width:180px;height:150px;object-fit:cover;border-radius:6px;border:1px solid #e5e7eb;display:block;">'
                        if note:
                            html += f'<p style="font-size:11px;color:#374151;margin:3px 0 0;max-width:180px;font-style:italic;">{note}</p>'
                        html += '</div>'
                        attach_idx += 1
                    except Exception as ex:
                        print(f'[postmark] photo read error: {ex}')
                if len(cat_ph_list) > 8:
                    html += f'<p style="color:#6b7280;font-size:12px;align-self:center;">+{len(cat_ph_list)-8} more</p>'
                html += '</div></div>'
    else:
        html += '<p style="color:#6b7280;margin-bottom:20px;">No photos submitted this visit.</p>'

    # WCP item changes (only shown if old inspector flow was used)
    if changes:
        html += '<h3 style="color:#374151;margin:20px 0 10px;">WCP Progress Updates</h3>'
        html += (
            '<table style="width:100%;border-collapse:collapse;margin-bottom:24px;font-size:14px;">'
            '<tr style="background:#f3f4f6;">'
            '<th style="padding:8px 12px;text-align:left;border:1px solid #e5e7eb;">Item</th>'
            '<th style="padding:8px 12px;text-align:center;border:1px solid #e5e7eb;width:60px;">Before</th>'
            '<th style="padding:8px 12px;text-align:center;border:1px solid #e5e7eb;width:60px;">After</th>'
            '</tr>'
        )
        for c in changes:
            color = '#16a34a' if c['after'] > c['before'] else '#374151'
            html += (
                f'<tr>'
                f'<td style="padding:8px 12px;border:1px solid #e5e7eb;">{c["name"]}</td>'
                f'<td style="padding:8px 12px;text-align:center;border:1px solid #e5e7eb;">{c["before"]}%</td>'
                f'<td style="padding:8px 12px;text-align:center;border:1px solid #e5e7eb;font-weight:bold;color:{color};">{c["after"]}%</td>'
                f'</tr>'
            )
        html += '</table>'

    html += '</div>'

    payload = json.dumps({
        'From': POSTMARK_FROM_EMAIL,
        'To': INSPECTION_NOTIFY_EMAIL,
        'Subject': subject,
        'HtmlBody': html,
        'Attachments': attachments,
        'MessageStream': 'outbound',
    }).encode()

    req = _urllib_req.Request(
        'https://api.postmarkapp.com/email',
        data=payload,
        headers={
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'X-Postmark-Server-Token': POSTMARK_SERVER_TOKEN,
        },
        method='POST',
    )
    try:
        _urllib_req.urlopen(req, timeout=15)
        print(f'[postmark] inspection email sent: {addr}')
    except Exception as e:
        print(f'[postmark] email failed: {e}')


def _get_inspector_token(data):
    import secrets
    token = (data.get('settings') or {}).get('inspector_token')
    if not token:
        token = secrets.token_urlsafe(16)
        data.setdefault('settings', {})['inspector_token'] = token
        save_data(data)
    return token


def _check_inspector_token():
    token = request.args.get('token') or request.headers.get('X-Inspector-Token')
    if not token:
        return False
    data = load_data()
    return token == (data.get('settings') or {}).get('inspector_token')


@app.route('/inspect')
def inspector_app():
    if not _check_inspector_token():
        return '<h2 style="font-family:sans-serif;padding:40px;color:#c00">Invalid or missing access token.</h2>', 403
    return render_template('inspector.html')


@app.route('/draw-package/<prop_id>')
def draw_package(prop_id):
    if not _check_inspector_token():
        return '<h2 style="font-family:sans-serif;padding:40px;color:#c00">Invalid or missing access token.</h2>', 403
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return '<h2 style="font-family:sans-serif;padding:40px;">Property not found.</h2>', 404

    addr = f"{prop.get('address', '')} {prop.get('city', '')} {prop.get('state', '')}".strip()
    today = datetime.now().strftime('%B %d, %Y')
    plan = prop.get('project_plan', {}) or {}
    scope = prop.get('scope_items', [])
    prior_draws = plan.get('draws', [])
    draw_number = len(prior_draws) + 1

    # Per-phase draw eligibility (completion - already-drawn)
    phase_buckets = {ph_order: {'phase': ph, 'phase_order': ph_order, 'items': [], 'subtotal': 0.0}
                     for ph, ph_order, _ in WCP_SCHEMA}
    total_eligible = 0.0
    draw_lines = []  # for copy-text

    for item in scope:
        drawn_pct = item.get('drawn_pct', 0) or 0
        eligible_pct = item['completion_pct'] - drawn_pct
        if eligible_pct <= 0:
            continue
        eligible_amt = round(item['budget'] * eligible_pct / 100.0, 2)
        if eligible_amt <= 0:
            continue
        pg = phase_buckets[item['phase_order']]
        pg['items'].append({
            'name': item['name'],
            'budget': item['budget'],
            'completion_pct': item['completion_pct'],
            'drawn_pct': drawn_pct,
            'eligible_pct': eligible_pct,
            'eligible_amt': eligible_amt,
        })
        pg['subtotal'] = round(pg['subtotal'] + eligible_amt, 2)
        total_eligible = round(total_eligible + eligible_amt, 2)

    active_phases = [phase_buckets[ph_order] for _, ph_order, _ in WCP_SCHEMA
                     if phase_buckets[ph_order]['items']]

    # Build plain-text draw summary for copy-to-clipboard
    text_lines = [
        f'DRAW REQUEST - {addr}',
        f'Date: {today}',
        f'Draw #{draw_number}',
        '',
    ]
    for pg in active_phases:
        text_lines.append(pg['phase'])
        for it in pg['items']:
            text_lines.append(f"  {it['name']}: {it['completion_pct']}% complete, ${it['eligible_amt']:,.2f}")
        text_lines.append(f"  Subtotal: ${pg['subtotal']:,.2f}")
        text_lines.append('')
    text_lines.append(f"TOTAL DRAW REQUEST: ${total_eligible:,.2f}")
    draw_text = '\n'.join(text_lines)

    # Collect inspection photos (most recent first), rebuild URLs with current token
    token = request.args.get('token') or request.headers.get('X-Inspector-Token', '')
    inspection_history = []
    for insp in reversed(plan.get('inspections', [])):
        photos = insp.get('photos', [])
        if photos:
            rebuilt = [{'filename': p['filename'],
                        'url': f'/photos/{prop_id}/session/{p["filename"]}?token={token}'}
                       for p in photos]
            inspection_history.append({'date': insp['date'], 'photos': rebuilt})

    return render_template('draw_package.html',
        addr=addr,
        today=today,
        draw_number=draw_number,
        total_eligible=total_eligible,
        active_phases=active_phases,
        inspection_history=inspection_history,
        draw_text=draw_text,
        token=token,
        prop_id=prop_id,
    )


@app.route('/photos/<prop_id>/<item_id>/<filename>')
def serve_photo(prop_id, item_id, filename):
    if not (check_auth() or _check_inspector_token()):
        return '', 403
    path = os.path.join(PHOTOS_DIR, prop_id, item_id, filename)
    if not os.path.exists(path):
        return '', 404
    return send_file(path)


@app.route('/api/settings/inspector-token', methods=['GET'])
@login_required
def get_inspector_token_route():
    data = load_data()
    token = _get_inspector_token(data)
    base = request.host_url.rstrip('/')
    return jsonify({'token': token, 'url': f'{base}/inspect?token={token}'})


@app.route('/api/flips/<prop_id>/scope/import', methods=['POST'])
@login_required
def import_scope(prop_id):
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    f = request.files['file']
    fname = (f.filename or '').lower()
    file_bytes = f.read()

    try:
        if fname.endswith('.pdf'):
            new_items = parse_wcp_pdf(file_bytes)
        elif fname.endswith(('.xlsx', '.xls')):
            new_items = parse_wcp_xlsx(file_bytes)
        else:
            return jsonify({'error': 'Upload a PDF or XLSX file'}), 400
    except Exception as e:
        return jsonify({'error': f'Parse error: {e}'}), 400

    if not new_items:
        return jsonify({'error': 'No line items found. Make sure this is a WCP Construction Budget document.'}), 400

    existing = {item['name']: item for item in prop.get('scope_items', [])}
    merged = []
    for item in new_items:
        if item['name'] in existing:
            existing[item['name']]['budget'] = item['budget']
            merged.append(existing[item['name']])
        else:
            merged.append(item)

    prop['scope_items'] = merged
    save_data(data)
    return jsonify({'ok': True, 'count': len(merged), 'scope_items': merged,
                    'metrics': calc_project_metrics(prop)})


@app.route('/api/flips/<prop_id>/scope/bootstrap', methods=['POST'])
@login_required
def bootstrap_scope_from_expenses(prop_id):
    """Create WCP scope items inferred from the property's existing expense log.
    Only creates items for categories where expenses exist; skips any WCP item name
    that already has a scope item (so it never clobbers a real WCP PDF import).
    Budget is set to actual dollars spent. Completion starts at 0 — inspector confirms.
    """
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404

    # Tally expense totals by category
    expense_totals = defaultdict(float)
    for exp in prop.get('expenses', []):
        cat = exp.get('category', '')
        amt = float(exp.get('amount', 0) or 0)
        expense_totals[cat] += amt

    # Map expense totals to WCP item budgets
    wcp_budgets = defaultdict(float)
    for cat, total in expense_totals.items():
        for _, _, item_name, fraction in (EXPENSE_TO_WCP_SCOPE.get(cat) or []):
            wcp_budgets[item_name] = round(wcp_budgets[item_name] + total * fraction, 2)

    # Create ALL WCP items so the inspector has the full checklist.
    # Budget comes from expense mapping where available; $0 otherwise.
    existing_names = {item['name'] for item in prop.get('scope_items', [])}
    new_items = []
    for phase_name, phase_order, item_names in WCP_SCHEMA:
        for item_name in item_names:
            if item_name in existing_names:
                continue
            budget = round(wcp_budgets.get(item_name, 0.0), 2)
            si = _make_scope_item(phase_name, phase_order, item_name, budget)
            if budget > 0:
                si['notes'] = 'Budget estimated from expense log. Confirm completion on site.'
            new_items.append(si)

    prop['scope_items'] = prop.get('scope_items', []) + new_items
    save_data(data)
    return jsonify({'ok': True, 'count': len(new_items),
                    'scope_items': prop['scope_items'],
                    'metrics': calc_project_metrics(prop)})


@app.route('/api/flips/<prop_id>/scope', methods=['GET'])
@login_required
def get_scope(prop_id):
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404
    return jsonify({'scope_items': prop.get('scope_items', []),
                    'project_plan': prop.get('project_plan', {}),
                    'metrics': calc_project_metrics(prop)})


@app.route('/api/flips/<prop_id>/scope/<item_id>', methods=['PUT'])
@login_required
def update_scope_item(prop_id, item_id):
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404
    item = next((i for i in prop.get('scope_items', []) if i['id'] == item_id), None)
    if not item:
        return jsonify({'error': 'Item not found'}), 404
    body = request.get_json(silent=True) or {}
    if 'completion_pct' in body:
        item['completion_pct'] = int(body['completion_pct'])
    if 'notes' in body:
        item['notes'] = body['notes']
    item['last_updated'] = datetime.now().strftime('%Y-%m-%d')
    item['updated_by'] = 'admin'
    save_data(data)
    return jsonify({'ok': True, 'metrics': calc_project_metrics(prop)})


@app.route('/api/flips/<prop_id>/scope/bulk-update', methods=['POST'])
@login_required
def bulk_update_scope(prop_id):
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404
    body = request.get_json(silent=True) or {}
    updates = body.get('updates', [])
    insp_date = body.get('inspection_date')
    today = datetime.now().strftime('%Y-%m-%d')
    item_by_id = {i['id']: i for i in prop.get('scope_items', [])}
    updated = 0
    for upd in updates:
        item = item_by_id.get(upd.get('item_id'))
        if item and 'completion_pct' in upd:
            item['completion_pct'] = int(upd['completion_pct'])
            item['last_updated'] = today
            item['updated_by'] = 'admin'
            updated += 1
    if insp_date:
        plan = prop.get('project_plan', {}) or {}
        for insp in plan.get('inspections', []):
            if insp.get('date') == insp_date:
                insp['reviewed'] = True
    save_data(data)
    return jsonify({'ok': True, 'updated': updated, 'metrics': calc_project_metrics(prop)})


@app.route('/api/flips/<prop_id>/project', methods=['POST'])
@login_required
def set_project_plan(prop_id):
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404
    body = request.get_json(silent=True) or {}
    plan = prop.get('project_plan', {}) or {}
    for field in ('start_date', 'contractor', 'notes'):
        if field in body:
            plan[field] = body[field]
    if 'projected_days' in body:
        plan['projected_days'] = int(body.get('projected_days') or 0)
    if 'daily_interest' in body:
        plan['daily_interest'] = float(body.get('daily_interest') or 0)
    prop['project_plan'] = plan
    save_data(data)
    return jsonify({'ok': True, 'metrics': calc_project_metrics(prop)})


@app.route('/api/flips/<prop_id>/draws/summary', methods=['GET'])
@login_required
def get_draw_summary(prop_id):
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404

    scope = prop.get('scope_items', [])
    draws = (prop.get('project_plan') or {}).get('draws', [])
    draw_number = len(draws) + 1
    addr = f"{prop.get('address', '')} {prop.get('city', '')} {prop.get('state', '')}".strip()
    today = datetime.now().strftime('%B %d, %Y')

    # Build per-phase eligible amounts: completion_pct minus already-drawn pct
    phase_buckets = {ph_order: {'phase': ph, 'phase_order': ph_order, 'items': [], 'subtotal': 0.0}
                     for ph, ph_order, _ in WCP_SCHEMA}
    total_eligible = 0.0
    draw_items_snapshot = []

    for item in scope:
        drawn_pct = item.get('drawn_pct', 0) or 0
        eligible_pct = item['completion_pct'] - drawn_pct
        if eligible_pct <= 0:
            continue
        eligible_amt = round(item['budget'] * eligible_pct / 100.0, 2)
        if eligible_amt <= 0:
            continue
        pg = phase_buckets[item['phase_order']]
        pg['items'].append({
            'id': item['id'],
            'name': item['name'],
            'budget': item['budget'],
            'completion_pct': item['completion_pct'],
            'drawn_pct': drawn_pct,
            'eligible_pct': eligible_pct,
            'eligible_amt': eligible_amt,
        })
        pg['subtotal'] = round(pg['subtotal'] + eligible_amt, 2)
        total_eligible = round(total_eligible + eligible_amt, 2)
        draw_items_snapshot.append({
            'item_id': item['id'],
            'name': item['name'],
            'phase': item['phase'],
            'eligible_pct': eligible_pct,
            'amount': eligible_amt,
        })

    # Preserve WCP phase order in output
    active_phases = [phase_buckets[ph_order] for _, ph_order, _ in WCP_SCHEMA
                     if phase_buckets[ph_order]['items']]

    if not active_phases:
        return jsonify({
            'ok': True, 'total_eligible': 0.0, 'phases': [],
            'draw_items': [], 'draw_number': draw_number,
            'formatted_text': 'No drawable amounts at this time.\nAll items are either at 0% or already drawn.',
        })

    # Plain-text summary formatted for copy/paste to lender
    lines = [
        f'DRAW REQUEST — {addr}',
        f'Date: {today}',
        f'Draw #{draw_number}',
        '',
    ]
    for pg in active_phases:
        budget_total = sum(i['budget'] for i in pg['items'])
        lines.append(f"{pg['phase']:<44} ${budget_total:>10,.2f} budget")
        for it in pg['items']:
            pct_str = f"{it['completion_pct']}% complete"
            lines.append(f"  {it['name']:<42} {pct_str:<18} ${it['eligible_amt']:>10,.2f}")
        lines.append(f"  {'Subtotal:':<60} ${pg['subtotal']:>10,.2f}")
        lines.append('')
    lines.append(f"{'TOTAL DRAW REQUEST:':<62} ${total_eligible:>10,.2f}")

    return jsonify({
        'ok': True,
        'draw_number': draw_number,
        'address': addr,
        'date': today,
        'phases': active_phases,
        'draw_items': draw_items_snapshot,
        'total_eligible': total_eligible,
        'formatted_text': '\n'.join(lines),
    })


@app.route('/api/flips/<prop_id>/draws', methods=['POST'])
@login_required
def add_draw_request(prop_id):
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404
    body = request.get_json(silent=True) or {}
    plan = prop.get('project_plan', {}) or {}
    draws = plan.get('draws', [])
    draw = {
        'id': str(uuid.uuid4()),
        'draw_number': len(draws) + 1,
        'requested_date': body.get('requested_date', datetime.now().strftime('%Y-%m-%d')),
        'total_requested': float(body.get('total_requested') or 0),
        'items': body.get('items', []),
        'status': 'pending',
        'amount_received': 0.0,
        'received_date': None,
        'notes': body.get('notes', ''),
    }
    draws.append(draw)
    plan['draws'] = draws
    prop['project_plan'] = plan
    save_data(data)
    return jsonify({'ok': True, 'draw': draw, 'metrics': calc_project_metrics(prop)})


@app.route('/api/flips/<prop_id>/draws/<draw_id>', methods=['PUT'])
@login_required
def update_draw(prop_id, draw_id):
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404
    plan = prop.get('project_plan', {}) or {}
    draw = next((d for d in plan.get('draws', []) if d['id'] == draw_id), None)
    if not draw:
        return jsonify({'error': 'Draw not found'}), 404
    body = request.get_json(silent=True) or {}
    prev_status = draw.get('status')
    for f in ('status', 'received_date', 'notes'):
        if f in body:
            draw[f] = body[f]
    if 'amount_received' in body:
        draw['amount_received'] = float(body['amount_received'] or 0)
    # When a draw transitions to received, stamp drawn_pct on included scope items
    # so the next draw summary doesn't double-count them.
    if body.get('status') == 'received' and prev_status != 'received':
        scope = prop.get('scope_items', [])
        item_by_id = {i['id']: i for i in scope}
        for di in draw.get('items', []):
            si = item_by_id.get(di.get('item_id'))
            if si:
                si['drawn_pct'] = min(
                    (si.get('drawn_pct', 0) or 0) + (di.get('eligible_pct', 0) or 0), 100
                )
    prop['project_plan'] = plan
    save_data(data)
    return jsonify({'ok': True, 'metrics': calc_project_metrics(prop)})


@app.route('/api/inspect/properties', methods=['GET'])
def inspect_properties():
    if not _check_inspector_token():
        return jsonify({'error': 'Unauthorized'}), 403
    data = load_data()
    result = []
    for prop in data.get('properties', []):
        if prop.get('status') == 'closed':
            continue
        scope_items = prop.get('scope_items', [])
        phase_map = defaultdict(list)
        for item in scope_items:
            phase_map[item.get('phase', 'Other')].append(item)
        phases = []
        for phase_name, items in phase_map.items():
            total_budget = sum(i.get('budget') or 0 for i in items)
            if total_budget > 0:
                weighted_pct = sum((i.get('budget') or 0) * (i.get('completion_pct') or 0) for i in items) / total_budget
            else:
                weighted_pct = sum(i.get('completion_pct') or 0 for i in items) / len(items)
            phases.append({'phase': phase_name, 'pct': round(weighted_pct), 'item_count': len(items)})

        result.append({
            'id': prop['id'],
            'address': prop.get('address', 'Unknown'),
            'scope_items': scope_items,
            'project_plan': prop.get('project_plan', {}),
            'metrics': calc_project_metrics(prop),
            'blocking': _compute_scope_blocking(prop),
            'phases': phases,
        })
    return jsonify({'properties': result})


@app.route('/api/inspect/session-photo/<prop_id>', methods=['POST'])
def upload_session_photo(prop_id):
    """Upload a general site photo not tied to a specific scope item."""
    if not _check_inspector_token():
        return jsonify({'error': 'Unauthorized'}), 403
    if 'photo' not in request.files:
        return jsonify({'error': 'No photo'}), 400
    f = request.files['photo']
    ext = os.path.splitext(f.filename or 'photo.jpg')[1].lower() or '.jpg'
    if ext not in ('.jpg', '.jpeg', '.png', '.heic', '.webp'):
        return jsonify({'error': 'Unsupported image type'}), 400
    photo_dir = _session_photos_dir(prop_id)
    filename = f'{datetime.now().strftime("%Y%m%d_%H%M%S")}_{uuid.uuid4().hex[:6]}{ext}'
    filepath = os.path.join(photo_dir, filename)
    f.save(filepath)
    token = request.args.get('token') or request.headers.get('X-Inspector-Token', '')
    photo_url = f'/photos/{prop_id}/session/{filename}?token={token}'
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    category = request.args.get('category') or request.form.get('category', '')
    if prop is not None:
        prop.setdefault('pending_photos', []).append({
            'filename': filename,
            'path': filepath,
            'url': photo_url,
            'uploaded': datetime.now().strftime('%Y-%m-%d %H:%M'),
            'category': category,
        })
        save_data(data)
    return jsonify({'ok': True, 'url': photo_url})


@app.route('/api/inspect/report', methods=['POST'])
def submit_inspection():
    if not _check_inspector_token():
        return jsonify({'error': 'Unauthorized'}), 403
    body = request.get_json(silent=True) or {}
    prop_id = body.get('prop_id')
    updates = body.get('updates', [])
    site_notes = (body.get('site_notes') or '').strip()
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404

    today = datetime.now().strftime('%Y-%m-%d')
    item_by_id = {i['id']: i for i in prop.get('scope_items', [])}

    # Apply any WCP item updates (backward compat — new inspector sends empty list)
    email_changes = []
    updated_count = 0
    for upd in updates:
        item = item_by_id.get(upd.get('item_id'))
        if item:
            before_pct = item.get('completion_pct', 0)
            if 'completion_pct' in upd:
                item['completion_pct'] = int(upd['completion_pct'])
            if 'notes' in upd and upd['notes']:
                item['notes'] = upd['notes']
            item['last_updated'] = today
            item['updated_by'] = 'inspector'
            updated_count += 1
            if item['completion_pct'] != before_pct:
                email_changes.append({
                    'name': item['name'],
                    'before': before_pct,
                    'after': item['completion_pct'],
                    'notes': upd.get('notes') or item.get('notes') or '',
                })

    # Store category-level progress snapshot from inspector (no WCP item writes — Barry reviews)
    category_updates = body.get('category_updates', [])

    # Pull pending photos and clear them
    pending_photos = prop.pop('pending_photos', []) or []

    # Match per-photo notes from submit body to uploaded photos
    photo_notes_map = {
        n['url'].split('?')[0]: n['note']
        for n in body.get('photo_notes', [])
        if n.get('url') and n.get('note', '').strip()
    }
    photos_for_record = []
    for p in pending_photos:
        entry = {'filename': p['filename'], 'url': p['url'], 'category': p.get('category', '')}
        note = photo_notes_map.get(p['url'].split('?')[0], '')
        if note:
            entry['note'] = note
        photos_for_record.append(entry)

    plan = prop.get('project_plan', {}) or {}
    plan.setdefault('inspections', []).append({
        'date': today,
        'items_updated': updated_count,
        'changes': email_changes,
        'site_notes': site_notes,
        'category_progress': category_updates,
        'photos': photos_for_record,
    })
    prop['project_plan'] = plan
    save_data(data)

    # Fire email in background so the APM doesn't wait on it
    if site_notes or email_changes or pending_photos:
        t = threading.Thread(
            target=_send_inspection_email,
            args=(prop, email_changes, pending_photos, site_notes),
            daemon=True,
        )
        t.start()

    metrics = calc_project_metrics(prop)
    return jsonify({'ok': True, 'updated': updated_count, 'photos': len(pending_photos),
                    'drawable_now': round(metrics.get('drawable_now', 0))})


@app.route('/api/flips/<prop_id>/inspections', methods=['GET'])
@login_required
def get_inspection_history(prop_id):
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if not prop:
        return jsonify({'error': 'Property not found'}), 404
    inspections = (prop.get('project_plan') or {}).get('inspections', [])
    return jsonify({'inspections': list(reversed(inspections))})


@app.route('/api/inspect/photo/<prop_id>/<item_id>', methods=['POST'])
def upload_inspect_photo(prop_id, item_id):
    if not _check_inspector_token():
        return jsonify({'error': 'Unauthorized'}), 403
    if 'file' not in request.files:
        return jsonify({'error': 'No file'}), 400
    f = request.files['file']
    ext = os.path.splitext(f.filename or 'photo.jpg')[1].lower() or '.jpg'
    if ext not in ('.jpg', '.jpeg', '.png', '.heic', '.webp'):
        return jsonify({'error': 'Unsupported image type'}), 400
    photo_dir = _ensure_photos_dir(prop_id, item_id)
    filename = f'{datetime.now().strftime("%Y%m%d_%H%M%S")}{ext}'
    filepath = os.path.join(photo_dir, filename)
    f.save(filepath)
    token = request.args.get('token') or request.headers.get('X-Inspector-Token', '')
    photo_url = f'/photos/{prop_id}/{item_id}/{filename}?token={token}'
    data = load_data()
    prop = next((p for p in data['properties'] if p.get('id') == prop_id), None)
    if prop:
        item = next((i for i in prop.get('scope_items', []) if i['id'] == item_id), None)
        if item:
            item.setdefault('photos', []).append({
                'url': photo_url,
                'filename': filename,
                'uploaded': datetime.now().strftime('%Y-%m-%d %H:%M'),
            })
            save_data(data)
    return jsonify({'ok': True, 'url': photo_url})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5002))
    app.run(debug=True, port=port, host='0.0.0.0')
