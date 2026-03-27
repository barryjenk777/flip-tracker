#!/usr/bin/env python3
"""
Flip Tracker - Professional Real Estate Flip Investment Dashboard
Standalone Flask application for tracking renovation flip investments.
"""

from flask import Flask, render_template, request, jsonify
import json
import os
from datetime import datetime

app = Flask(__name__)

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
        metrics = calc_prospect_metrics(p, settings)
        result.append({**p, 'metrics': metrics})
    return jsonify({'prospects': result, 'settings': settings})


@app.route('/api/prospects', methods=['POST'])
def add_prospect():
    data = load_data()
    p = request.json
    if not p.get('id'):
        slug = (p.get('address', 'prospect') or 'prospect').lower().replace(' ', '-')[:30]
        p['id'] = slug + '-' + datetime.now().strftime('%Y%m%d%H%M%S')
    p.setdefault('stage', 'new_lead')
    p.setdefault('date_added', datetime.now().strftime('%Y-%m-%d'))
    p.setdefault('verdict', None)
    p.setdefault('notes', '')
    p.setdefault('source', '')
    p.setdefault('beds', 0)
    p.setdefault('baths', 0)
    p.setdefault('sqft', 0)
    p.setdefault('year_built', 0)
    p.setdefault('monthly_rent_estimate', 0)
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
