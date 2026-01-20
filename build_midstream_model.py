#!/usr/bin/env python3
"""
Build Model 5: Midstream (Gas Gathering)
Infrastructure PE Interview Prep - Lotus Infrastructure Partners

Key features:
- Volume profile: Growth Y1-5 (3%/yr), Decline Y6-10 (5%/yr)
- Fee-for-service revenue ($/Mcf)
- Aggressive 75% cash sweep with 5-year short tenor
- Growth capex Y1-5 only
- Compressed exit multiple (6.5x) due to decline
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Formatting constants
YELLOW_FILL = PatternFill(start_color="FFFFC7", end_color="FFFFC7", fill_type="solid")
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
ORANGE_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
NAVY_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")

BLUE_BOLD = Font(color="0000FF", bold=True)
BLACK_FONT = Font(color="000000")
BLACK_BOLD = Font(color="000000", bold=True)
WHITE_BOLD = Font(color="FFFFFF", bold=True)

THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

CURRENCY_FORMAT = '#,##0.0_);(#,##0.0)'
PERCENT_FORMAT = '0.0%'
MULTIPLE_FORMAT = '0.00"x"'
INTEGER_FORMAT = '#,##0'

def col_letter(col_num):
    return get_column_letter(col_num)

def apply_input_style(ws, row, col, value=None):
    cell = ws.cell(row=row, column=col)
    if value is not None:
        cell.value = value
    cell.fill = YELLOW_FILL
    cell.font = BLUE_BOLD
    cell.border = THIN_BORDER
    return cell

def apply_calc_style(ws, row, col, formula=None):
    cell = ws.cell(row=row, column=col)
    if formula is not None:
        cell.value = formula
    cell.fill = PatternFill(fill_type=None)
    cell.font = BLACK_FONT
    cell.border = THIN_BORDER
    return cell

def apply_output_style(ws, row, col, formula=None):
    cell = ws.cell(row=row, column=col)
    if formula is not None:
        cell.value = formula
    cell.fill = GREEN_FILL
    cell.font = BLACK_BOLD
    cell.border = THIN_BORDER
    return cell

def apply_header_style(ws, row, col, text):
    cell = ws.cell(row=row, column=col)
    cell.value = text
    cell.fill = NAVY_FILL
    cell.font = WHITE_BOLD
    cell.alignment = Alignment(horizontal='left')
    return cell

def apply_warning_style(ws, row, col, formula=None):
    cell = ws.cell(row=row, column=col)
    if formula is not None:
        cell.value = formula
    cell.fill = ORANGE_FILL
    cell.font = BLACK_FONT
    cell.border = THIN_BORDER
    return cell

def set_label(ws, row, label, units="", notes=""):
    ws.cell(row=row, column=1, value=label).font = BLACK_FONT
    ws.cell(row=row, column=2, value=units).font = Font(color="666666", italic=True)
    ws.cell(row=row, column=3, value=notes).font = Font(color="666666", italic=True)

def build_midstream_model(output_path):
    """Build the Midstream model with growth-decline volume profile"""

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Model"

    rows = {}

    # Column widths
    ws.column_dimensions['A'].width = 32
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 38
    for col in range(4, 15):
        ws.column_dimensions[col_letter(col)].width = 14

    # ========================================================================
    # ROW 1: YEAR HEADERS
    # ========================================================================
    row = 1
    rows['year_header'] = row
    ws.cell(row=row, column=1, value="Midstream Model - Lotus Infrastructure Partners").font = Font(bold=True, size=14)
    ws.cell(row=row, column=4, value=0)
    for yr in range(1, 11):
        ws.cell(row=row, column=4+yr, value=yr)
    for col in range(4, 15):
        cell = ws.cell(row=row, column=col)
        cell.font = BLACK_BOLD
        cell.alignment = Alignment(horizontal='center')

    # ========================================================================
    # AUDIT STRIP (Rows 2-13)
    # ========================================================================
    row = 2
    rows['audit_header'] = row
    apply_header_style(ws, row, 1, "═══ AUDIT STRIP ═══")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 3; rows['audit_entry_ev'] = row
    set_label(ws, row, "Entry EV", "$mm")

    row = 4; rows['audit_y1_ebitda'] = row
    set_label(ws, row, "Y1 EBITDA", "$mm")

    row = 5; rows['audit_avg_cfads'] = row
    set_label(ws, row, "Avg CFADS (Y1-Y5)", "$mm")

    row = 6; rows['audit_min_dscr'] = row
    set_label(ws, row, "Min DSCR (Y1-Y5)", "x")

    row = 7; rows['audit_levered_irr'] = row
    set_label(ws, row, "Levered IRR", "%")

    row = 8; rows['audit_moic'] = row
    set_label(ws, row, "MOIC", "x")

    row = 9; rows['audit_unlevered_irr'] = row
    set_label(ws, row, "Unlevered IRR", "%")

    row = 10; rows['qa_header'] = row
    apply_header_style(ws, row, 1, "═══ QA CHECKS ═══")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 11; rows['qa_su_balance'] = row
    set_label(ws, row, "S&U Balance")

    row = 12; rows['qa_debt_payoff'] = row
    set_label(ws, row, "Debt Payoff Y5")

    row = 13; rows['qa_min_dscr_check'] = row
    set_label(ws, row, "Min DSCR ≥ 1.25x")

    row = 15

    # ========================================================================
    # INPUTS SECTION
    # ========================================================================

    # Transaction Inputs
    row = 16; rows['input_header_txn'] = row
    apply_header_style(ws, row, 1, "TRANSACTION INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 17; rows['entry_ev'] = row
    set_label(ws, row, "Entry EV", "$mm", "Purchase price")
    apply_input_style(ws, row, 4, 85)
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 18; rows['txn_fees'] = row
    set_label(ws, row, "Transaction Fees", "%", "Legal, advisory, financing")
    apply_input_style(ws, row, 4, 0.015)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 19; rows['exit_multiple'] = row
    set_label(ws, row, "Exit Multiple", "x", "Compressed for decline")
    apply_input_style(ws, row, 4, 6.5)
    ws.cell(row=row, column=4).number_format = MULTIPLE_FORMAT

    row = 20; rows['exit_year'] = row
    set_label(ws, row, "Exit Year", "#", "Hold period")
    apply_input_style(ws, row, 4, 7)
    ws.cell(row=row, column=4).number_format = INTEGER_FORMAT

    row = 22
    # Volume Inputs
    rows['input_header_volume'] = row
    apply_header_style(ws, row, 1, "VOLUME INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 23; rows['y1_volume'] = row
    set_label(ws, row, "Y1 Volume", "MMcf/d", "Daily throughput")
    apply_input_style(ws, row, 4, 80)
    ws.cell(row=row, column=4).number_format = INTEGER_FORMAT

    row = 24; rows['growth_rate'] = row
    set_label(ws, row, "Growth Rate Y1-Y5", "%/yr", "Drilling phase")
    apply_input_style(ws, row, 4, 0.03)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 25; rows['decline_rate'] = row
    set_label(ws, row, "Decline Rate Y6-Y10", "%/yr", "Depletion phase")
    apply_input_style(ws, row, 4, 0.05)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 27
    # Pricing Inputs
    rows['input_header_pricing'] = row
    apply_header_style(ws, row, 1, "PRICING INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 28; rows['total_fee'] = row
    set_label(ws, row, "Total Fee", "$/Mcf", "Gathering + compression")
    apply_input_style(ws, row, 4, 0.55)
    ws.cell(row=row, column=4).number_format = '#,##0.00'

    row = 29; rows['fee_escalation'] = row
    set_label(ws, row, "Fee Escalation", "%/yr", "Contract escalation")
    apply_input_style(ws, row, 4, 0.02)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 31
    # Cost Inputs
    rows['input_header_costs'] = row
    apply_header_style(ws, row, 1, "COST INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 32; rows['om_y1'] = row
    set_label(ws, row, "O&M Y1", "$mm", "Operating costs")
    apply_input_style(ws, row, 4, 3.5)
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 33; rows['om_escalation'] = row
    set_label(ws, row, "O&M Escalation", "%/yr", "Cost escalation")
    apply_input_style(ws, row, 4, 0.025)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 34; rows['growth_capex'] = row
    set_label(ws, row, "Growth Capex Y1-Y5", "$mm/yr", "Compression, looping")
    apply_input_style(ws, row, 4, 2)
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 36
    # Financing Inputs
    rows['input_header_fin'] = row
    apply_header_style(ws, row, 1, "FINANCING INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 37; rows['leverage'] = row
    set_label(ws, row, "Leverage", "%", "Debt / Entry EV")
    apply_input_style(ws, row, 4, 0.30)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 38; rows['interest_rate'] = row
    set_label(ws, row, "Interest Rate", "%", "Higher (basin risk)")
    apply_input_style(ws, row, 4, 0.07)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 39; rows['debt_tenor'] = row
    set_label(ws, row, "Debt Tenor", "years", "Short - pre-decline")
    apply_input_style(ws, row, 4, 5)
    ws.cell(row=row, column=4).number_format = INTEGER_FORMAT

    row = 40; rows['dsra_months'] = row
    set_label(ws, row, "DSRA Months", "months", "Debt service reserve")
    apply_input_style(ws, row, 4, 6)
    ws.cell(row=row, column=4).number_format = INTEGER_FORMAT

    row = 41; rows['sweep_pct'] = row
    set_label(ws, row, "Cash Sweep %", "%", "Aggressive paydown")
    apply_input_style(ws, row, 4, 0.75)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 42; rows['lockup_dscr'] = row
    set_label(ws, row, "Lock-up DSCR", "x", "Distribution threshold")
    apply_input_style(ws, row, 4, 1.10)
    ws.cell(row=row, column=4).number_format = MULTIPLE_FORMAT

    row = 44
    # Tax Inputs
    rows['input_header_tax'] = row
    apply_header_style(ws, row, 1, "TAX INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 45; rows['tax_rate'] = row
    set_label(ws, row, "Tax Rate", "%", "Blended federal + state")
    apply_input_style(ws, row, 4, 0.25)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 46; rows['depreciable_basis'] = row
    set_label(ws, row, "Depreciable Basis", "%", "% of EV for depreciation")
    apply_input_style(ws, row, 4, 0.80)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 47; rows['depreciation_life'] = row
    set_label(ws, row, "Depreciation Life", "years", "Straight-line")
    apply_input_style(ws, row, 4, 15)
    ws.cell(row=row, column=4).number_format = INTEGER_FORMAT

    row = 49
    # ========================================================================
    # CALCULATED ITEMS
    # ========================================================================
    rows['calc_header'] = row
    apply_header_style(ws, row, 1, "CALCULATED ITEMS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 50; rows['debt_amount'] = row
    set_label(ws, row, "Debt Amount", "$mm", "Entry EV × Leverage")
    apply_calc_style(ws, row, 4, f"=$D${rows['entry_ev']}*$D${rows['leverage']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 51; rows['annual_depreciation'] = row
    set_label(ws, row, "Annual Depreciation", "$mm", "EV × Basis / Life")
    apply_calc_style(ws, row, 4, f"=$D${rows['entry_ev']}*$D${rows['depreciable_basis']}/$D${rows['depreciation_life']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 52; rows['scheduled_principal'] = row
    set_label(ws, row, "Scheduled Principal", "$mm/yr", "Debt / Tenor")
    apply_calc_style(ws, row, 4, f"=$D${rows['debt_amount']}/$D${rows['debt_tenor']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 54
    # ========================================================================
    # OPERATING MODEL
    # ========================================================================
    rows['op_header'] = row
    apply_header_style(ws, row, 1, "OPERATING MODEL")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 55; rows['daily_volume'] = row
    set_label(ws, row, "Daily Volume", "MMcf/d", "Growth then decline")
    for col in range(5, 15):
        yr = col - 4
        # Y1-Y5: Growth, Y6-Y10: Decline from Y5 level
        if yr <= 5:
            formula = f"=$D${rows['y1_volume']}*(1+$D${rows['growth_rate']})^({yr}-1)"
        else:
            # Y5 volume × (1 - decline)^(yr-5)
            formula = f"=$D${rows['y1_volume']}*(1+$D${rows['growth_rate']})^4*(1-$D${rows['decline_rate']})^({yr}-5)"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = '#,##0.0'

    row = 56; rows['annual_volume'] = row
    set_label(ws, row, "Annual Volume", "Bcf", "Daily × 365 / 1000")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['daily_volume']}*365/1000"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = '#,##0.0'

    row = 57; rows['fee'] = row
    set_label(ws, row, "Fee", "$/Mcf", "Escalated")
    for col in range(5, 15):
        yr = col - 4
        formula = f"=$D${rows['total_fee']}*(1+$D${rows['fee_escalation']})^({yr}-1)"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = '#,##0.00'

    row = 59; rows['revenue'] = row
    set_label(ws, row, "Revenue", "$mm", "Volume(Bcf) × Fee × 1000")
    for col in range(5, 15):
        # Revenue = Volume (Bcf) × Fee ($/Mcf) × 1000 Mcf/Bcf × 1M / 1M = Volume × Fee × 1000 / 1M
        # But Bcf = 1000 MMcf, and Fee is $/Mcf, so:
        # Revenue ($mm) = Volume (Bcf) × 1000 (MMcf/Bcf) × 1000 (Mcf/MMcf) × Fee ($/Mcf) / 1,000,000
        # = Volume (Bcf) × Fee × 1000
        formula = f"={col_letter(col)}{rows['annual_volume']}*{col_letter(col)}{rows['fee']}*1000"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 60; rows['om_cost'] = row
    set_label(ws, row, "O&M Cost", "$mm", "Escalated")
    for col in range(5, 15):
        yr = col - 4
        formula = f"=$D${rows['om_y1']}*(1+$D${rows['om_escalation']})^({yr}-1)"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 62; rows['ebitda'] = row
    set_label(ws, row, "EBITDA", "$mm")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['revenue']}-{col_letter(col)}{rows['om_cost']}"
        apply_output_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 64
    # ========================================================================
    # CASH FLOW
    # ========================================================================
    rows['cf_header'] = row
    apply_header_style(ws, row, 1, "CASH FLOW")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 65; rows['ebit'] = row
    set_label(ws, row, "EBIT", "$mm", "EBITDA - Depreciation")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['ebitda']}-$D${rows['annual_depreciation']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 66; rows['taxes'] = row
    set_label(ws, row, "Taxes", "$mm", "MAX(0, EBIT) × Tax Rate")
    for col in range(5, 15):
        formula = f"=MAX(0,{col_letter(col)}{rows['ebit']})*$D${rows['tax_rate']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 67; rows['growth_capex_yr'] = row
    set_label(ws, row, "Growth Capex", "$mm", "Y1-Y5 only")
    for col in range(5, 15):
        yr = col - 4
        formula = f"=IF({yr}<=5,$D${rows['growth_capex']},0)"
        apply_warning_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 68; rows['cfads'] = row
    set_label(ws, row, "CFADS", "$mm", "EBITDA - Taxes - Capex")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['ebitda']}-{col_letter(col)}{rows['taxes']}-{col_letter(col)}{rows['growth_capex_yr']}"
        apply_output_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 70
    # ========================================================================
    # DEBT SCHEDULE (75% Cash Sweep, 5-Year Tenor)
    # ========================================================================
    rows['debt_header'] = row
    apply_header_style(ws, row, 1, "DEBT SCHEDULE (75% Sweep, 5-Year Tenor)")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 71; rows['debt_beg_bal'] = row
    set_label(ws, row, "Beginning Balance", "$mm")
    apply_calc_style(ws, row, 5, f"=$D${rows['debt_amount']}")
    ws.cell(row=row, column=5).number_format = CURRENCY_FORMAT
    for col in range(6, 15):
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 72; rows['interest'] = row
    set_label(ws, row, "Interest", "$mm", "Beg Bal × Int Rate")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['debt_beg_bal']}*$D${rows['interest_rate']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 73; rows['sched_principal'] = row
    set_label(ws, row, "Scheduled Principal", "$mm", "MIN(Sched, Beg Bal)")
    for col in range(5, 15):
        formula = f"=MIN($D${rows['scheduled_principal']},{col_letter(col)}{rows['debt_beg_bal']})"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 74; rows['debt_service_pre'] = row
    set_label(ws, row, "Debt Service (Pre-Sweep)", "$mm", "Interest + Sched Prin")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['interest']}+{col_letter(col)}{rows['sched_principal']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 75; rows['excess_cash'] = row
    set_label(ws, row, "Excess Cash", "$mm", "CFADS - DS - DSRA Funding")

    row = 76; rows['cash_sweep'] = row
    set_label(ws, row, "Cash Sweep", "$mm", "MIN(Excess×75%, Beg-Sched)")
    for col in range(5, 15):
        formula = f"=MIN(MAX(0,{col_letter(col)}{rows['excess_cash']})*$D${rows['sweep_pct']},MAX(0,{col_letter(col)}{rows['debt_beg_bal']}-{col_letter(col)}{rows['sched_principal']}))"
        apply_warning_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 77; rows['total_principal'] = row
    set_label(ws, row, "Total Principal", "$mm", "Scheduled + Sweep")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['sched_principal']}+{col_letter(col)}{rows['cash_sweep']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 78; rows['debt_end_bal'] = row
    set_label(ws, row, "Ending Balance", "$mm", "MAX(0, Beg - Total Prin)")
    for col in range(5, 15):
        formula = f"=MAX(0,{col_letter(col)}{rows['debt_beg_bal']}-{col_letter(col)}{rows['total_principal']})"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    # Fix beginning balance references
    for col in range(6, 15):
        prev_col = col_letter(col - 1)
        formula = f"={prev_col}{rows['debt_end_bal']}"
        apply_calc_style(ws, rows['debt_beg_bal'], col, formula)

    row = 79; rows['dscr'] = row
    set_label(ws, row, "DSCR", "x", "CFADS / Debt Service")
    for col in range(5, 15):
        formula = f"=IF({col_letter(col)}{rows['debt_service_pre']}>0,{col_letter(col)}{rows['cfads']}/{col_letter(col)}{rows['debt_service_pre']},0)"
        apply_output_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = MULTIPLE_FORMAT

    # Add Scheduled DS row (for DSRA calculation - breaks circularity)
    row = 80; rows['scheduled_ds'] = row
    set_label(ws, row, "Scheduled DS (for DSRA)", "$mm", "Based on straight-line amort, no sweep")
    for col in range(5, 15):
        yr = col - 4
        # Calculate scheduled DS based on straight-line amortization only
        # Scheduled DS = (Debt Amount - (yr-1) × Scheduled Principal) × Interest Rate + Scheduled Principal
        formula = f"=MAX(0,$D${rows['debt_amount']}-({yr}-1)*$D${rows['scheduled_principal']})*$D${rows['interest_rate']}+$D${rows['scheduled_principal']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 82
    # ========================================================================
    # DSRA
    # ========================================================================
    rows['dsra_header'] = row
    apply_header_style(ws, row, 1, "DSRA (Debt Service Reserve Account)")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    # DSRA Target uses scheduled_ds (not debt_service_pre) to break circular reference
    row = 83; rows['dsra_target'] = row
    set_label(ws, row, "DSRA Target", "$mm", "Next Yr Scheduled DS × DSRA Months / 12")
    apply_calc_style(ws, row, 4, f"=E{rows['scheduled_ds']}*$D${rows['dsra_months']}/12")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT
    for col in range(5, 14):
        next_col = col_letter(col + 1)
        formula = f"={next_col}{rows['scheduled_ds']}*$D${rows['dsra_months']}/12"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT
    apply_calc_style(ws, row, 14, "=0")
    ws.cell(row=row, column=14).number_format = CURRENCY_FORMAT

    row = 84; rows['dsra_beg'] = row
    set_label(ws, row, "DSRA Beginning", "$mm")
    apply_calc_style(ws, row, 4, "=0")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT
    for col in range(5, 15):
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 85; rows['dsra_funding'] = row
    set_label(ws, row, "DSRA Funding/(Release)", "$mm", "Target - Beginning")
    for col in range(4, 15):
        formula = f"={col_letter(col)}{rows['dsra_target']}-{col_letter(col)}{rows['dsra_beg']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 86; rows['dsra_end'] = row
    set_label(ws, row, "DSRA Ending", "$mm", "Beginning + Funding")
    for col in range(4, 15):
        formula = f"={col_letter(col)}{rows['dsra_beg']}+{col_letter(col)}{rows['dsra_funding']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    # Fix DSRA Beginning references
    for col in range(5, 15):
        prev_col = col_letter(col - 1)
        formula = f"={prev_col}{rows['dsra_end']}"
        apply_calc_style(ws, rows['dsra_beg'], col, formula)

    # Add Excess Cash formulas
    for col in range(5, 15):
        formula = f"=MAX(0,{col_letter(col)}{rows['cfads']}-{col_letter(col)}{rows['debt_service_pre']}-{col_letter(col)}{rows['dsra_funding']})"
        apply_calc_style(ws, rows['excess_cash'], col, formula)
        ws.cell(row=rows['excess_cash'], column=col).number_format = CURRENCY_FORMAT

    row = 87
    # ========================================================================
    # DISTRIBUTION WATERFALL
    # ========================================================================
    rows['dist_header'] = row
    apply_header_style(ws, row, 1, "DISTRIBUTION WATERFALL")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 88; rows['cash_available'] = row
    set_label(ws, row, "Cash Available for Distribution", "$mm", "CFADS - Total DS - DSRA Fund")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['cfads']}-({col_letter(col)}{rows['interest']}+{col_letter(col)}{rows['total_principal']})-{col_letter(col)}{rows['dsra_funding']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 89; rows['lockup_test'] = row
    set_label(ws, row, "Lock-up Test", "", "PASS if DSCR ≥ Lock-up")
    for col in range(5, 15):
        formula = f'=IF({col_letter(col)}{rows["dscr"]}>=$D${rows["lockup_dscr"]},"PASS","LOCKED")'
        apply_calc_style(ws, row, col, formula)

    row = 90; rows['distributable'] = row
    set_label(ws, row, "Distributable Cash", "$mm", "If PASS, MAX(0, Cash Avail)")
    for col in range(5, 15):
        formula = f'=IF({col_letter(col)}{rows["lockup_test"]}="PASS",MAX(0,{col_letter(col)}{rows["cash_available"]}),0)'
        apply_output_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 92
    # ========================================================================
    # SOURCES & USES
    # ========================================================================
    rows['su_header'] = row
    apply_header_style(ws, row, 1, "SOURCES & USES (Year 0)")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 93; rows['uses_header'] = row
    set_label(ws, row, "USES", "", "")
    ws.cell(row=row, column=1).font = BLACK_BOLD

    row = 94; rows['use_purchase'] = row
    set_label(ws, row, "  Purchase Price", "$mm")
    apply_calc_style(ws, row, 4, f"=$D${rows['entry_ev']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 95; rows['use_fees'] = row
    set_label(ws, row, "  Transaction Fees", "$mm")
    apply_calc_style(ws, row, 4, f"=$D${rows['entry_ev']}*$D${rows['txn_fees']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 96; rows['use_dsra'] = row
    set_label(ws, row, "  DSRA Funding", "$mm")
    apply_calc_style(ws, row, 4, f"=D{rows['dsra_target']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 97; rows['total_uses'] = row
    set_label(ws, row, "Total Uses", "$mm")
    apply_output_style(ws, row, 4, f"=D{rows['use_purchase']}+D{rows['use_fees']}+D{rows['use_dsra']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 99; rows['sources_header'] = row
    set_label(ws, row, "SOURCES", "", "")
    ws.cell(row=row, column=1).font = BLACK_BOLD

    row = 100; rows['source_debt'] = row
    set_label(ws, row, "  Debt", "$mm")
    apply_calc_style(ws, row, 4, f"=$D${rows['debt_amount']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 101; rows['source_equity'] = row
    set_label(ws, row, "  Equity", "$mm", "Plug to balance")
    apply_calc_style(ws, row, 4, f"=D{rows['total_uses']}-D{rows['source_debt']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 102; rows['total_sources'] = row
    set_label(ws, row, "Total Sources", "$mm")
    apply_output_style(ws, row, 4, f"=D{rows['source_debt']}+D{rows['source_equity']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 103; rows['su_check'] = row
    set_label(ws, row, "Balance Check", "", "")
    apply_output_style(ws, row, 4, f'=IF(ABS(D{rows["total_uses"]}-D{rows["total_sources"]})<0.01,"PASS","FAIL")')

    row = 105
    # ========================================================================
    # EXIT & EQUITY RETURNS
    # ========================================================================
    rows['exit_header'] = row
    apply_header_style(ws, row, 1, "EXIT CALCULATIONS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 106; rows['exit_ev'] = row
    set_label(ws, row, "Exit EV", "$mm", "Exit Multiple × Exit Yr EBITDA")
    # Exit in Y7 (column K)
    apply_calc_style(ws, row, 11, f"=$D${rows['exit_multiple']}*K{rows['ebitda']}")
    ws.cell(row=row, column=11).number_format = CURRENCY_FORMAT

    row = 107; rows['exit_debt'] = row
    set_label(ws, row, "Exit Debt Payoff", "$mm", "Ending Bal at Exit")
    apply_calc_style(ws, row, 11, f"=K{rows['debt_end_bal']}")
    ws.cell(row=row, column=11).number_format = CURRENCY_FORMAT

    row = 108; rows['exit_dsra_release'] = row
    set_label(ws, row, "DSRA Release", "$mm", "DSRA Ending at Exit")
    apply_calc_style(ws, row, 11, f"=K{rows['dsra_end']}")
    ws.cell(row=row, column=11).number_format = CURRENCY_FORMAT

    row = 109; rows['exit_equity_proceeds'] = row
    set_label(ws, row, "Exit Equity Proceeds", "$mm", "Exit EV - Debt + DSRA")
    apply_output_style(ws, row, 11, f"=K{rows['exit_ev']}-K{rows['exit_debt']}+K{rows['exit_dsra_release']}")
    ws.cell(row=row, column=11).number_format = CURRENCY_FORMAT

    row = 111; rows['returns_header'] = row
    apply_header_style(ws, row, 1, "EQUITY RETURNS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 112; rows['equity_cf'] = row
    set_label(ws, row, "Equity Cash Flow", "$mm")
    apply_calc_style(ws, row, 4, f"=-D{rows['source_equity']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT
    for col in range(5, 11):  # Y1-Y6
        formula = f"={col_letter(col)}{rows['distributable']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT
    # Y7 = Distributable + Exit Equity Proceeds
    apply_calc_style(ws, row, 11, f"=K{rows['distributable']}+K{rows['exit_equity_proceeds']}")
    ws.cell(row=row, column=11).number_format = CURRENCY_FORMAT
    # Y8-Y10 = 0 (post-exit)
    for col in range(12, 15):
        apply_calc_style(ws, row, col, "=0")
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 113; rows['levered_irr'] = row
    set_label(ws, row, "Levered IRR", "%")
    apply_output_style(ws, row, 4, f"=IRR(D{rows['equity_cf']}:K{rows['equity_cf']})")
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 114; rows['moic'] = row
    set_label(ws, row, "MOIC", "x", "Sum inflows / outflow")
    apply_output_style(ws, row, 4, f"=SUMIF(D{rows['equity_cf']}:K{rows['equity_cf']},\">0\")/ABS(D{rows['equity_cf']})")
    ws.cell(row=row, column=4).number_format = MULTIPLE_FORMAT

    row = 116; rows['unlev_header'] = row
    set_label(ws, row, "Unlevered Returns (for reference)", "")
    ws.cell(row=row, column=1).font = BLACK_BOLD

    row = 117; rows['unlev_cf'] = row
    set_label(ws, row, "Unlevered Cash Flow", "$mm")
    apply_calc_style(ws, row, 4, f"=-$D${rows['entry_ev']}-D{rows['use_fees']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT
    for col in range(5, 11):
        formula = f"={col_letter(col)}{rows['cfads']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT
    apply_calc_style(ws, row, 11, f"=K{rows['cfads']}+K{rows['exit_ev']}")
    ws.cell(row=row, column=11).number_format = CURRENCY_FORMAT
    for col in range(12, 15):
        apply_calc_style(ws, row, col, "=0")
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 118; rows['unlev_irr'] = row
    set_label(ws, row, "Unlevered IRR", "%")
    apply_output_style(ws, row, 4, f"=IRR(D{rows['unlev_cf']}:K{rows['unlev_cf']})")
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    # ========================================================================
    # POPULATE AUDIT STRIP
    # ========================================================================
    apply_output_style(ws, rows['audit_entry_ev'], 4, f"=$D${rows['entry_ev']}")
    ws.cell(row=rows['audit_entry_ev'], column=4).number_format = CURRENCY_FORMAT

    apply_output_style(ws, rows['audit_y1_ebitda'], 4, f"=E{rows['ebitda']}")
    ws.cell(row=rows['audit_y1_ebitda'], column=4).number_format = CURRENCY_FORMAT

    apply_output_style(ws, rows['audit_avg_cfads'], 4, f"=AVERAGE(E{rows['cfads']}:I{rows['cfads']})")
    ws.cell(row=rows['audit_avg_cfads'], column=4).number_format = CURRENCY_FORMAT

    apply_output_style(ws, rows['audit_min_dscr'], 4, f"=MIN(E{rows['dscr']}:I{rows['dscr']})")
    ws.cell(row=rows['audit_min_dscr'], column=4).number_format = MULTIPLE_FORMAT

    apply_output_style(ws, rows['audit_levered_irr'], 4, f"=D{rows['levered_irr']}")
    ws.cell(row=rows['audit_levered_irr'], column=4).number_format = PERCENT_FORMAT

    apply_output_style(ws, rows['audit_moic'], 4, f"=D{rows['moic']}")
    ws.cell(row=rows['audit_moic'], column=4).number_format = MULTIPLE_FORMAT

    apply_output_style(ws, rows['audit_unlevered_irr'], 4, f"=D{rows['unlev_irr']}")
    ws.cell(row=rows['audit_unlevered_irr'], column=4).number_format = PERCENT_FORMAT

    apply_output_style(ws, rows['qa_su_balance'], 4, f"=D{rows['su_check']}")
    apply_output_style(ws, rows['qa_debt_payoff'], 4, f'=IF(I{rows["debt_end_bal"]}<1,"PASS","FAIL")')
    apply_output_style(ws, rows['qa_min_dscr_check'], 4, f'=IF(D{rows["audit_min_dscr"]}>=1.25,"PASS","FAIL")')

    # Freeze panes
    ws.freeze_panes = 'E2'

    wb.save(output_path)
    print(f"Midstream model saved to: {output_path}")
    return rows

if __name__ == "__main__":
    output_path = "/mnt/user-data/outputs/Model_5_Midstream.xlsx"
    rows = build_midstream_model(output_path)
    print("Midstream model built successfully!")
