#!/usr/bin/env python3
"""
Build Model 3: Solar + BESS
Infrastructure PE Interview Prep - Lotus Infrastructure Partners

Key features:
- No fuel cost
- ITC (30%) reduces equity requirement
- Solar degradation (0.5%/year)
- SCULPTED debt (sized to target DSCR, not cash sweep)
- BESS augmentation capex in Y7
- MACRS tax shield Y1-Y6
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

def build_solar_bess_model(output_path):
    """Build the Solar + BESS model with sculpted debt"""

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
    ws.cell(row=row, column=1, value="Solar+BESS Model - Lotus Infrastructure Partners").font = Font(bold=True, size=14)
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
    set_label(ws, row, "Avg CFADS (Y1-Y10)", "$mm")

    row = 6; rows['audit_min_dscr'] = row
    set_label(ws, row, "Min DSCR (Y1-Y10)", "x")

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
    set_label(ws, row, "Debt Payoff Y10")

    row = 13; rows['qa_min_dscr_check'] = row
    set_label(ws, row, "Min DSCR ≥ 1.30x")

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
    apply_input_style(ws, row, 4, 145)
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 18; rows['txn_fees'] = row
    set_label(ws, row, "Transaction Fees", "%", "Legal, advisory, financing")
    apply_input_style(ws, row, 4, 0.015)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 19; rows['itc_rate'] = row
    set_label(ws, row, "ITC Rate", "%", "Investment Tax Credit")
    apply_input_style(ws, row, 4, 0.30)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 20; rows['exit_multiple'] = row
    set_label(ws, row, "Exit Multiple", "x", "On exit year EBITDA")
    apply_input_style(ws, row, 4, 8.0)
    ws.cell(row=row, column=4).number_format = MULTIPLE_FORMAT

    row = 21; rows['exit_year'] = row
    set_label(ws, row, "Exit Year", "#", "Hold period")
    apply_input_style(ws, row, 4, 10)
    ws.cell(row=row, column=4).number_format = INTEGER_FORMAT

    row = 23
    # Asset Inputs
    rows['input_header_asset'] = row
    apply_header_style(ws, row, 1, "ASSET INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 24; rows['ac_capacity'] = row
    set_label(ws, row, "AC Capacity", "MWac", "Inverter rating")
    apply_input_style(ws, row, 4, 115)
    ws.cell(row=row, column=4).number_format = INTEGER_FORMAT

    row = 25; rows['capacity_factor'] = row
    set_label(ws, row, "Capacity Factor", "%", "P50 estimate")
    apply_input_style(ws, row, 4, 0.26)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 26; rows['degradation'] = row
    set_label(ws, row, "Degradation", "%/yr", "Annual output decline")
    apply_input_style(ws, row, 4, 0.005)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 28
    # Pricing Inputs
    rows['input_header_pricing'] = row
    apply_header_style(ws, row, 1, "PRICING INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 29; rows['ppa_price_y1'] = row
    set_label(ws, row, "PPA Price Y1", "$/MWh", "Contracted price")
    apply_input_style(ws, row, 4, 35)
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 30; rows['ppa_escalation'] = row
    set_label(ws, row, "PPA Escalation", "%/yr", "Contract escalation")
    apply_input_style(ws, row, 4, 0.015)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 31; rows['bess_revenue_y1'] = row
    set_label(ws, row, "BESS Revenue Y1", "$mm", "Arbitrage + ancillary")
    apply_input_style(ws, row, 4, 4.5)
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 32; rows['bess_escalation'] = row
    set_label(ws, row, "BESS Revenue Escalation", "%/yr", "Revenue growth")
    apply_input_style(ws, row, 4, 0.02)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 34
    # Cost Inputs
    rows['input_header_costs'] = row
    apply_header_style(ws, row, 1, "COST INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 35; rows['om_cost'] = row
    set_label(ws, row, "O&M Cost", "$mm/yr", "Fixed operating costs")
    apply_input_style(ws, row, 4, 2.0)
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 36; rows['om_escalation'] = row
    set_label(ws, row, "O&M Escalation", "%/yr", "Cost escalation")
    apply_input_style(ws, row, 4, 0.02)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 37; rows['aug_year'] = row
    set_label(ws, row, "Augmentation Year", "#", "BESS capacity augment")
    apply_input_style(ws, row, 4, 7)
    ws.cell(row=row, column=4).number_format = INTEGER_FORMAT

    row = 38; rows['aug_cost'] = row
    set_label(ws, row, "Augmentation Cost", "$mm", "Battery replacement")
    apply_input_style(ws, row, 4, 5)
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 40
    # Financing Inputs
    rows['input_header_fin'] = row
    apply_header_style(ws, row, 1, "FINANCING INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 41; rows['leverage'] = row
    set_label(ws, row, "Leverage", "%", "Debt / Entry EV")
    apply_input_style(ws, row, 4, 0.35)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 42; rows['interest_rate'] = row
    set_label(ws, row, "Interest Rate", "%", "Lower (contracted)")
    apply_input_style(ws, row, 4, 0.055)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 43; rows['debt_tenor'] = row
    set_label(ws, row, "Debt Tenor", "years", "Longer for contracted")
    apply_input_style(ws, row, 4, 10)
    ws.cell(row=row, column=4).number_format = INTEGER_FORMAT

    row = 44; rows['dsra_months'] = row
    set_label(ws, row, "DSRA Months", "months", "Debt service reserve")
    apply_input_style(ws, row, 4, 6)
    ws.cell(row=row, column=4).number_format = INTEGER_FORMAT

    row = 45; rows['target_dscr'] = row
    set_label(ws, row, "Target DSCR", "x", "For sculpting")
    apply_input_style(ws, row, 4, 1.35)
    ws.cell(row=row, column=4).number_format = MULTIPLE_FORMAT

    row = 46; rows['lockup_dscr'] = row
    set_label(ws, row, "Lock-up DSCR", "x", "Distribution threshold")
    apply_input_style(ws, row, 4, 1.10)
    ws.cell(row=row, column=4).number_format = MULTIPLE_FORMAT

    row = 48
    # Tax Inputs
    rows['input_header_tax'] = row
    apply_header_style(ws, row, 1, "TAX INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 49; rows['tax_rate'] = row
    set_label(ws, row, "Tax Rate", "%", "Blended federal + state")
    apply_input_style(ws, row, 4, 0.25)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 50; rows['macrs_years'] = row
    set_label(ws, row, "MACRS Shield Years", "#", "Years with no tax")
    apply_input_style(ws, row, 4, 6)
    ws.cell(row=row, column=4).number_format = INTEGER_FORMAT

    row = 51; rows['post_macrs_tax'] = row
    set_label(ws, row, "Post-MACRS Tax Rate", "%", "Reduced effective rate")
    apply_input_style(ws, row, 4, 0.15)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 53
    # ========================================================================
    # CALCULATED ITEMS
    # ========================================================================
    rows['calc_header'] = row
    apply_header_style(ws, row, 1, "CALCULATED ITEMS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 54; rows['y1_generation'] = row
    set_label(ws, row, "Y1 Generation", "MWh", "AC Cap × CF × 8760")
    apply_calc_style(ws, row, 4, f"=$D${rows['ac_capacity']}*$D${rows['capacity_factor']}*8760")
    ws.cell(row=row, column=4).number_format = '#,##0'

    row = 55; rows['debt_amount'] = row
    set_label(ws, row, "Debt Amount", "$mm", "Entry EV × Leverage")
    apply_calc_style(ws, row, 4, f"=$D${rows['entry_ev']}*$D${rows['leverage']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 56; rows['itc_benefit'] = row
    set_label(ws, row, "ITC Benefit", "$mm", "Entry EV × ITC Rate")
    apply_calc_style(ws, row, 4, f"=$D${rows['entry_ev']}*$D${rows['itc_rate']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 58
    # ========================================================================
    # OPERATING MODEL
    # ========================================================================
    rows['op_header'] = row
    apply_header_style(ws, row, 1, "OPERATING MODEL")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 59; rows['generation'] = row
    set_label(ws, row, "Generation", "MWh", "With degradation")
    for col in range(5, 15):
        yr = col - 4
        # Generation with degradation: Y1 Gen × (1 - degradation)^(year-1)
        formula = f"=$D${rows['y1_generation']}*(1-$D${rows['degradation']})^({yr}-1)"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = '#,##0'

    row = 60; rows['ppa_price'] = row
    set_label(ws, row, "PPA Price", "$/MWh", "Escalated")
    for col in range(5, 15):
        yr = col - 4
        formula = f"=$D${rows['ppa_price_y1']}*(1+$D${rows['ppa_escalation']})^({yr}-1)"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 62; rows['ppa_revenue'] = row
    set_label(ws, row, "PPA Revenue", "$mm", "Gen × PPA Price / 1M")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['generation']}*{col_letter(col)}{rows['ppa_price']}/1000000"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 63; rows['bess_revenue'] = row
    set_label(ws, row, "BESS Revenue", "$mm", "Arbitrage + ancillary")
    for col in range(5, 15):
        yr = col - 4
        formula = f"=$D${rows['bess_revenue_y1']}*(1+$D${rows['bess_escalation']})^({yr}-1)"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 64; rows['total_revenue'] = row
    set_label(ws, row, "Total Revenue", "$mm")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['ppa_revenue']}+{col_letter(col)}{rows['bess_revenue']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT
        ws.cell(row=row, column=col).font = BLACK_BOLD

    row = 66; rows['om_cost_yr'] = row
    set_label(ws, row, "O&M Cost", "$mm", "Escalated")
    for col in range(5, 15):
        yr = col - 4
        formula = f"=$D${rows['om_cost']}*(1+$D${rows['om_escalation']})^({yr}-1)"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 68; rows['ebitda'] = row
    set_label(ws, row, "EBITDA", "$mm")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['total_revenue']}-{col_letter(col)}{rows['om_cost_yr']}"
        apply_output_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 70
    # ========================================================================
    # CASH FLOW
    # ========================================================================
    rows['cf_header'] = row
    apply_header_style(ws, row, 1, "CASH FLOW")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 71; rows['taxes'] = row
    set_label(ws, row, "Taxes", "$mm", "MACRS shield Y1-Y6")
    for col in range(5, 15):
        yr = col - 4
        # Y1-Y6: No taxes (MACRS shield), Y7+: EBITDA × post-MACRS rate
        formula = f"=IF({yr}<=$D${rows['macrs_years']},0,{col_letter(col)}{rows['ebitda']}*$D${rows['post_macrs_tax']})"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 72; rows['augmentation'] = row
    set_label(ws, row, "BESS Augmentation", "$mm", "Battery replacement")
    for col in range(5, 15):
        yr = col - 4
        formula = f"=IF({yr}=$D${rows['aug_year']},$D${rows['aug_cost']},0)"
        apply_warning_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 73; rows['cfads'] = row
    set_label(ws, row, "CFADS", "$mm", "EBITDA - Taxes - Augment")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['ebitda']}-{col_letter(col)}{rows['taxes']}-{col_letter(col)}{rows['augmentation']}"
        apply_output_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 75
    # ========================================================================
    # DEBT SCHEDULE (SCULPTED to Target DSCR)
    # ========================================================================
    rows['debt_header'] = row
    apply_header_style(ws, row, 1, "DEBT SCHEDULE (Sculpted to 1.35x DSCR)")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 76; rows['debt_beg_bal'] = row
    set_label(ws, row, "Beginning Balance", "$mm")
    apply_calc_style(ws, row, 5, f"=$D${rows['debt_amount']}")
    ws.cell(row=row, column=5).number_format = CURRENCY_FORMAT
    for col in range(6, 15):
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 77; rows['interest'] = row
    set_label(ws, row, "Interest", "$mm", "Beg Bal × Int Rate")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['debt_beg_bal']}*$D${rows['interest_rate']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 78; rows['target_ds'] = row
    set_label(ws, row, "Target Debt Service", "$mm", "CFADS / Target DSCR")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['cfads']}/$D${rows['target_dscr']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 79; rows['sculpted_principal'] = row
    set_label(ws, row, "Sculpted Principal", "$mm", "Target DS - Interest")
    for col in range(5, 15):
        # Principal = MIN(MAX(0, Target DS - Interest), Beginning Balance)
        formula = f"=MIN(MAX(0,{col_letter(col)}{rows['target_ds']}-{col_letter(col)}{rows['interest']}),{col_letter(col)}{rows['debt_beg_bal']})"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 80; rows['actual_ds'] = row
    set_label(ws, row, "Actual Debt Service", "$mm", "Interest + Principal")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['interest']}+{col_letter(col)}{rows['sculpted_principal']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 81; rows['debt_end_bal'] = row
    set_label(ws, row, "Ending Balance", "$mm", "Beg - Principal")
    for col in range(5, 15):
        formula = f"=MAX(0,{col_letter(col)}{rows['debt_beg_bal']}-{col_letter(col)}{rows['sculpted_principal']})"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    # Fix beginning balance references
    for col in range(6, 15):
        prev_col = col_letter(col - 1)
        formula = f"={prev_col}{rows['debt_end_bal']}"
        apply_calc_style(ws, rows['debt_beg_bal'], col, formula)

    row = 82; rows['actual_dscr'] = row
    set_label(ws, row, "Actual DSCR", "x", "CFADS / Actual DS")
    for col in range(5, 15):
        formula = f"=IF({col_letter(col)}{rows['actual_ds']}>0,{col_letter(col)}{rows['cfads']}/{col_letter(col)}{rows['actual_ds']},0)"
        apply_output_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = MULTIPLE_FORMAT

    row = 84
    # ========================================================================
    # DSRA
    # ========================================================================
    rows['dsra_header'] = row
    apply_header_style(ws, row, 1, "DSRA (Debt Service Reserve Account)")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 85; rows['dsra_target'] = row
    set_label(ws, row, "DSRA Target", "$mm", "Next Yr DS × DSRA Months / 12")
    apply_calc_style(ws, row, 4, f"=E{rows['actual_ds']}*$D${rows['dsra_months']}/12")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT
    for col in range(5, 14):
        next_col = col_letter(col + 1)
        formula = f"={next_col}{rows['actual_ds']}*$D${rows['dsra_months']}/12"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT
    apply_calc_style(ws, row, 14, "=0")
    ws.cell(row=row, column=14).number_format = CURRENCY_FORMAT

    row = 86; rows['dsra_beg'] = row
    set_label(ws, row, "DSRA Beginning", "$mm")
    apply_calc_style(ws, row, 4, "=0")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT
    for col in range(5, 15):
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 87; rows['dsra_funding'] = row
    set_label(ws, row, "DSRA Funding/(Release)", "$mm", "Target - Beginning")
    for col in range(4, 15):
        formula = f"={col_letter(col)}{rows['dsra_target']}-{col_letter(col)}{rows['dsra_beg']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 88; rows['dsra_end'] = row
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

    row = 90
    # ========================================================================
    # DISTRIBUTION WATERFALL
    # ========================================================================
    rows['dist_header'] = row
    apply_header_style(ws, row, 1, "DISTRIBUTION WATERFALL")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 91; rows['cash_available'] = row
    set_label(ws, row, "Cash Available for Distribution", "$mm", "CFADS - DS - DSRA Fund")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['cfads']}-{col_letter(col)}{rows['actual_ds']}-{col_letter(col)}{rows['dsra_funding']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 92; rows['lockup_test'] = row
    set_label(ws, row, "Lock-up Test", "", "PASS if DSCR ≥ Lock-up")
    for col in range(5, 15):
        formula = f'=IF({col_letter(col)}{rows["actual_dscr"]}>=$D${rows["lockup_dscr"]},"PASS","LOCKED")'
        apply_calc_style(ws, row, col, formula)

    row = 93; rows['distributable'] = row
    set_label(ws, row, "Distributable Cash", "$mm", "If PASS, MAX(0, Cash Avail)")
    for col in range(5, 15):
        formula = f'=IF({col_letter(col)}{rows["lockup_test"]}="PASS",MAX(0,{col_letter(col)}{rows["cash_available"]}),0)'
        apply_output_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 95
    # ========================================================================
    # SOURCES & USES (with ITC)
    # ========================================================================
    rows['su_header'] = row
    apply_header_style(ws, row, 1, "SOURCES & USES (Year 0)")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 96; rows['uses_header'] = row
    set_label(ws, row, "USES", "", "")
    ws.cell(row=row, column=1).font = BLACK_BOLD

    row = 97; rows['use_purchase'] = row
    set_label(ws, row, "  Purchase Price", "$mm")
    apply_calc_style(ws, row, 4, f"=$D${rows['entry_ev']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 98; rows['use_fees'] = row
    set_label(ws, row, "  Transaction Fees", "$mm")
    apply_calc_style(ws, row, 4, f"=$D${rows['entry_ev']}*$D${rows['txn_fees']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 99; rows['use_dsra'] = row
    set_label(ws, row, "  DSRA Funding", "$mm")
    apply_calc_style(ws, row, 4, f"=D{rows['dsra_target']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 100; rows['total_uses'] = row
    set_label(ws, row, "Total Uses", "$mm")
    apply_output_style(ws, row, 4, f"=D{rows['use_purchase']}+D{rows['use_fees']}+D{rows['use_dsra']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 102; rows['sources_header'] = row
    set_label(ws, row, "SOURCES", "", "")
    ws.cell(row=row, column=1).font = BLACK_BOLD

    row = 103; rows['source_debt'] = row
    set_label(ws, row, "  Debt", "$mm")
    apply_calc_style(ws, row, 4, f"=$D${rows['debt_amount']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 104; rows['pre_itc_equity'] = row
    set_label(ws, row, "  Pre-ITC Equity", "$mm", "Uses - Debt")
    apply_calc_style(ws, row, 4, f"=D{rows['total_uses']}-D{rows['source_debt']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 105; rows['itc_source'] = row
    set_label(ws, row, "  ITC Benefit", "$mm", "Reduces equity required")
    apply_calc_style(ws, row, 4, f"=$D${rows['itc_benefit']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 106; rows['net_equity'] = row
    set_label(ws, row, "  Net Equity", "$mm", "Pre-ITC - ITC")
    apply_output_style(ws, row, 4, f"=D{rows['pre_itc_equity']}-D{rows['itc_source']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 107; rows['total_sources'] = row
    set_label(ws, row, "Total Sources", "$mm")
    apply_output_style(ws, row, 4, f"=D{rows['source_debt']}+D{rows['net_equity']}+D{rows['itc_source']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 108; rows['su_check'] = row
    set_label(ws, row, "Balance Check", "", "")
    apply_output_style(ws, row, 4, f'=IF(ABS(D{rows["total_uses"]}-D{rows["total_sources"]})<0.01,"PASS","FAIL")')

    row = 110
    # ========================================================================
    # EXIT & EQUITY RETURNS
    # ========================================================================
    rows['exit_header'] = row
    apply_header_style(ws, row, 1, "EXIT CALCULATIONS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 111; rows['exit_ev'] = row
    set_label(ws, row, "Exit EV", "$mm", "Exit Multiple × Exit Yr EBITDA")
    # Exit in Y10 (column N)
    apply_calc_style(ws, row, 14, f"=$D${rows['exit_multiple']}*N{rows['ebitda']}")
    ws.cell(row=row, column=14).number_format = CURRENCY_FORMAT

    row = 112; rows['exit_debt'] = row
    set_label(ws, row, "Exit Debt Payoff", "$mm", "Ending Bal at Exit")
    apply_calc_style(ws, row, 14, f"=N{rows['debt_end_bal']}")
    ws.cell(row=row, column=14).number_format = CURRENCY_FORMAT

    row = 113; rows['exit_dsra_release'] = row
    set_label(ws, row, "DSRA Release", "$mm", "DSRA Ending at Exit")
    apply_calc_style(ws, row, 14, f"=N{rows['dsra_end']}")
    ws.cell(row=row, column=14).number_format = CURRENCY_FORMAT

    row = 114; rows['exit_equity_proceeds'] = row
    set_label(ws, row, "Exit Equity Proceeds", "$mm", "Exit EV - Debt + DSRA")
    apply_output_style(ws, row, 14, f"=N{rows['exit_ev']}-N{rows['exit_debt']}+N{rows['exit_dsra_release']}")
    ws.cell(row=row, column=14).number_format = CURRENCY_FORMAT

    row = 116; rows['returns_header'] = row
    apply_header_style(ws, row, 1, "EQUITY RETURNS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 117; rows['equity_cf'] = row
    set_label(ws, row, "Equity Cash Flow", "$mm")
    # Y0 = -Net Equity (outflow)
    apply_calc_style(ws, row, 4, f"=-D{rows['net_equity']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT
    # Y1-Y9 = Distributable Cash
    for col in range(5, 14):
        formula = f"={col_letter(col)}{rows['distributable']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT
    # Y10 = Distributable + Exit Equity Proceeds
    apply_calc_style(ws, row, 14, f"=N{rows['distributable']}+N{rows['exit_equity_proceeds']}")
    ws.cell(row=row, column=14).number_format = CURRENCY_FORMAT

    row = 118; rows['levered_irr'] = row
    set_label(ws, row, "Levered IRR", "%")
    apply_output_style(ws, row, 4, f"=IRR(D{rows['equity_cf']}:N{rows['equity_cf']})")
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 119; rows['moic'] = row
    set_label(ws, row, "MOIC", "x", "Sum inflows / outflow")
    apply_output_style(ws, row, 4, f"=SUMIF(D{rows['equity_cf']}:N{rows['equity_cf']},\">0\")/ABS(D{rows['equity_cf']})")
    ws.cell(row=row, column=4).number_format = MULTIPLE_FORMAT

    row = 121; rows['unlev_header'] = row
    set_label(ws, row, "Unlevered Returns (for reference)", "")
    ws.cell(row=row, column=1).font = BLACK_BOLD

    row = 122; rows['unlev_cf'] = row
    set_label(ws, row, "Unlevered Cash Flow", "$mm")
    apply_calc_style(ws, row, 4, f"=-$D${rows['entry_ev']}-D{rows['use_fees']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT
    for col in range(5, 14):
        formula = f"={col_letter(col)}{rows['cfads']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT
    apply_calc_style(ws, row, 14, f"=N{rows['cfads']}+N{rows['exit_ev']}")
    ws.cell(row=row, column=14).number_format = CURRENCY_FORMAT

    row = 123; rows['unlev_irr'] = row
    set_label(ws, row, "Unlevered IRR", "%")
    apply_output_style(ws, row, 4, f"=IRR(D{rows['unlev_cf']}:N{rows['unlev_cf']})")
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    # ========================================================================
    # POPULATE AUDIT STRIP
    # ========================================================================
    apply_output_style(ws, rows['audit_entry_ev'], 4, f"=$D${rows['entry_ev']}")
    ws.cell(row=rows['audit_entry_ev'], column=4).number_format = CURRENCY_FORMAT

    apply_output_style(ws, rows['audit_y1_ebitda'], 4, f"=E{rows['ebitda']}")
    ws.cell(row=rows['audit_y1_ebitda'], column=4).number_format = CURRENCY_FORMAT

    apply_output_style(ws, rows['audit_avg_cfads'], 4, f"=AVERAGE(E{rows['cfads']}:N{rows['cfads']})")
    ws.cell(row=rows['audit_avg_cfads'], column=4).number_format = CURRENCY_FORMAT

    apply_output_style(ws, rows['audit_min_dscr'], 4, f"=MIN(E{rows['actual_dscr']}:N{rows['actual_dscr']})")
    ws.cell(row=rows['audit_min_dscr'], column=4).number_format = MULTIPLE_FORMAT

    apply_output_style(ws, rows['audit_levered_irr'], 4, f"=D{rows['levered_irr']}")
    ws.cell(row=rows['audit_levered_irr'], column=4).number_format = PERCENT_FORMAT

    apply_output_style(ws, rows['audit_moic'], 4, f"=D{rows['moic']}")
    ws.cell(row=rows['audit_moic'], column=4).number_format = MULTIPLE_FORMAT

    apply_output_style(ws, rows['audit_unlevered_irr'], 4, f"=D{rows['unlev_irr']}")
    ws.cell(row=rows['audit_unlevered_irr'], column=4).number_format = PERCENT_FORMAT

    apply_output_style(ws, rows['qa_su_balance'], 4, f"=D{rows['su_check']}")
    apply_output_style(ws, rows['qa_debt_payoff'], 4, f'=IF(N{rows["debt_end_bal"]}<1,"PASS","FAIL")')
    apply_output_style(ws, rows['qa_min_dscr_check'], 4, f'=IF(D{rows["audit_min_dscr"]}>=1.30,"PASS","FAIL")')

    # Freeze panes
    ws.freeze_panes = 'E2'

    wb.save(output_path)
    print(f"Solar+BESS model saved to: {output_path}")
    return rows

if __name__ == "__main__":
    output_path = "/mnt/user-data/outputs/Model_3_SolarBESS.xlsx"
    rows = build_solar_bess_model(output_path)
    print("Solar+BESS model built successfully!")
