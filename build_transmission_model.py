#!/usr/bin/env python3
"""
Build Model 4: Transmission
Infrastructure PE Interview Prep - Lotus Infrastructure Partners

Key features:
- Construction period (Y0-Y2) with no revenue
- Construction interest capitalization into RAB
- Phased equity deployment
- Regulated tariff revenue (Y3+ only)
- Sculpted debt (Y3+) at 1.40x DSCR
- Exit at RAB multiple
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

def build_transmission_model(output_path):
    """Build the Transmission model with construction period"""

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Model"

    rows = {}

    # Column widths
    ws.column_dimensions['A'].width = 34
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 38
    for col in range(4, 15):
        ws.column_dimensions[col_letter(col)].width = 14

    # ========================================================================
    # ROW 1: YEAR HEADERS (Y0=Construction start, Y3=COD)
    # ========================================================================
    row = 1
    rows['year_header'] = row
    ws.cell(row=row, column=1, value="Transmission Model - Lotus Infrastructure Partners").font = Font(bold=True, size=14)
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

    row = 3; rows['audit_rab_cod'] = row
    set_label(ws, row, "RAB at COD", "$mm")

    row = 4; rows['audit_y3_ebitda'] = row
    set_label(ws, row, "Y3 EBITDA (First Op Year)", "$mm")

    row = 5; rows['audit_avg_cfads'] = row
    set_label(ws, row, "Avg CFADS (Y3-Y10)", "$mm")

    row = 6; rows['audit_min_dscr'] = row
    set_label(ws, row, "Min DSCR (Y3-Y10)", "x")

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
    set_label(ws, row, "Min DSCR ≥ 1.35x")

    row = 15

    # ========================================================================
    # INPUTS SECTION
    # ========================================================================

    # Construction Inputs
    row = 16; rows['input_header_const'] = row
    apply_header_style(ws, row, 1, "CONSTRUCTION INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 17; rows['dev_cost_y0'] = row
    set_label(ws, row, "Development Cost Y0", "$mm", "Pre-construction dev")
    apply_input_style(ws, row, 4, 15)
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 18; rows['dev_cost_y1'] = row
    set_label(ws, row, "Development Cost Y1", "$mm", "Continued development")
    apply_input_style(ws, row, 4, 20)
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 19; rows['construction_cost'] = row
    set_label(ws, row, "Construction Cost Y2", "$mm", "Main construction")
    apply_input_style(ws, row, 4, 550)
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 20; rows['cod_year'] = row
    set_label(ws, row, "COD Year", "#", "Commercial operation date")
    apply_input_style(ws, row, 4, 3)
    ws.cell(row=row, column=4).number_format = INTEGER_FORMAT

    row = 22
    # Asset Inputs
    rows['input_header_asset'] = row
    apply_header_style(ws, row, 1, "ASSET INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 23; rows['transfer_capacity'] = row
    set_label(ws, row, "Transfer Capacity", "MW", "Line rating")
    apply_input_style(ws, row, 4, 3200)
    ws.cell(row=row, column=4).number_format = INTEGER_FORMAT

    row = 24; rows['availability'] = row
    set_label(ws, row, "Availability", "%", "Expected uptime")
    apply_input_style(ws, row, 4, 0.98)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 25; rows['asset_life'] = row
    set_label(ws, row, "Asset Life", "years", "For depreciation")
    apply_input_style(ws, row, 4, 40)
    ws.cell(row=row, column=4).number_format = INTEGER_FORMAT

    row = 27
    # Pricing Inputs
    rows['input_header_pricing'] = row
    apply_header_style(ws, row, 1, "PRICING INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 28; rows['tariff_rate'] = row
    set_label(ws, row, "Tariff Rate", "$/kW-yr", "FERC-approved")
    apply_input_style(ws, row, 4, 32)
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 29; rows['tariff_escalation'] = row
    set_label(ws, row, "Tariff Escalation", "%/yr", "Rate increase")
    apply_input_style(ws, row, 4, 0.02)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 31
    # Cost Inputs
    rows['input_header_costs'] = row
    apply_header_style(ws, row, 1, "COST INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 32; rows['om_y3'] = row
    set_label(ws, row, "O&M Y3", "$mm", "Post-COD O&M")
    apply_input_style(ws, row, 4, 3)
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 33; rows['om_escalation'] = row
    set_label(ws, row, "O&M Escalation", "%/yr", "Cost escalation")
    apply_input_style(ws, row, 4, 0.025)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 35
    # Financing Inputs
    rows['input_header_fin'] = row
    apply_header_style(ws, row, 1, "FINANCING INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 36; rows['const_debt_pct'] = row
    set_label(ws, row, "Construction Debt %", "%", "Const loan leverage")
    apply_input_style(ws, row, 4, 0.85)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 37; rows['term_debt_pct_rab'] = row
    set_label(ws, row, "Term Debt % of RAB", "%", "Refi at COD")
    apply_input_style(ws, row, 4, 0.60)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 38; rows['interest_rate'] = row
    set_label(ws, row, "Interest Rate", "%", "All-in rate")
    apply_input_style(ws, row, 4, 0.05)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 39; rows['dsra_months'] = row
    set_label(ws, row, "DSRA Months", "months", "Debt service reserve")
    apply_input_style(ws, row, 4, 6)
    ws.cell(row=row, column=4).number_format = INTEGER_FORMAT

    row = 40; rows['target_dscr'] = row
    set_label(ws, row, "Target DSCR", "x", "For sculpting")
    apply_input_style(ws, row, 4, 1.40)
    ws.cell(row=row, column=4).number_format = MULTIPLE_FORMAT

    row = 41; rows['lockup_dscr'] = row
    set_label(ws, row, "Lock-up DSCR", "x", "Distribution threshold")
    apply_input_style(ws, row, 4, 1.15)
    ws.cell(row=row, column=4).number_format = MULTIPLE_FORMAT

    row = 43
    # Exit Inputs
    rows['input_header_exit'] = row
    apply_header_style(ws, row, 1, "EXIT INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 44; rows['exit_xrab'] = row
    set_label(ws, row, "Exit xRAB Multiple", "x", "Exit at RAB multiple")
    apply_input_style(ws, row, 4, 1.15)
    ws.cell(row=row, column=4).number_format = MULTIPLE_FORMAT

    row = 45; rows['exit_year'] = row
    set_label(ws, row, "Exit Year", "#", "Hold period")
    apply_input_style(ws, row, 4, 10)
    ws.cell(row=row, column=4).number_format = INTEGER_FORMAT

    row = 47
    # Tax Inputs
    rows['input_header_tax'] = row
    apply_header_style(ws, row, 1, "TAX INPUTS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 48; rows['tax_rate'] = row
    set_label(ws, row, "Tax Rate", "%", "Blended federal + state")
    apply_input_style(ws, row, 4, 0.25)
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 50
    # ========================================================================
    # CALCULATED ITEMS
    # ========================================================================
    rows['calc_header'] = row
    apply_header_style(ws, row, 1, "CALCULATED ITEMS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 51; rows['const_loan'] = row
    set_label(ws, row, "Construction Loan", "$mm", "Const Cost × Const Debt %")
    apply_calc_style(ws, row, 4, f"=$D${rows['construction_cost']}*$D${rows['const_debt_pct']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 52; rows['const_interest'] = row
    set_label(ws, row, "Construction Interest", "$mm", "Capitalized into RAB")
    apply_calc_style(ws, row, 4, f"=$D${rows['const_loan']}*$D${rows['interest_rate']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 53; rows['rab_at_cod'] = row
    set_label(ws, row, "RAB at COD", "$mm", "Dev + Const + Int")
    apply_calc_style(ws, row, 4, f"=$D${rows['dev_cost_y0']}+$D${rows['dev_cost_y1']}+$D${rows['construction_cost']}+$D${rows['const_interest']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 54; rows['term_debt'] = row
    set_label(ws, row, "Term Debt", "$mm", "RAB × Term Debt %")
    apply_calc_style(ws, row, 4, f"=$D${rows['rab_at_cod']}*$D${rows['term_debt_pct_rab']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 55; rows['annual_depreciation'] = row
    set_label(ws, row, "Annual Depreciation", "$mm", "RAB / Asset Life")
    apply_calc_style(ws, row, 4, f"=$D${rows['rab_at_cod']}/$D${rows['asset_life']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 57
    # ========================================================================
    # CONSTRUCTION PHASE (Y0-Y2)
    # ========================================================================
    rows['const_phase_header'] = row
    apply_header_style(ws, row, 1, "CONSTRUCTION PHASE (Y0-Y2)")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 58; rows['dev_equity_y0'] = row
    set_label(ws, row, "Development Equity Y0", "$mm")
    apply_warning_style(ws, row, 4, f"=$D${rows['dev_cost_y0']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 59; rows['dev_equity_y1'] = row
    set_label(ws, row, "Development Equity Y1", "$mm")
    apply_warning_style(ws, row, 5, f"=$D${rows['dev_cost_y1']}")
    ws.cell(row=row, column=5).number_format = CURRENCY_FORMAT

    row = 60; rows['const_equity_y2'] = row
    set_label(ws, row, "Construction Equity Y2", "$mm", "Const × (1-Debt%)")
    apply_warning_style(ws, row, 6, f"=$D${rows['construction_cost']}*(1-$D${rows['const_debt_pct']})")
    ws.cell(row=row, column=6).number_format = CURRENCY_FORMAT

    row = 61; rows['refi_equity_y3'] = row
    set_label(ws, row, "Refinancing Equity Y3", "$mm", "Const Loan - Term Debt")
    # At COD (Y3), we refinance: pay off construction loan with term debt + equity
    apply_warning_style(ws, row, 7, f"=($D${rows['const_loan']}+$D${rows['const_interest']})-$D${rows['term_debt']}")
    ws.cell(row=row, column=7).number_format = CURRENCY_FORMAT

    row = 62; rows['total_equity'] = row
    set_label(ws, row, "Total Equity Invested", "$mm")
    apply_output_style(ws, row, 4, f"=D{rows['dev_equity_y0']}+E{rows['dev_equity_y1']}+F{rows['const_equity_y2']}+G{rows['refi_equity_y3']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 64
    # ========================================================================
    # OPERATING MODEL (Y3+ only)
    # ========================================================================
    rows['op_header'] = row
    apply_header_style(ws, row, 1, "OPERATING MODEL (Y3+ Post-COD)")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 65; rows['tariff'] = row
    set_label(ws, row, "Tariff Rate", "$/kW-yr", "Escalated")
    # Y1-Y2: No revenue (construction), Y3+: Tariff with escalation
    for col in range(5, 7):  # Y1-Y2
        apply_calc_style(ws, row, col, "=0")
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT
    for col in range(7, 15):  # Y3-Y10
        yr = col - 4
        # Escalation from Y3: (year - 3) years of escalation
        formula = f"=$D${rows['tariff_rate']}*(1+$D${rows['tariff_escalation']})^({yr}-3)"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 67; rows['revenue'] = row
    set_label(ws, row, "Revenue", "$mm", "Cap × Tariff × Avail × 1000/1M")
    for col in range(5, 15):
        # Revenue = Transfer Capacity (MW) × Tariff ($/kW-yr) × Availability × 1000 kW/MW / 1,000,000
        formula = f"=$D${rows['transfer_capacity']}*{col_letter(col)}{rows['tariff']}*$D${rows['availability']}*1000/1000000"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 68; rows['om_cost'] = row
    set_label(ws, row, "O&M Cost", "$mm", "Post-COD only")
    for col in range(5, 7):  # Y1-Y2
        apply_calc_style(ws, row, col, "=0")
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT
    for col in range(7, 15):  # Y3-Y10
        yr = col - 4
        formula = f"=$D${rows['om_y3']}*(1+$D${rows['om_escalation']})^({yr}-3)"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 70; rows['ebitda'] = row
    set_label(ws, row, "EBITDA", "$mm")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['revenue']}-{col_letter(col)}{rows['om_cost']}"
        apply_output_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 72
    # ========================================================================
    # CASH FLOW
    # ========================================================================
    rows['cf_header'] = row
    apply_header_style(ws, row, 1, "CASH FLOW")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 73; rows['depreciation_yr'] = row
    set_label(ws, row, "Depreciation", "$mm", "Post-COD only")
    for col in range(5, 7):  # Y1-Y2
        apply_calc_style(ws, row, col, "=0")
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT
    for col in range(7, 15):  # Y3-Y10
        formula = f"=$D${rows['annual_depreciation']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 74; rows['ebit'] = row
    set_label(ws, row, "EBIT", "$mm", "EBITDA - Depreciation")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['ebitda']}-{col_letter(col)}{rows['depreciation_yr']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 75; rows['taxes'] = row
    set_label(ws, row, "Taxes", "$mm", "MAX(0, EBIT) × Tax Rate")
    for col in range(5, 15):
        formula = f"=MAX(0,{col_letter(col)}{rows['ebit']})*$D${rows['tax_rate']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 76; rows['cfads'] = row
    set_label(ws, row, "CFADS", "$mm", "EBITDA - Taxes")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['ebitda']}-{col_letter(col)}{rows['taxes']}"
        apply_output_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 78
    # ========================================================================
    # DEBT SCHEDULE (Sculpted Y3+, no debt Y1-Y2)
    # ========================================================================
    rows['debt_header'] = row
    apply_header_style(ws, row, 1, "DEBT SCHEDULE (Sculpted 1.40x, Y3+ Only)")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 79; rows['debt_beg_bal'] = row
    set_label(ws, row, "Beginning Balance", "$mm")
    # Y1-Y2: No term debt, Y3: Term debt starts
    for col in range(5, 7):  # Y1-Y2
        apply_calc_style(ws, row, col, "=0")
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT
    apply_calc_style(ws, row, 7, f"=$D${rows['term_debt']}")  # Y3
    ws.cell(row=row, column=7).number_format = CURRENCY_FORMAT
    for col in range(8, 15):  # Y4-Y10
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 80; rows['interest'] = row
    set_label(ws, row, "Interest", "$mm", "Beg Bal × Int Rate")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['debt_beg_bal']}*$D${rows['interest_rate']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 81; rows['target_ds'] = row
    set_label(ws, row, "Target Debt Service", "$mm", "CFADS / Target DSCR")
    for col in range(5, 15):
        formula = f"=IF({col_letter(col)}{rows['cfads']}>0,{col_letter(col)}{rows['cfads']}/$D${rows['target_dscr']},0)"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 82; rows['sculpted_principal'] = row
    set_label(ws, row, "Sculpted Principal", "$mm", "Target DS - Interest")
    for col in range(5, 15):
        formula = f"=MIN(MAX(0,{col_letter(col)}{rows['target_ds']}-{col_letter(col)}{rows['interest']}),{col_letter(col)}{rows['debt_beg_bal']})"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 83; rows['actual_ds'] = row
    set_label(ws, row, "Actual Debt Service", "$mm", "Interest + Principal")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['interest']}+{col_letter(col)}{rows['sculpted_principal']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 84; rows['debt_end_bal'] = row
    set_label(ws, row, "Ending Balance", "$mm", "Beg - Principal")
    for col in range(5, 15):
        formula = f"=MAX(0,{col_letter(col)}{rows['debt_beg_bal']}-{col_letter(col)}{rows['sculpted_principal']})"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    # Fix beginning balance references for Y4+
    for col in range(8, 15):
        prev_col = col_letter(col - 1)
        formula = f"={prev_col}{rows['debt_end_bal']}"
        apply_calc_style(ws, rows['debt_beg_bal'], col, formula)

    row = 85; rows['actual_dscr'] = row
    set_label(ws, row, "Actual DSCR", "x", "CFADS / Actual DS")
    for col in range(5, 15):
        formula = f"=IF({col_letter(col)}{rows['actual_ds']}>0,{col_letter(col)}{rows['cfads']}/{col_letter(col)}{rows['actual_ds']},0)"
        apply_output_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = MULTIPLE_FORMAT

    row = 87
    # ========================================================================
    # DSRA
    # ========================================================================
    rows['dsra_header'] = row
    apply_header_style(ws, row, 1, "DSRA (Debt Service Reserve Account)")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 88; rows['dsra_target'] = row
    set_label(ws, row, "DSRA Target", "$mm", "Next Yr DS × DSRA Months / 12")
    # Y0-Y2: Target based on Y3 DS, Y3+: Next year DS
    for col in range(4, 7):  # Y0-Y2
        formula = f"=G{rows['actual_ds']}*$D${rows['dsra_months']}/12"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT
    for col in range(7, 14):  # Y3-Y9
        next_col = col_letter(col + 1)
        formula = f"={next_col}{rows['actual_ds']}*$D${rows['dsra_months']}/12"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT
    apply_calc_style(ws, row, 14, "=0")  # Y10
    ws.cell(row=row, column=14).number_format = CURRENCY_FORMAT

    row = 89; rows['dsra_beg'] = row
    set_label(ws, row, "DSRA Beginning", "$mm")
    apply_calc_style(ws, row, 4, "=0")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT
    for col in range(5, 15):
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 90; rows['dsra_funding'] = row
    set_label(ws, row, "DSRA Funding/(Release)", "$mm", "Target - Beginning")
    for col in range(4, 15):
        formula = f"={col_letter(col)}{rows['dsra_target']}-{col_letter(col)}{rows['dsra_beg']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 91; rows['dsra_end'] = row
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

    row = 93
    # ========================================================================
    # DISTRIBUTION WATERFALL
    # ========================================================================
    rows['dist_header'] = row
    apply_header_style(ws, row, 1, "DISTRIBUTION WATERFALL")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 94; rows['cash_available'] = row
    set_label(ws, row, "Cash Available for Distribution", "$mm", "CFADS - DS - DSRA Fund")
    for col in range(5, 15):
        formula = f"={col_letter(col)}{rows['cfads']}-{col_letter(col)}{rows['actual_ds']}-{col_letter(col)}{rows['dsra_funding']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 95; rows['lockup_test'] = row
    set_label(ws, row, "Lock-up Test", "", "PASS if DSCR ≥ Lock-up")
    for col in range(5, 15):
        formula = f'=IF({col_letter(col)}{rows["actual_dscr"]}>=$D${rows["lockup_dscr"]},"PASS","LOCKED")'
        apply_calc_style(ws, row, col, formula)

    row = 96; rows['distributable'] = row
    set_label(ws, row, "Distributable Cash", "$mm", "If PASS, MAX(0, Cash Avail)")
    for col in range(5, 15):
        formula = f'=IF({col_letter(col)}{rows["lockup_test"]}="PASS",MAX(0,{col_letter(col)}{rows["cash_available"]}),0)'
        apply_output_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT

    row = 98
    # ========================================================================
    # SOURCES & USES (Construction Phase Funding)
    # ========================================================================
    rows['su_header'] = row
    apply_header_style(ws, row, 1, "SOURCES & USES (Total Construction)")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 99; rows['uses_header'] = row
    set_label(ws, row, "USES", "", "")
    ws.cell(row=row, column=1).font = BLACK_BOLD

    row = 100; rows['use_dev_y0'] = row
    set_label(ws, row, "  Development Y0", "$mm")
    apply_calc_style(ws, row, 4, f"=$D${rows['dev_cost_y0']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 101; rows['use_dev_y1'] = row
    set_label(ws, row, "  Development Y1", "$mm")
    apply_calc_style(ws, row, 4, f"=$D${rows['dev_cost_y1']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 102; rows['use_construction'] = row
    set_label(ws, row, "  Construction Y2", "$mm")
    apply_calc_style(ws, row, 4, f"=$D${rows['construction_cost']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 103; rows['use_const_interest'] = row
    set_label(ws, row, "  Construction Interest", "$mm")
    apply_calc_style(ws, row, 4, f"=$D${rows['const_interest']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 104; rows['use_dsra'] = row
    set_label(ws, row, "  DSRA Funding (at COD)", "$mm")
    apply_calc_style(ws, row, 4, f"=G{rows['dsra_target']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 105; rows['total_uses'] = row
    set_label(ws, row, "Total Uses", "$mm")
    apply_output_style(ws, row, 4, f"=SUM(D{rows['use_dev_y0']}:D{rows['use_dsra']})")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 107; rows['sources_header'] = row
    set_label(ws, row, "SOURCES", "", "")
    ws.cell(row=row, column=1).font = BLACK_BOLD

    row = 108; rows['source_const_loan'] = row
    set_label(ws, row, "  Construction Loan", "$mm")
    apply_calc_style(ws, row, 4, f"=$D${rows['const_loan']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 109; rows['source_term_debt'] = row
    set_label(ws, row, "  Term Debt (at COD)", "$mm")
    apply_calc_style(ws, row, 4, f"=$D${rows['term_debt']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 110; rows['source_equity'] = row
    set_label(ws, row, "  Total Equity", "$mm")
    apply_calc_style(ws, row, 4, f"=D{rows['total_equity']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 111; rows['total_sources'] = row
    set_label(ws, row, "Total Sources", "$mm")
    # Note: Sources = Const Loan (repaid at COD) + Term Debt + Equity
    # Uses = Dev + Const + Interest + DSRA
    # At refinancing: Const Loan + Interest repaid by Term Debt + Refi Equity
    apply_output_style(ws, row, 4, f"=D{rows['source_term_debt']}+D{rows['source_equity']}+D{rows['use_dsra']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT

    row = 112; rows['su_check'] = row
    set_label(ws, row, "Balance Check", "", "")
    # Check: Total Uses = RAB + DSRA = Term Debt + Equity + DSRA
    apply_output_style(ws, row, 4, f'=IF(ABS(D{rows["rab_at_cod"]}+D{rows["use_dsra"]}-D{rows["total_sources"]})<0.01,"PASS","FAIL")')

    row = 114
    # ========================================================================
    # EXIT CALCULATIONS
    # ========================================================================
    rows['exit_header'] = row
    apply_header_style(ws, row, 1, "EXIT CALCULATIONS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 115; rows['cumulative_depreciation'] = row
    set_label(ws, row, "Cumulative Depreciation", "$mm", "Dep × Operating Years")
    # Exit Y10, operating years = Y3-Y10 = 8 years
    apply_calc_style(ws, row, 14, f"=$D${rows['annual_depreciation']}*8")
    ws.cell(row=row, column=14).number_format = CURRENCY_FORMAT

    row = 116; rows['exit_rab'] = row
    set_label(ws, row, "Exit RAB", "$mm", "RAB at COD - Cum Dep")
    apply_calc_style(ws, row, 14, f"=$D${rows['rab_at_cod']}-N{rows['cumulative_depreciation']}")
    ws.cell(row=row, column=14).number_format = CURRENCY_FORMAT

    row = 117; rows['exit_ev'] = row
    set_label(ws, row, "Exit EV", "$mm", "Exit RAB × xRAB Multiple")
    apply_calc_style(ws, row, 14, f"=N{rows['exit_rab']}*$D${rows['exit_xrab']}")
    ws.cell(row=row, column=14).number_format = CURRENCY_FORMAT

    row = 118; rows['exit_debt'] = row
    set_label(ws, row, "Exit Debt Payoff", "$mm", "Ending Bal at Exit")
    apply_calc_style(ws, row, 14, f"=N{rows['debt_end_bal']}")
    ws.cell(row=row, column=14).number_format = CURRENCY_FORMAT

    row = 119; rows['exit_dsra_release'] = row
    set_label(ws, row, "DSRA Release", "$mm", "DSRA Ending at Exit")
    apply_calc_style(ws, row, 14, f"=N{rows['dsra_end']}")
    ws.cell(row=row, column=14).number_format = CURRENCY_FORMAT

    row = 120; rows['exit_equity_proceeds'] = row
    set_label(ws, row, "Exit Equity Proceeds", "$mm", "Exit EV - Debt + DSRA")
    apply_output_style(ws, row, 14, f"=N{rows['exit_ev']}-N{rows['exit_debt']}+N{rows['exit_dsra_release']}")
    ws.cell(row=row, column=14).number_format = CURRENCY_FORMAT

    row = 122
    # ========================================================================
    # EQUITY RETURNS
    # ========================================================================
    rows['returns_header'] = row
    apply_header_style(ws, row, 1, "EQUITY RETURNS")
    for col in range(2, 15):
        ws.cell(row=row, column=col).fill = NAVY_FILL

    row = 123; rows['equity_cf'] = row
    set_label(ws, row, "Equity Cash Flow", "$mm")
    # Y0: -Dev Equity Y0
    apply_calc_style(ws, row, 4, f"=-D{rows['dev_equity_y0']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT
    # Y1: -Dev Equity Y1
    apply_calc_style(ws, row, 5, f"=-E{rows['dev_equity_y1']}")
    ws.cell(row=row, column=5).number_format = CURRENCY_FORMAT
    # Y2: -Const Equity Y2
    apply_calc_style(ws, row, 6, f"=-F{rows['const_equity_y2']}")
    ws.cell(row=row, column=6).number_format = CURRENCY_FORMAT
    # Y3: -Refi Equity + Distributable - DSRA initial funding
    apply_calc_style(ws, row, 7, f"=-G{rows['refi_equity_y3']}-G{rows['dsra_funding']}+G{rows['distributable']}")
    ws.cell(row=row, column=7).number_format = CURRENCY_FORMAT
    # Y4-Y9: Distributable
    for col in range(8, 14):
        formula = f"={col_letter(col)}{rows['distributable']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT
    # Y10: Distributable + Exit Equity Proceeds
    apply_calc_style(ws, row, 14, f"=N{rows['distributable']}+N{rows['exit_equity_proceeds']}")
    ws.cell(row=row, column=14).number_format = CURRENCY_FORMAT

    row = 124; rows['levered_irr'] = row
    set_label(ws, row, "Levered IRR", "%")
    apply_output_style(ws, row, 4, f"=IRR(D{rows['equity_cf']}:N{rows['equity_cf']})")
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    row = 125; rows['moic'] = row
    set_label(ws, row, "MOIC", "x", "Sum inflows / outflows")
    apply_output_style(ws, row, 4, f"=SUMIF(D{rows['equity_cf']}:N{rows['equity_cf']},\">0\")/-SUMIF(D{rows['equity_cf']}:N{rows['equity_cf']},\"<0\")")
    ws.cell(row=row, column=4).number_format = MULTIPLE_FORMAT

    row = 127; rows['unlev_header'] = row
    set_label(ws, row, "Unlevered Returns (for reference)", "")
    ws.cell(row=row, column=1).font = BLACK_BOLD

    row = 128; rows['unlev_cf'] = row
    set_label(ws, row, "Unlevered Cash Flow", "$mm")
    # Y0: -Dev Cost Y0
    apply_calc_style(ws, row, 4, f"=-$D${rows['dev_cost_y0']}")
    ws.cell(row=row, column=4).number_format = CURRENCY_FORMAT
    # Y1: -Dev Cost Y1
    apply_calc_style(ws, row, 5, f"=-$D${rows['dev_cost_y1']}")
    ws.cell(row=row, column=5).number_format = CURRENCY_FORMAT
    # Y2: -Const Cost
    apply_calc_style(ws, row, 6, f"=-$D${rows['construction_cost']}")
    ws.cell(row=row, column=6).number_format = CURRENCY_FORMAT
    # Y3-Y9: CFADS
    for col in range(7, 14):
        formula = f"={col_letter(col)}{rows['cfads']}"
        apply_calc_style(ws, row, col, formula)
        ws.cell(row=row, column=col).number_format = CURRENCY_FORMAT
    # Y10: CFADS + Exit EV
    apply_calc_style(ws, row, 14, f"=N{rows['cfads']}+N{rows['exit_ev']}")
    ws.cell(row=row, column=14).number_format = CURRENCY_FORMAT

    row = 129; rows['unlev_irr'] = row
    set_label(ws, row, "Unlevered IRR", "%")
    apply_output_style(ws, row, 4, f"=IRR(D{rows['unlev_cf']}:N{rows['unlev_cf']})")
    ws.cell(row=row, column=4).number_format = PERCENT_FORMAT

    # ========================================================================
    # POPULATE AUDIT STRIP
    # ========================================================================
    apply_output_style(ws, rows['audit_rab_cod'], 4, f"=$D${rows['rab_at_cod']}")
    ws.cell(row=rows['audit_rab_cod'], column=4).number_format = CURRENCY_FORMAT

    apply_output_style(ws, rows['audit_y3_ebitda'], 4, f"=G{rows['ebitda']}")
    ws.cell(row=rows['audit_y3_ebitda'], column=4).number_format = CURRENCY_FORMAT

    apply_output_style(ws, rows['audit_avg_cfads'], 4, f"=AVERAGE(G{rows['cfads']}:N{rows['cfads']})")
    ws.cell(row=rows['audit_avg_cfads'], column=4).number_format = CURRENCY_FORMAT

    apply_output_style(ws, rows['audit_min_dscr'], 4, f"=MIN(G{rows['actual_dscr']}:N{rows['actual_dscr']})")
    ws.cell(row=rows['audit_min_dscr'], column=4).number_format = MULTIPLE_FORMAT

    apply_output_style(ws, rows['audit_levered_irr'], 4, f"=D{rows['levered_irr']}")
    ws.cell(row=rows['audit_levered_irr'], column=4).number_format = PERCENT_FORMAT

    apply_output_style(ws, rows['audit_moic'], 4, f"=D{rows['moic']}")
    ws.cell(row=rows['audit_moic'], column=4).number_format = MULTIPLE_FORMAT

    apply_output_style(ws, rows['audit_unlevered_irr'], 4, f"=D{rows['unlev_irr']}")
    ws.cell(row=rows['audit_unlevered_irr'], column=4).number_format = PERCENT_FORMAT

    apply_output_style(ws, rows['qa_su_balance'], 4, f"=D{rows['su_check']}")
    apply_output_style(ws, rows['qa_debt_payoff'], 4, f'=IF(N{rows["debt_end_bal"]}<1,"PASS","FAIL")')
    apply_output_style(ws, rows['qa_min_dscr_check'], 4, f'=IF(D{rows["audit_min_dscr"]}>=1.35,"PASS","FAIL")')

    # Freeze panes
    ws.freeze_panes = 'E2'

    wb.save(output_path)
    print(f"Transmission model saved to: {output_path}")
    return rows

if __name__ == "__main__":
    output_path = "/mnt/user-data/outputs/Model_4_Transmission.xlsx"
    rows = build_transmission_model(output_path)
    print("Transmission model built successfully!")
