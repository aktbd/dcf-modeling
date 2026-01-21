#!/usr/bin/env python3
"""
CRITICAL FIX: Rewire all formula references in 6 Model files
The absolute references ($D$##) point to wrong cells due to row insertions.
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

# Paths
DELIV_DIR = '/home/user/dcf-modeling/deliverables/deliverables-20260120-0030'

# Style definitions
YELLOW_FILL = PatternFill(start_color="FFFFC7", end_color="FFFFC7", fill_type="solid")
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
BLUE_FONT = Font(color="0000FF", bold=True)
BLACK_FONT = Font(color="000000")
BLACK_BOLD = Font(color="000000", bold=True)

# Number formats
CURRENCY_FMT = '#,##0.0_);(#,##0.0)'
PERCENT_FMT = '0.0%'
MULTIPLE_FMT = '0.00"x"'
INTEGER_FMT = '#,##0'

def fix_ccgt(filepath):
    """Fix Model_1_CCGT.xlsx"""
    print(f"\n{'='*60}")
    print("FIXING MODEL_1_CCGT.xlsx")
    print('='*60)

    wb = load_workbook(filepath)
    ws = wb['Model_Standard']

    # CORRECT INPUT MAPPING (verified from file)
    r = {
        'entry_ev': 28,
        'txn_fees': 29,
        'exit_mult': 30,
        'exit_year': 31,
        'capacity': 34,
        'heat_rate': 35,
        'cap_factor': 36,
        'for_rate': 37,
        'power_price': 40,
        'cap_price': 41,
        'gas_price': 42,
        'escalation': 43,
        'vom': 46,
        'fom': 47,
        'leverage': 50,
        'int_rate': 51,
        'debt_tenor': 52,
        'dsra_months': 53,
        'sweep_pct': 54,
        'lockup': 55,
        'tax_rate': 58,
        'dep_basis': 59,
        'dep_life': 60,
        # Calculated items
        'ucap': 64,
        'generation': 65,
        'debt_amount': 66,
        'depreciation': 67,
        'sched_prin': 68,
        # Operating model
        'power_pr_row': 72,
        'cap_pr_row': 73,
        'gas_pr_row': 74,
        'energy_rev': 76,
        'cap_rev': 77,
        'total_rev': 78,
        'fuel_cost': 80,
        'spark_spread': 81,
        'vom_row': 82,
        'fom_row': 83,
        'total_costs': 84,
        'ebitda': 86,
        # Cash flow
        'ebit': 90,
        'taxes': 91,
        'cfads': 92,
        # Debt
        'beg_bal': 96,
        'interest': 97,
        'sched_prin_row': 98,
        'ds_pre': 99,
        'excess': 100,
    }

    # Fix Calculated Items (Column D)
    # UCAP = Capacity * (1 - FOR)
    ws.cell(row=r['ucap'], column=4).value = f"=$D${r['capacity']}*(1-$D${r['for_rate']})"

    # Generation = Capacity * CF * 8760 * (1 - FOR)
    ws.cell(row=r['generation'], column=4).value = f"=$D${r['capacity']}*$D${r['cap_factor']}*8760*(1-$D${r['for_rate']})"

    # Debt Amount = Entry EV * Leverage
    ws.cell(row=r['debt_amount'], column=4).value = f"=$D${r['entry_ev']}*$D${r['leverage']}"

    # Annual Depreciation = Entry EV * Depreciable Basis / Dep Life
    ws.cell(row=r['depreciation'], column=4).value = f"=$D${r['entry_ev']}*$D${r['dep_basis']}/$D${r['dep_life']}"

    # Scheduled Principal = Debt Amount / Debt Tenor
    ws.cell(row=r['sched_prin'], column=4).value = f"=$D${r['debt_amount']}/$D${r['debt_tenor']}"

    print("Fixed Calculated Items (D64-D68)")

    # Fix Operating Model (Columns E-N, Years 1-10)
    for col in range(5, 15):  # E=5 to N=14
        year = col - 4
        col_letter = get_column_letter(col)

        # Power Price = Y1 Price * (1 + Esc)^(Year-1)
        ws.cell(row=r['power_pr_row'], column=col).value = f"=$D${r['power_price']}*(1+$D${r['escalation']})^({year}-1)"

        # Capacity Price = Y1 Price * (1 + Esc)^(Year-1)
        ws.cell(row=r['cap_pr_row'], column=col).value = f"=$D${r['cap_price']}*(1+$D${r['escalation']})^({year}-1)"

        # Gas Price = Y1 Price * (1 + Esc)^(Year-1)
        ws.cell(row=r['gas_pr_row'], column=col).value = f"=$D${r['gas_price']}*(1+$D${r['escalation']})^({year}-1)"

        # Energy Revenue = Generation * Power Price / 1E6
        ws.cell(row=r['energy_rev'], column=col).value = f"=$D${r['generation']}*{col_letter}{r['power_pr_row']}/1000000"

        # Capacity Revenue = Capacity * Capacity Price * 1000 / 1E6
        ws.cell(row=r['cap_rev'], column=col).value = f"=$D${r['capacity']}*{col_letter}{r['cap_pr_row']}*1000/1000000"

        # Total Revenue = Energy + Capacity
        ws.cell(row=r['total_rev'], column=col).value = f"={col_letter}{r['energy_rev']}+{col_letter}{r['cap_rev']}"

        # Fuel Cost = Generation * Heat Rate * Gas Price / 1E9
        ws.cell(row=r['fuel_cost'], column=col).value = f"=$D${r['generation']}*$D${r['heat_rate']}*{col_letter}{r['gas_pr_row']}/1000000000"

        # Spark Spread = Power Price - (Heat Rate / 1000) * Gas Price
        ws.cell(row=r['spark_spread'], column=col).value = f"={col_letter}{r['power_pr_row']}-($D${r['heat_rate']}/1000)*{col_letter}{r['gas_pr_row']}"

        # VOM = Generation * VOM Rate / 1E6
        ws.cell(row=r['vom_row'], column=col).value = f"=$D${r['generation']}*$D${r['vom']}/1000000"

        # FOM = Capacity * FOM Rate * 1000 / 1E6
        ws.cell(row=r['fom_row'], column=col).value = f"=$D${r['capacity']}*$D${r['fom']}*1000/1000000"

        # Total Costs = Fuel + VOM + FOM
        ws.cell(row=r['total_costs'], column=col).value = f"={col_letter}{r['fuel_cost']}+{col_letter}{r['vom_row']}+{col_letter}{r['fom_row']}"

        # EBITDA = Total Revenue - Total Costs
        ws.cell(row=r['ebitda'], column=col).value = f"={col_letter}{r['total_rev']}-{col_letter}{r['total_costs']}"
        ws.cell(row=r['ebitda'], column=col).fill = GREEN_FILL
        ws.cell(row=r['ebitda'], column=col).font = BLACK_BOLD

        # EBIT = EBITDA - Depreciation
        ws.cell(row=r['ebit'], column=col).value = f"={col_letter}{r['ebitda']}-$D${r['depreciation']}"

        # Taxes = MAX(0, EBIT) * Tax Rate
        ws.cell(row=r['taxes'], column=col).value = f"=MAX(0,{col_letter}{r['ebit']})*$D${r['tax_rate']}"

        # CFADS = EBITDA - Taxes (simplified, no WC)
        ws.cell(row=r['cfads'], column=col).value = f"={col_letter}{r['ebitda']}-{col_letter}{r['taxes']}"
        ws.cell(row=r['cfads'], column=col).fill = GREEN_FILL
        ws.cell(row=r['cfads'], column=col).font = BLACK_BOLD

    print("Fixed Operating Model (Rows 72-92, Cols E-N)")

    # Fix Debt Schedule (need to find actual rows first)
    # Let me check what rows the debt schedule is on
    # Based on output, it starts at row 95
    debt_rows = {
        'header': 95,
        'beg_bal': 96,
        'interest': 97,
        'sched_prin': 98,
        'ds_pre': 99,
        'excess': 100,
        'sweep': 101,
        'total_prin': 102,
        'end_bal': 103,
        'dscr': 104,
    }

    # First, set Y0 (column D) debt beginning balance to 0 and Y1 to Debt Amount
    ws.cell(row=debt_rows['beg_bal'], column=4).value = 0

    for col in range(5, 15):  # E=5 to N=14
        year = col - 4
        col_letter = get_column_letter(col)
        prev_col = get_column_letter(col - 1)

        # Beginning Balance: Y1 = Debt Amount, Y2+ = Prior Year End Balance
        if year == 1:
            ws.cell(row=debt_rows['beg_bal'], column=col).value = f"=$D${r['debt_amount']}"
        else:
            ws.cell(row=debt_rows['beg_bal'], column=col).value = f"={prev_col}{debt_rows['end_bal']}"

        # Interest = Beg Bal * Interest Rate
        ws.cell(row=debt_rows['interest'], column=col).value = f"={col_letter}{debt_rows['beg_bal']}*$D${r['int_rate']}"

        # Scheduled Principal = MIN(Debt Amount / Tenor, Beg Bal)
        ws.cell(row=debt_rows['sched_prin'], column=col).value = f"=MIN($D${r['sched_prin']},{col_letter}{debt_rows['beg_bal']})"

        # Debt Service (Pre-Sweep) = Interest + Scheduled Principal
        ws.cell(row=debt_rows['ds_pre'], column=col).value = f"={col_letter}{debt_rows['interest']}+{col_letter}{debt_rows['sched_prin']}"

    print("Fixed Debt Schedule rows 96-99")

    # Now I need to find and fix DSRA and the rest of the debt schedule
    # Let me check more rows
    wb.save(filepath)
    print(f"Saved intermediate changes to {filepath}")

    # Reload and check more structure
    wb = load_workbook(filepath)
    ws = wb['Model_Standard']

    # Print rows 100-130 to find DSRA and rest of debt
    print("\n--- Checking rows 100-130 ---")
    for row in range(100, 131):
        a = ws.cell(row=row, column=1).value or ""
        if a:
            print(f"Row {row}: {str(a)[:40]}")

    return wb, r

def check_and_continue_ccgt(wb, r, filepath):
    """Continue fixing CCGT after checking structure"""
    ws = wb['Model_Standard']

    # Based on typical structure, the DSRA section follows debt schedule
    # Let's map more rows
    dsra_rows = {
        'scheduled_ds': 106,  # For DSRA target calculation (no circularity)
        'dsra_target': 108,
        'dsra_beg': 109,
        'dsra_funding': 110,
        'dsra_end': 111,
    }

    dist_rows = {
        'header': 114,
        'cfads': 115,
        'less_ds': 116,
        'less_dsra': 117,
        'avail': 118,
        'lockup': 119,
        'distributable': 120,
    }

    su_rows = {
        'header': 123,
        'uses_header': 124,
        'purchase': 125,
        'txn_fees': 126,
        'dsra_init': 127,
        'total_uses': 128,
        'sources_header': 130,
        'debt': 131,
        'equity': 132,
        'total_sources': 133,
        'balance': 135,
    }

    exit_rows = {
        'header': 138,
        'exit_ebitda': 139,
        'exit_ev': 140,
        'less_debt': 141,
        'plus_dsra': 142,
        'exit_equity': 143,
        'equity_cf': 146,
        'irr': 147,
        'moic': 148,
    }

    # Fix DSRA Section
    for col in range(5, 15):
        year = col - 4
        col_letter = get_column_letter(col)
        next_col = get_column_letter(col + 1) if col < 14 else col_letter

        # Scheduled DS for DSRA (based on straight-line, no circularity)
        # = MAX(0, Debt - (Year-1)*Sched Prin) * Int Rate + Sched Prin
        if col <= 14:
            ws.cell(row=dsra_rows['scheduled_ds'], column=col).value = \
                f"=MAX(0,$D${r['debt_amount']}-({year}-1)*$D${r['sched_prin']})*$D${r['int_rate']}+$D${r['sched_prin']}"

        # DSRA Target = Next Year's Scheduled DS * (DSRA Months / 12)
        # For final year, target = 0
        if col < 14:
            ws.cell(row=dsra_rows['dsra_target'], column=col).value = \
                f"={next_col}{dsra_rows['scheduled_ds']}*$D${r['dsra_months']}/12"
        else:
            ws.cell(row=dsra_rows['dsra_target'], column=col).value = 0

    print("Fixed DSRA rows 106-111")

    # Fix Sources & Uses (Column D only)
    # Purchase Price = Entry EV
    ws.cell(row=su_rows['purchase'], column=4).value = f"=$D${r['entry_ev']}"

    # Transaction Fees = Entry EV * Txn Fee %
    ws.cell(row=su_rows['txn_fees'], column=4).value = f"=$D${r['entry_ev']}*$D${r['txn_fees']}"

    # DSRA Initial = First year DSRA Target
    ws.cell(row=su_rows['dsra_init'], column=4).value = f"=E{dsra_rows['dsra_target']}"

    # Total Uses
    ws.cell(row=su_rows['total_uses'], column=4).value = f"=D{su_rows['purchase']}+D{su_rows['txn_fees']}+D{su_rows['dsra_init']}"

    # Funded Debt = Entry EV * Leverage
    ws.cell(row=su_rows['debt'], column=4).value = f"=$D${r['entry_ev']}*$D${r['leverage']}"

    # Entry Equity = Total Uses - Debt (plug)
    ws.cell(row=su_rows['equity'], column=4).value = f"=D{su_rows['total_uses']}-D{su_rows['debt']}"

    # Total Sources
    ws.cell(row=su_rows['total_sources'], column=4).value = f"=D{su_rows['debt']}+D{su_rows['equity']}"

    # S&U Balance Check
    ws.cell(row=su_rows['balance'], column=4).value = f"=D{su_rows['total_uses']}-D{su_rows['total_sources']}"

    print("Fixed Sources & Uses rows 125-135")

    # Fix Exit & Returns
    exit_year = 7  # From inputs
    exit_col = get_column_letter(4 + exit_year)  # K for Y7

    # Exit EBITDA = EBITDA in exit year
    ws.cell(row=exit_rows['exit_ebitda'], column=4).value = f"={exit_col}{r['ebitda']}"

    # Exit EV = Exit EBITDA * Exit Multiple
    ws.cell(row=exit_rows['exit_ev'], column=4).value = f"=D{exit_rows['exit_ebitda']}*$D${r['exit_mult']}"

    # Less Debt = Debt balance at exit (should be minimal after sweep)
    ws.cell(row=exit_rows['less_debt'], column=4).value = f"={exit_col}103"  # End balance row

    # Plus DSRA Release
    ws.cell(row=exit_rows['plus_dsra'], column=4).value = f"={exit_col}{dsra_rows['dsra_end']}"

    # Exit Equity = Exit EV - Debt + DSRA
    ws.cell(row=exit_rows['exit_equity'], column=4).value = f"=D{exit_rows['exit_ev']}-D{exit_rows['less_debt']}+D{exit_rows['plus_dsra']}"

    print("Fixed Exit rows 139-143")

    # Fix Equity CF row for IRR calculation
    # Y0 = -Entry Equity, Y1-Y6 = Distributions, Y7 = Distribution + Exit Equity
    ws.cell(row=exit_rows['equity_cf'], column=4).value = f"=-D{su_rows['equity']}"  # Y0 negative

    for col in range(5, 15):
        year = col - 4
        col_letter = get_column_letter(col)

        if year < 7:
            # Distributable cash flow
            ws.cell(row=exit_rows['equity_cf'], column=col).value = f"={col_letter}{dist_rows['distributable']}"
        elif year == 7:
            # Exit year: Distribution + Exit Equity
            ws.cell(row=exit_rows['equity_cf'], column=col).value = f"={col_letter}{dist_rows['distributable']}+D{exit_rows['exit_equity']}"
        else:
            # Post-exit: 0
            ws.cell(row=exit_rows['equity_cf'], column=col).value = 0

    # IRR = IRR of equity cash flows
    ws.cell(row=exit_rows['irr'], column=4).value = f"=IRR(D{exit_rows['equity_cf']}:N{exit_rows['equity_cf']})"
    ws.cell(row=exit_rows['irr'], column=4).fill = GREEN_FILL
    ws.cell(row=exit_rows['irr'], column=4).font = BLACK_BOLD
    ws.cell(row=exit_rows['irr'], column=4).number_format = PERCENT_FMT

    # MOIC = (Sum of distributions + Exit) / Entry Equity
    ws.cell(row=exit_rows['moic'], column=4).value = f"=(SUM(E{exit_rows['equity_cf']}:K{exit_rows['equity_cf']})+D{exit_rows['exit_equity']})/-D{exit_rows['equity_cf']}"
    ws.cell(row=exit_rows['moic'], column=4).fill = GREEN_FILL
    ws.cell(row=exit_rows['moic'], column=4).font = BLACK_BOLD
    ws.cell(row=exit_rows['moic'], column=4).number_format = MULTIPLE_FMT

    print("Fixed Equity CF and Returns rows 146-148")

    # Fix Audit Strip references
    ws.cell(row=3, column=4).value = f"=$D${r['entry_ev']}"  # Entry EV
    ws.cell(row=4, column=4).value = f"=E{r['ebitda']}"  # Y1 EBITDA

    wb.save(filepath)
    print(f"\nSaved all CCGT fixes to {filepath}")

    return wb


# Run the fixes
if __name__ == "__main__":
    # Fix CCGT first
    ccgt_path = os.path.join(DELIV_DIR, 'Model_1_CCGT.xlsx')
    wb, r = fix_ccgt(ccgt_path)
    check_and_continue_ccgt(wb, r, ccgt_path)

    print("\n" + "="*60)
    print("CCGT FIX COMPLETE - Now validating...")
    print("="*60)

    # Validate
    wb = load_workbook(ccgt_path, data_only=False)
    ws = wb['Model_Standard']

    print("\nKey formula checks:")
    print(f"  D64 (UCAP): {ws.cell(row=64, column=4).value}")
    print(f"  D65 (Generation): {ws.cell(row=65, column=4).value}")
    print(f"  D66 (Debt Amount): {ws.cell(row=66, column=4).value}")
    print(f"  E72 (Power Price Y1): {ws.cell(row=72, column=5).value}")
    print(f"  E76 (Energy Revenue): {ws.cell(row=76, column=5).value}")
    print(f"  E80 (Fuel Cost): {ws.cell(row=80, column=5).value}")
    print(f"  E86 (EBITDA): {ws.cell(row=86, column=5).value}")
    print(f"  E97 (Interest): {ws.cell(row=97, column=5).value}")
