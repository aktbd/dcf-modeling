#!/usr/bin/env python3
"""
Complete fix for Model_5_Midstream.xlsx
Key features: Volume growth Y1-Y5, decline Y6+, Fee-based revenue, Cash sweep debt
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

DELIV_DIR = '/home/user/dcf-modeling/deliverables/deliverables-20260120-0030'

GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
BLACK_BOLD = Font(color="000000", bold=True)
PERCENT_FMT = '0.0%'
MULTIPLE_FMT = '0.00"x"'
CURRENCY_FMT = '#,##0.0_);(#,##0.0)'

def fix_midstream():
    filepath = f"{DELIV_DIR}/Model_5_Midstream.xlsx"
    wb = load_workbook(filepath)
    ws = wb['Model_Standard']

    # ===== INPUT ROW MAPPING =====
    inp = {
        'entry_ev': 27,       # 85
        'txn_fees': 28,       # 0.015
        'exit_mult': 29,      # 6.5
        'exit_year': 30,      # 7
        'y1_volume': 33,      # 80
        'growth_rate': 34,    # 0.03
        'decline_rate': 35,   # 0.05
        'total_fee': 38,      # 0.55
        'fee_esc': 39,        # 0.02
        'om_y1': 42,          # 3.5
        'om_esc': 43,         # 0.025
        'growth_capex': 44,   # 2
        'leverage': 47,       # 0.3
        'int_rate': 48,       # 0.07
        'debt_tenor': 49,     # 5
        'dsra_months': 50,    # 6
        'sweep_pct': 51,      # 0.75
        'lockup': 52,         # 1.1
        'tax_rate': 55,       # 0.25
        'dep_basis': 56,      # 0.8
        'dep_life': 57,       # 15
    }

    calc = {
        'debt_amount': 60,
        'depreciation': 61,
        'sched_prin': 62,
    }

    op = {
        'daily_vol': 65,
        'annual_vol': 66,
        'fee': 67,
        'revenue': 69,
        'om': 70,
        'ebitda': 72,
        'taxes': 75,
        'growth_capex': 76,
        'cfads': 77,
    }

    debt = {
        'beg_bal': 81,
        'interest': 82,
        'sched_prin': 83,
        'ds_pre': 84,
        'excess': 85,
        'sweep': 86,
        'total_prin': 87,
        'end_bal': 88,
        'dscr': 89,
        'sched_ds': 90,
    }

    dsra = {
        'target': 94,
        'beg': 95,
        'funding': 96,
        'end': 97,
    }

    dist = {
        'avail': 100,
        'lockup': 101,
        'distributable': 102,
    }

    su = {
        'purchase': 106,
        'txn_fees': 107,
        'dsra_init': 108,
        'total_uses': 109,
        'debt': 112,
        'equity': 113,
        'total_sources': 114,
        'balance': 116,
    }

    exit_r = {
        'exit_ebitda': 119,
        'exit_ev': 120,
        'exit_debt': 121,
        'exit_dsra': 122,
        'exit_equity': 123,
        'equity_cf': 126,
        'irr': 127,
        'moic': 128,
        'unlev_cf': 131,
        'unlev_irr': 132,
    }

    print("="*60)
    print("FIXING MODEL_5_MIDSTREAM.xlsx")
    print("="*60)

    # ===== 1. CALCULATED ITEMS =====
    print("[1] Fixing Calculated Items...")

    ws.cell(row=calc['debt_amount'], column=4).value = f"=$D${inp['entry_ev']}*$D${inp['leverage']}"
    ws.cell(row=calc['depreciation'], column=4).value = f"=$D${inp['entry_ev']}*$D${inp['dep_basis']}/$D${inp['dep_life']}"
    ws.cell(row=calc['sched_prin'], column=4).value = f"=$D${calc['debt_amount']}/$D${inp['debt_tenor']}"

    # ===== 2. OPERATING MODEL =====
    print("[2] Fixing Operating Model...")

    for col in range(5, 15):
        year = col - 4
        c = get_column_letter(col)

        # Volume: Growth Y1-Y5, Decline Y6+
        if year <= 5:
            ws.cell(row=op['daily_vol'], column=col).value = f"=$D${inp['y1_volume']}*(1+$D${inp['growth_rate']})^({year}-1)"
        else:
            # Y6 starts from Y5 volume and declines
            ws.cell(row=op['daily_vol'], column=col).value = \
                f"=$D${inp['y1_volume']}*(1+$D${inp['growth_rate']})^4*(1-$D${inp['decline_rate']})^({year}-5)"

        # Annual Volume = Daily * 365
        ws.cell(row=op['annual_vol'], column=col).value = f"={c}{op['daily_vol']}*365"

        # Fee with escalation
        ws.cell(row=op['fee'], column=col).value = f"=$D${inp['total_fee']}*(1+$D${inp['fee_esc']})^({year}-1)"

        # Revenue = Volume * Fee (NO ×1000 - fee is already per unit)
        ws.cell(row=op['revenue'], column=col).value = f"={c}{op['annual_vol']}*{c}{op['fee']}"

        # O&M with escalation
        ws.cell(row=op['om'], column=col).value = f"=$D${inp['om_y1']}*(1+$D${inp['om_esc']})^({year}-1)"

        # EBITDA = Revenue - O&M
        ws.cell(row=op['ebitda'], column=col).value = f"={c}{op['revenue']}-{c}{op['om']}"
        ws.cell(row=op['ebitda'], column=col).fill = GREEN_FILL
        ws.cell(row=op['ebitda'], column=col).font = BLACK_BOLD
        ws.cell(row=op['ebitda'], column=col).number_format = CURRENCY_FMT

        # Taxes
        ws.cell(row=op['taxes'], column=col).value = f"=MAX(0,{c}{op['ebitda']}-$D${calc['depreciation']})*$D${inp['tax_rate']}"

        # Growth CapEx (Y1-Y5 only)
        ws.cell(row=op['growth_capex'], column=col).value = f"=IF({year}<=5,$D${inp['growth_capex']},0)"

        # CFADS = EBITDA - Taxes - Growth CapEx
        ws.cell(row=op['cfads'], column=col).value = f"={c}{op['ebitda']}-{c}{op['taxes']}-{c}{op['growth_capex']}"
        ws.cell(row=op['cfads'], column=col).fill = GREEN_FILL
        ws.cell(row=op['cfads'], column=col).font = BLACK_BOLD

    # ===== 3. CASH SWEEP DEBT SCHEDULE =====
    print("[3] Fixing Cash Sweep Debt Schedule...")

    ws.cell(row=debt['beg_bal'], column=4).value = 0
    ws.cell(row=debt['end_bal'], column=4).value = f"=$D${calc['debt_amount']}"

    for col in range(5, 15):
        year = col - 4
        c = get_column_letter(col)
        prev = get_column_letter(col - 1)

        if year == 1:
            ws.cell(row=debt['beg_bal'], column=col).value = f"=$D${calc['debt_amount']}"
        else:
            ws.cell(row=debt['beg_bal'], column=col).value = f"={prev}{debt['end_bal']}"

        ws.cell(row=debt['interest'], column=col).value = f"={c}{debt['beg_bal']}*$D${inp['int_rate']}"
        ws.cell(row=debt['sched_prin'], column=col).value = f"=MIN($D${calc['sched_prin']},{c}{debt['beg_bal']})"
        ws.cell(row=debt['ds_pre'], column=col).value = f"={c}{debt['interest']}+{c}{debt['sched_prin']}"

        # Scheduled DS for DSRA (deterministic)
        ws.cell(row=debt['sched_ds'], column=col).value = \
            f"=MAX(0,$D${calc['debt_amount']}-({year}-1)*$D${calc['sched_prin']})*$D${inp['int_rate']}+$D${calc['sched_prin']}"

    # Excess, Sweep, etc. depend on DSRA funding
    for col in range(5, 15):
        year = col - 4
        c = get_column_letter(col)

        ws.cell(row=debt['excess'], column=col).value = f"=MAX(0,{c}{op['cfads']}-{c}{debt['ds_pre']}-{c}{dsra['funding']})"
        ws.cell(row=debt['sweep'], column=col).value = \
            f"=MIN({c}{debt['excess']}*$D${inp['sweep_pct']},MAX(0,{c}{debt['beg_bal']}-{c}{debt['sched_prin']}))"
        ws.cell(row=debt['total_prin'], column=col).value = f"={c}{debt['sched_prin']}+{c}{debt['sweep']}"
        ws.cell(row=debt['end_bal'], column=col).value = f"=MAX(0,{c}{debt['beg_bal']}-{c}{debt['total_prin']})"
        ws.cell(row=debt['dscr'], column=col).value = f"=IF({c}{debt['ds_pre']}>0,{c}{op['cfads']}/{c}{debt['ds_pre']},0)"
        ws.cell(row=debt['dscr'], column=col).number_format = MULTIPLE_FMT

    # ===== 4. DSRA =====
    print("[4] Fixing DSRA...")

    ws.cell(row=dsra['beg'], column=4).value = 0

    for col in range(5, 15):
        year = col - 4
        c = get_column_letter(col)
        prev = get_column_letter(col - 1)
        next_c = get_column_letter(col + 1) if col < 14 else c

        if col < 14:
            ws.cell(row=dsra['target'], column=col).value = f"={next_c}{debt['sched_ds']}*$D${inp['dsra_months']}/12"
        else:
            ws.cell(row=dsra['target'], column=col).value = 0

        if year == 1:
            ws.cell(row=dsra['beg'], column=col).value = f"=D{su['dsra_init']}"
        else:
            ws.cell(row=dsra['beg'], column=col).value = f"={prev}{dsra['end']}"

        ws.cell(row=dsra['funding'], column=col).value = f"={c}{dsra['target']}-{c}{dsra['beg']}"
        ws.cell(row=dsra['end'], column=col).value = f"={c}{dsra['beg']}+{c}{dsra['funding']}"

    # ===== 5. DISTRIBUTION WATERFALL =====
    print("[5] Fixing Distribution Waterfall...")

    for col in range(5, 15):
        c = get_column_letter(col)

        ws.cell(row=dist['avail'], column=col).value = \
            f"={c}{op['cfads']}-{c}{debt['interest']}-{c}{debt['total_prin']}-{c}{dsra['funding']}"
        ws.cell(row=dist['lockup'], column=col).value = f"=IF({c}{debt['dscr']}>=$D${inp['lockup']},\"PASS\",\"LOCKED\")"
        ws.cell(row=dist['distributable'], column=col).value = \
            f"=IF({c}{dist['lockup']}=\"PASS\",MAX(0,{c}{dist['avail']}),0)"

    # ===== 6. SOURCES & USES =====
    print("[6] Fixing Sources & Uses...")

    ws.cell(row=su['purchase'], column=4).value = f"=$D${inp['entry_ev']}"
    ws.cell(row=su['txn_fees'], column=4).value = f"=$D${inp['entry_ev']}*$D${inp['txn_fees']}"
    ws.cell(row=su['dsra_init'], column=4).value = f"=E{dsra['target']}"
    ws.cell(row=su['total_uses'], column=4).value = f"=D{su['purchase']}+D{su['txn_fees']}+D{su['dsra_init']}"

    ws.cell(row=su['debt'], column=4).value = f"=$D${inp['entry_ev']}*$D${inp['leverage']}"
    ws.cell(row=su['equity'], column=4).value = f"=D{su['total_uses']}-D{su['debt']}"
    ws.cell(row=su['total_sources'], column=4).value = f"=D{su['debt']}+D{su['equity']}"
    ws.cell(row=su['balance'], column=4).value = f"=IF(ABS(D{su['total_uses']}-D{su['total_sources']})<0.01,\"PASS\",\"FAIL\")"

    # ===== 7. EXIT & RETURNS =====
    print("[7] Fixing Exit & Returns...")

    exit_year = 7
    exit_col = get_column_letter(4 + exit_year)

    ws.cell(row=exit_r['exit_ebitda'], column=4).value = f"={exit_col}{op['ebitda']}"
    ws.cell(row=exit_r['exit_ev'], column=4).value = f"=D{exit_r['exit_ebitda']}*$D${inp['exit_mult']}"
    ws.cell(row=exit_r['exit_debt'], column=4).value = f"={exit_col}{debt['end_bal']}"
    ws.cell(row=exit_r['exit_dsra'], column=4).value = f"={exit_col}{dsra['end']}"
    ws.cell(row=exit_r['exit_equity'], column=4).value = f"=D{exit_r['exit_ev']}-D{exit_r['exit_debt']}+D{exit_r['exit_dsra']}"

    ws.cell(row=exit_r['equity_cf'], column=4).value = f"=-D{su['equity']}"
    for col in range(5, 15):
        year = col - 4
        c = get_column_letter(col)
        if year < exit_year:
            ws.cell(row=exit_r['equity_cf'], column=col).value = f"={c}{dist['distributable']}"
        elif year == exit_year:
            ws.cell(row=exit_r['equity_cf'], column=col).value = f"={c}{dist['distributable']}+D{exit_r['exit_equity']}"
        else:
            ws.cell(row=exit_r['equity_cf'], column=col).value = 0

    ws.cell(row=exit_r['irr'], column=4).value = f"=IRR(D{exit_r['equity_cf']}:N{exit_r['equity_cf']})"
    ws.cell(row=exit_r['irr'], column=4).fill = GREEN_FILL
    ws.cell(row=exit_r['irr'], column=4).font = BLACK_BOLD
    ws.cell(row=exit_r['irr'], column=4).number_format = PERCENT_FMT

    ws.cell(row=exit_r['moic'], column=4).value = \
        f"=(SUM(E{exit_r['equity_cf']}:K{exit_r['equity_cf']}))/-D{exit_r['equity_cf']}"
    ws.cell(row=exit_r['moic'], column=4).fill = GREEN_FILL
    ws.cell(row=exit_r['moic'], column=4).font = BLACK_BOLD
    ws.cell(row=exit_r['moic'], column=4).number_format = MULTIPLE_FMT

    # Unlevered
    ws.cell(row=exit_r['unlev_cf'], column=4).value = f"=-$D${inp['entry_ev']}-D{su['txn_fees']}"
    for col in range(5, 15):
        year = col - 4
        c = get_column_letter(col)
        if year < exit_year:
            ws.cell(row=exit_r['unlev_cf'], column=col).value = f"={c}{op['cfads']}"
        elif year == exit_year:
            ws.cell(row=exit_r['unlev_cf'], column=col).value = f"={c}{op['cfads']}+D{exit_r['exit_ev']}"
        else:
            ws.cell(row=exit_r['unlev_cf'], column=col).value = 0

    ws.cell(row=exit_r['unlev_irr'], column=4).value = f"=IRR(D{exit_r['unlev_cf']}:N{exit_r['unlev_cf']})"
    ws.cell(row=exit_r['unlev_irr'], column=4).fill = GREEN_FILL
    ws.cell(row=exit_r['unlev_irr'], column=4).font = BLACK_BOLD
    ws.cell(row=exit_r['unlev_irr'], column=4).number_format = PERCENT_FMT

    # ===== 8. AUDIT STRIP =====
    print("[8] Fixing Audit Strip...")

    ws.cell(row=3, column=4).value = f"=$D${inp['entry_ev']}"
    ws.cell(row=4, column=4).value = f"=E{op['ebitda']}"

    wb.save(filepath)
    print(f"\nSAVED: {filepath}")

    print("\n--- KEY FORMULAS ---")
    print(f"  D60 (Debt): {ws.cell(row=60, column=4).value}")
    print(f"  E65 (Daily Vol Y1): {ws.cell(row=65, column=5).value}")
    print(f"  E69 (Revenue): {ws.cell(row=69, column=5).value}")
    print(f"  E72 (EBITDA): {ws.cell(row=72, column=5).value}")

if __name__ == "__main__":
    fix_midstream()
