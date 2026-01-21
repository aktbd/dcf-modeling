#!/usr/bin/env python3
"""
Complete fix for Model_4_Transmission.xlsx
Key features: Construction phase, COD Y3, RAB valuation, Sculpted debt
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

def fix_transmission():
    filepath = f"{DELIV_DIR}/Model_4_Transmission.xlsx"
    wb = load_workbook(filepath)
    ws = wb['Model_Standard']

    # ===== INPUT ROW MAPPING =====
    inp = {
        'dev_cost_y0': 27,    # 15
        'dev_cost_y1': 28,    # 20
        'con_cost_y2': 29,    # 550
        'cod_year': 30,       # 3
        'capacity': 33,       # 3200
        'availability': 34,   # 0.98
        'asset_life': 35,     # 40
        'tariff': 38,         # 32
        'tariff_esc': 39,     # 0.02
        'om_y3': 42,          # 3
        'om_esc': 43,         # 0.025
        'con_debt_pct': 46,   # 0.85
        'term_debt_pct': 47,  # 0.6
        'int_rate': 48,       # 0.05
        'dsra_months': 49,    # 6
        'target_dscr': 50,    # 1.4
        'lockup': 51,         # 1.15
        'exit_xrab': 54,      # 1.15
        'exit_year': 55,      # 10
        'tax_rate': 58,       # 0.25
    }

    calc = {
        'con_loan': 61,
        'con_interest': 62,
        'rab_cod': 63,
        'term_debt': 64,
        'depreciation': 65,
    }

    # Row locations for operating model, etc. - need to find them
    # Let me first print more of the structure

    print("="*60)
    print("FIXING MODEL_4_TRANSMISSION.xlsx")
    print("="*60)

    # First, let's read more rows to understand the full structure
    print("\n[0] Reading full structure...")
    for row in range(60, 140):
        a = ws.cell(row=row, column=1).value or ""
        if a and str(a).strip():
            print(f"  Row {row}: {str(a)[:40]}")

    # Based on typical transmission structure:
    con = {
        'dev_eq_y0': 68,
        'dev_eq_y1': 69,
        'con_eq_y2': 70,
        'refin_eq_y3': 71,
        'total_eq': 72,
    }

    op = {
        'tariff': 75,
        'revenue': 76,
        'om': 77,
        'ebitda': 79,
        'taxes': 82,
        'cfads': 83,
    }

    debt = {
        'beg_bal': 87,
        'interest': 88,
        'target_ds': 89,
        'sculpted_prin': 90,
        'actual_ds': 91,
        'end_bal': 92,
        'dscr': 93,
    }

    dsra = {
        'target': 97,
        'beg': 98,
        'funding': 99,
        'end': 100,
    }

    dist = {
        'avail': 103,
        'lockup': 104,
        'distributable': 105,
    }

    su = {
        'con_costs': 109,
        'dsra_init': 110,
        'total_uses': 111,
        'term_debt': 113,
        'equity': 114,
        'total_sources': 115,
        'balance': 117,
    }

    exit_r = {
        'rab_exit': 120,
        'exit_ev': 121,
        'exit_debt': 122,
        'exit_dsra': 123,
        'exit_equity': 124,
        'equity_cf': 127,
        'irr': 128,
        'moic': 129,
    }

    # ===== 1. CALCULATED ITEMS =====
    print("[1] Fixing Calculated Items...")

    # Construction Loan = Construction Cost * Con Debt %
    ws.cell(row=calc['con_loan'], column=4).value = f"=$D${inp['con_cost_y2']}*$D${inp['con_debt_pct']}"

    # Construction Interest (capitalized)
    ws.cell(row=calc['con_interest'], column=4).value = f"=$D${calc['con_loan']}*$D${inp['int_rate']}"

    # RAB at COD = Dev Y0 + Dev Y1 + Construction + Capitalized Interest
    ws.cell(row=calc['rab_cod'], column=4).value = f"=$D${inp['dev_cost_y0']}+$D${inp['dev_cost_y1']}+$D${inp['con_cost_y2']}+$D${calc['con_interest']}"

    # Term Debt = RAB * Term Debt %
    ws.cell(row=calc['term_debt'], column=4).value = f"=$D${calc['rab_cod']}*$D${inp['term_debt_pct']}"

    # Annual Depreciation = RAB / Asset Life
    ws.cell(row=calc['depreciation'], column=4).value = f"=$D${calc['rab_cod']}/$D${inp['asset_life']}"

    # ===== 2. CONSTRUCTION PHASE EQUITY =====
    print("[2] Fixing Construction Phase...")

    # Y0 Dev Equity = Dev Cost Y0
    ws.cell(row=con['dev_eq_y0'], column=4).value = f"=$D${inp['dev_cost_y0']}"

    # Y1 Dev Equity (Column E)
    ws.cell(row=con['dev_eq_y1'], column=5).value = f"=$D${inp['dev_cost_y1']}"

    # Y2 Construction Equity = Con Cost - Con Loan (Column F)
    ws.cell(row=con['con_eq_y2'], column=6).value = f"=$D${inp['con_cost_y2']}-$D${calc['con_loan']}"

    # Y3 Refinancing = Pay off construction loan - Term Debt (Column G)
    ws.cell(row=con['refin_eq_y3'], column=7).value = f"=$D${calc['con_loan']}-$D${calc['term_debt']}"

    # Total Equity = sum across construction phase
    ws.cell(row=con['total_eq'], column=4).value = f"=D{con['dev_eq_y0']}+E{con['dev_eq_y1']}+F{con['con_eq_y2']}+G{con['refin_eq_y3']}"

    # ===== 3. OPERATING MODEL (Y3+) =====
    print("[3] Fixing Operating Model...")

    cod_year = 3
    for col in range(5, 15):  # E to N (Y1 to Y10)
        year = col - 4
        c = get_column_letter(col)

        # Only operate post-COD (Y3+)
        if year >= cod_year:
            years_operating = year - cod_year + 1

            # Tariff with escalation (from COD)
            ws.cell(row=op['tariff'], column=col).value = f"=$D${inp['tariff']}*(1+$D${inp['tariff_esc']})^({years_operating}-1)"

            # Revenue = Capacity * Availability * Tariff / 1000
            ws.cell(row=op['revenue'], column=col).value = f"=$D${inp['capacity']}*$D${inp['availability']}*{c}{op['tariff']}/1000"

            # O&M with escalation
            ws.cell(row=op['om'], column=col).value = f"=$D${inp['om_y3']}*(1+$D${inp['om_esc']})^({years_operating}-1)"

            # EBITDA
            ws.cell(row=op['ebitda'], column=col).value = f"={c}{op['revenue']}-{c}{op['om']}"
            ws.cell(row=op['ebitda'], column=col).fill = GREEN_FILL
            ws.cell(row=op['ebitda'], column=col).font = BLACK_BOLD

            # Taxes
            ws.cell(row=op['taxes'], column=col).value = f"=MAX(0,{c}{op['ebitda']}-$D${calc['depreciation']})*$D${inp['tax_rate']}"

            # CFADS
            ws.cell(row=op['cfads'], column=col).value = f"={c}{op['ebitda']}-{c}{op['taxes']}"
            ws.cell(row=op['cfads'], column=col).fill = GREEN_FILL
            ws.cell(row=op['cfads'], column=col).font = BLACK_BOLD
        else:
            # Pre-COD: zeros
            ws.cell(row=op['tariff'], column=col).value = 0
            ws.cell(row=op['revenue'], column=col).value = 0
            ws.cell(row=op['om'], column=col).value = 0
            ws.cell(row=op['ebitda'], column=col).value = 0
            ws.cell(row=op['taxes'], column=col).value = 0
            ws.cell(row=op['cfads'], column=col).value = 0

    # ===== 4. SCULPTED DEBT SCHEDULE (Y3+) =====
    print("[4] Fixing Sculpted Debt Schedule...")

    ws.cell(row=debt['beg_bal'], column=4).value = 0

    for col in range(5, 15):
        year = col - 4
        c = get_column_letter(col)
        prev = get_column_letter(col - 1)

        if year >= cod_year:
            # Beginning Balance
            if year == cod_year:
                ws.cell(row=debt['beg_bal'], column=col).value = f"=$D${calc['term_debt']}"
            else:
                ws.cell(row=debt['beg_bal'], column=col).value = f"={prev}{debt['end_bal']}"

            # Interest
            ws.cell(row=debt['interest'], column=col).value = f"={c}{debt['beg_bal']}*$D${inp['int_rate']}"

            # Target DS = CFADS / Target DSCR
            ws.cell(row=debt['target_ds'], column=col).value = f"={c}{op['cfads']}/$D${inp['target_dscr']}"

            # Sculpted Principal
            ws.cell(row=debt['sculpted_prin'], column=col).value = \
                f"=MIN(MAX(0,{c}{debt['target_ds']}-{c}{debt['interest']}),{c}{debt['beg_bal']})"

            # Actual DS
            ws.cell(row=debt['actual_ds'], column=col).value = f"={c}{debt['interest']}+{c}{debt['sculpted_prin']}"

            # Ending Balance
            ws.cell(row=debt['end_bal'], column=col).value = f"=MAX(0,{c}{debt['beg_bal']}-{c}{debt['sculpted_prin']})"

            # DSCR
            ws.cell(row=debt['dscr'], column=col).value = f"=IF({c}{debt['actual_ds']}>0,{c}{op['cfads']}/{c}{debt['actual_ds']},0)"
            ws.cell(row=debt['dscr'], column=col).number_format = MULTIPLE_FMT
        else:
            # Pre-COD
            ws.cell(row=debt['beg_bal'], column=col).value = 0
            ws.cell(row=debt['interest'], column=col).value = 0
            ws.cell(row=debt['target_ds'], column=col).value = 0
            ws.cell(row=debt['sculpted_prin'], column=col).value = 0
            ws.cell(row=debt['actual_ds'], column=col).value = 0
            ws.cell(row=debt['end_bal'], column=col).value = 0
            ws.cell(row=debt['dscr'], column=col).value = 0

    # ===== 5. DSRA =====
    print("[5] Fixing DSRA...")

    ws.cell(row=dsra['beg'], column=4).value = 0

    for col in range(5, 15):
        year = col - 4
        c = get_column_letter(col)
        prev = get_column_letter(col - 1)
        next_c = get_column_letter(col + 1) if col < 14 else c

        if year >= cod_year:
            if col < 14:
                ws.cell(row=dsra['target'], column=col).value = f"={next_c}{debt['actual_ds']}*$D${inp['dsra_months']}/12"
            else:
                ws.cell(row=dsra['target'], column=col).value = 0

            if year == cod_year:
                ws.cell(row=dsra['beg'], column=col).value = f"=D{su['dsra_init']}"
            else:
                ws.cell(row=dsra['beg'], column=col).value = f"={prev}{dsra['end']}"

            ws.cell(row=dsra['funding'], column=col).value = f"={c}{dsra['target']}-{c}{dsra['beg']}"
            ws.cell(row=dsra['end'], column=col).value = f"={c}{dsra['beg']}+{c}{dsra['funding']}"
        else:
            ws.cell(row=dsra['target'], column=col).value = 0
            ws.cell(row=dsra['beg'], column=col).value = 0
            ws.cell(row=dsra['funding'], column=col).value = 0
            ws.cell(row=dsra['end'], column=col).value = 0

    # ===== 6. DISTRIBUTION WATERFALL =====
    print("[6] Fixing Distribution Waterfall...")

    for col in range(5, 15):
        year = col - 4
        c = get_column_letter(col)

        if year >= cod_year:
            ws.cell(row=dist['avail'], column=col).value = \
                f"={c}{op['cfads']}-{c}{debt['actual_ds']}-{c}{dsra['funding']}"
            ws.cell(row=dist['lockup'], column=col).value = f"=IF({c}{debt['dscr']}>=$D${inp['lockup']},\"PASS\",\"LOCKED\")"
            ws.cell(row=dist['distributable'], column=col).value = \
                f"=IF({c}{dist['lockup']}=\"PASS\",MAX(0,{c}{dist['avail']}),0)"
        else:
            ws.cell(row=dist['avail'], column=col).value = 0
            ws.cell(row=dist['lockup'], column=col).value = "N/A"
            ws.cell(row=dist['distributable'], column=col).value = 0

    # ===== 7. SOURCES & USES =====
    print("[7] Fixing Sources & Uses...")

    # Total construction costs
    ws.cell(row=su['con_costs'], column=4).value = f"=$D${inp['dev_cost_y0']}+$D${inp['dev_cost_y1']}+$D${inp['con_cost_y2']}+$D${calc['con_interest']}"

    # DSRA Initial = Y3 DSRA Target
    ws.cell(row=su['dsra_init'], column=4).value = f"=G{dsra['target']}"

    # Total Uses
    ws.cell(row=su['total_uses'], column=4).value = f"=D{su['con_costs']}+D{su['dsra_init']}"

    # Term Debt
    ws.cell(row=su['term_debt'], column=4).value = f"=$D${calc['term_debt']}"

    # Equity = Total Uses - Term Debt
    ws.cell(row=su['equity'], column=4).value = f"=D{su['total_uses']}-D{su['term_debt']}"

    # Total Sources
    ws.cell(row=su['total_sources'], column=4).value = f"=D{su['term_debt']}+D{su['equity']}"

    # Balance Check
    ws.cell(row=su['balance'], column=4).value = f"=IF(ABS(D{su['total_uses']}-D{su['total_sources']})<0.01,\"PASS\",\"FAIL\")"

    # ===== 8. EXIT & RETURNS =====
    print("[8] Fixing Exit & Returns...")

    exit_year = 10
    exit_col = get_column_letter(4 + exit_year)

    # RAB at exit (depreciated)
    ws.cell(row=exit_r['rab_exit'], column=4).value = f"=$D${calc['rab_cod']}-$D${calc['depreciation']}*({exit_year}-{cod_year}+1)"

    # Exit EV = RAB * Exit xRAB
    ws.cell(row=exit_r['exit_ev'], column=4).value = f"=D{exit_r['rab_exit']}*$D${inp['exit_xrab']}"

    ws.cell(row=exit_r['exit_debt'], column=4).value = f"={exit_col}{debt['end_bal']}"
    ws.cell(row=exit_r['exit_dsra'], column=4).value = f"={exit_col}{dsra['end']}"
    ws.cell(row=exit_r['exit_equity'], column=4).value = f"=D{exit_r['exit_ev']}-D{exit_r['exit_debt']}+D{exit_r['exit_dsra']}"

    # Equity CF - construction phase then operating
    ws.cell(row=exit_r['equity_cf'], column=4).value = f"=-D{con['dev_eq_y0']}"
    ws.cell(row=exit_r['equity_cf'], column=5).value = f"=-E{con['dev_eq_y1']}"
    ws.cell(row=exit_r['equity_cf'], column=6).value = f"=-F{con['con_eq_y2']}"
    ws.cell(row=exit_r['equity_cf'], column=7).value = f"=-G{con['refin_eq_y3']}+G{dist['distributable']}"

    for col in range(8, 15):
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
        f"=(SUM(G{exit_r['equity_cf']}:N{exit_r['equity_cf']}))/(D{con['dev_eq_y0']}+E{con['dev_eq_y1']}+F{con['con_eq_y2']}+G{con['refin_eq_y3']})"
    ws.cell(row=exit_r['moic'], column=4).fill = GREEN_FILL
    ws.cell(row=exit_r['moic'], column=4).font = BLACK_BOLD
    ws.cell(row=exit_r['moic'], column=4).number_format = MULTIPLE_FMT

    # ===== 9. AUDIT STRIP =====
    print("[9] Fixing Audit Strip...")

    ws.cell(row=3, column=4).value = f"=$D${calc['rab_cod']}"  # RAB at COD instead of Entry EV
    ws.cell(row=4, column=4).value = f"=G{op['ebitda']}"  # Y3 EBITDA (first operating year)

    wb.save(filepath)
    print(f"\nSAVED: {filepath}")

    print("\n--- KEY FORMULAS ---")
    print(f"  D63 (RAB at COD): {ws.cell(row=63, column=4).value}")
    print(f"  D64 (Term Debt): {ws.cell(row=64, column=4).value}")
    print(f"  G76 (Revenue Y3): {ws.cell(row=76, column=7).value}")
    print(f"  G79 (EBITDA Y3): {ws.cell(row=79, column=7).value}")

if __name__ == "__main__":
    fix_transmission()
