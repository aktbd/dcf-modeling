#!/usr/bin/env python3
"""
Corrected fix for Model_4_Transmission.xlsx based on actual row structure
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

    # ===== ACTUAL ROW MAPPING (verified) =====
    inp = {
        'dev_cost_y0': 27,
        'dev_cost_y1': 28,
        'con_cost_y2': 29,
        'cod_year': 30,
        'capacity': 33,
        'availability': 34,
        'asset_life': 35,
        'tariff': 38,
        'tariff_esc': 39,
        'om_y3': 42,
        'om_esc': 43,
        'con_debt_pct': 46,
        'term_debt_pct': 47,
        'int_rate': 48,
        'dsra_months': 49,
        'target_dscr': 50,
        'lockup': 51,
        'exit_xrab': 54,
        'exit_year': 55,
        'tax_rate': 58,
    }

    calc = {
        'con_loan': 61,
        'con_interest': 62,
        'rab_cod': 63,
        'term_debt': 64,
        'depreciation': 65,
    }

    con = {
        'dev_eq_y0': 68,
        'dev_eq_y1': 69,
        'con_eq_y2': 70,
        'refin_eq_y3': 71,
        'total_eq': 72,
    }

    # ACTUAL operating model rows
    op = {
        'tariff': 75,
        'revenue': 77,
        'om': 78,
        'ebitda': 80,
        'depreciation': 83,
        'ebit': 84,
        'taxes': 85,
        'cfads': 86,
    }

    # ACTUAL debt schedule rows
    debt = {
        'beg_bal': 89,
        'interest': 90,
        'target_ds': 91,
        'sculpted_prin': 92,
        'actual_ds': 93,
        'end_bal': 94,
        'dscr': 95,
    }

    dsra = {
        'target': 98,
        'beg': 99,
        'funding': 100,
        'end': 101,
    }

    dist = {
        'avail': 104,
        'lockup': 105,
        'distributable': 106,
    }

    su = {
        'dev_y0': 110,
        'dev_y1': 111,
        'con_y2': 112,
        'con_int': 113,
        'dsra_init': 114,
        'total_uses': 115,
        'con_loan': 118,
        'term_debt': 119,
        'total_equity': 120,
        'total_sources': 121,
        'balance': 122,
    }

    exit_r = {
        'cum_dep': 125,
        'exit_rab': 126,
        'exit_ev': 127,
        'exit_debt': 128,
        'exit_dsra': 129,
        'exit_equity': 130,
        'equity_cf': 133,
        'irr': 134,
        'moic': 135,
        'unlev_cf': 138,
        'unlev_irr': 139,
    }

    print("="*60)
    print("FIXING MODEL_4_TRANSMISSION.xlsx (CORRECTED)")
    print("="*60)

    # ===== 1. CALCULATED ITEMS =====
    print("[1] Fixing Calculated Items...")

    ws.cell(row=calc['con_loan'], column=4).value = f"=$D${inp['con_cost_y2']}*$D${inp['con_debt_pct']}"
    ws.cell(row=calc['con_interest'], column=4).value = f"=$D${calc['con_loan']}*$D${inp['int_rate']}"
    ws.cell(row=calc['rab_cod'], column=4).value = f"=$D${inp['dev_cost_y0']}+$D${inp['dev_cost_y1']}+$D${inp['con_cost_y2']}+$D${calc['con_interest']}"
    ws.cell(row=calc['term_debt'], column=4).value = f"=$D${calc['rab_cod']}*$D${inp['term_debt_pct']}"
    ws.cell(row=calc['depreciation'], column=4).value = f"=$D${calc['rab_cod']}/$D${inp['asset_life']}"

    # ===== 2. CONSTRUCTION PHASE =====
    print("[2] Fixing Construction Phase...")

    ws.cell(row=con['dev_eq_y0'], column=4).value = f"=$D${inp['dev_cost_y0']}"
    ws.cell(row=con['dev_eq_y1'], column=5).value = f"=$D${inp['dev_cost_y1']}"
    ws.cell(row=con['con_eq_y2'], column=6).value = f"=$D${inp['con_cost_y2']}-$D${calc['con_loan']}"
    ws.cell(row=con['refin_eq_y3'], column=7).value = f"=$D${calc['con_loan']}-$D${calc['term_debt']}"
    ws.cell(row=con['total_eq'], column=4).value = f"=D{con['dev_eq_y0']}+E{con['dev_eq_y1']}+F{con['con_eq_y2']}+G{con['refin_eq_y3']}"

    # ===== 3. OPERATING MODEL (Y3+) =====
    print("[3] Fixing Operating Model...")

    cod_year = 3
    for col in range(5, 15):
        year = col - 4
        c = get_column_letter(col)

        if year >= cod_year:
            years_op = year - cod_year + 1

            ws.cell(row=op['tariff'], column=col).value = f"=$D${inp['tariff']}*(1+$D${inp['tariff_esc']})^({years_op}-1)"
            ws.cell(row=op['revenue'], column=col).value = f"=$D${inp['capacity']}*$D${inp['availability']}*{c}{op['tariff']}/1000"
            ws.cell(row=op['om'], column=col).value = f"=$D${inp['om_y3']}*(1+$D${inp['om_esc']})^({years_op}-1)"
            ws.cell(row=op['ebitda'], column=col).value = f"={c}{op['revenue']}-{c}{op['om']}"
            ws.cell(row=op['ebitda'], column=col).fill = GREEN_FILL
            ws.cell(row=op['ebitda'], column=col).font = BLACK_BOLD

            ws.cell(row=op['depreciation'], column=col).value = f"=$D${calc['depreciation']}"
            ws.cell(row=op['ebit'], column=col).value = f"={c}{op['ebitda']}-{c}{op['depreciation']}"
            ws.cell(row=op['taxes'], column=col).value = f"=MAX(0,{c}{op['ebit']})*$D${inp['tax_rate']}"
            ws.cell(row=op['cfads'], column=col).value = f"={c}{op['ebitda']}-{c}{op['taxes']}"
            ws.cell(row=op['cfads'], column=col).fill = GREEN_FILL
            ws.cell(row=op['cfads'], column=col).font = BLACK_BOLD
        else:
            for row in [op['tariff'], op['revenue'], op['om'], op['ebitda'], op['depreciation'], op['ebit'], op['taxes'], op['cfads']]:
                ws.cell(row=row, column=col).value = 0

    # ===== 4. SCULPTED DEBT (Y3+) =====
    print("[4] Fixing Sculpted Debt Schedule...")

    ws.cell(row=debt['beg_bal'], column=4).value = 0

    for col in range(5, 15):
        year = col - 4
        c = get_column_letter(col)
        prev = get_column_letter(col - 1)

        if year >= cod_year:
            if year == cod_year:
                ws.cell(row=debt['beg_bal'], column=col).value = f"=$D${calc['term_debt']}"
            else:
                ws.cell(row=debt['beg_bal'], column=col).value = f"={prev}{debt['end_bal']}"

            ws.cell(row=debt['interest'], column=col).value = f"={c}{debt['beg_bal']}*$D${inp['int_rate']}"
            ws.cell(row=debt['target_ds'], column=col).value = f"={c}{op['cfads']}/$D${inp['target_dscr']}"
            ws.cell(row=debt['sculpted_prin'], column=col).value = \
                f"=MIN(MAX(0,{c}{debt['target_ds']}-{c}{debt['interest']}),{c}{debt['beg_bal']})"
            ws.cell(row=debt['actual_ds'], column=col).value = f"={c}{debt['interest']}+{c}{debt['sculpted_prin']}"
            ws.cell(row=debt['end_bal'], column=col).value = f"=MAX(0,{c}{debt['beg_bal']}-{c}{debt['sculpted_prin']})"
            ws.cell(row=debt['dscr'], column=col).value = f"=IF({c}{debt['actual_ds']}>0,{c}{op['cfads']}/{c}{debt['actual_ds']},0)"
            ws.cell(row=debt['dscr'], column=col).number_format = MULTIPLE_FMT
        else:
            for row in [debt['beg_bal'], debt['interest'], debt['target_ds'], debt['sculpted_prin'], debt['actual_ds'], debt['end_bal'], debt['dscr']]:
                ws.cell(row=row, column=col).value = 0

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
            for row in [dsra['target'], dsra['beg'], dsra['funding'], dsra['end']]:
                ws.cell(row=row, column=col).value = 0

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

    ws.cell(row=su['dev_y0'], column=4).value = f"=$D${inp['dev_cost_y0']}"
    ws.cell(row=su['dev_y1'], column=4).value = f"=$D${inp['dev_cost_y1']}"
    ws.cell(row=su['con_y2'], column=4).value = f"=$D${inp['con_cost_y2']}"
    ws.cell(row=su['con_int'], column=4).value = f"=$D${calc['con_interest']}"
    ws.cell(row=su['dsra_init'], column=4).value = f"=G{dsra['target']}"
    ws.cell(row=su['total_uses'], column=4).value = f"=D{su['dev_y0']}+D{su['dev_y1']}+D{su['con_y2']}+D{su['con_int']}+D{su['dsra_init']}"

    ws.cell(row=su['con_loan'], column=4).value = f"=$D${calc['con_loan']}"
    ws.cell(row=su['term_debt'], column=4).value = f"=$D${calc['term_debt']}"
    ws.cell(row=su['total_equity'], column=4).value = f"=D{su['total_uses']}-D{su['con_loan']}-D{su['term_debt']}+D{su['con_loan']}"  # Net of refi
    ws.cell(row=su['total_sources'], column=4).value = f"=D{su['term_debt']}+D{su['total_equity']}"
    ws.cell(row=su['balance'], column=4).value = f"=IF(ABS(D{su['total_uses']}-D{su['total_sources']})<0.01,\"PASS\",\"FAIL\")"

    # ===== 8. EXIT & RETURNS =====
    print("[8] Fixing Exit & Returns...")

    exit_year = 10
    exit_col = get_column_letter(4 + exit_year)

    ws.cell(row=exit_r['cum_dep'], column=4).value = f"=$D${calc['depreciation']}*({exit_year}-{cod_year}+1)"
    ws.cell(row=exit_r['exit_rab'], column=4).value = f"=$D${calc['rab_cod']}-D{exit_r['cum_dep']}"
    ws.cell(row=exit_r['exit_ev'], column=4).value = f"=D{exit_r['exit_rab']}*$D${inp['exit_xrab']}"
    ws.cell(row=exit_r['exit_debt'], column=4).value = f"={exit_col}{debt['end_bal']}"
    ws.cell(row=exit_r['exit_dsra'], column=4).value = f"={exit_col}{dsra['end']}"
    ws.cell(row=exit_r['exit_equity'], column=4).value = f"=D{exit_r['exit_ev']}-D{exit_r['exit_debt']}+D{exit_r['exit_dsra']}"

    # Equity CF
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

    # Unlevered
    ws.cell(row=exit_r['unlev_cf'], column=4).value = f"=-$D${calc['rab_cod']}"
    for col in range(5, 15):
        year = col - 4
        c = get_column_letter(col)
        if year >= cod_year and year < exit_year:
            ws.cell(row=exit_r['unlev_cf'], column=col).value = f"={c}{op['cfads']}"
        elif year == exit_year:
            ws.cell(row=exit_r['unlev_cf'], column=col).value = f"={c}{op['cfads']}+D{exit_r['exit_ev']}"
        else:
            ws.cell(row=exit_r['unlev_cf'], column=col).value = 0

    ws.cell(row=exit_r['unlev_irr'], column=4).value = f"=IRR(D{exit_r['unlev_cf']}:N{exit_r['unlev_cf']})"
    ws.cell(row=exit_r['unlev_irr'], column=4).fill = GREEN_FILL
    ws.cell(row=exit_r['unlev_irr'], column=4).font = BLACK_BOLD
    ws.cell(row=exit_r['unlev_irr'], column=4).number_format = PERCENT_FMT

    # ===== 9. AUDIT STRIP =====
    print("[9] Fixing Audit Strip...")

    ws.cell(row=3, column=4).value = f"=$D${calc['rab_cod']}"
    ws.cell(row=4, column=4).value = f"=G{op['ebitda']}"

    wb.save(filepath)
    print(f"\nSAVED: {filepath}")

if __name__ == "__main__":
    fix_transmission()
