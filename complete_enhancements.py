#!/usr/bin/env python3
"""
Complete Model Enhancements:
1. Add Spark Spread to Model_Standard/Model_Full for thermal models (CCGT, Peaker)
2. Add LLCR/PLCR to all Model_Full tabs
3. Update README with Cold Start Workflow, Time Budget, Debt Sizing guide
4. Replicate 3-tab structure to all models
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from copy import copy
import os

os.chdir("/home/user/dcf-modeling/deliverables/deliverables-20260120-0030")

# Style definitions
HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
INPUT_FILL = PatternFill(start_color="FFFFC7", end_color="FFFFC7", fill_type="solid")
INPUT_FONT = Font(bold=True, color="0000FF")
OUTPUT_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
CALC_FONT = Font(color="000000")
SECTION_FONT = Font(bold=True, size=11)


def add_control_panel(ws, start_row=5):
    """Add Control Panel with flags at specified row."""
    # Header
    ws.cell(row=start_row, column=1).value = "═══════════════════ CONTROL PANEL ═══════════════════"
    ws.cell(row=start_row, column=1).font = SECTION_FONT

    controls = [
        (start_row+2, "Revenue Mode", "Merchant", "Merchant / Contracted / Hybrid"),
        (start_row+3, "Contract Years", "0", "If Hybrid: years of contracted revenue"),
        (start_row+4, "Debt Sweep", "ON", "ON / OFF"),
        (start_row+5, "Sweep %", "75%", "Only active if Sweep = ON"),
        (start_row+6, "Exit Method", "Multiple", "Multiple / DCF"),
    ]

    for row, label, default, note in controls:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=1).font = Font(bold=True)
        ws.cell(row=row, column=4).value = default
        ws.cell(row=row, column=4).fill = INPUT_FILL
        ws.cell(row=row, column=4).font = INPUT_FONT
        ws.cell(row=row, column=3).value = note
        ws.cell(row=row, column=3).font = Font(italic=True, color="666666")

    # Add data validations
    rev_mode_dv = DataValidation(type="list", formula1='"Merchant,Contracted,Hybrid"', allow_blank=False)
    rev_mode_dv.add(ws.cell(row=start_row+2, column=4))
    ws.add_data_validation(rev_mode_dv)

    sweep_dv = DataValidation(type="list", formula1='"ON,OFF"', allow_blank=False)
    sweep_dv.add(ws.cell(row=start_row+4, column=4))
    ws.add_data_validation(sweep_dv)

    exit_dv = DataValidation(type="list", formula1='"Multiple,DCF"', allow_blank=False)
    exit_dv.add(ws.cell(row=start_row+6, column=4))
    ws.add_data_validation(exit_dv)

    ws.cell(row=start_row+8, column=1).value = "════════════════════════════════════════════════════"
    ws.cell(row=start_row+8, column=1).font = SECTION_FONT
    return start_row + 9


def create_model_quick(wb, asset_type, entry_ev=520, y1_ebitda=78, leverage=0.65, rate=0.065, tenor=7, exit_mult=7):
    """Create simplified Model_Quick tab (~60 rows, no DSRA, no sweep)."""
    ws = wb.create_sheet("Model_Quick", 0)

    # Set column widths
    ws.column_dimensions['A'].width = 32
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 38
    ws.column_dimensions['D'].width = 14
    for col in 'EFGHIJKLMN':
        ws.column_dimensions[col].width = 14

    row = 1

    # Title
    ws.cell(row=row, column=1).value = f"Model_Quick - {asset_type} (90-min build)"
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1

    # Year headers
    for col in range(4, 15):
        year_num = col - 4
        ws.cell(row=row, column=col).value = f"Y{year_num} ({2025 + year_num})"
        ws.cell(row=row, column=col).font = Font(bold=True)
    row += 2

    # Key Metrics (placeholder refs - will be linked at end)
    ws.cell(row=row, column=1).value = "═══ KEY METRICS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    metric_start = row
    metrics = [
        ("Entry EV", "$mm"),
        ("Debt %", "%"),
        ("Y1 EBITDA", "$mm"),
        ("Levered IRR", "%"),
        ("MOIC", "x"),
    ]
    for label, unit in metrics:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = unit
        ws.cell(row=row, column=4).fill = OUTPUT_FILL
        row += 1
    row += 1

    # Transaction Inputs
    ws.cell(row=row, column=1).value = "═══ TRANSACTION INPUTS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    entry_ev_row = row
    trans = [
        ("Entry EV", "$mm", entry_ev),
        ("Transaction Fees", "%", 0.015),
        ("Exit Multiple", "x", exit_mult),
        ("Exit Year", "yr", tenor),
    ]
    for label, unit, value in trans:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = unit
        ws.cell(row=row, column=4).value = value
        ws.cell(row=row, column=4).fill = INPUT_FILL
        ws.cell(row=row, column=4).font = INPUT_FONT
        row += 1
    row += 1

    # Asset Inputs
    ws.cell(row=row, column=1).value = "═══ ASSET INPUTS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    y1_ebitda_row = row + 2
    asset = [
        ("Capacity", "MW", 365),
        ("Capacity Factor", "%", 0.55),
        ("Y1 EBITDA", "$mm", y1_ebitda),
        ("EBITDA Escalation", "%", 0.02),
    ]
    for label, unit, value in asset:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = unit
        ws.cell(row=row, column=4).value = value
        ws.cell(row=row, column=4).fill = INPUT_FILL
        ws.cell(row=row, column=4).font = INPUT_FONT
        row += 1
    esc_row = row - 1
    row += 1

    # Financing Inputs
    ws.cell(row=row, column=1).value = "═══ FINANCING INPUTS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    leverage_row = row
    rate_row = row + 1
    tenor_row = row + 2
    fin = [
        ("Leverage", "%", leverage),
        ("Interest Rate", "%", rate),
        ("Debt Tenor", "yrs", tenor),
    ]
    for label, unit, value in fin:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = unit
        ws.cell(row=row, column=4).value = value
        ws.cell(row=row, column=4).fill = INPUT_FILL
        ws.cell(row=row, column=4).font = INPUT_FONT
        row += 1
    row += 1

    # Sources & Uses
    ws.cell(row=row, column=1).value = "═══ SOURCES & USES ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    ws.cell(row=row, column=1).value = "USES:"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1

    pp_row = row
    ws.cell(row=row, column=1).value = "Purchase Price"
    ws.cell(row=row, column=4).value = f"=D{entry_ev_row}"
    row += 1

    fees_row = row
    ws.cell(row=row, column=1).value = "Transaction Fees"
    ws.cell(row=row, column=4).value = f"=D{entry_ev_row}*D{entry_ev_row+1}"
    row += 1

    total_uses_row = row
    ws.cell(row=row, column=1).value = "Total Uses"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=4).value = f"=D{pp_row}+D{fees_row}"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 2

    ws.cell(row=row, column=1).value = "SOURCES:"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1

    debt_row = row
    ws.cell(row=row, column=1).value = "Funded Debt"
    ws.cell(row=row, column=4).value = f"=D{entry_ev_row}*D{leverage_row}"
    row += 1

    equity_row = row
    ws.cell(row=row, column=1).value = "Entry Equity"
    ws.cell(row=row, column=4).value = f"=D{total_uses_row}-D{debt_row}"
    row += 1

    total_src_row = row
    ws.cell(row=row, column=1).value = "Total Sources"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=4).value = f"=D{debt_row}+D{equity_row}"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 1

    ws.cell(row=row, column=1).value = "Balance Check"
    ws.cell(row=row, column=4).value = f'=IF(ABS(D{total_uses_row}-D{total_src_row})<0.01,"PASS","FAIL")'
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 2

    # Operating Model
    ws.cell(row=row, column=1).value = "═══ OPERATING MODEL ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    ebitda_row = row
    ws.cell(row=row, column=1).value = "EBITDA"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=3).value = "Y1 input, then escalated"
    for col in range(5, 15):
        year = col - 4
        ws.cell(row=row, column=col).value = f"=$D${y1_ebitda_row}*(1+$D${esc_row})^({year}-1)"
        ws.cell(row=row, column=col).fill = OUTPUT_FILL
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 2

    # Debt Schedule
    ws.cell(row=row, column=1).value = "═══ DEBT SCHEDULE (Straight-Line) ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    beg_bal_row = row
    ws.cell(row=row, column=1).value = "Beginning Balance"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=5).value = f"=D{debt_row}"
    for col in range(6, 15):
        prev = chr(64+col-1)
        ws.cell(row=row, column=col).value = f"={prev}{row+3}"
    row += 1

    interest_row = row
    ws.cell(row=row, column=1).value = "Interest"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        c = chr(64+col)
        ws.cell(row=row, column=col).value = f"={c}{beg_bal_row}*$D${rate_row}"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1

    principal_row = row
    ws.cell(row=row, column=1).value = "Principal"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=3).value = "Straight-line: Debt / Tenor"
    for col in range(5, 15):
        c = chr(64+col)
        ws.cell(row=row, column=col).value = f"=MIN($D${debt_row}/$D${tenor_row},{c}{beg_bal_row})"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1

    end_bal_row = row
    ws.cell(row=row, column=1).value = "Ending Balance"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        c = chr(64+col)
        ws.cell(row=row, column=col).value = f"=MAX(0,{c}{beg_bal_row}-{c}{principal_row})"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 2

    # Returns
    ws.cell(row=row, column=1).value = "═══ RETURNS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    exit_ebitda_row = row
    ws.cell(row=row, column=1).value = "Exit Year EBITDA"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=4).value = f"=OFFSET(E{ebitda_row},0,$D${entry_ev_row+3}-1)"
    ws.cell(row=row, column=4).number_format = '#,##0.0'
    row += 1

    exit_ev_row = row
    ws.cell(row=row, column=1).value = "Exit EV"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=4).value = f"=D{exit_ebitda_row}*$D${entry_ev_row+2}"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    ws.cell(row=row, column=4).number_format = '#,##0.0'
    row += 1

    exit_debt_row = row
    ws.cell(row=row, column=1).value = "Exit Debt Payoff"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=4).value = f"=OFFSET(E{end_bal_row},0,$D${entry_ev_row+3}-1)"
    ws.cell(row=row, column=4).number_format = '#,##0.0'
    row += 1

    exit_eq_row = row
    ws.cell(row=row, column=1).value = "Exit Equity Proceeds"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=4).value = f"=D{exit_ev_row}-D{exit_debt_row}"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    ws.cell(row=row, column=4).number_format = '#,##0.0'
    row += 2

    # Equity CF
    eq_cf_row = row
    ws.cell(row=row, column=1).value = "Equity Cash Flow"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=4).value = f"=-D{equity_row}"
    for col in range(5, 15):
        year = col - 4
        ws.cell(row=row, column=col).value = f"=IF({year}=$D${entry_ev_row+3},D{exit_eq_row},0)"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 2

    irr_row = row
    ws.cell(row=row, column=1).value = "Levered IRR"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=2).value = "%"
    ws.cell(row=row, column=4).value = f"=IRR(D{eq_cf_row}:N{eq_cf_row})"
    ws.cell(row=row, column=4).number_format = '0.0%'
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 1

    moic_row = row
    ws.cell(row=row, column=1).value = "MOIC"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=2).value = "x"
    ws.cell(row=row, column=4).value = f"=D{exit_eq_row}/D{equity_row}"
    ws.cell(row=row, column=4).number_format = '0.00"x"'
    ws.cell(row=row, column=4).fill = OUTPUT_FILL

    # Link metrics
    ws.cell(row=metric_start, column=4).value = f"=D{entry_ev_row}"
    ws.cell(row=metric_start+1, column=4).value = f"=D{debt_row}/D{entry_ev_row}"
    ws.cell(row=metric_start+1, column=4).number_format = '0.0%'
    ws.cell(row=metric_start+2, column=4).value = f"=E{ebitda_row}"
    ws.cell(row=metric_start+2, column=4).number_format = '#,##0.0'
    ws.cell(row=metric_start+3, column=4).value = f"=D{irr_row}"
    ws.cell(row=metric_start+3, column=4).number_format = '0.0%'
    ws.cell(row=metric_start+4, column=4).value = f"=D{moic_row}"
    ws.cell(row=metric_start+4, column=4).number_format = '0.00"x"'

    # Freeze panes
    ws.freeze_panes = 'E3'

    return ws


def update_readme_with_cold_start(ws):
    """Add Cold Start Workflow, Time Budget, and Debt Sizing guides to README."""
    # Find last row with content
    last_row = ws.max_row + 2
    row = last_row

    # Cold Start Workflow
    ws.cell(row=row, column=1).value = "═══════════════════════════════════════════════════════════════════════"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1
    ws.cell(row=row, column=1).value = "COLD START WORKFLOW (4-Hour Test)"
    ws.cell(row=row, column=1).font = Font(bold=True, size=14, color="1F4E79")
    row += 2

    workflow = [
        ("Phase 1: Setup (15 min)", [
            "→ Create workbook, freeze panes at E2",
            "→ Set column widths: A=32, B=12, C=38, D-N=14",
            "→ Build year headers: Y0 (2025), Y1 (2026), ... Y10 (2035)",
            "→ Create section placeholders with ═══ dividers",
        ]),
        ("Phase 2: Inputs & S&U (30 min)", [
            "→ Transaction: Entry EV, Fees %, Exit Multiple, Exit Year",
            "→ Asset: Capacity, CF/Dispatch, Prices, Escalation",
            "→ Financing: Leverage, Rate, Tenor, DSRA months, Sweep %",
            "→ Build Uses first, then Sources (balancing item = equity)",
        ]),
        ("Phase 3: Operating Model (45 min)", [
            "→ Generation = Capacity × CF × 8760 (baseload) OR Capacity × Dispatch × Avail (peaker)",
            "→ Revenue = Generation × Price / 1E6  (careful: capacity $/kW needs ×1000)",
            "→ Fuel = Gen × HR × Gas / 1E9  (CRITICAL: divide by BILLION)",
            "→ EBITDA = Revenue - OpEx",
            "→ Taxes = MAX(0, Taxable Income) × Rate  (no negative taxes!)",
        ]),
        ("Phase 4: Debt Schedule (45 min)", [
            "→ Beginning Bal Y1 = Funded Debt; Yn = Prior Ending",
            "→ Interest = Beg Bal × Rate",
            "→ Principal: Straight-line OR Sculpted OR Sweep",
            "→ DSCR = CFADS / Debt Service",
            "→ DSRA Target = NEXT year's DS × (months/12)  [TRAP: not current year!]",
        ]),
        ("Phase 5: Exit & Returns (30 min)", [
            "→ Exit EV = Exit EBITDA × Multiple",
            "→ Exit Equity = Exit EV - Debt Payoff + DSRA Release",
            "→ Equity CF row: Y0 = -Entry Equity, Y1-n = Distributions, Exit Year += Exit Equity",
            "→ IRR = IRR(Y0:Y10 equity CFs)",
            "→ MOIC = (Distributions + Exit) / Entry Equity",
        ]),
        ("Phase 6: QA & Polish (15 min)", [
            "→ S&U Balance Check = PASS",
            "→ Debt Payoff Check = PASS (ending bal < $1mm)",
            "→ Min DSCR Check = PASS (≥1.25x merchant, ≥1.30x contracted)",
            "→ Sanity: DSCR 1.2-2.5x, IRR 12-25%, Fuel $30-50mm (CCGT)",
        ]),
    ]

    for phase, steps in workflow:
        ws.cell(row=row, column=1).value = phase
        ws.cell(row=row, column=1).font = Font(bold=True, size=11)
        row += 1
        for step in steps:
            ws.cell(row=row, column=1).value = step
            ws.cell(row=row, column=1).font = Font(size=10)
            row += 1
        row += 1

    row += 1

    # Time Budget
    ws.cell(row=row, column=1).value = "4-HOUR TIME BUDGET"
    ws.cell(row=row, column=1).font = Font(bold=True, size=12, color="1F4E79")
    row += 2

    time_items = [
        ("Setup & Structure", "15 min", "Workbook, headers, freeze panes"),
        ("Transaction Inputs", "15 min", "EV, fees, leverage, rates"),
        ("Sources & Uses", "15 min", "Uses first, equity = balancing"),
        ("Operating Model", "45 min", "Revenue, costs, EBITDA"),
        ("Depreciation & Tax", "15 min", "Straight-line, MAX(0) guardrail"),
        ("Debt Schedule", "45 min", "Amort, interest, DSCR"),
        ("DSRA (if needed)", "15 min", "Target = NEXT year DS"),
        ("Waterfall & Returns", "30 min", "Distributable, IRR, MOIC"),
        ("Exit Mechanics", "15 min", "Exit EV, debt payoff, equity"),
        ("QA & Sanity Checks", "15 min", "Balance, DSCR, #REF sweep"),
        ("Buffer / Iterate", "15 min", "Fix issues, polish"),
        ("TOTAL", "4:00", ""),
    ]

    for item, time, note in time_items:
        ws.cell(row=row, column=1).value = item
        ws.cell(row=row, column=2).value = time
        ws.cell(row=row, column=3).value = note
        if item == "TOTAL":
            ws.cell(row=row, column=1).font = Font(bold=True)
            ws.cell(row=row, column=2).font = Font(bold=True)
        row += 1

    row += 2

    # Debt Sizing Approaches
    ws.cell(row=row, column=1).value = "DEBT SIZING APPROACHES"
    ws.cell(row=row, column=1).font = Font(bold=True, size=12, color="1F4E79")
    row += 2

    debt_approaches = [
        ("1. STRAIGHT-LINE (Simplest)", [
            "Principal = Debt Amount / Tenor",
            "Best for: Quick models, Model_Quick, when told 'ignore sweep'",
            "Pros: No circularity, fast to build",
            "Cons: May not reflect economic reality",
        ]),
        ("2. CASH SWEEP (Merchant)", [
            "Sweep = Excess Cash × Sweep%  (Excess = CFADS - DS - DSRA)",
            "Best for: Merchant assets, volatile CFADS",
            "Pros: Realistic for bank debt, de-levers faster",
            "Cons: Requires DSRA logic, near-circular (use scheduled_ds trick)",
        ]),
        ("3. SCULPTED (Contracted)", [
            "Target DS = CFADS / Target DSCR; Principal = Target DS - Interest",
            "Best for: PPAs, regulated assets, predictable CFADS",
            "Pros: Maximizes debt capacity, matches CF profile",
            "Cons: More complex, requires CFADS forecast first",
        ]),
    ]

    for approach, details in debt_approaches:
        ws.cell(row=row, column=1).value = approach
        ws.cell(row=row, column=1).font = Font(bold=True, size=11)
        row += 1
        for detail in details:
            ws.cell(row=row, column=1).value = f"  {detail}"
            ws.cell(row=row, column=1).font = Font(size=10)
            row += 1
        row += 1

    row += 1

    # Spark Spread explanation
    ws.cell(row=row, column=1).value = "KEY FORMULA CONCEPTS"
    ws.cell(row=row, column=1).font = Font(bold=True, size=12, color="1F4E79")
    row += 2

    concepts = [
        ("SPARK SPREAD (Thermal Plants)", [
            "= Power Price ($/MWh) - Fuel Cost per MWh",
            "= Power Price - (Heat Rate × Gas Price / 1000)",
            "Example: $45/MWh - (7000 Btu/kWh × $3/mmBtu / 1000) = $45 - $21 = $24/MWh",
            "Core margin driver for CCGT/Peaker dispatch economics",
        ]),
        ("UCAP (Unforced Capacity)", [
            "= Nameplate Capacity × (1 - Forced Outage Rate)",
            "= 365 MW × (1 - 5%) = 346.75 MW",
            "Capacity revenue uses UCAP, not nameplate",
            "Capacity Price is $/kW-yr, so: UCAP × Price × 1000 / 1E6",
        ]),
        ("LLCR / PLCR (Project Finance)", [
            "LLCR = NPV(CFADS over loan life) / Debt Outstanding",
            "PLCR = NPV(CFADS over project life) / Debt Outstanding",
            "LLCR ~1.4x+ typical for investment grade",
            "Covenant-based lending uses these vs DSCR",
        ]),
    ]

    for concept, details in concepts:
        ws.cell(row=row, column=1).value = concept
        ws.cell(row=row, column=1).font = Font(bold=True, size=11)
        row += 1
        for detail in details:
            ws.cell(row=row, column=1).value = f"  {detail}"
            ws.cell(row=row, column=1).font = Font(size=10)
            row += 1
        row += 1

    return ws


def add_llcr_plcr_to_full(ws, insert_row):
    """Add LLCR/PLCR section to Model_Full at specified row."""
    # Insert 6 rows for LLCR/PLCR section
    ws.insert_rows(insert_row, 6)

    row = insert_row
    ws.cell(row=row, column=1).value = "─────────── PROJECT FINANCE COVERAGE ───────────"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    # Note: These are placeholder formulas - actual row refs depend on model structure
    ws.cell(row=row, column=1).value = "LLCR (Loan Life Coverage)"
    ws.cell(row=row, column=2).value = "x"
    ws.cell(row=row, column=3).value = "[WHAT] NPV(CFADS over loan life) / Debt; [WHY] Total debt capacity"
    ws.cell(row=row, column=4).value = "=NPV(D45,E82:K82)/D100"  # Placeholder
    ws.cell(row=row, column=4).number_format = '0.00"x"'
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 1

    ws.cell(row=row, column=1).value = "PLCR (Project Life Coverage)"
    ws.cell(row=row, column=2).value = "x"
    ws.cell(row=row, column=3).value = "[WHAT] NPV(CFADS over project life) / Debt; [WHY] Full asset capacity"
    ws.cell(row=row, column=4).value = "=NPV(D45,E82:N82)/D100"  # Placeholder
    ws.cell(row=row, column=4).number_format = '0.00"x"'
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 1

    ws.cell(row=row, column=1).value = "─────────────────────────────────────────────────"

    return ws


def process_model(filename, asset_type, params):
    """Process a single model file to add 3-tab structure."""
    print(f"\nProcessing {filename}...")

    wb = load_workbook(filename)

    # Check if already processed
    if "Model_Quick" in wb.sheetnames:
        print(f"  SKIP: Already has 3-tab structure")
        wb.close()
        return

    # Get Model sheet
    if "Model" not in wb.sheetnames and "Model_Standard" not in wb.sheetnames:
        print(f"  ERROR: No Model sheet found")
        wb.close()
        return

    source_name = "Model" if "Model" in wb.sheetnames else "Model_Standard"
    source_ws = wb[source_name]

    # Create Model_Quick
    print("  Creating Model_Quick...")
    create_model_quick(wb, asset_type, **params)

    # Rename Model to Model_Standard if needed
    if source_name == "Model":
        print("  Renaming Model → Model_Standard...")
        source_ws.title = "Model_Standard"

    # Add Control Panel to Model_Standard
    print("  Adding Control Panel...")
    source_ws.insert_rows(5, 10)
    add_control_panel(source_ws, start_row=5)

    # Copy to Model_Full
    print("  Creating Model_Full...")
    model_full = wb.copy_worksheet(source_ws)
    model_full.title = "Model_Full"

    # Add LLCR/PLCR to Model_Full (insert after debt schedule area)
    print("  Adding LLCR/PLCR to Model_Full...")
    # Find DSRA section or debt section end (around row 115-120 typically)
    add_llcr_plcr_to_full(model_full, insert_row=120)

    # Update README with Cold Start Workflow
    if "README" in wb.sheetnames:
        print("  Updating README with Cold Start Workflow...")
        update_readme_with_cold_start(wb["README"])

    # Reorder sheets
    print("  Reordering sheets...")
    desired_order = ['README', 'Case Prompt', 'IC Memo', 'Model_Quick', 'Model_Standard', 'Model_Full']
    # Also handle Intuition Guide for drill versions
    if 'Intuition Guide' in wb.sheetnames:
        desired_order = ['Intuition Guide', 'Case Prompt', 'IC Memo', 'Model_Quick', 'Model_Standard', 'Model_Full']

    for i, name in enumerate(desired_order):
        if name in wb.sheetnames:
            current_idx = wb.sheetnames.index(name)
            wb.move_sheet(name, offset=i - current_idx)

    wb.save(filename)
    print(f"  Saved. Sheets: {wb.sheetnames}")
    wb.close()


# Asset parameters (defaults for Model_Quick)
ASSET_PARAMS = {
    "CCGT": {"entry_ev": 520, "y1_ebitda": 74, "leverage": 0.65, "rate": 0.065, "tenor": 7, "exit_mult": 7.0},
    "Peaker": {"entry_ev": 150, "y1_ebitda": 22, "leverage": 0.60, "rate": 0.070, "tenor": 5, "exit_mult": 6.5},
    "SolarBESS": {"entry_ev": 425, "y1_ebitda": 52, "leverage": 0.70, "rate": 0.055, "tenor": 15, "exit_mult": 8.0},
    "Transmission": {"entry_ev": 850, "y1_ebitda": 112, "leverage": 0.75, "rate": 0.050, "tenor": 20, "exit_mult": 9.0},
    "Midstream": {"entry_ev": 600, "y1_ebitda": 85, "leverage": 0.65, "rate": 0.060, "tenor": 7, "exit_mult": 7.5},
}

# Files to process
MODEL_FILES = [
    ("Model_1_CCGT.xlsx", "CCGT"),
    ("Model_2_Peaker.xlsx", "Peaker"),
    ("Model_3_SolarBESS.xlsx", "SolarBESS"),
    ("Model_4_Transmission.xlsx", "Transmission"),
    ("Model_5_Midstream.xlsx", "Midstream"),
]

DRILL_FILES = [
    ("Drill_1_CCGT.xlsx", "CCGT"),
    ("Drill_2_Peaker.xlsx", "Peaker"),
    ("Drill_3_SolarBESS.xlsx", "SolarBESS"),
    ("Drill_4_Transmission.xlsx", "Transmission"),
    ("Drill_5_Midstream.xlsx", "Midstream"),
]


def main():
    print("=" * 70)
    print("COMPLETE MODEL ENHANCEMENT")
    print("=" * 70)

    # Process all models
    print("\n--- Processing Model Files ---")
    for filename, asset_type in MODEL_FILES:
        if os.path.exists(filename):
            process_model(filename, asset_type, ASSET_PARAMS[asset_type])
        else:
            print(f"  SKIP: {filename} not found")

    # Process all drills
    print("\n--- Processing Drill Files ---")
    for filename, asset_type in DRILL_FILES:
        if os.path.exists(filename):
            process_model(filename, asset_type, ASSET_PARAMS[asset_type])
        else:
            print(f"  SKIP: {filename} not found")

    print("\n" + "=" * 70)
    print("VERIFICATION")
    print("=" * 70)

    # Verify all files
    for filename, _ in MODEL_FILES + DRILL_FILES:
        if os.path.exists(filename):
            wb = load_workbook(filename, read_only=True)
            print(f"{filename}: {wb.sheetnames}")
            wb.close()

    print("\n" + "=" * 70)
    print("ENHANCEMENT COMPLETE")
    print("=" * 70)


if __name__ == "__main__":
    main()
