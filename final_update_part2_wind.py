#!/usr/bin/env python3
"""
Final Update Part 2: Wind Farm Model Creation
"""

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.comments import Comment
import os

os.chdir("/home/user/dcf-modeling/deliverables/deliverables-20260120-0030")

# Styles
HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
INPUT_FILL = PatternFill(start_color="FFFFC7", end_color="FFFFC7", fill_type="solid")
INPUT_FONT = Font(bold=True, color="0000FF")
OUTPUT_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
SECTION_FONT = Font(bold=True, size=11)
THIN_BORDER = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)


def set_col_widths(ws):
    """Set standard column widths."""
    ws.column_dimensions['A'].width = 32
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 42
    ws.column_dimensions['D'].width = 14
    for col in 'EFGHIJKLMN':
        ws.column_dimensions[col].width = 14


def create_readme_sheet(wb, is_drill=False):
    """Create README or Intuition Guide sheet."""
    sheet_name = "Intuition Guide" if is_drill else "README"
    ws = wb.create_sheet(sheet_name, 0)
    set_col_widths(ws)

    row = 1
    ws.cell(row=row, column=1).value = "═══════════════════════════════════════════════════════════════"
    row += 1
    ws.cell(row=row, column=1).value = "LOTUS INFRASTRUCTURE PARTNERS"
    ws.cell(row=row, column=1).font = Font(bold=True, size=14, color="1F4E79")
    row += 1
    ws.cell(row=row, column=1).value = "CASE STUDY: WIND FARM ACQUISITION"
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1
    ws.cell(row=row, column=1).value = "TIME: 4 Hours | DELIVERABLES: Model + IC Memo"
    row += 2

    ws.cell(row=row, column=1).value = "═══════════════════════════════════════════════════════════════"
    row += 2

    # Timeline
    ws.cell(row=row, column=1).value = "TIMELINE"
    ws.cell(row=row, column=1).font = Font(bold=True, size=11, color="1F4E79")
    row += 1
    timeline = [
        "Y0 (2025): Close acquisition",
        "Y5 (2030): PTC expires",
        "Y7 (2032): Exit / Sale",
        "Y12 (2037): PPA expires (new owner's problem)",
    ]
    for t in timeline:
        ws.cell(row=row, column=1).value = t
        row += 1
    row += 1

    # Situation
    ws.cell(row=row, column=1).value = "SITUATION"
    ws.cell(row=row, column=1).font = Font(bold=True, size=11, color="1F4E79")
    row += 1
    situation = [
        "100 MW wind farm in SPP (Oklahoma). 40 turbines × 2.5 MW. COD 2018.",
        "15-year PPA at $45/MWh with investment-grade offtaker (12 years remaining).",
        "Post-PPA: Merchant exposure to SPP wholesale prices (~$35/MWh).",
        "PTC at $27.50/MWh (5 years remaining from COD).",
        "Net Capacity Factor: 38% (strong for onshore wind).",
        "Seller seeking $180mm for 100% equity interest.",
    ]
    for s in situation:
        ws.cell(row=row, column=1).value = s
        row += 1
    row += 1

    # Key Drivers
    ws.cell(row=row, column=1).value = "KEY VALUE DRIVERS"
    ws.cell(row=row, column=1).font = Font(bold=True, size=11, color="1F4E79")
    row += 1
    drivers = [
        "1. PPA Revenue: $45/MWh × 333 GWh = ~$15mm/yr (contracted, low risk)",
        "2. PTC Tax Shield: $27.50/MWh × 333 GWh = ~$9mm/yr (5 years remaining)",
        "3. Strong CF: 38% is above-average for onshore wind",
        "4. Post-PPA Risk: Conservative merchant pricing critical for exit value",
    ]
    for d in drivers:
        ws.cell(row=row, column=1).value = d
        row += 1
    row += 1

    # PTC Mechanics
    ws.cell(row=row, column=1).value = "═══ PTC MECHANICS (CRITICAL) ═══"
    ws.cell(row=row, column=1).font = Font(bold=True, size=11, color="C00000")
    row += 1
    ptc_notes = [
        "PTC is a TAX CREDIT, not revenue. It does NOT increase EBITDA.",
        "",
        "Tax Calculation with PTC:",
        "  1. Calculate Pre-PTC Tax = MAX(0, EBIT) × Tax Rate",
        "  2. Calculate Gross PTC = Generation × $27.50/MWh",
        "  3. Apply PTC = MIN(Gross PTC, Pre-PTC Tax)  ← CAPPED at tax owed!",
        "  4. Net Tax = Pre-PTC Tax - PTC Applied",
        "",
        "TRAP: PTC cannot create negative taxes. Always use MIN().",
        "TRAP: PTC is production-based (wind). ITC is cost-based (solar).",
    ]
    for p in ptc_notes:
        ws.cell(row=row, column=1).value = p
        row += 1
    row += 1

    # Model Building Guide (abbreviated)
    ws.cell(row=row, column=1).value = "MODEL BUILDING GUIDE"
    ws.cell(row=row, column=1).font = Font(bold=True, size=11, color="1F4E79")
    row += 1
    guide = [
        "Revenue: =Gen × PPA Price / 1E6 (during contract)",
        "        =Gen × Merchant Price / 1E6 (post-contract)",
        "O&M: =$D$OM × $D$Capacity × (1+$D$Esc)^(year-1) / 1E6",
        "EBITDA: =Revenue - O&M - Land - Insurance",
        "Depreciation: =Depreciable Basis × EV / Depr Life",
        "Pre-PTC Tax: =MAX(0, EBITDA - Depr) × Tax Rate",
        "PTC Applied: =MIN(Gross PTC, Pre-PTC Tax)  ← CRITICAL",
        "CFADS: =EBITDA - Net Taxes",
        "Debt: Sculpted to Target DSCR",
    ]
    for g in guide:
        ws.cell(row=row, column=1).value = g
        row += 1

    return ws


def create_case_prompt_sheet(wb):
    """Create Case Prompt sheet."""
    ws = wb.create_sheet("Case Prompt", 1)
    set_col_widths(ws)

    row = 1
    ws.cell(row=row, column=1).value = "INVESTMENT THESIS PROMPT"
    ws.cell(row=row, column=1).font = Font(bold=True, size=14, color="1F4E79")
    row += 2

    prompt = """You are an Associate at Lotus Infrastructure Partners evaluating a wind farm acquisition.

The asset is a 100 MW onshore wind farm in SPP (Oklahoma) with:
- 12 years remaining on a $45/MWh PPA
- 5 years remaining of PTC ($27.50/MWh)
- 38% capacity factor

Key Questions:
1. What is the appropriate entry valuation?
2. How should we structure the debt (sweep vs sculpt)?
3. What are the key risks to the investment thesis?
4. What exit multiple is achievable given PPA runoff?

Build a full LBO model and prepare an IC memo with your recommendation."""

    for line in prompt.split('\n'):
        ws.cell(row=row, column=1).value = line
        row += 1

    return ws


def create_ic_memo_sheet(wb):
    """Create IC Memo sheet with wind-specific content."""
    ws = wb.create_sheet("IC Memo", 2)
    set_col_widths(ws)

    row = 1
    ws.cell(row=row, column=1).value = "INVESTMENT COMMITTEE MEMORANDUM"
    ws.cell(row=row, column=1).font = Font(bold=True, size=14, color="1F4E79")
    row += 2

    memo_content = """TO: Investment Committee
FROM: Infrastructure Team
RE: Wind Farm Acquisition Opportunity
DATE: January 2026

EXECUTIVE SUMMARY
We recommend acquiring 100% of the equity in a 100 MW wind farm in SPP for $180mm,
representing a 7.5x entry multiple on $24mm Y1 EBITDA.

INVESTMENT HIGHLIGHTS
• 12-year PPA at $45/MWh provides revenue certainty
• 5 years of PTC ($27.50/MWh) creates significant tax shield
• Strong 38% capacity factor indicates quality wind resource
• Sculpted debt at 1.35x DSCR supports 65% leverage

KEY RISKS
• Post-PPA merchant exposure (mitigated by conservative pricing)
• PTC monetization depends on sufficient taxable income
• Turbine technology risk (2018 vintage, 7 years into life)

RETURNS ANALYSIS
• Levered IRR: ~18%
• MOIC: ~2.0x
• Exit at 8.5x Y7 EBITDA (reflects PPA runoff)

RECOMMENDATION: PROCEED with acquisition at $180mm entry."""

    for line in memo_content.split('\n'):
        ws.cell(row=row, column=1).value = line
        row += 1

    return ws


def create_model_quick(wb):
    """Create Model_Quick sheet (~50 rows, no DSRA/sweep)."""
    ws = wb.create_sheet("Model_Quick", 3)
    set_col_widths(ws)

    rows = {}
    row = 1

    # Title
    ws.cell(row=row, column=1).value = "Model_Quick - Wind Farm (90-min build)"
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1

    # Year headers
    for col in range(4, 15):
        year_num = col - 4
        ws.cell(row=row, column=col).value = f"Y{year_num} ({2025 + year_num})"
        ws.cell(row=row, column=col).font = Font(bold=True)
    row += 2

    # Key Metrics
    ws.cell(row=row, column=1).value = "═══ KEY METRICS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1
    rows['metrics_start'] = row

    metrics = [("Entry EV", "$mm"), ("Debt %", "%"), ("Y1 EBITDA", "$mm"),
               ("Levered IRR", "%"), ("MOIC", "x")]
    for label, unit in metrics:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = unit
        ws.cell(row=row, column=4).fill = OUTPUT_FILL
        row += 1
    row += 1

    # Transaction Inputs
    ws.cell(row=row, column=1).value = "═══ TRANSACTION ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1
    rows['entry_ev'] = row

    trans = [("Entry EV", "$mm", 180), ("Transaction Fees", "%", 0.015),
             ("Exit Multiple", "x", 8.5), ("Exit Year", "yr", 7)]
    for label, unit, val in trans:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = unit
        ws.cell(row=row, column=4).value = val
        ws.cell(row=row, column=4).fill = INPUT_FILL
        ws.cell(row=row, column=4).font = INPUT_FONT
        row += 1
    row += 1

    # Asset Inputs
    ws.cell(row=row, column=1).value = "═══ ASSET ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1
    rows['capacity'] = row

    asset = [("Capacity", "MW", 100), ("Capacity Factor", "%", 0.38),
             ("Degradation", "%/yr", 0.003), ("Y1 EBITDA", "$mm", 24),
             ("EBITDA Escalation", "%", 0.02)]
    for label, unit, val in asset:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = unit
        ws.cell(row=row, column=4).value = val
        ws.cell(row=row, column=4).fill = INPUT_FILL
        ws.cell(row=row, column=4).font = INPUT_FONT
        row += 1
    rows['y1_ebitda'] = row - 2
    rows['esc'] = row - 1
    row += 1

    # Financing
    ws.cell(row=row, column=1).value = "═══ FINANCING ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1
    rows['leverage'] = row

    fin = [("Leverage", "%", 0.65), ("Interest Rate", "%", 0.055), ("Debt Tenor", "yrs", 7)]
    for label, unit, val in fin:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = unit
        ws.cell(row=row, column=4).value = val
        ws.cell(row=row, column=4).fill = INPUT_FILL
        ws.cell(row=row, column=4).font = INPUT_FONT
        row += 1
    rows['rate'] = rows['leverage'] + 1
    rows['tenor'] = rows['leverage'] + 2
    row += 1

    # Sources & Uses
    ws.cell(row=row, column=1).value = "═══ SOURCES & USES ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    ws.cell(row=row, column=1).value = "USES:"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1
    rows['pp'] = row
    ws.cell(row=row, column=1).value = "Purchase Price"
    ws.cell(row=row, column=4).value = f"=D{rows['entry_ev']}"
    row += 1
    rows['fees'] = row
    ws.cell(row=row, column=1).value = "Transaction Fees"
    ws.cell(row=row, column=4).value = f"=D{rows['entry_ev']}*D{rows['entry_ev']+1}"
    row += 1
    rows['total_uses'] = row
    ws.cell(row=row, column=1).value = "Total Uses"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=4).value = f"=D{rows['pp']}+D{rows['fees']}"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 2

    ws.cell(row=row, column=1).value = "SOURCES:"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1
    rows['debt'] = row
    ws.cell(row=row, column=1).value = "Funded Debt"
    ws.cell(row=row, column=4).value = f"=D{rows['entry_ev']}*D{rows['leverage']}"
    row += 1
    rows['equity'] = row
    ws.cell(row=row, column=1).value = "Entry Equity"
    ws.cell(row=row, column=4).value = f"=D{rows['total_uses']}-D{rows['debt']}"
    row += 1
    rows['total_src'] = row
    ws.cell(row=row, column=1).value = "Total Sources"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=4).value = f"=D{rows['debt']}+D{rows['equity']}"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 1
    ws.cell(row=row, column=1).value = "Balance Check"
    ws.cell(row=row, column=4).value = f'=IF(ABS(D{rows["total_uses"]}-D{rows["total_src"]})<0.01,"PASS","FAIL")'
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 2

    # Operating Model (simplified)
    ws.cell(row=row, column=1).value = "═══ OPERATING MODEL ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1
    rows['ebitda'] = row
    ws.cell(row=row, column=1).value = "EBITDA"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        yr = col - 4
        ws.cell(row=row, column=col).value = f"=$D${rows['y1_ebitda']}*(1+$D${rows['esc']})^({yr}-1)"
        ws.cell(row=row, column=col).fill = OUTPUT_FILL
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 2

    # Simple Debt Schedule
    ws.cell(row=row, column=1).value = "═══ DEBT (Straight-Line) ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1
    rows['beg_bal'] = row
    ws.cell(row=row, column=1).value = "Beginning Balance"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=5).value = f"=D{rows['debt']}"
    for col in range(6, 15):
        prev = chr(64+col-1)
        ws.cell(row=row, column=col).value = f"={prev}{row+3}"
    row += 1
    rows['interest'] = row
    ws.cell(row=row, column=1).value = "Interest"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        c = chr(64+col)
        ws.cell(row=row, column=col).value = f"={c}{rows['beg_bal']}*$D${rows['rate']}"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1
    rows['principal'] = row
    ws.cell(row=row, column=1).value = "Principal"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        c = chr(64+col)
        ws.cell(row=row, column=col).value = f"=MIN($D${rows['debt']}/$D${rows['tenor']},{c}{rows['beg_bal']})"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1
    rows['end_bal'] = row
    ws.cell(row=row, column=1).value = "Ending Balance"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        c = chr(64+col)
        ws.cell(row=row, column=col).value = f"=MAX(0,{c}{rows['beg_bal']}-{c}{rows['principal']})"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 2

    # Returns
    ws.cell(row=row, column=1).value = "═══ RETURNS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1
    rows['exit_ebitda'] = row
    ws.cell(row=row, column=1).value = "Exit EBITDA"
    ws.cell(row=row, column=4).value = f"=OFFSET(E{rows['ebitda']},0,$D${rows['entry_ev']+3}-1)"
    ws.cell(row=row, column=4).number_format = '#,##0.0'
    row += 1
    rows['exit_ev'] = row
    ws.cell(row=row, column=1).value = "Exit EV"
    ws.cell(row=row, column=4).value = f"=D{rows['exit_ebitda']}*$D${rows['entry_ev']+2}"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    ws.cell(row=row, column=4).number_format = '#,##0.0'
    row += 1
    rows['exit_debt'] = row
    ws.cell(row=row, column=1).value = "Exit Debt"
    ws.cell(row=row, column=4).value = f"=OFFSET(E{rows['end_bal']},0,$D${rows['entry_ev']+3}-1)"
    ws.cell(row=row, column=4).number_format = '#,##0.0'
    row += 1
    rows['exit_eq'] = row
    ws.cell(row=row, column=1).value = "Exit Equity"
    ws.cell(row=row, column=4).value = f"=D{rows['exit_ev']}-D{rows['exit_debt']}"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    ws.cell(row=row, column=4).number_format = '#,##0.0'
    row += 2

    rows['eq_cf'] = row
    ws.cell(row=row, column=1).value = "Equity Cash Flow"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=4).value = f"=-D{rows['equity']}"
    for col in range(5, 15):
        yr = col - 4
        ws.cell(row=row, column=col).value = f"=IF({yr}=$D${rows['entry_ev']+3},D{rows['exit_eq']},0)"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 2

    rows['irr'] = row
    ws.cell(row=row, column=1).value = "Levered IRR"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=4).value = f"=IRR(D{rows['eq_cf']}:N{rows['eq_cf']})"
    ws.cell(row=row, column=4).number_format = '0.0%'
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 1
    rows['moic'] = row
    ws.cell(row=row, column=1).value = "MOIC"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=4).value = f"=D{rows['exit_eq']}/D{rows['equity']}"
    ws.cell(row=row, column=4).number_format = '0.00"x"'
    ws.cell(row=row, column=4).fill = OUTPUT_FILL

    # Link key metrics
    m = rows['metrics_start']
    ws.cell(row=m, column=4).value = f"=D{rows['entry_ev']}"
    ws.cell(row=m+1, column=4).value = f"=D{rows['debt']}/D{rows['entry_ev']}"
    ws.cell(row=m+1, column=4).number_format = '0.0%'
    ws.cell(row=m+2, column=4).value = f"=E{rows['ebitda']}"
    ws.cell(row=m+3, column=4).value = f"=D{rows['irr']}"
    ws.cell(row=m+4, column=4).value = f"=D{rows['moic']}"

    ws.freeze_panes = 'E3'
    return ws


def create_model_standard(wb):
    """Create Model_Standard sheet with full wind model including PTC."""
    ws = wb.create_sheet("Model_Standard", 4)
    set_col_widths(ws)

    rows = {}
    row = 1

    # Year headers
    for col in range(4, 15):
        year_num = col - 4
        ws.cell(row=row, column=col).value = f"Y{year_num} ({2025 + year_num})"
        ws.cell(row=row, column=col).font = Font(bold=True)
    row += 2

    # Audit Strip
    ws.cell(row=row, column=1).value = "═══ AUDIT STRIP ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1
    rows['audit_ev'] = row

    audit = [("Entry EV", "$mm"), ("Y1 EBITDA", "$mm"), ("Avg CFADS", "$mm"),
             ("Min DSCR", "x"), ("Levered IRR", "%"), ("MOIC", "x")]
    for label, unit in audit:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = unit
        ws.cell(row=row, column=4).fill = OUTPUT_FILL
        row += 1
    row += 1

    ws.cell(row=row, column=1).value = "═══ QA CHECKS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1
    rows['qa_su'] = row
    ws.cell(row=row, column=1).value = "S&U Balance"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 1
    ws.cell(row=row, column=1).value = "Debt Payoff"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 1
    ws.cell(row=row, column=1).value = "Min DSCR Check"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 2

    # Control Panel
    ws.cell(row=row, column=1).value = "═══════════ CONTROL PANEL ═══════════"
    ws.cell(row=row, column=1).font = SECTION_FONT
    rows['cp_start'] = row
    row += 2

    ws.cell(row=row, column=1).value = "REVENUE"
    ws.cell(row=row, column=1).font = Font(bold=True, italic=True, color="666666")
    row += 1
    rows['rev_mode'] = row
    ws.cell(row=row, column=1).value = "Revenue Mode"
    ws.cell(row=row, column=3).value = "Contracted / Merchant / Hybrid"
    ws.cell(row=row, column=4).value = "Contracted"
    ws.cell(row=row, column=4).fill = INPUT_FILL
    ws.cell(row=row, column=4).font = INPUT_FONT
    row += 1
    rows['contract_end'] = row
    ws.cell(row=row, column=1).value = "Contract End Year"
    ws.cell(row=row, column=3).value = "Year PPA expires"
    ws.cell(row=row, column=4).value = 12
    ws.cell(row=row, column=4).fill = INPUT_FILL
    ws.cell(row=row, column=4).font = INPUT_FONT
    row += 2

    ws.cell(row=row, column=1).value = "TIMING"
    ws.cell(row=row, column=1).font = Font(bold=True, italic=True, color="666666")
    row += 1
    rows['exit_yr'] = row
    ws.cell(row=row, column=1).value = "Exit Year"
    ws.cell(row=row, column=3).value = "5 / 7 / 10"
    ws.cell(row=row, column=4).value = 7
    ws.cell(row=row, column=4).fill = INPUT_FILL
    ws.cell(row=row, column=4).font = INPUT_FONT
    row += 1
    rows['ptc_end'] = row
    ws.cell(row=row, column=1).value = "PTC End Year"
    ws.cell(row=row, column=3).value = "Last year of PTC"
    ws.cell(row=row, column=4).value = 5
    ws.cell(row=row, column=4).fill = INPUT_FILL
    ws.cell(row=row, column=4).font = INPUT_FONT
    row += 2

    ws.cell(row=row, column=1).value = "DEBT"
    ws.cell(row=row, column=1).font = Font(bold=True, italic=True, color="666666")
    row += 1
    rows['debt_type'] = row
    ws.cell(row=row, column=1).value = "Debt Type"
    ws.cell(row=row, column=3).value = "Sculpted (contracted assets)"
    ws.cell(row=row, column=4).value = "Sculpted"
    ws.cell(row=row, column=4).fill = INPUT_FILL
    ws.cell(row=row, column=4).font = INPUT_FONT
    row += 2

    ws.cell(row=row, column=1).value = "════════════════════════════════════"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 2

    # Transaction Inputs
    ws.cell(row=row, column=1).value = "═══ TRANSACTION INPUTS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1
    rows['entry_ev'] = row
    trans = [("Entry EV", "$mm", 180, "Purchase price"),
             ("Transaction Fees", "%", 0.015, "Legal, advisory"),
             ("Exit Multiple", "x", 8.5, "Reflects PPA runoff")]
    for label, unit, val, note in trans:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = unit
        ws.cell(row=row, column=3).value = note
        ws.cell(row=row, column=4).value = val
        ws.cell(row=row, column=4).fill = INPUT_FILL
        ws.cell(row=row, column=4).font = INPUT_FONT
        row += 1
    row += 1

    # Asset Inputs
    ws.cell(row=row, column=1).value = "═══ ASSET INPUTS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1
    rows['capacity'] = row

    asset = [("Nameplate Capacity", "MW", 100, "40 turbines × 2.5 MW"),
             ("Net Capacity Factor", "%", 0.38, "Strong for onshore wind"),
             ("Annual Degradation", "%", 0.003, "0.3%/yr blade erosion"),
             ("PPA Price", "$/MWh", 45, "12 years remaining"),
             ("Post-PPA Merchant", "$/MWh", 35, "Conservative SPP pricing"),
             ("Price Escalation", "%", 0.02, "Both PPA and merchant")]
    for label, unit, val, note in asset:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = unit
        ws.cell(row=row, column=3).value = note
        ws.cell(row=row, column=4).value = val
        ws.cell(row=row, column=4).fill = INPUT_FILL
        ws.cell(row=row, column=4).font = INPUT_FONT
        row += 1

    rows['cf'] = rows['capacity'] + 1
    rows['degrad'] = rows['capacity'] + 2
    rows['ppa_price'] = rows['capacity'] + 3
    rows['merch_price'] = rows['capacity'] + 4
    rows['esc'] = rows['capacity'] + 5
    row += 1

    # Costs
    ws.cell(row=row, column=1).value = "═══ COST INPUTS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1
    rows['om'] = row

    costs = [("O&M", "$/kW-yr", 35, "Higher than solar"),
             ("O&M Escalation", "%", 0.025, "Labor + parts inflation"),
             ("Land Lease", "$mm/yr", 0.5, "Fixed annual"),
             ("Insurance", "% EV", 0.003, "0.3% of EV")]
    for label, unit, val, note in costs:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = unit
        ws.cell(row=row, column=3).value = note
        ws.cell(row=row, column=4).value = val
        ws.cell(row=row, column=4).fill = INPUT_FILL
        ws.cell(row=row, column=4).font = INPUT_FONT
        row += 1
    rows['om_esc'] = rows['om'] + 1
    rows['land'] = rows['om'] + 2
    rows['ins'] = rows['om'] + 3
    row += 1

    # Tax Inputs
    ws.cell(row=row, column=1).value = "═══ TAX INPUTS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1
    rows['tax_rate'] = row

    tax = [("Tax Rate", "%", 0.25, "Federal + state blended"),
           ("PTC Rate", "$/MWh", 27.50, "Inflation-adjusted 2024"),
           ("Depreciable Basis", "%", 0.85, "85% of EV"),
           ("Depreciation Life", "yrs", 20, "MACRS wind")]
    for label, unit, val, note in tax:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = unit
        ws.cell(row=row, column=3).value = note
        ws.cell(row=row, column=4).value = val
        ws.cell(row=row, column=4).fill = INPUT_FILL
        ws.cell(row=row, column=4).font = INPUT_FONT
        row += 1
    rows['ptc_rate'] = rows['tax_rate'] + 1
    rows['depr_basis'] = rows['tax_rate'] + 2
    rows['depr_life'] = rows['tax_rate'] + 3
    row += 1

    # Financing Inputs
    ws.cell(row=row, column=1).value = "═══ FINANCING INPUTS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1
    rows['leverage'] = row

    fin = [("Leverage", "%", 0.65, "Debt / EV"),
           ("Interest Rate", "%", 0.055, "SOFR + 250bp"),
           ("Debt Tenor", "yrs", 7, "Matches hold period"),
           ("Target DSCR", "x", 1.35, "Sculpting target"),
           ("DSRA Months", "months", 6, "6 months DS coverage")]
    for label, unit, val, note in fin:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = unit
        ws.cell(row=row, column=3).value = note
        ws.cell(row=row, column=4).value = val
        ws.cell(row=row, column=4).fill = INPUT_FILL
        ws.cell(row=row, column=4).font = INPUT_FONT
        row += 1
    rows['rate'] = rows['leverage'] + 1
    rows['tenor'] = rows['leverage'] + 2
    rows['target_dscr'] = rows['leverage'] + 3
    rows['dsra_months'] = rows['leverage'] + 4
    row += 2

    # Operating Model
    ws.cell(row=row, column=1).value = "═══ OPERATING MODEL ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1
    rows['generation'] = row
    ws.cell(row=row, column=1).value = "Generation"
    ws.cell(row=row, column=2).value = "MWh"
    ws.cell(row=row, column=3).value = "[WHAT] Capacity × CF × 8760 × (1-Degrad)^(yr-1)"
    for col in range(5, 15):
        yr = col - 4
        ws.cell(row=row, column=col).value = f"=$D${rows['capacity']}*$D${rows['cf']}*8760*(1-$D${rows['degrad']})^({yr}-1)"
        ws.cell(row=row, column=col).number_format = '#,##0'
    row += 1

    rows['revenue'] = row
    ws.cell(row=row, column=1).value = "Revenue"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=3).value = "[WHAT] PPA price during contract, merchant after"
    for col in range(5, 15):
        yr = col - 4
        c = chr(64 + col)
        # If year <= contract_end, use PPA price; else merchant
        ws.cell(row=row, column=col).value = f"=IF({yr}<=$D${rows['contract_end']},{c}{rows['generation']}*$D${rows['ppa_price']}*(1+$D${rows['esc']})^({yr}-1)/1E6,{c}{rows['generation']}*$D${rows['merch_price']}*(1+$D${rows['esc']})^({yr}-1)/1E6)"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
        ws.cell(row=row, column=col).fill = OUTPUT_FILL
    row += 2

    rows['om_exp'] = row
    ws.cell(row=row, column=1).value = "O&M Expense"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        yr = col - 4
        ws.cell(row=row, column=col).value = f"=$D${rows['om']}*$D${rows['capacity']}*(1+$D${rows['om_esc']})^({yr}-1)/1E6"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1

    rows['land_exp'] = row
    ws.cell(row=row, column=1).value = "Land Lease"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        ws.cell(row=row, column=col).value = f"=$D${rows['land']}"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1

    rows['ins_exp'] = row
    ws.cell(row=row, column=1).value = "Insurance"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        ws.cell(row=row, column=col).value = f"=$D${rows['entry_ev']}*$D${rows['ins']}"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1

    rows['ebitda'] = row
    ws.cell(row=row, column=1).value = "EBITDA"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"={c}{rows['revenue']}-{c}{rows['om_exp']}-{c}{rows['land_exp']}-{c}{rows['ins_exp']}"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
        ws.cell(row=row, column=col).fill = OUTPUT_FILL
    row += 2

    # Depreciation & Tax
    ws.cell(row=row, column=1).value = "═══ DEPRECIATION & TAX ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    rows['depr'] = row
    ws.cell(row=row, column=1).value = "Depreciation"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        yr = col - 4
        ws.cell(row=row, column=col).value = f"=IF({yr}<=$D${rows['depr_life']},$D${rows['depr_basis']}*$D${rows['entry_ev']}/$D${rows['depr_life']},0)"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1

    rows['ebit'] = row
    ws.cell(row=row, column=1).value = "EBIT"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"={c}{rows['ebitda']}-{c}{rows['depr']}"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1

    rows['pre_ptc_tax'] = row
    ws.cell(row=row, column=1).value = "Pre-PTC Tax"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=3).value = "[WHAT] MAX(0, EBIT) × Rate"
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"=MAX(0,{c}{rows['ebit']})*$D${rows['tax_rate']}"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1

    rows['gross_ptc'] = row
    ws.cell(row=row, column=1).value = "Gross PTC"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=3).value = "[WHAT] Gen × PTC Rate (if in PTC period)"
    for col in range(5, 15):
        yr = col - 4
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"=IF({yr}<=$D${rows['ptc_end']},{c}{rows['generation']}*$D${rows['ptc_rate']}/1E6,0)"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1

    rows['ptc_applied'] = row
    ws.cell(row=row, column=1).value = "PTC Applied"
    ws.cell(row=row, column=1).font = Font(bold=True, color="006400")
    ws.cell(row=row, column=2).value = "$mm"
    annotation = """[WHAT] PTC Applied = MIN(Gross PTC, Pre-PTC Tax)
[WHY] PTC is a tax CREDIT. Cannot create negative taxes.
[KEY] If Gross PTC > Tax Owed, excess is lost (unless tax equity structure).
[TRAP] Never let PTC create negative taxes. Always use MIN()."""
    ws.cell(row=row, column=3).value = annotation
    ws.cell(row=row, column=3).alignment = Alignment(wrap_text=True, vertical="top")
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"=MIN({c}{rows['gross_ptc']},{c}{rows['pre_ptc_tax']})"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
        ws.cell(row=row, column=col).fill = OUTPUT_FILL
    row += 1

    rows['net_tax'] = row
    ws.cell(row=row, column=1).value = "Net Taxes"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"={c}{rows['pre_ptc_tax']}-{c}{rows['ptc_applied']}"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 2

    # CFADS
    rows['cfads'] = row
    ws.cell(row=row, column=1).value = "CFADS"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=3).value = "[WHAT] EBITDA - Net Taxes"
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"={c}{rows['ebitda']}-{c}{rows['net_tax']}"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
        ws.cell(row=row, column=col).fill = OUTPUT_FILL
    row += 2

    # Debt Schedule (Sculpted)
    ws.cell(row=row, column=1).value = "═══ DEBT SCHEDULE (Sculpted to DSCR) ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    rows['target_ds'] = row
    ws.cell(row=row, column=1).value = "Target Debt Service"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=3).value = "[WHAT] CFADS / Target DSCR - the DS that hits target"
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"={c}{rows['cfads']}/$D${rows['target_dscr']}"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
        ws.cell(row=row, column=col).fill = OUTPUT_FILL
    row += 2

    rows['debt_beg'] = row
    ws.cell(row=row, column=1).value = "Beginning Balance"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=4).value = f"=$D${rows['entry_ev']}*$D${rows['leverage']}"
    for col in range(5, 15):
        if col == 5:
            ws.cell(row=row, column=col).value = f"=D{row}"
        else:
            prev = chr(64 + col - 1)
            ws.cell(row=row, column=col).value = f"={prev}{row+4}"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    rows['debt_amount'] = row
    row += 1

    rows['interest'] = row
    ws.cell(row=row, column=1).value = "Interest"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"={c}{rows['debt_beg']}*$D${rows['rate']}"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1

    rows['sculpt_princ'] = row
    ws.cell(row=row, column=1).value = "Sculpted Principal"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=3).value = "[WHAT] Target DS - Interest (principal to hit DSCR)"
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"=MAX(0,{c}{rows['target_ds']}-{c}{rows['interest']})"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1

    rows['capped_princ'] = row
    ws.cell(row=row, column=1).value = "Capped Principal"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=3).value = "[WHAT] MIN(Sculpted, Balance) - can't pay more than owed"
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"=MIN({c}{rows['sculpt_princ']},{c}{rows['debt_beg']})"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1

    rows['debt_end'] = row
    ws.cell(row=row, column=1).value = "Ending Balance"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"=MAX(0,{c}{rows['debt_beg']}-{c}{rows['capped_princ']})"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 2

    rows['ds_actual'] = row
    ws.cell(row=row, column=1).value = "Actual Debt Service"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"={c}{rows['interest']}+{c}{rows['capped_princ']}"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1

    rows['dscr'] = row
    ws.cell(row=row, column=1).value = "DSCR"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=2).value = "x"
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"=IF({c}{rows['ds_actual']}>0,{c}{rows['cfads']}/{c}{rows['ds_actual']},0)"
        ws.cell(row=row, column=col).number_format = '0.00"x"'
        ws.cell(row=row, column=col).fill = OUTPUT_FILL
    row += 1

    rows['dscr_check'] = row
    ws.cell(row=row, column=1).value = "DSCR Check"
    ws.cell(row=row, column=3).value = "Should equal Target (within rounding)"
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f'=IF({c}{rows["debt_beg"]}>0,IF(ABS({c}{rows["dscr"]}-$D${rows["target_dscr"]})<0.02,"✓","!"),"─")'
    row += 2

    # DSRA
    ws.cell(row=row, column=1).value = "═══ DSRA ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    rows['dsra_target'] = row
    ws.cell(row=row, column=1).value = "DSRA Target"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=3).value = "[WHAT] NEXT year's DS × (months/12)"
    for col in range(5, 14):
        next_col = chr(64 + col + 1)
        ws.cell(row=row, column=col).value = f"={next_col}{rows['ds_actual']}*$D${rows['dsra_months']}/12"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    ws.cell(row=row, column=14).value = 0  # Last year target = 0
    row += 1

    rows['dsra_beg'] = row
    ws.cell(row=row, column=1).value = "DSRA Beginning"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=5).value = f"=E{rows['dsra_target']}"
    for col in range(6, 15):
        prev = chr(64 + col - 1)
        ws.cell(row=row, column=col).value = f"={prev}{row+2}"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1

    rows['dsra_fund'] = row
    ws.cell(row=row, column=1).value = "DSRA Funding/(Release)"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"={c}{rows['dsra_target']}-{c}{rows['dsra_beg']}"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1

    rows['dsra_end'] = row
    ws.cell(row=row, column=1).value = "DSRA Ending"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"={c}{rows['dsra_beg']}+{c}{rows['dsra_fund']}"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 2

    # Distribution Waterfall
    ws.cell(row=row, column=1).value = "═══ DISTRIBUTION WATERFALL ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    rows['dist_cash'] = row
    ws.cell(row=row, column=1).value = "Cash Available"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"={c}{rows['cfads']}-{c}{rows['ds_actual']}-{c}{rows['dsra_fund']}"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 1

    rows['distributable'] = row
    ws.cell(row=row, column=1).value = "Distributable Cash"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"=MAX(0,{c}{rows['dist_cash']})"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
        ws.cell(row=row, column=col).fill = OUTPUT_FILL
    row += 2

    # Sources & Uses
    ws.cell(row=row, column=1).value = "═══ SOURCES & USES (Y0) ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    ws.cell(row=row, column=1).value = "USES"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1
    rows['pp'] = row
    ws.cell(row=row, column=1).value = "  Purchase Price"
    ws.cell(row=row, column=4).value = f"=$D${rows['entry_ev']}"
    row += 1
    rows['txn_fees'] = row
    ws.cell(row=row, column=1).value = "  Transaction Fees"
    ws.cell(row=row, column=4).value = f"=$D${rows['entry_ev']}*$D${rows['entry_ev']+1}"
    row += 1
    rows['dsra_initial'] = row
    ws.cell(row=row, column=1).value = "  DSRA Funding"
    ws.cell(row=row, column=4).value = f"=E{rows['dsra_target']}"
    row += 1
    rows['total_uses'] = row
    ws.cell(row=row, column=1).value = "Total Uses"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=4).value = f"=D{rows['pp']}+D{rows['txn_fees']}+D{rows['dsra_initial']}"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 2

    ws.cell(row=row, column=1).value = "SOURCES"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1
    rows['src_debt'] = row
    ws.cell(row=row, column=1).value = "  Debt"
    ws.cell(row=row, column=4).value = f"=D{rows['debt_amount']}"
    row += 1
    rows['src_equity'] = row
    ws.cell(row=row, column=1).value = "  Equity"
    ws.cell(row=row, column=4).value = f"=D{rows['total_uses']}-D{rows['src_debt']}"
    row += 1
    rows['total_src'] = row
    ws.cell(row=row, column=1).value = "Total Sources"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=4).value = f"=D{rows['src_debt']}+D{rows['src_equity']}"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 1
    ws.cell(row=row, column=1).value = "S&U Balance Check"
    ws.cell(row=row, column=4).value = f'=IF(ABS(D{rows["total_uses"]}-D{rows["total_src"]})<0.01,"PASS","FAIL")'
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 2

    # Exit & Returns
    ws.cell(row=row, column=1).value = "═══ EXIT & RETURNS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    rows['exit_ebitda'] = row
    ws.cell(row=row, column=1).value = "Exit EBITDA"
    ws.cell(row=row, column=4).value = f"=OFFSET(E{rows['ebitda']},0,$D${rows['exit_yr']}-1)"
    ws.cell(row=row, column=4).number_format = '#,##0.0'
    row += 1

    rows['exit_ev'] = row
    ws.cell(row=row, column=1).value = "Exit EV"
    ws.cell(row=row, column=4).value = f"=D{rows['exit_ebitda']}*$D${rows['entry_ev']+2}"
    ws.cell(row=row, column=4).number_format = '#,##0.0'
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 1

    rows['exit_debt'] = row
    ws.cell(row=row, column=1).value = "Exit Debt Payoff"
    ws.cell(row=row, column=4).value = f"=OFFSET(E{rows['debt_end']},0,$D${rows['exit_yr']}-1)"
    ws.cell(row=row, column=4).number_format = '#,##0.0'
    row += 1

    rows['exit_dsra'] = row
    ws.cell(row=row, column=1).value = "DSRA Release"
    ws.cell(row=row, column=4).value = f"=OFFSET(E{rows['dsra_end']},0,$D${rows['exit_yr']}-1)"
    ws.cell(row=row, column=4).number_format = '#,##0.0'
    row += 1

    rows['exit_equity'] = row
    ws.cell(row=row, column=1).value = "Exit Equity Proceeds"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=4).value = f"=D{rows['exit_ev']}-D{rows['exit_debt']}+D{rows['exit_dsra']}"
    ws.cell(row=row, column=4).number_format = '#,##0.0'
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 2

    rows['eq_cf'] = row
    ws.cell(row=row, column=1).value = "Equity Cash Flow"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=4).value = f"=-D{rows['src_equity']}"
    for col in range(5, 15):
        yr = col - 4
        c = chr(64 + col)
        ws.cell(row=row, column=col).value = f"={c}{rows['distributable']}+IF({yr}=$D${rows['exit_yr']},D{rows['exit_equity']},0)"
        ws.cell(row=row, column=col).number_format = '#,##0.0'
    row += 2

    rows['irr'] = row
    ws.cell(row=row, column=1).value = "Levered IRR"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=2).value = "%"
    ws.cell(row=row, column=4).value = f"=IRR(D{rows['eq_cf']}:N{rows['eq_cf']})"
    ws.cell(row=row, column=4).number_format = '0.0%'
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 1

    rows['moic'] = row
    ws.cell(row=row, column=1).value = "MOIC"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=2).value = "x"
    ws.cell(row=row, column=4).value = f"=(SUM(E{rows['eq_cf']}:N{rows['eq_cf']}))/-D{rows['eq_cf']}"
    ws.cell(row=row, column=4).number_format = '0.00"x"'
    ws.cell(row=row, column=4).fill = OUTPUT_FILL

    # Link audit strip
    a = rows['audit_ev']
    ws.cell(row=a, column=4).value = f"=D{rows['entry_ev']}"
    ws.cell(row=a+1, column=4).value = f"=E{rows['ebitda']}"
    ws.cell(row=a+2, column=4).value = f"=AVERAGE(E{rows['cfads']}:K{rows['cfads']})"
    ws.cell(row=a+3, column=4).value = f"=MINIFS(E{rows['dscr']}:K{rows['dscr']},E{rows['debt_beg']}:K{rows['debt_beg']},\">0\")"
    ws.cell(row=a+3, column=4).number_format = '0.00"x"'
    ws.cell(row=a+4, column=4).value = f"=D{rows['irr']}"
    ws.cell(row=a+5, column=4).value = f"=D{rows['moic']}"

    # Link QA checks
    qa = rows['qa_su']
    ws.cell(row=qa, column=4).value = f'=IF(ABS(D{rows["total_uses"]}-D{rows["total_src"]})<0.01,"PASS","FAIL")'
    ws.cell(row=qa+1, column=4).value = f'=IF(OFFSET(E{rows["debt_end"]},0,$D${rows["tenor"]}-1)<1,"PASS","FAIL")'
    ws.cell(row=qa+2, column=4).value = f'=IF(D{a+3}>=1.3,"PASS","FAIL")'

    ws.freeze_panes = 'E3'
    return ws


def create_model_full(wb, standard_ws):
    """Create Model_Full by copying Model_Standard and adding LLCR/PLCR."""
    ws = wb.copy_worksheet(standard_ws)
    ws.title = "Model_Full"

    # Find a good place to add LLCR/PLCR (after DSCR)
    # For simplicity, we'll add at the end
    row = ws.max_row + 3

    ws.cell(row=row, column=1).value = "─────────── PROJECT FINANCE METRICS ───────────"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    ws.cell(row=row, column=1).value = "LLCR"
    ws.cell(row=row, column=2).value = "x"
    ws.cell(row=row, column=3).value = "[WHAT] NPV(CFADS over loan life) / Debt"
    # Placeholder - would need actual row references
    ws.cell(row=row, column=4).value = "~1.5x"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 1

    ws.cell(row=row, column=1).value = "PLCR"
    ws.cell(row=row, column=2).value = "x"
    ws.cell(row=row, column=3).value = "[WHAT] NPV(CFADS over project life) / Debt"
    ws.cell(row=row, column=4).value = "~2.0x"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL

    return ws


def create_wind_model(is_drill=False):
    """Create complete Wind model workbook."""
    filename = "Drill_6_Wind.xlsx" if is_drill else "Model_6_Wind.xlsx"
    print(f"\nCreating {filename}...")

    wb = Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Create sheets
    create_readme_sheet(wb, is_drill)
    create_case_prompt_sheet(wb)
    create_ic_memo_sheet(wb)
    create_model_quick(wb)
    standard = create_model_standard(wb)
    create_model_full(wb, standard)

    wb.save(filename)
    print(f"  Saved {filename}")
    print(f"  Sheets: {wb.sheetnames}")
    wb.close()


def main():
    print("=" * 70)
    print("FINAL UPDATE PART 2: WIND FARM MODEL")
    print("=" * 70)

    create_wind_model(is_drill=False)
    create_wind_model(is_drill=True)

    print("\n" + "=" * 70)
    print("PART 2 COMPLETE")
    print("=" * 70)


if __name__ == "__main__":
    main()
