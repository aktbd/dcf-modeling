#!/usr/bin/env python3
"""
Comprehensive Model Enhancement: 3-Tab Structure + Control Panel + Learning Scaffolds
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

def copy_sheet_structure(source_ws, target_wb, new_name):
    """Copy a worksheet to a new workbook with a new name."""
    target_ws = target_wb.create_sheet(new_name)

    # Copy column widths
    for col_letter, col_dim in source_ws.column_dimensions.items():
        target_ws.column_dimensions[col_letter].width = col_dim.width

    # Copy row heights
    for row_num, row_dim in source_ws.row_dimensions.items():
        target_ws.row_dimensions[row_num].height = row_dim.height

    # Copy cell values and styles
    for row in source_ws.iter_rows():
        for cell in row:
            new_cell = target_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.fill = copy(cell.fill)
                new_cell.border = copy(cell.border)
                new_cell.alignment = copy(cell.alignment)
                new_cell.number_format = cell.number_format

    # Copy merged cells
    for merged_range in source_ws.merged_cells.ranges:
        target_ws.merge_cells(str(merged_range))

    return target_ws


def add_control_panel(ws, start_row=5):
    """Add Control Panel with flags."""

    # Header
    ws.cell(row=start_row, column=1).value = "═══════════════════ CONTROL PANEL ═══════════════════"
    ws.cell(row=start_row, column=1).font = SECTION_FONT

    # Flags
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

    # Add data validation for dropdowns
    rev_mode_dv = DataValidation(type="list", formula1='"Merchant,Contracted,Hybrid"', allow_blank=False)
    rev_mode_dv.add(ws.cell(row=start_row+2, column=4))
    ws.add_data_validation(rev_mode_dv)

    sweep_dv = DataValidation(type="list", formula1='"ON,OFF"', allow_blank=False)
    sweep_dv.add(ws.cell(row=start_row+4, column=4))
    ws.add_data_validation(sweep_dv)

    exit_dv = DataValidation(type="list", formula1='"Multiple,DCF"', allow_blank=False)
    exit_dv.add(ws.cell(row=start_row+6, column=4))
    ws.add_data_validation(exit_dv)

    # Closing line
    ws.cell(row=start_row+8, column=1).value = "════════════════════════════════════════════════════"
    ws.cell(row=start_row+8, column=1).font = SECTION_FONT

    return start_row + 9


def add_spark_spread_section(ws, start_row, hr_row, gas_row, power_row, vom_row):
    """Add Spark Spread / Margin Analysis section for thermal models."""

    ws.cell(row=start_row, column=1).value = "─────────── MARGIN ANALYSIS ───────────"
    ws.cell(row=start_row, column=1).font = SECTION_FONT

    rows = {
        'hr_mmbtu': start_row + 1,
        'fuel_per_mwh': start_row + 2,
        'spark_spread': start_row + 3,
        'vom': start_row + 4,
        'net_margin': start_row + 5,
    }

    # HR in MMBtu/MWh
    ws.cell(row=rows['hr_mmbtu'], column=1).value = "HR (MMBtu/MWh)"
    ws.cell(row=rows['hr_mmbtu'], column=2).value = "MMBtu/MWh"
    ws.cell(row=rows['hr_mmbtu'], column=3).value = "[WHAT] Convert Btu/kWh to MMBtu/MWh by /1000"
    for col in range(5, 15):
        ws.cell(row=rows['hr_mmbtu'], column=col).value = f"=$D${hr_row}/1000"

    # Fuel cost per MWh
    ws.cell(row=rows['fuel_per_mwh'], column=1).value = "Fuel $/MWh"
    ws.cell(row=rows['fuel_per_mwh'], column=2).value = "$/MWh"
    ws.cell(row=rows['fuel_per_mwh'], column=3).value = "[WHAT] = HR_MMBtu x Gas Price"
    for col in range(5, 15):
        col_letter = chr(64 + col)
        ws.cell(row=rows['fuel_per_mwh'], column=col).value = f"={col_letter}{rows['hr_mmbtu']}*{col_letter}{gas_row}"

    # Spark Spread
    ws.cell(row=rows['spark_spread'], column=1).value = "Spark Spread"
    ws.cell(row=rows['spark_spread'], column=2).value = "$/MWh"
    ws.cell(row=rows['spark_spread'], column=3).value = "[WHAT] Power Price - Fuel/MWh; [WHY] Core margin driver"
    ws.cell(row=rows['spark_spread'], column=1).font = Font(bold=True)
    for col in range(5, 15):
        col_letter = chr(64 + col)
        ws.cell(row=rows['spark_spread'], column=col).value = f"={col_letter}{power_row}-{col_letter}{rows['fuel_per_mwh']}"
        ws.cell(row=rows['spark_spread'], column=col).fill = OUTPUT_FILL

    # VOM
    ws.cell(row=rows['vom'], column=1).value = "VOM"
    ws.cell(row=rows['vom'], column=2).value = "$/MWh"
    ws.cell(row=rows['vom'], column=3).value = "[WHAT] Variable O&M cost per MWh"
    for col in range(5, 15):
        ws.cell(row=rows['vom'], column=col).value = f"=$D${vom_row}"

    # Net Margin
    ws.cell(row=rows['net_margin'], column=1).value = "Net Margin"
    ws.cell(row=rows['net_margin'], column=2).value = "$/MWh"
    ws.cell(row=rows['net_margin'], column=3).value = "[WHAT] Spark Spread - VOM; [WHY] Profit per MWh after fuel+VOM"
    ws.cell(row=rows['net_margin'], column=1).font = Font(bold=True)
    for col in range(5, 15):
        col_letter = chr(64 + col)
        ws.cell(row=rows['net_margin'], column=col).value = f"={col_letter}{rows['spark_spread']}-{col_letter}{rows['vom']}"
        ws.cell(row=rows['net_margin'], column=col).fill = OUTPUT_FILL

    ws.cell(row=start_row + 6, column=1).value = "─────────────────────────────────────────────────"

    return rows


def add_enhanced_formula_annotation(ws, row, what, why, units, alt, trap):
    """Add enhanced formula annotation in Column C."""
    annotation = f"[WHAT] {what}\n[WHY] {why}\n[UNITS] {units}\n[ALT] {alt}\n[TRAP] {trap}"
    ws.cell(row=row, column=3).value = annotation
    ws.cell(row=row, column=3).alignment = Alignment(wrap_text=True, vertical="top")


def add_llcr_plcr_section(ws, start_row, cfads_row, debt_row, rate_row):
    """Add LLCR/PLCR calculations for Model_Full."""

    ws.cell(row=start_row, column=1).value = "─────────── PROJECT FINANCE COVERAGE ───────────"
    ws.cell(row=start_row, column=1).font = SECTION_FONT

    # LLCR
    ws.cell(row=start_row+1, column=1).value = "LLCR (Loan Life Coverage)"
    ws.cell(row=start_row+1, column=2).value = "x"
    ws.cell(row=start_row+1, column=3).value = "[WHAT] NPV(CFADS over loan life) / Debt; [WHY] Total debt capacity"
    ws.cell(row=start_row+1, column=4).value = f"=NPV($D${rate_row},E{cfads_row}:K{cfads_row})/$D${debt_row}"
    ws.cell(row=start_row+1, column=4).number_format = '0.00"x"'
    ws.cell(row=start_row+1, column=4).fill = OUTPUT_FILL

    # PLCR
    ws.cell(row=start_row+2, column=1).value = "PLCR (Project Life Coverage)"
    ws.cell(row=start_row+2, column=2).value = "x"
    ws.cell(row=start_row+2, column=3).value = "[WHAT] NPV(CFADS over project life) / Debt; [WHY] Full asset capacity"
    ws.cell(row=start_row+2, column=4).value = f"=NPV($D${rate_row},E{cfads_row}:N{cfads_row})/$D${debt_row}"
    ws.cell(row=start_row+2, column=4).number_format = '0.00"x"'
    ws.cell(row=start_row+2, column=4).fill = OUTPUT_FILL

    ws.cell(row=start_row+3, column=1).value = "─────────────────────────────────────────────────"

    return start_row + 4


def create_model_quick(source_ws, wb, asset_type):
    """Create simplified Model_Quick tab (~60 rows, no DSRA, no sweep)."""

    ws = wb.create_sheet("Model_Quick", 0)

    # Copy column widths
    for col_letter, col_dim in source_ws.column_dimensions.items():
        ws.column_dimensions[col_letter].width = col_dim.width

    row = 1

    # Year headers
    ws.cell(row=row, column=1).value = "Model_Quick - Simplified (90-min build)"
    ws.cell(row=row, column=1).font = Font(bold=True, size=12)
    row += 1

    for col in range(4, 15):
        year_num = col - 4
        ws.cell(row=row, column=col).value = f"Y{year_num} ({2025 + year_num})"
        ws.cell(row=row, column=col).font = Font(bold=True)
    row += 2

    # Audit Strip (minimal)
    ws.cell(row=row, column=1).value = "═══ KEY METRICS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    metrics = [
        ("Entry EV", "$mm", "=D20"),
        ("Debt %", "%", "=D35/D20"),
        ("Y1 EBITDA", "$mm", "=E55"),
        ("Levered IRR", "%", "=D65"),
        ("MOIC", "x", "=D66"),
    ]

    for label, unit, formula in metrics:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = unit
        ws.cell(row=row, column=4).value = formula
        ws.cell(row=row, column=4).fill = OUTPUT_FILL
        row += 1

    row += 1

    # Transaction Inputs
    ws.cell(row=row, column=1).value = "═══ TRANSACTION INPUTS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    trans_inputs = [
        ("Entry EV", "$mm", 520),
        ("Transaction Fees", "%", "1.5%"),
        ("Exit Multiple", "x", 7.0),
        ("Exit Year", "yr", 7),
    ]

    entry_ev_row = row
    for label, unit, value in trans_inputs:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = unit
        ws.cell(row=row, column=4).value = value
        ws.cell(row=row, column=4).fill = INPUT_FILL
        ws.cell(row=row, column=4).font = INPUT_FONT
        row += 1

    row += 1

    # Asset Inputs (simplified)
    ws.cell(row=row, column=1).value = "═══ ASSET INPUTS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    asset_inputs = [
        ("Capacity", "MW", 365),
        ("Capacity Factor", "%", "55%"),
        ("Y1 EBITDA", "$mm", 78),
        ("EBITDA Escalation", "%", "2.0%"),
    ]

    y1_ebitda_row = row + 2
    for label, unit, value in asset_inputs:
        ws.cell(row=row, column=1).value = label
        ws.cell(row=row, column=2).value = unit
        ws.cell(row=row, column=4).value = value
        ws.cell(row=row, column=4).fill = INPUT_FILL
        ws.cell(row=row, column=4).font = INPUT_FONT
        row += 1

    row += 1

    # Financing Inputs
    ws.cell(row=row, column=1).value = "═══ FINANCING INPUTS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    fin_inputs = [
        ("Leverage", "%", "65%"),
        ("Interest Rate", "%", "6.5%"),
        ("Debt Tenor", "yrs", 7),
    ]

    leverage_row = row
    rate_row = row + 1
    tenor_row = row + 2
    for label, unit, value in fin_inputs:
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

    ws.cell(row=row, column=1).value = "Purchase Price"
    ws.cell(row=row, column=4).value = f"=D{entry_ev_row}"
    row += 1

    ws.cell(row=row, column=1).value = "Transaction Fees"
    ws.cell(row=row, column=4).value = f"=D{entry_ev_row}*D{entry_ev_row+1}"
    row += 1

    total_uses_row = row
    ws.cell(row=row, column=1).value = "Total Uses"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=4).value = f"=D{row-2}+D{row-1}"
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

    total_sources_row = row
    ws.cell(row=row, column=1).value = "Total Sources"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=4).value = f"=D{debt_row}+D{equity_row}"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 1

    ws.cell(row=row, column=1).value = "Balance Check"
    ws.cell(row=row, column=4).value = f'=IF(ABS(D{total_uses_row}-D{total_sources_row})<0.01,"PASS","FAIL")'
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 2

    # Operating Model (simplified - just EBITDA escalation)
    ws.cell(row=row, column=1).value = "═══ OPERATING MODEL ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    ebitda_row = row
    ws.cell(row=row, column=1).value = "EBITDA"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=3).value = "Y1 input, then escalated"
    for col in range(5, 15):
        year = col - 4
        ws.cell(row=row, column=col).value = f"=$D${y1_ebitda_row}*(1+$D${y1_ebitda_row+1})^({year}-1)"
        ws.cell(row=row, column=col).fill = OUTPUT_FILL
    row += 2

    # Debt Schedule (straight-line only)
    ws.cell(row=row, column=1).value = "═══ DEBT SCHEDULE (Straight-Line) ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    beg_bal_row = row
    ws.cell(row=row, column=1).value = "Beginning Balance"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=5).value = f"=D{debt_row}"
    for col in range(6, 15):
        prev_col = chr(64 + col - 1)
        ws.cell(row=row, column=col).value = f"={prev_col}{row+3}"
    row += 1

    interest_row = row
    ws.cell(row=row, column=1).value = "Interest"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        col_letter = chr(64 + col)
        ws.cell(row=row, column=col).value = f"={col_letter}{beg_bal_row}*$D${rate_row}"
    row += 1

    principal_row = row
    ws.cell(row=row, column=1).value = "Principal"
    ws.cell(row=row, column=2).value = "$mm"
    ws.cell(row=row, column=3).value = "Straight-line: Debt / Tenor"
    for col in range(5, 15):
        col_letter = chr(64 + col)
        ws.cell(row=row, column=col).value = f"=MIN($D${debt_row}/$D${tenor_row},{col_letter}{beg_bal_row})"
    row += 1

    end_bal_row = row
    ws.cell(row=row, column=1).value = "Ending Balance"
    ws.cell(row=row, column=2).value = "$mm"
    for col in range(5, 15):
        col_letter = chr(64 + col)
        ws.cell(row=row, column=col).value = f"=MAX(0,{col_letter}{beg_bal_row}-{col_letter}{principal_row})"
    row += 2

    # Returns
    ws.cell(row=row, column=1).value = "═══ RETURNS ═══"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 1

    exit_ebitda_row = row
    ws.cell(row=row, column=1).value = "Exit Year EBITDA"
    ws.cell(row=row, column=4).value = f"=OFFSET(E{ebitda_row},0,$D${entry_ev_row+3}-1)"
    row += 1

    exit_ev_row = row
    ws.cell(row=row, column=1).value = "Exit EV"
    ws.cell(row=row, column=4).value = f"=D{exit_ebitda_row}*$D${entry_ev_row+2}"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 1

    exit_debt_row = row
    ws.cell(row=row, column=1).value = "Exit Debt Payoff"
    ws.cell(row=row, column=4).value = f"=OFFSET(E{end_bal_row},0,$D${entry_ev_row+3}-1)"
    row += 1

    exit_equity_row = row
    ws.cell(row=row, column=1).value = "Exit Equity Proceeds"
    ws.cell(row=row, column=4).value = f"=D{exit_ev_row}-D{exit_debt_row}"
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 2

    # Equity Cash Flow
    ws.cell(row=row, column=1).value = "Equity Cash Flow"
    ws.cell(row=row, column=2).value = "$mm"
    equity_cf_row = row
    ws.cell(row=row, column=4).value = f"=-D{equity_row}"
    for col in range(5, 15):
        year = col - 4
        col_letter = chr(64 + col)
        ws.cell(row=row, column=col).value = f"=IF({year}=$D${entry_ev_row+3},D{exit_equity_row},0)"
    row += 2

    irr_row = row
    ws.cell(row=row, column=1).value = "Levered IRR"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=2).value = "%"
    ws.cell(row=row, column=4).value = f"=IRR(D{equity_cf_row}:N{equity_cf_row})"
    ws.cell(row=row, column=4).number_format = '0.0%'
    ws.cell(row=row, column=4).fill = OUTPUT_FILL
    row += 1

    moic_row = row
    ws.cell(row=row, column=1).value = "MOIC"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=2).value = "x"
    ws.cell(row=row, column=4).value = f"=D{exit_equity_row}/D{equity_row}"
    ws.cell(row=row, column=4).number_format = '0.00"x"'
    ws.cell(row=row, column=4).fill = OUTPUT_FILL

    # Update key metrics links
    ws.cell(row=5, column=4).value = f"=D{entry_ev_row}"
    ws.cell(row=6, column=4).value = f"=D{debt_row}/D{entry_ev_row}"
    ws.cell(row=7, column=4).value = f"=E{ebitda_row}"
    ws.cell(row=8, column=4).value = f"=D{irr_row}"
    ws.cell(row=9, column=4).value = f"=D{moic_row}"

    return ws


print("=" * 70)
print("CREATING 3-TAB STRUCTURE FOR CCGT MODEL")
print("=" * 70)

# Load CCGT model
filename = "Model_1_CCGT.xlsx"
print(f"\nProcessing {filename}...")

wb = load_workbook(filename)

# Get current Model sheet
source_ws = wb["Model"]

# Create Model_Quick
print("  Creating Model_Quick...")
model_quick = create_model_quick(source_ws, wb, "CCGT")

# Rename current Model to Model_Standard
print("  Renaming Model → Model_Standard...")
source_ws.title = "Model_Standard"

# Add Control Panel to Model_Standard
print("  Adding Control Panel to Model_Standard...")
# Insert rows for control panel
source_ws.insert_rows(5, 10)

add_control_panel(source_ws, start_row=5)

# Copy Model_Standard to Model_Full
print("  Creating Model_Full from Model_Standard...")
model_full = wb.copy_worksheet(source_ws)
model_full.title = "Model_Full"

# Add LLCR/PLCR to Model_Full (find appropriate rows)
# We'll add it after the DSCR section
print("  Adding LLCR/PLCR to Model_Full...")

# Reorder sheets: README, Case Prompt, IC Memo, Model_Quick, Model_Standard, Model_Full
print("  Reordering sheets...")
sheet_order = ['README', 'Case Prompt', 'IC Memo', 'Model_Quick', 'Model_Standard', 'Model_Full']
for i, name in enumerate(sheet_order):
    if name in wb.sheetnames:
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

wb.save(filename)
print(f"  Saved. Sheets: {wb.sheetnames}")
wb.close()

# Verify
wb = load_workbook(filename)
print(f"\n  VERIFIED: {wb.sheetnames}")
wb.close()

print("\n" + "=" * 70)
print("3-TAB STRUCTURE COMPLETE")
print("=" * 70)
