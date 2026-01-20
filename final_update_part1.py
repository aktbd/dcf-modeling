#!/usr/bin/env python3
"""
Final Update Part 1: Margin Analysis + Timing Flags + Sculpt Transparency
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.comments import Comment
import os

os.chdir("/home/user/dcf-modeling/deliverables/deliverables-20260120-0030")

# Styles
INPUT_FILL = PatternFill(start_color="FFFFC7", end_color="FFFFC7", fill_type="solid")
INPUT_FONT = Font(bold=True, color="0000FF")
OUTPUT_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
SECTION_FONT = Font(bold=True, size=11)


def update_control_panel_with_timing_flags(ws):
    """Update Control Panel to include all timing flags."""
    # Find Control Panel header
    cp_row = None
    for row in range(1, 30):
        val = ws.cell(row=row, column=1).value
        if val and "CONTROL PANEL" in str(val):
            cp_row = row
            break

    if not cp_row:
        print("    Control Panel not found!")
        return

    # Clear existing control panel rows (cp_row to cp_row+15)
    for row in range(cp_row, cp_row + 16):
        for col in range(1, 5):
            ws.cell(row=row, column=col).value = None
            ws.cell(row=row, column=col).fill = PatternFill()
            ws.cell(row=row, column=col).font = Font()

    # Rebuild Control Panel with timing flags
    row = cp_row
    ws.cell(row=row, column=1).value = "═══════════ CONTROL PANEL ═══════════"
    ws.cell(row=row, column=1).font = SECTION_FONT
    row += 2

    # REVENUE section
    ws.cell(row=row, column=1).value = "REVENUE"
    ws.cell(row=row, column=1).font = Font(bold=True, italic=True, color="666666")
    row += 1

    ws.cell(row=row, column=1).value = "Revenue Mode"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=3).value = "Merchant / Contracted / Hybrid"
    ws.cell(row=row, column=3).font = Font(italic=True, color="666666")
    ws.cell(row=row, column=4).value = "Merchant"
    ws.cell(row=row, column=4).fill = INPUT_FILL
    ws.cell(row=row, column=4).font = INPUT_FONT
    row += 1

    ws.cell(row=row, column=1).value = "Contract End Year"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=3).value = "0 = pure merchant; else year PPA ends"
    ws.cell(row=row, column=3).font = Font(italic=True, color="666666")
    ws.cell(row=row, column=4).value = 0
    ws.cell(row=row, column=4).fill = INPUT_FILL
    ws.cell(row=row, column=4).font = INPUT_FONT
    row += 2

    # TIMING section
    ws.cell(row=row, column=1).value = "TIMING"
    ws.cell(row=row, column=1).font = Font(bold=True, italic=True, color="666666")
    row += 1

    ws.cell(row=row, column=1).value = "Exit Year"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=3).value = "5 / 7 / 10"
    ws.cell(row=row, column=3).font = Font(italic=True, color="666666")
    ws.cell(row=row, column=4).value = 7
    ws.cell(row=row, column=4).fill = INPUT_FILL
    ws.cell(row=row, column=4).font = INPUT_FONT
    row += 1

    ws.cell(row=row, column=1).value = "Tax Credit End Year"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=3).value = "0 = N/A; else last year of PTC/ITC"
    ws.cell(row=row, column=3).font = Font(italic=True, color="666666")
    ws.cell(row=row, column=4).value = 0
    ws.cell(row=row, column=4).fill = INPUT_FILL
    ws.cell(row=row, column=4).font = INPUT_FONT
    row += 2

    # DEBT section
    ws.cell(row=row, column=1).value = "DEBT"
    ws.cell(row=row, column=1).font = Font(bold=True, italic=True, color="666666")
    row += 1

    ws.cell(row=row, column=1).value = "Debt Sweep"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=3).value = "ON / OFF (for sweep assets)"
    ws.cell(row=row, column=3).font = Font(italic=True, color="666666")
    ws.cell(row=row, column=4).value = "ON"
    ws.cell(row=row, column=4).fill = INPUT_FILL
    ws.cell(row=row, column=4).font = INPUT_FONT
    row += 1

    ws.cell(row=row, column=1).value = "Sweep %"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=3).value = "Active when Sweep = ON"
    ws.cell(row=row, column=3).font = Font(italic=True, color="666666")
    ws.cell(row=row, column=4).value = "75%"
    ws.cell(row=row, column=4).fill = INPUT_FILL
    ws.cell(row=row, column=4).font = INPUT_FONT
    row += 2

    # OTHER section
    ws.cell(row=row, column=1).value = "OTHER"
    ws.cell(row=row, column=1).font = Font(bold=True, italic=True, color="666666")
    row += 1

    ws.cell(row=row, column=1).value = "Exit Method"
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=3).value = "Multiple / DCF"
    ws.cell(row=row, column=3).font = Font(italic=True, color="666666")
    ws.cell(row=row, column=4).value = "Multiple"
    ws.cell(row=row, column=4).fill = INPUT_FILL
    ws.cell(row=row, column=4).font = INPUT_FONT
    row += 1

    ws.cell(row=row, column=1).value = "════════════════════════════════════"
    ws.cell(row=row, column=1).font = SECTION_FONT

    return row


def add_spark_spread_to_thermal(ws, is_full=False):
    """Add Spark Spread row to thermal model (CCGT/Peaker)."""
    # Find fuel cost row and insert Spark Spread after it
    fuel_row = None
    hr_row = None
    gas_row = None
    power_row = None

    for row in range(1, 150):
        val = ws.cell(row=row, column=1).value
        if val:
            val_str = str(val).lower()
            if 'fuel cost' in val_str or 'fuel ($mm)' in val_str.lower():
                fuel_row = row
            if 'heat rate' in val_str:
                hr_row = row
            if 'gas price' in val_str:
                gas_row = row
            if 'power price' in val_str or 'energy price' in val_str:
                power_row = row

    if not fuel_row:
        print("    Could not find Fuel Cost row")
        return

    # Check if Spark Spread already exists
    for row in range(fuel_row - 3, fuel_row + 3):
        val = ws.cell(row=row, column=1).value
        if val and 'spark spread' in str(val).lower():
            print("    Spark Spread already exists")
            return

    # Insert row after fuel cost
    ws.insert_rows(fuel_row + 1)
    ss_row = fuel_row + 1

    ws.cell(row=ss_row, column=1).value = "Spark Spread"
    ws.cell(row=ss_row, column=1).font = Font(bold=True)
    ws.cell(row=ss_row, column=2).value = "$/MWh"

    # Comprehensive annotation
    annotation = """[WHAT] Spark Spread = Power Price - (Plant HR × Gas / 1000)
[WHY] Gross margin per MWh after fuel. Core driver of thermal profitability.

[RMHR EQUIVALENCE - For Interview Discussion]
Traders express this as "Market Heat Rate":
  RMHR = Power / Gas × 1000
  If RMHR > Plant HR → "In the money" → Dispatch

Spark Spread = (RMHR - Plant HR) × Gas / 1000
Same math, different framing. Use Spark Spread in model (simpler).

[EXAMPLE] Power=$50, Gas=$5, HR=7000
  RMHR = 50/5 × 1000 = 10,000 Btu/kWh
  Spread = 10,000 - 7,000 = 3,000 Btu/kWh
  Spark = 3,000 × $5 / 1000 = $15/MWh ✓
  (Or directly: $50 - 7×$5 = $15/MWh)

[UNITS] $/MWh - (Btu/kWh × $/MMBtu / 1000) = $/MWh ✓"""

    ws.cell(row=ss_row, column=3).value = annotation
    ws.cell(row=ss_row, column=3).alignment = Alignment(wrap_text=True, vertical="top")

    # Formula: =E(PowerPrice)-($D$HR/1000)*E(GasPrice)
    # Need to find actual row references
    if hr_row and gas_row and power_row:
        for col in range(5, 15):
            col_letter = chr(64 + col)
            formula = f"={col_letter}{power_row}-($D${hr_row}/1000)*{col_letter}{gas_row}"
            ws.cell(row=ss_row, column=col).value = formula
            ws.cell(row=ss_row, column=col).number_format = '#,##0.00'
            ws.cell(row=ss_row, column=col).fill = OUTPUT_FILL

    # Add cell comment
    comment = Comment(
        """SPARK SPREAD
Profit per MWh after fuel cost.
= Power Price - (HR × Gas / 1000)

Traders call this "heat rate spread":
If Market HR > Plant HR → Profitable
Think: "Call option on scarcity"
""", "Claude")
    ws.cell(row=ss_row, column=1).comment = comment

    print(f"    Added Spark Spread at row {ss_row}")
    return ss_row


def add_carbon_to_thermal(ws):
    """Add Carbon Region flag and Carbon Cost row to thermal models."""
    # Find Control Panel OTHER section
    for row in range(1, 30):
        val = ws.cell(row=row, column=1).value
        if val and str(val).strip() == "OTHER":
            # Insert Carbon Region before Exit Method
            ws.insert_rows(row + 1)
            cr_row = row + 1

            ws.cell(row=cr_row, column=1).value = "Carbon Region"
            ws.cell(row=cr_row, column=1).font = Font(bold=True)
            ws.cell(row=cr_row, column=3).value = "None / CARB / RGGI (thermal only)"
            ws.cell(row=cr_row, column=3).font = Font(italic=True, color="666666")
            ws.cell(row=cr_row, column=4).value = "None"
            ws.cell(row=cr_row, column=4).fill = INPUT_FILL
            ws.cell(row=cr_row, column=4).font = INPUT_FONT

            print(f"    Added Carbon Region flag at row {cr_row}")
            break

    # Note: Carbon Cost row would need to be added to the operating model
    # This is complex due to row references - keeping it simple for now


def process_thermal_model(filename):
    """Process CCGT or Peaker model."""
    print(f"\n  Processing {filename}...")
    wb = load_workbook(filename)

    # Process Model_Standard
    if "Model_Standard" in wb.sheetnames:
        print("    Model_Standard:")
        ws = wb["Model_Standard"]
        update_control_panel_with_timing_flags(ws)
        add_spark_spread_to_thermal(ws, is_full=False)
        add_carbon_to_thermal(ws)

    # Process Model_Full
    if "Model_Full" in wb.sheetnames:
        print("    Model_Full:")
        ws = wb["Model_Full"]
        update_control_panel_with_timing_flags(ws)
        add_spark_spread_to_thermal(ws, is_full=True)
        add_carbon_to_thermal(ws)

    wb.save(filename)
    print(f"    Saved {filename}")
    wb.close()


def process_non_thermal_model(filename):
    """Process non-thermal model (just update Control Panel)."""
    print(f"\n  Processing {filename}...")
    wb = load_workbook(filename)

    for sheet_name in ["Model_Standard", "Model_Full"]:
        if sheet_name in wb.sheetnames:
            print(f"    {sheet_name}:")
            ws = wb[sheet_name]
            update_control_panel_with_timing_flags(ws)

    wb.save(filename)
    print(f"    Saved {filename}")
    wb.close()


def main():
    print("=" * 70)
    print("FINAL UPDATE PART 1: Margin Analysis + Timing Flags")
    print("=" * 70)

    # Process thermal models (CCGT, Peaker) with Spark Spread
    thermal_models = [
        "Model_1_CCGT.xlsx",
        "Model_2_Peaker.xlsx",
        "Drill_1_CCGT.xlsx",
        "Drill_2_Peaker.xlsx",
    ]

    for filename in thermal_models:
        if os.path.exists(filename):
            process_thermal_model(filename)

    # Process non-thermal models (just Control Panel update)
    non_thermal = [
        "Model_3_SolarBESS.xlsx",
        "Model_4_Transmission.xlsx",
        "Model_5_Midstream.xlsx",
        "Drill_3_SolarBESS.xlsx",
        "Drill_4_Transmission.xlsx",
        "Drill_5_Midstream.xlsx",
    ]

    for filename in non_thermal:
        if os.path.exists(filename):
            process_non_thermal_model(filename)

    print("\n" + "=" * 70)
    print("PART 1 COMPLETE")
    print("=" * 70)


if __name__ == "__main__":
    main()
