#!/usr/bin/env python3
"""
Final Polish: Drill Blanking + Unlevered IRR + Sensitivity + IC Memo Templates
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
import os

os.chdir("/home/user/dcf-modeling/deliverables/deliverables-20260120-0030")

OUTPUT_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
SECTION_FONT = Font(bold=True, size=11)


def blank_drill_formulas(filename):
    """Blank key formulas in Drill file's Model_Standard tab."""
    print(f"\n  Processing {filename}...")
    wb = load_workbook(filename)

    if "Model_Standard" not in wb.sheetnames:
        print("    Model_Standard not found!")
        wb.close()
        return

    ws = wb["Model_Standard"]

    # Add DRILL MODE note at row 2
    ws.insert_rows(2)
    ws.cell(row=2, column=1).value = "═══ DRILL MODE: Key formulas blanked. Rebuild using Formula Logic hints in Column C. ═══"
    ws.cell(row=2, column=1).font = Font(bold=True, color="C00000", size=11)

    # Keywords to identify rows to blank
    blank_keywords = [
        # Revenue
        "energy revenue", "capacity revenue", "revenue", "total revenue",
        "gathering revenue", "tariff revenue", "ppa revenue",
        # Costs
        "fuel cost", "fuel ($mm)", "total costs", "total opex",
        # EBITDA
        "ebitda",
        # Tax
        "ebit", "taxes", "tax expense", "net taxes", "ptc applied", "pre-ptc tax",
        # CFADS
        "cfads",
        # Debt
        "interest", "principal", "sculpted principal", "capped principal",
        "ending balance", "dscr", "target debt service", "target ds",
        "cash sweep", "total principal",
        # Returns
        "exit ev", "exit equity", "exit proceeds", "levered irr", "irr", "moic",
        "equity cash flow", "distributable",
    ]

    # Rows to preserve (don't blank these even if they match keywords)
    preserve_keywords = [
        "label", "units", "entry", "input", "control", "audit", "qa",
        "sources", "uses", "section", "═══", "───",
    ]

    blanked_count = 0

    for row in range(1, min(ws.max_row + 1, 200)):
        label = ws.cell(row=row, column=1).value
        if not label:
            continue

        label_lower = str(label).lower().strip()

        # Check if this row should be preserved
        should_preserve = False
        for preserve in preserve_keywords:
            if preserve in label_lower:
                should_preserve = True
                break

        if should_preserve:
            continue

        # Check if this row should be blanked
        should_blank = False
        for keyword in blank_keywords:
            if keyword in label_lower:
                should_blank = True
                break

        if should_blank:
            # Blank columns E through N (Y1-Y10)
            for col in range(5, 15):
                cell = ws.cell(row=row, column=col)
                if cell.value and str(cell.value).startswith("="):
                    cell.value = ""
                    blanked_count += 1

            # Also blank D if it has a formula (Y0)
            cell_d = ws.cell(row=row, column=4)
            if cell_d.value and str(cell_d.value).startswith("="):
                # Keep some Y0 formulas (like entry equity)
                if "entry" not in label_lower and "source" not in label_lower and "use" not in label_lower:
                    cell_d.value = ""
                    blanked_count += 1

    wb.save(filename)
    print(f"    Blanked {blanked_count} formula cells")
    print(f"    Added DRILL MODE note at row 2")
    wb.close()


def add_unlevered_irr_and_sensitivity(filename):
    """Add Unlevered IRR row and sensitivity note to Model file."""
    print(f"\n  Processing {filename}...")
    wb = load_workbook(filename)

    if "Model_Standard" not in wb.sheetnames:
        print("    Model_Standard not found!")
        wb.close()
        return

    ws = wb["Model_Standard"]

    # Find the Levered IRR row
    irr_row = None
    moic_row = None
    cfads_row = None
    entry_ev_row = None
    exit_ev_row = None
    equity_cf_row = None

    for row in range(1, min(ws.max_row + 1, 200)):
        label = ws.cell(row=row, column=1).value
        if label:
            label_str = str(label).lower()
            if "levered irr" in label_str:
                irr_row = row
            if "moic" in label_str and moic_row is None:
                moic_row = row
            if "cfads" in label_str and cfads_row is None:
                cfads_row = row
            if "entry ev" in label_str and entry_ev_row is None:
                entry_ev_row = row
            if "exit ev" in label_str and exit_ev_row is None:
                exit_ev_row = row
            if "equity cash flow" in label_str:
                equity_cf_row = row

    if not irr_row:
        print("    Could not find Levered IRR row")
        wb.close()
        return

    # Check if Unlevered IRR already exists
    for row in range(irr_row - 5, irr_row + 5):
        if row > 0 and row <= ws.max_row:
            val = ws.cell(row=row, column=1).value
            if val and "unlevered irr" in str(val).lower():
                print("    Unlevered IRR already exists")
                # Just add sensitivity note
                add_sensitivity_note(ws)
                wb.save(filename)
                wb.close()
                return

    # Insert rows for Unlevered CF and Unlevered IRR after MOIC
    insert_row = moic_row + 1 if moic_row else irr_row + 1
    ws.insert_rows(insert_row, 3)

    # Unlevered CF row
    unlev_cf_row = insert_row + 1
    ws.cell(row=unlev_cf_row, column=1).value = "Unlevered CF"
    ws.cell(row=unlev_cf_row, column=2).value = "$mm"
    ws.cell(row=unlev_cf_row, column=3).value = "Y0: -Entry EV, Y1-n: CFADS, Exit: +Exit EV"

    # Build unlevered CF formula
    if entry_ev_row and cfads_row and exit_ev_row:
        # Y0: negative entry EV
        ws.cell(row=unlev_cf_row, column=4).value = f"=-D{entry_ev_row}"
        ws.cell(row=unlev_cf_row, column=4).number_format = '#,##0.0'

        # Y1-Y10: CFADS + Exit EV in exit year
        # Need to find exit year reference
        exit_yr_ref = None
        for row in range(1, 30):
            val = ws.cell(row=row, column=1).value
            if val and "exit year" in str(val).lower():
                exit_yr_ref = f"$D${row}"
                break

        for col in range(5, 15):
            yr = col - 4
            c = chr(64 + col)
            if exit_yr_ref:
                # CFADS + Exit EV if exit year
                ws.cell(row=unlev_cf_row, column=col).value = f"={c}{cfads_row}+IF({yr}={exit_yr_ref},D{exit_ev_row},0)"
            else:
                ws.cell(row=unlev_cf_row, column=col).value = f"={c}{cfads_row}"
            ws.cell(row=unlev_cf_row, column=col).number_format = '#,##0.0'

    # Unlevered IRR row
    unlev_irr_row = insert_row + 2
    ws.cell(row=unlev_irr_row, column=1).value = "Unlevered IRR"
    ws.cell(row=unlev_irr_row, column=1).font = Font(bold=True)
    ws.cell(row=unlev_irr_row, column=2).value = "%"
    ws.cell(row=unlev_irr_row, column=3).value = "Asset return before financing; typical 8-12%"
    ws.cell(row=unlev_irr_row, column=4).value = f"=IRR(D{unlev_cf_row}:N{unlev_cf_row})"
    ws.cell(row=unlev_irr_row, column=4).number_format = '0.0%'
    ws.cell(row=unlev_irr_row, column=4).fill = OUTPUT_FILL

    print(f"    Added Unlevered IRR at row {unlev_irr_row}")

    # Add sensitivity note
    add_sensitivity_note(ws)

    wb.save(filename)
    print(f"    Saved {filename}")
    wb.close()


def add_sensitivity_note(ws):
    """Add sensitivity rule-of-thumb note to Audit Strip area."""
    # Find empty cells in M-N columns near top
    ws.cell(row=3, column=13).value = "IRR Sensitivity:"
    ws.cell(row=3, column=13).font = Font(bold=True, size=9)
    ws.cell(row=4, column=13).value = "+/- 5% Entry → ~+/- 1% IRR"
    ws.cell(row=4, column=13).font = Font(size=9, color="666666")
    ws.cell(row=5, column=13).value = "+/- 0.5x Exit → ~+/- 1.5% IRR"
    ws.cell(row=5, column=13).font = Font(size=9, color="666666")
    print("    Added sensitivity note")


def trim_formula_logic(filename):
    """Trim Formula Logic column to max 4 lines in Model_Standard."""
    print(f"\n  Trimming formula logic in {filename}...")
    wb = load_workbook(filename)

    if "Model_Standard" not in wb.sheetnames:
        wb.close()
        return

    ws = wb["Model_Standard"]

    trimmed_count = 0
    for row in range(1, min(ws.max_row + 1, 200)):
        cell = ws.cell(row=row, column=3)  # Column C
        if cell.value and len(str(cell.value)) > 150:
            # Trim to first 4 meaningful lines
            lines = str(cell.value).split('\n')
            # Keep lines that are not empty and not just whitespace
            meaningful_lines = [l.strip() for l in lines if l.strip()][:4]
            cell.value = '\n'.join(meaningful_lines)
            trimmed_count += 1

    wb.save(filename)
    print(f"    Trimmed {trimmed_count} cells")
    wb.close()


def main():
    print("=" * 70)
    print("FINAL POLISH: Drills + Unlevered IRR + Sensitivity")
    print("=" * 70)

    # Process DRILL files - blank formulas
    print("\n--- BLANKING DRILL FORMULAS ---")
    drill_files = [
        "Drill_1_CCGT.xlsx",
        "Drill_2_Peaker.xlsx",
        "Drill_3_SolarBESS.xlsx",
        "Drill_4_Transmission.xlsx",
        "Drill_5_Midstream.xlsx",
        "Drill_6_Wind.xlsx",
    ]

    for filename in drill_files:
        if os.path.exists(filename):
            blank_drill_formulas(filename)

    # Process MODEL files - add Unlevered IRR and sensitivity
    print("\n--- ADDING UNLEVERED IRR & SENSITIVITY ---")
    model_files = [
        "Model_1_CCGT.xlsx",
        "Model_2_Peaker.xlsx",
        "Model_3_SolarBESS.xlsx",
        "Model_4_Transmission.xlsx",
        "Model_5_Midstream.xlsx",
        "Model_6_Wind.xlsx",
    ]

    for filename in model_files:
        if os.path.exists(filename):
            add_unlevered_irr_and_sensitivity(filename)

    # Trim formula logic in Model files
    print("\n--- TRIMMING FORMULA LOGIC ---")
    for filename in model_files:
        if os.path.exists(filename):
            trim_formula_logic(filename)

    print("\n" + "=" * 70)
    print("DRILL & MODEL UPDATES COMPLETE")
    print("=" * 70)


if __name__ == "__main__":
    main()
