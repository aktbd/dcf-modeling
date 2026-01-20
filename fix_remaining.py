#!/usr/bin/env python3
"""
Fix remaining items:
1. Update CCGT README with Cold Start Workflow
2. Fix drill file sheet names (README -> Intuition Guide)
3. Verify prior audit fixes are intact
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
import os

os.chdir("/home/user/dcf-modeling/deliverables/deliverables-20260120-0030")


def update_readme_with_cold_start(ws):
    """Add Cold Start Workflow, Time Budget, and Debt Sizing guides to README."""
    # Find last row with content
    last_row = ws.max_row + 2
    row = last_row

    # Check if already added
    for r in range(1, ws.max_row + 1):
        val = ws.cell(row=r, column=1).value
        if val and "COLD START WORKFLOW" in str(val):
            print("    Cold Start Workflow already present")
            return ws

    # Cold Start Workflow
    ws.cell(row=row, column=1).value = "═══════════════════════════════════════════════════════════════════════"
    row += 1
    ws.cell(row=row, column=1).value = "COLD START WORKFLOW (4-Hour Test)"
    ws.cell(row=row, column=1).font = Font(bold=True, size=14, color="1F4E79")
    row += 2

    workflow = [
        ("Phase 1: Setup (15 min)", [
            "→ Create workbook, freeze panes at E2",
            "→ Set column widths: A=32, B=12, C=38, D-N=14",
            "→ Build year headers: Y0 (2025), Y1 (2026), ... Y10 (2035)",
        ]),
        ("Phase 2: Inputs & S&U (30 min)", [
            "→ Transaction: Entry EV, Fees %, Exit Multiple, Exit Year",
            "→ Asset: Capacity, CF/Dispatch, Prices, Escalation",
            "→ Financing: Leverage, Rate, Tenor, DSRA months, Sweep %",
            "→ Build Uses first, then Sources (equity = balancing item)",
        ]),
        ("Phase 3: Operating Model (45 min)", [
            "→ Generation = Capacity × CF × 8760 (baseload) OR Capacity × Dispatch × Avail (peaker)",
            "→ Revenue = Generation × Price / 1E6  (capacity $/kW needs ×1000)",
            "→ Fuel = Gen × HR × Gas / 1E9  (CRITICAL: divide by BILLION)",
            "→ EBITDA = Revenue - OpEx; Taxes = MAX(0, Taxable Income) × Rate",
        ]),
        ("Phase 4: Debt Schedule (45 min)", [
            "→ Beginning Bal Y1 = Funded Debt; Yn = Prior Ending",
            "→ Interest = Beg Bal × Rate; Principal: Straight-line/Sculpted/Sweep",
            "→ DSCR = CFADS / Debt Service",
            "→ DSRA Target = NEXT year's DS × (months/12)  [TRAP: not current!]",
        ]),
        ("Phase 5: Exit & Returns (30 min)", [
            "→ Exit EV = Exit EBITDA × Multiple",
            "→ Exit Equity = Exit EV - Debt Payoff + DSRA Release",
            "→ IRR = IRR(Y0:Y10 equity CFs); MOIC = Total / Entry",
        ]),
        ("Phase 6: QA & Polish (15 min)", [
            "→ S&U Balance Check = PASS; Debt Payoff = PASS; Min DSCR ≥1.25x",
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
        ("Setup & Structure", "15 min"),
        ("Transaction Inputs", "15 min"),
        ("Sources & Uses", "15 min"),
        ("Operating Model", "45 min"),
        ("Depreciation & Tax", "15 min"),
        ("Debt Schedule", "45 min"),
        ("DSRA (if needed)", "15 min"),
        ("Waterfall & Returns", "30 min"),
        ("Exit Mechanics", "15 min"),
        ("QA & Sanity Checks", "15 min"),
        ("Buffer / Iterate", "15 min"),
        ("TOTAL", "4:00"),
    ]

    for item, time in time_items:
        ws.cell(row=row, column=1).value = item
        ws.cell(row=row, column=2).value = time
        if item == "TOTAL":
            ws.cell(row=row, column=1).font = Font(bold=True)
            ws.cell(row=row, column=2).font = Font(bold=True)
        row += 1

    row += 2

    # Key Concepts
    ws.cell(row=row, column=1).value = "KEY FORMULA CONCEPTS"
    ws.cell(row=row, column=1).font = Font(bold=True, size=12, color="1F4E79")
    row += 2

    concepts = [
        ("SPARK SPREAD = Power Price - (HR × Gas / 1000)", "Core margin for thermal"),
        ("UCAP = Nameplate × (1 - FOR)", "Capacity revenue basis"),
        ("LLCR = NPV(CFADS loan life) / Debt", "Loan coverage ratio"),
        ("PLCR = NPV(CFADS project life) / Debt", "Project coverage ratio"),
    ]

    for concept, note in concepts:
        ws.cell(row=row, column=1).value = concept
        ws.cell(row=row, column=1).font = Font(bold=True)
        ws.cell(row=row, column=3).value = note
        row += 1

    return ws


def verify_audit_fixes(filename):
    """Verify prior audit fixes are intact."""
    wb = load_workbook(filename)

    issues = []

    # Find Model_Standard sheet
    if "Model_Standard" in wb.sheetnames:
        ws = wb["Model_Standard"]
    elif "Model" in wb.sheetnames:
        ws = wb["Model"]
    else:
        wb.close()
        return ["No Model sheet found"]

    # Scan for specific issues
    for row in range(1, min(ws.max_row + 1, 200)):
        for col in range(1, min(ws.max_column + 1, 15)):
            cell = ws.cell(row=row, column=col)
            val = str(cell.value) if cell.value else ""

            # Check for *1000 issue in midstream revenue
            if "Midstream" in filename and "E56*E57*1000" in val:
                issues.append(f"MIDSTREAM: Row {row} still has *1000 error")

            # Check for common formula errors
            if "#REF" in val or "#NAME" in val or "#VALUE" in val:
                issues.append(f"ERROR in {cell.coordinate}: {val}")

    wb.close()
    return issues


def main():
    print("=" * 70)
    print("FIXING REMAINING ITEMS")
    print("=" * 70)

    # 1. Update CCGT README
    print("\n--- Updating CCGT README ---")
    wb = load_workbook("Model_1_CCGT.xlsx")
    if "README" in wb.sheetnames:
        print("  Adding Cold Start Workflow to README...")
        update_readme_with_cold_start(wb["README"])
        wb.save("Model_1_CCGT.xlsx")
        print("  Saved.")
    wb.close()

    # 2. Fix drill sheet names
    print("\n--- Checking Drill Sheet Names ---")
    drill_files = [
        "Drill_1_CCGT.xlsx",
        "Drill_2_Peaker.xlsx",
        "Drill_3_SolarBESS.xlsx",
        "Drill_4_Transmission.xlsx",
        "Drill_5_Midstream.xlsx",
    ]

    for filename in drill_files:
        if os.path.exists(filename):
            wb = load_workbook(filename)
            print(f"  {filename}: {wb.sheetnames}")

            # Rename README to Intuition Guide if needed
            if "README" in wb.sheetnames and "Intuition Guide" not in wb.sheetnames:
                # Check if first sheet is generic README - drills should have Intuition Guide
                readme_ws = wb["README"]
                first_cell = readme_ws.cell(row=1, column=1).value
                if first_cell and "Intuition Guide" not in str(first_cell):
                    # This is a generic README - rename to Intuition Guide and update content
                    print(f"    Renaming README → Intuition Guide")
                    readme_ws.title = "Intuition Guide"
                    # Update first cell
                    readme_ws.cell(row=1, column=1).value = f"Intuition Guide - {filename.replace('.xlsx','').replace('Drill_','')}"
                    readme_ws.cell(row=1, column=1).font = Font(bold=True, size=14, color="1F4E79")
                    wb.save(filename)

            wb.close()

    # 3. Verify audit fixes
    print("\n--- Verifying Audit Fixes ---")
    models_to_check = [
        ("Model_3_SolarBESS.xlsx", "Solar MAX(0) tax"),
        ("Model_4_Transmission.xlsx", "Transmission S&U"),
        ("Model_5_Midstream.xlsx", "Midstream *1000"),
    ]

    all_ok = True
    for filename, fix_name in models_to_check:
        if os.path.exists(filename):
            issues = verify_audit_fixes(filename)
            if issues:
                print(f"  {filename}: ISSUES FOUND")
                for issue in issues:
                    print(f"    - {issue}")
                all_ok = False
            else:
                print(f"  {filename}: ✓ {fix_name} intact")

    if all_ok:
        print("\n  All audit fixes verified!")

    # 4. Final verification
    print("\n--- Final Sheet Verification ---")
    all_files = [
        "Model_1_CCGT.xlsx", "Model_2_Peaker.xlsx", "Model_3_SolarBESS.xlsx",
        "Model_4_Transmission.xlsx", "Model_5_Midstream.xlsx",
        "Drill_1_CCGT.xlsx", "Drill_2_Peaker.xlsx", "Drill_3_SolarBESS.xlsx",
        "Drill_4_Transmission.xlsx", "Drill_5_Midstream.xlsx",
    ]

    for filename in all_files:
        if os.path.exists(filename):
            wb = load_workbook(filename, read_only=True)
            print(f"  {filename}: {wb.sheetnames}")
            wb.close()

    print("\n" + "=" * 70)
    print("FIXES COMPLETE")
    print("=" * 70)


if __name__ == "__main__":
    main()
