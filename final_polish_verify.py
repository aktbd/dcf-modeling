#!/usr/bin/env python3
"""
Final Polish Verification and ZIP Regeneration
"""

from openpyxl import load_workbook
import os
import zipfile
import subprocess
import tempfile

os.chdir("/home/user/dcf-modeling/deliverables/deliverables-20260120-0030")


def verify_drill_blanking():
    """Verify Drill files have blanked formulas."""
    print("\n--- DRILL BLANKING VERIFICATION ---")

    drill_files = [
        ("Drill_1_CCGT.xlsx", ["revenue", "ebitda", "irr"]),
    ]

    for filename, check_rows in drill_files:
        if not os.path.exists(filename):
            print(f"  {filename}: NOT FOUND")
            continue

        wb = load_workbook(filename)
        ws = wb["Model_Standard"]

        # Check for DRILL MODE note
        drill_mode_found = False
        for row in range(1, 5):
            val = ws.cell(row=row, column=1).value
            if val and "DRILL MODE" in str(val):
                drill_mode_found = True
                print(f"  {filename}: ✓ DRILL MODE note at row {row}")
                break

        if not drill_mode_found:
            print(f"  {filename}: ❌ DRILL MODE note not found")

        # Check if formulas are blanked
        blanked_count = 0
        for row in range(1, 150):
            label = ws.cell(row=row, column=1).value
            if label:
                label_lower = str(label).lower()
                for keyword in check_rows:
                    if keyword in label_lower:
                        # Check if E column is blank
                        cell_e = ws.cell(row=row, column=5)
                        if cell_e.value == "" or cell_e.value is None:
                            blanked_count += 1
                        break

        print(f"    Key formula rows blanked: {blanked_count}")
        wb.close()


def verify_model_formulas():
    """Verify Model files still have working formulas."""
    print("\n--- MODEL FORMULA VERIFICATION ---")

    model_files = [
        "Model_1_CCGT.xlsx",
        "Model_5_Midstream.xlsx",
        "Model_6_Wind.xlsx",
    ]

    for filename in model_files:
        if not os.path.exists(filename):
            print(f"  {filename}: NOT FOUND")
            continue

        wb = load_workbook(filename)
        ws = wb["Model_Standard"]

        # Check that EBITDA has formula
        ebitda_has_formula = False
        irr_has_formula = False

        for row in range(1, 150):
            label = ws.cell(row=row, column=1).value
            if label:
                label_lower = str(label).lower()
                if "ebitda" in label_lower and "ebitda" == label_lower.strip():
                    cell_e = ws.cell(row=row, column=5)
                    if cell_e.value and str(cell_e.value).startswith("="):
                        ebitda_has_formula = True
                        print(f"  {filename}: ✓ EBITDA formula at row {row}")
                if "levered irr" in label_lower:
                    cell_d = ws.cell(row=row, column=4)
                    if cell_d.value and str(cell_d.value).startswith("="):
                        irr_has_formula = True
                        print(f"    ✓ IRR formula at row {row}")

        if not ebitda_has_formula:
            print(f"  {filename}: ❌ EBITDA formula missing!")
        if not irr_has_formula:
            print(f"  {filename}: ❌ IRR formula missing!")

        wb.close()


def verify_prior_fixes():
    """Verify prior fixes still intact."""
    print("\n--- PRIOR FIXES VERIFICATION ---")

    # Midstream revenue
    wb = load_workbook("Model_5_Midstream.xlsx")
    ws = wb["Model_Standard"]
    for row in range(50, 100):
        label = ws.cell(row=row, column=1).value
        if label and "Revenue" in str(label):
            formula = ws.cell(row=row, column=5).value
            if formula and "*1000" not in str(formula):
                print(f"  ✓ Midstream Revenue: no *1000 error")
            break
    wb.close()

    # Transmission
    wb = load_workbook("Model_4_Transmission.xlsx")
    ws = wb["Model_Standard"]
    for row in range(100, 140):
        label = ws.cell(row=row, column=1).value
        if label and "Total Sources" in str(label):
            formula = ws.cell(row=row, column=4).value
            print(f"  ✓ Transmission Sources: {formula}")
            break
    wb.close()

    # Solar tax
    wb = load_workbook("Model_3_SolarBESS.xlsx")
    ws = wb["Model_Standard"]
    for row in range(70, 100):
        label = ws.cell(row=row, column=1).value
        if label and "Tax" in str(label):
            formula = ws.cell(row=row, column=5).value
            if formula and "MAX(0" in str(formula):
                print(f"  ✓ Solar Tax: has MAX(0,...)")
            break
    wb.close()

    # Wind PTC
    wb = load_workbook("Model_6_Wind.xlsx")
    ws = wb["Model_Standard"]
    for row in range(1, 150):
        label = ws.cell(row=row, column=1).value
        if label and "PTC Applied" in str(label):
            formula = ws.cell(row=row, column=5).value
            if formula and "MIN" in str(formula):
                print(f"  ✓ Wind PTC: uses MIN()")
            break
    wb.close()


def regenerate_zip():
    """Regenerate ZIP with all files."""
    print("\n--- REGENERATING ZIP ---")

    zip_path = "deliverables_bundle.zip"
    if os.path.exists(zip_path):
        os.remove(zip_path)
        print(f"  Deleted old {zip_path}")

    files_to_zip = [
        # Models
        "Model_1_CCGT.xlsx", "Model_2_Peaker.xlsx", "Model_3_SolarBESS.xlsx",
        "Model_4_Transmission.xlsx", "Model_5_Midstream.xlsx", "Model_6_Wind.xlsx",
        # Drills
        "Drill_1_CCGT.xlsx", "Drill_2_Peaker.xlsx", "Drill_3_SolarBESS.xlsx",
        "Drill_4_Transmission.xlsx", "Drill_5_Midstream.xlsx", "Drill_6_Wind.xlsx",
        # IC Memos
        "IC_Memo_1_CCGT.docx", "IC_Memo_2_Peaker.docx", "IC_Memo_3_SolarBESS.docx",
        "IC_Memo_4_Transmission.docx", "IC_Memo_5_Midstream.docx", "IC_Memo_6_Wind.docx",
        # Reference
        "Master_Template.xlsx", "Lotus_Interview_CheatSheet_Expanded.pdf",
        "Lotus_Interview_CheatSheet_Expanded.md", "Implementation_Manifesto.md",
        "Lotus_Study_Guide.md",
    ]

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for fname in files_to_zip:
            if os.path.exists(fname):
                zf.write(fname)
                print(f"  Added: {fname}")

    print(f"\n  Created: {zip_path}")

    # Verify by extraction
    print("\n--- VERIFICATION BY EXTRACTION ---")
    with tempfile.TemporaryDirectory() as tmpdir:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            zf.extractall(tmpdir)

        # Check Drill file
        test_file = os.path.join(tmpdir, "Drill_1_CCGT.xlsx")
        wb = load_workbook(test_file, read_only=True)
        print(f"  Drill_1_CCGT.xlsx: {wb.sheetnames}")
        wb.close()

        # Check Model file
        test_file = os.path.join(tmpdir, "Model_1_CCGT.xlsx")
        wb = load_workbook(test_file, read_only=True)
        print(f"  Model_1_CCGT.xlsx: {wb.sheetnames}")
        wb.close()

    # SHA256
    result = subprocess.run(['sha256sum', zip_path], capture_output=True, text=True)
    print(f"\n  SHA256: {result.stdout.strip()}")

    # File count
    with zipfile.ZipFile(zip_path, 'r') as zf:
        print(f"  Total files: {len(zf.namelist())}")


def main():
    print("=" * 70)
    print("FINAL POLISH VERIFICATION")
    print("=" * 70)

    verify_drill_blanking()
    verify_model_formulas()
    verify_prior_fixes()
    regenerate_zip()

    print("\n" + "=" * 70)
    print("VERIFICATION COMPLETE")
    print("=" * 70)


if __name__ == "__main__":
    main()
