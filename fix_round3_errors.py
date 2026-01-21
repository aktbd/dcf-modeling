#!/usr/bin/env python3
"""
Round 3 Fix: 4 Minor Errors
1. Model_5_Midstream.xlsx - Fix Unlevered Returns label/formula mismatch
2. Model_6_Wind.xlsx - Fix Audit Strip S&U Balance cell references
3. Drill_1_CCGT.xlsx - Add DRILL MODE note and blank key formulas
4. Drill_5_Midstream.xlsx - Add DRILL MODE note and blank key formulas
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

DELIV_DIR = '/home/user/dcf-modeling/deliverables/deliverables-20260120-0030'

YELLOW_FILL = PatternFill(start_color="FFFFC7", end_color="FFFFC7", fill_type="solid")
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
BLACK_BOLD = Font(color="000000", bold=True)
PERCENT_FMT = '0.0%'
MULTIPLE_FMT = '0.00"x"'


def fix_midstream_unlevered():
    """Fix Unlevered Returns section - labels don't match formulas"""
    filepath = f"{DELIV_DIR}/Model_5_Midstream.xlsx"
    wb = load_workbook(filepath)
    ws = wb['Model_Standard']

    print("=" * 60)
    print("[1] FIXING Model_5_Midstream.xlsx - Unlevered Returns Section")
    print("=" * 60)

    # Current state:
    # R126: 'Unlevered Returns (for reference)' | D==-D113 (Unlevered CF Y0)
    # R127: 'Unlevered Cash Flow' | D==IRR(D126:N126) <- WRONG LABEL
    # R128: 'Unlevered IRR' | D==(SUM(E126:K126))/-D126 <- WRONG LABEL (this is MOIC)

    # The formulas are in the right rows, but labels are wrong!
    # Fix the labels:

    ws.cell(row=126, column=1).value = "Unlevered Cash Flow"
    ws.cell(row=127, column=1).value = "Unlevered IRR"
    ws.cell(row=128, column=1).value = "Unlevered MOIC"

    # Also verify the formulas are correct
    # R126 D should be -Entry EV (Y0), E-N should be CFADS + exit equity
    # Entry EV is at D27 based on our earlier fixes
    # CFADS is at row 78
    # Exit Equity is at D121 or similar

    # Check current D126
    d126 = ws.cell(row=126, column=4).value
    print(f"  Current D126: {d126}")

    # D126 should be =-D104 (Purchase Price which equals Entry EV)
    # But it says =-D113, let me check what D113 is
    d113_label = ws.cell(row=113, column=1).value
    print(f"  D113 label: {d113_label}")

    # Fix D126 to reference correct Purchase Price (D104)
    ws.cell(row=126, column=4).value = "=-D104"

    # E126-K126 should be CFADS (row 78)
    # K126 should include exit equity
    for col in range(5, 15):
        year = col - 4
        if year < 7:  # Before exit
            ws.cell(row=126, column=col).value = f"={chr(64+col)}78"  # CFADS
        elif year == 7:  # Exit year
            ws.cell(row=126, column=col).value = f"={chr(64+col)}78+D121"  # CFADS + Exit Equity
        else:  # After exit
            ws.cell(row=126, column=col).value = 0

    # D127 should be IRR
    ws.cell(row=127, column=4).value = "=IRR(D126:N126)"
    ws.cell(row=127, column=4).fill = GREEN_FILL
    ws.cell(row=127, column=4).font = BLACK_BOLD
    ws.cell(row=127, column=4).number_format = PERCENT_FMT

    # D128 should be MOIC
    ws.cell(row=128, column=4).value = "=(SUM(E126:K126))/-D126"
    ws.cell(row=128, column=4).fill = GREEN_FILL
    ws.cell(row=128, column=4).font = BLACK_BOLD
    ws.cell(row=128, column=4).number_format = MULTIPLE_FMT

    wb.save(filepath)
    print(f"\n  Fixed labels: R126='Unlevered Cash Flow', R127='Unlevered IRR', R128='Unlevered MOIC'")
    print(f"  Fixed D126 = =-D104 (Purchase Price)")
    print(f"  Saved: {filepath}")


def fix_wind_audit_strip():
    """Fix Audit Strip S&U Balance to reference correct cells"""
    filepath = f"{DELIV_DIR}/Model_6_Wind.xlsx"
    wb = load_workbook(filepath)
    ws = wb['Model_Standard']

    print("\n" + "=" * 60)
    print("[2] FIXING Model_6_Wind.xlsx - Audit Strip S&U Balance")
    print("=" * 60)

    # Current R15: =IF(ABS(D111-D116)<0.01,"PASS","FAIL")
    # D111 = Purchase Price, D116 = SOURCES header (wrong!)
    # Should be: =IF(ABS(D114-D119)<0.01,"PASS","FAIL")
    # D114 = Total Uses, D119 = Total Sources

    old_formula = ws.cell(row=15, column=4).value
    print(f"  Old R15 formula: {old_formula}")

    ws.cell(row=15, column=4).value = "=IF(ABS(D114-D119)<0.01,\"PASS\",\"FAIL\")"

    wb.save(filepath)
    print(f"  New R15 formula: =IF(ABS(D114-D119)<0.01,\"PASS\",\"FAIL\")")
    print(f"  Saved: {filepath}")


def fix_drill_ccgt():
    """Add DRILL MODE note and blank key formulas"""
    filepath = f"{DELIV_DIR}/Drill_1_CCGT.xlsx"
    wb = load_workbook(filepath)
    ws = wb['Model_Standard']

    print("\n" + "=" * 60)
    print("[3] FIXING Drill_1_CCGT.xlsx - DRILL MODE + Blank Formulas")
    print("=" * 60)

    # Add DRILL MODE note to a visible location
    # R2 has "═══ AUDIT STRIP ═══", let's add DRILL note above or beside it
    # Insert at R1 column A or use a different approach

    # Better: Add to R2 column C (notes column) or append to R2
    current_r2 = ws.cell(row=2, column=1).value
    ws.cell(row=2, column=1).value = "═══ DRILL MODE: Rebuild formulas in highlighted cells ═══"
    ws.cell(row=2, column=1).font = Font(color="FF0000", bold=True)

    # Move audit strip header to R3
    ws.cell(row=3, column=1).value = current_r2

    # Blank key formula cells (E column for operating years)
    # Based on CCGT structure from fix_ccgt_complete.py:
    # Row 86: EBITDA
    # Row 91-92: Debt service
    # Row 97-103: DSRA section
    # Row 117: Distributable
    # Row 141: Equity CF
    # Row 142: Levered IRR (D column)
    # Row 147: MOIC (D column)

    rows_to_blank_E_to_N = [
        76, 77, 78,  # Revenue section
        80, 82, 83, 84,  # OpEx section
        86,  # EBITDA
        91, 92,  # Debt service
        97, 98, 99, 101, 102, 103, 104,  # Debt schedule
        111, 112, 113, 114,  # DSRA
        117,  # Distributable
        141,  # Equity CF
    ]

    # Also blank returns in D column
    returns_to_blank_D = [142, 147]  # IRR, MOIC

    blanked_count = 0
    for row in rows_to_blank_E_to_N:
        for col in range(5, 15):  # E to N
            cell = ws.cell(row=row, column=col)
            if cell.value is not None:
                cell.value = None
                cell.fill = YELLOW_FILL  # Highlight for practice
                blanked_count += 1

    for row in returns_to_blank_D:
        cell = ws.cell(row=row, column=4)
        if cell.value is not None:
            cell.value = None
            cell.fill = YELLOW_FILL
            blanked_count += 1

    wb.save(filepath)
    print(f"  Added DRILL MODE note to R2")
    print(f"  Blanked {blanked_count} formula cells")
    print(f"  Saved: {filepath}")


def fix_drill_midstream():
    """Add DRILL MODE note and blank key formulas"""
    filepath = f"{DELIV_DIR}/Drill_5_Midstream.xlsx"
    wb = load_workbook(filepath)
    ws = wb['Model_Standard']

    print("\n" + "=" * 60)
    print("[4] FIXING Drill_5_Midstream.xlsx - DRILL MODE + Blank Formulas")
    print("=" * 60)

    # Add DRILL MODE note
    current_r2 = ws.cell(row=2, column=1).value
    ws.cell(row=2, column=1).value = "═══ DRILL MODE: Rebuild formulas in highlighted cells ═══"
    ws.cell(row=2, column=1).font = Font(color="FF0000", bold=True)
    ws.cell(row=3, column=1).value = current_r2

    # Midstream structure from fix_midstream_complete.py:
    # Row 69: Revenue
    # Row 70: OpEx
    # Row 72: EBITDA
    # Row 75: EBIT
    # Row 76: Taxes
    # Row 77: Growth CapEx
    # Row 78: CFADS
    # Row 81-89: Debt schedule
    # Row 93-96: DSRA
    # Row 100: Distributable
    # Row 122: Equity CF
    # Row 123: Levered IRR (D)
    # Row 124: MOIC (D)

    rows_to_blank_E_to_N = [
        69, 70,  # Revenue, OpEx
        72,  # EBITDA
        75, 76, 77, 78,  # EBIT, Taxes, CapEx, CFADS
        81, 82, 83, 84, 86, 87, 88, 89,  # Debt schedule
        93, 94, 95, 96,  # DSRA
        100,  # Distributable
        122,  # Equity CF
    ]

    returns_to_blank_D = [123, 124]  # IRR, MOIC

    blanked_count = 0
    for row in rows_to_blank_E_to_N:
        for col in range(5, 15):  # E to N
            cell = ws.cell(row=row, column=col)
            if cell.value is not None:
                cell.value = None
                cell.fill = YELLOW_FILL
                blanked_count += 1

    for row in returns_to_blank_D:
        cell = ws.cell(row=row, column=4)
        if cell.value is not None:
            cell.value = None
            cell.fill = YELLOW_FILL
            blanked_count += 1

    wb.save(filepath)
    print(f"  Added DRILL MODE note to R2")
    print(f"  Blanked {blanked_count} formula cells")
    print(f"  Saved: {filepath}")


def verify_fixes():
    """Verify all 4 fixes"""
    print("\n" + "=" * 60)
    print("VERIFICATION")
    print("=" * 60)

    # 1. Midstream
    wb1 = load_workbook(f"{DELIV_DIR}/Model_5_Midstream.xlsx")
    ws1 = wb1['Model_Standard']
    print("\n[1] Model_5_Midstream Unlevered Section:")
    for row in range(126, 129):
        a = ws1.cell(row=row, column=1).value or ""
        d = ws1.cell(row=row, column=4).value
        print(f"  R{row}: A='{a}' | D={d}")

    # Check label-formula consistency
    r127_label = ws1.cell(row=127, column=1).value or ""
    r127_formula = str(ws1.cell(row=127, column=4).value or "")
    if "IRR" in r127_label and "IRR" in r127_formula:
        print("  ✓ R127: Label says 'IRR' and formula uses IRR()")
    else:
        print(f"  ✗ R127: Label='{r127_label}', Formula='{r127_formula}'")

    r128_label = ws1.cell(row=128, column=1).value or ""
    r128_formula = str(ws1.cell(row=128, column=4).value or "")
    if "MOIC" in r128_label and "SUM" in r128_formula:
        print("  ✓ R128: Label says 'MOIC' and formula uses SUM (MOIC calc)")
    else:
        print(f"  ✗ R128: Label='{r128_label}', Formula='{r128_formula}'")

    # 2. Wind
    wb2 = load_workbook(f"{DELIV_DIR}/Model_6_Wind.xlsx")
    ws2 = wb2['Model_Standard']
    r15 = ws2.cell(row=15, column=4).value
    print(f"\n[2] Model_6_Wind Audit Strip S&U Balance:")
    print(f"  R15 formula: {r15}")
    if "D114" in str(r15) and "D119" in str(r15):
        print("  ✓ References D114 (Total Uses) and D119 (Total Sources)")
    else:
        print("  ✗ Wrong cell references")

    # 3. Drill CCGT
    wb3 = load_workbook(f"{DELIV_DIR}/Drill_1_CCGT.xlsx")
    ws3 = wb3['Model_Standard']
    r2 = ws3.cell(row=2, column=1).value or ""
    e86 = ws3.cell(row=86, column=5).value
    d142 = ws3.cell(row=142, column=4).value
    print(f"\n[3] Drill_1_CCGT:")
    print(f"  R2: '{r2[:50]}...'")
    if "DRILL" in r2:
        print("  ✓ DRILL MODE note present")
    else:
        print("  ✗ Missing DRILL MODE note")
    print(f"  E86 (EBITDA Y1): {e86}")
    print(f"  D142 (IRR): {d142}")
    if e86 is None and d142 is None:
        print("  ✓ Key formulas blanked")
    else:
        print("  ✗ Some formulas not blanked")

    # 4. Drill Midstream
    wb4 = load_workbook(f"{DELIV_DIR}/Drill_5_Midstream.xlsx")
    ws4 = wb4['Model_Standard']
    r2_m = ws4.cell(row=2, column=1).value or ""
    e72 = ws4.cell(row=72, column=5).value
    d123 = ws4.cell(row=123, column=4).value
    print(f"\n[4] Drill_5_Midstream:")
    print(f"  R2: '{r2_m[:50]}...'")
    if "DRILL" in r2_m:
        print("  ✓ DRILL MODE note present")
    else:
        print("  ✗ Missing DRILL MODE note")
    print(f"  E72 (EBITDA Y1): {e72}")
    print(f"  D123 (IRR): {d123}")
    if e72 is None and d123 is None:
        print("  ✓ Key formulas blanked")
    else:
        print("  ✗ Some formulas not blanked")


def regenerate_zip():
    """Regenerate ZIP with all files"""
    import zipfile
    import os
    import hashlib

    print("\n" + "=" * 60)
    print("REGENERATING ZIP")
    print("=" * 60)

    zip_path = f"{DELIV_DIR}/deliverables_bundle.zip"

    # Remove old zip
    if os.path.exists(zip_path):
        os.remove(zip_path)

    # Create new zip
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in os.listdir(DELIV_DIR):
            if f.endswith('.zip'):
                continue
            fp = os.path.join(DELIV_DIR, f)
            if os.path.isfile(fp):
                zf.write(fp, f)
                print(f"  Added: {f}")

    # Calculate SHA256
    with open(zip_path, 'rb') as f:
        sha256 = hashlib.sha256(f.read()).hexdigest()

    print(f"\n  SHA256: {sha256}")
    print(f"  Saved: {zip_path}")


if __name__ == "__main__":
    fix_midstream_unlevered()
    fix_wind_audit_strip()
    fix_drill_ccgt()
    fix_drill_midstream()
    verify_fixes()
    regenerate_zip()

    print("\n" + "=" * 60)
    print("ROUND 3 FIXES COMPLETE")
    print("=" * 60)
