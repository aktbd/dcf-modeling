#!/usr/bin/env python3
"""
Validate Excel models - check for formula errors and sanity check outputs
"""

import openpyxl
from openpyxl.utils import get_column_letter
import re

def check_model(filepath):
    """Check a model for formula errors and output sanity"""
    print(f"\n{'='*60}")
    print(f"Validating: {filepath}")
    print(f"{'='*60}")

    wb = openpyxl.load_workbook(filepath, data_only=False)
    ws = wb.active

    errors = []
    warnings = []

    # Check all cells for potential formula issues
    formula_count = 0
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            if cell.value and str(cell.value).startswith('='):
                formula_count += 1
                formula = str(cell.value)

                # Check for common formula issues
                if '#REF' in formula:
                    errors.append(f"#REF error in {get_column_letter(col)}{row}: {formula}")
                if '#NAME' in formula:
                    errors.append(f"#NAME error in {get_column_letter(col)}{row}: {formula}")
                if '#VALUE' in formula:
                    errors.append(f"#VALUE error in {get_column_letter(col)}{row}: {formula}")

                # Check for suspicious patterns
                if re.search(r'\$[A-Z]\$0', formula):
                    errors.append(f"Reference to row 0 in {get_column_letter(col)}{row}: {formula}")
                if re.search(r'[A-Z]0[^0-9]', formula) or formula.endswith('0'):
                    if not 'E0' in formula:  # E0 is scientific notation, not column
                        pass  # This could be a valid E0 reference

    print(f"Total formulas found: {formula_count}")

    # Load with calculated values to check outputs
    wb_calc = openpyxl.load_workbook(filepath, data_only=True)
    ws_calc = wb_calc.active

    # Find key output cells by scanning for labels
    outputs = {}
    for row in range(1, ws_calc.max_row + 1):
        label = ws_calc.cell(row=row, column=1).value
        if label:
            label_str = str(label).strip()
            # Check column D for values
            val = ws_calc.cell(row=row, column=4).value
            if val is not None:
                outputs[label_str] = {'row': row, 'value': val}

    # Report key metrics if found
    print("\nKey Outputs (from column D):")
    key_labels = ['Entry EV', 'Y1 EBITDA', 'Min DSCR', 'Levered IRR', 'MOIC',
                  'Unlevered IRR', 'Debt Amount', 'UCAP', 'Annual Generation',
                  'S&U Balance', 'Balance Check', 'Debt Payoff Y7']

    for label in key_labels:
        for out_label, data in outputs.items():
            if label.lower() in out_label.lower():
                val = data['value']
                if isinstance(val, (int, float)):
                    if 'IRR' in out_label or '%' in out_label:
                        print(f"  {out_label}: {val:.1%}")
                    elif 'DSCR' in out_label or 'MOIC' in out_label or 'Multiple' in out_label:
                        print(f"  {out_label}: {val:.2f}x")
                    else:
                        print(f"  {out_label}: {val:,.1f}")
                else:
                    print(f"  {out_label}: {val}")
                break

    # Check for specific sanity values
    print("\nSanity Checks:")

    # Check if we can find EBITDA and fuel cost values in Year 1 (column E)
    for row in range(1, ws_calc.max_row + 1):
        label = ws_calc.cell(row=row, column=1).value
        if label:
            label_str = str(label).strip()

            # Check Year 1 values (column E)
            y1_val = ws_calc.cell(row=row, column=5).value

            if 'EBITDA' in label_str and y1_val is not None and isinstance(y1_val, (int, float)):
                if 30 <= y1_val <= 150:
                    print(f"  Y1 EBITDA: ${y1_val:.1f}mm - OK")
                else:
                    warnings.append(f"Y1 EBITDA = ${y1_val:.1f}mm seems unusual")

            if 'Fuel Cost' in label_str and y1_val is not None and isinstance(y1_val, (int, float)):
                if 20 <= y1_val <= 80:
                    print(f"  Y1 Fuel Cost: ${y1_val:.1f}mm - OK")
                else:
                    warnings.append(f"Y1 Fuel Cost = ${y1_val:.1f}mm seems unusual")

            if 'DSCR' in label_str and 'Min' not in label_str and y1_val is not None and isinstance(y1_val, (int, float)):
                if 1.0 <= y1_val <= 3.5:
                    print(f"  Y1 DSCR: {y1_val:.2f}x - OK")
                elif y1_val < 1.0:
                    errors.append(f"Y1 DSCR = {y1_val:.2f}x is below 1.0x!")

            if label_str == 'Levered IRR':
                irr_val = ws_calc.cell(row=row, column=4).value
                if irr_val is not None and isinstance(irr_val, (int, float)):
                    if 0.10 <= irr_val <= 0.35:
                        print(f"  Levered IRR: {irr_val:.1%} - OK")
                    else:
                        warnings.append(f"Levered IRR = {irr_val:.1%} seems unusual")

    # Check for error values in calculated columns
    print("\nChecking for #ERROR values in output cells...")
    error_cells = []
    for row in range(1, min(ws_calc.max_row + 1, 150)):
        for col in range(4, min(ws_calc.max_column + 1, 15)):
            val = ws_calc.cell(row=row, column=col).value
            if val is not None and isinstance(val, str) and val.startswith('#'):
                error_cells.append(f"{get_column_letter(col)}{row}: {val}")

    if error_cells:
        print(f"  Found {len(error_cells)} error values:")
        for ec in error_cells[:10]:
            print(f"    {ec}")
        errors.extend(error_cells)
    else:
        print("  No #ERROR values found - PASS")

    # Summary
    print("\n" + "="*60)
    if errors:
        print(f"ERRORS FOUND: {len(errors)}")
        for e in errors[:10]:
            print(f"  - {e}")
    else:
        print("NO ERRORS FOUND")

    if warnings:
        print(f"\nWARNINGS: {len(warnings)}")
        for w in warnings:
            print(f"  - {w}")

    return len(errors) == 0

if __name__ == "__main__":
    import sys

    # Default to CCGT model
    filepath = sys.argv[1] if len(sys.argv) > 1 else "/mnt/user-data/outputs/Model_1_CCGT.xlsx"

    success = check_model(filepath)
    sys.exit(0 if success else 1)
