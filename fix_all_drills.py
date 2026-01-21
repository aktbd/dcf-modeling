#!/usr/bin/env python3
"""
Fix all 6 Drill files by copying formulas from corresponding Model files,
then blanking out key practice cells.
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from copy import copy

DELIV_DIR = '/home/user/dcf-modeling/deliverables/deliverables-20260120-0030'

# Cells to blank in each drill (row ranges for E-N columns)
# These are the cells students should fill in themselves
BLANK_PATTERNS = {
    'CCGT': {
        'rows_to_blank': [76, 77, 78, 80, 82, 83, 84, 86, 91, 92, 97, 98, 99, 101, 102, 103, 104],
        'returns_rows': [147, 148],
    },
    'Peaker': {
        'rows_to_blank': [77, 78, 79, 81, 83, 84, 85, 87, 91, 92, 93, 97, 98, 99, 101, 102, 103, 104],
        'returns_rows': [138, 139],
    },
    'SolarBESS': {
        'rows_to_blank': [72, 73, 74, 76, 78, 81, 83, 87, 88, 89, 90, 91, 92],
        'returns_rows': [128, 129],
    },
    'Transmission': {
        'rows_to_blank': [77, 78, 80, 85, 86, 89, 90, 91, 92, 93, 94, 95],
        'returns_rows': [134, 135],
    },
    'Midstream': {
        'rows_to_blank': [69, 70, 72, 75, 77, 81, 82, 83, 84, 86, 87, 88, 89],
        'returns_rows': [127, 128],
    },
    'Wind': {
        'rows_to_blank': [68, 69, 71, 72, 73, 74, 79, 80, 81, 82, 84, 89, 90, 91, 92, 93, 95, 96],
        'returns_rows': [131, 132],
    },
}

def copy_sheet_formulas(model_ws, drill_ws, max_row=160, max_col=14):
    """Copy all formulas and values from model to drill sheet"""
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            model_cell = model_ws.cell(row=row, column=col)
            drill_cell = drill_ws.cell(row=row, column=col)

            # Copy value
            drill_cell.value = model_cell.value

            # Copy style if exists
            if model_cell.has_style:
                drill_cell.font = copy(model_cell.font)
                drill_cell.fill = copy(model_cell.fill)
                drill_cell.border = copy(model_cell.border)
                drill_cell.alignment = copy(model_cell.alignment)
                drill_cell.number_format = model_cell.number_format

def blank_practice_cells(ws, rows_to_blank, returns_rows):
    """Blank out cells that students should fill in"""
    # Blank formulas in E-N for specified rows
    for row in rows_to_blank:
        for col in range(5, 15):  # E to N
            cell = ws.cell(row=row, column=col)
            cell.value = None

    # Blank returns (column D)
    for row in returns_rows:
        cell = ws.cell(row=row, column=4)
        cell.value = None

def fix_drill(model_name, drill_name, asset_type):
    """Fix a single drill file"""
    model_path = f"{DELIV_DIR}/{model_name}"
    drill_path = f"{DELIV_DIR}/{drill_name}"

    print(f"\nFixing {drill_name}...")

    # Load both files
    model_wb = load_workbook(model_path)
    drill_wb = load_workbook(drill_path)

    # Copy formulas from Model_Standard to drill's Model_Standard
    if 'Model_Standard' in model_wb.sheetnames and 'Model_Standard' in drill_wb.sheetnames:
        model_ws = model_wb['Model_Standard']
        drill_ws = drill_wb['Model_Standard']

        print(f"  Copying formulas from Model to Drill...")
        copy_sheet_formulas(model_ws, drill_ws)

        # Blank practice cells
        if asset_type in BLANK_PATTERNS:
            pattern = BLANK_PATTERNS[asset_type]
            print(f"  Blanking {len(pattern['rows_to_blank'])} practice rows...")
            blank_practice_cells(drill_ws, pattern['rows_to_blank'], pattern['returns_rows'])

        # Add DRILL MODE note at row 2
        drill_ws.cell(row=2, column=1).value = "═══ DRILL MODE - Fill in yellow cells ═══"

    # Save drill
    drill_wb.save(drill_path)
    print(f"  Saved: {drill_path}")

def main():
    print("="*60)
    print("FIXING ALL 6 DRILL FILES")
    print("="*60)

    drills = [
        ('Model_1_CCGT.xlsx', 'Drill_1_CCGT.xlsx', 'CCGT'),
        ('Model_2_Peaker.xlsx', 'Drill_2_Peaker.xlsx', 'Peaker'),
        ('Model_3_SolarBESS.xlsx', 'Drill_3_SolarBESS.xlsx', 'SolarBESS'),
        ('Model_4_Transmission.xlsx', 'Drill_4_Transmission.xlsx', 'Transmission'),
        ('Model_5_Midstream.xlsx', 'Drill_5_Midstream.xlsx', 'Midstream'),
        ('Model_6_Wind.xlsx', 'Drill_6_Wind.xlsx', 'Wind'),
    ]

    for model_name, drill_name, asset_type in drills:
        fix_drill(model_name, drill_name, asset_type)

    print("\n" + "="*60)
    print("ALL DRILL FILES FIXED")
    print("="*60)

if __name__ == "__main__":
    main()
