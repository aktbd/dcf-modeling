# IMPLEMENTATION MANIFESTO
## Infrastructure PE Interview Prep Package - Technical Documentation

---

# OVERVIEW

This document provides technical documentation for replicating the Infrastructure PE Interview Prep Package built for Lotus Infrastructure Partners Associate-level interviews.

## Package Contents

| File | Description |
|------|-------------|
| Model_1_CCGT.xlsx | Combined cycle gas turbine model |
| Model_2_Peaker.xlsx | Simple cycle peaker model |
| Model_3_SolarBESS.xlsx | Solar + battery storage model |
| Model_4_Transmission.xlsx | Regulated transmission model |
| Model_5_Midstream.xlsx | Gas gathering system model |
| Drill_1-5_*.xlsx | Same models with Intuition Guide sheets |
| IC_Memo_1-5_*.docx | Investment committee memos |
| Master_Template.xlsx | Functional template for any asset type |
| Lotus_Interview_CheatSheet_Expanded.md | Comprehensive study guide |
| Lotus_Interview_CheatSheet_Expanded.pdf | PDF version of cheat sheet |

---

# TECHNICAL ARCHITECTURE

## Row Tracking Methodology

The critical innovation in this build is **dictionary-based row tracking**. Instead of hardcoding row numbers in formulas (which leads to broken references), we track every row's position as we build:

```python
rows = {}

# When placing an input, track its row
row = 17
rows['entry_ev'] = row
ws.cell(row, 1, "Entry EV")
apply_input_style(ws, row, 4, 520)
row += 1

rows['txn_fees'] = row
# ... etc

# When writing formulas, reference tracked positions
formula = f"=$D${rows['entry_ev']}*$D${rows['leverage']}"
```

This approach ensures:
1. All cell references are correct by construction
2. Inserting rows doesn't break formulas
3. Easy to audit and modify
4. Self-documenting code

## Column Convention

| Column | Purpose | Width |
|--------|---------|-------|
| A | Row labels | 32 |
| B | Units | 12 |
| C | Notes/explanations | 38 |
| D | Year 0 / Inputs | 14 |
| E-N | Years 1-10 | 14 each |

## Formatting Standards

### Cell Styles

| Style | Fill Color | Font | Use Case |
|-------|------------|------|----------|
| Input | #FFFFC7 (yellow) | Blue, Bold | User-modifiable inputs |
| Calculation | None (white) | Black | Intermediate calculations |
| Output | #C6EFCE (green) | Black, Bold | Key results |
| Warning | #FCE4D6 (orange) | Black | Cash sweeps, one-time items |
| Header | #1F4E79 (navy) | White, Bold | Section headers |

### Number Formats

```python
CURRENCY_FORMAT = '#,##0.0_);(#,##0.0)'      # $mm values
PERCENT_FORMAT = '0.0%'                       # Percentages
MULTIPLE_FORMAT = '0.00"x"'                   # DSCR, multiples
INTEGER_FORMAT = '#,##0'                      # Whole numbers
```

---

# MODEL-SPECIFIC NOTES

## Model 1: CCGT

### Key Characteristics
- Baseload plant with capacity factor (~55%)
- Dual revenue streams (energy + capacity)
- 75% cash sweep structure
- 7-year hold, 7-year debt tenor

### Critical Formulas

**Fuel Cost (MUST verify dimensional analysis):**
```
Fuel Cost ($mm) = Generation (MWh) × Heat Rate (Btu/kWh) × Gas Price ($/mmBtu) / 1,000,000,000
```

**UCAP:**
```
UCAP = Capacity × (1 - Forced Outage Rate)
```

**Capacity Revenue (note ×1000 for kW→MW):**
```
Capacity Revenue ($mm) = UCAP (MW) × Capacity Price ($/kW-yr) × 1,000 / 1,000,000
```

---

## Model 2: Peaker

### Key Differences from CCGT
1. **Generation formula uses dispatch hours, NOT capacity factor:**
   ```
   Generation = Capacity × Dispatch Hours × Availability
   NOT: Capacity × CF × 8,760
   ```

2. Lower leverage (30%) due to volatility
3. Lower sweep (50%) to retain maintenance reserves
4. Major maintenance in Y5 ($4mm)

### CFADS Adjustment
```
CFADS = EBITDA - Taxes - Major Maintenance (if applicable year)
Major Maint = IF(year = maint_year, maint_cost, 0)
```

---

## Model 3: Solar + BESS

### Key Differences
1. **No fuel cost** - sun is free
2. **ITC (30%)** reduces equity in Sources & Uses
3. **Degradation (0.5%/yr)** on solar generation
4. **Sculpted debt** (not sweep) - sized to target DSCR
5. **BESS augmentation** in Y7 ($5mm)
6. **MACRS shield** - zero taxes Y1-Y6

### Sculpted Debt Mechanics
```
Target Debt Service = CFADS / Target DSCR
Sculpted Principal = MIN(MAX(0, Target DS - Interest), Beginning Balance)
```

### ITC in Sources & Uses
```
Uses: Purchase + Fees + DSRA
Sources: Debt + Pre-ITC Equity - ITC Benefit = Net Equity + ITC + Debt
```

---

## Model 4: Transmission

### Key Differences
1. **Construction period (Y0-Y2)** with no revenue
2. **Construction interest capitalization** into RAB
3. **Phased equity deployment** over Y0-Y3
4. **Sculpted debt starting Y3** (COD)
5. **Exit at RAB multiple** (not EBITDA multiple)

### Construction Financing Flow
```
Y0: Dev Equity = Dev Cost Y0
Y1: Dev Equity = Dev Cost Y1
Y2: Const Equity = Const Cost × (1 - Const Debt %)
Y3: Refi Equity = (Const Loan + Capitalized Int) - Term Debt
```

### RAB Calculation
```
RAB at COD = Dev Y0 + Dev Y1 + Construction + Capitalized Interest
Exit RAB = RAB at COD - Cumulative Depreciation
Exit EV = Exit RAB × xRAB Multiple
```

---

## Model 5: Midstream

### Key Differences
1. **Volume profile: Growth (Y1-5) then Decline (Y6-10)**
2. **Fee-for-service revenue** (Volume × Fee)
3. **Short tenor (5 years)** - must pay off before decline
4. **Aggressive sweep (75%)** during growth phase
5. **Growth capex Y1-5 only**
6. **Compressed exit multiple (6.5x)**

### Volume Profile
```
Y1-Y5: Volume = Y1 Volume × (1 + Growth Rate)^(year-1)
Y6-Y10: Volume = Y5 Volume × (1 - Decline Rate)^(year-5)
```

---

# VALIDATION PROCEDURES

## Pre-Delivery Checklist

### 1. Error Check
- Open each file in Excel
- Ctrl+` to show formulas
- Look for #REF, #NAME, #VALUE, #DIV/0
- Every formula cell must return valid result

### 2. Reference Audit
For each model, click through these cells and verify precedents:
- [ ] UCAP / Generation
- [ ] Fuel Cost (if applicable)
- [ ] EBITDA
- [ ] CFADS
- [ ] Beginning Debt Balance
- [ ] Scheduled Principal
- [ ] Cash Sweep (if applicable)
- [ ] DSCR
- [ ] IRR

### 3. Sanity Check Outputs

| Metric | Acceptable Range | Red Flag |
|--------|------------------|----------|
| CCGT Fuel Cost | $30-50mm | >$100mm or <$10mm |
| DSCR | 1.2-2.5x | <1.0x or >5.0x |
| Levered IRR | 12-25% | <5% or >40% |
| MOIC | 1.5-3.0x | <1.0x or >5.0x |
| Debt at exit | <$5mm | >50% of original |

### 4. QA Check Results
All models should show:
- S&U Balance: PASS
- Debt Payoff: PASS
- Min DSCR ≥ threshold: PASS

### 5. Scenario Test
For each model:
1. Change one input (e.g., Capacity +10%)
2. Verify downstream cells update proportionally
3. Confirm no circular reference errors

---

# PYTHON DEPENDENCIES

```
openpyxl>=3.1.0      # Excel file creation
python-docx>=1.0.0   # Word document creation
reportlab>=4.0.0     # PDF generation (optional)
```

Install with:
```bash
pip install openpyxl python-docx reportlab
```

---

# BUILD SCRIPTS

| Script | Purpose |
|--------|---------|
| build_ccgt_model.py | Builds CCGT model |
| build_peaker_model.py | Builds Peaker model |
| build_solar_bess_model.py | Builds Solar+BESS model |
| build_transmission_model.py | Builds Transmission model |
| build_midstream_model.py | Builds Midstream model |
| create_drill_versions.py | Adds Intuition Guide sheets |
| create_ic_memos.py | Creates Word memos |
| create_master_template.py | Creates generic template |
| validate_model.py | Validates formulas |
| verify_formulas.py | Deep formula verification |

---

# COMMON ISSUES AND FIXES

### Issue: #REF errors after inserting rows
**Fix:** Use row tracking dictionary; never hardcode row numbers

### Issue: Fuel cost is way off
**Fix:** Check divisor - must be 1,000,000,000 (1e9), not 1,000,000 (1e6)

### Issue: Peaker generation too high
**Fix:** Use Dispatch Hours × Availability, not CF × 8,760

### Issue: DSRA target incorrect
**Fix:** Target references NEXT year's debt service, not current year

### Issue: IRR returns #NUM!
**Fix:** Ensure equity CF range includes Y0 negative outflow and at least one positive inflow

### Issue: S&U doesn't balance
**Fix:** Check DSRA is included as both Use (initial funding) and component of equity at exit

---

# EXTENSION OPPORTUNITIES

### Additional Asset Types
- Offshore wind (construction risk, higher capex)
- Battery standalone (merchant, ancillary focus)
- LNG terminal (commodity exposure, volume risk)
- Data center (demand risk, high utilization)

### Additional Features
- Sensitivity tables (data tables in Excel)
- Scenario toggles (dropdown selection)
- Debt sizing (goal seek to target coverage)
- Construction draw schedule
- Working capital modeling

---

# CONTACT & SUPPORT

For questions about this implementation:
1. Review this manifesto first
2. Check the Intuition Guide sheets in drill versions
3. Reference the Cheat Sheet for conceptual questions

---

*Build Version: 1.0*
*Last Updated: January 2026*
