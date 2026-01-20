# CLAUDE.md - Infrastructure PE Modeling Standards

This document is the authoritative reference for all infrastructure private equity financial modeling work in this repository. All models, templates, and build scripts must adhere to these conventions.

---

## 1. WORKBOOK STRUCTURE

### Sheet Organization
```
Sheet 1: "Model" (or asset-specific name)
Sheet 2: "Intuition Guide" (for drill versions only)
Sheet 3: "Instructions" (for templates only)
```

### Column Layout
| Column | Purpose | Width |
|--------|---------|-------|
| A | Row labels | 32 |
| B | Units | 12 |
| C | Notes/explanations | 38 |
| D | Year 0 / Inputs | 14 |
| E | Year 1 | 14 |
| F | Year 2 | 14 |
| ... | ... | ... |
| N | Year 10 | 14 |

### Row Structure (Approximate Ranges)
```
Rows 1:        Year headers (0, 1, 2, ... 10)
Rows 2-13:     AUDIT STRIP (key metrics + QA checks)
Rows 14-15:    Blank separators
Rows 16-50:    INPUTS (grouped by category)
Rows 51-60:    CALCULATED ITEMS
Rows 61-80:    OPERATING MODEL
Rows 81-90:    CASH FLOW
Rows 91-105:   DEBT SCHEDULE
Rows 106-115:  DSRA
Rows 116-125:  DISTRIBUTION WATERFALL
Rows 126-140:  SOURCES & USES
Rows 141-160:  EXIT & EQUITY RETURNS
```

### Freeze Panes
Always freeze at cell `E2` so row labels (column A-D) and year headers (row 1) remain visible during scrolling.

---

## 2. ROW-INDEX DICTIONARY APPROACH

**CRITICAL: Never hardcode row numbers in formulas.**

### Implementation Pattern
```python
rows = {}  # Initialize tracking dictionary

row = 17
rows['entry_ev'] = row
set_label(ws, row, "Entry EV", "$mm", "Purchase price")
apply_input_style(ws, row, 4, 520)
row += 1

rows['txn_fees'] = row
set_label(ws, row, "Transaction Fees", "%", "Legal, advisory")
apply_input_style(ws, row, 4, 0.015)
row += 1

# When writing formulas, ALWAYS reference tracked positions:
formula = f"=$D${rows['entry_ev']}*$D${rows['leverage']}"
```

### Benefits
1. All cell references are correct by construction
2. Inserting/deleting rows doesn't break formulas
3. Easy to audit and modify
4. Self-documenting code

### Required Dictionary Keys (Minimum)
Every model must track these rows at minimum:
- `entry_ev`, `txn_fees`, `exit_multiple`, `exit_year`
- `capacity`, `utilization` (or equivalent)
- `leverage`, `interest_rate`, `debt_tenor`, `dsra_months`, `sweep_pct` (or `target_dscr`), `lockup_dscr`
- `tax_rate`, `depreciable_basis`, `depreciation_life`
- `debt_amount`, `ebitda`, `cfads`
- `debt_beg_bal`, `debt_end_bal`, `interest`, `dscr`
- `dsra_target`, `dsra_beg`, `dsra_funding`, `dsra_end`
- `distributable`, `equity_cf`, `levered_irr`, `moic`

---

## 3. FORMATTING STANDARDS

### Cell Styles

| Style Name | Fill Color | Font Color | Font Weight | Use Case |
|------------|------------|------------|-------------|----------|
| Input | #FFFFC7 (yellow) | #0000FF (blue) | Bold | User-modifiable inputs |
| Calculation | None (white) | #000000 (black) | Normal | Intermediate calculations |
| Output | #C6EFCE (green) | #000000 (black) | Bold | Key results (EBITDA, CFADS, IRR) |
| Warning | #FCE4D6 (orange) | #000000 (black) | Normal | Cash sweeps, one-time items, maintenance |
| Header | #1F4E79 (navy) | #FFFFFF (white) | Bold | Section headers |

### Number Formats
```python
CURRENCY_FORMAT = '#,##0.0_);(#,##0.0)'      # For $mm values
PERCENT_FORMAT = '0.0%'                       # For percentages
MULTIPLE_FORMAT = '0.00"x"'                   # For DSCR, MOIC, multiples
INTEGER_FORMAT = '#,##0'                      # For whole numbers (MW, years)
DECIMAL_FORMAT = '#,##0.00'                   # For prices ($/MWh, $/mmBtu)
```

### Style Helper Functions
Every build script must implement these functions:
```python
def apply_input_style(ws, row, col, value=None)
def apply_calc_style(ws, row, col, formula=None)
def apply_output_style(ws, row, col, formula=None)
def apply_warning_style(ws, row, col, formula=None)
def apply_header_style(ws, row, col, text)
def set_label(ws, row, label, units="", notes="")
```

---

## 4. DEBT MECHANICS RULES

### Sweep vs Sculpt Decision
```
IF CFADS is predictable and contracted:
    → Use SCULPTED debt
    → Size debt service to maintain target DSCR
    → Principal = Target DS - Interest
    → Use for: PPAs, regulated assets, long-term contracts

IF CFADS is volatile or merchant:
    → Use CASH SWEEP debt
    → Mandatory prepayment from excess cash flow
    → Sweep = MIN(Excess × Sweep%, Remaining Balance)
    → Use for: Merchant power, midstream, commodity-exposed assets
```

### DSRA Logic

**Target Calculation:**
```
Y0 Target = Y1 Debt Service × (DSRA Months / 12)
Yn Target = Y(n+1) Debt Service × (DSRA Months / 12)
Final Year Target = 0 (released at exit)
```

**CRITICAL: DSRA target references NEXT year's debt service, not current year.**

**Funding/Release:**
```
DSRA Funding = Target - Beginning
  Positive = funding requirement (reduces distributable)
  Negative = release (increases distributable)

DSRA Ending = Beginning + Funding
```

### Debt Service Calculations
```
Scheduled Principal = Debt Amount / Tenor
Interest = Beginning Balance × Interest Rate
Debt Service (Pre-Sweep) = Interest + Scheduled Principal

For Sweep:
  Excess Cash = CFADS - Debt Service - DSRA Funding
  Cash Sweep = MIN(Excess × Sweep%, Beginning - Scheduled Principal)
  Total Principal = Scheduled + Sweep

For Sculpt:
  Target DS = CFADS / Target DSCR
  Sculpted Principal = MIN(MAX(0, Target DS - Interest), Beginning Balance)

Ending Balance = MAX(0, Beginning - Total Principal)
DSCR = IF(Debt Service > 0, CFADS / Debt Service, 0)
```

### Distribution Waterfall
```
1. Start with CFADS
2. Less: Interest
3. Less: Total Principal (scheduled + sweep or sculpted)
4. Less: DSRA Funding (if positive)
5. Plus: DSRA Release (if negative)
6. Equals: Cash Available for Distribution
7. Apply Lock-up Test: IF(DSCR >= Lock-up DSCR, "PASS", "LOCKED")
8. Distributable = IF(PASS, MAX(0, Cash Available), 0)
```

---

## 5. QA AND AUDIT STRIP REQUIREMENTS

### Audit Strip Contents (Rows 2-9)
Every model must display these metrics at the top:
```
Row 2: Section header "═══ AUDIT STRIP ═══"
Row 3: Entry EV ($mm)
Row 4: Y1 EBITDA ($mm)
Row 5: Avg CFADS over hold period ($mm)
Row 6: Min DSCR over debt tenor (x)
Row 7: Levered IRR (%)
Row 8: MOIC (x)
Row 9: Unlevered IRR (%)
```

### QA Checks (Rows 10-13)
Every model must include these automated checks:
```
Row 10: Section header "═══ QA CHECKS ═══"
Row 11: S&U Balance = IF(ABS(Uses - Sources) < 0.01, "PASS", "FAIL")
Row 12: Debt Payoff = IF(Ending Balance at Exit < 1, "PASS", "FAIL")
Row 13: Min DSCR Check = IF(Min DSCR >= threshold, "PASS", "FAIL")
```

### Threshold Requirements
| Check | Minimum Threshold |
|-------|-------------------|
| DSCR (merchant/sweep) | ≥ 1.25x |
| DSCR (contracted/sculpt) | ≥ 1.30x |
| DSCR (transmission/regulated) | ≥ 1.35x |

---

## 6. FORMULA CONVENTIONS

### Cell Reference Patterns
```
$D$[row] - Fixed reference to input in column D (most inputs)
[col][row] - Relative reference for time-series calculations
$D$[row]*[col][price_row] - Mixed: fixed input × varying price
```

### Escalation Pattern
All escalating values use this formula:
```
Yn Value = Y1 Value × (1 + Escalation Rate)^(n-1)

Excel: =$D$[y1_row]*(1+$D$[esc_row])^([year]-1)
```
Where `[year]` is the column's year number (1, 2, 3...).

### Conditional Formulas
Always use explicit IF statements for:
- Division (prevent #DIV/0): `=IF(denominator>0, numerator/denominator, 0)`
- One-time items: `=IF(year=target_year, amount, 0)`
- Lock-up tests: `=IF(DSCR>=threshold, "PASS", "LOCKED")`

---

## 7. ASSET-SPECIFIC RULES

### Generation Formulas

**CCGT/Baseload:**
```
Generation = Capacity × Capacity Factor × 8,760
```

**Peaker (CRITICAL - most common error):**
```
Generation = Capacity × Dispatch Hours × Availability
DO NOT USE: Capacity × CF × 8,760
```

**Solar (with degradation):**
```
Yn Generation = Y1 Generation × (1 - Degradation Rate)^(n-1)
```

### Revenue Formulas

**Energy Revenue:**
```
Energy ($mm) = Generation (MWh) × Price ($/MWh) / 1,000,000
```

**Capacity Revenue (note ×1000 for kW conversion):**
```
Capacity ($mm) = UCAP (MW) × Price ($/kW-yr) × 1,000 / 1,000,000
```

**UCAP:**
```
UCAP = Capacity × (1 - Forced Outage Rate)
```

### Fuel Cost (CRITICAL)
```
Fuel ($mm) = Generation (MWh) × Heat Rate (Btu/kWh) × Gas Price ($/mmBtu) / 1,000,000,000

THE DIVISOR IS 1 BILLION (1e9), NOT 1 MILLION (1e6)
```

**Verification formula:**
```
Variable Cost ($/MWh) ≈ Heat Rate × Gas Price / 1,000
Example: 7,000 × $3.00 / 1,000 = $21/MWh
```

### Midstream Volume Profile
```
Growth Phase (Y1-Y5):
  Volume = Y1 Volume × (1 + Growth Rate)^(year-1)

Decline Phase (Y6+):
  Volume = Y5 Volume × (1 - Decline Rate)^(year-5)
```

---

## 8. DO NOT RULES (INTERVIEW FAILURES)

### Absolute Prohibitions
1. **DO NOT** hardcode row numbers in formulas
2. **DO NOT** use `Capacity × CF × 8,760` for peakers
3. **DO NOT** divide fuel cost by 1,000,000 (must be 1,000,000,000)
4. **DO NOT** forget the ×1,000 conversion for capacity revenue
5. **DO NOT** reference current year DS for DSRA target (use next year)
6. **DO NOT** include Exit EV directly as exit equity (must subtract debt, add DSRA)
7. **DO NOT** forget DSRA in Sources & Uses
8. **DO NOT** leave #REF, #NAME, #VALUE, or #DIV/0 errors
9. **DO NOT** create circular references
10. **DO NOT** use IRR range that excludes Y0 negative equity outflow

### Common Mistakes to Avoid
| Mistake | Consequence | Correct Approach |
|---------|-------------|------------------|
| Fuel cost ÷ 1e6 | Fuel cost 1000x too high | Divide by 1e9 |
| Peaker CF × 8760 | Generation 10x too high | Use Dispatch Hours |
| Capacity $/MW | Revenue 1000x too low | Price is $/kW, multiply by 1000 |
| DSRA = current DS | DSRA off by one year | Target = NEXT year DS |
| Exit equity = Exit EV | Equity overstated | Exit EV - Debt + DSRA |
| IRR excludes Y0 | IRR returns #NUM! | Include Y0 negative equity |

---

## 9. OUTPUT VALIDATION

### Sanity Check Ranges
| Metric | Acceptable Range | Red Flag |
|--------|------------------|----------|
| Fuel Cost (CCGT) | $30-50mm for 365 MW | >$100mm or <$10mm |
| DSCR | 1.2x - 2.5x | <1.0x or >5.0x |
| Levered IRR | 12% - 25% | <5% or >40% |
| MOIC | 1.5x - 3.0x | <1.0x or >5.0x |
| Debt at tenor end | <$5mm | >50% of original |

### Pre-Delivery Checklist
Before any model is delivered:
- [ ] All QA checks show "PASS"
- [ ] No Excel errors (#REF, #NAME, #VALUE, #DIV/0)
- [ ] Click through 5+ formula cells to verify references
- [ ] Run scenario test (change one input, verify cascade)
- [ ] DSCR > 1.0x in all years
- [ ] Debt fully paid by tenor end
- [ ] S&U balances to zero

---

## 10. FILE NAMING CONVENTIONS

### Models
```
Model_[N]_[AssetType].xlsx
Example: Model_1_CCGT.xlsx
```

### Drill Versions
```
Drill_[N]_[AssetType].xlsx
Example: Drill_1_CCGT.xlsx
```

### IC Memos
```
IC_Memo_[N]_[AssetType].docx
Example: IC_Memo_1_CCGT.docx
```

### Reference Materials
```
Master_Template.xlsx
Lotus_Interview_CheatSheet_Expanded.md
Lotus_Interview_CheatSheet_Expanded.pdf
Implementation_Manifesto.md
```

---

## 11. BUILD SCRIPT REQUIREMENTS

Every Python build script must:

1. Import openpyxl and define all style constants
2. Implement all helper functions (apply_input_style, etc.)
3. Initialize `rows = {}` dictionary at start
4. Track every row position as it's created
5. Use f-strings with dictionary references for all formulas
6. Apply appropriate number formats to all cells
7. Set column widths per convention
8. Freeze panes at E2
9. Save to `/mnt/user-data/outputs/` directory
10. Print confirmation message on completion

---

## 12. DEFAULT DELIVERY CONTRACT

**Staging:** `/mnt/user-data/outputs/` (sandbox only, not durable)

**Durable delivery:** GitHub repo at `/deliverables/<tag>/`

**On every run that produces deliverables:**

1. **Build** to `/mnt/user-data/outputs/`
2. **Copy** final files to `/deliverables/deliverables-YYYYMMDD-HHMM/`
3. **Create** `deliverables_bundle.zip` in that folder
4. **Update** these files:
   - `STATE.md` — current commit, timestamp, file status
   - `OUTPUTS_INDEX.md` — latest tag, download links
   - `DELIVERABLES_CHANGELOG.md` — append entry with tag, changes, sha256
5. **Commit** all changes
6. **Push** to both `origin` and `github` remotes
7. **Output** clickable GitHub URLs (not /mnt paths)

**GitHub remote setup:**
```bash
git remote add github https://${GITHUB_TOKEN}@github.com/aktbd/dcf-modeling.git
```

**Key files:**
- `STATE.md` — Bridge file for Claude.ai to read current state
- `OUTPUTS_INDEX.md` — Landing page pointing to latest deliverables
- `DELIVERABLES_CHANGELOG.md` — Append-only version history

**Precedence:** `Explicit user instruction > Current prompt > CLAUDE.md`

*These are defaults, not laws. Deviate when sensible; resume defaults afterward.*

---

## 13. GOOGLE DRIVE DELIVERY (NOT CONFIGURED)

**Status:** Not available in current environment.

**Blockers identified:**
- MCP server for Google Drive: Not installed
- gcloud CLI: Not available or not authenticated
- rclone: Not configured

**To enable (user action required):**

Option A - MCP Server (Preferred):
```bash
npm install -g @anthropic/mcp-server-gdrive
# Then configure OAuth credentials
```

Option B - gcloud CLI:
```bash
gcloud auth login
pip install google-api-python-client google-auth-oauthlib
```

Option C - rclone:
```bash
rclone config  # Interactive setup for Google Drive
# Then: rclone copy file.xlsx gdrive:folder/
```

**Current workaround:** GitHub raw URLs for file delivery.

---

## 14. LESSONS LEARNED (Updated 2026-01-20)

### Modeling Errors Encountered & Fixes

#### 1. DSRA Circular Reference (CRITICAL)

**Error:** DSRA Target in cash sweep models created circular dependency.

```
Circular chain:
DSRA Target Y1 → debt_service_pre Y2 → beg_bal Y2 → end_bal Y1
→ total_principal Y1 → cash_sweep Y1 → excess_cash Y1
→ DSRA_funding Y1 → DSRA_target Y1 (CIRCULAR!)
```

**Root cause:** DSRA Target referenced `debt_service_pre` which depends on beginning balance, which depends on prior year's ending balance, which depends on cash sweep, which depends on DSRA funding.

**Fix:** Added `scheduled_ds` row that calculates debt service using straight-line amortization only (independent of actual sweep):

```python
# ❌ WRONG: Creates circular reference
formula = f"={next_col}{rows['debt_service_pre']}*$D${rows['dsra_months']}/12"

# ✅ RIGHT: Uses scheduled_ds (deterministic, no circularity)
row = 93; rows['scheduled_ds'] = row
set_label(ws, row, "Scheduled DS (for DSRA)", "$mm", "Based on straight-line amort")
for col in range(5, 15):
    yr = col - 4
    formula = f"=MAX(0,$D${rows['debt_amount']}-({yr}-1)*$D${rows['scheduled_principal']})*$D${rows['interest_rate']}+$D${rows['scheduled_principal']}"
    apply_calc_style(ws, row, col, formula)

# Then DSRA Target references scheduled_ds
formula = f"={next_col}{rows['scheduled_ds']}*$D${rows['dsra_months']}/12"
```

**Applies to:** Cash sweep models (CCGT, Peaker, Midstream)
**Does NOT apply to:** Sculpted debt models (Solar+BESS, Transmission) - their debt service is determined by CFADS/target_DSCR, independent of DSRA.

#### 2. Validation Gap

**Error:** Initial validation only checked for error *strings* in formulas, not computed values.

```python
# ❌ WRONG: Only catches literal error text
if '#REF' in str(cell.value):
    errors.append(cell)

# ✅ RIGHT: Use formula evaluation library
import formulas
xl_model = formulas.ExcelModel().loads(filepath).finish()
results = xl_model.calculate()  # Actually computes values
```

**Fix:** Installed `formulas` library and added compute-based validation.

#### 3. Download Links

**Error:** Initially provided sandbox filesystem paths instead of GitHub URLs.

```
❌ WRONG: /mnt/user-data/outputs/Model_1_CCGT.xlsx (404 for user)
✅ RIGHT: https://github.com/aktbd/dcf-modeling/raw/branch/deliverables/file.xlsx
```

**Fix:** Always provide GitHub raw URLs for file downloads.

### Validation Protocol (Pre-Delivery Checklist)

1. **Build all models** with Python scripts
2. **Formula evaluation test:**
   ```python
   import formulas
   xl_model = formulas.ExcelModel().loads(filepath).finish()
   results = xl_model.calculate()  # Must complete without error
   ```
3. **Structural check:**
   - DSRA Target references `scheduled_ds` (not `debt_service_pre`) for cash sweep models
   - No #REF, #NAME, #VALUE, #DIV/0 errors
4. **Sanity check ranges:**
   - DSCR: 1.2x - 2.5x
   - Levered IRR: 12% - 25%
   - Fuel cost (CCGT): $30-50mm
5. **Verify download links** work (GitHub raw URLs)

### Delivery Protocol (Working)

1. Build to `/mnt/user-data/outputs/`
2. Copy to `/deliverables/deliverables-YYYYMMDD-HHMM/`
3. Create ZIP with `sha256sum`
4. Update STATE.md, DELIVERABLES_CHANGELOG.md
5. Commit and push to `origin` (and `github` if configured)
6. Provide GitHub raw URLs (not filesystem paths)

---

## 15. SKILLS ASSESSMENT (2026-01-20)

### Environment Audit

| Resource | Status | Notes |
|----------|--------|-------|
| /mnt/skills/ | Not present | No skill files loaded |
| ~/.claude/instructions.md | Not present | No global instructions |
| CLAUDE.md (this file) | Present | Authoritative for this repo |

### Assessment

**No conflicting skills found.** This repository uses only the local CLAUDE.md for modeling standards.

**Recommendation:** Keep CLAUDE.md as the single source of truth. No external skills to manage or reconcile.

---

## 16. VERSION HISTORY

| Version | Date | Changes |
|---------|------|---------|
| 1.0 | 2026-01-19 | Initial creation - modeling standards |
| 1.1 | 2026-01-20 | Added delivery protocol (Section 12) |
| 1.2 | 2026-01-20 | Simplified delivery contract, added link verification |
| 1.3 | 2026-01-20 | Durable delivery via repo; OUTPUTS_INDEX.md as landing page |
| 1.4 | 2026-01-20 | Added GitHub bridge protocol; STATE.md as sync file |
| 1.5 | 2026-01-20 | Added Lessons Learned, Drive status, Skills assessment |

---

*This document is authoritative for all infrastructure PE modeling work in this repository. Any deviation requires explicit justification and documentation.*
