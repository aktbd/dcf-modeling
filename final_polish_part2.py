#!/usr/bin/env python3
"""
Final Polish Part 2: IC Memo Templates + Study Guide Updates
"""

import os

os.chdir("/home/user/dcf-modeling/deliverables/deliverables-20260120-0030")


IC_MEMO_TEMPLATE = '''═══════════════════════════════════════════════════════════════════
LOTUS INFRASTRUCTURE PARTNERS
INVESTMENT COMMITTEE MEMORANDUM

Asset: {asset_name}
Date: ____________
Prepared By: ____________
═══════════════════════════════════════════════════════════════════

EXECUTIVE SUMMARY
─────────────────────────────────────────────────────────────────

Transaction:     Acquisition of {asset_description}
Entry EV:        $______mm
Capital Structure: ____% Debt / ____% Equity
Levered IRR:     ____%
MOIC:            ____x
Recommendation:  [ ] INVEST  [ ] PASS  [ ] CONDITIONAL

─────────────────────────────────────────────────────────────────
INVESTMENT THESIS (3 Key Points)
─────────────────────────────────────────────────────────────────

1. ________________________________________________________________
   Supporting data: ________________________________________________

2. ________________________________________________________________
   Supporting data: ________________________________________________

3. ________________________________________________________________
   Supporting data: ________________________________________________

─────────────────────────────────────────────────────────────────
KEY RISKS & MITIGANTS
─────────────────────────────────────────────────────────────────

Risk 1: ____________________________________________________________
  Probability: [ ] High [ ] Med [ ] Low  |  Impact: [ ] High [ ] Med [ ] Low
  Mitigant: ________________________________________________________

Risk 2: ____________________________________________________________
  Probability: [ ] High [ ] Med [ ] Low  |  Impact: [ ] High [ ] Med [ ] Low
  Mitigant: ________________________________________________________

Risk 3: ____________________________________________________________
  Probability: [ ] High [ ] Med [ ] Low  |  Impact: [ ] High [ ] Med [ ] Low
  Mitigant: ________________________________________________________

─────────────────────────────────────────────────────────────────
SENSITIVITY ANALYSIS
─────────────────────────────────────────────────────────────────

Base Case IRR: ____%

Key Sensitivities:
  • Entry EV +5%:           IRR → ____%
  • Exit Multiple -0.5x:    IRR → ____%
  • {key_driver} -10%:      IRR → ____%

Downside Case IRR: ____% (assumptions: ___________________________)
Break-even {key_driver}: _______________

─────────────────────────────────────────────────────────────────
DEAL STRUCTURE
─────────────────────────────────────────────────────────────────

SOURCES                          USES
  Debt:      $______mm (___%)      Purchase Price:  $______mm
  Equity:    $______mm (___%)      Transaction Fees: $______mm
                                   DSRA Funding:     $______mm
  ─────────────────────            ─────────────────────
  Total:     $______mm             Total:           $______mm

Debt Terms:
  • Rate: ____% | Tenor: ____ years | Amortization: [ ] Straight [ ] Sweep [ ] Sculpt
  • DSCR Covenant: ____x | DSRA: ____ months

─────────────────────────────────────────────────────────────────
DUE DILIGENCE PRIORITIES
─────────────────────────────────────────────────────────────────

[ ] Technical: ______________________________________________________
[ ] Commercial: _____________________________________________________
[ ] Legal/Regulatory: _______________________________________________
[ ] Environmental: __________________________________________________

─────────────────────────────────────────────────────────────────
RECOMMENDATION
─────────────────────────────────────────────────────────────────

________________________________________________________________
________________________________________________________________
________________________________________________________________

═══════════════════════════════════════════════════════════════════
'''


ASSET_SPECIFICS = {
    "CCGT": {
        "asset_name": "[CCGT Plant Name]",
        "asset_description": "___MW CCGT in [ISO/Region]",
        "key_driver": "Spark spread",
        "specific_section": '''
─────────────────────────────────────────────────────────────────
THERMAL-SPECIFIC CONSIDERATIONS
─────────────────────────────────────────────────────────────────

[ ] Spark spread: Current $____/MWh, Break-even $____/MWh
[ ] Heat rate: Plant ____ Btu/kWh vs Market HR ____ Btu/kWh
[ ] Capacity factor: ____% (basis: ________________)
[ ] Gas basis risk: Hub + $____ basis differential
[ ] Environmental: Carbon exposure? [ ] Y [ ] N  If Y: $____/MWh impact
[ ] Capacity market: ____ $/kW-yr locked through ____

'''
    },
    "Peaker": {
        "asset_name": "[Peaker Plant Name]",
        "asset_description": "___MW Peaker in [ISO/Region]",
        "key_driver": "Capacity price",
        "specific_section": '''
─────────────────────────────────────────────────────────────────
PEAKER-SPECIFIC CONSIDERATIONS
─────────────────────────────────────────────────────────────────

[ ] Dispatch hours: ____ hrs/year (vs ____ market avg)
[ ] Heat rate: ____ Btu/kWh (acceptable for peaker: <11,000)
[ ] Start capability: ____ minutes (fast start premium: [ ] Y [ ] N)
[ ] Capacity revenue: ____% of total revenue
[ ] Dual fuel: [ ] Y [ ] N  If Y: ____ days oil storage
[ ] Scarcity pricing exposure: [ ] High [ ] Med [ ] Low

'''
    },
    "SolarBESS": {
        "asset_name": "[Solar Project Name]",
        "asset_description": "___MW Solar + ___MWh BESS in [State]",
        "key_driver": "PPA price",
        "specific_section": '''
─────────────────────────────────────────────────────────────────
SOLAR-SPECIFIC CONSIDERATIONS
─────────────────────────────────────────────────────────────────

[ ] Capacity factor: ____% (Y10 output = ____% of Y1)
[ ] Degradation: ____% per year
[ ] ITC/PTC status: [ ] Claimed [ ] Remaining [ ] N/A
[ ] Curtailment risk: ____% historical, ____% assumed
[ ] Battery revenue: [ ] Arbitrage [ ] Ancillary [ ] Capacity
[ ] Inverter replacement: Year ____, $____mm estimated
[ ] Panel warranty: ____ years remaining

'''
    },
    "Transmission": {
        "asset_name": "[Transmission Line Name]",
        "asset_description": "____MW transmission line from [A] to [B]",
        "key_driver": "Rate base ROE",
        "specific_section": '''
─────────────────────────────────────────────────────────────────
TRANSMISSION-SPECIFIC CONSIDERATIONS
─────────────────────────────────────────────────────────────────

[ ] Rate base: $____mm
[ ] Allowed ROE: ____% (FERC approved: [ ] Y [ ] N)
[ ] Remaining asset life: ____ years
[ ] Rate case timing: Next filing ____
[ ] Congestion revenue rights: [ ] Y [ ] N
[ ] Permitting status: [ ] Complete [ ] Pending: ________________

'''
    },
    "Midstream": {
        "asset_name": "[Midstream System Name]",
        "asset_description": "____mmcf/d gathering system in [Basin]",
        "key_driver": "Volume",
        "specific_section": '''
─────────────────────────────────────────────────────────────────
MIDSTREAM-SPECIFIC CONSIDERATIONS
─────────────────────────────────────────────────────────────────

[ ] Current throughput: ____ mmcf/d (____% of capacity)
[ ] MVC protection: ____% of capacity through ____
[ ] Acreage dedication: ____ acres, ____ years remaining
[ ] Producer concentration: Top producer = ____% of volume
[ ] Basin economics: Breakeven gas price $____/mcf
[ ] Decline rate: ____% after Year ____

'''
    },
    "Wind": {
        "asset_name": "[Wind Farm Name]",
        "asset_description": "___MW wind farm in [State/ISO]",
        "key_driver": "Capacity factor",
        "specific_section": '''
─────────────────────────────────────────────────────────────────
WIND-SPECIFIC CONSIDERATIONS
─────────────────────────────────────────────────────────────────

[ ] Net capacity factor: ____% (P50/P90 spread: ____%)
[ ] PTC remaining: ____ years, $____mm/year value
[ ] Post-PPA pricing: $____/MWh (____% below PPA)
[ ] Curtailment: ____% historical
[ ] O&M: $____/kW-yr (escalation: ____%)
[ ] Turbine technology: [ ] Modern [ ] Aging - major component risk?
[ ] Gearbox/blade status: Last inspection ____, condition ____

'''
    }
}


def create_ic_memo_template(asset_type):
    """Create IC Memo template for given asset type."""
    specs = ASSET_SPECIFICS.get(asset_type, {})

    template = IC_MEMO_TEMPLATE.format(
        asset_name=specs.get("asset_name", "[Asset Name]"),
        asset_description=specs.get("asset_description", "[Asset Description]"),
        key_driver=specs.get("key_driver", "[Key Driver]"),
    )

    # Add asset-specific section
    specific = specs.get("specific_section", "")
    template += specific

    return template


def update_ic_memos():
    """Update all IC Memos to template format."""
    print("\n--- UPDATING IC MEMO TEMPLATES ---")

    memo_mapping = {
        "IC_Memo_1_CCGT.docx": "CCGT",
        "IC_Memo_2_Peaker.docx": "Peaker",
        "IC_Memo_3_SolarBESS.docx": "SolarBESS",
        "IC_Memo_4_Transmission.docx": "Transmission",
        "IC_Memo_5_Midstream.docx": "Midstream",
        "IC_Memo_6_Wind.docx": "Wind",
    }

    for filename, asset_type in memo_mapping.items():
        print(f"\n  Creating template for {filename}...")
        template = create_ic_memo_template(asset_type)

        # Save as markdown (can be converted to DOCX)
        md_filename = filename.replace(".docx", "_Template.md")
        with open(md_filename, "w") as f:
            f.write(template)
        print(f"    Created {md_filename}")

        # Try to create DOCX
        try:
            from docx import Document
            from docx.shared import Pt
            from docx.enum.text import WD_ALIGN_PARAGRAPH

            doc = Document()

            # Add content as formatted paragraphs
            lines = template.split('\n')
            for line in lines:
                p = doc.add_paragraph(line)
                if '═══' in line or '───' in line:
                    p.runs[0].font.size = Pt(10)
                elif line.startswith('LOTUS') or 'EXECUTIVE' in line or 'THESIS' in line:
                    p.runs[0].font.bold = True
                    p.runs[0].font.size = Pt(12)

            doc.save(filename)
            print(f"    Updated {filename}")

        except ImportError:
            print(f"    (python-docx not available, created markdown only)")


STUDY_GUIDE_ADDITION = '''

---

## 11. ADVANCED SCENARIOS (Edge Cases)

### 11A. Preferred Equity / Mezz

When you see "mezz" or "preferred" in the capital structure:

```
Sources:
  Senior Debt:     $XXmm  (60%)
  Preferred:       $XXmm  (15%)  ← ADD THIS
  Common Equity:   $XXmm  (25%)
```

**Mechanics:**
- Pref Coupon = Pref Amount × Rate (typically 8-12%)
- Pref coupon comes OUT of distributable cash before common
- Common IRR calculated on common equity only
- Pref is NOT debt (doesn't count in DSCR denominator)

**Waterfall:**
```
CFADS
  Less: Debt Service
  Less: DSRA Funding
  Less: Pref Coupon    ← NEW
  Equals: Distributable to Common
```

### 11B. Time Series Pricing

When given a price curve instead of escalation rate:

```
Year    Power Price
1       $42
2       $45
3       $48
4       $44
5       $50
...
```

**Approach:**
- Input prices directly in a row
- Reference cells individually: `=E$25` not `=$D$25*(1+esc)^...`
- Watch for embedded "cliffs" or shocks in the curve
- Note any reversion to mean in outer years

### 11C. Construction Period

When COD (Commercial Operation Date) is in the future:

**Key Mechanics:**
- Y0-Y2: Zero revenue, construction costs
- AFUDC: Interest capitalizes to cost basis
- Two financing events: Construction loan → Term debt at COD
- Development fees often treated as Uses

**Simplified Approach (if time-constrained):**
- Start model at COD
- Treat all construction as "Entry EV"
- Note assumption in memo

### 11D. ERCOT (Texas) Modeling

Energy-only market characteristics:

```
Revenue = Energy Only (no capacity payments!)
  • Higher price volatility
  • Scarcity pricing: $5,000/MWh cap
  • More weather-dependent
```

**Key Adjustments:**
- Capacity Revenue row = $0
- Higher energy price assumptions for peakers
- Note as key risk in IC memo
- Consider hedging/PPA as mitigant

### 11E. Partnership Flip / Tax Equity

If mentioned, they will give you the structure. Key concepts:

```
Year 1-10: Tax equity gets 99% of tax benefits
Year 10+: "Flip" to sponsor at 5% residual value
```

**In a timed test:**
- Don't try to build full tax equity waterfall
- Model simplified: adjust effective tax rate or add tax benefit row
- Note in memo: "Assumes tax equity monetization per provided terms"

---

## 12. PROXY DEFAULTS (When Inputs Are Missing)

### 12.1 Thermal Plants

| Input | CCGT | Peaker | Units |
|-------|------|--------|-------|
| Heat Rate | 7,000 | 10,500 | Btu/kWh |
| Capacity Factor | 55% | 10% | % |
| Forced Outage Rate | 5% | 6% | % |
| Fixed O&M | $15 | $12 | $/kW-yr |
| Variable O&M | $3 | $5 | $/MWh |
| Start Cost | - | $50 | $/start |

### 12.2 Renewables

| Input | Solar | Wind | Units |
|-------|-------|------|-------|
| Capacity Factor | 25% | 38% | % |
| Degradation | 0.5% | 0.3% | %/yr |
| O&M | $15 | $40 | $/kW-yr |
| Inverter Life | 15 | - | years |

### 12.3 Market Prices

| Market | Power ($/MWh) | Capacity ($/kW-yr) | Gas ($/MMBtu) |
|--------|---------------|-------------------|---------------|
| PJM | $45 | $100 | $3.50 |
| ERCOT | $40 | $0 | $3.25 |
| NYISO | $55 | $150 | $4.00 |
| CAISO | $50 | $80 | $4.50 |
| SPP | $35 | $50 | $3.00 |
| MISO | $38 | $40 | $3.25 |
| ISO-NE | $48 | $80 | $5.00 |

### 12.4 Financing Defaults

| Input | Merchant | Contracted | Units |
|-------|----------|------------|-------|
| Leverage | 55-60% | 65-75% | % |
| Interest Rate | 6.5% | 5.5% | % |
| Tenor | 5-7 | 7-15 | years |
| Target DSCR | 1.25x | 1.35x | x |
| DSRA | 6 | 6 | months |
| Sweep % | 75% | 50% | % |

### 12.5 Transaction Defaults

| Input | Default | Notes |
|-------|---------|-------|
| Transaction Fees | 1.5% | of Entry EV |
| Exit Multiple | Entry + 0.5x | if assets improving |
| Exit Multiple | Entry | if assets stable |
| Exit Multiple | Entry - 0.5x | if assets declining |
| Hold Period | 7 | years (unless specified) |
| Tax Rate | 25% | blended federal + state |

### 12.6 When to Use Proxies

**Use proxy if:**
- Input not provided in case
- Interviewer says "use your judgment"
- Need to move forward to finish model

**Don't use proxy if:**
- Specific number given in case
- Interviewer provides clarification
- It's a key driver (ask first)

**Always note in memo:**
"[Input] assumed at [value] per market convention"

---

*End of Study Guide - Good luck on your interviews!*
'''


def update_study_guide():
    """Add Sections 11 & 12 to Study Guide."""
    print("\n--- UPDATING STUDY GUIDE ---")

    with open("Lotus_Study_Guide.md", "r") as f:
        content = f.read()

    # Check if already updated
    if "## 11. ADVANCED SCENARIOS" in content:
        print("  Sections 11 & 12 already exist")
        return

    # Remove old ending if present
    if "*This guide is for study" in content:
        content = content.split("*This guide is for study")[0]

    # Add new sections
    content += STUDY_GUIDE_ADDITION

    with open("Lotus_Study_Guide.md", "w") as f:
        f.write(content)

    print("  Added Section 11: Advanced Scenarios")
    print("  Added Section 12: Proxy Defaults")


def main():
    print("=" * 70)
    print("FINAL POLISH PART 2: IC Memos + Study Guide")
    print("=" * 70)

    update_ic_memos()
    update_study_guide()

    print("\n" + "=" * 70)
    print("IC MEMO & STUDY GUIDE UPDATES COMPLETE")
    print("=" * 70)


if __name__ == "__main__":
    main()
