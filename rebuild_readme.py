#!/usr/bin/env python3
"""Rebuild README sheets with comprehensive case content."""

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
import os

os.chdir("/home/user/dcf-modeling/deliverables/deliverables-20260120-0030")

CASE_CONTENT = {
    "CCGT": {
        "full_name": "Combined Cycle Gas Turbine",
        "location": "PJM West (AEP Zone)",
        "entry_ev": 520,

        "situation": """Lotus Infrastructure Partners has been approached by GenCo Partners regarding
the potential acquisition of a 365 MW combined-cycle gas turbine power plant
located in PJM West (AEP zone). The facility achieved commercial operation in
2015 and has approximately 15 years of remaining useful life.

The plant operates as a fully merchant facility, selling energy into PJM's
day-ahead and real-time markets and clearing in the RPM capacity auction.
Historical capacity factors have averaged 50-60%. The 7,000 Btu/kWh heat rate
makes it competitive against older thermal generation.

GenCo is divesting non-core assets to fund renewable development. They are
seeking $520mm with final bids expected by February 15, 2026.""",

        "timeline": """Year 0 (2025): Transaction close Q4; debt funding at close
Year 1 (2026): First full operating year under Lotus ownership
Year 7 (2032): Target exit / debt maturity (75% cash sweep structure)
Year 10 (2035): Model horizon end""",

        "key_drivers": [
            ("Spark Spread", "Power price minus (Heat Rate x Gas Price) - core margin driver"),
            ("Capacity Factor", "Drives MWh generation; modeled NET of outages"),
            ("Capacity Accreditation", "UCAP rating determines capacity revenue"),
            ("Gas Basis", "Spread between Henry Hub and local delivery point"),
        ],

        "model_guide": """STEP-BY-STEP MODEL BUILDING GUIDE

1. SET UP INPUTS (Column D)
   - Transaction: Entry EV ($520mm), fees (1.5%), exit multiple (7x), exit year (7)
   - Asset: Capacity (365 MW), heat rate (7,000), CF (55%), FOR (5%)
   - Pricing: Power ($48/MWh), capacity ($110/kW-yr), gas ($3/MMBtu), esc (2.5%)
   - Financing: Leverage (45%), rate (6.5%), tenor (7 yrs), DSRA (6 mo), sweep (75%)

2. BUILD CALCULATED ITEMS
   - UCAP = Capacity x (1 - FOR) = 365 x 0.95 = 346.75 MW
   - Generation = Capacity x CF x 8,760 = 1,758,900 MWh
   - Debt Amount = Entry EV x Leverage = $234mm

3. BUILD OPERATING MODEL (Columns E-N)
   - Escalate prices: Yn = Y1 x (1 + esc)^(n-1)
   - Energy Revenue = Generation x Power Price / 1,000,000
   - Capacity Revenue = UCAP x Cap Price x 1,000 / 1,000,000
   - Fuel Cost = Generation x Heat Rate x Gas / 1,000,000,000
   - EBITDA = Revenue - OpEx

4. BUILD CASH FLOW
   - EBIT = EBITDA - Depreciation
   - Taxes = MAX(0, EBIT) x Tax Rate
   - CFADS = EBITDA - Taxes

5. BUILD DEBT SCHEDULE (CASH SWEEP)
   a) Beginning Balance: Y1 = Debt; Yn = Prior End Bal
   b) Interest = Beg Bal x Rate
   c) Scheduled Principal = MIN(Debt/Tenor, Beg Bal)
   d) Debt Service (Pre-Sweep) = Interest + Sched Prin
   e) Scheduled DS (for DSRA) = Independent calc using straight-line
   f) DSRA Target = NEXT Year's Sched DS x (DSRA Mo / 12)
   g) Excess Cash = CFADS - DS - DSRA Funding
   h) Cash Sweep = MIN(Excess x 75%, Beg Bal - Sched Prin)
   i) Ending Balance = MAX(0, Beg - Total Principal)
   j) DSCR = CFADS / DS (Pre-Sweep)

6. BUILD S&U, EXIT, RETURNS
   - S&U must balance (Uses = Sources)
   - Exit EV = EBITDA x Multiple
   - Levered IRR = IRR of equity cash flows""",

        "circularity": """HOW TO AVOID DSRA CIRCULARITY

The Problem:
DSRA Target -> Debt Service -> Sweep -> Excess Cash -> DSRA Funding -> DSRA Target

The Solution:
Use "Scheduled DS" (straight-line, no sweep) for DSRA sizing:
- Scheduled DS assumes NO cash sweep, just scheduled payments
- It's calculated INDEPENDENTLY of actual debt paydown
- DSRA Target references this independent row

Formula: Scheduled DS = MAX(0, Debt - (Yr-1) x Sched Prin) x Rate + Sched Prin""",

        "alternatives": """ALTERNATIVE SCENARIOS

A. Sculpted Debt (not sweep):
   - Principal = CFADS / Target DSCR - Interest
   - No sweep mechanics needed

B. Contracted Revenue (PPA/Hedge):
   - Use contract price, not merchant
   - More debt capacity

C. Major Maintenance Reserve:
   - Add MRA similar to DSRA
   - Funds hot gas path inspection""",

        "qa_checks": """QA CHECKS TO VERIFY MODEL

1. S&U Balance: =IF(ABS(Uses-Sources)<0.01,"PASS","FAIL")
2. Debt Payoff: Ending Bal near $0 by tenor end
3. Min DSCR: Use MINIFS to exclude post-payoff zeros
4. DSRA: Never negative; releases at debt payoff
5. Sign Convention: Outflows negative, inflows positive
6. No Circular Refs: Excel shouldn't show warning""",

        "hints": [
            "Spark spread: $1/MWh x 1.76M MWh = $1.76mm EBITDA impact",
            "CF is NET of outages - don't double-count FOR in generation",
            "FOR only derates UCAP for capacity revenue",
            "Fuel divisor is 1E9 (billion), not 1E6 (million)"
        ],

        "questions": [
            "Walk me through how $5/MWh power price change affects IRR.",
            "Why cash sweep instead of sculpted debt?",
            "What would make you walk away from this deal?",
            "How would you hedge the spark spread exposure?"
        ]
    },

    "Peaker": {
        "full_name": "Simple Cycle Gas Turbine (Peaker)",
        "location": "NYISO Zone J (New York City)",
        "entry_ev": 95,

        "situation": """Lotus Infrastructure Partners is evaluating a 200 MW simple-cycle gas
turbine peaking facility in NYISO Zone J (New York City). The plant is 8 years
old with approximately 2,500 equivalent operating starts.

Zone J has strong capacity prices due to transmission constraints. The facility
provides reliability during summer peaks and winter cold snaps. It operates
approximately 400 hours per year with revenue heavily weighted to capacity
payments (~70% of total).

The 10,500 Btu/kWh heat rate is higher than CCGT but acceptable for peaking
duty. Recent planning studies indicate potential capacity additions over 3-5
years including offshore wind and battery storage.

The seller is asking $95mm with bids due by January 31, 2026.""",

        "timeline": """Year 0 (2025): Transaction close Q4
Year 1 (2026): First full year; start-based maintenance tracking
Year 5 (2030): Potential hot gas path inspection
Year 7 (2032): Target exit / debt maturity
Year 10 (2035): Model horizon end""",

        "key_drivers": [
            ("Capacity Price", "Peakers earn 70%+ of revenue from capacity auctions"),
            ("Dispatch Hours", "Annual run hours drive energy revenue and fuel cost"),
            ("Starts", "Each start costs money and accelerates maintenance"),
            ("Availability", "UCAP derate risk if availability drops below threshold"),
        ],

        "model_guide": """STEP-BY-STEP MODEL BUILDING GUIDE FOR PEAKERS

KEY DIFFERENCE: Generation = Capacity x Dispatch Hours x Availability
(NOT Capacity x CF x 8,760 like CCGT!)

1. SET UP INPUTS
   - Capacity: 200 MW
   - Heat Rate: 10,500 Btu/kWh (worse than CCGT)
   - Availability: 94% (mechanical availability, not CF)
   - FOR: 5% (for UCAP derate)
   - Dispatch Hours: 400 hrs/yr
   - Starts per Year: ~50-100

2. BUILD GENERATION
   Generation = Capacity x Dispatch Hours x Availability
   = 200 MW x 400 hrs x 0.94 = 75,200 MWh

3. CAPACITY REVENUE IS DOMINANT
   UCAP = Capacity x (1 - FOR) = 200 x 0.95 = 190 MW
   Capacity Revenue = UCAP x $/kW-yr x 1000 / 1E6
   Verify: Capacity should be ~70% of total revenue

4. DEBT STRUCTURE
   50% cash sweep (vs 75% for CCGT)
   - Preserve flexibility for lumpy maintenance
   - Peaker cash flows more volatile""",

        "circularity": """Same approach as CCGT - use Scheduled DS for DSRA Target to avoid circularity.""",

        "alternatives": """ALTERNATIVE SCENARIOS

A. Ancillary Services Revenue:
   - Spinning/non-spinning reserves
   - Can be 10-15% of total revenue

B. Black Start Capability:
   - Premium for grid restoration
   - Adds contracted revenue

C. Capacity Contract vs Merchant:
   - Different risk profile
   - Affects debt sizing""",

        "qa_checks": """QA CHECKS

1. Capacity revenue should be 60-75% of total (peaker profile)
2. Generation = Dispatch Hours x Availability (NOT CF x 8760)
3. Energy margin may be negative in some hours (heat rate penalty)
4. DSCR >= 1.25x during debt tenor""",

        "hints": [
            "Capacity is ~70% of revenue - focus sensitivity there",
            "Each start costs money plus accelerates maintenance",
            "Zone J premium can erode if storage/transmission added",
            "Don't model generation as CF x 8760 - use Dispatch Hours"
        ],

        "questions": [
            "How would 100 MW of battery storage in Zone J affect capacity prices?",
            "Walk me through starts-based vs hours-based maintenance.",
            "Why is availability so critical for peaker economics?",
            "How would you model the revenue mix between capacity and energy?"
        ]
    },

    "SolarBESS": {
        "full_name": "Solar PV + Battery Energy Storage",
        "location": "ERCOT West Zone",
        "entry_ev": 280,

        "situation": """Lotus Infrastructure Partners is evaluating a 150 MW solar PV facility
with co-located 75 MW / 300 MWh battery storage in ERCOT West zone. The
project achieved COD in Q1 2024 (operating ~2 years).

The solar facility has a 15-year fixed-price PPA at $35/MWh with an IG utility
offtaker (A- rating). Approximately 13 years remain from close. Post-PPA
revenue depends on merchant market conditions.

The battery generates revenue through energy arbitrage and ancillary services,
limited to ~365 cycles/year for warranty preservation. The project received
30% ITC at COD (already monetized) with MACRS depreciation through Year 6.

The seller is asking $280mm with bids due by February 28, 2026.""",

        "timeline": """Year 0 (2025): Transaction close Q4
Year 1 (2026): First full year; PPA Year 3
Year 13 (2038): PPA expiration; merchant transition
Year 15 (2040): End of typical solar asset life""",

        "key_drivers": [
            ("PPA Price/Tenor", "Contracted revenue at $35/MWh for 13 remaining years"),
            ("Degradation", "0.5%/year panel efficiency loss - compounds over time"),
            ("Battery Cycling", "365 cycles/year limit; affects arbitrage opportunity"),
            ("Post-PPA Merchant", "Tail value uncertain; conservative assumptions key"),
        ],

        "model_guide": """STEP-BY-STEP MODEL BUILDING GUIDE FOR SOLAR+BESS

KEY DIFFERENCES:
- NO fuel cost - margins effectively 100%
- Degradation reduces output each year (compounds!)
- PPA provides contracted revenue
- Tax treatment: ITC taken, MACRS shield through Y6

1. SET UP INPUTS
   - Solar Capacity: 150 MW
   - Capacity Factor: 25-28%
   - Degradation: 0.5%/year (COMPOUNDS)
   - PPA Price: $35/MWh (fixed)
   - PPA Tenor: 13 years remaining
   - Post-PPA Price: Conservative merchant

2. BUILD GENERATION WITH DEGRADATION
   Year 1 Generation = Capacity x CF x 8760
   Year n Generation = Year 1 x (1 - Degradation)^(n-1)
   Example: Y10 = Y1 x 0.995^9 = ~95.6% of Y1

3. BUILD TAX SCHEDULE
   - Years 1-6: MACRS shield = ZERO taxes
   - Years 7+: Normal taxation
   Formula: =IF(Year <= Shield_Years, 0, MAX(0, EBIT) x Rate)

4. SCULPTED DEBT (NOT Cash Sweep)
   Solar+BESS with PPA uses SCULPTED debt:
   - Cash flows predictable (contracted)
   - Lenders comfortable with 1.30-1.40x DSCR
   Sculpted Principal = MIN(MAX(0, CFADS/Target_DSCR - Interest), Beg_Bal)""",

        "circularity": """SCULPTED DEBT - NO CIRCULARITY ISSUE

Unlike cash sweep, sculpted debt is calculated FROM CFADS:
1. Know CFADS (revenue - opex - taxes)
2. Calculate Target DS = CFADS / Target DSCR
3. Calculate Principal = Target DS - Interest
4. DSRA Target references this known DS

No circular loop because everything flows from CFADS downward.""",

        "alternatives": """ALTERNATIVE SCENARIOS

A. Tax Equity Partnership:
   - If ITC not yet monetized
   - Complex waterfall with flip date

B. Merchant Solar (No PPA):
   - Much riskier - lower leverage
   - May not be financeable with project debt

C. Standalone BESS:
   - No solar, just storage
   - Revenue from pure arbitrage""",

        "qa_checks": """QA CHECKS

1. Degradation compounds correctly (Y10 < 95% of Y1)
2. Tax is $0 for MACRS shield years
3. Tax formula has MAX(0,...) guardrail
4. PPA revenue switches to merchant at correct year
5. DSCR is ~1.30-1.40x (sculpted to target)""",

        "hints": [
            "Degradation at 0.5%/yr means Y10 is ~95.6% of Y1 (not 95%)",
            "ITC already taken - no tax equity modeling needed",
            "Post-PPA price should be conservative (70-80% of PPA)",
            "MACRS shield means ~6 years of zero taxes"
        ],

        "questions": [
            "How does 0.5% vs 0.7% annual degradation affect exit value?",
            "What's your post-PPA merchant assumption and why?",
            "Why use sculpted debt instead of cash sweep?",
            "What would make you walk away from this deal?"
        ]
    },

    "Transmission": {
        "full_name": "Electric Transmission Line",
        "location": "SPP (Southwest Power Pool)",
        "entry_ev": 180,

        "situation": """Lotus Infrastructure Partners is evaluating equity investment in a
150-mile, 345 kV transmission line under development in SPP. FERC has
approved 10.5% ROE with 50/50 regulatory capital structure.

Development began in 2025 (Y0), with construction in 2026-2027 and COD
expected in Year 3 (2028). Total project cost is ~$585mm including
development, construction, and AFUDC during construction.

Upon COD, revenue comes from FERC-approved tariff based on rate base and
allowed ROE. The line has 40+ years expected useful life.

Lotus is offered $180mm equity during construction. Commitment needed by
March 15, 2026.""",

        "timeline": """Year 0 (2025): Development costs; construction loan drawn
Year 1 (2026): Major construction; AFUDC accrues
Year 2 (2027): Construction completion
Year 3 (2028): COD - Term debt takeout; tariff revenue starts
Year 10 (2035): Model horizon (asset life 40+ years)""",

        "key_drivers": [
            ("FERC ROE", "Allowed return on equity sets revenue requirement"),
            ("Rate Base", "Capital invested that earns the allowed return"),
            ("Construction Timing", "Delays reduce returns; AFUDC accrues during build"),
            ("Regulatory Risk", "ROE can change in FERC rate case proceedings"),
        ],

        "model_guide": """STEP-BY-STEP MODEL BUILDING GUIDE FOR TRANSMISSION

KEY DIFFERENCES:
- CONSTRUCTION PERIOD before operations
- AFUDC (capitalized interest) accrues to rate base
- Revenue is REGULATED formula, not market prices
- Two financing events: Construction loan -> Term debt takeout

1. CONSTRUCTION PERIOD (Years 0-2)
   Development Costs: $35mm
   Construction Costs: $550mm
   Construction Loan: 85% of construction
   AFUDC = Const Loan x Interest x Avg Period = ~$35mm
   AFUDC accrues to Rate Base (non-cash during construction)

2. RATE BASE AT COD (Year 3)
   Rate Base = Development + Construction + AFUDC
   = $35mm + $550mm + ~$35mm = ~$620mm

3. TERM DEBT TAKEOUT AT COD
   At COD, Construction Loan repaid by Term Debt
   Term Debt = Rate Base x Regulatory Debt % = $620mm x 50% = $310mm
   Use RATE BASE (not construction cost) for debt sizing

4. REVENUE CALCULATION (Years 3+)
   Revenue = Rate Base x (ROE + Depreciation/Life + O&M/RB)

5. S&U - TWO EVENTS
   At Close (Y0): Development costs, fees -> Const loan, equity
   At COD (Y3): Const loan payoff -> Term debt, permanent equity
   DSRA is a USE (funded by debt/equity), NOT a Source

6. SCULPTED DEBT
   Regulated cash flows support sculpted debt at 1.40x DSCR""",

        "circularity": """TRANSMISSION - SPECIAL CONSIDERATIONS

No sweep mechanics - sculpted debt only.
AFUDC creates complexity but not circularity.

S&U RULES:
- DSRA is a Use, NOT a Source
- Balance check compares Total Uses vs Total Sources
- Construction loan payoff must be included at COD""",

        "alternatives": """ALTERNATIVE SCENARIOS

A. Acquisition at COD (No Construction):
   - Skip construction period
   - Buy operating asset at Rate Base multiple

B. Rate Case During Hold:
   - ROE may change in rate proceedings
   - Model +/- 50bp ROE sensitivity

C. Rate Base Growth:
   - Capital additions over time
   - Growing rate base = growing earnings""",

        "qa_checks": """QA CHECKS

1. AFUDC included in Rate Base at COD
2. S&U balances at BOTH close and COD
3. DSRA is only a Use, never a Source
4. Revenue ties to rate base x allowed returns
5. ROE sensitivity shows material impact""",

        "hints": [
            "AFUDC is often $20-40mm - don't forget to capitalize it",
            "50bp ROE change at 10x rate base = 5% equity value impact",
            "Construction delays kill returns - model delay scenarios",
            "DSRA is a Use funded by debt/equity, NOT a source"
        ],

        "questions": [
            "Walk me through AFUDC and how it affects rate base at COD.",
            "How does a 50bp ROE change affect your equity returns?",
            "What construction risks would you prioritize in diligence?",
            "How would you structure the construction loan takeout?"
        ]
    },

    "Midstream": {
        "full_name": "Natural Gas Gathering System",
        "location": "Permian Basin (Delaware Sub-basin)",
        "entry_ev": 85,

        "situation": """Lotus Infrastructure Partners is evaluating a natural gas gathering
system in the Permian Basin (Delaware sub-basin). Nameplate capacity is
500 MMcf/d with current throughput of ~80 MMcf/d from 15 producer connections.

The anchor shipper contract covers 50 MMcf/d at $0.55/Mcf through 2030
(5 years remaining) with 80% take-or-pay. Remaining ~30 MMcf/d is under
short-term (1-2 year) gathering agreements.

Basin activity is strong with rig counts up 15% YoY. However, producer
economics are sensitive to WTI and Henry Hub prices. The system has
expansion capacity without major capex.

The seller (PE fund approaching end of fund life) is asking $85mm with
bids due by February 10, 2026.""",

        "timeline": """Year 0 (2025): Transaction close Q4
Year 1-5 (2026-2030): Growth phase; volumes increase with new connections
Year 5 (2030): Anchor shipper contract expires; recontracting risk
Year 6-10 (2031-2035): Potential decline phase if no new drilling""",

        "key_drivers": [
            ("Volume Trajectory", "Producer drilling activity drives throughput"),
            ("Take-or-Pay", "80% MVC provides revenue floor on contracted volumes"),
            ("Recontracting Risk", "Anchor contract expiry in 2030 is key risk"),
            ("Basin Economics", "WTI/HH prices drive producer drilling decisions"),
        ],

        "model_guide": """STEP-BY-STEP MODEL BUILDING GUIDE FOR MIDSTREAM

KEY DIFFERENCES:
- Volume-based revenue (not capacity x price)
- Producer counterparty credit risk
- Volume decline curves for existing wells
- Recontracting risk at contract expiry
- Gathering fee is $/Mcf

1. SET UP INPUTS
   - Y1 Volume: 80 MMcf/d (current throughput)
   - Gathering Fee: $0.55/Mcf
   - Growth Rate Y1-5: 3%/year
   - Decline Rate Y6-10: 5%/year
   - O&M: $3.5mm/year base

2. BUILD VOLUME TRAJECTORY
   Daily Volume Y1 = 80 MMcf/d
   Daily Volume Yn = Y(n-1) x (1 + Growth/Decline Rate)
   Annual Volume (Bcf) = Daily Volume x 365 / 1000

3. BUILD REVENUE
   Revenue ($mm) = Annual Volume (Bcf) x Fee ($/Mcf)

   Why this works:
   - 1 Bcf = 1,000,000 Mcf
   - Revenue ($) = Bcf x 1,000,000 x $/Mcf
   - Revenue ($mm) = Revenue ($) / 1,000,000
   - Therefore: Revenue ($mm) = Bcf x $/Mcf (factors cancel)

   NO *1000 multiplier needed!

   Sanity Check: 29 Bcf x $0.55/Mcf = ~$16mm revenue
   This is reasonable for $85mm EV (5x EV/Revenue)

4. BUILD CFADS
   CFADS = Revenue - O&M - Taxes - Growth Capex
   Growth capex (wellhead connections) is BEFORE CFADS for midstream

5. DEBT STRUCTURE
   75% cash sweep (volume uncertainty warrants sweep)
   Shorter tenor (5 years) due to recontracting risk""",

        "circularity": """Same as CCGT - use Scheduled DS for DSRA Target to avoid circularity.""",

        "alternatives": """ALTERNATIVE SCENARIOS

A. Processing vs Gathering:
   - If system includes processing plant
   - Processing revenue is $/MMBtu (different unit)

B. Commodity Exposure:
   - If fee has commodity-linked component
   - Must model WTI/HH price scenarios

C. Acreage Dedication:
   - If contract includes acreage dedication
   - Provides volume upside if producer drills more""",

        "qa_checks": """QA CHECKS

1. Revenue formula does NOT have erroneous *1000
2. Y1 Revenue / Entry EV = reasonable multiple (4-6x typical)
3. Volume trajectory makes sense (growth then decline)
4. Take-or-pay floor revenue = MVC x Fee
5. Recontracting scenario modeled for Year 5+
6. DSCR >= 1.25x during debt tenor""",

        "hints": [
            "Revenue = Bcf x $/Mcf (NO conversion factor - they cancel)",
            "Y1 Revenue of ~$16mm on $85mm EV is ~5.3x, reasonable",
            "Recontracting at Year 5 is THE key risk - show sensitivity",
            "Take-or-pay provides 80% floor on contracted volume"
        ],

        "questions": [
            "Walk me through how you modeled volume decline.",
            "What's your recontracting assumption when anchor expires?",
            "How would a $10/bbl drop in WTI affect volume projections?",
            "What producer credit diligence would you conduct?"
        ]
    }
}


def rebuild_readme(wb, asset_type, content):
    """Rebuild README as comprehensive case prompt."""

    if "README" in wb.sheetnames:
        del wb["README"]

    ws = wb.create_sheet("README", 0)

    row = 1

    # Title block
    ws.cell(row=row, column=1).value = "=" * 75
    row += 1
    ws.cell(row=row, column=1).value = "LOTUS INFRASTRUCTURE PARTNERS"
    ws.cell(row=row, column=1).font = Font(bold=True, size=14)
    row += 1
    ws.cell(row=row, column=1).value = f"CASE STUDY: {content['full_name'].upper()} ACQUISITION"
    ws.cell(row=row, column=1).font = Font(bold=True, size=14)
    row += 1
    ws.cell(row=row, column=1).value = "=" * 75
    row += 2

    # Metadata
    ws.cell(row=row, column=1).value = f"TIME: 4 Hours | DATE: January 2026 | DELIVERABLES: Model + IC Memo"
    row += 2

    # Timeline
    ws.cell(row=row, column=1).value = "--- MODEL TIMELINE ---"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1
    for line in content["timeline"].strip().split('\n'):
        ws.cell(row=row, column=1).value = line
        row += 1
    row += 1

    # Situation
    ws.cell(row=row, column=1).value = "--- SITUATION ---"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1
    for para in content["situation"].split('\n\n'):
        ws.cell(row=row, column=1).value = para.strip()
        ws.cell(row=row, column=1).alignment = Alignment(wrap_text=True)
        row += 2

    # Key Drivers
    ws.cell(row=row, column=1).value = "--- KEY VALUE DRIVERS ---"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1
    for driver, desc in content["key_drivers"]:
        ws.cell(row=row, column=1).value = f"* {driver}: {desc}"
        row += 1
    row += 1

    # Model Building Guide
    ws.cell(row=row, column=1).value = "--- STEP-BY-STEP MODEL BUILDING GUIDE ---"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1
    for line in content["model_guide"].strip().split('\n'):
        ws.cell(row=row, column=1).value = line
        ws.cell(row=row, column=1).alignment = Alignment(wrap_text=True)
        row += 1
    row += 1

    # Circularity Guide
    ws.cell(row=row, column=1).value = "--- HOW TO AVOID CIRCULARITY ---"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1
    for line in content["circularity"].strip().split('\n'):
        ws.cell(row=row, column=1).value = line
        row += 1
    row += 1

    # Alternative Scenarios
    ws.cell(row=row, column=1).value = "--- ALTERNATIVE SCENARIOS ---"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1
    for line in content["alternatives"].strip().split('\n'):
        ws.cell(row=row, column=1).value = line
        row += 1
    row += 1

    # QA Checks
    ws.cell(row=row, column=1).value = "--- QA CHECKS ---"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1
    for line in content["qa_checks"].strip().split('\n'):
        ws.cell(row=row, column=1).value = line
        row += 1
    row += 1

    # Hints
    ws.cell(row=row, column=1).value = "--- HINTS & PITFALLS ---"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1
    for hint in content["hints"]:
        ws.cell(row=row, column=1).value = f"* {hint}"
        ws.cell(row=row, column=1).font = Font(italic=True, color="0066CC")
        row += 1
    row += 1

    # Interview Questions
    ws.cell(row=row, column=1).value = "--- INTERVIEW QUESTIONS ---"
    ws.cell(row=row, column=1).font = Font(bold=True)
    row += 1
    for q in content["questions"]:
        ws.cell(row=row, column=1).value = f"Q: {q}"
        row += 1

    ws.column_dimensions['A'].width = 95

    return row  # Return row count


print("=" * 60)
print("TASK 2: REBUILDING README SHEETS")
print("=" * 60)

# Process Models
for filename, asset in [
    ("Model_1_CCGT.xlsx", "CCGT"),
    ("Model_2_Peaker.xlsx", "Peaker"),
    ("Model_3_SolarBESS.xlsx", "SolarBESS"),
    ("Model_4_Transmission.xlsx", "Transmission"),
    ("Model_5_Midstream.xlsx", "Midstream"),
]:
    print(f"\n{filename}:")
    wb = load_workbook(filename)

    row_count = rebuild_readme(wb, asset, CASE_CONTENT[asset])
    print(f"  README rebuilt with {row_count} rows")

    # Reorder sheets
    order = ['README', 'Case Prompt', 'IC Memo', 'Model']
    for i, name in enumerate(order):
        if name in wb.sheetnames:
            wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

    wb.save(filename)
    print(f"  Sheets: {wb.sheetnames}")
    wb.close()

# Process Drills
for filename, asset in [
    ("Drill_1_CCGT.xlsx", "CCGT"),
    ("Drill_2_Peaker.xlsx", "Peaker"),
    ("Drill_3_SolarBESS.xlsx", "SolarBESS"),
    ("Drill_4_Transmission.xlsx", "Transmission"),
    ("Drill_5_Midstream.xlsx", "Midstream"),
]:
    print(f"\n{filename}:")
    wb = load_workbook(filename)

    # Remove old Intuition Guide if present
    if "Intuition Guide" in wb.sheetnames:
        del wb["Intuition Guide"]

    row_count = rebuild_readme(wb, asset, CASE_CONTENT[asset])
    print(f"  README rebuilt with {row_count} rows")

    # Reorder
    order = ['README', 'Case Prompt', 'IC Memo', 'Model']
    for i, name in enumerate(order):
        if name in wb.sheetnames:
            wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

    wb.save(filename)
    print(f"  Sheets: {wb.sheetnames}")
    wb.close()

print("\n" + "=" * 60)
print("README REBUILD COMPLETE")
print("=" * 60)
