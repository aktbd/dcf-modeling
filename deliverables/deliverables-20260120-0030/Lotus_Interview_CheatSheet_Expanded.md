# LOTUS INFRASTRUCTURE PARTNERS
## Infrastructure PE Interview Cheat Sheet - Expanded Edition

---

# 1. FIRM INTELLIGENCE

## About Lotus Infrastructure Partners
- **Formerly:** Starwood Energy Group (rebranded 2023)
- **Founded:** 2005
- **Headquarters:** Greenwich, CT
- **AUM:** $3B+
- **Focus:** North American power and energy infrastructure

## Key Portfolio Assets
| Asset | Type | Location | Capacity |
|-------|------|----------|----------|
| Caithness Long Island | CCGT | New York | 350 MW |
| Ten West Link | Transmission | AZ-CA | 3,200 MW |
| Lavaca | Gas Storage | Texas | 16 Bcf |
| Allium | Renewables Platform | Multi-state | 2+ GW pipeline |

## Investment Thesis
- Target **contracted and regulated** assets with visible cash flows
- Focus on **energy transition** enabling infrastructure
- Prefer **operating assets** over greenfield (de-risks execution)
- Typical hold period: **5-10 years**
- Target returns: **Mid-teens levered IRR**, **1.8-2.5x MOIC**

---

# 2. ASSET CLASS DEEP DIVES

## CCGT (Combined Cycle Gas Turbine)

### What It Does
Efficient gas-fired power plant combining gas and steam turbines. "Combined cycle" = waste heat from gas turbine powers steam turbine, achieving ~60% efficiency vs ~35% for simple cycle.

### Revenue Drivers
- **Energy revenue:** Spark spread = Power price - (Heat rate × Gas price)
- **Capacity revenue:** Payment for being available ($/kW-year from capacity auctions)

### Why It Matters
Bridge fuel for energy transition. Runs when renewables can't (no sun/wind). Lower emissions than coal (50%). Fast-ramping complements intermittent renewables.

### What Makes a Good Asset
- Heat rate <7,500 Btu/kWh (efficiency)
- Location in constrained transmission zone
- Dual fuel capability (gas + oil backup)
- Young age (<15 years)
- Recent major maintenance completed

### What Kills Deals
- Old, inefficient units (>10,000 Btu/kWh)
- Environmental compliance issues
- Single offtaker concentration
- Unfavorable grid position (congestion risk)
- Pending EPA regulations

---

## Peaker (Simple Cycle Gas Turbine)

### What It Does
Fast-start gas turbine that runs only during peak demand (~400 hours/year). Less efficient but can start in 10-15 minutes. "Peakers" because they serve peak load.

### Revenue Drivers
- **Capacity revenue:** 60-70% of total (paid for availability)
- **Energy revenue:** High prices during scarcity hours

### Why It Matters
Essential for grid reliability. Batteries can handle 4-hour peaks; peakers needed for multi-day events, extreme weather, renewable droughts. Complementary to storage.

### What Makes a Good Asset
- Location in capacity-constrained zone (NYISO Zone J, PJM MAAC)
- Fast-start capability (<10 minutes)
- Dual fuel (gas + oil for fuel security)
- Recent hot gas path inspection

### What Kills Deals
- Location in capacity-long market
- Single customer/PPA concentration
- Deferred maintenance
- Environmental non-compliance

---

## Solar + BESS

### What It Does
Utility-scale solar panels + battery energy storage. Solar generates during daylight; batteries provide time-shifting and grid services.

### Revenue Drivers
- **PPA revenue:** Contracted $/MWh for solar output
- **BESS revenue:** Arbitrage (buy low/sell high) + ancillary services (frequency regulation)

### Why It Matters
Lowest LCOE generation. ITC (30%) dramatically improves economics. Batteries solve intermittency, enabling higher renewable penetration.

### What Makes a Good Asset
- Investment-grade offtaker on PPA
- High capacity factor location (Southwest US)
- BESS sized for 4+ hours duration
- ITC qualification secured

### What Kills Deals
- Weak offtaker credit
- Curtailment risk (transmission constraints)
- Panel degradation above warranty
- BESS technology risk (unproven chemistry)

---

## Transmission

### What It Does
High-voltage lines connecting generation to load centers. Regulated asset earning FERC-approved ROE on rate base.

### Revenue Drivers
- **Tariff revenue:** $/kW-year for transfer capacity
- Return ON and OF rate base (RAB)

### Why It Matters
Critical infrastructure bottleneck. New transmission is hard to permit (10+ year timeline). Stable, regulated returns.

### What Makes a Good Asset
- FERC-approved tariff in place
- Connects low-cost generation to high-demand load
- Long remaining asset life (30+ years)
- Minimal permitting/environmental risk

### What Kills Deals
- Permitting delays/uncertainty
- Contested rate cases
- Single utility counterparty
- Technology obsolescence risk

---

## Midstream (Gas Gathering)

### What It Does
Pipeline systems collecting natural gas from wellheads and delivering to processing plants or trunk lines. Fee-for-service model.

### Revenue Drivers
- **Gathering fee:** $/Mcf for moving gas
- **Compression fee:** Additional fee for compression services
- Volume × Fee = Revenue

### Why It Matters
Essential infrastructure for production. No commodity exposure (fee-based). Tied to producer drilling activity.

### What Makes a Good Asset
- Acreage dedication (exclusive rights)
- Active drilling by producer
- Multiple producer diversification
- Minimum volume commitments (MVCs)

### What Kills Deals
- Single producer concentration
- Basin in decline
- No MVC protection
- Producer credit risk

---


---

## Wind Farm

### What It Does
Onshore or offshore wind turbines generating electricity from wind energy. Modern turbines: 2-5 MW each, 80-150m hub height.

### Revenue Drivers
- **PPA revenue:** Contracted $/MWh for electricity output
- **Merchant revenue:** Wholesale market prices (post-PPA)
- **PTC:** Production Tax Credit ($27.50/MWh for 10 years from COD)

### Why It Matters
Mature, proven renewable technology. PTC creates significant tax shield. Complements solar (wind often blows at night). Essential for grid decarbonization.

### What Makes a Good Asset
- Capacity factor >35% (good wind resource)
- Long-term PPA with IG offtaker
- PTC eligibility secured
- Modern turbines (<10 years old)
- Strong O&M track record

### What Kills Deals
- Low CF location (<30%)
- Merchant exposure without hedge
- Aging turbines (gearbox risk)
- Curtailment risk (transmission)
- Weak offtaker credit

### Key Modeling Notes
- **Generation:** Capacity × CF × 8,760 × (1-Degradation)^year
- **PTC:** Tax CREDIT, not revenue. Applied = MIN(Gross PTC, Tax Owed)
- **Degradation:** 0.2-0.5%/yr (blade erosion)
- **O&M:** $30-50/kW-yr (higher than solar)
- **Debt:** Sculpted (contracted revenue)

### PTC Mechanics (CRITICAL)
```
Pre-PTC Tax = MAX(0, EBIT) × Tax Rate
Gross PTC = Generation × $27.50/MWh
PTC Applied = MIN(Gross PTC, Pre-PTC Tax)  ← CANNOT create negative taxes!
Net Tax = Pre-PTC Tax - PTC Applied
```

---

# 3. REGIONAL MARKET CONSIDERATIONS

## PJM (Mid-Atlantic)
- **Structure:** Capacity market (RPM auction)
- **Key Dynamic:** 3-year forward capacity procurement provides visibility
- **Opportunity:** Coal retirements creating capacity value for gas
- **Risk:** Aggressive renewable mandates in PA, NJ
- **Common Question:** "How do you underwrite PJM capacity prices?"
  - *Answer:* Use forward auction results for years locked, consultant forecast for remaining. Stress test with 20% decline scenario.

## ERCOT (Texas)
- **Structure:** Energy-only market (NO capacity payments)
- **Key Dynamic:** Scarcity pricing drives revenues ($9,000/MWh cap)
- **Opportunity:** Highest price volatility = highest upside potential
- **Risk:** No capacity payment floor; relies on scarcity events
- **Common Question:** "Why no capacity market in ERCOT?"
  - *Answer:* Political/philosophical choice - Texas believes energy-only market provides proper price signals. Counter: 2021 winter storm raised reliability concerns.

## CAISO (California)
- **Structure:** Capacity market (RA requirements)
- **Key Dynamic:** Duck curve creates storage opportunity
- **Opportunity:** Aggressive decarbonization = premium for clean capacity
- **Risk:** Retail rate death spiral; increasing NEM solar
- **Common Question:** "Explain the duck curve."
  - *Answer:* Solar peaks midday → net load minimum → steep ramp to evening peak. Creates opportunity for 4-hour storage to capture arbitrage.

## NYISO (New York)
- **Structure:** Capacity market (ICAP/UCAP)
- **Key Dynamic:** Zone J (NYC) structural shortage = premium prices
- **Opportunity:** Constrained transmission into NYC supports local generation
- **Risk:** Aggressive CLCPA targets; Indian Point retirement pressures grid
- **Common Question:** "What's Zone J premium?"
  - *Answer:* Historically 2-3x Rest of State prices due to transmission constraints into NYC.

## MISO (Midwest)
- **Structure:** Voluntary capacity market
- **Key Dynamic:** Wind-rich; transmission constraints to load centers
- **Opportunity:** Transmission projects connecting wind to demand
- **Risk:** Oversupply in some zones; weak capacity prices
- **Common Question:** "Why are MISO capacity prices so low?"
  - *Answer:* Capacity surplus + voluntary market = minimal scarcity. But pockets of shortage (Zone 4) have higher prices.

## SPP (Southwest Power Pool)
- **Structure:** Energy-only (no capacity market)
- **Key Dynamic:** Wind saturated; negative prices common
- **Opportunity:** Transmission to move wind to demand centers
- **Risk:** Extreme oversupply; curtailment risk
- **Common Question:** "How do you underwrite SPP wind?"
  - *Answer:* Contracted PPA essential; merchant exposure = significant curtailment risk.

## ISO-NE (New England)
- **Structure:** Capacity market (FCM)
- **Key Dynamic:** Pipeline-constrained; winter gas/power price spikes
- **Opportunity:** Dual-fuel capability premium during winter
- **Risk:** Mystic cost-of-service ending; capacity uncertainty
- **Common Question:** "What's the Mystic issue?"
  - *Answer:* ISO-NE kept Mystic running under cost-of-service for reliability. Ending creates capacity uncertainty but potential upside for remaining generators.

---

# 4. DEBT MECHANICS

## Cash Sweep vs Sculpted - Decision Tree

```
Is CFADS predictable and contracted?
├── YES → Use SCULPTED debt
│         • Size debt service to target DSCR
│         • More efficient - matches payment to cash flow
│         • Works for: PPAs, regulated assets, long-term contracts
│
└── NO → Use CASH SWEEP debt
          • Mandatory prepayment from excess cash
          • Accelerates paydown during good years
          • Works for: Merchant power, midstream, volatile assets
```

## DSRA Mechanics

**Purpose:** Liquidity reserve to cover debt service during temporary CFADS shortfall

**Sizing:**
- Typically 6 months of next year's debt service
- Formula: `DSRA Target = Next Year DS × (DSRA Months / 12)`

**Funding:**
- Initial: Funded at close from equity (part of Uses)
- Ongoing: Tops up if target increases; releases if target decreases
- Exit: Released at exit (adds to equity proceeds)

**Cash Flow Impact:**
```
DSRA Funding = DSRA Target - DSRA Beginning
• Positive = funding requirement (reduces distributable cash)
• Negative = release (increases distributable cash)
```

## Distribution Waterfall

```
CFADS (Cash Flow Available for Debt Service)
  └── Less: Interest Payment
  └── Less: Scheduled Principal
  └── Less: Cash Sweep (if applicable)
  └── Less: DSRA Funding
  └── Equals: Cash Available for Distribution
      └── Lock-up Test: Is DSCR ≥ Lock-up threshold?
          ├── YES (PASS) → Distribute to equity
          └── NO (LOCKED) → Retain / additional debt paydown
```

## Covenant Structures

| Covenant | Typical Level | Purpose |
|----------|---------------|---------|
| Min DSCR | 1.20-1.40x | Ensure debt coverage |
| Lock-up DSCR | 1.10-1.15x | Protect distributions |
| Debt/EBITDA | 4.0-6.0x | Limit leverage |
| Max Capex | Varies | Control cash outflow |

---

# 5. KEY FORMULAS (With Dimensional Analysis)

## Generation Calculations

**CCGT Annual Generation:**
```
Generation (MWh) = Capacity (MW) × Capacity Factor × 8,760 hours/year

Example: 365 MW × 55% × 8,760 = 1,758,390 MWh
```

**Peaker Annual Generation:**
```
Generation (MWh) = Capacity (MW) × Dispatch Hours × Availability

Example: 200 MW × 400 hours × 94% = 75,200 MWh

⚠️ DO NOT use Capacity Factor × 8,760 for peakers!
```

**UCAP (Unforced Capacity):**
```
UCAP (MW) = Capacity (MW) × (1 - Forced Outage Rate)

Example: 365 MW × (1 - 5%) = 346.75 MW
```

## Revenue Calculations

**Energy Revenue:**
```
Revenue ($mm) = Generation (MWh) × Price ($/MWh) / 1,000,000

Example: 1,758,390 MWh × $48/MWh / 1,000,000 = $84.4mm
```

**Capacity Revenue:**
```
Revenue ($mm) = UCAP (MW) × Capacity Price ($/kW-yr) × 1,000 / 1,000,000

Example: 346.75 MW × $110/kW-yr × 1,000 / 1,000,000 = $38.1mm

Note: Multiply by 1,000 because price is per kW, not MW
```

## Fuel Cost (CRITICAL)

```
Fuel Cost ($mm) = Generation (MWh) × Heat Rate (Btu/kWh) × Gas Price ($/mmBtu) / 1,000,000,000

Example: 1,758,390 MWh × 7,000 Btu/kWh × $3.00/mmBtu / 1,000,000,000 = $36.9mm

Dimensional check:
MWh × (Btu/kWh) × ($/mmBtu) / 10^9 = $ × 10^6 / 10^9 = $mm ✓
```

**Quick Check Formula:**
```
Variable Cost ($/MWh) ≈ Heat Rate × Gas Price / 1,000
Example: 7,000 × $3.00 / 1,000 = $21/MWh
```

## Debt Service Calculations

**Scheduled Principal:**
```
Scheduled Principal ($mm/yr) = Debt Amount / Tenor

Example: $234mm / 7 years = $33.4mm/year
```

**Interest:**
```
Interest ($mm) = Beginning Balance × Interest Rate

Example: $234mm × 6.5% = $15.2mm
```

**DSCR:**
```
DSCR = CFADS / Debt Service

Example: $55mm / $40mm = 1.38x
```

## Return Calculations

**Levered IRR:**
```
=IRR(Equity Cash Flow Range)

Equity CF: Y0 = -Equity, Y1-Y6 = Distributions, Y7 = Distribution + Exit Proceeds
```

**MOIC:**
```
MOIC = Sum of Positive Cash Flows / Sum of Negative Cash Flows (absolute value)

Or: =SUMIF(range,">0") / ABS(SUMIF(range,"<0"))
```

---

# 6. BUILD SEQUENCE (4-Hour Interview)

## Hour 1 (Minutes 0-60): Setup & Inputs
- [ ] Read case materials carefully (10 min)
- [ ] Set up Excel structure (5 min)
- [ ] Input all assumptions (15 min)
- [ ] Build calculated items (10 min)
- [ ] Start operating model (20 min)

## Hour 2 (Minutes 60-120): Core Model
- [ ] Complete operating model (20 min)
- [ ] Build cash flow section (15 min)
- [ ] Build debt schedule (25 min)

## Hour 3 (Minutes 120-180): Returns & Review
- [ ] Complete DSRA (15 min)
- [ ] Build distribution waterfall (10 min)
- [ ] Build Sources & Uses (10 min)
- [ ] Build returns section (15 min)
- [ ] Run sanity checks (10 min)

## Hour 4 (Minutes 180-240): Polish & Memo
- [ ] Fix any errors (15 min)
- [ ] Build sensitivities if time (15 min)
- [ ] Draft IC memo (20 min)
- [ ] Final review (10 min)

---

# 7. BRAIN TEASERS - DEAL SOURCING

**Q: How would you source a deal in this market?**
A: Multiple channels: (1) Banker relationships for marketed processes, (2) Direct outreach to known asset owners, (3) Developer partnerships for early-stage opportunities, (4) Distressed situations via restructuring advisors. Key is maintaining relationships to see deals early.

**Q: Auction vs proprietary - which is better?**
A: Proprietary preferred - less competition, better pricing, more time for diligence. But reality: most quality assets run processes. Our edge: (1) Speed to close, (2) Certainty of funding, (3) Sector expertise reduces perceived risk.

**Q: What are your kill criteria at initial screen?**
A: Immediate passes: (1) Environmental contamination, (2) Pending adverse regulation, (3) Technology risk beyond comfort zone, (4) Counterparty credit below threshold, (5) Returns below hurdle even at base case.

**Q: How do you assess seller motivations?**
A: Key questions: Why selling now? What's the funding status? Are they a strategic or financial owner? Distressed vs opportunistic sale? Understanding motivation helps structure bid and anticipate negotiation points.

**Q: What makes an asset "core" vs "opportunistic"?**
A: Core: Contracted/regulated, investment-grade counterparties, stable cash flows, 8-12% unlevered returns. Opportunistic: Merchant exposure, smaller scale, higher returns (15%+) with more risk. Portfolio should balance both.

---

# 8. BRAIN TEASERS - UNDERWRITING

**Q: How do you determine the right entry multiple?**
A: Triangulate: (1) Comparable transactions in asset class, (2) Replacement cost / NAV analysis, (3) DCF output at target returns, (4) Public market comps (if applicable). Adjust for asset-specific factors: age, location, contract coverage.

**Q: Your MOIC is 2.0x but IRR is 12% - explain the disconnect.**
A: Long hold period. MOIC ignores timing; IRR accounts for time value. 2.0x over 10 years = ~7% annualized; 2.0x over 5 years = ~15% annualized. Both matter: MOIC for absolute return, IRR for return per unit time.

**Q: Why would you accept lower IRR for contracted vs merchant?**
A: Risk-adjusted basis. Contracted provides certainty; merchant provides optionality but downside risk. Lower contracted IRR may have better risk-adjusted return (Sharpe ratio perspective). Also: lender appetite differs significantly.

**Q: How do you stress test this investment?**
A: Key drivers: (1) Commodity prices ±20%, (2) Volume/capacity factor ±10%, (3) Capex overrun +15%, (4) Exit multiple compression 0.5-1.0x. Show IRR impact of each; combination downside scenario.

**Q: What's your conviction level on this deal?**
A: [Framework] High conviction if: (1) Clear value drivers within control, (2) Downside protected by contract/regulation, (3) Multiple exit paths, (4) Favorable competitive dynamics. Medium if one factor weak; Low if multiple concerns.

---

# 9. BRAIN TEASERS - MARKET/TECHNICAL

**Q: Why doesn't ERCOT have a capacity market?**
A: Texas philosophy: energy-only market provides proper price signals via scarcity pricing. $9,000/MWh cap supposed to incentivize new build. Counter-argument: Winter Storm Uri showed reliability risk. Debate ongoing about market reform.

**Q: Explain the duck curve and its investment implications.**
A: Midday solar surplus → net load minimum → steep ramp to evening peak. Shape looks like duck. Implications: (1) Midday prices can go negative, (2) 4-hour storage captures arbitrage, (3) Flexible generation (gas) needed for ramp, (4) Solar+storage paired investments.

**Q: What's ELCC and why does it matter?**
A: Effective Load Carrying Capability - measures how much a resource contributes to reliability. Solar ELCC declining as more solar added (all produces at same time). Matters for capacity value - lower ELCC = less capacity revenue. BESS helps improve ELCC.

**Q: How does carbon regulation affect different asset classes?**
A: Gas: Potentially helpful - carbon cost raises power prices, improving spark spread vs zero-marginal-cost renewables. Coal: Negative - accelerates retirement. Renewables: Positive - value of clean attribute increases. Transmission: Positive - needed for renewable integration.

**Q: What's basis risk in power markets?**
A: Difference between hub price and node price due to transmission congestion. Generator at congested node receives lower price than hub. Matters for hedging - most hedges reference hub, but physical settlement at node. Locational basis can add/subtract $5-15/MWh.

---

# 10. COMMON MISTAKES & DAY-OF CHECKLIST

## Common Modeling Mistakes

| Mistake | Correct Approach |
|---------|------------------|
| Fuel cost divisor wrong | Divide by 1,000,000,000 (not 1,000,000) |
| Peaker uses CF × 8,760 | Use Dispatch Hours × Availability |
| Capacity price per MW | Price is per kW - multiply by 1,000 |
| DSRA uses current year DS | DSRA target = NEXT year's debt service |
| IRR includes wrong years | IRR range must include Y0 (negative equity) |
| Exit equity = Exit EV | Exit equity = Exit EV - Debt + DSRA release |
| Forgetting DSRA in S&U | DSRA is a Use and a Source at exit |

## Day-Of Checklist

### Before Starting
- [ ] Read ALL materials before touching Excel
- [ ] Identify asset type and key drivers
- [ ] Note any unusual features (construction, ITC, degradation)
- [ ] Plan your build order

### During Build
- [ ] Track row numbers in a dictionary/notes
- [ ] Use consistent cell reference style ($D$ for inputs)
- [ ] Build one section, verify, then move on
- [ ] Check dimensional analysis on any conversion

### Before Submitting
- [ ] S&U balances to zero (or PASS check)
- [ ] DSCR > 1.0x all years (ideally > 1.25x)
- [ ] Debt pays off by tenor end
- [ ] IRR is reasonable (12-25% for infra)
- [ ] MOIC makes sense given IRR and hold period
- [ ] Click through 5+ formula cells to verify references
- [ ] Run a quick sensitivity (change one input, verify cascade)

---

## QUICK REFERENCE CARD

### Standard Assumptions
| Input | Typical Range |
|-------|---------------|
| Leverage | 30-60% |
| Interest Rate | 5-8% |
| DSRA | 6 months |
| Tax Rate | 25% |
| Depreciation | 15-20 years SL |
| Exit Multiple | 6-8x EBITDA |

### Red Flag Outputs
| Metric | Concern If |
|--------|------------|
| DSCR | < 1.25x |
| IRR | < 10% or > 35% |
| MOIC | < 1.5x |
| Debt payoff | > 50% remaining at exit |

### Unit Conversions
- 1 MW = 1,000 kW
- 1 MWh = 1,000 kWh
- 1 mmBtu = 1,000,000 Btu
- 1 Bcf = 1,000 MMcf = 1,000,000 Mcf
- 8,760 hours/year

---

*Good luck! Remember: accuracy > speed. A working model with simple structure beats a complex model with errors.*


# MARGIN ANALYSIS (Thermal Plants)

## Spark Spread (Quick Method)
```
Spark Spread ($/MWh) = Power Price - (Heat Rate × Gas Price / 1,000)
```

**Example:** Power=$50, Gas=$5, HR=7,000
```
Spark = $50 - (7,000 × $5 / 1,000) = $50 - $35 = $15/MWh
```

## Market Heat Rate (Trader Method)
```
RMHR (Btu/kWh) = Power Price / Gas Price × 1,000
```

**Interpretation:**
- RMHR > Plant HR → "In the money" → Dispatch
- RMHR < Plant HR → "Out of money" → Don't run

**Same Example:**
```
RMHR = $50 / $5 × 1,000 = 10,000 Btu/kWh
Since 10,000 > 7,000 → Profitable
Margin = (10,000 - 7,000) × $5 / 1,000 = $15/MWh ✓
```

**Key Insight:** Same answer, different framing. Use Spark Spread in model; use RMHR in interview discussions.

---



# PROXY ASSUMPTIONS REFERENCE

## Thermal Power Plants
| Parameter | CCGT | Peaker |
|-----------|------|--------|
| Heat Rate | 6,500-7,500 | 9,500-11,000 | Btu/kWh |
| Capacity Factor | 45-65% | 5-15% (dispatch hrs) |
| Forced Outage Rate | 4-6% | 5-8% |
| Fixed O&M | $12-18/kW-yr | $8-15/kW-yr |
| Variable O&M | $2-4/MWh | $3-6/MWh |

## Renewables
| Parameter | Solar | Wind |
|-----------|-------|------|
| Capacity Factor | 20-30% | 30-45% |
| Degradation | 0.4-0.6%/yr | 0.2-0.5%/yr |
| O&M | $10-20/kW-yr | $30-50/kW-yr |
| PPA Price | $25-45/MWh | $30-50/MWh |

## Power Prices ($/MWh)
| Market | Range | Notes |
|--------|-------|-------|
| PJM | $35-50 | Capacity market |
| ERCOT | $30-45 | Energy-only, volatile |
| NYISO | $40-70 | Zone J premium |
| CAISO | $45-75 | + carbon (~$35/ton) |
| SPP/MISO | $25-40 | Wind-rich |

## Capacity Prices ($/kW-yr)
| Market | Range | Notes |
|--------|-------|-------|
| PJM | $50-150 | Varies by zone |
| NYISO Zone J | $100-200 | Scarcity premium |
| ISO-NE | $40-100 | FCM auction |
| ERCOT | $0 | Energy-only market |

## Gas Prices
| Index | Range |
|-------|-------|
| Henry Hub | $2.50-4.00/MMBtu |
| + Basis (regional) | $0.25-1.00 |

## Financing Defaults
| Parameter | Range | Notes |
|-----------|-------|-------|
| Leverage | 55-70% | Contracted higher |
| Interest Rate | 5.5-7.0% | SOFR + 250-350bp |
| Tenor | 5-7 years | Matches hold |
| DSCR Target | 1.25-1.35x | Merchant |
| | 1.35-1.45x | Contracted |
| DSRA | 6 months DS | |
| Cash Sweep | 50-75% | |

## Transaction
| Parameter | Default |
|-----------|---------|
| Fees | 1.0-2.0% of EV |
| Exit Multiple | Same or +0.5x vs entry |
| Hold Period | 5-7 years |
| Tax Rate | 25% |

---
