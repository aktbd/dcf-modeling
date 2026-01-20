# LOTUS INFRASTRUCTURE PE STUDY GUIDE

> Deep conceptual explanations for energy infrastructure modeling.
> Keep Excel annotations concise — the deep dives live here.

---

## 1. What is an "Energy LBO"?

An infrastructure LBO differs from a traditional corporate LBO:

| Aspect | Corporate LBO | Infrastructure LBO |
|--------|---------------|-------------------|
| Target | Operating company | Single asset or portfolio |
| Revenue | Multiple business lines | Physical operations (power, transport) |
| Debt | Corporate loans | Project finance debt |
| Cash Flow | Volatile, operational | Predictable, contracted |
| Exit | Sale or IPO | Asset sale to strategic/sponsor |

**Key Insight:** Infrastructure deals are "levered equity bets on cash flow." The model answers: "What equity check do I write, and what distributions do I receive?"

---

## 2. The Cash Flow Waterfall

Understanding the flow from revenue to equity distributions:

```
ASSET OPERATIONS
    │
    ├── Capacity (MW) × Availability
    │         │
    │         ▼
    ├── Generation (MWh) = Capacity × CF × 8760
    │         │
    │         ▼
    └── Revenue ($mm) = Generation × Price / 1E6
              │
              ▼
         ┌────────────────┐
         │    EBITDA      │ ← Revenue - OpEx
         └───────┬────────┘
                 │
                 ▼
         ┌────────────────┐
         │    CFADS       │ ← EBITDA - Taxes
         └───────┬────────┘
                 │
    ┌────────────┼────────────┐
    │            │            │
    ▼            ▼            ▼
 Interest   Principal    DSRA Funding
    │            │            │
    └────────────┴────────────┘
                 │
                 ▼
         ┌────────────────┐
         │ Distributable  │ ← If DSCR passes lock-up
         └───────┬────────┘
                 │
                 ▼
         ┌────────────────┐
         │  Equity CF     │ ← Distributions + Exit
         └────────────────┘
```

---

## 3. Power Market Economics

### 3.1 Spark Spread (The Quick Way)

**Formula:**
```
Spark Spread ($/MWh) = Power Price - (Heat Rate × Gas Price / 1000)
```

**Example:**
- Power Price: $50/MWh
- Heat Rate: 7,000 Btu/kWh
- Gas Price: $5/MMBtu

```
Spark Spread = $50 - (7,000 × $5 / 1,000)
             = $50 - $35
             = $15/MWh
```

This is the **gross margin per MWh after fuel cost**. It's what matters for thermal plant profitability.

### 3.2 Market Heat Rate (The Trader's Way)

Traders often think in terms of "Realized Market Heat Rate" (RMHR):

**Formula:**
```
RMHR (Btu/kWh) = Power Price / Gas Price × 1,000
```

**Interpretation:**
- If RMHR > Plant Heat Rate → Plant is "in the money" → Dispatch
- If RMHR < Plant Heat Rate → Plant is "out of the money" → Don't run

**Same Example:**
```
RMHR = $50 / $5 × 1,000 = 10,000 Btu/kWh

Since 10,000 > 7,000 (plant HR), the plant is profitable.
Heat Rate Spread = 10,000 - 7,000 = 3,000 Btu/kWh
Margin = 3,000 × $5 / 1,000 = $15/MWh ← Same answer!
```

**Why Both Methods?**
- Spark Spread is simpler for modeling (use this in Excel)
- RMHR is how traders quote and think (use this in interviews)
- Both give the same answer — just different framing

### 3.3 Carbon Costs (When to Include)

Only include carbon costs if the case mentions:
- California, CARB, or cap-and-trade → ~$35/ton
- Northeast, RGGI → ~$15/ton
- Federal carbon tax (hypothetical)

**Formula:**
```
Carbon Cost ($mm) = Generation × (HR/1000) × (117/2000) × $/ton / 1E6

Where:
- HR/1000 converts Btu/kWh to MMBtu/MWh
- 117 lb CO2/MMBtu is natural gas emission factor
- /2000 converts lb to tons
```

**Rough Rule:** At HR=7,000 and $35/ton → ~$1.50/MWh

---

## 4. Asset Class Profiles

### 4.1 CCGT (Combined Cycle Gas Turbine)

| Parameter | Typical Range |
|-----------|---------------|
| Capacity | 200-600 MW |
| Heat Rate | 6,500-7,500 Btu/kWh |
| Capacity Factor | 45-65% |
| Fixed O&M | $12-18/kW-yr |
| Variable O&M | $2-4/MWh |
| Leverage | 60-70% |
| DSCR Target | 1.25-1.40x |
| Debt Approach | Cash Sweep |

**Revenue:** Energy + Capacity
**Key Driver:** Spark spread (gas-power correlation)
**Risk:** Merchant price exposure, gas basis

### 4.2 Peaker (Simple Cycle Gas Turbine)

| Parameter | Typical Range |
|-----------|---------------|
| Capacity | 100-300 MW |
| Heat Rate | 9,500-11,000 Btu/kWh |
| Dispatch Hours | 500-1,500 hrs/yr |
| Fixed O&M | $8-15/kW-yr |
| Variable O&M | $3-6/MWh |
| Leverage | 55-65% |
| DSCR Target | 1.25-1.35x |
| Debt Approach | Cash Sweep |

**Revenue:** Mostly Capacity (limited energy)
**Key Driver:** Scarcity pricing, capacity auctions
**CRITICAL:** DO NOT use CF × 8760 — use Dispatch Hours!

### 4.3 Solar + BESS

| Parameter | Typical Range |
|-----------|---------------|
| Capacity | 50-200 MW |
| Capacity Factor | 20-30% |
| Degradation | 0.4-0.6%/yr |
| O&M | $10-20/kW-yr |
| ITC | 30-40% of cost |
| Leverage | 65-75% |
| DSCR Target | 1.30-1.45x |
| Debt Approach | Sculpted |

**Revenue:** PPA (contracted) or Merchant
**Key Driver:** PPA price, degradation rate
**Tax:** ITC (investment tax credit) — reduces basis

### 4.4 Wind

| Parameter | Typical Range |
|-----------|---------------|
| Capacity | 50-300 MW |
| Capacity Factor | 30-45% |
| Degradation | 0.2-0.5%/yr |
| O&M | $30-50/kW-yr |
| PTC | $27.50/MWh (10 yrs) |
| Leverage | 60-70% |
| DSCR Target | 1.30-1.40x |
| Debt Approach | Sculpted |

**Revenue:** PPA (contracted) or Merchant
**Key Driver:** Capacity factor, PTC utilization
**Tax:** PTC (production tax credit) — tax CREDIT not revenue!

### 4.5 Transmission

| Parameter | Typical Range |
|-----------|---------------|
| Capacity | 100-500 MW |
| Availability | 97-99% |
| Rate Base | $500mm-$2B |
| ROE | 9-11% (regulated) |
| Leverage | 70-80% |
| DSCR Target | 1.35-1.50x |
| Debt Approach | Sculpted |

**Revenue:** Regulated return on rate base
**Key Driver:** Rate case outcomes, capex efficiency
**Risk:** Regulatory, permitting

### 4.6 Midstream

| Parameter | Typical Range |
|-----------|---------------|
| Volumes | 100-500 mmcf/d |
| Take-or-Pay | 70-90% |
| Growth Phase | Y1-5 |
| Decline Phase | Y6+ |
| Leverage | 60-70% |
| DSCR Target | 1.25-1.35x |
| Debt Approach | Cash Sweep |

**Revenue:** Volume × Fee ($/mcf)
**Key Driver:** Producer activity, basin economics
**Risk:** Volume decline, counterparty credit

---

## 5. Debt Mechanics

### 5.1 Straight-Line Amortization

**When to Use:** Quick models, when told "ignore sweep"

```
Annual Principal = Debt Amount / Tenor
```

**Pros:** Simple, no circularity
**Cons:** Doesn't reflect economic reality

### 5.2 Cash Sweep

**When to Use:** Merchant assets (CCGT, Peaker, Midstream)

```
Excess Cash = CFADS - Scheduled DS - DSRA Funding
Cash Sweep = MIN(Excess × Sweep%, Remaining Balance)
Total Principal = Scheduled + Sweep
```

**Pros:** Realistic for bank debt, de-levers faster in good years
**Cons:** Near-circular (use scheduled_ds trick for DSRA)

### 5.3 Sculpted Debt

**When to Use:** Contracted assets (Solar, Wind, Transmission)

```
Target Debt Service = CFADS / Target DSCR
Sculpted Principal = MAX(0, Target DS - Interest)
Capped Principal = MIN(Sculpted, Beginning Balance)
```

**Pros:** Maximizes debt capacity, matches CF profile
**Cons:** More complex, requires CFADS forecast first

### 5.4 How to Choose

```
Is CFADS predictable and contracted?
    │
    ├── YES → Sculpted Debt (target DSCR, size to CFADS)
    │         Examples: Solar+PPA, Wind+PPA, Transmission
    │
    └── NO → Cash Sweep (mandatory prepayment from excess)
              Examples: Merchant CCGT, Peaker, Midstream
```

### 5.5 PV Debt Sizing (Educational)

In project finance, lenders size debt as:

```
Maximum Debt = NPV(Target Debt Service stream, discount = cost of debt)
```

This ensures the present value of all debt service payments equals the loan amount.

**Example:**
- Target DS = $20mm/yr for 7 years
- Cost of debt = 6%
- NPV = $111.6mm = Maximum supportable debt

In PE/LBO cases, debt is usually GIVEN. But understanding this helps answer "How would you size debt?"

---

## 6. Returns Math

### 6.1 IRR (Internal Rate of Return)

The discount rate that makes NPV = 0.

**Equity Cash Flow Timeline:**
```
Y0: -$63mm (entry equity)
Y1: +$5mm (distribution)
Y2: +$6mm
...
Y7: +$120mm (distribution + exit)
```

**IRR Interpretation:**
- 15% IRR → Beats cost of equity, acceptable
- 20%+ IRR → Strong return
- <12% IRR → May not meet hurdle rate

### 6.2 MOIC (Multiple on Invested Capital)

```
MOIC = Total Cash Inflows / Total Cash Outflows
     = (Sum of Distributions + Exit Proceeds) / Entry Equity
```

**MOIC Interpretation:**
- 1.5x → 50% gain
- 2.0x → Doubled money
- 2.5x+ → Strong multiple

### 6.3 IRR vs MOIC Trade-off

| Scenario | IRR | MOIC | Comment |
|----------|-----|------|---------|
| Quick exit | High | Low | Time value helps IRR |
| Long hold | Low | High | Compounding helps MOIC |
| Optimal | Both | Both | Sweet spot at 5-7 years |

---

## 7. Common Mistakes (Top 10)

1. **Fuel Cost ÷ 1E6 instead of 1E9**
   - Results in fuel cost 1000x too high
   - Fix: Always use /1E9 for fuel

2. **Peaker: CF × 8760**
   - Results in generation 10x too high
   - Fix: Use Dispatch Hours × Availability

3. **Capacity Revenue: Price in $/MW**
   - Results in revenue 1000x too low
   - Fix: Price is $/kW-yr, multiply by 1000

4. **DSRA Target = Current Year DS**
   - Results in DSRA off by one year
   - Fix: Target = NEXT year's debt service

5. **Exit Equity = Exit EV**
   - Overstates equity proceeds
   - Fix: Exit EV - Debt Payoff + DSRA Release

6. **PTC Creates Negative Taxes**
   - Tax credit exceeds tax owed
   - Fix: PTC Applied = MIN(Gross PTC, Pre-PTC Tax)

7. **IRR Range Excludes Y0**
   - Returns #NUM! error
   - Fix: IRR(D:N) must include Y0 negative

8. **DSCR Uses Post-Sweep DS**
   - Understates true coverage
   - Fix: Use Pre-Sweep DS in denominator

9. **Min DSCR Includes Zero Years**
   - Returns 0 after debt paid off
   - Fix: Use MINIFS with balance > 0

10. **Circular Reference in DSRA**
    - Excel can't solve
    - Fix: Use scheduled_ds row (deterministic)

---

## 8. Excel Formula Reference

### 8.1 The Big 5 Functions

**MAX(a, b)**
```
=MAX(0, E50)  → Returns E50 if positive, else 0
Use: Prevent negative values (taxes, distributions)
```

**MIN(a, b)**
```
=MIN(E45, E50)  → Returns smaller of two values
Use: Cap values (PTC at tax, sweep at balance)
```

**IF(test, true, false)**
```
=IF(E$1<=7, "Active", "Inactive")
Use: Conditional logic (contract years, exit year)
```

**IRR(range)**
```
=IRR(D100:N100)  → Returns rate making NPV=0
TRAP: Range MUST include Y0 negative cash flow
```

**NPV(rate, range)**
```
=NPV(D45, E82:K82)  → Present value of future cash flows
Note: NPV excludes Y0, so add Y0 separately if needed
```

### 8.2 Key Formula Patterns

**Escalation:**
```
=Base × (1 + Rate)^(Year - 1)
Example: =$D$25*(1+$D$26)^(E$1-1)
```

**Generation (Baseload):**
```
=Capacity × CF × 8760
Example: =$D$30*$D$31*8760
```

**Generation (Peaker):**
```
=Capacity × Dispatch Hours × Availability
Example: =$D$30*$D$32*$D$33
```

**Fuel Cost:**
```
=Generation × Heat Rate × Gas Price / 1E9
Example: =E45*$D$24*E27/1E9
```

**DSCR:**
```
=IF(DS>0, CFADS/DS, 0)
Example: =IF(E95>0,E82/E95,0)
```

### 8.3 Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| F2 | Edit cell |
| F4 | Toggle absolute reference ($) |
| Ctrl+D | Fill down |
| Ctrl+R | Fill right |
| Ctrl+` | Show/hide formulas |
| Ctrl+[ | Trace precedents |
| Ctrl+Shift+Enter | Array formula |

---

## 9. Four-Hour Time Budget

| Phase | Time | Activities |
|-------|------|------------|
| **Setup** | 15 min | Workbook, headers, freeze panes, sections |
| **Transaction** | 15 min | Entry EV, fees, exit inputs |
| **S&U** | 15 min | Uses first, equity = balancing item |
| **Operating** | 45 min | Revenue, costs, EBITDA |
| **Depreciation** | 15 min | Straight-line, tax rate |
| **Debt** | 45 min | Schedule, DSCR, sweep/sculpt |
| **DSRA** | 15 min | Target, funding, release |
| **Waterfall** | 30 min | Distributable, lock-up |
| **Exit** | 15 min | Exit EV, proceeds |
| **QA** | 15 min | Checks, sanity, errors |
| **Buffer** | 15 min | Fix issues, polish |
| **TOTAL** | **4:00** | |

---

## 10. Last-Minute Checklist

### Before Starting
- [ ] Read entire case prompt
- [ ] Note special instructions
- [ ] Identify asset type and debt approach
- [ ] Plan sections mentally

### During Build
- [ ] Save frequently (Ctrl+S)
- [ ] Check formulas as you go (click through)
- [ ] Watch for #REF, #NAME, #VALUE
- [ ] Keep units consistent

### Before Submitting
- [ ] S&U Balance = PASS
- [ ] Debt Payoff = PASS
- [ ] Min DSCR ≥ 1.25x
- [ ] No Excel errors
- [ ] Sanity check key outputs
- [ ] Return to audit strip

---

*This guide is for study and preparation. Keep Excel annotations concise — reference this document for deep understanding.*
