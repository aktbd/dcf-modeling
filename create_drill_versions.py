#!/usr/bin/env python3
"""
Create Drill versions of all models with Intuition Guide sheets
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import shutil

NAVY_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
LIGHT_BLUE_FILL = PatternFill(start_color="DEEAF6", end_color="DEEAF6", fill_type="solid")
WHITE_BOLD = Font(color="FFFFFF", bold=True)
BLACK_FONT = Font(color="000000")
BLACK_BOLD = Font(color="000000", bold=True)

THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

def add_intuition_guide_ccgt(wb):
    """Add Intuition Guide sheet for CCGT model"""
    ws = wb.create_sheet("Intuition Guide", 0)

    # Column widths
    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 80

    content = [
        ("CCGT MODEL - INTUITION GUIDE", ""),
        ("", ""),
        ("SECTION", "WHY IT MATTERS / KEY LOGIC"),
        ("", ""),
        ("═══ BUSINESS CONTEXT ═══", ""),
        ("What is a CCGT?", "Combined Cycle Gas Turbine - efficient baseload gas plant that uses waste heat from gas turbines to power steam turbines. 'Combined' refers to combining gas and steam cycles for ~60% efficiency vs ~35% for simple cycle."),
        ("Revenue Streams", "1) Energy revenue (selling power at spot/contract prices) - driven by spark spread (power price - fuel cost)\n2) Capacity revenue (payment for being available) - driven by capacity auction prices\nRevenue volatility justifies cash sweep structure to accelerate debt paydown during good years."),
        ("", ""),
        ("═══ INPUTS LOGIC ═══", ""),
        ("Entry EV ($520mm)", "Purchase price. EV = Equity + Debt, represents total asset value regardless of financing."),
        ("Capacity (365 MW)", "Nameplate rating - maximum output. Used to calculate generation and capacity revenue."),
        ("Heat Rate (7,000 Btu/kWh)", "Fuel efficiency - how much gas needed per kWh. Lower = better. 7,000 is efficient CCGT; peakers are 10,000+."),
        ("Capacity Factor (55%)", "Actual generation / max possible generation. 55% is typical merchant CCGT - runs when profitable but not always."),
        ("Forced Outage Rate (5%)", "Unplanned downtime. Reduces UCAP (capacity revenue) because you're paid for reliable capacity, not nameplate."),
        ("Leverage (45%)", "Debt / EV. 45% is moderate for merchant power - balances returns enhancement with covenant headroom."),
        ("Cash Sweep (75%)", "Mandatory prepayment of debt from excess cash flow. Higher sweep = faster paydown = lower refinancing risk."),
        ("", ""),
        ("═══ CALCULATED ITEMS ═══", ""),
        ("UCAP (346.75 MW)", "Unforced Capacity = Nameplate × (1 - FOR). This is what you get paid for in capacity markets - your reliable capacity."),
        ("Generation (1,758,390 MWh)", "Annual output = Capacity × CF × 8,760 hours/year. This drives energy revenue."),
        ("", ""),
        ("═══ OPERATING MODEL ═══", ""),
        ("Price Escalation", "All prices/costs grow at escalation rate (2.5%). Reflects inflation and market dynamics."),
        ("Energy Revenue", "Generation × Power Price / 1,000,000. Converting MWh × $/MWh to $mm."),
        ("Capacity Revenue", "UCAP (MW) × $/kW-yr × 1,000 kW/MW / 1,000,000. Note: capacity price is per kW, not MW."),
        ("Fuel Cost (CRITICAL)", "Generation × Heat Rate × Gas Price / 1,000,000,000.\nDimensional analysis: MWh × Btu/kWh × $/mmBtu / 10^9 = $mm\nThis is the #1 error candidates make - always verify the divisor."),
        ("", ""),
        ("═══ DEBT MECHANICS ═══", ""),
        ("Why Cash Sweep?", "Merchant assets have volatile CFADS. Sweep structure pays down debt faster during good years, providing cushion for bad years. Alternative is sculpting (seen in contracted assets)."),
        ("DSRA Purpose", "6 months of debt service held in reserve. If CFADS drops, DSRA provides liquidity to make debt payments and avoid covenant breach."),
        ("Lock-up Test", "If DSCR < 1.10x, distributions are 'locked up' - all cash goes to debt paydown. Protects lenders."),
        ("", ""),
        ("═══ RETURNS ═══", ""),
        ("IRR vs MOIC", "IRR measures time-weighted return (sensitive to timing). MOIC measures total cash-on-cash return (ignores timing). Both matter."),
        ("Levered vs Unlevered", "Levered IRR = return to equity. Unlevered IRR = return to total capital. Difference is the 'leverage benefit.'"),
        ("Exit at 7.0x EBITDA", "Standard infrastructure exit assumption. Buyer universe: other PE funds, utilities, infrastructure yield vehicles."),
        ("", ""),
        ("═══ COMMON IC QUESTIONS ═══", ""),
        ("Q: Walk me through IRR bridge", "Start with unlevered IRR, add leverage benefit (cheaper debt vs equity), subtract financing costs (interest, fees)."),
        ("Q: What if gas prices spike?", "Higher fuel costs reduce spark spread and EBITDA. Run sensitivity showing EBITDA and DSCR impact."),
        ("Q: Why 45% leverage?", "Balances return enhancement with covenant headroom. Higher leverage = higher returns but less cushion for volatility."),
    ]

    row = 1
    for label, explanation in content:
        if label.startswith("═══"):
            ws.cell(row=row, column=1, value=label).font = WHITE_BOLD
            ws.cell(row=row, column=1).fill = NAVY_FILL
            ws.cell(row=row, column=2).fill = NAVY_FILL
        elif label == "SECTION":
            ws.cell(row=row, column=1, value=label).font = BLACK_BOLD
            ws.cell(row=row, column=2, value=explanation).font = BLACK_BOLD
        elif label.startswith("Q:"):
            ws.cell(row=row, column=1, value=label).font = BLACK_BOLD
            ws.cell(row=row, column=2, value=explanation).font = BLACK_FONT
            ws.cell(row=row, column=1).fill = LIGHT_BLUE_FILL
            ws.cell(row=row, column=2).fill = LIGHT_BLUE_FILL
        else:
            ws.cell(row=row, column=1, value=label).font = BLACK_BOLD
            ws.cell(row=row, column=2, value=explanation).font = BLACK_FONT

        ws.cell(row=row, column=2).alignment = Alignment(wrap_text=True)
        row += 1

    # Set row heights for wrapped text
    for r in range(1, row):
        ws.row_dimensions[r].height = 30

def add_intuition_guide_peaker(wb):
    """Add Intuition Guide sheet for Peaker model"""
    ws = wb.create_sheet("Intuition Guide", 0)

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 80

    content = [
        ("PEAKER MODEL - INTUITION GUIDE", ""),
        ("", ""),
        ("SECTION", "WHY IT MATTERS / KEY LOGIC"),
        ("", ""),
        ("═══ BUSINESS CONTEXT ═══", ""),
        ("What is a Peaker?", "Simple cycle gas turbine - fast-start units that run only during peak demand (~400 hours/year vs 4,800 for CCGT). Less efficient but can start in 10-15 minutes."),
        ("Revenue Mix", "60%+ from capacity payments for being available. Energy revenue is smaller (few running hours) but at high prices (dispatched during peaks)."),
        ("Why Lower Leverage (30%)?", "More volatile revenue profile than CCGT. Capacity revenues more stable, but energy upside is lumpier. Conservative structure provides cushion."),
        ("", ""),
        ("═══ KEY DIFFERENCES FROM CCGT ═══", ""),
        ("Dispatch Hours vs CF", "CRITICAL: Peakers use dispatch hours (400), NOT capacity factor × 8,760. This is the most common modeling error."),
        ("Heat Rate (10,500 Btu/kWh)", "Less efficient than CCGT (7,000). Acceptable because fuel cost is small % of revenue - availability matters more."),
        ("Lower Cash Sweep (50%)", "Need to retain cash for major maintenance. HGP (hot gas path) inspection every 5 years is $4mm+."),
        ("Major Maintenance Y5", "Hot gas path inspection - internal combustion parts require overhaul. Must fund from retained cash, not debt."),
        ("", ""),
        ("═══ GENERATION FORMULA ═══", ""),
        ("Correct Formula", "Generation = Capacity × Dispatch Hours × Availability = 200 MW × 400 hrs × 94% = 75,200 MWh"),
        ("WRONG Formula", "DO NOT use: Capacity × CF × 8,760. That's for baseload plants, not peakers."),
        ("", ""),
        ("═══ MAINTENANCE IMPACT ═══", ""),
        ("Y5 Cash Flow", "CFADS drops by $4mm in Y5 due to HGP inspection. This hits DSCR - make sure covenant still passes."),
        ("Why Not Debt-Fund Maint?", "Maintenance is operational expense, not capital. Lenders expect you to fund from operations. That's why sweep is only 50%."),
        ("", ""),
        ("═══ COMMON IC QUESTIONS ═══", ""),
        ("Q: Why is energy price $95/MWh?", "Dispatch-weighted average. Peakers only run during high-price hours, so their realized price > average market price."),
        ("Q: What's UCAP vs ICAP?", "UCAP = Unforced Capacity (adjusted for reliability). ICAP = Installed Capacity (nameplate). Markets pay for UCAP."),
        ("Q: Downside if peaker never dispatches?", "Still receive capacity payments. Energy revenue is upside, not base case. Worst case = capacity-only revenue."),
    ]

    row = 1
    for label, explanation in content:
        if label.startswith("═══"):
            ws.cell(row=row, column=1, value=label).font = WHITE_BOLD
            ws.cell(row=row, column=1).fill = NAVY_FILL
            ws.cell(row=row, column=2).fill = NAVY_FILL
        elif label == "SECTION":
            ws.cell(row=row, column=1, value=label).font = BLACK_BOLD
            ws.cell(row=row, column=2, value=explanation).font = BLACK_BOLD
        elif label.startswith("Q:"):
            ws.cell(row=row, column=1, value=label).font = BLACK_BOLD
            ws.cell(row=row, column=2, value=explanation).font = BLACK_FONT
            ws.cell(row=row, column=1).fill = LIGHT_BLUE_FILL
            ws.cell(row=row, column=2).fill = LIGHT_BLUE_FILL
        else:
            ws.cell(row=row, column=1, value=label).font = BLACK_BOLD
            ws.cell(row=row, column=2, value=explanation).font = BLACK_FONT

        ws.cell(row=row, column=2).alignment = Alignment(wrap_text=True)
        row += 1

    for r in range(1, row):
        ws.row_dimensions[r].height = 30

def add_intuition_guide_solar_bess(wb):
    """Add Intuition Guide sheet for Solar+BESS model"""
    ws = wb.create_sheet("Intuition Guide", 0)

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 80

    content = [
        ("SOLAR+BESS MODEL - INTUITION GUIDE", ""),
        ("", ""),
        ("SECTION", "WHY IT MATTERS / KEY LOGIC"),
        ("", ""),
        ("═══ BUSINESS CONTEXT ═══", ""),
        ("Why Solar + Battery?", "Solar provides contracted PPA revenue. Battery adds arbitrage (buy low/sell high) and ancillary services (grid stability). Combined improves risk-adjusted returns."),
        ("Revenue Stability", "PPA contract provides revenue certainty. This supports sculpted debt (vs sweep) because cash flows are predictable."),
        ("ITC Impact", "30% Investment Tax Credit reduces equity requirement. Government subsidy makes project economics work."),
        ("", ""),
        ("═══ KEY DIFFERENCES FROM THERMAL ═══", ""),
        ("No Fuel Cost", "Sun is free. No commodity exposure. Gross margin = revenue - O&M only."),
        ("Degradation (0.5%/yr)", "Solar panels lose efficiency over time. Y10 generation is ~95% of Y1. Must model declining output."),
        ("Sculpted vs Sweep Debt", "Contracted cash flows allow sculpting - size debt service to hit target DSCR. More efficient than sweep."),
        ("ITC in Sources & Uses", "ITC is a 'source' that offsets equity. Pre-ITC Equity - ITC = Net Equity (your actual check)."),
        ("", ""),
        ("═══ DEBT SCULPTING LOGIC ═══", ""),
        ("How It Works", "Target DS = CFADS / Target DSCR. Then: Sculpted Principal = Target DS - Interest. Debt service sized to maintain constant coverage."),
        ("Why 1.35x Target?", "Contracted assets get lower DSCR targets (less volatile). 1.35x provides cushion while maximizing debt capacity."),
        ("Benefit", "Efficient structure - pay principal when you have cash, not on fixed schedule. Better matches asset profile."),
        ("", ""),
        ("═══ BESS AUGMENTATION ═══", ""),
        ("Why Y7 Capex?", "Battery capacity degrades over time. Augmentation restores capacity to maintain contracted services. Budget $5mm."),
        ("Impact on CFADS", "Y7 CFADS reduced by augmentation capex. Check that DSCR still passes in that year."),
        ("", ""),
        ("═══ TAX TREATMENT ═══", ""),
        ("MACRS Shield Y1-Y6", "Accelerated depreciation creates tax losses, shielding income. Zero cash taxes early years."),
        ("Post-MACRS (Y7+)", "MACRS exhausted. Now pay taxes, but at reduced effective rate (15% vs 25%) due to remaining depreciation."),
        ("", ""),
        ("═══ COMMON IC QUESTIONS ═══", ""),
        ("Q: What happens if panels degrade faster?", "Lower generation = lower PPA revenue. Run sensitivity on degradation rate showing EBITDA and DSCR impact."),
        ("Q: Why sculpted vs sweep?", "Contracted cash flows are predictable, allowing efficient sculpting. Sweeps are for volatile merchant assets."),
        ("Q: What's BESS basis risk?", "Battery revenue depends on price volatility. Low volatility = less arbitrage opportunity. PPA hedges this."),
    ]

    row = 1
    for label, explanation in content:
        if label.startswith("═══"):
            ws.cell(row=row, column=1, value=label).font = WHITE_BOLD
            ws.cell(row=row, column=1).fill = NAVY_FILL
            ws.cell(row=row, column=2).fill = NAVY_FILL
        elif label == "SECTION":
            ws.cell(row=row, column=1, value=label).font = BLACK_BOLD
            ws.cell(row=row, column=2, value=explanation).font = BLACK_BOLD
        elif label.startswith("Q:"):
            ws.cell(row=row, column=1, value=label).font = BLACK_BOLD
            ws.cell(row=row, column=2, value=explanation).font = BLACK_FONT
            ws.cell(row=row, column=1).fill = LIGHT_BLUE_FILL
            ws.cell(row=row, column=2).fill = LIGHT_BLUE_FILL
        else:
            ws.cell(row=row, column=1, value=label).font = BLACK_BOLD
            ws.cell(row=row, column=2, value=explanation).font = BLACK_FONT

        ws.cell(row=row, column=2).alignment = Alignment(wrap_text=True)
        row += 1

    for r in range(1, row):
        ws.row_dimensions[r].height = 30

def add_intuition_guide_transmission(wb):
    """Add Intuition Guide sheet for Transmission model"""
    ws = wb.create_sheet("Intuition Guide", 0)

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 80

    content = [
        ("TRANSMISSION MODEL - INTUITION GUIDE", ""),
        ("", ""),
        ("SECTION", "WHY IT MATTERS / KEY LOGIC"),
        ("", ""),
        ("═══ BUSINESS CONTEXT ═══", ""),
        ("What is Transmission?", "High-voltage lines connecting generation to load centers. Regulated asset - earns FERC-approved ROE on rate base."),
        ("RAB-Based Returns", "Rate base (RAB) = total invested capital. Tariff set to provide return ON and OF capital. Very stable, predictable."),
        ("Construction Period", "Y0-Y2 is construction with no revenue. Must fund construction, then refinance at COD (Commercial Operation Date)."),
        ("", ""),
        ("═══ CONSTRUCTION FINANCING ═══", ""),
        ("Construction Loan (85%)", "Short-term loan to fund construction. Higher leverage acceptable because it's temporary."),
        ("Interest Capitalization", "Construction interest added to RAB. You're earning on the interest you paid - improves economics."),
        ("Phased Equity", "Equity deployed over multiple years as construction progresses. IRR calculation must account for timing."),
        ("", ""),
        ("═══ RAB CALCULATION ═══", ""),
        ("RAB at COD", "Dev Y0 + Dev Y1 + Construction + Capitalized Interest. This is your 'invested capital' for rate-setting."),
        ("RAB Depreciation", "RAB declines as you recover capital through depreciation. Exit RAB = COD RAB - Cumulative Depreciation."),
        ("Exit at xRAB Multiple", "Transmission trades at premium to RAB (1.15x) because of stability and growth. Exit EV = Exit RAB × 1.15x."),
        ("", ""),
        ("═══ REFINANCING AT COD ═══", ""),
        ("What Happens at COD", "Construction loan + capitalized interest repaid. Replaced with term debt (60% of RAB) + refinancing equity."),
        ("Refi Equity Calculation", "(Construction Loan + Capitalized Interest) - Term Debt = equity needed to bridge the gap."),
        ("", ""),
        ("═══ DEBT STRUCTURE ═══", ""),
        ("Why Sculpted?", "Regulated tariff provides very predictable cash flows. Sculpting is efficient - size debt service to DSCR target."),
        ("Higher DSCR (1.40x)", "Single-asset, long-lived infrastructure. Higher covenant provides cushion for multi-decade asset life."),
        ("No Debt Y1-Y2", "Construction period - no revenue to service debt. Term debt starts at COD (Y3)."),
        ("", ""),
        ("═══ COMMON IC QUESTIONS ═══", ""),
        ("Q: What's the regulatory risk?", "FERC could reduce allowed ROE. Mitigant: historically stable, bipartisan support for transmission investment."),
        ("Q: Why does transmission trade at premium?", "Stable regulated returns, long asset life (40+ years), critical infrastructure status. Scarcity value."),
        ("Q: What if construction is delayed?", "Higher capitalized interest, later revenue start. Run sensitivity on COD delay showing IRR impact."),
    ]

    row = 1
    for label, explanation in content:
        if label.startswith("═══"):
            ws.cell(row=row, column=1, value=label).font = WHITE_BOLD
            ws.cell(row=row, column=1).fill = NAVY_FILL
            ws.cell(row=row, column=2).fill = NAVY_FILL
        elif label == "SECTION":
            ws.cell(row=row, column=1, value=label).font = BLACK_BOLD
            ws.cell(row=row, column=2, value=explanation).font = BLACK_BOLD
        elif label.startswith("Q:"):
            ws.cell(row=row, column=1, value=label).font = BLACK_BOLD
            ws.cell(row=row, column=2, value=explanation).font = BLACK_FONT
            ws.cell(row=row, column=1).fill = LIGHT_BLUE_FILL
            ws.cell(row=row, column=2).fill = LIGHT_BLUE_FILL
        else:
            ws.cell(row=row, column=1, value=label).font = BLACK_BOLD
            ws.cell(row=row, column=2, value=explanation).font = BLACK_FONT

        ws.cell(row=row, column=2).alignment = Alignment(wrap_text=True)
        row += 1

    for r in range(1, row):
        ws.row_dimensions[r].height = 30

def add_intuition_guide_midstream(wb):
    """Add Intuition Guide sheet for Midstream model"""
    ws = wb.create_sheet("Intuition Guide", 0)

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 80

    content = [
        ("MIDSTREAM MODEL - INTUITION GUIDE", ""),
        ("", ""),
        ("SECTION", "WHY IT MATTERS / KEY LOGIC"),
        ("", ""),
        ("═══ BUSINESS CONTEXT ═══", ""),
        ("What is Midstream?", "Gas gathering systems collect production from wellheads and deliver to processing plants/pipelines. Fee-for-service model."),
        ("Fee-for-Service", "Revenue = Volume × Fee ($/Mcf). No commodity exposure - you're paid to move gas, not sell it."),
        ("Volume Risk", "Key risk is producer activity. If drilling slows, volumes decline. Hence the growth-then-decline profile."),
        ("", ""),
        ("═══ VOLUME PROFILE ═══", ""),
        ("Growth Phase (Y1-Y5)", "Active drilling - new wells connected. Volume grows 3%/year as producer develops acreage."),
        ("Decline Phase (Y6-Y10)", "Wells deplete naturally. 5%/year decline as production tails off. No new drilling assumed."),
        ("Y5 Peak Volume", "Y1 Volume × (1.03)^4 = peak. Then Y6 = Peak × (0.95), Y7 = Peak × (0.95)^2, etc."),
        ("", ""),
        ("═══ FINANCING RATIONALE ═══", ""),
        ("Short Tenor (5 years)", "Must pay off debt BEFORE decline phase. Can't service debt with declining cash flows. De-risk early."),
        ("Aggressive Sweep (75%)", "Maximize paydown during growth phase. Every dollar of sweep reduces refinancing risk."),
        ("Lower Leverage (30%)", "Basin/producer risk. More conservative than contracted assets. Single counterparty exposure."),
        ("Higher Rate (7%)", "Basin-specific risk premium. Smaller asset, producer concentration, decline exposure."),
        ("", ""),
        ("═══ GROWTH CAPEX ═══", ""),
        ("Y1-Y5 Only ($2mm/yr)", "Compression, looping to handle growing volumes. Stops in Y6 when growth ends."),
        ("Impact on CFADS", "Capex reduces CFADS and thus available cash for debt paydown. Factor into covenant analysis."),
        ("", ""),
        ("═══ EXIT CONSIDERATIONS ═══", ""),
        ("Compressed Multiple (6.5x)", "Declining asset commands lower multiple than stable/growing assets. Reflects PDP-like value."),
        ("Exit at Y7 (Not Y10)", "Better exit in early decline (Y7) when volumes still meaningful, vs late decline (Y10) with lower EBITDA."),
        ("Buyer Universe", "Producer looking for vertical integration, larger midstream looking for bolt-on, PE doing roll-up."),
        ("", ""),
        ("═══ COMMON IC QUESTIONS ═══", ""),
        ("Q: What if producer goes bankrupt?", "Volume risk - gatherer has no MVC protection typically. Counter: producer needs gatherer to monetize gas."),
        ("Q: Why not longer hold period?", "Exit before steep decline. Y7 EBITDA > Y10 EBITDA. Don't ride declining asset down."),
        ("Q: Commodity exposure?", "Fee-for-service = no direct commodity risk. But indirectly: low gas prices = less drilling = lower volumes."),
    ]

    row = 1
    for label, explanation in content:
        if label.startswith("═══"):
            ws.cell(row=row, column=1, value=label).font = WHITE_BOLD
            ws.cell(row=row, column=1).fill = NAVY_FILL
            ws.cell(row=row, column=2).fill = NAVY_FILL
        elif label == "SECTION":
            ws.cell(row=row, column=1, value=label).font = BLACK_BOLD
            ws.cell(row=row, column=2, value=explanation).font = BLACK_BOLD
        elif label.startswith("Q:"):
            ws.cell(row=row, column=1, value=label).font = BLACK_BOLD
            ws.cell(row=row, column=2, value=explanation).font = BLACK_FONT
            ws.cell(row=row, column=1).fill = LIGHT_BLUE_FILL
            ws.cell(row=row, column=2).fill = LIGHT_BLUE_FILL
        else:
            ws.cell(row=row, column=1, value=label).font = BLACK_BOLD
            ws.cell(row=row, column=2, value=explanation).font = BLACK_FONT

        ws.cell(row=row, column=2).alignment = Alignment(wrap_text=True)
        row += 1

    for r in range(1, row):
        ws.row_dimensions[r].height = 30

def create_drill_versions():
    """Create drill versions of all models"""

    models = [
        ("Model_1_CCGT.xlsx", "Drill_1_CCGT.xlsx", add_intuition_guide_ccgt),
        ("Model_2_Peaker.xlsx", "Drill_2_Peaker.xlsx", add_intuition_guide_peaker),
        ("Model_3_SolarBESS.xlsx", "Drill_3_SolarBESS.xlsx", add_intuition_guide_solar_bess),
        ("Model_4_Transmission.xlsx", "Drill_4_Transmission.xlsx", add_intuition_guide_transmission),
        ("Model_5_Midstream.xlsx", "Drill_5_Midstream.xlsx", add_intuition_guide_midstream),
    ]

    base_path = "/mnt/user-data/outputs/"

    for source, dest, guide_func in models:
        source_path = base_path + source
        dest_path = base_path + dest

        # Copy the original model
        shutil.copy(source_path, dest_path)

        # Open and add intuition guide
        wb = openpyxl.load_workbook(dest_path)
        guide_func(wb)
        wb.save(dest_path)

        print(f"Created drill version: {dest}")

if __name__ == "__main__":
    create_drill_versions()
    print("\nAll drill versions created successfully!")
