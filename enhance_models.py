#!/usr/bin/env python3
"""
Enhance all models with:
1. Column C formula logic with hover comments
2. IC Memo sheet (Models only)
3. Quick Reference sheet (Master_Template only)
"""

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import os

# ============================================================================
# COMMENT DICTIONARIES BY ASSET TYPE
# ============================================================================

COMMON_COMMENTS = {
    'entry ev': {
        'text': 'Purchase price for asset',
        'comment': 'Enterprise Value = Equity + Debt. Represents total asset value regardless of financing structure. Key input for returns analysis.'
    },
    'transaction fees': {
        'text': '% of EV for legal, advisory',
        'comment': 'Typical 1-2% of EV for legal, accounting, advisory fees. Funded at close, increases total uses.'
    },
    'exit multiple': {
        'text': 'Exit EV / Exit EBITDA',
        'comment': 'Assumed sale multiple at exit. Conservative approach: use entry multiple. Upside case: multiple expansion from derisking.'
    },
    'exit year': {
        'text': 'Hold period in years',
        'comment': 'Typical PE hold is 5-7 years. Often aligned with debt tenor to avoid refinancing risk.'
    },
    'leverage': {
        'text': 'Debt / EV',
        'comment': 'Higher leverage amplifies equity returns but increases risk. Merchant assets: 40-50%. Contracted: 60-70%.'
    },
    'interest rate': {
        'text': 'All-in cost of debt',
        'comment': 'Includes spread over benchmark (SOFR). Investment grade infra: 150-250 bps spread. Sub-IG: 300-500 bps.'
    },
    'debt tenor': {
        'text': 'Years to maturity',
        'comment': 'Shorter tenor = faster amortization = lower refinancing risk but higher annual debt service.'
    },
    'dsra months': {
        'text': 'Months of DS in reserve',
        'comment': '6 months is market standard. Provides lender runway if borrower misses payment. Funded at close.'
    },
    'sweep': {
        'text': '% of excess cash to prepay',
        'comment': 'Mandatory prepayment from excess cash. Lenders require this for merchant assets to accelerate deleveraging.'
    },
    'lock-up dscr': {
        'text': 'Min DSCR to distribute',
        'comment': 'If DSCR falls below this level, cash is trapped and cannot be distributed to equity. Typically 1.10-1.15x.'
    },
    'tax rate': {
        'text': 'Effective tax rate',
        'comment': 'US federal ~21% + state. Use blended effective rate. Shields reduce taxable income.'
    },
    'depreciable basis': {
        'text': '% of EV that depreciates',
        'comment': 'Excludes land (~15% of EV). Depreciation creates tax shield, improves after-tax cash flow.'
    },
    'ebitda': {
        'text': '= Revenue − OpEx',
        'comment': 'Earnings Before Interest, Taxes, Depreciation, Amortization. Cash earnings available to service debt and pay equity.'
    },
    'cfads': {
        'text': '= EBITDA − Taxes',
        'comment': 'Cash Flow Available for Debt Service. What remains after taxes to pay interest and principal. DSCR = CFADS ÷ DS.'
    },
    'interest': {
        'text': '= Beg Bal × Rate',
        'comment': 'Interest accrues on outstanding balance. As principal amortizes, interest drops — the deleveraging benefit to equity.'
    },
    'beginning balance': {
        'text': '= Prior Yr End Bal',
        'comment': 'Debt outstanding at start of year. Decreases as principal is paid or swept.'
    },
    'ending balance': {
        'text': '= MAX(0, Beg − Principal)',
        'comment': 'Debt remaining after payments. Should reach zero by tenor end. MAX(0,...) prevents negative balance.'
    },
    'dscr': {
        'text': '= CFADS / Debt Service',
        'comment': 'Debt Service Coverage Ratio. Key lender metric. <1.0x means insufficient cash to pay debt. Target: 1.25-1.40x.'
    },
    'dsra target': {
        'text': '= Next Yr Sched DS × Mo/12',
        'comment': 'Reserve sized to cover NEXT years debt service. Uses SCHEDULED DS (not actual) to avoid circular reference. Released at exit.'
    },
    'dsra funding': {
        'text': '= Target − Beginning',
        'comment': 'Positive = funding requirement (reduces distributable). Negative = release (increases distributable).'
    },
    'distributable': {
        'text': '= IF(DSCR≥Lock-up, Cash, 0)',
        'comment': 'Cash available to equity after debt service and DSRA. Trapped if DSCR below lock-up threshold.'
    },
    'levered irr': {
        'text': '= IRR(Equity CF)',
        'comment': 'Return to equity after debt service. Should exceed cost of equity (12-15% for infra). Leverage amplifies returns.'
    },
    'moic': {
        'text': '= Total In / Total Out',
        'comment': 'Multiple on Invested Capital. Cash-on-cash return. 2.0x = doubled money. Less sensitive to timing than IRR.'
    },
    'unlevered irr': {
        'text': '= IRR(Asset CF)',
        'comment': 'Return on total capital (as if no debt). Measures asset quality independent of financing. Compare to WACC.'
    },
}

CCGT_COMMENTS = {
    'capacity': {
        'text': 'Nameplate MW',
        'comment': 'Maximum output rating. Used to calculate generation and capacity revenue. Actual dispatch varies with economics.'
    },
    'heat rate': {
        'text': 'Btu/kWh efficiency',
        'comment': 'Measures thermal efficiency. 7,000 Btu/kWh is efficient modern CCGT. Higher HR = less efficient = higher fuel cost per MWh.'
    },
    'capacity factor': {
        'text': 'Actual / Max generation',
        'comment': 'Expected dispatch rate. 55% typical for mid-merit CCGT. Driven by spark spread economics — runs when profitable.'
    },
    'forced outage': {
        'text': 'Unplanned downtime %',
        'comment': 'Reduces UCAP (capacity revenue). Lenders/buyers discount nameplate by this factor. 5% is typical well-maintained plant.'
    },
    'ucap': {
        'text': '= Capacity × (1 − FOR)',
        'comment': 'Unforced Capacity — what you can sell into capacity market. FOR reflects forced outages outside your control.'
    },
    'generation': {
        'text': '= Capacity × CF × 8760',
        'comment': 'Annual MWh output. CF reflects expected dispatch based on merit order. 8760 = hours/year. 55% CF = ~4800 hrs/yr.'
    },
    'energy revenue': {
        'text': '= Gen × Price / 1E6',
        'comment': 'Merchant energy revenue from selling MWh into wholesale market. Divide by 1E6 to convert $ to $mm.'
    },
    'capacity revenue': {
        'text': '= UCAP × Price × 1000 / 1E6',
        'comment': 'Payment for being available. Price is $/kW-yr, so multiply by 1000 to convert MW to kW.'
    },
    'fuel cost': {
        'text': '= Gen × HR × Gas / 1E9',
        'comment': 'Largest variable cost for thermal plants. HR measures efficiency. Divide by 1E9: MWh × Btu/kWh × $/mmBtu → $mm.'
    },
    'spark spread': {
        'text': '= Power Price − Fuel/MWh',
        'comment': 'When positive, plant is in the money and dispatches. Drives profitability. Negative spread = plant stays offline.'
    },
    'cash sweep': {
        'text': '= MIN(Excess×75%, Bal−Sched)',
        'comment': 'Mandatory prepayment from excess cash. MIN ensures you dont sweep more than remaining balance after scheduled principal.'
    },
    'scheduled ds': {
        'text': '= Straight-line amort + Int',
        'comment': 'Debt service using straight-line principal only (no sweep). Used for DSRA target to avoid circular reference.'
    },
}

PEAKER_COMMENTS = {
    'capacity': {
        'text': 'Nameplate MW',
        'comment': 'Maximum output. Peakers are smaller than baseload but valuable for reliability during peak demand.'
    },
    'dispatch hours': {
        'text': 'Hours run per year',
        'comment': 'Unlike baseload, peakers run only during high-price hours. 400-800 hrs/yr typical. NOT capacity factor × 8760.'
    },
    'availability': {
        'text': '% of time plant can run',
        'comment': 'Availability for dispatch when called. Higher than CF because plant can run but chooses not to when uneconomic.'
    },
    'generation': {
        'text': '= Capacity × Dispatch × Avail',
        'comment': 'CRITICAL: Use dispatch hours, NOT CF × 8760. Peakers run few hours but earn most revenue from capacity payments.'
    },
    'capacity revenue': {
        'text': '= UCAP × Price × 1000 / 1E6',
        'comment': 'Peakers earn most revenue from capacity payments, not energy. Market values reliability during peak demand.'
    },
    'fuel cost': {
        'text': '= Gen × HR × Gas / 1E9',
        'comment': 'Peaker HR (10,000+) is worse than CCGT (7,000) because simple cycle. But low dispatch hours limit total fuel spend.'
    },
    'maintenance': {
        'text': 'One-time major overhaul',
        'comment': 'Each start-stop cycle causes wear. Major maintenance every 5-7 years. Factor into cash flow timing.'
    },
}

SOLAR_BESS_COMMENTS = {
    'capacity': {
        'text': 'DC nameplate MW',
        'comment': 'DC rating at module level. AC output is lower due to inverter losses. Use DC for generation calc.'
    },
    'capacity factor': {
        'text': 'Actual / Max generation',
        'comment': 'Solar CF driven by irradiance (location) and tracking. 25-30% typical US utility scale.'
    },
    'degradation': {
        'text': 'Annual output loss %',
        'comment': 'Panels lose ~0.5%/yr efficiency. Year 10 output is ~95% of Year 1. Must model declining generation.'
    },
    'generation': {
        'text': '= Y1 Gen × (1−Deg)^(n−1)',
        'comment': 'Generation declines each year due to panel degradation. Compound effect over 20+ year life.'
    },
    'ppa price': {
        'text': '$/MWh contracted price',
        'comment': 'Power Purchase Agreement locks in price for 15-25 years. Removes merchant risk, enables sculpted debt.'
    },
    'ppa revenue': {
        'text': '= Gen × PPA / 1E6',
        'comment': 'Contracted revenue. PPA provides predictable cash flow that supports higher leverage via sculpted debt.'
    },
    'itc': {
        'text': 'Investment Tax Credit %',
        'comment': '30% ITC reduces equity basis at close. Net Equity = Gross − (EV × ITC%). Significantly improves IRR.'
    },
    'bess augmentation': {
        'text': 'Battery replacement capex',
        'comment': 'Battery capacity degrades faster than solar (~2-3%/yr). Augmentation in Y5-7 restores capacity to meet PPA.'
    },
    'sculpted principal': {
        'text': '= MIN(MAX(0, DS−Int), Bal)',
        'comment': 'Principal sized to hit target DSCR given CFADS. No Goal Seek needed — formula calculates directly.'
    },
    'target dscr': {
        'text': 'DSCR for debt sizing',
        'comment': 'Contracted cash flow allows sizing debt to target DSCR. More efficient than sweep — maximizes leverage.'
    },
}

TRANSMISSION_COMMENTS = {
    'transfer capacity': {
        'text': 'MW transfer rating',
        'comment': 'Maximum power that can flow through line. Determines tariff revenue base.'
    },
    'availability': {
        'text': '% of time line available',
        'comment': 'Transmission availability typically 97-99%. Revenue tied to availability, not actual flow.'
    },
    'tariff': {
        'text': '$/kW-yr regulated rate',
        'comment': 'Regulated return on RAB. Predictable but subject to regulatory reset risk. Often inflation-linked.'
    },
    'construction period': {
        'text': 'Years before COD',
        'comment': 'No revenue during build (Y0-Y2). Interest capitalizes into RAB. Equity deployed in phases.'
    },
    'construction capex': {
        'text': 'Total build cost',
        'comment': 'Spent over construction period. Adds to RAB. No revenue until Commercial Operation Date (COD).'
    },
    'rab': {
        'text': 'Regulated Asset Base',
        'comment': 'Exit value = RAB × multiple, not EBITDA × multiple. RAB grows with capex, shrinks with depreciation.'
    },
    'capitalized interest': {
        'text': 'Interest during construction',
        'comment': 'Interest accrued before revenue starts. Adds to RAB, funded by debt. Rolled into opening balance at COD.'
    },
}

MIDSTREAM_COMMENTS = {
    'volume y1': {
        'text': 'Starting throughput Mcf/d',
        'comment': 'Daily volume at project start. Drives fee revenue. Check MVC for downside protection.'
    },
    'growth rate': {
        'text': 'Annual volume growth %',
        'comment': 'Volume growth as new wells come online. Typical growth phase Y1-5 as basin develops.'
    },
    'decline rate': {
        'text': 'Annual volume decline %',
        'comment': 'Volume decline as reservoir depletes. 5%/yr typical for mature basin. Accelerates in late life.'
    },
    'volume': {
        'text': '= Growth then Decline',
        'comment': 'Growth phase (Y1-5) as wells connect, then decline (Y6+) as reservoir depletes. Model inflection carefully.'
    },
    'fee': {
        'text': '$/Mcf gathering fee',
        'comment': 'Fee-for-service revenue. Predictable but volume-dependent. Check MVC terms.'
    },
    'mvc': {
        'text': 'Minimum Volume Commitment',
        'comment': 'Shipper pays for contracted volume even if actual is lower. Provides downside protection during decline.'
    },
    'fee revenue': {
        'text': '= MAX(Vol, MVC) × Fee',
        'comment': 'Revenue is greater of actual volume or MVC times fee. MVC protects during decline phase.'
    },
    'exit timing': {
        'text': 'Exit before steep decline',
        'comment': 'Exit multiple compresses with declining volumes. Sell before decline accelerates to preserve value.'
    },
}

def get_comments_for_asset(asset_type):
    """Return combined common + asset-specific comments"""
    comments = COMMON_COMMENTS.copy()

    if 'ccgt' in asset_type.lower():
        comments.update(CCGT_COMMENTS)
    elif 'peaker' in asset_type.lower():
        comments.update(PEAKER_COMMENTS)
    elif 'solar' in asset_type.lower() or 'bess' in asset_type.lower():
        comments.update(SOLAR_BESS_COMMENTS)
    elif 'transmission' in asset_type.lower():
        comments.update(TRANSMISSION_COMMENTS)
    elif 'midstream' in asset_type.lower():
        comments.update(MIDSTREAM_COMMENTS)

    return comments

def add_formula_column_with_comments(filepath, asset_type):
    """Add Column C formula logic with hover comments"""
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    comments = get_comments_for_asset(asset_type)

    # Ensure Column C header
    if ws['C1'].value != "Notes":
        # Don't overwrite existing notes column
        pass

    # Process each row
    for row in range(2, 160):
        label_cell = ws.cell(row=row, column=1)
        if not label_cell.value:
            continue

        label = str(label_cell.value).lower().strip()

        # Find matching comment
        for key, content in comments.items():
            if key in label:
                # Add/update Column C text if not already populated with something meaningful
                c_cell = ws.cell(row=row, column=3)
                current_val = c_cell.value

                # Only update if empty or very short
                if not current_val or len(str(current_val)) < 10:
                    c_cell.value = content['text']

                # Add comment box
                comment = Comment(content['comment'], 'Lotus Infrastructure')
                comment.width = 350
                comment.height = 80
                c_cell.comment = comment
                break

    wb.save(filepath)
    print(f"  Added comments to: {filepath}")

def add_ic_memo_sheet(filepath, asset_type):
    """Add IC Memo sheet to Model files"""
    wb = openpyxl.load_workbook(filepath)

    # Check if sheet already exists
    if "IC Memo" in wb.sheetnames:
        del wb["IC Memo"]

    ws = wb.create_sheet("IC Memo")

    # Formatting
    ws.column_dimensions['A'].width = 90
    header_font = Font(bold=True, size=14)
    section_font = Font(bold=True, size=11)
    body_font = Font(size=10)

    # Asset-specific content
    memos = {
        'ccgt': {
            'title': 'Greenfield CCGT Power Station',
            'summary': [
                'Acquire 365 MW combined-cycle gas turbine at $520mm ($1,425/kW)',
                'Mid-merit dispatch profile with 55% capacity factor',
                'Dual revenue streams: energy ($45/MWh) + capacity ($75/kW-yr)',
                'Target returns: 12-14% levered IRR, 1.8-2.0x MOIC'
            ],
            'thesis': [
                ('Efficient heat rate provides dispatch advantage', 'At 7,000 Btu/kWh, plant dispatches ahead of older, less efficient fleet. Structural cost advantage in merit order.'),
                ('Capacity revenue provides downside protection', '40% of revenue from capacity payments regardless of dispatch. Reduces merchant exposure.'),
                ('Cash sweep structure accelerates deleveraging', '75% mandatory sweep rapidly pays down debt, reducing refinancing risk and building equity cushion.')
            ],
            'risks': [
                ('Gas price volatility', 'Hedge via heat rate call options; PPA negotiations'),
                ('Regulatory changes', 'Monitor capacity market reforms; diversify across ISOs'),
                ('Execution risk', 'Experienced operator; performance guarantees')
            ],
            'metrics': {'EV': '$520mm', 'Leverage': '45%', 'DSCR': '1.29x', 'IRR': '12.0%', 'MOIC': '1.75x'}
        },
        'peaker': {
            'title': 'Natural Gas Peaking Facility',
            'summary': [
                'Acquire 180 MW simple-cycle peaker at $90mm ($500/kW)',
                'Low dispatch (400 hrs/yr) but high capacity value',
                'Revenue weighted toward capacity payments (70%+)',
                'Target returns: 14-16% levered IRR, 2.0-2.2x MOIC'
            ],
            'thesis': [
                ('Capacity-weighted revenue model', 'Peakers earn on reliability, not runtime. Low dispatch hours limit fuel cost exposure.'),
                ('Grid reliability premium', 'As renewables penetrate, fast-start peakers become more valuable for grid stability.'),
                ('Lower capital intensity', 'Simple cycle costs $500/kW vs $1,400/kW for CCGT. Lower basis improves returns.')
            ],
            'risks': [
                ('Capacity price volatility', 'Long-term capacity contracts where available'),
                ('Start-cycle wear', 'Maintenance reserves; long-term service agreement'),
                ('Renewable displacement', 'Battery storage competition in 5-10 years')
            ],
            'metrics': {'EV': '$90mm', 'Leverage': '40%', 'DSCR': '1.32x', 'IRR': '15.0%', 'MOIC': '2.1x'}
        },
        'solar': {
            'title': 'Utility-Scale Solar + BESS Project',
            'summary': [
                'Acquire 150 MW solar + 50 MW/200 MWh BESS at $280mm',
                '20-year PPA with investment-grade offtaker at $55/MWh',
                'Contracted cash flows enable sculpted debt structure',
                'Target returns: 10-12% levered IRR, 1.6-1.8x MOIC'
            ],
            'thesis': [
                ('Long-term contracted revenue', '20-year PPA eliminates merchant risk. Creditworthy offtaker provides revenue certainty.'),
                ('ITC benefit enhances returns', '30% Investment Tax Credit reduces equity basis, significantly improving IRR.'),
                ('Storage adds value stack', 'BESS enables arbitrage, ancillary services, and capacity revenue beyond solar-only project.')
            ],
            'risks': [
                ('Panel degradation', 'Tier-1 modules with performance guarantees; conservative degradation assumptions'),
                ('Battery degradation', 'Augmentation capex budgeted for Year 5-7'),
                ('Curtailment risk', 'PPAincludes curtailment provisions; transmission study completed')
            ],
            'metrics': {'EV': '$280mm', 'Leverage': '65%', 'DSCR': '1.35x', 'IRR': '11.0%', 'MOIC': '1.7x'}
        },
        'transmission': {
            'title': 'High-Voltage Transmission Line',
            'summary': [
                'Develop 200-mile, 500 kV transmission line at $850mm total cost',
                'Regulated asset with 40-year tariff at $180/kW-yr',
                '3-year construction, 37-year operating period',
                'Target returns: 9-11% levered IRR, 1.5-1.7x MOIC'
            ],
            'thesis': [
                ('Regulated, predictable cash flows', 'FERC-approved tariff provides long-term revenue certainty. Inflation escalators protect real returns.'),
                ('Essential infrastructure', 'Transmission is critical for renewable integration. Regulatory support for new build.'),
                ('RAB-based exit valuation', 'Exit at RAB multiple provides valuation floor independent of market conditions.')
            ],
            'risks': [
                ('Construction execution', 'EPC contract with liquidated damages; experienced contractor'),
                ('Regulatory reset risk', 'Long tariff term (40 yr); track record of fair treatment'),
                ('Interest rate sensitivity', 'Fixed-rate debt; long tenor matches asset life')
            ],
            'metrics': {'EV': '$850mm', 'Leverage': '70%', 'DSCR': '1.40x', 'IRR': '10.0%', 'MOIC': '1.6x'}
        },
        'midstream': {
            'title': 'Natural Gas Gathering System',
            'summary': [
                'Acquire gas gathering system at $85mm (6.5x EBITDA)',
                'Fee-based revenue with minimum volume commitments',
                'Growth phase Y1-5 (3%/yr), decline Y6+ (5%/yr)',
                'Target returns: 16-18% levered IRR, 2.2-2.5x MOIC'
            ],
            'thesis': [
                ('Fee-based, volume-protected revenue', 'MVCs provide downside protection. Fee structure insulates from commodity prices.'),
                ('Near-term growth optionality', 'Undeveloped acreage dedication provides volume upside as producer drills.'),
                ('Aggressive deleveraging', '75% cash sweep with 5-year tenor rapidly pays down debt before decline phase.')
            ],
            'risks': [
                ('Producer credit risk', 'Investment-grade anchor shipper; diversified producer base'),
                ('Volume decline steeper than modeled', 'Conservative decline assumptions; exit before steep decline'),
                ('Commodity price impact on drilling', 'MVC protection; acreage dedication locks in volumes')
            ],
            'metrics': {'EV': '$85mm', 'Leverage': '55%', 'DSCR': '1.35x', 'IRR': '17.0%', 'MOIC': '2.3x'}
        }
    }

    # Determine which memo to use
    memo_key = None
    for key in memos:
        if key in asset_type.lower():
            memo_key = key
            break

    if not memo_key:
        memo_key = 'ccgt'  # Default

    memo = memos[memo_key]

    # Write memo content
    row = 1

    # Title
    ws.cell(row=row, column=1, value="INVESTMENT COMMITTEE MEMORANDUM").font = Font(bold=True, size=16)
    row += 1
    ws.cell(row=row, column=1, value=memo['title']).font = Font(bold=True, size=14)
    row += 2

    # Executive Summary
    ws.cell(row=row, column=1, value="EXECUTIVE SUMMARY").font = section_font
    row += 1
    for bullet in memo['summary']:
        ws.cell(row=row, column=1, value=f"• {bullet}").font = body_font
        row += 1
    row += 1

    # Investment Thesis
    ws.cell(row=row, column=1, value="INVESTMENT THESIS").font = section_font
    row += 1
    for i, (point, rationale) in enumerate(memo['thesis'], 1):
        ws.cell(row=row, column=1, value=f"{i}. {point}").font = Font(bold=True, size=10)
        row += 1
        ws.cell(row=row, column=1, value=f"   {rationale}").font = body_font
        row += 1
    row += 1

    # Key Risks
    ws.cell(row=row, column=1, value="KEY RISKS & MITIGANTS").font = section_font
    row += 1
    ws.cell(row=row, column=1, value="Risk | Mitigant").font = Font(bold=True, size=10)
    row += 1
    for risk, mitigant in memo['risks']:
        ws.cell(row=row, column=1, value=f"{risk} | {mitigant}").font = body_font
        row += 1
    row += 1

    # Financial Summary
    ws.cell(row=row, column=1, value="FINANCIAL SUMMARY").font = section_font
    row += 1
    metrics = memo['metrics']
    ws.cell(row=row, column=1, value=f"Entry EV: {metrics['EV']} | Leverage: {metrics['Leverage']} | Min DSCR: {metrics['DSCR']}").font = body_font
    row += 1
    ws.cell(row=row, column=1, value=f"Levered IRR: {metrics['IRR']} | MOIC: {metrics['MOIC']}").font = body_font
    row += 2

    # Recommendation
    ws.cell(row=row, column=1, value="RECOMMENDATION").font = section_font
    row += 1
    ws.cell(row=row, column=1, value="Recommend APPROVAL subject to satisfactory completion of confirmatory due diligence.").font = body_font

    # Set row heights
    for r in range(1, row + 1):
        ws.row_dimensions[r].height = 18

    wb.save(filepath)
    print(f"  Added IC Memo sheet to: {filepath}")

def add_quick_reference_sheet(filepath):
    """Add Quick Reference sheet to Master Template"""
    wb = openpyxl.load_workbook(filepath)

    if "Quick Reference" in wb.sheetnames:
        del wb["Quick Reference"]

    ws = wb.create_sheet("Quick Reference", 0)

    # Formatting
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 45

    header_font = Font(bold=True, size=12)
    subheader_font = Font(bold=True, size=10)
    body_font = Font(size=9)

    row = 1

    # Section A: Debt Structure
    ws.cell(row=row, column=1, value="═══ DEBT STRUCTURE DECISION TREE ═══").font = header_font
    row += 2
    ws.cell(row=row, column=1, value="Cash Flow Type").font = subheader_font
    ws.cell(row=row, column=2, value="Structure").font = subheader_font
    ws.cell(row=row, column=3, value="Key Formula").font = subheader_font
    row += 1

    structures = [
        ("Contracted (PPA, tariff)", "Sculpted to 1.30-1.40x DSCR", "Prin = MIN(MAX(0, CFADS/DSCR − Int), Bal)"),
        ("Merchant (power, midstream)", "75% Cash Sweep, 1.25x min", "Sweep = MIN(Excess × 75%, Bal − Sched)"),
        ("Hybrid/Semi-contracted", "Sweep + tighter covenant", "Same as sweep with 1.30x DSCR"),
    ]
    for cf_type, structure, formula in structures:
        ws.cell(row=row, column=1, value=cf_type).font = body_font
        ws.cell(row=row, column=2, value=structure).font = body_font
        ws.cell(row=row, column=3, value=formula).font = body_font
        row += 1

    row += 2

    # Section B: Unit Conversions
    ws.cell(row=row, column=1, value="═══ UNIT CONVERSIONS ═══").font = header_font
    row += 2
    ws.cell(row=row, column=1, value="Calculation").font = subheader_font
    ws.cell(row=row, column=2, value="Formula").font = subheader_font
    ws.cell(row=row, column=3, value="Common Error").font = subheader_font
    row += 1

    conversions = [
        ("Fuel Cost ($mm)", "Gen × HR × Gas / 1E9", "÷1E6 → 1000× too high"),
        ("Energy Revenue ($mm)", "Gen × Price / 1E6", "Forgetting ÷1E6"),
        ("Capacity Revenue ($mm)", "UCAP × Price × 1000 / 1E6", "Missing ×1000 (price is $/kW)"),
        ("Peaker Generation (MWh)", "Capacity × Dispatch Hrs × Avail", "Using CF × 8760 (10× too high)"),
    ]
    for calc, formula, error in conversions:
        ws.cell(row=row, column=1, value=calc).font = body_font
        ws.cell(row=row, column=2, value=formula).font = body_font
        ws.cell(row=row, column=3, value=error).font = body_font
        row += 1

    row += 2

    # Section C: DSRA Logic
    ws.cell(row=row, column=1, value="═══ DSRA LOGIC ═══").font = header_font
    row += 2
    ws.cell(row=row, column=1, value="DSRA Target Y(n) = Scheduled DS Y(n+1) × (DSRA Months / 12)").font = subheader_font
    row += 2

    dsra_notes = [
        "• Reference NEXT year's scheduled DS, not current year",
        "• Use scheduled DS (straight-line), not actual (includes sweep)",
        "• This avoids circular reference in cash sweep models",
        "• Sculpted models: use actual_ds (no circularity — DS derived from CFADS)",
        "• DSRA released at exit (target = 0 in final year)",
    ]
    for note in dsra_notes:
        ws.cell(row=row, column=1, value=note).font = body_font
        row += 1

    row += 2

    # Section D: QA Checks
    ws.cell(row=row, column=1, value="═══ QA CHECKS ═══").font = header_font
    row += 2
    ws.cell(row=row, column=1, value="Check").font = subheader_font
    ws.cell(row=row, column=2, value="Formula").font = subheader_font
    ws.cell(row=row, column=3, value="Pass Condition").font = subheader_font
    row += 1

    checks = [
        ("S&U Balance", "ABS(Total Uses − Total Sources)", "< $0.01mm"),
        ("Debt Payoff", "Ending Balance at Debt Tenor", "< $1mm"),
        ("Min DSCR", "MIN(DSCR Y1:Yn)", "≥ 1.25x (sweep) or 1.30x (sculpt)"),
        ("No Errors", "Scan for #REF, #NAME, #VALUE, #DIV/0", "None found"),
    ]
    for check, formula, condition in checks:
        ws.cell(row=row, column=1, value=check).font = body_font
        ws.cell(row=row, column=2, value=formula).font = body_font
        ws.cell(row=row, column=3, value=condition).font = body_font
        row += 1

    row += 2

    # Section E: Common Mistakes
    ws.cell(row=row, column=1, value="═══ COMMON MISTAKES ═══").font = header_font
    row += 2

    mistakes = [
        ("Fuel cost ÷ 1E6", "Fuel = $30B (should be $30mm)", "Use ÷ 1E9"),
        ("DSRA = current DS", "Circular reference error", "Reference NEXT column"),
        ("Exit Equity = Exit EV", "IRR 5-10% too high", "Subtract debt, add DSRA"),
        ("Peaker CF × 8760", "Generation 10× too high", "Use Dispatch Hours × Avail"),
        ("Hardcoded row numbers", "Formulas break when rows move", "Use rows dict: rows['ebitda']"),
    ]
    for mistake, symptom, fix in mistakes:
        ws.cell(row=row, column=1, value=mistake).font = body_font
        ws.cell(row=row, column=2, value=symptom).font = body_font
        ws.cell(row=row, column=3, value=fix).font = body_font
        row += 1

    wb.save(filepath)
    print(f"  Added Quick Reference sheet to: {filepath}")

def main():
    output_dir = "/mnt/user-data/outputs"

    print("=" * 60)
    print("ENHANCING MODELS WITH COMMENTS AND SHEETS")
    print("=" * 60)

    # Process Models (Col C + Comments + IC Memo)
    models = [
        ("Model_1_CCGT.xlsx", "CCGT"),
        ("Model_2_Peaker.xlsx", "Peaker"),
        ("Model_3_SolarBESS.xlsx", "Solar BESS"),
        ("Model_4_Transmission.xlsx", "Transmission"),
        ("Model_5_Midstream.xlsx", "Midstream"),
    ]

    print("\n--- Processing Models ---")
    for filename, asset_type in models:
        filepath = os.path.join(output_dir, filename)
        if os.path.exists(filepath):
            add_formula_column_with_comments(filepath, asset_type)
            add_ic_memo_sheet(filepath, asset_type)
        else:
            print(f"  WARNING: {filepath} not found")

    # Process Drills (Col C + Comments only)
    drills = [
        ("Drill_1_CCGT.xlsx", "CCGT"),
        ("Drill_2_Peaker.xlsx", "Peaker"),
        ("Drill_3_SolarBESS.xlsx", "Solar BESS"),
        ("Drill_4_Transmission.xlsx", "Transmission"),
        ("Drill_5_Midstream.xlsx", "Midstream"),
    ]

    print("\n--- Processing Drills ---")
    for filename, asset_type in drills:
        filepath = os.path.join(output_dir, filename)
        if os.path.exists(filepath):
            add_formula_column_with_comments(filepath, asset_type)
        else:
            print(f"  WARNING: {filepath} not found")

    # Process Master Template (Quick Reference sheet)
    print("\n--- Processing Master Template ---")
    template_path = os.path.join(output_dir, "Master_Template.xlsx")
    if os.path.exists(template_path):
        add_quick_reference_sheet(template_path)
    else:
        print(f"  WARNING: {template_path} not found")

    print("\n" + "=" * 60)
    print("ENHANCEMENT COMPLETE")
    print("=" * 60)

if __name__ == "__main__":
    main()
