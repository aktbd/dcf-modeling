#!/usr/bin/env python3
"""
Create IC Memos for all 5 models
"""

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_cell_shading(cell, color):
    """Set cell background color"""
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), color)
    cell._tc.get_or_add_tcPr().append(shading_elm)

def add_heading_style(doc, text, level=1):
    """Add a styled heading"""
    heading = doc.add_heading(text, level=level)
    return heading

def create_metrics_table(doc, metrics):
    """Create a metrics table"""
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'

    # Header row
    hdr_cells = table.rows[0].cells
    headers = ['Metric', 'Base Case', 'Downside', 'Upside']
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
        hdr_cells[i].paragraphs[0].runs[0].bold = True
        set_cell_shading(hdr_cells[i], '1F4E79')
        hdr_cells[i].paragraphs[0].runs[0].font.color.rgb = None

    # Data rows
    for metric in metrics:
        row_cells = table.add_row().cells
        for i, value in enumerate(metric):
            row_cells[i].text = str(value)

    return table

def create_risk_table(doc, risks):
    """Create a risks table"""
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'

    hdr_cells = table.rows[0].cells
    headers = ['Risk', 'Probability', 'Mitigant']
    for i, header in enumerate(headers):
        hdr_cells[i].text = header
        hdr_cells[i].paragraphs[0].runs[0].bold = True
        set_cell_shading(hdr_cells[i], '1F4E79')

    for risk in risks:
        row_cells = table.add_row().cells
        for i, value in enumerate(risk):
            row_cells[i].text = str(value)

    return table

def create_ccgt_memo():
    """Create CCGT IC Memo"""
    doc = Document()

    # Title
    title = doc.add_heading('PROJECT SPARTAN', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = doc.add_paragraph('INVESTMENT COMMITTEE MEMORANDUM')
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].bold = True

    doc.add_paragraph('365 MW Combined Cycle Gas Turbine – PJM West')
    doc.add_paragraph()

    # Recommendation
    rec = doc.add_paragraph()
    rec.add_run('RECOMMENDATION: ').bold = True
    rec.add_run('Proceed to final bid at $520mm')

    doc.add_paragraph('─' * 60)

    # Key Metrics
    add_heading_style(doc, 'KEY METRICS', 2)
    metrics = [
        ('Entry EV', '$520mm', '-', '-'),
        ('Entry Multiple', '7.8x Y1 EBITDA', '-', '-'),
        ('Leverage', '45%', '-', '-'),
        ('Levered IRR', '18.5%', '14.2%', '22.8%'),
        ('MOIC', '2.15x', '1.85x', '2.45x'),
        ('Min DSCR', '1.42x', '1.28x', '1.58x'),
        ('Unlevered IRR', '12.1%', '-', '-'),
    ]
    create_metrics_table(doc, metrics)

    doc.add_paragraph()
    doc.add_paragraph('─' * 60)

    # Investment Thesis
    add_heading_style(doc, 'INVESTMENT THESIS', 2)
    thesis_points = [
        'Strategic PJM West location with dual revenue streams (energy + capacity) provides diversified cash flow profile',
        'Efficient 7,000 Btu/kWh heat rate positions asset competitively on dispatch stack, ensuring consistent energy margins',
        'Recent PJM capacity auction results support sustained capacity revenue outlook through the hold period',
        'Clear exit path to strategic buyers (utilities seeking gas generation) and infrastructure yield vehicles'
    ]
    for point in thesis_points:
        doc.add_paragraph(f'• {point}')

    doc.add_paragraph()
    doc.add_paragraph('─' * 60)

    # Key Risks
    add_heading_style(doc, 'KEY RISKS & MITIGANTS', 2)
    risks = [
        ('Gas price volatility', 'Medium', 'Partial hedge program (50% Y1-Y3); efficient heat rate provides margin cushion'),
        ('Capacity price decline', 'Medium', 'PJM capacity construct provides 3-year forward visibility; locked through Y3'),
        ('Forced outage risk', 'Low', 'Comprehensive LTSA with OEM; historical 95% availability; 6-month DSRA'),
        ('Environmental regulation', 'Medium', 'No pending EPA action on unit; brownfield expansion rights provide optionality'),
    ]
    create_risk_table(doc, risks)

    doc.add_paragraph()
    doc.add_paragraph('─' * 60)

    # Financing Structure
    add_heading_style(doc, 'FINANCING STRUCTURE RATIONALE', 2)
    doc.add_paragraph('Structure: 7-year term loan with 75% cash sweep and 6-month DSRA')
    doc.add_paragraph()
    doc.add_paragraph('Rationale:')
    rationale = [
        'Cash sweep structure appropriate for merchant asset with variable CFADS – accelerates paydown in strong years',
        '45% leverage balances return enhancement with covenant headroom for commodity cycles',
        '6-month DSRA provides liquidity buffer, sizing is market standard for merchant power',
        'Lock-up DSCR at 1.10x protects lenders while allowing distributions in base case'
    ]
    for point in rationale:
        doc.add_paragraph(f'• {point}')

    doc.add_paragraph()
    doc.add_paragraph('─' * 60)

    # IC Q&A
    add_heading_style(doc, 'ANTICIPATED IC Q&A', 2)

    qa_sections = {
        'VALUATION': [
            ('Q: How did you arrive at 7.8x entry multiple?', 'A: Based on recent PJM CCGT transactions ranging 7.0-8.5x, adjusted down for merchant exposure vs contracted comps. Premium to replacement cost reflects strategic location and grid constraints.'),
            ('Q: What\'s the exit buyer universe?', 'A: Primary: Utilities seeking gas generation for reliability (AES, NRG, Vistra). Secondary: Infrastructure funds (KKR, Brookfield, GIP) seeking yield with upside. Tertiary: IPPs for portfolio buildout.'),
            ('Q: Why 7.0x exit multiple vs 7.8x entry?', 'A: Conservative assumption reflects shorter remaining life at exit. Sensitivity shows 0.5x multiple change = 200bps IRR impact.'),
        ],
        'FINANCING': [
            ('Q: Why sweep instead of sculpt?', 'A: Merchant CFADS volatility makes sculpting impractical – you\'d need to re-sculpt every year. Sweep naturally adjusts paydown to actual cash flow.'),
            ('Q: How did you size the DSRA?', 'A: 6-month debt service is market standard for merchant power. Provides 6 months of runway if spark spreads compress or outage occurs.'),
            ('Q: What happens if DSCR drops below 1.10x?', 'A: Distributions locked up – all excess cash goes to debt paydown. Likely scenario: temporary commodity dislocation. Recovery expected within 1-2 years.'),
        ],
        'RETURNS': [
            ('Q: Walk me through the IRR bridge from unlevered to levered.', 'A: Unlevered 12.1% + leverage benefit (cheaper debt vs equity cost) ≈ +7% – financing costs (interest, fees) ≈ -0.6% = Levered 18.5%. Leverage adds ~6.5% to returns.'),
            ('Q: Why is MOIC 2.15x lower than I\'d expect for 18.5% IRR?', 'A: 7-year hold period. MOIC = cash-on-cash regardless of timing. 18.5% IRR over 7 years compounds to 2.15x.'),
            ('Q: What\'s the downside scenario?', 'A: Gas +25%, power flat = squeezed spark spreads. IRR drops to 14.2%, still above hurdle. DSCR 1.28x maintains covenant compliance.'),
        ],
        'MARKET': [
            ('Q: What happens if PJM capacity prices collapse?', 'A: Downside scenario assumes 20% capacity price decline. Still economic due to energy revenue diversification. Extreme downside: asset becomes peaker-like, relies on scarcity pricing.'),
            ('Q: How does carbon regulation affect the investment?', 'A: Gas is bridge fuel – CCGTs run when renewables can\'t. Carbon cost would increase power prices, potentially improving spark spreads. Natural gas is 50% lower emissions than coal.'),
        ],
        'OPERATIONS': [
            ('Q: What are the key operational risks?', 'A: Equipment failure (mitigated by LTSA), fuel supply disruption (mitigated by pipeline redundancy), grid curtailment (mitigated by must-run status).'),
            ('Q: What value creation levers exist?', 'A: 1) Heat rate improvement project (+$2mm EBITDA), 2) Ancillary services expansion, 3) Capacity factor optimization through better dispatch strategy.'),
        ],
    }

    for section, qas in qa_sections.items():
        doc.add_paragraph(section, style='Heading 3')
        for q, a in qas:
            p = doc.add_paragraph()
            p.add_run(q).bold = True
            doc.add_paragraph(a)
        doc.add_paragraph()

    doc.add_paragraph('─' * 60)

    # Diligence Plan
    add_heading_style(doc, 'DILIGENCE PLAN', 2)

    diligence = {
        'Technical': [
            'Independent engineer report (Wood Mackenzie)',
            'Equipment condition assessment – turbine hours remaining',
            'Environmental compliance review – air permits, water discharge',
            'Grid interconnection study – transmission constraints',
        ],
        'Commercial': [
            'PPA/hedging contract review',
            'Market study validation (ERCOT/PJM consultant)',
            'Fuel supply and transportation agreements',
            'Counterparty credit analysis',
        ],
        'Legal': [
            'Title and land rights review',
            'Permitting status and renewal timeline',
            'Environmental liability assessment',
            'Regulatory status and rate case history',
        ],
        'Financial': [
            'Quality of earnings (trailing 3-year analysis)',
            'Working capital normalization',
            'Tax structure optimization',
            'Insurance adequacy review',
        ],
    }

    for category, items in diligence.items():
        doc.add_paragraph(f'{category}:', style='Heading 3')
        for item in items:
            doc.add_paragraph(f'☐ {item}')

    doc.save('/mnt/user-data/outputs/IC_Memo_1_CCGT.docx')
    print('Created IC_Memo_1_CCGT.docx')

def create_peaker_memo():
    """Create Peaker IC Memo"""
    doc = Document()

    title = doc.add_heading('PROJECT PHOENIX', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = doc.add_paragraph('INVESTMENT COMMITTEE MEMORANDUM')
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].bold = True

    doc.add_paragraph('200 MW Simple Cycle Peaker – NYISO Zone J')
    doc.add_paragraph()

    rec = doc.add_paragraph()
    rec.add_run('RECOMMENDATION: ').bold = True
    rec.add_run('Proceed to final bid at $110mm')

    doc.add_paragraph('─' * 60)

    add_heading_style(doc, 'KEY METRICS', 2)
    metrics = [
        ('Entry EV', '$110mm', '-', '-'),
        ('Entry Multiple', '8.5x Y1 EBITDA', '-', '-'),
        ('Leverage', '30%', '-', '-'),
        ('Levered IRR', '16.8%', '12.5%', '21.2%'),
        ('MOIC', '1.95x', '1.65x', '2.25x'),
        ('Min DSCR', '1.55x', '1.35x', '1.75x'),
        ('Unlevered IRR', '13.2%', '-', '-'),
    ]
    create_metrics_table(doc, metrics)

    doc.add_paragraph()
    add_heading_style(doc, 'INVESTMENT THESIS', 2)
    thesis = [
        'NYISO Zone J capacity scarcity drives premium capacity prices ($60/kW-yr vs $40 market average)',
        'Fast-start capability (10-minute dispatch) positions asset for ancillary services revenue growth',
        'Conservative 30% leverage with 50% sweep retains cash for major maintenance',
        'Exit to utility seeking reliability assets in constrained zone'
    ]
    for point in thesis:
        doc.add_paragraph(f'• {point}')

    doc.add_paragraph()
    add_heading_style(doc, 'KEY RISKS & MITIGANTS', 2)
    risks = [
        ('Capacity price decline', 'Medium', 'Zone J structural scarcity; 3-year forward locked'),
        ('Major maintenance timing', 'Medium', 'HGP in Y5 budgeted; LTSA provides cost certainty'),
        ('Limited dispatch hours', 'Low', 'Capacity revenue = 65% of total; energy is upside'),
        ('Battery competition', 'Medium', 'Batteries complement, not replace – need gas for multi-day events'),
    ]
    create_risk_table(doc, risks)

    doc.add_paragraph()
    add_heading_style(doc, 'ANTICIPATED IC Q&A', 2)

    qa = [
        ('Q: Why dispatch hours instead of capacity factor?', 'A: Peakers run only during peak demand (~400 hours). Using CF × 8,760 would overstate generation by 10x. This is the most common modeling error.'),
        ('Q: What if the peaker never dispatches?', 'A: Still receive capacity payments (65% of revenue). Energy revenue is upside, not base case. Worst case = capacity-only returns still above hurdle.'),
        ('Q: Why 50% sweep vs 75% for CCGT?', 'A: Need to retain cash for major maintenance (HGP every 5 years = $4mm). Higher sweep would require external funding for maintenance.'),
        ('Q: How does battery storage affect peakers?', 'A: 4-hour batteries handle daily peaks. Peakers needed for multi-day events, extreme weather, renewable droughts. Complementary, not substitutive.'),
    ]

    for q, a in qa:
        p = doc.add_paragraph()
        p.add_run(q).bold = True
        doc.add_paragraph(a)

    doc.save('/mnt/user-data/outputs/IC_Memo_2_Peaker.docx')
    print('Created IC_Memo_2_Peaker.docx')

def create_solar_bess_memo():
    """Create Solar+BESS IC Memo"""
    doc = Document()

    title = doc.add_heading('PROJECT SUNVAULT', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = doc.add_paragraph('INVESTMENT COMMITTEE MEMORANDUM')
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].bold = True

    doc.add_paragraph('115 MWac Solar + 50 MW/200 MWh BESS – ERCOT West')
    doc.add_paragraph()

    rec = doc.add_paragraph()
    rec.add_run('RECOMMENDATION: ').bold = True
    rec.add_run('Proceed to final bid at $145mm')

    doc.add_paragraph('─' * 60)

    add_heading_style(doc, 'KEY METRICS', 2)
    metrics = [
        ('Entry EV', '$145mm', '-', '-'),
        ('Entry Multiple', '9.2x Y1 EBITDA', '-', '-'),
        ('Leverage', '35%', '-', '-'),
        ('Net Equity (post-ITC)', '$55mm', '-', '-'),
        ('Levered IRR', '17.2%', '13.8%', '20.5%'),
        ('MOIC', '2.35x', '2.05x', '2.65x'),
        ('Min DSCR', '1.35x', '1.28x', '1.42x'),
    ]
    create_metrics_table(doc, metrics)

    doc.add_paragraph()
    add_heading_style(doc, 'INVESTMENT THESIS', 2)
    thesis = [
        '10-year investment-grade PPA provides contracted revenue certainty, supporting sculpted debt structure',
        '30% ITC reduces equity requirement by $43.5mm, significantly enhancing returns',
        'BESS provides additional revenue streams (arbitrage, ancillary) with limited incremental risk',
        'Exit to yield-oriented infrastructure buyer at premium multiple (8.0x vs 6.5x for thermal)'
    ]
    for point in thesis:
        doc.add_paragraph(f'• {point}')

    doc.add_paragraph()
    add_heading_style(doc, 'KEY RISKS & MITIGANTS', 2)
    risks = [
        ('Degradation exceeds forecast', 'Low', 'Panel warranty covers 80% output at Y25; 0.5%/yr is conservative'),
        ('BESS revenue volatility', 'Medium', 'BESS = 20% of revenue; PPA provides floor'),
        ('Offtaker credit risk', 'Low', 'Investment-grade utility; standard security package'),
        ('ITC recapture', 'Low', '5-year recapture period; no change of control planned'),
    ]
    create_risk_table(doc, risks)

    doc.add_paragraph()
    add_heading_style(doc, 'ANTICIPATED IC Q&A', 2)

    qa = [
        ('Q: Why sculpted debt instead of sweep?', 'A: Contracted PPA provides predictable CFADS. Sculpting efficiently sizes debt service to maintain target DSCR, maximizing debt capacity.'),
        ('Q: How does ITC flow through S&U?', 'A: ITC is a "source" offsetting equity. Pre-ITC Equity ($98mm) - ITC ($43.5mm) = Net Equity ($55mm). You write a $55mm check, not $98mm.'),
        ('Q: What happens to BESS revenue if volatility drops?', 'A: Arbitrage revenue declines but ancillary services (frequency regulation) remain stable. PPA revenue is 80% of total – BESS is upside.'),
        ('Q: Why Y7 augmentation capex?', 'A: Battery capacity degrades 2-3% annually. Y7 augmentation ($5mm) restores capacity to maintain contracted services. Factor into CFADS.'),
    ]

    for q, a in qa:
        p = doc.add_paragraph()
        p.add_run(q).bold = True
        doc.add_paragraph(a)

    doc.save('/mnt/user-data/outputs/IC_Memo_3_SolarBESS.docx')
    print('Created IC_Memo_3_SolarBESS.docx')

def create_transmission_memo():
    """Create Transmission IC Memo"""
    doc = Document()

    title = doc.add_heading('PROJECT CORRIDOR', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = doc.add_paragraph('INVESTMENT COMMITTEE MEMORANDUM')
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].bold = True

    doc.add_paragraph('3,200 MW HVDC Transmission Line – SPP to MISO')
    doc.add_paragraph()

    rec = doc.add_paragraph()
    rec.add_run('RECOMMENDATION: ').bold = True
    rec.add_run('Proceed to development at $585mm total investment')

    doc.add_paragraph('─' * 60)

    add_heading_style(doc, 'KEY METRICS', 2)
    metrics = [
        ('RAB at COD', '$608mm', '-', '-'),
        ('Total Equity', '$265mm', '-', '-'),
        ('Term Debt % RAB', '60%', '-', '-'),
        ('Levered IRR', '14.5%', '12.2%', '16.8%'),
        ('MOIC', '2.05x', '1.85x', '2.25x'),
        ('Min DSCR (Y3-Y10)', '1.40x', '1.32x', '1.48x'),
    ]
    create_metrics_table(doc, metrics)

    doc.add_paragraph()
    add_heading_style(doc, 'INVESTMENT THESIS', 2)
    thesis = [
        'FERC-approved tariff provides stable, regulated return on 40-year asset life',
        'Critical infrastructure connecting low-cost wind to high-demand load centers',
        'Construction interest capitalization into RAB improves economics',
        'Exit at 1.15x RAB to infrastructure yield vehicles seeking stable returns'
    ]
    for point in thesis:
        doc.add_paragraph(f'• {point}')

    doc.add_paragraph()
    add_heading_style(doc, 'KEY RISKS & MITIGANTS', 2)
    risks = [
        ('Construction delay/overrun', 'Medium', 'Fixed-price EPC; liquidated damages; contingency budget'),
        ('Regulatory ROE reduction', 'Low', 'FERC historically stable; bipartisan transmission support'),
        ('Permitting challenges', 'Medium', 'Pre-approved route; tribal and environmental clearances secured'),
        ('Counterparty default', 'Low', 'Utility offtakers are investment-grade; regulatory backstop'),
    ]
    create_risk_table(doc, risks)

    doc.add_paragraph()
    add_heading_style(doc, 'ANTICIPATED IC Q&A', 2)

    qa = [
        ('Q: Why does transmission trade at premium to RAB?', 'A: Scarcity value – new transmission is hard to permit. 40+ year life with stable regulated returns. Critical infrastructure status. 1.15x is conservative vs recent deals at 1.25x+.'),
        ('Q: Walk me through construction financing.', 'A: Y0-Y1: Equity funds development ($35mm). Y2: Construction loan (85% × $550mm = $467.5mm) + equity funds construction. Y3 (COD): Refi into term debt (60% × RAB) + equity to repay construction loan.'),
        ('Q: What if construction is delayed?', 'A: Higher capitalized interest, later revenue start. 6-month delay = ~80bps IRR reduction. Mitigated by fixed-price EPC with liquidated damages.'),
        ('Q: What\'s the regulatory risk?', 'A: FERC could reduce allowed ROE. Historically stable at 10-11%. Even at 9%, returns still exceed cost of capital due to 60% leverage.'),
    ]

    for q, a in qa:
        p = doc.add_paragraph()
        p.add_run(q).bold = True
        doc.add_paragraph(a)

    doc.save('/mnt/user-data/outputs/IC_Memo_4_Transmission.docx')
    print('Created IC_Memo_4_Transmission.docx')

def create_midstream_memo():
    """Create Midstream IC Memo"""
    doc = Document()

    title = doc.add_heading('PROJECT PERMIAN GATHERING', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = doc.add_paragraph('INVESTMENT COMMITTEE MEMORANDUM')
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].bold = True

    doc.add_paragraph('80 MMcf/d Gas Gathering System – Permian Basin')
    doc.add_paragraph()

    rec = doc.add_paragraph()
    rec.add_run('RECOMMENDATION: ').bold = True
    rec.add_run('Proceed to final bid at $85mm')

    doc.add_paragraph('─' * 60)

    add_heading_style(doc, 'KEY METRICS', 2)
    metrics = [
        ('Entry EV', '$85mm', '-', '-'),
        ('Entry Multiple', '5.8x Y1 EBITDA', '-', '-'),
        ('Leverage', '30%', '-', '-'),
        ('Levered IRR', '19.2%', '15.5%', '23.0%'),
        ('MOIC', '2.05x', '1.75x', '2.35x'),
        ('Min DSCR (Y1-Y5)', '1.65x', '1.45x', '1.85x'),
    ]
    create_metrics_table(doc, metrics)

    doc.add_paragraph()
    add_heading_style(doc, 'INVESTMENT THESIS', 2)
    thesis = [
        'Fee-for-service model eliminates commodity price exposure – paid to move gas, not sell it',
        'Active drilling program supports 3%/year volume growth through Y5',
        'Aggressive 75% sweep with 5-year tenor ensures debt payoff before decline phase',
        'Exit to producer seeking vertical integration or larger midstream for basin consolidation'
    ]
    for point in thesis:
        doc.add_paragraph(f'• {point}')

    doc.add_paragraph()
    add_heading_style(doc, 'KEY RISKS & MITIGANTS', 2)
    risks = [
        ('Producer curtails drilling', 'Medium', 'Acreage dedication provides volume floor; producer needs gatherer to monetize'),
        ('Volume decline steeper than forecast', 'Medium', '5%/year decline is conservative for Permian; exit at Y7 before steep decline'),
        ('Single counterparty concentration', 'Medium', 'Producer is well-capitalized; consider credit support'),
        ('Basin differentials widen', 'Low', 'Fee-based, not commodity-based; producer bears basis risk'),
    ]
    create_risk_table(doc, risks)

    doc.add_paragraph()
    add_heading_style(doc, 'ANTICIPATED IC Q&A', 2)

    qa = [
        ('Q: Why the growth-then-decline profile?', 'A: Y1-Y5: Active drilling connects new wells (3%/year growth). Y6+: Drilling slows, existing wells deplete naturally (5%/year decline). Must structure debt to pay off before decline.'),
        ('Q: Why 5-year tenor vs 7 years?', 'A: Debt must be repaid during growth phase when CFADS is strong. Cannot service debt with declining cash flows. Short tenor + aggressive sweep = de-risked debt.'),
        ('Q: What\'s the commodity exposure?', 'A: Direct: None – fee-for-service. Indirect: Low gas prices → less drilling → lower volumes. Mitigated by acreage dedication and producer economics (breakeven $2.50/mmBtu).'),
        ('Q: Why compressed 6.5x exit multiple?', 'A: Declining asset commands lower multiple than stable/growing. At exit (Y7), asset has 3-4 years of decline remaining. Buyer pricing reflects PDP-like valuation.'),
    ]

    for q, a in qa:
        p = doc.add_paragraph()
        p.add_run(q).bold = True
        doc.add_paragraph(a)

    doc.save('/mnt/user-data/outputs/IC_Memo_5_Midstream.docx')
    print('Created IC_Memo_5_Midstream.docx')

if __name__ == "__main__":
    create_ccgt_memo()
    create_peaker_memo()
    create_solar_bess_memo()
    create_transmission_memo()
    create_midstream_memo()
    print("\nAll IC Memos created successfully!")
