#!/usr/bin/env python3
"""
Create PDF version of the cheat sheet using reportlab
"""

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT

def create_cheat_sheet_pdf():
    """Create a formatted PDF cheat sheet"""

    doc = SimpleDocTemplate(
        "/mnt/user-data/outputs/Lotus_Interview_CheatSheet_Expanded.pdf",
        pagesize=letter,
        rightMargin=0.5*inch,
        leftMargin=0.5*inch,
        topMargin=0.5*inch,
        bottomMargin=0.5*inch
    )

    styles = getSampleStyleSheet()

    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        spaceAfter=12,
        alignment=TA_CENTER,
        textColor=colors.HexColor('#1F4E79')
    )

    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=14,
        spaceBefore=12,
        spaceAfter=6,
        textColor=colors.HexColor('#1F4E79')
    )

    subheading_style = ParagraphStyle(
        'CustomSubheading',
        parent=styles['Heading3'],
        fontSize=11,
        spaceBefore=8,
        spaceAfter=4,
        textColor=colors.black
    )

    body_style = ParagraphStyle(
        'CustomBody',
        parent=styles['Normal'],
        fontSize=9,
        spaceBefore=2,
        spaceAfter=2
    )

    bullet_style = ParagraphStyle(
        'CustomBullet',
        parent=styles['Normal'],
        fontSize=9,
        leftIndent=20,
        spaceBefore=1,
        spaceAfter=1
    )

    story = []

    # Title
    story.append(Paragraph("LOTUS INFRASTRUCTURE PARTNERS", title_style))
    story.append(Paragraph("Infrastructure PE Interview Cheat Sheet", styles['Heading2']))
    story.append(Spacer(1, 12))

    # Section 1: Firm Intelligence
    story.append(Paragraph("1. FIRM INTELLIGENCE", heading_style))
    story.append(Paragraph("<b>About Lotus Infrastructure Partners</b>", subheading_style))
    story.append(Paragraph("• Formerly Starwood Energy Group (rebranded 2023)", bullet_style))
    story.append(Paragraph("• Founded: 2005 | HQ: Greenwich, CT | AUM: $3B+", bullet_style))
    story.append(Paragraph("• Focus: North American power and energy infrastructure", bullet_style))
    story.append(Spacer(1, 6))

    story.append(Paragraph("<b>Investment Thesis</b>", subheading_style))
    story.append(Paragraph("• Target contracted and regulated assets with visible cash flows", bullet_style))
    story.append(Paragraph("• Focus on energy transition enabling infrastructure", bullet_style))
    story.append(Paragraph("• Typical hold: 5-10 years | Target: Mid-teens IRR, 1.8-2.5x MOIC", bullet_style))
    story.append(Spacer(1, 12))

    # Section 2: Asset Classes
    story.append(Paragraph("2. ASSET CLASS QUICK REFERENCE", heading_style))

    # Asset table
    asset_data = [
        ['Asset Type', 'Revenue Model', 'Typical Leverage', 'Key Risk'],
        ['CCGT', 'Energy + Capacity', '40-50%', 'Spark spread'],
        ['Peaker', 'Capacity (60%+)', '25-35%', 'Dispatch hours'],
        ['Solar+BESS', 'PPA + Arbitrage', '30-40%', 'Degradation'],
        ['Transmission', 'Regulated tariff', '55-65%', 'Regulatory'],
        ['Midstream', 'Fee × Volume', '25-35%', 'Volume decline'],
    ]

    asset_table = Table(asset_data, colWidths=[1.2*inch, 1.3*inch, 1*inch, 1.2*inch])
    asset_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E79')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ]))
    story.append(asset_table)
    story.append(Spacer(1, 12))

    # Section 3: Key Formulas
    story.append(Paragraph("3. KEY FORMULAS", heading_style))

    story.append(Paragraph("<b>Generation Calculations</b>", subheading_style))
    story.append(Paragraph("CCGT: Generation = Capacity × CF × 8,760", body_style))
    story.append(Paragraph("Peaker: Generation = Capacity × Dispatch Hours × Availability (NOT CF × 8,760!)", body_style))
    story.append(Paragraph("UCAP = Capacity × (1 - Forced Outage Rate)", body_style))
    story.append(Spacer(1, 6))

    story.append(Paragraph("<b>Revenue Calculations</b>", subheading_style))
    story.append(Paragraph("Energy Revenue ($mm) = Generation × Price / 1,000,000", body_style))
    story.append(Paragraph("Capacity Revenue ($mm) = UCAP × $/kW-yr × 1,000 / 1,000,000", body_style))
    story.append(Spacer(1, 6))

    story.append(Paragraph("<b>Fuel Cost (CRITICAL)</b>", subheading_style))
    story.append(Paragraph("Fuel ($mm) = Generation × Heat Rate × Gas Price / 1,000,000,000", body_style))
    story.append(Paragraph("Quick check: Variable cost ≈ Heat Rate × Gas / 1,000 ($/MWh)", body_style))
    story.append(Spacer(1, 12))

    # Section 4: Debt Mechanics
    story.append(Paragraph("4. DEBT MECHANICS", heading_style))

    story.append(Paragraph("<b>Sweep vs Sculpt Decision</b>", subheading_style))
    story.append(Paragraph("• CFADS predictable/contracted → Use SCULPTED (size DS to target DSCR)", bullet_style))
    story.append(Paragraph("• CFADS volatile/merchant → Use SWEEP (prepay excess cash)", bullet_style))
    story.append(Spacer(1, 6))

    story.append(Paragraph("<b>DSRA Mechanics</b>", subheading_style))
    story.append(Paragraph("• Target = NEXT year's DS × (DSRA Months / 12)", bullet_style))
    story.append(Paragraph("• Funding = Target - Beginning (positive = fund, negative = release)", bullet_style))
    story.append(Spacer(1, 12))

    # Section 5: Common Mistakes
    story.append(Paragraph("5. COMMON MISTAKES TO AVOID", heading_style))

    mistakes_data = [
        ['Mistake', 'Correct Approach'],
        ['Fuel cost divisor wrong', 'Divide by 1e9, not 1e6'],
        ['Peaker uses CF × 8,760', 'Use Dispatch Hours × Availability'],
        ['Capacity price per MW', 'Price is per kW - multiply by 1,000'],
        ['DSRA uses current year DS', 'Target = NEXT year DS'],
        ['Exit equity = Exit EV', 'Exit equity = EV - Debt + DSRA'],
    ]

    mistakes_table = Table(mistakes_data, colWidths=[2.2*inch, 2.8*inch])
    mistakes_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E79')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ]))
    story.append(mistakes_table)
    story.append(Spacer(1, 12))

    # Page break for markets
    story.append(PageBreak())

    # Section 6: Regional Markets
    story.append(Paragraph("6. REGIONAL MARKET QUICK REFERENCE", heading_style))

    markets_data = [
        ['Market', 'Structure', 'Key Dynamic', 'Interview Question'],
        ['PJM', 'Capacity (RPM)', '3-yr forward visibility', 'How underwrite cap prices?'],
        ['ERCOT', 'Energy-only', '$9K scarcity pricing', 'Why no capacity market?'],
        ['CAISO', 'Capacity (RA)', 'Duck curve opportunity', 'Explain duck curve'],
        ['NYISO', 'Capacity (ICAP)', 'Zone J premium (2-3x)', 'What drives Zone J?'],
        ['MISO', 'Voluntary cap', 'Wind-rich, weak cap', 'Why low cap prices?'],
        ['ISO-NE', 'Capacity (FCM)', 'Winter gas spikes', 'Mystic reliability?'],
    ]

    markets_table = Table(markets_data, colWidths=[0.7*inch, 1*inch, 1.4*inch, 1.8*inch])
    markets_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E79')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ]))
    story.append(markets_table)
    story.append(Spacer(1, 12))

    # Section 7: Build Timeline
    story.append(Paragraph("7. 4-HOUR BUILD SEQUENCE", heading_style))

    story.append(Paragraph("<b>Hour 1:</b> Read materials, set up structure, input assumptions, start operating model", body_style))
    story.append(Paragraph("<b>Hour 2:</b> Complete operating model, cash flow, debt schedule", body_style))
    story.append(Paragraph("<b>Hour 3:</b> DSRA, waterfall, S&U, returns, sanity checks", body_style))
    story.append(Paragraph("<b>Hour 4:</b> Fix errors, sensitivities, IC memo, final review", body_style))
    story.append(Spacer(1, 12))

    # Section 8: Quick Reference
    story.append(Paragraph("8. QUICK REFERENCE CARD", heading_style))

    story.append(Paragraph("<b>Standard Assumptions</b>", subheading_style))
    story.append(Paragraph("Leverage: 30-60% | Interest: 5-8% | DSRA: 6 months | Tax: 25%", body_style))
    story.append(Paragraph("Depreciation: 15-20 yrs SL | Exit Multiple: 6-8x EBITDA", body_style))
    story.append(Spacer(1, 6))

    story.append(Paragraph("<b>Red Flag Outputs</b>", subheading_style))
    story.append(Paragraph("DSCR < 1.25x | IRR < 10% or > 35% | MOIC < 1.5x | Debt > 50% at exit", body_style))
    story.append(Spacer(1, 6))

    story.append(Paragraph("<b>Unit Conversions</b>", subheading_style))
    story.append(Paragraph("1 MW = 1,000 kW | 1 MWh = 1,000 kWh | 8,760 hours/year", body_style))
    story.append(Paragraph("1 mmBtu = 1,000,000 Btu | 1 Bcf = 1,000 MMcf", body_style))
    story.append(Spacer(1, 12))

    # Footer
    story.append(Spacer(1, 24))
    story.append(Paragraph("<i>Good luck! Remember: accuracy > speed. A working model beats a complex model with errors.</i>",
                          ParagraphStyle('Footer', parent=body_style, alignment=TA_CENTER, textColor=colors.grey)))

    doc.build(story)
    print("Created Lotus_Interview_CheatSheet_Expanded.pdf")

if __name__ == "__main__":
    create_cheat_sheet_pdf()
