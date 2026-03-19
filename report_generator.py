"""
-----------------------------------------------
Name            : Sahana V
College         : [Seshadripuram College]
Internship Domain : Python Internship


Description:
This project is developed as part of my internship task.
It demonstrates the implementation of Python-based
solutions using libraries such as scikit-learn, NLTK,
and spaCy for building intelligent applications.

Technologies Used:
- Python
- Machine Learning
- Natural Language Processing
- VS Code
-----------------------------------------------
"""
import sys
import csv
import os
from datetime import datetime
from collections import defaultdict

# ReportLab imports

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether
)
from reportlab.graphics.shapes import Drawing, Rect, String, Line
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics import renderPDF
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT



#  COLOR PALETTE

BRAND_BLUE    = colors.HexColor("#1A3C6E")
BRAND_ACCENT  = colors.HexColor("#2E86C1")
BRAND_LIGHT   = colors.HexColor("#D6EAF8")
BRAND_GREEN   = colors.HexColor("#1E8449")
BRAND_RED     = colors.HexColor("#C0392B")
BRAND_ORANGE  = colors.HexColor("#E67E22")
GRAY_DARK     = colors.HexColor("#2C3E50")
GRAY_MID      = colors.HexColor("#7F8C8D")
GRAY_LIGHT    = colors.HexColor("#ECF0F1")
WHITE         = colors.white

CHART_COLORS = [
    colors.HexColor("#2E86C1"),
    colors.HexColor("#E67E22"),
    colors.HexColor("#1E8449"),
    colors.HexColor("#8E44AD"),
    colors.HexColor("#C0392B"),
    colors.HexColor("#16A085"),
]



#  DATA LOADING & ANALYSIS


def load_data(filepath: str) -> list[dict]:
    """Load CSV data and cast numeric fields."""
    records = []
    with open(filepath, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            row["Units_Sold"]  = int(row["Units_Sold"])
            row["Unit_Price"]  = float(row["Unit_Price"])
            row["Revenue"]     = float(row["Revenue"])
            row["Target"]      = float(row["Target"])
            records.append(row)
    return records


def analyze(records: list[dict]) -> dict:
    """Run all analytics and return a results dictionary."""
    total_revenue  = sum(r["Revenue"] for r in records)
    total_units    = sum(r["Units_Sold"] for r in records)
    total_target   = sum(r["Target"] for r in records)
    achievement    = (total_revenue / total_target * 100) if total_target else 0

    # ── By region ──────────────────────────────
    region_rev   = defaultdict(float)
    region_units = defaultdict(int)
    for r in records:
        region_rev[r["Region"]]   += r["Revenue"]
        region_units[r["Region"]] += r["Units_Sold"]

    # ── By salesperson ─────────────────────────
    sales_rev    = defaultdict(float)
    sales_units  = defaultdict(int)
    sales_target = defaultdict(float)
    for r in records:
        sales_rev[r["Salesperson"]]    += r["Revenue"]
        sales_units[r["Salesperson"]]  += r["Units_Sold"]
        sales_target[r["Salesperson"]] += r["Target"]

    # ── By product ─────────────────────────────
    prod_rev   = defaultdict(float)
    prod_units = defaultdict(int)
    for r in records:
        prod_rev[r["Product"]]   += r["Revenue"]
        prod_units[r["Product"]] += r["Units_Sold"]

    # ── By month ───────────────────────────────
    month_rev   = defaultdict(float)
    month_units = defaultdict(int)
    for r in records:
        month = datetime.strptime(r["Date"], "%Y-%m-%d").strftime("%b %Y")
        month_rev[month]   += r["Revenue"]
        month_units[month] += r["Units_Sold"]

    # ── Top performer ──────────────────────────
    top_person = max(sales_rev, key=sales_rev.get)

    # ── Best & worst region ────────────────────
    best_region  = max(region_rev, key=region_rev.get)
    worst_region = min(region_rev, key=region_rev.get)

    return {
        "total_revenue":  total_revenue,
        "total_units":    total_units,
        "total_target":   total_target,
        "achievement_pct": achievement,
        "num_records":    len(records),
        "region_rev":     dict(region_rev),
        "region_units":   dict(region_units),
        "sales_rev":      dict(sales_rev),
        "sales_units":    dict(sales_units),
        "sales_target":   dict(sales_target),
        "prod_rev":       dict(prod_rev),
        "prod_units":     dict(prod_units),
        "month_rev":      dict(month_rev),
        "month_units":    dict(month_units),
        "top_person":     top_person,
        "best_region":    best_region,
        "worst_region":   worst_region,
    }



#  CUSTOM PAGE TEMPLATE  (header + footer)


def make_page_template(canvas, doc):
    """Draw the header banner and footer on every page."""
    w, h = letter
    canvas.saveState()

    # ── Header banner ──────────────────────────
    canvas.setFillColor(BRAND_BLUE)
    canvas.rect(0, h - 0.65 * inch, w, 0.65 * inch, fill=1, stroke=0)

    canvas.setFillColor(WHITE)
    canvas.setFont("Helvetica-Bold", 14)
    canvas.drawString(0.5 * inch, h - 0.42 * inch, "SALES PERFORMANCE REPORT")

    canvas.setFont("Helvetica", 9)
    canvas.drawRightString(w - 0.5 * inch, h - 0.42 * inch,
                           f"Generated: {datetime.now().strftime('%B %d, %Y  %H:%M')}")

    # ── Accent stripe under banner ─────────────
    canvas.setFillColor(BRAND_ACCENT)
    canvas.rect(0, h - 0.72 * inch, w, 0.07 * inch, fill=1, stroke=0)

    # ── Footer ─────────────────────────────────
    canvas.setFillColor(GRAY_LIGHT)
    canvas.rect(0, 0, w, 0.45 * inch, fill=1, stroke=0)

    canvas.setFillColor(GRAY_MID)
    canvas.setFont("Helvetica", 8)
    canvas.drawString(0.5 * inch, 0.16 * inch, "Confidential — Internal Use Only")
    canvas.drawCentredString(w / 2, 0.16 * inch, "Automated Report Generator  |  ReportLab")
    canvas.drawRightString(w - 0.5 * inch, 0.16 * inch, f"Page {doc.page}")

    canvas.restoreState()



#  STYLE HELPERS


def build_styles() -> dict:
    base = getSampleStyleSheet()

    styles = {
        "title": ParagraphStyle("ReportTitle",
            fontSize=22, textColor=BRAND_BLUE, spaceAfter=4,
            fontName="Helvetica-Bold", alignment=TA_CENTER),

        "subtitle": ParagraphStyle("ReportSubtitle",
            fontSize=11, textColor=GRAY_MID, spaceAfter=20,
            fontName="Helvetica", alignment=TA_CENTER),

        "h1": ParagraphStyle("H1",
            fontSize=13, textColor=WHITE, spaceBefore=14, spaceAfter=4,
            fontName="Helvetica-Bold", leftIndent=8),

        "h2": ParagraphStyle("H2",
            fontSize=11, textColor=BRAND_BLUE, spaceBefore=10, spaceAfter=4,
            fontName="Helvetica-Bold"),

        "body": ParagraphStyle("Body",
            fontSize=9.5, textColor=GRAY_DARK, spaceAfter=6,
            fontName="Helvetica", leading=14),

        "kpi_label": ParagraphStyle("KpiLabel",
            fontSize=8, textColor=GRAY_MID, spaceAfter=1,
            fontName="Helvetica", alignment=TA_CENTER),

        "kpi_value": ParagraphStyle("KpiValue",
            fontSize=18, textColor=BRAND_BLUE, spaceAfter=0,
            fontName="Helvetica-Bold", alignment=TA_CENTER),

        "kpi_sub": ParagraphStyle("KpiSub",
            fontSize=8, textColor=GRAY_MID, spaceAfter=0,
            fontName="Helvetica", alignment=TA_CENTER),

        "insight": ParagraphStyle("Insight",
            fontSize=9, textColor=GRAY_DARK, spaceAfter=4,
            fontName="Helvetica", leftIndent=12, leading=14),

        "table_header": ParagraphStyle("TableHeader",
            fontSize=9, textColor=WHITE, fontName="Helvetica-Bold",
            alignment=TA_CENTER),

        "table_cell": ParagraphStyle("TableCell",
            fontSize=9, textColor=GRAY_DARK, fontName="Helvetica",
            alignment=TA_CENTER),
    }
    return styles



#  SECTION BUILDERS


def section_banner(title: str, styles: dict) -> list:
    """A full-width colored section header."""
    tbl = Table([[Paragraph(f"  {title}", styles["h1"])]],
                colWidths=[7.5 * inch])
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), BRAND_BLUE),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [BRAND_BLUE]),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
    ]))
    return [tbl, Spacer(1, 6)]


def kpi_cards(stats: dict, styles: dict) -> list:
    """Four KPI boxes in a 2×2 grid."""
    achievement = stats["achievement_pct"]
    ach_color   = BRAND_GREEN if achievement >= 100 else (
                  BRAND_ORANGE if achievement >= 80 else BRAND_RED)

    def card(label, value, sub=""):
        inner = Table([
            [Paragraph(label, styles["kpi_label"])],
            [Paragraph(value, styles["kpi_value"])],
            [Paragraph(sub,   styles["kpi_sub"])],
        ], colWidths=[1.75 * inch])
        inner.setStyle(TableStyle([
            ("ALIGN",         (0, 0), (-1, -1), "CENTER"),
            ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",    (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ]))
        return inner

    kpi1 = card("TOTAL REVENUE",
                f"${stats['total_revenue']:,.0f}",
                f"vs target ${stats['total_target']:,.0f}")
    kpi2 = card("UNITS SOLD",
                f"{stats['total_units']:,}",
                "across all regions")
    kpi3 = card("TARGET ACHIEVEMENT",
                f"{achievement:.1f}%",
                "of revenue goal")
    kpi4 = card("TRANSACTIONS",
                str(stats["num_records"]),
                "data records")

    # Style override for achievement colour
    ach_style = ParagraphStyle("AchVal",
        fontSize=18, textColor=ach_color, spaceAfter=0,
        fontName="Helvetica-Bold", alignment=TA_CENTER)
    kpi3 = Table([
        [Paragraph("TARGET ACHIEVEMENT", styles["kpi_label"])],
        [Paragraph(f"{achievement:.1f}%", ach_style)],
        [Paragraph("of revenue goal", styles["kpi_sub"])],
    ], colWidths=[1.75 * inch])
    kpi3.setStyle(TableStyle([
        ("ALIGN",  (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))

    row = Table([[kpi1, kpi2, kpi3, kpi4]], colWidths=[1.875 * inch] * 4)
    row.setStyle(TableStyle([
        ("BOX",         (0, 0), (-1, -1), 0.5, BRAND_ACCENT),
        ("INNERGRID",   (0, 0), (-1, -1), 0.5, BRAND_LIGHT),
        ("BACKGROUND",  (0, 0), (-1, -1), BRAND_LIGHT),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [BRAND_LIGHT]),
    ]))
    return [row, Spacer(1, 14)]


def regional_table(stats: dict, styles: dict) -> list:
    """Revenue & units by region."""
    header = ["Region", "Revenue ($)", "Units Sold", "Avg Unit Revenue ($)"]
    rows   = [header]

    for reg in sorted(stats["region_rev"]):
        rev   = stats["region_rev"][reg]
        units = stats["region_units"][reg]
        avg   = rev / units if units else 0
        rows.append([reg, f"{rev:,.2f}", f"{units:,}", f"{avg:.2f}"])

    tbl = Table(rows, colWidths=[1.6*inch, 1.8*inch, 1.8*inch, 2.3*inch])
    tbl.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1,  0), BRAND_BLUE),
        ("TEXTCOLOR",    (0, 0), (-1,  0), WHITE),
        ("FONTNAME",     (0, 0), (-1,  0), "Helvetica-Bold"),
        ("FONTSIZE",     (0, 0), (-1, -1), 9),
        ("ALIGN",        (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [WHITE, GRAY_LIGHT]),
        ("BOX",          (0, 0), (-1, -1), 0.5, BRAND_ACCENT),
        ("INNERGRID",    (0, 0), (-1, -1), 0.25, colors.HexColor("#BDC3C7")),
        ("TOPPADDING",   (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 7),
    ]))
    return [tbl, Spacer(1, 12)]


def salesperson_table(stats: dict, styles: dict) -> list:
    """Detailed per-salesperson breakdown."""
    header = ["Salesperson", "Revenue ($)", "Target ($)", "Achievement", "Units"]
    rows   = [header]

    for sp in sorted(stats["sales_rev"], key=lambda x: -stats["sales_rev"][x]):
        rev    = stats["sales_rev"][sp]
        tgt    = stats["sales_target"][sp]
        units  = stats["sales_units"][sp]
        ach    = (rev / tgt * 100) if tgt else 0
        rows.append([sp, f"{rev:,.2f}", f"{tgt:,.2f}", f"{ach:.1f}%", f"{units:,}"])

    tbl = Table(rows, colWidths=[1.8*inch, 1.6*inch, 1.6*inch, 1.3*inch, 1.2*inch])
    tbl.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1,  0), BRAND_BLUE),
        ("TEXTCOLOR",    (0, 0), (-1,  0), WHITE),
        ("FONTNAME",     (0, 0), (-1,  0), "Helvetica-Bold"),
        ("FONTSIZE",     (0, 0), (-1, -1), 9),
        ("ALIGN",        (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [WHITE, GRAY_LIGHT]),
        ("BOX",          (0, 0), (-1, -1), 0.5, BRAND_ACCENT),
        ("INNERGRID",    (0, 0), (-1, -1), 0.25, colors.HexColor("#BDC3C7")),
        ("TOPPADDING",   (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 7),
    ]))
    return [tbl, Spacer(1, 12)]


def product_table(stats: dict, styles: dict) -> list:
    """Revenue and units per product."""
    total_rev = stats["total_revenue"]
    header    = ["Product", "Revenue ($)", "Units Sold", "Revenue Share"]
    rows      = [header]

    for prod in sorted(stats["prod_rev"], key=lambda x: -stats["prod_rev"][x]):
        rev   = stats["prod_rev"][prod]
        units = stats["prod_units"][prod]
        share = (rev / total_rev * 100) if total_rev else 0
        rows.append([prod, f"{rev:,.2f}", f"{units:,}", f"{share:.1f}%"])

    tbl = Table(rows, colWidths=[2.2*inch, 1.8*inch, 1.6*inch, 1.9*inch])
    tbl.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1,  0), BRAND_BLUE),
        ("TEXTCOLOR",    (0, 0), (-1,  0), WHITE),
        ("FONTNAME",     (0, 0), (-1,  0), "Helvetica-Bold"),
        ("FONTSIZE",     (0, 0), (-1, -1), 9),
        ("ALIGN",        (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [WHITE, GRAY_LIGHT]),
        ("BOX",          (0, 0), (-1, -1), 0.5, BRAND_ACCENT),
        ("INNERGRID",    (0, 0), (-1, -1), 0.25, colors.HexColor("#BDC3C7")),
        ("TOPPADDING",   (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 7),
    ]))
    return [tbl, Spacer(1, 12)]


def bar_chart_revenue_by_region(stats: dict) -> list:
    """Vertical bar chart — revenue by region."""
    regions = sorted(stats["region_rev"].keys())
    values  = [stats["region_rev"][r] for r in regions]
    max_val = max(values) * 1.15

    drawing = Drawing(500, 200)

    bc = VerticalBarChart()
    bc.x            = 50
    bc.y            = 30
    bc.width        = 420
    bc.height       = 150
    bc.data         = [values]
    bc.strokeColor  = None
    bc.fillColor    = None
    bc.groupSpacing = 10
    bc.barSpacing   = 2

    bc.bars[0].fillColor   = BRAND_ACCENT
    bc.bars[0].strokeColor = None

    bc.valueAxis.valueMin       = 0
    bc.valueAxis.valueMax       = max_val
    bc.valueAxis.valueStep      = max_val / 5
    bc.valueAxis.labelTextFormat = "$%.0f"
    bc.valueAxis.labels.fontSize = 7
    bc.valueAxis.labels.fontName = "Helvetica"
    bc.valueAxis.gridStrokeColor = colors.HexColor("#E0E0E0")

    bc.categoryAxis.categoryNames  = regions
    bc.categoryAxis.labels.angle   = 0
    bc.categoryAxis.labels.dy      = -5
    bc.categoryAxis.labels.fontSize = 8
    bc.categoryAxis.labels.fontName = "Helvetica"
    bc.categoryAxis.strokeColor     = GRAY_MID

    drawing.add(bc)
    return [drawing, Spacer(1, 10)]


def pie_chart_product_mix(stats: dict) -> list:
    """Pie chart — revenue share by product."""
    products = list(stats["prod_rev"].keys())
    values   = [stats["prod_rev"][p] for p in products]

    drawing = Drawing(260, 160)

    pie = Pie()
    pie.x        = 30
    pie.y        = 20
    pie.width    = 130
    pie.height   = 130
    pie.data     = values
    pie.labels   = [f"{p}\n{v/sum(values)*100:.1f}%" for p, v in zip(products, values)]
    pie.sideLabels = True

    for i, col in enumerate(CHART_COLORS[:len(products)]):
        pie.slices[i].fillColor   = col
        pie.slices[i].strokeColor = WHITE
        pie.slices[i].strokeWidth = 1.5
        pie.slices[i].labelRadius = 1.35
        pie.slices[i].fontName    = "Helvetica"
        pie.slices[i].fontSize    = 7

    drawing.add(pie)
    return [drawing, Spacer(1, 10)]


def monthly_trend_chart(stats: dict) -> list:
    """Simple hand-drawn bar chart for monthly revenue trend."""
    months = list(stats["month_rev"].keys())
    values = [stats["month_rev"][m] for m in months]
    max_v  = max(values)

    chart_w = 440
    chart_h = 130
    left    = 60
    bottom  = 30
    bar_w   = (chart_w - 20) / len(months)

    drawing = Drawing(520, chart_h + bottom + 20)

    # Draw bars
    for i, (month, val) in enumerate(zip(months, values)):
        bh  = (val / max_v) * chart_h
        bx  = left + i * bar_w + bar_w * 0.1
        by  = bottom
        bw  = bar_w * 0.8
        col = BRAND_ACCENT if val == max_v else BRAND_LIGHT

        drawing.add(Rect(bx, by, bw, bh,
                         fillColor=col, strokeColor=BRAND_BLUE, strokeWidth=0.3))

        # Value label
        drawing.add(String(bx + bw / 2, by + bh + 3,
                           f"${val/1000:.0f}k",
                           fontSize=6, fontName="Helvetica",
                           fillColor=GRAY_DARK, textAnchor="middle"))

        # Month label
        drawing.add(String(bx + bw / 2, bottom - 12,
                           month,
                           fontSize=6.5, fontName="Helvetica",
                           fillColor=GRAY_DARK, textAnchor="middle"))

    # Axis lines
    drawing.add(Line(left, bottom, left + chart_w - 20, bottom,
                     strokeColor=GRAY_MID, strokeWidth=0.5))
    drawing.add(Line(left, bottom, left, bottom + chart_h,
                     strokeColor=GRAY_MID, strokeWidth=0.5))

    return [drawing, Spacer(1, 10)]


def insights_section(stats: dict, styles: dict) -> list:
    """Key auto-generated insights as bullet paragraphs."""
    ach = stats["achievement_pct"]
    ach_msg = (f"exceeded the revenue target by {ach-100:.1f}%"
               if ach >= 100 else f"achieved {ach:.1f}% of the revenue target")

    top = stats["top_person"]
    top_rev = stats["sales_rev"][top]
    top_tgt = stats["sales_target"][top]
    top_ach = (top_rev / top_tgt * 100) if top_tgt else 0

    best_prod = max(stats["prod_rev"], key=stats["prod_rev"].get)
    best_prod_share = stats["prod_rev"][best_prod] / stats["total_revenue"] * 100

    bullets = [
        f"<b>Overall Performance:</b> The team {ach_msg}, generating total revenue "
        f"of <b>${stats['total_revenue']:,.2f}</b> against a target of <b>${stats['total_target']:,.2f}</b>.",

        f"<b>Top Performer:</b> <b>{top}</b> led all salespeople with "
        f"<b>${top_rev:,.2f}</b> in revenue ({top_ach:.1f}% of personal target).",

        f"<b>Best Region:</b> The <b>{stats['best_region']}</b> region recorded the highest "
        f"revenue at <b>${stats['region_rev'][stats['best_region']]:,.2f}</b>. "
        f"The <b>{stats['worst_region']}</b> region had the lowest performance and "
        f"may benefit from additional support.",

        f"<b>Product Mix:</b> <b>{best_prod}</b> dominated the portfolio, accounting for "
        f"<b>{best_prod_share:.1f}%</b> of total revenue.",

        f"<b>Volume:</b> A total of <b>{stats['total_units']:,}</b> units were sold across "
        f"<b>{stats['num_records']}</b> transactions in the reporting period.",
    ]

    flowables = []
    for b in bullets:
        flowables.append(Paragraph(f"&#8226;  {b}", styles["insight"]))
        flowables.append(Spacer(1, 4))
    return flowables



#  MAIN REPORT BUILDER


def generate_report(input_csv: str, output_pdf: str) -> None:
    print(f"[1/4] Loading data from '{input_csv}' ...")
    records = load_data(input_csv)

    print("[2/4] Analysing data ...")
    stats = analyze(records)

    print(f"[3/4] Building PDF report → '{output_pdf}' ...")
    styles = build_styles()

    doc = SimpleDocTemplate(
        output_pdf,
        pagesize=letter,
        topMargin=0.85 * inch,
        bottomMargin=0.6  * inch,
        leftMargin=0.5  * inch,
        rightMargin=0.5  * inch,
        title="Sales Performance Report",
        author="Automated Report Generator",
    )

    story = []

    # ── Cover / title area ─────────────────────
    story.append(Spacer(1, 18))
    story.append(Paragraph("Sales Performance Report", styles["title"]))
    story.append(Paragraph(
        f"Reporting Period: January – June 2024  |  Generated {datetime.now().strftime('%B %d, %Y')}",
        styles["subtitle"]))
    story.append(HRFlowable(width="100%", thickness=1.5, color=BRAND_ACCENT, spaceAfter=12))

    # ── KPI cards ──────────────────────────────
    story += section_banner("Executive Summary — Key Metrics", styles)
    story += kpi_cards(stats, styles)

    # ── Key insights ───────────────────────────
    story += section_banner("Key Insights", styles)
    story += insights_section(stats, styles)
    story.append(Spacer(1, 8))

    # ── Revenue by region chart ────────────────
    story += section_banner("Revenue by Region", styles)
    story.append(Paragraph(
        "The chart below compares total revenue generated by each sales region "
        "over the reporting period.",
        styles["body"]))
    story += bar_chart_revenue_by_region(stats)
    story += regional_table(stats, styles)

    # ── Monthly trend ──────────────────────────
    story += section_banner("Monthly Revenue Trend", styles)
    story.append(Paragraph(
        "Month-over-month revenue performance. The highlighted bar represents the "
        "peak revenue month.",
        styles["body"]))
    story += monthly_trend_chart(stats)
    story.append(Spacer(1, 6))

    # ── Salesperson breakdown ──────────────────
    story += section_banner("Salesperson Performance", styles)
    story.append(Paragraph(
        "Individual salesperson results ranked by total revenue, including target "
        "achievement and unit volumes.",
        styles["body"]))
    story += salesperson_table(stats, styles)

    # ── Product analysis ───────────────────────
    story += section_banner("Product Analysis", styles)

    prod_row = Table(
        [[pie_chart_product_mix(stats)[0], Spacer(0.2*inch, 1),
          Table(
              [[""] + product_table(stats, styles)],
              colWidths=[0.1*inch, 6.9*inch]
          )]],
        colWidths=[2.7*inch, 0.3*inch, 4.5*inch],
    )
    prod_row.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))

    story.append(Paragraph(
        "Product revenue distribution and unit volumes for the period.",
        styles["body"]))
    story += product_table(stats, styles)

    # ── Appendix: raw data ─────────────────────
    story.append(PageBreak())
    story += section_banner("Appendix — Raw Transaction Data", styles)
    story.append(Paragraph(
        "All individual transaction records used to compile this report.",
        styles["body"]))
    story.append(Spacer(1, 6))

    # Reload to get original string values
    with open(input_csv, newline="", encoding="utf-8") as f:
        raw_rows = list(csv.reader(f))

    col_widths = [0.9*inch, 0.75*inch, 1.35*inch, 1.1*inch,
                  0.75*inch, 0.75*inch, 1.05*inch, 0.85*inch]
    data_tbl = Table(raw_rows, colWidths=col_widths, repeatRows=1)
    data_tbl.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1,  0), BRAND_BLUE),
        ("TEXTCOLOR",    (0, 0), (-1,  0), WHITE),
        ("FONTNAME",     (0, 0), (-1,  0), "Helvetica-Bold"),
        ("FONTSIZE",     (0, 0), (-1, -1), 7),
        ("ALIGN",        (0, 0), (-1, -1), "CENTER"),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [WHITE, GRAY_LIGHT]),
        ("BOX",          (0, 0), (-1, -1), 0.4, BRAND_ACCENT),
        ("INNERGRID",    (0, 0), (-1, -1), 0.2, colors.HexColor("#BDC3C7")),
        ("TOPPADDING",   (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
    ]))
    story.append(data_tbl)

    # ── Build ───────────────────────────────────
    doc.build(story, onFirstPage=make_page_template, onLaterPages=make_page_template)
    print(f"[4/4] Done! Report saved to '{output_pdf}'")



#  ENTRY POINT


if __name__ == "__main__":
    input_file  = sys.argv[1] if len(sys.argv) > 1 else "sales_data.csv"
    output_file = sys.argv[2] if len(sys.argv) > 2 else "sales_report.pdf"

    if not os.path.exists(input_file):
        print(f"Error: '{input_file}' not found.")
        sys.exit(1)

    generate_report(input_file, output_file)
