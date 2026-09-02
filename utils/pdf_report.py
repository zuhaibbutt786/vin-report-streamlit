"""Professional PDF report generator using ReportLab."""

from __future__ import annotations

from datetime import datetime
from io import BytesIO
from typing import Any

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
)

NA = "Not available from free sources"
HEADER_BG = colors.Color(30 / 255, 58 / 255, 95 / 255)


def _v(val: Any, fallback: str = NA) -> str:
    if val is None:
        return fallback
    s = str(val).strip()
    return s if s else fallback


def _kv_table(pairs: list[tuple[str, str]], col_widths=None) -> Table:
    data = [[k, v] for k, v in pairs]
    t = Table(data, colWidths=col_widths or [55 * mm, 120 * mm])
    t.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("TEXTCOLOR", (0, 0), (0, -1), colors.Color(0.3, 0.3, 0.3)),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("LINEBELOW", (0, 0), (-1, -2), 0.3, colors.Color(0.85, 0.85, 0.85)),
            ]
        )
    )
    return t


def build_pdf(report: dict, branding: dict) -> bytes:
    buf = BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=14 * mm,
        rightMargin=14 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "Title",
        parent=styles["Heading1"],
        fontSize=18,
        textColor=HEADER_BG,
        spaceAfter=6,
    )
    h2 = ParagraphStyle(
        "H2",
        parent=styles["Heading2"],
        fontSize=12,
        textColor=HEADER_BG,
        spaceBefore=12,
        spaceAfter=6,
    )
    body = ParagraphStyle("Body", parent=styles["Normal"], fontSize=9, leading=12)
    small = ParagraphStyle(
        "Small", parent=styles["Normal"], fontSize=8, textColor=colors.Color(0.35, 0.35, 0.35)
    )
    disclaimer = ParagraphStyle(
        "Disc",
        parent=styles["Normal"],
        fontSize=8,
        leading=11,
        textColor=colors.Color(0.25, 0.25, 0.25),
    )

    ident = report.get("identification") or {}
    engine = report.get("engine") or {}
    drive = report.get("drivetrain") or {}
    dims = report.get("dimensions") or {}
    mfg = report.get("manufacturing") or {}
    safety = report.get("safety") or {}
    manual = report.get("manual") or {}
    wmi = report.get("wmi") or {}
    sources = report.get("sources") or []

    vehicle_title = (
        " ".join(
            x for x in [ident.get("model_year"), ident.get("make"), ident.get("model")] if x
        )
        or "Vehicle Report"
    )

    story = []

    company = branding.get("company") or branding.get("name") or "VIN Vehicle Report"
    story.append(
        Paragraph(_v(branding.get("default_report_title"), "VEHICLE VIN REPORT"), title_style)
    )
    story.append(Paragraph(company, body))
    contact = "  \u00b7  ".join(x for x in [branding.get("email"), branding.get("phone")] if x)
    if contact:
        story.append(Paragraph(contact, small))
    story.append(Spacer(1, 8 * mm))
    story.append(
        Paragraph(
            vehicle_title,
            ParagraphStyle("VT", parent=styles["Heading1"], fontSize=14),
        )
    )
    story.append(Paragraph(f"VIN: <b>{report.get('vin', '')}</b>", body))
    story.append(Paragraph(f"Report No: {report.get('id', '')}", small))
    story.append(Paragraph(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}", small))
    story.append(Spacer(1, 6 * mm))

    eng_bits = [
        f"{engine['cylinders']} cyl" if engine.get("cylinders") else None,
        engine.get("displacement"),
        f"{engine['horsepower']} hp" if engine.get("horsepower") else None,
    ]
    story.append(
        _kv_table(
            [
                ("Make", _v(ident.get("make"))),
                ("Model", _v(ident.get("model"))),
                ("Model Year", _v(ident.get("model_year"))),
                ("Trim", _v(ident.get("trim"))),
                ("Body Class", _v(ident.get("body_class"))),
                ("Drive Type", _v(drive.get("drive_type"))),
                ("Fuel Type", _v(engine.get("fuel_type"))),
                ("Engine", _v(" \u00b7 ".join(x for x in eng_bits if x) or None)),
                ("Data Completeness", f"{report.get('data_completeness', 0)}%"),
                (
                    "Overall Condition",
                    (manual.get("condition") or {})
                    .get("overall", "not_assessed")
                    .replace("_", " "),
                ),
            ]
        )
    )
    story.append(Spacer(1, 4 * mm))
    story.append(
        Paragraph(
            "Data completeness indicates how many requested fields were populated. It does not guarantee factual accuracy.",
            small,
        )
    )

    story.append(Paragraph("Vehicle Identification", h2))
    story.append(
        _kv_table(
            [
                ("VIN", report.get("vin", "")),
                ("Make", _v(ident.get("make"))),
                ("Manufacturer", _v(ident.get("manufacturer"))),
                ("Model", _v(ident.get("model"))),
                ("Model Year", _v(ident.get("model_year"))),
                ("Trim", _v(ident.get("trim"))),
                ("Series", _v(ident.get("series"))),
                ("Vehicle Type", _v(ident.get("vehicle_type"))),
                ("Body Class", _v(ident.get("body_class"))),
                ("Vehicle Class", _v(ident.get("vehicle_class"))),
            ]
        )
    )

    story.append(Paragraph("WMI-Derived Information", h2))
    story.append(
        _kv_table(
            [
                ("WMI Code", _v(wmi.get("code"), "\u2014")),
                ("Manufacturer (WMI)", _v(wmi.get("manufacturer"))),
                ("Country (WMI)", _v(wmi.get("country"))),
                ("Region (WMI)", _v(wmi.get("region"))),
            ]
        )
    )

    story.append(PageBreak())
    story.append(Paragraph("Engine & Drivetrain", h2))
    story.append(
        _kv_table(
            [
                ("Engine Model", _v(engine.get("engine_model"))),
                ("Displacement", _v(engine.get("displacement"))),
                ("Cylinders", _v(engine.get("cylinders"))),
                ("Configuration", _v(engine.get("configuration"))),
                ("Fuel Type", _v(engine.get("fuel_type"))),
                ("Horsepower", _v(engine.get("horsepower"))),
                ("Engine Manufacturer", _v(engine.get("manufacturer"))),
                ("Transmission", _v(drive.get("transmission"))),
                ("Transmission Speeds", _v(drive.get("transmission_speeds"))),
                ("Drive Type", _v(drive.get("drive_type"))),
                ("Doors", _v(dims.get("doors"))),
                ("Seats", _v(dims.get("seats"))),
                ("GVWR", _v(dims.get("gvwr"))),
            ]
        )
    )

    story.append(Paragraph("Manufacturing", h2))
    story.append(
        _kv_table(
            [
                ("Plant Country", _v(mfg.get("plant_country"))),
                ("Plant State", _v(mfg.get("plant_state"))),
                ("Plant City", _v(mfg.get("plant_city"))),
                ("Plant Company", _v(mfg.get("plant_company"))),
            ]
        )
    )

    story.append(Paragraph("Safety / NHTSA Recalls", h2))
    recalls = safety.get("recalls") or []
    if not safety.get("recalls_available") or not recalls:
        story.append(
            Paragraph(
                _v(
                    safety.get("recalls_note"),
                    "Recall information was not available from the connected free source.",
                ),
                body,
            )
        )
        story.append(
            Paragraph(
                "Always verify open recalls at https://www.nhtsa.gov/recalls using the exact VIN.",
                small,
            )
        )
    else:
        story.append(
            Paragraph(
                f"{len(recalls)} recall campaign(s) found for this Year / Make / Model.",
                body,
            )
        )
        for r in recalls[:10]:
            story.append(Spacer(1, 3 * mm))
            story.append(
                Paragraph(
                    f"<b>{r.get('campaign_number', '')}</b> \u2014 {_v(r.get('component'), 'Component N/A')}",
                    body,
                )
            )
            if r.get("summary"):
                story.append(Paragraph(r["summary"], small))
            if r.get("remedy"):
                story.append(Paragraph(f"<b>Remedy:</b> {r['remedy']}", small))

    story.append(PageBreak())
    story.append(Paragraph("Manual Inspection & Client Information (User-provided)", h2))
    cond = manual.get("condition") or {}
    story.append(
        _kv_table(
            [
                ("Client Name", _v(manual.get("client_name"), "\u2014")),
                ("Client Reference", _v(manual.get("client_reference"), "\u2014")),
                ("Registration", _v(manual.get("registration"), "\u2014")),
                ("Mileage", _v(manual.get("mileage"), "\u2014")),
                ("Purchase Price", _v(manual.get("purchase_price"), "\u2014")),
                ("Inspection Date", _v(manual.get("inspection_date"), "\u2014")),
                ("Inspector", _v(manual.get("inspector_name"), "\u2014")),
                ("Overall", cond.get("overall", "not_assessed").replace("_", " ")),
                ("Exterior", cond.get("exterior", "not_inspected").replace("_", " ")),
                ("Interior", cond.get("interior", "not_inspected").replace("_", " ")),
                ("Engine", cond.get("engine", "not_inspected").replace("_", " ")),
                ("Transmission", cond.get("transmission", "not_inspected").replace("_", " ")),
                ("Electrical", cond.get("electrical", "not_inspected").replace("_", " ")),
                ("Tires", cond.get("tires", "not_inspected").replace("_", " ")),
                ("Brakes", cond.get("brakes", "not_inspected").replace("_", " ")),
            ]
        )
    )
    if manual.get("notes"):
        story.append(Paragraph("Notes", h2))
        story.append(Paragraph(manual["notes"], body))
    if manual.get("additional_findings"):
        story.append(Paragraph("Additional Findings", h2))
        story.append(Paragraph(manual["additional_findings"], body))

    story.append(Paragraph("Data Sources", h2))
    for s in sources:
        status = "OK" if s.get("success") else "unavailable"
        story.append(Paragraph(f"<b>{s.get('name', '')}</b> ({status})", body))
        if s.get("url"):
            story.append(Paragraph(s["url"], small))
        if s.get("error"):
            story.append(Paragraph(s["error"], small))

    story.append(Paragraph("Important Limitations", h2))
    limits = [
        "This report uses free public data primarily from the U.S. National Highway Traffic Safety Administration (NHTSA).",
        "It does NOT include ownership history, accident records, title brands, odometer readings, service history, auction history, or insurance claims.",
        "Missing information does not mean the vehicle has a clean history.",
        'Never assume "no accidents", "clean title", or "one owner" unless verified by an authoritative paid source.',
        "Always verify open recalls at nhtsa.gov/recalls using the exact VIN.",
        "User-provided information is not independently verified by this tool.",
    ]
    for line in limits:
        story.append(Paragraph(f"\u2022 {line}", disclaimer))

    story.append(Paragraph("Disclaimer", h2))
    story.append(
        Paragraph(
            branding.get("disclaimer")
            or (
                "Information in this report is compiled from available external data sources and user-provided information. "
                "Availability and accuracy depend on the underlying sources. The absence of information does not confirm that an event did not occur. "
                "This report is not a substitute for a professional mechanical inspection, official title verification, or a comprehensive paid vehicle-history report."
            ),
            disclaimer,
        )
    )

    footer_text = branding.get("report_footer") or "Confidential \u2013 For client use only"

    def _footer(canvas, doc_):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.Color(0.45, 0.45, 0.45))
        canvas.drawString(14 * mm, 10 * mm, footer_text)
        canvas.drawRightString(A4[0] - 14 * mm, 10 * mm, f"Page {doc_.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    return buf.getvalue()
