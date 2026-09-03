"""Generate a highly detailed vehicle inspection narrative via Groq LLM and export to Word (.docx)."""

from __future__ import annotations

import json
import os
from datetime import datetime
from io import BytesIO
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from docx.enum.style import WD_STYLE_TYPE

try:
    from groq import Groq
except ImportError:  # pragma: no cover
    Groq = None  # type: ignore


NA = "Not available from free sources"


def _v(val: Any, fallback: str = NA) -> str:
    if val is None:
        return fallback
    s = str(val).strip()
    return s if s else fallback


def _build_context(report: dict) -> str:
    """Flatten the VIN report into a clean context string for the LLM."""
    ident = report.get("identification") or {}
    engine = report.get("engine") or {}
    drive = report.get("drivetrain") or {}
    dims = report.get("dimensions") or {}
    mfg = report.get("manufacturing") or {}
    safety = report.get("safety") or {}
    manual = report.get("manual") or {}
    wmi = report.get("wmi") or {}
    cond = manual.get("condition") or {}

    recalls_text = []
    for r in (safety.get("recalls") or [])[:4]:
        summary = (r.get("summary") or "")[:120]
        recalls_text.append(
            f"- {r.get('campaign_number')}: {r.get('component') or 'N/A'} — {summary}"
        )
    recalls_block = "\n".join(recalls_text) if recalls_text else "No recalls for Year/Make/Model."

    notes = (_v(manual.get("notes"), "None") or "None")[:400]
    findings = (_v(manual.get("additional_findings"), "None") or "None")[:300]

    ctx = f"""VIN: {report.get('vin')}
Report: {report.get('id')} | Completeness: {report.get('data_completeness', 0)}%

IDENTIFICATION
Make/Model/Year: {_v(ident.get('make'))} {_v(ident.get('model'))} {_v(ident.get('model_year'))}
Trim/Series: {_v(ident.get('trim'))} / {_v(ident.get('series'))}
Body/Type/Class: {_v(ident.get('body_class'))} / {_v(ident.get('vehicle_type'))} / {_v(ident.get('vehicle_class'))}
Manufacturer: {_v(ident.get('manufacturer'))}

ENGINE & DRIVETRAIN
Engine: {_v(engine.get('engine_model'))} | {_v(engine.get('displacement'))} | {_v(engine.get('cylinders'))} cyl | {_v(engine.get('configuration'))}
Fuel/HP: {_v(engine.get('fuel_type'))} | {_v(engine.get('horsepower'))} hp | Mfr: {_v(engine.get('manufacturer'))}
Trans/Drive: {_v(drive.get('transmission'))} ({_v(drive.get('transmission_speeds'))} spd) | {_v(drive.get('drive_type'))}
Doors/Seats/GVWR: {_v(dims.get('doors'))} / {_v(dims.get('seats'))} / {_v(dims.get('gvwr'))}

MANUFACTURING
Plant: {_v(mfg.get('plant_city'))}, {_v(mfg.get('plant_state'))}, {_v(mfg.get('plant_country'))} ({_v(mfg.get('plant_company'))})
WMI: {_v(wmi.get('code'))} | {_v(wmi.get('manufacturer'))} | {_v(wmi.get('country'))}

RECALLS (Year/Make/Model — verify VIN at nhtsa.gov/recalls)
{recalls_block}

INSPECTOR DATA
Client: {_v(manual.get('client_name'), '—')} | Ref: {_v(manual.get('client_reference'), '—')}
Reg/Mileage/Price: {_v(manual.get('registration'), '—')} | {_v(manual.get('mileage'), '—')} | {_v(manual.get('purchase_price'), '—')}
Date/Inspector: {_v(manual.get('inspection_date'), '—')} | {_v(manual.get('inspector_name'), '—')}
Condition overall/ext/int/eng/trans/elec/tires/brakes: {cond.get('overall','n/a')} / {cond.get('exterior','n/a')} / {cond.get('interior','n/a')} / {cond.get('engine','n/a')} / {cond.get('transmission','n/a')} / {cond.get('electrical','n/a')} / {cond.get('tires','n/a')} / {cond.get('brakes','n/a')}
Notes: {notes}
Findings: {findings}
"""
    return ctx.strip()


SYSTEM_PROMPT = """You are an expert automotive inspector and technical writer.
Write a detailed, professional vehicle inspection report from the supplied VIN data only.

RULES:
- Never invent accidents, ownership, title brands, odometer, or service history.
- If data is missing, say it is not available from free public sources.
- Treat recalls as Year/Make/Model campaigns; remind to verify exact VIN at nhtsa.gov/recalls.
- Incorporate inspector notes/ratings when present; label them as inspector observations.
- Use clear professional English. Do not use code fences.

Structure EXACTLY with these ## headings:
## Executive Summary
## Vehicle Identification
## Technical Specifications
## Manufacturing Origin
## Safety & Recall Analysis
## Condition Assessment (Inspector Observations)
## Recommended Inspection Checklist
## Known Common Issues for this Model
## Limitations & Disclaimer
## Closing Recommendation

Checklist: comprehensive bullets for exterior, underbody, engine bay, interior, fluids, electronics, tires/brakes, road test — tailored to class/age.
Common issues: only well-known patterns for that generation, or say research is needed.
Be thorough and factual.
"""


# Preferred order when preferred model fails (rate limit, 413, decommissioned, etc.)
# Updated Sep 2026 — older Llama 3.x / Mixtral / Gemma IDs are decommissioned on free/dev tier.
# See: https://console.groq.com/docs/deprecations
FALLBACK_MODELS = [
    "openai/gpt-oss-120b",
    "openai/gpt-oss-20b",
    "qwen/qwen3.6-27b",
    "qwen/qwen3.8-27b",
]


def generate_inspection_narrative(
    report: dict,
    api_key: str | None = None,
    model: str = "openai/gpt-oss-120b",
    fallback: bool = True,
) -> tuple[str, str]:
    """Call Groq to produce a detailed inspection narrative from the VIN report.

    Returns (narrative_text, model_used).
    If ``fallback`` is True, tries FALLBACK_MODELS (preferred model first) on failure.
    """
    if Groq is None:
        raise RuntimeError("groq package is not installed. Run: pip install groq")

    key = api_key or os.environ.get("GROQ_API_KEY")
    if not key:
        raise ValueError(
            "Groq API key is required. Set GROQ_API_KEY environment variable "
            "or enter it in the Streamlit UI."
        )

    client = Groq(api_key=key)
    context = _build_context(report)
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {
            "role": "user",
            "content": (
                "Produce a detailed professional vehicle inspection report "
                "from this VIN-decoded and inspector data:\n\n" + context
            ),
        },
    ]

    # Preferred model first, then remaining fallbacks without duplicates
    candidates: list[str] = []
    for m in [model] + (FALLBACK_MODELS if fallback else []):
        if m and m not in candidates:
            candidates.append(m)

    errors: list[str] = []
    for candidate in candidates:
        try:
            response = client.chat.completions.create(
                model=candidate,
                messages=messages,
                temperature=0.25,
                max_tokens=4096,
            )
            content = response.choices[0].message.content
            if not content:
                errors.append(f"{candidate}: empty response")
                continue
            return content.strip(), candidate
        except Exception as e:
            errors.append(f"{candidate}: {e}")
            continue

    detail = " | ".join(errors) if errors else "unknown error"
    raise RuntimeError(f"All Groq models failed. Attempts: {detail}")


def _set_run_font(run, name: str = "Calibri", size_pt: int = 11, bold: bool = False, color: RGBColor | None = None):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size_pt)
    run.bold = bold
    if color:
        run.font.color.rgb = color


def build_inspection_docx(
    report: dict,
    narrative_md: str,
    branding: dict | None = None,
) -> bytes:
    """Turn the Groq Markdown narrative + report metadata into a polished .docx."""
    branding = branding or {}
    doc = Document()

    # Narrow margins
    for section in doc.sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.7)
        section.left_margin = Inches(0.85)
        section.right_margin = Inches(0.85)

    styles = doc.styles
    # Title style tweak
    title_style = styles["Title"]
    title_style.font.name = "Calibri"
    title_style.font.size = Pt(22)
    title_style.font.color.rgb = RGBColor(0x1E, 0x3A, 0x5F)

    heading1 = styles["Heading 1"]
    heading1.font.name = "Calibri"
    heading1.font.size = Pt(14)
    heading1.font.color.rgb = RGBColor(0x1E, 0x3A, 0x5F)
    heading1.font.bold = True

    heading2 = styles["Heading 2"]
    heading2.font.name = "Calibri"
    heading2.font.size = Pt(12)
    heading2.font.color.rgb = RGBColor(0x2C, 0x52, 0x7A)

    normal = styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)

    ident = report.get("identification") or {}
    vehicle_title = (
        " ".join(x for x in [ident.get("model_year"), ident.get("make"), ident.get("model")] if x)
        or "Vehicle Inspection Report"
    )
    company = branding.get("company") or branding.get("name") or "VIN Vehicle Report"

    # Header block
    p = doc.add_paragraph()
    run = p.add_run((branding.get("default_report_title") or "DETAILED VEHICLE INSPECTION REPORT").upper())
    _set_run_font(run, size_pt=18, bold=True, color=RGBColor(0x1E, 0x3A, 0x5F))
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    p = doc.add_paragraph()
    run = p.add_run(company)
    _set_run_font(run, size_pt=11, bold=True)

    contact_bits = [x for x in [branding.get("email"), branding.get("phone"), branding.get("website")] if x]
    if contact_bits:
        p = doc.add_paragraph()
        run = p.add_run("  ·  ".join(contact_bits))
        _set_run_font(run, size_pt=9, color=RGBColor(0x55, 0x55, 0x55))

    doc.add_paragraph()

    p = doc.add_paragraph()
    run = p.add_run(vehicle_title)
    _set_run_font(run, size_pt=14, bold=True)

    p = doc.add_paragraph()
    run = p.add_run(f"VIN: {report.get('vin', '')}")
    _set_run_font(run, size_pt=11, bold=True)

    meta = (
        f"Report No: {report.get('id', '')}   |   "
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}   |   "
        f"Data completeness: {report.get('data_completeness', 0)}%"
    )
    p = doc.add_paragraph()
    run = p.add_run(meta)
    _set_run_font(run, size_pt=9, color=RGBColor(0x55, 0x55, 0x55))

    # Horizontal rule via bottom border on empty paragraph is complex; use spacer
    doc.add_paragraph()

    # Parse simple Markdown-ish narrative into paragraphs / headings / bullets
    lines = narrative_md.splitlines()
    for raw in lines:
        line = raw.rstrip()
        if not line.strip():
            continue

        if line.startswith("## "):
            text = line[3:].strip()
            doc.add_heading(text, level=1)
            continue
        if line.startswith("### "):
            text = line[4:].strip()
            doc.add_heading(text, level=2)
            continue

        # Bullets
        if line.lstrip().startswith(("- ", "* ", "• ")):
            text = line.lstrip()[2:].strip()
            p = doc.add_paragraph(style="List Bullet")
            run = p.add_run(text)
            _set_run_font(run, size_pt=10)
            continue

        # Numbered
        stripped = line.lstrip()
        if len(stripped) > 2 and stripped[0].isdigit() and stripped[1] in ".)":
            text = stripped[2:].strip() if stripped[1] == "." else stripped[stripped.find(" ")+1:].strip()
            p = doc.add_paragraph(style="List Number")
            run = p.add_run(text)
            _set_run_font(run, size_pt=10)
            continue

        # Bold markers **text**
        p = doc.add_paragraph()
        parts = line.split("**")
        for i, part in enumerate(parts):
            run = p.add_run(part)
            _set_run_font(run, size_pt=10, bold=(i % 2 == 1))

    # Footer disclaimer block
    doc.add_paragraph()
    p = doc.add_paragraph()
    run = p.add_run("Important Limitations")
    _set_run_font(run, size_pt=11, bold=True, color=RGBColor(0x1E, 0x3A, 0x5F))

    limits = [
        "This report is generated from free public NHTSA data and optional inspector notes. It does not include ownership history, accident records, title brands, odometer history, service records, or auction data.",
        "Absence of information does not confirm a clean history.",
        "Always verify open recalls at https://www.nhtsa.gov/recalls using the exact VIN.",
        "This document is not a substitute for a professional mechanical inspection or a paid vehicle-history report.",
    ]
    for lim in limits:
        p = doc.add_paragraph(style="List Bullet")
        run = p.add_run(lim)
        _set_run_font(run, size_pt=8, color=RGBColor(0x44, 0x44, 0x44))

    disc = branding.get("disclaimer") or (
        "Information in this report is compiled from available external data sources and user-provided information. "
        "Availability and accuracy depend on the underlying sources. This report is not a substitute for a professional "
        "mechanical inspection, official title verification, or a comprehensive paid vehicle-history report."
    )
    p = doc.add_paragraph()
    run = p.add_run("Disclaimer: ")
    _set_run_font(run, size_pt=8, bold=True, color=RGBColor(0x44, 0x44, 0x44))
    run = p.add_run(disc)
    _set_run_font(run, size_pt=8, color=RGBColor(0x44, 0x44, 0x44))

    footer_text = branding.get("report_footer") or "Confidential – For client use only"
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(footer_text)
    _set_run_font(run, size_pt=8, color=RGBColor(0x66, 0x66, 0x66))

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()
