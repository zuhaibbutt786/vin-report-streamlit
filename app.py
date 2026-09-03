"""
VIN Vehicle Report Generator — Streamlit (personal / offline-friendly tool)
Decode VIN via free NHTSA APIs → review → add inspection notes → download PDF
"""

from __future__ import annotations

import json
import random
from datetime import datetime
from pathlib import Path

import streamlit as st

from utils.vin_validate import validate_vin, clean_vin
from utils.nhtsa import decode_vin, fetch_recalls, lookup_wmi
from utils.pdf_report import build_pdf
from utils.groq_inspection import generate_inspection_narrative, build_inspection_docx

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
HISTORY_FILE = DATA_DIR / "report_history.json"
SETTINGS_FILE = DATA_DIR / "settings.json"

DEFAULT_SETTINGS = {
    "name": "",
    "company": "",
    "phone": "",
    "email": "",
    "website": "",
    "address": "",
    "report_footer": "Confidential – For client use only",
    "default_report_title": "Vehicle VIN Report",
    "disclaimer": (
        "Information in this report is compiled from available external data sources and user-provided information. "
        "Availability and accuracy depend on the underlying sources. The absence of information does not confirm that "
        "an event did not occur. This report is not a substitute for a professional mechanical inspection, official "
        "title verification, or a comprehensive paid vehicle-history report."
    ),
}

NA = "Not available from free sources"


def load_json(path: Path, default):
    try:
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        pass
    return default


def save_json(path: Path, data) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def load_settings() -> dict:
    return {**DEFAULT_SETTINGS, **load_json(SETTINGS_FILE, {})}


def save_settings(s: dict) -> None:
    save_json(SETTINGS_FILE, s)


def load_history() -> list:
    return load_json(HISTORY_FILE, [])


def save_report_to_history(report: dict) -> None:
    history = load_history()
    history = [h for h in history if h.get("id") != report.get("id")]
    label = " ".join(
        x
        for x in [
            (report.get("identification") or {}).get("model_year"),
            (report.get("identification") or {}).get("make"),
            (report.get("identification") or {}).get("model"),
        ]
        if x
    ) or "Unknown Vehicle"
    history.insert(
        0,
        {
            "id": report["id"],
            "vin": report["vin"],
            "label": label,
            "created_at": report.get("created_at"),
            "report": report,
        },
    )
    save_json(HISTORY_FILE, history[:100])


def new_report_id() -> str:
    now = datetime.now()
    return f"VR-{now.strftime('%Y%m%d')}-{random.randint(100, 999)}"


def calc_completeness(report: dict) -> int:
    ident = report.get("identification") or {}
    engine = report.get("engine") or {}
    drive = report.get("drivetrain") or {}
    dims = report.get("dimensions") or {}
    mfg = report.get("manufacturing") or {}
    fields = [
        ident.get("make"),
        ident.get("model"),
        ident.get("model_year"),
        ident.get("trim"),
        ident.get("body_class"),
        engine.get("displacement"),
        engine.get("cylinders"),
        engine.get("fuel_type"),
        drive.get("drive_type"),
        dims.get("doors"),
        mfg.get("plant_country"),
        mfg.get("plant_city"),
    ]
    filled = sum(1 for v in fields if v and str(v).strip())
    return round(filled / len(fields) * 100) if fields else 0


def empty_manual() -> dict:
    return {
        "client_name": "",
        "client_reference": "",
        "registration": "",
        "mileage": "",
        "purchase_price": "",
        "inspection_date": "",
        "inspector_name": "",
        "notes": "",
        "additional_findings": "",
        "condition": {
            "overall": "not_assessed",
            "exterior": "not_inspected",
            "interior": "not_inspected",
            "engine": "not_inspected",
            "transmission": "not_inspected",
            "electrical": "not_inspected",
            "tires": "not_inspected",
            "brakes": "not_inspected",
        },
    }


def build_report(vin: str) -> dict:
    now = datetime.now().isoformat(timespec="seconds")
    wmi = lookup_wmi(vin)
    sources = []
    errors = []

    vpic = decode_vin(vin)
    sources.append(
        {
            "name": "NHTSA vPIC",
            "url": "https://vpic.nhtsa.dot.gov/api/",
            "success": vpic["success"],
            "error": vpic.get("error"),
            "retrieved_at": now,
        }
    )
    if not vpic["success"] and vpic.get("error"):
        errors.append(vpic["error"])

    report = {
        "id": new_report_id(),
        "vin": vin,
        "created_at": now,
        "updated_at": now,
        "identification": vpic.get("identification") or {},
        "engine": vpic.get("engine") or {},
        "drivetrain": vpic.get("drivetrain") or {},
        "dimensions": vpic.get("dimensions") or {},
        "manufacturing": vpic.get("manufacturing") or {},
        "wmi": wmi,
        "safety": {
            "recalls": [],
            "recalls_available": False,
            "recalls_note": "Recall information was not available from the connected free source.",
        },
        "sources": sources,
        "data_completeness": 0,
        "manual": empty_manual(),
        "errors": errors,
    }

    if not report["manufacturing"].get("plant_country") and wmi.get("country"):
        report["manufacturing"]["plant_country"] = wmi["country"]

    make = report["identification"].get("make")
    model = report["identification"].get("model")
    year = report["identification"].get("model_year")

    if make and model and year:
        recalls = fetch_recalls(make, model, year)
        sources.append(
            {
                "name": "NHTSA Recalls",
                "url": "https://api.nhtsa.gov/recalls/recallsByVehicle",
                "success": recalls["success"],
                "error": recalls.get("error"),
                "retrieved_at": datetime.now().isoformat(timespec="seconds"),
            }
        )
        if recalls["success"]:
            report["safety"]["recalls"] = recalls["items"]
            report["safety"]["recalls_available"] = True
            if recalls["count"] == 0:
                report["safety"]["recalls_note"] = (
                    "No safety recalls found for this Year / Make / Model in the NHTSA database. "
                    "Always verify with the exact VIN at nhtsa.gov/recalls."
                )
            else:
                report["safety"]["recalls_note"] = None
        else:
            report["safety"]["recalls_note"] = recalls.get("error") or report["safety"]["recalls_note"]

    report["sources"] = sources
    report["data_completeness"] = calc_completeness(report)
    return report


def field(label: str, value) -> None:
    display = value if value and str(value).strip() else NA
    st.markdown(f"**{label}**  \n{display}")


st.set_page_config(
    page_title="VIN Vehicle Report Generator",
    page_icon="\U0001F697",
    layout="wide",
    initial_sidebar_state="expanded",
)

if "report" not in st.session_state:
    st.session_state.report = None
if "page" not in st.session_state:
    st.session_state.page = "home"

settings = load_settings()

with st.sidebar:
    st.title("VIN Report")
    st.caption("Personal tool · Free NHTSA data")
    if st.button("\U0001F3E0 Dashboard", use_container_width=True):
        st.session_state.page = "home"
        st.rerun()
    if st.button("\U0001F50E New VIN Decode", use_container_width=True):
        st.session_state.page = "decode"
        st.rerun()
    if st.button("\u2699\uFE0F Settings", use_container_width=True):
        st.session_state.page = "settings"
        st.rerun()
    st.divider()
    st.markdown(
        """
**Data limits**

Does **not** include accidents, title brands, ownership, or odometer history.

Missing data ≠ clean history.
"""
    )

if st.session_state.page == "home":
    st.header("VIN Vehicle Report Generator")
    st.write(
        "Decode a VIN with free public NHTSA data, add your inspection notes, and download a professional PDF."
    )

    col1, col2 = st.columns([1, 3])
    with col1:
        if st.button("Enter VIN →", type="primary", use_container_width=True):
            st.session_state.page = "decode"
            st.rerun()

    history = load_history()
    st.subheader("Recent reports")
    if not history:
        st.info("No reports yet. Create your first report above.")
    else:
        for h in history[:20]:
            c1, c2, c3, c4 = st.columns([2, 3, 3, 1])
            c1.write(h.get("id", ""))
            c2.write(h.get("label", ""))
            c3.code(h.get("vin", ""))
            if c4.button("Open", key=f"open_{h['id']}"):
                st.session_state.report = h.get("report")
                st.session_state.page = "report"
                st.rerun()

elif st.session_state.page == "decode":
    st.header("Enter Vehicle Identification Number")
    vin_input = st.text_input(
        "VIN",
        max_chars=17,
        placeholder="1HGCM82633A004352",
        help="17 characters. Letters I, O, Q are not used in VINs.",
    )
    sample = st.selectbox(
        "Sample VINs (optional)",
        ["", "1HGCM82633A004352", "5YJSA1E14HF000001", "1FTFW1ET5DFC10312"],
    )
    if sample:
        vin_input = sample

    if st.button("Decode VIN", type="primary"):
        result = validate_vin(vin_input)
        if not result["valid"]:
            st.error(result.get("error") or "Invalid VIN")
        else:
            if result.get("warning"):
                st.warning(result["warning"])
            with st.spinner("Retrieving vehicle information from NHTSA…"):
                report = build_report(result["cleaned"])
            save_report_to_history(report)
            st.session_state.report = report
            st.session_state.page = "report"
            st.rerun()

elif st.session_state.page == "settings":
    st.header("Settings / Branding")
    st.caption("Appears on generated PDFs. Saved locally on this computer.")
    s = load_settings()
    with st.form("settings_form"):
        c1, c2 = st.columns(2)
        s["name"] = c1.text_input("Your Name", s.get("name", ""))
        s["company"] = c2.text_input("Company", s.get("company", ""))
        s["phone"] = c1.text_input("Phone", s.get("phone", ""))
        s["email"] = c2.text_input("Email", s.get("email", ""))
        s["website"] = c1.text_input("Website", s.get("website", ""))
        s["address"] = c2.text_input("Address", s.get("address", ""))
        s["default_report_title"] = st.text_input(
            "Default Report Title", s.get("default_report_title", "")
        )
        s["report_footer"] = st.text_input("Report Footer", s.get("report_footer", ""))
        s["disclaimer"] = st.text_area("Disclaimer", s.get("disclaimer", ""), height=120)
        if st.form_submit_button("Save Settings", type="primary"):
            save_settings(s)
            st.success("Settings saved.")

elif st.session_state.page == "report":
    report = st.session_state.report
    if not report:
        st.warning("No report loaded.")
        st.session_state.page = "home"
        st.rerun()

    ident = report.get("identification") or {}
    engine = report.get("engine") or {}
    drive = report.get("drivetrain") or {}
    dims = report.get("dimensions") or {}
    mfg = report.get("manufacturing") or {}
    safety = report.get("safety") or {}
    wmi = report.get("wmi") or {}
    manual = report.get("manual") or empty_manual()

    title = (
        " ".join(x for x in [ident.get("model_year"), ident.get("make"), ident.get("model")] if x)
        or "Vehicle Report"
    )
    st.header(title)
    st.code(report.get("vin", ""))
    st.caption(f"Report {report.get('id')} · Completeness {report.get('data_completeness', 0)}%")

    tab_overview, tab_edit, tab_pdf, tab_ai = st.tabs(
        ["Overview", "Edit / Inspection", "PDF", "AI Word Report"]
    )

    with tab_overview:
        st.subheader("Identification")
        cols = st.columns(3)
        with cols[0]:
            field("Make", ident.get("make"))
            field("Model", ident.get("model"))
            field("Year", ident.get("model_year"))
        with cols[1]:
            field("Trim", ident.get("trim"))
            field("Body Class", ident.get("body_class"))
            field("Vehicle Type", ident.get("vehicle_type"))
        with cols[2]:
            field("Manufacturer", ident.get("manufacturer"))
            field("Series", ident.get("series"))
            field("Vehicle Class", ident.get("vehicle_class"))

        st.subheader("Engine & Drivetrain")
        cols = st.columns(3)
        with cols[0]:
            field("Displacement", engine.get("displacement"))
            field("Cylinders", engine.get("cylinders"))
            field("Fuel Type", engine.get("fuel_type"))
        with cols[1]:
            field("Horsepower", engine.get("horsepower"))
            field("Configuration", engine.get("configuration"))
            field("Engine Model", engine.get("engine_model"))
        with cols[2]:
            field("Transmission", drive.get("transmission"))
            field("Drive Type", drive.get("drive_type"))
            field("Doors", dims.get("doors"))

        st.subheader("Manufacturing / WMI")
        cols = st.columns(3)
        with cols[0]:
            field("Plant City", mfg.get("plant_city"))
            field("Plant State", mfg.get("plant_state"))
        with cols[1]:
            field("Plant Country", mfg.get("plant_country"))
            field("Plant Company", mfg.get("plant_company"))
        with cols[2]:
            field("WMI Code", wmi.get("code"))
            field("WMI Country", wmi.get("country"))

        st.subheader("Safety Recalls")
        if not safety.get("recalls_available") or not safety.get("recalls"):
            st.info(
                safety.get("recalls_note")
                or "Recall information was not available from the connected free source."
            )
        else:
            st.warning(f"{len(safety['recalls'])} recall campaign(s) for this Year / Make / Model.")
            for r in safety["recalls"]:
                with st.expander(
                    f"{r.get('campaign_number', '')} — {r.get('component') or 'Component N/A'}"
                ):
                    if r.get("summary"):
                        st.write(r["summary"])
                    if r.get("remedy"):
                        st.markdown(f"**Remedy:** {r['remedy']}")

        st.subheader("Sources")
        for s in report.get("sources") or []:
            icon = "✅" if s.get("success") else "❌"
            st.write(
                f"{icon} **{s.get('name')}**"
                + (f" — {s['error']}" if s.get("error") else "")
            )

        st.caption(
            "This tool does **not** include ownership history, accidents, title brands, or odometer records. "
            "Missing information does not mean the vehicle has a clean history."
        )

    with tab_edit:
        st.subheader("Client & inspection (user-provided)")
        c1, c2 = st.columns(2)
        manual["client_name"] = c1.text_input("Client Name", manual.get("client_name", ""))
        manual["client_reference"] = c2.text_input(
            "Client Reference", manual.get("client_reference", "")
        )
        manual["registration"] = c1.text_input("Registration", manual.get("registration", ""))
        manual["mileage"] = c2.text_input("Mileage", manual.get("mileage", ""))
        manual["purchase_price"] = c1.text_input(
            "Purchase Price", manual.get("purchase_price", "")
        )
        manual["inspection_date"] = c2.text_input(
            "Inspection Date", manual.get("inspection_date", "")
        )
        manual["inspector_name"] = c1.text_input(
            "Inspector Name", manual.get("inspector_name", "")
        )

        st.markdown("**Condition**")
        cond = manual.get("condition") or {}
        overall_opts = ["not_assessed", "excellent", "very_good", "good", "fair", "poor"]
        part_opts = ["not_inspected", "excellent", "good", "fair", "poor"]
        c1, c2, c3, c4 = st.columns(4)
        cond["overall"] = c1.selectbox(
            "Overall",
            overall_opts,
            index=overall_opts.index(cond.get("overall", "not_assessed"))
            if cond.get("overall") in overall_opts
            else 0,
        )
        cond["exterior"] = c2.selectbox(
            "Exterior",
            part_opts,
            index=part_opts.index(cond.get("exterior", "not_inspected"))
            if cond.get("exterior") in part_opts
            else 0,
        )
        cond["interior"] = c3.selectbox(
            "Interior",
            part_opts,
            index=part_opts.index(cond.get("interior", "not_inspected"))
            if cond.get("interior") in part_opts
            else 0,
        )
        cond["engine"] = c4.selectbox(
            "Engine",
            part_opts,
            index=part_opts.index(cond.get("engine", "not_inspected"))
            if cond.get("engine") in part_opts
            else 0,
        )
        c1, c2, c3, c4 = st.columns(4)
        cond["transmission"] = c1.selectbox(
            "Transmission",
            part_opts,
            index=part_opts.index(cond.get("transmission", "not_inspected"))
            if cond.get("transmission") in part_opts
            else 0,
        )
        cond["electrical"] = c2.selectbox(
            "Electrical",
            part_opts,
            index=part_opts.index(cond.get("electrical", "not_inspected"))
            if cond.get("electrical") in part_opts
            else 0,
        )
        cond["tires"] = c3.selectbox(
            "Tires",
            part_opts,
            index=part_opts.index(cond.get("tires", "not_inspected"))
            if cond.get("tires") in part_opts
            else 0,
        )
        cond["brakes"] = c4.selectbox(
            "Brakes",
            part_opts,
            index=part_opts.index(cond.get("brakes", "not_inspected"))
            if cond.get("brakes") in part_opts
            else 0,
        )
        manual["condition"] = cond

        manual["notes"] = st.text_area("Notes", manual.get("notes", ""), height=100)
        manual["additional_findings"] = st.text_area(
            "Additional findings", manual.get("additional_findings", ""), height=80
        )

        if st.button("Save changes", type="primary"):
            report["manual"] = manual
            report["updated_at"] = datetime.now().isoformat(timespec="seconds")
            st.session_state.report = report
            save_report_to_history(report)
            st.success("Saved.")

    with tab_pdf:
        st.subheader("Generate PDF")
        st.write(
            "Creates a professional multi-page client report from current data + your notes."
        )
        if st.button("Build PDF", type="primary"):
            report["manual"] = manual
            pdf_bytes = build_pdf(report, load_settings())
            fname = f"Vehicle_Report_{report.get('vin', 'VIN')}.pdf"
            st.download_button(
                "⬇️ Download PDF",
                data=pdf_bytes,
                file_name=fname,
                mime="application/pdf",
                type="primary",
            )
            st.success("PDF ready — click Download.")

    with tab_ai:
        st.subheader("AI Detailed Inspection Report (Word)")
        st.write(
            "Uses **Groq LLM** to expand the VIN decode + your inspection notes into a "
            "very detailed professional inspection narrative, then exports a **.docx** Word file."
        )
        st.caption(
            "Requires a free Groq API key from https://console.groq.com — "
            "never invents accident/title/ownership history."
        )

        groq_key = st.text_input(
            "Groq API Key",
            type="password",
            value=st.session_state.get("groq_api_key", ""),
            help="Stored only in this browser session. Prefer setting GROQ_API_KEY env var.",
            key="groq_key_input",
        )
        if groq_key:
            st.session_state.groq_api_key = groq_key

        model_choice = st.selectbox(
            "Preferred Groq model",
            [
                "llama-3.3-70b-versatile",
                "llama-3.1-70b-versatile",
                "llama-3.1-8b-instant",
                "gemma2-9b-it",
                "mixtral-8x7b-32768",
            ],
            index=0,
            help="If this model fails (rate limit, 413, unavailable), others are tried automatically.",
        )
        use_fallback = st.checkbox(
            "Enable model fallback",
            value=True,
            help="On failure, try: llama-3.3-70b → 3.1-70b → 3.1-8b-instant → gemma2-9b → mixtral",
        )

        if st.button("Generate detailed Word report", type="primary"):
            report["manual"] = manual
            st.session_state.report = report
            save_report_to_history(report)
            key = st.session_state.get("groq_api_key") or None
            try:
                with st.spinner("Calling Groq (with model fallback) and building Word document…"):
                    narrative, model_used = generate_inspection_narrative(
                        report,
                        api_key=key,
                        model=model_choice,
                        fallback=use_fallback,
                    )
                    docx_bytes = build_inspection_docx(
                        report, narrative, load_settings()
                    )
                st.session_state.ai_narrative = narrative
                st.session_state.ai_docx = docx_bytes
                st.session_state.ai_model_used = model_used
                if model_used != model_choice:
                    st.warning(
                        f"Preferred model unavailable — used fallback: **{model_used}**"
                    )
                st.success(f"Detailed inspection report ready (model: `{model_used}`).")
            except Exception as e:
                st.error(f"Could not reach Groq API: {e}")
                st.session_state.ai_docx = None

        if st.session_state.get("ai_docx"):
            fname = f"Vehicle_Inspection_{report.get('vin', 'VIN')}.docx"
            st.download_button(
                "⬇️ Download Word (.docx)",
                data=st.session_state.ai_docx,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )
            if st.session_state.get("ai_narrative"):
                with st.expander("Preview AI narrative (Markdown)"):
                    st.markdown(st.session_state.ai_narrative)
