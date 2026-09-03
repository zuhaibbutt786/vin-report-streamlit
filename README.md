# VIN Vehicle Report Generator (Streamlit)

Personal tool to decode a VIN using **free NHTSA public APIs**, add your own inspection notes, download a **professional PDF**, and optionally generate a **very detailed AI inspection Word (.docx)** report via **Groq**.

No Node.js, no npm, no backend server.

## Features

- VIN format + check-digit validation
- NHTSA vPIC factory decode (make, model, year, engine, plant, etc.)
- NHTSA safety recalls by Year / Make / Model
- Static WMI manufacturer / country lookup
- Manual inspection fields & condition ratings
- Professional multi-page PDF (ReportLab)
- **AI detailed inspection Word report (Groq LLM → .docx)**
- Local history saved as JSON on disk
- Branding settings for client PDFs

## Important limitations

This tool **never invents** vehicle history.

It does **not** include:

- Ownership history
- Accident / damage records
- Title brands (salvage, flood, etc.)
- Odometer history
- Service or auction records

Missing information does **not** mean the vehicle has a clean history.

## Requirements

- Python 3.10+
- Internet (for NHTSA decode / recalls and optional Groq)
- Optional: free [Groq API key](https://console.groq.com) for AI Word reports

## Install & run

```bash
cd vin-report-streamlit
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```

Browser opens at **http://localhost:8501**

Optional: export `GROQ_API_KEY=gsk_...` so you do not need to paste the key every time.

## Usage

1. **New VIN Decode** → enter 17-character VIN
2. Review decoded specs and recalls
3. **Edit / Inspection** → client name, mileage, condition, notes
4. **PDF** → Build PDF → Download
5. **AI Word Report** → paste Groq key (or use env) → Generate → Download `.docx`

## Project structure

```
vin-report-streamlit/
├── app.py                 # Main Streamlit UI
├── requirements.txt
├── data/
│   ├── wmi.json           # WMI manufacturer lookup
│   ├── settings.json      # Created at runtime
│   └── report_history.json
└── utils/
    ├── vin_validate.py
    ├── nhtsa.py           # vPIC + recalls
    ├── pdf_report.py      # ReportLab PDF
    └── groq_inspection.py # Groq narrative + Word (.docx)
```

## Sample VINs

- `1HGCM82633A004352`
- `5YJSA1E14HF000001`
- `1FTFW1ET5DFC10312`

## License

MIT — use freely for personal and client reports.
