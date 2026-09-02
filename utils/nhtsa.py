"""NHTSA vPIC decode + recalls (free public APIs, no keys required)."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

import requests

VPIC_URL = "https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVinValues/{vin}?format=json"
RECALLS_URL = "https://api.nhtsa.gov/recalls/recallsByVehicle"
TIMEOUT = 15

WMI_PATH = Path(__file__).resolve().parent.parent / "data" / "wmi.json"


def _load_wmi() -> dict:
    try:
        with open(WMI_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def lookup_wmi(vin: str) -> dict:
    data = _load_wmi()
    code4 = vin[:4]
    code3 = vin[:3]
    entry = data.get(code4) or data.get(code3) or {}
    return {
        "code": code4 if code4 in data else code3,
        "manufacturer": entry.get("manufacturer"),
        "country": entry.get("country"),
        "region": entry.get("region"),
    }


def _clean(val: Any) -> str | None:
    if val is None:
        return None
    s = str(val).strip()
    if not s or s.lower() in ("null", "none", "not applicable", "n/a", ""):
        return None
    return s


def decode_vin(vin: str) -> dict:
    """Decode VIN via NHTSA vPIC. Never invents data."""
    result = {
        "success": False,
        "error": None,
        "raw": {},
        "identification": {},
        "engine": {},
        "drivetrain": {},
        "dimensions": {},
        "manufacturing": {},
    }
    try:
        r = requests.get(VPIC_URL.format(vin=vin), timeout=TIMEOUT)
        r.raise_for_status()
        payload = r.json()
        rows = payload.get("Results") or []
        if not rows:
            result["error"] = "No decode results returned by NHTSA vPIC."
            return result
        d = rows[0]
        result["raw"] = d
        result["success"] = True

        result["identification"] = {
            "make": _clean(d.get("Make")),
            "manufacturer": _clean(d.get("Manufacturer")) or _clean(d.get("Make")),
            "model": _clean(d.get("Model")),
            "model_year": _clean(d.get("ModelYear")),
            "trim": _clean(d.get("Trim")),
            "series": _clean(d.get("Series")),
            "vehicle_type": _clean(d.get("VehicleType")),
            "body_class": _clean(d.get("BodyClass")),
            "vehicle_class": _clean(d.get("VehicleClass")),
        }
        disp = None
        if _clean(d.get("DisplacementL")):
            disp = f"{_clean(d.get('DisplacementL'))} L"
        elif _clean(d.get("DisplacementCC")):
            disp = f"{_clean(d.get('DisplacementCC'))} cc"

        result["engine"] = {
            "engine_model": _clean(d.get("EngineModel")),
            "displacement": disp,
            "cylinders": _clean(d.get("EngineCylinders")),
            "fuel_type": _clean(d.get("FuelTypePrimary")),
            "configuration": _clean(d.get("EngineConfiguration")),
            "horsepower": _clean(d.get("EngineHP")),
            "manufacturer": _clean(d.get("EngineManufacturer")),
        }
        result["drivetrain"] = {
            "transmission": _clean(d.get("TransmissionStyle")) or _clean(d.get("Transmission")),
            "transmission_speeds": _clean(d.get("TransmissionSpeeds")),
            "drive_type": _clean(d.get("DriveType")),
        }
        result["dimensions"] = {
            "doors": _clean(d.get("Doors")),
            "seats": _clean(d.get("Seats")) or _clean(d.get("SeatRows")),
            "gvwr": _clean(d.get("GVWR")) or _clean(d.get("GrossVehicleWeightRatingFrom")),
            "bed_type": _clean(d.get("BedType")),
        }
        result["manufacturing"] = {
            "plant_country": _clean(d.get("PlantCountry")),
            "plant_state": _clean(d.get("PlantState")),
            "plant_city": _clean(d.get("PlantCity")),
            "plant_company": _clean(d.get("PlantCompanyName")),
        }
    except requests.Timeout:
        result["error"] = "NHTSA vPIC request timed out."
    except requests.RequestException as e:
        result["error"] = f"Could not reach NHTSA vPIC: {e}"
    except Exception as e:
        result["error"] = f"Decode failed: {e}"
    return result


def fetch_recalls(make: str, model: str, model_year: str) -> dict:
    """Fetch recalls by Year/Make/Model. Not VIN-specific precision."""
    out = {"success": False, "error": None, "count": 0, "items": []}
    if not (make and model and model_year):
        out["error"] = "Make, model, and year required for recall lookup."
        return out
    try:
        r = requests.get(
            RECALLS_URL,
            params={"make": make, "model": model, "modelYear": model_year},
            timeout=TIMEOUT,
        )
        r.raise_for_status()
        payload = r.json()
        rows = payload.get("results") or []
        items = []
        for row in rows:
            items.append(
                {
                    "campaign_number": str(row.get("NHTSACampaignNumber") or ""),
                    "manufacturer": _clean(row.get("Manufacturer")),
                    "component": _clean(row.get("Component")),
                    "summary": _clean(row.get("Summary")),
                    "consequence": _clean(row.get("Consequence")),
                    "remedy": _clean(row.get("Remedy")),
                    "report_date": _clean(row.get("ReportReceivedDate")),
                }
            )
        out["success"] = True
        out["count"] = len(items)
        out["items"] = items
    except requests.Timeout:
        out["error"] = "Recall lookup timed out."
    except requests.RequestException as e:
        out["error"] = f"Recall information was not available from the connected free source ({e})."
    except Exception as e:
        out["error"] = str(e)
    return out
