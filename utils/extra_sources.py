"""Extra free public sources: complaints, NCAP ratings, EPA MPG, secondary decode."""

from __future__ import annotations

from typing import Any
from xml.etree import ElementTree as ET

import requests

TIMEOUT = 15
HEADERS_JSON = {"Accept": "application/json"}


def _clean(val: Any) -> str | None:
    if val is None:
        return None
    s = str(val).strip()
    if not s or s.lower() in ("null", "none", "not applicable", "n/a", ""):
        return None
    return s


def fetch_complaints(make: str, model: str, model_year: str, limit: int = 25) -> dict:
    """NHTSA owner complaints by Year/Make/Model."""
    out = {"success": False, "error": None, "count": 0, "items": [], "crash_count": 0, "fire_count": 0, "injury_count": 0}
    if not (make and model and model_year):
        out["error"] = "Make, model, and year required."
        return out
    try:
        r = requests.get(
            "https://api.nhtsa.gov/complaints/complaintsByVehicle",
            params={"make": make, "model": model, "modelYear": model_year},
            timeout=TIMEOUT,
        )
        r.raise_for_status()
        payload = r.json()
        rows = payload.get("results") or []
        out["count"] = int(payload.get("count") or len(rows))
        items = []
        crash = fire = injury = 0
        for row in rows[:limit]:
            if row.get("crash"):
                crash += 1
            if row.get("fire"):
                fire += 1
            injuries = int(row.get("numberOfInjuries") or 0)
            if injuries:
                injury += 1
            items.append(
                {
                    "odi_number": str(row.get("odiNumber") or ""),
                    "components": _clean(row.get("components")),
                    "summary": _clean(row.get("summary")),
                    "crash": bool(row.get("crash")),
                    "fire": bool(row.get("fire")),
                    "injuries": injuries,
                    "deaths": int(row.get("numberOfDeaths") or 0),
                    "date_incident": _clean(row.get("dateOfIncident")),
                    "date_filed": _clean(row.get("dateComplaintFiled")),
                }
            )
        out["success"] = True
        out["items"] = items
        out["crash_count"] = crash
        out["fire_count"] = fire
        out["injury_count"] = injury
    except requests.Timeout:
        out["error"] = "Complaints lookup timed out."
    except Exception as e:
        out["error"] = f"Complaints not available: {e}"
    return out


def fetch_ncap_ratings(make: str, model: str, model_year: str) -> dict:
    """NHTSA 5-Star Safety Ratings (NCAP) by Year/Make/Model."""
    out = {
        "success": False,
        "error": None,
        "variants": [],
        "overall": None,
        "frontal": None,
        "side": None,
        "rollover": None,
        "vehicle_description": None,
    }
    if not (make and model and model_year):
        out["error"] = "Make, model, and year required."
        return out
    try:
        r = requests.get(
            f"https://api.nhtsa.gov/SafetyRatings/modelyear/{model_year}/make/{make}/model/{model}",
            timeout=TIMEOUT,
        )
        r.raise_for_status()
        payload = r.json()
        results = payload.get("Results") or []
        if not results:
            out["error"] = "No NCAP safety ratings found for this Year / Make / Model."
            return out
        out["variants"] = [
            {"vehicle_id": str(v.get("VehicleId")), "description": _clean(v.get("VehicleDescription"))}
            for v in results
        ]
        vid = results[0].get("VehicleId")
        r2 = requests.get(f"https://api.nhtsa.gov/SafetyRatings/VehicleId/{vid}", timeout=TIMEOUT)
        r2.raise_for_status()
        detail = (r2.json().get("Results") or [{}])[0]
        out["success"] = True
        out["vehicle_description"] = _clean(detail.get("VehicleDescription")) or out["variants"][0]["description"]
        out["overall"] = _clean(detail.get("OverallRating"))
        out["frontal"] = _clean(detail.get("OverallFrontCrashRating"))
        out["side"] = _clean(detail.get("OverallSideCrashRating"))
        out["rollover"] = _clean(detail.get("RolloverRating"))
        out["front_driver"] = _clean(detail.get("FrontCrashDriversideRating"))
        out["front_passenger"] = _clean(detail.get("FrontCrashPassengersideRating"))
        out["nhtsa_url"] = f"https://www.nhtsa.gov/vehicle/{vid}"
    except requests.Timeout:
        out["error"] = "Safety ratings lookup timed out."
    except Exception as e:
        out["error"] = f"Safety ratings not available: {e}"
    return out


def _xml_text(el, tag: str) -> str | None:
    node = el.find(tag)
    if node is not None and node.text:
        return node.text.strip()
    return None


def fetch_epa_fuel(make: str, model: str, model_year: str) -> dict:
    """EPA / FuelEconomy.gov MPG via official free web service."""
    out = {
        "success": False,
        "error": None,
        "options": [],
        "city": None,
        "highway": None,
        "combined": None,
        "fuel_type": None,
        "annual_fuel_cost": None,
        "co2": None,
        "vehicle_id": None,
        "option_text": None,
    }
    if not (make and model and model_year):
        out["error"] = "Make, model, and year required."
        return out
    try:
        r = requests.get(
            "https://www.fueleconomy.gov/ws/rest/vehicle/menu/options",
            params={"year": model_year, "make": make, "model": model},
            headers=HEADERS_JSON,
            timeout=TIMEOUT,
        )
        if r.status_code != 200:
            out["error"] = "EPA options lookup failed."
            return out
        data = r.json()
        items = data.get("menuItem") or []
        if isinstance(items, dict):
            items = [items]
        if not items:
            out["error"] = "No EPA fuel-economy configurations found for this Year / Make / Model."
            return out
        out["options"] = [{"text": i.get("text"), "value": str(i.get("value"))} for i in items if i.get("value")]
        vid = out["options"][0]["value"]
        out["option_text"] = out["options"][0]["text"]
        out["vehicle_id"] = vid

        r2 = requests.get(
            f"https://www.fueleconomy.gov/ws/rest/vehicle/{vid}",
            headers=HEADERS_JSON,
            timeout=TIMEOUT,
        )
        if "json" in r2.headers.get("content-type", ""):
            v = r2.json()
            out["city"] = _clean(v.get("city08"))
            out["highway"] = _clean(v.get("highway08"))
            out["combined"] = _clean(v.get("comb08"))
            out["fuel_type"] = _clean(v.get("fuelType1") or v.get("fuelType"))
            out["annual_fuel_cost"] = _clean(v.get("fuelCost08"))
            out["co2"] = _clean(v.get("co2TailpipeGpm") or v.get("co2"))
        else:
            root = ET.fromstring(r2.content)
            out["city"] = _xml_text(root, "city08")
            out["highway"] = _xml_text(root, "highway08")
            out["combined"] = _xml_text(root, "comb08")
            out["fuel_type"] = _xml_text(root, "fuelType1") or _xml_text(root, "fuelType")
            out["annual_fuel_cost"] = _xml_text(root, "fuelCost08")
            out["co2"] = _xml_text(root, "co2TailpipeGpm") or _xml_text(root, "co2")
        out["success"] = True
    except requests.Timeout:
        out["error"] = "EPA fuel economy lookup timed out."
    except Exception as e:
        out["error"] = f"EPA fuel economy not available: {e}"
    return out


def secondary_decode(vin: str) -> dict:
    """Secondary free decode via DecodeVin for cross-check."""
    out = {"success": False, "error": None, "fields": {}}
    try:
        r = requests.get(
            f"https://vpic.nhtsa.dot.gov/api/vehicles/DecodeVin/{vin}",
            params={"format": "json"},
            timeout=TIMEOUT,
        )
        r.raise_for_status()
        results = r.json().get("Results") or []
        fields = {}
        for row in results:
            var = (row.get("Variable") or "").strip()
            val = _clean(row.get("Value"))
            if var and val and var not in ("Error Code", "Error Text", "Additional Error Text"):
                fields[var] = val
        out["success"] = bool(fields)
        out["fields"] = fields
        if not fields:
            out["error"] = "Secondary decode returned no useful fields."
    except Exception as e:
        out["error"] = f"Secondary decode unavailable: {e}"
    return out


def market_context_from_epa(epa: dict) -> dict:
    """Operating-cost context from EPA only — no marketplace scraping."""
    if not epa.get("success"):
        return {
            "success": False,
            "note": "Market listing prices are not available from free public APIs. "
            "Use a licensed valuation service for sale price estimates.",
        }
    return {
        "success": True,
        "annual_fuel_cost_usd": epa.get("annual_fuel_cost"),
        "combined_mpg": epa.get("combined"),
        "note": (
            "Annual fuel cost is an EPA estimate for typical U.S. driving, not a sale-price valuation. "
            "Transaction prices and residual values require paid market data sources."
        ),
    }
