"""VIN format validation and check-digit verification."""

from __future__ import annotations

TRANSLITERATION = {
    "A": 1, "B": 2, "C": 3, "D": 4, "E": 5, "F": 6, "G": 7, "H": 8,
    "J": 1, "K": 2, "L": 3, "M": 4, "N": 5, "P": 7, "R": 9,
    "S": 2, "T": 3, "U": 4, "V": 5, "W": 6, "X": 7, "Y": 8, "Z": 9,
    "0": 0, "1": 1, "2": 2, "3": 3, "4": 4, "5": 5, "6": 6, "7": 7, "8": 8, "9": 9,
}

WEIGHTS = [8, 7, 6, 5, 4, 3, 2, 10, 0, 9, 8, 7, 6, 5, 4, 3, 2]


def clean_vin(raw: str) -> str:
    return "".join(c for c in (raw or "").upper() if c.isalnum())


def validate_vin(raw: str) -> dict:
    vin = clean_vin(raw)

    if len(vin) != 17:
        return {"valid": False, "cleaned": vin, "error": f"VIN must be 17 characters (got {len(vin)})."}

    if any(c in vin for c in "IOQ"):
        return {"valid": False, "cleaned": vin, "error": "VIN cannot contain the letters I, O, or Q."}

    total = 0
    for i, ch in enumerate(vin):
        if ch not in TRANSLITERATION:
            return {"valid": False, "cleaned": vin, "error": f"Invalid character '{ch}' in VIN."}
        total += TRANSLITERATION[ch] * WEIGHTS[i]

    remainder = total % 11
    expected = "X" if remainder == 10 else str(remainder)
    actual = vin[8]

    if actual != expected:
        return {
            "valid": True,
            "cleaned": vin,
            "check_digit_ok": False,
            "warning": f"Check digit mismatch (expected {expected}, found {actual}). Decoding will still be attempted.",
        }

    return {"valid": True, "cleaned": vin, "check_digit_ok": True}


def get_year_hint(vin: str) -> str | None:
    if len(vin) < 10:
        return None
    code = vin[9]
    modern = {
        "A": "2010", "B": "2011", "C": "2012", "D": "2013", "E": "2014", "F": "2015",
        "G": "2016", "H": "2017", "J": "2018", "K": "2019", "L": "2020", "M": "2021",
        "N": "2022", "P": "2023", "R": "2024", "S": "2025", "T": "2026",
        "V": "2027", "W": "2028", "X": "2029", "Y": "2030",
        "1": "2001", "2": "2002", "3": "2003", "4": "2004", "5": "2005",
        "6": "2006", "7": "2007", "8": "2008", "9": "2009",
    }
    return modern.get(code)
