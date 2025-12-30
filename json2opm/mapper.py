from typing import Any, Dict


# Explicit field order matters for .opm consumers
OPM_FIELD_ORDER = [
    "JsonVersion",
    "TestDateTime",
    "MeasurementId",
    "MeasurementName",
    "Identification",
    "Identifiers",
    "Hardware",
    "Reporting",
    "Context",
    "Measurement",
    "GlobalVerdict",
]


def _as_dict(v: Any) -> Dict[str, Any]:
    return v if isinstance(v, dict) else {}


def _normalize_identification(ident: Any) -> Dict[str, Any]:
    """
    Normalize Identification to be compatible with "working" OPM files and
    avoid fields that appear to break some importers.

    Changes:
    - Remove Geolocation / GeolocationDetails (observed difference: Bad vs known-working)
    - Map Exchange lowercase keys to schema-style keys:
        company  -> CompanyName
        customer -> CustomerName
    - Keep everything else as-is (JobId, OperatorA/B, Comment, etc.)
    """
    d = dict(_as_dict(ident))  # shallow copy

    # Drop fields that are known to be present in Bad but absent in a known-working OPM
    d.pop("Geolocation", None)
    d.pop("GeolocationDetails", None)

    # Normalize key naming (keep schema-style keys if already present)
    if "CompanyName" not in d and "company" in d:
        d["CompanyName"] = d.get("company")
    if "CustomerName" not in d and "customer" in d:
        d["CustomerName"] = d.get("customer")

    # Remove the lowercase originals to avoid duplicate/conflicting identity fields
    d.pop("company", None)
    d.pop("customer", None)

    return d


def map_pxm_json_to_opm(src: Dict[str, Any]) -> Dict[str, Any]:
    """
    Convert a PXM/Exchange JSON result into an OPM-compatible JSON structure.

    Rules:
    - `brief` is authoritative
    - Fields are copied verbatim EXCEPT for a small normalization step on Identification
      (to match known-working .opm JSON and avoid problematic fields)
    - No regeneration of measurement values
    - Only explicitly mapped fields are emitted
    """
    if "brief" not in src:
        raise ValueError("Source JSON missing required 'brief' section")

    brief = _as_dict(src["brief"])
    opm: Dict[str, Any] = {}

    for field in OPM_FIELD_ORDER:
        if field not in brief:
            raise ValueError(f"Missing required field in brief: {field}")

        if field == "Identification":
            opm[field] = _normalize_identification(brief.get(field))
        else:
            opm[field] = brief[field]

    return opm
