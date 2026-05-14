"""Realworks Aankoop v3 koppeling.

Alleen het endpoint:
    GET https://api.realworks.nl/aankoop/v3/objecten/{afdelingscode}

Authenticatie: token in Authorization-header.
Configuratie via environment variables:
    REALWORKS_TOKEN
    REALWORKS_AFDELINGSCODE
"""

from __future__ import annotations

import os
from typing import Optional

import requests

BASE_URL = "https://api.realworks.nl/aankoop/v3"
TIMEOUT = 20

# Velden die we exposen (en daarmee terug naar de UI / Excel sturen)
VELDEN = (
    "systemid",
    "straat",
    "huisnummer",
    "huisnummerToevoeging",
    "postcode",
    "plaatsnaam",
    "status",
    "bouwjaar",
    "woonoppervlakte",
    "aantalKamers",
    "soortObject",
    "typeAppartement",
    "woningType",
    "vraagprijs",
    "transactieprijs",
    "datumTransport",
)


def _token() -> str:
    return os.environ.get("REALWORKS_TOKEN", "").strip()


def _afdelingscode() -> str:
    return os.environ.get("REALWORKS_AFDELINGSCODE", "").strip()


def is_geconfigureerd() -> bool:
    return bool(_token() and _afdelingscode())


def _media_link(obj: dict) -> Optional[str]:
    """Pakt de eerste bruikbare URL uit een Realworks media-veld.

    Realworks levert media meestal als lijst van dicts met 'link' of 'url'.
    """
    media = obj.get("media") or obj.get("foto") or obj.get("fotos")
    if not media:
        return None
    if isinstance(media, list):
        items = media
    elif isinstance(media, dict):
        items = [media]
    else:
        return None
    for item in items:
        if isinstance(item, str):
            return item
        if isinstance(item, dict):
            for k in ("link", "url", "href", "uri", "src"):
                if item.get(k):
                    return item[k]
    return None


def _compact(obj: dict) -> dict:
    """Reduceert een Realworks-object tot de gewenste velden + media_url."""
    compact = {f: obj.get(f) for f in VELDEN}
    compact["media_url"] = _media_link(obj)
    return compact


def adres_string(obj: dict) -> str:
    """Bouwt een adresregel uit losse Realworks-velden voor PDOK-lookup."""
    parts = []
    if obj.get("straat"):
        parts.append(str(obj["straat"]).strip())
    if obj.get("huisnummer") not in (None, ""):
        parts.append(str(obj["huisnummer"]).strip())
    if obj.get("huisnummerToevoeging"):
        parts.append(str(obj["huisnummerToevoeging"]).strip())
    if obj.get("postcode"):
        parts.append(str(obj["postcode"]).strip())
    if obj.get("plaatsnaam"):
        parts.append(str(obj["plaatsnaam"]).strip())
    return " ".join(parts).strip()


class RealworksError(Exception):
    """Wrapper voor fouten van/naar Realworks."""

    def __init__(self, message: str, status_code: Optional[int] = None):
        super().__init__(message)
        self.status_code = status_code


def haal_objecten(
    *,
    vanaf: int = 0,
    aantal: int = 100,
    status: Optional[str] = None,
) -> dict:
    """Eén pagina Realworks-objecten ophalen.

    Returnt dict:
        {
            "objecten": [...],
            "vanaf": int, "aantal": int, "ontvangen": int,
            "meer_beschikbaar": bool,
            "totaal_in_api": Optional[int],  # alleen als Realworks dit teruggeeft
        }

    Raises RealworksError bij configuratie- of HTTP-fouten.
    """
    if not is_geconfigureerd():
        raise RealworksError(
            "Realworks niet geconfigureerd: zet REALWORKS_TOKEN en REALWORKS_AFDELINGSCODE."
        )

    if aantal < 1 or aantal > 100:
        aantal = 100
    if vanaf < 0:
        vanaf = 0

    url = f"{BASE_URL}/objecten/{_afdelingscode()}"
    params = {"aantal": aantal, "vanaf": vanaf, "actief": "true"}
    if status:
        params["status"] = status
    headers = {"Authorization": _token(), "Accept": "application/json"}

    try:
        r = requests.get(url, params=params, headers=headers, timeout=TIMEOUT)
    except requests.RequestException as e:
        raise RealworksError(f"Netwerkfout: {e}") from e

    if r.status_code == 401:
        raise RealworksError("Niet geautoriseerd — controleer REALWORKS_TOKEN.", 401)
    if r.status_code == 403:
        raise RealworksError(
            "Geen toegang tot deze afdelingscode — controleer REALWORKS_AFDELINGSCODE.", 403
        )
    if r.status_code == 404:
        raise RealworksError("Endpoint of afdelingscode niet gevonden (404).", 404)
    if not r.ok:
        raise RealworksError(f"Realworks gaf status {r.status_code}: {r.text[:200]}", r.status_code)

    try:
        payload = r.json()
    except ValueError as e:
        raise RealworksError(f"Ongeldige JSON-respons van Realworks: {e}") from e

    # De response kan een lijst zijn óf een dict met paginering. Beide afhandelen.
    if isinstance(payload, list):
        raw_objecten = payload
        totaal_in_api = None
    elif isinstance(payload, dict):
        raw_objecten = (
            payload.get("resultaten")
            or payload.get("results")
            or payload.get("objecten")
            or payload.get("data")
            or []
        )
        totaal_in_api = (
            payload.get("totaal")
            or payload.get("total")
            or payload.get("totaalAantal")
            or payload.get("totalCount")
        )
    else:
        raw_objecten = []
        totaal_in_api = None

    objecten = [_compact(o) for o in raw_objecten if isinstance(o, dict)]
    ontvangen = len(objecten)
    meer = ontvangen >= aantal
    return {
        "objecten": objecten,
        "vanaf": vanaf,
        "aantal": aantal,
        "ontvangen": ontvangen,
        "meer_beschikbaar": meer,
        "totaal_in_api": totaal_in_api,
    }


def haal_alle_objecten(
    *,
    status: Optional[str] = None,
    max_objecten: int = 1000,
    pagesize: int = 100,
) -> dict:
    """Loopt door alle pagina's heen tot max_objecten of einde-lijst.

    Returnt:
        {"objecten": [...], "totaal": int, "afgebroken": bool}
    """
    alle: list[dict] = []
    vanaf = 0
    afgebroken = False
    while True:
        page = haal_objecten(vanaf=vanaf, aantal=pagesize, status=status)
        batch = page["objecten"]
        alle.extend(batch)
        if not page["meer_beschikbaar"] or len(batch) == 0:
            break
        if len(alle) >= max_objecten:
            afgebroken = True
            alle = alle[:max_objecten]
            break
        vanaf += pagesize
    return {"objecten": alle, "totaal": len(alle), "afgebroken": afgebroken}
