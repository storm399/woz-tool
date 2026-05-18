"""WOZ-waardeloket webtool.

Lokaal Flask-webserver dat:
- PDOK Locatieserver gebruikt om adres -> nummeraanduiding te resolven
- WOZ-waardeloket API (api.kadaster.nl) aanroept om WOZ-waarden op te vragen
- Enkel adres in browser laat invoeren OF batch Excel upload + download

Start: python3 app.py  ->  open http://127.0.0.1:5005
"""

import io
import os
import time
import re
from typing import Optional

import requests
from flask import Flask, render_template, request, jsonify, send_file, abort

import realworks_service
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule


def _load_dotenv(path: str = ".env") -> None:
    """Mini-loader voor .env (alleen lokaal; op Render zet je env vars in dashboard)."""
    if not os.path.exists(path):
        return
    with open(path) as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            os.environ.setdefault(key.strip(), val.strip().strip('"').strip("'"))


_load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

app = Flask(__name__)

PDOK_FREE = "https://api.pdok.nl/bzk/locatieserver/search/v3_1/free"
PDOK_SUGGEST = "https://api.pdok.nl/bzk/locatieserver/search/v3_1/suggest"
WOZ_API = "https://api.kadaster.nl/lvwoz/wozwaardeloket-api/v1/wozwaarde/nummeraanduiding/{}"
WOZ_HEADERS = {
    "Accept": "application/json",
    "Origin": "https://www.wozwaardeloket.nl",
    "Referer": "https://www.wozwaardeloket.nl/",
    "User-Agent": "Mozilla/5.0 (WOZ-tool lokaal)",
}

EPO_API = "https://public.ep-online.nl/api/v5/PandEnergielabel/AdresseerbaarObject/{}"
EPO_API_KEY = os.environ.get("EPONLINE_API_KEY", "").strip()

BAG_WFS = "https://service.pdok.nl/lv/bag/wfs/v2_0"
CBS_WFS = "https://service.pdok.nl/cbs/wijkenbuurten/2023/wfs/v1_0"
RCE_SPARQL = "https://api.linkeddata.cultureelerfgoed.nl/datasets/rce/cho/sparql"

REQUEST_TIMEOUT = 15
session = requests.Session()


def _normaliseer(s: str) -> str:
    """Normaliseert adres-string voor match-vergelijking."""
    if not s:
        return ""
    s = s.lower()
    s = re.sub(r"[^\w\s]", " ", s, flags=re.UNICODE)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _check_adres_match(invoer: str, gevonden_straat: str, gevonden_huisnummer) -> tuple:
    """Returnt (match: bool, reden: str).

    Match = True als de gevonden straatnaam (genormaliseerd) als substring
    in de invoer voorkomt EN het gevonden huisnummer in de invoer staat.
    """
    inv = _normaliseer(invoer)
    if not inv:
        return False, "Adres leeg"

    straat_norm = _normaliseer(gevonden_straat or "")
    nr = str(gevonden_huisnummer) if gevonden_huisnummer is not None else ""

    straat_ok = bool(straat_norm) and straat_norm in inv
    if not straat_ok:
        return False, f"Straat ‘{gevonden_straat}’ niet in invoer"

    if nr:
        invoer_nums = re.findall(r"\d+", inv)
        if nr not in invoer_nums:
            return False, f"Huisnummer {nr} niet in invoer"

    return True, "OK"


def _bereken_afgeleiden(res: dict) -> None:
    """Vult res in-place met afgeleide velden: WOZ/m², %-stijging, match-flag."""
    waarden = res.get("wozWaarden") or []
    bag = res.get("bag") or {}
    opp = bag.get("gebruiksoppervlakte")
    laatste = waarden[0]["vastgesteldeWaarde"] if waarden and waarden[0].get("vastgesteldeWaarde") else None

    res["woz_per_m2"] = int(round(laatste / opp)) if (laatste and opp) else None

    # Stijgingspercentages
    if len(waarden) >= 2 and waarden[0].get("vastgesteldeWaarde") and waarden[1].get("vastgesteldeWaarde"):
        cur = waarden[0]["vastgesteldeWaarde"]
        prev = waarden[1]["vastgesteldeWaarde"]
        res["pct_1jr"] = round((cur - prev) / prev * 100, 1)
    else:
        res["pct_1jr"] = None

    def vergelijk_x_jaar_geleden(jaren):
        if not waarden or not laatste:
            return None
        peildatum_doel = None
        from datetime import datetime
        try:
            cur_year = int(waarden[0]["peildatum"][:4])
        except (KeyError, ValueError, TypeError):
            return None
        doel_year = cur_year - jaren
        for w in waarden:
            if w.get("peildatum", "").startswith(str(doel_year)):
                doel = w.get("vastgesteldeWaarde")
                if doel:
                    return round((laatste - doel) / doel * 100, 1)
        # Anders: pak oudst beschikbaar
        oudste = waarden[-1].get("vastgesteldeWaarde") if waarden else None
        return round((laatste - oudste) / oudste * 100, 1) if oudste else None

    res["pct_5jr"] = vergelijk_x_jaar_geleden(5)

    # Vergelijking met oudst beschikbaar (vaak 2014)
    if waarden and waarden[-1].get("vastgesteldeWaarde") and laatste:
        oudste = waarden[-1]["vastgesteldeWaarde"]
        res["pct_sinds_oudst"] = round((laatste - oudste) / oudste * 100, 1)
        res["oudste_peiljaar"] = (waarden[-1].get("peildatum") or "")[:4]
    else:
        res["pct_sinds_oudst"] = None
        res["oudste_peiljaar"] = None

    # Adres-match-check: gebruik losse straat+huisnummer ipv hele weergavenaam
    bag_straat = res.get("straat") or (res.get("bag") or {}).get("straat")
    bag_nr = res.get("huisnummer")
    match, reden = _check_adres_match(res.get("adres_invoer", ""), bag_straat, bag_nr)
    res["adres_match"] = match
    res["adres_match_reden"] = reden


def zoek_adres(query: str) -> Optional[dict]:
    """Fuzzy adres-zoek via PDOK (single-lookup). Returnt eerste hit of None."""
    try:
        r = session.get(
            PDOK_FREE,
            params={"q": query, "fq": "type:adres", "rows": 1},
            timeout=REQUEST_TIMEOUT,
        )
        r.raise_for_status()
        docs = r.json().get("response", {}).get("docs", [])
        return docs[0] if docs else None
    except requests.RequestException:
        return None


def _norm_postcode(pc: str) -> str:
    """Normaliseert een NL-postcode naar '1234AB' (uppercase, geen spaties)."""
    if not pc:
        return ""
    return re.sub(r"\s+", "", str(pc)).upper()


def _norm_letter(s) -> Optional[str]:
    if s is None or s == "":
        return None
    s = str(s).strip()
    return s.upper() if s else None


def _norm_toev(s) -> Optional[str]:
    if s is None or s == "":
        return None
    s = str(s).strip()
    return s if s else None


def zoek_adres_strict(
    postcode: str,
    huisnummer,
    huisletter=None,
    huisnummertoevoeging=None,
) -> Optional[dict]:
    """Strict adres-zoek op postcode + huisnummer (primair) + letter/toevoeging.

    Gebruikt PDOK Locatieserver met filterquery, geen fuzzy match op straat.
    Filtert vervolgens op huisletter en huisnummertoevoeging als die zijn gegeven.
    """
    pc = _norm_postcode(postcode)
    if not pc or huisnummer in (None, ""):
        return None
    try:
        nr = int(str(huisnummer).strip())
    except (TypeError, ValueError):
        return None

    letter = _norm_letter(huisletter)
    toev = _norm_toev(huisnummertoevoeging)

    try:
        r = session.get(
            PDOK_FREE,
            params=[
                ("fq", "type:adres"),
                ("fq", f"postcode:{pc}"),
                ("fq", f"huisnummer:{nr}"),
                ("rows", 25),
            ],
            timeout=REQUEST_TIMEOUT,
        )
        r.raise_for_status()
        docs = r.json().get("response", {}).get("docs", []) or []
    except requests.RequestException:
        return None

    if not docs:
        return None

    def candidate_score(d):
        # Hoger = beter
        score = 0
        d_letter = _norm_letter(d.get("huisletter"))
        d_toev = _norm_toev(d.get("huisnummertoevoeging"))
        if letter and d_letter == letter:
            score += 10
        elif not letter and not d_letter:
            score += 5
        if toev and d_toev == toev:
            score += 10
        elif not toev and not d_toev:
            score += 5
        # Strikt geval: beide gevuld én matchen
        if letter and toev and d_letter == letter and d_toev == toev:
            score += 50
        return score

    docs.sort(key=candidate_score, reverse=True)
    return docs[0]


def haal_woz(nummeraanduiding_id: str) -> Optional[dict]:
    """Haalt WOZ-waarden op via Kadaster-API. Returnt dict of None bij 404."""
    try:
        r = session.get(
            WOZ_API.format(nummeraanduiding_id),
            headers=WOZ_HEADERS,
            timeout=REQUEST_TIMEOUT,
        )
        if r.status_code == 404:
            return None
        r.raise_for_status()
        return r.json()
    except requests.RequestException:
        return None


def haal_bag(adresseerbaar_object_id: str) -> Optional[dict]:
    """Haalt BAG-verblijfsobject op via PDOK WFS (geen API-key nodig).

    Returnt dict met bouwjaar, gebruiksoppervlakte (m²), gebruiksdoel,
    pandidentificatie en statussen.
    """
    if not adresseerbaar_object_id:
        return None
    fil = (
        f"<Filter><PropertyIsEqualTo>"
        f"<PropertyName>bag:identificatie</PropertyName>"
        f"<Literal>{adresseerbaar_object_id}</Literal>"
        f"</PropertyIsEqualTo></Filter>"
    )
    try:
        r = session.get(
            BAG_WFS,
            params={
                "service": "WFS",
                "version": "2.0.0",
                "request": "GetFeature",
                "typeNames": "bag:verblijfsobject",
                "outputFormat": "application/json",
                "srsName": "EPSG:4326",
                "filter": fil,
            },
            timeout=REQUEST_TIMEOUT,
        )
        r.raise_for_status()
        feats = r.json().get("features", [])
        if not feats:
            return None
        p = feats[0].get("properties", {}) or {}
        return {
            "bouwjaar": p.get("bouwjaar"),
            "gebruiksoppervlakte": p.get("oppervlakte"),
            "gebruiksdoel": p.get("gebruiksdoel"),
            "vbo_status": p.get("status"),
            "pand_id": p.get("pandidentificatie"),
            "pand_status": p.get("pandstatus"),
            "straat": p.get("openbare_ruimte"),
            "huisnummer": p.get("huisnummer"),
        }
    except (requests.RequestException, ValueError):
        return None


def haal_cbs_buurt(buurtcode: str) -> Optional[dict]:
    """Haalt CBS-kerncijfers voor een buurt op via PDOK Wijken-en-Buurten WFS.

    Verwacht buurtcode in formaat 'BU05990112'. Eenheden:
    - gemiddeldeWoningwaarde: in duizenden euro's (we vermenigvuldigen ×1000)
    - gemiddeldInkomenPerInkomensontvanger: idem ×1000
    """
    if not buurtcode:
        return None
    fil = (
        f"<Filter><PropertyIsEqualTo>"
        f"<PropertyName>buurtcode</PropertyName>"
        f"<Literal>{buurtcode}</Literal>"
        f"</PropertyIsEqualTo></Filter>"
    )
    try:
        r = session.get(
            CBS_WFS,
            params={
                "service": "WFS",
                "version": "2.0.0",
                "request": "GetFeature",
                "typeNames": "wijkenbuurten:buurten",
                "outputFormat": "application/json",
                "srsName": "EPSG:4326",
                "filter": fil,
                "count": 1,
            },
            timeout=REQUEST_TIMEOUT,
        )
        r.raise_for_status()
        feats = r.json().get("features", [])
        if not feats:
            return None
        p = feats[0].get("properties", {}) or {}

        def _g(key, mult=1):
            v = p.get(key)
            if v is None or (isinstance(v, (int, float)) and v <= -99000):
                return None
            return v * mult if isinstance(v, (int, float)) else v

        return {
            "buurtnaam": _g("buurtnaam"),
            "gemeentenaam": _g("gemeentenaam"),
            "aantal_inwoners": _g("aantalInwoners"),
            "bevolkingsdichtheid": _g("bevolkingsdichtheidInwonersPerKm2"),
            "gem_woningwaarde": _g("gemiddeldeWoningwaarde", 1000),
            "gem_inkomen": _g("gemiddeldInkomenPerInkomensontvanger", 1000),
            "pct_huur": _g("percentageHuurwoningen"),
            "pct_koop": _g("percentageKoopwoningen"),
            "huishoudgrootte": _g("gemiddeldeHuishoudsgrootte"),
            "woningvoorraad": _g("woningvoorraad"),
        }
    except (requests.RequestException, ValueError):
        return None


def haal_monument(pand_id: str) -> Optional[dict]:
    """Checkt of een BAG-pand een rijksmonument is via RCE Linked Data SPARQL.

    Returnt dict met monumentnummer + adres uit RCE, of None.
    """
    if not pand_id:
        return None
    query = (
        "PREFIX ceo: <https://linkeddata.cultureelerfgoed.nl/def/ceo#>\n"
        "SELECT ?nummer ?straat ?huisnr WHERE {\n"
        "  ?monument a ceo:Rijksmonument ;\n"
        "            ceo:rijksmonumentnummer ?nummer ;\n"
        "            ceo:heeftBasisregistratieRelatie/ceo:heeftBAGRelatie ?bagrel .\n"
        f"  ?bagrel ceo:pandIdentificatie \"{pand_id}\" .\n"
        "  OPTIONAL { ?bagrel ceo:openbareRuimte ?straat ; ceo:huisnummer ?huisnr }\n"
        "} LIMIT 1"
    )
    try:
        r = session.get(
            RCE_SPARQL,
            params={"query": query, "format": "json"},
            timeout=REQUEST_TIMEOUT,
        )
        r.raise_for_status()
        rows = r.json()
        if not rows:
            return None
        row = rows[0]
        return {
            "monumentnummer": row.get("nummer"),
            "monument_straat": row.get("straat"),
            "monument_huisnummer": row.get("huisnr"),
        }
    except (requests.RequestException, ValueError):
        return None


def haal_energielabel(adresseerbaar_object_id: str) -> Optional[dict]:
    """Haalt het meest recente energielabel op via EP-online.

    Returnt dict met velden energieklasse, registratiedatum, geldig_tot,
    bouwjaar, gebouwtype — of None als er geen label is.
    """
    if not EPO_API_KEY:
        return None
    try:
        r = session.get(
            EPO_API.format(adresseerbaar_object_id),
            headers={"Authorization": EPO_API_KEY, "Accept": "application/json"},
            timeout=REQUEST_TIMEOUT,
        )
        if r.status_code == 404:
            return None
        r.raise_for_status()
        data = r.json()
        if not data:
            return None
        # Meest recente label (hoogste registratiedatum)
        labels = sorted(
            data,
            key=lambda x: x.get("Registratiedatum") or "",
            reverse=True,
        )
        latest = labels[0]
        return {
            "energieklasse": latest.get("Energieklasse"),
            "registratiedatum": (latest.get("Registratiedatum") or "")[:10],
            "geldig_tot": (latest.get("Geldig_tot") or "")[:10],
            "bouwjaar": latest.get("Bouwjaar"),
            "gebouwtype": latest.get("Gebouwtype"),
            "gebouwklasse": latest.get("Gebouwklasse"),
            "berekend_energieverbruik": latest.get("BerekendeEnergieverbruik"),
        }
    except requests.RequestException:
        return None


def _resultaat_uit_hit(hit: dict, adres_invoer: str, ref_straat: Optional[str] = None) -> dict:
    """Gemeenschappelijke verrijkingsflow: vanuit een PDOK-hit alle koppelingen draaien.

    `ref_straat` (optioneel) wordt alleen gebruikt voor straat-verificatie:
    als ingevulde straat duidelijk afwijkt van de gevonden straat, krijgt het
    resultaat een 'straat_verschilt'-flag (true-positief is niet leidend, het
    matchen zelf gebeurt op postcode+huisnummer).
    """
    weergavenaam = hit.get("weergavenaam", "")
    nummeraanduiding = hit.get("nummeraanduiding_id")
    adresseerbaar_object_id = hit.get("adresseerbaarobject_id")
    if not nummeraanduiding:
        return {"adres_invoer": adres_invoer, "weergavenaam": weergavenaam,
                "status": "Geen nummeraanduiding", "wozWaarden": []}

    buurtcode = hit.get("buurtcode")
    buurtnaam = hit.get("buurtnaam")

    woz = haal_woz(nummeraanduiding)
    bag = haal_bag(adresseerbaar_object_id) if adresseerbaar_object_id else None
    energielabel = haal_energielabel(adresseerbaar_object_id) if adresseerbaar_object_id else None
    cbs = haal_cbs_buurt(buurtcode) if buurtcode else None
    monument = haal_monument(bag.get("pand_id") if bag else None)

    base = {
        "adres_invoer": adres_invoer,
        "weergavenaam": weergavenaam,
        "nummeraanduiding": nummeraanduiding,
        "adresseerbaarobject_id": adresseerbaar_object_id,
        "buurtnaam": buurtnaam,
        "buurtcode": buurtcode,
        "bag": bag,
        "energielabel": energielabel,
        "cbs": cbs,
        "monument": monument,
    }

    if not woz:
        base.update({
            "status": "Geen WOZ-waarde bekend (niet-woning of onbekend)",
            "wozWaarden": [],
        })
        _bereken_afgeleiden(base)
        _verifieer_straat(base, ref_straat, hit)
        return base

    obj = woz.get("wozObject", {}) or {}
    waarden = woz.get("wozWaarden", []) or []
    base.update({
        "status": "OK",
        "wozobjectnummer": obj.get("wozobjectnummer"),
        "straat": obj.get("openbareruimtenaam") or hit.get("straatnaam"),
        "huisnummer": obj.get("huisnummer") or hit.get("huisnummer"),
        "huisletter": obj.get("huisletter") or hit.get("huisletter"),
        "huisnummertoevoeging": obj.get("huisnummertoevoeging") or hit.get("huisnummertoevoeging"),
        "postcode": obj.get("postcode") or hit.get("postcode"),
        "woonplaats": obj.get("woonplaatsnaam") or hit.get("woonplaatsnaam"),
        "grondoppervlakte": obj.get("grondoppervlakte"),
        "wozWaarden": sorted(
            [
                {"peildatum": w.get("peildatum"), "vastgesteldeWaarde": w.get("vastgesteldeWaarde")}
                for w in waarden
            ],
            key=lambda x: x["peildatum"] or "",
            reverse=True,
        ),
    })
    _bereken_afgeleiden(base)
    _verifieer_straat(base, ref_straat, hit)
    return base


def _verifieer_straat(res: dict, ref_straat: Optional[str], hit: dict) -> None:
    """Markeert in res of een opgegeven referentie-straat afwijkt van wat PDOK vond.

    Alleen ter informatie — de match zelf is al gedaan op postcode+huisnummer.
    """
    if not ref_straat:
        return
    gevonden = res.get("straat") or hit.get("straatnaam") or ""
    ref_norm = _normaliseer(ref_straat)
    gev_norm = _normaliseer(gevonden)
    if not gev_norm or not ref_norm:
        return
    # Match als één in de ander zit (compenseert "Wittedewithstraat" vs "Witte de Withstraat")
    ok = ref_norm in gev_norm or gev_norm in ref_norm
    if not ok:
        res["straat_verschilt"] = True
        res["straat_invoer"] = ref_straat
        res["straat_gevonden"] = gevonden


def maak_resultaat(adres: str) -> dict:
    """Fuzzy adres-flow voor losse opvragingen. Houdt huidig gedrag voor /api/woz."""
    hit = zoek_adres(adres)
    if not hit:
        return {"adres_invoer": adres, "status": "Adres niet gevonden", "wozWaarden": []}
    return _resultaat_uit_hit(hit, adres_invoer=adres)


def maak_resultaat_strict(
    postcode: str,
    huisnummer,
    huisletter=None,
    huisnummertoevoeging=None,
    straat: Optional[str] = None,
    plaats: Optional[str] = None,
    adres_invoer: Optional[str] = None,
) -> dict:
    """Strict match: postcode + huisnummer leidend. Straat alleen ter verificatie."""
    invoer_label = adres_invoer or " ".join(
        str(p) for p in (straat, huisnummer, huisletter, huisnummertoevoeging,
                          postcode, plaats) if p not in (None, "")
    ).strip()

    hit = zoek_adres_strict(postcode, huisnummer, huisletter, huisnummertoevoeging)
    if not hit:
        return {
            "adres_invoer": invoer_label,
            "status": "Adres niet gevonden op postcode + huisnummer",
            "wozWaarden": [],
            "match_modus": "strict",
        }
    res = _resultaat_uit_hit(hit, adres_invoer=invoer_label, ref_straat=straat)
    res["match_modus"] = "strict"
    return res


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/suggest")
def api_suggest():
    """Auto-suggest endpoint voor het zoekveld."""
    q = (request.args.get("q") or "").strip()
    if len(q) < 3:
        return jsonify([])
    try:
        r = session.get(
            PDOK_SUGGEST,
            params={"q": q, "fq": "type:adres", "rows": 8},
            timeout=REQUEST_TIMEOUT,
        )
        r.raise_for_status()
        docs = r.json().get("response", {}).get("docs", [])
        return jsonify([{"id": d.get("id"), "weergavenaam": d.get("weergavenaam")} for d in docs])
    except requests.RequestException as e:
        return jsonify({"error": str(e)}), 502


@app.route("/api/woz")
def api_woz():
    """Single-address lookup endpoint."""
    adres = (request.args.get("adres") or "").strip()
    if not adres:
        return jsonify({"error": "Parameter 'adres' is verplicht"}), 400
    return jsonify(maak_resultaat(adres))


@app.route("/api/excel", methods=["POST"])
def api_excel():
    """Batch-verwerking: Excel in -> Excel out met WOZ-data."""
    if "file" not in request.files:
        abort(400, "Geen bestand geüpload (veldnaam: 'file').")
    upload = request.files["file"]
    if not upload.filename.lower().endswith((".xlsx", ".xlsm")):
        abort(400, "Alleen .xlsx of .xlsm bestanden worden ondersteund.")

    try:
        wb_in = load_workbook(io.BytesIO(upload.read()), data_only=True)
    except Exception as e:
        abort(400, f"Kan Excel niet lezen: {e}")

    ws_in = wb_in.active
    rows = list(ws_in.iter_rows(values_only=True))
    if not rows:
        abort(400, "Het Excel-bestand is leeg.")

    header = [str(c) if c is not None else "" for c in rows[0]]
    header_lower = [h.lower().strip() for h in header]

    # Probeer een adres-kolom te vinden. Anders: hele rij plakken.
    adres_idx = None
    for i, h in enumerate(header_lower):
        if h in ("adres", "address", "weergavenaam", "volledig adres", "volledig_adres"):
            adres_idx = i
            break

    straat_idx = next((i for i, h in enumerate(header_lower) if h in ("straat", "straatnaam", "openbareruimte", "openbareruimtenaam")), None)
    huisnummer_idx = next((i for i, h in enumerate(header_lower) if h in ("huisnummer", "nr", "nummer", "huisnr")), None)
    letter_idx = next((i for i, h in enumerate(header_lower) if h in ("huisletter", "letter")), None)
    toevoeging_idx = next((i for i, h in enumerate(header_lower) if h in ("toevoeging", "huisnummertoevoeging", "huisnrtoevoeging")), None)
    postcode_idx = next((i for i, h in enumerate(header_lower) if h in ("postcode", "pc")), None)
    plaats_idx = next((i for i, h in enumerate(header_lower) if h in ("plaats", "woonplaats", "stad", "gemeente")), None)

    def _val(row, idx):
        if idx is None:
            return None
        v = row[idx] if idx < len(row) else None
        if v in (None, ""):
            return None
        return v

    def _adres_label(row) -> str:
        if adres_idx is not None and row[adres_idx]:
            return str(row[adres_idx]).strip()
        parts = []
        for idx in (straat_idx, huisnummer_idx, letter_idx, toevoeging_idx, postcode_idx, plaats_idx):
            v = _val(row, idx)
            if v is not None:
                parts.append(str(v).strip())
        return " ".join(parts).strip()

    def _parse_huisnr_combo(s):
        """Splitst '12A' / '12-2' / '12 bis' in (nummer, letter, toevoeging)."""
        s = str(s).strip()
        m = re.match(r"^(\d+)\s*([A-Za-z])?\s*[-/]?\s*(.*)$", s)
        if not m:
            return None, None, None
        return m.group(1), (m.group(2) or None), (m.group(3) or None)

    def _strict_kwargs(row):
        """Probeert postcode + huisnummer uit de rij te halen voor strict-match."""
        pc = _val(row, postcode_idx)
        nr = _val(row, huisnummer_idx)
        letter = _val(row, letter_idx)
        toev = _val(row, toevoeging_idx)
        straat = _val(row, straat_idx)
        plaats = _val(row, plaats_idx)

        if pc and nr is not None:
            # Als nummer in het veld als '12A' staat, splits letter eruit
            if letter is None and isinstance(nr, str) and re.search(r"[A-Za-z]", nr):
                num, lt, tv = _parse_huisnr_combo(nr)
                nr = num or nr
                letter = letter or lt
                toev = toev or tv
            return {
                "postcode": pc,
                "huisnummer": nr,
                "huisletter": letter,
                "huisnummertoevoeging": toev,
                "straat": straat,
                "plaats": plaats,
            }
        # Probeer ook uit een vrije adres-string een postcode+nr te halen
        if adres_idx is not None and row[adres_idx]:
            adres = str(row[adres_idx])
            pc_match = re.search(r"\b([1-9]\d{3}\s?[A-Za-z]{2})\b", adres)
            nr_match = re.search(r"\b(\d{1,5})\b", adres)
            if pc_match and nr_match:
                # Eerste cijferreeks die NIET binnen postcode valt
                nrs = re.findall(r"\b(\d{1,5})\b", adres)
                pc_digits = pc_match.group(1)[:4]
                nrs_clean = [n for n in nrs if n != pc_digits]
                if nrs_clean:
                    return {
                        "postcode": pc_match.group(1),
                        "huisnummer": nrs_clean[0],
                        "huisletter": None,
                        "huisnummertoevoeging": None,
                        "straat": straat,
                        "plaats": plaats,
                    }
        return None

    # Verzamel resultaten — postcode+huisnummer leidend, fallback op fuzzy
    resultaten = []
    peildata = set()
    for row in rows[1:]:
        adres_label = _adres_label(row)
        if not adres_label:
            resultaten.append({"adres_invoer": "", "status": "Leeg", "wozWaarden": []})
            continue

        strict = _strict_kwargs(row)
        if strict:
            res = maak_resultaat_strict(adres_invoer=adres_label, **strict)
            # Als strict niets vond, val terug op fuzzy als noodgreep
            if res.get("status", "").startswith("Adres niet gevonden"):
                fb = maak_resultaat(adres_label)
                if fb.get("status") == "OK":
                    fb["match_modus"] = "fuzzy-fallback"
                    fb["fallback_reden"] = "strict op postcode+huisnummer leverde geen hit"
                    res = fb
        else:
            res = maak_resultaat(adres_label)
            res["match_modus"] = "fuzzy"

        for w in res.get("wozWaarden") or []:
            if w.get("peildatum"):
                peildata.add(w["peildatum"])
        resultaten.append(res)
        time.sleep(0.15)  # voorkom rate-limit (60/min)

    peildata_sorted = sorted(peildata, reverse=True)

    wb_out = _bouw_excel(header, rows, resultaten, peildata_sorted)
    buf = io.BytesIO()
    wb_out.save(buf)
    buf.seek(0)
    naam = re.sub(r"\.xlsx?$", "", upload.filename, flags=re.I) + "_woz.xlsx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=naam,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# --- Excel-stijlen ---
_RED_DARK = "1F4E79"
_RED_MID = "2E6BA8"
_GREY = "F1F4F8"
_BORDER = Side(style="thin", color="C8D2DD")
_BORDER_THICK = Side(style="medium", color=_RED_DARK)

_LABEL_FILL = {
    "A++++": "009639", "A+++": "009639", "A++": "009639", "A+": "009639",
    "A": "009639", "B": "4CAF50", "C": "8BC34A",
    "D": "FFC107", "E": "FF9800", "F": "FF5722", "G": "D32F2F",
}


def _bouw_excel(header_in, rows, resultaten, peildata_sorted) -> Workbook:
    """Bouwt een net opgemaakt resultaat-workbook met twee tabbladen."""
    wb = Workbook()

    # ---- Tab 1: Resultaten ----
    ws = wb.active
    ws.title = "Resultaten"

    oudste_peiljaar = peildata_sorted[-1][:4] if peildata_sorted else "oudste"

    invoer_kols = list(header_in) or ["adres"]
    locatie_kols = ["status", "match", "gevonden adres", "postcode", "woonplaats", "BAG-id"]
    bag_kols = ["bouwjaar", "opp. (m²)", "gebruiksdoel"]
    buurt_kols = ["buurt", "gem. WOZ buurt", "gem. inkomen", "% huur"]
    label_kols = ["energielabel", "gebouwtype", "geldig t/m"]
    mon_kols = ["rijksmonument"]
    afgeleid_kols = ["€/m²", "% 1jr", "% 5jr", f"% sinds {oudste_peiljaar}"]
    woz_kols = [p[:4] for p in peildata_sorted]

    sizes = {
        "in": len(invoer_kols), "loc": len(locatie_kols), "bag": len(bag_kols),
        "buurt": len(buurt_kols), "lab": len(label_kols), "mon": len(mon_kols),
        "afg": len(afgeleid_kols), "woz": len(woz_kols),
    }

    # Rij 1: groepheaders (merged)
    groepen = [
        ("Invoer", sizes["in"], "6B7785"),
        ("Adres & WOZ-object", sizes["loc"], _RED_DARK),
        ("BAG-gegevens", sizes["bag"], "8B5A2B"),
        ("Buurt (CBS 2023)", sizes["buurt"], "0F766E"),
        ("Energielabel (EP-online)", sizes["lab"], "1B7A3A"),
        ("Monument (RCE)", sizes["mon"], "7C3AED"),
        ("Trend", sizes["afg"], "B45309"),
        ("WOZ-waarden", sizes["woz"], _RED_MID),
    ]
    col = 1
    for naam, n, kleur in groepen:
        if n == 0:
            continue
        ws.cell(row=1, column=col, value=naam)
        if n > 1:
            ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + n - 1)
        c = ws.cell(row=1, column=col)
        c.font = Font(bold=True, color="FFFFFF", size=11)
        c.fill = PatternFill("solid", fgColor=kleur)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = Border(top=_BORDER_THICK, left=_BORDER, right=_BORDER, bottom=_BORDER)
        col += n
    ws.row_dimensions[1].height = 24

    # Rij 2: kolomheaders
    headers_row2 = (invoer_kols + locatie_kols + bag_kols + buurt_kols
                    + label_kols + mon_kols + afgeleid_kols + woz_kols)
    for i, h in enumerate(headers_row2, start=1):
        c = ws.cell(row=2, column=i, value=h)
        c.font = Font(bold=True, color="1F2933", size=10)
        c.fill = PatternFill("solid", fgColor=_GREY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(left=_BORDER, right=_BORDER, bottom=_BORDER)
    ws.row_dimensions[2].height = 30

    # Data
    mismatched_rows = []
    for row, res in zip(rows[1:], resultaten):
        bag = res.get("bag") or {}
        ep = res.get("energielabel") or {}
        cbs = res.get("cbs") or {}
        mon = res.get("monument") or {}
        out_row = list(row)
        out_row += [None] * (sizes["in"] - len(out_row))
        out_row = out_row[:sizes["in"]]

        match_ok = res.get("adres_match")
        match_cell = "OK" if match_ok else (res.get("adres_match_reden") or "—")

        out_row += [
            res.get("status"),
            match_cell,
            res.get("weergavenaam"),
            res.get("postcode"),
            res.get("woonplaats"),
            res.get("adresseerbaarobject_id") or res.get("nummeraanduiding"),
        ]
        out_row += [bag.get("bouwjaar"), bag.get("gebruiksoppervlakte"), bag.get("gebruiksdoel")]
        out_row += [
            cbs.get("buurtnaam"),
            cbs.get("gem_woningwaarde"),
            cbs.get("gem_inkomen"),
            cbs.get("pct_huur"),
        ]
        out_row += [ep.get("energieklasse"), ep.get("gebouwtype"), ep.get("geldig_tot")]
        out_row += [mon.get("monumentnummer") or "—"]
        out_row += [
            res.get("woz_per_m2"),
            res.get("pct_1jr"),
            res.get("pct_5jr"),
            res.get("pct_sinds_oudst"),
        ]
        waarden_map = {w["peildatum"]: w["vastgesteldeWaarde"] for w in res.get("wozWaarden", [])}
        for p in peildata_sorted:
            out_row.append(waarden_map.get(p))
        ws.append(out_row)
        if match_ok is False:
            mismatched_rows.append(ws.max_row)

    max_row = ws.max_row
    max_col = len(headers_row2)

    # Borders + zebra
    for r in range(3, max_row + 1):
        for c_idx in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c_idx)
            cell.border = Border(left=_BORDER, right=_BORDER, bottom=_BORDER)
            if r % 2 == 1:
                cell.fill = PatternFill("solid", fgColor="FAFBFC")

    # Mismatch-rijen oranje
    mismatch_fill = PatternFill("solid", fgColor="FFE8D6")
    for r in mismatched_rows:
        for c_idx in range(1, max_col + 1):
            ws.cell(row=r, column=c_idx).fill = mismatch_fill
        # Match-cel donkerder kleuren
        match_col = sizes["in"] + 2
        c = ws.cell(row=r, column=match_col)
        c.fill = PatternFill("solid", fgColor="FF9800")
        c.font = Font(bold=True, color="FFFFFF")

    # Energielabel-cel kleuren
    label_col = sizes["in"] + sizes["loc"] + sizes["bag"] + sizes["buurt"] + 1
    for r in range(3, max_row + 1):
        cell = ws.cell(row=r, column=label_col)
        lab = (cell.value or "").strip() if isinstance(cell.value, str) else ""
        if lab in _LABEL_FILL:
            cell.fill = PatternFill("solid", fgColor=_LABEL_FILL[lab])
            light = lab in ("D",)
            cell.font = Font(bold=True, color="1F2933" if light else "FFFFFF", size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Monument-cel highlight als gevuld (niet —)
    mon_col = sizes["in"] + sizes["loc"] + sizes["bag"] + sizes["buurt"] + sizes["lab"] + 1
    for r in range(3, max_row + 1):
        cell = ws.cell(row=r, column=mon_col)
        if cell.value and cell.value != "—":
            cell.fill = PatternFill("solid", fgColor="7C3AED")
            cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")

    # BAG-bouwjaar + opp formattering
    bouwjaar_col = sizes["in"] + sizes["loc"] + 1
    opp_col = sizes["in"] + sizes["loc"] + 2
    for r in range(3, max_row + 1):
        ws.cell(row=r, column=bouwjaar_col).number_format = "0"
        ws.cell(row=r, column=bouwjaar_col).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=opp_col).number_format = "0"
        ws.cell(row=r, column=opp_col).alignment = Alignment(horizontal="center")

    # CBS-kolommen formattering
    buurt_start = sizes["in"] + sizes["loc"] + sizes["bag"] + 1
    # buurtnaam | gem WOZ | gem inkomen | % huur
    for r in range(3, max_row + 1):
        ws.cell(row=r, column=buurt_start + 1).number_format = '"€" #,##0'
        ws.cell(row=r, column=buurt_start + 2).number_format = '"€" #,##0'
        ws.cell(row=r, column=buurt_start + 3).number_format = "0\"%\""
        ws.cell(row=r, column=buurt_start + 3).alignment = Alignment(horizontal="center")

    # Afgeleide kolommen: €/m², %1jr, %5jr, %sinds
    afg_start = sizes["in"] + sizes["loc"] + sizes["bag"] + sizes["buurt"] + sizes["lab"] + sizes["mon"] + 1
    for r in range(3, max_row + 1):
        ws.cell(row=r, column=afg_start).number_format = '"€" #,##0'
        for off in (1, 2, 3):
            c = ws.cell(row=r, column=afg_start + off)
            c.number_format = "0.0\"%\""
            c.alignment = Alignment(horizontal="right")
    # Conditional formatting op %-kolommen: groen positief, rood negatief
    for off in (1, 2, 3):
        col_letter = get_column_letter(afg_start + off)
        rng = f"{col_letter}3:{col_letter}{max_row}"
        ws.conditional_formatting.add(rng, ColorScaleRule(
            start_type="num", start_value=-20, start_color="E57373",
            mid_type="num", mid_value=0, mid_color="FFFFFF",
            end_type="num", end_value=50, end_color="81C784",
        ))

    # Euro-format + ColorScale voor WOZ-kolommen
    woz_start = sizes["in"] + sizes["loc"] + sizes["bag"] + sizes["buurt"] + sizes["lab"] + sizes["mon"] + sizes["afg"] + 1
    woz_end = woz_start + sizes["woz"] - 1
    for c_idx in range(woz_start, woz_end + 1):
        for r in range(3, max_row + 1):
            ws.cell(row=r, column=c_idx).number_format = '"€" #,##0'
        if max_row >= 3:
            rng = f"{get_column_letter(c_idx)}3:{get_column_letter(c_idx)}{max_row}"
            ws.conditional_formatting.add(rng, ColorScaleRule(
                start_type="min", start_color="FFFFFF",
                mid_type="percentile", mid_value=50, mid_color="DCE6F2",
                end_type="max", end_color="9BB8DC",
            ))

    # Kolombreedtes
    def set_width(col_letter, w):
        ws.column_dimensions[col_letter].width = w

    # Invoer-kolommen: auto-fit op werkelijke inhoud (min 8, max 38)
    for i in range(1, sizes["in"] + 1):
        max_len = len(str(invoer_kols[i - 1]))
        for r in range(3, ws.max_row + 1):
            v = ws.cell(row=r, column=i).value
            if v is not None:
                max_len = max(max_len, len(str(v)))
        set_width(get_column_letter(i), max(8, min(max_len + 2, 38)))
    loc_widths = [11, 8, 38, 10, 16, 18]
    for i, w in enumerate(loc_widths):
        set_width(get_column_letter(sizes["in"] + 1 + i), w)
    bag_widths = [10, 10, 20]
    for i, w in enumerate(bag_widths):
        set_width(get_column_letter(sizes["in"] + sizes["loc"] + 1 + i), w)
    buurt_widths = [22, 14, 14, 9]
    for i, w in enumerate(buurt_widths):
        set_width(get_column_letter(sizes["in"] + sizes["loc"] + sizes["bag"] + 1 + i), w)
    lab_widths = [11, 22, 12]
    for i, w in enumerate(lab_widths):
        set_width(get_column_letter(sizes["in"] + sizes["loc"] + sizes["bag"] + sizes["buurt"] + 1 + i), w)
    set_width(get_column_letter(mon_col), 14)
    afg_widths = [10, 9, 9, 12]
    for i, w in enumerate(afg_widths):
        set_width(get_column_letter(afg_start + i), w)
    for i in range(sizes["woz"]):
        set_width(get_column_letter(woz_start + i), 12)

    ws.freeze_panes = ws.cell(row=3, column=sizes["in"] + 1)
    ws.auto_filter.ref = f"A2:{get_column_letter(max_col)}{max_row}"

    # ---- Tab 2: Samenvatting ----
    ws2 = wb.create_sheet("Samenvatting")
    ws2.column_dimensions["A"].width = 36
    ws2.column_dimensions["B"].width = 18

    title = ws2.cell(row=1, column=1, value="Samenvatting WOZ + energielabel")
    title.font = Font(bold=True, size=14, color="FFFFFF")
    title.fill = PatternFill("solid", fgColor=_RED_DARK)
    ws2.merge_cells("A1:B1")
    ws2.row_dimensions[1].height = 28
    ws2["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)

    totaal = len(resultaten)
    ok = sum(1 for r in resultaten if r.get("status") == "OK")
    niet_gevonden = sum(1 for r in resultaten if r.get("status") in ("Adres niet gevonden", "Geen nummeraanduiding"))
    geen_woz = totaal - ok - niet_gevonden

    woz_actueel = [r["wozWaarden"][0]["vastgesteldeWaarde"]
                   for r in resultaten if r.get("wozWaarden") and r["wozWaarden"][0].get("vastgesteldeWaarde")]
    gem_woz = int(sum(woz_actueel) / len(woz_actueel)) if woz_actueel else None
    min_woz = min(woz_actueel) if woz_actueel else None
    max_woz = max(woz_actueel) if woz_actueel else None

    labels_lijst = [(r.get("energielabel") or {}).get("energieklasse") for r in resultaten]
    labels_lijst = [l for l in labels_lijst if l]
    label_count = {}
    for l in labels_lijst:
        label_count[l] = label_count.get(l, 0) + 1

    # Bouwjaar uit BAG (autoritatief)
    bouwjaren = [(r.get("bag") or {}).get("bouwjaar") for r in resultaten]
    bouwjaren = [b for b in bouwjaren if isinstance(b, int) and b > 1500]
    gem_bj = int(sum(bouwjaren) / len(bouwjaren)) if bouwjaren else None
    oudste = min(bouwjaren) if bouwjaren else None
    nieuwste = max(bouwjaren) if bouwjaren else None

    opp = [(r.get("bag") or {}).get("gebruiksoppervlakte") for r in resultaten]
    opp = [o for o in opp if isinstance(o, (int, float)) and o > 0]
    gem_opp = round(sum(opp) / len(opp)) if opp else None
    min_opp = min(opp) if opp else None
    max_opp = max(opp) if opp else None

    laatste_peildatum = peildata_sorted[0] if peildata_sorted else "—"

    # Nieuw: WOZ per m²
    per_m2 = [r.get("woz_per_m2") for r in resultaten if r.get("woz_per_m2")]
    gem_per_m2 = int(sum(per_m2) / len(per_m2)) if per_m2 else None
    min_per_m2 = min(per_m2) if per_m2 else None
    max_per_m2 = max(per_m2) if per_m2 else None

    # Stijgingen
    p5 = [r.get("pct_5jr") for r in resultaten if r.get("pct_5jr") is not None]
    gem_p5 = round(sum(p5) / len(p5), 1) if p5 else None
    p_oudst = [r.get("pct_sinds_oudst") for r in resultaten if r.get("pct_sinds_oudst") is not None]
    gem_p_oudst = round(sum(p_oudst) / len(p_oudst), 1) if p_oudst else None

    # Mismatches
    mismatches = sum(1 for r in resultaten if r.get("adres_match") is False)

    # Monumenten
    monumenten = sum(1 for r in resultaten if (r.get("monument") or {}).get("monumentnummer"))

    blokken = [
        ("Aantal adressen", [
            ("Totaal verwerkt", totaal),
            ("Succesvol gevonden", ok),
            ("Adres-match-waarschuwingen", mismatches),
            ("Adres niet gevonden", niet_gevonden),
            ("Geen WOZ-waarde", geen_woz),
            ("Rijksmonumenten", monumenten),
        ]),
        (f"WOZ-waarde ({laatste_peildatum})", [
            ("Aantal met WOZ", len(woz_actueel)),
            ("Gemiddeld", gem_woz),
            ("Laagste", min_woz),
            ("Hoogste", max_woz),
        ]),
        ("WOZ per m² (BAG-opp.)", [
            ("Aantal met €/m²", len(per_m2)),
            ("Gemiddeld €/m²", gem_per_m2),
            ("Laagste €/m²", min_per_m2),
            ("Hoogste €/m²", max_per_m2),
        ]),
        ("WOZ-stijging (gemiddeld)", [
            ("% 5 jaar", gem_p5),
            (f"% sinds oudste peiljaar", gem_p_oudst),
        ]),
        ("Energielabels", [(f"Label {k}", v) for k, v in sorted(label_count.items())] or [("Geen labels gevonden", 0)]),
        ("Bouwjaar (BAG)", [
            ("Aantal met bouwjaar", len(bouwjaren)),
            ("Gemiddeld", gem_bj),
            ("Oudst", oudste),
            ("Nieuwst", nieuwste),
        ]),
        ("Gebruiksoppervlakte (BAG, m²)", [
            ("Aantal met oppervlakte", len(opp)),
            ("Gemiddeld", gem_opp),
            ("Kleinst", min_opp),
            ("Grootst", max_opp),
        ]),
    ]

    r = 3
    for titel, items in blokken:
        c = ws2.cell(row=r, column=1, value=titel)
        c.font = Font(bold=True, color="FFFFFF", size=11)
        c.fill = PatternFill("solid", fgColor=_RED_MID)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws2.row_dimensions[r].height = 20
        r += 1
        for k, v in items:
            ws2.cell(row=r, column=1, value=k).alignment = Alignment(indent=1)
            cv = ws2.cell(row=r, column=2, value=v)
            cv.alignment = Alignment(horizontal="right")
            if titel.startswith("WOZ-waarde") and k != "Aantal met WOZ":
                cv.number_format = '"€" #,##0'
            elif titel.startswith("WOZ per m²") and "€/m²" in k:
                cv.number_format = '"€" #,##0'
            elif titel.startswith("WOZ-stijging"):
                cv.number_format = "0.0\"%\""
            for col_idx in (1, 2):
                ws2.cell(row=r, column=col_idx).border = Border(bottom=_BORDER)
            r += 1
        r += 1  # lege tussen blokken

    # ---- Tabs 3-7: deelviews per onderwerp ----
    _bouw_subtabs(wb, resultaten, peildata_sorted)

    return wb


def _bouw_subtabs(wb: Workbook, resultaten: list, peildata_sorted: list) -> None:
    """Voegt aparte tabbladen toe per onderwerp, voor snel filteren per categorie."""
    _bouw_tab_woz(wb, resultaten, peildata_sorted)
    _bouw_tab_energielabel(wb, resultaten)
    _bouw_tab_bag(wb, resultaten)
    _bouw_tab_buurt(wb, resultaten)
    _bouw_tab_monumenten(wb, resultaten)
    _bouw_tab_matchwarnings(wb, resultaten)


def _adres_kolommen(res: dict) -> list:
    """Standaard set adres-identifier kolommen voor elke deelview."""
    return [
        res.get("postcode") or "",
        res.get("huisnummer") or "",
        res.get("huisletter") or "",
        res.get("huisnummertoevoeging") or "",
        res.get("straat") or "",
        res.get("woonplaats") or "",
        res.get("weergavenaam") or res.get("adres_invoer") or "",
    ]


_ADRES_HEADERS = ["postcode", "huisnr", "letter", "toev", "straat", "plaats", "gevonden adres"]


def _stijl_subtab_header(ws, kleur: str, kolomnamen: list, breedtes: list) -> None:
    """Past de standaard 2-rij header toe op een sub-tabblad."""
    titel = ws.title
    ws.cell(row=1, column=1, value=titel)
    n = len(kolomnamen)
    if n > 1:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    c = ws.cell(row=1, column=1)
    c.font = Font(bold=True, color="FFFFFF", size=12)
    c.fill = PatternFill("solid", fgColor=kleur)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 26

    for i, h in enumerate(kolomnamen, start=1):
        cell = ws.cell(row=2, column=i, value=h)
        cell.font = Font(bold=True, color="1F2933", size=10)
        cell.fill = PatternFill("solid", fgColor=_GREY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(left=_BORDER, right=_BORDER, bottom=_BORDER)
    ws.row_dimensions[2].height = 30

    for i, w in enumerate(breedtes, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _zebra_en_borders(ws, max_col: int) -> None:
    for r in range(3, ws.max_row + 1):
        for c_idx in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c_idx)
            cell.border = Border(left=_BORDER, right=_BORDER, bottom=_BORDER)
            if r % 2 == 1:
                cell.fill = PatternFill("solid", fgColor="FAFBFC")


def _bouw_tab_woz(wb, resultaten, peildata_sorted) -> None:
    ws = wb.create_sheet("WOZ-waardeloket")
    oudst_jaar = peildata_sorted[-1][:4] if peildata_sorted else "oudst"
    woz_jaren = [p[:4] for p in peildata_sorted]
    headers = _ADRES_HEADERS + ["€/m²", "% 1jr", "% 5jr", f"% sinds {oudst_jaar}"] + woz_jaren
    widths = [11, 8, 6, 8, 28, 16, 38] + [10, 9, 9, 12] + [12] * len(woz_jaren)
    _stijl_subtab_header(ws, _RED_MID, headers, widths)

    n_adres = len(_ADRES_HEADERS)
    n_afg = 4

    for res in resultaten:
        if not res.get("wozWaarden"):
            continue
        row = _adres_kolommen(res) + [
            res.get("woz_per_m2"),
            res.get("pct_1jr"),
            res.get("pct_5jr"),
            res.get("pct_sinds_oudst"),
        ]
        waarden_map = {w["peildatum"]: w["vastgesteldeWaarde"] for w in res.get("wozWaarden") or []}
        for p in peildata_sorted:
            row.append(waarden_map.get(p))
        ws.append(row)

    max_row = ws.max_row
    max_col = len(headers)
    _zebra_en_borders(ws, max_col)

    # Formats: €/m², percentages, WOZ-kolommen
    afg_start = n_adres + 1
    for r in range(3, max_row + 1):
        ws.cell(row=r, column=afg_start).number_format = '"€" #,##0'
        for off in (1, 2, 3):
            ws.cell(row=r, column=afg_start + off).number_format = "0.0\"%\""
            ws.cell(row=r, column=afg_start + off).alignment = Alignment(horizontal="right")

    # Diverging color-scale op trend-kolommen
    if max_row >= 3:
        for off in (1, 2, 3):
            col_letter = get_column_letter(afg_start + off)
            rng = f"{col_letter}3:{col_letter}{max_row}"
            ws.conditional_formatting.add(rng, ColorScaleRule(
                start_type="num", start_value=-20, start_color="E57373",
                mid_type="num", mid_value=0, mid_color="FFFFFF",
                end_type="num", end_value=50, end_color="81C784",
            ))

    # WOZ-kolommen euro + heatmap
    woz_start = afg_start + n_afg
    for c_idx in range(woz_start, max_col + 1):
        for r in range(3, max_row + 1):
            ws.cell(row=r, column=c_idx).number_format = '"€" #,##0'
        if max_row >= 3:
            rng = f"{get_column_letter(c_idx)}3:{get_column_letter(c_idx)}{max_row}"
            ws.conditional_formatting.add(rng, ColorScaleRule(
                start_type="min", start_color="FFFFFF",
                mid_type="percentile", mid_value=50, mid_color="DCE6F2",
                end_type="max", end_color="9BB8DC",
            ))

    ws.freeze_panes = ws.cell(row=3, column=n_adres + 1)
    if max_row >= 2:
        ws.auto_filter.ref = f"A2:{get_column_letter(max_col)}{max_row}"


def _bouw_tab_energielabel(wb, resultaten) -> None:
    ws = wb.create_sheet("Energielabel")
    headers = _ADRES_HEADERS + ["label", "bouwjaar (EP)", "gebouwklasse", "gebouwtype",
                                 "geldig t/m", "geregistreerd", "berekend verbruik"]
    widths = [11, 8, 6, 8, 28, 16, 38, 9, 12, 14, 22, 12, 14, 14]
    _stijl_subtab_header(ws, "1B7A3A", headers, widths)

    n_adres = len(_ADRES_HEADERS)
    for res in resultaten:
        ep = res.get("energielabel") or {}
        if not ep.get("energieklasse"):
            continue
        ws.append(_adres_kolommen(res) + [
            ep.get("energieklasse"),
            ep.get("bouwjaar"),
            ep.get("gebouwklasse"),
            ep.get("gebouwtype"),
            ep.get("geldig_tot"),
            ep.get("registratiedatum"),
            ep.get("berekend_energieverbruik"),
        ])

    max_row = ws.max_row
    max_col = len(headers)
    _zebra_en_borders(ws, max_col)

    # Kleur de label-cel
    label_col = n_adres + 1
    for r in range(3, max_row + 1):
        c = ws.cell(row=r, column=label_col)
        lab = (c.value or "").strip() if isinstance(c.value, str) else ""
        if lab in _LABEL_FILL:
            c.fill = PatternFill("solid", fgColor=_LABEL_FILL[lab])
            c.font = Font(bold=True, color="1F2933" if lab == "D" else "FFFFFF", size=11)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r, column=n_adres + 2).number_format = "0"
        ws.cell(row=r, column=n_adres + 2).alignment = Alignment(horizontal="center")

    ws.freeze_panes = ws.cell(row=3, column=n_adres + 1)
    if max_row >= 2:
        ws.auto_filter.ref = f"A2:{get_column_letter(max_col)}{max_row}"


def _bouw_tab_bag(wb, resultaten) -> None:
    ws = wb.create_sheet("BAG")
    headers = _ADRES_HEADERS + ["bouwjaar", "gebr.opp. (m²)", "gebruiksdoel",
                                 "vbo-status", "pand-id", "pand-status"]
    widths = [11, 8, 6, 8, 28, 16, 38, 10, 12, 22, 22, 18, 22]
    _stijl_subtab_header(ws, "8B5A2B", headers, widths)

    n_adres = len(_ADRES_HEADERS)
    for res in resultaten:
        bag = res.get("bag") or {}
        if not bag.get("bouwjaar") and not bag.get("gebruiksoppervlakte"):
            continue
        ws.append(_adres_kolommen(res) + [
            bag.get("bouwjaar"),
            bag.get("gebruiksoppervlakte"),
            bag.get("gebruiksdoel"),
            bag.get("vbo_status"),
            bag.get("pand_id"),
            bag.get("pand_status"),
        ])

    max_row = ws.max_row
    max_col = len(headers)
    _zebra_en_borders(ws, max_col)

    for r in range(3, max_row + 1):
        ws.cell(row=r, column=n_adres + 1).number_format = "0"
        ws.cell(row=r, column=n_adres + 1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=n_adres + 2).number_format = "0"
        ws.cell(row=r, column=n_adres + 2).alignment = Alignment(horizontal="center")

    ws.freeze_panes = ws.cell(row=3, column=n_adres + 1)
    if max_row >= 2:
        ws.auto_filter.ref = f"A2:{get_column_letter(max_col)}{max_row}"


def _bouw_tab_buurt(wb, resultaten) -> None:
    ws = wb.create_sheet("CBS-buurt")
    headers = _ADRES_HEADERS + ["buurt", "gemeente", "inwoners", "dichtheid /km²",
                                 "gem. WOZ buurt", "gem. inkomen", "% huur", "% koop",
                                 "huishoudgrootte", "woningvoorraad"]
    widths = [11, 8, 6, 8, 28, 16, 38, 22, 16, 10, 13, 14, 14, 8, 8, 13, 13]
    _stijl_subtab_header(ws, "0F766E", headers, widths)

    n_adres = len(_ADRES_HEADERS)
    for res in resultaten:
        cbs = res.get("cbs") or {}
        if not cbs.get("buurtnaam"):
            continue
        ws.append(_adres_kolommen(res) + [
            cbs.get("buurtnaam"),
            cbs.get("gemeentenaam"),
            cbs.get("aantal_inwoners"),
            cbs.get("bevolkingsdichtheid"),
            cbs.get("gem_woningwaarde"),
            cbs.get("gem_inkomen"),
            cbs.get("pct_huur"),
            cbs.get("pct_koop"),
            cbs.get("huishoudgrootte"),
            cbs.get("woningvoorraad"),
        ])

    max_row = ws.max_row
    max_col = len(headers)
    _zebra_en_borders(ws, max_col)

    for r in range(3, max_row + 1):
        ws.cell(row=r, column=n_adres + 5).number_format = '"€" #,##0'  # gem WOZ
        ws.cell(row=r, column=n_adres + 6).number_format = '"€" #,##0'  # gem inkomen
        for off in (7, 8):
            ws.cell(row=r, column=n_adres + off).number_format = "0\"%\""
            ws.cell(row=r, column=n_adres + off).alignment = Alignment(horizontal="center")

    ws.freeze_panes = ws.cell(row=3, column=n_adres + 1)
    if max_row >= 2:
        ws.auto_filter.ref = f"A2:{get_column_letter(max_col)}{max_row}"


def _bouw_tab_monumenten(wb, resultaten) -> None:
    ws = wb.create_sheet("Monumenten")
    headers = _ADRES_HEADERS + ["monumentnummer", "RCE straat", "RCE huisnr"]
    widths = [11, 8, 6, 8, 28, 16, 38, 14, 22, 12]
    _stijl_subtab_header(ws, "7C3AED", headers, widths)

    n_adres = len(_ADRES_HEADERS)
    for res in resultaten:
        mon = res.get("monument") or {}
        if not mon.get("monumentnummer"):
            continue
        ws.append(_adres_kolommen(res) + [
            mon.get("monumentnummer"),
            mon.get("monument_straat"),
            mon.get("monument_huisnummer"),
        ])

    max_row = ws.max_row
    max_col = len(headers)
    _zebra_en_borders(ws, max_col)

    # Kleur monumentnummer-cel
    for r in range(3, max_row + 1):
        c = ws.cell(row=r, column=n_adres + 1)
        c.fill = PatternFill("solid", fgColor="7C3AED")
        c.font = Font(bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center")

    ws.freeze_panes = ws.cell(row=3, column=n_adres + 1)
    if max_row >= 2:
        ws.auto_filter.ref = f"A2:{get_column_letter(max_col)}{max_row}"

    if max_row < 3:
        info = ws.cell(row=3, column=1, value="Geen rijksmonumenten in deze batch.")
        info.font = Font(italic=True, color="6B7785")
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=max_col)


def _bouw_tab_matchwarnings(wb, resultaten) -> None:
    """Lijst van rijen waar match-status of straat-verificatie iets opviel."""
    ws = wb.create_sheet("Match-controle")
    headers = ["status", "match-modus", "adres-invoer", "postcode", "huisnr",
               "letter", "toev", "gevonden straat", "straat-invoer",
               "straat verschilt?", "fallback?", "fallback-reden"]
    widths = [22, 14, 38, 11, 8, 6, 8, 22, 22, 9, 9, 30]
    _stijl_subtab_header(ws, "B45309", headers, widths)

    for res in resultaten:
        status = res.get("status", "")
        modus = res.get("match_modus", "")
        verschil = bool(res.get("straat_verschilt"))
        fallback = modus == "fuzzy-fallback"
        # Toon alleen rijen die opvallend zijn
        if status == "OK" and not verschil and not fallback and modus == "strict":
            continue
        ws.append([
            status,
            modus,
            res.get("adres_invoer", ""),
            res.get("postcode") or "",
            res.get("huisnummer") or "",
            res.get("huisletter") or "",
            res.get("huisnummertoevoeging") or "",
            res.get("straat") or "",
            res.get("straat_invoer") or "",
            "ja" if verschil else "",
            "ja" if fallback else "",
            res.get("fallback_reden") or "",
        ])

    max_row = ws.max_row
    max_col = len(headers)
    _zebra_en_borders(ws, max_col)

    # Markeer rijen met straat-verschil oranje
    verschil_col = 10
    fallback_col = 11
    for r in range(3, max_row + 1):
        if ws.cell(row=r, column=verschil_col).value == "ja":
            for c_idx in range(1, max_col + 1):
                ws.cell(row=r, column=c_idx).fill = PatternFill("solid", fgColor="FFE8D6")
        if ws.cell(row=r, column=fallback_col).value == "ja":
            ws.cell(row=r, column=fallback_col).fill = PatternFill("solid", fgColor="FFF4D6")
            ws.cell(row=r, column=fallback_col).font = Font(bold=True, color="8a5a00")

    ws.freeze_panes = ws.cell(row=3, column=4)
    if max_row >= 2:
        ws.auto_filter.ref = f"A2:{get_column_letter(max_col)}{max_row}"

    if max_row < 3:
        info = ws.cell(row=3, column=1, value="Alle adressen strict op postcode+huisnummer gematcht, geen afwijkingen.")
        info.font = Font(italic=True, color="1f7a3a")
        ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=max_col)


@app.route("/api/template")
def api_template():
    """Genereer een lege voorbeeld-Excel die gebruiker kan invullen."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Adressen"
    ws.append(["adres"])
    voorbeelden = [
        "Witte de Withstraat 1D, 3012BK Rotterdam",
        "Vondelstraat 70, 1054GG Amsterdam",
        "Bloemstraat 1, 3581WB Utrecht",
    ]
    for v in voorbeelden:
        ws.append([v])

    ws["A1"].font = Font(bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="1F4E79")
    ws.column_dimensions["A"].width = 55

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="woz_template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# Realworks-koppeling
# ---------------------------------------------------------------------------


@app.route("/realworks")
def realworks_page():
    """HTML-pagina met overzicht van Realworks-objecten."""
    return render_template(
        "realworks.html",
        geconfigureerd=realworks_service.is_geconfigureerd(),
    )


@app.route("/api/realworks")
def api_realworks():
    """JSON-endpoint: één pagina Realworks-objecten."""
    try:
        vanaf = int(request.args.get("vanaf", 0))
        aantal = int(request.args.get("aantal", 100))
    except ValueError:
        return jsonify({"error": "Parameter 'vanaf' en 'aantal' moeten getallen zijn."}), 400
    status = (request.args.get("status") or "").strip() or None

    try:
        page = realworks_service.haal_objecten(vanaf=vanaf, aantal=aantal, status=status)
    except realworks_service.RealworksError as e:
        return jsonify({"error": str(e)}), e.status_code or 502
    return jsonify(page)


@app.route("/api/realworks/alle")
def api_realworks_alle():
    """JSON-endpoint: alle objecten via auto-paginatie (gecapt)."""
    status = (request.args.get("status") or "").strip() or None
    try:
        max_obj = int(request.args.get("max", 500))
    except ValueError:
        max_obj = 500
    try:
        result = realworks_service.haal_alle_objecten(status=status, max_objecten=max_obj)
    except realworks_service.RealworksError as e:
        return jsonify({"error": str(e)}), e.status_code or 502
    return jsonify(result)


@app.route("/api/realworks/verrijk", methods=["POST"])
def api_realworks_verrijk():
    """Neemt een lijst Realworks-objecten en draait WOZ/BAG/label/CBS/monument."""
    payload = request.get_json(silent=True) or {}
    objecten = payload.get("objecten") or []
    if not objecten:
        return jsonify({"error": "Geen objecten meegegeven."}), 400

    verrijkt = []
    for obj in objecten:
        adres = realworks_service.adres_string(obj)
        res = maak_resultaat(adres) if adres else {"adres_invoer": "", "status": "Leeg"}
        verrijkt.append({"realworks": obj, "resultaat": res})
        time.sleep(0.15)
    return jsonify({"verrijkt": verrijkt, "aantal": len(verrijkt)})


if __name__ == "__main__":
    print("\n  WOZ-tool gestart op  http://127.0.0.1:5005\n")
    app.run(host="127.0.0.1", port=5005, debug=False)
