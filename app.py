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

REQUEST_TIMEOUT = 15
session = requests.Session()


def zoek_adres(query: str) -> Optional[dict]:
    """Zoekt een adres via PDOK Locatieserver. Returnt eerste hit of None."""
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


def maak_resultaat(adres: str) -> dict:
    """Combineert adres-resolve + WOZ-call + energielabel tot één resultaat-dict."""
    hit = zoek_adres(adres)
    if not hit:
        return {"adres_invoer": adres, "status": "Adres niet gevonden", "wozWaarden": []}

    weergavenaam = hit.get("weergavenaam", "")
    nummeraanduiding = hit.get("nummeraanduiding_id")
    adresseerbaar_object_id = hit.get("adresseerbaarobject_id")
    if not nummeraanduiding:
        return {"adres_invoer": adres, "status": "Geen nummeraanduiding", "wozWaarden": []}

    woz = haal_woz(nummeraanduiding)
    bag = haal_bag(adresseerbaar_object_id) if adresseerbaar_object_id else None
    energielabel = haal_energielabel(adresseerbaar_object_id) if adresseerbaar_object_id else None

    if not woz:
        return {
            "adres_invoer": adres,
            "weergavenaam": weergavenaam,
            "nummeraanduiding": nummeraanduiding,
            "adresseerbaarobject_id": adresseerbaar_object_id,
            "status": "Geen WOZ-waarde bekend (niet-woning of onbekend)",
            "wozWaarden": [],
            "bag": bag,
            "energielabel": energielabel,
        }

    obj = woz.get("wozObject", {}) or {}
    waarden = woz.get("wozWaarden", []) or []

    return {
        "adres_invoer": adres,
        "weergavenaam": weergavenaam,
        "nummeraanduiding": nummeraanduiding,
        "adresseerbaarobject_id": adresseerbaar_object_id,
        "status": "OK",
        "wozobjectnummer": obj.get("wozobjectnummer"),
        "straat": obj.get("openbareruimtenaam"),
        "huisnummer": obj.get("huisnummer"),
        "huisletter": obj.get("huisletter"),
        "huisnummertoevoeging": obj.get("huisnummertoevoeging"),
        "postcode": obj.get("postcode"),
        "woonplaats": obj.get("woonplaatsnaam"),
        "grondoppervlakte": obj.get("grondoppervlakte"),
        "wozWaarden": sorted(
            [
                {"peildatum": w.get("peildatum"), "vastgesteldeWaarde": w.get("vastgesteldeWaarde")}
                for w in waarden
            ],
            key=lambda x: x["peildatum"] or "",
            reverse=True,
        ),
        "bag": bag,
        "energielabel": energielabel,
    }


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

    straat_idx = next((i for i, h in enumerate(header_lower) if h in ("straat", "straatnaam", "openbareruimte")), None)
    huisnummer_idx = next((i for i, h in enumerate(header_lower) if h in ("huisnummer", "nr", "nummer")), None)
    toevoeging_idx = next((i for i, h in enumerate(header_lower) if h in ("toevoeging", "huisletter", "huisnummertoevoeging")), None)
    postcode_idx = next((i for i, h in enumerate(header_lower) if h in ("postcode", "pc")), None)
    plaats_idx = next((i for i, h in enumerate(header_lower) if h in ("plaats", "woonplaats", "stad", "gemeente")), None)

    def adres_uit_rij(row):
        if adres_idx is not None and row[adres_idx]:
            return str(row[adres_idx]).strip()
        parts = []
        if straat_idx is not None and row[straat_idx]:
            parts.append(str(row[straat_idx]).strip())
        if huisnummer_idx is not None and row[huisnummer_idx] not in (None, ""):
            parts.append(str(row[huisnummer_idx]).strip())
        if toevoeging_idx is not None and row[toevoeging_idx]:
            parts.append(str(row[toevoeging_idx]).strip())
        if postcode_idx is not None and row[postcode_idx]:
            parts.append(str(row[postcode_idx]).strip())
        if plaats_idx is not None and row[plaats_idx]:
            parts.append(str(row[plaats_idx]).strip())
        return " ".join(parts).strip()

    # Verzamel resultaten en bepaal welke peildata voorkomen
    resultaten = []
    peildata = set()
    for row in rows[1:]:
        adres = adres_uit_rij(row)
        if not adres:
            resultaten.append({"adres_invoer": "", "status": "Leeg", "wozWaarden": []})
            continue
        res = maak_resultaat(adres)
        for w in res["wozWaarden"]:
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

    invoer_kols = list(header_in) or ["adres"]
    locatie_kols = ["status", "gevonden adres", "postcode", "woonplaats", "BAG-id"]
    bag_kols = ["bouwjaar", "gebr.oppervlakte (m²)", "gebruiksdoel"]
    label_kols = ["energielabel", "gebouwtype", "geldig t/m"]
    woz_kols = [p[:4] for p in peildata_sorted]  # alleen jaartal

    n_in = len(invoer_kols)
    n_loc = len(locatie_kols)
    n_bag = len(bag_kols)
    n_lab = len(label_kols)
    n_woz = len(woz_kols)

    # Rij 1: groepheaders (merged)
    groepen = [("Invoer", n_in, "6B7785"),
               ("Adres & WOZ-object", n_loc, _RED_DARK),
               ("BAG-gegevens", n_bag, "8B5A2B"),
               ("Energielabel (EP-online)", n_lab, "1B7A3A"),
               ("WOZ-waarden", n_woz, _RED_MID)]
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
    headers_row2 = invoer_kols + locatie_kols + bag_kols + label_kols + woz_kols
    for i, h in enumerate(headers_row2, start=1):
        c = ws.cell(row=2, column=i, value=h)
        c.font = Font(bold=True, color="1F2933", size=10)
        c.fill = PatternFill("solid", fgColor=_GREY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(left=_BORDER, right=_BORDER, bottom=_BORDER)
    ws.row_dimensions[2].height = 30

    # Data
    for row, res in zip(rows[1:], resultaten):
        bag = res.get("bag") or {}
        ep = res.get("energielabel") or {}
        out_row = list(row)
        # Pad indien rij korter dan header
        out_row += [None] * (n_in - len(out_row))
        out_row = out_row[:n_in]
        out_row += [
            res.get("status"),
            res.get("weergavenaam"),
            res.get("postcode"),
            res.get("woonplaats"),
            res.get("adresseerbaarobject_id") or res.get("nummeraanduiding"),
            bag.get("bouwjaar"),
            bag.get("gebruiksoppervlakte"),
            bag.get("gebruiksdoel"),
            ep.get("energieklasse"),
            ep.get("gebouwtype"),
            ep.get("geldig_tot"),
        ]
        waarden_map = {w["peildatum"]: w["vastgesteldeWaarde"] for w in res.get("wozWaarden", [])}
        for p in peildata_sorted:
            out_row.append(waarden_map.get(p))
        ws.append(out_row)

    max_row = ws.max_row
    max_col = len(headers_row2)

    # Borders op alle cellen
    for r in range(3, max_row + 1):
        for c_idx in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c_idx)
            cell.border = Border(left=_BORDER, right=_BORDER, bottom=_BORDER)
            if r % 2 == 1:
                cell.fill = PatternFill("solid", fgColor="FAFBFC")

    # Kleur energielabel-cel (col na invoer+locatie+bag)
    label_col_idx = n_in + n_loc + n_bag + 1  # 1-based
    for r in range(3, max_row + 1):
        cell = ws.cell(row=r, column=label_col_idx)
        lab = (cell.value or "").strip() if isinstance(cell.value, str) else ""
        if lab in _LABEL_FILL:
            cell.fill = PatternFill("solid", fgColor=_LABEL_FILL[lab])
            light = lab in ("D",)
            cell.font = Font(bold=True, color="1F2933" if light else "FFFFFF", size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Euro-format + ColorScale voor WOZ-kolommen
    woz_start = n_in + n_loc + n_bag + n_lab + 1
    woz_end = woz_start + n_woz - 1
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

    # BAG bouwjaar als getal zonder duizendtal, gebruiksoppervlakte centered
    bouwjaar_col = n_in + n_loc + 1
    oppervlakte_col = n_in + n_loc + 2
    for r in range(3, max_row + 1):
        ws.cell(row=r, column=bouwjaar_col).number_format = "0"
        ws.cell(row=r, column=bouwjaar_col).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=oppervlakte_col).number_format = "0"
        ws.cell(row=r, column=oppervlakte_col).alignment = Alignment(horizontal="center")

    # Kolombreedtes — zinvolle defaults
    def set_width(col_letter, w):
        ws.column_dimensions[col_letter].width = w

    for i in range(1, n_in + 1):
        set_width(get_column_letter(i), 38)
    locatie_widths = [11, 38, 10, 18, 18]
    for i, w in enumerate(locatie_widths):
        set_width(get_column_letter(n_in + 1 + i), w)
    bag_widths = [10, 12, 22]
    for i, w in enumerate(bag_widths):
        set_width(get_column_letter(n_in + n_loc + 1 + i), w)
    label_widths = [11, 22, 12]
    for i, w in enumerate(label_widths):
        set_width(get_column_letter(n_in + n_loc + n_bag + 1 + i), w)
    for i in range(n_woz):
        set_width(get_column_letter(woz_start + i), 12)

    ws.freeze_panes = ws.cell(row=3, column=n_in + 1)
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

    blokken = [
        ("Aantal adressen", [
            ("Totaal verwerkt", totaal),
            ("Succesvol gevonden", ok),
            ("Adres niet gevonden", niet_gevonden),
            ("Geen WOZ-waarde", geen_woz),
        ]),
        (f"WOZ-waarde ({laatste_peildatum})", [
            ("Aantal met WOZ", len(woz_actueel)),
            ("Gemiddeld", gem_woz),
            ("Laagste", min_woz),
            ("Hoogste", max_woz),
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
            if titel.startswith("WOZ") and k != "Aantal met WOZ":
                cv.number_format = '"€" #,##0'
            for col_idx in (1, 2):
                ws2.cell(row=r, column=col_idx).border = Border(bottom=_BORDER)
            r += 1
        r += 1  # lege tussen blokken

    return wb


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


if __name__ == "__main__":
    print("\n  WOZ-tool gestart op  http://127.0.0.1:5005\n")
    app.run(host="127.0.0.1", port=5005, debug=False)
