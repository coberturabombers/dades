#!/usr/bin/env python3
"""
parse_tri.py — Motor de lectura del TRI (Assemblea Bombers de Catalunya)
========================================================================

Llegeix el fitxer Excel "Personal de guàrdia" (el TRI diari), n'extreu NOMÉS les
dades de cobertura (parc, mínim de torn, efectius reals) de les 7 regions, i
desa una "foto" del dia a data/history.json.

Pensat per executar-se cada dia automàticament (GitHub Actions).

IMPORTANT SOBRE PRIVADESA: aquest script descarta deliberadament tota la resta
d'informació de l'Excel (noms de comandaments, menús, vehicles, serveis...).
Només s'extreuen i es desen les tres dades de cobertura per parc.

Ús:
    python parse_tri.py --input <fitxer.xlsx> [--date YYYY-MM-DD]
    python parse_tri.py --input entrada/TRI.xlsx [--date YYYY-MM-DD]
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime, date

try:
    import openpyxl
except ImportError:
    print("ERROR: cal instal·lar openpyxl  ->  pip install openpyxl", file=sys.stderr)
    sys.exit(1)

# --- Configuració fixa de l'estructura de l'Excel -------------------------

# Columna on comença cada bloc de regió al full "Dades" (col PARC de cada regió)
REGION_START_COLS = {2: "REC", 14: "REG", 26: "REMN", 38: "REL",
                     50: "REMS", 62: "RET", 74: "RETE"}
REGION_ORDER = ["REC", "REG", "REMN", "REL", "REMS", "RET", "RETE"]
REGION_NAMES = {
    "REC": "Centre", "REG": "Girona", "REMN": "Metropolitana Nord",
    "REL": "Lleida", "REMS": "Metropolitana Sud", "RET": "Tarragona",
    "RETE": "Terres de l'Ebre",
}

# Parcs sense codi numèric al davant (excepcions conegudes)
KNOWN_NOCODE = {"GROS"}

# Suma de mínims esperada per regió (dada FIXA i verificada). Serveix de
# checksum: si no quadra, el fitxer ha canviat d'estructura i avisem.
MIN_CHECKSUM = {"REC": 45, "REG": 71, "REMN": 69, "REL": 45,
                "REMS": 55, "RET": 42, "RETE": 27}

# Coordenades dels parcs (aproximades on calgui). Clau = codi TRI.
PARC_COORDS = {
    "03 BER": (42.104, 1.845), "08 GUA": (42.203, 1.857), "15 PUIG": (41.974, 1.881),
    "19 SOLS": (41.995, 1.517), "04 CAF": (41.734, 1.512), "05 CAR": (41.913, 1.681),
    "10 MAN": (41.728, 1.827), "11 MOI": (41.812, 2.096), "14 PRA": (42.010, 2.032),
    "20 TOR": (42.049, 2.259), "21 VIC": (41.930, 2.254),
    "10 FIGU": (42.267, 2.961), "14 LLAN": (42.363, 3.152), "22 ROSE": (42.262, 3.176),
    "26 TORR": (42.043, 3.127), "11 GIRO": (41.983, 2.824), "17 OLOT": (42.181, 2.489),
    "21 RIPO": (42.201, 2.190), "03 BANY": (42.119, 2.767), "15 LLOR": (41.700, 2.845),
    "16 MAÇA": (41.777, 2.734), "25 SCFA": (41.860, 2.669), "01 AMER": (41.968, 2.601),
    "18 PALA": (41.918, 3.163), "28 VALL": (41.818, 3.033), "04 PERA": (41.981, 2.965),
    "08 CASS": (41.885, 2.874),
    "GROS": (41.386, 2.170), "13 RUB": (41.493, 2.033), "14 SAB": (41.548, 2.107),
    "20 TER": (41.563, 2.010), "06 GRA": (41.608, 2.288), "10 MOL": (41.540, 2.213),
    "16 SCE": (41.689, 2.491), "02 BAD": (41.450, 2.247), "18 SCG": (41.452, 2.208),
    "09 MAT": (41.538, 2.445), "12 PIN": (41.627, 2.689),
    "09 CERV": (41.671, 1.272), "19 LLEI": (41.615, 0.626), "20 MOLL": (41.632, 0.895),
    "31 TAR": (41.647, 1.140), "06 BALA": (41.790, 0.807), "29 SEU": (42.358, 1.461),
    "25 PONT": (42.408, 0.741), "30 SORT": (42.412, 1.128), "33 TREM": (42.166, 0.895),
    "04 COR": (41.354, 2.070), "05 GAV": (41.306, 2.000), "07 HOS": (41.360, 2.100),
    "12 PLL": (41.327, 2.095), "13 SBOI": (41.345, 2.037), "15 SFE": (41.383, 2.044),
    "08 IGU": (41.579, 1.617), "10 MAR": (41.474, 1.930), "18 VIF": (41.346, 1.699),
    "19 VIL": (41.224, 1.725),
    "05 MONT": (41.375, 1.161), "13 VALS": (41.287, 1.249), "12 TARR": (41.119, 1.245),
    "14 VEND": (41.220, 1.535), "02 CAMB": (41.067, 1.058), "04 FALS": (41.145, 0.821),
    "05 HOSP": (40.995, 0.938), "08 REUS": (41.156, 1.107),
    "22 ASCO": (41.183, 0.564), "26 GAND": (41.053, 0.437), "28 MÓRA": (41.093, 0.643),
    "20 AMET": (40.884, 0.802), "21 AMPO": (40.708, 0.581), "30 TORT": (40.812, 0.521),
    "31 ULLD": (40.596, 0.451),
}

PARC_NAMES = {
    "03 BER": "Berga", "08 GUA": "Guardiola de B.", "15 PUIG": "Puig-reig",
    "19 SOLS": "Solsona", "04 CAF": "Calaf", "05 CAR": "Cardona", "10 MAN": "Manresa",
    "11 MOI": "Moià", "14 PRA": "Prats de Lluçanés", "20 TOR": "Torelló", "21 VIC": "Vic",
    "10 FIGU": "Figueres", "14 LLAN": "Llançà", "22 ROSE": "Roses",
    "26 TORR": "Torroella de Montgrí", "11 GIRO": "Girona", "17 OLOT": "Olot",
    "21 RIPO": "Ripoll", "03 BANY": "Banyoles", "15 LLOR": "Lloret de Mar",
    "16 MAÇA": "Maçanet", "25 SCFA": "Santa Coloma de Farners", "01 AMER": "Amer",
    "18 PALA": "Palafrugell", "28 VALL": "Vall d'Aro", "04 PERA": "La Pera",
    "08 CASS": "Cassà de la Selva",
    "GROS": "Barcelona (Gros)", "13 RUB": "Rubí", "14 SAB": "Sabadell",
    "20 TER": "Terrassa", "06 GRA": "Granollers", "10 MOL": "Mollet",
    "16 SCE": "Sant Celoni", "02 BAD": "Badalona", "18 SCG": "Santa Coloma de Gramanet",
    "09 MAT": "Mataró", "12 PIN": "Pineda",
    "09 CERV": "Cervera", "19 LLEI": "Lleida", "20 MOLL": "Mollerussa", "31 TAR": "Tàrrega",
    "06 BALA": "Balaguer", "29 SEU": "La Seu d'Urgell", "25 PONT": "El Pont de Suert",
    "30 SORT": "Sort", "33 TREM": "Tremp",
    "04 COR": "Cornellà de Ll.", "05 GAV": "Gavà", "07 HOS": "L'Hospitalet de Ll.",
    "12 PLL": "El Prat de Llobregat", "13 SBOI": "Sant Boi de Ll.", "15 SFE": "S. Feliu de Ll.",
    "08 IGU": "Igualada", "10 MAR": "Martorell", "18 VIF": "Vilafranca del P.",
    "19 VIL": "Vilanova i la Geltrú",
    "05 MONT": "Montblanc", "13 VALS": "Valls", "12 TARR": "Tarragona",
    "14 VEND": "El Vendrell", "02 CAMB": "Cambrils", "04 FALS": "Falset",
    "05 HOSP": "L'Hospitalet de l'Infant", "08 REUS": "Reus",
    "22 ASCO": "Ascó", "26 GAND": "Gandesa", "28 MÓRA": "Móra d'Ebre",
    "20 AMET": "L'Ametlla de Mar", "21 AMPO": "Amposta", "30 TORT": "Tortosa",
    "31 ULLD": "Ulldecona",
}


def is_parc(value):
    if not value:
        return False
    s = str(value).strip()
    if s in KNOWN_NOCODE:
        return True
    return bool(re.match(r"^\d{2}\s+[A-ZÀ-Ú]", s))


def extract_excel_date(wb):
    """Intenta llegir la data del full Resum; si no, retorna None."""
    try:
        ws = wb["Resum"]
        v = ws.cell(row=4, column=2).value
        if isinstance(v, datetime):
            return v.date().isoformat()
        if isinstance(v, date):
            return v.isoformat()
        if isinstance(v, str):
            m = re.search(r"(\d{4})-(\d{2})-(\d{2})", v)
            if m:
                return m.group(0)
    except Exception:
        pass
    return None


def parse_workbook(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Dades"]

    excel_date = extract_excel_date(wb)
    parcs = []
    warnings = []

    for start_col, region in REGION_START_COLS.items():
        region_min_sum = 0
        for r in range(5, 60):
            code = ws.cell(row=r, column=start_col).value
            if not is_parc(code):
                continue
            code = str(code).strip()
            mn = ws.cell(row=r, column=start_col + 1).value
            rl = ws.cell(row=r, column=start_col + 2).value
            mn = int(mn) if isinstance(mn, (int, float)) else None
            rl = int(rl) if isinstance(rl, (int, float)) else None
            if mn is not None:
                region_min_sum += mn
            lat, lon = PARC_COORDS.get(code, (None, None))
            parcs.append({
                "code": code,
                "name": PARC_NAMES.get(code, code),
                "region": region,
                "min": mn,
                "real": rl,
                "lat": lat,
                "lon": lon,
            })
        # checksum de mínims
        expected = MIN_CHECKSUM.get(region)
        if expected is not None and region_min_sum != expected:
            warnings.append(
                f"ATENCIÓ: la suma de mínims de {region} és {region_min_sum}, "
                f"s'esperava {expected}. Pot indicar un canvi d'estructura de l'Excel.")

    return excel_date, parcs, warnings


def build_snapshot(the_date, parcs):
    """Construeix la foto del dia amb els totals derivats."""
    regions = {rc: {"sobre_minims": 0, "minims": 0, "inacceptable": 0,
                    "greu": 0, "critic": 0, "tancat": 0} for rc in REGION_ORDER}
    total_min = 0
    total_real = 0
    sota_minims = 0
    tancats = 0

    for p in parcs:
        rc = p["region"]
        mn, rl = p["min"], p["real"]
        if mn is None or rl is None:
            continue
        total_min += mn
        total_real += rl
        falten = mn - rl
        if rl == 0:
            regions[rc]["tancat"] += 1
            tancats += 1
            sota_minims += 1
        elif falten <= 0:
            if rl > mn:
                regions[rc]["sobre_minims"] += 1
            else:
                regions[rc]["minims"] += 1
        else:
            sota_minims += 1
            if falten == 1:
                regions[rc]["inacceptable"] += 1
            elif falten == 2:
                regions[rc]["greu"] += 1
            else:
                regions[rc]["critic"] += 1

    return {
        "date": the_date,
        "parcs": parcs,
        "regions": regions,
        "catalunya": {
            "min_total": total_min,
            "real_total": total_real,
            "diferencia": total_real - total_min,
            "sota_minims": sota_minims,
            "tancats": tancats,
            "n_parcs": len([p for p in parcs if p["real"] is not None]),
        },
    }


def _download_from_drive(url, dest):
    """Descarrega un fitxer d'un enllaç. Gestiona el cas de Google Drive, que per a
    fitxers pot mostrar una pàgina de confirmació en lloc del fitxer directament."""
    import urllib.request
    import urllib.parse
    import re as _re
    import http.cookiejar

    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
    opener.addheaders = [("User-Agent", "Mozilla/5.0")]

    resp = opener.open(url)
    data = resp.read()

    # Si ens han tornat un HTML (pàgina de confirmació de Google Drive), buscar el
    # token de confirmació o l'enllaç real i tornar-ho a demanar.
    head = data[:2000].lstrip()[:15].lower()
    if head.startswith(b"<!doctype html") or head.startswith(b"<html"):
        text = data.decode("utf-8", "replace")
        # cas nou de Drive: formulari amb action a drive.usercontent.google.com
        m_action = _re.search(r'action="([^"]+)"', text)
        confirm = _re.search(r'name="confirm"\s+value="([^"]+)"', text)
        uuid = _re.search(r'name="uuid"\s+value="([^"]+)"', text)
        idm = _re.search(r'name="id"\s+value="([^"]+)"', text)
        if m_action and confirm:
            params = {"confirm": confirm.group(1)}
            if uuid:
                params["uuid"] = uuid.group(1)
            if idm:
                params["id"] = idm.group(1)
            action = m_action.group(1).replace("&amp;", "&")
            new_url = action + ("&" if "?" in action else "?") + urllib.parse.urlencode(params)
            data = opener.open(new_url).read()
        else:
            # cas antic: paràmetre confirm=XXXX
            m = _re.search(r"confirm=([0-9A-Za-z_\-]+)", text)
            if m:
                sep = "&" if "?" in url else "?"
                data = opener.open(url + sep + "confirm=" + m.group(1)).read()

    with open(dest, "wb") as f:
        f.write(data)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", help="Fitxer .xlsx local")
    ap.add_argument("--url", help="Enllaç de descàrrega directa (Google Drive) del .xlsx")
    ap.add_argument("--date", help="Data a assignar (YYYY-MM-DD). Si no, s'intenta llegir de l'Excel o s'usa avui.")
    ap.add_argument("--history", default="data/history.json", help="Fitxer d'històric a actualitzar")
    args = ap.parse_args()

    # Obtenir el fitxer
    xlsx_path = args.input
    if args.url:
        xlsx_path = "/tmp/_tri_download.xlsx"
        print("Descarregant Excel des de l'enllaç...")
        _download_from_drive(args.url, xlsx_path)
    if not xlsx_path or not os.path.exists(xlsx_path):
        print("ERROR: cal --input <fitxer> o --url <enllaç> vàlid.", file=sys.stderr)
        sys.exit(1)

    excel_date, parcs, warnings = parse_workbook(xlsx_path)
    for w in warnings:
        print(w, file=sys.stderr)

    the_date = args.date or excel_date or date.today().isoformat()
    print(f"Data de la foto: {the_date}  ({len(parcs)} parcs llegits)")

    snapshot = build_snapshot(the_date, parcs)

    # Carregar històric existent i afegir/actualitzar el dia
    history = {}
    if os.path.exists(args.history):
        with open(args.history, encoding="utf-8") as f:
            history = json.load(f)
    history[the_date] = snapshot

    os.makedirs(os.path.dirname(args.history) or ".", exist_ok=True)
    with open(args.history, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, separators=(",", ":"))

    # També desem un fitxer "latest.json" amb l'últim dia, per comoditat
    latest_path = os.path.join(os.path.dirname(args.history) or ".", "latest.json")
    with open(latest_path, "w", encoding="utf-8") as f:
        json.dump(snapshot, f, ensure_ascii=False, separators=(",", ":"))

    cat = snapshot["catalunya"]
    print(f"OK. Sota mínims: {cat['sota_minims']}  Tancats: {cat['tancats']}  "
          f"Diferència: {cat['diferencia']}")
    print(f"Històric actualitzat: {args.history}  ({len(history)} dies)")


if __name__ == "__main__":
    main()
