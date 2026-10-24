"""
Spelarprofiler och trender från matchfilerna (flik "SkottSpelare" + "Sammanställning").

Användning:
    python spelarprofiler.py [mapp med matchfiler] [utmapp]
Utan argument (Run-knappen i VS Code) läses matcherna i mappen ovanför skriptet
och resultatet sparas i "profiler_24_25" bredvid skriptet.

Filnamnet ska sluta på datum (ddmmåå), t.ex. 'fch150425.xlsx' = 15 april 2025.

Vad som läses:
  - SkottSpelare: ett block per spelare, 4 perioder, 10 zoner, raderna
    P (på), T (täckt), U (utanför), M (mål). Siffran står till höger om bokstaven.
  - Sammanställning: xG-vikten per zon (cellen ovanför varje zons P-rad).
  - p1-p4: defensiva avslut mot oss (DEF, lagnivå) samt varje spelares övriga siffror
    (kolumn AS-AY: täck def, bollerövring, stopp, dueller, regel med, bollförlust).
    OFF-summorna i p1-p4 används också som KONTROLL av att allt stämmer.
  - Sammanställning (nederst): kontringar (√/×), högpress, återerövring (initiativ, √, ×)
    och bollinnehav. De visas i lag_oversikt.png.
Avslut = P + T + U (som i din mall). xG = avslut x zonens vikt.
DEF-zonerna finns bara för laget (inte per spelare). Spindeln jämför varje spelare mot laget:
100 % = den spelare i laget som har flest av den sortens händelse.
"""
import math
import re
import sys
import warnings
from datetime import datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")  # ritar till filer, öppnar inga fönster
import matplotlib.patheffects as pe
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.patches import Polygon
from PIL import Image, ImageDraw, ImageOps
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string as KOL
from openpyxl.utils import get_column_letter as KOLBOKSTAV

warnings.filterwarnings("ignore", category=UserWarning)  # openpyxl klagar på former i mallen

TYPER = ["P", "T", "U", "M"]  # På, Täckt, Utanför, Mål

# Kolumnen där bokstäverna (P/T/U/M) står i SkottSpelare, per period.
# Tre rader av zoner: 3 + 3 + 4 = 10 zoner. Zon 1-3 = rad 1, 4-6 = rad 2, 7-10 = rad 3.
BOKSTAVSKOLUMNER = {
    1: (["D", "H", "L"], ["D", "H", "L"], ["D", "F", "I", "L"]),
    2: (["R", "V", "Z"], ["R", "V", "Z"], ["R", "T", "W", "Z"]),
    3: (["AF", "AJ", "AN"], ["AF", "AJ", "AN"], ["AF", "AH", "AK", "AN"]),
    # I några block ligger sista zonen i period 4 i BB i stället för BC ("BC/BB" = prova båda).
    4: (["AT", "AX", "BC/BB"], ["AT", "AX", "BC/BB"], ["AT", "AV", "AY", "BC/BB"]),
}

# Planbilden (halvplan.png, 970 x 942 px, mål överst) ligger bredvid skriptet.
# Zonerna är ritade enligt zonkartan i mallen: tre band (3 + 3 + 4 zoner), där snedlinjerna går
# från målburens hörn. Alla koordinater är pixlar i planbilden (x åt höger, y nedåt).
PLANBILD = "halvplan.png"
BILD_W, BILD_H = 970, 942
_V, _H, _C = 20, 943, 481.5          # vänster vägg, höger vägg, mitten
_G, _K, _T, _M = 178, 332, 595, 940  # mållinje, krease-kant, tredje linjen, mittlinje
_xa, _xb = 188, 750                  # lodräta linjer i nedersta bandet
_xk1, _xk2 = 336, 618                # snedlinjerna där de möter krease-kanten
ZONPOLYGONER = {
    1: [(_V, _G), (423, _G), (_xk1, _K), (_V, _K)],
    2: [(423, _G), (540, _G), (_xk2, _K), (_xk1, _K)],
    3: [(540, _G), (_H, _G), (_H, _K), (_xk2, _K)],
    4: [(_V, _K), (_xk1, _K), (_xa, _T), (_V, _T)],
    5: [(_xk1, _K), (_xk2, _K), (_xb, _T), (_xa, _T)],
    6: [(_xk2, _K), (_H, _K), (_H, _T), (_xb, _T)],
    7: [(_V, _T), (_xa, _T), (_xa, _M), (_V, _M)],
    8: [(_xa, _T), (_C, _T), (_C, _M), (_xa, _M)],
    9: [(_C, _T), (_xb, _T), (_xb, _M), (_C, _M)],
    10: [(_xb, _T), (_H, _T), (_H, _M), (_xb, _M)],
}
# Namn på zonerna (används i texten bredvid kartan). Vänster/höger följer planbilden.
ZONNAMN = {1: "nära mål, vänster", 2: "nära mål, mitt (slot)", 3: "nära mål, höger",
           4: "mellanzon, vänster", 5: "mellanzon, mitt", 6: "mellanzon, höger",
           7: "bortre, ytter vänster", 8: "bortre, mitt vänster", 9: "bortre, mitt höger",
           10: "bortre, ytter höger"}
# Var texten (antal mål) placeras i varje zon.
ZONTEXT = {1: (190, 258), 2: (481, 284), 3: (775, 258), 4: (125, 465), 5: (481, 465), 6: (838, 465),
           7: (104, 767), 8: (335, 767), 9: (615, 767), 10: (846, 767)}
FARGER = {"P": "#c0392b", "T": "#7f8c8d", "U": "#d5d8dc"}

# Spelarnas övriga siffror: kolumn AS-AY i p1-p4 (en rad per spelare, spelarnumret står i AL).
HANDELSEKOLUMNER = {"tack_def": "AS", "boll_eroverring": "AT", "stopp": "AU", "off_duell": "AV",
                    "def_duell": "AW", "regel_med": "AX", "bollforlust": "AY"}
# Lagstatistik som står i Sammanställning (läses som cachade värden) och var samma siffror skrivs in i
# period-bladen: nyckel -> (cell i Sammanställning, celler som summeras i varje p-blad).
# ⫻ = antal initiativ, √ = lyckade, × = misslyckade. Bollinnehav är minuter (Sammanställning summerar p1-p3).
LAGSTAT = {
    "kontring_ok": ("A29", ["A29"]),
    "kontring_miss": ("C29", ["C29"]),
    "hogpress_initiativ": ("T24", ["T24", "U24"]),
    "hogpress_ok": ("V24", ["V24", "W24"]),
    "hogpress_miss": ("X24", ["X24", "Y24"]),
    "aterero_initiativ": ("T31", ["T31", "U31"]),
    "aterero_ok": ("V31", ["V31", "W31"]),
    "aterero_miss": ("X31", ["X31", "Y31"]),
    "bollinnehav_min": ("N33", ["N31"]),
}
GRON, ROD, GRA = "#2e8b57", "#c0392b", "#7f8c8d"

SPINDEL = [("tack_def", "Täck def"), ("boll_eroverring", "Boll-\nerövring"), ("stopp", "Stopp"),
           ("off_duell", "Off duell"), ("def_duell", "Def duell"), ("regel_med", "Regel med"),
           ("bollforlust", "Boll-\nförlust")]


# ---------- Filnamn ----------
def tolka_filnamn(fil: Path):
    """'fch150425' -> (datetime(2025, 4, 15), '15/04 fch'). None om datum saknas."""
    m = re.search(r"(\d{6})", fil.stem)
    if not m:
        return None, None
    try:
        datum = datetime.strptime(m.group(1), "%d%m%y")
    except ValueError:
        return None, None
    namn = (fil.stem[:m.start()] + fil.stem[m.end():]).strip(" _-")
    return datum, f"{datum:%d/%m} {namn}".strip()


# ---------- Inläsning av en match ----------
def tal(v, var, plats):
    """Gör om en cell till ett tal. Tom cell = 0. Konstigt innehåll = 0 + varning."""
    if v is None or v == "":
        return 0
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        var.add((plats.split(" | ")[0], f"cellen {plats.split(' | ')[1]} innehåller '{v}' (inte ett tal) – räknas som 0."))
        return 0


def las_vikter(ws, kol_fran=1, kol_till=17):
    """xG-vikt per zon: står i cellen ovanför varje zons P-rad i Sammanställning.
    OFF-sidan ligger i kolumn 1-16 och DEF-sidan i kolumn 17-39."""
    vikter, zon = {}, 1
    for p_rad in (4, 9, 14):
        for c in range(kol_fran, kol_till):
            if ws.cell(p_rad, c).value == "P":
                vikter[zon] = float(ws.cell(p_rad - 1, c).value)
                zon += 1
    if len(vikter) != 10:
        raise ValueError(f"hittade {len(vikter)} zoner med xG-vikt i Sammanställning, kolumn "
                         f"{kol_fran}-{kol_till - 1} (väntade 10)")
    return vikter


def las_skottspelare(ws, etikett, var):
    rader, block = [], {}
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 3).value
        if isinstance(v, str) and v.strip().startswith("#") and not re.fullmatch(r"#\d+", v.strip()):
            # Block utan spelarnummer (t.ex. "#??"): ignoreras, men varna om det finns siffror i det.
            siffror = [c for rr in range(r, r + 14) for c in range(4, 58)
                       if isinstance(ws.cell(rr, c).value, (int, float)) and ws.cell(rr, c).value > 0]
            if siffror:
                var.add((etikett, f"Blocket '{v.strip()}' på rad {r} i SkottSpelare har siffror men inget spelarnummer "
                                  "– det ignoreras. Skriv in ett nummer i blockets rubrik."))
            continue
        if not (isinstance(v, str) and re.fullmatch(r"#\d+", v.strip())):
            continue
        nr = int(v.strip()[1:])
        block.setdefault(nr, []).append(r)
        for per, grupper in BOKSTAVSKOLUMNER.items():
            zon = 1
            for g, kolumner in enumerate(grupper):
                rad0 = r + 2 + 4 * g
                for k in kolumner:
                    # prova alla kolumner som står som alternativ ("BC/BB") och ta den där bokstäverna finns
                    col = None
                    for alt in k.split("/"):
                        c_try = KOL(alt)
                        if all(ws.cell(rad0 + j, c_try).value == TYPER[j] for j in range(4)):
                            col = c_try
                            break
                    if col is None:
                        var.add((etikett, f"#{nr} saknar zonbokstäverna P/T/U/M i {k.split('/')[0]}{rad0} "
                                          f"(period {per}, zon {zon}) – zonen räknas som 0."))
                    else:
                        for j, typ in enumerate(TYPER):
                            cell = ws.cell(rad0 + j, col + 1)
                            antal = tal(cell.value, var, f"{etikett} | {cell.coordinate}")
                            if antal > 0:
                                rader.append((etikett, nr, per, zon, typ, antal))
                    zon += 1
    for nr, rr in block.items():
        if len(rr) > 1:
            var.add((etikett, f"#{nr} förekommer {len(rr)} gånger i SkottSpelare "
                              f"(rad {', '.join(map(str, rr))}) – siffrorna slås ihop."))
    return pd.DataFrame(rader, columns=["match", "spelare", "period", "zon", "typ", "antal"])


def kontrollera_mot_periodblad(wb, df, etikett, var):
    """Jämför våra summor med p1-p4 (OFF-sidan). Avvikelse = en formel i mallen missar en cell."""
    for per in range(1, 5):
        namn = f"p{per}"
        if namn not in wb.sheetnames:
            continue
        ws = wb[namn]
        blad = {t: 0 for t in TYPER}
        for r in range(1, 20):
            for c in range(1, 17):
                v = ws.cell(r, c).value
                if v in TYPER:
                    x = ws.cell(r, c + 1).value
                    blad[v] += x if isinstance(x, (int, float)) else 0
        mina = df[df.period == per].groupby("typ")["antal"].sum()
        for t in TYPER:
            m = int(mina.get(t, 0))
            if m != blad[t]:
                var.add((etikett, f"{namn} visar {blad[t]} ({t}) men SkottSpelare summerar till {m}. "
                                  f"En formel i {namn} missar troligen en cell. Skriptet räknar från SkottSpelare."))


def las_def_zoner(wb, etikett, var):
    """Motståndarens avslut mot oss (DEF). Lagnivå, inskrivet för hand i p1-p4. Samma 10 zoner som OFF."""
    rader = []
    for per in range(1, 5):
        namn = f"p{per}"
        if namn not in wb.sheetnames:
            continue
        ws = wb[namn]
        zon = 1
        for p_rad in (4, 9, 14):
            for c in range(17, 40):
                if ws.cell(p_rad, c).value != "P":
                    continue
                for j, typ in enumerate(TYPER):
                    if ws.cell(p_rad + j, c).value != typ:
                        var.add((etikett, f"{namn}: DEF-zon {zon} saknar bokstaven {typ} i "
                                          f"{KOLBOKSTAV(c)}{p_rad + j}."))
                        continue
                    cell = ws.cell(p_rad + j, c + 1)
                    antal = tal(cell.value, var, f"{etikett} | {namn}!{cell.coordinate}")
                    if antal > 0:
                        rader.append((etikett, per, zon, typ, antal))
                zon += 1
        if zon != 11:
            var.add((etikett, f"{namn}: hittade {zon - 1} DEF-zoner (väntade 10)."))
    return pd.DataFrame(rader, columns=["match", "period", "zon", "typ", "antal"])


def _nr(v):
    return f"#{v}" if v else "tom"


def las_handelser(wb, etikett, var):
    """Spelarnas övriga siffror (p1-p4, kolumn AS-AY), summerade över perioderna, per spelarnummer."""
    rader, listor = [], {}
    for per in range(1, 5):
        namn = f"p{per}"
        if namn not in wb.sheetnames:
            continue
        ws = wb[namn]
        lista = {}
        for r in range(2, 61):
            nr = ws.cell(r, KOL("AL")).value
            if not isinstance(nr, (int, float)):
                continue
            lista[r] = int(nr)
            vard = {k: tal(ws.cell(r, KOL(kol)).value, var, f"{etikett} | {namn}!{kol}{r}")
                    for k, kol in HANDELSEKOLUMNER.items()}
            rader.append({"match": etikett, "spelare": int(nr), **vard})
        listor[per] = lista
    # Kontroll: ligger spelarna i samma ordning i alla perioder? Annars blir 3D-summorna i mallen fel.
    for per, lista in listor.items():
        if per == 1 or 1 not in listor:
            continue
        skillnad = [r for r in sorted(set(lista) | set(listor[1])) if lista.get(r) != listor[1].get(r)]
        if skillnad:
            r = skillnad[0]
            var.add((etikett, f"spelarlistan (kolumn AL) skiljer sig mellan p1 och p{per} på {len(skillnad)} rad(er), "
                              f"t.ex. rad {r}: {_nr(listor[1].get(r))} i p1 men {_nr(lista.get(r))} i p{per}. "
                              "Summorna i Sammanställning kan då blanda ihop spelare. Skriptet räknar per spelarnummer."))
    kolumner = list(HANDELSEKOLUMNER)
    if not rader:
        return pd.DataFrame(columns=["match", "spelare"] + kolumner)
    return pd.DataFrame(rader).groupby(["match", "spelare"], as_index=False)[kolumner].sum()


def kontrollera_handelser(wb, h, etikett, var):
    """Jämför med spelartabellen i Sammanställning (kolumn AS-AY). En varning per spelarrad."""
    ws = wb["Sammanställning"]
    mina = h.set_index("spelare")
    for r in range(2, 61):
        nr = ws.cell(r, KOL("AL")).value
        if not isinstance(nr, (int, float)):
            continue
        nr = int(nr)
        avvik = []
        for k, kol in HANDELSEKOLUMNER.items():
            blad = tal(ws.cell(r, KOL(kol)).value, var, f"{etikett} | Sammanställning!{kol}{r}")
            m = mina[k].get(nr, 0) if nr in mina.index else 0
            if blad != m:
                avvik.append(f"{k} {blad:g} (p1-p4: {m:g})")
        if avvik:
            var.add((etikett, f"Sammanställning rad {r} (#{nr}): {', '.join(avvik)}. "
                              "Kolla formlerna på raden – den verkar räkna på fel rad i p1-p4."))


def las_lagstat(wb, etikett, var):
    """Kontringar, högpress, återerövring och bollinnehav från Sammanställning – kontrollerade mot p1-p4."""
    sam = wb["Sammanställning"]
    ut = {"match": etikett}
    for k, (cell, raa) in LAGSTAT.items():
        v = tal(sam[cell].value, var, f"{etikett} | Sammanställning!{cell}")
        ut[k] = v
        perioder = ("p1", "p2", "p3") if k == "bollinnehav_min" else ("p1", "p2", "p3", "p4")
        summa = 0
        for pn in perioder:
            if pn in wb.sheetnames:
                for c in raa:
                    x = wb[pn][c].value
                    summa += x if isinstance(x, (int, float)) else 0
        if abs(v - summa) > 1e-6:
            var.add((etikett, f"Sammanställning!{cell} visar {v:g} men {'+'.join(perioder)} summerar till {summa:g}. "
                              "Spara filen i Excel så uppdateras värdena, eller kolla formeln."))
    if "p4" in wb.sheetnames:
        p4 = wb["p4"]["N31"].value
        if isinstance(p4, (int, float)) and p4 > 0:
            var.add((etikett, "p4 har bollinnehav (N31) men Sammanställning summerar bara p1-p3 (60 min). "
                              "Övertiden räknas inte med."))
    ut["lag_namn"] = sam["K36"].value or "Vi"
    ut["mot_namn"] = sam["M36"].value or "Motståndare"
    return ut


def las_match(fil: Path, datum, etikett, var):
    wb = load_workbook(fil, data_only=True)
    for namn in ("SkottSpelare", "Sammanställning"):
        if namn not in wb.sheetnames:
            raise ValueError(f"saknar bladet '{namn}'. Bladen heter: {wb.sheetnames}")
    sam = wb["Sammanställning"]
    vikter_off = las_vikter(sam, 1, 17)
    vikter_def = las_vikter(sam, 17, 40)
    df = las_skottspelare(wb["SkottSpelare"], etikett, var)
    kontrollera_mot_periodblad(wb, df, etikett, var)
    df["datum"] = datum
    df["vikt"] = df["zon"].map(vikter_off)
    d = las_def_zoner(wb, etikett, var)
    d["datum"] = datum
    d["vikt"] = d["zon"].map(vikter_def)
    h = las_handelser(wb, etikett, var)
    kontrollera_handelser(wb, h, etikett, var)
    h["datum"] = datum
    lag = las_lagstat(wb, etikett, var)
    lag["datum"] = datum
    return df, h, d, lag


def las_alla(sokvag: Path, var):
    filer = [sokvag] if sokvag.is_file() else sorted(
        p for p in sokvag.iterdir()
        if p.suffix.lower() in (".xlsm", ".xlsx") and not p.name.startswith("~$")
    )
    delar = []
    for f in filer:
        datum, etikett = tolka_filnamn(f)
        if datum is None:
            print(f"Hoppar över {f.name}: hittar inget datum (ddmmåå) i filnamnet.")
            continue
        try:
            delar.append(las_match(f, datum, etikett, var))
        except Exception as fel:  # visar det RIKTIGA felet, inte en gissning
            print(f"Hoppar över {f.name}: {type(fel).__name__}: {fel}")
    if not delar:
        sys.exit("Hittade inga användbara matchfiler.")
    print(f"Läste in {len(delar)} matcher.")
    df = pd.concat([x[0] for x in delar], ignore_index=True)
    handelser = pd.concat([x[1] for x in delar], ignore_index=True)
    defz = pd.concat([x[2] for x in delar], ignore_index=True)
    lagstat = pd.DataFrame([x[3] for x in delar])
    return df, handelser, defz, lagstat


# ---------- Beräkning ----------
def bygg_tabeller(df):
    """Zonrader per match och spelare med P, T, U, M, avslut och xG."""
    bred = (df.pivot_table(index=["datum", "match", "spelare", "zon", "vikt"],
                           columns="typ", values="antal", aggfunc="sum", fill_value=0)
              .reindex(columns=TYPER, fill_value=0).reset_index())
    bred["avslut"] = bred["P"] + bred["T"] + bred["U"]
    bred["xg"] = bred["avslut"] * bred["vikt"]
    return bred


def bygg_def(d):
    """Motståndarens avslut mot oss per match och zon (P, T, U, M, avslut, xG)."""
    if d.empty:
        return pd.DataFrame(columns=["datum", "match", "zon", "vikt"] + TYPER + ["avslut", "xg"])
    bred = (d.pivot_table(index=["datum", "match", "zon", "vikt"], columns="typ", values="antal",
                          aggfunc="sum", fill_value=0)
              .reindex(columns=TYPER, fill_value=0).reset_index())
    bred["avslut"] = bred["P"] + bred["T"] + bred["U"]
    bred["xg"] = bred["avslut"] * bred["vikt"]
    return bred


def lag_summering(lagstat):
    """Totaler för hela serien och procent. Bollinnehav räknas på matcher där det är ifyllt (60 min per match)."""
    t = {k: float(lagstat[k].sum()) for k in LAGSTAT}
    t["matcher"] = len(lagstat)
    mm = lagstat[lagstat["bollinnehav_min"] > 0]
    t["bollinnehav_procent"] = 100 * mm["bollinnehav_min"].sum() / (60 * len(mm)) if len(mm) else None
    for namn in ("kontring", "hogpress", "aterero"):
        ok, miss = t[f"{namn}_ok"], t[f"{namn}_miss"]
        t[f"{namn}_procent_lyckade"] = 100 * ok / (ok + miss) if ok + miss else None
    t["lag_namn"] = lagstat["lag_namn"].iloc[0]
    t["mot_namn"] = lagstat["mot_namn"].iloc[0]
    return t


def summera(bred, extra=()):
    rader = []
    for nr, e in bred.groupby("spelare"):
        t = e[["P", "T", "U", "M", "avslut", "xg"]].sum()
        if t["avslut"] == 0 and t["M"] == 0:
            continue
        zoner = e.groupby("zon")["avslut"].sum()
        topp = zoner.idxmax()
        rader.append({
            "spelare": int(nr),
            "avslut": int(t["avslut"]),
            "pa_mal": int(t["P"]), "tackta": int(t["T"]), "utanfor": int(t["U"]), "mal": int(t["M"]),
            "pa_procent": round(100 * t["P"] / t["avslut"]) if t["avslut"] else 0,
            "mal_procent": round(100 * t["M"] / t["avslut"]) if t["avslut"] else 0,
            "xg": round(t["xg"], 2),
            "mal_minus_xg": round(t["M"] - t["xg"], 2),
            "matcher_med_avslut": int(e[e.avslut > 0]["match"].nunique()),
            "vanligaste_zon": f"Z{int(topp)}",
            "andel_i_vanligaste_zon_procent": round(100 * zoner.max() / t["avslut"]) if t["avslut"] else 0,
        })
    # Spelare som har defensiva siffror men inga avslut tas också med (tom skottkarta).
    finns = {r["spelare"] for r in rader}
    for nr in extra:
        if int(nr) not in finns:
            rader.append({"spelare": int(nr), "avslut": 0, "pa_mal": 0, "tackta": 0, "utanfor": 0, "mal": 0,
                          "pa_procent": 0, "mal_procent": 0, "xg": 0.0, "mal_minus_xg": 0.0,
                          "matcher_med_avslut": 0, "vanligaste_zon": "–",
                          "andel_i_vanligaste_zon_procent": 0})
    kolumner = ["spelare", "avslut", "pa_mal", "tackta", "utanfor", "mal", "pa_procent", "mal_procent", "xg",
                "mal_minus_xg", "matcher_med_avslut", "vanligaste_zon", "andel_i_vanligaste_zon_procent"]
    if not rader:  # t.ex. en helt tom mall: inga spelare med avslut
        return pd.DataFrame(columns=kolumner)
    return pd.DataFrame(rader).sort_values(["avslut", "xg", "spelare"], ascending=[False, False, True])


# ---------- Ritning ----------
def ladda_plan(skriptmapp):
    """Läser planbilden (och rättar rotation från skärmdumpar). Saknas filen ritas en enkel plan."""
    fil = skriptmapp / PLANBILD
    if fil.exists():
        bild = ImageOps.exif_transpose(Image.open(fil)).convert("RGB")
        if bild.size != (BILD_W, BILD_H):
            print(f"OBS: {PLANBILD} är {bild.size[0]}x{bild.size[1]} px (väntade {BILD_W}x{BILD_H}). "
                  "Bilden sträcks, så zonerna kan hamna lite fel.")
            bild = bild.resize((BILD_W, BILD_H))
        return np.array(bild)
    print(f"OBS: hittar inte {PLANBILD} bredvid skriptet – ritar en enkel plan i stället.")
    bild = Image.new("RGB", (BILD_W, BILD_H), (122, 176, 220))
    d = ImageDraw.Draw(bild)
    d.rectangle([_V, 16, _H, _M], outline="white", width=2)
    d.rectangle([366, 147, 597, 332], fill=(11, 91, 147), outline="white", width=2)
    d.rectangle([423, 178, 540, 224], fill=(122, 176, 220), outline="white", width=2)
    return np.array(bild)


def rita_zonkarta(ax, avslut, mal, xg, titel, plan):
    """Planbild med zoner. Varje zon visar mål, avslut och xG. Färgen visar antal avslut."""
    ax.imshow(plan)
    maxv = max(float(avslut.max()), 1.0) if len(avslut) else 1.0
    farg = plt.get_cmap("YlOrRd")
    tjock = [pe.withStroke(linewidth=3, foreground="black")]
    tunn = [pe.withStroke(linewidth=2, foreground="black")]
    for zon, hornen in ZONPOLYGONER.items():
        n, g, v = int(avslut.get(zon, 0)), int(mal.get(zon, 0)), float(xg.get(zon, 0))
        ax.add_patch(Polygon(hornen, closed=True, edgecolor="white", linewidth=1.6,
                             facecolor=farg(0.15 + 0.85 * n / maxv) if n else "none",
                             alpha=0.8 if n else 1.0))
        if n == 0:
            continue
        x, y = ZONTEXT[zon]
        ax.text(x, y - 28, f"{g} mål", ha="center", va="center", color="white", fontsize=15,
                fontweight="bold", alpha=1.0 if g else 0.8, path_effects=tjock)
        ax.text(x, y + 4, f"{n} avslut", ha="center", va="center", color="white", fontsize=10,
                path_effects=tunn)
        ax.text(x, y + 28, f"{v:.2f} xG", ha="center", va="center", color="white", fontsize=10,
                path_effects=tunn)
    ax.set_xlim(0, BILD_W)
    ax.set_ylim(BILD_H, 0)
    ax.axis("off")
    ax.set_title(titel, fontsize=11)


def rita_matchtrend(ax, pm, matcher, titel):
    """Staplar: på / täckt / utanför per match. Linje: xG. Stjärna: match med mål."""
    pm = pm.reindex(matcher).fillna(0)
    x = range(len(matcher))
    botten = pd.Series(0.0, index=pm.index)
    for typ, namn in (("P", "På mål"), ("T", "Täckt"), ("U", "Utanför")):
        ax.bar(x, pm[typ], bottom=botten, color=FARGER[typ], label=namn)
        botten += pm[typ]
    ax.plot(x, pm["xg"], color="black", marker="o", markersize=3, linewidth=1, label="xG")
    mal = [i for i, v in enumerate(pm["M"]) if v > 0]
    if mal:
        ax.scatter(mal, [botten.iloc[i] + 0.4 for i in mal], marker="*", s=90, color="gold",
                   edgecolor="black", label="mål", zorder=3)
    ax.set_xticks(list(x))
    ax.set_xticklabels(matcher, rotation=90, fontsize=7)
    ax.set_ylabel("avslut")
    ax.set_title(titel, fontsize=11)
    ax.legend(fontsize=7, loc="upper left")


def rita_spindel(ax, vardena, maxar, titel, snitt=None, n_snitt=0, spelarnamn="Spelaren", liten=False):
    """Spindeldiagram. Varje axel: spelarens värde / högsta värdet i laget (100 % = flest i laget).
    Den grå ytan är snittet för spelare som har statistik (n_snitt st), på samma skala."""
    n = len(SPINDEL)
    vinklar = [2 * math.pi * i / n for i in range(n)]
    rel = [100 * vardena[k] / maxar[k] if maxar[k] else 0 for k, _ in SPINDEL]
    ax.set_theta_offset(math.pi / 2)
    ax.set_theta_direction(-1)
    if snitt is not None and n_snitt:
        rel_snitt = [100 * snitt[k] / maxar[k] if maxar[k] else 0 for k, _ in SPINDEL]
        ax.fill(vinklar + vinklar[:1], rel_snitt + rel_snitt[:1], color=GRA, alpha=0.5, zorder=1)
        ax.plot(vinklar + vinklar[:1], rel_snitt + rel_snitt[:1], color="#566573", linewidth=1.4,
                linestyle="--", label=f"Snitt ({n_snitt} spelare)", zorder=2)
    ax.fill(vinklar + vinklar[:1], rel + rel[:1], color=ROD, alpha=0.3, zorder=3)
    ax.plot(vinklar + vinklar[:1], rel + rel[:1], color=ROD, linewidth=2, label=spelarnamn, zorder=4)
    etiketter = []
    for (k, namn), r in zip(SPINDEL, rel):
        rad = f"{namn}\n{int(vardena[k])} st ({r:.0f} %)"
        if snitt is not None and n_snitt and not liten:
            rad += f"\nsnitt {snitt[k]:.1f}"
        etiketter.append(rad)
    ax.set_xticks(vinklar)
    ax.set_xticklabels(etiketter, fontsize=8 if liten else 9)
    ax.tick_params(axis="x", pad=10 if liten else 18)
    ax.set_ylim(0, 100)
    ax.set_yticks([25, 50, 75, 100])
    ax.set_yticklabels(["25 %", "50 %", "75 %", "100 %"], fontsize=6 if liten else 7, color="gray")
    ax.set_rlabel_position(180 / n)
    ax.set_title(titel, fontsize=10 if liten else 11, pad=22 if liten else 36)
    if liten:  # liten spindel (lagöversikten): legenden under diagrammet
        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.14), ncol=2, fontsize=8, frameon=False)
    else:
        ax.legend(loc="upper right", bbox_to_anchor=(1.3, 1.18), fontsize=8, frameon=False)


def rita_spelare(bred, rad, matcher, utmapp, plan, h_tot, h_max, snitt, n_snitt):
    nr = rad["spelare"]
    e = bred[bred.spelare == nr]
    fig = plt.figure(figsize=(18, 12))
    gs = fig.add_gridspec(2, 3, width_ratios=[1.6, 1.0, 1.0], height_ratios=[1.5, 1])
    ax_karta = fig.add_subplot(gs[0, 0])
    ax_text = fig.add_subplot(gs[1, 0])
    ax_trend = fig.add_subplot(gs[0, 1:])
    ax_spindel = fig.add_subplot(gs[1, 1:], polar=True)

    z = e.groupby("zon")[["avslut", "M", "xg"]].sum()
    rita_zonkarta(ax_karta, z["avslut"], z["M"], z["xg"],
                  f"#{nr} – mål, avslut och xG per zon (hela serien)", plan)
    inga_avslut = rad["avslut"] == 0 and rad["mal"] == 0
    if inga_avslut:
        ax_trend.axis("off")
        ax_trend.text(0.5, 0.5, "Inga avslut registrerade", ha="center", va="center",
                      fontsize=14, color="gray")
    else:
        pm = e.groupby("match")[["P", "T", "U", "M", "xg"]].sum()
        rita_matchtrend(ax_trend, pm, matcher, "Avslut och xG per match")

    ax_text.axis("off")
    zonnamn = ZONNAMN.get(int(rad["vanligaste_zon"][1:]), "–") if str(rad["vanligaste_zon"]).startswith("Z") else "–"
    text = (f"Spelare #{nr}\n\nInga avslut registrerade." if inga_avslut else
            f"Spelare #{nr}\n\n"
            f"Avslut: {rad['avslut']}\n"
            f"På mål: {rad['pa_mal']} ({rad['pa_procent']} %)\n"
            f"Täckta: {rad['tackta']}\nUtanför: {rad['utanfor']}\n"
            f"Mål: {rad['mal']} ({rad['mal_procent']} % av avslut)\n\n"
            f"xG: {rad['xg']:.2f}\nMål – xG: {rad['mal_minus_xg']:+.2f}\n\n"
            f"Matcher med avslut: {rad['matcher_med_avslut']}\n"
            f"Flest avslut:\n{zonnamn}\n"
            f"({rad['andel_i_vanligaste_zon_procent']} % av avsluten)")
    ax_text.text(0.05, 0.95, text, va="top", fontsize=11, family="monospace")

    vardena = h_tot.loc[nr] if nr in h_tot.index else pd.Series(0, index=h_max.index)
    rita_spindel(ax_spindel, vardena, h_max,
                 "Defensiv- och duellprofil – 100 % = flest i laget (hela serien)",
                 snitt=snitt, n_snitt=n_snitt, spelarnamn=f"#{nr}")
    fig.tight_layout()
    fig.savefig(utmapp / f"spelare_{nr}.png", dpi=110)
    plt.close(fig)


def rita_lyckade(ax, titel, ok, miss, undertitel=""):
    """Två staplar: lyckade (✓) och misslyckade (✗)."""
    ax.set_title(titel, fontsize=12, pad=28)
    if ok + miss == 0:
        ax.axis("off")
        ax.text(0.5, 0.5, "Ingen data i\nSammanställning", ha="center", va="center",
                color="gray", fontsize=11, transform=ax.transAxes)
        return
    staplar = ax.bar(["✓ lyckade", "✗ misslyckade"], [ok, miss], color=[GRON, ROD])
    for st, v in zip(staplar, (ok, miss)):
        ax.text(st.get_x() + st.get_width() / 2, v, f"{int(v)}", ha="center", va="bottom",
                fontsize=14, fontweight="bold")
    ax.set_ylim(0, max(ok, miss) * 1.3)
    ax.text(0.5, 1.03, undertitel, transform=ax.transAxes, ha="center", va="bottom", fontsize=10, color="dimgray")
    ax.spines[["top", "right"]].set_visible(False)
    ax.tick_params(axis="x", labelsize=11)


def rita_bollinnehav(ax, procent, lag, mot):
    ax.set_title("Bollinnehav (hela serien)", fontsize=12, pad=28)
    ax.axis("off")
    if procent is None:
        ax.text(0.5, 0.5, "Ingen data i\nSammanställning", ha="center", va="center",
                color="gray", fontsize=11, transform=ax.transAxes)
        return
    ax.barh([0], [procent], color=GRON, height=0.55)
    ax.barh([0], [100 - procent], left=[procent], color="#b0b7bb", height=0.55)
    ax.text(procent / 2, 0, f"{lag}\n{procent:.0f} %", ha="center", va="center", color="white",
            fontsize=13, fontweight="bold")
    ax.text(procent + (100 - procent) / 2, 0, f"{mot}\n{100 - procent:.0f} %", ha="center", va="center",
            color="#2c3e50", fontsize=13, fontweight="bold")
    ax.set_xlim(0, 100)
    ax.set_ylim(-0.6, 0.6)


def _procent_text(v):
    return f"{v:.0f} % lyckade" if v is not None else ""


def rita_lag(bred, bred_def, sammanfattning, matcher, utmapp, plan, lag_tot, h_tot, h_max, snitt, n_snitt, topp3):
    fig = plt.figure(figsize=(17, 21))
    gs = fig.add_gridspec(4, 1, height_ratios=[1.7, 0.7, 1.0, 0.9], hspace=0.42)

    # 1) Zonkartor: offensiv och defensiv planhalva
    g0 = gs[0].subgridspec(1, 2, wspace=0.05)
    ax_off, ax_def = fig.add_subplot(g0[0, 0]), fig.add_subplot(g0[0, 1])
    z = bred.groupby("zon")[["avslut", "M", "xg"]].sum()
    rita_zonkarta(ax_off, z["avslut"], z["M"], z["xg"],
                  "OFFENSIV planhalva – våra avslut (hela serien)\n(mål / avslut / xG)", plan)
    zd = bred_def.groupby("zon")[["avslut", "M", "xg"]].sum()
    rita_zonkarta(ax_def, zd["avslut"], zd["M"], zd["xg"],
                  "DEFENSIV planhalva – motståndarens avslut mot oss (hela serien)\n"
                  "(mål mot / avslut mot / xG mot)", plan)

    # 2) Lagstatistik från Sammanställning
    g1 = gs[1].subgridspec(1, 4, wspace=0.4)
    ax_k, ax_b, ax_h, ax_a = (fig.add_subplot(g1[0, i]) for i in range(4))
    k_ok, k_miss = lag_tot["kontring_ok"], lag_tot["kontring_miss"]
    rita_lyckade(ax_k, "Kontringar", k_ok, k_miss,
                 f"{_procent_text(lag_tot['kontring_procent_lyckade'])} ({int(k_ok)} av {int(k_ok + k_miss)})")
    rita_bollinnehav(ax_b, lag_tot["bollinnehav_procent"], lag_tot["lag_namn"], lag_tot["mot_namn"])
    for ax, namn, nyckel in ((ax_h, "Högpress", "hogpress"), (ax_a, "Återerövring", "aterero")):
        ok, miss, init = lag_tot[f"{nyckel}_ok"], lag_tot[f"{nyckel}_miss"], lag_tot[f"{nyckel}_initiativ"]
        delar = [f"{int(init)} initiativ"] if init else []
        if _procent_text(lag_tot[f"{nyckel}_procent_lyckade"]):
            delar.append(_procent_text(lag_tot[f"{nyckel}_procent_lyckade"]))
        rita_lyckade(ax, namn, ok, miss, " · ".join(delar))

    # 3) Topp 3 spindlar
    ax_rad = fig.add_subplot(gs[2])
    ax_rad.axis("off")
    g2 = gs[2].subgridspec(1, 3, wspace=0.55)
    for i, nr in enumerate(topp3):
        ax = fig.add_subplot(g2[0, i], polar=True)
        rita_spindel(ax, h_tot.loc[nr], h_max, f"Topp {i + 1}: #{nr}", snitt=snitt, n_snitt=n_snitt,
                     spelarnamn=f"#{nr}", liten=True)
    if topp3:
        ax_rad.text(0.5, -0.24, "Topp 3 = högst summa av procent på de sex axlarna (bollförlust räknas inte med). "
                                "Grå yta = snitt för spelare med statistik.",
                    ha="center", va="top", fontsize=9, color="dimgray", transform=ax_rad.transAxes)

    # 4) Trend och avslut per spelare
    g3 = gs[3].subgridspec(1, 2, wspace=0.25)
    ax_trend, ax_bar = fig.add_subplot(g3[0, 0]), fig.add_subplot(g3[0, 1])
    pm = bred.groupby("match")[["P", "T", "U", "M", "xg"]].sum()
    rita_matchtrend(ax_trend, pm, matcher, "Lagets avslut och xG per match")
    s = sammanfattning[sammanfattning.avslut > 0].head(15).iloc[::-1]
    botten = pd.Series(0, index=s.index)
    for kol, typ in (("pa_mal", "P"), ("tackta", "T"), ("utanfor", "U")):
        ax_bar.barh(s["spelare"].astype(str), s[kol], left=botten, color=FARGER[typ])
        botten += s[kol]
    ax_bar.set_title("Avslut per spelare (topp 15)", fontsize=11)
    ax_bar.set_xlabel("avslut")

    antal = int(lag_tot["matcher"])
    fig.suptitle(f"Lagöversikt – {antal} {'match' if antal == 1 else 'matcher'}", fontsize=16, y=0.985)
    fig.subplots_adjust(left=0.05, right=0.97, top=0.955, bottom=0.04)
    fig.savefig(utmapp / "lag_oversikt.png", dpi=100)
    plt.close(fig)


def spara_csv(df, sokvag):
    df.to_csv(sokvag, index=False, sep=";", decimal=",", encoding="utf-8-sig")  # öppnas rätt i svensk Excel
def skapa_portal(utmapp, sammanfattning):
     print("SKAPAR PORTAL")
    html = """
    <html>
    <head>
        <title>Väsby Analytics</title>
    </head>
    <body>
        <h1>Väsby Analytics</h1>

        <h2>Lagöversikt</h2>
        <img src="lag_oversikt.png" width="1000">

        <h2>Spelare</h2>
    """

    for _, rad in sammanfattning.iterrows():
        nr = int(rad["spelare"])

        html += f"""
        <p>
            <a href="spelare_{nr}.html">
                Spelare #{nr}
            </a>
        </p>
        """

    html += "</body></html>"

    with open(utmapp / "index.html", "w", encoding="utf-8") as f:
        f.write(html)
def skapa_spelarsidor(utmapp, sammanfattning):
    print("SKAPAR SPELARSIDOR")
    for _, rad in sammanfattning.iterrows():

        nr = int(rad["spelare"])

        html = f"""
        <html>
        <body>

        <a href="index.html">
        ← Tillbaka
        </a>

        <h1>Spelare #{nr}</h1>

        <img src="spelare_{nr}.png" width="1400">

        </body>
        </html>
        """

        with open(
            utmapp / f"spelare_{nr}.html",
            "w",
            encoding="utf-8"
        ) as f:
            f.write(html)


def main():
    skriptmapp = Path(__file__).resolve().parent
    # Statistikbladen ligger i Serie/, tre nivåer ovanför skriptet i files/
    sokvag = Path(sys.argv[1]) if len(sys.argv) > 1 else skriptmapp.parent.parent
    utmapp = Path(sys.argv[2]) if len(sys.argv) > 2 else skriptmapp / "profiler_24_25"
    utmapp.mkdir(parents=True, exist_ok=True)

    var = set()  # varningar samlas och skrivs ut sist
    df, handelser, defz, lagstat = las_alla(sokvag, var)
    bred = bygg_tabeller(df)
    bred_def = bygg_def(defz)
    matcher = list(bred.sort_values("datum", kind="stable")["match"].drop_duplicates())

    kolumner = list(HANDELSEKOLUMNER)
    h_tot = handelser.groupby("spelare")[kolumner].sum() if not handelser.empty else pd.DataFrame(columns=kolumner)
    h_max = h_tot.max() if len(h_tot) else pd.Series(0, index=kolumner)
    extra = h_tot.index[h_tot.sum(axis=1) > 0]
    sammanfattning = summera(bred, extra)

    # Snitt för spelare som har statistik (grå spindel) och topp 3 till lagöversikten
    aktiva = h_tot[h_tot.sum(axis=1) > 0]
    n_snitt = len(aktiva)
    snitt = aktiva.mean() if n_snitt else pd.Series(0.0, index=kolumner)
    rank_k = [k for k, _ in SPINDEL if k != "bollforlust"]
    if n_snitt:
        poang = (100 * aktiva[rank_k] / h_max[rank_k].replace(0, np.nan)).fillna(0).sum(axis=1)
        topp3 = list(poang.sort_values(ascending=False).head(3).index)
    else:
        topp3 = []
    lag_tot = lag_summering(lagstat)
    spara_csv(lagstat.sort_values("datum"), utmapp / "lag_statistik_per_match.csv")
    spara_csv(pd.DataFrame([lag_tot]), utmapp / "lag_statistik_serie.csv")

    spara_csv(sammanfattning.merge(h_tot.reset_index(), on="spelare", how="left").fillna(0),
              utmapp / "spelarsammanfattning.csv")
    per_match = (bred.groupby(["datum", "match", "spelare"])[["P", "T", "U", "M", "avslut", "xg"]]
                     .sum().round(2).reset_index())
    spara_csv(per_match, utmapp / "spelare_per_match.csv")
    spara_csv(handelser.sort_values(["datum", "spelare"]), utmapp / "forsvar_och_dueller_per_match.csv")
    rel = (100 * h_tot / h_max.replace(0, float("nan"))).round(0).add_suffix("_procent_av_max").fillna(0)
    spara_csv(h_tot.join(rel).reset_index(), utmapp / "forsvar_och_dueller_spelare.csv")
    zoner = pd.concat([
        bred.groupby("zon")[["P", "T", "U", "M", "avslut", "xg"]].sum().round(2).reset_index().assign(riktning="OFF (våra avslut)"),
        bred_def.groupby("zon")[["P", "T", "U", "M", "avslut", "xg"]].sum().round(2).reset_index().assign(riktning="DEF (avslut mot oss)"),
    ])
    zoner["zonnamn"] = zoner["zon"].map(ZONNAMN)
    spara_csv(zoner[["riktning", "zon", "zonnamn", "P", "T", "U", "M", "avslut", "xg"]], utmapp / "lag_zoner.csv")

    plan = ladda_plan(skriptmapp)
    rita_lag(bred, bred_def, sammanfattning, matcher, utmapp, plan, lag_tot, h_tot, h_max, snitt, n_snitt, topp3)
    for _, rad in sammanfattning.iterrows():
        rita_spelare(bred, rad, matcher, utmapp, plan, h_tot, h_max, snitt, n_snitt)
    skapa_portal(utmapp, sammanfattning)
    skapa_spelarsidor(utmapp, sammanfattning)
    print(sammanfattning.to_string(index=False))
    if var:
        grupper = {}
        for match, text in sorted(var):
            grupper.setdefault(text, []).append(match)
        print("\nOBS – saker att kolla i dina filer:")
        for text, lista in grupper.items():
            var_ = ", ".join(lista) if len(lista) <= 4 else f"{', '.join(lista[:3])} m.fl. ({len(lista)} matcher)"
            print(f" - [{var_}] {text}")
    print(f"\nFiler sparade i: {utmapp.resolve()}")


if __name__ == "__main__":
    main()
