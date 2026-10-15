"""
Actualiza los datos del dashboard (index.html) a partir del Excel de auditoría.

Uso:  python scripts/build_data.py [ruta_al_excel]
Sin argumento toma el Excel más reciente de la carpeta data/.

Hojas que lee:
  - BASE GENERAL            -> visitas, POP, apreciación y competencia (columna MES)
  - BASE PRECIOS ...        -> precios (la hoja "BASE PRECIOS" con columna MES y más filas)
  - BASE AGO / BASE SEP ... -> solo para las fechas de cada visita (columna "Marca temporal")
"""
import json, re, subprocess, sys, unicodedata
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
HTML = ROOT / "index.html"
START, END = "<!--YC_DATA_START-->", "<!--YC_DATA_END-->"

MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO",
         "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
ABR = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "set", "oct", "nov", "dic"]
MER = {"FERIA LOS INCAS": "Feria Los Incas", "MERCADO ZAMACOLA": "Mercado Zamacola"}
MAT = {"AFICHE": "Afiche", "COLGANTE": "Colgante", "FLANGER O JALAVISTA": "Flanger / jalavista",
       "VENTANA CON CENEFA": "Ventana con cenefa", "MARCA PRECIOS": "Marca precios",
       "CINTA FRONTERIZADORA": "Cinta fronterizadora"}
MARCA = {"FLORIDA": "Florida", "COMPASS": "Compass", "MARINERO": "Marinero", "ACONCAGUA": "Aconcagua",
         "MC COLINS": "McColin's", "MC COLLINS": "McColin's", "A1": "A1"}
COMP = {"D' GUSTO": "D' Gusto", "DE VALLE": "De Valle", "VALLE NORTE": "Valle Norte", "2 CABALLOS": "2 Caballos"}
PRODUCTOS = {  # nombres ya usados en el dashboard; los nuevos se formatean solos
    "ACONCAGUA COCKTAIL FRUTA 12/822G": "Aconcagua Cocktail Fruta 12/822 g",
    "ACONCAGUA DURAZ. MITAD A/FACIL 12/822G": "Aconcagua Duraznos Mitad A/Fácil 12/822 g",
    "COMPASS COCKTAIL DE FRUTAS 12/820GR": "Compass Cocktail de Frutas 12/820 g",
    "COMPASS DURAZNOS MITADES (G) 12/820GR": "Compass Duraznos Mitades 12/820 g",
    "COMPASS FILETE AT AC VEG 150/100": "Compass Filete AT AC Veg 150/100",
    "COMPASS GRATED AT AC VEG 140/100": "Compass Grated AT AC Veg 140/100",
    "COMPASS PIÑA EN RODAJAS 24/565 GR": "Compass Piña en Rodajas 24/565 g",
    "COMPASS PIÑA GOLD EN AG COCO 24/425": "Compass Piña Gold en agua de coco 24/425",
    "COMPASS PIÑA RODAJAS 24/560 GR": "Compass Piña Rodajas 24/560 g",
    "COMPASS TROZOS AT AC VEG 140/100": "Compass Trozos AT AC Veg 140/100",
    "COMPASS TROZOS AT AC VEG 150/100": "Compass Trozos AT AC Veg 150/100",
    "FLORIDA FILETE AGUA SAL 110K 140/100": "Florida Filete Agua Sal 110 K 140/100",
    "FLORIDA FILETE AT AC VEG 140/100": "Florida Filete AT AC Veg 140/100",
    "FLORIDA FILETE DE ATÚN AT AC GIR 140/100": "Florida Filete de Atún AT AC Gir 140/100",
    "FLORIDA GRATED AT AC VEG 140/115": "Florida Grated AT AC Veg 140/115",
    "FLORIDA SOLIDO AT AC VEG 140/100": "Florida Sólido AT AC Veg 140/100",
    "FLORIDA TROZOS DE ATÚN AC. VEG. 140GR": "Florida Trozos de Atún AC Veg 140 g",
    "MARINERO FILETE AT AC VEG 140/100": "Marinero Filete AT AC Veg 140/100",
    "MARINERO FILETE CABALL AC 140/100": "Marinero Filete Caballa AC 140/100",
    "MARINERO GRATED AT AC VEG 140/100": "Marinero Grated AT AC Veg 140/100",
    "MARINERO TROZ AT AC VEG 140/110": "Marinero Trozos AT AC Veg 140/110",
}
COLS = ["cliente", "mes", "mercado", "visita", "abierto", "permitio", "marca", "material", "encontro",
        "visible", "contaminado", "buen", "cant_pop", "validos", "florida", "compass", "marinero",
        "aconcagua", "mccolins", "a1", "exh_visible", "exh_limpia", "comp_q", "marca_comp", "obs", "puntos"]
AP_COLS = ["florida", "compass", "marinero", "aconcagua", "mccolins", "a1", "exh_visible", "exh_limpia", "comp_q"]


def clean(s):
    return " ".join(str(s).split()) if isinstance(s, str) else s

def up(s):
    return clean(s).upper() if isinstance(s, str) else ""

def mer(x):
    return MER.get(up(x), clean(str(x)).title())

def vis(x):
    m = re.search(r"(\d+)", str(x))
    return f"V{m.group(1)}" if m else clean(str(x))

def num(x):
    if x is None or pd.isna(x):
        return ""
    x = float(x)
    return str(int(x)) if x.is_integer() else str(x)

def mes_label(m):
    return up(m).capitalize()

def producto(p):
    p = clean(p)
    if p.upper() in PRODUCTOS:
        return PRODUCTOS[p.upper()]
    s = p.title()
    s = re.sub(r"^Mc Coll?ins", "McColin's", s)
    s = re.sub(r"\b(At|Ac)\b", lambda m: m.group(1).upper(), s)
    s = re.sub(r"\b(De|En|Y|Con)\b", lambda m: m.group(1).lower(), s)
    s = re.sub(r"(\d)\s*(Gr|G)\b", r"\1 g", s)
    return s

def estado(abierto, permitio, obs):
    if "NO EXISTE" in up(obs):
        return "No existe"
    if up(abierto) == "NO":
        return "Cerrado"
    if up(permitio) == "NO":
        return "No permitió auditoría"
    return "Auditado"


def leer_base_general(xl):
    hoja = next((h for h in xl.sheet_names if up(h) == "BASE GENERAL"), None)
    if hoja is None:
        sys.exit("ERROR: no encontré la hoja 'BASE GENERAL' en el Excel.")
    bg = pd.read_excel(xl, hoja).iloc[:, :26]
    bg.columns = COLS
    bg = bg[bg.cliente.notna() & bg.mes.notna()].copy()
    bg["cliente"] = bg.cliente.map(clean)
    bg["mes"] = bg.mes.map(up)
    return bg


def tablas_del_mes(m):
    k = ["cliente", "visita"]
    det = m[m.marca.notna()]
    st = det.groupby(k, sort=False).agg(
        mercado=("mercado", "first"), abierto=("abierto", "first"), permitio=("permitio", "first"),
        pop=("encontro", "sum"), visibles=("visible", "sum"), contaminados=("contaminado", "sum"),
        buen=("buen", "sum"))
    validos = m[m.cant_pop.notna() | m.validos.notna()].groupby(k).validos.first()
    puntos = m[m.puntos.notna()].groupby(k).puntos.first()
    obs = m[m.obs.notna()].groupby(k).obs.first()
    ap = m[m[AP_COLS].notna().any(axis=1)].groupby(k).first()

    visitas, comp = [], {}
    for (cli, v), r in st.iterrows():
        est = estado(r.abierto, r.permitio, obs.get((cli, v)))
        aud, V, M = est == "Auditado", vis(v), mer(r.mercado)
        p = puntos.get((cli, v))
        p = 0 if p is None or pd.isna(p) else int(p)
        a = ap.loc[(cli, v)] if (cli, v) in ap.index else None
        g = lambda c: up(a[c]) if (a is not None and aud) else ""
        val = validos.get((cli, v))
        visitas.append({
            "id": f"{cli} · {V}", "pdv": cli, "mercado": M, "visita": V, "estado": est,
            "auditado": "1" if aud else "0",
            "pop": num(r["pop"]) if aud else "", "visibles": num(r.visibles) if aud else "",
            "contaminados": num(r.contaminados) if aud else "", "buen_estado": num(r.buen) if aud else "",
            "validos": (num(val) or "0") if aud else "",
            "puntos": str(p), "ganador": "1" if p >= 5 else "0",
            "exh_visible": g("exh_visible"), "exh_limpia": g("exh_limpia"), "florida": g("florida"),
            "compass": g("compass"), "marinero": g("marinero"), "aconcagua": g("aconcagua"),
            "mccolins": g("mccolins"), "a1": g("a1")})
        if aud and a is not None and isinstance(a.comp_q, str):
            for b in {COMP.get(up(x), up(x).title()) for x in a.comp_q.split(",") if x.strip()}:
                comp[(M, V, b)] = comp.get((M, V, b), 0) + 1

    auditados = {(x["pdv"], x["visita"]) for x in visitas if x["auditado"] == "1"}
    d = det.assign(V=det.visita.map(vis), M=det.mercado.map(mer))
    d = d[[(c, v) in auditados for c, v in zip(d.cliente, d.V)]]
    pop = []
    for (M, V, b, mt), g in d.groupby(["M", "V", "marca", "material"]):
        B, MT = MARCA.get(up(b), str(b).title()), MAT.get(up(mt), str(mt).capitalize())
        pop.append({"id": f"{M}|{V}|{B}|{MT}", "mercado": M, "visita": V, "marca": B, "material": MT,
                    "pop": num(g.encontro.fillna(0).sum())})
    competencia = [{"id": f"{M}|{V}|{b}", "mercado": M, "visita": V, "marca": b, "pdv": str(n)}
                   for (M, V, b), n in sorted(comp.items())]
    return visitas, pop, competencia


def leer_precios(xl):
    cand = []
    for h in xl.sheet_names:
        if not up(h).startswith("BASE PRECIOS") or "COMPETENCIA" in up(h):
            continue
        df = pd.read_excel(xl, h)
        if {"MES", "PRODUCTO", "MODALIDAD DE VENTA"} <= set(df.columns):
            cand.append(df)
    if not cand:
        print("AVISO: no encontré la hoja 'BASE PRECIOS' con columna MES; el tab de precios quedará vacío.")
        return pd.DataFrame()
    df = max(cand, key=len).copy()
    mod = df["MODALIDAD DE VENTA"].map(up)
    df = df[mod.isin(["POR UNIDAD", "POR CAJA"])].copy()
    df["modalidad"] = df["MODALIDAD DE VENTA"].map(up).map({"POR UNIDAD": "Unidad", "POR CAJA": "Caja"})
    df["precio"] = df["PRECIO POR UNIDAD"].where(df.modalidad == "Unidad", df["PRECIO POR CAJA"])
    df = df[pd.to_numeric(df.precio, errors="coerce").notna()]
    df["mes"] = df.MES.map(up)
    return df


def precios_del_mes(df, mes):
    out = []
    for _, r in df[df.mes == mes].iterrows():
        cli, V, n = clean(r["NOMBRE DEL CLIENTE"]), vis(r["NÚMERO DE VISITA"]), producto(r.PRODUCTO)
        out.append({"id": f"{cli}|{V}|{n}|{r.modalidad}", "pdv": cli, "mercado": mer(r.MERCADO), "visita": V,
                    "marca": MARCA.get(up(r.MARCA), clean(str(r.MARCA)).title()), "nombre": n,
                    "modalidad": r.modalidad, "precio": str(float(r.precio))})
    return out


def leer_fechas(xl):
    """Rango de fechas de cada visita, desde las hojas BASE AGO / BASE SEP / BASE OCT ..."""
    fechas = {}
    for h in xl.sheet_names:
        H = up(h)
        if not H.startswith("BASE ") or "PRECIO" in H or H == "BASE GENERAL":
            continue
        clave = H[5:8].replace("SET", "SEP")
        mes = next((m for m in MESES if m.startswith(clave)), None)
        if not mes:
            continue
        raw = pd.read_excel(xl, h, header=None, nrows=10)
        fila = next((i for i, row in raw.iterrows() if any(up(c) == "MARCA TEMPORAL" for c in row)), None)
        if fila is None:
            continue
        df = pd.read_excel(xl, h, header=fila)
        df.columns = [up(c) for c in df.columns]
        if "NÚMERO DE VISITA" not in df.columns:
            continue
        df["f"] = pd.to_datetime(df["MARCA TEMPORAL"], errors="coerce")
        for v, g in df.dropna(subset=["f"]).groupby("NÚMERO DE VISITA"):
            dias = g.f.dt.normalize().value_counts()
            if len(g) >= 10 and (dias > 1).any():  # ignora registros sueltos cargados días después
                g = g[g.f.dt.normalize().isin(dias[dias > 1].index)]
            a, b = g.f.min(), g.f.max()
            if (a.month, a.day) == (b.month, b.day):
                txt = f"{a.day} {ABR[a.month-1]}"
            elif a.month == b.month:
                txt = f"{a.day}–{b.day} {ABR[a.month-1]}"
            else:
                txt = f"{a.day} {ABR[a.month-1]}–{b.day} {ABR[b.month-1]}"
            fechas.setdefault(mes_label(mes), {})[vis(v)] = txt
    return fechas


def excel_mas_reciente():
    files = [p for p in (ROOT / "data").glob("*.xls*") if not p.name.startswith("~$")]
    if not files:
        sys.exit("ERROR: no hay ningún Excel en la carpeta data/.")
    def t(p):
        try:
            out = subprocess.run(["git", "log", "-1", "--format=%ct", "--", str(p)], cwd=ROOT,
                                 capture_output=True, text=True).stdout.strip()
            return int(out) if out else p.stat().st_mtime
        except Exception:
            return p.stat().st_mtime
    return max(files, key=t)


def main():
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else excel_mas_reciente()
    print(f"Leyendo {path.name}")
    xl = pd.ExcelFile(path)
    bg = leer_base_general(xl)
    precios = leer_precios(xl)
    fechas = leer_fechas(xl)
    meses = sorted(bg.mes.unique(), key=lambda m: MESES.index(m) if m in MESES else 99)

    data = {}
    for m in meses:
        visitas, pop, comp = tablas_del_mes(bg[bg.mes == m])
        pr = precios_del_mes(precios, m) if len(precios) else []
        L = mes_label(m)
        data[L] = {"data": {"visitas": visitas, "pop": pop, "precios": pr, "competencia": comp},
                   "fechas": fechas.get(L, {})}
        aud = sum(v["auditado"] == "1" for v in visitas)
        gan = sum(v["ganador"] == "1" for v in visitas)
        print(f"  {L}: {len(visitas)} visitas ({aud} auditadas, {gan} ganadoras), {len(pop)} filas POP, "
              f"{len(pr)} precios, {len(comp)} filas competencia, fechas {data[L]['fechas']}")

    orden = list(data.keys())
    js = (f"{START}<script>window.YC_MESES={json.dumps(data, ensure_ascii=False)};"
          f"window.YC_MES_ORDEN={json.dumps(orden, ensure_ascii=False)};"
          f"window.YC_DATA=window.YC_MESES[window.YC_MES_ORDEN[window.YC_MES_ORDEN.length-1]].data;</script>{END}")
    html = HTML.read_text(encoding="utf-8")
    i, j = html.find(START), html.find(END)
    if i < 0 or j < 0:
        sys.exit("ERROR: index.html no tiene los marcadores YC_DATA_START / YC_DATA_END.")
    HTML.write_text(html[:i] + js + html[j + len(END):], encoding="utf-8")
    print(f"index.html actualizado · meses: {', '.join(orden)}")


if __name__ == "__main__":
    main()
