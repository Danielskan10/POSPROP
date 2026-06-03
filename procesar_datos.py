# ╔══════════════════════════════════════════════════════════════════╗
# ║  SAPIENZA · Procesador de Posiciones de Portafolio              ║
# ║  Flujo:                                                         ║
# ║    1. Lee CSV/XLSX nuevos (caché incremental — no reprocesa)    ║
# ║    2. Genera data.json y lo sube a GitHub                       ║
# ║    3. GitHub Pages sirve el dashboard actualizado               ║
# ║  Universal: funciona en cualquier PC sin cambiar rutas          ║
# ╚══════════════════════════════════════════════════════════════════╝

import os, re, glob, json, warnings, hashlib
from datetime import datetime
from pathlib import Path

import pandas as pd
import numpy as np

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════
#  ★  CONFIGURA ESTAS DOS COSAS — lo demás no lo toques
# ══════════════════════════════════════════════════════════════════

# 1. Carpeta donde están tus archivos CSV/XLSX
#    Puede ser cualquier ruta del PC. Ej: r"C:\Datos\Posiciones"
CARPETA_DATOS = r"C:\Users\danie\Sapienza\POSPRO"

# 2. Token de GitHub para subir los datos
#    Crea uno en: https://github.com/settings/tokens
#    Permisos necesarios: repo (full control)
GITHUB_TOKEN = ""   # ← pegar tu token aquí

# ══════════════════════════════════════════════════════════════════
#  Repositorio de destino — cambiar si usas uno diferente
GITHUB_REPO  = "Danielskan10/POSPROP"    # usuario/repositorio
GITHUB_BRANCH = "main"
# ══════════════════════════════════════════════════════════════════

# Rutas internas (no tocar)
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT        = CARPETA_DATOS.strip() if CARPETA_DATOS.strip() else _SCRIPT_DIR

# ┌─────────────────────────────────────────────────────────────────┐
# │  Lo demás no necesitas tocarlo                                   │
# └─────────────────────────────────────────────────────────────────┘
CFG = {
    "output_json":   os.path.join(_SCRIPT_DIR, "data.json"),
    "cache_file":    os.path.join(_SCRIPT_DIR, ".cache_procesado.json"),

    # ── Git ──────────────────────────────────────────────────────
    "git_push":   True,
    "git_msg":    "data: posiciones {fecha}",
    "git_branch": "main",

    # ── Identidad ────────────────────────────────────────────────
    "org": "Skandia Colombia",
    "sub": "Dashboard de Posiciones de Portafolio",

    # ── Archivos ignorados al buscar datos fuente ────────────────
    "ignorar": {
        "posiciones_consolidadas.xlsx", "index.html",
        "dashboard.html", "data.json", "_dashboard_tpl.html",
    },

    # ════════════════════════════════════════════════════════════
    #  MAPAS EDITABLES — se publican en data.json y el dashboard
    #  los lee para mostrar nombres legibles. Puedes agregar,
    #  quitar o renombrar cualquier entrada.
    # ════════════════════════════════════════════════════════════

    # Código portafolio (campo "Por" antes del guión) → nombre legible
    "ports": {
        "21": "Portafolio 21",
        "41": "Portafolio 41",
        "51": "Portafolio 51",
        "HC": "HC Cesantías",
        "HO": "HO Obligatorio",
        "HE": "HE Especial",
    },

    # Código tipo de activo (campo "Por" después del guión) → nombre legible
    "activos": {
        "F":  "Renta Fija (TES)",
        "H":  "Liquidez",
        "E":  "Fondos Colectivos",
        "O":  "Otros Títulos",
        "D":  "Depósitos/Cash",
        "Y":  "Fondos SPC",
        "V":  "AOR",
        "L":  "Acciones",
        "G":  "Fondos",
        "DF": "Derivados",
        "T":  "TIDIS",
        "DB": "Depósitos Banco",
        "HB": "Liquidez HB",
        "W":  "AOR Propios",
        "P":  "Otros",
    },

    # Código moneda del sistema → nombre legible
    "monedas": {
        "COP":     "COP",
        "USCOP":   "USD",
        "EUO":     "EUR",
        "UKCOP":   "GBP",
        "ETQACOP": "Fdo ETQA",
        "ETRACOP": "Fdo ETRA",
        "ESJWX":   "Fdo Alt A",
        "ESY0O":   "Fdo UnoMas",
        "EIBMS":   "Fdo Efect A",
        "EIEMS":   "Fdo Efect D",
        "EABFR":   "Fdo Occirenta",
        "ETGWX":   "Fdo Alt D",
        "EDXR5":   "SPC Corto Plazo",
        "EDY1U":   "SPC Largo Plazo",
        "EDT1U":   "SPC Conservador",
        "EDU1U":   "SPC Mayor Riesgo",
        "EDSR6":   "SPC Moderado",
        "EDV1U":   "SPC Retiro Programado",
        "ESUMS":   "Fdo Umas",
    },

    # Código modalidad → descripción legible
    "modalidades": {
        "AV":  "Año Vencido",
        "DV":  "Día Vencido",
        "Dto": "Descuento",
        "MV":  "Mes Vencido",
        "TV":  "Trimestre Vencido",
        "NAp": "No Aplica",
    },

    # Código estado → descripción legible
    "estados": {
        "Pend": "Pendiente",
        "Reci": "En Cartera",
        "Vend": "Vendido",
        "Frac": "Fraccionado",
    },

    # Código método de valoración → descripción legible
    "metodos": {
        "QSI":   "Quantil (Fondos)",
        "QES-SI":"Quantil Renta Fija",
        "MC4-E": "Modelo Equity",
        "MC1-I": "Modelo Interés",
    },

    # Reglas de clasificación por tipo de instrumento.
    # Cada regla: {"contiene": ["texto"], "tipo": "TES"}
    # Se evalúan en orden; la primera que coincide gana.
    # El texto se compara en MAYÚSCULAS con el nombre de la especie.
    "tipo_reglas": [
        {"contiene": ["TES"],                              "tipo": "TES"},
        {"contiene": ["CDT", "CREDITO", "CRÉDITO"],        "tipo": "CDT"},
        {"contiene": ["CASH", "CTA AHO", "CTA AH", "CTA "],"tipo": "Liquidez"},
        {"contiene": ["FIC", "FCPD", "FCPE", "FCP", "P SPC"],"tipo": "Fondos"},
        {"contiene": ["DER.", "DERIV"],                    "tipo": "Derivados"},
        {"contiene": ["AOR"],                              "tipo": "AOR"},
        {"contiene": ["TIDIS"],                            "tipo": "TIDIS"},
        {"contiene": ["TITULARIZ"],                        "tipo": "Titularizaciones"},
        {"contiene": ["ACC."],                             "tipo": "Acciones"},
    ],

    # Colores de los tipos de instrumento en las gráficas
    "tipo_colores": {
        "TES":             "#3b82f6",
        "Fondos":          "#00854A",
        "Liquidez":        "#10b981",
        "Derivados":       "#ef4444",
        "AOR":             "#f59e0b",
        "TIDIS":           "#8b5cf6",
        "Titularizaciones":"#f97316",
        "Acciones":        "#ec4899",
        "CDT":             "#06b6d4",
        "Otro":            "#6b7280",
    },

    # Reclasificación manual de especies específicas
    # (se puede gestionar también desde el dashboard)
    "tipo_map": {
        # "Credito Bancolombia 365 $": "CDT",
    },
}
# ──────────────────────────────────────────────────────────────────

COLS_RAW = [
    "Especie","Titulo","Inver","F_Vcto","Vlr_Nominal","Facial","Mod",
    "Desde","Hasta","Vlr_Mer_Ant","Vlr_Mer_Hoy","Adeudados",
    "Causacion_Mer","Causacion_TIR","ISIN_Nemot","Met","Precio",
    "Marg_Efec","TIR_Mercado","Moneda","Mnd_Val_An","Mnd_Val",
    "Dif_cambio","Causacion_Moneda","Causacion_Tasa","Dias","Por","Est"
]
COLS_NUM = [
    "Vlr_Nominal","Vlr_Mer_Ant","Vlr_Mer_Hoy","Adeudados",
    "Causacion_Mer","Causacion_TIR","Precio","TIR_Mercado",
    "Mnd_Val_An","Mnd_Val","Dif_cambio","Causacion_Moneda",
    "Causacion_Tasa","Dias","Inver","Titulo"
]
COLS_DT = ["F_Vcto","Desde","Hasta"]

# Accesos cortos a los mapas del CFG (se usan en transformación)
def _PORTS():    return CFG["ports"]
def _ACTIVOS():  return CFG["activos"]
def _MONEDAS():  return CFG["monedas"]
def _MOD():      return CFG["modalidades"]
def _EST():      return CFG["estados"]
def _MET():      return CFG["metodos"]

# ══════════════════════════════════════════════════════════════════
#  CACHÉ INCREMENTAL
# ══════════════════════════════════════════════════════════════════

def _hash_file(path: str) -> str:
    """MD5 rápido del archivo (primeros 512 KB bastan para detectar cambios)."""
    h = hashlib.md5()
    with open(path, "rb") as f:
        h.update(f.read(524288))
    return h.hexdigest()

def cargar_cache() -> dict:
    p = CFG["cache_file"]
    if os.path.exists(p):
        try:
            with open(p, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def guardar_cache(cache: dict):
    with open(CFG["cache_file"], "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

def archivo_ya_procesado(path: str, cache: dict) -> bool:
    """True si el archivo está en caché y su hash no cambió."""
    key = os.path.relpath(path, ROOT)
    if key not in cache:
        return False
    return cache[key].get("hash") == _hash_file(path)

def marcar_procesado(path: str, cache: dict):
    key = os.path.relpath(path, ROOT)
    cache[key] = {
        "hash":     _hash_file(path),
        "procesado": datetime.now().isoformat(timespec="seconds"),
    }

# ══════════════════════════════════════════════════════════════════
#  LIMPIEZA / PARSEO
# ══════════════════════════════════════════════════════════════════

def _num(v):
    if pd.isna(v): return np.nan
    try: return float(str(v).strip().replace(",","").replace(" ",""))
    except: return np.nan

def _dt(v):
    if pd.isna(v): return pd.NaT
    for fmt in ("%Y/%m/%d","%Y-%m-%d","%d/%m/%Y","%m/%d/%Y"):
        try: return pd.to_datetime(str(v).strip(), format=fmt)
        except: pass
    return pd.NaT

def _fecha_path(p):
    m = re.search(r"(\d{8})", os.path.basename(p))
    if m:
        try: return datetime.strptime(m.group(1), "%Y%m%d").date()
        except: pass
    return None

def _tipo(e):
    """Clasifica especie según tipo_reglas del CFG, luego tipo_map."""
    # 1. Reclasificación manual explícita (mayor prioridad)
    tipo_map = CFG.get("tipo_map", {})
    if e in tipo_map:
        return tipo_map[e]
    # 2. Reglas por contenido de texto
    u = str(e).upper()
    for regla in CFG.get("tipo_reglas", []):
        if any(x in u for x in regla["contiene"]):
            return regla["tipo"]
    return "Otro"

# ══════════════════════════════════════════════════════════════════
#  LECTURA DE ARCHIVOS
# ══════════════════════════════════════════════════════════════════

def _leer_csv(path):
    try:
        df = pd.read_csv(
            path, sep=";", header=None, skiprows=2, names=COLS_RAW,
            encoding="latin1", dtype=str, on_bad_lines="skip"
        )
        df = df[~df["Especie"].str.strip().str.startswith("---", na=True)]
        df = df.dropna(how="all")
        return df[df["Especie"].str.strip().ne("")]
    except Exception as e:
        print(f"  [WARN] {os.path.basename(path)}: {e}")
        return None

def _leer_xlsx(path):
    try:
        frames = []
        for sh in pd.ExcelFile(path).sheet_names:
            raw = pd.read_excel(path, sheet_name=sh, header=None, dtype=str)
            hr = next((i for i, r in raw.iterrows()
                       if any("Especie" in str(v) for v in r.values)), None)
            st = (hr + 2) if hr is not None else 2
            d  = raw.iloc[st:].reset_index(drop=True)
            n  = len(COLS_RAW)
            while len(d.columns) < n:
                d[len(d.columns)] = np.nan
            d = d.iloc[:, :n]
            d.columns = COLS_RAW
            frames.append(d)
        return pd.concat(frames, ignore_index=True) if frames else None
    except Exception as e:
        print(f"  [WARN] {os.path.basename(path)}: {e}")
        return None

# ══════════════════════════════════════════════════════════════════
#  CONVERSIÓN CSV → XLSX
# ══════════════════════════════════════════════════════════════════

def convertir_csvs(cache: dict):
    csvs = list({os.path.normcase(p): p for p in
        glob.glob(os.path.join(ROOT, "**", "*.CSV"), recursive=True) +
        glob.glob(os.path.join(ROOT, "**", "*.csv"), recursive=True)
    }.values())

    convertidos = 0
    for csv_path in sorted(csvs):
        xlsx_path = os.path.splitext(csv_path)[0] + ".xlsx"
        if (os.path.exists(xlsx_path) and
                os.path.getmtime(xlsx_path) >= os.path.getmtime(csv_path)):
            continue
        df = _leer_csv(csv_path)
        if df is not None and not df.empty:
            df.to_excel(xlsx_path, index=False)
            rel = os.path.relpath(xlsx_path, ROOT)
            print(f"  Convertido: {rel}")
            convertidos += 1

    if convertidos:
        print(f"  {convertidos} CSV convertidos a XLSX.")
    else:
        print("  Sin CSV nuevos para convertir.")

# ══════════════════════════════════════════════════════════════════
#  TRANSFORMACIÓN
# ══════════════════════════════════════════════════════════════════

def _transformar(df, fecha):
    df = df.copy()
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].str.strip()
    for c in COLS_NUM:
        if c in df.columns:
            df[c] = df[c].apply(_num)
    for c in COLS_DT:
        if c in df.columns:
            df[c] = df[c].apply(_dt)

    df["Fecha_Posicion"] = pd.to_datetime(fecha)
    df["Especie"]        = df["Especie"].str.strip()
    df["ISIN_Nemot"]     = df["ISIN_Nemot"].str.strip().replace("", np.nan)
    df["Est"]            = df["Est"].str.strip()
    df["Mod"]            = df["Mod"].str.strip()
    df["Met"]            = df["Met"].str.strip()
    df["Facial"]         = df["Facial"].str.strip()
    df["Tipo"]           = df["Especie"].apply(_tipo)
    df["Por"]            = df["Por"].str.strip()
    df["Port_Cod"]       = df["Por"].str.extract(r"^(\d+|H[COEB]+)", expand=False)
    df["Act_Cod"]        = df["Por"].str.extract(r"-(\w+)$",         expand=False)
    df["Port_Nom"]       = df["Port_Cod"].map(_PORTS()).fillna(df["Port_Cod"])
    df["Act_Nom"]        = df["Act_Cod"].map(_ACTIVOS()).fillna(df["Act_Cod"])
    df["Moneda"]         = df["Moneda"].str.replace("$", "COP", regex=False)
    df["Mon_Desc"]       = df["Moneda"].map(_MONEDAS()).fillna(df["Moneda"])
    df["Es_Ext"]         = df["Moneda"].isin(["USCOP","EUO","UKCOP","USD","EUR","GBP"])
    df["PL"]             = df["Vlr_Mer_Hoy"] - df["Vlr_Mer_Ant"]
    df["Rend_Pct"]       = np.where(
        df["Vlr_Mer_Ant"] != 0,
        df["PL"] / df["Vlr_Mer_Ant"] * 100, np.nan
    )
    df["Caus_Spread"]    = df["Causacion_Mer"] - df["Causacion_TIR"]
    if "F_Vcto" in df.columns:
        df["Dias_Vcto"]  = (df["F_Vcto"] - df["Fecha_Posicion"]).dt.days
    df["Mod_Desc"]       = df["Mod"].map(_MOD()).fillna(df["Mod"])
    df["Est_Desc"]       = df["Est"].map(_EST()).fillna(df["Est"])
    df["Met_Desc"]       = df["Met"].map(_MET()).fillna(df["Met"])
    return df

# ══════════════════════════════════════════════════════════════════
#  CARGA CON CACHÉ INCREMENTAL
# ══════════════════════════════════════════════════════════════════

def cargar(cache: dict):
    """
    Carga todos los XLSX.
    - Si existe data.json previo: extrae fechas ya procesadas.
    - Solo procesa archivos nuevos o modificados (hash distinto).
    - Combina datos nuevos con datos previos del JSON.
    """
    ign = CFG["ignorar"]
    xlsxs = [p for p in
        glob.glob(os.path.join(ROOT, "**", "*.xlsx"), recursive=True) +
        glob.glob(os.path.join(ROOT, "**", "*.xls"),  recursive=True)
        if os.path.basename(p) not in ign
    ]
    xlsxs = list({os.path.normcase(p): p for p in xlsxs}.values())

    # Separar archivos nuevos/modificados de los ya procesados
    nuevos    = [p for p in xlsxs if not archivo_ya_procesado(p, cache)]
    ya_ok     = [p for p in xlsxs if     archivo_ya_procesado(p, cache)]

    print(f"XLSX encontrados: {len(xlsxs)}  "
          f"({len(nuevos)} nuevos/modificados, {len(ya_ok)} en caché)")

    # Cargar datos previos del JSON si existe
    datos_previos = pd.DataFrame()
    json_path = CFG["output_json"]
    if os.path.exists(json_path) and ya_ok:
        try:
            with open(json_path, encoding="utf-8") as f:
                prev_json = json.load(f)
            # Reconstruir DataFrame desde la tabla del JSON
            filas = prev_json.get("_raw_tabla", [])
            if filas:
                datos_previos = pd.DataFrame(filas)
                datos_previos["Fecha_Posicion"] = pd.to_datetime(
                    datos_previos["Fecha_Posicion"]
                )
                print(f"  Datos previos cargados del JSON: "
                      f"{len(datos_previos):,} filas | "
                      f"{datos_previos['Fecha_Posicion'].nunique()} fechas")
        except Exception as e:
            print(f"  [WARN] No se pudo leer datos previos del JSON: {e}")

    # Procesar archivos nuevos
    todos_nuevos = []
    for p in sorted(nuevos):
        f  = _fecha_path(p)
        df = _leer_xlsx(p)
        if df is not None and not df.empty:
            df = _transformar(df, f)
            df["_src"] = os.path.basename(p)
            todos_nuevos.append(df)
            marcar_procesado(p, cache)
            print(f"  [NUEVO] {os.path.relpath(p, ROOT)} -> "
                  f"{len(df)} filas | fecha={f}")
        else:
            print(f"  [SKIP]  {os.path.relpath(p, ROOT)} vacío")

    if not todos_nuevos and datos_previos.empty:
        raise SystemExit("[ERROR] No se encontraron datos.")

    # Combinar previos + nuevos
    partes = []
    if not datos_previos.empty:
        partes.append(datos_previos)
    if todos_nuevos:
        partes.append(pd.concat(todos_nuevos, ignore_index=True))

    master = (pd.concat(partes, ignore_index=True)
              .sort_values(["Fecha_Posicion","Especie"], ignore_index=True))

    # Deduplicar: si una misma fecha+especie+portafolio aparece dos veces,
    # quedarse con la más reciente (el archivo más nuevo tiene prioridad)
    master = master.drop_duplicates(
        subset=["Fecha_Posicion","Especie","Por"], keep="last"
    )

    print(f"\nTotal combinado: {len(master):,} registros | "
          f"{master['Fecha_Posicion'].nunique()} fechas | "
          f"{master['Especie'].nunique()} especies")
    return master, bool(todos_nuevos)

# ══════════════════════════════════════════════════════════════════
#  SERIALIZACIÓN
# ══════════════════════════════════════════════════════════════════

def sf(v):
    if v is None: return None
    try:
        f = float(v)
        return None if (np.isnan(f) or np.isinf(f)) else round(f, 4)
    except: return None

def sd(v):
    try: return pd.Timestamp(v).strftime("%Y-%m-%d") if pd.notna(v) else None
    except: return None

def rows(df):
    out = []
    for _, r in df.iterrows():
        row = {}
        for k, v in r.items():
            if isinstance(v, (float, np.floating)):
                row[k] = sf(v)
            elif isinstance(v, (int, np.integer)) and not isinstance(v, bool):
                row[k] = int(v)
            elif isinstance(v, bool) or isinstance(v, np.bool_):
                row[k] = bool(v)
            elif pd.isna(v) if not isinstance(v, (list, dict)) else False:
                row[k] = None
            else:
                row[k] = str(v)
        out.append(row)
    return out

def _raw_row(r):
    """Serializa una fila del master para guardar en _raw_tabla (caché JSON)."""
    keep = [
        "Especie","Titulo","Inver","F_Vcto","Vlr_Nominal","Facial","Mod",
        "Desde","Hasta","Vlr_Mer_Ant","Vlr_Mer_Hoy","Adeudados",
        "Causacion_Mer","Causacion_TIR","ISIN_Nemot","Met","Precio",
        "Marg_Efec","TIR_Mercado","Moneda","Mnd_Val_An","Mnd_Val",
        "Dif_cambio","Causacion_Moneda","Causacion_Tasa","Dias","Por","Est",
        "Fecha_Posicion","Tipo","Port_Cod","Act_Cod","Port_Nom","Act_Nom",
        "Mon_Desc","Es_Ext","PL","Rend_Pct","Caus_Spread",
        "Dias_Vcto","Mod_Desc","Est_Desc","Met_Desc","_src",
    ]
    row = {}
    for k in keep:
        if k not in r.index: continue
        v = r[k]
        if isinstance(v, (float, np.floating)):
            row[k] = sf(v)
        elif isinstance(v, (int, np.integer)) and not isinstance(v, bool):
            row[k] = int(v)
        elif isinstance(v, bool) or isinstance(v, np.bool_):
            row[k] = bool(v)
        elif pd.isna(v) if not isinstance(v, (list, dict)) else False:
            row[k] = None
        elif isinstance(v, pd.Timestamp):
            row[k] = v.isoformat()
        else:
            row[k] = str(v)
    return row

# ══════════════════════════════════════════════════════════════════
#  EXTRAS
# ══════════════════════════════════════════════════════════════════

def _build_extras(m, hoy, evol, fechas_ord):
    nom_evol = m.groupby("Fecha_Posicion").agg(
        nominal=("Vlr_Nominal","sum"),
        mercado=("Vlr_Mer_Hoy","sum"),
        n_titulos=("Especie","nunique"),
        adeudados=("Adeudados","sum"),
    ).reset_index().sort_values("Fecha_Posicion")
    nom_evol["spread_pct"] = (
        (nom_evol["mercado"] - nom_evol["nominal"])
        / nom_evol["nominal"].replace(0, np.nan) * 100
    )

    vcto = hoy[hoy["Dias_Vcto"].notna() & (hoy["Dias_Vcto"] > 0)].copy()
    bins   = [0,30,60,90,180,365,730,9999]
    labels = ["<30d","30-60d","60-90d","90-180d","180d-1a","1a-2a",">2a"]
    vcto["bucket"] = pd.cut(vcto["Dias_Vcto"], bins=bins, labels=labels)
    vcto_buck = vcto.groupby("bucket", observed=True).agg(
        valor=("Vlr_Mer_Hoy","sum"), n=("Especie","count")
    ).reset_index()

    est_evol  = m.groupby(["Fecha_Posicion","Est_Desc"])["Vlr_Mer_Hoy"].sum().reset_index()
    estados_u = sorted(m["Est_Desc"].dropna().unique().tolist())
    est_dict  = {}
    for e in estados_u:
        sub = est_evol[est_evol["Est_Desc"]==e].sort_values("Fecha_Posicion")
        est_dict[e] = [sf(v) for v in sub["Vlr_Mer_Hoy"]]

    total_hoy = hoy["Vlr_Mer_Hoy"].sum()
    top10 = hoy.groupby("Especie")["Vlr_Mer_Hoy"].sum().sort_values(ascending=False).head(10)
    concentracion = [
        {"esp":k,"valor":sf(v),"pct":sf(v/total_hoy*100)}
        for k,v in top10.items()
    ]

    cross = hoy.groupby(["Port_Nom","Tipo"])["Vlr_Mer_Hoy"].sum().reset_index()
    cross_dict = {}
    for _, r in cross.iterrows():
        p = str(r["Port_Nom"]); t = str(r["Tipo"])
        if p not in cross_dict: cross_dict[p] = {}
        cross_dict[p][t] = sf(r["Vlr_Mer_Hoy"])

    ports = m["Port_Nom"].dropna().unique()
    nom_port_dict = {}
    n_port_dict   = {}
    for p in ports:
        sub_n = m[m["Port_Nom"]==p].groupby("Fecha_Posicion")["Vlr_Nominal"].sum().reset_index().sort_values("Fecha_Posicion")
        nom_port_dict[str(p)] = [sf(v) for v in sub_n["Vlr_Nominal"]]
        sub_c = m[m["Port_Nom"]==p].groupby("Fecha_Posicion")["Especie"].nunique().reset_index().sort_values("Fecha_Posicion")
        n_port_dict[str(p)] = [int(v) for v in sub_c["Especie"]]

    # Rentabilidad mensual
    m2 = m.copy()
    m2["mes"] = m2["Fecha_Posicion"].dt.to_period("M").astype(str)
    rent_mes = m2.groupby("mes")["PL"].sum().reset_index()
    rent_mes_dict = {r["mes"]: sf(r["PL"]) for _, r in rent_mes.iterrows()}

    # Correlación portafolios (últimos 30 días si hay suficientes datos)
    corr_dict = {}
    try:
        if len(fechas_ord) >= 5:
            pivot = m.pivot_table(
                index="Fecha_Posicion", columns="Port_Nom",
                values="PL", aggfunc="sum"
            )
            corr = pivot.corr().round(3)
            corr_dict = {
                col: {row: sf(corr.loc[row, col]) for row in corr.index}
                for col in corr.columns
            }
    except Exception:
        pass

    return {
        "nom_evol": {
            "nominal":   [sf(v) for v in nom_evol["nominal"]],
            "mercado":   [sf(v) for v in nom_evol["mercado"]],
            "n_titulos": [int(v) for v in nom_evol["n_titulos"]],
            "adeudados": [sf(v) for v in nom_evol["adeudados"]],
            "spread_pct":[sf(v) for v in nom_evol["spread_pct"]],
        },
        "vcto_buckets": {
            "labels": [str(r["bucket"]) for _,r in vcto_buck.iterrows()],
            "valores": [sf(r["valor"]) for _,r in vcto_buck.iterrows()],
            "n":       [int(r["n"]) for _,r in vcto_buck.iterrows()],
        },
        "estados":        est_dict,
        "estados_u":      estados_u,
        "concentracion":  concentracion,
        "cross":          cross_dict,
        "nom_port":       nom_port_dict,
        "n_port":         n_port_dict,
        "rent_mensual":   rent_mes_dict,
        "correlacion":    corr_dict,
    }

# ══════════════════════════════════════════════════════════════════
#  BUILD JSON
# ══════════════════════════════════════════════════════════════════

def build_json(m):
    fechas_ord = sorted(m["Fecha_Posicion"].unique())
    ult  = fechas_ord[-1]
    hoy  = m[m["Fecha_Posicion"] == ult]

    # Series temporales
    evol = (m.groupby("Fecha_Posicion").agg(
        total=("Vlr_Mer_Hoy","sum"), pl=("PL","sum"),
        caus_mer=("Causacion_Mer","sum"), caus_tir=("Causacion_TIR","sum"),
        caus_mon=("Causacion_Moneda","sum"), caus_tasa=("Causacion_Tasa","sum"),
        adeudados=("Adeudados","sum"), n=("Especie","count"),
        nominal=("Vlr_Nominal","sum"),
    ).reset_index().sort_values("Fecha_Posicion"))
    evol["pl_acum"]   = evol["pl"].cumsum()
    evol["caus_acum"] = evol["caus_mer"].cumsum()
    evol["rend_pct"]  = evol["pl"] / evol["total"].shift(1).replace(0, np.nan) * 100

    # Por portafolio × fecha
    ep = m.groupby(["Fecha_Posicion","Port_Nom"]).agg(
        total=("Vlr_Mer_Hoy","sum"), pl=("PL","sum"),
        caus=("Causacion_Mer","sum"), n=("Especie","count"),
        nominal=("Vlr_Nominal","sum"),
    ).reset_index()
    ports_u = sorted(m["Port_Nom"].dropna().unique().tolist())
    evol_port = {}
    for p in ports_u:
        s = ep[ep["Port_Nom"]==p].sort_values("Fecha_Posicion")
        evol_port[p] = {
            "total":   [sf(v) for v in s["total"]],
            "pl":      [sf(v) for v in s["pl"]],
            "caus":    [sf(v) for v in s["caus"]],
            "nominal": [sf(v) for v in s["nominal"]],
        }

    # Por tipo × fecha
    et = m.groupby(["Fecha_Posicion","Tipo"]).agg(
        total=("Vlr_Mer_Hoy","sum"), pl=("PL","sum")
    ).reset_index()
    tipos_u = sorted(m["Tipo"].unique().tolist())
    evol_tipo = {}
    for t in tipos_u:
        s = et[et["Tipo"]==t].sort_values("Fecha_Posicion")
        evol_tipo[t] = [sf(v) for v in s["total"]]

    # Historia por especie
    esp_hist = {}
    for esp, grp in m.groupby("Especie"):
        g = grp.sort_values("Fecha_Posicion")
        esp_hist[esp] = {
            "fechas":    [sd(d) for d in g["Fecha_Posicion"]],
            "total":     [sf(v) for v in g["Vlr_Mer_Hoy"]],
            "pl":        [sf(v) for v in g["PL"]],
            "caus":      [sf(v) for v in g["Causacion_Mer"]],
            "caus_tir":  [sf(v) for v in g["Causacion_TIR"]],
            "caus_mon":  [sf(v) for v in g["Causacion_Moneda"]],
            "caus_tasa": [sf(v) for v in g["Causacion_Tasa"]],
            "tir":       [sf(v) for v in g["TIR_Mercado"]],
            "precio":    [sf(v) for v in g["Precio"]],
            "port":      str(g["Port_Nom"].iloc[-1]) if pd.notna(g["Port_Nom"].iloc[-1]) else "",
            "tipo":      str(g["Tipo"].iloc[-1]),
        }

    # Tabla detalle último día
    def _det(r):
        return {
            "esp":       str(r["Especie"]),
            "port":      str(r["Port_Nom"])   if pd.notna(r["Port_Nom"])   else "",
            "act":       str(r["Act_Nom"])    if pd.notna(r["Act_Nom"])    else "",
            "tipo":      str(r["Tipo"]),
            "isin":      str(r["ISIN_Nemot"]) if pd.notna(r["ISIN_Nemot"])else "",
            "nominal":   sf(r["Vlr_Nominal"]),
            "v_ant":     sf(r["Vlr_Mer_Ant"]),
            "valor":     sf(r["Vlr_Mer_Hoy"]),
            "pl":        sf(r["PL"]),
            "rend":      sf(r["Rend_Pct"]),
            "caus_mer":  sf(r["Causacion_Mer"]),
            "caus_tir":  sf(r["Causacion_TIR"]),
            "caus_mon":  sf(r["Causacion_Moneda"]),
            "caus_tasa": sf(r["Causacion_Tasa"]),
            "caus_diff": sf(r.get("Caus_Spread", 0)),
            "adeudados": sf(r["Adeudados"]),
            "tir":       sf(r["TIR_Mercado"]),
            "precio":    sf(r["Precio"]),
            "moneda":    str(r["Mon_Desc"]),
            "ext":       bool(r["Es_Ext"]),
            "mnd_hoy":   sf(r["Mnd_Val"]),
            "mnd_ant":   sf(r["Mnd_Val_An"]),
            "dif_fx":    sf(r["Dif_cambio"]),
            "vcto":      sd(r.get("F_Vcto")),
            "dias_vcto": sf(r.get("Dias_Vcto")),
            "facial":    str(r["Facial"])    if pd.notna(r["Facial"])    else "",
            "mod":       str(r["Mod_Desc"])  if pd.notna(r["Mod_Desc"])  else "",
            "est":       str(r["Est_Desc"])  if pd.notna(r["Est_Desc"])  else "",
            "met":       str(r["Met_Desc"])  if pd.notna(r["Met_Desc"])  else "",
            "por":       str(r["Por"]),
            "desde":     sd(r.get("Desde")),
            "hasta":     sd(r.get("Hasta")),
            "dias":      sf(r["Dias"]),
        }

    tabla = [_det(r) for _, r in hoy.sort_values("Vlr_Mer_Hoy", ascending=False).iterrows()]
    tes   = [_det(r) for _, r in
             hoy[hoy["Tipo"].isin(["TES","TIDIS","Titularizaciones","CDT"])]
             .sort_values("Vlr_Mer_Hoy", ascending=False).iterrows()]

    # Resúmenes
    def _gagg(grp_col):
        agg = {"total":("Vlr_Mer_Hoy","sum"),"pl":("PL","sum"),
               "caus":("Causacion_Mer","sum"),"caus_tir":("Causacion_TIR","sum"),
               "caus_mon":("Causacion_Moneda","sum"),"caus_tasa":("Causacion_Tasa","sum"),
               "n":("Especie","count")}
        df_g = (hoy.groupby(grp_col).agg(**{k:v for k,v in agg.items()})
                .reset_index().sort_values("total", ascending=False))
        df_g["pct"] = df_g["total"] / df_g["total"].sum() * 100
        return rows(df_g)

    by_port = _gagg("Port_Nom")
    by_tipo = _gagg("Tipo")
    by_mon  = _gagg("Mon_Desc")
    by_act  = _gagg("Act_Nom")
    by_mod  = _gagg("Mod")

    # Causación histórica por portafolio
    caus_hist = {}
    for p in ports_u:
        sub = m[m["Port_Nom"]==p].groupby("Fecha_Posicion").agg(
            caus_mer=("Causacion_Mer","sum"), caus_tir=("Causacion_TIR","sum"),
            caus_mon=("Causacion_Moneda","sum"), caus_tasa=("Causacion_Tasa","sum"),
        ).reset_index().sort_values("Fecha_Posicion")
        caus_hist[p] = {
            "mer":  [sf(v) for v in sub["caus_mer"]],
            "tir":  [sf(v) for v in sub["caus_tir"]],
            "mon":  [sf(v) for v in sub["caus_mon"]],
            "tasa": [sf(v) for v in sub["caus_tasa"]],
        }

    comp = rows(hoy.groupby(["Port_Nom","Tipo"])["Vlr_Mer_Hoy"].sum().reset_index())

    # KPIs
    tot    = sf(hoy["Vlr_Mer_Hoy"].sum())
    ant    = sf(hoy["Vlr_Mer_Ant"].sum())
    pl_d   = sf(hoy["PL"].sum())
    caus_d = sf(hoy["Causacion_Mer"].sum())
    caus_t = sf(hoy["Causacion_TIR"].sum())
    caus_m = sf(hoy["Causacion_Moneda"].sum())
    caus_s = sf(hoy["Causacion_Tasa"].sum())
    adeud  = sf(hoy["Adeudados"].sum())
    n_pos  = int(len(hoy))
    var_p  = sf(pl_d / ant * 100) if ant else 0

    mask  = hoy["TIR_Mercado"].notna() & (hoy["TIR_Mercado"] > 0.0001) & (hoy["TIR_Mercado"] < 30)
    tir_p = sf(
        (hoy.loc[mask,"Vlr_Mer_Hoy"] * hoy.loc[mask,"TIR_Mercado"]).sum()
        / hoy.loc[mask,"Vlr_Mer_Hoy"].sum()
    ) if mask.any() else None

    rf_mask = hoy["Tipo"].isin(["TES","CDT","TIDIS","Titularizaciones"]) & hoy["Dias_Vcto"].notna()
    dur_p   = sf(
        (hoy.loc[rf_mask,"Vlr_Mer_Hoy"] * hoy.loc[rf_mask,"Dias_Vcto"]).sum()
        / hoy.loc[rf_mask,"Vlr_Mer_Hoy"].sum()
    ) if rf_mask.any() else None

    fx_total = sf(hoy[hoy["Es_Ext"]]["Vlr_Mer_Hoy"].sum())
    fx_pct   = sf(fx_total / tot * 100) if (tot and fx_total) else 0

    # Estadísticas período
    pl_arr  = evol["pl"].values
    tot_arr = evol["total"].values

    def _dd(arr):
        pk = arr[0]; mx = 0
        for v in arr:
            if v > pk: pk = v
            d = (pk - v) / pk if pk > 0 else 0
            if d > mx: mx = d
        return mx * 100

    sharpe = sf(
        float(pl_arr.mean()) / float(pl_arr.std()) * np.sqrt(252)
        if pl_arr.std() > 0 else 0
    )

    stats = {
        "pl_max":     sf(float(pl_arr.max())),
        "pl_min":     sf(float(pl_arr.min())),
        "pl_avg":     sf(float(pl_arr.mean())),
        "pl_std":     sf(float(pl_arr.std())),
        "pl_acum":    sf(float(pl_arr.sum())),
        "dias_pos":   int((pl_arr > 0).sum()),
        "dias_neg":   int((pl_arr < 0).sum()),
        "dias_tot":   len(pl_arr),
        "sharpe":     sharpe,
        "drawdown":   sf(_dd(tot_arr)),
        "n_fechas":   len(fechas_ord),
        "n_esp":      int(m["Especie"].nunique()),
        "primer_dia": sd(fechas_ord[0]),
        "ultimo_dia": sd(fechas_ord[-1]),
        "var_periodo":sf((tot_arr[-1] - tot_arr[0]) / tot_arr[0] * 100) if tot_arr[0] else 0,
        "vol_anual":  sf(float(pl_arr.std()) * np.sqrt(252)),
        "hit_rate":   sf(float((pl_arr > 0).sum()) / len(pl_arr) * 100) if len(pl_arr) else 0,
    }

    # Insights
    insights = []
    idx_max = int(evol["pl"].idxmax())
    idx_min = int(evol["pl"].idxmin())
    insights.append({"tipo":"positive","txt":
        f"Mejor día: {sd(evol.iloc[idx_max]['Fecha_Posicion'])} → P&L ${evol.iloc[idx_max]['pl']:+,.0f}"})
    insights.append({"tipo":"negative","txt":
        f"Peor día: {sd(evol.iloc[idx_min]['Fecha_Posicion'])} → P&L ${evol.iloc[idx_min]['pl']:+,.0f}"})
    if stats["sharpe"]:
        tone = "positive" if stats["sharpe"] > 0.5 else "warning" if stats["sharpe"] > 0 else "negative"
        insights.append({"tipo":tone,"txt":f"Sharpe anualizado: {stats['sharpe']:.3f}"})
    if stats["drawdown"] and stats["drawdown"] > 0.5:
        insights.append({"tipo":"warning","txt":f"Máx. drawdown: {stats['drawdown']:.2f}%"})
    if fx_pct and fx_pct > 1:
        insights.append({"tipo":"info","txt":f"Exposición FX: {fx_pct:.1f}% del portafolio"})
    if tir_p:
        insights.append({"tipo":"info","txt":f"TIR ponderada RF: {tir_p:.2f}%"})
    vcto_90 = hoy[(hoy.get("Dias_Vcto", pd.Series(dtype=float)).fillna(999) < 90) &
                  (hoy.get("Dias_Vcto", pd.Series(dtype=float)).fillna(999) > 0)] \
        if "Dias_Vcto" in hoy.columns else pd.DataFrame()
    if len(vcto_90):
        insights.append({"tipo":"warning","txt":
            f"{len(vcto_90)} posiciones vencen en <90 días (${vcto_90['Vlr_Mer_Hoy'].sum():,.0f})"})
    pl_ac = stats["pl_acum"]
    if pl_ac is not None:
        tone = "positive" if pl_ac >= 0 else "negative"
        insights.append({"tipo":tone,"txt":f"P&L acumulado período: ${pl_ac:+,.0f}"})
    if stats["hit_rate"] is not None:
        insights.append({"tipo":"info","txt":
            f"Hit rate: {stats['hit_rate']:.1f}% días positivos de {stats['dias_tot']}"})

    # Tabla raw para caché incremental — guarda todas las columnas calculadas
    raw_tabla = [_raw_row(r) for _, r in m.iterrows()]

    return {
        "meta": {
            "generado":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "org":         CFG["org"],
            "sub":         CFG["sub"],
            "ultimo_dia":  sd(ult),
            "primer_dia":  sd(fechas_ord[0]),
            "n_fechas":    len(fechas_ord),
            "data_url":    "",   # se llena en main() tras detectar el remote
        },
        # ── Mapas editables — el dashboard los lee y permite modificarlos ──
        "mapas": {
            "ports":        CFG["ports"],
            "activos":      CFG["activos"],
            "monedas":      CFG["monedas"],
            "modalidades":  CFG["modalidades"],
            "estados":      CFG["estados"],
            "metodos":      CFG["metodos"],
            "tipo_reglas":  CFG["tipo_reglas"],
            "tipo_colores": CFG["tipo_colores"],
            "tipo_map":     CFG.get("tipo_map", {}),
        },
        "fechas":     [sd(d) for d in evol["Fecha_Posicion"]],
        "kpis": {
            "total":tot,"ant":ant,"pl":pl_d,"var_pct":var_p,
            "caus_mer":caus_d,"caus_tir":caus_t,
            "caus_mon":caus_m,"caus_tasa":caus_s,
            "adeudados":adeud,
            "n_pos":n_pos,"tir_pond":tir_p,
            "dur_pond":dur_p,"fx_total":fx_total,"fx_pct":fx_pct,
        },
        "stats":      stats,
        "insights":   insights,
        "evol": {
            "total":    [sf(v) for v in evol["total"]],
            "nominal":  [sf(v) for v in evol["nominal"]],
            "pl":       [sf(v) for v in evol["pl"]],
            "pl_acum":  [sf(v) for v in evol["pl_acum"]],
            "caus_mer": [sf(v) for v in evol["caus_mer"]],
            "caus_tir": [sf(v) for v in evol["caus_tir"]],
            "caus_mon": [sf(v) for v in evol["caus_mon"]],
            "caus_tasa":[sf(v) for v in evol["caus_tasa"]],
            "caus_acum":[sf(v) for v in evol["caus_acum"]],
            "rend_pct": [sf(v) for v in evol["rend_pct"]],
            "n":        [int(v) for v in evol["n"]],
        },
        "ports":      ports_u,
        "tipos":      tipos_u,
        "evol_port":  evol_port,
        "evol_tipo":  evol_tipo,
        "caus_hist":  caus_hist,
        "esp_hist":   esp_hist,
        "by_port":    by_port,
        "by_tipo":    by_tipo,
        "by_mon":     by_mon,
        "by_act":     by_act,
        "by_mod":     by_mod,
        "comp":       comp,
        "tabla":      tabla,
        "tes":        tes,
        "extras":     _build_extras(m, hoy, evol, fechas_ord),
        # Datos crudos para caché incremental (no se usan en el dashboard)
        "_raw_tabla": raw_tabla,
    }

# ══════════════════════════════════════════════════════════════════
#  HELPERS GIT
# ══════════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════════
#  SUBIDA A GITHUB VIA API — no necesita git instalado ni repo
#  clonado. Solo necesita el token configurado arriba.
# ══════════════════════════════════════════════════════════════════

def _github_api(method, endpoint, body=None):
    """Llama a la API de GitHub con urllib (sin dependencias externas)."""
    import urllib.request, urllib.error
    token = GITHUB_TOKEN.strip()
    if not token:
        raise ValueError("GITHUB_TOKEN no configurado.")
    url = f"https://api.github.com/{endpoint}"
    data = json.dumps(body).encode() if body else None
    req  = urllib.request.Request(url, data=data, method=method)
    req.add_header("Authorization", f"token {token}")
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req) as resp:
            return json.loads(resp.read()), resp.status
    except urllib.error.HTTPError as e:
        body_err = e.read().decode()
        raise RuntimeError(f"GitHub API {e.code}: {body_err}") from e

def github_push_file(local_path, repo_path, commit_msg):
    """
    Sube un archivo a GitHub via API.
    - local_path : ruta local del archivo a subir
    - repo_path  : ruta dentro del repo (ej: "data.json")
    - commit_msg : mensaje del commit
    Devuelve True si tuvo éxito.
    """
    import base64
    repo = GITHUB_REPO.strip()

    # Leer el archivo
    with open(local_path, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode()

    # Obtener el SHA actual del archivo (necesario para actualizar)
    sha = None
    try:
        info, _ = _github_api("GET", f"repos/{repo}/contents/{repo_path}")
        sha = info.get("sha")
    except RuntimeError as e:
        if "404" not in str(e):
            raise   # si es otro error, relanzar
        # 404 = archivo nuevo, no hay SHA

    # Preparar el body del commit
    body = {
        "message": commit_msg,
        "content": content_b64,
        "branch":  GITHUB_BRANCH,
    }
    if sha:
        body["sha"] = sha

    _github_api("PUT", f"repos/{repo}/contents/{repo_path}", body)
    return True

def verificar_token():
    """Verifica que el token de GitHub sea válido y tenga permisos."""
    token = GITHUB_TOKEN.strip()
    if not token:
        print("\n  [ERROR] GITHUB_TOKEN está vacío.")
        print("  Pasos:")
        print("  1. Ve a https://github.com/settings/tokens")
        print("  2. 'Generate new token (classic)'")
        print("  3. Marca el permiso: repo (Full control)")
        print("  4. Copia el token y pégalo en GITHUB_TOKEN en este script.")
        return False
    try:
        info, _ = _github_api("GET", f"repos/{GITHUB_REPO}")
        print(f"  Repositorio : {info['html_url']}")
        pages_url = f"https://{info['owner']['login']}.github.io/{info['name']}/"
        print(f"  Dashboard   : {pages_url}")
        return True
    except RuntimeError as e:
        if "401" in str(e):
            print("\n  [ERROR] Token inválido o expirado.")
            print("  Genera uno nuevo en: https://github.com/settings/tokens")
        elif "404" in str(e):
            print(f"\n  [ERROR] Repositorio '{GITHUB_REPO}' no encontrado.")
            print("  Verifica que GITHUB_REPO esté bien escrito.")
        else:
            print(f"\n  [ERROR] {e}")
        return False

# ══════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════

def main():
    repo        = GITHUB_REPO.strip()
    gh_user     = repo.split("/")[0] if "/" in repo else ""
    gh_repo_name= repo.split("/")[1] if "/" in repo else repo
    pages_url   = f"https://{gh_user}.github.io/{gh_repo_name}/" if gh_user else ""
    repo_url    = f"https://github.com/{repo}"

    print(f"\n  {'═'*56}")
    print(f"  ║  SAPIENZA · Procesador de Posiciones de Portafolio  ║")
    print(f"  {'═'*56}")
    print(f"\n  Repositorio : {repo_url}")
    print(f"  Dashboard   : {pages_url or '[configura GITHUB_REPO]'}")
    print(f"  Datos       : {ROOT}")

    # ── 0. Verificar token antes de procesar ─────────────────────
    print()
    if not verificar_token():
        raise SystemExit(1)

    # ── 1. Caché + CSV → XLSX ────────────────────────────────────
    cache = cargar_cache()
    print(f"\n  Caché       : {len(cache)} archivos registrados")
    print(f"\n[1/4] Convirtiendo CSV a XLSX...")
    convertir_csvs(cache)

    # ── 2. Cargar datos ──────────────────────────────────────────
    print("\n[2/4] Cargando datos (caché incremental)...")
    master, hay_nuevos = cargar(cache)

    if not hay_nuevos:
        print("\n  ✓ Sin archivos nuevos — data.json ya está actualizado.")
        print(f"\n  Dashboard: {pages_url}")
        return

    # ── 3. Construir data.json ───────────────────────────────────
    guardar_cache(cache)
    print("\n[3/4] Construyendo data.json...")
    data = build_json(master)
    data["meta"]["data_url"] = f"https://{gh_user}.github.io/{gh_repo_name}/data.json"

    json_out = CFG["output_json"]
    with open(json_out, "w", encoding="utf-8") as f:
        data_pub = {k: v for k, v in data.items() if k != "_raw_tabla"}
        json.dump(data_pub, f, ensure_ascii=False, separators=(",", ":"), default=str)
    kb = os.path.getsize(json_out) // 1024
    print(f"  data.json  : {kb} KB  ({data['meta']['n_fechas']} fechas)")

    k = data["kpis"]; s = data["stats"]
    print(f"\n  {'─'*48}")
    print(f"  Fecha posición   : {data['meta']['ultimo_dia']}")
    print(f"  Total portafolio : ${k['total']:>22,.0f}")
    print(f"  P&L del día      : ${k['pl']:>+22,.0f}  ({k['var_pct']:+.3f}%)")
    print(f"  P&L acumulado    : ${s['pl_acum']:>+22,.0f}")
    print(f"  Hit rate         : {s.get('hit_rate',0):.1f}%   Sharpe: {s['sharpe']}")
    print(f"  TIR ponderada    : {str(k['tir_pond'])+'%':>24}")
    print(f"  Exp. FX          : ${k['fx_total']:>22,.0f}  ({k['fx_pct']:.1f}%)")
    print(f"  {'─'*48}")

    # ── 4. Subir data.json via API de GitHub ─────────────────────
    print("\n[4/4] Subiendo data.json a GitHub...")
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M")
    msg   = CFG["git_msg"].format(fecha=fecha)
    try:
        github_push_file(json_out, "data.json", msg)
        print()
        print(f"  ╔{'═'*52}╗")
        print(f"  ║  ✓ Datos publicados correctamente               ║")
        print(f"  ╠{'═'*52}╣")
        print(f"  ║  Dashboard : {pages_url:<38} ║")
        print(f"  ║  Repo      : {repo_url:<38} ║")
        print(f"  ╚{'═'*52}╝")
        print(f"\n  El dashboard mostrará los datos nuevos en ~1 minuto.")
    except Exception as e:
        print(f"\n  [ERROR] No se pudo subir data.json: {e}")
        print(f"  El archivo quedó guardado localmente en: {json_out}")
    print("\n  Abriendo dashboard local…")


if __name__ == "__main__":
    main()
