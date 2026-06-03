# ╔══════════════════════════════════════════════════════════════════╗
# ║  SAPIENZA · Procesador de Posiciones de Portafolio              ║
# ║  Convierte CSV→XLSX · genera data.json + index.html · git push  ║
# ╚══════════════════════════════════════════════════════════════════╝
#
# COLUMNAS DEL SISTEMA FUENTE (referencia):
#   Especie          - Nombre del instrumento
#   Titulo/Inver     - Códigos internos del sistema de inversiones
#   F_Vcto           - Fecha de vencimiento del instrumento
#   Vlr_Nominal      - Valor nominal (face value) de la posición
#   Facial           - Tasa facial / referencia (ej: 7.5%, IBR3M+1.2%, BCL01+0F)
#   Mod              - Modalidad: AV=Año Vencido, DV=Día Vencido, Dto=Descuento,
#                      MV=Mes Vencido, TV=Trimestre Vencido, NAp=No Aplica
#   Desde/Hasta      - Período de causación (fecha inicio y fin del período actual)
#   Vlr_Mer_Ant      - Valor de mercado día ANTERIOR (en COP)
#   Vlr_Mer_Hoy      - Valor de mercado HOYA (en COP)
#   Adeudados        - Cupones o intereses adeudados pendientes de pago
#   Causacion_Mer    - Causación a precios de mercado (rendimiento devengado diario)
#   Causacion_TIR    - Causación a TIR (devengado teórico por TIR de compra)
#   ISIN_Nemot       - Código ISIN o nemotécnico del título
#   Met              - Método de valoración:
#                      QSI=Quantil (fondos/liquidez), QES-SI=Quantil renta fija,
#                      MC4-E=Modelo 4 Equity, MC1-I=Modelo 1 Interés
#   Precio           - Precio de mercado del título (% del nominal, ej: 101.376)
#   TIR_Mercado      - Tasa Interna de Retorno a precios de mercado (%)
#   Moneda           - Moneda de denominación (COP, USCOP=USD, EUO=EUR, etc.)
#   Mnd_Val_An       - Valor nominal en moneda nativa AYER (para calcular FX)
#   Mnd_Val          - Valor nominal en moneda nativa HOY
#   Dif_cambio       - Diferencial de tasa de cambio (Mnd_Val - Mnd_Val_An)
#   Causacion_Moneda - Parte de causación atribuible al efecto cambiario (FX)
#   Causacion_Tasa   - Parte de causación atribuible al efecto tasa de interés
#   Dias             - Días de la posición en el período actual
#   Por              - Código portafolio-activo (ej: "21-F" = Portafolio 21, Renta Fija)
#   Est              - Estado: Pend=Pendiente liquidación, Reci=Recibido/En cartera,
#                      Vend=Vendido, Frac=Fraccionado

import os, re, glob, json, subprocess, warnings
from datetime import datetime
from pathlib import Path

import pandas as pd
import numpy as np

warnings.filterwarnings("ignore")

# ┌─────────────────────────────────────────────────────────────────┐
# │  ★  CONFIGURACIÓN — solo editar aquí                            │
# └─────────────────────────────────────────────────────────────────┘
CFG = {
    # Carpeta raíz con CSV/XLSX (subcarpetas por mes son detectadas automáticamente)
    "carpeta_datos": r"C:\Users\danie\Sapienza\POSPRO",

    # Archivos generados (se sobreescriben en cada ejecución)
    "output_html":   r"C:\Users\danie\Sapienza\POSPRO\index.html",
    "output_json":   r"C:\Users\danie\Sapienza\POSPRO\data.json",

    # Git: commit + push automático al terminar
    "git_push": True,
    "git_msg":  "data: posiciones {fecha}",

    # Archivos que NO son datos fuente (se ignoran al buscar)
    "ignorar": {
        "posiciones_consolidadas.xlsx", "index.html",
        "dashboard.html", "data.json", "_dashboard_tpl.html",
    },

    # Identidad del dashboard
    "org":  "Skandia Colombia",
    "sub":  "Dashboard de Posiciones de Portafolio",
}
# ──────────────────────────────────────────────────────────────────

ROOT = CFG["carpeta_datos"]

# Columnas del archivo fuente (en orden exacto del CSV)
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

# Mapas de códigos → nombres legibles
PORTS = {
    "21":"Portafolio 21","41":"Portafolio 41","51":"Portafolio 51",
    "HC":"HC Cesantías","HO":"HO Obligatorio","HE":"HE Especial",
}
ACTIVOS = {
    "F":"Renta Fija (TES)","H":"Liquidez","E":"Fondos Colectivos",
    "O":"Otros Títulos","D":"Depósitos/Cash","Y":"Fondos SPC",
    "V":"AOR","L":"Acciones","G":"Fondos","DF":"Derivados",
    "T":"TIDIS","DB":"Depósitos Banco","HB":"Liquidez HB",
    "W":"AOR Propios","P":"Otros",
}
MONEDAS = {
    "COP":"COP","USCOP":"USD","EUO":"EUR","UKCOP":"GBP",
    "ETQACOP":"Fdo ETQA","ETRACOP":"Fdo ETRA",
    "ESJWX":"Fdo Alt A","ESY0O":"Fdo UnoMas",
    "EIBMS":"Fdo Efect A","EIEMS":"Fdo Efect D",
    "EABFR":"Fdo Occirenta","ETGWX":"Fdo Alt D",
    "EDXR5":"SPC Corto Plazo","EDY1U":"SPC Largo Plazo",
    "EDT1U":"SPC Conservador","EDU1U":"SPC Mayor Riesgo",
    "EDSR6":"SPC Moderado","EDV1U":"SPC Retiro Programado",
    "ESUMS":"Fdo Umas",
}
MOD_DESC = {
    "AV":"Año Vencido","DV":"Día Vencido","Dto":"Descuento",
    "MV":"Mes Vencido","TV":"Trimestre Vencido","NAp":"No Aplica",
}
EST_DESC = {
    "Pend":"Pendiente","Reci":"En Cartera","Vend":"Vendido",
    "Frac":"Fraccionado",
}
MET_DESC = {
    "QSI":"Quantil (Fondos)","QES-SI":"Quantil Renta Fija",
    "MC4-E":"Modelo Equity","MC1-I":"Modelo Interés",
}

# ── Funciones de limpieza ─────────────────────────────────────────

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
    """Clasifica el instrumento por nombre de especie."""
    u = str(e).upper()
    if "TES"        in u:                                   return "TES"
    if any(x in u for x in ["CDT","CREDITO","CRÉDITO"]):   return "CDT"
    if any(x in u for x in ["CASH","CTA AHO","CTA AH","CTA "]): return "Liquidez"
    if any(x in u for x in ["FIC","FCPD","FCPE","FCP","P SPC"]): return "Fondos"
    if any(x in u for x in ["DER.","DERIV"]):              return "Derivados"
    if "AOR"        in u:                                   return "AOR"
    if "TIDIS"      in u:                                   return "TIDIS"
    if "TITULARIZ"  in u:                                   return "Titularizaciones"
    if "ACC."       in u:                                   return "Acciones"
    return "Otro"

# ── Lectura de archivos ───────────────────────────────────────────

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

# ── Conversión CSV → XLSX ─────────────────────────────────────────

def convertir_csvs():
    """
    Convierte todos los CSV de la carpeta (y subcarpetas) a XLSX
    en la misma ubicación. Si el XLSX ya existe y es más reciente, lo salta.
    """
    csvs = list({os.path.normcase(p): p for p in
        glob.glob(os.path.join(ROOT, "**", "*.CSV"), recursive=True) +
        glob.glob(os.path.join(ROOT, "**", "*.csv"), recursive=True)
    }.values())

    convertidos = 0
    for csv_path in sorted(csvs):
        xlsx_path = os.path.splitext(csv_path)[0] + ".xlsx"
        # Saltar si XLSX ya existe y es más reciente que el CSV
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

# ── Transformación ────────────────────────────────────────────────

def _transformar(df, fecha):
    df = df.copy()

    # Limpiar strings
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].str.strip()

    # Convertir numéricos
    for c in COLS_NUM:
        if c in df.columns:
            df[c] = df[c].apply(_num)

    # Convertir fechas
    for c in COLS_DT:
        if c in df.columns:
            df[c] = df[c].apply(_dt)

    # Campos calculados base
    df["Fecha_Posicion"] = pd.to_datetime(fecha)
    df["Especie"]        = df["Especie"].str.strip()
    df["ISIN_Nemot"]     = df["ISIN_Nemot"].str.strip().replace("", np.nan)
    df["Est"]            = df["Est"].str.strip()
    df["Mod"]            = df["Mod"].str.strip()
    df["Met"]            = df["Met"].str.strip()
    df["Facial"]         = df["Facial"].str.strip()

    # Clasificación de instrumento
    df["Tipo"] = df["Especie"].apply(_tipo)

    # Portafolio y tipo de activo desde código "Por"
    df["Por"]      = df["Por"].str.strip()
    df["Port_Cod"] = df["Por"].str.extract(r"^(\d+|H[COEB]+)", expand=False)
    df["Act_Cod"]  = df["Por"].str.extract(r"-(\w+)$",         expand=False)
    df["Port_Nom"] = df["Port_Cod"].map(PORTS).fillna(df["Port_Cod"])
    df["Act_Nom"]  = df["Act_Cod"].map(ACTIVOS).fillna(df["Act_Cod"])

    # Moneda
    df["Moneda"]   = df["Moneda"].str.replace("$", "COP", regex=False)
    df["Mon_Desc"] = df["Moneda"].map(MONEDAS).fillna(df["Moneda"])
    df["Es_Ext"]   = df["Moneda"].isin(["USCOP","EUO","UKCOP","USD","EUR","GBP"])

    # P&L y rendimiento
    df["PL"]       = df["Vlr_Mer_Hoy"] - df["Vlr_Mer_Ant"]
    df["Rend_Pct"] = np.where(
        df["Vlr_Mer_Ant"] != 0,
        df["PL"] / df["Vlr_Mer_Ant"] * 100,
        np.nan
    )

    # Descomposición de causación
    # Causacion_Mer   = total causación a mercado
    # Causacion_TIR   = causación teórica a TIR de compra
    # Causacion_Tasa  = componente por movimiento de tasas
    # Causacion_Moneda= componente por movimiento de FX
    df["Caus_Spread"] = df["Causacion_Mer"] - df["Causacion_TIR"]  # diferencia mercado vs TIR

    # Días al vencimiento
    if "F_Vcto" in df.columns:
        df["Dias_Vcto"] = (df["F_Vcto"] - df["Fecha_Posicion"]).dt.days

    # Descripción de modalidad y estado
    df["Mod_Desc"] = df["Mod"].map(MOD_DESC).fillna(df["Mod"])
    df["Est_Desc"] = df["Est"].map(EST_DESC).fillna(df["Est"])
    df["Met_Desc"] = df["Met"].map(MET_DESC).fillna(df["Met"])

    return df

# ── Carga completa ────────────────────────────────────────────────

def cargar():
    ign = CFG["ignorar"]
    xlsxs = [p for p in
        glob.glob(os.path.join(ROOT, "**", "*.xlsx"), recursive=True) +
        glob.glob(os.path.join(ROOT, "**", "*.xls"),  recursive=True)
        if os.path.basename(p) not in ign
    ]
    # Deduplicar por ruta normalizada
    xlsxs = list({os.path.normcase(p): p for p in xlsxs}.values())

    todos = []
    print(f"XLSX: {len(xlsxs)} archivos encontrados")
    for p in sorted(xlsxs):
        f  = _fecha_path(p)
        df = _leer_xlsx(p)
        if df is not None and not df.empty:
            df = _transformar(df, f)
            df["_src"] = os.path.basename(p)
            todos.append(df)
            print(f"  {os.path.relpath(p, ROOT)} -> {len(df)} filas | fecha={f}")

    if not todos:
        raise SystemExit("[ERROR] No se encontraron archivos de datos.")

    master = (pd.concat(todos, ignore_index=True)
              .sort_values(["Fecha_Posicion","Especie"], ignore_index=True))

    print(f"\nTotal: {len(master):,} registros | "
          f"{master['Fecha_Posicion'].nunique()} fechas | "
          f"{master['Especie'].nunique()} especies")
    return master

# ── Helpers de serialización ──────────────────────────────────────

def sf(v):
    """float seguro para JSON"""
    if v is None: return None
    try:
        f = float(v)
        return None if (np.isnan(f) or np.isinf(f)) else round(f, 4)
    except: return None

def sd(v):
    """datetime → "YYYY-MM-DD" """
    try: return pd.Timestamp(v).strftime("%Y-%m-%d") if pd.notna(v) else None
    except: return None

def rows(df):
    """DataFrame → lista de dicts con valores seguros para JSON"""
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

# ── Construcción del JSON ─────────────────────────────────────────

def build_json(m):
    fechas_ord = sorted(m["Fecha_Posicion"].unique())
    ult = fechas_ord[-1]
    hoy = m[m["Fecha_Posicion"] == ult]

    # ── 1. Series temporales completas ───────────────────────────
    evol = (m.groupby("Fecha_Posicion").agg(
        total=("Vlr_Mer_Hoy","sum"),
        pl=("PL","sum"),
        caus_mer=("Causacion_Mer","sum"),
        caus_tir=("Causacion_TIR","sum"),
        caus_mon=("Causacion_Moneda","sum"),
        caus_tasa=("Causacion_Tasa","sum"),
        adeudados=("Adeudados","sum"),
        n=("Especie","count"),
    ).reset_index().sort_values("Fecha_Posicion"))
    evol["pl_acum"]      = evol["pl"].cumsum()
    evol["caus_acum"]    = evol["caus_mer"].cumsum()
    evol["rend_pct"]     = evol["pl"] / evol["total"].shift(1).replace(0, np.nan) * 100

    # ── 2. Por portafolio × fecha ─────────────────────────────────
    ep = m.groupby(["Fecha_Posicion","Port_Nom"]).agg(
        total=("Vlr_Mer_Hoy","sum"), pl=("PL","sum"),
        caus=("Causacion_Mer","sum"), n=("Especie","count")
    ).reset_index()
    ports_u = sorted(m["Port_Nom"].dropna().unique().tolist())
    evol_port = {}
    for p in ports_u:
        s = ep[ep["Port_Nom"]==p].sort_values("Fecha_Posicion")
        evol_port[p] = {
            "total": [sf(v) for v in s["total"]],
            "pl":    [sf(v) for v in s["pl"]],
            "caus":  [sf(v) for v in s["caus"]],
        }

    # ── 3. Por tipo × fecha ───────────────────────────────────────
    et = m.groupby(["Fecha_Posicion","Tipo"]).agg(
        total=("Vlr_Mer_Hoy","sum"), pl=("PL","sum")
    ).reset_index()
    tipos_u = sorted(m["Tipo"].unique().tolist())
    evol_tipo = {}
    for t in tipos_u:
        s = et[et["Tipo"]==t].sort_values("Fecha_Posicion")
        evol_tipo[t] = [sf(v) for v in s["total"]]

    # ── 4. Historia completa por especie ──────────────────────────
    esp_hist = {}
    for esp, grp in m.groupby("Especie"):
        g = grp.sort_values("Fecha_Posicion")
        esp_hist[esp] = {
            "fechas":  [sd(d) for d in g["Fecha_Posicion"]],
            "total":   [sf(v) for v in g["Vlr_Mer_Hoy"]],
            "pl":      [sf(v) for v in g["PL"]],
            "caus":    [sf(v) for v in g["Causacion_Mer"]],
            "caus_tir":[sf(v) for v in g["Causacion_TIR"]],
            "caus_mon":[sf(v) for v in g["Causacion_Moneda"]],
            "caus_tasa":[sf(v) for v in g["Causacion_Tasa"]],
            "tir":     [sf(v) for v in g["TIR_Mercado"]],
            "precio":  [sf(v) for v in g["Precio"]],
            "port":    str(g["Port_Nom"].iloc[-1]) if pd.notna(g["Port_Nom"].iloc[-1]) else "",
            "tipo":    str(g["Tipo"].iloc[-1]),
        }

    # ── 5. Último día — tabla detalle completa ───────────────────
    def _det(r):
        return {
            "esp":       str(r["Especie"]),
            "port":      str(r["Port_Nom"])  if pd.notna(r["Port_Nom"])  else "",
            "act":       str(r["Act_Nom"])   if pd.notna(r["Act_Nom"])   else "",
            "tipo":      str(r["Tipo"]),
            "isin":      str(r["ISIN_Nemot"])if pd.notna(r["ISIN_Nemot"])else "",
            "nominal":   sf(r["Vlr_Nominal"]),
            "v_ant":     sf(r["Vlr_Mer_Ant"]),
            "valor":     sf(r["Vlr_Mer_Hoy"]),
            "pl":        sf(r["PL"]),
            "rend":      sf(r["Rend_Pct"]),
            # Causaciones descompuestas
            "caus_mer":  sf(r["Causacion_Mer"]),
            "caus_tir":  sf(r["Causacion_TIR"]),
            "caus_mon":  sf(r["Causacion_Moneda"]),
            "caus_tasa": sf(r["Causacion_Tasa"]),
            "caus_diff": sf(r.get("Caus_Spread", 0)),
            "adeudados": sf(r["Adeudados"]),
            # Precio y TIR
            "tir":       sf(r["TIR_Mercado"]),
            "precio":    sf(r["Precio"]),
            # Moneda / FX
            "moneda":    str(r["Mon_Desc"]),
            "ext":       bool(r["Es_Ext"]),
            "mnd_hoy":   sf(r["Mnd_Val"]),
            "mnd_ant":   sf(r["Mnd_Val_An"]),
            "dif_fx":    sf(r["Dif_cambio"]),
            # Vencimiento
            "vcto":      sd(r.get("F_Vcto")),
            "dias_vcto": sf(r.get("Dias_Vcto")),
            # Descriptivos
            "facial":    str(r["Facial"])   if pd.notna(r["Facial"])   else "",
            "mod":       str(r["Mod_Desc"]) if pd.notna(r["Mod_Desc"]) else "",
            "est":       str(r["Est_Desc"]) if pd.notna(r["Est_Desc"]) else "",
            "met":       str(r["Met_Desc"]) if pd.notna(r["Met_Desc"]) else "",
            "por":       str(r["Por"]),
            "desde":     sd(r.get("Desde")),
            "hasta":     sd(r.get("Hasta")),
            "dias":      sf(r["Dias"]),
        }

    tabla = [_det(r) for _, r in hoy.sort_values("Vlr_Mer_Hoy", ascending=False).iterrows()]

    # Renta fija por separado
    tes  = [_det(r) for _, r in
            hoy[hoy["Tipo"].isin(["TES","TIDIS","Titularizaciones","CDT"])]
            .sort_values("Vlr_Mer_Hoy", ascending=False).iterrows()]

    # ── 6. Resúmenes ─────────────────────────────────────────────

    def _gagg(grp_col, extra_aggs=None):
        agg = {"total":("Vlr_Mer_Hoy","sum"),"pl":("PL","sum"),
               "caus":("Causacion_Mer","sum"),"caus_tir":("Causacion_TIR","sum"),
               "caus_mon":("Causacion_Moneda","sum"),"caus_tasa":("Causacion_Tasa","sum"),
               "n":("Especie","count")}
        if extra_aggs: agg.update(extra_aggs)
        df_g = (hoy.groupby(grp_col).agg(**{k:v for k,v in agg.items()})
                .reset_index().sort_values("total", ascending=False))
        df_g["pct"] = df_g["total"] / df_g["total"].sum() * 100
        return rows(df_g)

    by_port = _gagg("Port_Nom")
    by_tipo = _gagg("Tipo")
    by_mon  = _gagg("Mon_Desc")
    by_act  = _gagg("Act_Nom")
    by_mod  = _gagg("Mod")

    # Causación descompuesta por portafolio (histórico)
    caus_hist = {}
    for p in ports_u:
        sub = m[m["Port_Nom"]==p].groupby("Fecha_Posicion").agg(
            caus_mer=("Causacion_Mer","sum"),
            caus_tir=("Causacion_TIR","sum"),
            caus_mon=("Causacion_Moneda","sum"),
            caus_tasa=("Causacion_Tasa","sum"),
        ).reset_index().sort_values("Fecha_Posicion")
        caus_hist[p] = {
            "mer":  [sf(v) for v in sub["caus_mer"]],
            "tir":  [sf(v) for v in sub["caus_tir"]],
            "mon":  [sf(v) for v in sub["caus_mon"]],
            "tasa": [sf(v) for v in sub["caus_tasa"]],
        }

    # Composición por portafolio × tipo (último día)
    comp = rows(hoy.groupby(["Port_Nom","Tipo"])["Vlr_Mer_Hoy"].sum().reset_index())

    # ── 7. KPIs último día ────────────────────────────────────────
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

    # TIR ponderada (solo instrumentos con TIR significativa)
    mask   = hoy["TIR_Mercado"].notna() & (hoy["TIR_Mercado"] > 0.0001) & (hoy["TIR_Mercado"] < 30)
    tir_p  = sf(
        (hoy.loc[mask,"Vlr_Mer_Hoy"] * hoy.loc[mask,"TIR_Mercado"]).sum()
        / hoy.loc[mask,"Vlr_Mer_Hoy"].sum()
    ) if mask.any() else None

    # Duración promedio ponderada (solo RF)
    rf_mask = hoy["Tipo"].isin(["TES","CDT","TIDIS","Titularizaciones"]) & hoy["Dias_Vcto"].notna()
    dur_p   = sf(
        (hoy.loc[rf_mask,"Vlr_Mer_Hoy"] * hoy.loc[rf_mask,"Dias_Vcto"]).sum()
        / hoy.loc[rf_mask,"Vlr_Mer_Hoy"].sum()
    ) if rf_mask.any() else None

    # Exposición FX total
    fx_total = sf(hoy[hoy["Es_Ext"]]["Vlr_Mer_Hoy"].sum())
    fx_pct   = sf(fx_total / tot * 100) if (tot and fx_total) else 0

    # ── 8. Estadísticas del período ───────────────────────────────
    pl_arr = evol["pl"].values
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
        "pl_max":   sf(float(pl_arr.max())),
        "pl_min":   sf(float(pl_arr.min())),
        "pl_avg":   sf(float(pl_arr.mean())),
        "pl_std":   sf(float(pl_arr.std())),
        "pl_acum":  sf(float(pl_arr.sum())),
        "dias_pos": int((pl_arr > 0).sum()),
        "dias_neg": int((pl_arr < 0).sum()),
        "dias_tot": len(pl_arr),
        "sharpe":   sharpe,
        "drawdown": sf(_dd(tot_arr)),
        "n_fechas": len(fechas_ord),
        "n_esp":    int(m["Especie"].nunique()),
        "primer_dia": sd(fechas_ord[0]),
        "ultimo_dia": sd(fechas_ord[-1]),
        "var_periodo": sf((tot_arr[-1] - tot_arr[0]) / tot_arr[0] * 100) if tot_arr[0] else 0,
    }

    # ── 9. Insights automáticos ───────────────────────────────────
    insights = []
    idx_max = int(evol["pl"].idxmax())
    idx_min = int(evol["pl"].idxmin())
    insights.append({"tipo":"positive","txt":
        f"Mejor día: {sd(evol.iloc[idx_max]['Fecha_Posicion'])} → P&L ${evol.iloc[idx_max]['pl']:+,.0f}"})
    insights.append({"tipo":"negative","txt":
        f"Peor día: {sd(evol.iloc[idx_min]['Fecha_Posicion'])} → P&L ${evol.iloc[idx_min]['pl']:+,.0f}"})
    if stats["sharpe"]:
        tone = "positive" if stats["sharpe"] > 0.5 else "warning" if stats["sharpe"] > 0 else "negative"
        insights.append({"tipo":tone,"txt":f"Sharpe anualizado del período: {stats['sharpe']:.3f}"})
    if stats["drawdown"] and stats["drawdown"] > 0.5:
        insights.append({"tipo":"warning","txt":f"Máximo drawdown del período: {stats['drawdown']:.2f}%"})
    if fx_pct and fx_pct > 1:
        insights.append({"tipo":"info","txt":f"Exposición en moneda extranjera: {fx_pct:.1f}% del portafolio"})
    if tir_p:
        insights.append({"tipo":"info","txt":f"TIR ponderada de renta fija: {tir_p:.2f}%"})
    vcto_90 = hoy[(hoy.get("Dias_Vcto", pd.Series(dtype=float)).fillna(999) < 90) &
                  (hoy.get("Dias_Vcto", pd.Series(dtype=float)).fillna(999) > 0)] \
        if "Dias_Vcto" in hoy.columns else pd.DataFrame()
    if len(vcto_90):
        v90_val = vcto_90["Vlr_Mer_Hoy"].sum()
        insights.append({"tipo":"warning","txt":
            f"{len(vcto_90)} posiciones vencen <90 días (${v90_val:,.0f})"})
    if stats["pl_acum"] and stats["pl_acum"] < 0:
        insights.append({"tipo":"negative","txt":
            f"P&L acumulado del período negativo: ${stats['pl_acum']:+,.0f}"})
    else:
        insights.append({"tipo":"positive","txt":
            f"P&L acumulado del período: ${stats['pl_acum']:+,.0f}"})

    return {
        "meta": {
            "generado":   datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "org":        CFG["org"],
            "sub":        CFG["sub"],
            "ultimo_dia": sd(ult),
            "primer_dia": sd(fechas_ord[0]),
        },
        "fechas":    [sd(d) for d in evol["Fecha_Posicion"]],
        "kpis": {
            "total":tot,"ant":ant,"pl":pl_d,"var_pct":var_p,
            "caus_mer":caus_d,"caus_tir":caus_t,
            "caus_mon":caus_m,"caus_tasa":caus_s,
            "adeudados":adeud,
            "n_pos":n_pos,"tir_pond":tir_p,
            "dur_pond":dur_p,"fx_total":fx_total,"fx_pct":fx_pct,
        },
        "stats":     stats,
        "insights":  insights,
        "evol": {
            "total":    [sf(v) for v in evol["total"]],
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
        "ports":     ports_u,
        "tipos":     tipos_u,
        "evol_port": evol_port,
        "evol_tipo": evol_tipo,
        "caus_hist": caus_hist,
        "esp_hist":  esp_hist,
        "by_port":   by_port,
        "by_tipo":   by_tipo,
        "by_mon":    by_mon,
        "by_act":    by_act,
        "by_mod":    by_mod,
        "comp":      comp,
        "tabla":     tabla,
        "tes":       tes,
    }

# ── Git push ──────────────────────────────────────────────────────

def git_push(files):
    repo  = ROOT
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M")
    msg   = CFG["git_msg"].format(fecha=fecha)
    for f in files:
        subprocess.run(["git","-C",repo,"add",os.path.relpath(f, repo)],
                       capture_output=True)
    r = subprocess.run(["git","-C",repo,"commit","-m",msg], capture_output=True, text=True)
    ok = r.returncode == 0 or "nothing to commit" in r.stdout + r.stderr
    print(f"  git commit: {'OK' if ok else r.stderr.strip()}")
    r = subprocess.run(["git","-C",repo,"push"], capture_output=True, text=True)
    print(f"  git push:   {'OK' if r.returncode==0 else r.stderr.strip()}")

# ── Main ──────────────────────────────────────────────────────────

def main():
    print("=" * 64)
    print(" SAPIENZA — Procesador de Posiciones de Portafolio")
    print(f" Carpeta datos: {ROOT}")
    print("=" * 64 + "\n")

    # 1. Convertir CSV → XLSX
    print("[1/4] Convirtiendo CSV a XLSX...")
    convertir_csvs()

    # 2. Cargar todos los XLSX
    print("\n[2/4] Cargando datos...")
    master = cargar()

    # 3. Construir JSON
    print("\n[3/4] Construyendo JSON de datos...")
    data = build_json(master)

    json_out = CFG["output_json"]
    with open(json_out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",",":"), default=str)
    print(f"  data.json: {os.path.getsize(json_out)//1024} KB")

    # 4. Generar HTML
    print("\n[4/4] Generando index.html...")
    tpl_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_dashboard_tpl.html")
    if not os.path.exists(tpl_path):
        raise SystemExit(f"[ERROR] No se encuentra la plantilla: {tpl_path}")
    with open(tpl_path, "r", encoding="utf-8") as f:
        tpl = f.read()

    data_str = json.dumps(data, ensure_ascii=False, separators=(",",":"), default=str)
    html = tpl.replace("__DATA_JSON__", data_str).replace("__ORG__", CFG["org"])
    html_out = CFG["output_html"]
    with open(html_out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  index.html: {os.path.getsize(html_out)//1024} KB")

    # Resumen en consola
    k = data["kpis"]; s = data["stats"]
    print(f"\n{'='*64}")
    print(f"  Fecha posicion   : {data['meta']['ultimo_dia']}")
    print(f"  Total portafolio : ${k['total']:>22,.0f}")
    print(f"  P&L del dia      : ${k['pl']:>+22,.0f}  ({k['var_pct']:+.3f}%)")
    print(f"  P&L acumulado    : ${s['pl_acum']:>+22,.0f}")
    print(f"  Causacion (mer)  : ${k['caus_mer']:>22,.0f}")
    print(f"  Causacion (TIR)  : ${k['caus_tir']:>22,.0f}")
    print(f"  TIR ponderada    :  {str(k['tir_pond'])+'%':>23}")
    print(f"  Drawdown max     :  {str(s['drawdown'])+'%':>23}")
    print(f"  Sharpe anualiz.  :  {str(s['sharpe']):>23}")
    print(f"  Exp. FX          : ${k['fx_total']:>22,.0f}  ({k['fx_pct']:.1f}%)")
    print(f"{'='*64}")

    # Git push
    if CFG["git_push"]:
        print("\n[Git] Subiendo a GitHub...")
        git_push([html_out, json_out])
        r = subprocess.run(["git","-C",ROOT,"remote","get-url","origin"],
                           capture_output=True, text=True)
        remote = r.stdout.strip()
        if "github.com" in remote:
            m2 = re.search(r"github\.com[:/]([^/]+)/([^/.]+)", remote)
            if m2:
                user, repo = m2.group(1), m2.group(2)
                print(f"\n  URL del dashboard:")
                print(f"  https://{user}.github.io/{repo}/")

    import webbrowser
    webbrowser.open(html_out)


if __name__ == "__main__":
    main()
