"""
Picking Orchestrator v4.0 — Beccacece Hnos SA
Streamlit unificado para automatización de picking (DPO 2.1 — Pilar Almacén)

Sucesor de app_v3_9_1.py (Planilla de Carga). No reemplaza la v3.9: se ejecuta
en paralelo. Reutiliza literalmente las funciones de la v3.9 para Tabs 1 y 2.

SPRINT 1 (este archivo):
  - Esqueleto con 6 tabs + sidebar (toggles dry-run y snapshot)
  - Tab 1: Planilla de Carga (reuso 100% v3.9 — generate_pdf)
  - Tab 2: Resumen por Camión (reuso 100% v3.9 — build_resumen_carga_pdf)
  - Tab 5: Extraíbles Sheets (Matriz Pall., Agregados AE, Reposición AE filtrada J<0, fx Picking)
  - Tabs 3, 4, 6: placeholders con scope documentado (Sprints 2-5)
  - Snapshot local provisional ./snapshots/YYYY-MM-DD/<run_id>/

ARQUITECTURA:
  - 1 solo archivo .py autocontenido (decisión Lucas)
  - Stack: streamlit, pandas, openpyxl, reportlab, loguru
  - MASTER (78 MB) leído con openpyxl(data_only=True, read_only=True)

EJECUCIÓN:
  pip install streamlit pandas openpyxl reportlab loguru
  streamlit run app_v4_0.py
"""
import io, math, hashlib, unicodedata, os, shutil, json
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import openpyxl
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as rl_canvas
from loguru import logger

# ─── VERSIÓN Y CONFIG GLOBAL ────────────────────────────────────────────────
APP_VERSION = "4.0.0-sprint1"
SNAPSHOT_DIR = Path("./snapshots")

# Colores T2 (Sprint 3)
T2_COLORS = {
    "verde": "#22c55e",     # Pallet pura completa
    "amarillo": "#eab308",  # Pallet picking (parcial)
    "gris": "#9ca3af",      # Vacío
    "rojo": "#ef4444",      # Alerta (excedido)
}

# Setup logger único (no duplicar handlers en re-runs Streamlit)
if "logger_configured" not in st.session_state:
    logger.remove()
    logger.add(
        lambda m: None,  # Stdout silenciado en Streamlit (lo capturamos en sesión)
        level="INFO",
        format="{time:HH:mm:ss} | {level} | {message}",
    )
    st.session_state["logger_configured"] = True
    st.session_state["log_buffer"] = []

def log_event(level: str, msg: str):
    """Log a archivo (futuro) y a buffer de sesión (para mostrar en UI)."""
    ts = datetime.now().strftime("%H:%M:%S")
    entry = f"[{ts}] {level.upper()}: {msg}"
    st.session_state.setdefault("log_buffer", []).append(entry)
    logger.log(level.upper(), msg)

# ============================================================================
# ║              BLOQUE REUSADO LITERAL DE app_v3_9_1.py                     ║
# ║                                                                          ║
# ║  Funciones críticas para Tabs 1 y 2:                                     ║
# ║   - load_car, load_frescura, load_agr                                    ║
# ║   - generate_pdf (Planilla de Carga)                                     ║
# ║   - build_resumen_carga_pdf (Resumen por Camión)                         ║
# ║   + todas las helpers de dibujo del PDF (draw_*, header_*, etc)          ║
# ║                                                                          ║
# ║  ⚠ NO MODIFICAR — lógica validada en producción.                         ║
# ============================================================================

import io, math, hashlib, unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as rl_canvas

# ─── MEDIDAS FIJAS (puntos) ──────────────────────────────────────────────────
PAGE_W, PAGE_H = A4
MARGIN  = 12 * mm
CW      = PAGE_W - 2 * MARGIN

H_TITLE     = 26
H_TRANS     = 18
H_PARTIDA   = 16
H_CTRL      = 56
H_LEMA      = 18
H_CANCHA    = 20
H_THDR      = 14
H_ALM       = 16
H_ROW       = 14
H_TOT       = 14
H_ESPACIO   = 62
H_ROUTE     = 16
H_FOOT      = 28
GAP         = 2

# ─── COLORES ────────────────────────────────────────────────────────────────
DARK_BLUE  = colors.HexColor('#1a3a6b')
MED_BLUE   = colors.HexColor('#2e5fa3')
LIGHT_GRAY = colors.HexColor('#EFEFEF')
BG_GRAY    = colors.HexColor('#F4F4F4')
DARK_ROW   = colors.HexColor('#AAAAAA')
MED_ROW    = colors.HexColor('#D8D8D8')
AMBER      = colors.HexColor('#FFE0B2')
RED_ALERT  = colors.HexColor('#C0392B')
FOOT_GRAY  = colors.HexColor('#555555')
BORDER     = colors.HexColor('#AAAAAA')
HDR_BG     = colors.HexColor('#E8E8E8')
YELLOW_RTE = colors.HexColor('#FFF3A0')
WM_GRAY    = colors.HexColor('#888888')
ALM_BG     = colors.HexColor('#DDEEFF')

EXCLUDED_SKUS = {2730, 2731, 2776, 2780, 5192}
EXCLUDED_PATS = ['Q CERVEZAS', 'Q PLAS', 'BOT 1/1', 'BOT AMBAR', 'ARACELI']
CANCHA_ORDER  = ['CANCHA I', 'CANCHA II', 'CANCHA III', 'CANCHA IV', 'CANCHA V']
COL_W = [50, CW - 50 - 62 - 52 - 52, 62, 52, 52]

LEMAS = [
    "Cada bulto bien puesto es una entrega perfecta. Dale con todo!",
    "El picking preciso empieza con vos. Conta, verifica, confia.",
    "Hoy es dia de hacer historia en el deposito. A romperla!",
    "Rapido no es apurado. Rapido es preciso y sin errores.",
    "El mejor pickero no es el mas veloz, es el mas exacto.",
    "Tu equipo cuenta con vos. Cada bulto importa.",
    "Zapatos abrochados, espalda cuidada, mente enfocada. Arrancamos!",
    "Paleta perfecta = cliente feliz = empresa fuerte. Todo empieza aca.",
    "Cero errores de picking es el estandar. Vos podes.",
    "Si dudas del SKU, verifica. Un segundo de control evita una hora de correccion.",
    "La diferencia entre bueno y excelente esta en los detalles. Se excelente.",
    "Arrancar bien es terminar mejor. Que empiece el picking!",
    "Trabajo limpio, paleta ordenada, conciencia tranquila.",
    "Sos parte del corazon de la operacion. El deposito te necesita al 100%.",
    "Pensa antes de mover. El orden en la paleta es orden en la entrega.",
    "Tecnica correcta al levantar = cuerpo sano manana tambien.",
    "Verifica el SKU, verifica la cantidad, verifica el lote. Tres checks, cero errores!",
    "Cada noche de picking sin errores construye la reputacion de todos.",
    "El deposito funciona porque vos funcionas. Gracias por el compromiso.",
    "Paleta bien armada, camion bien cargado, cliente bien atendido.",
    "El cansancio es real. El orgullo del trabajo bien hecho tambien.",
    "Reporta lo que esta mal antes de moverlo. Un reporte a tiempo vale oro.",
    "Somos un equipo. Si uno falla, fallamos todos. Si uno gana, ganamos todos.",
    "Segui la planilla al pie de la letra. Esta hecha para que llegues sin errores.",
    "Esta noche vas a hacer 500 movimientos correctos. Demostralo!",
    "Seguridad primero: EPI puesto, pasillo libre, mente clara.",
    "Los mejores pickeros no nacen, se hacen en noches como esta.",
    "Picking nocturno = concentracion maxima. El dia descansa, vos construis.",
    "Cada producto en su lugar es una promesa cumplida al cliente.",
    "Ultima caja, mismo cuidado que la primera. Asi se hace el trabajo bien hecho.",
]

# ─── UTILIDADES ─────────────────────────────────────────────────────────────

def _norm(s):
    """Normaliza string: sin tildes, sin espacios extra, lowercase."""
    if s is None: return ''
    s = str(s).strip()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    return s.lower()

def find_col(df, *candidates):
    """Busca columna en df por nombre normalizado (tolerante a tildes/case)."""
    norm_map = {_norm(c): c for c in df.columns}
    for cand in candidates:
        c = _norm(cand)
        if c in norm_map:
            return norm_map[c]
    return None

def to_int_sku(val):
    """Convierte SKU a int de forma robusta. Devuelve None si no se puede."""
    if val is None: return None
    if isinstance(val, float) and math.isnan(val): return None
    try:
        s = str(val).strip()
        if s in ('', 'nan', 'NaN', 'None'): return None
        return int(float(s))
    except (ValueError, TypeError):
        return None

def is_envase(sku, desc):
    if sku in EXCLUDED_SKUS: return True
    d = str(desc).upper()
    for p in EXCLUDED_PATS:
        if p in d: return True
    return d.strip() == 'RET'

def normalize_cancha(val):
    if val is None or (isinstance(val, float) and math.isnan(val)):
        return 'SIN CANCHA'
    s = str(val).strip().upper()
    if s in ('', '0', 'NAN', 'NONE'): return 'SIN CANCHA'
    # Eliminar tildes
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    MAP = {'CANCHA I':'CANCHA I','CANCHA II':'CANCHA II','CANCHA III':'CANCHA III',
           'CANCHA IV':'CANCHA IV','CANCHA V':'CANCHA V',
           'MKPL':'CANCHA V','MERCH':'CANCHA V','MERCHANDISING':'CANCHA V',
           'MARKETPLACE':'CANCHA V'}
    # Match exacto primero
    if s in MAP: return MAP[s]
    # Match por prefijo (más específico primero)
    for k in sorted(MAP.keys(), key=len, reverse=True):
        if s.startswith(k): return MAP[k]
    return 'SIN CANCHA'

# ─── CARGA DE DATOS ─────────────────────────────────────────────────────────

def load_car(file, ddm=None):
    """
    v3.8 — Si se pasa ddm, se usa para decidir si los huérfanos del bloque azul
    son paletas enteras (no se agregan) o tienen fracción / no tienen BXP
    (se agregan al df). Si ddm es None, no se procesan huérfanos (compat retro).

    Retorna: (df, blue_audit) donde blue_audit es una lista de dicts con la
    trazabilidad de cada fila del bloque azul y su destino.
    """
    # ── Leer hoja completa sin header para manejar filas por índice ──────────
    raw = pd.read_excel(file, sheet_name=0, header=None)

    # Fila 0 = header (Excel fila 1)
    header = raw.iloc[0].tolist()

    # ── FILAS AZULES: Excel filas 2-41 → índices 1-40 ────────────────────────
    # Columna T = índice 19 (A=0 … T=19)
    blue_raw = raw.iloc[1:41].copy()
    blue_raw.columns = header

    blue_raw["Artículo"]   = pd.to_numeric(blue_raw["Artículo"], errors="coerce")
    blue_raw["Transporte"] = pd.to_numeric(blue_raw["Transporte"], errors="coerce")
    blue_raw["_bultos_azul"] = pd.to_numeric(blue_raw.iloc[:, 19], errors="coerce").fillna(0)
    blue_raw = blue_raw.dropna(subset=["Artículo", "Transporte"])
    blue_raw = blue_raw[blue_raw["Artículo"] > 0]
    blue_raw = blue_raw[blue_raw["_bultos_azul"] > 0]
    blue_raw["Artículo"]   = blue_raw["Artículo"].astype(int)
    blue_raw["Transporte"] = blue_raw["Transporte"].apply(lambda x: int(float(x)) if pd.notna(x) else x)
    # Excluir envases de las filas azules
    blue_raw = blue_raw[~blue_raw.apply(
        lambda r: is_envase(int(r["Artículo"]), str(r.get("Descripción Artículo", ""))), axis=1
    )]

    # Guardar descripción del primer registro azul por SKU (para huérfanos)
    blue_desc_map = (blue_raw.dropna(subset=["Descripción Artículo"])
                     .drop_duplicates(subset=["Artículo"], keep="first")
                     .set_index("Artículo")["Descripción Artículo"].to_dict())

    # Agregar bultos azules por Transporte + SKU
    blue_agg = (blue_raw.groupby(["Transporte", "Artículo"], as_index=False)
                .agg(_bultos_azul=("_bultos_azul", "sum")))
    blue_total_in = float(blue_agg["_bultos_azul"].sum()) if len(blue_agg) else 0.0

    # ── FILAS NORMALES: Excel fila 42 en adelante → índice 41+ ───────────────
    normal_raw = raw.iloc[41:].copy()
    normal_raw.columns = header

    df = normal_raw.copy()
    df["Artículo"]   = pd.to_numeric(df["Artículo"], errors="coerce")
    df["Transporte"] = pd.to_numeric(df["Transporte"], errors="coerce")
    df = df.dropna(subset=["Artículo", "Transporte"])
    df = df[df["Artículo"] > 0]
    df["Artículo"]   = df["Artículo"].astype(int)
    df["Transporte"] = df["Transporte"].apply(lambda x: int(float(x)) if pd.notna(x) else x)
    df["Bultos"]     = pd.to_numeric(df["Bultos"], errors="coerce").fillna(0)
    df["Unids"]      = pd.to_numeric(df["Unids"],  errors="coerce").fillna(0)
    df = df[~df.apply(lambda r: is_envase(r["Artículo"], r["Descripción Artículo"]), axis=1)]

    # Inicializar flag de huérfano sin DDM (siempre False por defecto)
    df["_orphan_no_ddm"] = False
    blue_audit = []

    # ── MERGE: sumar bultos azules sobre los normales ─────────────────────────
    blue_consumed_merge = 0.0
    blue_consumed_orphans = 0.0
    blue_discarded_pal_entera = 0.0

    if not blue_agg.empty:
        # Set de pares (T, SKU) que existen en filas normales
        normal_keys = set(zip(df["Transporte"], df["Artículo"]))

        # Partir blue_agg en dos: los que tienen match y los huérfanos
        blue_agg["_has_match"] = blue_agg.apply(
            lambda r: (r["Transporte"], r["Artículo"]) in normal_keys, axis=1)

        blue_match    = blue_agg[blue_agg["_has_match"]].copy()
        blue_orphans  = blue_agg[~blue_agg["_has_match"]].copy()

        # ── Caso 1: con match → merge clásico (sumar bultos azules) ───────────
        if not blue_match.empty:
            blue_match_clean = blue_match[["Transporte","Artículo","_bultos_azul"]]
            df = df.merge(blue_match_clean, on=["Transporte", "Artículo"], how="left")
            df["_bultos_azul"] = df["_bultos_azul"].fillna(0)
            df["Bultos"] = df["Bultos"] + df["_bultos_azul"]
            blue_consumed_merge = float(blue_match_clean["_bultos_azul"].sum())
            df = df.drop(columns=["_bultos_azul"])

            # Trazabilidad: filas sumadas
            for _, br in blue_match.iterrows():
                blue_audit.append({
                    "Transporte": int(br["Transporte"]),
                    "SKU":        int(br["Artículo"]),
                    "Descripción": str(blue_desc_map.get(int(br["Artículo"]), "")),
                    "Bultos azules": float(br["_bultos_azul"]),
                    "BXP": None,
                    "Acción": "Sumado a fila normal",
                })

        # ── Caso 2: huérfanos → decidir según DDM ─────────────────────────────
        if not blue_orphans.empty and ddm is not None:
            # Tomar info de reparto y descripción de transporte de cualquier
            # fila normal del mismo transporte (siempre existe)
            transp_info = (df.dropna(subset=["Número"])
                           .drop_duplicates(subset=["Transporte"], keep="first")
                           .set_index("Transporte")
                           [["Número","Fecha Mvto","Descripción Transporte","Depósito","Descripción Depósito"]]
                           .to_dict("index"))

            new_rows = []
            for _, br in blue_orphans.iterrows():
                t   = int(br["Transporte"]); sku = int(br["Artículo"])
                blt = float(br["_bultos_azul"])
                desc = str(blue_desc_map.get(sku, f"SKU {sku}"))

                d = ddm[ddm["sku"] == sku]
                bxp = float(d.iloc[0]["bxp"]) if len(d) and pd.notna(d.iloc[0]["bxp"]) else None

                if bxp and bxp > 0 and (blt % bxp == 0):
                    # Paleta entera exacta → NO agregar (no es picking)
                    blue_discarded_pal_entera += blt
                    blue_audit.append({
                        "Transporte": t, "SKU": sku, "Descripción": desc,
                        "Bultos azules": blt, "BXP": bxp,
                        "Acción": "Descartado (paleta entera)",
                    })
                    continue

                # Necesita aparecer en PDF: tiene fracción o no tiene BXP
                if t not in transp_info:
                    # No hay info de reparto → no se puede asignar a un reparto
                    # No debería pasar nunca pero registramos
                    blue_audit.append({
                        "Transporte": t, "SKU": sku, "Descripción": desc,
                        "Bultos azules": blt, "BXP": bxp,
                        "Acción": "ERROR: sin info de reparto",
                    })
                    continue

                info = transp_info[t]
                no_ddm = (bxp is None or bxp <= 0)

                new_row = {col: None for col in df.columns}
                new_row["Depósito"]              = info.get("Depósito")
                new_row["Descripción Depósito"]  = info.get("Descripción Depósito")
                new_row["Número"]                = info["Número"]
                new_row["Fecha Mvto"]            = info["Fecha Mvto"]
                new_row["Transporte"]            = t
                new_row["Descripción Transporte"]= info["Descripción Transporte"]
                new_row["Artículo"]              = sku
                new_row["Descripción Artículo"]  = desc
                new_row["Bultos"]                = blt
                new_row["Unids"]                 = 0
                new_row["_orphan_no_ddm"]        = no_ddm

                new_rows.append(new_row)
                blue_consumed_orphans += blt

                blue_audit.append({
                    "Transporte": t, "SKU": sku, "Descripción": desc,
                    "Bultos azules": blt, "BXP": bxp,
                    "Acción": ("Agregado como picking (SIN BXP — REVISAR)"
                               if no_ddm else "Agregado como picking (fracción)"),
                })

            if new_rows:
                # NOTA: el header del CAR puede tener columnas con nombres
                # duplicados (espacios non-breaking). Para evitar fallar en
                # pd.concat con InvalidIndexError, construimos el dataframe
                # nuevo respetando posiciones de columnas, no nombres.
                new_df = pd.DataFrame(new_rows)
                # Reindexar usando solo las columnas que SÍ existen en df
                # (asignación posicional por nombre único)
                aligned = pd.DataFrame(columns=df.columns)
                for nr in new_rows:
                    aligned_row = []
                    for col in df.columns:
                        # Si la columna está repetida tomamos el primer valor encontrado
                        aligned_row.append(nr.get(col, None))
                    aligned.loc[len(aligned)] = aligned_row
                df = pd.concat([df, aligned], ignore_index=True, sort=False)

        elif not blue_orphans.empty and ddm is None:
            # Sin DDM no podemos decidir → registrar pero no agregar (compat)
            for _, br in blue_orphans.iterrows():
                blue_audit.append({
                    "Transporte": int(br["Transporte"]),
                    "SKU":        int(br["Artículo"]),
                    "Descripción": str(blue_desc_map.get(int(br["Artículo"]), "")),
                    "Bultos azules": float(br["_bultos_azul"]),
                    "BXP": None,
                    "Acción": "PERDIDO (sin DDM para decidir)",
                })

    # ── Invariante de validación ─────────────────────────────────────────────
    blue_total_out = (blue_consumed_merge + blue_consumed_orphans
                      + blue_discarded_pal_entera)
    if blue_total_in > 0 and abs(blue_total_in - blue_total_out) > 0.01 and ddm is not None:
        raise ValueError(
            f"Invariante de bloque azul rota: entrada={blue_total_in:.0f} "
            f"bultos, salida={blue_total_out:.0f} bultos "
            f"(merge={blue_consumed_merge:.0f} + huérfanos={blue_consumed_orphans:.0f} "
            f"+ descartados pal entera={blue_discarded_pal_entera:.0f}). "
            f"Diferencia: {blue_total_in - blue_total_out:.0f} bultos perdidos."
        )

    return df, blue_audit

def load_frescura(file):
    """
    FIX v3.4 — Construcción correcta de cancha por SKU:
      1) Lee la hoja API completa.
      2) SKU está en columna A ('Artículo'). Cada SKU tiene un Nº de almacén
         (columna 'almacén' o similar).
      3) Lee el bloque I:M de la misma hoja API (tabla maestra de almacenes):
         I=Nº Almacén | J=Detalle | K=Calibre | L=Apilabilidad | M=Cancha
      4) Mergea SKU → Nº Almacén → Cancha.
    """
    xls = pd.ExcelFile(file)

    # ── HOJA API: lectura completa ──────────────────────────────────────────
    api_full = pd.read_excel(xls, sheet_name='API', header=0)

    # Columna SKU (col A)
    col_sku = find_col(api_full, 'Artículo', 'Articulo', 'SKU', 'Cód', 'Cod')
    if col_sku is None:
        raise ValueError("Hoja API: no se encontró la columna de SKU (Artículo).")

    # Columna almacén por SKU (entre A y H, la que asocia SKU con su almacén)
    col_alm_sku = find_col(api_full, 'almacén', 'almacen', 'Almacén', 'Almacen', 'Alm')
    if col_alm_sku is None:
        raise ValueError("Hoja API: no se encontró la columna 'almacén' que asocia SKU con almacén.")

    # Tabla SKU → Nº Almacén
    sku_alm = api_full[[col_sku, col_alm_sku]].copy()
    sku_alm.columns = ['sku_raw', 'alm_num_raw']
    sku_alm['sku'] = sku_alm['sku_raw'].apply(to_int_sku)
    sku_alm['alm_num'] = pd.to_numeric(sku_alm['alm_num_raw'], errors='coerce')
    sku_alm = sku_alm.dropna(subset=['sku'])
    sku_alm['sku'] = sku_alm['sku'].astype(int)
    # Regla de prevalencia: la última fila con almacén válido prevalece
    sku_alm = sku_alm.dropna(subset=['alm_num'])
    sku_alm['alm_num'] = sku_alm['alm_num'].astype(int)
    sku_alm = sku_alm.drop_duplicates(subset=['sku'], keep='last')[['sku', 'alm_num']]

    # ── BLOQUE I:M (tabla maestra de almacenes), header en fila 1 ───────────
    # usecols='I:M', header=0 → lee solo las cols I-M con su header propio.
    alm_master = pd.read_excel(
        xls, sheet_name='API', header=0, usecols='I:M'
    )
    # Mapeo posicional según definición del usuario:
    # I = Nº Almacén | J = Detalle | K = Calibre | L = Apilabilidad | M = Cancha
    if alm_master.shape[1] < 5:
        raise ValueError(
            f"Bloque I:M de la hoja API tiene {alm_master.shape[1]} columnas, se esperaban 5."
        )
    alm_master = alm_master.iloc[:, :5].copy()
    alm_master.columns = ['alm_num', 'alm_det', 'calibre', 'apil', 'cancha_raw']
    alm_master['alm_num'] = pd.to_numeric(alm_master['alm_num'], errors='coerce')
    alm_master = alm_master.dropna(subset=['alm_num'])
    alm_master['alm_num'] = alm_master['alm_num'].astype(int)
    # Prevalece la última definición del almacén
    alm_master = alm_master.drop_duplicates(subset=['alm_num'], keep='last')

    # ── MERGE: SKU → Almacén → Cancha ───────────────────────────────────────
    api = sku_alm.merge(alm_master, on='alm_num', how='left')
    api['cancha'] = api['cancha_raw'].apply(normalize_cancha)
    api = api.rename(columns={'alm_num': 'alm_id'})
    api['alm_det'] = api['alm_det'].fillna('sin datos').astype(str)
    api = api[['sku', 'cancha', 'alm_id', 'alm_det']]

    # ── HOJA DDM (Bultos por pallet) ────────────────────────────────────────
    ddm_full = pd.read_excel(xls, sheet_name='DDM')
    col_ddm_sku = find_col(ddm_full, 'ARTÍCULO', 'ARTICULO', 'Artículo', 'Articulo', 'SKU')
    col_bxp     = find_col(ddm_full, 'BULTOS X PALLET', 'BULTOS_X_PALLET', 'BXP')
    if not col_ddm_sku or not col_bxp:
        raise ValueError("Hoja DDM: no se encontraron las columnas ARTÍCULO y/o BULTOS X PALLET.")
    ddm = ddm_full[[col_ddm_sku, col_bxp]].copy()
    ddm.columns = ['sku', 'bxp']
    ddm['sku'] = ddm['sku'].apply(to_int_sku)
    ddm = ddm.dropna(subset=['sku'])
    ddm['sku'] = ddm['sku'].astype(int)
    ddm['bxp'] = pd.to_numeric(ddm['bxp'], errors='coerce')
    ddm = ddm.drop_duplicates(subset=['sku'], keep='last')

    # ── HOJA Frescura ───────────────────────────────────────────────────────
    fr_full = pd.read_excel(xls, sheet_name='Frescura')
    col_fr_sku    = find_col(fr_full, 'Cód', 'Cod', 'SKU', 'Artículo', 'Articulo')
    col_fr_fecha  = find_col(fr_full, 'FECHA DE VENC.', 'FECHA DE VENC', 'Fecha de Venc', 'Vencimiento')
    col_fr_status = find_col(fr_full, 'Status', 'Estado')
    if not all([col_fr_sku, col_fr_fecha, col_fr_status]):
        raise ValueError("Hoja Frescura: faltan columnas Cód / FECHA DE VENC. / Status.")
    fr = fr_full[[col_fr_sku, col_fr_fecha, col_fr_status]].copy()
    fr.columns = ['sku', 'fecha', 'status']
    fr['sku'] = fr['sku'].apply(to_int_sku)
    fr = fr.dropna(subset=['sku'])
    fr['sku'] = fr['sku'].astype(int)
    fr = fr.sort_values('fecha').drop_duplicates(subset=['sku'], keep='first')

    # Diagnóstico para debug
    diag = {
        'api_full_cols': list(api_full.columns),
        'col_sku_used': col_sku,
        'col_alm_sku_used': col_alm_sku,
        'alm_master_cols_raw': list(pd.read_excel(xls, sheet_name='API', header=0, usecols='I:M').columns),
        'sku_count_api': len(api),
        'alm_master_count': len(alm_master),
        'sku_sin_cancha': int((api['cancha'] == 'SIN CANCHA').sum()),
        'alm_master_preview': alm_master.head(20).to_dict('records'),
    }
    return api, ddm, fr, diag

def get_lema(transport, idx=0):
    h = int(hashlib.md5(str(transport).encode()).hexdigest(), 16)
    return LEMAS[(h + idx) % len(LEMAS)]

def pallet_reading(val):
    if not val or val == 0.0:
        return ('0.00', '0 paletas', 'Sin carga de picking en esta cancha')
    vm = f'{val:.2f}'; n = int(val); fr = val - n
    if val < 1.0:  return (vm, '1 paleta',    'Menos de una paleta')
    if val == 1.0: return (vm, '1 paleta',    '1 paleta exacta')
    if val < 1.5:  return (vm, '2 paletas',   'Mas de una paleta')
    if val == 1.5: return (vm, '2 paletas',   '1 paleta y media')
    if val < 2.0:  return (vm, '2 paletas',   'Casi 2 paletas')
    if fr == 0.0:  return (vm, f'{n} paletas',   f'{n} paletas exactas')
    if fr < 0.5:   return (vm, f'{n+1} paletas', f'Mas de {n} paletas')
    if fr == 0.5:  return (vm, f'{n+1} paletas', f'{n} paletas y media')
    return          (vm, f'{n+1} paletas', f'Casi {n+1} paletas')

def process_reparto(rep_df, api, ddm, fr):
    rows = []
    for _, r in rep_df.iterrows():
        sku = int(r['Artículo'])
        blts_raw = float(r['Bultos']); unids = float(r['Unids'])
        desc = str(r['Descripción Artículo'])
        orphan_no_ddm = bool(r.get('_orphan_no_ddm', False))

        a = api[api['sku'] == sku]
        if len(a):
            cancha     = str(a.iloc[0]['cancha'])
            alm_id_i   = int(a.iloc[0]['alm_id']) if pd.notna(a.iloc[0]['alm_id']) else None
            alm_det    = str(a.iloc[0]['alm_det'])
        else:
            cancha, alm_id_i, alm_det = 'SIN CANCHA', None, 'sin datos'

        d = ddm[ddm['sku'] == sku]
        bxp = float(d.iloc[0]['bxp']) if len(d) and pd.notna(d.iloc[0]['bxp']) else None

        if bxp and bxp > 0:
            pal_ent = int(blts_raw / bxp); blts_pick = blts_raw - pal_ent * bxp
            has_pal = pal_ent >= 1; miss_bxp = False
        else:
            pal_ent = 0; blts_pick = blts_raw; has_pal = False; miss_bxp = True

        skip_color = (sku == 7038) or (alm_id_i == 46)
        f = fr[fr['sku'] == sku]
        if len(f) and not skip_color:
            fv = f.iloc[0]['fecha']; st_ = str(f.iloc[0]['status']).strip().upper()
            if pd.isna(fv): fecha_s, row_st = '', 'NONE'
            else:
                fecha_s = pd.Timestamp(fv).strftime('%d/%m/%y')
                row_st  = st_ if st_ in ('RED','YELLOW','GREEN') else 'NONE'
        else:
            fecha_s, row_st = '', 'NONE'

        rows.append({'sku':sku,'desc':desc,'blts_raw':blts_raw,'unids':unids,
                     'blts_pick':blts_pick,'pal_ent':pal_ent,'has_pal':has_pal,
                     'miss_bxp':miss_bxp,'bxp':bxp,'cancha':cancha,
                     'alm_id':alm_id_i,'alm_det':alm_det,'fecha_s':fecha_s,'row_st':row_st,
                     'orphan_no_ddm':orphan_no_ddm})
    return rows

def build_alm_groups(c_rows):
    groups = {}; order = []
    for r in c_rows:
        key = (r['alm_id'], r['alm_det'])
        if key not in groups: groups[key] = []; order.append(key)
        groups[key].append(r)
    for key in order:
        groups[key].sort(key=lambda x: -x['blts_raw'])
    # Ordenar almacenes de menor a mayor por número de almacén
    order.sort(key=lambda k: (k[0] is None, k[0] if k[0] is not None else 0))
    return [(k[0], k[1], groups[k]) for k in order]

def compute_pall_value(rows, cancha):
    return sum(r['blts_pick']/r['bxp'] for r in rows
               if r['cancha'] == cancha and r['bxp'] and r['bxp'] > 0)

# ─── DRAWING HELPERS ─────────────────────────────────────────────────────────

def ry(y_top): return PAGE_H - y_top

def rfill(c, x, y_top, w, h, fill, stroke=None, lw=0.4):
    c.setFillColor(fill)
    if stroke:
        c.setStrokeColor(stroke); c.setLineWidth(lw)
        c.rect(x, ry(y_top+h), w, h, fill=1, stroke=1)
    else:
        c.rect(x, ry(y_top+h), w, h, fill=1, stroke=0)

def txt(c, x, y_top, s, font='Helvetica', sz=8, col=colors.black, align='left', mw=None):
    c.setFont(font, sz); c.setFillColor(col)
    if mw:
        while s and c.stringWidth(s, font, sz) > mw: s = s[:-1]
    fn = {'left': c.drawString, 'center': c.drawCentredString, 'right': c.drawRightString}[align]
    fn(x, ry(y_top), s)

def draw_watermark(c):
    c.saveState()
    c.setFillColor(WM_GRAY); c.setFillAlpha(0.07)
    c.setFont('Helvetica-Bold', 68)
    c.translate(PAGE_W/2, PAGE_H/2); c.rotate(35)
    c.drawCentredString(0, 0, 'BECCACECE HNOS')
    c.restoreState()

def draw_title_bar(c, y, numero, fecha_str):
    rfill(c, MARGIN, y, CW, H_TITLE, DARK_BLUE)
    label = f'PLANILLA DE CARGA  |  Reparto Nro: {numero}  |  Fecha: {fecha_str}'
    txt(c, MARGIN+CW/2, y+17, label, 'Helvetica-Bold', 10.5, colors.white, 'center', CW-16)
    return H_TITLE

def fmt_transport(t):
    """Elimina el .0 final del número de transporte (128.0 → 128)."""
    try:
        v = float(t)
        if v == int(v):
            return str(int(v))
    except (ValueError, TypeError):
        pass
    return str(t)

def draw_transport_line(c, y, transport, chofer):
    rfill(c, MARGIN, y, CW, H_TRANS, colors.HexColor('#EEF2F8'))
    txt(c, MARGIN+6,      y+12, f'Transporte: {fmt_transport(transport)} - {chofer}', 'Helvetica-Bold', 9)
    txt(c, MARGIN+CW-6,   y+12, 'Depósito: 001 - CASA CENTRAL',        'Helvetica-Bold', 9, align='right')
    return H_TRANS

def draw_partida_regreso(c, y):
    txt(c, MARGIN+6,     y+11, 'F. y H. Est. de Partida: ________ Hs.', 'Helvetica', 9)
    txt(c, MARGIN+CW-6,  y+11, 'F. y H. Est. de Regreso: ________ Hs.','Helvetica', 9, align='right')
    return H_PARTIDA

def draw_control_carga(c, y):
    hdr_h = 15; box_h = H_CTRL - hdr_h; half = CW/2
    rfill(c, MARGIN,        y, half, hdr_h, HDR_BG, BORDER)
    rfill(c, MARGIN+half,   y, half, hdr_h, HDR_BG, BORDER)
    txt(c, MARGIN+half/2,         y+10, 'CONTROL DE CARGA',    'Helvetica-Bold', 9, align='center')
    txt(c, MARGIN+half+half/2,    y+10, 'CONTROL DE DESCARGA', 'Helvetica-Bold', 9, align='center')
    c.setStrokeColor(BORDER); c.setLineWidth(0.4)
    c.rect(MARGIN, ry(y+hdr_h+box_h), CW, box_h, fill=0, stroke=1)
    c.line(MARGIN+half, ry(y+hdr_h), MARGIN+half, ry(y+hdr_h+box_h))
    return H_CTRL

def draw_lema(c, y, lema):
    rfill(c, MARGIN, y, CW, H_LEMA, DARK_BLUE)
    txt(c, MARGIN+CW/2, y+12, f'"{lema}"', 'Helvetica-Oblique', 9, colors.white, 'center', CW-16)
    return H_LEMA

def draw_cancha_header(c, y, label):
    rfill(c, MARGIN, y, CW, H_CANCHA, MED_BLUE)
    txt(c, MARGIN+10, y+13, f'◀  {label}', 'Helvetica-Bold', 11, colors.white)
    return H_CANCHA

def draw_sin_cancha_header(c, y):
    rfill(c, MARGIN, y, CW, H_CANCHA, RED_ALERT)
    txt(c, MARGIN+CW/2, y+13, '⚠  SIN CANCHA ASIGNADA — REVISAR', 'Helvetica-Bold', 11, colors.white, 'center')
    return H_CANCHA

def draw_table_header(c, y):
    labels = ['SKU', 'Descripción', 'Venc.', 'Bultos', 'Unids']
    rfill(c, MARGIN, y, CW, H_THDR, HDR_BG, BORDER)
    x = MARGIN
    for lbl, w in zip(labels, COL_W):
        c.setStrokeColor(BORDER); c.setLineWidth(0.4)
        c.rect(x, ry(y+H_THDR), w, H_THDR, fill=0, stroke=1)
        txt(c, x+w/2, y+9.5, lbl, 'Helvetica-Bold', 8, align='center')
        x += w
    return H_THDR

def draw_alm_header(c, y, alm_id, alm_det):
    label = f'Almacén {alm_id} — {alm_det}' if alm_id is not None else 'Almacén: SIN DATOS'
    rfill(c, MARGIN, y, CW, H_ALM, ALM_BG, BORDER)
    txt(c, MARGIN+6, y+11, label, 'Helvetica-Bold', 9)
    return H_ALM

def draw_product_row(c, y, r):
    st_ = r['row_st']
    is_orphan_no_ddm = r.get('orphan_no_ddm', False)
    bg  = DARK_ROW if st_=='RED' else MED_ROW if st_=='YELLOW' else (
          AMBER if r.get('is_sin_cancha') else
          LIGHT_GRAY if r['has_pal'] else colors.white)

    rfill(c, MARGIN, y, CW, H_ROW, bg, BORDER)
    c.setStrokeColor(BORDER); c.setLineWidth(0.4)

    # Marca de huérfano sin DDM: (*) rojo después del SKU
    sku_s = str(r['sku'])
    if is_orphan_no_ddm:
        sku_s = f"{r['sku']}(*)"
    elif r['miss_bxp']:
        sku_s = f"{r['sku']}(!)"

    bp     = r['blts_pick']; bp_s = str(int(bp)) if bp == int(bp) else f'{bp:.0f}'
    un_s   = str(int(r['unids']))
    ind    = {'RED':'■','YELLOW':'▲','GREEN':'○'}.get(st_,'')
    venc_s = f'{ind} {r["fecha_s"]}' if r['fecha_s'] else ''
    font   = 'Helvetica-Bold' if st_=='RED' else 'Helvetica'

    # Descripción con aviso embebido si es huérfano sin DDM
    desc_show = r['desc']
    if is_orphan_no_ddm:
        desc_show = f"{r['desc']}  ⚠ SIN BXP EN DDM — REVISAR"

    cells = [(sku_s,'center'),(desc_show,'left'),(venc_s,'center'),(bp_s,'center'),(un_s,'center')]
    x = MARGIN
    for i, ((s, align), w) in enumerate(zip(cells, COL_W)):
        c.rect(x, ry(y+H_ROW), w, H_ROW, fill=0, stroke=1)
        # Color rojo en SKU y descripción si es huérfano sin DDM
        col_text = RED_ALERT if (is_orphan_no_ddm and i in (0, 1)) else colors.black
        font_use = 'Helvetica-Bold' if (is_orphan_no_ddm and i in (0, 1)) else font
        if align == 'center':
            txt(c, x+w/2, y+9.5, s, font_use, 9, col=col_text, align='center', mw=w-3)
        else:
            txt(c, x+3,   y+9.5, s, font_use, 9, col=col_text, mw=w-5)
        # Subrayado SOLO en columna Bultos (índice 3)
        if r['has_pal'] and i == 3:
            tw = c.stringWidth(s, font, 9)
            cx = x+w/2
            c.setStrokeColor(colors.black); c.setLineWidth(0.6)
            c.line(cx-tw/2, ry(y+H_ROW)+1, cx+tw/2, ry(y+H_ROW)+1)
            c.setLineWidth(0.4)
        x += w
    return H_ROW

def draw_total_row(c, y, tot_b, tot_u):
    mw = COL_W[0]+COL_W[1]+COL_W[2]
    rfill(c, MARGIN, y, CW, H_TOT, colors.HexColor('#E8E8E8'), BORDER)
    c.setStrokeColor(BORDER); c.setLineWidth(0.4)
    c.rect(MARGIN,    ry(y+H_TOT), mw,      H_TOT, fill=0, stroke=1)
    c.rect(MARGIN+mw, ry(y+H_TOT), COL_W[3],H_TOT, fill=0, stroke=1)
    c.rect(MARGIN+mw+COL_W[3], ry(y+H_TOT), COL_W[4], H_TOT, fill=0, stroke=1)
    txt(c, MARGIN+mw/2, y+9.5, 'T O T A L  A L M A C É N', 'Helvetica-Bold', 8.5, align='center')
    for val, xoff in [(tot_b, mw), (tot_u, mw+COL_W[3])]:
        s = str(int(val)) if val == int(val) else f'{val:.0f}'
        txt(c, MARGIN+xoff+COL_W[3 if xoff==mw else 4]/2, y+9.5, s, 'Helvetica-Bold', 8.5, align='center')
    return H_TOT

def draw_espacio_asignado(c, y, cancha, pv):
    rfill(c, MARGIN, y, CW, H_ESPACIO, BG_GRAY, BORDER)
    txt(c, MARGIN+CW/2, y+10, f'ESPACIO ASIGNADO — {cancha}', 'Helvetica-Bold', 9, align='center')
    col_w = CW/3
    hdrs  = ['Valor mínimo','Valor entero','Lectura']
    bx = MARGIN; by = y+14; rh = 13
    c.setStrokeColor(BORDER); c.setLineWidth(0.4)
    for h in hdrs:
        rfill(c, bx, by, col_w, rh, HDR_BG, BORDER)
        txt(c, bx+col_w/2, by+9, h, 'Helvetica-Bold', 8, align='center')
        bx += col_w
    vm, ve, lec = pallet_reading(pv)
    bx = MARGIN; by2 = by+rh
    for v in [vm, ve, lec]:
        rfill(c, bx, by2, col_w, rh, colors.white, BORDER)
        txt(c, bx+col_w/2, by2+9, v, 'Helvetica', 8.5, align='center', mw=col_w-4)
        bx += col_w
    return H_ESPACIO

def draw_route_to(c, y, dest):
    rfill(c, MARGIN, y, CW, H_ROUTE, YELLOW_RTE, BORDER)
    txt(c, MARGIN+CW/2, y+11, f'▶  DIRIGIR PALETA A: {dest}', 'Helvetica-Bold', 9.5, align='center')
    return H_ROUTE

def draw_route_recv(c, y):
    rfill(c, MARGIN, y, CW, H_ROUTE, LIGHT_GRAY, BORDER)
    txt(c, MARGIN+CW/2, y+11, '◀  RECIBE PALETA DESDE: CANCHA I', 'Helvetica-Bold', 9.5, align='center')
    return H_ROUTE

def draw_footer(c, date_str, page_num):
    yb = PAGE_H - MARGIN + 2
    c.setStrokeColor(colors.HexColor('#CCCCCC')); c.setLineWidth(0.5)
    c.line(MARGIN, ry(yb-4), MARGIN+CW, ry(yb-4))
    txt(c, MARGIN+CW/2, yb+6, 'Almacen Digital 3.0  |  Beccacece Hnos SA  |  Distribuidor Oficial CMQ — Desde 1963',
        'Helvetica', 8, FOOT_GRAY, 'center')
    txt(c, MARGIN,    yb+16, date_str,           'Helvetica', 8, FOOT_GRAY)
    txt(c, MARGIN+CW, yb+16, f'Página: {page_num}','Helvetica', 8, FOOT_GRAY, 'right')

def header_height(is_first_rep_page: bool) -> float:
    h = H_TITLE + GAP + H_TRANS + GAP
    if is_first_rep_page:
        h += H_PARTIDA + GAP + H_CTRL + GAP
    h += H_LEMA + GAP + H_CANCHA + GAP + H_THDR
    return h

def footer_reserve(has_route: bool) -> float:
    return H_ESPACIO + (H_ROUTE if has_route else 0) + GAP + H_FOOT

def content_avail(is_first_rep_page: bool, has_route: bool) -> float:
    return PAGE_H - 2*MARGIN - header_height(is_first_rep_page) - footer_reserve(has_route)

ALM_GAP = 8   # espacio entre grupos de almacén (más legible)

def alm_group_height(n_rows: int) -> float:
    return ALM_GAP + H_ALM + n_rows*H_ROW + H_TOT

def draw_cancha_pages(c, reparto_rows, cancha, is_first_rep, numero, transport,
                       chofer, lema, fecha_str, date_str, pc,
                       pall_values, route_dest, receives_route):
    c_rows     = [r for r in reparto_rows if r['cancha'] == cancha]
    alm_groups = build_alm_groups(c_rows)

    remaining = list(alm_groups)
    pages     = []
    first_pg  = True

    while True:
        is_frp    = is_first_rep and first_pg
        avail     = content_avail(is_frp, has_route=False)
        used      = 0
        pg_groups = []

        for g in remaining[:]:
            gh = alm_group_height(len(g[2]))
            if used + gh > avail and pg_groups:
                break
            pg_groups.append(remaining.pop(0))
            used += gh

        pages.append((is_frp, pg_groups))
        first_pg = False
        if not remaining:
            break

    n_pages = len(pages)
    for pi, (is_frp, pg_groups) in enumerate(pages):
        is_last = (pi == n_pages - 1)
        has_rt  = is_last and (route_dest or receives_route)

        if pc['n'] > 0:
            c.showPage()
        pc['n'] += 1

        draw_watermark(c)
        y = MARGIN

        y += draw_title_bar(c, y, numero, fecha_str);  y += GAP
        y += draw_transport_line(c, y, transport, chofer); y += GAP

        if is_frp:
            y += draw_partida_regreso(c, y); y += GAP
            y += draw_control_carga(c, y);   y += GAP

        y += draw_lema(c, y, lema); y += GAP

        lbl = cancha if pi == 0 else f'{cancha}  (continuación {pi+1})'
        y += draw_cancha_header(c, y, lbl); y += GAP
        y += draw_table_header(c, y)

        if not pg_groups and not c_rows:
            txt(c, MARGIN+CW/2, y+13, f'— Sin carga de {cancha} para este camión —',
                'Helvetica-Oblique', 8, colors.HexColor('#888888'), 'center')
        else:
            for alm_id, alm_det, rows in pg_groups:
                y += ALM_GAP
                y += draw_alm_header(c, y, alm_id, alm_det)
                for r in rows:
                    y += draw_product_row(c, y, r)
                y += draw_total_row(c, y,
                                    sum(r['blts_pick'] for r in rows),
                                    sum(r['unids']     for r in rows))

        if is_last:
            pv = pall_values.get(cancha, 0.0)
            route_h   = H_ROUTE if has_rt else 0
            foot_y    = PAGE_H - MARGIN - H_FOOT - H_ESPACIO - route_h - GAP
            draw_espacio_asignado(c, foot_y, cancha, pv)
            if route_dest:
                draw_route_to(c, foot_y + H_ESPACIO + GAP, route_dest)
            elif receives_route:
                draw_route_recv(c, foot_y + H_ESPACIO + GAP)

        draw_footer(c, date_str, pc['n'])


def draw_sin_cancha_page(c, row, numero, transport, chofer, lema, fecha_str, date_str, pc):
    if pc['n'] > 0:
        c.showPage()
    pc['n'] += 1
    draw_watermark(c)
    y = MARGIN

    y += draw_title_bar(c, y, numero, fecha_str);  y += GAP
    y += draw_transport_line(c, y, transport, chofer); y += GAP
    y += draw_lema(c, y, lema); y += GAP
    y += draw_sin_cancha_header(c, y); y += GAP
    y += draw_alm_header(c, y, None, 'sin asignación en base de datos'); y += GAP
    y += draw_table_header(c, y)

    rc = dict(row); rc['is_sin_cancha'] = True
    y += draw_product_row(c, y, rc); y += 8

    rfill(c, MARGIN, y, CW, 22, colors.HexColor('#FFF3E0'), colors.HexColor('#FFB300'), lw=0.6)
    note = f"SKU {row['sku']} no tiene cancha asignada. Validar con Jefe de Almacén antes de cargar."
    txt(c, MARGIN+CW/2, y+14, note, 'Helvetica', 7.5, colors.HexColor('#C0392B'), 'center', CW-10)

    draw_footer(c, date_str, pc['n'])

# ─── GENERADOR PRINCIPAL ─────────────────────────────────────────────────────

def generate_pdf(car_df, api, ddm, fr):
    buf = io.BytesIO()
    c   = rl_canvas.Canvas(buf, pagesize=A4)
    pc  = {'n': 0}

    if 'Fecha Mvto' in car_df.columns and len(car_df):
        try:
            dt = pd.Timestamp(car_df['Fecha Mvto'].dropna().iloc[0])
            fecha_str = date_str = dt.strftime('%d/%m/%Y')
        except:
            fecha_str = date_str = datetime.today().strftime('%d/%m/%Y')
    else:
        fecha_str = date_str = datetime.today().strftime('%d/%m/%Y')

    stats = {'repartos':[], 'red':[], 'yellow':[], 'pallet_applied':[],
             'miss_bxp':[], 'no_fecha':[], 'sin_cancha_skus':[],
             'orphans_no_ddm':[],
             'total_pages':0, 'trace':[]}

    for rep_idx, (numero, rep_df) in enumerate(car_df.groupby('Número', sort=False)):
        transport = str(rep_df['Transporte'].iloc[0])
        chofer    = str(rep_df['Descripción Transporte'].iloc[0])
        lema      = get_lema(transport, rep_idx)
        stats['repartos'].append({'numero': numero, 'transport': transport, 'chofer': chofer})

        rows        = process_reparto(rep_df, api, ddm, fr)
        pall_values = {ch: compute_pall_value(rows, ch) for ch in CANCHA_ORDER}

        # Trazabilidad de cancha por SKU para el debug
        for r in rows:
            stats['trace'].append({
                'Reparto': numero, 'Camión': transport,
                'SKU': r['sku'], 'Descripción': r['desc'][:50],
                'Almacén': r['alm_id'], 'Detalle Almacén': r['alm_det'],
                'Cancha': r['cancha'], 'Bultos CAR': r['blts_raw'],
                'A pickear': r['blts_pick'],
                'Huérfano sin DDM': '⚠' if r.get('orphan_no_ddm') else '',
            })

        pv1 = pall_values['CANCHA I']; frac1 = pv1 - int(pv1)
        if frac1 > 0.001:
            pv2, pv4 = pall_values['CANCHA II'], pall_values['CANCHA IV']
            if pv2 == 0 and pv4 == 0:  route_dest = route_recv = None
            elif pv4 > pv2:             route_dest = route_recv = 'CANCHA IV'
            else:                       route_dest = route_recv = 'CANCHA II'
        else:
            route_dest = route_recv = None

        for ci, cancha in enumerate(CANCHA_ORDER):
            is_first_rep = (ci == 0)
            receives = (cancha == route_recv and not is_first_rep)
            draw_cancha_pages(c, rows, cancha, is_first_rep, numero, transport, chofer,
                              lema, fecha_str, date_str, pc, pall_values,
                              route_dest if is_first_rep else None, receives)

        for r in [r for r in rows if r['cancha'] == 'SIN CANCHA']:
            draw_sin_cancha_page(c, r, numero, transport, chofer, lema, fecha_str, date_str, pc)
            stats['sin_cancha_skus'].append(f"{r['sku']} — {r['desc']}")

        for r in rows:
            if r['row_st'] == 'RED':   stats['red'].append(r['sku'])
            if r['row_st'] == 'YELLOW':stats['yellow'].append(r['sku'])
            if r['has_pal']:           stats['pallet_applied'].append(r['sku'])
            if r['miss_bxp']:          stats['miss_bxp'].append(r['sku'])
            if not r['fecha_s']:       stats['no_fecha'].append(r['sku'])
            if r.get('orphan_no_ddm'): stats['orphans_no_ddm'].append(
                f"{r['sku']} — {r['desc']} (T{transport}, R{numero})")

    c.save(); buf.seek(0)
    stats['total_pages'] = pc['n']
    return buf.read(), stats

# ─── RESUMEN DE CARGA POR CAMIÓN (v3.9, hoja AGR) ───────────────────────────

def load_agr(file):
    """Lee la hoja AGR del CAR y devuelve DF consolidado por camión + SKU.

    Columnas resultantes: chofer | cod_producto | descripcion | bultos
    Una fila por (camión, SKU) con bultos sumados.
    Ignora Cliente / Razón social / Importes (no son relevantes para este doc).
    """
    file.seek(0)
    raw = pd.read_excel(file, sheet_name='AGR', header=None)

    # Buscar la fila de header (contiene 'Cliente' o 'Chofer')
    header_row = None
    for i in range(min(10, len(raw))):
        row_str = ' '.join(str(v) for v in raw.iloc[i].values if pd.notna(v))
        if 'Chofer' in row_str and 'Cantidad' in row_str:
            header_row = i
            break
    if header_row is None:
        raise ValueError("No se encontró el header de la hoja AGR (esperaba 'Chofer' y 'Cantidad').")

    df = pd.read_excel(file, sheet_name='AGR', header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    col_cod   = find_col(df, 'Cod Producto', 'CodProducto', 'Codigo Producto', 'Codigo', 'Código')
    col_desc  = find_col(df, 'DESCRIPCION', 'Descripcion', 'Descripción')
    col_cant  = find_col(df, 'Cantidad', 'Bultos')
    col_chof  = find_col(df, 'Chofer', 'Camion', 'Camión')

    if not all([col_cod, col_desc, col_cant, col_chof]):
        raise ValueError(
            f"Faltan columnas en hoja AGR. Encontradas: {list(df.columns)}. "
            f"Necesito: Cod Producto, Descripcion, Cantidad, Chofer."
        )

    df = df[[col_cod, col_desc, col_cant, col_chof]].copy()
    df.columns = ['cod_producto', 'descripcion', 'bultos', 'chofer']
    df = df.dropna(subset=['chofer', 'cod_producto', 'bultos'])

    df['chofer']       = df['chofer'].apply(to_int_sku)
    df['cod_producto'] = df['cod_producto'].apply(to_int_sku)
    df['bultos']       = pd.to_numeric(df['bultos'], errors='coerce').fillna(0).astype(int)
    df['descripcion']  = df['descripcion'].astype(str).str.strip()
    df = df[df['bultos'] > 0]

    # Consolidación: total por SKU por camión
    out = (df.groupby(['chofer', 'cod_producto', 'descripcion'], as_index=False)['bultos']
             .sum()
             .sort_values(['chofer', 'cod_producto']))
    return out


def _draw_resumen_header(c, fecha_str):
    """Cabecera del documento (título + subtítulo + línea separadora)."""
    margin = 15 * mm
    y_top = PAGE_H - margin

    c.setFillColor(colors.HexColor('#1a1a1a'))
    c.setFont('Helvetica-Bold', 16)
    c.drawString(margin, y_top - 14, 'Resumen de Carga por Camión')

    c.setFillColor(colors.HexColor('#666666'))
    c.setFont('Helvetica', 9)
    subt = f"Beccacece Hnos. S.A.   ·   La Reja, Moreno   ·   Fecha: {fecha_str}"
    c.drawString(margin, y_top - 28, subt)

    # Línea separadora
    c.setStrokeColor(colors.HexColor('#1a1a1a'))
    c.setLineWidth(0.8)
    c.line(margin, y_top - 36, PAGE_W - margin, y_top - 36)

    return y_top - 44  # y donde puede empezar el contenido (en coord top-down)


def _truck_block_height(n_rows):
    """Altura total de un bloque de camión (header + filas + spacer)."""
    return 7*mm + 6*mm + n_rows * 7*mm + 6*mm  # header + thead + filas + spacer


def _draw_truck_block(c, y_top, cam, rows):
    """Dibuja un bloque (header de camión + tabla) usando coords top-down.
    Devuelve el nuevo y_top (más abajo) después del bloque.
    """
    margin = 15 * mm
    total_w = PAGE_W - 2 * margin  # 180mm

    # Anchos de columna: Código 22 | Descripción 122 | Bultos 26 | Check 10 (mm)
    w_cod, w_desc, w_blt, w_chk = 22*mm, 122*mm, 26*mm, 10*mm

    # ── 1) Header del camión (barra negra) ──────────────────────────────────
    h_hdr = 7 * mm
    y_hdr = y_top - h_hdr
    c.setFillColor(colors.HexColor('#1a1a1a'))
    c.rect(margin, y_hdr, total_w, h_hdr, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont('Helvetica-Bold', 11)
    c.drawString(margin + 6, y_hdr + 2.3*mm, f'Camión {cam}')

    # ── 2) Header de columnas (banda gris clara) ────────────────────────────
    h_thd = 6 * mm
    y_thd = y_hdr - h_thd
    c.setFillColor(colors.HexColor('#f0f0f0'))
    c.rect(margin, y_thd, total_w, h_thd, fill=1, stroke=0)
    c.setStrokeColor(colors.HexColor('#cccccc'))
    c.setLineWidth(0.5)
    c.line(margin, y_thd, margin + total_w, y_thd)  # línea inferior

    c.setFillColor(colors.HexColor('#333333'))
    c.setFont('Helvetica-Bold', 8.5)
    # Centrado en la celda
    c.drawCentredString(margin + w_cod/2,                                y_thd + 1.7*mm, 'Código')
    c.drawString       (margin + w_cod + 6,                              y_thd + 1.7*mm, 'Descripción')
    c.drawCentredString(margin + w_cod + w_desc + w_blt/2,               y_thd + 1.7*mm, 'Bultos')
    c.drawCentredString(margin + w_cod + w_desc + w_blt + w_chk/2,       y_thd + 1.7*mm, 'Check')

    # ── 3) Filas ─────────────────────────────────────────────────────────────
    h_row = 7 * mm
    y_cur = y_thd
    c.setStrokeColor(colors.HexColor('#cccccc'))

    for i, r in enumerate(rows):
        y_cur -= h_row

        # Texto
        c.setFillColor(colors.HexColor('#1a1a1a'))
        c.setFont('Helvetica', 9)
        c.drawCentredString(margin + w_cod/2,                          y_cur + 2.2*mm, str(r['cod_producto']))
        c.drawString       (margin + w_cod + 6,                        y_cur + 2.2*mm, str(r['descripcion']))
        c.setFont('Helvetica-Bold', 9)
        c.drawCentredString(margin + w_cod + w_desc + w_blt/2,         y_cur + 2.2*mm, str(int(r['bultos'])))

        # Línea separadora entre filas (excepto la última, que la cierra el BOX)
        if i < len(rows) - 1:
            c.setStrokeColor(colors.HexColor('#e8e8e8'))
            c.setLineWidth(0.25)
            c.line(margin, y_cur, margin + total_w, y_cur)

        # Recuadro de check (casilla destacada con borde más marcado)
        chk_x = margin + w_cod + w_desc + w_blt
        # padding interno para que el cuadrado no toque los bordes de la fila
        pad_x, pad_y = 2.5*mm, 1.3*mm
        c.setStrokeColor(colors.HexColor('#666666'))
        c.setLineWidth(0.8)
        c.rect(chk_x + pad_x, y_cur + pad_y,
               w_chk - 2*pad_x, h_row - 2*pad_y,
               fill=0, stroke=1)

    # ── 4) Box exterior de toda la tabla ─────────────────────────────────────
    tabla_h = h_thd + len(rows) * h_row
    c.setStrokeColor(colors.HexColor('#cccccc'))
    c.setLineWidth(0.5)
    c.rect(margin, y_thd - len(rows) * h_row, total_w, tabla_h, fill=0, stroke=1)

    # Línea vertical antes de columna check
    chk_x = margin + w_cod + w_desc + w_blt
    c.line(chk_x, y_thd - len(rows)*h_row, chk_x, y_thd)

    # Devolver y_top siguiente (con un spacer de 6mm)
    return y_cur - 6 * mm


def _draw_resumen_footer(c, page_num, total_pages):
    """Pie de página: 'Página X de Y' centrado."""
    c.setFillColor(colors.HexColor('#666666'))
    c.setFont('Helvetica', 8)
    c.drawCentredString(PAGE_W / 2, 10 * mm, f'Página {page_num} de {total_pages}')


def build_resumen_carga_pdf(car_file):
    """Genera el PDF 'Resumen de Carga por Camión' a partir de la hoja AGR.

    Lógica:
      1) Lee hoja AGR y consolida bultos por (camión, SKU).
      2) Ordena camiones ascendentemente.
      3) Imprime en bloques compactos: varios camiones por página si entran.
      4) Ningún bloque se parte entre páginas.
      5) Pie de página 'Página X de Y'.
    """
    resumen = load_agr(car_file)
    if resumen.empty:
        raise ValueError("La hoja AGR no contiene datos válidos para generar el resumen.")

    margin = 15 * mm
    fecha_str = datetime.today().strftime('%d/%m/%Y')

    # ── PRIMERA PASADA: paginar (sin dibujar) para conocer total_pages ──────
    camiones = sorted(resumen['chofer'].unique())
    bloques = []  # lista de (cam, rows_list, alto)
    for cam in camiones:
        sub = resumen[resumen['chofer'] == cam].to_dict('records')
        h = _truck_block_height(len(sub))
        bloques.append((cam, sub, h))

    # Reserva para cabecera (solo en página 1) y pie (en todas)
    H_HEADER_PAGE1 = 44  # 'Resumen de Carga por Camión' + subt + línea + spacer
    H_FOOTER      = 18 * mm  # margen inferior efectivo donde reservamos espacio

    # Espacio disponible para contenido (top → bottom)
    def avail(is_first_page):
        top    = PAGE_H - margin
        bottom = H_FOOTER  # piso del contenido
        if is_first_page:
            top -= H_HEADER_PAGE1
        return top - bottom

    # Asignar bloques a páginas
    pages = []   # lista de lista de (cam, rows)
    current = []
    space_left = avail(True)
    is_first   = True
    for cam, rows, h in bloques:
        if h <= space_left:
            current.append((cam, rows))
            space_left -= h
        else:
            # Cerrar página actual y abrir nueva
            pages.append(current)
            is_first = False
            current = [(cam, rows)]
            space_left = avail(False) - h
    if current:
        pages.append(current)

    total_pages = len(pages)

    # ── SEGUNDA PASADA: dibujar ─────────────────────────────────────────────
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    c.setTitle('Resumen de Carga por Camión')

    for page_idx, page_content in enumerate(pages):
        is_first_page = (page_idx == 0)
        if is_first_page:
            y_top = _draw_resumen_header(c, fecha_str)
        else:
            y_top = PAGE_H - margin

        for cam, rows in page_content:
            y_top = _draw_truck_block(c, y_top, cam, rows)

        _draw_resumen_footer(c, page_idx + 1, total_pages)
        c.showPage()

    c.save()
    pdf_bytes = buf.getvalue()
    return pdf_bytes, {
        'total_pages': total_pages,
        'camiones': len(camiones),
        'filas_sku': len(resumen),
    }


# ============================================================================
# ║                                                                          ║
# ║        EXTENSIONES v4.0 — lectores MASTER, T2, extraíbles, UI            ║
# ║                                                                          ║
# ============================================================================

# ─── LECTORES MASTER (78 MB) ────────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_master_matriz_pall(file_bytes: bytes) -> tuple[pd.DataFrame, datetime | None]:
    """Lee la hoja 'Matriz Pall.' del MASTER.

    Estructura confirmada por inspección:
      - Celda A1: fecha del día (datetime)
      - Fila 1 cols B-Y: headers
      - Datos desde fila 2

    Returns:
        df: DataFrame con columnas tipadas
        fecha: datetime de A1 (o None si no se pudo leer)
    """
    wb = openpyxl.load_workbook(
        io.BytesIO(file_bytes), data_only=True, read_only=True,
    )
    if "Matriz Pall." not in wb.sheetnames:
        raise ValueError("Hoja 'Matriz Pall.' no encontrada en MASTER")
    ws = wb["Matriz Pall."]

    rows = list(ws.iter_rows(min_row=1, values_only=True))
    if not rows:
        raise ValueError("Hoja 'Matriz Pall.' vacía")

    # A1 = fecha del día; los headers de las columnas B-Y están en fila 1
    fecha_a1 = rows[0][0] if isinstance(rows[0][0], datetime) else None

    # Headers (col B en adelante)
    headers_raw = list(rows[0][1:25])  # B-Y
    headers = [
        "Camion", "Reparto",
        "UP_PICKING_AGR_FACT",
        "PICK_C1", "PICK_C2", "PICK_C3", "PICK_C4", "PICK_C5", "PICK_MKPL",
        "UP_AE_PALL_AGR_FACT",
        "AE_C1", "AE_C2", "AE_C3", "AE_C4", "AE_C5", "AE_MKPL",
        "KG_MATRIZ",
        "KG_C1", "KG_C2", "KG_C3", "KG_C4", "KG_C5", "KG_MKPL",
        "CLASIFICACION",
    ]

    # Datos: fila 2 en adelante, cols B-Y
    data = []
    for r in rows[1:]:
        if r[1] is None:  # Camion vacío → fin de datos
            break
        data.append(r[1:25])

    df = pd.DataFrame(data, columns=headers)
    # Tipos
    df["Camion"] = pd.to_numeric(df["Camion"], errors="coerce").astype("Int64")
    df["Reparto"] = df["Reparto"].astype(str).str.strip().str.upper()
    for col in df.columns:
        if col not in ("Camion", "Reparto"):
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    wb.close()
    return df, fecha_a1


@st.cache_data(show_spinner=False)
def load_master_matriz_picking_app(file_bytes: bytes) -> pd.DataFrame:
    """Lee la hoja 'Matriz Picking APP' del MASTER.

    Estructura: fila 1 = headers con A1 = fecha del día, 64 cols total.
    Para Sprint 1 (extraíble fx Picking) basta con devolver el bloque
    como DataFrame plano.
    """
    wb = openpyxl.load_workbook(
        io.BytesIO(file_bytes), data_only=True, read_only=True,
    )
    if "Matriz Picking APP" not in wb.sheetnames:
        raise ValueError("Hoja 'Matriz Picking APP' no encontrada en MASTER")
    ws = wb["Matriz Picking APP"]

    rows = list(ws.iter_rows(min_row=1, values_only=True))
    if not rows:
        return pd.DataFrame()

    headers_raw = list(rows[0])
    # Header A es la fecha → reemplazamos por 'FECHA'
    headers = ["FECHA"] + [str(h) if h is not None else f"COL_{i}"
                           for i, h in enumerate(headers_raw[1:], start=2)]
    # Recortar al ancho real
    n = len(headers)

    data = []
    for r in rows[1:]:
        if all(v is None for v in r[:n]):
            continue
        data.append(list(r[:n]))

    df = pd.DataFrame(data, columns=headers[:len(data[0]) if data else n])
    wb.close()
    return df


@st.cache_data(show_spinner=False)
def load_master_reposicion_ae(file_bytes: bytes) -> pd.DataFrame:
    """Lee la hoja 'Reposición AE' del MASTER.

    Columnas A-J:
        A: Fecha | B: Almacén | C: Descripción | D: Cancha | E: bxp
        F: Posiciones | G: En cancha | H: AGR | I: PICK | J: Reposición Pall
    """
    wb = openpyxl.load_workbook(
        io.BytesIO(file_bytes), data_only=True, read_only=True,
    )
    if "Reposición AE" not in wb.sheetnames:
        raise ValueError("Hoja 'Reposición AE' no encontrada en MASTER")
    ws = wb["Reposición AE"]

    headers = ["Fecha", "Almacen", "Descripcion", "Cancha", "bxp",
               "Posiciones", "En_cancha", "AGR", "PICK", "Reposicion_Pall"]
    data = []
    for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True)):
        if r[1] is None:  # Almacén vacío → fin
            break
        data.append(list(r[:10]))

    df = pd.DataFrame(data, columns=headers)
    df["Almacen"] = pd.to_numeric(df["Almacen"], errors="coerce").astype("Int64")
    for col in ("bxp", "Posiciones", "En_cancha", "AGR", "PICK", "Reposicion_Pall"):
        df[col] = pd.to_numeric(df[col], errors="coerce")
    wb.close()
    return df


@st.cache_data(show_spinner=False)
def load_master_agregados_ae(file_bytes: bytes) -> pd.DataFrame:
    """Placeholder para Tab 5 — Agregados AE.

    Por ahora devuelve un DataFrame derivado de Matriz Pall. con las cols
    de AE (K-Q). Cuando Lucas confirme la hoja exacta destino se ajusta.
    """
    df_pall, _ = load_master_matriz_pall(file_bytes)
    # Subset: Camion + Reparto + UP AE + 6 canchas AE
    cols = ["Camion", "Reparto", "UP_AE_PALL_AGR_FACT",
            "AE_C1", "AE_C2", "AE_C3", "AE_C4", "AE_C5", "AE_MKPL"]
    return df_pall[cols].copy()


# ─── LECTOR T2 Pall. Camiones (referencia estática) ─────────────────────────

@st.cache_data(show_spinner=False)
def load_pall_camiones(file_bytes: bytes) -> pd.DataFrame:
    """Lee la hoja 'Pall. Camiones' del T2_Status_Carga (datos maestros).

    Headers reales A-AF (32 cols). El archivo se carga una sola vez por
    sesión y queda en cache (referencia estática para Tab 3 — Sprint 3).
    """
    wb = openpyxl.load_workbook(
        io.BytesIO(file_bytes), data_only=True, read_only=True,
    )
    if "Pall. Camiones" not in wb.sheetnames:
        raise ValueError("Hoja 'Pall. Camiones' no encontrada en T2_Status")
    ws = wb["Pall. Camiones"]

    rows = list(ws.iter_rows(min_row=1, values_only=True))
    headers = [str(h).strip() if h else f"COL_{i}" for i, h in enumerate(rows[0])]
    # Renombre clave: TARA → CAPACIDAD MÁX (decisión cerrada handoff #9)
    # En esta hoja el header ya es 'Capacidad Maxima (kg)' en col AC, así que
    # el rename solo aplica al PDF rediseñado del Tab 3 — la fuente se respeta.

    data = []
    for r in rows[1:]:
        if r[1] is None:
            continue
        data.append(list(r[:len(headers)]))
    df = pd.DataFrame(data, columns=headers)
    df["Camion"] = pd.to_numeric(df["Camion"], errors="coerce").astype("Int64")
    wb.close()
    return df


# ─── EXTRAÍBLES SHEETS (TAB 5) ──────────────────────────────────────────────

def df_to_tsv(df: pd.DataFrame) -> bytes:
    """TSV: pegar directo en Google Sheets sin separadores ambiguos."""
    return df.to_csv(sep="\t", index=False, lineterminator="\n").encode("utf-8")

def df_to_csv(df: pd.DataFrame) -> bytes:
    return df.to_csv(sep=",", index=False, lineterminator="\n").encode("utf-8")

def df_to_xlsx(df: pd.DataFrame, sheet_name: str = "Hoja1") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as wr:
        df.to_excel(wr, index=False, sheet_name=sheet_name[:31])
    buf.seek(0)
    return buf.read()

def filter_reposicion_negativa(df: pd.DataFrame) -> pd.DataFrame:
    """Tab 5 — filtro decisión cerrada: solo filas con Reposicion_Pall < 0."""
    return df[df["Reposicion_Pall"] < 0].copy().reset_index(drop=True)


# ─── SNAPSHOT LOCAL (Sprint 5 → Drive) ──────────────────────────────────────

def make_run_id() -> str:
    """Run ID único por corrida (timestamp + hash corto)."""
    ts = datetime.now().strftime("%H%M%S")
    h = hashlib.sha1(ts.encode()).hexdigest()[:6]
    return f"{ts}_{h}"

def snapshot_local(run_id: str, files: dict[str, bytes], log_text: str):
    """Guarda los archivos y el log en ./snapshots/YYYY-MM-DD/<run_id>/."""
    fecha = datetime.today().strftime("%Y-%m-%d")
    base = SNAPSHOT_DIR / fecha / run_id
    base.mkdir(parents=True, exist_ok=True)
    for name, data in files.items():
        with open(base / name, "wb") as f:
            f.write(data)
    with open(base / "run.log", "w", encoding="utf-8") as f:
        f.write(log_text)
    return base


# ─── HELPERS UI ─────────────────────────────────────────────────────────────

def _stamp(prefix: str, ext: str) -> str:
    """Nombre de archivo con timestamp para descargas."""
    return f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"

def _download_trio(df: pd.DataFrame, base_name: str, sheet_name: str, key_prefix: str):
    """Renderiza 3 botones de descarga (TSV / CSV / XLSX) en columnas."""
    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button(
            "⬇ TSV (pegar Sheets)",
            data=df_to_tsv(df),
            file_name=_stamp(base_name, "tsv"),
            mime="text/tab-separated-values",
            key=f"{key_prefix}_tsv",
            use_container_width=True,
        )
    with c2:
        st.download_button(
            "⬇ CSV",
            data=df_to_csv(df),
            file_name=_stamp(base_name, "csv"),
            mime="text/csv",
            key=f"{key_prefix}_csv",
            use_container_width=True,
        )
    with c3:
        st.download_button(
            "⬇ XLSX",
            data=df_to_xlsx(df, sheet_name),
            file_name=_stamp(base_name, "xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"{key_prefix}_xlsx",
            use_container_width=True,
        )


# ============================================================================
# ║                            UI STREAMLIT                                  ║
# ============================================================================

def render_sidebar():
    """Sidebar global con toggles y badges de versión."""
    with st.sidebar:
        st.markdown(f"### 🟦 Picking Orchestrator")
        st.caption(f"v{APP_VERSION}")
        st.caption("Beccacece Hnos SA — DPO 2.1")
        st.divider()

        st.session_state["dry_run"] = st.toggle(
            "🧪 Modo Dry-Run",
            value=st.session_state.get("dry_run", False),
            help="Procesa todo pero NO genera PDFs finales (solo preview en pantalla). "
                 "Recomendado durante la primera semana de validación.",
        )
        st.session_state["snapshot_on"] = st.toggle(
            "💾 Snapshot activo",
            value=st.session_state.get("snapshot_on", True),
            help=f"Guarda CAR + MASTER + PDFs + log en {SNAPSHOT_DIR}/<fecha>/<run_id>/ "
                 "cada corrida. Default ON. Local hasta Sprint 5 (después → Drive).",
        )

        st.divider()
        st.caption(f"📅 Hoy: {datetime.today().strftime('%d/%m/%Y')}")
        if st.session_state.get("dry_run"):
            st.warning("🧪 DRY-RUN ACTIVO\n\nNo se generan PDFs finales.")

        st.divider()
        with st.expander("🛠 Stack técnico", expanded=False):
            st.code(
                f"streamlit {st.__version__}\n"
                f"pandas {pd.__version__}\n"
                f"openpyxl {openpyxl.__version__}",
                language="text",
            )


def _file_uploader(label: str, types: list[str], key: str, required: bool = True):
    """Wrapper de st.file_uploader con label rico."""
    suffix = " *" if required else " (opcional)"
    return st.file_uploader(
        label + suffix, type=types, key=key, accept_multiple_files=False,
    )


# ── TAB 1 — Planilla de Carga (REUSO v3.9) ─────────────────────────────────

def render_tab_planilla():
    st.subheader("📦 Planilla de Carga")
    st.caption(
        "Genera el PDF multi-página de la planilla diaria de picking. "
        "Lógica idéntica a la app v3.9.1 validada en producción."
    )

    col_a, col_b = st.columns(2)
    with col_a:
        car_file = _file_uploader(
            "CAR.xlsx (export Chess)", ["xlsx"], key="t1_car"
        )
    with col_b:
        fr_file = _file_uploader(
            "Frescura 3.0.xlsx", ["xlsx"], key="t1_fr"
        )

    if not (car_file and fr_file):
        st.info("Subí ambos archivos para generar la Planilla.")
        return

    if st.button("🚀 Generar Planilla de Carga", type="primary", key="t1_gen"):
        try:
            with st.spinner("Cargando Frescura..."):
                api, ddm, fr, fr_diag = load_frescura(fr_file)
            log_event("info", f"Frescura cargada: API={len(api)} | DDM={len(ddm)} | FR={len(fr)}")

            with st.spinner("Cargando CAR..."):
                car_df, blue_audit = load_car(car_file, ddm=ddm)
            log_event("info", f"CAR cargado: {len(car_df)} filas | bloque azul={len(blue_audit)}")

            if st.session_state.get("dry_run"):
                st.success(f"✓ DRY-RUN OK — {len(car_df)} filas listas. PDF no generado.")
                with st.expander("Preview CAR (primeras 20 filas)"):
                    st.dataframe(car_df.head(20), use_container_width=True)
                return

            with st.spinner("Generando PDF (puede tardar)..."):
                pdf_bytes, stats = generate_pdf(car_df, api, ddm, fr)

            log_event("info", f"Planilla generada: {stats['total_pages']} páginas, "
                              f"{len(stats['repartos'])} repartos")

            st.success(
                f"✓ Planilla lista — {stats['total_pages']} páginas | "
                f"{len(stats['repartos'])} repartos"
            )

            st.download_button(
                "⬇ Descargar Planilla de Carga (PDF)",
                data=pdf_bytes,
                file_name=_stamp("Planilla_Carga", "pdf"),
                mime="application/pdf",
                type="primary",
                use_container_width=True,
            )

            with st.expander(f"📊 Alertas y diagnósticos"):
                a, b, c, d = st.columns(4)
                a.metric("🔴 RED", len(stats["red"]))
                b.metric("🟡 YELLOW", len(stats["yellow"]))
                c.metric("📦 c/Pallet", len(stats["pallet_applied"]))
                d.metric("⚠ Sin BXP", len(stats["miss_bxp"]))
                if stats["orphans_no_ddm"]:
                    st.warning("Huérfanos sin DDM:")
                    for o in stats["orphans_no_ddm"]:
                        st.write(f"- {o}")

            # Stash para snapshot (Sprint 1: solo en memoria)
            st.session_state["last_planilla_pdf"] = pdf_bytes

        except Exception as e:
            log_event("error", f"Error en Planilla: {e}")
            st.error(f"❌ Error generando Planilla: {e}")
            with st.expander("Stack trace"):
                import traceback
                st.code(traceback.format_exc())


# ── TAB 2 — Resumen por Camión (REUSO v3.9) ────────────────────────────────

def render_tab_resumen():
    st.subheader("📋 Resumen de Carga por Camión")
    st.caption(
        "Genera un PDF compacto con bultos consolidados por SKU por camión "
        "(hoja AGR del CAR). Útil para tildar al final de la carga."
    )

    car_file = _file_uploader(
        "CAR.xlsx (export Chess — debe tener hoja AGR)", ["xlsx"], key="t2_car"
    )

    if not car_file:
        st.info("Subí el CAR.xlsx con la hoja AGR para generar el Resumen.")
        return

    if st.button("🚀 Generar Resumen por Camión", type="primary", key="t2_gen"):
        try:
            if st.session_state.get("dry_run"):
                with st.spinner("Leyendo hoja AGR..."):
                    df_agr = load_agr(car_file)
                st.success(
                    f"✓ DRY-RUN OK — {len(df_agr)} filas SKU/camión, "
                    f"{df_agr['chofer'].nunique()} camiones."
                )
                st.dataframe(df_agr.head(30), use_container_width=True)
                return

            with st.spinner("Generando PDF resumen..."):
                pdf_bytes, stats = build_resumen_carga_pdf(car_file)

            log_event("info", f"Resumen generado: {stats['camiones']} camiones, "
                              f"{stats['filas_sku']} filas SKU, "
                              f"{stats['total_pages']} páginas")

            st.success(
                f"✓ Resumen listo — {stats['camiones']} camiones | "
                f"{stats['filas_sku']} filas SKU | {stats['total_pages']} páginas"
            )

            st.download_button(
                "⬇ Descargar Resumen por Camión (PDF)",
                data=pdf_bytes,
                file_name=_stamp("Resumen_Camiones", "pdf"),
                mime="application/pdf",
                type="primary",
                use_container_width=True,
            )

            st.session_state["last_resumen_pdf"] = pdf_bytes

        except Exception as e:
            log_event("error", f"Error en Resumen: {e}")
            st.error(f"❌ Error generando Resumen: {e}")
            with st.expander("Stack trace"):
                import traceback
                st.code(traceback.format_exc())


# ── TAB 3 — Camiones T2 (PLACEHOLDER, Sprint 3) ────────────────────────────

def render_tab_t2():
    st.subheader("🚛 Camiones T2 — Hojas de Carga")
    st.caption(
        "Reemplaza el Sheet T2 Status Carga. Lee MASTER → Matriz Pall. "
        "y genera 1 PDF batch con N páginas (1 por camión con Reparto=SI y TOTAL PALL>0). "
        "**Sprint 3** — pendiente."
    )
    st.info(
        "🔧 Esta tab está en desarrollo (Sprint 3).\n\n"
        "**Alcance previsto:**\n"
        "- Input: MASTER + T2_Status (referencia estática Pall. Camiones)\n"
        "- Lógica plano del camión: 6×10 celdas, fórmulas portadas (sección 9 handoff)\n"
        "- Rediseño visual profesional (NO réplica del Sheet actual — decisión #3)\n"
        "- Output: PDF A4 horizontal, ~17 páginas\n"
        "- Lógica simplificada: lee 1 sola columna por cancha (decisión #8 — sin E+Z / L+AG)"
    )

    with st.expander("📐 Vista previa de datos disponibles", expanded=False):
        st.markdown("Subí MASTER y T2_Status para inspeccionar los datos que va a usar Tab 3:")
        master = _file_uploader(
            "MASTER del día.xlsx (78 MB)", ["xlsx"], key="t3_master", required=False
        )
        t2 = _file_uploader(
            "T2_Status_Carga (referencia)", ["xlsx"], key="t3_t2", required=False
        )

        if master:
            try:
                df_pall, fecha = load_master_matriz_pall(master.getvalue())
                st.success(f"Matriz Pall. cargada: {len(df_pall)} camiones | Fecha A1: {fecha}")

                # Filtro preview: candidatos a Tab 3
                df_si = df_pall[
                    (df_pall["Reparto"] == "SI")
                    & ((df_pall["UP_PICKING_AGR_FACT"] + df_pall["UP_AE_PALL_AGR_FACT"]) > 0)
                ]
                st.info(
                    f"Camiones que entrarían al PDF T2 (Reparto=SI ∧ TOTAL>0): "
                    f"**{len(df_si)}** de {len(df_pall)}"
                )
                st.dataframe(df_pall.head(15), use_container_width=True)
            except Exception as e:
                st.error(f"Error leyendo MASTER: {e}")

        if t2:
            try:
                df_cam = load_pall_camiones(t2.getvalue())
                st.success(f"Pall. Camiones cargado: {len(df_cam)} camiones maestros")
                st.dataframe(df_cam.head(10), use_container_width=True)
            except Exception as e:
                st.error(f"Error leyendo T2_Status: {e}")


# ── TAB 4 — Proyección Picking ×4 (PLACEHOLDER, Sprint 4) ──────────────────

def render_tab_proyeccion():
    st.subheader("📊 Proyección Picking ×4")
    st.caption(
        "Editor interactivo de la proyección de picking + PDF con 4 páginas "
        "idénticas (1 copia por cancha). **Sprint 4** — pendiente."
    )
    st.info(
        "🔧 En desarrollo (Sprint 4).\n\n"
        "**Alcance previsto:**\n"
        "- Input: MASTER → hoja `Proyección Picking`\n"
        "- Editor `st.data_editor` para reasignar carga entre canchas\n"
        "- Output: PDF con 4 páginas idénticas (CANCHA I, II, III, IV)\n"
        "- Lucas valida balance antes de imprimir"
    )


# ── TAB 5 — Extraíbles Sheets ──────────────────────────────────────────────

def render_tab_extraibles():
    st.subheader("📤 Extraíbles para Google Sheets")
    st.caption(
        "Bajá los bloques del MASTER en formato TSV/CSV/XLSX para pegar manual "
        "en el Ecosistema AD 3.0. Sin push automático en v4.0 (decisión #2)."
    )

    master = _file_uploader(
        "MASTER del día.xlsx (78 MB)", ["xlsx"], key="t5_master"
    )

    if not master:
        st.info("Subí el MASTER para acceder a los 4 bloques.")
        return

    file_bytes = master.getvalue()

    try:
        df_pall, fecha_a1 = load_master_matriz_pall(file_bytes)
        df_rep = load_master_reposicion_ae(file_bytes)
        df_pick = load_master_matriz_picking_app(file_bytes)
        df_agr = load_master_agregados_ae(file_bytes)
        log_event("info", f"MASTER cargado para extraíbles. Fecha A1: {fecha_a1}")
    except Exception as e:
        st.error(f"❌ Error leyendo MASTER: {e}")
        return

    fecha_str = fecha_a1.strftime("%d/%m/%Y") if fecha_a1 else "—"
    st.success(f"✓ MASTER cargado — Fecha en A1: **{fecha_str}**")

    st.divider()

    # ── Bloque 1: Matriz Pall. ──────────────────────────────────────────
    with st.expander("1️⃣  Matriz Pall. (29 camiones × 24 cols)", expanded=True):
        st.dataframe(df_pall, use_container_width=True, height=300)
        _download_trio(df_pall, "Matriz_Pall", "Matriz Pall.", "t5_pall")

    # ── Bloque 2: Agregados AE ──────────────────────────────────────────
    with st.expander("2️⃣  Agregados AE (subset Matriz Pall. — cols AE)"):
        st.caption(
            "⚠ Provisional: deriva de Matriz Pall. cols K-Q. "
            "Cuando confirmes hoja destino exacta en Ecosistema AD 3.0, se ajusta."
        )
        st.dataframe(df_agr, use_container_width=True, height=300)
        _download_trio(df_agr, "Agregados_AE", "Agregados AE", "t5_agr")

    # ── Bloque 3: Reposición AE (filtrado J<0) ──────────────────────────
    with st.expander("3️⃣  Reposición AE (solo Reposicion_Pall < 0)", expanded=True):
        df_rep_neg = filter_reposicion_negativa(df_rep)
        st.metric("Filas con reposición negativa", f"{len(df_rep_neg)} / {len(df_rep)}")
        if len(df_rep_neg) == 0:
            st.warning("No hay filas con Reposicion_Pall < 0. Nada para pegar hoy.")
        else:
            st.dataframe(df_rep_neg, use_container_width=True, height=300)
            _download_trio(df_rep_neg, "Reposicion_AE_neg", "Reposición AE", "t5_rep")

    # ── Bloque 4: (fx) Picking ──────────────────────────────────────────
    with st.expander("4️⃣  (fx) Picking — fila resumen Matriz Picking APP"):
        st.caption(
            "Vista plana de la hoja Matriz Picking APP. Para Sprint 1 se exporta "
            "completa; en Sprint 2 se filtra/recorta según destino exacto."
        )
        st.dataframe(df_pick.head(50), use_container_width=True, height=300)
        _download_trio(df_pick, "fx_Picking", "(fx) Picking", "t5_pick")


# ── TAB 6 — Validación + Log (PLACEHOLDER, Sprint 2) ───────────────────────

def render_tab_validacion():
    st.subheader("✅ Validación + Reglas + Log")
    st.caption(
        "Pre-checks + reglas de negocio + validación cruzada CAR↔MASTER. "
        "**Sprint 2** — pendiente."
    )

    st.info(
        "🔧 En desarrollo (Sprint 2).\n\n"
        "**Pre-checks previstos:**\n"
        "- CAR.xlsx modificado hoy\n"
        "- ANR (rechazos1.xlsx) modificado hoy (warning)\n"
        "- MASTER → Matriz Pall.!A1 == hoy\n"
        "- Suma TOTAL PALL > 0\n\n"
        "**Reglas:**\n"
        "- Camión TOTAL=0 → ocultar y loguear\n"
        "- Camión Reparto=NO → tabla checkbox, default OFF (decisión #1)\n"
        "- Cross-validate CAR↔MASTER (WARNING, no bloqueante — decisión #6)\n"
        "- Sanity checks bloqueantes pre-publicación"
    )

    st.divider()
    st.markdown("### 📜 Log de la corrida actual")
    log_lines = st.session_state.get("log_buffer", [])
    if not log_lines:
        st.caption("(sin eventos aún)")
    else:
        st.code("\n".join(log_lines[-200:]), language="text")
        st.download_button(
            "⬇ Descargar log completo",
            data="\n".join(log_lines).encode("utf-8"),
            file_name=_stamp("run_log", "txt"),
            mime="text/plain",
        )


# ── MAIN ────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title=f"Picking Orchestrator v{APP_VERSION}",
        page_icon="📦",
        layout="wide",
    )

    # Banner header
    st.markdown(
        f"""
        <div style="
            background: linear-gradient(135deg, #1a3a6b 0%, #2e5fa3 100%);
            padding: 14px 22px;
            border-radius: 8px;
            margin-bottom: 18px;
            color: white;
        ">
            <div style="font-size: 22px; font-weight: 600;">
                📦 Picking Orchestrator
            </div>
            <div style="font-size: 13px; opacity: 0.85; margin-top: 2px;">
                Beccacece Hnos SA — Distribuidor CMQ &nbsp;·&nbsp; DPO 2.1 — Pilar Almacén
                &nbsp;·&nbsp; v{APP_VERSION}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    render_sidebar()

    tabs = st.tabs([
        "📦 Planilla Carga",
        "📋 Resumen Camiones",
        "🚛 Camiones T2",
        "📊 Proyección Picking ×4",
        "📤 Extraíbles Sheets",
        "✅ Validación + Log",
    ])
    with tabs[0]: render_tab_planilla()
    with tabs[1]: render_tab_resumen()
    with tabs[2]: render_tab_t2()
    with tabs[3]: render_tab_proyeccion()
    with tabs[4]: render_tab_extraibles()
    with tabs[5]: render_tab_validacion()


if __name__ == "__main__":
    main()
