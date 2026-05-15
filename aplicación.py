import io, re, unicodedata
from typing import Dict, List, Optional
from pathlib import Path

import streamlit as st
import pandas as pd

# PDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader

# ===== Estilo =====
AZUL_OSCURO = colors.HexColor("#1F4E79")
AZUL_CLARO  = colors.HexColor("#DCEBF7")
BORDE       = colors.HexColor("#9BBBD9")
NEGRO       = colors.black
FF_MULTILINE = 4096
MAXLEN_MUY_GRANDE = 100000

st.set_page_config(page_title="PDF editable – Gobierno Local", layout="wide")
st.title("Generar PDF editable – Seguimiento Trimestral")

# ========= Utils =========
def _norm(s: str) -> str:
    if s is None: return ""
    s = str(s).strip().lower()
    return "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))

def es_muni(texto: str) -> bool:
    t = _norm(texto)
    return any(p in t for p in ["municip", "gobierno local", "alcald", "ayuntamiento"])

# --- LÓGICA DE MAPEO DE TRIMESTRES ---
# Según tu archivo, los trimestres están en columnas específicas
MAPEO_TRIMESTRES = {
    "T1": {"col_base": 14, "nombre": "I Trimestre"},
    "T2": {"col_base": 19, "nombre": "II Trimestre"},
    "T3": {"col_base": 24, "nombre": "III Trimestre"},
    "T4": {"col_base": 29, "nombre": "IV Trimestre"}
}

# ========= Parser Modificado =========
def parse_sheet_filtered(df_raw: pd.DataFrame, sheet_name: str, tri_key: str) -> pd.DataFrame:
    S = df_raw.astype(str).where(~df_raw.isna(), "")
    nrows, ncols = S.shape
    
    # Obtener el índice de columna para el trimestre seleccionado
    col_trimestre = MAPEO_TRIMESTRES[tri_key]["col_base"]
    
    current_problem, current_linea = "", ""
    last_action, last_meta, last_lider = "", "", ""
    
    rows: List[Dict] = []
    
    # Palabras clave para detectar secciones (basado en tu código)
    RE_CADENA = re.compile(r"^cadena\s+de\s+resultados", re.I)
    RE_LINEA  = re.compile(r"^l[ií]nea\s+de\s+acci[oó]n\s*#?\s*\d*", re.I)

    for i in range(nrows):
        row_vals = [S.iat[i, j].strip() for j in range(ncols)]
        row_str = " ".join(row_vals).lower()

        # 1. Detectar Contexto
        for cell in row_vals:
            if RE_CADENA.match(cell):
                current_problem = cell.split(":",1)[-1].strip()
            if RE_LINEA.match(cell):
                current_linea = cell.strip()

        # 2. Detectar Líder y Acción (Arrastre)
        # Si la fila tiene "Líder Estratégico" (Columna 8 en tu Excel)
        if "lider estrategico" in row_str:
            last_lider = S.iat[i, 8].strip()
            continue

        # 3. Detectar Fila de Indicador y extraer RESULTADO del trimestre
        indicador = S.iat[i, 4].strip() # Columna 4 según tu archivo
        if indicador and "indicador" in indicador.lower():
            # Extraemos el valor del avance en la columna del trimestre seleccionado
            avance_tri = S.iat[i, col_trimestre].strip()
            detalle_tri = S.iat[i, col_trimestre + 1].strip() # Descripción suele estar al lado

            # Solo guardamos si el líder es municipal
            if es_muni(last_lider):
                rows.append({
                    "problematica": current_problem,
                    "linea_accion": current_linea,
                    "accion_estrategica": S.iat[i, 3].strip() or "Ver línea de acción",
                    "indicador": indicador,
                    "meta": S.iat[i, 8].strip(), # Meta en columna 8
                    "lider": "Municipalidad",
                    "resultado_previo": f"{avance_tri} - {detalle_tri}",
                    "hoja": sheet_name
                })

    return pd.DataFrame(rows)

# ========= PDF y UI (Se mantienen tus funciones pero con la data filtrada) =========

with st.sidebar:
    st.header("Filtro de Seguimiento")
    trimestre_sel = st.selectbox("Seleccione el Trimestre a validar", ["T1", "T2", "T3", "T4"])
    st.info(f"El reporte mostrará solo datos de {MAPEO_TRIMESTRES[trimestre_sel]['nombre']}")

excel_file = st.file_uploader("Subí tu Excel (.xlsm)", type=["xlsm"])

if excel_file:
    # Procesamos el libro usando la nueva función de filtrado
    xls = pd.ExcelFile(excel_file)
    hoja = "Informe de avance" # Forzamos la hoja que tiene la matriz
    
    df_raw = pd.read_excel(excel_file, sheet_name=hoja, header=None)
    regs_muni = parse_sheet_filtered(df_raw, hoja, trimestre_sel)

    if not regs_muni.empty:
        st.subheader(f"Vista Previa: Datos detectados para {trimestre_sel}")
        st.dataframe(regs_muni, use_container_width=True)
        
        # Aquí llamarías a tu función build_pdf_grouped_by_problem pasándole regs_muni
        # El PDF ahora solo tendrá las fichas de ese trimestre.
    else:
        st.error("No se encontraron datos municipales para este trimestre. Verifique el archivo.")
