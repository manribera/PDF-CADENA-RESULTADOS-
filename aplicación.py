import streamlit as st
import pandas as pd
import io, re, unicodedata
from pathlib import Path

# PDF Libraries
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors

# ===== CONFIGURACIÓN DE ESTILO =====
AZUL_OSCURO = colors.HexColor("#1F4E79")
AZUL_CLARO  = colors.HexColor("#DCEBF7")
BORDE       = colors.HexColor("#9BBBD9")
NEGRO       = colors.black

st.set_page_config(page_title="Validador de Informes 2026", layout="wide")

# ========= UTILIDADES Y LIMPIEZA =========
def _norm(s: str) -> str:
    if s is None: return ""
    s = str(s).strip().lower()
    return "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))

def es_muni(texto: str) -> bool:
    t = _norm(texto)
    return any(p in t for p in ["municip", "gobierno local", "alcald", "ayuntamiento"])

# ========= LÓGICA DE EXTRACCIÓN POR TRIMESTRE =========
def parse_trimestre_especifico(file, tri_objetivo):
    """
    Extrae datos del Excel basándose en el trimestre seleccionado.
    T1: Col 14, T2: Col 19, T3: Col 24, T4: Col 29 (Índices aproximados del archivo)
    """
    df = pd.read_excel(file, sheet_name='Informe de avance', header=None)
    
    # Mapeo de columnas según la estructura de tu archivo .xlsm
    mapeo = {
        "T1": {"avance": 14, "desc": 15},
        "T2": {"avance": 19, "desc": 20},
        "T3": {"avance": 24, "desc": 25},
        "T4": {"avance": 29, "desc": 30}
    }
    
    col_idx = mapeo[tri_objetivo]
    rows = []
    
    # Variables para arrastrar datos de celdas combinadas
    curr_linea = ""
    curr_lider = ""
    delegacion = str(df.iloc[1, 7]).strip() # D39-La Unión o similar

    for i in range(len(df)):
        row_vals = df.iloc[i].astype(str).tolist()
        row_str = " ".join(row_vals).lower()

        # Identificar secciones
        if "linea de accion #" in row_str:
            curr_linea = str(df.iloc[i, 3]).strip()
        if "lider estrategico" in row_str:
            curr_lider = str(df.iloc[i, 8]).strip()

        # Detectar fila de Indicador
        indicador = str(df.iloc[i, 4]).strip()
        if indicador and "indicador" in indicador.lower():
            avance_val = str(df.iloc[i, col_idx["avance"]]).strip()
            desc_val = str(df.iloc[i, col_idx["desc"]]).strip()

            # SOLO AGREGAR SI PERTENECE A LA MUNICIPALIDAD Y TIENE DATOS EN EL TRIMESTRE
            if es_muni(curr_lider) and (avance_val != "nan" or desc_val != "nan"):
                rows.append({
                    "Delegacion": delegacion,
                    "Linea": curr_linea,
                    "Lider": "Municipalidad",
                    "Indicador": indicador,
                    "Meta": str(df.iloc[i, 8]).strip(),
                    "Avance": avance_val if avance_val != "nan" else "0%",
                    "Resultado": desc_val if desc_val != "nan" else "Sin descripción",
                    "Trimestre": tri_objetivo
                })
    
    return pd.DataFrame(rows)

# ========= GENERACIÓN DE PDF =========
def generar_pdf_reporte(df, tri, canton):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    # Portada
    c.setFillColor(AZUL_OSCURO)
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(w/2, h - 10*cm, "INFORME DE VALIDACIÓN ESTRATÉGICA")
    c.setFont("Helvetica", 16)
    c.drawCentredString(w/2, h - 11.5*cm, f"Gobierno Local de {canton}")
    c.drawCentredString(w/2, h - 12.5*cm, f"Periodo: {tri} - 2026")
    c.showPage()

    # Contenido
    y = h - 3*cm
    for idx, row in df.iterrows():
        if y < 6*cm:
            c.showPage()
            y = h - 3*cm

        # Caja de Acción
        c.setStrokeColor(BORDE)
        c.rect(1.5*cm, y-5*cm, w-3*cm, 4.5*cm)
        
        c.setFillColor(AZUL_CLARO)
        c.rect(1.5*cm, y-0.8*cm, w-3*cm, 0.8*cm, fill=1, stroke=0)
        
        c.setFillColor(AZUL_OSCURO)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(1.7*cm, y-0.5*cm, f"INDICADOR {idx+1}")

        c.setFillColor(NEGRO)
        c.setFont("Helvetica", 9)
        c.drawString(1.7*cm, y-1.5*cm, f"Línea: {row['Linea'][:100]}")
        c.drawString(1.7*cm, y-2.1*cm, f"Meta: {row['Meta'][:110]}")
        
        # Bloque de Resultado
        c.setFillColor(colors.whitesmoke)
        c.rect(1.7*cm, y-4.6*cm, w-3.4*cm, 2.2*cm, fill=1, stroke=1)
        c.setFillColor(AZUL_OSCURO)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(2*cm, y-2.9*cm, f"AVANCE {tri}: {row['Avance']}")
        c.setFillColor(NEGRO)
        c.setFont("Helvetica", 9)
        
        # Ajuste de texto para el resultado (multilínea simple)
        text_obj = c.beginText(2*cm, y-3.5*cm)
        lines = [row['Resultado'][i:i+85] for i in range(0, len(row['Resultado']), 85)]
        for line in lines[:3]:
            text_obj.textLine(line)
        c.drawText(text_obj)

        y -= 5.5*cm

    c.save()
    buf.seek(0)
    return buf

# ========= INTERFAZ DE USUARIO (UI) =========
st.title("🚀 Sistema de Seguimiento Sembremos Seguridad")
st.markdown("Filtra y genera reportes de validación específicos por trimestre.")

with st.sidebar:
    st.header("Configuración")
    tri_sel = st.selectbox("Trimestre a validar", ["T1", "T2", "T3", "T4"])
    st.warning("El reporte solo incluirá acciones del Líder Municipal.")

uploaded_file = st.file_uploader("Subir archivo Excel (.xlsm)", type=["xlsm"])

if uploaded_file:
    # Extraer el nombre del cantón del archivo
    canton = uploaded_file.name.split(" - ")[0]
    
    with st.spinner("Procesando trimestre seleccionado..."):
        df_tri = parse_trimestre_especifico(uploaded_file, tri_sel)

    if not df_tri.empty:
        st.success(f"Se encontraron {len(df_tri)} registros para el {tri_sel}")
        
        st.subheader(f"Vista Previa - {tri_sel} ({canton})")
        st.dataframe(df_tri[["Linea", "Indicador", "Avance", "Resultado"]], use_container_width=True)

        # Botón para generar y descargar el PDF
        if st.button("Generar Reporte PDF"):
            pdf = generar_pdf_reporte(df_tri, tri_sel, canton)
            st.download_button(
                label="⬇️ Descargar PDF de Validación",
                data=pdf,
                file_name=f"Reporte_{tri_sel}_{canton}.pdf",
                mime="application/pdf"
            )
    else:
        st.error(f"No se encontraron datos de avance municipal para el {tri_sel}. Verifique el archivo.")
