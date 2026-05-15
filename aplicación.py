import streamlit as st
import pandas as pd
import io
import re
import unicodedata
from pathlib import Path

# Librerías para el PDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors

# --- CONFIGURACIÓN DE ESTILO ---
AZUL_OSCURO = colors.HexColor("#1F4E79")
AZUL_CLARO  = colors.HexColor("#DCEBF7")
BORDE       = colors.HexColor("#9BBBD9")
NEGRO       = colors.black

st.set_page_config(page_title="Validador Trimestral 2026", layout="wide")

# --- FUNCIONES DE UTILIDAD ---
def _norm(s: str) -> str:
    if s is None: return ""
    s = str(s).strip().lower()
    return "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))

def es_muni(texto: str) -> bool:
    t = _norm(texto)
    return any(p in t for p in ["municip", "gobierno local", "alcald", "ayuntamiento", "gl"])

# --- MOTOR DE EXTRACCIÓN ---
def extraer_datos_trimestre(file, tri_objetivo):
    """
    Procesa el archivo y extrae solo las columnas del trimestre seleccionado.
    """
    df = pd.read_excel(file, sheet_name='Informe de avance', header=None)
    
    # Mapeo de columnas basado en la estructura del .xlsm
    # T1: 14, T2: 19, T3: 24, T4: 29
    mapeo = {
        "T1": {"avance": 14, "desc": 15, "cant": 16},
        "T2": {"avance": 19, "desc": 20, "cant": 21},
        "T3": {"avance": 24, "desc": 25, "cant": 26},
        "T4": {"avance": 29, "desc": 30, "cant": 31}
    }
    
    col_idx = mapeo[tri_objetivo]
    registros = []
    
    # Datos de cabecera
    delegacion = str(df.iloc[1, 7]).strip() if pd.notna(df.iloc[1, 7]) else "No especificada"
    
    # Variables de arrastre
    linea_actual = ""
    problematica_actual = ""
    lider_actual = ""

    for i in range(len(df)):
        row = df.iloc[i]
        row_str = " ".join(row.astype(str)).lower()

        # Detección de bloques
        if "linea de accion #" in row_str:
            linea_actual = str(row[3]).strip()
        if "problematica" in row_str:
            problematica_actual = str(row[5]).strip()
        if "lider estrategico" in row_str:
            lider_actual = str(row[8]).strip()

        # Fila de indicador
        indicador = str(row[4]).strip()
        if indicador and "indicador" in indicador.lower():
            avance = str(row[col_idx["avance"]]).strip()
            detalle = str(row[col_idx["desc"]]).strip()

            # Filtrar: Solo Municipalidad y solo si hay datos en este trimestre
            if es_muni(lider_actual) and (avance != "nan" or detalle != "nan"):
                registros.append({
                    "Delegacion": delegacion,
                    "Linea": linea_actual,
                    "Lider": "Gobierno Local",
                    "Indicador": indicador,
                    "Meta": str(row[8]).strip(),
                    "Avance (%)": avance if avance != "nan" else "0%",
                    "Resultado": detalle if detalle != "nan" else "Sin reporte",
                    "Trimestre": tri_objetivo
                })
    
    return pd.DataFrame(registros)

# --- GENERADOR DE PDF ---
def crear_pdf(df, tri, nombre_muni):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    # Portada
    c.setFillColor(AZUL_OSCURO)
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(w/2, h - 10*cm, "INFORME DE VALIDACIÓN TRIMESTRAL")
    c.setFont("Helvetica", 16)
    c.drawCentredString(w/2, h - 11.5*cm, f"Gobierno Local de {nombre_muni}")
    c.drawCentredString(w/2, h - 12.5*cm, f"Seguimiento: {tri} - 2026")
    c.showPage()

    # Cuerpo del reporte
    y = h - 2.5*cm
    for idx, row in df.iterrows():
        if y < 6*cm:
            c.showPage()
            y = h - 2.5*cm

        # Recuadro de información
        c.setStrokeColor(BORDE)
        c.rect(1.5*cm, y-5*cm, w-3*cm, 4.8*cm)
        
        c.setFillColor(AZUL_CLARO)
        c.rect(1.5*cm, y-0.8*cm, w-3*cm, 0.8*cm, fill=1, stroke=0)
        
        c.setFillColor(AZUL_OSCURO)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(1.7*cm, y-0.5*cm, f"ACCIÓN ESTRATÉGICA {idx+1}")

        c.setFillColor(NEGRO)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(1.7*cm, y-1.4*cm, "Línea:")
        c.setFont("Helvetica", 9)
        c.drawString(3.5*cm, y-1.4*cm, row['Linea'][:80])

        c.setFont("Helvetica-Bold", 9)
        c.drawString(1.7*cm, y-2*cm, "Indicador:")
        c.setFont("Helvetica", 9)
        c.drawString(3.5*cm, y-2*cm, row['Indicador'][:100])

        # Espacio de Resultado
        c.setFillColor(colors.whitesmoke)
        c.rect(1.7*cm, y-4.6*cm, w-3.4*cm, 2.2*cm, fill=1, stroke=1)
        c.setFillColor(AZUL_OSCURO)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(2*cm, y-3*cm, f"RESULTADO {tri}: (Avance: {row['Avance (%)']})")
        
        c.setFillColor(NEGRO)
        c.setFont("Helvetica", 9)
        res_text = row['Resultado']
        text_obj = c.beginText(2*cm, y-3.6*cm)
        for i in range(0, len(res_text), 90):
            text_obj.textLine(res_text[i:i+90])
        c.drawText(text_obj)

        y -= 5.5*cm

    c.save()
    buf.seek(0)
    return buf

# --- INTERFAZ ---
st.title("📑 Validador de Líneas Estratégicas 2026")
st.markdown("Herramienta para procesar informes trimestrales individuales y validar resultados municipales.")

with st.sidebar:
    st.header("Opciones")
    tri_seleccionado = st.selectbox("Seleccionar Trimestre", ["T1", "T2", "T3", "T4"])
    st.info("El sistema solo extraerá datos donde el Líder sea Municipal.")

archivo_excel = st.file_uploader("Cargar Informe de Avance (.xlsm)", type=["xlsm"])

if archivo_excel:
    # Detectar cantón por nombre de archivo
    muni_name = archivo_excel.name.split(" - ")[0]
    
    with st.spinner("Procesando datos del trimestre..."):
        df_tri = extraer_datos_trimestre(archivo_excel, tri_seleccionado)

    if not df_tri.empty:
        st.success(f"Éxito: {len(df_tri)} registros municipales encontrados.")
        
        st.subheader(f"Vista previa de Validación - {tri_seleccionado}")
        st.dataframe(df_tri[["Linea", "Indicador", "Avance (%)", "Resultado"]], use_container_width=True)

        if st.button("Generar PDF de Validación"):
            pdf_out = crear_pdf(df_tri, tri_seleccionado, muni_name)
            st.download_button(
                label="⬇️ Descargar Reporte PDF",
                data=pdf_out,
                file_name=f"Validacion_{tri_seleccionado}_{muni_name}.pdf",
                mime="application/pdf"
            )
    else:
        st.warning(f"No se detectaron reportes de la Municipalidad para el {tri_seleccionado} en este archivo.")
