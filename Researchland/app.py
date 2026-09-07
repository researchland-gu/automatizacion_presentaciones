# ==============================================================================
# APP WEB DEL REPORTE RESEARCHLAND
#
# Cómo correrla (desde la terminal de PyCharm, dentro de la carpeta Researchland):
#     pip install -r requirements.txt
#     streamlit run app.py
#
# Se abrirá la URL http://localhost:8501 en tu navegador, donde seleccionas
# dinámicamente FILE_ACTUAL, FILE_ANTERIOR, el periodo (solo para el nombre
# de salida) y OUTPUT_PPTX, y generas el reporte con un clic.
# ==============================================================================
import io
import os
import tempfile
from contextlib import redirect_stdout

import streamlit as st

import config
import main as pipeline

st.set_page_config(page_title="Reporte Researchland", page_icon="📊", layout="centered")

st.title("📊 Generador de Reporte Researchland")
st.caption("Selecciona las bases y el periodo; el reporte se genera con un clic.")


def _guardar_upload(uploaded_file):
    """Guarda un archivo subido en un temporal y devuelve su ruta."""
    sufijo = os.path.splitext(uploaded_file.name)[1] or '.xlsx'
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=sufijo)
    tmp.write(uploaded_file.getbuffer())
    tmp.close()
    return tmp.name


# ------------------------------------------------------------------------------
# 1. SELECCIÓN DE BASES (subir archivo o escribir ruta local)
# ------------------------------------------------------------------------------
st.header("1. Bases de resultados")

modo = st.radio(
    "¿Cómo quieres indicar las bases?",
    ["Subir archivos", "Escribir ruta local"],
    horizontal=True,
    help="'Subir archivos' abre un selector; 'Ruta local' usa archivos ya guardados en tu disco (más rápido para bases grandes).",
)

ruta_actual = None
ruta_anterior = None

if modo == "Subir archivos":
    up_actual = st.file_uploader("Base del trimestre ACTUAL (FILE_ACTUAL)", type=['xlsx', 'xlsm'], key='up_act')
    up_anterior = st.file_uploader("Base del trimestre ANTERIOR (FILE_ANTERIOR)", type=['xlsx', 'xlsm'], key='up_ant')
    if up_actual is not None:
        ruta_actual = _guardar_upload(up_actual)
        st.success(f"Base actual cargada: {up_actual.name}")
    if up_anterior is not None:
        ruta_anterior = _guardar_upload(up_anterior)
        st.success(f"Base anterior cargada: {up_anterior.name}")
else:
    ruta_actual = st.text_input("Ruta de la base ACTUAL (FILE_ACTUAL)", value=config.FILE_ACTUAL)
    ruta_anterior = st.text_input("Ruta de la base ANTERIOR (FILE_ANTERIOR)", value=config.FILE_ANTERIOR)
    for etiqueta, ruta in [("actual", ruta_actual), ("anterior", ruta_anterior)]:
        if ruta and not os.path.exists(ruta):
            st.warning(f"⚠ La ruta de la base {etiqueta} no existe todavía en este equipo.")

# ------------------------------------------------------------------------------
# 2. PERIODO Y ARCHIVO DE SALIDA
# ------------------------------------------------------------------------------
st.header("2. Periodo y salida")

col1, col2 = st.columns(2)
with col1:
    trimestre = st.selectbox("Trimestre", ["1Q", "2Q", "3Q", "4Q"], index=1)
    anio = st.number_input("Año", min_value=2020, max_value=2050, value=2026, step=1)
    nuevo_periodo = f"{trimestre}´{str(anio)[-2:]}"
    st.markdown(f"**Etiqueta del periodo:** `{nuevo_periodo}`")
with col2:
    output_pptx = st.text_input(
        "Nombre del archivo de salida (OUTPUT_PPTX)",
        value=f"Reporte Researchland_{trimestre} {anio}.pptx",
    )
    if output_pptx and not output_pptx.lower().endswith('.pptx'):
        output_pptx += '.pptx'

with st.expander("Opciones avanzadas (plantilla)"):
    template_pptx = st.text_input("Ruta de la plantilla (TEMPLATE_PPTX)", value=config.TEMPLATE_PPTX)
    st.caption(f"El histórico del slide 6 actualiza el Excel: {config.FILE_HISTORICO_RL}")

# ------------------------------------------------------------------------------
# 3. GENERAR REPORTE
# ------------------------------------------------------------------------------
st.header("3. Generar")

listo = bool(ruta_actual) and bool(ruta_anterior) and bool(output_pptx)
if not listo:
    st.info("Indica ambas bases y el nombre de salida para habilitar el botón.")

if st.button("🚀 Generar reporte", type="primary", disabled=not listo):
    log_buffer = io.StringIO()
    try:
        with st.spinner("Procesando bases e inyectando datos en la presentación..."):
            with redirect_stdout(log_buffer):
                ruta_salida, errores = pipeline.ejecutar_reporte(
                    file_actual=ruta_actual,
                    file_anterior=ruta_anterior,
                    output_pptx=output_pptx,
                    nuevo_periodo=nuevo_periodo,
                    template_pptx=template_pptx,
                )

        if errores:
            st.warning(f"Reporte generado con {len(errores)} bloque(s) con error: {', '.join(errores)}")
        else:
            st.success("✔ Reporte generado sin errores.")

        with open(ruta_salida, 'rb') as f:
            st.download_button(
                "⬇ Descargar presentación",
                data=f.read(),
                file_name=os.path.basename(ruta_salida),
                mime='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            )
    except Exception as e:
        st.error(f"❌ El proceso falló: {e}")

    with st.expander("Ver log completo del proceso", expanded=bool(log_buffer.getvalue())):
        st.code(log_buffer.getvalue() or "(sin salida)", language=None)
