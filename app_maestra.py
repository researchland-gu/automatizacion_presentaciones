# ==============================================================================
# APP MAESTRA DE REPORTES IPN COLABORADORES (VERSIÓN BLINDADA + DINÁMICA)
# ==============================================================================
import io
import os
import sys
import tempfile
import importlib.util
from contextlib import contextmanager, redirect_stdout
import streamlit as st

RAIZ = os.path.dirname(os.path.abspath(__file__))
CARPETA_PLANTILLAS = os.path.join(RAIZ, "Plantillas_pptx")


# ------------------------------------------------------------------------------
# CONTEXT MANAGER PARA DIRECTORIO DE TRABAJO
# ------------------------------------------------------------------------------
@contextmanager
def cambiar_directorio(ruta_destino):
    """Garantiza que rutas relativas en calculos, inyectores o main funcionen."""
    antiguo_dir = os.getcwd()
    os.chdir(ruta_destino)
    try:
        yield
    finally:
        os.chdir(antiguo_dir)


# ------------------------------------------------------------------------------
# DESCUBRIMIENTO Y CARGA AISLADA DE MÓDULOS
# ------------------------------------------------------------------------------
def descubrir_proyectos():
    """Detecta subcarpetas que contienen al menos config.py y main.py."""
    proyectos = {}
    carpetas_ignorar = {'.git', '__pycache__', '.streamlit', 'Plantillas_pptx'}
    for nombre in sorted(os.listdir(RAIZ), key=str.lower):
        if nombre in carpetas_ignorar:
            continue
        ruta = os.path.join(RAIZ, nombre)
        if (os.path.isdir(ruta)
                and os.path.exists(os.path.join(ruta, 'config.py'))
                and os.path.exists(os.path.join(ruta, 'main.py'))):
            proyectos[nombre] = ruta
    return proyectos


def importar_modulo_aislado(nombre_modulo, ruta_archivo):
    """Carga un script Python directamente por ruta física sin colisiones."""
    spec = importlib.util.spec_from_file_location(nombre_modulo, ruta_archivo)
    if spec is None or spec.loader is None:
        raise ImportError(f"No se pudo cargar {nombre_modulo} desde {ruta_archivo}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[nombre_modulo] = mod
    spec.loader.exec_module(mod)
    return mod


def cargar_proyecto(ruta_proyecto, nombre_proyecto):
    """Limpia la memoria e importa los archivos del proyecto seleccionado."""
    # Desregistrar versiones previas de memoria
    for mod in ['config', 'calculos', 'inyectores', 'main']:
        sys.modules.pop(mod, None)

    if ruta_proyecto not in sys.path:
        sys.path.insert(0, ruta_proyecto)

    try:
        with cambiar_directorio(ruta_proyecto):
            config = importar_modulo_aislado("config", os.path.join(ruta_proyecto, "config.py"))
            # Pre-cargar dependencias locales si existen
            for dep in ["calculos", "inyectores"]:
                p_dep = os.path.join(ruta_proyecto, f"{dep}.py")
                if os.path.exists(p_dep):
                    importar_modulo_aislado(dep, p_dep)
            pipeline = importar_modulo_aislado("main", os.path.join(ruta_proyecto, "main.py"))
    finally:
        if ruta_proyecto in sys.path:
            sys.path.remove(ruta_proyecto)

    return config, pipeline


def _guardar_upload(uploaded_file):
    sufijo = os.path.splitext(uploaded_file.name)[1] or '.xlsx'
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=sufijo)
    tmp.write(uploaded_file.getbuffer())
    tmp.close()
    return tmp.name


# ------------------------------------------------------------------------------
# INTERFAZ DE USUARIO (STREAMLIT)
# ------------------------------------------------------------------------------
st.set_page_config(page_title="Central IPN Colaboradores", page_icon="📊", layout="wide")
st.title("📊 Central de Reportes IPN Colaboradores")
st.caption("Consolidación en una sola liga para todas las unidades de negocio.")

proyectos = descubrir_proyectos()
if not proyectos:
    st.error("No se encontraron carpetas de proyecto válidas (con config.py y main.py).")
    st.stop()

# --- 1. SELECTOR PRINCIPAL ---
st.sidebar.header("Configuración de Unidad")
nombre_proyecto = st.sidebar.selectbox(
    "Selecciona la presentación a generar:",
    list(proyectos.keys())
)
ruta_proyecto = proyectos[nombre_proyecto]

try:
    config, pipeline = cargar_proyecto(ruta_proyecto, nombre_proyecto)
except Exception as e:
    st.error(f"Error cargando los scripts de '{nombre_proyecto}': {e}")
    st.stop()

# Información rápida en barra lateral
st.sidebar.markdown(f"**Directorio:** `{nombre_proyecto}`")
historicos = [getattr(config, attr) for attr in
              ('FILE_HISTORICO', 'FILE_HISTORICO_AGRUPADORES', 'FILE_HISTORICO_UPAX', 'FILE_HISTORICO_RL')
              if getattr(config, attr, None)]
if historicos:
    st.sidebar.warning(f"⚠ Modifica {len(historicos)} archivo(s) histórico(s).")

# Resumen de métricas detectadas
resumen = []
if getattr(config, 'BLOQUES_PARTICIPACION', None):
    resumen.append(f"{len(config.BLOQUES_PARTICIPACION)} bloques de participación")
if getattr(config, 'BLOQUES_NPS', None):
    resumen.append(f"{len(config.BLOQUES_NPS)} bloques NPS")
if getattr(config, 'BLOQUES_MENCIONES', None):
    resumen.append(f"{len(config.BLOQUES_MENCIONES)} bloques de menciones")
if getattr(config, 'HISTORICO', None):
    resumen.append("seguimiento histórico")

st.info(
    f"**Unidad seleccionada:** {nombre_proyecto} | **Estructura:** {', '.join(resumen) if resumen else 'Ver config.py'}")

col_izq, col_der = st.columns(2)

# --- 2. BASES DE ENTRADA (CON PERSISTENCIA) ---
with col_izq:
    st.subheader("1. Insumos (Bases de datos)")
    modo = st.radio("Método de entrada:", ["Escribir ruta local", "Subir archivos Excel"], horizontal=True)

    key_act = f"path_act_{nombre_proyecto}"
    key_ant = f"path_ant_{nombre_proyecto}"

    if key_act not in st.session_state:
        st.session_state[key_act] = None
    if key_ant not in st.session_state:
        st.session_state[key_ant] = None

    if modo == "Subir archivos Excel":
        up_actual = st.file_uploader("Base trimestre ACTUAL", type=['xlsx', 'xlsm'], key=f'act_{nombre_proyecto}')
        up_anterior = st.file_uploader("Base trimestre ANTERIOR", type=['xlsx', 'xlsm'], key=f'ant_{nombre_proyecto}')

        if up_actual is not None:
            st.session_state[key_act] = _guardar_upload(up_actual)
        if up_anterior is not None:
            st.session_state[key_ant] = _guardar_upload(up_anterior)

        ruta_actual = st.session_state[key_act]
        ruta_anterior = st.session_state[key_ant]
    else:
        def_actual = getattr(config, 'FILE_ACTUAL', '')
        def_anterior = getattr(config, 'FILE_ANTERIOR', '')
        ruta_actual = st.text_input("Ruta base ACTUAL:", value=def_actual)
        ruta_anterior = st.text_input("Ruta base ANTERIOR:", value=def_anterior)

        for tag, path in [("actual", ruta_actual), ("anterior", ruta_anterior)]:
            if path and not os.path.exists(path):
                st.caption(f":orange[Nota: La ruta {tag} no existe en este disco actualmente.]")

# --- 3. CONFIGURACIÓN DEL PERIODO Y PLANTILLA PPTX ---
with col_der:
    st.subheader("2. Parámetros del Reporte")
    c_periodo1, c_periodo2 = st.columns(2)
    with c_periodo1:
        trimestre = st.selectbox("Trimestre", ["1Q", "2Q", "3Q", "4Q"], index=1)
    with c_periodo2:
        anio = st.number_input("Año", min_value=2024, max_value=2030, value=2026, step=1)

    nuevo_periodo = f"{trimestre}´{str(anio)[-2:]}"
    st.caption(f"Etiqueta generada: **{nuevo_periodo}**")

    output_pptx = st.text_input(
        "Nombre de salida (.pptx):",
        value=f"{nombre_proyecto}_{trimestre}_{anio}.pptx",
        key=f"salida_{nombre_proyecto}"
    )
    if output_pptx and not output_pptx.lower().endswith('.pptx'):
        output_pptx += '.pptx'

    # Detección inteligente de la plantilla (Plantillas_pptx/ -> proyecto/ -> config)
    tpl_raw = getattr(config, 'PLANTILLA_PPTX', getattr(config, 'TEMPLATE_PPTX', 'plantilla.pptx'))
    nombre_archivo_ppt = os.path.basename(tpl_raw)

    ruta_en_plantillas = os.path.join(CARPETA_PLANTILLAS, nombre_archivo_ppt)
    ruta_en_proyecto = os.path.join(ruta_proyecto, nombre_archivo_ppt)

    if os.path.exists(ruta_en_plantillas):
        plantilla_default = ruta_en_plantillas
    elif os.path.exists(ruta_en_proyecto):
        plantilla_default = ruta_en_proyecto
    else:
        plantilla_default = tpl_raw

    with st.expander("Plantilla PowerPoint"):
        template_pptx = st.text_input("Ruta/Nombre plantilla:", value=plantilla_default)
        if not os.path.exists(template_pptx):
            st.warning(
                f"⚠ No se localizó el archivo: `{template_pptx}`. Verifica que exista en la carpeta `Plantillas_pptx/`.")

# --- 4. EJECUCIÓN ---
st.write("---")
listo = bool(ruta_actual) and bool(ruta_anterior) and bool(output_pptx) and os.path.exists(template_pptx)

if st.button(f"🚀 Ejecutar automatización para {nombre_proyecto}", type="primary", disabled=not listo):
    ruta_salida_destino = os.path.join(tempfile.gettempdir(), output_pptx)
    log_buffer = io.StringIO()

    try:
        with st.spinner(f"Procesando métricas e inyectando gráficos en {nombre_proyecto}..."):
            with redirect_stdout(log_buffer):
                with cambiar_directorio(ruta_proyecto):
                    if hasattr(pipeline, 'ejecutar_reporte'):
                        resultado = pipeline.ejecutar_reporte(
                            file_actual=ruta_actual,
                            file_anterior=ruta_anterior,
                            output_pptx=ruta_salida_destino,
                            nuevo_periodo=nuevo_periodo,
                            template_pptx=template_pptx,
                        )
                        ruta_salida = resultado[0] if isinstance(resultado, tuple) else ruta_salida_destino
                        errores = resultado[1] if isinstance(resultado, tuple) and len(resultado) > 1 else []
                    elif hasattr(pipeline, 'ejecutar_automatizacion'):
                        config.FILE_ACTUAL = ruta_actual
                        config.FILE_ANTERIOR = ruta_anterior
                        config.ETIQUETA_TRIMESTRE = nuevo_periodo
                        config.PLANTILLA_PPTX = template_pptx
                        config.OUTPUT_PPTX = ruta_salida_destino

                        pipeline.ejecutar_automatizacion(
                            ruta_actual=ruta_actual,
                            ruta_anterior=ruta_anterior,
                            etiqueta=nuevo_periodo,
                            output_pptx=ruta_salida_destino
                        )
                        ruta_salida = ruta_salida_destino
                        errores = []
                    else:
                        raise AttributeError(
                            "No se encontró 'ejecutar_reporte' ni 'ejecutar_automatizacion' en main.py.")

        if errores:
            st.warning(f"Completado con advertencias en los bloques: {', '.join(map(str, errores))}")
        else:
            st.success("✔ Presentación generada exitosamente.")

        st.caption(f"Ubicación: `{ruta_salida}`")

        if os.path.exists(ruta_salida):
            with open(ruta_salida, 'rb') as f:
                st.download_button(
                    label="⬇ Descargar PowerPoint (.pptx)",
                    data=f.read(),
                    file_name=os.path.basename(ruta_salida),
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

    except Exception as e:
        st.error(f"❌ Error durante el proceso: {e}")

    with st.expander("Consola de salida (Log de ejecución)", expanded=False):
        st.code(log_buffer.getvalue() or "Sin mensajes de consola emitidos.", language=None)