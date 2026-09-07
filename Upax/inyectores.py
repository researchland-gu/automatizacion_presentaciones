# ==============================================================================
# INYECTORES: todo lo que escribe en shapes de PowerPoint. No calcula nada.
# ==============================================================================
import re
import copy
import difflib

import pandas as pd

from pptx.chart.data import CategoryChartData, ChartData
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE
from pptx.oxml.ns import qn

from config import (FUENTE, MOTIVOS_POSITIVOS, MOTIVOS_NEGATIVOS,
                    GRUPOS_ORDENADOS, UMBRAL_CASOS_MINIMOS)


# ------------------------------------------------------------------------------
# LIMPIEZA DE CEROS (presentación limpia)
# Cualquier valor calculado que sea 0 se muestra vacío en tablas y se omite
# en las etiquetas de las gráficas, para que el reporte final no se llene
# de "0", "0.0%" o "(0)".
# ------------------------------------------------------------------------------
def _es_cero(valor):
    """True si el valor es None, NaN o numéricamente cero."""
    if valor is None:
        return True
    try:
        if pd.isna(valor):
            return True
    except (TypeError, ValueError):
        pass
    try:
        return float(valor) == 0
    except (TypeError, ValueError):
        return False


def fmt_miles(valor):
    """'12,345' — o cadena vacía si el valor es 0/None/NaN."""
    return '' if _es_cero(valor) else f"{int(valor):,}"


def fmt_pct(valor, decimales=1):
    """'12.3%' — o cadena vacía si el valor es 0/None/NaN."""
    return '' if _es_cero(valor) else f"{float(valor):.{decimales}f}%"


def limpiar_ceros_serie(valores):
    """Convierte los ceros/NaN de una serie de gráfica en None, para que
    python-pptx no dibuje la etiqueta '0.0%' en ese punto."""
    return [None if _es_cero(v) else v for v in valores]


def _formatear_parrafo(p, alignment=None, size=None, bold=None, color=None):
    """Aplica alineación y formato al primer run del párrafo, tolerando
    celdas vacías (sin runs) sin lanzar IndexError."""
    if alignment is not None:
        p.alignment = alignment
    if not p.runs:
        return
    run = p.runs[0]
    run.font.name = FUENTE
    if size is not None:
        run.font.size = size
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def buscar_shape(slide, nombre, requiere_tabla=False, requiere_chart=False):
    """Busca una shape por nombre (ignorando espacios en extremos)."""
    objetivo = nombre.strip()
    for shape in slide.shapes:
        if shape.name.strip() == objetivo:
            if requiere_tabla and not shape.has_table:
                continue
            if requiere_chart and not shape.has_chart:
                continue
            return shape
    return None


# ------------------------------------------------------------------------------
# TIPO 1: Participación
# ------------------------------------------------------------------------------
def inyectar_tabla_pct(shape, validos_global, pct_global):
    """Tabla de participación global (número de válidos + porcentaje).
    Los valores en cero se muestran vacíos."""
    tabla = shape.table
    num_cols = len(tabla.columns)

    txt_validos = fmt_miles(validos_global)
    txt_pct = fmt_pct(pct_global)

    if num_cols >= 2:
        tabla.cell(0, 0).text = txt_validos
        _formatear_parrafo(tabla.cell(0, 0).text_frame.paragraphs[0],
                           alignment=PP_ALIGN.RIGHT, size=Pt(8), bold=False)

        tabla.cell(0, 1).text = txt_pct
        _formatear_parrafo(tabla.cell(0, 1).text_frame.paragraphs[0],
                           alignment=PP_ALIGN.LEFT, size=Pt(8), bold=True)
    else:
        cell = tabla.cell(0, 0)
        cell.text_frame.clear()
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER

        if txt_validos:
            r1 = p.add_run()
            r1.text = f"{txt_validos}   "
            r1.font.name, r1.font.size, r1.font.bold = FUENTE, Pt(8), False

        if txt_pct:
            r2 = p.add_run()
            r2.text = txt_pct
            r2.font.name, r2.font.size, r2.font.bold = FUENTE, Pt(8), True

    print(f"Éxito: '{shape.name}' actualizada.")


def inyectar_tabla_pct_compuesta(shape, validos_global, pct_global):
    """Celda única en el 2do renglón con la cadena 'IPN – N – %'."""
    tabla = shape.table

    if len(tabla.rows) >= 2:
        cell = tabla.cell(1, 0)
        cell.text_frame.clear()

        # Si no hay datos, la celda queda limpia en lugar de 'IPN – 0 – 0.0%'
        if not _es_cero(validos_global):
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER

            run = p.add_run()
            run.text = f"IPN – {int(validos_global):,} – {pct_global:.1f}%"
            run.font.name = FUENTE
            run.font.size = Pt(14)
            run.font.bold = False

        print(f"Éxito: '{shape.name}' actualizada en el segundo renglón.")
    else:
        print(f"⚠ Error: La tabla '{shape.name}' necesita al menos 2 renglones.")


def inyectar_tabla_desglose(shape, df_resumen):
    """Tabla de desglose de negocios (nombre + válidos + porcentaje por fila).
    Las filas cuyo segmento tiene 0 válidos quedan completamente en blanco."""
    tabla = shape.table
    num_cols = len(tabla.columns)

    for i, row in df_resumen.iterrows():
        f_idx = i + 1  # Fila 0 es el encabezado
        if f_idx >= len(tabla.rows):
            break

        nombre_negocio = row[df_resumen.columns[0]]
        validos = row['Validos']
        pct = row['Pct']

        # Fila sin datos: se limpia por completo (nombre incluido)
        if _es_cero(validos):
            for c in range(num_cols):
                tabla.cell(f_idx, c).text = ''
            continue

        tabla.cell(f_idx, 0).text = str(nombre_negocio)
        _formatear_parrafo(tabla.cell(f_idx, 0).text_frame.paragraphs[0],
                           alignment=PP_ALIGN.LEFT, size=Pt(8))

        if num_cols == 3:
            tabla.cell(f_idx, 1).text = fmt_miles(validos)
            _formatear_parrafo(tabla.cell(f_idx, 1).text_frame.paragraphs[0],
                               alignment=PP_ALIGN.RIGHT, size=Pt(8))

            tabla.cell(f_idx, 2).text = fmt_pct(pct)
            _formatear_parrafo(tabla.cell(f_idx, 2).text_frame.paragraphs[0],
                               alignment=PP_ALIGN.RIGHT, size=Pt(8), bold=True)

        elif num_cols == 2:
            cell = tabla.cell(f_idx, 1)
            cell.text_frame.clear()
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT

            txt_val, txt_pct = fmt_miles(validos), fmt_pct(pct)
            if txt_val:
                r1 = p.add_run()
                r1.text = f"{txt_val}   "
                r1.font.name, r1.font.size = FUENTE, Pt(8)
            if txt_pct:
                r2 = p.add_run()
                r2.text = txt_pct
                r2.font.name, r2.font.size, r2.font.bold = FUENTE, Pt(8), True

    print(f"Éxito: '{shape.name}' actualizada.")


# ------------------------------------------------------------------------------
# TIPO 2: Bloques NPS (barras + diferencia + dona)
# ------------------------------------------------------------------------------
def listar_shapes(slide):
    """Nombres de todas las shapes de una diapositiva (para diagnóstico)."""
    return sorted(shape.name for shape in slide.shapes)


def inyectar_grafica_barras(slide, nombre_grafico, categorias, lista_det, lista_pas, lista_prom):
    shape = buscar_shape(slide, nombre_grafico, requiere_chart=True)
    if shape is None:
        print(f"⚠ No se encontró la gráfica '{nombre_grafico}'.")
        print(f"  Shapes disponibles en esta diapositiva: {listar_shapes(slide)}")
        return

    chart_data = CategoryChartData()
    chart_data.categories = categorias
    # Los ceros van como None para no dibujar etiquetas '0.0%'
    chart_data.add_series('Detractores', limpiar_ceros_serie(lista_det))
    chart_data.add_series('Pasivos', limpiar_ceros_serie(lista_pas))
    chart_data.add_series('Promotores', limpiar_ceros_serie(lista_prom))
    shape.chart.replace_data(chart_data)

    for series in shape.chart.series:
        series.has_data_labels = True
        series.data_labels.number_format = '0.0"%"'
        font = series.data_labels.font
        font.name = FUENTE
        font.size = Pt(9)
        font.bold = True

    print(f"Gráfica de barras '{nombre_grafico}' actualizada con éxito.")


def inyectar_tablas_diferencia(slide, bloque, df_final):
    etiqueta_total = bloque['etiqueta_total']
    try:
        total_hist = df_final[df_final['tipo_geografia_general'] == etiqueta_total].iloc[0]
        ipn_anterior_total = total_hist['ipn_3q']
        casos_anterior_total = total_hist['n_3q']
    except IndexError:
        ipn_anterior_total = 0.0
        casos_anterior_total = 0

    # A. Tabla principal de diferencias
    shape_dif = buscar_shape(slide, bloque['tabla_dif'], requiere_tabla=True)
    if shape_dif is not None:
        tabla = shape_dif.table
        num_cols = len(tabla.columns)

        for i, row_data in df_final.iterrows():
            f_idx = i + 1
            if f_idx >= len(tabla.rows):
                break

            # Fila sin casos en el periodo actual: se limpia por completo
            if _es_cero(row_data['n_4q']):
                for c in range(num_cols):
                    tabla.cell(f_idx, c).text = ''
                continue

            # Confidencialidad: con pocos casos se muestra el conteo pero se
            # enmascara el IPN actual con '*' y el anterior/diferencia con '–'
            if int(row_data['n_4q']) <= UMBRAL_CASOS_MINIMOS:
                if num_cols > 0:
                    tabla.cell(f_idx, 0).text = fmt_miles(row_data['n_4q'])
                if num_cols > 1:
                    tabla.cell(f_idx, 1).text = '*'
                if num_cols > 2:
                    tabla.cell(f_idx, 2).text = '–'
                if num_cols > 3:
                    tabla.cell(f_idx, 3).text = '–'

                for col_idx in range(min(num_cols, 4)):
                    _formatear_parrafo(tabla.cell(f_idx, col_idx).text_frame.paragraphs[0],
                                       alignment=PP_ALIGN.CENTER, size=Pt(10), bold=True)
                continue

            if num_cols > 0:
                tabla.cell(f_idx, 0).text = fmt_miles(row_data['n_4q'])
            if num_cols > 1:
                tabla.cell(f_idx, 1).text = fmt_pct(row_data['ipn_4q'])
            if num_cols > 2:
                tabla.cell(f_idx, 2).text = fmt_pct(row_data['ipn_3q'])
            if num_cols > 3:
                d_val = row_data['dif']
                if _es_cero(d_val):
                    tabla.cell(f_idx, 3).text = ''
                else:
                    signo = "+" if d_val > 0 else ""
                    tabla.cell(f_idx, 3).text = f"{signo}{d_val:.1f} pp"

            for col_idx in range(min(num_cols, 4)):
                _formatear_parrafo(tabla.cell(f_idx, col_idx).text_frame.paragraphs[0],
                                   alignment=PP_ALIGN.CENTER, size=Pt(10), bold=True)

        print(f"Éxito: '{bloque['tabla_dif']}' actualizada.")
    else:
        print(f"⚠ No se encontró la tabla '{bloque['tabla_dif']}'.")

    # B. Tabla resumen del periodo anterior
    shape_ant = buscar_shape(slide, bloque['tabla_anterior'], requiere_tabla=True)
    if shape_ant is not None:
        tabla_ant = shape_ant.table
        if len(tabla_ant.rows) >= 2 and len(tabla_ant.columns) >= 2:
            tabla_ant.cell(0, 1).text = fmt_pct(ipn_anterior_total) if not _es_cero(casos_anterior_total) else ''
            tabla_ant.cell(1, 1).text = fmt_miles(casos_anterior_total)

            for f_idx in [0, 1]:
                _formatear_parrafo(tabla_ant.cell(f_idx, 1).text_frame.paragraphs[0],
                                   alignment=PP_ALIGN.CENTER, size=Pt(10), bold=True)

            print(f"Éxito: '{bloque['tabla_anterior']}' actualizada.")
    else:
        print(f"⚠ No se encontró la tabla '{bloque['tabla_anterior']}'.")


def aplicar_formato_texto(text_frame, titulo, porcentaje, conteo):
    """Formato de 3 líneas usado en las tablas resumen de la dona.
    Las líneas con texto vacío (valores en cero) no agregan runs."""
    text_frame.clear()
    p1 = text_frame.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    if titulo:
        run1 = p1.add_run()
        run1.text = titulo
        run1.font.name, run1.font.size, run1.font.bold = 'Montserrat ExtraBold', Pt(14), True

    p2 = text_frame.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    if porcentaje:
        run2 = p2.add_run()
        run2.text = porcentaje
        run2.font.name, run2.font.size, run2.font.bold = 'Montserrat ExtraBold', Pt(14), True

    p3 = text_frame.add_paragraph()
    p3.alignment = PP_ALIGN.CENTER
    if conteo:
        run3 = p3.add_run()
        run3.text = conteo
        run3.font.name, run3.font.size, run3.font.bold = 'Montserrat Light', Pt(10.5), False


def inyectar_dona(slide, bloque, datos_dona):
    total_n, counts, p_prom, p_pas, p_det, ipn = datos_dona
    suf = bloque['sufijo_dona']

    def _conteo(n):
        txt = fmt_miles(n)
        return f"({txt})" if txt else ''

    config_tablas = {
        f'Tabla_PRO_{suf}': ["Promotores", fmt_pct(p_prom), _conteo(counts.get('Promotor', 0))],
        f'Tabla_PAS_{suf}': ["Pasivos", fmt_pct(p_pas), _conteo(counts.get('Pasivo', 0))],
        f'Tabla_DET_{suf}': ["Detractores", fmt_pct(p_det), _conteo(counts.get('Detractor', 0))],
        f'Tabla_IPN_{suf}': [fmt_pct(ipn), "IPN", _conteo(total_n)],
    }

    nombre_dona = bloque['grafico_dona'].strip()

    for shape in slide.shapes:
        s_name = shape.name.strip()

        if s_name == nombre_dona and shape.has_chart:
            chart_data = CategoryChartData()
            chart_data.categories = ['Promotores', 'Pasivos', 'Detractores']
            chart_data.add_series('calificacion_nps', (p_prom, p_pas, p_det))
            shape.chart.replace_data(chart_data)

        elif s_name in config_tablas:
            datos = config_tablas[s_name]
            text_frame = shape.table.cell(0, 0).text_frame if shape.has_table else shape.text_frame
            aplicar_formato_texto(text_frame, datos[0], datos[1], datos[2])

    print(f"Dona '{nombre_dona}' y sus tablas actualizadas.")


# ------------------------------------------------------------------------------
# TIPO 3: Menciones Top 5
# ------------------------------------------------------------------------------
def _limpiar_para_match(texto):
    t = str(texto).lower().strip()
    reemplazos = {'á': 'a', 'é': 'e', 'í': 'i', 'ó': 'o', 'ú': 'u', 'ü': 'u', 'ñ': 'n'}
    for orig, nuevo in reemplazos.items():
        t = t.replace(orig, nuevo)
    return re.sub(r'[^a-z0-9]', '', t)


_POSITIVAS_NORM = [_limpiar_para_match(m) for m in MOTIVOS_POSITIVOS]
_NEGATIVAS_NORM = [_limpiar_para_match(m) for m in MOTIVOS_NEGATIVOS]


def obtener_color_mencion(motivo):
    m = _limpiar_para_match(motivo)
    if m in _POSITIVAS_NORM:
        return RGBColor(0, 176, 80)
    if m in _NEGATIVAS_NORM:
        return RGBColor(255, 0, 0)
    if difflib.get_close_matches(m, _POSITIVAS_NORM, n=1, cutoff=0.7):
        return RGBColor(0, 176, 80)
    if difflib.get_close_matches(m, _NEGATIVAS_NORM, n=1, cutoff=0.7):
        return RGBColor(255, 0, 0)
    return RGBColor(0, 0, 0)


def _inyectar_celda_compuesta(cell, motivo, menciones, color_rgb):
    cell.text_frame.clear()
    cell.text_frame.word_wrap = True
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    p = cell.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER

    run1 = p.add_run()
    run1.text = f"{motivo}\n"
    run1.font.name = FUENTE
    run1.font.size = Pt(9.5)
    run1.font.bold = False
    run1.font.color.rgb = color_rgb

    run2 = p.add_run()
    run2.text = f"{int(menciones):,}"
    run2.font.name = FUENTE
    run2.font.size = Pt(10)
    run2.font.bold = True
    run2.font.color.rgb = color_rgb


def inyectar_menciones(slide, nombre_tabla, resultados_por_agrupador):
    shape = buscar_shape(slide, nombre_tabla, requiere_tabla=True)
    if shape is None:
        print(f"⚠ No se encontró la tabla '{nombre_tabla}'.")
        return

    tabla = shape.table
    print(f"Inyectando motivos y menciones en '{nombre_tabla}'...")

    for c_idx, grupo in enumerate(GRUPOS_ORDENADOS):
        df_top5 = resultados_por_agrupador[grupo]

        for r_idx in range(5):
            f_idx = r_idx + 1  # +1 para saltar encabezados
            if f_idx >= len(tabla.rows) or c_idx >= len(tabla.columns):
                continue

            if r_idx < len(df_top5):
                motivo = str(df_top5.loc[r_idx, 'Motivo'])
                menciones = df_top5.loc[r_idx, 'Menciones']
                color = obtener_color_mencion(motivo)
                _inyectar_celda_compuesta(tabla.cell(f_idx, c_idx), motivo, menciones, color)
            else:
                tabla.cell(f_idx, c_idx).text_frame.clear()

    print(f"Éxito: '{nombre_tabla}' actualizada.")


# ------------------------------------------------------------------------------
# TIPO 4: Histórico
# ------------------------------------------------------------------------------
def inyectar_historico(slide, nombre_grafico, df_hist):
    shape = buscar_shape(slide, nombre_grafico, requiere_chart=True)
    if shape is None:
        print(f"⚠ No se encontró la gráfica '{nombre_grafico}'.")
        return

    chart_data = ChartData()
    chart_data.categories = df_hist['Periodo'].tolist()
    chart_data.add_series('Promotores', df_hist['Promotores'].tolist())
    chart_data.add_series('IPN', df_hist['IPN'].tolist())
    chart_data.add_series('Pasivos', df_hist['Pasivos'].tolist())
    chart_data.add_series('Detractores', df_hist['Detractores'].tolist())
    shape.chart.replace_data(chart_data)

    estilos = {
        'Promotores':  (RGBColor(0, 176, 80),   MSO_LINE.DASH),
        'IPN':         (RGBColor(112, 48, 160), MSO_LINE.SOLID),
        'Pasivos':     (RGBColor(255, 192, 0),  MSO_LINE.DASH),
        'Detractores': (RGBColor(255, 0, 0),    MSO_LINE.DASH),
    }

    for series in shape.chart.series:
        series.has_data_labels = True
        series.data_labels.number_format = '0.0%'
        font = series.data_labels.font
        font.name = FUENTE
        font.size = Pt(9)
        font.bold = False

        if series.name in estilos:
            color, dash = estilos[series.name]
            series.format.line.color.rgb = color
            series.format.line.dash_style = dash

    print(f"Gráfica histórica '{nombre_grafico}' actualizada.")


def agregar_columna_tabla(shape):
    """Manipula el XML de la tabla para clonar la última columna y agregar
    una nueva al final de forma segura (celdas clonadas quedan vacías)."""
    tbl = shape._element.graphic.graphicData.tbl
    tblGrid = tbl.tblGrid

    new_gridCol = copy.deepcopy(tblGrid.gridCol_lst[-1])
    tblGrid.append(new_gridCol)

    for tr in tbl.tr_lst:
        tcs = tr.findall(qn('a:tc'))
        if not tcs:
            continue

        last_tc = tcs[-1]
        new_tc = copy.deepcopy(last_tc)

        # Vaciamos el texto previo del clon
        if new_tc.txBody is not None:
            for p in new_tc.txBody.p_lst:
                for r in p.r_lst:
                    p.remove(r)

        # Inserción segura: justo después de la última celda
        last_tc.addnext(new_tc)


def inyectar_tabla_pct_historicos(slide, nombre_tabla, metricas, col_objetivo):
    """Inyecta % de participación y número de participantes del periodo en la
    columna objetivo, creando columnas nuevas si el periodo no cabe."""
    shape = buscar_shape(slide, nombre_tabla, requiere_tabla=True)
    if shape is None:
        print(f"⚠ No se encontró la tabla '{nombre_tabla}'.")
        return

    tabla = shape.table
    while col_objetivo >= len(tabla.columns):
        agregar_columna_tabla(shape)
        print(f"  -> Se creó una nueva columna en {nombre_tabla}")

    pct_val = metricas['Pct_Participacion']
    texto_pct = '' if _es_cero(pct_val) else f"{pct_val * 100:.1f}%".replace('.0%', '%')
    part_val = metricas['Participantes']
    texto_part = fmt_miles(part_val)

    if len(tabla.rows) >= 2:
        # Dos filas separadas: % arriba, participantes abajo
        tabla.cell(0, col_objetivo).text = texto_pct
        _formatear_parrafo(tabla.cell(0, col_objetivo).text_frame.paragraphs[0],
                           alignment=PP_ALIGN.CENTER, size=Pt(9.5), bold=True)

        tabla.cell(1, col_objetivo).text = texto_part
        _formatear_parrafo(tabla.cell(1, col_objetivo).text_frame.paragraphs[0],
                           alignment=PP_ALIGN.CENTER, size=Pt(9), bold=True,
                           color=RGBColor(124, 177, 201))
    else:
        # Una sola fila: ambos textos apilados en la misma celda
        cell = tabla.cell(0, col_objetivo)
        cell.text_frame.clear()

        p0 = cell.text_frame.paragraphs[0]
        p0.alignment = PP_ALIGN.CENTER
        if texto_pct:
            r0 = p0.add_run()
            r0.text = texto_pct
            r0.font.name, r0.font.size, r0.font.bold = FUENTE, Pt(9.5), True

        p1 = cell.text_frame.add_paragraph()
        p1.alignment = PP_ALIGN.CENTER
        if texto_part:
            r1 = p1.add_run()
            r1.text = texto_part
            r1.font.name, r1.font.size, r1.font.bold = FUENTE, Pt(9), True
            r1.font.color.rgb = RGBColor(124, 177, 201)

    print(f"Éxito: Columna {col_objetivo} de '{nombre_tabla}' actualizada.")


def inyectar_tabla_casos_historicos(slide, nombre_tabla, metricas, col_objetivo):
    """Inyecta los conteos absolutos (promotores, pasivos, detractores) del
    periodo en la columna objetivo, creando columnas nuevas si no cabe."""
    shape = buscar_shape(slide, nombre_tabla, requiere_tabla=True)
    if shape is None:
        print(f"⚠ No se encontró la tabla '{nombre_tabla}'.")
        return

    tabla = shape.table
    while col_objetivo >= len(tabla.columns):
        agregar_columna_tabla(shape)
        print(f"  -> Se creó una nueva columna en {nombre_tabla}")

    tabla.cell(0, col_objetivo).text = fmt_miles(metricas['Num_Promotores'])
    tabla.cell(1, col_objetivo).text = fmt_miles(metricas['Num_Pasivos'])
    tabla.cell(2, col_objetivo).text = fmt_miles(metricas['Num_Detractores'])

    for r_idx in range(3):
        _formatear_parrafo(tabla.cell(r_idx, col_objetivo).text_frame.paragraphs[0],
                           alignment=PP_ALIGN.CENTER, size=Pt(9), bold=False)

    print(f"Éxito: Columna {col_objetivo} de '{nombre_tabla}' actualizada.")


# ------------------------------------------------------------------------------
# TIPO 5: Histórico por agrupador (8 gráficas de barras, slide 12)
# ------------------------------------------------------------------------------
def inyectar_historicos_agrupadores(slide, mapa_graficas, df_hist):
    """Inyecta el histórico de IPN de cada agrupador en su gráfica de barras.
    La barra del periodo más reciente va en gris oscuro; el resto en gris claro."""
    nombres_a_agrupador = {v: k for k, v in mapa_graficas.items()}
    graficas_actualizadas = 0

    for shape in slide.shapes:
        if shape.has_chart and shape.name.strip() in nombres_a_agrupador:
            agrupador = nombres_a_agrupador[shape.name.strip()]

            chart_data = ChartData()
            chart_data.categories = df_hist['Periodo'].tolist()

            # Vacíos (NaN) y ceros no se grafican ni muestran etiqueta
            lista_ipn = limpiar_ceros_serie(df_hist[agrupador].tolist())
            chart_data.add_series('IPN', lista_ipn)

            shape.chart.replace_data(chart_data)

            for series in shape.chart.series:
                series.has_data_labels = True
                series.data_labels.number_format = '0.0%'

                font = series.data_labels.font
                font.name = FUENTE
                font.size = Pt(8)
                font.bold = False

                for i, point in enumerate(series.points):
                    fill = point.format.fill
                    fill.solid()
                    if i == len(series.points) - 1:
                        fill.fore_color.rgb = RGBColor(105, 105, 105)  # Más reciente
                    else:
                        fill.fore_color.rgb = RGBColor(192, 192, 192)  # Históricas

            graficas_actualizadas += 1
            print(f"Éxito: Gráfica '{shape.name}' ({agrupador}) actualizada.")

    if graficas_actualizadas == 0:
        print("⚠ Error: No se encontró ninguna gráfica de agrupadores en la diapositiva.")

    return graficas_actualizadas


# ------------------------------------------------------------------------------
# TIPO 6: Formatos de tarjetas de dona (paso final)
# ------------------------------------------------------------------------------
def _aplicar_estilo_parrafo(p, regla):
    """Aplica fuente/tamaño/negritas/color a todos los runs de un párrafo.
    Si el párrafo tiene texto pero no runs, lo reconstruye primero."""
    if not p.runs and p.text:
        txt = p.text
        p.text = ''
        run = p.add_run()
        run.text = txt
    for run in p.runs:
        run.font.name = FUENTE
        run.font.size = Pt(regla['size'])
        run.font.bold = regla['bold']
        run.font.color.rgb = RGBColor(*regla['color'])


def _aplicar_formato_tarjeta(shape, reglas):
    """Aplica las reglas por línea (0,1,2) a una tarjeta, sea tabla o cuadro
    de texto. En tablas de 3+ filas cada regla va a una fila; en tablas de una
    celda o cuadros de texto, cada regla va a un párrafo."""
    if shape.has_table:
        tabla = shape.table
        if len(tabla.rows) >= 3:
            for r_idx, regla in reglas.items():
                if r_idx < len(tabla.rows):
                    cell = tabla.cell(r_idx, 0)
                    for p in cell.text_frame.paragraphs:
                        _aplicar_estilo_parrafo(p, regla)
        else:
            cell = tabla.cell(0, 0)
            for p_idx, p in enumerate(cell.text_frame.paragraphs):
                if p_idx in reglas:
                    _aplicar_estilo_parrafo(p, reglas[p_idx])
    elif shape.has_text_frame:
        for p_idx, p in enumerate(shape.text_frame.paragraphs):
            if p_idx in reglas:
                _aplicar_estilo_parrafo(p, reglas[p_idx])


def aplicar_formatos_dona(slide, sufijo, estilos):
    """Formatea las 4 tarjetas de la dona de un bloque NPS:
    Tabla_DET/PAS/PRO_{suf} con la regla 'DET_PAS_PRO' y Tabla_IPN_{suf}
    con la regla 'IPN'. Devuelve cuántas tarjetas se formatearon."""
    objetivos = {
        f'Tabla_DET_{sufijo}': 'DET_PAS_PRO',
        f'Tabla_PAS_{sufijo}': 'DET_PAS_PRO',
        f'Tabla_PRO_{sufijo}': 'DET_PAS_PRO',
        f'Tabla_IPN_{sufijo}': 'IPN',
    }

    formateadas = 0
    for shape in slide.shapes:
        s_name = shape.name.strip()
        if s_name in objetivos:
            _aplicar_formato_tarjeta(shape, estilos[objetivos[s_name]])
            formateadas += 1

    return formateadas


# ------------------------------------------------------------------------------
# MENCIONES TOP 10 CON RUBRO (slides 7-9 del reporte Despachos)
# ------------------------------------------------------------------------------
def _inyectar_texto_celda(cell, texto, alineacion, color_rgb, font_size=14, is_bold=True):
    """Escribe una celda con word-wrap y centrado vertical activados."""
    cell.text_frame.clear()
    cell.text_frame.word_wrap = True
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE

    p = cell.text_frame.paragraphs[0]
    p.alignment = alineacion

    if str(texto):
        run = p.add_run()
        run.text = str(texto)
        run.font.name = FUENTE
        run.font.size = Pt(font_size)
        run.font.bold = is_bold
        run.font.color.rgb = color_rgb


def inyectar_menciones_rubro(slide, nombre_tabla, top_motivos):
    """Inyecta el Top N de motivos con rubro en una tabla de 3 columnas:
    motivo (izquierda, coloreado), menciones (centro, coloreado) y
    rubro (centro, gris). Las filas sobrantes de la tabla se limpian."""
    shape = buscar_shape(slide, nombre_tabla, requiere_tabla=True)
    if shape is None:
        print(f"⚠ No se encontró la tabla '{nombre_tabla}'.")
        return

    tabla = shape.table
    num_cols = len(tabla.columns)
    print(f"Inyectando Top {len(top_motivos)} de motivos en '{nombre_tabla}'...")

    for i, row_data in top_motivos.iterrows():
        f_idx = i + 1
        if f_idx >= len(tabla.rows):
            break

        motivo = str(row_data['Motivo'])
        color = obtener_color_mencion(motivo)

        if num_cols > 0:
            _inyectar_texto_celda(tabla.cell(f_idx, 0), motivo, PP_ALIGN.LEFT, color)
        if num_cols > 1:
            _inyectar_texto_celda(tabla.cell(f_idx, 1), fmt_miles(row_data['Menciones']),
                                  PP_ALIGN.CENTER, color)
        if num_cols > 2:
            _inyectar_texto_celda(tabla.cell(f_idx, 2), str(row_data['Rubro']),
                                  PP_ALIGN.CENTER, RGBColor(128, 128, 128))

    # Filas sobrantes (menos motivos que renglones): quedan limpias
    for f_idx in range(len(top_motivos) + 1, len(tabla.rows)):
        for c in range(num_cols):
            tabla.cell(f_idx, c).text_frame.clear()

    print(f"Éxito: '{nombre_tabla}' actualizada.")


# ------------------------------------------------------------------------------
# HISTÓRICO UPAX (slide 8): gráficas históricas, tarjetas y diferencias
# ------------------------------------------------------------------------------
def _formatear_celda_hist(cell, texto, font_name, font_size_pt, color_rgb,
                          bold=False, align=PP_ALIGN.CENTER):
    """Aplica texto y propiedades tipográficas a una celda específica."""
    tf = cell.text_frame
    tf.word_wrap = False
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = align
    if str(texto):
        run = p.add_run()
        run.text = str(texto)
        run.font.name = font_name
        run.font.size = Pt(font_size_pt)
        run.font.bold = bold
        run.font.color.rgb = color_rgb


def inyectar_graficas_historico_upax(slide, mapa_graficas, periodos, series_historicas,
                                     formatear_ipn):
    """Actualiza las 4 gráficas de barras históricas (DET/PAS/PRO/IPN);
    la barra del periodo más reciente va en gris oscuro. La gráfica de IPN
    usa etiquetas de periodo en dos renglones."""
    graficas = 0
    nombres = {v: k for k, v in mapa_graficas.items()}

    for shape in slide.shapes:
        if shape.has_chart and shape.name.strip() in nombres:
            metrica = nombres[shape.name.strip()]
            valores = series_historicas[metrica]

            chart_data = ChartData()
            if metrica == 'IPN':
                chart_data.categories = [formatear_ipn(p) for p in periodos]
            else:
                chart_data.categories = periodos
            chart_data.add_series(metrica, [v if v is not None else 0.0 for v in valores])
            shape.chart.replace_data(chart_data)

            for series in shape.chart.series:
                series.has_data_labels = True
                series.data_labels.number_format = '0.0%'
                font = series.data_labels.font
                font.name = FUENTE
                font.size = Pt(8)
                font.bold = False

                for i, point in enumerate(series.points):
                    fill = point.format.fill
                    fill.solid()
                    if i == len(series.points) - 1:
                        fill.fore_color.rgb = RGBColor(105, 105, 105)
                    else:
                        fill.fore_color.rgb = RGBColor(192, 192, 192)

            graficas += 1
            print(f"   Éxito: Gráfica '{shape.name}' ({metrica}) actualizada.")

    if graficas == 0:
        print("⚠ No se encontró ninguna gráfica del histórico UPAX en la diapositiva.")
        print(f"  Shapes disponibles: {listar_shapes(slide)}")
    return graficas


def inyectar_tablas_historico_upax(slide, mapa_hist, mapa_dif, datos_tarjetas, diferencias):
    """Inyecta las 4 tarjetas de métricas del periodo (porcentaje + conteo)
    y las tablas de diferencia vs levantamiento anterior (flecha Webdings
    verde/roja/gris + porcentaje con signo)."""
    hist_lookup = {k.upper(): v for k, v in mapa_hist.items()}
    dif_lookup = {k.upper(): v for k, v in mapa_dif.items()}
    n_hist, n_dif = 0, 0

    for shape in slide.shapes:
        if not shape.has_table:
            continue

        nombre = shape.name.strip().upper()
        tabla = shape.table

        if nombre in hist_lookup:
            metrica = hist_lookup[nombre]
            val_pct, val_casos = datos_tarjetas[metrica]

            if len(tabla.rows) >= 3:
                _formatear_celda_hist(tabla.cell(1, 0), val_pct, FUENTE, 16, RGBColor(0, 0, 0), bold=True)
                _formatear_celda_hist(tabla.cell(2, 0), val_casos, FUENTE, 11, RGBColor(127, 127, 127))
            elif len(tabla.rows) == 2:
                cell = tabla.cell(1, 0)
                tf = cell.text_frame
                tf.clear()

                p1 = tf.paragraphs[0]
                p1.alignment = PP_ALIGN.CENTER
                if val_pct:
                    r1 = p1.add_run()
                    r1.text = val_pct
                    r1.font.name, r1.font.size, r1.font.bold = FUENTE, Pt(16), True
                    r1.font.color.rgb = RGBColor(0, 0, 0)

                p2 = tf.add_paragraph()
                p2.alignment = PP_ALIGN.CENTER
                if val_casos:
                    r2 = p2.add_run()
                    r2.text = val_casos
                    r2.font.name, r2.font.size, r2.font.bold = FUENTE, Pt(11), False
                    r2.font.color.rgb = RGBColor(127, 127, 127)

            n_hist += 1
            print(f"   Éxito: Tabla métrica '{shape.name}' ({metrica}) actualizada.")

        elif nombre in dif_lookup:
            metrica = dif_lookup[nombre]
            dif_valor = diferencias[metrica]

            if dif_valor > 0:
                simbolo, color_simbolo = '5', RGBColor(0, 176, 80)   # Flecha arriba verde
            elif dif_valor < 0:
                simbolo, color_simbolo = '6', RGBColor(255, 0, 0)    # Flecha abajo roja
            else:
                simbolo, color_simbolo = '-', RGBColor(127, 127, 127)

            target_row = 1 if len(tabla.rows) > 1 else 0
            _formatear_celda_hist(tabla.cell(target_row, 0), simbolo, 'Webdings', 22,
                                  color_simbolo, bold=True)
            _formatear_celda_hist(tabla.cell(target_row, 1), f"{dif_valor * 100:+.1f}%",
                                  FUENTE, 12, RGBColor(0, 0, 0))

            n_dif += 1
            print(f"   Éxito: Tabla diferencia '{shape.name}' ({metrica}): {dif_valor * 100:+.1f}%")

    return n_hist, n_dif
