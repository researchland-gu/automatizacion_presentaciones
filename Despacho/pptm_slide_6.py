import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN

# ==============================================================================
# 0. CONFIGURACIÓN DE RUTAS
# ==============================================================================
FILE_1Q_2026 = r'E:\Users\954108\Desktop\codigos_py\Presentaciones\IPN_Colaboradores\Base_resultados\Base para dashboard 1Q 2026.xlsx'
FILE_4Q_2025 = r'E:\Users\954108\Desktop\codigos_py\Presentaciones\IPN_Colaboradores\Base_resultados\Base para dashboard 4Q 2025 liberada.xlsx'
TEMPLATE_PPTX = r'E:\Users\954108\Desktop\Satisfaccion  y Asuntos Publicos\Automatizaciones_PPT\IPN Colaboradores\IPN_Colaboradores_Despachos_automatizada.pptx'
OUTPUT_PPTX = 'Reporte_Automatizado.pptx'


# ==============================================================================
# 1. FUNCIONES COMPARTIDAS (mapeo de geografía, clasificación NPS, filtros)
# ==============================================================================
def construir_mapeo_geografia(df, col_n3='Nivel 3', col_resp='Resonsable_N3'):
    """Genera el diccionario dinámico {Nivel 3 -> 'Nivel 3 - Responsable'}."""
    df[col_n3] = df[col_n3].fillna(' ').astype(str)
    df[col_resp] = df[col_resp].fillna('Sin Responsable').astype(str)

    mapeo = {}
    for n_orig, resp_orig in df[[col_n3, col_resp]].drop_duplicates().values:
        n_limpio = n_orig.strip()
        resp_limpio = resp_orig.strip()

        if n_limpio == "":
            mapeo[n_orig] = 'Sin Etiqueta'
        else:
            mapeo[n_orig] = f"{n_limpio.title()}\n{resp_limpio.title()}"
    return mapeo


def clasificar_nps(calificacion):
    """Clasifica una calificación en Detractor / Pasivo / Promotor."""
    if pd.isna(calificacion):
        return None
    elif calificacion <= 6:
        return 'Detractor'
    elif calificacion <= 8:
        return 'Pasivo'
    else:
        return 'Promotor'


def cargar_y_filtrar_base(path, excluir_no_identificados=True):
    df = pd.read_excel(path)

    if excluir_no_identificados:
        df = df[df['Nivel 3'] != 'CASOS NO IDENTIFICADOS'].copy()

    mapeo_geografia = construir_mapeo_geografia(df)
    df['tipo_geografia_general'] = df['Nivel 3'].map(mapeo_geografia)
    df['calificacion_nps'] = df['Calificacion'].apply(clasificar_nps)

    cats_desglose = ['Detractor', 'Pasivo', 'Promotor']
    df_1 = df[df['calificacion_nps'].isin(cats_desglose)].copy()
    df_2 = df_1[df_1['Respondida'] == 'Si'].copy()
    df_final = df_2[df_2['Agrupador'] == 'DESPACHOS'].copy()
    return df_final


def aplicar_formato_texto(text_frame, titulo, porcentaje, conteo):
    """Formato de texto de 3 líneas usado en las tablas resumen de la dona."""
    text_frame.clear()
    p1 = text_frame.paragraphs[0]
    p1.alignment = PP_ALIGN.CENTER
    run1 = p1.add_run()
    run1.text = titulo
    run1.font.name, run1.font.size, run1.font.bold = 'Montserrat ExtraBold', Pt(14), True

    p2 = text_frame.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    run2 = p2.add_run()
    run2.text = porcentaje
    run2.font.name, run2.font.size, run2.font.bold = 'Montserrat ExtraBold', Pt(14), True

    p3 = text_frame.add_paragraph()
    p3.alignment = PP_ALIGN.CENTER
    run3 = p3.add_run()
    run3.text = conteo
    run3.font.name, run3.font.size, run3.font.bold = 'Montserrat Light', Pt(10.5), False


# ==============================================================================
# 2. SECCIÓN BARRAS — cálculo de distribución apilada (Slide 13)
# ==============================================================================
def calcular_datos_barras(df):
    """Devuelve (categorias_eje, lista_det, lista_pas, lista_prom) ordenados
    de mayor a menor número de casos, con 'Grupo Elektra' fijo al inicio."""
    dist_general = df['calificacion_nps'].value_counts(normalize=True) * 100
    prom_gen = round(dist_general.get('Promotor', 0), 1)
    pas_gen = round(dist_general.get('Pasivo', 0), 1)
    det_gen = round(dist_general.get('Detractor', 0), 1)

    dist_seg = pd.crosstab(df['tipo_geografia_general'], df['calificacion_nps'], normalize='index') * 100
    for col in ['Promotor', 'Pasivo', 'Detractor']:
        if col not in dist_seg.columns:
            dist_seg[col] = 0.0

    dist_seg['num_casos'] = dist_seg.index.map(df['tipo_geografia_general'].value_counts())
    segmentos_ordenados = dist_seg.sort_values(by='num_casos', ascending=False).index.tolist()
    categorias_eje = ['Despachos - RBS'] + segmentos_ordenados

    lista_det = [det_gen] + [round(dist_seg.loc[cat, 'Detractor'], 1) for cat in categorias_eje[1:]]
    lista_pas = [pas_gen] + [round(dist_seg.loc[cat, 'Pasivo'], 1) for cat in categorias_eje[1:]]
    lista_prom = [prom_gen] + [round(dist_seg.loc[cat, 'Promotor'], 1) for cat in categorias_eje[1:]]

    return categorias_eje, lista_det, lista_pas, lista_prom


def actualizar_grafica_barras(slide, df):
    categorias_eje, lista_det, lista_pas, lista_prom = calcular_datos_barras(df)

    for shape in slide.shapes:
        if shape.name == 'Grafico_bar_DP' and shape.has_chart:
            chart_data = CategoryChartData()
            chart_data.categories = categorias_eje
            chart_data.add_series('Detractores', lista_det)
            chart_data.add_series('Pasivos', lista_pas)
            chart_data.add_series('Promotores', lista_prom)
            shape.chart.replace_data(chart_data)

            for series in shape.chart.series:
                series.has_data_labels = True
                series.data_labels.number_format = '0.0"%"'
                font = series.data_labels.font
                font.name = 'Montserrat'
                font.size = Pt(9)
                font.bold = True

            print("Gráfica de barras apiladas 'Grafico_bar_DP' actualizada y ordenada con éxito.")


# ==============================================================================
# 3. SECCIÓN DIFERENCIA — comparación entre periodos (Slide 13)
# ==============================================================================
def calcular_resumen_nps(df):
    """Calcula n e IPN por segmento, más el total 'Grupo Elektra'."""
    resumen = df.groupby('tipo_geografia_general').apply(lambda x: pd.Series({
        'n': float(len(x)),
        'ipn': round((x['calificacion_nps'].value_counts(normalize=True).get('Promotor', 0) * 100), 5) -
               round((x['calificacion_nps'].value_counts(normalize=True).get('Detractor', 0) * 100), 5)
    }), include_groups=False).reset_index()

    t_n = float(len(df))
    t_ipn = round((df['calificacion_nps'].value_counts(normalize=True).get('Promotor', 0) * 100), 5) - \
            round((df['calificacion_nps'].value_counts(normalize=True).get('Detractor', 0) * 100), 5)

    df_total = pd.DataFrame([{'tipo_geografia_general': 'Despachos - RBS', 'n': t_n, 'ipn': t_ipn}])
    return pd.concat([df_total, resumen], ignore_index=True)


def calcular_diferencia_periodos(df_actual, df_anterior):
    """Cruza el periodo actual (1Q 2026) contra el anterior (4Q 2025) y
    devuelve el DataFrame final ordenado, listo para inyectar en la tabla."""
    data_actual = calcular_resumen_nps(df_actual)
    data_anterior = calcular_resumen_nps(df_anterior)

    df_final = data_actual.merge(data_anterior, on='tipo_geografia_general', suffixes=('_4q', '_3q'))
    df_final['dif'] = df_final['ipn_4q'] - df_final['ipn_3q']

    df_total_row = df_final[df_final['tipo_geografia_general'] == 'Despachos - RBS']
    df_segmentos = df_final[df_final['tipo_geografia_general'] != 'Despachos - RBS'].copy()
    df_segmentos = df_segmentos.sort_values(by='n_4q', ascending=False)

    df_final = pd.concat([df_total_row, df_segmentos], ignore_index=True)
    return df_final


def actualizar_tablas_diferencia(slide, df_final):
    try:
        total_historico = df_final[df_final['tipo_geografia_general'] == 'Despachos - RBS'].iloc[0]
        ipn_anterior_total = total_historico['ipn_3q']
        casos_anterior_total = total_historico['n_3q']
    except IndexError:
        ipn_anterior_total = 0.0
        casos_anterior_total = 0

    for shape in slide.shapes:
        # A. TABLA PRINCIPAL DE DIFERENCIAS DETALLADAS
        if shape.name == 'Tabla_DIF_DP' and shape.has_table:
            tabla = shape.table
            num_cols = len(tabla.columns)

            for i, row_data in df_final.iterrows():
                f_idx = i + 1
                if f_idx >= len(tabla.rows):
                    break

                if num_cols > 0:
                    tabla.cell(f_idx, 0).text = f"{int(row_data['n_4q']):,}"
                if num_cols > 1:
                    tabla.cell(f_idx, 1).text = f"{row_data['ipn_4q']:.1f}%"
                if num_cols > 2:
                    tabla.cell(f_idx, 2).text = f"{row_data['ipn_3q']:.1f}%"
                if num_cols > 3:
                    d_val = row_data['dif']
                    signo = "+" if d_val > 0 else ""
                    tabla.cell(f_idx, 3).text = f"{signo}{d_val:.1f} pp"

                for col_idx in range(num_cols):
                    if col_idx in [0, 1, 2, 3]:
                        p = tabla.cell(f_idx, col_idx).text_frame.paragraphs[0]
                        p.alignment = PP_ALIGN.CENTER
                        run = p.runs[0] if p.runs else p.add_run()
                        run.font.name = 'Montserrat'
                        run.font.size = Pt(10)
                        run.font.bold = True

                print(f"Fila {f_idx}: {row_data['tipo_geografia_general']} actualizada en Tabla_DIF_SL.")

        # B. TABLA RESUMEN ANTERIOR MAESTRA
        elif shape.name == 'Tabla_ANTERIO_DP' and shape.has_table:
            tabla_ant = shape.table
            num_filas_ant = len(tabla_ant.rows)
            num_cols_ant = len(tabla_ant.columns)

            if num_filas_ant >= 2 and num_cols_ant >= 2:
                tabla_ant.cell(0, 1).text = f"{ipn_anterior_total:.1f}%"
                tabla_ant.cell(1, 1).text = f"{int(casos_anterior_total):,}"

                for f_idx in [0, 1]:
                    p = tabla_ant.cell(f_idx, 1).text_frame.paragraphs[0]
                    p.alignment = PP_ALIGN.CENTER
                    run = p.runs[0] if p.runs else p.add_run()
                    run.font.name = 'Montserrat'
                    run.font.size = Pt(10)
                    run.font.bold = True

                print("Éxito: 'Tabla_ANTERIO_DP' actualizada con métricas del periodo anterior.")


# ==============================================================================
# 4. SECCIÓN DONA — gráfica de dona + tablas resumen (Slide 6)
# ==============================================================================
def calcular_datos_dona(df_red):
    total_n_red = len(df_red)
    counts_red = df_red['calificacion_nps'].value_counts()
    p_prom = round((counts_red.get('Promotor', 0) / total_n_red) * 100, 1)
    p_pas = round((counts_red.get('Pasivo', 0) / total_n_red) * 100, 1)
    p_det = round((counts_red.get('Detractor', 0) / total_n_red) * 100, 1)
    ipn_red = round(p_prom - p_det, 1)
    return total_n_red, counts_red, p_prom, p_pas, p_det, ipn_red


def actualizar_dona(slide, df_red):
    total_n_red, counts_red, p_prom, p_pas, p_det, ipn_red = calcular_datos_dona(df_red)

    config_nps_red = {
        'Tabla_PRO_DP': ["Promotores", f"{p_prom:.1f}%", f"({int(counts_red.get('Promotor', 0)):,})"],
        'Tabla_PAS_DP': ["Pasivos", f"{p_pas:.1f}%", f"({int(counts_red.get('Pasivo', 0)):,})"],
        'Tabla_DET_DP': ["Detractores", f"{p_det:.1f}%", f"({int(counts_red.get('Detractor', 0)):,})"],
        'Tabla_IPN_DP': [f"{ipn_red:.1f}%", "IPN", f"({total_n_red:,})", "Colaboradores"]
    }

    for shape in slide.shapes:
        s_name = shape.name.strip()

        if s_name == 'Grafico _don_DP' and shape.has_chart:
            chart_data = CategoryChartData()
            chart_data.categories = ['Promotores', 'Pasivos', 'Detractores']
            chart_data.add_series('calificacion_nps', (p_prom, p_pas, p_det))
            shape.chart.replace_data(chart_data)

        elif s_name in config_nps_red:
            datos = config_nps_red[s_name]
            text_frame = shape.table.cell(0, 0).text_frame if shape.has_table else shape.text_frame
            aplicar_formato_texto(text_frame, datos[0], datos[1], datos[2])

    print("Slide de dona (Despacho) actualizada exitosamente.")


# ==============================================================================
# 5. EJECUCIÓN PRINCIPAL
# ==============================================================================
def main():
    print("Procesando archivos de Excel...")

    # Base actual (1Q 2026): una versión SIN excluir 'CASOS NO IDENTIFICADOS'
    # (para replicar exactamente el comportamiento original de la sección de
    # barras) y otra versión CON la exclusión (para dona y diferencia).
    df_1q_barras = cargar_y_filtrar_base(FILE_1Q_2026, excluir_no_identificados=False)
    df_1q_filtrada = cargar_y_filtrar_base(FILE_1Q_2026, excluir_no_identificados=True)

    # Base anterior (4Q 2025), usada solo por la sección de diferencia
    df_4q_filtrada = cargar_y_filtrar_base(FILE_4Q_2025, excluir_no_identificados=True)

    df_diferencia_final = calcular_diferencia_periodos(df_1q_filtrada, df_4q_filtrada)

    print("Abriendo presentación e inyectando datos...")
    prs = Presentation(TEMPLATE_PPTX)

    slide_5 = prs.slides[5]  # Índice 5 corresponde al Slide 6 (barras + diferencia + dona)

    actualizar_grafica_barras(slide_5, df_1q_barras)
    actualizar_tablas_diferencia(slide_5, df_diferencia_final)
    actualizar_dona(slide_5, df_1q_filtrada)

    prs.save(OUTPUT_PPTX)
    print(f"\n--- Proceso concluido: presentación guardada en '{OUTPUT_PPTX}' ---")


if __name__ == '__main__':
    main()