# ==============================================================================
# ORQUESTADOR DEL REPORTE IPN COLABORADORES
# Lee las bases UNA vez, abre la plantilla UNA vez, procesa los 4 tipos de
# inyección para todos los slides y guarda UN solo archivo final.
#
# Uso en PyCharm: correr este archivo. Cada trimestre solo edita config.py.
# ==============================================================================
import os
import traceback

import pandas as pd
from pptx import Presentation

import config
import calculos
import inyectores


def procesar_participacion(prs, df_actual, errores):
    print("\n================ TIPO 1: PARTICIPACIÓN (slides 3-4) ================")
    for bloque in config.BLOQUES_PARTICIPACION:
        nombre = f"Participación {bloque['filtro_val']} (slide idx {bloque['slide_index']})"
        try:
            print(f"\n--- {nombre} ---")
            slide = prs.slides[bloque['slide_index']]
            df_resumen, validos_global, pct_global = calculos.calcular_participacion_nps(
                df_actual, bloque['filtro_col'], bloque['filtro_val']
            )

            shape_pct = inyectores.buscar_shape(slide, bloque['tabla_pct'], requiere_tabla=True)
            if shape_pct is not None:
                if bloque['modo'] == 'compuesto':
                    inyectores.inyectar_tabla_pct_compuesta(shape_pct, validos_global, pct_global)
                else:
                    inyectores.inyectar_tabla_pct(shape_pct, validos_global, pct_global)
            else:
                print(f"⚠ No se encontró la tabla '{bloque['tabla_pct']}'.")

            if bloque['tabla_desglose']:
                shape_desg = inyectores.buscar_shape(slide, bloque['tabla_desglose'], requiere_tabla=True)
                if shape_desg is not None:
                    inyectores.inyectar_tabla_desglose(shape_desg, df_resumen)
                else:
                    print(f"⚠ No se encontró la tabla '{bloque['tabla_desglose']}'.")
        except Exception:
            errores.append(nombre)
            print(f"❌ Error en {nombre}:\n{traceback.format_exc()}")


def procesar_bloques_nps(prs, df_actual, df_anterior, errores):
    print("\n================ TIPO 2: NPS BARRAS + DIFERENCIA + DONA ================")
    for bloque in config.BLOQUES_NPS:
        nombre = f"NPS {bloque['sufijo_dona']} (slide idx {bloque['slide_index']})"
        try:
            print(f"\n--- {nombre} ---")
            slide = prs.slides[bloque['slide_index']]

            # Barras: SIN excluir 'CASOS NO IDENTIFICADOS' (comportamiento original)
            df_barras = calculos.preparar_base_nps(df_actual, bloque, excluir_no_identificados=False)
            # Dona y diferencia: CON exclusión
            df_1q = calculos.preparar_base_nps(df_actual, bloque, excluir_no_identificados=True)
            df_4q = calculos.preparar_base_nps(df_anterior, bloque, excluir_no_identificados=True)

            categorias, l_det, l_pas, l_prom = calculos.calcular_datos_barras(
                df_barras, bloque['etiqueta_total'])
            inyectores.inyectar_grafica_barras(
                slide, bloque['grafico_barras'], categorias, l_det, l_pas, l_prom)

            df_dif = calculos.calcular_diferencia_periodos(df_1q, df_4q, bloque['etiqueta_total'])
            inyectores.inyectar_tablas_diferencia(slide, bloque, df_dif)

            datos_dona = calculos.calcular_datos_dona(df_1q)
            inyectores.inyectar_dona(slide, bloque, datos_dona)
        except Exception:
            errores.append(nombre)
            print(f"❌ Error en {nombre}:\n{traceback.format_exc()}")


def procesar_menciones(prs, df_actual, errores):
    print("\n================ TIPO 3: MENCIONES TOP 5 ================")
    for bloque in config.BLOQUES_MENCIONES:
        nombre = f"Menciones {bloque['categoria']} (slide idx {bloque['slide_index']})"
        try:
            print(f"\n--- {nombre} ---")
            slide = prs.slides[bloque['slide_index']]
            resultados = calculos.calcular_top5_menciones(df_actual, bloque['categoria'])
            inyectores.inyectar_menciones(slide, bloque['tabla'], resultados)
        except Exception:
            errores.append(nombre)
            print(f"❌ Error en {nombre}:\n{traceback.format_exc()}")


def procesar_historico(prs, df_actual, errores):
    print("\n================ TIPO 4: HISTÓRICO ================")
    nombre = "Histórico"
    try:
        metricas = calculos.calcular_metricas_historico(
            df_actual, config.HISTORICO['filtro_responsable'])

        # Al Excel histórico solo van las métricas de la gráfica de líneas;
        # participación y conteos se inyectan directo en las tablas del PPT.
        metricas_grafica = {k: metricas[k] for k in ['IPN', 'Promotores', 'Pasivos', 'Detractores']}

        print("Actualizando archivo histórico Excel (solo métricas de gráfica)...")
        df_hist = pd.read_excel(config.FILE_HISTORICO)
        df_hist = calculos.actualizar_base_historica(df_hist, metricas_grafica, config.NUEVO_PERIODO)
        df_hist.to_excel(config.FILE_HISTORICO, index=False)

        # Columna objetivo en las tablas históricas del PPT (+1 por los títulos)
        col_objetivo = df_hist.index[df_hist['Periodo'] == config.NUEVO_PERIODO].tolist()[0] + 1

        slide = prs.slides[config.HISTORICO['slide_index']]
        inyectores.inyectar_historico(slide, config.HISTORICO['grafico'], df_hist)
        inyectores.inyectar_tabla_pct_historicos(
            slide, config.HISTORICO['tabla_pct'], metricas, col_objetivo)
        inyectores.inyectar_tabla_casos_historicos(
            slide, config.HISTORICO['tabla_casos'], metricas, col_objetivo)
    except Exception:
        errores.append(nombre)
        print(f"❌ Error en {nombre}:\n{traceback.format_exc()}")


def procesar_historico_agrupadores(prs, df_actual, errores):
    print("\n================ TIPO 5: HISTÓRICO POR AGRUPADOR ================")
    nombre = "Histórico por agrupador"
    try:
        cfg = config.HISTORICO_AGRUPADORES
        agrupadores = list(cfg['mapa_graficas'].keys())

        metricas = calculos.calcular_ipn_por_agrupador(df_actual, agrupadores)

        print("Actualizando Excel histórico de agrupadores...")
        if os.path.exists(config.FILE_HISTORICO_AGRUPADORES):
            df_hist = pd.read_excel(config.FILE_HISTORICO_AGRUPADORES)
        else:
            df_hist = pd.DataFrame(columns=['Periodo'] + agrupadores)

        df_hist = calculos.actualizar_base_historica(df_hist, metricas, config.NUEVO_PERIODO)
        df_hist.to_excel(config.FILE_HISTORICO_AGRUPADORES, index=False)

        slide = prs.slides[cfg['slide_index']]
        n = inyectores.inyectar_historicos_agrupadores(slide, cfg['mapa_graficas'], df_hist)
        print(f"{n}/{len(agrupadores)} gráficas de agrupadores actualizadas.")
    except Exception:
        errores.append(nombre)
        print(f"❌ Error en {nombre}:\n{traceback.format_exc()}")


def procesar_formatos(prs, errores):
    """Paso final: aplica los estilos tipográficos a las tarjetas de la dona
    de todos los bloques NPS (el catálogo se deriva de BLOQUES_NPS)."""
    print("\n================ TIPO 6: FORMATOS DE TARJETAS ================")
    nombre = "Formatos de tarjetas"
    try:
        total = 0
        for bloque in config.BLOQUES_NPS:
            if bloque['slide_index'] < len(prs.slides):
                slide = prs.slides[bloque['slide_index']]
                total += inyectores.aplicar_formatos_dona(
                    slide, bloque['sufijo_dona'], config.ESTILOS_TARJETAS)
        print(f"Total aplicado: {total} tarjetas formateadas "
              f"(esperadas: {len(config.BLOQUES_NPS) * 4}).")
    except Exception:
        errores.append(nombre)
        print(f"❌ Error en {nombre}:\n{traceback.format_exc()}")


def ejecutar_reporte(file_actual=None, file_anterior=None, output_pptx=None,
                     nuevo_periodo=None, template_pptx=None):
    """Corre el pipeline completo. Los parámetros son opcionales: si no se
    pasan, se usan los valores por defecto de config.py. Esto permite invocar
    el reporte dinámicamente desde la app web (app.py) o desde otro script.

    Devuelve (ruta_del_pptx_generado, lista_de_errores)."""
    # Overrides dinámicos sobre config
    if file_actual:
        config.FILE_ACTUAL = file_actual
    if file_anterior:
        config.FILE_ANTERIOR = file_anterior
    if output_pptx:
        config.OUTPUT_PPTX = output_pptx
    if nuevo_periodo:
        config.NUEVO_PERIODO = nuevo_periodo
    if template_pptx:
        config.TEMPLATE_PPTX = template_pptx

    errores = []

    print(f"Periodo a procesar: {config.NUEVO_PERIODO}")
    print("Leyendo bases de Excel (una sola vez cada una)...")
    print(f"  Base actual:   {config.FILE_ACTUAL}")
    print(f"  Base anterior: {config.FILE_ANTERIOR}")
    df_actual = pd.read_excel(config.FILE_ACTUAL)
    df_anterior = pd.read_excel(config.FILE_ANTERIOR)

    print("Abriendo plantilla de PowerPoint...")
    prs = Presentation(config.TEMPLATE_PPTX)

    procesar_participacion(prs, df_actual, errores)
    procesar_bloques_nps(prs, df_actual, df_anterior, errores)
    procesar_menciones(prs, df_actual, errores)
    procesar_historico(prs, df_actual, errores)
    procesar_historico_agrupadores(prs, df_actual, errores)
    procesar_formatos(prs, errores)  # Siempre al final: pule lo ya inyectado

    prs.save(config.OUTPUT_PPTX)

    print("\n" + "=" * 70)
    print(f"Proceso concluido: presentación guardada en '{config.OUTPUT_PPTX}'")
    if errores:
        print(f"\n⚠ {len(errores)} bloque(s) fallaron (el resto se procesó bien):")
        for e in errores:
            print(f"   - {e}")
    else:
        print("Todos los bloques se procesaron sin errores. ✔")

    return config.OUTPUT_PPTX, errores


def main():
    """Ejecución clásica desde PyCharm/terminal con los valores de config.py."""
    ejecutar_reporte()


if __name__ == '__main__':
    main()
