# ==============================================================================
# ORQUESTADOR DEL REPORTE OFICINA DE LA PRESIDENCIA
# (slide 3: NPS Educación; slides 6-8: menciones Top 10 con rubro)
# Lee las bases UNA vez, abre la plantilla UNA vez, procesa los bloques NPS
# (slides 4 y 6), las menciones Top 10 con rubro (slides 7-9), aplica los
# formatos de tarjetas y guarda UN solo archivo final.
#
# Uso en PyCharm: correr este archivo, o levantar la app web con
#     streamlit run app.py
# ==============================================================================
import traceback

import pandas as pd
from pptx import Presentation

import config
import calculos
import inyectores


def procesar_bloques_nps(prs, df_actual, df_anterior, errores):
    print("\n================ NPS: BARRAS + DIFERENCIA + DONA ================")
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

            if len(df_1q) == 0:
                print(f"⚠ El filtro de este bloque no arrojó ningún caso en la base actual.")
                print(f"  Filtros aplicados: {bloque['filtros']}")
                df_diag = df_actual.copy()
                for col, val in bloque['filtros']:
                    if col in df_diag.columns:
                        disponibles = sorted(df_diag[col].dropna().astype(str).str.strip().unique())[:15]
                        print(f"  Valores disponibles en '{col}': {disponibles}")
                        df_diag = df_diag[df_diag[col] == val]
                    else:
                        print(f"  ⚠ La columna '{col}' NO existe en la base.")
                print(f"  Corrige columnas/valores en config.py para que coincidan con la base.")
                errores.append(f"{nombre} (filtro sin casos)")
                continue

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


def procesar_menciones_rubro(prs, df_actual, errores):
    print("\n================ MENCIONES TOP 10 CON RUBRO ================")
    for bloque in config.BLOQUES_MENCIONES_RUBRO:
        nombre = f"Menciones {bloque['categoria']} (slide idx {bloque['slide_index']})"
        try:
            print(f"\n--- {nombre} ---")
            slide = prs.slides[bloque['slide_index']]
            top = calculos.calcular_top10_menciones_rubro(
                df_actual, bloque['categoria'], config.FILTRO_MENCIONES,
                top_n=config.TOP_N_MENCIONES)
            inyectores.inyectar_menciones_rubro(slide, bloque['tabla'], top)
        except Exception:
            errores.append(nombre)
            print(f"❌ Error en {nombre}:\n{traceback.format_exc()}")


def procesar_formatos(prs, errores):
    """Paso final: aplica los estilos tipográficos a las tarjetas de la dona
    de todos los bloques NPS (el catálogo se deriva de BLOQUES_NPS)."""
    print("\n================ FORMATOS DE TARJETAS ================")
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

    print("Leyendo bases de Excel (una sola vez cada una)...")
    print(f"  Base actual:   {config.FILE_ACTUAL}")
    print(f"  Base anterior: {config.FILE_ANTERIOR}")
    df_actual = pd.read_excel(config.FILE_ACTUAL)
    df_anterior = pd.read_excel(config.FILE_ANTERIOR)

    print("Abriendo plantilla de PowerPoint...")
    prs = Presentation(config.TEMPLATE_PPTX)

    procesar_bloques_nps(prs, df_actual, df_anterior, errores)
    procesar_menciones_rubro(prs, df_actual, errores)
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
