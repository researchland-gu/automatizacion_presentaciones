# ==============================================================================
# CONFIGURACIÓN CENTRAL DEL REPORTE IPN COLABORADORES
# Este es el ÚNICO archivo que necesitas editar cada trimestre:
#   1. Actualiza las rutas de FILE_ACTUAL / FILE_ANTERIOR
#   2. Actualiza NUEVO_PERIODO para el histórico
#   3. Si la plantilla cambió de nombres de shapes o índices, ajústalos aquí
# ==============================================================================

# ------------------------------------------------------------------------------
# RUTAS
# ------------------------------------------------------------------------------
FILE_ACTUAL = r'E:\Users\954108\Desktop\codigos_py\Presentaciones\IPN_Colaboradores\Base_resultados\Base para dashboard 1Q 2026.xlsx'
FILE_ANTERIOR = r'E:\Users\954108\Desktop\codigos_py\Presentaciones\IPN_Colaboradores\Base_resultados\Base para dashboard 4Q 2025 liberada.xlsx'
FILE_HISTORICO = r'E:\Users\954108\Desktop\codigos_py\Presentaciones\IPN_Colaboradores\Colaboradores\Base_historica_Colaboradores.xlsx'
FILE_HISTORICO_AGRUPADORES = r'E:\Users\954108\Desktop\codigos_py\Presentaciones\IPN_Colaboradores\Colaboradores\Base_historica_Grupo_Empresarial.xlsx'

# OJO: tus scripts originales usaban 2 plantillas distintas
# (la de "Satisfaccion y Asuntos Publicos" para slides 3-4 y "Reporte_Automatizado"
# para el resto). Verifica cuál es la definitiva y ponla aquí.
TEMPLATE_PPTX = r'E:\Users\954108\Desktop\codigos_py\Presentaciones\IPN_Colaboradores\Reporte_Automatizado.pptx'
OUTPUT_PPTX = 'Reporte_Automatizado_completo.pptx'

# Etiqueta del nuevo periodo en la gráfica histórica (eje X)
NUEVO_PERIODO = "2Q´26"

FUENTE = 'Montserrat'

# Confidencialidad estadística en tablas de diferencia (Tabla_DIF):
# los segmentos con este número de casos o menos muestran el conteo,
# pero enmascaran el IPN actual con '*' y el anterior/diferencia con '–'.
UMBRAL_CASOS_MINIMOS = 3

# ------------------------------------------------------------------------------
# TIPO 1: BLOQUES DE PARTICIPACIÓN (slides 3 y 4)
#   modo 'estandar'  -> tabla pct (N + %) y tabla de desglose por Nivel 3
#   modo 'compuesto' -> celda única "IPN – N – %" en el 2do renglón
# ------------------------------------------------------------------------------
COL_DIVISION_PARTICIPACION = 'Nivel 3'

BLOQUES_PARTICIPACION = [
    # --- Slide 3 (índice 2) ---
    {'slide_index': 2, 'filtro_col': 'TIPO',      'filtro_val': 'Normal',          'tabla_pct': 'Tabla_pct_00', 'tabla_desglose': None,       'modo': 'compuesto'},
    {'slide_index': 2, 'filtro_col': 'Agrupador', 'filtro_val': 'GRUPO ELEKTRA',   'tabla_pct': 'Tabla_pct_01', 'tabla_desglose': 'Tabla_01', 'modo': 'estandar'},
    {'slide_index': 2, 'filtro_col': 'Agrupador', 'filtro_val': 'DESPACHOS',       'tabla_pct': 'Tabla_pct_02', 'tabla_desglose': 'Tabla_02', 'modo': 'estandar'},
    {'slide_index': 2, 'filtro_col': 'Agrupador', 'filtro_val': 'GRUPO TOTALPLAY', 'tabla_pct': 'Tabla_pct_03', 'tabla_desglose': 'Tabla_03', 'modo': 'estandar'},
    {'slide_index': 2, 'filtro_col': 'Agrupador', 'filtro_val': 'OTROS NEGOCIOS',  'tabla_pct': 'Tabla_pct_04', 'tabla_desglose': 'Tabla_04', 'modo': 'estandar'},
    {'slide_index': 2, 'filtro_col': 'Agrupador', 'filtro_val': 'GRUPO TV AZTECA', 'tabla_pct': 'Tabla_pct_05', 'tabla_desglose': 'Tabla_05', 'modo': 'estandar'},
    # --- Slide 4 (índice 3) ---
    {'slide_index': 3, 'filtro_col': 'TIPO',      'filtro_val': 'Normal',                    'tabla_pct': 'Tabla_pct_10', 'tabla_desglose': None,       'modo': 'compuesto'},
    {'slide_index': 3, 'filtro_col': 'Agrupador', 'filtro_val': 'NEGOCIOS FAMILIA',          'tabla_pct': 'Tabla_pct_06', 'tabla_desglose': 'Tabla_06', 'modo': 'estandar'},
    {'slide_index': 3, 'filtro_col': 'Agrupador', 'filtro_val': 'SOPORTE ESTRATÉGICO',       'tabla_pct': 'Tabla_pct_07', 'tabla_desglose': 'Tabla_07', 'modo': 'estandar'},
    {'slide_index': 3, 'filtro_col': 'Agrupador', 'filtro_val': 'EDUCACIÓN',                 'tabla_pct': 'Tabla_pct_08', 'tabla_desglose': 'Tabla_08', 'modo': 'estandar'},
    {'slide_index': 3, 'filtro_col': 'Agrupador', 'filtro_val': 'OFICINA DE LA PRESIDENCIA', 'tabla_pct': 'Tabla_pct_09', 'tabla_desglose': 'Tabla_09', 'modo': 'estandar'},
]

# ------------------------------------------------------------------------------
# TIPO 2: BLOQUES NPS (barras apiladas + tabla de diferencia + dona)
# Slides 6 y 13-23. Cada bloque define:
#   - slide_index      : índice base 0 de la diapositiva
#   - nivel_col        : columna que segmenta la gráfica ('Nivel 4', 'Agrupador', ...)
#   - resp_col         : columna del responsable para la etiqueta "Nivel - Responsable"
#   - filtros          : lista de (columna, valor) aplicados en cascada
#   - etiqueta_total   : nombre de la fila/barra total (puede llevar \n)
#   - grafico_barras / grafico_dona / tabla_dif / tabla_anterior : nombres de shapes
#     (los nombres de dona a veces llevan espacio: 'Grafico _don_GS'; se copian tal cual)
#   - sufijo_dona      : sufijo de las 4 tablas de la dona (Tabla_PRO_{suf}, etc.)
# ------------------------------------------------------------------------------
FILTRO_GE = [('Agrupador', 'GRUPO ELEKTRA'), ('Nivel 3', 'MÉXICO')]

BLOQUES_NPS = [
    # Slide 6: Grupo Salinas por Agrupador
    {'slide_index': 5, 'nivel_col': 'Agrupador', 'resp_col': 'Resonsable_N2',
     'filtros': [('Resonsable_N1', 'RICARDO BENJAMIN SALINAS PLIEGO')],
     'etiqueta_total': 'Grupo Salinas',
     'grafico_barras': 'Grafico_bar_GS', 'grafico_dona': 'Grafico _don_GS',
     'tabla_dif': 'Tabla_DIF_GS', 'tabla_anterior': 'Tabla_ANTERIO_GS', 'sufijo_dona': 'GS'},

    # Slide 13: Grupo Elektra por Nivel 3
    {'slide_index': 12, 'nivel_col': 'Nivel 3', 'resp_col': 'Resonsable_N3',
     'filtros': [('Agrupador', 'GRUPO ELEKTRA')],
     'etiqueta_total': 'Grupo Elektra',
     'grafico_barras': 'Grafico_bar_GE', 'grafico_dona': 'Grafico _don_GE',
     'tabla_dif': 'Tabla_DIF_GE', 'tabla_anterior': 'Tabla_ANTERIO_UPAX', 'sufijo_dona': 'GE'},

    # Slide 14: GE México por Nivel 4
    {'slide_index': 13, 'nivel_col': 'Nivel 4', 'resp_col': 'Resonsable_N4',
     'filtros': FILTRO_GE,
     'etiqueta_total': 'Grupo Elektra - México',
     'grafico_barras': 'Grafico_bar_GE_ME', 'grafico_dona': 'Grafico _don_GE_ME',
     'tabla_dif': 'Tabla_DIF_GE_ME', 'tabla_anterior': 'Tabla_ANTERIO_GE_ME', 'sufijo_dona': 'GE_ME'},

    # Slide 15: Sector Mercado Masivo por Nivel 5
    {'slide_index': 14, 'nivel_col': 'Nivel 5', 'resp_col': 'Resonsable_N5',
     'filtros': FILTRO_GE + [('Nivel 4', 'SECTOR MERCADO MASIVO')],
     'etiqueta_total': 'Grupo Elektra\nMéxico\nSector Mercado Masivo',
     'grafico_barras': 'Grafico_bar_GE_ME_SMM', 'grafico_dona': 'Grafico _don_GE_ME_SMM',
     'tabla_dif': 'Tabla_DIF_GE_ME_SMM', 'tabla_anterior': 'Tabla_ANTERIO_GE_ME_SMM', 'sufijo_dona': 'GE_ME_SMM'},

    # Slide 16: Canales de Venta por Nivel 6
    {'slide_index': 15, 'nivel_col': 'Nivel 6', 'resp_col': 'Resonsable_N6',
     'filtros': FILTRO_GE + [('Nivel 4', 'SECTOR MERCADO MASIVO'), ('Nivel 5', 'CANALES DE VENTA')],
     'etiqueta_total': 'Grupo Elektra\nMéxico\nSector Mercado Masivo\nCanales de Venta',
     'grafico_barras': 'Grafico_bar_GE_ME_SMM_CV', 'grafico_dona': 'Grafico_don_GE_ME_SMM_CV',
     'tabla_dif': 'Tabla_DIF_GE_ME_SMM_CV', 'tabla_anterior': 'Tabla_ANTERIO_GE_ME_SMM_CV', 'sufijo_dona': 'GE_ME_SMM_CV'},

    # Slide 17: Canales Digitales por Nivel 7
    {'slide_index': 16, 'nivel_col': 'Nivel 7', 'resp_col': 'Resonsable_N7',
     'filtros': FILTRO_GE + [('Nivel 4', 'SECTOR MERCADO MASIVO'), ('Nivel 5', 'CANALES DE VENTA'), ('Nivel 6', 'CANALES DIGITALES')],
     'etiqueta_total': 'Grupo Elektra\nMéxico\nSector Mercado Masivo\nCanales de Venta\nCanales Digitales',
     'grafico_barras': 'Grafico_bar_GE_CD', 'grafico_dona': 'Grafico _don_GE_CD',
     'tabla_dif': 'Tabla_DIF_GE_CD', 'tabla_anterior': 'Tabla_ANTERIO_GE_CD', 'sufijo_dona': 'GE_CD'},

    # Slide 18: Empresas por Nivel 6
    {'slide_index': 17, 'nivel_col': 'Nivel 6', 'resp_col': 'Resonsable_N6',
     'filtros': FILTRO_GE + [('Nivel 4', 'SECTOR MERCADO MASIVO'), ('Nivel 5', 'EMPRESAS')],
     'etiqueta_total': 'Grupo Elektra\nMéxico\nSector Mercado Masivo\nEmpresas',
     'grafico_barras': 'Grafico_bar_GE_E', 'grafico_dona': 'Grafico _don_GE_E',
     'tabla_dif': 'Tabla_DIF_GE_E', 'tabla_anterior': 'Tabla_ANTERIO_GE_E', 'sufijo_dona': 'GE_E'},

    # Slide 19: Unidades de Negocio por Nivel 6
    {'slide_index': 18, 'nivel_col': 'Nivel 6', 'resp_col': 'Resonsable_N6',
     'filtros': FILTRO_GE + [('Nivel 4', 'SECTOR MERCADO MASIVO'), ('Nivel 5', 'UNIDADES DE NEGOCIO')],
     'etiqueta_total': 'Grupo Elektra\nMéxico\nSector Mercado Masivo\nUnidades de Negocio',
     'grafico_barras': 'Grafico_bar_GE_UD', 'grafico_dona': 'Grafico _don_GE_UD',
     'tabla_dif': 'Tabla_DIF_GE_UD', 'tabla_anterior': 'Tabla_ANTERIO_GE_UD', 'sufijo_dona': 'GE_UD'},

    # Slide 20: Sector Mercado Corporativo por Nivel 5
    {'slide_index': 19, 'nivel_col': 'Nivel 5', 'resp_col': 'Resonsable_N5',
     'filtros': FILTRO_GE + [('Nivel 4', 'SECTOR MERCADO CORPORATIVO')],
     'etiqueta_total': 'Grupo Elektra\nMéxico\nSector Mercado Corporativo',
     'grafico_barras': 'Grafico_bar_GE_SMC', 'grafico_dona': 'Grafico _don_GE_SMC',
     'tabla_dif': 'Tabla_DIF_GE_SMC', 'tabla_anterior': 'Tabla_ANTERIO_GE_SMC', 'sufijo_dona': 'GE_SMC'},

    # Slide 21: BIG por Nivel 6
    {'slide_index': 20, 'nivel_col': 'Nivel 6', 'resp_col': 'Resonsable_N6',
     'filtros': FILTRO_GE + [('Nivel 4', 'SECTOR MERCADO CORPORATIVO'), ('Nivel 5', 'BIG')],
     'etiqueta_total': 'Grupo Elektra\nMéxico\nSector Mercado Corporativo\nBig',
     'grafico_barras': 'Grafico_bar_GE_BIG', 'grafico_dona': 'Grafico _don_GE_BIG',
     'tabla_dif': 'Tabla_DIF_GE_BIG', 'tabla_anterior': 'Tabla_ANTERIO_GE_BIG', 'sufijo_dona': 'GE_BIG'},

    # Slide 22: BIG Canales por Nivel 7
    {'slide_index': 21, 'nivel_col': 'Nivel 7', 'resp_col': 'Resonsable_N7',
     'filtros': FILTRO_GE + [('Nivel 4', 'SECTOR MERCADO CORPORATIVO'), ('Nivel 5', 'BIG'), ('Nivel 6', 'CANALES')],
     'etiqueta_total': 'Grupo Elektra\nMéxico\nSector Mercado Corporativo\nBig\nCanales',
     'grafico_barras': 'Grafico_bar_GE_CANAL', 'grafico_dona': 'Grafico _don_GE_CANAL',
     'tabla_dif': 'Tabla_DIF_GE_CANAL', 'tabla_anterior': 'Tabla_ANTERIO_GE_CANAL', 'sufijo_dona': 'GE_CANAL'},

    # Slide 23: BIG Segmentos por Nivel 7
    {'slide_index': 22, 'nivel_col': 'Nivel 7', 'resp_col': 'Resonsable_N7',
     'filtros': FILTRO_GE + [('Nivel 4', 'SECTOR MERCADO CORPORATIVO'), ('Nivel 5', 'BIG'), ('Nivel 6', 'SEGMENTOS')],
     'etiqueta_total': 'Grupo Elektra\nMéxico\nSector Mercado Corporativo\nBig\nSegmentos',
     'grafico_barras': 'Grafico_bar_GE_SEG', 'grafico_dona': 'Grafico _don_GE_SEG',
     'tabla_dif': 'Tabla_DIF_GE_SEG', 'tabla_anterior': 'Tabla_ANTERIO_GE_SEG', 'sufijo_dona': 'GE_SEG'},

    # Slide 24: BIG Unidades de Negocio por Nivel 7
    {'slide_index': 23, 'nivel_col': 'Nivel 7', 'resp_col': 'Resonsable_N7',
     'filtros': FILTRO_GE + [('Nivel 4', 'SECTOR MERCADO CORPORATIVO'), ('Nivel 5', 'BIG'), ('Nivel 6', 'UNIDADES DE NEGOCIO')],
     'etiqueta_total': 'Grupo Elektra\nMéxico\nSector Mercado Corporativo\nBig\nUnidades de Negocios',
     'grafico_barras': 'Grafico_bar_GE_UDN', 'grafico_dona': 'Grafico _don_GE_UDN',
     'tabla_dif': 'Tabla_DIF_GE_UDN', 'tabla_anterior': 'Tabla_ANTERIO_GE_UDN', 'sufijo_dona': 'GE_UDN'},

    # Slide 25: Empresas de Datos por Nivel 6
    {'slide_index': 24, 'nivel_col': 'Nivel 6', 'resp_col': 'Resonsable_N6',
     'filtros': FILTRO_GE + [('Nivel 4', 'SECTOR MERCADO CORPORATIVO'), ('Nivel 5', 'EMPRESAS DE DATOS')],
     'etiqueta_total': 'Grupo Elektra\nMéxico\nSector Mercado Corporativo\nEmpresas de Datos',
     'grafico_barras': 'Grafico_bar_GE_EDD', 'grafico_dona': 'Grafico _don_GE_EDD',
     'tabla_dif': 'Tabla_DIF_GE_EDD', 'tabla_anterior': 'Tabla_ANTERIO_GE_EDD', 'sufijo_dona': 'GE_EDD'},

    # Slide 26: GE Comité de Supervisión Países por Nivel 4
    {'slide_index': 25, 'nivel_col': 'Nivel 4', 'resp_col': 'Resonsable_N4',
     'filtros': [('Agrupador', 'GRUPO ELEKTRA'), ('Nivel 3', 'COMITÉ DE SUPERVISIÓN PAÍSES')],
     'etiqueta_total': 'Grupo Elektra\nComité de Supervisión Países',
     'grafico_barras': 'Grafico_bar_GE_COP', 'grafico_dona': 'Grafico _don_GE_COP',
     'tabla_dif': 'Tabla_DIF_GE_COP', 'tabla_anterior': 'Tabla_ANTERIO_GE_COP', 'sufijo_dona': 'GE_COP'},

    # Slide 27: Despachos por Nivel 3
    {'slide_index': 26, 'nivel_col': 'Nivel 3', 'resp_col': 'Resonsable_N3',
     'filtros': [('Agrupador', 'DESPACHOS')],
     'etiqueta_total': 'Despachos\nRicardo Benjamin Salinas Pliego',
     'grafico_barras': 'Grafico_bar_GE_DP', 'grafico_dona': 'Grafico _don_GE_DP',
     'tabla_dif': 'Tabla_DIF_GE_DP', 'tabla_anterior': 'Tabla_ANTERIO_GE_DP', 'sufijo_dona': 'GE_DP'},

    # Slide 28: Grupo TV Azteca por Nivel 4
    {'slide_index': 27, 'nivel_col': 'Nivel 4', 'resp_col': 'Resonsable_N4',
     'filtros': [('Agrupador', 'GRUPO TV AZTECA')],
     'etiqueta_total': 'Grupo Tv Azteca',
     'grafico_barras': 'Grafico_bar_TV', 'grafico_dona': 'Grafico _don_GE_TV',
     'tabla_dif': 'Tabla_DIF_TV', 'tabla_anterior': 'Tabla_ANTERIO_TV', 'sufijo_dona': 'GE_TV'},

    # Slide 29: Grupo Totalplay por Nivel 3
    {'slide_index': 28, 'nivel_col': 'Nivel 3', 'resp_col': 'Resonsable_N3',
     'filtros': [('Agrupador', 'GRUPO TOTALPLAY')],
     'etiqueta_total': 'Totalplay',
     'grafico_barras': 'Grafico_bar_TP', 'grafico_dona': 'Grafico _don_GE_TP',
     'tabla_dif': 'Tabla_DIF_TP', 'tabla_anterior': 'Tabla_ANTERIO_TP', 'sufijo_dona': 'GE_TP'},

    # Slide 30: Negocios Familia por Nivel 4
    {'slide_index': 29, 'nivel_col': 'Nivel 4', 'resp_col': 'Resonsable_N4',
     'filtros': [('Agrupador', 'NEGOCIOS FAMILIA')],
     'etiqueta_total': 'Negocios Familia\nRBS',
     'grafico_barras': 'Grafico_bar_NG', 'grafico_dona': 'Grafico _don_GE_NG',
     'tabla_dif': 'Tabla_DIF_NG', 'tabla_anterior': 'Tabla_ANTERIO_NG', 'sufijo_dona': 'GE_NG'},

    # Slide 31: Otros Negocios por Nivel 3
    {'slide_index': 30, 'nivel_col': 'Nivel 3', 'resp_col': 'Resonsable_N3',
     'filtros': [('Agrupador', 'OTROS NEGOCIOS')],
     'etiqueta_total': 'Otros Negocios\nRBS',
     'grafico_barras': 'Grafico_bar_GE_ON', 'grafico_dona': 'Grafico _don_GE_ON',
     'tabla_dif': 'Tabla_DIF_GE_ON', 'tabla_anterior': 'Tabla_ANTERIO_GE_ON', 'sufijo_dona': 'GE_ON'},

    # Slide 32: Soporte Estratégico por Nivel 3
    {'slide_index': 31, 'nivel_col': 'Nivel 3', 'resp_col': 'Resonsable_N3',
     'filtros': [('Agrupador', 'SOPORTE ESTRATÉGICO')],
     'etiqueta_total': 'Soporte Estratégico\nRBS',
     'grafico_barras': 'Grafico_bar_GE_SP', 'grafico_dona': 'Grafico _don_GE_SP',
     'tabla_dif': 'Tabla_DIF_GE_SP', 'tabla_anterior': 'Tabla_ANTERIO_GE_SP', 'sufijo_dona': 'GE_SP'},

    # Slide 33: Educación por Nivel 3
    {'slide_index': 32, 'nivel_col': 'Nivel 3', 'resp_col': 'Resonsable_N3',
     'filtros': [('Agrupador', 'EDUCACIÓN')],
     'etiqueta_total': 'Educación',
     'grafico_barras': 'Grafico_bar_GE_ED', 'grafico_dona': 'Grafico _don_GE_ED',
     'tabla_dif': 'Tabla_DIF_GE_ED', 'tabla_anterior': 'Tabla_ANTERIO_GE_ED', 'sufijo_dona': 'GE_ED'},

    # Slide 34: Oficina de la Presidencia por Nivel 3
    {'slide_index': 33, 'nivel_col': 'Nivel 3', 'resp_col': 'Resonsable_N3',
     'filtros': [('Agrupador', 'OFICINA DE LA PRESIDENCIA')],
     'etiqueta_total': 'Oficina de la Presidencia',
     'grafico_barras': 'Grafico_bar_OP', 'grafico_dona': 'Grafico _don_GE_OP',
     'tabla_dif': 'Tabla_DIF_OP', 'tabla_anterior': 'Tabla_ANTERIO_OP', 'sufijo_dona': 'GE_OP'},
]

# ------------------------------------------------------------------------------
# TIPO 3: MENCIONES TOP 5 (slides 8, 9 y 10)
# ------------------------------------------------------------------------------
GRUPOS_ORDENADOS = [
    'GRUPO ELEKTRA', 'DESPACHOS', 'GRUPO TOTALPLAY', 'GRUPO TV AZTECA',
    'NEGOCIOS FAMILIA', 'OTROS NEGOCIOS', 'SOPORTE ESTRATÉGICO', 'EDUCACIÓN', 'OFICINA DE LA PRESIDENCIA'
]

BLOQUES_MENCIONES = [
    {'slide_index': 7, 'categoria': 'Promotor',  'tabla': 'Tabla_Men_Prom'},
    {'slide_index': 8, 'categoria': 'Pasivo',    'tabla': 'Tabla_Men_Pas'},
    {'slide_index': 9, 'categoria': 'Detractor', 'tabla': 'Tabla_Men_Detr'},
]

# ------------------------------------------------------------------------------
# TIPO 4: HISTÓRICO (slide 7)
# ------------------------------------------------------------------------------
HISTORICO = {
    'slide_index': 6,
    'grafico': 'Grafico_historico_GS',
    'tabla_pct': 'Tabla_pct_historicos',
    'tabla_casos': 'Tabla_casos_historicos',
    'filtro_responsable': ('Resonsable_N1', 'RICARDO BENJAMIN SALINAS PLIEGO'),
}

# ------------------------------------------------------------------------------
# TIPO 5: HISTÓRICO POR AGRUPADOR (slide 12)
# 8 gráficas de barras (una por agrupador) alimentadas por un segundo Excel
# histórico. La barra del periodo más reciente se pinta en gris oscuro.
# ------------------------------------------------------------------------------
HISTORICO_AGRUPADORES = {
    'slide_index': 11,
    'mapa_graficas': {
        'GRUPO ELEKTRA': 'Grafico_Historico_GE',
        'DESPACHOS': 'Grafico_Historico_DP',
        'GRUPO TOTALPLAY': 'Grafico_Historico_GT',
        'GRUPO TV AZTECA': 'Grafico_Historico_TA',
        'OTROS NEGOCIOS': 'Grafico_Historico_ON',
        'NEGOCIOS FAMILIA': 'Grafico_Historico_NF',
        'SOPORTE ESTRATÉGICO': 'Grafico_Historico_SE',
        'OFICINA DE LA PRESIDENCIA': 'Grafico_Historico_OP',
    },
}

# ------------------------------------------------------------------------------
# TIPO 6: FORMATOS DE TARJETAS DE DONA (paso final)
# Estilos tipográficos por línea (0, 1, 2) de las tarjetas Tabla_DET/PAS/PRO
# (regla 'DET_PAS_PRO') y Tabla_IPN (regla 'IPN') de cada bloque NPS.
# El catálogo de slides/tablas se deriva automáticamente de BLOQUES_NPS,
# así que un bloque nuevo se formatea solo, sin tocar nada más.
# Tamaños en puntos; colores en RGB.
# ------------------------------------------------------------------------------
ESTILOS_TARJETAS = {
    'DET_PAS_PRO': {
        0: {'size': 14,   'color': (0, 0, 0),       'bold': True},
        1: {'size': 14,   'color': (0, 0, 0),       'bold': True},
        2: {'size': 10.5, 'color': (127, 127, 127), 'bold': False},
    },
    'IPN': {
        0: {'size': 32, 'color': (0, 0, 0),       'bold': True},
        1: {'size': 11, 'color': (0, 0, 0),       'bold': True},
        2: {'size': 10, 'color': (127, 127, 127), 'bold': False},
    },
}

# ------------------------------------------------------------------------------
# LISTAS DE MOTIVOS PARA COLOREAR MENCIONES
# ------------------------------------------------------------------------------
MOTIVOS_NEGATIVOS = [
    "No hay apoyo de Capital Humano", "Hay demasiada presión para los colaboradores",
    "No se preocupan por sus colaboradores", "No se cotiza al 100%", "No cumplen lo que promete",
    "NO SABE / NO CONTESTÓ", "El ambiente laboral es malo", "Horario de trabajo extenso / no veo a mi familia",
    "Solo un día de descanso / descanso entre semana", "Me hacen descuentos adicionales (merma, faltantes, errores)",
    "No hay suficientes prestaciones/beneficios",
    "El trabajo no es para cualquier persona / debe ser apto para el puesto",
    "El sueldo es poco / sueldo garantía", "Ha habido muchos cambios que han afectado",
    "La empresa no da utilidades / pocas utilidades", "Falta organización / estructura / mejoras en el área",
    "Inestabilidad laboral / amenazas de despido / alta rotación de personal",
    "No me gusta el esquema de pago / compensación",
    "Se trabajan días festivos", "No respetan el día de descanso / te hablan o mandan mensaje",
    "Falta o poca comunicación con los colaboradores", "No hay home office", "Los líderes tienen mala actitud",
    "Hay favoritismo para crecer", "No se pagan horas extras",
    "No hay recursos suficientes (gasolina, transporte, uniforme, teléfono, computadoras)",
    "Los líderes / jefes no están preparados", "No hay oportunidades de crecimiento",
    "No hay claridad en la forma de pagos", "No hay claridad en procesos de planes de crecimiento / desarrollo",
    "No se reconoce tu esfuerzo / Falta de motivación", "Cambian de horarios a conveniencia",
    "Distribución injusta de cargas de trabajo", "No permiten la contratación de familiares",
    "Los trámites son burocráticos / tardados", "No hay trabajo en equipo",
    "Perjuicios derivados de la unificación (Altos impuestos, desaparición de bonos)",
    "Empresa con mal prestigio (engañan a los clientes / malos productos / procesos lentos)",
    "No hay igualdad en todos los negocios", "No hay capacitación / mejorar la capacitación", "Falta personal",
    "Ya he recomendado amigos y familiares y no los aceptan",
    "No me agrada el esquema /las  condiciones de contratación",
    "Fallan en los pagos/hay atrasos en los sueldos", "Instalaciones deficientes",
    "No se respetan las políticas de la empresa", "No hay claridad en prestaciones/beneficios",
    "Pocos días de vacaciones / no me dan vacaciones", "Empresa sin valores / principios",
    "Para no afectar el resultado/ Subir el IPN", "No me informan de los cambios",
    "Las vacantes son de ventas / campo / andar en la calle / caminar", "Me obligan a calificar con 10",
    "Herramientas no adecuadas", "Más seguridad / debería haber un guardia", "No hay incentivos/ bonos",
    "Mal manejo de la pandemia", "No toman en cuenta mis opiniones",
    "Falta de actividades grupales de convivencia/ esparcimiento"
]

MOTIVOS_POSITIVOS = [
    "La empresa está a la vanguardia / Cambios positivos",
    "Estoy orgulloso de pertenecer al grupo / Es una buena empresa /recomendable",
    "Es una empresa con valores", "La empresa reconoce mi esfuerzo", "La empresa te brinda su apoyo en todo momento",
    "La empresa me permite aprender y crecer", "Buen salario", "Hay buen ambiente de trabajo", "Me gusta mi trabajo",
    "La empresa me da estabilidad (laboral/económica)",
    "La empresa me ayuda a mejorar mi calidad de vida y la de mi familia",
    "La empresa tiene buenas prestaciones y beneficios", "Me dan todas las herramientas para trabajar",
    "Toman en cuenta mis opiniones", "Los horarios son buenos / cómodos / flexibles",
    "Se trabaja en equipo / hay compañerismo / unión", "Gano de acuerdo a mi esfuerzo/ Buenas comisiones",
    "Te asignan a sucursales cerca del domicilio", "Me dan capacitación constante", "La empresa da buenas utilidades",
    "Buena organización / definición de actividades", "Líderes te apoyan/ Te escuchan", "Comisiones bajas",
    "Te asignan a sucursales lejos del domicilio", "Los líderes / jefes están preparados",
    "Me gusta el esquema de pago",
    "Cuenta con home office", "Los sueldos / salarios son puntuales", "Líderes / jefes con buena actitud",
    "Para no tener represalias", "Empresa brinda un gran servicio al cliente",
    "Hay gente comprometida / profesional dentro del grupo", "Las instalaciones son las idóneas",
    "Porque me agrada la encuesta que se está realizando", "Buen manejo de la pandemia",
    "Me informan de manera oportuna de los cambios", "Para no tener afectaciones económicas"
]
