# ==============================================================================
# CONFIGURACIÓN CENTRAL DEL REPORTE RESEARCHLAND
# Este es el ÚNICO archivo que necesitas editar cada trimestre:
#   1. Actualiza las rutas de FILE_ACTUAL / FILE_ANTERIOR
#   2. Verifica NUEVO_PERIODO (etiqueta del eje X del histórico)
#
# Plantilla oficial: 'IPN_Colaboradores Researchland_automatizada.pptx'.
# El reporte cubre: NPS de Upax (slide 3, shapes UPAX), Research Land
# (slide 4, shapes RL) y Customer Insight (slide 5, shapes CI); histórico de
# líneas de Research Land (slide 6, Grafico_historico_RL + Base_historica_RL);
# y menciones Top 10 con rubro Tabla_PRO/PAS/DET_MENS_RL (slides 7-9).
#
# NOTA: en este reporte NINGÚN bloque filtra TIPO == 'Normal' (fiel a los
# scripts originales) y las etiquetas de geografía usan guion (' - ').
# ==============================================================================

# ------------------------------------------------------------------------------
# RUTAS
# ------------------------------------------------------------------------------
FILE_ACTUAL = r'E:\Users\954108\Desktop\codigos_py\Presentaciones\IPN_Colaboradores\Base_resultados\Base para dashboard 1Q 2026.xlsx'
FILE_ANTERIOR = r'E:\Users\954108\Desktop\codigos_py\Presentaciones\IPN_Colaboradores\Base_resultados\Base para dashboard 4Q 2025 liberada.xlsx'
TEMPLATE_PPTX = r'E:\Users\954108\Desktop\Satisfaccion  y Asuntos Publicos\Automatizaciones_PPT\IPN Colaboradores\IPN_Colaboradores Researchland_automatizada.pptx'
OUTPUT_PPTX = 'Reporte_Researchland.pptx'

# Este reporte no actualiza Excel históricos, pero la app y main comparten
# interfaz con el proyecto principal, por eso el periodo sigue existiendo
# (se usa solo para nombrar el archivo de salida desde la app web).
NUEVO_PERIODO = "2Q´26"

FUENTE = 'Montserrat'

# En este reporte las etiquetas de geografía van en dos líneas
# ('Nombre\nResponsable'); en el reporte principal van con ' - '.
SEPARADOR_GEOGRAFIA = ' - '

# Confidencialidad estadística en tablas de diferencia (Tabla_DIF):
# los segmentos con este número de casos o menos muestran el conteo,
# pero enmascaran el IPN actual con '*' y el anterior/diferencia con '–'.
UMBRAL_CASOS_MINIMOS = 3

# (Usada por funciones compartidas de participación; no aplica en este reporte)
COL_DIVISION_PARTICIPACION = 'Nivel 3'
BLOQUES_PARTICIPACION = []

# ------------------------------------------------------------------------------
# BLOQUES NPS (barras apiladas + tabla de diferencia + dona)
# ------------------------------------------------------------------------------
BLOQUES_NPS = [
    # Slide 3: Upax por Nivel 4
    {'slide_index': 2, 'nivel_col': 'Nivel 4', 'resp_col': 'Resonsable_N4',
     'filtros': [('Nivel 3', 'UPAX')],
     'etiqueta_total': 'Upax - Cecilia Laura Fallabrino',
     'grafico_barras': 'Grafico_bar_UPAX', 'grafico_dona': 'Grafico _don_UPAX',
     'tabla_dif': 'Tabla_DIF_UPAX', 'tabla_anterior': 'Tabla_ANTERIO_UPAX', 'sufijo_dona': 'UPAX'},

    # Slide 4: Research Land por Nivel 5 (dona SIN espacio en el nombre)
    {'slide_index': 3, 'nivel_col': 'Nivel 5', 'resp_col': 'Resonsable_N5',
     'filtros': [('Nivel 4', 'RESEARCH LAND')],
     'etiqueta_total': 'Researchland - Pablo Federico Levy Covarrubias',
     'grafico_barras': 'Grafico_bar_RL', 'grafico_dona': 'Grafico_don_RL',
     'tabla_dif': 'Tabla_DIF_RL', 'tabla_anterior': 'Tabla_ANTERIO_RL', 'sufijo_dona': 'RL'},

    # Slide 5: Customer Insight por Nivel 6 (dona SIN espacio en el nombre)
    {'slide_index': 4, 'nivel_col': 'Nivel 6', 'resp_col': 'Resonsable_N6',
     'filtros': [('Nivel 5', 'CUSTOMER INSIGHT')],
     'etiqueta_total': 'Customer Insight - Alicia Sarkis Saad',
     'grafico_barras': 'Grafico_bar_CI', 'grafico_dona': 'Grafico_don_CI',
     'tabla_dif': 'Tabla_DIF_CI', 'tabla_anterior': 'Tabla_ANTERIO_CI', 'sufijo_dona': 'CI'},
]

# ------------------------------------------------------------------------------
# MENCIONES TOP 10 CON RUBRO (slides 7, 8 y 9)
# Una tabla por slide: motivo (color según lista), menciones (color) y
# rubro (gris). Filtro fijo del reporte: Nivel 4 == RESEARCH LAND.
# ------------------------------------------------------------------------------
FILTRO_MENCIONES = ('Nivel 4', 'RESEARCH LAND')
TOP_N_MENCIONES = 10

BLOQUES_MENCIONES_RUBRO = [
    {'slide_index': 6, 'categoria': 'Promotor',  'tabla': 'Tabla_PRO_MENS_RL'},
    {'slide_index': 7, 'categoria': 'Pasivo',    'tabla': 'Tabla_PAS_MENS_RL'},
    {'slide_index': 8, 'categoria': 'Detractor', 'tabla': 'Tabla_DET_MENS_RL'},
]

# (Lista usada por el matcheo de colores; compartida con el proyecto principal)
GRUPOS_ORDENADOS = ['RESEARCH LAND']

# ------------------------------------------------------------------------------
# FORMATOS DE TARJETAS DE DONA (paso final; derivado de BLOQUES_NPS)
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


# ------------------------------------------------------------------------------
# HISTÓRICO RESEARCH LAND (slide 6)
# Actualiza el Excel histórico vertical (columnas Periodo/IPN/Promotores/
# Pasivos/Detractores) e inyecta la gráfica de líneas con 4 series.
# Filtro: calificación válida + Respondida == 'Si' + Nivel 4 == RESEARCH LAND
# (sin filtro de TIPO, fiel al script original).
# ------------------------------------------------------------------------------
FILE_HISTORICO_RL = r'E:\Users\954108\Desktop\codigos_py\Presentaciones\IPN_Colaboradores\Base_historica_RL.xlsx'

HISTORICO_RL = {
    'slide_index': 5,
    'grafico': 'Grafico_historico_RL',
    'filtro': ('Nivel 4', 'RESEARCH LAND'),
}
