# ==============================================================================
# CÁLCULOS: toda la lógica de pandas. No toca PowerPoint.
# ==============================================================================
import pandas as pd

from config import COL_DIVISION_PARTICIPACION, GRUPOS_ORDENADOS, SEPARADOR_GEOGRAFIA


# ------------------------------------------------------------------------------
# Funciones base compartidas
# ------------------------------------------------------------------------------
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


CATS_VALIDAS = ['Detractor', 'Pasivo', 'Promotor']


def construir_mapeo_geografia(df, col_nivel, col_resp):
    """Diccionario dinámico {nivel -> 'Nivel<SEP>Responsable'} en Title Case.

    FIX: si un mismo nivel tiene más de un responsable capturado, antes se
    quedaba con el ÚLTIMO que aparecía en drop_duplicates() (orden arbitrario
    del DataFrame). Ahora se usa el MÁS FRECUENTE (moda) y se avisa en
    consola cuando hay ambigüedad, para que el nombre mostrado sea estable
    entre corridas aunque no cambien los datos reales.
    """
    df[col_nivel] = df[col_nivel].fillna(' ').astype(str)
    df[col_resp] = df[col_resp].fillna('Sin Responsable').astype(str)

    mapeo = {}
    conteo_resp_por_nivel = df.groupby(col_nivel)[col_resp].apply(lambda s: s.str.strip().value_counts())

    for n_orig in df[col_nivel].unique():
        n_limpio = n_orig.strip()
        if n_limpio == "":
            mapeo[n_orig] = 'Sin Etiqueta'
            continue

        conteo = conteo_resp_por_nivel.get(n_orig)
        if conteo is None or len(conteo) == 0:
            resp_limpio = 'Sin Responsable'
        else:
            if len(conteo) > 1:
                print(f"AVISO: el nivel '{n_orig}' tiene {len(conteo)} responsables distintos "
                      f"({list(conteo.index)}); se usará el más frecuente: '{conteo.index[0]}'.")
            resp_limpio = conteo.index[0]

        mapeo[n_orig] = f"{n_limpio.title()}{SEPARADOR_GEOGRAFIA}{resp_limpio.title()}"
    return mapeo


# ------------------------------------------------------------------------------
# TIPO 1: Participación
# ------------------------------------------------------------------------------
def calcular_participacion_nps(df_raw, filtro_col, filtro_val):
    """Universo total vs respuestas válidas para un segmento específico."""
    df_ge = df_raw[df_raw[filtro_col] == filtro_val].copy()
    df_ge['calificacion_nps'] = df_ge['Calificacion'].apply(clasificar_nps)

    col = COL_DIVISION_PARTICIPACION
    df_ge[col] = df_ge[col].fillna('Sin Etiqueta').astype(str).str.title()

    total_universo = df_ge.groupby(col).size()
    df_validos = df_ge[df_ge['calificacion_nps'].isin(CATS_VALIDAS) & (df_ge['Respondida'] == 'Si')]
    total_validos = df_validos.groupby(col).size()

    df_resumen = pd.DataFrame({'Validos': total_validos, 'Total': total_universo}).fillna(0)
    df_resumen['Pct'] = (df_resumen['Validos'] / df_resumen['Total']) * 100

    validos_global = df_resumen['Validos'].sum()
    total_global = df_resumen['Total'].sum()
    pct_global = (validos_global / total_global) * 100 if total_global > 0 else 0

    df_resumen = df_resumen.sort_values(by='Validos', ascending=False).reset_index()
    return df_resumen, validos_global, pct_global


# ------------------------------------------------------------------------------
# TIPO 2: Bloques NPS (barras + diferencia + dona)
# ------------------------------------------------------------------------------
def preparar_base_nps(df_raw, bloque, excluir_no_identificados=True):
    """Aplica a la base cruda el mapeo de geografía y la cascada de filtros
    de un bloque NPS. Devuelve el DataFrame filtrado listo para calcular."""
    df = df_raw.copy()
    nivel_col = bloque['nivel_col']

    if excluir_no_identificados:
        df = df[df[nivel_col] != 'CASOS NO IDENTIFICADOS'].copy()

    mapeo = construir_mapeo_geografia(df, nivel_col, bloque['resp_col'])
    df['tipo_geografia_general'] = df[nivel_col].map(mapeo)
    # FIX: clave_nivel es el valor crudo (antes de etiquetar) y es lo que se
    # usa para cruzar periodo actual vs. anterior, para no depender del
    # nombre del responsable (que puede cambiar entre periodos).
    df['clave_nivel'] = df[nivel_col].fillna(' ').astype(str)
    df['calificacion_nps'] = df['Calificacion'].apply(clasificar_nps)

    df = df[df['calificacion_nps'].isin(CATS_VALIDAS)].copy()
    df = df[df['Respondida'] == 'Si'].copy()
    df = df[df['TIPO'] == 'Normal'].copy()

    for col, val in bloque['filtros']:
        df = df[df[col] == val].copy()

    return df


def calcular_datos_barras(df, etiqueta_total):
    """(categorias_eje, lista_det, lista_pas, lista_prom) ordenados de mayor a
    menor número de casos, con la etiqueta del total fija al inicio."""
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
    categorias_eje = [etiqueta_total] + segmentos_ordenados

    lista_det = [det_gen] + [round(dist_seg.loc[cat, 'Detractor'], 1) for cat in categorias_eje[1:]]
    lista_pas = [pas_gen] + [round(dist_seg.loc[cat, 'Pasivo'], 1) for cat in categorias_eje[1:]]
    lista_prom = [prom_gen] + [round(dist_seg.loc[cat, 'Promotor'], 1) for cat in categorias_eje[1:]]

    return categorias_eje, lista_det, lista_pas, lista_prom


def calcular_resumen_nps(df, etiqueta_total):
    """n e IPN por segmento, más la fila del total.

    FIX: se agrupa por 'clave_nivel' (el valor crudo, estable entre periodos)
    en vez de 'tipo_geografia_general' (que incluye el nombre del responsable
    y puede cambiar de un periodo a otro). 'tipo_geografia_general' se
    conserva como columna de despliegue (la etiqueta que se ve en la tabla).
    """
    resumen = df.groupby('clave_nivel').apply(lambda x: pd.Series({
        'tipo_geografia_general': x['tipo_geografia_general'].iloc[0],
        'n': float(len(x)),
        'ipn': round((x['calificacion_nps'].value_counts(normalize=True).get('Promotor', 0) * 100), 5) -
               round((x['calificacion_nps'].value_counts(normalize=True).get('Detractor', 0) * 100), 5)
    }), include_groups=False).reset_index()

    t_n = float(len(df))
    t_ipn = round((df['calificacion_nps'].value_counts(normalize=True).get('Promotor', 0) * 100), 5) - \
            round((df['calificacion_nps'].value_counts(normalize=True).get('Detractor', 0) * 100), 5)

    df_total = pd.DataFrame([{'clave_nivel': etiqueta_total, 'tipo_geografia_general': etiqueta_total,
                              'n': t_n, 'ipn': t_ipn}])
    return pd.concat([df_total, resumen], ignore_index=True)


def calcular_diferencia_periodos(df_actual, df_anterior, etiqueta_total):
    """Cruza periodo actual vs anterior; devuelve DataFrame listo para inyectar.

    FIX: el cruce ahora es por 'clave_nivel' (estable) con how='left', para
    que un segmento que solo cambió de responsable entre periodos NO se
    pierda (antes, el merge por defecto era 'inner' y lo eliminaba en
    silencio, recorriendo todas las filas siguientes de la tabla).
    """
    data_actual = calcular_resumen_nps(df_actual, etiqueta_total)
    data_anterior = calcular_resumen_nps(df_anterior, etiqueta_total)

    df_final = data_actual.merge(
        data_anterior[['clave_nivel', 'n', 'ipn']],
        on='clave_nivel', how='left', suffixes=('_4q', '_3q'),
    )
    df_final['dif'] = df_final['ipn_4q'] - df_final['ipn_3q']

    df_total_row = df_final[df_final['clave_nivel'] == etiqueta_total]
    df_segmentos = df_final[df_final['clave_nivel'] != etiqueta_total].copy()
    df_segmentos = df_segmentos.sort_values(by='n_4q', ascending=False)

    return pd.concat([df_total_row, df_segmentos], ignore_index=True)


def calcular_datos_dona(df):
    total_n = len(df)
    counts = df['calificacion_nps'].value_counts()
    if total_n == 0:
        # Sin casos: todo en cero; la limpieza de ceros dejará las tarjetas vacías
        return 0, counts, 0.0, 0.0, 0.0, 0.0
    p_prom = round((counts.get('Promotor', 0) / total_n) * 100, 1)
    p_pas = round((counts.get('Pasivo', 0) / total_n) * 100, 1)
    p_det = round((counts.get('Detractor', 0) / total_n) * 100, 1)
    ipn = round(p_prom - p_det, 1)
    return total_n, counts, p_prom, p_pas, p_det, ipn


# ------------------------------------------------------------------------------
# TIPO 3: Menciones Top 5
# ------------------------------------------------------------------------------
def calcular_top5_menciones(df_raw, categoria):
    """Top 5 de motivos (SUBNETO1..7) por Agrupador para una categoría NPS."""
    df = df_raw.copy()
    df['calificacion_nps'] = df['Calificacion'].apply(clasificar_nps)

    df = df[df['calificacion_nps'].isin(CATS_VALIDAS)].copy()
    df = df[df['Respondida'] == 'Si'].copy()
    df = df[df['TIPO'] == 'Normal'].copy()
    df_cat = df[df['calificacion_nps'] == categoria].copy()

    valores_basura = ['', 'nan', '0', '0.0', ' ', 'none']
    resultados = {}

    for grupo in GRUPOS_ORDENADOS:
        df_g = df_cat[df_cat['Agrupador'] == grupo]

        lista_pares = []
        for i in range(1, 8):
            col_motivo = f"SUBNETO{i}"
            if col_motivo in df_g.columns:
                temp = df_g[[col_motivo]].dropna()
                temp.columns = ['Motivo']
                lista_pares.append(temp)

        if lista_pares:
            df_cons = pd.concat(lista_pares, ignore_index=True)
            df_cons['Motivo'] = df_cons['Motivo'].apply(lambda x: str(x).strip())
            df_cons = df_cons[~df_cons['Motivo'].str.lower().isin(valores_basura)]

            top_5 = df_cons.groupby('Motivo').size().reset_index(name='Menciones')
            top_5 = top_5.sort_values(by='Menciones', ascending=False).head(5).reset_index(drop=True)
            resultados[grupo] = top_5
        else:
            resultados[grupo] = pd.DataFrame(columns=['Motivo', 'Menciones'])

    return resultados


# ------------------------------------------------------------------------------
# TIPO 4: Histórico
# ------------------------------------------------------------------------------
def calcular_metricas_historico(df_raw, filtro_responsable):
    """Métricas del trimestre para el histórico (slide 7): IPN y distribución
    porcentual, más participación (universo vs respondida) y conteos absolutos."""
    df = df_raw.copy()
    df['calificacion_nps'] = df['Calificacion'].apply(clasificar_nps)

    col, val = filtro_responsable

    # Participación: universo total (TIPO Normal + responsable) vs respondidas
    df_universo = df[(df['TIPO'] == 'Normal') & (df[col] == val)].copy()
    total_universo = len(df_universo)

    df_respondida = df_universo[df_universo['Respondida'] == 'Si'].copy()
    total_participantes = len(df_respondida)
    pct_participacion = (total_participantes / total_universo) if total_universo > 0 else 0

    # NPS y casos: solo calificaciones válidas dentro de las respondidas
    df_validos = df_respondida[df_respondida['calificacion_nps'].isin(CATS_VALIDAS)]

    dist = df_validos['calificacion_nps'].value_counts(normalize=True)
    promotores = dist.get('Promotor', 0)
    pasivos = dist.get('Pasivo', 0)
    detractores = dist.get('Detractor', 0)

    counts = df_validos['calificacion_nps'].value_counts()

    return {
        'IPN': promotores - detractores,
        'Promotores': promotores,
        'Pasivos': pasivos,
        'Detractores': detractores,
        'Pct_Participacion': pct_participacion,
        'Participantes': total_participantes,
        'Num_Promotores': counts.get('Promotor', 0),
        'Num_Pasivos': counts.get('Pasivo', 0),
        'Num_Detractores': counts.get('Detractor', 0),
    }


def actualizar_base_historica(df_hist, nuevas_metricas, nuevo_periodo):
    """Si el periodo ya existe lo actualiza; si no, lo agrega al final.
    Funciona tanto para el histórico global como para el de agrupadores."""
    nuevas_metricas = dict(nuevas_metricas)
    nuevas_metricas['Periodo'] = nuevo_periodo

    if nuevo_periodo in df_hist['Periodo'].values:
        idx = df_hist[df_hist['Periodo'] == nuevo_periodo].index[0]
        for col, val in nuevas_metricas.items():
            if col != 'Periodo':
                df_hist.at[idx, col] = val
    else:
        df_hist.loc[len(df_hist)] = nuevas_metricas

    return df_hist


# ------------------------------------------------------------------------------
# TIPO 5: Histórico por agrupador (slide 12)
# ------------------------------------------------------------------------------
def calcular_ipn_por_agrupador(df_raw, agrupadores):
    """IPN del trimestre actual para cada agrupador. Devuelve {agrupador: ipn},
    con None cuando el agrupador no tiene casos (evita graficar ceros falsos)."""
    df = df_raw.copy()
    df['calificacion_nps'] = df['Calificacion'].apply(clasificar_nps)

    df = df[df['Respondida'] == 'Si'].copy()
    df_val = df[df['TIPO'] == 'Normal'].copy()

    if 'Agrupador' not in df_val.columns:
        raise KeyError("No se encontró la columna 'Agrupador' en la base.")
    df_val['Agrupador_Clean'] = df_val['Agrupador'].astype(str).str.strip().str.upper()

    resultados = {}
    for agrupador in agrupadores:
        df_grupo = df_val[df_val['Agrupador_Clean'] == agrupador]
        if len(df_grupo) > 0:
            dist = df_grupo['calificacion_nps'].value_counts(normalize=True)
            resultados[agrupador] = dist.get('Promotor', 0) - dist.get('Detractor', 0)
        else:
            resultados[agrupador] = None

    return resultados


# ------------------------------------------------------------------------------
# MENCIONES TOP 10 CON RUBRO (slides 7-9)
# ------------------------------------------------------------------------------
def calcular_top10_menciones_rubro(df_raw, categoria, filtro, top_n=10):
    """Top N de motivos (SUBNETO1..7) con su rubro (NETO 1..7) para una
    categoría NPS, dentro del segmento definido por `filtro` (col, val)."""
    df = df_raw.copy()
    df['calificacion_nps'] = df['Calificacion'].apply(clasificar_nps)

    df = df[df['calificacion_nps'].isin(CATS_VALIDAS)].copy()
    df = df[df['Respondida'] == 'Si'].copy()
    df = df[df['TIPO'] == 'Normal'].copy()
    col, val = filtro
    df = df[df[col] == val].copy()
    df_cat = df[df['calificacion_nps'] == categoria].copy()

    lista_pares = []
    for i in range(1, 8):
        col_motivo = f"SUBNETO{i}"
        col_rubro = f"NETO {i}"
        if col_motivo in df_cat.columns and col_rubro in df_cat.columns:
            temp = df_cat[[col_motivo, col_rubro]].dropna(subset=[col_motivo])
            temp.columns = ['Motivo', 'Rubro']
            lista_pares.append(temp)

    if not lista_pares:
        return pd.DataFrame(columns=['Motivo', 'Rubro', 'Menciones'])

    df_cons = pd.concat(lista_pares, ignore_index=True)
    df_cons['Motivo'] = df_cons['Motivo'].apply(lambda x: str(x).strip())
    df_cons['Rubro'] = df_cons['Rubro'].apply(lambda x: str(x).strip())

    valores_basura = ['', 'nan', '0', '0.0', ' ', 'none']
    df_cons = df_cons[~df_cons['Motivo'].str.lower().isin(valores_basura)]

    top = (df_cons.groupby(['Motivo', 'Rubro'])
           .size()
           .reset_index(name='Menciones'))
    top = top.sort_values(by='Menciones', ascending=False).head(top_n).reset_index(drop=True)
    return top


# ------------------------------------------------------------------------------
# HISTÓRICO RESEARCH LAND (slide 6): métricas simples del periodo
# ------------------------------------------------------------------------------
def calcular_metricas_historico_rl(df_raw, filtro):
    """Distribución NPS del periodo para el histórico RL (versión simple:
    calificación válida + Respondida == 'Si' + TIPO Normal + filtro)."""
    df = df_raw.copy()
    df['calificacion_nps'] = df['Calificacion'].apply(clasificar_nps)

    df = df[df['calificacion_nps'].isin(CATS_VALIDAS)].copy()
    df = df[df['Respondida'] == 'Si'].copy()
    df = df[df['TIPO'] == 'Normal'].copy()
    col, val = filtro
    df = df[df[col] == val].copy()

    dist = df['calificacion_nps'].value_counts(normalize=True)
    promotores = dist.get('Promotor', 0)
    pasivos = dist.get('Pasivo', 0)
    detractores = dist.get('Detractor', 0)

    return {
        'IPN': promotores - detractores,
        'Promotores': promotores,
        'Pasivos': pasivos,
        'Detractores': detractores,
    }
