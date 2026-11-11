"""Alertas tributarias: retenciones vs. perfil de la empresa y autorretenedores, IVA, estado DIAN, periodo y soporte."""
import re
from datetime import date

import numpy as np
import pandas as pd

from core.config import (
    CONCEPTOS_ESTRICTOS, ETIQUETAS_TRATAMIENTO, PATRON_IMPUESTO_EN_COSTO, TARIFA_MAXIMA_IVA, TARIFA_RETEIVA,
    TIPOS_CONTRIBUYENTE, TOLERANCIA_RELATIVA, TOLERANCIA_TARIFA_RETENCION, TRATAMIENTOS_LISTADO, UMBRAL_TOLERANCIA,
)
from data import parametros_repo, tabla_retenciones_repo
from shared.cuentas import codigo_cuenta, nombre_cuenta
from shared.dataframes import a_numero, buscar_columna, contiene, sumar_columnas
from shared.texto import normalizar_ascii, normalizar_nit

COL_ALERTAS = 'Alertas'
NO_RESPONSABLE_IVA = "No responsable de IVA"
PERIODOS_IVA = {"Bimestral": (2, 3), "Cuatrimestral": (4, 1)}  # (meses por periodo, periodos siguientes permitidos)
PATRON_NOTA_CREDITO = r'cr[eé]dito'
COL_TARIFA_RET, COL_DETALLE_RET = '% RETENCION', 'DETALLE RETENCION'
# Concepto de la tabla de retenciones (gana la primera coincidencia) y fila general de cada concepto
FAMILIAS_TABLA = (
    ('agropecuario', r'agr[ií]col.*sin procesamiento'), ('combustibles', r'combustible'),
    ('honorarios', r'^honorarios'), ('arrendamientos', r'^arrendamiento'), ('servicios', r'^servicios'),
    ('compras', r'^compras'),
)
FILA_GENERAL = {'compras': r'generales', 'servicios': r'generales', 'arrendamientos': r'inmuebles',
                'honorarios': r'jur[ií]dicas|no declarantes'}


def agregar_alerta(serie, mask, texto):
    """texto: str o Series (un texto por fila)."""
    mask = mask.reindex(serie.index, fill_value=False).fillna(False).astype(bool)
    t = texto[mask].astype(str) if isinstance(texto, pd.Series) else texto
    serie.loc[mask] = np.where(serie[mask] == '', t, serie[mask] + ' / ' + t)
    return serie


def _pesos(valor):
    return f"${valor:,.0f}".replace(',', '.')


def _porcentaje(valor):
    return f"{valor * 100:.2f} %".replace('.', ',')


def diccionario_autorretenedores(df_auto):
    cols = df_auto.columns
    col_nit = next((c for c in cols if 'NIT' in str(c).upper()), None)
    col_com = next((c for c in cols if 'COMENT' in str(c).upper()), None) or \
        next((c for c in cols if 'NOMBRE' in str(c).upper()), None)
    if not (col_nit and col_com):
        return {}
    dic = {}
    for nit, comentario in zip(df_auto[col_nit], df_auto[col_com]):
        nit = normalizar_nit(nit)
        if nit:
            dic[nit] = str(comentario).strip() or 'AUTORRETENEDOR'
    return dic


def tratamiento_listado(comentario):
    """(clave, tratamiento, explicación) del comentario del listado; observación sin explicación si no coincide."""
    texto = str(comentario).upper()
    if not texto.strip():
        return '', '', ''
    return next(((k, t, e) for k, t, e in TRATAMIENTOS_LISTADO if k in texto), ('', 'observacion', ''))


def resumen_tratamientos(dict_auto):
    """Líneas para la consola: cuántos terceros del listado reciben cada tratamiento y por qué."""
    conteo = {}
    for comentario in dict_auto.values():
        clave, tratamiento, explicacion = tratamiento_listado(comentario)
        llave = (tratamiento, clave or str(comentario).upper()[:40], explicacion)
        conteo[llave] = conteo.get(llave, 0) + 1
    def cantidad(n):
        return f"{n:,}".replace(',', '.') + (" tercero" if n == 1 else " terceros")

    lineas = [f"  Listado de autorretenedores: {cantidad(len(dict_auto))}. Tratamiento según su comentario:"]
    for (tratamiento, clave, explicacion), n in sorted(conteo.items(), key=lambda x: (x[0][0], -x[1])):
        lineas.append(f"    • \"{clave}\" ({cantidad(n)}): {ETIQUETAS_TRATAMIENTO[tratamiento]}"
                      + (f" — el tercero {explicacion}." if explicacion else "."))
    return lineas


def _corto(texto, largo=60):
    texto = str(texto)
    return texto if len(texto) <= largo else texto[:largo].rstrip() + '…'


def agentes_retencion(perfil):
    """(agente de retención en la fuente, agente de retención de IVA) según el tipo de contribuyente."""
    return TIPOS_CONTRIBUYENTE.get((perfil or {}).get('tipo'), (True, False))


def tolerancia(valor):
    return np.maximum(UMBRAL_TOLERANCIA, np.abs(valor) * TOLERANCIA_RELATIVA)


def _concepto_cuenta(columna, mapa=None, defecto='compras'):
    mapa = parametros_repo.cargar_conceptos()['gasto'] if mapa is None else mapa
    return parametros_repo.concepto_de(columna, mapa, defecto)[0]


def tipo_persona(nit):
    """'PJ' si es NIT de 9 dígitos que inicia en 8 o 9; 'PN' (cédula) en otro caso."""
    return pd.Series(np.where(nit.str.fullmatch(r'[89]\d{8}', na=False), 'PJ', 'PN'), index=nit.index)


def _reglas_tabla(mapa_retencion):
    reglas = tabla_retenciones_repo.cargar()
    for r in reglas:
        familia = next((k for k, patron in FAMILIAS_TABLA if re.search(patron, r['concepto'], re.I)), '')
        r['clave'] = (familia if familia == 'agropecuario' or not r['cuenta']
                      else _concepto_cuenta(r['cuenta'], mapa_retencion, '') or familia)
    return reglas


def tarifa_de_nombre(nombre):
    """Tarifa escrita en el nombre de la cuenta ('Compras (2.5%)' -> 0.025) o None."""
    m = re.search(r'(\d+(?:[.,]\d+)?)\s*%', str(nombre))
    return float(m.group(1).replace(',', '.')) / 100 if m else None


def _palabras(texto):
    return {p for p in re.findall(r'[a-z]{4,}', normalizar_ascii(texto))}


def _regla(reglas, tipo, cuenta='', clave='', nombre=''):
    """Regla de la tabla: por cuenta de retención con cuenta en la tabla (prefijo más largo); si no, entre las filas
    del concepto, la que coincide con la tarifa y las palabras del nombre de la cuenta de retención o, en su defecto,
    la fila general. Prima el tipo de NIT."""
    candidatas = []
    if cuenta:
        candidatas = [r for r in reglas if r['cuenta'] and cuenta.startswith(r['cuenta'])]
        largo = max((len(r['cuenta']) for r in candidatas), default=0)
        candidatas = [r for r in candidatas if len(r['cuenta']) == largo]
    if not candidatas:
        candidatas = [r for r in reglas if clave and r['clave'] == clave]
    candidatas = [r for r in candidatas if r['tipo'] in (tipo, '')] or candidatas
    if not candidatas:
        return None
    tarifa, palabras, general = tarifa_de_nombre(nombre), _palabras(nombre), FILA_GENERAL.get(clave, '')
    return max(candidatas, key=lambda r: (tarifa is not None and abs(r['tarifa'] - tarifa) < 1e-6,
                                          len(palabras & _palabras(r['concepto'])),
                                          bool(general and re.search(general, r['concepto'], re.I)),
                                          r['tipo'] == tipo))


def base_minima(uvt_concepto, uvt_anio):
    """UVT del concepto × UVT del año, aproximada al múltiplo de $1.000 más cercano (art. 868 E.T.)."""
    return float(np.floor(uvt_concepto * uvt_anio / 1000 + 0.5) * 1000)


def _detalle_cuenta(columna):
    """'236540005 - Compras (2.5%)' -> 'Compras (2.5%)'; solo código -> nombre del PUC."""
    return re.sub(r'^\s*\d+\s*[-–:.]?\s*', '', str(columna)).strip() or nombre_cuenta(columna) or str(columna)


def detalle_retencion(valores, base):
    """(tarifas, detalles) de las cuentas de retención con valor, de mayor a menor, separadas por ' / '. La tarifa
    sale del nombre de la cuenta; si no la trae, se calcula sobre la base."""
    cuentas = valores[valores > UMBRAL_TOLERANCIA].sort_values(ascending=False)
    tarifas = []
    for columna, valor in cuentas.items():
        tarifa = tarifa_de_nombre(columna)
        tarifa = tarifa if tarifa is not None else (valor / base if base else None)
        tarifas.append(_porcentaje(tarifa) if tarifa is not None else '')
    return ' / '.join(t for t in tarifas if t), ' / '.join(_detalle_cuenta(c) for c in cuentas.index)


def a_fecha(serie):
    if pd.api.types.is_datetime64_any_dtype(serie):
        return serie
    texto = serie.map(lambda v: v.isoformat() if hasattr(v, 'isoformat') else str(v)).str.strip()
    iso = texto.str.match(r'^\d{4}-\d{1,2}-\d{1,2}')
    fechas = pd.Series(pd.NaT, index=serie.index, dtype='datetime64[ns]')
    fechas[iso] = pd.to_datetime(texto[iso].str[:10], errors='coerce', format='%Y-%m-%d')
    fechas[~iso] = pd.to_datetime(texto[~iso], errors='coerce', dayfirst=True, format='mixed')
    return fechas


def _suma_abs(df, columnas):
    cols = [c for c in columnas if c in df.columns]
    return sumar_columnas(df, cols).abs() if cols else pd.Series(0.0, index=df.index)


def alertas_retencion(df_aud, seriales_base, seriales_iva, seriales_ret, seriales_reteiva, dict_auto, perfil):
    """Por fila de AUD-COMP: DataFrame con RETEFUENTE, RETEIVA y ALERTA."""
    idx = df_aud.index
    retencion, reteiva, iva = (_suma_abs(df_aud, s) for s in (seriales_ret, seriales_reteiva, seriales_iva))
    alertas = pd.Series('', index=idx)
    resultado = pd.DataFrame({'RETEFUENTE': retencion, COL_TARIFA_RET: '', COL_DETALLE_RET: '', 'RETEIVA': reteiva,
                              'ALERTA': alertas})
    cols_ret = [c for c in seriales_ret if c in df_aud.columns]
    if cols_ret and not df_aud.empty:
        valores_ret = df_aud[cols_ret].apply(pd.to_numeric, errors='coerce').fillna(0).abs()
        base_total = _suma_abs(df_aud, seriales_base)
        detalle = [detalle_retencion(valores_ret.loc[i], base_total.loc[i]) for i in idx]
        resultado[COL_TARIFA_RET] = [t for t, _ in detalle]
        resultado[COL_DETALLE_RET] = [d for _, d in detalle]

    if (perfil or {}).get('periodicidad') == NO_RESPONSABLE_IVA:
        agregar_alerta(alertas, iva > UMBRAL_TOLERANCIA, 'IVA descontable en empresa no responsable de IVA')

    cols_base = [c for c in seriales_base if c in df_aud.columns]
    if not cols_base or df_aud.empty:
        resultado['ALERTA'] = alertas
        return resultado

    valores_base = df_aud[cols_base].apply(pd.to_numeric, errors='coerce').fillna(0).abs()
    base = valores_base.sum(axis=1)
    # Base de retención: sin impuestos llevados al costo (consumo, saludables, IVA mayor valor): no son ingreso del proveedor
    cols_impuesto = [c for c in cols_base if re.search(PATRON_IMPUESTO_EN_COSTO, str(c), re.I)]
    base_ret = base - valores_base[cols_impuesto].sum(axis=1) if cols_impuesto else base
    conceptos = parametros_repo.cargar_conceptos()
    concepto_gasto = valores_base.idxmax(axis=1).map(lambda c: _concepto_cuenta(c, conceptos['gasto']))

    col_nit = buscar_columna(df_aud.columns, 'Nit/C.C.', 'NIT')
    nit = df_aud[col_nit].map(normalizar_nit) if col_nit else pd.Series('', index=idx)
    col_tipo = buscar_columna(df_aud.columns, 'Tipo', 'TIPO')
    tipo = df_aud[col_tipo].fillna('').astype(str).str.upper() if col_tipo else pd.Series('', index=idx)
    col_fecha = buscar_columna(df_aud.columns, 'Fecha', 'FECHA')
    fecha = a_fecha(df_aud[col_fecha]) if col_fecha else pd.Series(pd.NaT, index=idx)
    comentario = nit.map(dict_auto).fillna('').astype(str).str.upper()
    tratado = comentario.map(tratamiento_listado)
    tratamiento = tratado.str[1]
    en_listado = nit.isin(dict_auto.keys())

    con_base = (base > 0) & (nit != '')
    agregar_alerta(alertas, (nit != '') & (iva > UMBRAL_TOLERANCIA) & (base <= UMBRAL_TOLERANCIA), 'IVA sin base gravada')

    # Concepto: cuenta de retención usada > historial del tercero > cuenta de gasto
    concepto_ret, col_ret = pd.Series('', index=idx), pd.Series(None, index=idx, dtype=object)
    unica = pd.Series(False, index=idx)
    if cols_ret:
        unica = (valores_ret > UMBRAL_TOLERANCIA).sum(axis=1) == 1
        col_ret = valores_ret.idxmax(axis=1).where(retencion > UMBRAL_TOLERANCIA)
        concepto_ret = col_ret.map(lambda c: _concepto_cuenta(c, conceptos['retencion'], '') if isinstance(c, str) else '')
    con_concepto = concepto_ret != ''
    historial = concepto_ret[con_concepto].groupby(nit[con_concepto]).agg(lambda s: s.mode().iat[0])
    concepto = concepto_ret.where(con_concepto, nit.map(historial)).fillna(concepto_gasto)
    concepto = concepto.where(concepto != '', concepto_gasto)
    agro = en_listado & (tratamiento == 'agropecuario')
    concepto = concepto.where(~agro, 'agropecuario')

    # Tabla de retenciones: fila según la cuenta de retención usada (o la del historial del tercero) y su concepto
    reglas = _reglas_tabla(conceptos['retencion'])
    persona = tipo_persona(nit)
    nombre_ret = col_ret.where(con_concepto)
    historial_ret = nombre_ret[con_concepto].groupby(nit[con_concepto]).agg(lambda s: s.mode().iat[0])
    nombre_ret = nombre_ret.where(con_concepto, nit.map(historial_ret)).where(~agro).fillna('')
    cuenta_ref = nombre_ret.map(codigo_cuenta)
    clave = concepto.replace('mixto', 'compras')
    regla = pd.Series([_regla(reglas, p, c, k, n) for p, c, k, n in zip(persona, cuenta_ref, clave, nombre_ret)],
                      index=idx, dtype=object)
    revisado = regla.notna()
    parametros = parametros_repo.cargar()
    anio = fecha.dt.year.fillna(date.today().year).astype(int)
    base_min = pd.Series([base_minima(r['uvt'], parametros_repo.uvt_vigente(parametros, a)) if r else np.nan
                          for r, a in zip(regla, anio)], index=idx, dtype=float)
    tarifa = pd.Series([r['tarifa'] if r else np.nan for r in regla], index=idx, dtype=float)
    # Cédula: no se sabe si declara renta; también vale la tarifa de declarante (fila PJ/declarantes)
    tarifa_declarante = pd.Series([(_regla(reglas, 'PJ', c, k, n) or {}).get('tarifa', np.nan) if p == 'PN' else np.nan
                                   for p, c, k, n in zip(persona, cuenta_ref, clave, nombre_ret)], index=idx, dtype=float)
    etiqueta = regla.map(lambda r: _corto(r['concepto'], 45) if r else '')
    cifras = base_ret.map(_pesos) + pd.Series(np.where(base_ret >= base_min, ' ≥ ', ' < '), index=idx) \
        + base_min.map(lambda v: _pesos(v) if pd.notna(v) else '')
    supera = revisado & (base_ret >= base_min)
    simple = tipo.str.contains('SIMPLE', na=False) | comentario.str.contains('SIMPLE', na=False)
    agente_rf, agente_iva = agentes_retencion(perfil)

    if seriales_ret:
        con_ret = retencion > UMBRAL_TOLERANCIA
        if agente_rf:
            exime = tratamiento == 'no_sujeto'
            no_sujeto = (en_listado & exime) | simple
            observado = en_listado & (tratamiento == 'observacion') & (comentario != '')
            reducida = en_listado & (tratamiento == 'tarifa_reducida')
            sujeto = revisado & con_base & ~no_sujeto
            agregar_alerta(alertas, sujeto & ~observado & supera & ~con_ret,
                           'Sin retención: la base supera el mínimo de ' + etiqueta + ' (' + cifras + ')')
            agregar_alerta(alertas, revisado & con_base & no_sujeto & con_ret,
                           'Retención practicada a no sujeto ('
                           + comentario.map(lambda c: _corto(c, 30)).where(~(simple & ~en_listado), 'RÉGIMEN SIMPLE') + ')')
            agregar_alerta(alertas, sujeto & con_ret & (base_ret < base_min),
                           'Retención con base bajo el mínimo de ' + etiqueta + ' (' + cifras + ')')

            esperada = tarifa.where(~reducida, np.nan)

            def _desvio(t):
                return (retencion - base_ret * t).abs() > np.maximum(UMBRAL_TOLERANCIA, base_ret * t * TOLERANCIA_TARIFA_RETENCION)
            desvio = _desvio(esperada) & (tarifa_declarante.isna() | _desvio(tarifa_declarante))
            aplicada = (retencion / base_ret.replace(0, np.nan)).map(lambda v: _porcentaje(v) if pd.notna(v) else '')
            agregar_alerta(alertas, sujeto & con_ret & unica & esperada.notna() & desvio,
                           'Tarifa aplicada distinta a la tabla (' + aplicada + ' vs '
                           + esperada.map(lambda v: _porcentaje(v) if pd.notna(v) else '') + ')')

            cruce = revisado & con_ret & con_concepto & (concepto_gasto != 'mixto') & (concepto_ret != concepto_gasto) & (
                concepto_ret.isin(CONCEPTOS_ESTRICTOS) | concepto_gasto.isin(CONCEPTOS_ESTRICTOS))
            agregar_alerta(alertas, cruce, 'Cuenta de retención no corresponde al gasto (retención por '
                           + concepto_ret + ', gasto de ' + concepto_gasto + ')')
        else:
            agregar_alerta(alertas, con_ret, 'Retención practicada sin ser agente de retención')

    if seriales_reteiva:
        con_reteiva = reteiva > UMBRAL_TOLERANCIA
        if agente_iva:
            exento = simple | comentario.str.contains('GRAN CONTRIBUYENTE', na=False)
            agregar_alerta(alertas, con_base & supera & (iva > UMBRAL_TOLERANCIA) & ~con_reteiva & ~exento,
                           'Sin reteIVA: agente de IVA y proveedor con IVA')
            difiere = (reteiva - iva * TARIFA_RETEIVA).abs() > np.maximum(UMBRAL_TOLERANCIA, iva * 0.01)
            agregar_alerta(alertas, con_reteiva & difiere, 'ReteIVA distinta al 15% del IVA')
        else:
            agregar_alerta(alertas, con_reteiva, 'ReteIVA practicada sin ser agente de retención de IVA')

    resultado['ALERTA'] = alertas
    return resultado


def _indice_periodo(fechas, meses):
    return fechas.dt.year * (12 // meses) + (fechas.dt.month - 1) // meses


def _nit_empresa(df, es_dian, perfil):
    nit = normalizar_nit((perfil or {}).get('nit', ''))
    if nit:
        return nit
    recibidos = es_dian & contiene(df['Grupo'], 'recib')
    receptores = df.loc[recibidos, 'NIT Receptor'].map(normalizar_nit)
    receptores = receptores[receptores != '']
    return receptores.mode().iat[0] if not receptores.empty else ''


def alertas_conciliacion(df, prioridad_dian, prioridad_cont, mask_pairs, perfil=None):
    """Alertas sobre DIAN VS CONT. Requiere Box_Color calculado."""
    alertas = df[COL_ALERTAS].fillna('').astype(str) if COL_ALERTAS in df.columns else pd.Series('', index=df.index)
    es_dian = df['Prioridad_Fila'] == prioridad_dian
    es_cont = df['Prioridad_Fila'] == prioridad_cont
    contable = ~es_dian
    claves = df['Num.Ext_Clean']

    def por_grupo(serie, func):
        return serie[mask_pairs].groupby(claves[mask_pairs]).transform(func).reindex(df.index)

    # Estado y documento DIAN
    estado = df['Estado'].fillna('').astype(str).str.upper()
    agregar_alerta(alertas, es_dian & estado.str.contains('RECHAZ', na=False), 'Factura rechazada en DIAN')
    es_nc = es_dian & contiene(df['Tipo de documento'], PATRON_NOTA_CREDITO)
    agregar_alerta(alertas, mask_pairs & (por_grupo(es_nc, 'any') == True) & (por_grupo(es_cont, 'any') == True),  # noqa: E712
                   'Nota crédito DIAN cruzada contra compras')

    nit_empresa = _nit_empresa(df, es_dian, perfil)
    if nit_empresa:
        receptor = df['NIT Receptor'].map(normalizar_nit)
        agregar_alerta(alertas, es_dian & contiene(df['Grupo'], 'recib') & (receptor != '') & (receptor != nit_empresa),
                       'Factura a nombre de otro adquirente (art. 771-2 E.T.)')

    # Valores
    base, iva = a_numero(df['BASE']).abs(), a_numero(df['IVA']).abs()
    agregar_alerta(alertas, (base > 0) & (iva > base * TARIFA_MAXIMA_IVA + UMBRAL_TOLERANCIA),
                   'IVA supera el 19% de la base')
    total_dian = por_grupo(a_numero(df['Total']).abs().where(es_dian, 0), 'sum')
    otros_dian = por_grupo(a_numero(df['OTROS IMPUESTOS']).abs().where(es_dian, 0), 'sum')
    total_cont = por_grupo((base + a_numero(df['BASE_2']).abs() + iva).where(contable, 0), 'sum')
    # La BASE DIAN (Total - IVA) ya incluye los otros impuestos de la factura: el total cuadra si coincide con lo
    # contabilizado, o con lo contabilizado más los otros impuestos (cuando estos se registraron aparte y sin marcar)
    tol_total = tolerancia(total_dian)
    descuadre = ((total_dian - total_cont).abs() > tol_total) & ((total_dian - total_cont - otros_dian).abs() > tol_total)
    agregar_alerta(alertas, mask_pairs & (df['Box_Color'] == 'BLACK') & (total_dian > 0) & descuadre,
                   'Total DIAN no cuadra (BASE_2 u otros impuestos)')

    # Periodo de contabilización
    emision = por_grupo(a_fecha(df['Fecha Emisión']).where(es_dian), 'max')
    f_cont = a_fecha(df['Fecha'])
    validas = mask_pairs & contable & emision.notna() & f_cont.notna()
    agregar_alerta(alertas, validas & (f_cont < emision.dt.normalize()), 'Contabilizada antes de la emisión')
    agregar_alerta(alertas, validas & (f_cont.dt.year != emision.dt.year), 'Contabilizada en año distinto a la emisión')
    periodo = PERIODOS_IVA.get((perfil or {}).get('periodicidad'))
    if periodo:
        meses, permitidos = periodo
        fuera = (_indice_periodo(f_cont, meses) - _indice_periodo(emision, meses)) > permitidos
        agregar_alerta(alertas, validas & (iva > UMBRAL_TOLERANCIA) & fuera,
                       'IVA descontable fuera del término (art. 496 E.T.)')

    # Soporte electrónico
    novedad = df['Novedad'].fillna('').astype(str)
    sin_soporte = es_cont & (novedad.str.contains('CONT - Sin Pareja en DIAN', regex=False)
                             | novedad.str.contains('CONT - Sin Num.Ext', regex=False))
    agregar_alerta(alertas, sin_soporte, 'Sin factura/documento soporte electrónico en DIAN (art. 771-2 E.T.)')
    personal_cont = contable & novedad.str.startswith('CONT - Gasto Personal contabilizado')
    agregar_alerta(alertas, personal_cont, 'Gasto personal contabilizado: no deducible (art. 107 E.T.)')
    return alertas

# Texto con que empieza (o que contiene) cada alerta -> nombre en GUIA_ALERTAS
_FIRMAS_ALERTA = (
    (r'^Sin retención:', 'Sin retención'),
    (r'practicada a no sujeto', 'Retención practicada a no sujeto'),
    (r'^Retención con base bajo el mínimo', 'Retención con base bajo el mínimo'),
    (r'^Tarifa aplicada distinta', 'Tarifa aplicada distinta a la tabla de retenciones'),
    (r'^Cuenta de retención no corresponde', 'Cuenta de retención no corresponde al gasto'),
    (r'^IVA sin base gravada', 'IVA sin base gravada'),
    (r'^Factura rechazada', 'Factura rechazada en DIAN'),
    (r'^Nota crédito DIAN cruzada', 'Nota crédito DIAN cruzada contra compras'),
    (r'^Factura a nombre de otro adquirente', 'Factura a nombre de otro adquirente (art. 771-2 E.T.)'),
    (r'^IVA supera el 19', 'IVA supera el 19 % de la base'),
    (r'^Total DIAN no cuadra', 'Total DIAN no cuadra (BASE_2 u otros impuestos)'),
    (r'^Contabilizada antes', 'Contabilizada antes de la emisión'),
    (r'^Contabilizada en año distinto', 'Contabilizada en año distinto a la emisión'),
    (r'^Sin factura/documento soporte', 'Sin factura/documento soporte electrónico en DIAN (art. 771-2 E.T.)'),
    (r'^Gasto personal contabilizado', 'Gasto personal contabilizado'),
    (r'^IVA descontable en empresa no responsable', 'IVA descontable en empresa no responsable de IVA'),
    (r'^IVA descontable fuera del término', 'IVA descontable fuera del término (art. 496 E.T.)'),
    (r'^Retención practicada sin ser agente', 'Retención practicada sin ser agente de retención'),
    (r'^Sin reteIVA', 'Sin reteIVA: agente de IVA y proveedor con IVA'),
    (r'^ReteIVA distinta', 'ReteIVA distinta al 15 % del IVA'),
    (r'^ReteIVA practicada sin ser agente', 'ReteIVA practicada sin ser agente de retención de IVA'),
)


def tipo_alerta(texto):
    """Nombre de la alerta en GUIA_ALERTAS para un texto de alerta (con valores); '' si no se reconoce."""
    return next((nombre for patron, nombre in _FIRMAS_ALERTA if re.search(patron, str(texto))), '')


def explicacion_alerta(nombre):
    """(sección, cuándo aparece, qué significa) de la guía para el nombre de alerta."""
    from core.config import GUIA_ALERTAS
    for seccion, _, alertas in GUIA_ALERTAS:
        for alerta, cuando, significa in alertas:
            if alerta == nombre:
                return seccion, cuando, significa
    return '', '', ''