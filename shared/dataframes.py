import pandas as pd

from shared.texto import VALORES_NULOS


def limpiar_encabezados(df, mayus=False):
    cols = df.columns.astype(str).str.replace(r'[\r\n\t]', '', regex=True).str.strip()
    df.columns = cols.str.upper() if mayus else cols
    return df


def a_numero(serie):
    return pd.to_numeric(serie, errors='coerce').fillna(0)


def sumar_columnas(df, columnas):
    return df[columnas].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1)


def buscar_columna(columnas, *nombres):
    objetivos = {n.strip().lower() for n in nombres}
    return next((c for c in columnas if str(c).strip().lower() in objetivos), None)


def columna_o_vacio(df, columna):
    return df[columna] if columna in df.columns else ''


def quitar_filas_vacias(df, columnas):
    """Elimina filas donde todas las columnas indicadas están vacías o con valores nulos textuales."""
    columnas = list(columnas)
    if not columnas:
        return df
    vacias = df[columnas].apply(
        lambda c: c.fillna('').astype(str).str.strip().str.lower().replace(list(VALORES_NULOS), '') == ''
    ).all(axis=1)
    return df[~vacias].reset_index(drop=True)


def contiene(serie, texto):
    return serie.astype(str).str.contains(texto, case=False, na=False)


def es_personal(df):
    """Gasto personal: CONCEPTO (TIPO + DETALLE de Contabilidad) o TIPO contienen 'PERSONAL'. Los Emitidos (ventas)
    no son gasto personal, salvo el documento soporte (compra a no obligados). Misma regla en el formateo del Token,
    la conciliación y el informe de IVA."""
    marca = pd.Series(False, index=df.index)
    for col in ('CONCEPTO', 'TIPO'):
        if col in df.columns:
            marca |= df[col].fillna('').astype(str).str.upper().str.contains('PERSONAL', regex=False)
    col_grupo, col_doc = buscar_columna(df.columns, 'Grupo'), buscar_columna(df.columns, 'Tipo de documento')
    if col_grupo:
        emitido = contiene(df[col_grupo], 'emitido')
        if col_doc:
            emitido &= ~contiene(df[col_doc], 'soporte')
        marca &= ~emitido
    return marca