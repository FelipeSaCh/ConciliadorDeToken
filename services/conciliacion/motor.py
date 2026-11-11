from pathlib import Path

import pandas as pd

from core.config import PERFIL_TRIBUTARIO_ACTIVO, TOKEN_SHEETS
from core.errors import ErrorSistema, ErrorUsuario, HojaNoEncontradaError, Reportador, logger
from data import tabla_retenciones_repo
from data.autorretenedores_repo import GestorAutorretenedores
from shared.dataframes import buscar_columna, limpiar_encabezados, quitar_filas_vacias
from shared.excel_io import cargar_hoja_encabezado_variable, hojas_de
from shared.hojas import resolver_hojas
from shared.totales import quitar_fila_totales
from services.conciliacion import preparacion as prep
from services.conciliacion.alertas import alertas_retencion, diccionario_autorretenedores, tipo_alerta
from services.conciliacion.emparejamiento import unificar_dian_vs_cont
from services.conciliacion.escritura import escribir_resultados
from services.token_service import COLUMNAS_CALCULADAS, FormateadorToken


class ConciliadorAuditoria(Reportador):
    def __init__(self, file_path, sheet_names, seriales_iva=None, seriales_base=None,
                 seriales_base2=None, seriales_ret=None, seriales_iva_dc=None,
                 seriales_base_dc=None, seriales_base2_dc=None, filtro_grupo="Ambos", progress_callback=None,
                 seriales_reteiva=None, perfil=None):
        self.file_path = Path(file_path)
        self.sheet_names = dict(sheet_names)
        self.progress_callback = progress_callback
        self.seriales_iva = seriales_iva or []
        self.seriales_base = seriales_base or []
        self.seriales_base2 = seriales_base2 or []
        self.seriales_ret = seriales_ret or []
        self.seriales_reteiva = (seriales_reteiva or []) if PERFIL_TRIBUTARIO_ACTIVO else []
        self.perfil = (perfil or {}) if PERFIL_TRIBUTARIO_ACTIVO else {}
        self.seriales_iva_dc = seriales_iva_dc or []
        self.seriales_base_dc = seriales_base_dc or []
        self.seriales_base2_dc = seriales_base2_dc or []
        self.filtro_grupo = filtro_grupo

    def _validar_archivo(self):
        if not self.file_path.exists():
            raise ErrorUsuario(f"El archivo no existe: {self.file_path}")
        try:
            hojas = hojas_de(self.file_path)
        except Exception as e:
            raise ErrorSistema(f"No se pudo abrir el archivo Excel: {e}") from e
        self._reportar("Ubicando las hojas Token, Contabilidad y Terceros en el libro...")
        self.sheet_names, mensajes = resolver_hojas(self.file_path, self.sheet_names)
        for mensaje in mensajes:
            self._reportar(mensaje)
        for clave in ('principal', 'aud_comp'):
            nombre = self.sheet_names.get(clave)
            if not nombre:
                raise ErrorUsuario(f"No se especificó el nombre de la hoja '{clave}'.")
            if nombre not in hojas:
                raise HojaNoEncontradaError(nombre, hojas)
        return hojas

    def _cargar_token(self):
        try:
            df = quitar_fila_totales(pd.read_excel(self.file_path, sheet_name=self.sheet_names['principal']))
            ultimo = df.dropna(how='all').index.max()
            if pd.notna(ultimo):
                df = df.loc[:ultimo].copy()
            return limpiar_encabezados(df)
        except Exception as e:
            raise ErrorSistema(f"Error leyendo la hoja principal: {e}") from e

    @staticmethod
    def _token_requiere_formato(df):
        """Sin columnas calculadas, o con filas de factura cuyo CONCEPTO / Num.Ext / BASE quedó vacío."""
        if any(c not in df.columns for c in COLUMNAS_CALCULADAS):
            return True
        col_folio = buscar_columna(df.columns, 'FOLIO')
        filas = df[df[col_folio].notna()] if col_folio else df
        vacias = filas[['CONCEPTO', 'Num.Ext', 'BASE']].apply(
            lambda c: c.isna() | (c.astype(str).str.strip().isin(['', 'nan', 'None'])))
        return bool(vacias.any(axis=None))

    def _asegurar_token_formateado(self):
        """Formatea SIEMPRE el Token con la Contabilidad/Terceros actuales, también en re-procesos de auditorías ya
        hechas: nunca se reutilizan CONCEPTO, TERCERO, BASE ni Num.Ext de una ejecución anterior."""
        if self._token_requiere_formato(self._cargar_token()):
            self._reportar("El Token no está formateado: ejecutando 'Formatear token' automáticamente...")
        else:
            self._reportar("Reformateando el Token con los datos actuales de Contabilidad y Terceros...")
        FormateadorToken(
            self.file_path, {k: self.sheet_names.get(k) for k in TOKEN_SHEETS}, progress_callback=self.progress_callback,
            hojas_resueltas=True,
        ).ejecutar()

    def _cargar_auditoria(self, nombre):
        return prep.limpiar_hoja_auditoria(
            limpiar_encabezados(cargar_hoja_encabezado_variable(self.file_path, nombre))
        )

    def _cargar_devoluciones(self, nombre, hojas):
        if not nombre or nombre not in hojas:
            return None
        try:
            return self._cargar_auditoria(nombre)
        except Exception as e:
            logger.warning(f"No se pudo leer la hoja de auditoría de devoluciones '{nombre}': {e}")
            return None

    def _validar_retenciones(self, df_aud_comp, df_res_auditoria, df_autorretenedores):
        if not self.seriales_ret:
            self._reportar("Sin columnas de retefuente marcadas: se omite la validación de retención en la fuente.")
        if not PERFIL_TRIBUTARIO_ACTIVO:
            self._reportar("Validando retención en la fuente vs. autorretenedores...")
        else:
            if not self.seriales_reteiva:
                self._reportar("Sin columnas de reteIVA marcadas: se omite la validación de reteIVA.")
            self._reportar(f"Validando retenciones según perfil: {self.perfil.get('tipo') or 'sin definir'}...")
        reglas = tabla_retenciones_repo.cargar()
        self._reportar(f"  Tabla de retenciones: {len(reglas)} concepto(s) revisado(s)" if reglas
                       else "  Tabla de retenciones no disponible: no se revisan retenciones.")
        res = alertas_retencion(
            df_aud_comp, self.seriales_base, self.seriales_iva, self.seriales_ret, self.seriales_reteiva,
            diccionario_autorretenedores(df_autorretenedores), self.perfil,
        )
        df_res_auditoria[res.columns] = res

    def _reportar_alertas(self, alertas):
        textos = alertas[alertas != ''].str.split(' / ').explode()
        conteo = textos.map(lambda t: tipo_alerta(t) or t.split(':')[0]).value_counts()
        for texto, cantidad in conteo.items():
            self._reportar(f"  Alerta - {texto}: {cantidad}")

    def ejecutar(self):
        hojas = self._validar_archivo()
        self._asegurar_token_formateado()

        self._reportar("Cargando hoja de Token procesado...")
        df_token = prep.completar_columnas_token(self._cargar_token())
        mapas_impuestos = prep.mapas_impuestos(df_token)
        df_token = quitar_filas_vacias(
            df_token, [c for c in ('CONCEPTO', 'TERCERO', 'TIPO', 'BASE', 'Num.Ext') if c in df_token.columns]
        )

        self._reportar("Cargando hoja de auditoría de comprobantes...")
        df_aud_comp = self._cargar_auditoria(self.sheet_names['aud_comp'])
        df_autorretenedores = GestorAutorretenedores.como_dataframe()

        self._reportar("Filtrando registros y mapeando motivos de diferencias...")
        nombre_dc = self.sheet_names.get('Aud_dc')
        df_aud_dc = self._cargar_devoluciones(nombre_dc, hojas)

        df_token, personales, mask_ignorar, n_duplicados = prep.marcar_novedades_token(df_token, self.filtro_grupo)
        self._reportar(f"Se detectaron {n_duplicados} registros duplicados en Token (excluyendo Personales)")
        df_resultado = prep.construir_resultado(df_token, personales, mask_ignorar)
        df_auditoria = prep.construir_auditoria(df_token)

        self._reportar("Procesando hoja de auditoría de comprobantes...")
        df_res_auditoria = prep.procesar_seriales(
            df_aud_comp, self.seriales_iva, self.seriales_base, self.seriales_base2, mapas_impuestos
        )
        self._validar_retenciones(df_aud_comp, df_res_auditoria, df_autorretenedores)

        self._reportar("Unificando y conciliando DIAN vs Contabilidad...")
        res = unificar_dian_vs_cont(
            df_auditoria, df_aud_comp, df_res_auditoria, df_autorretenedores,
            df_aud_dc=df_aud_dc, seriales_iva_dc=self.seriales_iva_dc,
            seriales_base_dc=self.seriales_base_dc, seriales_base2_dc=self.seriales_base2_dc,
            mapa_otros_imp=mapas_impuestos, perfil=self.perfil,
        )
        self._reportar("Alertas tributarias detectadas:")
        self._reportar_alertas(res.df_completo['Alertas'])

        self._reportar("Escribiendo resultados en el archivo Excel...")
        escribir_resultados(
            self.file_path, self.sheet_names, nombre_dc, df_resultado,
            df_auditoria.drop(columns=['Novedad'], errors='ignore'), df_res_auditoria, df_autorretenedores, res,
        )

        self._reportar("Proceso completado con éxito.")
        return {
            "filas_procesadas": len(df_token),
            "filas_personales": int(personales.sum()),
            "filas_sin_pareja": int(sum(res.incompletas)),
            "df_completo": res.df_completo,
        }