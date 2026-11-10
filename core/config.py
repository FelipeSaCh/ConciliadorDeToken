from pathlib import Path

from version import __version__

APP_NAME = "Conciliador DIAN vs Contabilidad"
APP_VERSION = __version__

BASE_DIR = Path.home() / "ConciliadorDIAN"
LOG_DIR = BASE_DIR / "logs"
LOG_FILE = LOG_DIR / "conciliador.log"
ASSETS_DIR = Path(__file__).resolve().parent.parent / "assets"
ICON_FILE = ASSETS_DIR / "icon.ico"

CARPETA_AUTORRETENEDORES = Path(r"\\SRV-CONTAB\Servidor de Archivos\2025 SERVIDOR\APLICATIVO FELIPE\Documentos_auditor")
ARCHIVO_AUTORRETENEDORES = "autorretenedores.json"

DEFAULT_SHEET_NAMES = {
    "principal": "Token",
    "contabilidad": "contabilidad",
    "terceros": "TERCEROS",
    "aud_comp": "AUD-COMP",
    "autorretenedores": "AUTORRETENEDORES",
    "Aud_dc": "AUDITORIA DE DEVOLUCIONES",
}

SHEET_LABELS = {
    "principal": "Hoja Principal (Token)",
    "contabilidad": "Hoja Contabilidad",
    "terceros": "Hoja Terceros",
    "aud_comp": "Hoja Auditoría Comprobantes (AUD-COMP)",
    "Aud_dc": "Hoja Auditoría de Devoluciones (opcional)",
}

SHEET_ORDER = ["principal", "contabilidad", "terceros", "aud_comp", "Aud_dc"]
TOKEN_SHEETS = ("principal", "contabilidad", "terceros")
REQUIRED_SHEETS = ["principal", "contabilidad", "terceros", "aud_comp"]
OPTIONAL_SHEETS = ["contabilidad", "terceros", "Aud_dc"]

ORDEN = ['Tipo de documento', 'CUFE/CUDE', 'Folio', 'Prefijo', 'Num.Ext', 'Divisa', 'Forma de Pago',
         'Medio de Pago', 'Fecha Emisión', 'Fecha Recepción', 'NIT Emisor', 'Nombre Emisor',
         'NIT Receptor', 'Nombre Receptor', 'BASE', 'IVA', 'BASE_2', 'Total', 'Estado', 'Grupo']

COLUMNAS_DIAN_VS_CONT = [
    'Tipo de documento', 'CUFE/CUDE', 'Folio', 'Prefijo', 'Num.Ext', 'Divisa', 'Fecha',
    'Forma de Pago', 'Medio de Pago', 'Fecha Emisión', 'Fecha Recepción', 'NIT Emisor',
    'Nombre Emisor', 'NIT Receptor', 'Nombre Receptor', 'BASE', 'IVA', 'BASE_2', 'Total', 'Estado', 'Grupo'
]

COLUMNAS_EXCLUIDAS_AUD_COMP = ['Tipo', 'Número', 'Num.Ext', 'Fecha', 'Nit/C.C.', 'Tercero', 'Detalle', 'Sumas', 'IVA', 'BASE', "TIPO", "NUMERO", "CHEQ.NO",
                               "NO.EXT", "AÑO", "MES", "DIA", "NIT", "TERCERO", "SUC.PTO", "DOCRELA", "CCOSTO", "SCCOSTO", "DETALLE", "Unnamed: 17",
                               "ELABORO", "REF1", "REF2", "REF3", "REF4", "No. AUTORIZACION(DATAFONO)",
                               # Columnas que inyecta la conciliación (re-procesos): nunca son seriales
                               "AUTORRETENCION", "BASE_2", "OTROS IMPUESTOS", "RETENCIONES", "RETEFUENTE", "RETEIVA",
                               "ALERTA"]

# Tributos que suman al total de la factura vs. retenciones (informativas, no suman)
COLUMNAS_TRIBUTOS = [
    "ICA", "IC", "INC", "Timbre", "INC Timbre", "INC Bolsas", "IN Carbono",
    "IN Combustibles", "IC Datos", "ICL", "INPP", "IBUA", "ICUI",
]
COLUMNAS_RETENCIONES = ["Rete IVA", "Rete Renta", "Rete ICA"]
COLUMNAS_OTROS_IMPUESTOS = COLUMNAS_TRIBUTOS + COLUMNAS_RETENCIONES
IMPUESTOS_ADICIONALES = {"OTROS IMPUESTOS": COLUMNAS_TRIBUTOS, "RETENCIONES": COLUMNAS_RETENCIONES}

OUTPUT_SHEETS_TO_HIDE = ['Resultados', 'auditoria', 'resultados-auditoria']

UMBRAL_TOLERANCIA = 50.0

RED_FILL_COLOR = "FFC7CE"
DUPLICADO_FILL_COLOR = "FF6B6B"
CURRENCY_FORMAT = '"$"#,##0'

PREVIEW_MAX_ROWS = 300
PREVIEW_MAX_COLS = 60

# Interruptor del módulo de perfil tributario (tipo de contribuyente, periodicidad IVA, reteIVA).
# Desactivado: la auditoría asume que la empresa es agente de retención en la fuente y solo valida retefuente.
# Ver PROPUESTA_PERFIL_TRIBUTARIO.txt
PERFIL_TRIBUTARIO_ACTIVO = False

CATEGORY_COLORS = {
    "iva": "#EB6425",
    "base": "#16A34A",
    "base2": "#33EAEA",
    "autorretenedor": "#cc99ff",
    **({"reteiva": "#F472B6"} if PERFIL_TRIBUTARIO_ACTIVO else {}),
}
CATEGORIAS_EXCLUYENTES = tuple(c for c in ("iva", "base", "autorretenedor", "reteiva") if c in CATEGORY_COLORS)

PUC_FILE = ASSETS_DIR / "puc.json"  # código y nombre de cuenta

# Preselección automática de seriales por PUC: gana el prefijo más largo entre "incluir" y "excluir".
REGLAS_CATEGORIAS_PUC = {
    "iva": {
        "incluir": ("240811", "240813", "240814", "240815", "240816", "135530",  # IVA descontable
                    "240802", "240804", "240805"),  # IVA en devoluciones en compras/servicios/gastos
        "excluir": (),
    },
    "base": {
        "incluir": ("14", "15", "16", "1705", "1710", "51", "52", "530505", "530515", "61", "62", "71", "73", "74",
                    "510551", "520551"),  # dotación a trabajadores: compra a terceros (prima sobre la exclusión de nómina)
        "excluir": (
            "1410", "1430", "1499",                                   # producción interna / provisiones
            "1592", "1596", "1597", "1598", "1599", "1698", "1699",   # depreciaciones, amortizaciones, provisiones
            "171076",                                                  # impuesto diferido
            "5105", "5205",                                            # gastos de personal (nómina)
            "5115", "5215",                                            # impuestos (incluye IVA llevado al gasto)
            "5125", "5225",                                            # contribuciones y afiliaciones
            "5160", "5165", "5260", "5265",                            # depreciaciones y amortizaciones
            "5196", "5296", "5199", "5299", "5270", "5275",            # seguridad social, provisiones, otros
            "7305",                                                    # sueldos (CIF)
        ),
    },
    "autorretenedor": {
        "incluir": ("2365",),
        "excluir": ("236505", "236510", "236545", "236555", "236565", "236575", "236590", "236599"),
        # laborales, dividendos, loterías, ingresos del exterior, timbre, autorretenciones, CREE, pagos
    },
    "reteiva": {"incluir": ("2367",), "excluir": ()},
}
JERARQUIA_SERIALES = ["2408", "5315", "15", "62", "51", "52", "53", "72", "2365", "2367", "2368", "2205", "2335"]

CARACTERES_ESPECIALES = [
    ' ', '-', '_', '.', ',', ';', ':', '/', '\\', '|',
    '!', '"', '#', '$', '%', '&', '(', ')', '=', '?',
    '¡', '¿', '+', '*', '~', '{', '}', '[', ']', '^',
    '`', '´', "'", '<', '>', '@'
]

# Retención en la fuente (renta): UVT por año y bases mínimas en UVT por vigencia
UVT_POR_ANIO = {2024: 47065, 2025: 49799, 2026: 52374}
BASES_MINIMAS_RETENCION_UVT = (  # (vigente desde, {concepto: UVT})
    ("2000-01-01", {"compras": 27, "servicios": 4, "arrendamientos": 27, "honorarios": 0, "combustibles": 0,
                    "agropecuario": 92}),
    ("2025-06-01", {"compras": 10, "servicios": 2, "arrendamientos": 10, "honorarios": 0, "combustibles": 0,
                    "agropecuario": 70}),  # Dcto. 572/2025
    ("2026-05-08", {"compras": 27, "servicios": 4, "arrendamientos": 27, "honorarios": 0, "combustibles": 0,
                    "agropecuario": 92}),  # Suspensión C.E.
    ("2026-07-01", {"compras": 10, "servicios": 2, "arrendamientos": 10, "honorarios": 0, "combustibles": 0,
                    "agropecuario": 70}),  # Revocatoria suspensión (Auto C.E. 02-jun-2026)
)
CONCEPTO_RETENCION_POR_CUENTA = {  # prefijo de la cuenta de gasto/costo -> concepto; resto = compras
    "5110": "honorarios", "5210": "honorarios", "7310": "honorarios",
    "5120": "arrendamientos", "5220": "arrendamientos", "7320": "arrendamientos",
    "519535": "combustibles", "529535": "combustibles",
    "170510": "honorarios", "170525": "arrendamientos", "170540": "servicios", "530515": "servicios",
    "5135": "servicios", "5235": "servicios", "7335": "servicios",
    "519560": "servicios", "529560": "servicios",  # casino y restaurante
    # Mixtas (repuestos/materiales o mano de obra): "sin retención" con base de compras, "bajo base" con la de servicios
    "5145": "mixto", "5245": "mixto", "5150": "mixto", "5250": "mixto", "7345": "mixto",
}
CONCEPTO_POR_CUENTA_RETENCION = {  # subcuenta 2365 usada -> concepto (prima sobre la cuenta de gasto)
    "236515": "honorarios", "236520": "honorarios", "236525": "servicios", "236530": "arrendamientos",
    "236540": "compras", "236540004": "combustibles",
}
# Comentario del listado de autorretenedores -> tratamiento en las alertas de retención (gana la primera coincidencia).
#   no_sujeto: no se exige retención | agropecuario: base mínima y tarifa de productos agropecuarios
#   tarifa_reducida: no se compara la tarifa de la cuenta | observacion: solo informa (revisión manual)
TRATAMIENTOS_LISTADO = (
    ("AUTORRET", "no_sujeto", "es autorretenedor: él mismo se practica la retención en la fuente"),
    ("SIMPLE", "no_sujeto", "pertenece al régimen SIMPLE: no se le practica retención en la fuente (art. 911 E.T.)"),
    ("PATRIMONIO", "no_sujeto", "es un patrimonio autónomo: la retención la maneja la sociedad fiduciaria"),
    ("FIDEICOMISO", "no_sujeto", "es un fideicomiso (patrimonio autónomo): la retención la maneja la sociedad fiduciaria"),
    ("ABSTEN", "no_sujeto", "pidió abstenerse de practicar retención; conserve el soporte de esa solicitud"),
    ("NO SUJETO", "no_sujeto", "está marcado como no sujeto a retención"),
    ("NO CONTRIBUYENTE", "no_sujeto", "está marcado como no contribuyente del impuesto de renta"),
    ("NO APLICA POR", "no_sujeto", "está marcado como 'no aplica' por la norma citada en el listado; conserve el soporte"),
    ("AGRICOLA", "agropecuario", "vende productos agropecuarios sin procesamiento industrial: tienen base mínima y "
                                 "tarifa propias (1,5 %)"),
    ("ZESE", "tarifa_reducida", "pertenece al régimen especial ZESE: según el año de permanencia la retención puede ser "
                                "0 % o el 50 % de la tarifa general"),
    ("ZOMAC", "tarifa_reducida", "es sociedad ZOMAC: su tarifa de renta es reducida y la retención se aplica en esa proporción"),
    ("50%", "tarifa_reducida", "pidió que se le aplique el 50 % de la tarifa de retención"),
    ("LEY 1819", "observacion", "se acoge a la Ley 1819 de 2016: si es persona natural que no contrata 2 o más trabajadores, la retención se calcula con la "
                                "tabla del art. 383 E.T.; eso no significa que no se le retenga"),
    ("BOLSA", "observacion", "puede vender por bolsa de productos: si la compra se hizo por bolsa, la retención la "
                             "practica la bolsa, no el comprador"),
    ("RGB", "observacion", "tiene clasificación RGB en el listado: revise el soporte para definir la retención"),
    ("SIN ANIMO", "observacion", "es entidad sin ánimo de lucro: si pertenece al régimen tributario especial puede no estar "
                                 "sujeta a retención; verifique su RUT"),
    ("REVISAR FACTURA", "observacion", "tiene la indicación de revisar la factura para definir el tratamiento"),
)
ETIQUETAS_TRATAMIENTO = {
    "no_sujeto": "No se exige retención",
    "agropecuario": "Se evalúa como compra de productos agropecuarios (base mínima y tarifa 1,5 % propias)",
    "tarifa_reducida": "Se exige retención, pero no se compara la tarifa con la de la cuenta",
    "observacion": "Solo informa: no se alerta \"Sin retención\"; revisión manual",
}
COMENTARIOS_NO_SUJETOS = tuple(k for k, t, _ in TRATAMIENTOS_LISTADO if t == "no_sujeto")
TARIFA_RETENCION_AGROPECUARIO = 0.015
# Cuentas de costo/gasto/activo que registran impuestos (consumo, saludables, IVA llevado al costo): no son base de retención
PATRON_IMPUESTO_EN_COSTO = r"\bIVA\b|IMPOCONSUMO|IMPUESTO\s*(NACIONAL\s*)?(AL\s*)?CONSUMO|SALUDABLE|ULTRAPROCESAD|AZUCARAD|\bINC\b|\bICUI\b|\bIBUA\b|BOLSAS?\s*PL"
CONCEPTOS_ESTRICTOS = ("honorarios", "arrendamientos", "combustibles")  # su cruce con otra cuenta de retención se alerta
TOLERANCIA_TARIFA_RETENCION = 0.15  # desviación relativa admitida frente a la tarifa del nombre de la cuenta
TARIFA_MAXIMA_RETENCION = 0.11
TARIFA_RETEIVA = 0.15
TOLERANCIA_RELATIVA = 0.001  # 0,1 % del valor de la factura (se usa el mayor entre esto y UMBRAL_TOLERANCIA)
TARIFA_MAXIMA_IVA = 0.19
ALERTA_FILL_COLOR = "F8CBAD"

# Perfil tributario de la empresa auditada (se guarda por NIT)
ARCHIVO_PERFILES = BASE_DIR / "perfiles_empresas.json"
TIPOS_CONTRIBUYENTE = {  # tipo: (agente de retención en la fuente, agente de retención de IVA)
    "Persona jurídica (ordinaria)": (True, False),
    "Agente de retención de IVA designado por la DIAN": (True, True),
    "Gran contribuyente": (True, True),
    "Entidad estatal": (True, True),
    "Persona natural agente de retención": (True, False),
    "Persona natural no agente de retención": (False, False),
    "Régimen SIMPLE de tributación": (False, False),
}
PERIODICIDADES_IVA = ("Bimestral", "Cuatrimestral", "No responsable de IVA")

INFO_TIPOS_CONTRIBUYENTE = (
    "Define qué retenciones debe practicar la empresa auditada:\n\n"
    "• Persona jurídica (ordinaria): retiene en la fuente (renta). No practica reteIVA salvo designación.\n\n"
    "• Agente de retención de IVA designado por la DIAN: retefuente + reteIVA (15 % del IVA) "
    "a responsables de IVA (art. 437-2 E.T.).\n\n"
    "• Gran contribuyente: retefuente + reteIVA. No retiene IVA a otros grandes contribuyentes.\n\n"
    "• Entidad estatal: retefuente + reteIVA.\n\n"
    "• Persona natural agente de retención: comerciante con patrimonio o ingresos del año anterior "
    "superiores a 30.000 UVT (art. 368-2 E.T.). Retiene en la fuente.\n\n"
    "• Persona natural no agente de retención: por debajo de ese tope. No practica retenciones.\n\n"
    "• Régimen SIMPLE: no practica retención en la fuente, salvo pagos laborales (art. 911 E.T.).\n\n"
    "Los proveedores registrados en Autorretenedores no se les exige retención en la fuente."
)
INFO_PERIODICIDAD_IVA = (
    "Define el término para tomar el IVA descontable (art. 496 E.T.):\n\n"
    "• Bimestral: periodo de causación o uno de los tres bimestres siguientes.\n\n"
    "• Cuatrimestral: periodo de causación o el cuatrimestre siguiente.\n\n"
    "• No responsable de IVA: el IVA pagado es mayor valor del costo o gasto; "
    "si aparece en cuentas de IVA descontable se genera alerta."
)

# Parámetros tributarios editables desde la app (UVT y bases mínimas). Los valores de arriba son los de fábrica.
ARCHIVO_PARAMETROS = CARPETA_AUTORRETENEDORES / "parametros_tributarios.json"
ARCHIVO_PARAMETROS_LOCAL = BASE_DIR / "parametros_tributarios.json"
# Tabla de retenciones: solo se revisan los conceptos con cuenta (col. F); tarifa y base (UVT) salen de la tabla
ARCHIVO_TABLA_RETENCIONES = CARPETA_AUTORRETENEDORES / "TABLA_DE_RETENCIONES.xlsx"
ARCHIVO_TABLA_RETENCIONES_LOCAL = BASE_DIR / "TABLA_DE_RETENCIONES.xlsx"
CONCEPTOS_BASE_MINIMA = ("compras", "servicios", "arrendamientos", "honorarios", "combustibles", "agropecuario")
INFO_PARAMETROS_UVT = (
    "UVT por año: valor fijado cada año por resolución de la DIAN.\n\n"
    "Bases mínimas por vigencia: número de UVT a partir del cual se practica retención en la fuente. "
    "Cada fila rige desde su fecha hasta la siguiente vigencia.\n\n"
    "Para cada comprobante se toma su fecha, se busca la UVT de ese año y la vigencia aplicable, y "
    "base mínima = UVT del año × UVT del concepto.\n\n"
    "Concepto: se toma de la cuenta de retención usada (2365xx); si no hay, del concepto que tuvo el "
    "mismo tercero en el periodo; si no, de la cuenta de gasto. 0 UVT = sin base mínima.\n\n"
    "Agropecuario: se aplica a los terceros marcados \"AGRICOLA\" en Autorretenedores (compra de productos "
    "agropecuarios sin procesamiento industrial, tarifa 1,5 %). Si un archivo de parámetros guardado antes no trae "
    "esta columna, se usa el valor de fábrica de la vigencia.\n\n"
    "Los cambios se guardan en la carpeta compartida y aplican a todos los usuarios."
)

# Reglas de categorías de cuentas editables desde la app (valores de fábrica: REGLAS_CATEGORIAS_PUC)
ARCHIVO_CATEGORIAS = CARPETA_AUTORRETENEDORES / "categorias_cuentas.json"
ARCHIVO_CATEGORIAS_LOCAL = BASE_DIR / "categorias_cuentas.json"
ETIQUETAS_CATEGORIAS = {"iva": "IVA", "base": "Base", "base2": "Base 2", "autorretenedor": "Retefuente",
                        "reteiva": "ReteIVA"}
INFO_CATEGORIAS = (
    "Define qué cuentas se marcan automáticamente en cada categoría de seriales.\n\n"
    "• Incluir: prefijos de cuenta que pertenecen a la categoría (ej. 2365 = toda la retención en la fuente).\n\n"
    "• Excluir: prefijos que, aun estando dentro de uno incluido, no pertenecen (ej. 236505 = retención laboral).\n\n"
    "Gana el prefijo más largo: si se incluye 51 y se excluye 5105, la 510506 queda excluida y la 513520 incluida.\n\n"
    "Una cuenta no puede quedar en IVA y Base a la vez: si se marca en una, se desmarca la otra.\n\n"
    "Los cambios se guardan en la carpeta compartida y aplican a todos los usuarios al recargar las hojas."
)

# Cuentas por concepto de base mínima (editable desde la app; valores de fábrica arriba)
ARCHIVO_CONCEPTOS = CARPETA_AUTORRETENEDORES / "conceptos_retencion.json"
ARCHIVO_CONCEPTOS_LOCAL = BASE_DIR / "conceptos_retencion.json"
ETIQUETAS_CONCEPTOS = {
    "compras": "Compras (por defecto)", "servicios": "Servicios", "arrendamientos": "Arrendamientos",
    "honorarios": "Honorarios", "combustibles": "Combustibles", "mixto": "Mixto (compra o servicio)",
}
INFO_CONCEPTOS = (
    "Define qué columna de bases mínimas (UVT) se aplica a cada comprobante.\n\n"
    "• Cuentas de retención (2365xx): si el comprobante tiene retención, su cuenta decide el concepto. "
    "Es el criterio más confiable porque refleja lo que aplicó el contador.\n\n"
    "• Cuentas de gasto/costo: se usan cuando el comprobante no tiene retención ni historial del tercero. "
    "Decide la cuenta de base de mayor valor.\n\n"
    "• Compras es el concepto por defecto: toda cuenta de gasto sin asignar se trata como compra.\n\n"
    "• Mixto: cuentas que pueden ser compra o servicio (mantenimiento, reparaciones). Para 'sin retención' se "
    "usa la base de compras y para 'bajo base mínima' la de servicios.\n\n"
    "Gana el prefijo más largo (ej. 236540 = compras, 236540004 = combustibles)."
)
# Guía de alertas (botón "?" en Conciliar y capítulo 11 del manual): (sección, nota, [(alerta, cuándo aparece, qué significa)])
GUIA_ALERTAS = (
    ("Retención en la fuente (cada comprobante contable)",
     "Solo se revisan los conceptos de TABLA_DE_RETENCIONES.xlsx que tienen cuenta asignada; base mínima (UVT) y tarifa "
     "salen de la tabla según el tipo de NIT (PJ: NIT 8/9 de 9 dígitos; PN: cédula). La base es la suma de las cuentas "
     "marcadas como Base, sin impuestos llevados al costo.",
     (
         ("Sin retención",
          "La compra supera la base mínima del concepto de la tabla y no se registró retención.",
          "Retención posiblemente omitida: la empresa responde por ella (art. 370 E.T.). Verifique si se retuvo al pagar."),
         ("Retención practicada a no sujeto",
          "Se retuvo a un proveedor no sujeto según el listado de Autorretenedores o el régimen SIMPLE.",
          "Retención indebida: el proveedor puede pedir el reintegro."),
         ("Retención con base bajo el mínimo",
          "Se retuvo en una compra que no alcanza la base mínima del concepto de la tabla.",
          "Retención sin causa."),
         ("Tarifa aplicada distinta a la tabla de retenciones",
          "Lo retenido se aparta más del 15 % de la tarifa de la tabla.",
          "Tarifa mal aplicada o base que incluye valores que no son base."),
         ("Cuenta de retención no corresponde al gasto",
          "La retención se registró con un concepto distinto al del gasto y uno de los dos es honorarios, arrendamientos o "
          "combustibles (p. ej. arriendo con retención de compras).",
          "Concepto de retención mal aplicado: tarifa y base mínima pueden ser otras."),
     )),
    ("IVA (cada comprobante contable)", "",
     (
         ("IVA sin base gravada",
          "El comprobante tiene IVA descontable pero ninguna cuenta marcada como Base tiene valor.",
          "Falta marcar la cuenta del gasto o costo, o el IVA está mal contabilizado."),
     )),
    ("Factura DIAN vs contabilidad (hoja DIAN VS CONT)", "",
     (
         ("Factura rechazada en DIAN", "El estado de la factura en el Token dice \"Rechazada\".",
          "La factura no sirve como soporte de costo ni de IVA."),
         ("Nota crédito DIAN cruzada contra compras",
          "Una nota crédito de la DIAN quedó emparejada con un comprobante de compra.",
          "Debió registrarse como devolución."),
         ("Factura a nombre de otro adquirente (art. 771-2 E.T.)",
          "Una factura recibida está a nombre de un NIT distinto al de la empresa.",
          "Ese costo no es deducible y el IVA no es descontable."),
         ("IVA supera el 19 % de la base", "El IVA es mayor al 19 % de la base.",
          "Ninguna tarifa de IVA es mayor: error de valores o de cuentas."),
         ("Total DIAN no cuadra (BASE_2 u otros impuestos)",
          "La pareja cuadra en base e IVA, pero el total de la factura no coincide con lo contabilizado, aun descontando "
          "los otros impuestos de la factura (consumo, bolsas, etc.).",
          "Falta marcar una cuenta en Base 2 o un impuesto adicional sin contabilizar."),
         ("Contabilizada antes de la emisión", "La fecha contable es anterior a la emisión de la factura.",
          "Causación anticipada."),
         ("Contabilizada en año distinto a la emisión", "La factura es de un año y se contabilizó en otro.",
          "Afecta el año en que se toma el costo (corte de periodo)."),
         ("Sin factura/documento soporte electrónico en DIAN (art. 771-2 E.T.)",
          "Comprobante contable sin factura electrónica en la DIAN (sin pareja o sin Num.Ext).",
          "Costo e IVA sin soporte: solicitar la factura o el documento soporte."),
         ("Gasto personal contabilizado",
          "Un comprobante contable tiene el Num.Ext (y NIT) de una factura DIAN que la hoja Contabilidad marca como gasto "
          "personal (TIPO o concepto con \"PERSONAL\").",
          "No es costo ni gasto deducible (art. 107 E.T.) y su IVA no es descontable (art. 488 E.T.)."),
     )),
    ("Comentarios del listado de autorretenedores (cómo se tratan)",
     "El programa busca estas palabras en el comentario de cada tercero del listado; gana la primera que coincida. "
     "Un comentario que no coincide con ninguna se trata como observación (solo informa).",
     tuple((f"Comentario con \"{clave}\"", ETIQUETAS_TRATAMIENTO[tratamiento], f"El tercero {explicacion}.")
           for clave, tratamiento, explicacion in TRATAMIENTOS_LISTADO)),
    ("Perfil tributario (hoy desactivado: no aparecen)",
     "Dependen del tipo de contribuyente y de la periodicidad de IVA de la empresa.",
     (
         ("IVA descontable en empresa no responsable de IVA",
          "La empresa no es responsable de IVA y registra IVA descontable.", "Ese IVA es mayor valor del costo o gasto."),
         ("IVA descontable fuera del término (art. 496 E.T.)",
          "El IVA se tomó después del periodo permitido (bimestral: 3 bimestres siguientes; cuatrimestral: el siguiente).",
          "El IVA ya no se puede descontar."),
         ("Retención practicada sin ser agente de retención",
          "La empresa no está obligada a retener y registró retenciones.", "Retención indebida."),
         ("Sin reteIVA: agente de IVA y proveedor con IVA",
          "La empresa es agente de retención de IVA, el proveedor cobró IVA y no se retuvo el 15 %.", "ReteIVA omitida."),
         ("ReteIVA distinta al 15 % del IVA", "Se retuvo IVA, pero no es el 15 % del IVA de la factura.",
          "Valor de reteIVA mal calculado."),
         ("ReteIVA practicada sin ser agente de retención de IVA",
          "Se retuvo IVA sin que la empresa esté obligada.", "Retención indebida."),
     )),
)