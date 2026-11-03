import io
import re
import calendar
import unicodedata
from datetime import datetime

import pandas as pd
import streamlit as st
import gspread

from google.oauth2.service_account import Credentials


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Consolidado de Producción y Costos",
    page_icon="📊",
    layout="wide"
)

SPREADSHEET_ID = "1gNq2nuNGfogcgcq7gsIqmQ5BZmMuWiSgAbDLpytJTIY"
HOJA_DESTINO = "CONSOLIDADO_GENERAL"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

PROCESOS = {
    "Actividad-CP": "11-12-13-14 Actividad CP",
    "Lecturas": "14 Lecturas",
    "Averías": "15. Averias",
    "Notificaciones": "16. Notific",
    "Itinerante": "17 Itinerant",
}

COLUMNAS_FINALES = [
    "periodo",
    "fecha",
    "suministro",
    "cuadrilla",
    "zona",
    "codigo",
    "cantidad",
    "actividad",
    "costo_unitario",
    "costo_total",
    "proceso",
    "hoja_origen",
    "archivo_origen",
]

# ------------------------------------------------------------
# ANCHOS FIJOS DE COLUMNAS PARA LA VISTA PREVIA
# Evita que las columnas se deformen al hacer clic sobre
# una fila o celda del st.dataframe.
# ------------------------------------------------------------

CONFIG_COLUMNAS = {
    "periodo": st.column_config.Column(
        "Periodo", width="small", pinned=True
    ),
    "fecha": st.column_config.Column("Fecha", width="small"),
    "suministro": st.column_config.Column("Suministro", width="small"),
    "cuadrilla": st.column_config.Column("Cuadrilla", width="medium"),
    "zona": st.column_config.Column("Zona", width="small"),
    "codigo": st.column_config.Column("Código", width="small"),
    "cantidad": st.column_config.NumberColumn(
        "Cantidad", width="small", format="%.2f"
    ),
    "actividad": st.column_config.Column(
        "Actividad", width="large"
    ),
    "costo_unitario": st.column_config.NumberColumn(
        "Costo unitario", width="small", format="%.2f"
    ),
    "costo_total": st.column_config.NumberColumn(
        "Costo total", width="small", format="%.2f"
    ),
    "proceso": st.column_config.Column("Proceso", width="medium"),
    "hoja_origen": st.column_config.Column(
        "Hoja origen", width="medium"
    ),
    "archivo_origen": st.column_config.Column(
        "Archivo origen", width="medium"
    ),
}


# ============================================================
# FUNCIONES GENERALES
# ============================================================

def normalizar_texto(valor):
    """Convierte textos a mayúsculas y elimina tildes."""
    if pd.isna(valor):
        return ""

    texto = str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(
        caracter for caracter in texto
        if not unicodedata.combining(caracter)
    )


def limpiar_codigo(valor):
    """Evita códigos con .0 al leerlos desde Excel."""
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    if re.fullmatch(r"\d+\.0", texto):
        texto = texto[:-2]

    return texto
def es_codigo_actividad(valor):
    """
    Verifica que el valor tenga formato de código
    de actividad, por ejemplo:
    114AB40, 112MP01, 113CP24.
    """
    codigo = limpiar_codigo(valor).upper()

    return bool(
        re.fullmatch(r"\d{3}[A-Z]{2}\d{2}", codigo)
    )


def normalizar_zona(valor):
    zona = normalizar_texto(valor)

    equivalencias = {
        "UB": "UB",
        "UR": "UR",
        "R": "R",
        "URBANO": "UB",
        "URBANO RURAL": "UR",
        "RURAL": "R",
        "RL": "R",
    }

    return equivalencias.get(zona, "")


def convertir_numero(valor):
    """Convierte números y textos numéricos a formato decimal."""
    if pd.isna(valor):
        return None

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip().replace(" ", "")

    if not texto or texto.upper() in {"NAN", "NONE", "-"}:
        return None

    # Formato peruano: 12.286,46
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")

    elif "," in texto:
        texto = texto.replace(",", ".")

    try:
        return float(texto)
    except (ValueError, TypeError):
        return None


def convertir_fecha(valor):
    """Normaliza fechas, incluyendo periodos expresados entre dos fechas."""
    if pd.isna(valor):
        return pd.NaT

    if isinstance(valor, (pd.Timestamp, datetime)):
        return pd.Timestamp(valor).normalize()

    texto = str(valor).strip()

    # Ejemplo: 01/12/2025 al 31/12/2025.
    if re.search(r"\bal\b", texto, flags=re.IGNORECASE):
        coincidencias = re.findall(
            r"\d{1,2}[/-]\d{1,2}[/-]\d{4}",
            texto
        )

        if coincidencias:
            fecha_inicio = pd.to_datetime(
                coincidencias[0],
                dayfirst=True,
                errors="coerce"
            )

            if pd.notna(fecha_inicio):
                ultimo_dia = calendar.monthrange(
                    fecha_inicio.year,
                    fecha_inicio.month
                )[1]

                return pd.Timestamp(
                    fecha_inicio.year,
                    fecha_inicio.month,
                    ultimo_dia
                )

    return pd.to_datetime(
        texto,
        dayfirst=True,
        errors="coerce"
    )

def obtener_ultima_fila(df, col_fecha, fila_datos, proceso):
    """Determina la última fila con registros."""

    # ITINERANTE Y NOTIFICACIONES:
    # No depender de que FECHA tenga un valor.
    if proceso in ["Itinerante", "Notificaciones"]:

        datos = df.iloc[fila_datos:, :]

        filas_con_datos = datos.notna().any(axis=1)

        indices_validos = filas_con_datos[
            filas_con_datos
        ].index

        if len(indices_validos) == 0:
            raise ValueError(
                f"No se encontraron registros en {proceso}."
            )

        return int(indices_validos[-1])

    # DEMÁS PROCESOS:
    # Mantener el comportamiento original.
    fechas = df.iloc[fila_datos:, col_fecha].apply(
        convertir_fecha
    )

    indices_validos = fechas[fechas.notna()].index

    if len(indices_validos) == 0:
        raise ValueError(
            f"No se encontraron fechas válidas en {proceso}."
        )

    return int(indices_validos[-1])

def buscar_columnas_base(df, fila_encabezado, proceso):
    """Identifica las columnas de fecha, suministro y cuadrilla."""
    col_fecha = None
    col_suministro = None
    col_cuadrilla = None

    for col in range(df.shape[1]):
        encabezado = normalizar_texto(
            df.iloc[fila_encabezado, col]
        )

        # Identificar la fecha.
        if "FECHA" in encabezado:
            col_fecha = col

        # ----------------------------------------------------
        # LECTURAS: Nº DE OT está en la columna C.
        # ----------------------------------------------------
        if proceso == "Lecturas":

            # La columna C de Excel corresponde al índice 2.
            col_suministro = 2

            encabezado_sin_simbolos = (
                encabezado.replace("°", "")
                .replace("º", "")
            )

            encabezado_sin_simbolos = re.sub(
                r"\s+", " ", encabezado_sin_simbolos
            ).strip()

            if encabezado_sin_simbolos == "CUADRILLA":
                col_cuadrilla = col

        # ----------------------------------------------------
        # AVERÍAS Y NOTIFICACIONES
        # ----------------------------------------------------
        elif proceso in ["Averías", "Notificaciones"]:
            encabezado_sin_simbolos = (
                encabezado.replace("°", "")
                .replace("º", "")
            )

            if encabezado_sin_simbolos in [
                "SUMINISTRO",
                "SUMINISTRO/SED",
            ]:
                col_suministro = col

            if any(
                termino in encabezado_sin_simbolos
                for termino in [
                    "CUADRILLA",
                    "TECNICO",
                    "TECNICO RESPONSABLE",
                ]
            ):
                col_cuadrilla = col

    # Validar fecha.
    if col_fecha is None:
        raise ValueError(
            f"No se encontró la columna FECHA en {proceso}."
        )

    # Validar suministro.
    if col_suministro is None:
        raise ValueError(
            f"No se encontró SUMINISTRO/OT en {proceso}."
        )

    # Validar técnico o cuadrilla.
    if col_cuadrilla is None:
        raise ValueError(
            f"No se encontró la columna de técnico o cuadrilla en {proceso}."
        )

    return col_fecha, col_suministro, col_cuadrilla


def detectar_columnas_produccion(
    df,
    fila_zona,
    fila_codigo,
    proceso
):
    """Detecta códigos válidos asociados a su zona."""

    zonas_validas = {"UB", "UR", "R"}

    # Leer zonas de la fila 3
    zonas = df.iloc[fila_zona, :].apply(normalizar_zona)

    # Completar las celdas combinadas hacia la derecha
    zonas = zonas.replace("", pd.NA).ffill().fillna("")

    columnas = []

    for col in range(df.shape[1]):

        zona = zonas.iloc[col]
        codigo = limpiar_codigo(
            df.iloc[fila_codigo, col]
        ).upper()

        # Solo zonas reconocidas
        if zona not in zonas_validas:
            continue

        # Excluir MEDIDOR, PRECINTO, ETC.
        if not es_codigo_actividad(codigo):
            continue

        columnas.append({
            "columna": col,
            "zona": zona,
            "codigo": codigo,
        })

    if not columnas:
        raise ValueError(
            f"No se detectaron códigos válidos en {proceso}."
        )

    return columnas

# ============================================================
# PROCESAMIENTO DE CADA HOJA
# ============================================================

def procesar_hoja(
    archivo_excel,
    hoja,
    proceso,
    nombre_archivo
):
    """Extrae los registros de una hoja y los devuelve en formato común."""

    df = pd.read_excel(
        io.BytesIO(archivo_excel),
        sheet_name=hoja,
        header=None
    )

    fila_zona = 2
    fila_codigo = 3
    fila_encabezado = 4
    fila_datos = 5

    # --------------------------------------------------------
    # CONFIGURACIÓN ESPECIAL: ACTIVIDAD-CP
    # --------------------------------------------------------
    if proceso == "Actividad-CP":
        col_fecha = 1
        col_suministro = 4
        col_cuadrilla = 7

        ultima_fila = obtener_ultima_fila(
            df, col_fecha, fila_datos, proceso
        )

        columnas = detectar_columnas_produccion(
            df, fila_zona, fila_codigo, proceso
        )

    # --------------------------------------------------------
    # CONFIGURACIÓN ESPECIAL: ITINERANTE
    # Buscar los encabezados FECHA, MES y CUADRILLA
    # --------------------------------------------------------
    elif proceso == "Itinerante":

        col_fecha = None
        col_suministro = None
        col_cuadrilla = None

        for col in range(df.shape[1]):
            encabezado = normalizar_texto(
                df.iloc[fila_encabezado, col]
            )

            if encabezado == "FECHA":
                col_fecha = col

            elif encabezado == "MES":
                col_suministro = col

            elif encabezado == "CUADRILLA":
                col_cuadrilla = col

        faltantes = []

        if col_fecha is None:
            faltantes.append("FECHA")

        if col_suministro is None:
            faltantes.append("MES")

        if col_cuadrilla is None:
            faltantes.append("CUADRILLA")

        if faltantes:
            raise ValueError(
                "No se encontraron los encabezados "
                "en Itinerante: " + ", ".join(faltantes)
            )

        ultima_fila = obtener_ultima_fila(
            df, col_fecha, fila_datos, proceso
        )

        columnas = []

        for col in range(df.shape[1]):
            zona_original = df.iloc[fila_zona, col]
            codigo = df.iloc[fila_codigo, col]

            zona = normalizar_zona(zona_original)

            if zona and pd.notna(codigo):
                codigo_limpio = limpiar_codigo(codigo)
            
                if es_codigo_actividad(codigo_limpio):
                    columnas.append({
                        "columna": col,
                        "zona": zona,
                        "codigo": codigo_limpio,
                    })

        if not columnas:
            raise ValueError(
                "No se detectaron columnas de producción en Itinerante."
            )


        ultima_fila = obtener_ultima_fila(
            df, col_fecha, fila_datos, proceso
        )

        columnas = []

        for col in range(df.shape[1]):
            zona_original = df.iloc[fila_zona, col]
            codigo = df.iloc[fila_codigo, col]

            zona = normalizar_zona(zona_original)

            if zona and pd.notna(codigo):
                codigo_limpio = limpiar_codigo(codigo)
            
                if es_codigo_actividad(codigo_limpio):
                    columnas.append({
                        "columna": col,
                        "zona": zona,
                        "codigo": codigo_limpio,
                    })

        if not columnas:
            raise ValueError(
                "No se detectaron columnas de producción en Itinerante."
            )

    # --------------------------------------------------------
    # CONFIGURACIÓN: LECTURAS, AVERÍAS Y NOTIFICACIONES
    # --------------------------------------------------------
    else:
        (
            col_fecha,
            col_suministro,
            col_cuadrilla,
        ) = buscar_columnas_base(
            df,
            fila_encabezado,
            proceso
        )

        ultima_fila = obtener_ultima_fila(
            df, col_fecha, fila_datos, proceso
        )

        columnas = detectar_columnas_produccion(
            df, fila_zona, fila_codigo, proceso
        )

    # --------------------------------------------------------
    # GENERAR REGISTROS
    # --------------------------------------------------------
    registros = []

    for fila in range(fila_datos, ultima_fila + 1):

        
        fecha_original = df.iloc[fila, col_fecha]
        
        if proceso == "Itinerante":
            # Mantener el funcionamiento original.
            if pd.isna(fecha_original):
                continue
        
            fecha = str(fecha_original).strip()
        
            if not fecha:
                continue
        
        elif proceso == "Notificaciones":
            # Aceptar fechas vacías sin descartar el registro.
            if pd.isna(fecha_original):
                fecha = ""
        
            elif isinstance(fecha_original, (pd.Timestamp, datetime)):
                fecha = fecha_original.strftime("%d/%m/%Y")
        
            else:
                fecha = str(fecha_original).strip()
        
        else:
            # Mantener el comportamiento de los demás procesos.
            fecha = convertir_fecha(fecha_original)
        
            if pd.isna(fecha):
                continue


        if col_suministro is None:
            suministro = ""
        else:
            suministro = df.iloc[fila, col_suministro]

            if pd.isna(suministro):
                suministro = ""
            else:
                suministro = str(suministro).strip()

        cuadrilla = df.iloc[fila, col_cuadrilla]

        if pd.isna(cuadrilla):
            cuadrilla = ""
        else:
            cuadrilla = str(cuadrilla).strip()

        for item in columnas:
            cantidad_original = df.iloc[
                fila, item["columna"]
            ]

            cantidad = convertir_numero(cantidad_original)

            if cantidad is None or cantidad == 0:
                continue

            registros.append({
                "fecha": fecha,
                "suministro": suministro,
                "cuadrilla": cuadrilla,
                "zona": item["zona"],
                "codigo": item["codigo"],
                "cantidad": cantidad,
                "actividad": "",
                "costo_unitario": None,
                "costo_total": None,
                "proceso": proceso,
                "hoja_origen": hoja,
                "archivo_origen": nombre_archivo,
            })

    if not registros:
        return pd.DataFrame(columns=COLUMNAS_FINALES)

    return pd.DataFrame(registros, columns=COLUMNAS_FINALES)


# ============================================================
# FORMATO PASCO: CÓDIGOS EN FILA 2 Y DATOS DESDE FILA 4
# ============================================================
HOJAS_PASCO = {
    "MC-UB PAS": "UB", "MC-UR PAS": "UR",
    "SS-DEN. UB": "UB", "SS-DEN. UR": "UR",
    "NOT. UB": "UB", "NOT. UR": "UR",
    "AP-PASCO UB": "UB", "AP-PASCO UR": "UR",
    "AP-YANAHUANCA": "R", "AP-PAUCARTAMBO": "R", "AP-HUARIACA": "R",
    "CI-YANAHUANCA": "R", "CI-PAUCARTAMBO": "R", "CI-HUARIACA": "R",
}
PROCESOS_PASCO = {
    # ACTIVIDAD CP
    "MC-UB PAS": "Actividad-CP",
    "MC-UR PAS": "Actividad-CP",

    # LECTURAS
    "AP-PASCO UB": "Lecturas",
    "AP-PASCO UR": "Lecturas",
    "AP-YANAHUANCA": "Lecturas",
    "AP-PAUCARTAMBO": "Lecturas",
    "AP-HUARIACA": "Lecturas",

    # AVERÍAS
    "SS-DEN. UB": "Averias",
    "SS-DEN. UR": "Averias",

    # ITINERANTE
    "CI-YANAHUANCA": "Itinerante",
    "CI-PAUCARTAMBO": "Itinerante",
    "CI-HUARIACA": "Itinerante",

    # NOTIFICACIONES
    "NOT. UB": "Notificaciones",
    "NOT. UR": "Notificaciones",
}

def procesar_hoja_pasco(archivo_excel, hoja, nombre_archivo):
    df = pd.read_excel(io.BytesIO(archivo_excel), sheet_name=hoja, header=None)
    zona = HOJAS_PASCO[hoja.strip().upper()]
    encabezados = {normalizar_texto(v): i for i, v in enumerate(df.iloc[2]) if pd.notna(v)}
    def columna(*nombres):
        for nombre in nombres:
            if normalizar_texto(nombre) in encabezados:
                return encabezados[normalizar_texto(nombre)]
        return None

    es_ap = hoja.strip().upper().startswith("AP-")
    col_suministro = columna("EELECTRICO", "SUMINISTRO") if es_ap else columna("SUMINISTRO")
    col_cuadrilla = columna("TÉCNICO RESPONSABLE", "TECNICO RESPONSABLE")
    col_fecha = columna("FECHA DE EJECUCIÓN", "FECHA", "FECHA DE EJECUCION")
    if all(c is None for c in (col_fecha, col_suministro, col_cuadrilla)):
        raise ValueError(f"No se encontró FECHA, SUMINISTRO ni CUADRILLA en {hoja}")

    columnas = []
    for col in range(df.shape[1]):
        codigo = limpiar_codigo(df.iat[1, col]).upper()
        if es_codigo_actividad(codigo):
            columnas.append((col, codigo))
    if not columnas:
        raise ValueError(f"No se encontraron códigos de subactividad en fila 2: {hoja}")

    # Validar datos REALES de cada fila, nunca metadatos automáticos.
    col_item = columna("ITEM")
    col_orden = columna("ORDEN DE TRABAJO", "ORDTRABAJO")

    def dato_real(valor):
        if pd.isna(valor):
            return False
        texto = normalizar_texto(valor).strip()
        return texto not in {
            "", "NAN", "NONE", "NAT", "-", "0", "TOTAL",
            "SUBTOTAL", "TOTALES", "SUMA", "FECHA", "SUMINISTRO",
            "EELECTRICO", "TECNICO RESPONSABLE", "CUADRILLA"
        }

    def fecha_real(valor):
        return dato_real(valor) and pd.notna(convertir_fecha(valor))

    registros = []
    for fila in range(3, len(df)):
        suministro = df.iat[fila, col_suministro] if col_suministro is not None else None
        cuadrilla = df.iat[fila, col_cuadrilla] if col_cuadrilla is not None else None
        original_fecha = df.iat[fila, col_fecha] if col_fecha is not None else None

        # No sumar filas de totales o resúmenes.
        campos_control = [col_item, col_suministro, col_cuadrilla, col_orden]
        if any(
            c is not None and normalizar_texto(df.iat[fila, c]).strip()
            in {"TOTAL", "SUBTOTAL", "TOTALES", "SUMA"}
            for c in campos_control
        ):
            continue

        # ITEM es un identificador de registro, no un total ni celda vacía.
        # Hay filas auxiliares con técnico y subactividades marcadas con 1,
        # pero sin ITEM ni SUMINISTRO: no son trabajos ejecutados.
        if col_item is not None:
            item_original = df.iat[fila, col_item]
            item_numero = convertir_numero(item_original)
            if (item_numero is None or not item_numero.is_integer()
                    or item_numero <= 0):
                continue

        # Solo cuenta si hay un dato real en FECHA, SUMINISTRO o CUADRILLA.
        # periodo, zona y codigo son automáticos y NO cuentan como registro.
        if not (fecha_real(original_fecha) or dato_real(suministro) or dato_real(cuadrilla)):
            continue

        fecha = ""
        if fecha_real(original_fecha):
            fecha = convertir_fecha(original_fecha)
        for col, codigo in columnas:
            # PASCO: contar exclusivamente las celdas con valor 1.
            # No sumar valores 2, 3, decimales diferentes ni textos.
            valor = convertir_numero(df.iat[fila, col])
            if valor != 1:
                continue
            cantidad = 1
            registros.append({
                "fecha": fecha,
                "suministro": "" if pd.isna(suministro) else str(suministro).strip(),
                "cuadrilla": "" if pd.isna(cuadrilla) else str(cuadrilla).strip(),
                "zona": zona, "codigo": codigo, "cantidad": cantidad,
                "actividad": "", "costo_unitario": None, "costo_total": None,
                #"proceso": "Pasco - " + hoja.split("-")[0].strip(),
                "proceso": PROCESOS_PASCO[hoja.strip().upper()],
                "hoja_origen": hoja, "archivo_origen": nombre_archivo,
            })
    return pd.DataFrame(registros, columns=COLUMNAS_FINALES)


# ============================================================
# PROCESAMIENTO DE COSTOS
# ============================================================

def procesar_costos(archivo_costos, hoja_costos):
    df = pd.read_excel(
        io.BytesIO(archivo_costos),
        sheet_name=hoja_costos
    )

    # Normalizar nombres de columnas.
    mapa_columnas = {}

    for columna in df.columns:
        nombre = normalizar_texto(columna)

        if nombre == "ITEM":
            mapa_columnas[columna] = "codigo"

        elif nombre in [
            "ACTIVIDAD/SUBACTIVIDAD",
            "ACTIVIDAD / SUBACTIVIDAD",
        ]:
            mapa_columnas[columna] = "actividad"

        elif nombre in ["UB", "UR", "R", "RL"]:
            mapa_columnas[columna] = "R" if nombre == "RL" else nombre

    df = df.rename(columns=mapa_columnas)

    requeridas = ["codigo", "actividad", "UB", "UR", "R"]
    faltantes = [
        columna for columna in requeridas
        if columna not in df.columns
    ]

    if faltantes:
        raise ValueError(
            "El archivo de costos no contiene todas las columnas "
            "necesarias: " + ", ".join(faltantes)
        )

    df = df[requeridas].copy()

    df["codigo"] = df["codigo"].apply(limpiar_codigo)

    df["actividad"] = df["actividad"].fillna("").astype(str).str.strip()

    for zona in ["UB", "UR", "R"]:
        df[zona] = df[zona].apply(convertir_numero)

    df_costos = df.melt(
        id_vars=["codigo", "actividad"],
        value_vars=["UB", "UR", "R"],
        var_name="zona",
        value_name="costo_unitario"
    )

    df_costos = df_costos.dropna(
        subset=["costo_unitario"]
    )

    df_costos = df_costos[df_costos["codigo"] != ""]

    duplicados = df_costos.duplicated(
        subset=["codigo", "zona"],
        keep=False
    )

    cantidad_duplicados = int(duplicados.sum())

    if cantidad_duplicados:
        df_costos = df_costos.drop_duplicates(
            subset=["codigo", "zona"],
            keep="first"
        )

    return df_costos, cantidad_duplicados


def cruzar_costos(df_produccion, df_costos):
    if df_produccion.empty:
        return df_produccion

    df = df_produccion.copy()

    df["codigo"] = df["codigo"].apply(limpiar_codigo)
    df_costos = df_costos.copy()
    df_costos["codigo"] = df_costos["codigo"].apply(limpiar_codigo)

    df = df.drop(
        columns=["actividad", "costo_unitario", "costo_total"],
        errors="ignore"
    )

    df = df.merge(
        df_costos,
        on=["codigo", "zona"],
        how="left"
    )

    df["costo_total"] = (
        pd.to_numeric(df["cantidad"], errors="coerce")
        * pd.to_numeric(df["costo_unitario"], errors="coerce")
    )

    for columna in COLUMNAS_FINALES:
        if columna not in df.columns:
            df[columna] = ""

    return df[COLUMNAS_FINALES]


# ============================================================
# CONEXIÓN CON GOOGLE SHEETS
# ============================================================

def conectar_google_sheets():
    if "gcp_service_account" not in st.secrets:
        raise ValueError(
            "Falta configurar gcp_service_account en los Secrets "
            "de Streamlit."
        )

    credenciales = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=SCOPES
    )

    cliente = gspread.authorize(credenciales)

    return cliente.open_by_key(SPREADSHEET_ID)


def preparar_valores_sheets(df):
    """Convierte el DataFrame en valores compatibles con Google Sheets."""
    datos = df.copy()

    if "fecha" in datos.columns:

        def preparar_fecha(valor):
            # Primero comprobar si el valor está vacío o es NaT.
            if pd.isna(valor):
                return ""

            # Fechas reales: convertir a texto YYYY-MM-DD.
            if isinstance(valor, (pd.Timestamp, datetime)):
                return valor.strftime("%Y-%m-%d")

            # Periodos y listas de días de Itinerante:
            # conservarlos como texto.
            return str(valor).strip()

        datos["fecha"] = datos["fecha"].apply(preparar_fecha)

    # Convertir los demás valores NaN, NaT y None en celdas vacías.
    datos = datos.astype(object).where(
        pd.notna(datos),
        ""
    )

    # Devolver encabezados y filas para Google Sheets.
    return [
        datos.columns.tolist()
    ] + datos.values.tolist()


def exportar_google_sheets(df, modo):
    libro = conectar_google_sheets()

    try:
        hoja = libro.worksheet(HOJA_DESTINO)
    except gspread.WorksheetNotFound:
        hoja = libro.add_worksheet(
            title=HOJA_DESTINO,
            rows=max(len(df) + 100, 1000),
            cols=len(COLUMNAS_FINALES)
        )

    valores = preparar_valores_sheets(df)

    if modo == "Reemplazar toda la hoja":
        hoja.clear()

        hoja.resize(
            rows=max(len(valores), 100),
            cols=len(COLUMNAS_FINALES)
        )

        hoja.update(
            range_name="A1",
            values=valores,
            value_input_option="USER_ENTERED"
        )

        return (
            f"Se reemplazó el contenido de {HOJA_DESTINO} "
            f"con {len(df):,} registros."
        )

    # Modo agregar: crea encabezados si la hoja está vacía.
    hoja_vacia = not hoja.get_all_values()

    if hoja_vacia:
        hoja.resize(
            rows=max(len(valores) + 100, 1000),
            cols=len(COLUMNAS_FINALES)
        )

        hoja.update(
            range_name="A1",
            values=valores,
            value_input_option="USER_ENTERED"
        )

        return (
            f"Se creó {HOJA_DESTINO} y se agregaron "
            f"{len(df):,} registros."
        )

    encabezados_existentes = hoja.row_values(1)

    if encabezados_existentes != COLUMNAS_FINALES:
        raise ValueError(
            f"La hoja {HOJA_DESTINO} ya tiene encabezados distintos. "
            "Usa 'Reemplazar toda la hoja' para unificar su estructura."
        )

    if len(df) == 0:
        return "No hay registros nuevos para agregar."

    hoja.append_rows(
        valores[1:],
        value_input_option="USER_ENTERED"
    )

    return (
        f"Se agregaron {len(df):,} registros a "
        f"{HOJA_DESTINO}. Recuerda que agregar no elimina duplicados."
    )


# ============================================================
# INTERFAZ STREAMLIT
# ============================================================

st.title("📊 CONSOLIDADO DE PRODUCCIÓN Y COSTOS")

st.write(
    "Procesa las hojas de producción, cruza los costos por código "
    "y zona, y consolida los resultados en una única pestaña "
    "de Google Sheets."
)

st.info(
    "Sube el libro Excel que contiene las hojas de producción "
    "y el archivo Excel de costos."
)

st.subheader("1. Seleccionar procesos")
formato_pasco = st.checkbox("Procesar hojas PASCO (códigos en fila 2)", value=True)


procesos_seleccionados = st.multiselect(
    "¿Qué hojas deseas procesar?",
    options=list(PROCESOS.keys()),
    default=[] if True else list(PROCESOS.keys())
)

st.subheader("2. Cargar archivos")

col1, col2 = st.columns(2)

with col1:
    archivo_produccion = st.file_uploader(
        "Archivo Excel de producción",
        type=["xlsx", "xls", "xlsm"],
        key="produccion"
    )

with col2:
    archivo_costos = st.file_uploader(
        "Archivo Excel de costos",
        type=["xlsx", "xls", "xlsm"],
        key="costos"
    )

hoja_costos = None

if archivo_costos:
    try:
        excel_costos = pd.ExcelFile(
            io.BytesIO(archivo_costos.getvalue())
        )

        hoja_costos = st.selectbox(
            "Seleccionar hoja de costos",
            options=excel_costos.sheet_names
        )

    except Exception as e:
        st.error(f"No se pudo leer el archivo de costos: {e}")

st.subheader("3. Enviar los datos a Google Sheets")
periodo_exportacion = st.text_input(
    "Periodo de exportación (Año y mes)",
    value="202609",
    max_chars=6,
    placeholder="Ejemplo: 202609",
    help="Ingresa el periodo en formato AAAAMM. Ejemplo: 202609 para septiembre de 2026."
).strip()

modo_exportacion = st.radio(
    "¿Cómo deseas guardar los registros?",
    options=[
        "Reemplazar toda la hoja",
        "Agregar al final"
    ],
    help=(
        "Reemplazar borra el contenido anterior de la pestaña destino. "
        "Agregar conserva lo existente, pero puede duplicar registros "
        "si procesas los mismos archivos más de una vez."
    )
)

if st.button(
    "🚀 PROCESAR Y CONSOLIDAR",
    type="primary",
    use_container_width=True
):
    if not re.fullmatch(r"\d{6}", periodo_exportacion) or not (
        1 <= int(periodo_exportacion[4:6]) <= 12
    ):
        st.error(
            "Ingresa un periodo válido en formato AAAAMM, "
            "por ejemplo: 202609 o 202610."
        )

    elif not procesos_seleccionados and not formato_pasco:
        st.error("Selecciona por lo menos un proceso.")

    elif not archivo_produccion:
        st.error("Carga el archivo Excel de producción.")

    elif not archivo_costos or not hoja_costos:
        st.error("Carga el archivo Excel de costos y selecciona su hoja.")

    else:
        try:
            bytes_produccion = archivo_produccion.getvalue()
            bytes_costos = archivo_costos.getvalue()

            nombre_archivo = archivo_produccion.name

            excel_produccion = pd.ExcelFile(
                io.BytesIO(bytes_produccion)
            )

            hojas_disponibles = {
                nombre.strip(): nombre
                for nombre in excel_produccion.sheet_names
            }

            resultados = []
            errores = []
            progreso = st.progress(0)

            hojas_pasco_seleccionadas = [h for h in HOJAS_PASCO if h in {x.strip().upper() for x in excel_produccion.sheet_names}] if formato_pasco else []
            total_procesos = len(hojas_pasco_seleccionadas) if formato_pasco else len(procesos_seleccionados)

            if total_procesos == 0:
                st.warning("No se encontraron las hojas PASCO esperadas en el archivo.")

            for indice, proceso in enumerate(
                (hojas_pasco_seleccionadas if formato_pasco else procesos_seleccionados),
                start=1
            ):
                hoja_origen = proceso if formato_pasco else PROCESOS[proceso]

                if hoja_origen.strip() not in hojas_disponibles:
                    errores.append(
                        f"{proceso}: no existe la hoja "
                        f"'{hoja_origen}' en el archivo de producción."
                    )
                    progreso.progress(indice / total_procesos)
                    continue
                
                # Usar el nombre real de la hoja, conservando sus espacios originales.
                hoja_origen = hojas_disponibles[hoja_origen.strip()]

                try:
                    df_proceso = (procesar_hoja_pasco(bytes_produccion, hoja_origen, nombre_archivo)
                                  if formato_pasco else procesar_hoja(
                        archivo_excel=bytes_produccion,
                        hoja=hoja_origen,
                        proceso=proceso,
                        nombre_archivo=nombre_archivo
                    ))

                    resultados.append(df_proceso)

                except Exception as e:
                    errores.append(f"{proceso}: {str(e)}")

                progreso.progress(indice / total_procesos)

            if not resultados or all(
                df.empty for df in resultados
            ):
                st.error(
                    "No se obtuvieron registros para consolidar."
                )

                for error in errores:
                    st.warning(error)

            else:
                df_produccion = pd.concat(
                    resultados,
                    ignore_index=True
                )

                df_costos, cantidad_duplicados = procesar_costos(
                    bytes_costos,
                    hoja_costos
                )

                df_final = cruzar_costos(
                    df_produccion,
                    df_costos
                )

                # Asignar el periodo a todos los registros.
                df_final["periodo"] = periodo_exportacion

                # Periodo será la primera columna.
                df_final = df_final[COLUMNAS_FINALES]

                st.session_state["df_consolidado"] = df_final

                # Guardar el periodo usado para este consolidado.
                st.session_state["periodo_consolidado"] = periodo_exportacion

                st.success(
                    f"Procesamiento completado: "
                    f"{len(df_final):,} registros."
                )

                resumen = (
                    df_final.groupby(
                        ["proceso", "hoja_origen"],
                        dropna=False
                    )
                    .size()
                    .reset_index(name="registros")
                )

                st.subheader("Resumen por proceso")
                st.dataframe(
                    resumen,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "proceso": st.column_config.Column(
                            "Proceso", width="medium"
                        ),
                        "hoja_origen": st.column_config.Column(
                            "Hoja origen", width="medium"
                        ),
                        "registros": st.column_config.Column(
                            "Registros", width="small"
                        ),
                    },
                )

                if cantidad_duplicados:
                    st.warning(
                        f"El archivo de costos tenía "
                        f"{cantidad_duplicados} filas involucradas "
                        "en códigos/zona duplicados. Se conservó "
                        "la primera coincidencia por código y zona."
                    )

                if errores:
                    st.warning(
                        "Algunos procesos no pudieron procesarse:"
                    )

                    for error in errores:
                        st.write(f"- {error}")

        except Exception as e:
            st.error(f"Error durante el procesamiento: {e}")


# ============================================================
# VISTA PREVIA Y DESCARGA
# ============================================================

if "df_consolidado" in st.session_state:
    df_mostrar = st.session_state["df_consolidado"]

    st.divider()
    st.subheader("4. Vista previa del consolidado")

    c1, c2, c3 = st.columns(3)

    c1.metric("Total de registros", f"{len(df_mostrar):,}")

    c2.metric(
        "Registros con costo",
        f"{df_mostrar['costo_unitario'].notna().sum():,}"
    )

    c3.metric(
        "Costo total",
        f"{df_mostrar['costo_total'].sum(skipna=True):,.2f}"
    )

    # --------------------------------------------------------
    # Copia para mostrar: la fecha se convierte a texto para
    # que la celda no cambie de tamaño al ser seleccionada.
    # --------------------------------------------------------
    df_vista = df_mostrar.head(500).copy()

    df_vista["fecha"] = df_vista["fecha"].apply(
        lambda v: (
            ""
            if pd.isna(v)
            else v.strftime("%d/%m/%Y")
            if isinstance(v, (pd.Timestamp, datetime))
            else str(v)
        )
    )

    st.dataframe(
        df_vista,
        use_container_width=True,
        hide_index=True,
        column_config=CONFIG_COLUMNAS,
    )

    salida_excel = io.BytesIO()

    with pd.ExcelWriter(
        salida_excel,
        engine="openpyxl"
    ) as writer:
        df_mostrar.to_excel(
            writer,
            index=False,
            sheet_name="Consolidado"
        )

    st.download_button(
        label="📥 Descargar consolidado en Excel",
        data=salida_excel.getvalue(),
        file_name="CONSOLIDADO_GENERAL.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

    if st.button(
        "☁️ EXPORTAR A GOOGLE SHEETS",
        type="primary",
        use_container_width=True
    ):
        try:
            with st.spinner(
                "Enviando registros a Google Sheets..."
            ):
                mensaje = exportar_google_sheets(
                    df_mostrar,
                    modo_exportacion
                )

            st.success(mensaje)

            st.link_button(
                "Abrir Google Sheets",
                (
                    "https://docs.google.com/spreadsheets/d/"
                    + SPREADSHEET_ID
                    + "/edit"
                ),
                use_container_width=True
            )

        except Exception as e:
            st.error(
                "No se pudo exportar a Google Sheets. "
                f"Detalle: {e}"
            )
