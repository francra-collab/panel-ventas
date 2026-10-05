import streamlit as st
import pandas as pd
import requests
import io
import base64
import hashlib

# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Panel de Ventas",
    page_icon="📊",
    layout="wide",
)

HISTORIAL_PATH = "data/historial.csv"

COLUMNAS_REQUERIDAS = [
    "Fecha",
    "Cod",
    "Articulo",
    "Cantidad",
    "Cliente",
    "Precio Costo",
    "Total Costo",
    "Rentabilidad",
    "Precio Vent.",
    "Total Venta",
    "Forma Pago",
]

COLUMNAS_NUMERICAS = [
    "Cantidad",
    "Precio Costo",
    "Total Costo",
    "Rentabilidad",
    "Precio Vent.",
    "Total Venta",
]

COLUMNAS_TEXTO = [
    "Cod",
    "Articulo",
    "Cliente",
    "Forma Pago",
]

# ============================================================
# FUNCIONES GENERALES
# ============================================================

def numero(valor):
    """Convierte números argentinos y números normales a float."""
    if pd.isna(valor):
        return 0.0

    texto = str(valor).strip()

    if texto == "":
        return 0.0

    texto = texto.replace("$", "").replace(" ", "")

    try:
        if "." in texto and "," in texto:
            texto = texto.replace(".", "").replace(",", ".")
        elif "," in texto:
            texto = texto.replace(",", ".")
        return float(texto)
    except (ValueError, TypeError):
        return 0.0


def limpiar_columnas(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(axis=1, how="all")
    return df


def leer_infoventas(archivo):
    """Intenta leer XLSX/XLS y, si no funciona, TXT/CSV."""
    contenido = archivo.getvalue()
    nombre = str(getattr(archivo, "name", "")).lower()

    if nombre.endswith((".xlsx", ".xls", ".xlsm")):
        for kwargs in (
            {"engine": "openpyxl"},
            {"engine": "xlrd"},
        ):
            try:
                df = pd.read_excel(io.BytesIO(contenido), **kwargs)
                df = limpiar_columnas(df)
                if len(df.columns) > 3:
                    return df
            except Exception:
                pass

    try:
        df = pd.read_excel(io.BytesIO(contenido))
        df = limpiar_columnas(df)
        if len(df.columns) > 3:
            return df
    except Exception:
        pass

    for encoding in ("cp1252", "latin1", "utf-8-sig", "utf-8"):
        try:
            df = pd.read_csv(
                io.BytesIO(contenido),
                sep="\t",
                encoding=encoding,
            )
            df = limpiar_columnas(df)
            if len(df.columns) > 3:
                return df
        except Exception:
            pass

    for encoding in ("cp1252", "latin1", "utf-8-sig", "utf-8"):
        try:
            df = pd.read_csv(
                io.BytesIO(contenido),
                encoding=encoding,
            )
            df = limpiar_columnas(df)
            if len(df.columns) > 3:
                return df
        except Exception:
            pass

    raise ValueError(
        "No pude reconocer el archivo. "
        "Probá con el Excel original de INFOVENTAS (.xls o .xlsx)."
    )


def preparar_datos(df):
    df = limpiar_columnas(df)

    faltantes = [
        c for c in COLUMNAS_REQUERIDAS
        if c not in df.columns
    ]

    if faltantes:
        raise ValueError(
            "Faltan columnas obligatorias: "
            + ", ".join(faltantes)
        )

    df["Fecha"] = pd.to_datetime(
        df["Fecha"],
        dayfirst=True,
        errors="coerce",
    )

    if df["Fecha"].isna().sum() == len(df):
        raise ValueError(
            "No pude reconocer ninguna fecha de la columna 'Fecha'."
        )

    for columna in COLUMNAS_NUMERICAS:
        df[columna] = df[columna].apply(numero)

    for columna in COLUMNAS_TEXTO:
        df[columna] = (
            df[columna]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    for columna in [
        "TC",
        "Suc.",
        "Nº Comprobante",
        "Lista Utilizada",
        "Porcentaje Comisión",
        "Monto Comisión",
    ]:
        if columna not in df.columns:
            df[columna] = ""

    # Precio real por unidad
    df["Precio Unitario"] = 0.0
    mascara = df["Cantidad"] != 0
    df.loc[mascara, "Precio Unitario"] = (
        df.loc[mascara, "Total Venta"]
        / df.loc[mascara, "Cantidad"]
    )

    # Este cálculo es independiente de la columna
    # Rentabilidad (%) que ya viene en INFOVENTAS.
    df["Resultado Venta-Costo $"] = (
        df["Total Venta"] - df["Total Costo"]
    )

    # Agrupación solicitada: contado / cuenta corriente / otro
    forma = df["Forma Pago"].str.upper()

    df["Tipo Pago"] = "OTRO"

    df.loc[
        forma.str.contains("CONTADO", na=False),
        "Tipo Pago"
    ] = "CONTADO"

    df.loc[
        forma.str.contains("CUENTA CORRIENTE", na=False),
        "Tipo Pago"
    ] = "CUENTA CORRIENTE"

    df["Dia"] = df["Fecha"].dt.date.astype(str)

    # Número de fila original para distinguir dos renglones iguales.
    df["_FilaArchivo"] = range(1, len(df) + 1)

    # Firma de fila para evitar volver a cargar exactamente
    # el mismo renglón del mismo archivo.
    columnas_firma = [
        "Fecha",
        "Cod",
        "Articulo",
        "Cantidad",
        "Cliente",
        "Total Venta",
        "Total Costo",
        "Nº Comprobante",
        "_FilaArchivo",
    ]

    def firma_fila(fila):
        partes = []

        for columna in columnas_firma:
            valor = fila.get(columna, "")

            if isinstance(valor, pd.Timestamp):
                valor = valor.isoformat()

            partes.append(str(valor))

        texto = "|".join(partes)

        return hashlib.sha256(
            texto.encode("utf-8")
        ).hexdigest()

    df["_FirmaFila"] = df.apply(
        firma_fila,
        axis=1,
    )

    # ID compatible con el histórico anterior.
    df["ID Operacion"] = (
        df["Fecha"].astype(str)
        + "|"
        + df["Cod"]
        + "|"
        + df["Cliente"]
        + "|"
        + df["Total Venta"].round(2).astype(str)
        + "|"
        + df["Cantidad"].round(2).astype(str)
    )

    return df


# ============================================================
# GITHUB - HISTORIAL
# ============================================================

def github_configurado():
    try:
        return all(
            clave in st.secrets
            and str(st.secrets[clave]).strip() != ""
            for clave in (
                "GITHUB_TOKEN",
                "GITHUB_REPO",
                "GITHUB_USER",
            )
        )
    except Exception:
        return False


def github_headers():
    return {
        "Authorization": (
            f"Bearer {st.secrets['GITHUB_TOKEN']}"
        ),
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def github_url():
    return (
        "https://api.github.com/repos/"
        f"{st.secrets['GITHUB_USER']}/"
        f"{st.secrets['GITHUB_REPO']}/contents/"
        f"{HISTORIAL_PATH}"
    )


def cargar_historial():
    if not github_configurado():
        return pd.DataFrame()

    try:
        respuesta = requests.get(
            github_url(),
            headers=github_headers(),
            timeout=20,
        )

        if respuesta.status_code == 404:
            return pd.DataFrame()

        respuesta.raise_for_status()

        datos = respuesta.json()

        if "content" not in datos:
            return pd.DataFrame()

        contenido = base64.b64decode(
            datos["content"]
        )

        df = pd.read_csv(
            io.BytesIO(contenido)
        )

        if "Fecha" in df.columns:
            df["Fecha"] = pd.to_datetime(
                df["Fecha"],
                errors="coerce",
            )

        return df

    except Exception as e:
        st.error(
            "No se pudo leer el historial de GitHub. "
            f"Detalle: {e}"
        )
        return pd.DataFrame()


def guardar_historial(df):
    if not github_configurado():
        st.error(
            "GitHub no está configurado en los Secrets de Streamlit."
        )
        return False

    try:
        df_guardar = df.copy()

        if "Fecha" in df_guardar.columns:
            df_guardar = df_guardar.sort_values(
                "Fecha",
                ascending=True,
                na_position="last",
            )

        contenido_csv = df_guardar.to_csv(
            index=False
        ).encode("utf-8")

        contenido_base64 = base64.b64encode(
            contenido_csv
        ).decode("utf-8")

        respuesta = requests.get(
            github_url(),
            headers=github_headers(),
            timeout=20,
        )

        sha = None

        if respuesta.status_code == 200:
            sha = respuesta.json().get("sha")
        elif respuesta.status_code != 404:
            respuesta.raise_for_status()

        datos = {
            "message": "Actualizar historial de ventas",
            "content": contenido_base64,
            "branch": "main",
        }

        if sha:
            datos["sha"] = sha

        respuesta = requests.put(
            github_url(),
            headers=github_headers(),
            json=datos,
            timeout=60,
        )

        respuesta.raise_for_status()

        return True

    except Exception as e:
        st.error(
            "No se pudo guardar el historial en GitHub. "
            f"Detalle: {e}"
        )
        return False


# ============================================================
# PRESENTACIÓN
# ============================================================

def formato_pesos(valor):
    try:
        return f"$ {float(valor):,.2f}"
    except Exception:
        return "$ 0.00"


# ============================================================
# TÍTULO
# ============================================================

st.title("📊 Panel de Ventas")

st.caption(
    "INFOVENTAS — análisis diario e histórico"
)

# ============================================================
# CARGAR HISTORIAL
# ============================================================

historial = cargar_historial()

if not historial.empty:
    historial["Fecha"] = pd.to_datetime(
        historial["Fecha"],
        errors="coerce",
    )

# ============================================================
# SUBIR ARCHIVO
# ============================================================

st.subheader("📁 Cargar INFOVENTAS")

archivo = st.file_uploader(
    "Subí el INFOVENTAS del día",
    type=[
        "xlsx",
        "xls",
        "xlsm",
        "csv",
        "txt",
    ],
)

if archivo is not None:

    try:
        df_nuevo = leer_infoventas(
            archivo
        )

        df_nuevo = preparar_datos(
            df_nuevo
        )

        st.success(
            f"Archivo leído correctamente: "
            f"{len(df_nuevo):,} filas."
        )

        fechas_archivo = sorted(
            df_nuevo["Fecha"]
            .dropna()
            .dt.date
            .unique()
        )

        if fechas_archivo:
            st.info(
                "Fecha/s detectada/s: "
                + ", ".join(
                    str(f)
                    for f in fechas_archivo
                )
            )

        # ====================================================
        # GUARDADO AUTOMÁTICO
        # ====================================================

        if not github_configurado():

            st.warning(
                "⚠️ GitHub no está configurado en los Secrets. "
                "El dashboard funciona, pero todavía no "
                "guardará el histórico."
            )

        else:

            if historial.empty:

                df_para_agregar = df_nuevo.copy()

            elif "_FirmaFila" in historial.columns:

                firmas_existentes = set(
                    historial["_FirmaFila"]
                    .dropna()
                    .astype(str)
                )

                df_para_agregar = df_nuevo[
                    ~df_nuevo["_FirmaFila"].isin(
                        firmas_existentes
                    )
                ].copy()

            else:

                ids_existentes = set(
                    historial.get(
                        "ID Operacion",
                        pd.Series(dtype=str),
                    )
                    .dropna()
                    .astype(str)
                )

                df_para_agregar = df_nuevo[
                    ~df_nuevo["ID Operacion"].isin(
                        ids_existentes
                    )
                ].copy()

            if not df_para_agregar.empty:

                with st.spinner(
                    "Guardando automáticamente en el histórico..."
                ):

                    historial_actualizado = pd.concat(
                        [
                            historial,
                            df_para_agregar,
                        ],
                        ignore_index=True,
                    )

                    if "_FirmaFila" in historial_actualizado.columns:

                        historial_actualizado = (
                            historial_actualizado
                            .drop_duplicates(
                                subset=["_FirmaFila"],
                                keep="first",
                            )
                        )

                    elif "ID Operacion" in historial_actualizado.columns:

                        historial_actualizado = (
                            historial_actualizado
                            .drop_duplicates(
                                subset=["ID Operacion"],
                                keep="first",
                            )
                        )

                    if guardar_historial(
                        historial_actualizado
                    ):

                        historial = (
                            historial_actualizado
                        )

                        st.success(
                            f"✅ Guardado automático: "
                            f"{len(df_para_agregar):,} "
                            f"filas nuevas."
                        )

            else:

                st.info(
                    "ℹ️ Este archivo ya está cargado "
                    "en el histórico. No se agregaron duplicados."
                )

    except Exception as e:

        st.error(
            "❌ Error procesando el archivo: "
            f"{e}"
        )

        st.stop()

# ============================================================
# SIN HISTÓRICO
# ============================================================

if historial.empty:

    st.info(
        "Subí un INFOVENTAS para comenzar."
    )

    st.stop()

# ============================================================
# PREPARAR DATASET
# ============================================================

df = historial.copy()

df["Fecha"] = pd.to_datetime(
    df["Fecha"],
    errors="coerce",
)

for columna in [
    "Suc.",
    "Lista Utilizada",
    "Nº Comprobante",
]:
    if columna not in df.columns:
        df[columna] = ""

# Si el histórico viejo no tiene estas columnas,
# las reconstruimos para mantener compatibilidad.
if "Resultado Venta-Costo $" not in df.columns:
    df["Resultado Venta-Costo $"] = (
        df["Total Venta"] - df["Total Costo"]
    )

if "Tipo Pago" not in df.columns:
    forma = (
        df["Forma Pago"]
        .fillna("")
        .astype(str)
        .str.upper()
    )

    df["Tipo Pago"] = "OTRO"

    df.loc[
        forma.str.contains("CONTADO", na=False),
        "Tipo Pago",
    ] = "CONTADO"

    df.loc[
        forma.str.contains(
            "CUENTA CORRIENTE",
            na=False,
        ),
        "Tipo Pago",
    ] = "CUENTA CORRIENTE"

# ============================================================
# FILTROS
# ============================================================

st.sidebar.header("🔎 Filtros")

fechas_validas = df["Fecha"].dropna()

if fechas_validas.empty:

    st.error(
        "No hay fechas válidas en el histórico."
    )

    st.stop()

fecha_min = fechas_validas.min().date()
fecha_max = fechas_validas.max().date()

rango_fechas = st.sidebar.date_input(
    "Período",
    value=(
        fecha_min,
        fecha_max,
    ),
    min_value=fecha_min,
    max_value=fecha_max,
)

if (
    isinstance(rango_fechas, (tuple, list))
    and len(rango_fechas) == 2
):

    fecha_inicio = pd.Timestamp(
        rango_fechas[0]
    )

    fecha_fin = (
        pd.Timestamp(rango_fechas[1])
        + pd.Timedelta(days=1)
    )

    df = df[
        (df["Fecha"] >= fecha_inicio)
        & (df["Fecha"] < fecha_fin)
    ]

# Sucursal
sucursales = sorted(
    df["Suc."]
    .fillna("")
    .astype(str)
    .str.strip()
    .replace("", pd.NA)
    .dropna()
    .unique()
)

sucursal = st.sidebar.multiselect(
    "Sucursal",
    sucursales,
)

if sucursal:
    df = df[
        df["Suc."].astype(str).isin(
            sucursal
        )
    ]

# Cliente
clientes = sorted(
    df["Cliente"]
    .fillna("")
    .astype(str)
    .str.strip()
    .replace("", pd.NA)
    .dropna()
    .unique()
)

cliente = st.sidebar.multiselect(
    "Cliente",
    clientes,
)

if cliente:
    df = df[
        df["Cliente"].isin(
            cliente
        )
    ]

# Tipo de pago
tipos_pago = sorted(
    df["Tipo Pago"]
    .fillna("OTRO")
    .astype(str)
    .unique()
)

tipo_pago = st.sidebar.multiselect(
    "Tipo de pago",
    tipos_pago,
)

if tipo_pago:
    df = df[
        df["Tipo Pago"].isin(
            tipo_pago
        )
    ]

# Forma exacta
formas_pago = sorted(
    df["Forma Pago"]
    .fillna("")
    .astype(str)
    .str.strip()
    .replace("", pd.NA)
    .dropna()
    .unique()
)

forma_pago = st.sidebar.multiselect(
    "Forma de pago exacta",
    formas_pago,
)

if forma_pago:
    df = df[
        df["Forma Pago"].isin(
            forma_pago
        )
    ]

# Lista
listas = sorted(
    df["Lista Utilizada"]
    .fillna("")
    .astype(str)
    .str.strip()
    .replace("", pd.NA)
    .dropna()
    .unique()
)

lista = st.sidebar.multiselect(
    "Lista utilizada",
    listas,
)

if lista:
    df = df[
        df["Lista Utilizada"].isin(
            lista
        )
    ]

# Artículo
articulos = sorted(
    df["Articulo"]
    .fillna("")
    .astype(str)
    .str.strip()
    .replace("", pd.NA)
    .dropna()
    .unique()
)

articulo = st.sidebar.multiselect(
    "Artículo",
    articulos,
)

if articulo:
    df = df[
        df["Articulo"].isin(
            articulo
        )
    ]

# ============================================================
# SIN RESULTADOS
# ============================================================

if df.empty:

    st.warning(
        "⚠️ No hay operaciones que coincidan "
        "con los filtros."
    )

    st.stop()

# ============================================================
# KPIs
# ============================================================

total_venta = df["Total Venta"].sum()

unidades = df["Cantidad"].sum()

comprobantes_validos = (
    df["Nº Comprobante"]
    .fillna("")
    .astype(str)
    .str.strip()
)

comprobantes = (
    comprobantes_validos[
        comprobantes_validos != ""
    ].nunique()
)

if comprobantes == 0:
    comprobantes = len(df)

clientes_unicos = (
    df["Cliente"]
    .replace("", pd.NA)
    .dropna()
    .nunique()
)

articulos_unicos = (
    df["Cod"]
    .replace("", pd.NA)
    .dropna()
    .nunique()
)

cantidad_total = df["Cantidad"].sum()

precio_promedio = (
    total_venta / cantidad_total
    if cantidad_total != 0
    else 0
)

ticket_promedio = (
    total_venta / comprobantes
    if comprobantes != 0
    else 0
)

resultado_total = (
    df["Resultado Venta-Costo $"].sum()
)

rentabilidad_promedio = (
    df["Rentabilidad"].mean()
    if len(df) > 0
    else 0
)

# ============================================================
# RESUMEN SUPERIOR
# ============================================================

st.subheader("📌 Resumen")

c1, c2, c3, c4 = st.columns(4)

c1.metric(
    "💰 Facturación",
    formato_pesos(total_venta),
)

c2.metric(
    "📦 Unidades",
    f"{unidades:,.0f}",
)

c3.metric(
    "🧾 Comprobantes",
    f"{comprobantes:,}",
)

c4.metric(
    "👥 Clientes",
    f"{clientes_unicos:,}",
)

c5, c6, c7, c8 = st.columns(4)

c5.metric(
    "💵 Precio promedio",
    formato_pesos(precio_promedio),
)

c6.metric(
    "🛒 Ticket promedio",
    formato_pesos(ticket_promedio),
)

c7.metric(
    "📈 Rentabilidad % promedio",
    f"{rentabilidad_promedio:.2f}%",
)

c8.metric(
    "💰 Venta - Costo",
    formato_pesos(resultado_total),
)

st.caption(
    "Rentabilidad % = columna original de INFOVENTAS. "
    "Venta - Costo = diferencia matemática entre ambas columnas."
)

st.divider()

# ============================================================
# TABS
# ============================================================

tab_resumen, tab_articulos, tab_clientes, tab_pagos, tab_evolucion, tab_detalle = st.tabs(
    [
        "📊 Resumen",
        "🔧 Artículos",
        "👥 Compradores",
        "💳 Pagos",
        "📅 Evolución",
        "📋 Detalle",
    ]
)

# ============================================================
# RESUMEN
# ============================================================

with tab_resumen:

    st.subheader(
        "📊 Principales indicadores"
    )

    resumen = pd.DataFrame(
        {
            "Indicador": [
                "Facturación",
                "Unidades",
                "Comprobantes",
                "Clientes",
                "Artículos",
                "Precio promedio",
                "Ticket promedio",
                "Rentabilidad % promedio",
                "Venta - Costo",
            ],
            "Valor": [
                formato_pesos(total_venta),
                f"{unidades:,.0f}",
                f"{comprobantes:,}",
                f"{clientes_unicos:,}",
                f"{articulos_unicos:,}",
                formato_pesos(precio_promedio),
                formato_pesos(ticket_promedio),
                f"{rentabilidad_promedio:.2f}%",
                formato_pesos(resultado_total),
            ],
        }
    )

    st.dataframe(
        resumen,
        use_container_width=True,
        hide_index=True,
    )

# ============================================================
# ARTÍCULOS
# ============================================================

with tab_articulos:

    st.subheader(
        "🔧 Ranking de artículos"
    )

    ranking_articulos = (
        df.groupby(
            ["Cod", "Articulo"],
            as_index=False,
            dropna=False,
        )
        .agg(
            Unidades=("Cantidad", "sum"),
            Facturacion=("Total Venta", "sum"),
            Resultado=("Resultado Venta-Costo $", "sum"),
            Rentabilidad_Promedio=("Rentabilidad", "mean"),
            Operaciones=("Cod", "size"),
        )
        .sort_values(
            "Facturacion",
            ascending=False,
        )
        .reset_index(drop=True)
    )

    ranking_articulos.insert(
        0,
        "Puesto",
        range(
            1,
            len(ranking_articulos) + 1,
        ),
    )

    vista = ranking_articulos.copy()

    vista["Facturación"] = (
        vista["Facturacion"]
        .apply(formato_pesos)
    )

    vista["Resultado $"] = (
        vista["Resultado"]
        .apply(formato_pesos)
    )

    vista["Rentabilidad %"] = (
        vista["Rentabilidad_Promedio"]
        .apply(
            lambda x: f"{x:.2f}%"
        )
    )

    vista = vista[
        [
            "Puesto",
            "Cod",
            "Articulo",
            "Unidades",
            "Facturación",
            "Resultado $",
            "Rentabilidad %",
            "Operaciones",
        ]
    ]

    st.dataframe(
        vista.head(50),
        use_container_width=True,
        hide_index=True,
    )

    st.subheader(
        "📊 Artículos con mayor facturación"
    )

    grafico = (
        ranking_articulos
        .sort_values(
            "Facturacion",
            ascending=False,
        )
        .head(15)
        .copy()
    )

    grafico["Artículo"] = (
        grafico["Cod"]
        + " - "
        + grafico["Articulo"]
    )

    st.bar_chart(
        grafico.set_index(
            "Artículo"
        )["Facturacion"]
    )

    st.subheader(
        "📦 Artículos por cantidad"
    )

    grafico_cantidad = (
        ranking_articulos
        .sort_values(
            "Unidades",
            ascending=False,
        )
        .head(15)
        .copy()
    )

    grafico_cantidad["Artículo"] = (
        grafico_cantidad["Cod"]
        + " - "
        + grafico_cantidad["Articulo"]
    )

    st.bar_chart(
        grafico_cantidad.set_index(
            "Artículo"
        )["Unidades"]
    )

# ============================================================
# CLIENTES
# ============================================================

with tab_clientes:

    st.subheader(
        "👥 Ranking de compradores"
    )

    ranking_clientes = (
        df.groupby(
            "Cliente",
            as_index=False,
            dropna=False,
        )
        .agg(
            Facturacion=("Total Venta", "sum"),
            Unidades=("Cantidad", "sum"),
            Resultado=("Resultado Venta-Costo $", "sum"),
            Rentabilidad_Promedio=("Rentabilidad", "mean"),
            Operaciones=("Cliente", "size"),
        )
        .sort_values(
            "Facturacion",
            ascending=False,
        )
        .reset_index(drop=True)
    )

    ranking_clientes.insert(
        0,
        "Puesto",
        range(
            1,
            len(ranking_clientes) + 1,
        ),
    )

    vista = ranking_clientes.copy()

    vista["Facturación"] = (
        vista["Facturacion"]
        .apply(formato_pesos)
    )

    vista["Resultado $"] = (
        vista["Resultado"]
        .apply(formato_pesos)
    )

    vista["Rentabilidad %"] = (
        vista["Rentabilidad_Promedio"]
        .apply(
            lambda x: f"{x:.2f}%"
        )
    )

    vista = vista[
        [
            "Puesto",
            "Cliente",
            "Facturación",
            "Unidades",
            "Resultado $",
            "Rentabilidad %",
            "Operaciones",
        ]
    ]

    # IMPORTANTE:
    # hide_index=True elimina el número automático
    # que aparecía a la izquierda de la tabla.
    st.dataframe(
        vista.head(50),
        use_container_width=True,
        hide_index=True,
    )

    st.subheader(
        "📊 Compradores con mayor facturación"
    )

    grafico = (
        ranking_clientes
        .head(15)
        .copy()
    )

    st.bar_chart(
        grafico.set_index(
            "Cliente"
        )["Facturacion"]
    )

# ============================================================
# PAGOS
# ============================================================

with tab_pagos:

    st.subheader(
        "💳 Contado vs cuenta corriente"
    )

    ranking_tipo_pago = (
        df.groupby(
            "Tipo Pago",
            as_index=False,
        )
        .agg(
            Facturacion=("Total Venta", "sum"),
            Unidades=("Cantidad", "sum"),
            Operaciones=("Tipo Pago", "size"),
        )
        .sort_values(
            "Facturacion",
            ascending=False,
        )
        .reset_index(drop=True)
    )

    ranking_tipo_pago.insert(
        0,
        "Puesto",
        range(
            1,
            len(ranking_tipo_pago) + 1,
        ),
    )

    vista = ranking_tipo_pago.copy()

    vista["Facturación"] = (
        vista["Facturacion"]
        .apply(formato_pesos)
    )

    st.dataframe(
        vista[
            [
                "Puesto",
                "Tipo Pago",
                "Facturación",
                "Unidades",
                "Operaciones",
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )

    st.subheader(
        "📊 Facturación por tipo de pago"
    )

    st.bar_chart(
        ranking_tipo_pago.set_index(
            "Tipo Pago"
        )["Facturacion"]
    )

    st.subheader(
        "🥧 Distribución de facturación"
    )

    try:
        import plotly.express as px

        pie = ranking_tipo_pago[
            ranking_tipo_pago["Facturacion"] != 0
        ].copy()

        if not pie.empty:

            fig = px.pie(
                pie,
                names="Tipo Pago",
                values="Facturacion",
                hole=0.35,
            )

            fig.update_layout(
                margin=dict(
                    l=10,
                    r=10,
                    t=30,
                    b=10,
                ),
            )

            st.plotly_chart(
                fig,
                use_container_width=True,
            )

        else:

            st.info(
                "No hay facturación distinta de cero "
                "para mostrar."
            )

    except ImportError:

        st.warning(
            "Falta Plotly. Verificá requirements.txt."
        )

    st.subheader(
        "💳 Formas de pago exactas"
    )

    ranking_forma_pago = (
        df.groupby(
            "Forma Pago",
            as_index=False,
        )
        .agg(
            Facturacion=("Total Venta", "sum"),
            Unidades=("Cantidad", "sum"),
        )
        .sort_values(
            "Facturacion",
            ascending=False,
        )
        .reset_index(drop=True)
    )

    ranking_forma_pago["Facturación"] = (
        ranking_forma_pago["Facturacion"]
        .apply(formato_pesos)
    )

    st.dataframe(
        ranking_forma_pago[
            [
                "Forma Pago",
                "Facturación",
                "Unidades",
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )

# ============================================================
# EVOLUCIÓN
# ============================================================

with tab_evolucion:

    st.subheader(
        "📅 Evolución diaria"
    )

    ventas_dia = (
        df.assign(
            DiaFecha=df["Fecha"].dt.date
        )
        .groupby(
            "DiaFecha",
            as_index=False,
        )
        .agg(
            Facturacion=("Total Venta", "sum"),
            Unidades=("Cantidad", "sum"),
            Resultado=("Resultado Venta-Costo $", "sum"),
        )
        .sort_values(
            "DiaFecha"
        )
    )

    st.line_chart(
        ventas_dia.set_index(
            "DiaFecha"
        )[
            ["Facturacion"]
        ]
    )

    st.subheader(
        "📦 Unidades por día"
    )

    st.line_chart(
        ventas_dia.set_index(
            "DiaFecha"
        )[
            ["Unidades"]
        ]
    )

    vista = ventas_dia.copy()

    vista["Facturación"] = (
        vista["Facturacion"]
        .apply(formato_pesos)
    )

    vista["Resultado $"] = (
        vista["Resultado"]
        .apply(formato_pesos)
    )

    st.dataframe(
        vista[
            [
                "DiaFecha",
                "Facturación",
                "Unidades",
                "Resultado $",
            ]
        ].sort_values(
            "DiaFecha",
            ascending=False,
        ),
        use_container_width=True,
        hide_index=True,
    )

# ============================================================
# DETALLE
# ============================================================

with tab_detalle:

    st.subheader(
        "📋 Detalle de operaciones"
    )

    columnas_mostrar = [
        "Fecha",
        "Cod",
        "Articulo",
        "Cantidad",
        "Cliente",
        "Precio Unitario",
        "Total Venta",
        "Forma Pago",
        "Tipo Pago",
        "Rentabilidad",
        "Resultado Venta-Costo $",
    ]

    columnas_mostrar = [
        c
        for c in columnas_mostrar
        if c in df.columns
    ]

    detalle = (
        df[columnas_mostrar]
        .sort_values(
            "Fecha",
            ascending=False,
        )
        .reset_index(drop=True)
    )

    detalle.insert(
        0,
        "Nº",
        range(
            1,
            len(detalle) + 1,
        ),
    )

    detalle["Fecha"] = (
        detalle["Fecha"]
        .dt.strftime("%d/%m/%Y")
    )

    for columna in [
        "Precio Unitario",
        "Total Venta",
        "Resultado Venta-Costo $",
    ]:

        if columna in detalle.columns:
            detalle[columna] = (
                detalle[columna]
                .apply(formato_pesos)
            )

    if "Rentabilidad" in detalle.columns:

        detalle["Rentabilidad"] = (
            detalle["Rentabilidad"]
            .apply(
                lambda x: f"{x:.2f}%"
            )
        )

    st.dataframe(
        detalle,
        use_container_width=True,
        hide_index=True,
    )

    # Descargar sin columnas internas.
    csv = (
        df.drop(
            columns=[
                "_FilaArchivo",
                "_FirmaFila",
            ],
            errors="ignore",
        )
        .to_csv(index=False)
        .encode("utf-8-sig")
    )

    st.download_button(
        "📥 Descargar datos filtrados",
        data=csv,
        file_name="ventas_filtradas.csv",
        mime="text/csv",
    )
