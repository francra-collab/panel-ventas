import streamlit as st
import pandas as pd
import requests
import io
import base64
from datetime import datetime

# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Panel de Ventas",
    page_icon="📊",
    layout="wide"
)

HISTORIAL_PATH = "data/historial.csv"

# ============================================================
# FUNCIONES GENERALES
# ============================================================

def numero(valor):
    """
    Convierte números argentinos:
    12.345,67
    12345,67
    12345.67
    """
    if pd.isna(valor):
        return 0.0

    texto = str(valor).strip()

    if texto == "":
        return 0.0

    try:
        # Si tiene punto y coma, asumimos formato argentino
        if "." in texto and "," in texto:
            texto = texto.replace(".", "").replace(",", ".")
        elif "," in texto:
            texto = texto.replace(",", ".")

        return float(texto)

    except Exception:
        return 0.0


def leer_infoventas(archivo):
    """
    Intenta leer:
    - XLSX
    - XLS
    - TXT/CSV/tabulado
    """

    contenido = archivo.getvalue()

    # Primero intentamos Excel
    try:
        df = pd.read_excel(io.BytesIO(contenido))
        if len(df.columns) > 3:
            return df
    except Exception:
        pass

    # Si no es Excel, intentamos texto tabulado
    for encoding in ["cp1252", "latin1", "utf-8"]:
        try:
            df = pd.read_csv(
                io.BytesIO(contenido),
                sep="\t",
                encoding=encoding
            )

            if len(df.columns) > 3:
                return df
        except Exception:
            pass

    # Último intento: CSV normal
    for encoding in ["cp1252", "latin1", "utf-8"]:
        try:
            df = pd.read_csv(
                io.BytesIO(contenido),
                encoding=encoding
            )

            if len(df.columns) > 3:
                return df
        except Exception:
            pass

    raise ValueError(
        "No pude reconocer el formato del archivo INFOVENTAS."
    )


def preparar_datos(df):

    # Limpiar nombres de columnas
    df.columns = [
        str(c).strip()
        for c in df.columns
    ]

    # Eliminar columnas completamente vacías
    df = df.dropna(axis=1, how="all")

    columnas_necesarias = [
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
        "Forma Pago"
    ]

    faltantes = [
        c for c in columnas_necesarias
        if c not in df.columns
    ]

    if faltantes:
        raise ValueError(
            "Faltan estas columnas en el INFOVENTAS: "
            + ", ".join(faltantes)
        )

    # --------------------------------------------------------
    # Fechas
    # --------------------------------------------------------

    df["Fecha"] = pd.to_datetime(
        df["Fecha"],
        dayfirst=True,
        errors="coerce"
    )

    # --------------------------------------------------------
    # Números
    # --------------------------------------------------------

    columnas_numericas = [
        "Cantidad",
        "Precio Costo",
        "Total Costo",
        "Rentabilidad",
        "Precio Vent.",
        "Total Venta"
    ]

    for columna in columnas_numericas:
        df[columna] = df[columna].apply(numero)

    # --------------------------------------------------------
    # Texto
    # --------------------------------------------------------

    columnas_texto = [
        "Cod",
        "Articulo",
        "Cliente",
        "Forma Pago"
    ]

    for columna in columnas_texto:
        df[columna] = (
            df[columna]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    # --------------------------------------------------------
    # Precio real por unidad
    # --------------------------------------------------------

    df["Precio Unitario"] = 0.0

    mask = df["Cantidad"] != 0

    df.loc[mask, "Precio Unitario"] = (
        df.loc[mask, "Total Venta"]
        / df.loc[mask, "Cantidad"]
    )

    # --------------------------------------------------------
    # Rentabilidad en pesos
    # --------------------------------------------------------

    df["Rentabilidad $"] = (
        df["Total Venta"]
        - df["Total Costo"]
    )

    # --------------------------------------------------------
    # Día
    # --------------------------------------------------------

    df["Dia"] = df["Fecha"].dt.date.astype(str)

    # --------------------------------------------------------
    # Identificador de operación
    # --------------------------------------------------------

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

    return (
        "GITHUB_TOKEN" in st.secrets
        and "GITHUB_REPO" in st.secrets
        and "GITHUB_USER" in st.secrets
    )


def github_headers():

    return {
        "Authorization": f"Bearer {st.secrets['GITHUB_TOKEN']}",
        "Accept": "application/vnd.github+json"
    }


def github_url():

    return (
        f"https://api.github.com/repos/"
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
            timeout=20
        )

        if respuesta.status_code == 404:
            return pd.DataFrame()

        respuesta.raise_for_status()

        contenido = respuesta.json()["content"]

        datos = base64.b64decode(contenido)

        df = pd.read_csv(
            io.BytesIO(datos)
        )

        if "Fecha" in df.columns:
            df["Fecha"] = pd.to_datetime(
                df["Fecha"],
                errors="coerce"
            )

        return df

    except Exception as e:

        st.warning(
            f"No se pudo cargar el historial: {e}"
        )

        return pd.DataFrame()


def guardar_historial(df):

    if not github_configurado():

        st.error(
            "Todavía no está configurado GitHub para guardar "
            "el historial."
        )

        return False

    try:

        contenido_csv = df.to_csv(
            index=False
        ).encode("utf-8")

        contenido_base64 = base64.b64encode(
            contenido_csv
        ).decode("utf-8")

        # Buscar archivo existente
        respuesta = requests.get(
            github_url(),
            headers=github_headers(),
            timeout=20
        )

        sha = None

        if respuesta.status_code == 200:
            sha = respuesta.json()["sha"]

        datos = {
            "message": "Actualizar historial de ventas",
            "content": contenido_base64,
            "branch": "main"
        }

        if sha:
            datos["sha"] = sha

        respuesta = requests.put(
            github_url(),
            headers=github_headers(),
            json=datos,
            timeout=30
        )

        respuesta.raise_for_status()

        return True

    except Exception as e:

        st.error(
            f"No se pudo guardar el historial: {e}"
        )

        return False


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
        errors="coerce"
    )

# ============================================================
# SUBIR ARCHIVO
# ============================================================

st.subheader("📁 Cargar INFOVENTAS")

archivo = st.file_uploader(
    "Subí el INFOVENTAS del día",
    type=["xlsx", "xls", "csv", "txt"]
)

if archivo is not None:

    try:

        df_nuevo = leer_infoventas(archivo)

        df_nuevo = preparar_datos(df_nuevo)

        st.success(
            f"Archivo cargado correctamente: "
            f"{len(df_nuevo):,} operaciones."
        )

        # ----------------------------------------------------
        # FECHAS DEL ARCHIVO
        # ----------------------------------------------------

        fechas_archivo = sorted(
            df_nuevo["Fecha"]
            .dropna()
            .dt.date
            .unique()
        )

        st.info(
            "Fecha/s detectada/s: "
            + ", ".join(
                str(f)
                for f in fechas_archivo
            )
        )

        # ----------------------------------------------------
        # GUARDAR HISTÓRICO
        # ----------------------------------------------------

        if github_configurado():

            ids_nuevos = set(
                df_nuevo["ID Operacion"]
            )

            if not historial.empty:

                ids_existentes = set(
                    historial["ID Operacion"]
                    if "ID Operacion" in historial.columns
                    else []
                )

                df_para_agregar = df_nuevo[
                    ~df_nuevo["ID Operacion"].isin(
                        ids_existentes
                    )
                ].copy()

            else:

                df_para_agregar = df_nuevo.copy()

            if len(df_para_agregar) > 0:

                if historial.empty:

                    historial_actualizado = (
                        df_para_agregar.copy()
                    )

                else:

                    historial_actualizado = pd.concat(
                        [
                            historial,
                            df_para_agregar
                        ],
                        ignore_index=True
                    )

                historial_actualizado = (
                    historial_actualizado
                    .drop_duplicates(
                        subset=["ID Operacion"]
                    )
                )

                if st.button(
                    "💾 Guardar este día en el histórico",
                    type="primary"
                ):

                    with st.spinner(
                        "Guardando histórico..."
                    ):

                        if guardar_historial(
                            historial_actualizado
                        ):

                            st.success(
                                "✅ Día guardado correctamente."
                            )

                            historial = (
                                historial_actualizado
                            )

                            st.rerun()

            else:

                st.warning(
                    "⚠️ Este archivo ya está cargado "
                    "en el histórico."
                )

        else:

            st.warning(
                "⚠️ El histórico todavía no está "
                "configurado. El dashboard funciona, "
                "pero todavía no guardará los días."
            )

    except Exception as e:

        st.error(
            f"Error procesando el archivo: {e}"
        )

        st.stop()

# ============================================================
# ELEGIR DATOS PARA MOSTRAR
# ============================================================

if historial.empty:

    st.info(
        "Subí un INFOVENTAS para comenzar."
    )

    st.stop()

df = historial.copy()

df["Fecha"] = pd.to_datetime(
    df["Fecha"],
    errors="coerce"
)

# ============================================================
# FILTROS
# ============================================================

st.sidebar.header("🔎 Filtros")

fecha_min = df["Fecha"].min().date()
fecha_max = df["Fecha"].max().date()

rango_fechas = st.sidebar.date_input(
    "Período",
    value=(fecha_min, fecha_max),
    min_value=fecha_min,
    max_value=fecha_max
)

if isinstance(rango_fechas, tuple) and len(rango_fechas) == 2:

    fecha_inicio = pd.Timestamp(
        rango_fechas[0]
    )

    fecha_fin = pd.Timestamp(
        rango_fechas[1]
    ) + pd.Timedelta(days=1)

    df = df[
        (df["Fecha"] >= fecha_inicio)
        & (df["Fecha"] < fecha_fin)
    ]

# ------------------------------------------------------------
# SUCURSAL
# ------------------------------------------------------------

if "Suc." in df.columns:

    sucursales = sorted(
        df["Suc."]
        .dropna()
        .astype(str)
        .unique()
    )

    sucursal = st.sidebar.multiselect(
        "Sucursal",
        sucursales
    )

    if sucursal:

        df = df[
            df["Suc."].astype(str).isin(
                sucursal
            )
        ]

# ------------------------------------------------------------
# CLIENTE
# ------------------------------------------------------------

clientes = sorted(
    df["Cliente"]
    .dropna()
    .unique()
)

cliente = st.sidebar.multiselect(
    "Cliente",
    clientes
)

if cliente:

    df = df[
        df["Cliente"].isin(cliente)
    ]

# ------------------------------------------------------------
# FORMA DE PAGO
# ------------------------------------------------------------

pagos = sorted(
    df["Forma Pago"]
    .dropna()
    .unique()
)

pago = st.sidebar.multiselect(
    "Forma de pago",
    pagos
)

if pago:

    df = df[
        df["Forma Pago"].isin(pago)
    ]

# ------------------------------------------------------------
# ARTÍCULO
# ------------------------------------------------------------

articulos = sorted(
    df["Articulo"]
    .dropna()
    .unique()
)

articulo = st.sidebar.multiselect(
    "Artículo",
    articulos
)

if articulo:

    df = df[
        df["Articulo"].isin(articulo)
    ]

# ============================================================
# SI NO HAY DATOS
# ============================================================

if df.empty:

    st.warning(
        "No hay operaciones que coincidan "
        "con los filtros."
    )

    st.stop()

# ============================================================
# KPIs
# ============================================================

total_venta = df["Total Venta"].sum()

unidades = df["Cantidad"].sum()

comprobantes = (
    df["Nº Comprobante"].nunique()
    if "Nº Comprobante" in df.columns
    else len(df)
)

clientes_unicos = df["Cliente"].nunique()

articulos_unicos = df["Cod"].nunique()

precio_promedio = (
    df["Total Venta"].sum()
    / df["Cantidad"].sum()
    if df["Cantidad"].sum() != 0
    else 0
)

ticket_promedio = (
    total_venta / comprobantes
    if comprobantes != 0
    else 0
)

rentabilidad_total = df["Rentabilidad $"].sum()

rentabilidad_promedio = (
    df["Rentabilidad"].mean()
)

# ============================================================
# TABLA PRINCIPAL DE KPIs
# ============================================================

st.subheader("📌 Resumen")

col1, col2, col3, col4 = st.columns(4)

col1.metric(
    "💰 Facturación",
    f"$ {total_venta:,.2f}"
)

col2.metric(
    "📦 Unidades",
    f"{unidades:,.0f}"
)

col3.metric(
    "🧾 Comprobantes",
    f"{comprobantes:,}"
)

col4.metric(
    "👥 Clientes",
    f"{clientes_unicos:,}"
)

col5, col6, col7, col8 = st.columns(4)

col5.metric(
    "💵 Precio promedio",
    f"$ {precio_promedio:,.2f}"
)

col6.metric(
    "🛒 Ticket promedio",
    f"$ {ticket_promedio:,.2f}"
)

col7.metric(
    "📈 Rentabilidad promedio",
    f"{rentabilidad_promedio:.2f}%"
)

col8.metric(
    "💰 Rentabilidad $",
    f"$ {rentabilidad_total:,.2f}"
)

st.divider()

# ============================================================
# RANKING ARTÍCULOS
# ============================================================

st.subheader("🏆 Ranking de artículos")

ranking_articulos = (
    df.groupby(
        ["Cod", "Articulo"],
        as_index=False
    )
    .agg(
        Unidades=("Cantidad", "sum"),
        Facturacion=("Total Venta", "sum"),
        Rentabilidad=("Rentabilidad $", "sum"),
        Rentabilidad_Promedio=("Rentabilidad", "mean")
    )
    .sort_values(
        "Facturacion",
        ascending=False
    )
)

tab1, tab2, tab3, tab4 = st.tabs(
    [
        "💰 Facturación",
        "📦 Cantidad",
        "📈 Rentabilidad $",
        "%",
    ]
)

with tab1:

    st.dataframe(
        ranking_articulos[
            [
                "Cod",
                "Articulo",
                "Unidades",
                "Facturacion",
                "Rentabilidad_Promedio"
            ]
        ].head(30),
        use_container_width=True
    )

with tab2:

    st.dataframe(
        ranking_articulos
        .sort_values(
            "Unidades",
            ascending=False
        )
        [
            [
                "Cod",
                "Articulo",
                "Unidades",
                "Facturacion"
            ]
        ]
        .head(30),
        use_container_width=True
    )

with tab3:

    st.dataframe(
        ranking_articulos
        .sort_values(
            "Rentabilidad",
            ascending=False
        )
        [
            [
                "Cod",
                "Articulo",
                "Facturacion",
                "Rentabilidad"
            ]
        ]
        .head(30),
        use_container_width=True
    )

with tab4:

    st.dataframe(
        ranking_articulos
        .sort_values(
            "Rentabilidad_Promedio",
            ascending=False
        )
        [
            [
                "Cod",
                "Articulo",
                "Rentabilidad_Promedio"
            ]
        ]
        .head(30),
        use_container_width=True
    )

# ============================================================
# GRÁFICO ARTÍCULOS
# ============================================================

st.subheader("📊 Artículos con mayor facturación")

grafico_articulos = (
    ranking_articulos
    .sort_values(
        "Facturacion",
        ascending=False
    )
    .head(15)
    .copy()
)

grafico_articulos["Nombre"] = (
    grafico_articulos["Cod"]
    + " - "
    + grafico_articulos["Articulo"]
)

st.bar_chart(
    grafico_articulos.set_index(
        "Nombre"
    )["Facturacion"]
)

# ============================================================
# CLIENTES
# ============================================================

st.subheader("👥 Ranking de clientes")

ranking_clientes = (
    df.groupby(
        "Cliente",
        as_index=False
    )
    .agg(
        Facturacion=("Total Venta", "sum"),
        Unidades=("Cantidad", "sum"),
        Rentabilidad=("Rentabilidad $", "sum"),
        Rentabilidad_Promedio=("Rentabilidad", "mean"),
        Operaciones=("Cliente", "size")
    )
    .sort_values(
        "Facturacion",
        ascending=False
    )
)

st.dataframe(
    ranking_clientes.head(30),
    use_container_width=True
)

# ============================================================
# GRÁFICO CLIENTES
# ============================================================

grafico_clientes = (
    ranking_clientes
    .head(15)
    .copy()
)

st.bar_chart(
    grafico_clientes.set_index(
        "Cliente"
    )["Facturacion"]
)

# ============================================================
# FORMA DE PAGO
# ============================================================

st.subheader("💳 Forma de pago")

ranking_pago = (
    df.groupby(
        "Forma Pago",
        as_index=False
    )
    .agg(
        Facturacion=("Total Venta", "sum"),
        Unidades=("Cantidad", "sum"),
        Rentabilidad=("Rentabilidad $", "sum")
    )
    .sort_values(
        "Facturacion",
        ascending=False
    )
)

st.dataframe(
    ranking_pago,
    use_container_width=True
)

st.bar_chart(
    ranking_pago.set_index(
        "Forma Pago"
    )["Facturacion"]
)

# ============================================================
# VENTAS POR DÍA
# ============================================================

st.subheader("📅 Evolución de ventas")

ventas_dia = (
    df.groupby(
        df["Fecha"].dt.date
    )
    .agg(
        Facturacion=("Total Venta", "sum"),
        Unidades=("Cantidad", "sum"),
        Rentabilidad=("Rentabilidad $", "sum")
    )
)

st.line_chart(
    ventas_dia[
        [
            "Facturacion",
            "Rentabilidad"
        ]
    ]
)

# ============================================================
# TABLA DE OPERACIONES
# ============================================================

st.subheader("📋 Detalle de operaciones")

columnas_mostrar = [
    "Fecha",
    "Cod",
    "Articulo",
    "Cantidad",
    "Cliente",
    "Precio Unitario",
    "Total Venta",
    "Forma Pago",
    "Rentabilidad",
    "Rentabilidad $"
]

columnas_mostrar = [
    c for c in columnas_mostrar
    if c in df.columns
]

st.dataframe(
    df[columnas_mostrar]
    .sort_values(
        "Fecha",
        ascending=False
    ),
    use_container_width=True
)

# ============================================================
# DESCARGAR DATOS FILTRADOS
# ============================================================

st.subheader("⬇️ Descargar")

csv = df.to_csv(
    index=False
).encode("utf-8-sig")

st.download_button(
    "📥 Descargar datos filtrados",
    data=csv,
    file_name="ventas_filtradas.csv",
    mime="text/csv"
)
