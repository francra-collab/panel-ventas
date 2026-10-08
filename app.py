import base64
import io
import textwrap

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st

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

COLUMNAS_OPCIONALES = [
    "TC",
    "Suc.",
    "Nº Comprobante",
    "Lista Utilizada",
    "Porcentaje Comisión",
    "Monto Comisión",
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
    "TC",
    "Suc.",
    "Nº Comprobante",
    "Lista Utilizada",
]

# IVA incluido en "Total Venta" (se usa para estimar la ganancia real).
IVA = 0.21

# Columnas que se guardan en el histórico (todo lo demás se recalcula al cargar)
COLUMNAS_GUARDAR = COLUMNAS_REQUERIDAS + COLUMNAS_OPCIONALES + ["_Clave"]

DIAS_SEMANA = {
    0: "Lunes",
    1: "Martes",
    2: "Miércoles",
    3: "Jueves",
    4: "Viernes",
    5: "Sábado",
    6: "Domingo",
}

# Paleta de los gráficos
AZUL = "#2563EB"
VERDE = "#16A34A"
ROJO = "#DC2626"
NARANJA = "#F59E0B"
GRIS = "#64748B"
PALETA = [AZUL, VERDE, NARANJA, "#7C3AED", "#0891B2", ROJO, GRIS]

# Colores fijos por forma de pago: cada una se distingue en todos los gráficos
COLORES_TIPO_PAGO = {
    "CONTADO": VERDE,
    "CUENTA CORRIENTE": AZUL,
    "OTRO": GRIS,
}
COLORES_FORMA_PAGO = {
    "CONTADO": VERDE,
    "CUENTA CORRIENTE 7 DIAS": NARANJA,
    "CUENTA CORRIENTE 10 DIAS": AZUL,
    "CUENTA CORRIENTE 15 DIAS": "#7C3AED",
}
EXTRA_COLORES = ["#0891B2", "#DB2777", "#65A30D", "#92400E", "#0F766E"]


def color_pago(nombre, tipo="forma"):
    mapa = COLORES_FORMA_PAGO if tipo == "forma" else COLORES_TIPO_PAGO
    nombre = str(nombre).strip().upper()
    if nombre in mapa:
        return mapa[nombre]
    # Forma de pago nueva que no conocemos: color estable según su nombre
    return EXTRA_COLORES[sum(ord(c) for c in nombre) % len(EXTRA_COLORES)]


# Los gráficos no son interactivos con la rueda del mouse, así que
# al scrollear la página no se mueven ni se achican.
CONFIG_GRAFICO = {
    "displayModeBar": False,
    "scrollZoom": False,
    "doubleClick": False,
    "responsive": True,
}

st.markdown(
    """
    <style>
        .block-container {padding-top: 2rem; padding-bottom: 3rem;}
        [data-testid="stMetricValue"] {font-size: 1.55rem;}
        [data-testid="stMetricLabel"] p {font-size: 0.9rem; color: #64748B;}
        div[data-testid="stMetric"] {
            background: #F8FAFC;
            border: 1px solid #E2E8F0;
            border-radius: 12px;
            padding: 14px 16px;
        }
        div[data-testid="stPlotlyChart"] {width: 100%;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# FORMATO (estilo argentino: 1.234,56 y dd/mm/aaaa)
# ============================================================

def formato_numero(valor, decimales=2):
    try:
        if pd.isna(valor):
            valor = 0
        texto = f"{float(valor):,.{decimales}f}"
    except (ValueError, TypeError):
        texto = f"{0:,.{decimales}f}"
    return texto.replace(",", "§").replace(".", ",").replace("§", ".")


def formato_pesos(valor):
    try:
        valor = 0.0 if pd.isna(valor) else float(valor)
    except (ValueError, TypeError):
        valor = 0.0
    signo = "-" if valor < 0 else ""
    return f"{signo}$ {formato_numero(abs(valor), 2)}"


def formato_entero(valor):
    return formato_numero(valor, 0)


def formato_porcentaje(valor):
    return f"{formato_numero(valor, 2)} %"


def formato_fecha(serie):
    return pd.to_datetime(serie, errors="coerce").dt.strftime("%d/%m/%Y")


# ============================================================
# LECTURA Y LIMPIEZA
# ============================================================

def numero(valor):
    """Convierte números argentinos (1.234,56) y normales a float."""
    if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip().replace("$", "").replace(" ", "")

    if texto in ("", "nan", "None", "-"):
        return 0.0

    try:
        if "." in texto and "," in texto:
            texto = texto.replace(".", "").replace(",", ".")
        elif "," in texto:
            texto = texto.replace(",", ".")
        elif texto.count(".") > 1:
            # 1.234.567 → miles
            texto = texto.replace(".", "")
        return float(texto)
    except (ValueError, TypeError):
        return 0.0


def limpiar_columnas(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, [c for c in df.columns if not c.startswith("Unnamed")]]
    df = df.dropna(axis=1, how="all")
    return df


def leer_infoventas(archivo):
    """
    Lee el INFOVENTAS. Aunque se llame .xls, el sistema suele exportar
    un texto separado por tabulaciones, así que detectamos el formato real.
    """
    contenido = archivo.getvalue()

    es_excel_real = contenido[:4] in (b"PK\x03\x04", b"\xd0\xcf\x11\xe0")

    if es_excel_real:
        for motor in ("openpyxl", "xlrd"):
            try:
                df = pd.read_excel(
                    io.BytesIO(contenido),
                    engine=motor,
                    dtype=str,
                )
                df = limpiar_columnas(df)
                if len(df.columns) > 3:
                    return df
            except Exception:
                pass

    for separador in ("\t", ";", ","):
        for encoding in ("cp1252", "latin1", "utf-8-sig", "utf-8"):
            try:
                df = pd.read_csv(
                    io.BytesIO(contenido),
                    sep=separador,
                    encoding=encoding,
                    dtype=str,
                )
                df = limpiar_columnas(df)
                if len(df.columns) > 3:
                    return df
            except Exception:
                pass

    raise ValueError(
        "No pude reconocer el archivo. "
        "Probá con el archivo original de INFOVENTAS."
    )


def parsear_fecha(serie):
    """Fechas dd/mm/aaaa primero; si no, cualquier otro formato (ISO, Excel)."""
    serie = serie.astype(str).str.strip()
    fecha = pd.to_datetime(serie, format="%d/%m/%Y", errors="coerce")
    resto = fecha.isna()
    if resto.any():
        fecha.loc[resto] = pd.to_datetime(
            serie[resto],
            errors="coerce",
        )
    return fecha


def normalizar_comprobante(serie):
    return (
        serie.fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.lstrip("0")
    )


def agregar_clave(df):
    """
    Clave única por renglón: sirve para no cargar dos veces lo mismo.
    Dos renglones realmente idénticos se distinguen con un contador,
    así que re-subir un archivo no duplica nada y no depende del orden.
    """
    base = (
        pd.to_datetime(df["Fecha"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
        + "|" + df["TC"].fillna("").astype(str).str.strip()
        + "|" + normalizar_comprobante(df["Nº Comprobante"])
        + "|" + df["Cod"].fillna("").astype(str).str.strip()
        + "|" + df["Cliente"].fillna("").astype(str).str.strip()
        + "|" + df["Cantidad"].round(2).astype(str)
        + "|" + df["Total Venta"].round(2).astype(str)
        + "|" + df["Total Costo"].round(2).astype(str)
    )
    contador = base.groupby(base).cumcount().astype(str)
    df["_Clave"] = base + "#" + contador
    return df


def preparar_columnas_base(df):
    """Deja el dataframe con tipos correctos (sirve para archivo nuevo e histórico)."""
    df = limpiar_columnas(df)

    faltantes = [c for c in COLUMNAS_REQUERIDAS if c not in df.columns]
    if faltantes:
        raise ValueError(
            "Faltan columnas obligatorias: " + ", ".join(faltantes)
        )

    for columna in COLUMNAS_OPCIONALES:
        if columna not in df.columns:
            df[columna] = ""

    df["Fecha"] = parsear_fecha(df["Fecha"])
    df = df[df["Fecha"].notna()].copy()

    if df.empty:
        raise ValueError(
            "No pude reconocer ninguna fecha en la columna 'Fecha'."
        )

    for columna in COLUMNAS_NUMERICAS:
        df[columna] = df[columna].apply(numero)

    for columna in COLUMNAS_TEXTO:
        df[columna] = df[columna].fillna("").astype(str).str.strip()

    return df.reset_index(drop=True)


def clasificar_pago(serie):
    """CONTADO / CUENTA CORRIENTE / OTRO (tolera 'CORRENTE', 'CTA CTE', etc.)."""
    forma = serie.fillna("").astype(str).str.upper()

    tipo = pd.Series("OTRO", index=serie.index)
    tipo[forma.str.contains("CONTADO", na=False)] = "CONTADO"
    tipo[
        forma.str.contains(r"CUENTA\s+COR+I?ENTE|CTA\.?\s*CTE|CTA\.?\s*CORR", regex=True, na=False)
    ] = "CUENTA CORRIENTE"
    return tipo


def enriquecer(df):
    """
    Calcula todo lo derivado. Se aplica SIEMPRE sobre el histórico completo,
    así los datos viejos también quedan corregidos.

    Claves del arreglo de "valores raros / negativos":
    - INFOVENTAS informa la Cantidad SIEMPRE positiva, incluso en notas de
      crédito (devoluciones, saldos cancelados). Los importes sí vienen en
      negativo. Antes se sumaban las unidades de las NC como si fueran
      ventas y el precio unitario salía mal.
    - Acá se detectan las devoluciones y se les pone cantidad negativa.
    - La rentabilidad se calcula ponderada por importes, no como promedio
      de porcentajes (los renglones de NC traen 0 % y arruinaban el promedio).
    """
    df = df.copy()

    df["Forma Pago"] = (
        df["Forma Pago"]
        .str.upper()
        .str.replace("CORRENTE", "CORRIENTE", regex=False)
    )
    df["Tipo Pago"] = clasificar_pago(df["Forma Pago"])

    tc = df["TC"].str.upper().str.strip()
    df["Es Devolucion"] = (
        tc.str.startswith("NC")
        | (df["Total Venta"] < 0)
        | (df["Total Costo"] < 0)
    )

    signo = df["Es Devolucion"].map({True: -1.0, False: 1.0})
    df["Cantidad Neta"] = df["Cantidad"].abs() * signo
    df["Tipo Operacion"] = df["Es Devolucion"].map(
        {True: "Devolución / Nota de crédito", False: "Venta"}
    )

    df["Comprobante"] = (
        df["TC"].str.strip() + " " + normalizar_comprobante(df["Nº Comprobante"])
    ).str.strip()

    df["Precio Unitario"] = 0.0
    mascara = df["Cantidad"] != 0
    df.loc[mascara, "Precio Unitario"] = (
        df.loc[mascara, "Total Venta"].abs() / df.loc[mascara, "Cantidad"].abs()
    )

    # "Total Costo" de INFOVENTAS NO es el costo real: es el precio de lista
    # sin IVA (Total Venta / Total Costo = 1,21). Restarlos da casi solo IVA.
    # El costo real es Precio Costo x Cantidad. Las notas de crédito y los
    # renglones "TODO" vienen con Precio Costo = 0 (sin costo conocido) y
    # no entran en el cálculo de ganancia.
    df["Con Costo"] = df["Precio Costo"] > 0
    df["Costo Real"] = df["Precio Costo"] * df["Cantidad Neta"]
    df["Venta sin IVA"] = df["Total Venta"] / (1 + IVA)

    df["Resultado Venta-Costo $"] = 0.0
    df.loc[df["Con Costo"], "Resultado Venta-Costo $"] = (
        df.loc[df["Con Costo"], "Venta sin IVA"] - df.loc[df["Con Costo"], "Costo Real"]
    )

    df["Rentabilidad Real"] = 0.0
    cc = df["Costo Real"] != 0
    df.loc[cc, "Rentabilidad Real"] = (
        df.loc[cc, "Resultado Venta-Costo $"] / df.loc[cc, "Costo Real"] * 100
    )

    df["Dia"] = df["Fecha"].dt.normalize()

    return df


def preparar_datos(df):
    df = preparar_columnas_base(df)
    df = agregar_clave(df)
    return df


# ============================================================
# MÉTRICAS (todas ponderadas por importes)
# ============================================================

def rentabilidad_ponderada(df):
    """Ganancia / costo real, ponderada por importes."""
    costo = df["Costo Real"].sum()
    if costo == 0:
        return 0.0
    return df["Resultado Venta-Costo $"].sum() / costo * 100


def margen_sobre_venta(df):
    venta = df.loc[df["Con Costo"], "Venta sin IVA"].sum()
    if venta == 0:
        return 0.0
    return df["Resultado Venta-Costo $"].sum() / venta * 100


def resumen_por(df, columnas):
    """Agrupa y calcula ventas, devoluciones y neto de forma consistente."""
    ventas = df[~df["Es Devolucion"]]
    devol = df[df["Es Devolucion"]]

    g_total = df.groupby(columnas, dropna=False)
    base = g_total.agg(
        Venta_Neta=("Total Venta", "sum"),
        Costo_Neto=("Costo Real", "sum"),
        Resultado=("Resultado Venta-Costo $", "sum"),
        Unidades=("Cantidad Neta", "sum"),
        Renglones=("Cantidad", "size"),
    )

    ventas_brutas = (
        ventas.groupby(columnas, dropna=False)["Total Venta"].sum().rename("Ventas_Brutas")
    )
    devoluciones = (
        devol.groupby(columnas, dropna=False)["Total Venta"].sum().rename("Devoluciones")
    )

    out = base.join(ventas_brutas, how="left").join(devoluciones, how="left")
    out[["Ventas_Brutas", "Devoluciones"]] = out[["Ventas_Brutas", "Devoluciones"]].fillna(0.0)

    out["Rentabilidad"] = 0.0
    con_costo = out["Costo_Neto"] != 0
    out.loc[con_costo, "Rentabilidad"] = (
        out.loc[con_costo, "Resultado"] / out.loc[con_costo, "Costo_Neto"] * 100
    )

    return out.reset_index()


# ============================================================
# GITHUB - HISTORIAL
# ============================================================

def github_configurado():
    try:
        return all(
            clave in st.secrets and str(st.secrets[clave]).strip() != ""
            for clave in ("GITHUB_TOKEN", "GITHUB_REPO", "GITHUB_USER")
        )
    except Exception:
        return False


def github_headers():
    return {
        "Authorization": f"Bearer {st.secrets['GITHUB_TOKEN']}",
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


@st.cache_data(ttl=600, show_spinner=False)
def _leer_historial_github(_firma_config):
    """Se cachea 10 minutos para no consultar GitHub en cada click."""
    respuesta = requests.get(github_url(), headers=github_headers(), timeout=20)

    if respuesta.status_code == 404:
        return pd.DataFrame()

    respuesta.raise_for_status()
    datos = respuesta.json()

    if "content" not in datos:
        # Archivos de más de 1 MB: GitHub no devuelve el contenido inline
        descarga = requests.get(
            datos["download_url"], headers=github_headers(), timeout=30
        )
        descarga.raise_for_status()
        contenido = descarga.content
    else:
        contenido = base64.b64decode(datos["content"])

    return pd.read_csv(io.BytesIO(contenido), dtype=str)


def cargar_historial():
    if not github_configurado():
        return pd.DataFrame()

    try:
        crudo = _leer_historial_github(st.secrets["GITHUB_REPO"])

        if crudo.empty:
            return pd.DataFrame()

        df = preparar_columnas_base(crudo)
        # Se recalcula la clave sobre todo el histórico: así los datos
        # guardados con versiones anteriores también quedan deduplicados.
        return agregar_clave(df)

    except Exception as e:
        st.error(f"No se pudo leer el historial de GitHub. Detalle: {e}")
        return pd.DataFrame()


def guardar_historial(df):
    if not github_configurado():
        st.error("GitHub no está configurado en los Secrets de Streamlit.")
        return False

    try:
        df_guardar = df[[c for c in COLUMNAS_GUARDAR if c in df.columns]].copy()
        df_guardar = df_guardar.sort_values(
            "Fecha", ascending=True, na_position="last", kind="stable"
        )

        contenido_csv = df_guardar.to_csv(index=False).encode("utf-8")
        contenido_base64 = base64.b64encode(contenido_csv).decode("utf-8")

        respuesta = requests.get(github_url(), headers=github_headers(), timeout=20)

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
            github_url(), headers=github_headers(), json=datos, timeout=60
        )
        respuesta.raise_for_status()

        _leer_historial_github.clear()
        return True

    except Exception as e:
        st.error(f"No se pudo guardar el historial en GitHub. Detalle: {e}")
        return False


# ============================================================
# GRÁFICOS (Plotly, tamaño fijo)
# ============================================================

def _layout_base(fig, titulo, alto):
    fig.update_layout(
        title=dict(text=f"<b>{titulo}</b>", x=0, xanchor="left", font=dict(size=17)),
        height=alto,
        autosize=True,
        template="plotly_white",
        separators=",.",  # decimales con coma, miles con punto
        font=dict(family="Inter, Segoe UI, Arial, sans-serif", size=13, color="#0F172A"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=10, r=30, t=60, b=40),
        dragmode=False,
        hoverlabel=dict(bgcolor="white", font_size=13),
        legend=dict(orientation="h", y=-0.15, x=0),
    )
    return fig


def etiqueta_articulo(cod, descripcion, ancho=48):
    """Código en negrita y descripción en varias líneas HORIZONTALES."""
    lineas = textwrap.wrap(
        str(descripcion), width=ancho, max_lines=3, placeholder="…"
    ) or ["-"]
    return f"<b>{cod}</b><br>" + "<br>".join(lineas)


def etiqueta_texto(texto, ancho=40):
    lineas = textwrap.wrap(str(texto), width=ancho, max_lines=2, placeholder="…") or ["-"]
    return "<br>".join(lineas)


def grafico_barras_horizontal(etiquetas, valores, titulo, formato, color=AZUL, alto_por_barra=58, colores=None):
    valores = list(valores)
    textos = [formato(v) for v in valores]
    n = len(valores)
    maximo = max([abs(v) for v in valores] + [1])
    minimo = min(valores + [0])

    if colores is None:
        colores = [color if v >= 0 else ROJO for v in valores]

    fig = go.Figure(
        go.Bar(
            x=valores,
            y=list(etiquetas),
            orientation="h",
            text=textos,
            textposition="outside",
            cliponaxis=False,
            marker=dict(color=colores, line=dict(width=0)),
            hovertemplate="%{y}<br><b>%{text}</b><extra></extra>",
        )
    )

    fig = _layout_base(fig, titulo, max(380, alto_por_barra * n + 120))
    fig.update_layout(bargap=0.35, showlegend=False)
    fig.update_xaxes(
        range=[minimo * 1.3 if minimo < 0 else 0, maximo * 1.3],
        tickformat=",.0f",
        showgrid=True,
        gridcolor="#E2E8F0",
        zeroline=True,
        zerolinecolor="#94A3B8",
        fixedrange=True,
    )
    fig.update_yaxes(
        autorange="reversed",  # el puesto 1 arriba
        automargin=True,
        fixedrange=True,
        tickfont=dict(size=12),
    )
    return fig


def grafico_dona(etiquetas, valores, titulo, colores=None):
    fig = go.Figure(
        go.Pie(
            labels=list(etiquetas),
            values=list(valores),
            hole=0.55,
            sort=False,
            marker=dict(colors=colores or PALETA, line=dict(color="white", width=2)),
            textinfo="percent",
            textfont=dict(size=14),
            hovertemplate="<b>%{label}</b><br>%{value:,.2f}<br>%{percent}<extra></extra>",
        )
    )
    fig = _layout_base(fig, titulo, 400)
    fig.update_layout(legend=dict(orientation="h", y=-0.05, x=0.5, xanchor="center"))
    return fig


def mostrar_grafico(fig, clave):
    st.plotly_chart(
        fig,
        use_container_width=True,
        config=CONFIG_GRAFICO,
        key=clave,
    )


# ============================================================
# TÍTULO
# ============================================================

st.title("📊 Panel de Ventas")
st.caption("INFOVENTAS — análisis diario e histórico")

# ============================================================
# CARGAR HISTORIAL
# ============================================================

historial = cargar_historial()

# ============================================================
# SUBIR ARCHIVO
# ============================================================

st.subheader("📁 Cargar INFOVENTAS")

archivo = st.file_uploader(
    "Subí el INFOVENTAS del día",
    type=["xlsx", "xls", "xlsm", "csv", "txt"],
)

if archivo is not None:
    try:
        df_nuevo = preparar_datos(leer_infoventas(archivo))

        st.success(f"Archivo leído correctamente: {len(df_nuevo):,} filas.".replace(",", "."))

        fechas_archivo = sorted(df_nuevo["Fecha"].dt.date.unique())
        st.info(
            "Fecha/s detectada/s: "
            + ", ".join(f.strftime("%d/%m/%Y") for f in fechas_archivo)
        )

        n_dev = int(enriquecer(df_nuevo)["Es Devolucion"].sum())
        if n_dev:
            st.info(
                f"🔁 El archivo tiene {n_dev} renglones de notas de crédito / "
                "devoluciones. Se restan de las ventas."
            )

        if not github_configurado():
            st.warning(
                "⚠️ GitHub no está configurado en los Secrets. "
                "El panel funciona, pero todavía no guardará el histórico."
            )
        else:
            if historial.empty:
                df_para_agregar = df_nuevo.copy()
            else:
                existentes = set(historial["_Clave"])
                df_para_agregar = df_nuevo[~df_nuevo["_Clave"].isin(existentes)].copy()

            if not df_para_agregar.empty:
                with st.spinner("Guardando automáticamente en el histórico..."):
                    historial_actualizado = pd.concat(
                        [historial, df_para_agregar], ignore_index=True
                    ).drop_duplicates(subset=["_Clave"], keep="first")

                    if guardar_historial(historial_actualizado):
                        historial = historial_actualizado
                        st.success(
                            "✅ Guardado automático: "
                            f"{len(df_para_agregar):,} filas nuevas.".replace(",", ".")
                        )
            else:
                st.info(
                    "ℹ️ Este archivo ya está cargado en el histórico. "
                    "No se agregaron duplicados."
                )

    except Exception as e:
        st.error(f"❌ Error procesando el archivo: {e}")
        st.stop()

# ============================================================
# SIN HISTÓRICO
# ============================================================

if historial.empty:
    st.info("Subí un INFOVENTAS para comenzar.")
    st.stop()

df = enriquecer(historial)

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
    max_value=fecha_max,
    format="DD/MM/YYYY",
)

if isinstance(rango_fechas, (tuple, list)):
    if len(rango_fechas) == 2:
        desde, hasta = rango_fechas
    elif len(rango_fechas) == 1:
        desde = hasta = rango_fechas[0]
    else:
        desde, hasta = fecha_min, fecha_max
else:
    desde = hasta = rango_fechas

df = df[
    (df["Fecha"] >= pd.Timestamp(desde))
    & (df["Fecha"] < pd.Timestamp(hasta) + pd.Timedelta(days=1))
]


def opciones(columna):
    return sorted(
        df[columna]
        .fillna("")
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .unique()
    )


tipo_operacion = st.sidebar.multiselect(
    "Tipo de operación",
    ["Venta", "Devolución / Nota de crédito"],
)
if tipo_operacion:
    df = df[df["Tipo Operacion"].isin(tipo_operacion)]

cliente = st.sidebar.multiselect("Cliente", opciones("Cliente"))
if cliente:
    df = df[df["Cliente"].isin(cliente)]

tipo_pago = st.sidebar.multiselect("Tipo de pago", opciones("Tipo Pago"))
if tipo_pago:
    df = df[df["Tipo Pago"].isin(tipo_pago)]

forma_pago = st.sidebar.multiselect("Forma de pago exacta", opciones("Forma Pago"))
if forma_pago:
    df = df[df["Forma Pago"].isin(forma_pago)]

articulo = st.sidebar.multiselect("Artículo", opciones("Articulo"))
if articulo:
    df = df[df["Articulo"].isin(articulo)]

if df.empty:
    st.warning("⚠️ No hay operaciones que coincidan con los filtros.")
    st.stop()

# ============================================================
# KPIs
# ============================================================

ventas_df = df[~df["Es Devolucion"]]
devol_df = df[df["Es Devolucion"]]

ventas_brutas = ventas_df["Total Venta"].sum()
devoluciones = devol_df["Total Venta"].sum()  # negativo
venta_neta = df["Total Venta"].sum()

unidades_netas = df["Cantidad Neta"].sum()
unidades_vendidas = ventas_df["Cantidad Neta"].sum()
unidades_devueltas = devol_df["Cantidad Neta"].sum()  # negativo

comprobantes_venta = ventas_df.loc[ventas_df["Comprobante"] != "", "Comprobante"].nunique()
comprobantes_nc = devol_df.loc[devol_df["Comprobante"] != "", "Comprobante"].nunique()
if comprobantes_venta == 0:
    comprobantes_venta = len(ventas_df)

clientes_activos = ventas_df.loc[ventas_df["Cliente"] != "", "Cliente"].nunique()
articulos_unicos = df.loc[df["Cod"] != "", "Cod"].nunique()

precio_promedio = ventas_brutas / unidades_vendidas if unidades_vendidas else 0
ticket_promedio = ventas_brutas / comprobantes_venta if comprobantes_venta else 0

resultado_total = df["Resultado Venta-Costo $"].sum()
rent_ponderada = rentabilidad_ponderada(df)
margen_venta = margen_sobre_venta(df)

# ============================================================
# RESUMEN SUPERIOR
# ============================================================

st.subheader("📌 Resumen")

c1, c2, c3, c4 = st.columns(4)
c1.metric("💰 Venta neta", formato_pesos(venta_neta))
c2.metric("🧾 Ventas brutas", formato_pesos(ventas_brutas))
c3.metric("🔁 Devoluciones / NC", formato_pesos(devoluciones))
c4.metric("💵 Ganancia estimada (sin IVA)", formato_pesos(resultado_total))

c5, c6, c7, c8 = st.columns(4)
c5.metric("📦 Unidades netas", formato_entero(unidades_netas))
c6.metric("🛒 Ticket promedio", formato_pesos(ticket_promedio))
c7.metric("📈 Rentabilidad s/costo", formato_porcentaje(rent_ponderada))
c8.metric("👥 Clientes con compras", formato_entero(clientes_activos))

sin_costo = int((~df["Con Costo"]).sum())

st.caption(
    "Venta neta = ventas brutas menos notas de crédito / devoluciones. "
    "Ganancia estimada = venta sin IVA − (Precio Costo × cantidad). "
    "Rentabilidad s/costo = ganancia / costo, ponderada por importes. "
    f"{sin_costo} renglones (notas de crédito, artículos 'TODO', etc.) vienen "
    "sin costo en INFOVENTAS y no entran en la ganancia."
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
    st.subheader("📊 Principales indicadores")

    resumen = pd.DataFrame(
        {
            "Indicador": [
                "Ventas brutas",
                "Devoluciones / notas de crédito",
                "Venta neta",
                "Unidades vendidas",
                "Unidades devueltas",
                "Unidades netas",
                "Comprobantes de venta",
                "Notas de crédito",
                "Clientes con compras",
                "Artículos distintos",
                "Precio promedio por unidad",
                "Ticket promedio",
                "Ganancia estimada (sin IVA)",
                "Rentabilidad s/costo",
                "Margen s/venta (sin IVA)",
            ],
            "Valor": [
                formato_pesos(ventas_brutas),
                formato_pesos(devoluciones),
                formato_pesos(venta_neta),
                formato_entero(unidades_vendidas),
                formato_entero(unidades_devueltas),
                formato_entero(unidades_netas),
                formato_entero(comprobantes_venta),
                formato_entero(comprobantes_nc),
                formato_entero(clientes_activos),
                formato_entero(articulos_unicos),
                formato_pesos(precio_promedio),
                formato_pesos(ticket_promedio),
                formato_pesos(resultado_total),
                formato_porcentaje(rent_ponderada),
                formato_porcentaje(margen_venta),
            ],
        }
    )

    izq, der = st.columns([1, 1])

    with izq:
        st.dataframe(resumen, use_container_width=True, hide_index=True, height=565)

    with der:
        fig = go.Figure(
            go.Waterfall(
                x=["Ventas brutas", "Devoluciones / NC", "Venta neta"],
                y=[ventas_brutas, devoluciones, 0],
                measure=["absolute", "relative", "total"],
                text=[
                    formato_pesos(ventas_brutas),
                    formato_pesos(devoluciones),
                    formato_pesos(venta_neta),
                ],
                textposition="outside",
                increasing=dict(marker=dict(color=VERDE)),
                decreasing=dict(marker=dict(color=ROJO)),
                totals=dict(marker=dict(color=AZUL)),
                connector=dict(line=dict(color="#94A3B8")),
                hovertemplate="%{x}<br><b>%{text}</b><extra></extra>",
            )
        )
        fig = _layout_base(fig, "De ventas brutas a venta neta", 565)
        fig.update_layout(showlegend=False)
        fig.update_xaxes(fixedrange=True)
        fig.update_yaxes(
            fixedrange=True,
            tickformat=",.0f",
            showgrid=True,
            gridcolor="#E2E8F0",
            range=[0, ventas_brutas * 1.2 if ventas_brutas > 0 else 1],
        )
        mostrar_grafico(fig, "grafico_cascada")

# ============================================================
# ARTÍCULOS
# ============================================================

with tab_articulos:
    st.subheader("🔧 Ranking de artículos")

    ranking_articulos = (
        resumen_por(df, ["Cod", "Articulo"])
        .sort_values("Venta_Neta", ascending=False)
        .reset_index(drop=True)
    )
    ranking_articulos.insert(0, "Puesto", range(1, len(ranking_articulos) + 1))

    vista = pd.DataFrame(
        {
            "Puesto": ranking_articulos["Puesto"],
            "Código": ranking_articulos["Cod"],
            "Artículo": ranking_articulos["Articulo"],
            "Unidades netas": ranking_articulos["Unidades"].apply(formato_entero),
            "Venta neta": ranking_articulos["Venta_Neta"].apply(formato_pesos),
            "Devoluciones": ranking_articulos["Devoluciones"].apply(formato_pesos),
            "Ganancia $": ranking_articulos["Resultado"].apply(formato_pesos),
            "Rentabilidad": ranking_articulos["Rentabilidad"].apply(formato_porcentaje),
            "Renglones": ranking_articulos["Renglones"],
        }
    )

    st.dataframe(
        vista.head(100),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Artículo": st.column_config.TextColumn("Artículo", width="large"),
        },
    )

    top_venta = ranking_articulos[ranking_articulos["Venta_Neta"] > 0].head(15)
    if not top_venta.empty:
        fig = grafico_barras_horizontal(
            [etiqueta_articulo(c, a) for c, a in zip(top_venta["Cod"], top_venta["Articulo"])],
            top_venta["Venta_Neta"],
            "Artículos con mayor venta neta",
            formato_pesos,
            color=AZUL,
            alto_por_barra=66,
        )
        mostrar_grafico(fig, "grafico_articulos_venta")

    top_unidades = (
        ranking_articulos[ranking_articulos["Unidades"] > 0]
        .sort_values("Unidades", ascending=False)
        .head(15)
    )
    if not top_unidades.empty:
        fig = grafico_barras_horizontal(
            [etiqueta_articulo(c, a) for c, a in zip(top_unidades["Cod"], top_unidades["Articulo"])],
            top_unidades["Unidades"],
            "Artículos más vendidos (unidades netas)",
            formato_entero,
            color=VERDE,
            alto_por_barra=66,
        )
        mostrar_grafico(fig, "grafico_articulos_unidades")

# ============================================================
# CLIENTES
# ============================================================

with tab_clientes:
    st.subheader("👥 Ranking de compradores")

    ranking_clientes = (
        resumen_por(df, ["Cliente"])
        .sort_values("Venta_Neta", ascending=False)
        .reset_index(drop=True)
    )
    ranking_clientes.insert(0, "Puesto", range(1, len(ranking_clientes) + 1))

    def estado_cliente(fila):
        if fila["Devoluciones"] < 0 and fila["Venta_Neta"] <= 0:
            return "Saldo cancelado por NC / devoluciones"
        if fila["Devoluciones"] < 0:
            return "Con devoluciones"
        return "Normal"

    ranking_clientes["Estado"] = ranking_clientes.apply(estado_cliente, axis=1)

    vista = pd.DataFrame(
        {
            "Puesto": ranking_clientes["Puesto"],
            "Cliente": ranking_clientes["Cliente"],
            "Ventas brutas": ranking_clientes["Ventas_Brutas"].apply(formato_pesos),
            "Devoluciones": ranking_clientes["Devoluciones"].apply(formato_pesos),
            "Venta neta": ranking_clientes["Venta_Neta"].apply(formato_pesos),
            "Unidades netas": ranking_clientes["Unidades"].apply(formato_entero),
            "Ganancia $": ranking_clientes["Resultado"].apply(formato_pesos),
            "Rentabilidad": ranking_clientes["Rentabilidad"].apply(formato_porcentaje),
            "Estado": ranking_clientes["Estado"],
        }
    )

    st.dataframe(vista.head(100), use_container_width=True, hide_index=True)

    top_clientes = ranking_clientes[ranking_clientes["Venta_Neta"] > 0].head(15)
    if not top_clientes.empty:
        fig = grafico_barras_horizontal(
            [etiqueta_texto(c) for c in top_clientes["Cliente"]],
            top_clientes["Venta_Neta"],
            "Compradores con mayor venta neta",
            formato_pesos,
            color=AZUL,
        )
        mostrar_grafico(fig, "grafico_clientes")

    anulados = ranking_clientes[ranking_clientes["Estado"] != "Normal"]
    if not anulados.empty:
        st.caption(
            "Los clientes con venta neta en cero o negativa (por notas de crédito "
            "o saldos cancelados) no entran en el gráfico, pero sí figuran en la "
            "tabla con su estado."
        )

# ============================================================
# PAGOS
# ============================================================

with tab_pagos:
    st.subheader("💳 Contado vs cuenta corriente")

    ranking_tipo_pago = (
        resumen_por(df, ["Tipo Pago"])
        .sort_values("Venta_Neta", ascending=False)
        .reset_index(drop=True)
    )

    vista = pd.DataFrame(
        {
            "Tipo de pago": ranking_tipo_pago["Tipo Pago"],
            "Venta neta": ranking_tipo_pago["Venta_Neta"].apply(formato_pesos),
            "Devoluciones": ranking_tipo_pago["Devoluciones"].apply(formato_pesos),
            "Unidades netas": ranking_tipo_pago["Unidades"].apply(formato_entero),
            "Renglones": ranking_tipo_pago["Renglones"],
        }
    )
    st.dataframe(vista, use_container_width=True, hide_index=True)

    col_a, col_b = st.columns(2)

    positivos = ranking_tipo_pago[ranking_tipo_pago["Venta_Neta"] > 0]

    with col_a:
        if not positivos.empty:
            fig = grafico_dona(
                positivos["Tipo Pago"],
                positivos["Venta_Neta"],
                "Distribución de la venta neta",
                colores=[color_pago(t, "tipo") for t in positivos["Tipo Pago"]],
            )
            mostrar_grafico(fig, "grafico_dona_pago")
        else:
            st.info("No hay venta neta positiva para mostrar.")

    with col_b:
        fig = grafico_barras_horizontal(
            list(ranking_tipo_pago["Tipo Pago"]),
            ranking_tipo_pago["Venta_Neta"],
            "Venta neta por tipo de pago",
            formato_pesos,
            color=AZUL,
            alto_por_barra=90,
            colores=[color_pago(t, "tipo") for t in ranking_tipo_pago["Tipo Pago"]],
        )
        fig.update_layout(height=400)
        mostrar_grafico(fig, "grafico_barras_pago")

    st.subheader("💳 Formas de pago exactas")

    ranking_forma = (
        resumen_por(df, ["Forma Pago"])
        .sort_values("Venta_Neta", ascending=False)
        .reset_index(drop=True)
    )

    st.dataframe(
        pd.DataFrame(
            {
                "Forma de pago": ranking_forma["Forma Pago"],
                "Venta neta": ranking_forma["Venta_Neta"].apply(formato_pesos),
                "Unidades netas": ranking_forma["Unidades"].apply(formato_entero),
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

    positivas = ranking_forma[ranking_forma["Venta_Neta"] > 0]
    if not positivas.empty:
        col_c, col_d = st.columns(2)

        with col_c:
            fig = grafico_dona(
                positivas["Forma Pago"],
                positivas["Venta_Neta"],
                "Distribución por forma de pago",
                colores=[color_pago(f) for f in positivas["Forma Pago"]],
            )
            mostrar_grafico(fig, "grafico_dona_forma_pago")

        with col_d:
            fig = grafico_barras_horizontal(
                [etiqueta_texto(f) for f in positivas["Forma Pago"]],
                positivas["Venta_Neta"],
                "Venta neta por forma de pago",
                formato_pesos,
                alto_por_barra=70,
                colores=[color_pago(f) for f in positivas["Forma Pago"]],
            )
            fig.update_layout(height=400)
            mostrar_grafico(fig, "grafico_forma_pago")

# ============================================================
# EVOLUCIÓN
# ============================================================

with tab_evolucion:
    st.subheader("📅 Evolución diaria")

    ventas_dia = (
        resumen_por(df, ["Dia"]).sort_values("Dia").reset_index(drop=True)
    )

    if len(ventas_dia) == 1:
        st.info(
            "Hay un solo día en el período seleccionado. "
            "Cargá más días para ver la evolución."
        )

    x_fechas = ventas_dia["Dia"]

    fig = go.Figure(
        go.Scatter(
            x=x_fechas,
            y=ventas_dia["Venta_Neta"],
            mode="lines+markers",
            line=dict(color=AZUL, width=3),
            marker=dict(size=9, color=AZUL, line=dict(color="white", width=2)),
            fill="tozeroy",
            fillcolor="rgba(37, 99, 235, 0.10)",
            customdata=[formato_pesos(v) for v in ventas_dia["Venta_Neta"]],
            hovertemplate="%{x|%d/%m/%Y}<br><b>%{customdata}</b><extra></extra>",
        )
    )
    fig = _layout_base(fig, "Venta neta por día", 420)
    fig.update_layout(showlegend=False)
    fig.update_xaxes(
        tickformat="%d/%m/%Y",
        type="date",
        fixedrange=True,
        showgrid=False,
    )
    fig.update_yaxes(
        tickformat=",.0f",
        fixedrange=True,
        showgrid=True,
        gridcolor="#E2E8F0",
        rangemode="tozero",
    )
    mostrar_grafico(fig, "grafico_evolucion_venta")

    fig = go.Figure(
        go.Bar(
            x=x_fechas,
            y=ventas_dia["Unidades"],
            marker=dict(color=VERDE),
            customdata=[formato_entero(v) for v in ventas_dia["Unidades"]],
            hovertemplate="%{x|%d/%m/%Y}<br><b>%{customdata} unidades</b><extra></extra>",
        )
    )
    fig = _layout_base(fig, "Unidades netas por día", 380)
    fig.update_layout(showlegend=False, bargap=0.4)
    fig.update_xaxes(tickformat="%d/%m/%Y", type="date", fixedrange=True)
    fig.update_yaxes(
        tickformat=",.0f", fixedrange=True, showgrid=True, gridcolor="#E2E8F0"
    )
    mostrar_grafico(fig, "grafico_evolucion_unidades")

    por_pago = (
        resumen_por(df, ["Dia", "Forma Pago"]).sort_values(["Dia", "Forma Pago"])
    )
    fig = go.Figure()
    for tipo in sorted(por_pago["Forma Pago"].unique()):
        sub = por_pago[por_pago["Forma Pago"] == tipo]
        fig.add_trace(
            go.Bar(
                x=sub["Dia"],
                y=sub["Venta_Neta"],
                name=tipo,
                marker=dict(color=color_pago(tipo)),
                customdata=[formato_pesos(v) for v in sub["Venta_Neta"]],
                hovertemplate="%{x|%d/%m/%Y}<br>" + tipo + ": <b>%{customdata}</b><extra></extra>",
            )
        )
    fig = _layout_base(fig, "Venta neta por día según forma de pago", 420)
    fig.update_layout(barmode="stack", bargap=0.4)
    fig.update_xaxes(tickformat="%d/%m/%Y", type="date", fixedrange=True)
    fig.update_yaxes(
        tickformat=",.0f", fixedrange=True, showgrid=True, gridcolor="#E2E8F0"
    )
    mostrar_grafico(fig, "grafico_evolucion_pago")

    tabla_dia = ventas_dia.sort_values("Dia", ascending=False)
    st.dataframe(
        pd.DataFrame(
            {
                "Fecha": tabla_dia["Dia"].dt.strftime("%d/%m/%Y"),
                "Día": tabla_dia["Dia"].dt.dayofweek.map(DIAS_SEMANA),
                "Ventas brutas": tabla_dia["Ventas_Brutas"].apply(formato_pesos),
                "Devoluciones": tabla_dia["Devoluciones"].apply(formato_pesos),
                "Venta neta": tabla_dia["Venta_Neta"].apply(formato_pesos),
                "Unidades netas": tabla_dia["Unidades"].apply(formato_entero),
                "Ganancia $": tabla_dia["Resultado"].apply(formato_pesos),
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

# ============================================================
# DETALLE
# ============================================================

with tab_detalle:
    st.subheader("📋 Detalle de operaciones")

    detalle = df.sort_values(["Fecha", "Comprobante"], ascending=[False, False]).reset_index(drop=True)

    tabla = pd.DataFrame(
        {
            "Nº": range(1, len(detalle) + 1),
            "Fecha": formato_fecha(detalle["Fecha"]),
            "Comprobante": detalle["Comprobante"],
            "Tipo": detalle["Tipo Operacion"],
            "Código": detalle["Cod"],
            "Artículo": detalle["Articulo"],
            "Cantidad": detalle["Cantidad Neta"].apply(formato_entero),
            "Cliente": detalle["Cliente"],
            "Precio unitario": detalle["Precio Unitario"].apply(formato_pesos),
            "Total venta": detalle["Total Venta"].apply(formato_pesos),
            "Forma de pago": detalle["Forma Pago"],
            "Tipo de pago": detalle["Tipo Pago"],
            "Rentabilidad": detalle["Rentabilidad Real"].apply(formato_porcentaje),
            "Ganancia $": detalle["Resultado Venta-Costo $"].apply(formato_pesos),
        }
    )

    st.dataframe(
        tabla,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Artículo": st.column_config.TextColumn("Artículo", width="large"),
        },
    )

    # Descarga lista para abrir en Excel argentino (separador ; y coma decimal)
    export = pd.DataFrame(
        {
            "Fecha": formato_fecha(detalle["Fecha"]),
            "Comprobante": detalle["Comprobante"],
            "Tipo": detalle["Tipo Operacion"],
            "Código": detalle["Cod"],
            "Artículo": detalle["Articulo"],
            "Cantidad": detalle["Cantidad Neta"],
            "Cliente": detalle["Cliente"],
            "Precio unitario": detalle["Precio Unitario"],
            "Costo real": detalle["Costo Real"],
            "Total venta": detalle["Total Venta"],
            "Forma de pago": detalle["Forma Pago"],
            "Tipo de pago": detalle["Tipo Pago"],
            "Rentabilidad %": detalle["Rentabilidad Real"].round(2),
            "Ganancia $": detalle["Resultado Venta-Costo $"],
        }
    )

    csv = export.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")

    st.download_button(
        "📥 Descargar datos filtrados",
        data=csv,
        file_name="ventas_filtradas.csv",
        mime="text/csv",
    )
