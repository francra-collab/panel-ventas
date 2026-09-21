import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Panel de Ventas",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Panel de Ventas")

st.write("Subí tu archivo INFOVENTAS para comenzar.")

archivo = st.file_uploader(
    "📁 Subir INFOVENTAS",
    type=["xlsx", "xls"]
)

if archivo is not None:

    df = pd.read_excel(archivo)

    st.success("Excel cargado correctamente.")

    st.write("Cantidad de registros:", len(df))

    st.dataframe(
        df,
        use_container_width=True
    )