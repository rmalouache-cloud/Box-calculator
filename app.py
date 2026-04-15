import streamlit as st
import pandas as pd

# ==============================
# CONFIG PAGE
# ==============================
st.set_page_config(page_title="Packing Summary", layout="wide")

st.title("📦 Packing Summary by Model & Type")

# ==============================
# UPLOAD FILE
# ==============================
uploaded_file = st.file_uploader("📥 Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:

    # ==============================
    # READ DATA
    # ==============================
    df = pd.read_excel(uploaded_file)

    # Nettoyage des colonnes
    df.columns = df.columns.str.strip()

    # ==============================
    # AFFICHER DATA ORIGINALE
    # ==============================
    st.subheader("📄 Original Data")
    st.dataframe(df, use_container_width=True)

    # ==============================
    # FILTRE PAR MODELE (OPTIONNEL)
    # ==============================
    models = st.multiselect(
        "🔎 Filter by Model",
        options=df["Model"].unique(),
        default=df["Model"].unique()
    )

    df_filtered = df[df["Model"].isin(models)]

    # ==============================
    # GROUPBY + SOMME
    # ==============================
    result = df_filtered.groupby(
        ["Model", "TYPE"], as_index=False
    ).agg({
        "CTN QTY": "sum",
        "TOTAL N W (KG)": "sum",
        "TOTAL G W (KG)": "sum",
        "TOTAL VOLUME (CBM)": "sum"
    })

    # ==============================
    # AFFICHER RESULTAT
    # ==============================
    st.subheader("📊 Summary Result")
    st.dataframe(result, use_container_width=True)

    # ==============================
    # KPI GLOBAL (BONUS 🔥)
    # ==============================
    st.subheader("📈 Global Totals")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("📦 Total CTN", int(result["CTN QTY"].sum()))
    col2.metric("⚖️ Total Net Weight", round(result["TOTAL N W (KG)"].sum(), 2))
    col3.metric("⚖️ Total Gross Weight", round(result["TOTAL G W (KG)"].sum(), 2))
    col4.metric("📐 Total Volume", round(result["TOTAL VOLUME (CBM)"].sum(), 3))

    # ==============================
    # DOWNLOAD RESULT
    # ==============================
    csv = result.to_csv(index=False).encode('utf-8')

    st.download_button(
        label="📥 Download Result (CSV)",
        data=csv,
        file_name="packing_summary.csv",
        mime="text/csv"
    )

else:
    st.info("👆 Upload an Excel file to start")