import streamlit as st
import pandas as pd

# ==============================
# CONFIG
# ==============================
st.set_page_config(page_title="Packing Summary", layout="wide")

st.title("📦 Packing Summary by Model & Type")

# ==============================
# UPLOAD FILE
# ==============================
uploaded_file = st.file_uploader("📥 Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:

    # ==============================
    # READ FILE
    # ==============================
    df = pd.read_excel(uploaded_file)

    # ==============================
    # CLEAN COLUMNS (🔥 IMPORTANT)
    # ==============================
    df.columns = (
        df.columns
        .str.strip()
        .str.replace("\n", " ", regex=True)
        .str.replace(r"\s+", " ", regex=True)
    )

    st.write("✅ Colonnes détectées :", df.columns.tolist())

    # ==============================
    # CHECK REQUIRED COLUMNS
    # ==============================
    if "Model" not in df.columns or "TYPE" not in df.columns:
        st.error("❌ Les colonnes 'Model' et 'TYPE' sont obligatoires")
        st.stop()

    # ==============================
    # AUTO DETECT COLUMNS 🔍
    # ==============================
    col_ctn = None
    col_nw = None
    col_gw = None
    col_vol = None

    for col in df.columns:
        col_upper = col.upper()

        if "CTN" in col_upper:
            col_ctn = col
        elif "N W" in col_upper or "NET" in col_upper:
            col_nw = col
        elif "G W" in col_upper or "GROSS" in col_upper:
            col_gw = col
        elif "VOLUME" in col_upper or "CBM" in col_upper:
            col_vol = col

    # Vérification
    if not all([col_ctn, col_nw, col_gw, col_vol]):
        st.error("❌ Impossible de détecter toutes les colonnes nécessaires")
        st.write("Colonnes trouvées :", df.columns.tolist())
        st.stop()

    # Affichage mapping
    st.write("🔎 Mapping utilisé :")
    st.write({
        "CTN QTY": col_ctn,
        "TOTAL N W (KG)": col_nw,
        "TOTAL G W (KG)": col_gw,
        "TOTAL VOLUME (CBM)": col_vol
    })

    # ==============================
    # CONVERT TO NUMERIC 🔢
    # ==============================
    for col in [col_ctn, col_nw, col_gw, col_vol]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # ==============================
    # FILTER (OPTIONAL)
    # ==============================
    models = st.multiselect(
        "🔎 Filter by Model",
        options=df["Model"].unique(),
        default=df["Model"].unique()
    )

    df_filtered = df[df["Model"].isin(models)]

    # ==============================
    # GROUPBY + SUM
    # ==============================
    result = df_filtered.groupby(
        ["Model", "TYPE"], as_index=False
    ).agg({
        col_ctn: "sum",
        col_nw: "sum",
        col_gw: "sum",
        col_vol: "sum"
    })

    # Rename clean columns
    result.columns = [
        "Model", "TYPE",
        "CTN QTY",
        "TOTAL N W (KG)",
        "TOTAL G W (KG)",
        "TOTAL VOLUME (CBM)"
    ]

    # ==============================
    # DISPLAY
    # ==============================
    st.subheader("📊 Summary Result")
    st.dataframe(result, use_container_width=True)

    # ==============================
    # KPI TOTAL 🔥
    # ==============================
    st.subheader("📈 Global Totals")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("📦 Total CTN", int(result["CTN QTY"].sum()))
    col2.metric("⚖️ Net Weight", round(result["TOTAL N W (KG)"].sum(), 2))
    col3.metric("⚖️ Gross Weight", round(result["TOTAL G W (KG)"].sum(), 2))
    col4.metric("📐 Volume", round(result["TOTAL VOLUME (CBM)"].sum(), 3))

    # ==============================
    # DOWNLOAD
    # ==============================
    csv = result.to_csv(index=False).encode("utf-8")

    st.download_button(
        label="📥 Download Result",
        data=csv,
        file_name="packing_summary.csv",
        mime="text/csv"
    )

else:
    st.info("👆 Upload your Excel file to start")
