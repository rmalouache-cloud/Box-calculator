import streamlit as st
import pandas as pd
from io import BytesIO

# ==============================
# CONFIG
# ==============================
st.set_page_config(page_title="Packing Summary", layout="wide")

st.title("📦 Packing Summary by Model & Type")

# ==============================
# UPLOAD
# ==============================
uploaded_file = st.file_uploader("📥 Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:

    df = pd.read_excel(uploaded_file)

    # Nettoyage colonnes
    df.columns = (
        df.columns
        .str.strip()
        .str.replace("\n", " ", regex=True)
        .str.replace(r"\s+", " ", regex=True)
    )

    # Détection colonnes
    col_ctn = [c for c in df.columns if "CTN" in c.upper()][0]
    col_nw = [c for c in df.columns if "N W" in c.upper() or "NET" in c.upper()][0]
    col_gw = [c for c in df.columns if "G W" in c.upper() or "GROSS" in c.upper()][0]
    col_vol = [c for c in df.columns if "VOLUME" in c.upper() or "CBM" in c.upper()][0]

    # Convert numeric
    for col in [col_ctn, col_nw, col_gw, col_vol]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # ==============================
    # LOT INPUT
    # ==============================
    st.subheader("📥 Enter LOT Quantity per Model")

    models = df["Model"].unique()
    lot_qty_dict = {}

    cols = st.columns(len(models))

    for i, model in enumerate(models):
        lot_qty = cols[i].number_input(f"{model}", min_value=0, value=0)
        lot_qty_dict[model] = lot_qty

    # ==============================
    # GROUPBY
    # ==============================
    result = df.groupby(
        ["Model", "TYPE"], as_index=False
    ).agg({
        col_ctn: "sum",
        col_nw: "sum",
        col_gw: "sum",
        col_vol: "sum"
    })

    result.columns = [
        "Model", "TYPE",
        "CTN QTY",
        "TOTAL N W (KG)",
        "TOTAL G W (KG)",
        "TOTAL VOLUME (CBM)"
    ]

    # Ajouter LOT
    result["LOT QTY"] = result["Model"].map(lot_qty_dict)

    # ==============================
    # DISPLAY
    # ==============================
    st.subheader("📊 Summary Result")
    st.dataframe(result, use_container_width=True)

    # ==============================
    # DOWNLOAD CSV (FIX FR)
    # ==============================
    csv = result.to_csv(index=False, sep=";").encode("utf-8")

    st.download_button(
        label="📥 Download CSV (Excel FR compatible)",
        data=csv,
        file_name="packing_summary.csv",
        mime="text/csv"
    )

    # ==============================
    # DOWNLOAD EXCEL (PRO 🔥)
    # ==============================
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result.to_excel(writer, index=False, sheet_name='Summary')

    st.download_button(
        label="📥 Download Excel (Best)",
        data=output.getvalue(),
        file_name="packing_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👆 Upload your Excel file to start")
