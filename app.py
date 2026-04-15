import streamlit as st
import pandas as pd
from io import BytesIO
from PIL import Image

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Packing Dashboard", layout="wide")

# =========================
# LOGOS + HEADER (PLUS GRANDS)
# =========================
container_logo = Image.open("conteneur_logo.png")
stream_logo = Image.open("stream_logo.png")

col1, col2, col3 = st.columns([1, 5, 1])

with col1:
    st.image(container_logo, width=200)

with col2:
    st.title("Container Filling Industrial Dashboard")
    st.caption("Packing Summary by Model & Type")

with col3:
    st.image(stream_logo, width=200)

st.markdown("---")

# =========================
# INPUT GLOBAL
# =========================
st.subheader("📝 Shipment Information")

apna = st.text_input("APNA")
order_shipment = st.text_input("Order of Shipment")

st.markdown("---")

# =========================
# UPLOAD FILE
# =========================
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

    # =========================
    # FILTER (OPTIONAL)
    # =========================
    models_all = df["Model"].dropna().unique()

    models_selected = st.multiselect(
        "🔎 Filter by Model",
        options=models_all,
        default=models_all
    )

    df_filtered = df[df["Model"].isin(models_selected)]

    # =========================
    # DETECT COLUMNS
    # =========================
    col_ctn = [c for c in df_filtered.columns if "CTN" in c.upper()][0]
    col_nw = [c for c in df_filtered.columns if "N W" in c.upper() or "NET" in c.upper()][0]
    col_gw = [c for c in df_filtered.columns if "G W" in c.upper() or "GROSS" in c.upper()][0]
    col_vol = [c for c in df_filtered.columns if "VOLUME" in c.upper() or "CBM" in c.upper()][0]

    # Convert numeric
    for col in [col_ctn, col_nw, col_gw, col_vol]:
        df_filtered[col] = pd.to_numeric(df_filtered[col], errors="coerce").fillna(0)

    # =========================
    # LOT INPUT
    # =========================
    st.subheader("📥 Enter LOT Quantity per Model")

    models = df_filtered["Model"].unique()
    lot_qty_dict = {}

    cols = st.columns(len(models) if len(models) > 0 else 1)

    for i, model in enumerate(models):
        lot_qty_dict[model] = cols[i].number_input(
            f"{model}",
            min_value=0,
            value=0,
            step=100
        )

    # =========================
    # GROUPBY
    # =========================
    result = df_filtered.groupby(
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

    result["LOT QTY"] = result["Model"].map(lot_qty_dict)

    # =========================
    # ADD INFO INTO FILE
    # =========================
    result["APNA"] = apna
    result["Order of Shipment"] = order_shipment

    # =========================
    # STYLE TABLE
    # =========================
    def style_table(df):
        return df.style.set_properties(**{
            'border': '1px solid black',
            'text-align': 'center'
        }).set_table_styles([
            {'selector': 'th',
             'props': [('border', '1px solid black'),
                       ('background-color', '#f2f2f2'),
                       ('text-align', 'center')]}
        ])

    st.subheader("📦 Summary Result (Styled)")
    st.dataframe(style_table(result), use_container_width=True)

    # =========================
    # KPI
    # =========================
    st.subheader("📈 Global Totals")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("📦 Total CTN", int(result["CTN QTY"].sum()))
    col2.metric("⚖️ Net Weight", round(result["TOTAL N W (KG)"].sum(), 2))
    col3.metric("⚖️ Gross Weight", round(result["TOTAL G W (KG)"].sum(), 2))
    col4.metric("📐 Volume", round(result["TOTAL VOLUME (CBM)"].sum(), 3))

    # =========================
    # FILE NAME CLEANING
    # =========================
    safe_apna = apna.replace(" ", "_") if apna else "NO_APNA"
    safe_order = order_shipment.replace(" ", "_") if order_shipment else "NO_ORDER"

    file_name = f"packing_summary_{safe_apna}_{safe_order}.xlsx"

    # =========================
    # DOWNLOAD EXCEL
    # =========================
    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        result.to_excel(writer, index=False, sheet_name="Summary")

    st.download_button(
        label="📥 Download Excel Report",
        data=output.getvalue(),
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("👆 Upload your Excel file to start")
