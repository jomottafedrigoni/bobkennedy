from pathlib import Path

import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parents[2]

CACHE = ROOT / "data" / "cache"

st.title("Visualizzazione Transazioni e Verifica Riconciliabilità")

try:

    enriched_df = pd.read_parquet(
        CACHE / "enriched_transactions.parquet"
    )

except Exception as e:

    st.error(
        "Lancia prima run_etl.py oppure premi Aggiorna Dati"
    )

    st.stop()

# =====================
# FILTRI SIDEBAR
# =====================

st.sidebar.header("Filtri")

filtered_df = enriched_df.copy()

# Conto
conti = sorted(
    filtered_df["Conto principale"]
    .dropna()
    .astype(str)
    .unique()
)

conto_sel = st.sidebar.multiselect(
    "Conto",
    conti
)

if conto_sel:
    filtered_df = filtered_df[
        filtered_df["Conto principale"]
        .astype(str)
        .isin(conto_sel)
    ]


conti_bilancio = sorted(
    filtered_df["Conto Bilancio"]
    .dropna()
    .astype(str)
    .unique()
)

conto_bilancio_sel = st.sidebar.multiselect(
    "Conto Bilancio",
    conti_bilancio
)

if conto_bilancio_sel:

    filtered_df = filtered_df[
        filtered_df["Conto Bilancio"]
        .astype(str)
        .isin(conto_bilancio_sel)
    ]


# Cost Center
cost_centers = sorted(
    filtered_df["Cost Center"]
    .dropna()
    .astype(str)
    .unique()
)

cc_sel = st.sidebar.multiselect(
    "Cost Center",
    cost_centers
)

if cc_sel:
    filtered_df = filtered_df[
        filtered_df["Cost Center"]
        .astype(str)
        .isin(cc_sel)
    ]

# Plant
plants = sorted(
    filtered_df["Plant"]
    .dropna()
    .astype(str)
    .unique()
)

plant_sel = st.sidebar.multiselect(
    "Plant",
    plants
)

if plant_sel:
    filtered_df = filtered_df[
        filtered_df["Plant"]
        .astype(str)
        .isin(plant_sel)
    ]

# Data
if "Data" in filtered_df.columns:

    filtered_df["Data"] = pd.to_datetime(
        filtered_df["Data"]
    )

    min_date = filtered_df["Data"].min().date()
    max_date = filtered_df["Data"].max().date()

    data_sel = st.sidebar.date_input(
        "Intervallo Data",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )

    if len(data_sel) == 2:

        start_date, end_date = data_sel

        filtered_df = filtered_df[
            (filtered_df["Data"].dt.date >= start_date)
            &
            (filtered_df["Data"].dt.date <= end_date)
        ]

# =====================
# KPI
# =====================

mapped = filtered_df[
    filtered_df["Conto Bilancio"].notna()
]

unmapped = filtered_df[
    filtered_df["Conto Bilancio"].isna()
]

col1, col2, col3, col4 = st.columns(4)

col1.metric(
    "Transazioni",
    f"{len(filtered_df):,}"
)

col2.metric(
    "Importo Totale",
    f"{filtered_df['Importo'].sum():,.2f}"
)

col3.metric(
    "Mappate",
    f"{len(mapped):,}"
)

col4.metric(
    "Non Mappate",
    f"{len(unmapped):,}"
)

# =====================
# TAB
# =====================

tab1, tab2 = st.tabs(
    [
        "Transazioni",
        "Non Mappate"
    ]
)

with tab1:

    st.dataframe(
        filtered_df,
        use_container_width=True
    )

with tab2:

    st.dataframe(
        unmapped,
        use_container_width=True
    )