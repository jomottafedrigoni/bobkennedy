from pathlib import Path
import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parents[2]

CACHE = ROOT / "data" / "cache"

st.title("Riconciliazione Dynamics vs Tagetik")

reconciliation_df = pd.read_parquet(
    CACHE / "reconciliation.parquet"
)

months = [
    "01","02","03","04",
    "05","06","07","08",
    "09","10","11","12"
]

delta_cols = [
    f"{m}_DELTA"
    for m in months
    if f"{m}_DELTA" in reconciliation_df.columns
]

reconciliation_df["Delta Totale"] = (
    reconciliation_df[delta_cols]
    .abs()
    .sum(axis=1)
)

solo_differenze = st.checkbox(
    "Mostra solo differenze",
    value=True
)

if solo_differenze:

    reconciliation_df = reconciliation_df[
        reconciliation_df["Delta Totale"] > 0.01
    ]

st.metric(
    "Delta Totale",
    f"{reconciliation_df['Delta Totale'].sum():,.2f}"
)

st.dataframe(
    reconciliation_df,
    use_container_width=True
)