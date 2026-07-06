from pathlib import Path

import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parents[2]

CACHE = ROOT / "data" / "cache"

st.title("Confronto Dynamics vs Tagetik")

try:

    dynamics_pivot = pd.read_parquet(
        CACHE / "dynamics_pivot.parquet"
    )

    tagetik_pivot_df = pd.read_parquet(
        CACHE / "tagetik_pivot.parquet"
    )

except Exception:

    st.error(
        "Cache non trovata. Esegui prima run_etl.py"
    )

    st.stop()

# =====================
# COVERAGE
# =====================

dynamics_accounts = set(
    dynamics_pivot["Conto principale"]
    .astype(str)
    .unique()
)

tagetik_accounts = set(
    tagetik_pivot_df["LOCAL_ACCOUNT"]
    .astype(str)
    .unique()
)

common_accounts = (
    dynamics_accounts
    .intersection(tagetik_accounts)
)

only_dynamics = (
    dynamics_accounts
    - tagetik_accounts
)

only_tagetik = (
    tagetik_accounts
    - dynamics_accounts
)

# =====================
# KPI
# =====================

c1, c2, c3 = st.columns(3)

c1.metric(
    "Conti Dynamics",
    len(dynamics_accounts)
)

c2.metric(
    "Conti Tagetik",
    len(tagetik_accounts)
)

c3.metric(
    "Conti Comuni",
    len(common_accounts)
)

# =====================
# TAB
# =====================

tab1, tab2, tab3 = st.tabs(
    [
        "Dynamics",
        "Tagetik",
        "Coverage"
    ]
)

with tab1:

    st.subheader(
        "Dynamics aggregato"
    )

    st.dataframe(
        dynamics_pivot,
        use_container_width=True
    )

with tab2:

    st.subheader(
        "Tagetik aggregato"
    )

    st.dataframe(
        tagetik_pivot_df,
        use_container_width=True
    )

with tab3:

    st.write(
        f"Conti solo Dynamics: {len(only_dynamics)}"
    )

    st.write(
        f"Conti solo Tagetik: {len(only_tagetik)}"
    )

    if len(only_dynamics):

        st.subheader(
            "Presenti solo in Dynamics"
        )

        st.dataframe(
            pd.DataFrame(
                sorted(list(only_dynamics)),
                columns=["Conto"]
            ),
            use_container_width=True
        )

    if len(only_tagetik):

        st.subheader(
            "Presenti solo in Tagetik"
        )

        st.dataframe(
            pd.DataFrame(
                sorted(list(only_tagetik)),
                columns=["Conto"]
            ),
            use_container_width=True
        )