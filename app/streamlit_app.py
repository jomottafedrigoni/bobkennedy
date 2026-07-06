from pathlib import Path
import subprocess

import streamlit as st

ROOT = Path(__file__).resolve().parent.parent

st.set_page_config(
    page_title="Bob Kennedy",
    layout="wide"
)

st.title("Bob Kennedy")

with st.sidebar:

    st.header("ETL")

    if st.button("Aggiorna Dati"):

        with st.spinner("Elaborazione in corso..."):

            result = subprocess.run(
                [
                    "python",
                    str(ROOT / "scripts" / "run_etl.py")
                ],
                capture_output=True,
                text=True
            )

        st.success("ETL completato")