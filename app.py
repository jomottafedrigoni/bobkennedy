import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="Rielaborazione Report Transazioni",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Rielaborazione Report Transazioni")
st.markdown("Carica il file Excel delle transazioni per applicare la mappatura ed esplorare le analisi mensili.")

# --- DIZIONARIO DI MAPPATURA AGGIORNATO ---
MAPPATURA_DEFAULT = {
    "07.02.02.740.00RITR_CUSU": "Canteen expenses",
    "07.02.02.740.00RITR_HSE": "Canteen expenses",
    "07.02.02.740.00RITR_MAINT": "Canteen expenses",
    "07.02.02.740.00RITR_OPIND": "Canteen expenses",
    "07.02.02.740.00RITR_ORDF": "Canteen expenses",
    "07.02.02.740.00RITR_PRODIND": "Canteen expenses",
    "07.02.02.740.00RITR_QUAIND": "Canteen expenses",
    "07.02.02.740.00RITR_WARIND": "Canteen expenses",
    "07.02.02.780.00RITR_PLA": "Cleaning expenses",
    "07.02.02.340.04RITR_PLA": "Consultancy",
    "07.02.02.340.04RITR_HSE": "Consultancy",
    "07.02.02.340.04RITR_QUAIND": "Consultancy",
    "07.02.02.410.06RITR_PLA": "Insurance",
    "07.02.02.410.05RITR_PLA": "Insurance",
    "07.02.02.200.00RITR_PLA": "Intercompany",
    "07.02.02.170.01RITR_HSE": "Maintenance (Fix)",
    "07.02.02.170.01RITR_MAINT": "Maintenance (Fix)",
    "07.02.02.750.00RITR_HSE": "Other",
    "07.04.04.510.04RITR_PLA": "Other",
    "07.01.01.210.00RITR_HSE": "Spare parts and Equipments",
    "07.01.01.210.00RITR_MAINT": "VIC - Maintenance",
    "07.04.04.240.00RITR_QUAIND": "Subscription and Associations fees",
    "07.02.02.430.02RITR_OPIND": "Travel expenses",
    "07.02.02.430.07RITR_OPIND": "Travel expenses",
    "07.02.02.430.02RITR_CUSU": "Travel expenses",
    "07.02.02.430.07RITR_CUSU": "Travel expenses",
    "07.02.02.430.02RITR_QUAIND": "Travel expenses",
    "07.02.02.430.07RITR_QUAIND": "Travel expenses",
    "07.02.02.430.03RITR_MAINT": "Travel expenses",
    "07.02.02.430.03RITR_OPIND": "Travel expenses",
    "07.02.02.430.03RITR_QUAIND": "Travel expenses",
    "07.02.02.430.06RITR_OPIND": "Travel expenses",
    "07.02.02.430.06RITR_QUAIND": "Travel expenses",
    "07.02.02.430.02RITR_MAINT": "Travel expenses",
    "07.02.02.430.03RITR_CUSU": "Travel expenses",
    "07.02.02.430.03RITR_ORDF": "Travel expenses",
    "07.02.02.430.01RITR_OPIND": "Travel expenses",
    "07.02.02.430.01RITR_QUAIND": "Travel expenses",
    "07.02.02.430.05RITR_OPIND": "Travel expenses",
    "07.01.01.180.02RITR_MAINT": "VIC - Maintenance",
    "07.01.01.180.01RITR_MAINT": "VIC - Maintenance"
}

# --- SEZIONE INFO DYNAMICS ---
with st.expander("📌 Codici Conto per Filtro Dynamics (Clicca per espandere)", expanded=True):
    # Estrazione dei soli codici conto (formato ##.##.##.###.##) dai primi 15 caratteri delle chiavi
    codici_conto_unici = sorted(list(set([chiave[:15] for chiave in MAPPATURA_DEFAULT.keys()])))
    
    st.markdown("Usa questi codici per filtrare le transazioni direttamente su **Dynamics**.")
    
    col_tab, col_stringa = st.columns([1, 2])
    
    with col_tab:
        df_codici = pd.DataFrame(codici_conto_unici, columns=["Codice Conto"])
        st.dataframe(df_codici, hide_index=True, use_container_width=True)
        
    with col_stringa:
        st.markdown("**Stringa pronta per il filtro Dynamics (OR):**")
        stringa_dynamics = "|".join(codici_conto_unici)
        st.code(stringa_dynamics, language="text")

st.markdown("---")

def elabora_report(df):
    df_out = df.copy()

    # 1. DROP PROVVCH
    col_giustificativo = "Giustificativo"
    if col_giustificativo in df_out.columns:
        maschera_provvch = df_out[col_giustificativo].astype(str).str.startswith("PROVVCH", na=False)
        n_provvch = maschera_provvch.sum()
        df_out = df_out[~maschera_provvch]
    else:
        n_provvch = 0

    # 2. ESTRAZIONE E MAPPATURA
    col_valore = "Valore visualizzato conto"
    if col_valore in df_out.columns:
        def estrai_chiave(val):
            if pd.isna(val):
                return ""
            parti = str(val).split('-')
            if len(parti) >= 2:
                return f"{parti[0].strip()}{parti[1].strip()}"
            return str(val).strip()

        df_out["Codice Mappatura Estratto"] = df_out[col_valore].apply(estrai_chiave)
        
        # Mappatura e Drop di tutte le voci fuori lista
        df_out["Mappatura"] = df_out["Codice Mappatura Estratto"].map(MAPPATURA_DEFAULT)
        
        # Righe non presenti nel dizionario
        maschera_non_mappate = df_out["Mappatura"].isna()
        n_non_mappate = maschera_non_mappate.sum()
        
        # DROPPING DELLE RIGHE ESCLUSE
        df_out = df_out[~maschera_non_mappate]

        # Posizionamento colonna
        idx = df_out.columns.get_loc(col_valore) + 1
        cols = list(df_out.columns)
        mappatura_col = cols.pop(cols.index("Mappatura"))
        cols.insert(idx, mappatura_col)
        df_out = df_out[cols]
    else:
        st.error(f"Impossibile trovare la colonna '{col_valore}' nel file Excel.")
        n_non_mappate = 0

    return df_out, n_provvch, n_non_mappate

# --- CARICAMENTO FILE ---
uploaded_excel = st.file_uploader("Carica il file Excel del Report (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_excel is not None:
    try:
        df_input = pd.read_excel(uploaded_excel)
        df_completo, n_provvch, n_non_mappate = elabora_report(df_input)

        # Notifiche di scarto
        if n_provvch > 0 or n_non_mappate > 0:
            st.warning(
                f"🧹 **Filtro applicato:** Rimosse **{n_provvch}** righe 'PROVVCH' "
                f"e **{n_non_mappate}** righe non appartenenti alle categorie mappate."
            )

        # --- MENU A COMPARSA PER IL FILTRO ---
        st.sidebar.header("🔍 Filtri Report")
        
        categorie_disponibili = ["TUTTE LE CATEGORIE"] + sorted(df_completo["Mappatura"].unique())
        
        categoria_selezionata = st.sidebar.selectbox(
            "Seleziona la Mappatura da visualizzare:",
            options=categorie_disponibili,
            index=0
        )

        if categoria_selezionata == "TUTTE LE CATEGORIE":
            df_filtered = df_completo.copy()
        else:
            df_filtered = df_completo[df_completo["Mappatura"] == categoria_selezionata]

        # KPI
        c1, c2 = st.columns(2)
        c1.metric("Totale Transazioni Valide", len(df_filtered))
        
        col_importo = "Importo" if "Importo" in df_filtered.columns else None
        if col_importo:
            tot_spesa = pd.to_numeric(df_filtered[col_importo], errors='coerce').sum()
            c2.metric("Spesa Totale Mappata", f"€ {tot_spesa:,.2f}")

        st.markdown("---")

        # --- SEZIONE REPORTING MENSILE ---
        st.subheader("📈 Report Sintetico per Mese e Mappatura")
        
        col_data = "Data" if "Data" in df_filtered.columns else ("Data documento" if "Data documento" in df_filtered.columns else None)

        if col_data and col_importo and not df_filtered.empty:
            df_filtered["Mese"] = pd.to_datetime(df_filtered[col_data], dayfirst=True, errors='coerce').dt.to_period('M').astype(str)
            df_filtered[col_importo] = pd.to_numeric(df_filtered[col_importo], errors='coerce').fillna(0)

            # Tabella Pivot base (Mese x Mappatura)
            pivot_df = df_filtered.pivot_table(
                index="Mese", 
                columns="Mappatura", 
                values=col_importo, 
                aggfunc="sum", 
                fill_value=0
            )

            # 1. Calcolo Somma Mensile di tutte le categorie
            totale_mensile = pivot_df.sum(axis=1)

            # 2. Inseriamo la colonna TOTALE MENSILE in PRIMA POSIZIONE
            pivot_df.insert(0, "TOTALE MENSILE", totale_mensile)

            # Visualizzazione tabella con Totale in prima colonna
            st.dataframe(pivot_df.style.format("€ {:,.2f}"), use_container_width=True)

            # Grafico a barre (escludendo la colonna TOTALE MENSILE dal grafico)
            st.subheader("📊 Andamento Mensile delle Spese")
            chart_data = pivot_df.drop(columns=["TOTALE MENSILE"], errors="ignore")
            st.bar_chart(chart_data)

        elif df_filtered.empty:
            st.warning("Nessun dato presente dopo l'applicazione dei filtri.")
        else:
            st.info("Assicurati che le colonne 'Data' e 'Importo' siano presenti nel file.")

        st.markdown("---")
        st.subheader("📋 Dettaglio Transazioni Mappate")
        st.dataframe(df_filtered, use_container_width=True)

        # Download Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_filtered.to_excel(writer, index=False, sheet_name='Dettaglio_Transazioni')
            if col_data and col_importo and not df_filtered.empty:
                pivot_df.to_excel(writer, sheet_name='Sintesi_Mensile')
        buffer.seek(0)

        st.download_button(
            label="📥 Scarica Report Filtrato in Excel",
            data=buffer,
            file_name="Report_Transazioni_Mappato.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Errore durante l'elaborazione del file Excel: {e}")