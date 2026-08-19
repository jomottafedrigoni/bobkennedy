import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="Rielaborazione Report Transazioni",
    page_icon="📊",
    layout="wide"
)

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
    "07.01.01.210.00RITR_MAINT": "VIC - Maintenance (VAR)",
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
    # --- COSTI VARIABILI (VIC) ---
    "07.01.01.180.02RITR_MAINT": "VIC - Maintenance (VAR)",
    "07.01.01.180.01RITR_MAINT": "VIC - Maintenance (VAR)"
}

st.title("📊 Rielaborazione Report Transazioni")

# --- CARICAMENTO FILE IN SIDEBAR PER ESSERE SEMPRE DISPONIBILE ---
st.sidebar.header("📂 Caricamento Dati")
uploaded_excel = st.sidebar.file_uploader("Carica Excel (.xlsx, .xls)", type=["xlsx", "xls"])

# --- SEZIONE INFO DYNAMICS (Visibile SOLO finché non carichi l'Excel) ---
if uploaded_excel is None:
    st.info("👋 **Benvenuto!** Copia i codici sottostanti per filtrare su Dynamics, poi carica il file esportato dalla barra laterale a sinistra.")
    
    with st.expander("📌 Codici Conto per Filtro Dynamics", expanded=True):
        codici_conto_unici = sorted(list(set([chiave[:15] for chiave in MAPPATURA_DEFAULT.keys()])))
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**1. Elenco in colonna (seleziona e copia stile Excel):**")
            testo_colonna = "\n".join(codici_conto_unici)
            st.text_area("Copia colonna", testo_colonna, height=220)
            
        with col2:
            st.markdown("**2. Stringa filtro con '|' (Filtro rapido Dynamics):**")
            stringa_dynamics = "|".join(codici_conto_unici)
            st.code(stringa_dynamics, language="text")

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
        df_out["Mappatura"] = df_out["Codice Mappatura Estratto"].map(MAPPATURA_DEFAULT)
        
        # Rimozione non mappati
        maschera_non_mappate = df_out["Mappatura"].isna()
        n_non_mappate = maschera_non_mappate.sum()
        df_out = df_out[~maschera_non_mappate]

        # Posizionamento colonna Mappatura
        idx = df_out.columns.get_loc(col_valore) + 1
        cols = list(df_out.columns)
        mappatura_col = cols.pop(cols.index("Mappatura"))
        cols.insert(idx, mappatura_col)
        df_out = df_out[cols]
    else:
        st.error(f"Impossibile trovare la colonna '{col_valore}' nel file Excel.")
        n_non_mappate = 0

    return df_out, n_provvch, n_non_mappate


# --- ELABORAZIONE E NAVIGAZIONE PER PAGINE ---
if uploaded_excel is not None:
    try:
        df_input = pd.read_excel(uploaded_excel)
        df_completo, n_provvch, n_non_mappate = elabora_report(df_input)

        # SELEZIONE PAGINA
        st.sidebar.markdown("---")
        st.sidebar.header("🌐 Navigazione")
        pagina = st.sidebar.radio(
            "Seleziona la Vista:",
            ["📌 Fixed Costs (Costi Fissi)", "⚙️ Variable Industrial Costs (VIC)"]
        )

        # SEPARAZIONE NETTA DEI DATI TRA FISSI E VARIABILI
        if pagina == "📌 Fixed Costs (Costi Fissi)":
            df_sezione = df_completo[df_completo["Mappatura"] != "VIC - Maintenance (VAR)"].copy()
            st.header("📌 Fixed Costs (Costi Fissi)")
        else:
            df_sezione = df_completo[df_completo["Mappatura"] == "VIC - Maintenance (VAR)"].copy()
            st.header("⚙️ Variable Industrial Costs - Maintenance (VAR)")

        # Notifiche di scarto generali
        if n_provvch > 0 or n_non_mappate > 0:
            st.caption(f"ℹ️ Dati filtrati globali: Rimosse {n_provvch} righe 'PROVVCH' e {n_non_mappate} righe non mappate.")

        # FILTRO SOTTO-CATEGORIA (Solo se siamo in Fixed Costs)
        if pagina == "📌 Fixed Costs (Costi Fissi)":
            st.sidebar.markdown("---")
            st.sidebar.header("🔍 Filtri Categoria")
            categorie_disponibili = ["TUTTE LE CATEGORIE FISSE"] + sorted(df_sezione["Mappatura"].unique())
            cat_sel = st.sidebar.selectbox("Seleziona la Mappatura:", options=categorie_disponibili)
            if cat_sel != "TUTTE LE CATEGORIE FISSE":
                df_sezione = df_sezione[df_sezione["Mappatura"] == cat_sel]

        # KPI DELLA SEZIONE SELEZIONATA
        c1, c2 = st.columns(2)
        c1.metric("Transazioni Trovate", len(df_sezione))
        
        col_importo = "Importo" if "Importo" in df_sezione.columns else None
        if col_importo:
            tot_spesa = pd.to_numeric(df_sezione[col_importo], errors='coerce').sum()
            c2.metric("Spesa Totale Sezione", f"€ {tot_spesa:,.2f}")

        st.markdown("---")

        # --- SEZIONE REPORTING MENSILE ---
        st.subheader("📈 Report Sintetico Mensile")
        col_data = "Data" if "Data" in df_sezione.columns else ("Data documento" if "Data documento" in df_sezione.columns else None)

        if col_data and col_importo and not df_sezione.empty:
            df_sezione["Mese"] = pd.to_datetime(df_sezione[col_data], dayfirst=True, errors='coerce').dt.to_period('M').astype(str)
            df_sezione[col_importo] = -pd.to_numeric(df_sezione[col_importo], errors='coerce').fillna(0)

            # Tabella Pivot
            pivot_df = df_sezione.pivot_table(
                index="Mese", 
                columns="Mappatura", 
                values=-col_importo, 
                aggfunc="sum", 
                fill_value=0
            )

            # Inseriamo la colonna Totale Mensile
            totale_mensile = pivot_df.sum(axis=1)
            pivot_df.insert(0, "TOTALE MENSILE", totale_mensile)

            # Visualizzazione tabella
            st.dataframe(pivot_df.style.format("€ {:,.2f}"), use_container_width=True)

            # Grafico a barre
            st.subheader("📊 Andamento Mensile")
            chart_data = pivot_df.drop(columns=["TOTALE MENSILE"], errors="ignore")
            st.bar_chart(chart_data)

        elif df_sezione.empty:
            st.warning("Nessuna transazione trovata per questa sezione o filtro selezionato.")
        else:
            st.info("Assicurati che le colonne 'Data' e 'Importo' siano presenti nel file.")

        st.markdown("---")
        st.subheader("📋 Dettaglio Transazioni")
        st.dataframe(df_sezione, use_container_width=True)

        # Download Excel dedicato alla sezione attiva
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_sezione.to_excel(writer, index=False, sheet_name='Dettaglio_Transazioni')
            if col_data and col_importo and not df_sezione.empty:
                pivot_df.to_excel(writer, sheet_name='Sintesi_Mensile')
        buffer.seek(0)

        nome_file = "Report_Fixed_Costs.xlsx" if pagina == "📌 Fixed Costs (Costi Fissi)" else "Report_Variable_Costs_VIC.xlsx"

        st.download_button(
            label=f"📥 Scarica Excel ({pagina})",
            data=buffer,
            file_name=nome_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Errore durante l'elaborazione del file Excel: {e}")