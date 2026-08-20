import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(
    page_title="Rielaborazione Report Transazioni & Budget",
    page_icon="📊",
    layout="wide"
)

# --- DIZIONARIO DI MAPPATURA DEFAULT ---
MAPPATURA_DEFAULT = {
    "07.02.02.740.00RITR_CUSU": "Canteen Expenses",
    "07.02.02.740.00RITR_HSE": "Canteen Expenses",
    "07.02.02.740.00RITR_MAINT": "Canteen Expenses",
    "07.02.02.740.00RITR_OPIND": "Canteen Expenses",
    "07.02.02.740.00RITR_ORDF": "Canteen Expenses",
    "07.02.02.740.00RITR_PRODIND": "Canteen Expenses",
    "07.02.02.740.00RITR_QUAIND": "Canteen Expenses",
    "07.02.02.740.00RITR_WARIND": "Canteen Expenses",
    "07.02.02.780.00RITR_PLA": "Cleaning Expenses",
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
    "07.01.01.210.00RITR_HSE": "Spare Parts and Equipments",
    "07.04.04.240.00RITR_QUAIND": "Subscription and Associations Fees",
    "07.02.02.430.02RITR_OPIND": "Travel Expenses",
    "07.02.02.430.07RITR_OPIND": "Travel Expenses",
    "07.02.02.430.02RITR_CUSU": "Travel Expenses",
    "07.02.02.430.07RITR_CUSU": "Travel Expenses",
    "07.02.02.430.02RITR_QUAIND": "Travel Expenses",
    "07.02.02.430.07RITR_QUAIND": "Travel Expenses",
    "07.02.02.430.03RITR_MAINT": "Travel Expenses",
    "07.02.02.430.03RITR_OPIND": "Travel Expenses",
    "07.02.02.430.03RITR_QUAIND": "Travel Expenses",
    "07.02.02.430.06RITR_OPIND": "Travel Expenses",
    "07.02.02.430.06RITR_QUAIND": "Travel Expenses",
    "07.02.02.430.02RITR_MAINT": "Travel Expenses",
    "07.02.02.430.03RITR_CUSU": "Travel Expenses",
    "07.02.02.430.03RITR_ORDF": "Travel Expenses",
    "07.02.02.430.01RITR_OPIND": "Travel Expenses",
    "07.02.02.430.01RITR_QUAIND": "Travel Expenses",
    "07.02.02.430.05RITR_OPIND": "Travel Expenses",
    "07.01.01.180.02RITR_MAINT": "VIC - Maintenance (VAR)",
    "07.01.01.210.00RITR_MAINT": "VIC - Maintenance (VAR)",
    "07.01.01.180.01RITR_MAINT": "VIC - Maintenance (VAR)"
}

MAP_MESI = {
    'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 
    'May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08', 
    'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'
}

MESI_LISTA = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

st.title("📊 RITRAMA Report Costi Fissi e Manutenzione Variabile")

# --- CARICAMENTO FILE IN SIDEBAR ---
st.sidebar.header("📂 Caricamento File")
uploaded_transazioni = st.sidebar.file_uploader("1. File Transazioni (ACT)", type=["xlsx", "xls"], key="transazioni")
uploaded_fornitori = st.sidebar.file_uploader("2. File Anagrafica Fornitori", type=["xlsx", "xls"], key="fornitori")
uploaded_budget = st.sidebar.file_uploader("3. File Budget (BDG)", type=["xlsx", "xls"], key="budget")

def formatta_k_euro(val):
    if pd.isna(val) or round(val, 2) == 0:
        return "-"
    if val < 0:
        return f"({abs(val):,.2f})".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def colora_colonne_fornitori(df):
    if not isinstance(df, pd.DataFrame):
        return df

    styles = pd.DataFrame('', index=df.index, columns=df.columns)
    
    for col in df.columns:
        col_str = str(col)
        is_tot_col = "TOTALE" in col_str.upper()

        # Stile base per colonna normale o colonna di totale
        if "ACT" in col_str:
            bg_color = '#d9d9d9' if is_tot_col else '#f2f2f2'  # Grigio più scuro per i totali
            base_style = f'background-color: {bg_color}; color: black;'
        elif "BDG" in col_str:
            bg_color = '#e6e6e6' if is_tot_col else '#ffffff'  # Grigio chiaro/scuro per budget
            base_style = f'background-color: {bg_color}; color: black;'
        elif "Delta" in col_str:
            bg_color = '#b3d8ff' if is_tot_col else '#e6f2ff'  # Blu più scuro per i totali Delta
            base_style = f'background-color: {bg_color};'
        else:
            base_style = ''

        # Applicazione stile per riga e colonna
        for idx in df.index:
            idx_str = str(idx).upper()
            is_tot_row = "TOTALE" in idx_str
            
            style = base_style

            # Se si tratta di un Delta, applica la logica dei colori (Rosso/Verde)
            if "Delta" in col_str:
                val = df.loc[idx, col]
                if isinstance(val, (int, float)):
                    if val > 0.01:
                        style += ' color: red;'
                    elif val < -0.01:
                        style += ' color: green;'
                    else:
                        style += ' color: black;'

            # Applica Bold se è una colonna di Totale O una riga di Totale
            if is_tot_col or is_tot_row:
                style += ' font-weight: bold;'
                # Se la riga è TOTALE e non è Delta, scurisce leggermente lo sfondo
                if is_tot_row and "Delta" not in col_str:
                    style += ' background-color: #d0d0d0;'
                elif is_tot_row and "Delta" in col_str:
                    style += ' background-color: #99ccff;'

            styles.loc[idx, col] = style

    return styles

def processa_budget(df_bdg, anno_ref=2026):
    cols_mesi = [m for m in MESI_LISTA if m in df_bdg.columns]
    for m in cols_mesi:
        df_bdg[m] = pd.to_numeric(df_bdg[m], errors='coerce').fillna(0)
    
    df_melt = df_bdg.melt(
        id_vars=['PLANT', 'Mappatura', 'Supplier'],
        value_vars=cols_mesi,
        var_name='Mese_Nome',
        value_name='Importo_BDG_Orig'
    )
    
    df_melt['Mese_Num'] = df_melt['Mese_Nome'].map(MAP_MESI)
    df_melt['Mese'] = str(anno_ref) + "-" + df_melt['Mese_Num']
    df_melt['Importo_BDG_kEUR'] = (df_melt['Importo_BDG_Orig'] * -1) / 1000.0
    
    df_melt['PLANT'] = df_melt['PLANT'].astype(str).str.strip().str.upper()
    df_melt['Mappatura'] = df_melt['Mappatura'].astype(str).str.strip().str.title()
    df_melt['Supplier'] = df_melt['Supplier'].astype(str).str.strip()
    
    return df_melt

def elabora_report(df_trans, df_forn):
    df_out = df_trans.copy()

    if "Giustificativo" in df_out.columns:
        maschera_provvch = df_out["Giustificativo"].astype(str).str.startswith("PROVVCH", na=False)
        df_out = df_out[~maschera_provvch]

    for col in ["Importo nella valuta della transazione", "Importo", "Importo nella valuta di dichiarazione"]:
        if col in df_out.columns:
            df_out[col] = pd.to_numeric(df_out[col], errors='coerce').fillna(0)

    mppa_fornitori = {}
    if "Account fornitore" in df_forn.columns and "Nome" in df_forn.columns:
        df_forn["Account_Clean"] = df_forn["Account fornitore"].astype(str).str.strip()
        mppa_fornitori = dict(zip(df_forn["Account_Clean"], df_forn["Nome"]))

    if "Valore visualizzato conto" in df_out.columns:
        def estrai_chiave(val):
            parti = str(val).split('-')
            return f"{parti[0].strip()}{parti[1].strip()}" if len(parti) >= 2 else str(val).strip()

        def estrai_stab(val):
            match = re.search(r'(RITR-(?:SAS|CAP|BAS|HQ|CPN))', str(val), re.IGNORECASE)
            return match.group(1).upper() if match else "Altro/Non Spec."

        def estrai_cod_forn(val):
            val_str = str(val)
            if '--' in val_str:
                return val_str.split('--')[-1].split('-')[0].strip()
            parti = [p.strip() for p in val_str.split('-') if p.strip()]
            return parti[-1] if parti else ""

        df_out["Codice Mappatura Estratto"] = df_out["Valore visualizzato conto"].apply(estrai_chiave)
        df_out["Mappatura"] = df_out["Codice Mappatura Estratto"].map(MAPPATURA_DEFAULT).str.title()
        df_out["Stabilimento"] = df_out["Valore visualizzato conto"].apply(estrai_stab)
        df_out["Codice Fornitore Estratto"] = df_out["Valore visualizzato conto"].apply(estrai_cod_forn)
        
        df_out["Fornitore"] = df_out["Codice Fornitore Estratto"].map(mppa_fornitori)
        df_out["Fornitore"] = df_out["Fornitore"].fillna(
            df_out["Codice Fornitore Estratto"].apply(lambda x: f"Cod. {x}" if x else "Non Definito")
        )

        df_out = df_out[df_out["Mappatura"].notna()]
    return df_out

if uploaded_transazioni and uploaded_fornitori and uploaded_budget:
    try:
        df_act_raw = pd.read_excel(uploaded_transazioni)
        df_forn_raw = pd.read_excel(uploaded_fornitori)
        df_bdg_raw = pd.read_excel(uploaded_budget)

        df_act = elabora_report(df_act_raw, df_forn_raw)

        col_data = "Data" if "Data" in df_act.columns else ("Data documento" if "Data documento" in df_act.columns else None)
        col_imp = "Importo" if "Importo" in df_act.columns else None

        if col_data and col_imp:
            df_act["dt_temp"] = pd.to_datetime(df_act[col_data], dayfirst=True, errors='coerce')
            df_act["Anno"] = df_act["dt_temp"].dt.year.astype(str)
            df_act["Mese_Num"] = df_act["dt_temp"].dt.strftime('%m')
            df_act["Mese"] = df_act["Anno"] + "-" + df_act["Mese_Num"]
            df_act["Importo_ACT_kEUR"] = (df_act[col_imp] * -1) / 1000.0

            st.sidebar.markdown("---")
            st.sidebar.header("🌐 Selezione Pagina")
            pagina = st.sidebar.radio("Seleziona Pagina:", ["📌 Fixed Costs (Costi Fissi)", "⚙️ Variable Industrial Costs (VIC)"])

            st.sidebar.markdown("---")
            st.sidebar.header("🔍 Filtri")
            
            # 1. Filtro Anno
            anni_disponibili = sorted([a for a in df_act["Anno"].unique() if a != "nan"], reverse=True)
            anno_sel = st.sidebar.selectbox("Seleziona Anno:", anni_disponibili, index=0)

            df_bdg = processa_budget(df_bdg_raw, anno_ref=int(anno_sel) if anno_sel else 2026)

            # 2. Filtro Plant
            plants = sorted(list(set(df_act["Stabilimento"].unique()).union(set(df_bdg["PLANT"].unique()))))
            plant_sel = st.sidebar.selectbox("Seleziona Plant:", ["TUTTI I PLANT"] + plants)

            # 3. Filtro Mappatura
            mappature_totali = sorted(list(set(df_act["Mappatura"].unique()).union(set(df_bdg["Mappatura"].unique()))))
            mappatura_sel = st.sidebar.multiselect("Seleziona Voce di Spesa:", options=mappature_totali, default=[])

            # 4. Filtro Fornitore
            fornitori_totali = sorted(list(set(df_act["Fornitore"].unique()).union(set(df_bdg["Supplier"].unique()))))
            fornitore_sel = st.sidebar.selectbox("Seleziona Fornitore:", ["TUTTI I FORNITORI"] + fornitori_totali)

            # Applicazione Filtri Base
            df_act_f = df_act[df_act["Anno"] == anno_sel]
            df_bdg_f = df_bdg.copy()

            if plant_sel != "TUTTI I PLANT":
                df_act_f = df_act_f[df_act_f["Stabilimento"] == plant_sel]
                df_bdg_f = df_bdg_f[df_bdg_f["PLANT"] == plant_sel]

            if mappatura_sel:
                df_act_f = df_act_f[df_act_f["Mappatura"].isin(mappatura_sel)]
                df_bdg_f = df_bdg_f[df_bdg_f["Mappatura"].isin(mappatura_sel)]

            if fornitore_sel != "TUTTI I FORNITORI":
                df_act_f = df_act_f[df_act_f["Fornitore"] == fornitore_sel]
                df_bdg_f = df_bdg_f[df_bdg_f["Supplier"] == fornitore_sel]

            # 5. Filtro Mesi
            mesi_disponibili = sorted(list(set(df_act_f["Mese"].unique()).union(set(df_bdg_f["Mese"].unique()))))
            mesi_sel = st.sidebar.multiselect("Seleziona Mesi:", options=mesi_disponibili, default=[])

            if mesi_sel:
                df_act_f_mesi = df_act_f[df_act_f["Mese"].isin(mesi_sel)]
            else:
                df_act_f_mesi = df_act_f.copy()

            tab_dettaglio, tab_act_bdg = st.tabs(["📊 Report Analisi Spese", "⚖️ ACT vs BDG"])

            # ==========================================
            # TAB 1: METRICHE E REPORTE SULLE TRANSAZIONI
            # ==========================================
            with tab_dettaglio:
                if pagina == "📌 Fixed Costs (Costi Fissi)":
                    df_sezione = df_act_f_mesi[df_act_f_mesi["Mappatura"] != "Vic - Maintenance (Var)"].copy()
                    st.header("📌 Fixed Costs - Dettaglio Analisi Spese")
                else:
                    df_sezione = df_act_f_mesi[df_act_f_mesi["Mappatura"] == "Vic - Maintenance (Var)"].copy()
                    st.header("⚙️ Variable Costs - Dettaglio Analisi Spese")

                oggi = datetime.now()
                mese_prec_num = 12 if oggi.month == 1 else oggi.month - 1
                mese_chiuso_ref = f"{anno_sel}-{mese_prec_num:02d}"

                if mesi_sel:
                    m_max = max(mesi_sel)
                else:
                    mesi_act_validi = [m for m in df_act_f["Mese"].unique() if pd.notna(m) and m <= mese_chiuso_ref]
                    m_max = max(mesi_act_validi) if mesi_act_validi else (max(df_act_f["Mese"].unique()) if len(df_act_f) > 0 else "")

                df_act_sezione_base = df_act_f[df_act_f["Mappatura"] != "Vic - Maintenance (Var)"] if pagina == "📌 Fixed Costs (Costi Fissi)" else df_act_f[df_act_f["Mappatura"] == "Vic - Maintenance (Var)"]
                
                s_mtd = df_act_sezione_base[df_act_sezione_base["Mese"] == m_max]["Importo_ACT_kEUR"].sum() if m_max else 0
                s_ytd = df_act_sezione_base[df_act_sezione_base["Mese"] <= m_max]["Importo_ACT_kEUR"].sum() if m_max else 0

                c1, c2 = st.columns(2)
                c1.metric(f"Totale YTD (fino a {m_max})", f"{formatta_k_euro(s_ytd)} k€")
                c2.metric(f"Totale MTD ({m_max})", f"{formatta_k_euro(s_mtd)} k€")
                st.markdown("---")

                st.subheader("📈 Actual MTD (in k€)")
                piv_mensile = df_sezione.pivot_table(index="Mese", columns="Mappatura", values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
                piv_mensile["TOTALE MENSILE"] = piv_mensile.sum(axis=1)
                tot_map = piv_mensile.sum(axis=0)
                tot_map.name = "TOTALE PER MAPPATURA"
                piv_mensile_full = pd.concat([piv_mensile, pd.DataFrame(tot_map).T])
                st.dataframe(piv_mensile_full.style.format(formatta_k_euro), use_container_width=True)

                st.subheader("📊 ACT MTD per Plant (k€)")
                piv_plant = df_sezione.pivot_table(index="Mese", columns="Stabilimento", values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
                st.bar_chart(piv_plant)

                st.markdown("---")
                st.subheader("🏢 ACT per Fornitore YTD (in k€)")
                piv_ytd = df_sezione.pivot_table(index="Fornitore", columns="Mappatura", values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
                piv_ytd["TOTALE FORNITORE"] = piv_ytd.sum(axis=1)
                piv_ytd = piv_ytd.sort_values(by="TOTALE FORNITORE", ascending=False)
                tot_cat = piv_ytd.sum(axis=0)
                tot_cat.name = "TOTALE CATEGORIA"
                piv_ytd_full = pd.concat([piv_ytd, pd.DataFrame(tot_cat).T])
                st.dataframe(piv_ytd_full.style.format(formatta_k_euro), use_container_width=True)

            # ==========================================
            # TAB 2: ACT VS BDG (CON LOGICA MTD / YTD)
            # ==========================================
            with tab_act_bdg:
                tipo_vista = st.radio("Orizzonte Temporale:", ["MTD (Valori Mensili Puntuali)", "YTD (Cumulato Progressivo Mese per Mese)"], horizontal=True)

                if pagina == "📌 Fixed Costs (Costi Fissi)":
                    st.header("📌 Fixed Costs - Analisi ACT vs BDG")
                    
                    st.subheader("MTD / YTD Actual vs Budget per Voce di Spesa")

                    df_act_m1 = df_act_f[df_act_f["Mappatura"] != "Vic - Maintenance (Var)"]
                    df_bdg_m1 = df_bdg_f[df_bdg_f["Mappatura"] != "Vic - Maintenance (Var)"]

                    piv_act_m = df_act_m1.pivot_table(index="Mese", columns="Mappatura", values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
                    piv_bdg_m = df_bdg_m1.pivot_table(index="Mese", columns="Mappatura", values="Importo_BDG_kEUR", aggfunc="sum", fill_value=0)

                    # Reindicizza entrambe le pivot su tutti i 12 mesi dell'anno
                    mesi_completi = [f"{anno_sel}-{idx:02d}" for idx in range(1, 13)]
                    piv_act_m = piv_act_m.reindex(mesi_completi, fill_value=0)
                    piv_bdg_m = piv_bdg_m.reindex(mesi_completi, fill_value=0)

                    if "YTD" in tipo_vista:
                        # 1. Trova l'ultimo mese in cui esistono transazioni ACT reali
                        ultimo_mese_act = df_act_m1["Mese"].max() if len(df_act_m1) > 0 else None

                        # 2. Calcola il cumulato progressivo
                        piv_act_m = piv_act_m.cumsum(axis=0)
                        piv_bdg_m = piv_bdg_m.cumsum(axis=0)

                        # 3. Applica il Forward Fill dei dati ACT dopo l'ultimo mese consuntivato
                        if ultimo_mese_act and ultimo_mese_act in piv_act_m.index:
                            idx_ultimo = piv_act_m.index.get_loc(ultimo_mese_act)
                            # Trascina il valore dell'ultimo mese consuntivato a tutti i mesi successivi
                            for c in piv_act_m.columns:
                                piv_act_m.iloc[idx_ultimo+1:, piv_act_m.columns.get_loc(c)] = piv_act_m.iloc[idx_ultimo][c]

                    all_maps = sorted(list(set(piv_act_m.columns).union(set(piv_bdg_m.columns))))
                    
                    rows_m1 = []
                    for idx, m_code in enumerate(MESI_LISTA, 1):
                        m_str = f"{anno_sel}-{idx:02d}"
                        r = {"Mese": f"{m_code} ({m_str})"}
                        tot_act_mese = 0.0
                        tot_bdg_mese = 0.0

                        for m_cat in all_maps:
                            v_act = piv_act_m.loc[m_str, m_cat] if (m_str in piv_act_m.index and m_cat in piv_act_m.columns) else 0.0
                            v_bdg = piv_bdg_m.loc[m_str, m_cat] if (m_str in piv_bdg_m.index and m_cat in piv_bdg_m.columns) else 0.0
                            
                            r[f"{m_cat} ACT"] = v_act
                            r[f"{m_cat} BDG"] = v_bdg
                            r[f"{m_cat} Delta"] = v_bdg - v_act

                            tot_act_mese += v_act
                            tot_bdg_mese += v_bdg

                        r["TOTALE ACT"] = tot_act_mese
                        r["TOTALE BDG"] = tot_bdg_mese
                        r["TOTALE Delta"] = tot_bdg_mese - tot_act_mese

                        rows_m1.append(r)

                    df_mod1 = pd.DataFrame(rows_m1).set_index("Mese")

                    # Calcolo Totali Anno dinamico
                    tot_dict = {}
                    if "YTD" in tipo_vista:
                        for col in df_mod1.columns:
                            if "Delta" in col:
                                base_col = col.replace(" Delta", "")
                                tot_dict[col] = tot_dict[f"{base_col} BDG"] - tot_dict[f"{base_col} ACT"]
                            else:
                                tot_dict[col] = df_mod1.iloc[-1][col]
                    else:
                        for col in df_mod1.columns:
                            if "Delta" in col:
                                base_col = col.replace(" Delta", "")
                                tot_dict[col] = tot_dict[f"{base_col} BDG"] - tot_dict[f"{base_col} ACT"]
                            else:
                                tot_dict[col] = df_mod1[col].sum()
                    
                    df_tot_m1 = pd.DataFrame([tot_dict], index=["TOTALE ANNO"])
                    df_mod1_full = pd.concat([df_mod1, df_tot_m1])

                    st.dataframe(
                        df_mod1_full.style.format(formatta_k_euro).apply(colora_colonne_fornitori, axis=None),
                        use_container_width=True
                    )

                    st.markdown("---")

                    st.subheader("MTD Actual vs Budget per Fornitore e Voce di Spesa")
                    map_singola = st.selectbox("Seleziona Voce di Spesa:", all_maps if all_maps else ["Tutte"])

                    if map_singola:
                        df_act_m2 = df_act_f[df_act_f["Mappatura"] == map_singola]
                        df_bdg_m2 = df_bdg_f[df_bdg_f["Mappatura"] == map_singola]

                        fornitori_unici = sorted(list(set(df_act_m2["Fornitore"].unique()).union(set(df_bdg_m2["Supplier"].unique()))))

                        piv_act_f = df_act_m2.pivot_table(index="Fornitore", columns="Mese", values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
                        piv_bdg_f = df_bdg_m2.pivot_table(index="Supplier", columns="Mese", values="Importo_BDG_kEUR", aggfunc="sum", fill_value=0)

                        if "YTD" in tipo_vista:
                            piv_act_f = piv_act_f.cumsum(axis=1).ffill(axis=1)
                            piv_bdg_f = piv_bdg_f.cumsum(axis=1).ffill(axis=1)

                        rows_mod2 = []
                        for forn in fornitori_unici:
                            r = {"Fornitore": forn}
                            for idx, m_code in enumerate(MESI_LISTA, 1):
                                m_str = f"{anno_sel}-{idx:02d}"
                                v_act = piv_act_f.loc[forn, m_str] if (forn in piv_act_f.index and m_str in piv_act_f.columns) else 0.0
                                v_bdg = piv_bdg_f.loc[forn, m_str] if (forn in piv_bdg_f.index and m_str in piv_bdg_f.columns) else 0.0
                                
                                r[f"ACT {m_code}"] = v_act
                                r[f"BDG {m_code}"] = v_bdg
                                r[f"Delta {m_code}"] = v_bdg - v_act

                            if "YTD" in tipo_vista:
                                r["TOTALE ACT"] = r[f"ACT {MESI_LISTA[-1]}"]
                                r["TOTALE BDG"] = r[f"BDG {MESI_LISTA[-1]}"]
                            else:
                                r["TOTALE ACT"] = sum([r[f"ACT {m_code}"] for m_code in MESI_LISTA])
                                r["TOTALE BDG"] = sum([r[f"BDG {m_code}"] for m_code in MESI_LISTA])

                            r["TOTALE Delta"] = r["TOTALE BDG"] - r["TOTALE ACT"]
                            rows_mod2.append(r)

                        df_mod2 = pd.DataFrame(rows_mod2).set_index("Fornitore")

                        tot_cols = df_mod2.sum(axis=0)
                        df_tot_mod2 = pd.DataFrame(tot_cols).T
                        df_tot_mod2.index = ["TOTALE FORNITORI"]
                        df_mod2_full = pd.concat([df_mod2, df_tot_mod2])

                        st.dataframe(
                            df_mod2_full.style.format(formatta_k_euro).apply(colora_colonne_fornitori, axis=None),
                            use_container_width=True
                        )

                else:
                    st.header("⚙️ Variable Costs - Maintenance (VAR)")
                    st.subheader("MTD / YTD Actual vs Budget")

                    df_act_v = df_act_f[df_act_f["Mappatura"] == "Vic - Maintenance (Var)"]
                    df_bdg_v = df_bdg_f[df_bdg_f["Mappatura"] == "Vic - Maintenance (Var)"]

                    piv_act_v = df_act_v.pivot_table(index="Mese", values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
                    piv_bdg_v = df_bdg_v.pivot_table(index="Mese", values="Importo_BDG_kEUR", aggfunc="sum", fill_value=0)

                    # Reindicizza sui 12 mesi completi prima di fare cumsum + ffill
                    mesi_completi = [f"{anno_sel}-{idx:02d}" for idx in range(1, 13)]
                    piv_act_v = piv_act_v.reindex(mesi_completi, fill_value=0)
                    piv_bdg_v = piv_bdg_v.reindex(mesi_completi, fill_value=0)

                    if "YTD" in tipo_vista:
                        piv_act_v = piv_act_v.cumsum(axis=0).ffill()
                        piv_bdg_v = piv_bdg_v.cumsum(axis=0).ffill()

                    rows_v = []
                    for idx, m_code in enumerate(MESI_LISTA, 1):
                        m_str = f"{anno_sel}-{idx:02d}"
                        v_act = piv_act_v.loc[m_str, "Importo_ACT_kEUR"] if m_str in piv_act_v.index else 0.0
                        v_bdg = piv_bdg_v.loc[m_str, "Importo_BDG_kEUR"] if m_str in piv_bdg_v.index else 0.0
                        
                        rows_v.append({
                            "Mese": f"{m_code} ({m_str})",
                            "Man Var ACT k€": v_act,
                            "Man Var BDG k€": v_bdg,
                            "Delta k€": v_bdg - v_act
                        })

                    df_mod3 = pd.DataFrame(rows_v).set_index("Mese")

                    # Calcolo Totali Anno VIC
                    if "YTD" in tipo_vista:
                        tot_act_v = df_mod3["Man Var ACT k€"].iloc[-1]
                        tot_bdg_v = df_mod3["Man Var BDG k€"].iloc[-1]
                    else:
                        tot_act_v = df_mod3["Man Var ACT k€"].sum()
                        tot_bdg_v = df_mod3["Man Var BDG k€"].sum()

                    df_tot_mod3 = pd.DataFrame([{
                        "Man Var ACT k€": tot_act_v,
                        "Man Var BDG k€": tot_bdg_v,
                        "Delta k€": tot_bdg_v - tot_act_v
                    }], index=["TOTALE ANNO"])

                    df_mod3_full = pd.concat([df_mod3, df_tot_mod3])

                    st.dataframe(
                        df_mod3_full.style.format(formatta_k_euro).apply(colora_colonne_fornitori, axis=None),
                        use_container_width=True
                    )

    except Exception as e:
        st.error(f"Errore durante l'elaborazione dei dati: {e}")