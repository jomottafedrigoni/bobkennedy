import streamlit as st
import pandas as pd
import re
import io
import requests
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

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
    "07.01.01.180.02RITR_MAINT": "Vic - Maintenance (Var)",
    "07.01.01.210.00RITR_MAINT": "Vic - Maintenance (Var)",
    "07.01.01.180.01RITR_MAINT": "Vic - Maintenance (Var)",
    "07.01.01.210.00":"Vic - Maintenance (Var)"
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

def genera_pdf_report(df_act_f, df_bdg_f, anno_sel, mese_ref):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=20,
        leftMargin=20,
        topMargin=30,
        bottomMargin=20
    )
    
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('CoverTitle', parent=styles['Title'], fontSize=24, leading=28, alignment=1, textColor=colors.HexColor('#000000'))
    subtitle_style = ParagraphStyle('CoverSubtitle', parent=styles['Normal'], fontSize=11, leading=16, alignment=1, textColor=colors.HexColor('#444444'))
    disclaimer_style = ParagraphStyle('CoverDisclaimer', parent=styles['Normal'], fontSize=9, leading=13, alignment=1, textColor=colors.HexColor('#666666'))
    
    header_style = ParagraphStyle('PageHeader', parent=styles['Heading1'], fontSize=13, leading=16, textColor=colors.HexColor('#1A1A1A'), spaceAfter=10)
    sub_header_style = ParagraphStyle('SubHeader', parent=styles['Heading2'], fontSize=10, leading=12, textColor=colors.HexColor('#222222'), spaceBefore=8, spaceAfter=6)
    
    header_cell = ParagraphStyle('HCell', parent=styles['Normal'], fontSize=7, leading=8, fontName='Helvetica-Bold', textColor=colors.white, alignment=1)
    num_style = ParagraphStyle('NCell', parent=styles['Normal'], fontSize=7, leading=8, textColor=colors.black, alignment=2)
    num_bold = ParagraphStyle('NBold', parent=styles['Normal'], fontSize=7, leading=8, fontName='Helvetica-Bold', textColor=colors.black, alignment=2)
    txt_style = ParagraphStyle('TCell', parent=styles['Normal'], fontSize=7, leading=8, textColor=colors.black, alignment=0)
    txt_bold = ParagraphStyle('TBold', parent=styles['Normal'], fontSize=7, leading=8, fontName='Helvetica-Bold', textColor=colors.black, alignment=0)

    def formatta_contabile(val):
        if val == 0 or pd.isna(val):
            return "-"
        val_ass = abs(val)
        str_val = f"{val_ass:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"({str_val})" if val < 0 else str_val

    story = []

    # ==========================================
    # PAGINA 0: COPERTINA
    # ==========================================
    story.append(Spacer(1, 40))
    story.append(Paragraph("Report Mensile Costi Fissi e Manutenzione Variabile", title_style))
    story.append(Spacer(1, 10))
    story.append(Paragraph(f"Data Report: {datetime.now().strftime('%d/%m/%Y')}", subtitle_style))
    story.append(Spacer(1, 30))
    
    subtext = (
        "<b>Ritrama S.p.A. - Industrial Controlling Team</b><br/><br/>"
        "Il presente documento è rivolto unicamente ai Plant Director degli stabilimenti di Ritrama S.p.A. "
        "di Basiano, Caponago e Sassoferrato.<br/>"
        "Le informazioni contenute nel report sono strettamente confidenziali, riservate e destinate ad uso interno. "
        "Ne è severamente vietata la riproduzione, la diffusione o la condivisione con soggetti non autorizzati, "
        "sia all'interno che all'esterno dell'organizzazione.<br/><br/>"
        "Per chiarimenti sui dati contenuti nel report o richieste di approfondimento analitico, fare riferimento a:<br/>"
        "• <b>Industrial Controller Italy:</b> Jona Motta (jona.motta@fedrigoni.com)<br/>"
        "• <b>Group Industrial Controller:</b> Roberto Ghirardi (roberto.ghirardi@fedrigoni.com)"
    )
    story.append(Paragraph(subtext, disclaimer_style))
    story.append(PageBreak())

    def crea_tabella_actual(df_source, is_vic=False, escludi_voci=[]):
        MAP_CONTI_VIC = {
            "07.01.01.180.01": "MATERIALI DI CONSUMO - MANUTENZIONI",
            "07.01.01.210.00": "MATERIALI PER ANTINFORTUNISTICO",
            "07.01.01.180.02": "SPESE DI MANUTENZIONE"
        }

        if is_vic:
            df_sub = df_source[df_source["Mappatura"] == "Vic - Maintenance (Var)"].copy()
            col_group = "Conto Contabile" if "Conto Contabile" in df_sub.columns else "Mappatura"
            
            if col_group == "Conto Contabile":
                df_sub[col_group] = df_sub[col_group].astype(str).str.strip()
                df_sub[col_group] = df_sub[col_group].map(lambda x: MAP_CONTI_VIC.get(x, x))
        else:
            df_sub = df_source[~df_source["Mappatura"].isin(escludi_voci)].copy()
            col_group = "Mappatura"

        if df_sub.empty:
            return Paragraph("<i>Nessun dato Actual disponibile.</i>", txt_style)
            
        piv = df_sub.pivot_table(index="Mese", columns=col_group, values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
        piv["TOTALE MENSILE"] = piv.sum(axis=1)
        tot_map = piv.sum(axis=0)
        tot_map.name = "TOTALE"
        piv_full = pd.concat([piv, pd.DataFrame(tot_map).T])

        headers = ["Mese"] + [str(c) for c in piv_full.columns]
        table_data = [[Paragraph(h, header_cell) for h in headers]]

        for idx, row in piv_full.iterrows():
            is_tot = (str(idx) == "TOTALE")
            r_data = [Paragraph(str(idx), txt_bold if is_tot else txt_style)]
            for val in row:
                r_data.append(Paragraph(formatta_contabile(val), num_bold if is_tot else num_style))
            table_data.append(r_data)

        col_w = [45] + [(735 / (len(headers) - 1))] * (len(headers) - 1)
        t = Table(table_data, colWidths=col_w, repeatRows=1)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1A1A1A')),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#CCCCCC')),
            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#F4F4F4')),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
        ]))
        return t

    def crea_tabella_act_vs_bdg(df_act_sub, df_bdg_sub, is_vic=False, escludi_voci=[]):
        delta_red = ParagraphStyle('DRed', parent=styles['Normal'], fontSize=6, leading=7, textColor=colors.HexColor('#C00000'), alignment=2)
        delta_red_bold = ParagraphStyle('DRedB', parent=styles['Normal'], fontSize=6, leading=7, fontName='Helvetica-Bold', textColor=colors.HexColor('#C00000'), alignment=2)
        delta_green = ParagraphStyle('DGreen', parent=styles['Normal'], fontSize=6, leading=7, textColor=colors.HexColor('#008000'), alignment=2)
        delta_green_bold = ParagraphStyle('DGreenB', parent=styles['Normal'], fontSize=6, leading=7, fontName='Helvetica-Bold', textColor=colors.HexColor('#008000'), alignment=2)
        
        num_s = ParagraphStyle('NCellS', parent=styles['Normal'], fontSize=6, leading=7, textColor=colors.black, alignment=2)
        num_b = ParagraphStyle('NBoldS', parent=styles['Normal'], fontSize=6, leading=7, fontName='Helvetica-Bold', textColor=colors.black, alignment=2)
        txt_s = ParagraphStyle('TCellS', parent=styles['Normal'], fontSize=6, leading=7, textColor=colors.black, alignment=0)
        txt_b = ParagraphStyle('TBoldS', parent=styles['Normal'], fontSize=6, leading=7, fontName='Helvetica-Bold', textColor=colors.black, alignment=0)

        def get_delta_style(v_delta, is_bold=False):
            if v_delta < -0.01:
                return delta_green_bold if is_bold else delta_green
            elif v_delta > 0.01:
                return delta_red_bold if is_bold else delta_red
            return num_b if is_bold else num_s

        if is_vic:
            df_act_cf = df_act_sub[df_act_sub["Mappatura"] == "Vic - Maintenance (Var)"]
            df_bdg_cf = df_bdg_sub[df_bdg_sub["Mappatura"] == "Vic - Maintenance (Var)"]

            piv_act_tot = df_act_cf.pivot_table(index="Mese", values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
            piv_bdg_tot = df_bdg_cf.pivot_table(index="Mese", values="Importo_BDG_kEUR", aggfunc="sum", fill_value=0)

            header_row1 = [Paragraph("Mese", header_cell), Paragraph("VIC TOTALE ACT", header_cell), Paragraph("VIC TOTALE BDG", header_cell), Paragraph("VIC TOTALE Delta", header_cell)]
            table_data = [header_row1]
            mesi_list = [f"{anno_sel}-{i:02d}" for i in range(1, 13)]
            
            tot_act_gen, tot_bdg_gen = 0.0, 0.0
            for m_str in mesi_list:
                m_breve = m_str[5:] if len(m_str) >= 7 else m_str
                v_act = piv_act_tot.loc[m_str, "Importo_ACT_kEUR"] if m_str in piv_act_tot.index else 0.0
                v_bdg = piv_bdg_tot.loc[m_str, "Importo_BDG_kEUR"] if m_str in piv_bdg_tot.index else 0.0
                v_delta = v_bdg - v_act
                tot_act_gen += v_act
                tot_bdg_gen += v_bdg

                table_data.append([
                    Paragraph(m_breve, txt_s),
                    Paragraph(formatta_contabile(v_act), num_s),
                    Paragraph(formatta_contabile(v_bdg), num_s),
                    Paragraph(formatta_contabile(v_delta), get_delta_style(v_delta))
                ])

            tot_delta_gen = tot_bdg_gen - tot_act_gen
            table_data.append([
                Paragraph("TOT", txt_b),
                Paragraph(formatta_contabile(tot_act_gen), num_b),
                Paragraph(formatta_contabile(tot_bdg_gen), num_b),
                Paragraph(formatta_contabile(tot_delta_gen), get_delta_style(tot_delta_gen, is_bold=True))
            ])

            t = Table(table_data, colWidths=[60, 230, 230, 230], repeatRows=1)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1A1A1A')),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#CCCCCC')),
                ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#F4F4F4')),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
                ('TOPPADDING', (0, 0), (-1, -1), 2),
            ]))
            return t

        df_act_cf = df_act_sub[~df_act_sub["Mappatura"].isin(escludi_voci)]
        df_bdg_cf = df_bdg_sub[~df_bdg_sub["Mappatura"].isin(escludi_voci)]

        piv_act = df_act_cf.pivot_table(index="Mese", columns="Mappatura", values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
        piv_bdg = df_bdg_cf.pivot_table(index="Mese", columns="Mappatura", values="Importo_BDG_kEUR", aggfunc="sum", fill_value=0)

        cats = sorted(list(set(piv_act.columns).union(set(piv_bdg.columns))))
        if not cats:
            return Paragraph("<i>Nessun dato Costi Fissi disponibile.</i>", txt_s)

        h_row1 = [Paragraph("M", header_cell)]
        for c in cats:
            h_row1.extend([Paragraph(c, header_cell), "", ""])
        h_row1.extend([Paragraph("TOT", header_cell), "", ""])

        h_row2 = [""]
        for _ in range(len(cats) + 1):
            h_row2.extend([Paragraph("ACT", header_cell), Paragraph("BDG", header_cell), Paragraph("&Delta;", header_cell)])

        table_data = [h_row1, h_row2]
        mesi_list = [f"{anno_sel}-{i:02d}" for i in range(1, 13)]

        tot_act_cats = {c: 0.0 for c in cats}
        tot_bdg_cats = {c: 0.0 for c in cats}
        tot_act_gen, tot_bdg_gen = 0.0, 0.0

        for m_str in mesi_list:
            m_breve = m_str[5:] if len(m_str) >= 7 else m_str
            row = [Paragraph(m_breve, txt_s)]
            m_act_tot, m_bdg_tot = 0.0, 0.0

            for c in cats:
                v_act = piv_act.loc[m_str, c] if (m_str in piv_act.index and c in piv_act.columns) else 0.0
                v_bdg = piv_bdg.loc[m_str, c] if (m_str in piv_bdg.index and c in piv_bdg.columns) else 0.0
                v_delta = v_bdg - v_act

                tot_act_cats[c] += v_act
                tot_bdg_cats[c] += v_bdg
                m_act_tot += v_act
                m_bdg_tot += v_bdg

                row.extend([
                    Paragraph(formatta_contabile(v_act), num_s),
                    Paragraph(formatta_contabile(v_bdg), num_s),
                    Paragraph(formatta_contabile(v_delta), get_delta_style(v_delta))
                ])

            m_delta_tot = m_bdg_tot - m_act_tot
            tot_act_gen += m_act_tot
            tot_bdg_gen += m_bdg_tot

            row.extend([
                Paragraph(formatta_contabile(m_act_tot), num_b),
                Paragraph(formatta_contabile(m_bdg_tot), num_b),
                Paragraph(formatta_contabile(m_delta_tot), get_delta_style(m_delta_tot, is_bold=True))
            ])
            table_data.append(row)

        tot_row = [Paragraph("TOT", txt_b)]
        for c in cats:
            v_act = tot_act_cats[c]
            v_bdg = tot_bdg_cats[c]
            v_delta = v_bdg - v_act
            tot_row.extend([
                Paragraph(formatta_contabile(v_act), num_b),
                Paragraph(formatta_contabile(v_bdg), num_b),
                Paragraph(formatta_contabile(v_delta), get_delta_style(v_delta, is_bold=True))
            ])

        tot_delta_gen = tot_bdg_gen - tot_act_gen
        tot_row.extend([
            Paragraph(formatta_contabile(tot_act_gen), num_b),
            Paragraph(formatta_contabile(tot_bdg_gen), num_b),
            Paragraph(formatta_contabile(tot_delta_gen), get_delta_style(tot_delta_gen, is_bold=True))
        ])
        table_data.append(tot_row)

        num_cols = 1 + (len(cats) + 1) * 3
        col_w = [25] + [(732 / (num_cols - 1))] * (num_cols - 1)

        t_style = [
            ('BACKGROUND', (0, 0), (-1, 1), colors.HexColor('#1A1A1A')),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor('#CCCCCC')),
            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#F4F4F4')),
            ('SPAN', (0, 0), (0, 1)),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 1.5),
            ('TOPPADDING', (0, 0), (-1, -1), 1.5),
        ]

        col_idx = 1
        for _ in range(len(cats) + 1):
            t_style.append(('SPAN', (col_idx, 0), (col_idx + 2, 0)))
            col_idx += 3

        t = Table(table_data, colWidths=col_w, repeatRows=2)
        t.setStyle(TableStyle(t_style))
        return t

    voci_fisse_escluse_act = ["Vic - Maintenance (Var)", "Insurance", "Intercompany"]
    voci_fisse_escluse_bdg = voci_fisse_escluse_act + ["Spare Parts And Equipments", "Subscription And Associations Fees"]

    sezioni = [
        ("Ritrama S.p.A. - Tutti gli Stabilimenti", df_act_f, df_bdg_f),
        ("Ritrama S.p.A. - Stabilimento di Basiano", df_act_f[df_act_f["Stabilimento"] == "RITR-BAS"], df_bdg_f[df_bdg_f["PLANT"] == "RITR-BAS"]),
        ("Ritrama S.p.A. - Stabilimento di Caponago", df_act_f[df_act_f["Stabilimento"] == "RITR-CAP"], df_bdg_f[df_bdg_f["PLANT"] == "RITR-CAP"]),
        ("Ritrama S.p.A. - Stabilimento di Sassoferrato", df_act_f[df_act_f["Stabilimento"] == "RITR-SAS"], df_bdg_f[df_bdg_f["PLANT"] == "RITR-SAS"]),
        ("Ritrama S.p.A. - Stabilimento di Headquarter", df_act_f[df_act_f["Stabilimento"] == "RITR-HQ"], df_bdg_f[df_bdg_f["PLANT"] == "RITR-HQ"])
    ]

    for i, (titolo, df_act_sub, df_bdg_sub) in enumerate(sezioni):
        story.append(Paragraph(f"<b>{titolo} - Costi Fissi</b>", header_style))
        story.append(Paragraph("1. Costi Fissi - Actual MTD (in k€)", sub_header_style))
        story.append(crea_tabella_actual(df_act_sub, is_vic=False, escludi_voci=voci_fisse_escluse_act))
        story.append(Spacer(1, 10))
        story.append(Paragraph("2. Costi Fissi - Actual vs Budget MTD (in k€)", sub_header_style))
        story.append(crea_tabella_act_vs_bdg(df_act_sub, df_bdg_sub, is_vic=False, escludi_voci=voci_fisse_escluse_bdg))
        story.append(PageBreak())

        story.append(Paragraph(f"<b>{titolo} - Costi Variabili di Manutenzione (VIC)</b>", header_style))
        story.append(Paragraph("1. Costi Variabili per Conto - Actual MTD (in k€)", sub_header_style))
        story.append(crea_tabella_actual(df_act_sub, is_vic=True))
        story.append(Spacer(1, 10))
        story.append(Paragraph("2. Costi Variabili - Actual vs Budget MTD (in k€)", sub_header_style))
        story.append(crea_tabella_act_vs_bdg(df_act_sub, df_bdg_sub, is_vic=True))

        if i < len(sezioni) - 1:
            story.append(PageBreak())

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

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

        if "ACT" in col_str:
            bg_color = '#d9d9d9' if is_tot_col else '#f2f2f2'
            base_style = f'background-color: {bg_color}; color: black;'
        elif "BDG" in col_str:
            bg_color = '#e6e6e6' if is_tot_col else '#ffffff'
            base_style = f'background-color: {bg_color}; color: black;'
        elif "Delta" in col_str:
            bg_color = '#b3d8ff' if is_tot_col else '#e6f2ff'
            base_style = f'background-color: {bg_color};'
        else:
            base_style = ''

        for idx in df.index:
            idx_str = str(idx).upper()
            is_tot_row = "TOTALE" in idx_str
            style = base_style

            if "Delta" in col_str:
                val = df.loc[idx, col]
                if isinstance(val, (int, float)):
                    if val > 0.01:
                        style += ' color: red;'
                    elif val < -0.01:
                        style += ' color: green;'
                    else:
                        style += ' color: black;'

            if is_tot_col or is_tot_row:
                style += ' font-weight: bold;'
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

def to_excel_download(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Budget_Elaborato')
    return output.getvalue()

def elabora_report(df_trans, df_forn):
    df_out = df_trans.copy()
    df_out.columns = df_out.columns.astype(str).str.strip()

    col_desc_trovata = None
    nomi_possibili = [
        "Descrizione", "Testo", "Descrizione transazione", 
        "Testo dell'intestazione del documento", "Testo posizione", "Designazione"
    ]
    
    for col in nomi_possibili:
        if col in df_out.columns:
            col_desc_trovata = col
            break

    df_out["Descrizione"] = df_out[col_desc_trovata].fillna("-").astype(str) if col_desc_trovata else "-"

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

        def estrai_conto_contabile(val):
            match = re.search(r'(\d{2}\.\d{2}\.\d{2}\.\d{3}\.\d{2})', str(val))
            return match.group(1) if match else str(val).split('-')[0].strip()

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
        df_out["Conto Contabile"] = df_out["Valore visualizzato conto"].apply(estrai_conto_contabile)
        df_out["Mappatura"] = df_out["Codice Mappatura Estratto"].map(MAPPATURA_DEFAULT).str.title()
        df_out["Stabilimento"] = df_out["Valore visualizzato conto"].apply(estrai_stab)
        df_out["Codice Fornitore Estratto"] = df_out["Valore visualizzato conto"].apply(estrai_cod_forn)
        
        df_out["Fornitore"] = df_out["Codice Fornitore Estratto"].map(mppa_fornitori)
        df_out["Fornitore"] = df_out["Fornitore"].fillna(
            df_out["Codice Fornitore Estratto"].apply(lambda x: f"Cod. {x}" if x else "Non Definito")
        )

        df_out = df_out[df_out["Mappatura"].notna()].copy()

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
            
            anni_disponibili = sorted([a for a in df_act["Anno"].unique() if a != "nan"], reverse=True)
            anno_sel = st.sidebar.selectbox("Seleziona Anno:", anni_disponibili, index=0)

            df_bdg = processa_budget(df_bdg_raw, anno_ref=int(anno_sel) if anno_sel else 2026)

            plants = sorted(list(set(df_act["Stabilimento"].unique()).union(set(df_bdg["PLANT"].unique()))))
            plant_sel = st.sidebar.selectbox("Seleziona Plant:", ["TUTTI I PLANT"] + plants)

            mappature_totali = sorted(list(set(df_act["Mappatura"].unique()).union(set(df_bdg["Mappatura"].unique()))))
            mappatura_sel = st.sidebar.multiselect("Seleziona Voce di Spesa:", options=mappature_totali, default=[])

            fornitori_totali = sorted(list(set(df_act["Fornitore"].unique()).union(set(df_bdg["Supplier"].unique()))))
            fornitore_sel = st.sidebar.selectbox("Seleziona Fornitore:", ["TUTTI I FORNITORI"] + fornitori_totali)

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

            mesi_disponibili = sorted(list(set(df_act_f["Mese"].unique()).union(set(df_bdg_f["Mese"].unique()))))
            mesi_sel = st.sidebar.multiselect("Seleziona Mesi:", options=mesi_disponibili, default=[])

            if mesi_sel:
                df_act_f_mesi = df_act_f[df_act_f["Mese"].isin(mesi_sel)]
            else:
                df_act_f_mesi = df_act_f.copy()

            tab_dettaglio, tab_act_bdg = st.tabs(["📊 Report Analisi Spese", "⚖️ ACT vs BDG"])

            # ==========================================
            # TAB 1: METRICHE E REPORT TRANSAZIONI
            # ==========================================
            with tab_dettaglio:
                if pagina == "📌 Fixed Costs (Costi Fissi)":
                    df_sezione = df_act_f_mesi[df_act_f_mesi["Mappatura"] != "Vic - Maintenance (Var)"].copy()
                    st.header("📌 Fixed Costs - Dettaglio Analisi Spese")
                    col_raggruppamento = "Mappatura"
                else:
                    df_sezione = df_act_f_mesi[df_act_f_mesi["Mappatura"] == "Vic - Maintenance (Var)"].copy()
                    st.header("⚙️ Variable Costs - Dettaglio Analisi Spese per Conto Contabile")
                    col_raggruppamento = "Conto Contabile" if "Conto Contabile" in df_sezione.columns else "Mappatura"

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

                st.subheader(f"📈 Actual MTD per {col_raggruppamento} (in k€)")
                piv_mensile = df_sezione.pivot_table(index="Mese", columns=col_raggruppamento, values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
                piv_mensile["TOTALE MENSILE"] = piv_mensile.sum(axis=1)
                tot_map = piv_mensile.sum(axis=0)
                tot_map.name = "TOTALE"
                piv_mensile_full = pd.concat([piv_mensile, pd.DataFrame(tot_map).T])
                st.dataframe(piv_mensile_full.style.format(formatta_k_euro), use_container_width=True)

                st.subheader("📊 ACT MTD per Plant (k€)")
                piv_plant = df_sezione.pivot_table(index="Mese", columns="Stabilimento", values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
                st.bar_chart(piv_plant)

                st.markdown("---")
                st.subheader(f"🏢 ACT per Fornitore YTD - {col_raggruppamento} (in k€)")
                piv_ytd = df_sezione.pivot_table(index="Fornitore", columns=col_raggruppamento, values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
                piv_ytd["TOTALE FORNITORE"] = piv_ytd.sum(axis=1)
                piv_ytd = piv_ytd.sort_values(by="TOTALE FORNITORE", ascending=False)
                tot_cat = piv_ytd.sum(axis=0)
                tot_cat.name = "TOTALE CATEGORIA"
                piv_ytd_full = pd.concat([piv_ytd, pd.DataFrame(tot_cat).T])
                st.dataframe(piv_ytd_full.style.format(formatta_k_euro), use_container_width=True)

                st.markdown("---")
                st.subheader("📋 Dettaglio Analitico Transazioni per Fornitore")
                
                cols_dettaglio = [c for c in [col_data, "Conto Contabile", "Mappatura", "Fornitore", "Stabilimento", "Importo_ACT_kEUR", "Descrizione"] if c in df_sezione.columns]
                df_dett = df_sezione[cols_dettaglio].copy()
                if "Importo_ACT_kEUR" in df_dett.columns:
                    df_dett = df_dett.rename(columns={"Importo_ACT_kEUR": "Importo (k€)"})
                
                st.dataframe(
                    df_dett.style.format({"Importo (k€)": formatta_k_euro}),
                    use_container_width=True
                )

                pdf_bytes = genera_pdf_report(df_act_f, df_bdg_f, anno_sel, mese_chiuso_ref)
                st.download_button(
                    label="📄 Scarica Report PDF per Plant Directors",
                    data=pdf_bytes,
                    file_name=f"Report_Costi_{mese_chiuso_ref}.pdf",
                    mime="application/pdf"
                )

            # ==========================================
            # TAB 2: ACT VS BDG + EXCEL BUDGET
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

                    mesi_completi = [f"{anno_sel}-{idx:02d}" for idx in range(1, 13)]
                    piv_act_m = piv_act_m.reindex(mesi_completi, fill_value=0)
                    piv_bdg_m = piv_bdg_m.reindex(mesi_completi, fill_value=0)

                    if "YTD" in tipo_vista:
                        ultimo_mese_act = df_act_m1["Mese"].max() if len(df_act_m1) > 0 else None
                        piv_act_m = piv_act_m.cumsum(axis=0)
                        piv_bdg_m = piv_bdg_m.cumsum(axis=0)

                        if ultimo_mese_act and ultimo_mese_act in piv_act_m.index:
                            idx_ultimo = piv_act_m.index.get_loc(ultimo_mese_act)
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

                else:
                    st.header("⚙️ Variable Costs - Maintenance (VAR)")
                    st.subheader("MTD / YTD Actual vs Budget Totale VIC")

                    df_act_v = df_act_f[df_act_f["Mappatura"] == "Vic - Maintenance (Var)"]
                    df_bdg_v = df_bdg_f[df_bdg_f["Mappatura"] == "Vic - Maintenance (Var)"]

                    piv_act_v_tot = df_act_v.pivot_table(index="Mese", values="Importo_ACT_kEUR", aggfunc="sum", fill_value=0)
                    piv_bdg_v_tot = df_bdg_v.pivot_table(index="Mese", values="Importo_BDG_kEUR", aggfunc="sum", fill_value=0)

                    mesi_completi = [f"{anno_sel}-{idx:02d}" for idx in range(1, 13)]
                    piv_act_v_tot = piv_act_v_tot.reindex(mesi_completi, fill_value=0)
                    piv_bdg_v_tot = piv_bdg_v_tot.reindex(mesi_completi, fill_value=0)

                    if "YTD" in tipo_vista:
                        piv_act_v_tot = piv_act_v_tot.cumsum(axis=0).ffill()
                        piv_bdg_v_tot = piv_bdg_v_tot.cumsum(axis=0).ffill()

                    rows_v = []
                    
                    for idx, m_code in enumerate(MESI_LISTA, 1):
                        m_str = f"{anno_sel}-{idx:02d}"
                        r = {"Mese": f"{m_code} ({m_str})"}

                        v_act_tot = piv_act_v_tot.loc[m_str, "Importo_ACT_kEUR"] if m_str in piv_act_v_tot.index else 0.0
                        v_bdg_tot = piv_bdg_v_tot.loc[m_str, "Importo_BDG_kEUR"] if m_str in piv_bdg_v_tot.index else 0.0

                        r["VIC Manutenzione ACT"] = v_act_tot
                        r["VIC Manutenzione BDG"] = v_bdg_tot
                        r["VIC Manutenzione Delta"] = v_bdg_tot - v_act_tot
                        rows_v.append(r)

                    df_mod3 = pd.DataFrame(rows_v).set_index("Mese")

                    tot_dict_v = {}
                    if "YTD" in tipo_vista:
                        for col in df_mod3.columns:
                            tot_dict_v[col] = df_mod3.iloc[-1][col]
                        tot_dict_v["VIC Manutenzione Delta"] = tot_dict_v["VIC Manutenzione BDG"] - tot_dict_v["VIC Manutenzione ACT"]
                    else:
                        for col in df_mod3.columns:
                            tot_dict_v[col] = df_mod3[col].sum()
                        tot_dict_v["VIC Manutenzione Delta"] = tot_dict_v["VIC Manutenzione BDG"] - tot_dict_v["VIC Manutenzione ACT"]

                    df_tot_mod3 = pd.DataFrame([tot_dict_v], index=["TOTALE ANNO"])
                    df_mod3_full = pd.concat([df_mod3, df_tot_mod3])

                    st.dataframe(
                        df_mod3_full.style.format(formatta_k_euro).apply(colora_colonne_fornitori, axis=None),
                        use_container_width=True
                    )

                st.markdown("---")
                st.subheader("📥 Dettaglio e Download Budget Elaborato")
                st.dataframe(
                    df_bdg_f.style.format({"Importo_BDG_kEUR": formatta_k_euro, "Importo_BDG_Orig": "{:,.2f}"}),
                    use_container_width=True
                )

                excel_data = to_excel_download(df_bdg_f)
                st.download_button(
                    label="📥 Scarica Budget in Excel (.xlsx)",
                    data=excel_data,
                    file_name=f"Budget_Elaborato_{anno_sel}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    except Exception as e:
        st.error(f"Errore durante l'elaborazione dei dati: {e}")