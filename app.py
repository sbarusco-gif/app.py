import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import datetime
from io import BytesIO

# --- CONFIGURAZIONE PAGINA ---
st.set_page_config(page_title="Nota Spese Tributaria - Indeterminato", page_icon="⚖️")

st.title("⚖️ Nota Spese Tributaria Professionale")

# --- INPUT DATI (SIDEBAR) ---
st.sidebar.header("Dati Procedimento")
sede = st.sidebar.text_input("Sede Corte", "BIELLA").upper()
grado = st.sidebar.selectbox("Grado", ["I GRADO", "II GRADO"])
rgr = st.sidebar.text_input("R.G.R. n.", "123/2024")

# GESTIONE VALORE INDETERMINATO
indeterminato = st.sidebar.checkbox("Valore Indeterminato")

if indeterminato:
    complessita = st.sidebar.selectbox("Complessità della causa", ["Bassa", "Media (Standard)", "Alta"])
    if complessita == "Bassa":
        valore_per_calcolo = 25000 # Scaglione 5k-26k
        testo_valore = "Valore Indeterminato (complessità bassa: scaglione € 5.201 - € 26.000)"
    elif complessita == "Media (Standard)":
        valore_per_calcolo = 250000 # Scaglione 26k-260k
        testo_valore = "Valore Indeterminato (complessità media: scaglione € 26.001 - € 260.000)"
    else:
        valore_per_calcolo = 500000 # Scaglione 260k-520k
        testo_valore = "Valore Indeterminato (complessità alta: scaglione € 260.001 - € 520.000)"
else:
    valore_lite = st.sidebar.number_input("Valore della lite (€)", min_value=0.0, value=15000.0)
    valore_per_calcolo = valore_lite
    testo_valore = f"€ {valore_lite:,.2f}"

st.sidebar.header("Dati Cliente")
cliente = st.sidebar.text_input("Nome Cliente", "Federica Benzi")
luogo_nascita = st.sidebar.text_input("Luogo di nascita", "Montevideo (EE)")
data_nascita = st.sidebar.text_input("Data di nascita", "21/03/1959")
cf_cliente = st.sidebar.text_input("C.F. Cliente", "BNZFRC59C61Z613C")
residenza = st.sidebar.text_input("Residenza", "Sagliano Micca (BI), via Grosso n. 8")

# --- LOGICA CALCOLO PARAMETRI ---
def get_params(v):
    if v <= 1100: return (270, 270, 140, 340, "fino a € 1.100")
    elif v <= 5200: return (485, 405, 270, 610, "da € 1.101 a € 5.200")
    elif v <= 26000: return (1150, 875, 675, 1350, "da € 5.201 a € 26.000")
    elif v <= 52000: return (1960, 1285, 1080, 2365, "da € 26.001 a € 52.000")
    elif v <= 260000: return (3375, 2230, 1620, 3850, "da € 52.001 a € 260.000")
    elif v <= 520000: return (5065, 3310, 2430, 5805, "da € 260.001 a € 520.000")
    else: return (7225, 4795, 3445, 8710, "oltre € 520.000")

p = get_params(valore_per_calcolo)
scaglione_testo = p[4]
fase_studio, fase_intro, fase_dec = p[0], p[1], p[3]
comp_tabellare = fase_studio + fase_intro + fase_dec

# Calcoli Fiscali
spese_gen = comp_tabellare * 0.15
cpa = (comp_tabellare + spese_gen) * 0.04
imponibile = comp_tabellare + spese_gen + cpa
iva = imponibile * 0.22
totale_liquidabile = imponibile + iva

# --- GENERAZIONE WORD (IDENTICO AL MODELLO) ---
def create_professional_word():
    doc = Document()
    font_name = 'Calibri'
    
    # 1. INTESTAZIONE
    h1 = doc.add_paragraph()
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = h1.add_run(f"ALLA CORTE DI GIUSTIZIA TRIBUTARIA DI {grado} DI {sede}")
    r1.bold = True
    r1.font.name = font_name
    
    for _ in range(2):
        p_ast = doc.add_paragraph("* * *")
        p_ast.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if _ == 0:
            h2 = doc.add_paragraph()
            h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r2 = h2.add_run("Nota spese ex art. 15, D.Lgs. 546/1992")
            r2.bold = True
            r2.font.name = font_name

    # 2. CORPO
    p_intro = doc.add_paragraph()
    p_intro.alignment = WD_ALIGN_PARAGRAPH.BOTH
    p_intro.add_run("Il ").bold = False
    p_intro.add_run("prof. dott. Mario Rovetti").bold = True
    p_intro.add_run(" (C.F. RVTMRA63T23A859T), iscritto all’Albo dei Dottori Commercialisti e degli Esperti Contabili di Biella al n. 106/A ed il ")
    p_intro.add_run("dott. Sebastiano Barusco").bold = True
    p_intro.add_run(" (C.F. BRSSST66C17L736M), iscritto all’Albo dei Dottori Commercialisti e degli Esperti Contabili di Padova al n. 667/A ed, con studio in Padova, via Cavazzana 5 (")
    
    # Sottolineatura contatti
    run_und = p_intro.add_run("Barusco Rovetti & Associati, tel. 049-8752918, e-mail: segreteria@baruscorovetti.it, indirizzi di posta elettronica certificata: sebastiano.barusco@legalmail.it e studiorovetti@pec.studiorovetti.it")
    run_und.underline = True
    
    p_intro.add_run(f"), in qualità di rappresentanti, difensori e domiciliatari, come da mandato in atti, della signora ")
    p_intro.add_run(f"{cliente}").bold = True
    p_intro.add_run(f", nata a {luogo_nascita} il {data_nascita}, C.F. {cf_cliente}, residente in {residenza}")

    doc.add_paragraph("\nDEPOSITANO").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("la presente nota spese:")
    doc.add_paragraph(f"Competenza: Corte di giustizia tributaria di {grado.lower()}")
    doc.add_paragraph(f"Valore della causa: {testo_valore}")

    # 3. TABELLE
    # Tabella Compensi
    table = doc.add_table(rows=1, cols=2)
    table.columns[0].width = Cm(10)
    table.columns[1].width = Cm(5)
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Fase'
    hdr_cells[1].text = 'Compenso'
    hdr_cells[0].paragraphs[0].runs[0].bold = True
    hdr_cells[1].paragraphs[0].runs[0].bold = True
    hdr_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    fasi_list = [
        ("Fase di studio della controversia:", f"€ {fase_studio:,.2f}"),
        ("Fase introduttiva del giudizio:", f"€ {fase_intro:,.2f}"),
        ("Fase decisionale:", f"€ {fase_dec:,.2f}"),
        ("Compenso tabellare", f"€ {comp_tabellare:,.2f}")
    ]

    for f_n, f_v in fasi_list:
        row = table.add_row().cells
        row[0].text = f_n
        row[1].text = f_v
        row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if f_n == "Compenso tabellare":
            row[0].paragraphs[0].runs[0].bold = True
            row[1].paragraphs[0].runs[0].bold = True

    # Prospetto Finale
    doc.add_paragraph("\nPROSPETTO FINALE").bold = True
    table2 = doc.add_table(rows=0, cols=2)
    final_rows = [
        ("Compenso tabellare", f"€ {comp_tabellare:,.2f}"),
        ("Spese generali (15% sul compenso totale)", f"€ {spese_gen:,.2f}"),
        ("Cassa Previdenza (4%)", f"€ {cpa:,.2f}"),
        ("Totale imponibile", f"€ {imponibile:,.2f}"),
        ("IVA 22% su Imponibile", f"€ {iva:,.2f}"),
        ("IPOTESI DI COMPENSO LIQUIDABILE", f"€ {totale_liquidabile:,.2f}")
    ]

    for desc, val in final_rows:
        r = table2.add_row().cells
        r[0].text = desc
        r[1].text = val
        r[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if "TOTALE" in desc.upper() or "IPOTESI" in desc.upper():
            r[0].paragraphs[0].runs[0].bold = True
            r[1].paragraphs[0].runs[0].bold = True

    doc.add_paragraph("\nCon osservanza.")
    doc.add_paragraph("\nSebastiano Barusco - Firmato digitalmente")

    # Footer paginazione
    section = doc.sections[0]
    footer = section.footer
    p_footer = footer.paragraphs[0]
    p_footer.text = "\t\tPagina 1 di 1"
    p_footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    target = BytesIO()
    doc.save(target)
    return target.getvalue()

# --- INTERFACCIA ---
st.info(f"Scaglione applicato: {scaglione_testo}")
if st.button("Genera Documento Word"):
    file_bytes = create_professional_word()
    st.download_button(
        label="📥 Scarica Word",
        data=file_bytes,
        file_name=f"Nota_Spese_{cliente.replace(' ','_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
