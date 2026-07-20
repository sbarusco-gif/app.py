import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import datetime
from io import BytesIO

# --- CONFIGURAZIONE PAGINA STREAMLIT ---
st.set_page_config(page_title="Nota Spese Tributaria", page_icon="⚖️")

st.title("⚖️ Nota Spese Tributaria Professionale")

# --- INPUT DATI (SIDEBAR) ---
st.sidebar.header("Dati Procedimento")
grado_selezione = st.sidebar.selectbox("Grado", ["I GRADO", "II GRADO"])

if grado_selezione == "I GRADO":
    sede_input = st.sidebar.text_input("Sede (Città)", "BIELLA").upper()
    grado_header = f"DI PRIMO GRADO DI {sede_input}"
    grado_competenza = f"di primo grado di {sede_input.lower()}"
    localita_firma = sede_input.capitalize()
else:
    territorio_input = st.sidebar.text_input("Regione (es. del Veneto, della Lombardia)", "del Veneto")
    grado_header = f"DI SECONDO GRADO {territorio_input.upper()}"
    grado_competenza = f"di secondo grado {territorio_input.lower()}"
    localita_firma = "Padova" # Default per lo studio

rgr = st.sidebar.text_input("R.G.R. n.", "123/2024")

indeterminato = st.sidebar.checkbox("Valore Indeterminato")

if indeterminato:
    complessita = st.sidebar.selectbox("Complessità della causa", ["Bassa", "Media (Standard)", "Alta"])
    if complessita == "Bassa":
        valore_per_calcolo = 25000 
        testo_valore = "Valore Indeterminato (complessità bassa: scaglione € 5.201 - € 26.000)"
    elif complessita == "Media (Standard)":
        valore_per_calcolo = 250000 
        testo_valore = "Valore Indeterminato (complessità media: scaglione € 26.001 - € 260.000)"
    else:
        valore_per_calcolo = 500000 
        testo_valore = "Valore Indeterminato (complessità alta: scaglione € 260.001 - € 520.000)"
else:
    valore_lite = st.sidebar.number_input("Valore della lite (€)", min_value=0.0, value=15000.0)
    valore_per_calcolo = valore_lite
    testo_valore = "" # Verrà riempito dallo scaglione sotto

# --- LOGICA CALCOLO PARAMETRI (DM 147/2022) ---
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
if not indeterminato: testo_valore = f"da € {scaglione_testo}"

fase_studio, fase_intro, fase_dec = p[0], p[1], p[3]
comp_tabellare = fase_studio + fase_intro + fase_dec
spese_gen = comp_tabellare * 0.15
cpa = (comp_tabellare + spese_gen) * 0.04
imponibile = comp_tabellare + spese_gen + cpa
iva = imponibile * 0.22
totale_liquidabile = imponibile + iva

st.sidebar.header("Dati Cliente")
cliente = st.sidebar.text_input("Nome Cliente", "Federica Benzi")
luogo_nascita = st.sidebar.text_input("Luogo di nascita", "Montevideo (EE)")
data_nascita = st.sidebar.text_input("Data di nascita", "21/03/1959")
cf_cliente = st.sidebar.text_input("C.F. Cliente", "BNZFRC59C61Z613C")
residenza = st.sidebar.text_input("Residenza", "Sagliano Micca (BI), via Grosso n. 8")

# --- GENERAZIONE WORD ---
def create_professional_word():
    doc = Document()
    
    # STILE: CALIBRI 12, INTERLINEA 1.5
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(12)
    style.paragraph_format.line_spacing = 1.5
    style.paragraph_format.space_after = Pt(6)

    # 1. INTESTAZIONE
    h1 = doc.add_paragraph()
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = h1.add_run(f"ALLA CORTE DI GIUSTIZIA TRIBUTARIA {grado_header}")
    run1.bold = True

    doc.add_paragraph("* * *").alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    h2 = doc.add_paragraph()
    h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = h2.add_run("Nota spese ex art. 15, D.Lgs. 546/1992")
    run2.bold = True

    doc.add_paragraph("* * *").alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 2. CORPO TESTO
    p_intro = doc.add_paragraph()
    p_intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_intro.add_run("Il ").bold = False
    p_intro.add_run("prof. dott. Mario Rovetti").bold = True
    p_intro.add_run(" (C.F. RVTMRA63T23A859T), iscritto all’Albo dei Dottori Commercialisti e degli Esperti Contabili di Biella al n. 106/A ed il ")
    p_intro.add_run("dott. Sebastiano Barusco").bold = True
    p_intro.add_run(" (C.F. BRSSST66C17L736M), iscritto all’Albo dei Dottori Commercialisti e degli Esperti Contabili di Padova al n. 667/A ed, con studio in Padova, via Cavazzana 5 (")
    
    run_und = p_intro.add_run("Barusco Rovetti & Associati, tel. 049-8752918, e-mail: segreteria@baruscorovetti.it, indirizzi di posta elettronica certificata: sebastiano.barusco@legalmail.it e studiorovetti@pec.studiorovetti.it")
    run_und.underline = True
    
    p_intro.add_run(f"), in qualità di rappresentanti, difensori e domiciliatari, come da mandato in atti, della signora ")
    p_intro.add_run(f"{cliente}").bold = True
    p_intro.add_run(f", nata a {luogo_nascita} il {data_nascita}, C.F. {cf_cliente}, residente in {residenza}")

    doc.add_paragraph("\nDEPOSITANO").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("la presente nota spese:")
    doc.add_paragraph(f"Competenza: Corte di giustizia tributaria {grado_competenza}")
    doc.add_paragraph(f"Valore della causa: {testo_valore}")

    # 3. TABELLA COMPENSI
    table = doc.add_table(rows=1, cols=2)
    table.columns[0].width = Cm(10)
    table.columns[1].width = Cm(5)
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Fase'
    hdr_cells[1].text = 'Compenso'
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

    # 4. PROSPETTO FINALE
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
        if desc in ["Totale imponibile", "IPOTESI DI COMPENSO LIQUIDABILE"]:
            r[0].paragraphs[0].runs[0].bold = True
            r[1].paragraphs[0].runs[0].bold = True

    doc.add_paragraph("\nCon osservanza.")
    doc.add_paragraph(f"{localita_firma}, {datetime.date.today().strftime('%d/%m/%Y')}")
    doc.add_paragraph("\nSebastiano Barusco - Firmato digitalmente")

    target = BytesIO()
    doc.save(target)
    return target.getvalue()

# --- INTERFACCIA ---
if st.button("Genera Nota Spese Word"):
    file_bytes = create_professional_word()
    st.download_button(
        label="📥 Scarica Documento Word",
        data=file_bytes,
        file_name=f"Nota_Spese_{cliente.replace(' ','_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
