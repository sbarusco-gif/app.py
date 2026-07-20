import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import datetime
from io import BytesIO

# --- DATABASE SEDI ---
REGIONI_II_GRADO = [
    "dell'Abruzzo", "della Basilicata", "della Calabria", "della Campania", 
    "dell'Emilia-Romagna", "del Friuli-Venezia Giulia", "del Lazio", "della Liguria", 
    "della Lombardia", "delle Marche", "del Molise", "del Piemonte", "della Puglia", 
    "della Sardegna", "della Sicilia", "della Toscana", "del Trentino-Alto Adige", 
    "dell'Umbria", "della Valle d'Aosta", "del Veneto"
]

CITTA_I_GRADO = [
    "Agrigento", "Alessandria", "Ancona", "Aosta", "Arezzo", "Ascoli Piceno", "Asti", "Avellino", "Bari", 
    "Barletta", "Belluno", "Benevento", "Bergamo", "Biella", "Bologna", "Bolzano", "Brescia", "Brindisi", 
    "Cagliari", "Caltanissetta", "Campobasso", "Caserta", "Cassino", "Catania", "Catanzaro", "Chieti", 
    "Como", "Cosenza", "Cremona", "Crotone", "Cuneo", "Enna", "Fermo", "Ferrara", "Firenze", "Foggia", 
    "Forlì", "Frosinone", "Genova", "Gorizia", "Grosseto", "Imperia", "Isernia", "L'Aquila", "La Spezia", 
    "Latina", "Lecce", "Lecco", "Livorno", "Lodi", "Lucca", "Macerata", "Mantova", "Massa Carrara", 
    "Matera", "Messina", "Milano", "Modena", "Monza", "Napoli", "Novara", "Nuoro", "Oristano", "Padova", 
    "Palermo", "Parma", "Pavia", "Perugia", "Pesaro", "Pescara", "Piacenza", "Pisa", "Pistoia", "Pordenone", 
    "Potenza", "Prato", "Ragusa", "Ravenna", "Reggio Calabria", "Reggio Emilia", "Rieti", "Rimini", "Roma", 
    "Rovigo", "Salerno", "Sassari", "Savona", "Siena", "Siracusa", "Sondrio", "Taranto", "Teramo", "Terni", 
    "Torino", "Trani", "Trapani", "Trento", "Treviso", "Trieste", "Udine", "Varese", "Venezia", "Verbano-Cusio-Ossola", 
    "Vercelli", "Verona", "Vibo Valentia", "Vicenza", "Viterbo"
]

# --- CONFIGURAZIONE PAGINA ---
st.set_page_config(page_title="Nota Spese Tributaria", page_icon="⚖️")
st.title("⚖️ Nota Spese Tributaria Professionale")

# --- SIDEBAR: DATI PROCEDIMENTO ---
st.sidebar.header("Dati Procedimento")
grado_selezione = st.sidebar.selectbox("Grado di Giudizio", ["I GRADO", "II GRADO"])

if grado_selezione == "I GRADO":
    sede_citta = st.sidebar.selectbox("Sede della Corte (Città)", sorted(CITTA_I_GRADO))
    grado_header = f"DI PRIMO GRADO DI {sede_citta.upper()}"
    grado_competenza = f"di primo grado di {sede_citta.lower()}"
else:
    regione_scelta = st.sidebar.selectbox("Sede della Corte (Regione)", sorted(REGIONI_II_GRADO))
    grado_header = f"DI SECONDO GRADO {regione_scelta.upper()}"
    grado_competenza = f"di secondo grado {regione_scelta.lower()}"

rgr = st.sidebar.text_input("R.G.R. n.", "123/2024")

# GESTIONE VALORE
indeterminato = st.sidebar.checkbox("Valore Indeterminato")
if indeterminato:
    complessita = st.sidebar.selectbox("Complessità", ["Bassa", "Media", "Alta"])
    valore_per_calcolo = 25000 if complessita == "Bassa" else 250000 if complessita == "Media" else 500000
    testo_valore = f"Valore Indeterminato (complessità {complessita.lower()})"
else:
    valore_lite = st.sidebar.number_input("Valore della lite (€)", min_value=0.0, value=15000.0)
    valore_per_calcolo = valore_lite
    testo_valore = ""

# --- LOGICA CALCOLO ---
def get_params(v):
    if v <= 1100: return (270, 270, 140, 340, "fino a € 1.100")
    elif v <= 5200: return (485, 405, 270, 610, "da € 1.101 a € 5.200")
    elif v <= 26000: return (1150, 875, 675, 1350, "da € 5.201 a € 26.000")
    elif v <= 52000: return (1960, 1285, 1080, 2365, "da € 26.001 a € 52.000")
    elif v <= 260000: return (3375, 2230, 1620, 3850, "da € 52.001 a € 260.000")
    elif v <= 520000: return (5065, 3310, 2430, 5805, "da € 260.001 a € 520.000")
    else: return (7225, 4795, 3445, 8710, "oltre € 520.000")

p = get_params(valore_per_calcolo)
scaglione_esteso = p[4]
if not indeterminato: testo_valore = scaglione_esteso

fase_studio, fase_intro, fase_dec = p[0], p[1], p[3]
comp_tabellare = fase_studio + fase_intro + fase_dec
spese_gen = comp_tabellare * 0.15
cpa = (comp_tabellare + spese_gen) * 0.04
imponibile = comp_tabellare + spese_gen + cpa
iva = imponibile * 0.22
totale_liquidabile = imponibile + iva

# --- DATI CLIENTE ---
st.sidebar.header("Dati Cliente")
cliente = st.sidebar.text_input("Nome Cliente", "Federica Benzi")
luogo_nascita = st.sidebar.text_input("Luogo di nascita", "Montevideo (EE)")
data_nascita = st.sidebar.text_input("Data di nascita", "21/03/1959")
cf_cliente = st.sidebar.text_input("C.F. Cliente", "BNZFRC59C61Z613C")
residenza = st.sidebar.text_input("Residenza", "Sagliano Micca (BI), via Grosso n. 8")

# --- GENERAZIONE WORD ---
def create_professional_word():
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(12)
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
    h2.add_run("Nota spese ex art. 15, D.Lgs. 546/1992").bold = True
    doc.add_paragraph("* * *").alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 2. CORPO
    p_intro = doc.add_paragraph()
    p_intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_intro.add_run("Il ").bold = False
    p_intro.add_run("prof. dott. Mario Rovetti").bold = True
    p_intro.add_run(" (C.F. RVTMRA63T23A859T) ed il ")
    p_intro.add_run("dott. Sebastiano Barusco").bold = True
    p_intro.add_run(" (C.F. BRSSST66C17L736M), con studio in Padova, via Cavazzana 5 (")
    run_und = p_intro.add_run("Barusco Rovetti & Associati, tel. 049-8752918, e-mail: segreteria@baruscorovetti.it, indirizzi PEC: sebastiano.barusco@legalmail.it e studiorovetti@pec.studiorovetti.it")
    run_und.underline = True
    p_intro.add_run(f"), in qualità di difensori della signora ")
    p_intro.add_run(f"{cliente}").bold = True
    p_intro.add_run(f", nata a {luogo_nascita} il {data_nascita}, C.F. {cf_cliente}, residente in {residenza}")

    doc.add_paragraph("\nDEPOSITANO").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Competenza: Corte di giustizia tributaria {grado_competenza}")
    doc.add_paragraph(f"Valore della causa: {testo_valore}")

    # 3. TABELLA COMPENSI
    table = doc.add_table(rows=1, cols=2)
    table.columns[0].width = Cm(10)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Fase'; hdr_cells[1].text = 'Compenso'
    hdr_cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    fasi = [("Fase di studio della controversia:", f"€ {fase_studio:,.2f}"), ("Fase introduttiva del giudizio:", f"€ {fase_intro:,.2f}"), 
            ("Fase decisionale:", f"€ {fase_dec:,.2f}"), ("Compenso tabellare", f"€ {comp_tabellare:,.2f}")]
    for n, v in fasi:
        row = table.add_row().cells
        row[0].text = n; row[1].text = v
        row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if n == "Compenso tabellare": row[0].paragraphs[0].runs[0].bold = True; row[1].paragraphs[0].runs[0].bold = True

    doc.add_paragraph("\nPROSPETTO FINALE").bold = True
    table2 = doc.add_table(rows=0, cols=2)
    finali = [("Compenso tabellare", f"€ {comp_tabellare:,.2f}"), ("Spese generali (15% sul compenso totale)", f"€ {spese_gen:,.2f}"),
              ("Cassa Previdenza (4%)", f"€ {cpa:,.2f}"), ("Totale imponibile", f"€ {imponibile:,.2f}"),
              ("IVA 22% su Imponibile", f"€ {iva:,.2f}"), ("IPOTESI DI COMPENSO LIQUIDABILE", f"€ {totale_liquidabile:,.2f}")]
    for n, v in finali:
        r = table2.add_row().cells
        r[0].text = n; r[1].text = v
        r[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if n in ["Totale imponibile", "IPOTESI DI COMPENSO LIQUIDABILE"]: r[0].paragraphs[0].runs[0].bold = True; r[1].paragraphs[0].runs[0].bold = True

    doc.add_paragraph("\nCon osservanza.")
    
    # --- LOGICA DATA IN LETTERE ---
    mesi = {1: "Gennaio", 2: "Febbraio", 3: "Marzo", 4: "Aprile", 5: "Maggio", 6: "Giugno",
            7: "Luglio", 8: "Agosto", 9: "Settembre", 10: "Ottobre", 11: "Novembre", 12: "Dicembre"}
    oggi = datetime.date.today()
    data_estesa = f"Padova, {oggi.day} {mesi[oggi.month]} {oggi.year}"
    
    doc.add_paragraph(data_estesa)
    
    # Firma con firmato digitalmente a capo
    p_firma = doc.add_paragraph("\nSebastiano Barusco")
    p_firma.add_run("\nFirmato digitalmente")

    target = BytesIO()
    doc.save(target)
    return target.getvalue()

# --- INTERFACCIA ---
if st.button("Genera Word"):
    st.download_button(label="📥 Scarica Word", data=create_professional_word(), 
                       file_name=f"Nota_Spese_{cliente.replace(' ','_')}.docx")
