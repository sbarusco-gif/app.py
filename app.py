import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime
from io import BytesIO

# --- DATABASE SEDI ---
REGIONI_II_GRADO = ["dell'Abruzzo", "della Basilicata", "della Calabria", "della Campania", "dell'Emilia-Romagna", "del Friuli-Venezia Giulia", "del Lazio", "della Liguria", "della Lombardia", "delle Marche", "del Molise", "del Piemonte", "della Puglia", "della Sardegna", "della Sicilia", "della Toscana", "del Trentino-Alto Adige", "dell'Umbria", "della Valle d'Aosta", "del Veneto"]
CITTA_I_GRADO = ["Agrigento", "Alessandria", "Ancona", "Aosta", "Arezzo", "Ascoli Piceno", "Asti", "Avellino", "Bari", "Barletta", "Belluno", "Benevento", "Bergamo", "Biella", "Bologna", "Bolzano", "Brescia", "Brindisi", "Cagliari", "Caltanissetta", "Campobasso", "Caserta", "Cassino", "Catania", "Catanzaro", "Chieti", "Como", "Cosenza", "Cremona", "Crotone", "Cuneo", "Enna", "Fermo", "Ferrara", "Firenze", "Foggia", "Forlì", "Frosinone", "Genova", "Gorizia", "Grosseto", "Imperia", "Isernia", "L'Aquila", "La Spezia", "Latina", "Lecce", "Lecco", "Livorno", "Lodi", "Lucca", "Macerata", "Mantova", "Massa Carrara", "Matera", "Messina", "Milano", "Modena", "Monza", "Napoli", "Novara", "Nuoro", "Oristano", "Padova", "Palermo", "Parma", "Pavia", "Perugia", "Pesaro", "Pescara", "Piacenza", "Pisa", "Pistoia", "Pordenone", "Potenza", "Prato", "Ragusa", "Ravenna", "Reggio Calabria", "Reggio Emilia", "Rieti", "Rimini", "Roma", "Rovigo", "Salerno", "Sassari", "Savona", "Siena", "Siracusa", "Sondrio", "Taranto", "Teramo", "Terni", "Torino", "Trani", "Trapani", "Trento", "Treviso", "Trieste", "Udine", "Varese", "Venezia", "Verbano-Cusio-Ossola", "Vercelli", "Verona", "Vibo Valentia", "Vicenza", "Viterbo"]

# --- FUNZIONI SUPPORTO ---
def get_params(v):
    if v <= 1100: return (270.0, 270.0, 140.0, 340.0, "fino a € 1.100")
    elif v <= 5200: return (485.0, 405.0, 270.0, 610.0, "da € 1.101 a € 5.200")
    elif v <= 26000: return (1150.0, 875.0, 675.0, 1350.0, "da € 5.201 a € 26.000")
    elif v <= 52000: return (1960.0, 1285.0, 1080.0, 2365.0, "da € 26.001 a € 52.000")
    elif v <= 260000: return (3375.0, 2230.0, 1620.0, 3850.0, "da € 52.001 a € 260.000")
    else: return (5065.0, 3310.0, 2430.0, 5805.0, "da € 260.001 a € 520.000")

def set_cell_shading(cell, color):
    tc_pr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color)
    tc_pr.append(shd)

# --- CONFIGURAZIONE ---
st.set_page_config(page_title="Nota Spese Pro", page_icon="⚖️")
st.title("⚖️ Nota Spese Tributaria Professionale")

# --- SIDEBAR: PROCEDIMENTO ---
st.sidebar.header("Procedimento")
grado_sel = st.sidebar.selectbox("Grado", ["I GRADO", "II GRADO"])
if grado_sel == "I GRADO":
    sede = st.sidebar.selectbox("Città", sorted(CITTA_I_GRADO))
    g_h, g_c = f"DI PRIMO GRADO DI {sede.upper()}", f"di Primo Grado di {sede}"
else:
    sede = st.sidebar.selectbox("Regione", sorted(REGIONI_II_GRADO))
    g_h, g_c = f"DI SECONDO GRADO {sede.upper()}", f"di Secondo Grado {sede}"

rgr = st.sidebar.text_input("R.G.R. n.", "123/2024")
sez = st.sidebar.text_input("Sezione", "")
usa_u = st.sidebar.checkbox("Includere data udienza?")
dt_u = st.sidebar.date_input("Udienza", datetime.date.today()).strftime("%d.%m.%Y") if usa_u else ""

indet = st.sidebar.checkbox("Valore Indeterminato")
if indet:
    compl = st.sidebar.selectbox("Complessità", ["Bassa", "Media", "Alta"])
    val_calc = 25000.0 if compl == "Bassa" else 250000.0 if compl == "Media" else 500000.0
    txt_val = f"Indeterminato (complessità {compl.lower()})"
else:
    val_lite = st.sidebar.number_input("Valore della lite (€)", min_value=0.0, value=15000.0)
    val_calc = float(val_lite)
    txt_val = ""

# --- SPESE ESENTI ---
cut_importo = st.sidebar.number_input("Contributo Unificato (C.U.T.) €", min_value=0.0, value=0.0, step=10.0)

# --- SIDEBAR: CLIENTE ---
st.sidebar.header("Cliente")
genere = st.sidebar.radio("Sesso del cliente", ["Femminile", "Maschile"])
cli = st.sidebar.text_input("Nome e Cognome", "Federica Benzi")
ln = st.sidebar.text_input("Luogo di nascita", "Montevideo (EE)")
dn = st.sidebar.text_input("Data di nascita", "21/03/1959")
cf = st.sidebar.text_input("Codice Fiscale", "BNZFRC59C61Z613C")
res = st.sidebar.text_input("Residenza", "Sagliano Micca (BI), via Grosso n. 8")

prep_t = "della signora" if genere == "Femminile" else "del signor"
nasc_c = "nata" if genere == "Femminile" else "nato"

# --- GENERATORE DOCUMENTO ---
def create_doc():
    doc = Document()
    style = doc.styles['Normal']
    style.font.name, style.font.size = 'Calibri', Pt(12)
    style.paragraph_format.line_spacing, style.paragraph_format.space_after = 1.5, Pt(6)

    # Calcoli
    p = get_params(val_calc)
    scaglione_lab = txt_val if txt_val != "" else p[4]
    c_st, c_in, c_de = float(p[0]), float(p[1]), float(p[3])
    comp_t = c_st + c_in + c_de
    s_ge = comp_t * 0.15
    cpa = (comp_t + s_ge) * 0.04
    imp = comp_t + s_ge + cpa
    iva = imp * 0.22
    tot_liq = imp + iva + cut_importo

    # Intestazione
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run(f"ALLA CORTE DI GIUSTIZIA TRIBUTARIA {g_h}")
    r.bold = True
    if sez: h.add_run(f"\nSEZIONE {sez.upper()}").bold = True
    if dt_u: h.add_run(f"\nUDIENZA DEL {dt_u}").bold = True
    
    doc.add_paragraph("* * *").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("Nota spese ex art. 15, D.Lgs. 546/1992").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("* * *").alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Corpo
    p_b = doc.add_paragraph()
    p_b.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_b.add_run("Il prof. dott. Mario Rovetti ed il dott. Sebastiano Barusco con studio in Padova, via Cavazzana 5 (Barusco Rovetti & Associati, tel. 049-8752918, PEC: sebastiano.barusco@legalmail.it), difensori ")
    p_b.add_run(f"{prep_t} {cli}").bold = True
    p_b.add_run(f", {nasc_c} a {ln} il {dn}, C.F. {cf}, residente in {res}")

    doc.add_paragraph("\nDEPOSITANO").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("la seguente nota spese:").alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph(f"Competenza: Corte di Giustizia Tributaria {g_c}")

    # Tabella Professionale
    table = doc.add_table(rows=0, cols=2)
    table.style = 'Table Grid'
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def add_row(label, val_str, is_bold=False, bg=None):
        row = table.add_row().cells
        row[0].text = label
        row[1].text = val_str
        row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if is_bold:
            row[0].paragraphs[0].runs[0].bold = True
            row[1].paragraphs[0].runs[0].bold = True
        if bg:
            set_cell_shading(row[0], bg)
            set_cell_shading(row[1], bg)

    add_row("VALORE DELLA CAUSA", scaglione_lab, True, "E7E7E7")
    add_row("Fase di studio della controversia", f"€ {c_st:,.2f}")
    add_row("Fase introduttiva del giudizio", f"€ {c_in:,.2f}")
    add_row("Fase decisionale", f"€ {c_de:,.2f}")
    add_row("COMPENSO TABELLARE", f"€ {comp_t:,.2f}", True, "F2F2F2")
    add_row("Spese generali (15% su compenso)", f"€ {s_ge:,.2f}")
    add_row("C.P.A. (4% su imponibile)", f"€ {cpa:,.2f}")
    add_row("TOTALE IMPONIBILE", f"€ {imp:,.2f}", True)
    add_row("I.V.A. (22%)", f"€ {iva:,.2f}")
    if cut_importo > 0:
        add_row("Contributo Unificato Tributario (C.U.T.)", f"€ {cut_importo:,.2f}")
    add_row("IPOTESI DI COMPENSO LIQUIDABILE", f"€ {tot_liq:,.2f}", True, "D9D9D9")

    # Firma
    mesi = {1:"Gennaio", 2:"Febbraio", 3:"Marzo", 4:"Aprile", 5:"Maggio", 6:"Giugno", 7:"Luglio", 8:"Agosto", 9:"Settembre", 10:"Ottobre", 11:"Novembre", 12:"Dicembre"}
    oggi = datetime.date.today()
    doc.add_paragraph(f"\nPadova, {oggi.day} {mesi[oggi.month]} {oggi.year}")
    f = doc.add_paragraph("\nSebastiano Barusco")
    f.add_run("\nFirmato digitalmente")

    out = BytesIO()
    doc.save(out)
    return out.getvalue()

# --- INTERFACCIA STREAMLIT ---
if st.button("Genera Nota Spese"):
    try:
        st.download_button("📥 Scarica Word", create_doc(), f"Nota_{cli.replace(' ','_')}.docx")
    except Exception as e:
        st.error(f"Errore: {e}")
