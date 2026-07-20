import streamlit as st
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime
from io import BytesIO

# --- DATABASE SEDI ---
REGIONI_II_GRADO = ["dell'Abruzzo", "della Basilicata", "della Calabria", "della Campania", "dell'Emilia-Romagna", "del Friuli-Venezia Giulia", "del Lazio", "della Liguria", "della Lombardia", "delle Marche", "del Molise", "del Piemonte", "della Puglia", "della Sardegna", "della Sicilia", "della Toscana", "del Trentino-Alto Adige", "dell'Umbria", "della Valle d'Aosta", "del Veneto"]
CITTA_I_GRADO = ["Agrigento", "Alessandria", "Ancona", "Aosta", "Arezzo", "Ascoli Piceno", "Asti", "Avellino", "Bari", "Barletta", "Belluno", "Benevento", "Bergamo", "Biella", "Bologna", "Bolzano", "Brescia", "Brindisi", "Cagliari", "Caltanissetta", "Campobasso", "Caserta", "Cassino", "Catania", "Catanzaro", "Chieti", "Como", "Cosenza", "Cremona", "Crotone", "Cuneo", "Enna", "Fermo", "Ferrara", "Firenze", "Foggia", "Forlì", "Frosinone", "Genova", "Gorizia", "Grosseto", "Imperia", "Isernia", "L'Aquila", "La Spezia", "Latina", "Lecce", "Lecco", "Livorno", "Lodi", "Lucca", "Macerata", "Mantova", "Massa Carrara", "Matera", "Messina", "Milano", "Modena", "Monza", "Napoli", "Novara", "Nuoro", "Oristano", "Padova", "Palermo", "Parma", "Pavia", "Perugia", "Pesaro", "Pescara", "Piacenza", "Pisa", "Pistoia", "Pordenone", "Potenza", "Prato", "Ragusa", "Ravenna", "Reggio Calabria", "Reggio Emilia", "Rieti", "Rimini", "Roma", "Rovigo", "Salerno", "Sassari", "Savona", "Siena", "Siracusa", "Sondrio", "Taranto", "Teramo", "Terni", "Torino", "Trani", "Trapani", "Trento", "Treviso", "Trieste", "Udine", "Varese", "Venezia", "Verbano-Cusio-Ossola", "Vercelli", "Verona", "Vibo Valentia", "Vicenza", "Viterbo"]

# --- CONFIGURAZIONE ---
st.set_page_config(page_title="Nota Spese Pro", page_icon="⚖️")
st.title("⚖️ Nota Spese Tributaria Professionale")

# --- SIDEBAR ---
st.sidebar.header("Procedimento")
grado_selezione = st.sidebar.selectbox("Grado", ["I GRADO", "II GRADO"])
if grado_selezione == "I GRADO":
    sede = st.sidebar.selectbox("Città", sorted(CITTA_I_GRADO))
    grado_h, grado_c = f"DI PRIMO GRADO DI {sede.upper()}", f"di primo grado di {sede.lower()}"
else:
    sede = st.sidebar.selectbox("Regione", sorted(REGIONI_II_GRADO))
    grado_h, grado_c = f"DI SECONDO GRADO {sede.upper()}", f"di secondo grado {sede.lower()}"

rgr = st.sidebar.text_input("R.G.R. n.", "123/2024")
sez = st.sidebar.text_input("Sezione", "")
usa_udienza = st.sidebar.checkbox("Data udienza?")
dt_udienza = st.sidebar.date_input("Udienza", datetime.date.today()).strftime("%d.%m.%Y") if usa_udienza else ""

indet = st.sidebar.checkbox("Valore Indeterminato")
if indet:
    compl = st.sidebar.selectbox("Complessità", ["Bassa", "Media", "Alta"])
    val_calc = 25000 if compl == "Bassa" else 250000 if compl == "Media" else 500000
    txt_val = f"Indeterminato ({compl.lower()})"
else:
    val_lite = st.sidebar.number_input("Valore (€)", min_value=0.0, value=15000.0)
    val_calc = val_lite
    txt_val = ""

# --- CALCOLI ---
def get_params(v):
    if v <= 1100: return (270, 270, 140, 340, "fino a € 1.100")
    elif v <= 5200: return (485, 405, 270, 610, "da € 1.101 a € 5.200")
    elif v <= 26000: return (1150, 875, 675, 1350, "da € 5.201 a € 26.000")
    elif v <= 52000: return (1960, 1285, 1080, 2365, "da € 26.001 a € 52.000")
    elif v <= 260000: return (3375, 2230, 1620, 3850, "da € 52.001 a € 260.000")
    else: return (5065, 3310, 2430, 5805, "da € 260.001 a € 520.000")

p = get_params(val_calc)
if not indet: txt_val = p[4]
comp = p[0] + p[1] + p[3]
s_gen, cpa = comp * 0.15, (comp * 1.15) * 0.04
imp = comp + s_gen + cpa
iva = imp * 0.22
tot = imp + iva

# --- DATI CLIENTE ---
st.sidebar.header("Cliente")
cli = st.sidebar.text_input("Nome", "Federica Benzi")
ln = st.sidebar.text_input("Nascita (Luogo)", "Montevideo (EE)")
dn = st.sidebar.text_input("Nascita (Data)", "21/03/1959")
cf = st.sidebar.text_input("C.F.", "BNZFRC59C61Z613C")
res = st.sidebar.text_input("Residenza", "Sagliano Micca (BI), via Grosso n. 8")

# --- FUNZIONE GRAFICA TABELLA ---
def set_cell_shading(cell, color):
    tc_pr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color)
    tc_pr.append(shd)

# --- WORD GENERATOR ---
def create_doc():
    doc = Document()
    style = doc.styles['Normal']
    style.font.name, style.font.size = 'Calibri', Pt(12)
    style.paragraph_format.line_spacing, style.paragraph_format.space_after = 1.5, Pt(6)

    # Intestazione
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = h.add_run(f"ALLA CORTE DI GIUSTIZIA TRIBUTARIA {grado_h}")
    r.bold = True
    if sez: h.add_run(f"\nSEZIONE {sez.upper()}").bold = True
    if usa_udienza: h.add_run(f"\nUDIENZA DEL {dt_udienza}").bold = True
    
    doc.add_paragraph("* * *").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("Nota spese ex art. 15, D.Lgs. 546/1992").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("* * *").alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Corpo
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run("Il ").bold = False
    p.add_run("prof. dott. Mario Rovetti").bold = True
    p.add_run(" ed il ").bold = False
    p.add_run("dott. Sebastiano Barusco").bold = True
    p.add_run(" con studio in Padova, via Cavazzana 5 (")
    u = p.add_run("Barusco Rovetti & Associati, tel. 049-8752918, PEC: sebastiano.barusco@legalmail.it e studiorovetti@pec.studiorovetti.it")
    u.underline = True
    p.add_run(f"), difensori della signora ").bold = False
    p.add_run(f"{cli}").bold = True
    p.add_run(f", nata a {ln} il {dn}, C.F. {cf}, residente in {res}")

    doc.add_paragraph("\nDEPOSITANO").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Competenza: Corte di giustizia tributaria {grado_c}")

    # --- TABELLA UNIFICATA PROFESSIONALE ---
    table = doc.add_table(rows=0, cols=2)
    table.style = 'Table Grid'
    table.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def add_row(label, value, is_bold=False, bg_color=None):
        row = table.add_row().cells
        row[0].text = label
        row[1].text = value
        row[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        if is_bold:
            row[0].paragraphs[0].runs[0].bold = True
            row[1].paragraphs[0].runs[0].bold = True
        if bg_color:
            set_cell_shading(row[0], bg_color)
            set_cell_shading(row[1], bg_color)

    # Righe Tabella
    add_row("VALORE DELLA CAUSA", txt_val, True, "E7E7E7")
    add_row("Fase di studio della controversia", f"€ {p[0]:,.2f}")
    add_row("Fase introduttiva del giudizio", f"€ {p[1]:,.2f}")
    add_row("Fase decisionale", f"€ {p[3]:,.2f}")
    add_row("COMPENSO TABELLARE", f"€ {comp:,.2f}", True, "F2F2F2")
    add_row("Spese generali (15% su compenso)", f"€ {s_gen:,.2f}")
    add_row("C.P.A. (4% su imponibile)", f"€ {cpa:,.2f}")
    add_row("TOTALE IMPONIBILE", f"€ {imp:,.2f}", True)
    add_row("I.V.A. (22%)", f"€ {iva:,.2f}")
    add_row("IPOTESI DI COMPENSO LIQUIDABILE", f"€ {tot:,.2f}", True, "D9D9D9")

    # Firma
    mesi = {1:"Gennaio", 2:"Febbraio", 3:"Marzo", 4:"Aprile", 5:"Maggio", 6:"Giugno", 7:"Luglio", 8:"Agosto", 9:"Settembre", 10:"Ottobre", 11:"Novembre", 12:"Dicembre"}
    oggi = datetime.date.today()
    doc.add_paragraph(f"\nPadova, {oggi.day} {mesi[oggi.month]} {oggi.year}")
    f = doc.add_paragraph("\nSebastiano Barusco")
    f.add_run("\nFirmato digitalmente")

    out = BytesIO()
    doc.save(out)
    return out.getvalue()

if st.button("Genera Nota Spese"):
    st.download_button("📥 Scarica Word", create_doc(), f"Nota_{cli.replace(' ','_')}.docx")
