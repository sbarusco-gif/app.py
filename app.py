import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import datetime
from io import BytesIO

# Configurazione Pagina
st.set_page_config(page_title="Generatore Nota Spese Tributaria", page_icon="⚖️")

st.title("⚖️ Nota Spese Processo Tributario")
st.subheader("Modello Barusco-Rovetti - Parametri DM 147/2022")

# --- SIDEBAR: DATI PROCEDIMENTO ---
st.sidebar.header("Dati Procedimento")
sede = st.sidebar.text_input("Sede Corte", "PADOVA").upper()
rgr = st.sidebar.text_input("Numero R.G.R.", "123/2024")
cliente = st.sidebar.text_input("Nome Cliente", "Mario Rossi").upper()
grado = st.sidebar.selectbox("Grado", ["I Grado", "II Grado"])
valore = st.sidebar.number_input("Valore della lite (€)", min_value=0.0, value=10000.0, step=100.0)

# --- FASI DA INCLUDERE ---
st.header("Fasi e Compensi")
col1, col2 = st.columns(2)
with col1:
    f_studio = st.checkbox("Fase di Studio", value=True)
    f_intro = st.checkbox("Fase Introduttiva", value=True)
with col2:
    f_istr = st.checkbox("Fase Istruttoria/Trattazione", value=False)
    f_dec = st.checkbox("Fase Decisionale", value=True)

cut = st.number_input("Contributo Unificato / Spese Vive (€)", min_value=0.0, value=0.0)

# --- LOGICA CALCOLO ---
def calcola_parametri(v):
    if v <= 1100: return (270, 270, 140, 340)
    elif v <= 5200: return (485, 405, 270, 610)
    elif v <= 26000: return (1150, 875, 675, 1350)
    elif v <= 52000: return (1960, 1285, 1080, 2365)
    elif v <= 260000: return (3375, 2230, 1620, 3850)
    elif v <= 520000: return (5065, 3310, 2430, 5805)
    else: return (7225, 4795, 3445, 8710)

p = calcola_parametri(valore)
tot_comp = 0
if f_studio: tot_comp += p[0]
if f_intro: tot_comp += p[1]
if f_istr: tot_comp += p[2]
if f_dec: tot_comp += p[3]

# Calcoli Fiscali
spese_gen = tot_comp * 0.15
cpa = (tot_comp + spese_gen) * 0.04
imponibile = tot_comp + spese_gen + cpa
iva = imponibile * 0.22
totale_lordo = imponibile + iva + cut

# --- ANTEPRIMA A SCHERMO ---
st.info(f"**Totale Compenso Tabellare:** € {tot_comp:,.2f}")
st.write(f"Imponibile (con Spese Gen. e CPA): € {imponibile:,.2f}")
st.success(f"### TOTALE LIQUIDABILE: € {totale_lordo:,.2f}")

# --- GENERAZIONE WORD ---
def create_word():
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Verdana'
    style.font.size = Pt(10)

    # Intestazione
    p_head = doc.add_paragraph()
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_head.add_run(f"ALLA CORTE DI GIUSTIZIA TRIBUTARIA DI {grado.upper()} DI {sede}\n* * *\n")
    run.bold = True
    
    corpo = doc.add_paragraph()
    corpo.add_run(f"Il prof. dott. Mario Rovetti ed il dott. Sebastiano Barusco, in qualità di difensori del/della ")
    corpo.add_run(f"{cliente}").bold = True
    corpo.add_run(f", depositano la presente nota spese per il procedimento ")
    corpo.add_run(f"R.G.R. n. {rgr}").bold = True
    corpo.add_run(".")

    doc.add_paragraph(f"\nValore della causa: € {valore:,.2f}\n")

    # Tabella
    table = doc.add_table(rows=0, cols=2)
    data = [
        ("Fase di studio", f"€ {p[0]:,.2f}") if f_studio else None,
        ("Fase introduttiva", f"€ {p[1]:,.2f}") if f_intro else None,
        ("Fase istruttoria", f"€ {p[2]:,.2f}") if f_istr else None,
        ("Fase decisionale", f"€ {p[3]:,.2f}") if f_dec else None,
        ("TOTALE COMPENSI", f"€ {tot_comp:,.2f}"),
        ("Spese generali (15%)", f"€ {spese_gen:,.2f}"),
        ("C.P.A. (4%)", f"€ {cpa:,.2f}"),
        ("IVA (22%)", f"€ {iva:,.2f}"),
        ("Spese vive (C.U.T.)", f"€ {cut:,.2f}"),
    ]
    for item in data:
        if item:
            row = table.add_row().cells
            row[0].text, row[1].text = item

    doc.add_paragraph(f"\nTOTALE DA LIQUIDARE: € {totale_lordo:,.2f}").bold = True
    doc.add_paragraph(f"\n{sede}, {datetime.date.today().strftime('%d/%m/%Y')}\n\nDott. Sebastiano Barusco")
    
    target = BytesIO()
    doc.save(target)
    return target.getvalue()

# Tasto Download
word_data = create_word()
st.download_button(
    label="📄 Scarica Nota Spese in Word",
    data=word_data,
    file_name=f"Nota_Spese_{rgr.replace('/','_')}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)
