import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import os
import subprocess
import base64

# Biblioteki Google i PDF
from google.oauth2.service_account import Credentials as SACredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from pptx import Presentation
from pypdf import PdfWriter, PdfReader

# ReportLab - Generowanie PDF z zachowaniem czcionek CI
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# --- KONFIGURACJA CI DĘBOGÓRA ---
CI = {
    "dark_green": "#00622f",
    "light_green": "#e8ece6",
    "gray": "#333333",
    "white": "#ffffff"
}

# Folder główny z podfolderami (nowy link)
ROOT_FOLDER_ID = "1tU6mo1YWpTep8vl5CRR5DhsZAINeWnHz"

# Rejestracja czcionek (Pliki muszą być w repozytorium GitHub)
try:
    pdfmetrics.registerFont(TTFont('Lora-Bold', 'Lora-Bold.ttf'))
    pdfmetrics.registerFont(TTFont('Lora-Regular', 'Lora-Regular.ttf'))
    pdfmetrics.registerFont(TTFont('PTSans-Regular', 'PTSans-Regular.ttf'))
except:
    pass # Jeśli nie ma plików, użyje standardowych

st.set_page_config(page_title="System Ofertowania - Dębogóra", layout="wide")

# --- STYLE CSS ---
st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Lora:wght@400;700&family=PT+Sans:wght@400;700&display=swap');
    .stApp {{ background-color: {CI['white']}; font-family: 'PT Sans', sans-serif; }}
    h1, h2, h3 {{ font-family: 'Lora', serif !important; color: {CI['dark_green']} !important; font-weight: 700 !important; }}
    div[data-testid="stVerticalBlock"] > div > div[data-testid="stVerticalBlock"] {{
        background-color: {CI['light_green']}; padding: 2.5rem; border-left: 5px solid {CI['dark_green']};
    }}
    div.stButton > button {{
        background-color: {CI['dark_green']} !important; color: white !important;
        border-radius: 0px !important; font-family: 'Lora', serif !important; padding: 0.8rem 3rem !important;
    }}
    </style>
""", unsafe_allow_html=True)

# --- GOOGLE DRIVE LOGIC ---
@st.cache_resource
def get_drive_service():
    info = st.secrets["gcp_service_account"]
    creds = SACredentials.from_service_account_info(info)
    return build('drive', 'v3', credentials=creds)

def find_file_by_name(name_substring, folder_id=ROOT_FOLDER_ID):
    service = get_drive_service()
    # Szukaj rekurencyjnie w podfolderach
    query = f"name contains '{name_substring}' and mimeType = 'application/pdf' and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    return results.get('files', [])

def download_file(file_id):
    service = get_drive_service()
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    fh.seek(0)
    return fh

# --- LOGIKA BIZNESOWA ---
CENNIK = {
    "nocleg_1_noc": 220, "nocleg_2_noce": 170, "doplata_domek": 40,
    "domki": {
        "Muuu 1": {"baza": 700, "max_os": 4, "pdf": "Muuu1"}, 
        "Muuu 2": {"baza": 700, "max_os": 4, "pdf": "Muuu2"},
        "Muuu 3": {"baza": 1050, "max_os": 6, "pdf": "Muuu3"}, 
        "Muuu 4": {"baza": 1050, "max_os": 6, "pdf": "Muuu4"},
        "Muuu 5": {"baza": 700, "max_os": 3, "pdf": "Muuu5"}, 
        "Muuu 6": {"baza": 700, "max_os": 3, "pdf": "Muuu6"}
    },
    "wyzywienie": {"Brak": 0, "Śniadanie": 50, "Śniadanie + Obiadokolacja": 120},
    "atrakcje": {
        "Kajaki": {"cena": 140, "typ": "osoba", "pdf": "kajak"},
        "Sauna Olchowa": {"cena": 400, "typ": "grupa", "pdf": "sauna"},
        "Balia": {"cena": 300, "typ": "grupa", "pdf": "balia"},
        "Paintball": {"cena": 150, "typ": "osoba", "pdf": "paintball"},
        "Skarby Dębogóry": {"cena": 200, "typ": "osoba", "pdf": "skarby"},
        "Ognisko": {"cena": 150, "typ": "grupa", "pdf": "ognisko"}
    }
}
POKOJE_DWOREK = {f"Pokój nr {i}": (1 if i==1 else 4 if i==11 else 3 if i in [7,9,10,12] else 2) for i in range(1,13)}

# --- INTERFEJS ---
try:
    st.image("logo.png", width=120)
except: pass

st.markdown("<h1 style='text-align: center;'>System Ofertowania Dębogóra</h1>", unsafe_allow_html=True)

if "l_osob_total" not in st.session_state: st.session_state.l_osob_total = 10

def auto_alloc():
    total = st.session_state.l_osob_total
    st.session_state.wybrane_p = []
    st.session_state.wybrane_d = []
    for p, cap in POKOJE_DWOREK.items():
        if total > 0:
            st.session_state.wybrane_p.append(p)
            val = min(cap, total); st.session_state[f"os_{p}"] = val; total -= val
    for d, par in CENNIK["domki"].items():
        if total > 0:
            st.session_state.wybrane_d.append(d)
            val = min(par["max_os"], total); st.session_state[f"os_{d}"] = val; total -= val

pozycje_kosztowe = []

with st.container():
    st.subheader("1. Dane Klienta")
    c1, c2 = st.columns(2)
    with c1:
        klient_imie = st.text_input("Imię i nazwisko *")
        firma_n = st.text_input("Firma")
        nip_n = st.text_input("NIP")
        st.number_input("Liczba osób", 1, 100, key="l_osob_total")
        st.button("🤖 Automatycznie rozmieść gości", on_click=auto_alloc)
    with c2:
        email_n = st.text_input("Email")
        tel_n = st.text_input("Telefon")
        cd1, cd2 = st.columns(2)
        with cd1: d_in = st.date_input("Przyjazd", date.today())
        with cd2: d_out = st.date_input("Wyjazd", date.today()+timedelta(1))
        dni = max(1, (d_out - d_in).days)

with st.container():
    st.subheader("2. Zakwaterowanie")
    stawka_dw = CENNIK["nocleg_1_noc"] if dni == 1 else CENNIK["nocleg_2_noce"]
    col_dw, col_dm = st.columns(2)
    sum_os = 0
    with col_dw:
        p_sel = st.multiselect("Dworek", list(POKOJE_DWOREK.keys()), key="wybrane_p")
        for p in p_sel:
            ile = st.number_input(f"{p}", 1, POKOJE_DWOREK[p], key=f"os_{p}")
            sum_os += ile
            pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"{p} (os: {ile})", "Ilość": ile, "Cena": stawka_dw*dni, "Suma": ile*stawka_dw*dni, "pdf_keyword": "Dworek"})
    with col_dm:
        d_sel = st.multiselect("Domki", list(CENNIK["domki"].keys()), key="wybrane_d")
        for d in d_sel:
            ile = st.number_input(f"{d}", 1, CENNIK["domki"][d]["max_os"], key=f"os_{d}")
            sum_os += ile
            cena_d = (CENNIK["domki"][d]["baza"] + (max(0, ile-1)*CENNIK["doplata_domek"]))*dni
            pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"{d} ({ile} os.)", "Ilość": 1, "Cena": cena_d, "Suma": cena_d, "pdf_keyword": CENNIK["domki"][d]["pdf"]})

with st.container():
    st.subheader("3. Wyżywienie i Atrakcje")
    c_w1, c_w2 = st.columns(2)
    with c_w1:
        wyz_opt = st.selectbox("Wyżywienie", list(CENNIK["wyzywienie"].keys()))
        if wyz_opt != "Brak":
            pozycje_kosztowe.append({"Kategoria": "Gastronomia", "Opis": wyz_opt, "Ilość": st.session_state.l_osob_total*dni, "Cena": CENNIK["wyzywienie"][wyz_opt], "Suma": CENNIK["wyzywienie"][wyz_opt]*st.session_state.l_osob_total*dni, "pdf_keyword": "wyżywienie"})
    with c_w2:
        atr_sel = st.multiselect("Dodaj atrakcje", list(CENNIK["atrakcje"].keys()))
        for a in atr_sel:
            a_data = CENNIK["atrakcje"][a]
            ile = st.number_input(f"Ilość: {a}", 1, 100, st.session_state.l_osob_total if a_data["typ"]=="osoba" else 1)
            pozycje_kosztowe.append({"Kategoria": "Atrakcje", "Opis": a, "Ilość": ile, "Cena": a_data["cena"], "Suma": ile*a_data["cena"], "pdf_keyword": a_data["pdf"]})

# --- GENEROWANIE ---
with st.container():
    st.subheader("4. Kosztorys i Eksport")
    df = pd.DataFrame(pozycje_kosztowe)
    if not df.empty:
        edf = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        if st.button("GENERUJ FINALNY PDF"):
            with st.spinner("Budowanie oferty..."):
                merger = PdfWriter()
                
                # 1. OKŁADKA (PPTX -> PDF)
                try:
                    okl_files = find_file_by_name("okładka_02")
                    ppt_stream = download_file(okl_files[0]['id'])
                    prs = Presentation(ppt_stream)
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if hasattr(shape, "text") and "{{nazwa firmy}}" in shape.text:
                                shape.text = shape.text.replace("{{nazwa firmy}}", firma_n if firma_n else klient_imie)
                    prs.save("okladka.pptx")
                    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "okladka.pptx"])
                    merger.append("okladka.pdf")
                except: st.error("Nie znaleziono pliku okładki na Dysku.")

                # 2. AUTOMATYCZNE KARTY PRODUKTÓW
                # Pobieramy unikalne słowa kluczowe z wybranych pozycji
                keywords = list(set(edf["pdf_keyword"].tolist()))
                for kw in keywords:
                    found = find_file_by_name(kw)
                    if found:
                        merger.append(download_file(found[0]['id']))

                # 3. STRONA OFERTOWA (ReportLab - ZGODNA Z CI)
                buf = io.BytesIO()
                doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
                
                # Definicja Stylów CI
                styles = getSampleStyleSheet()
                header_style = ParagraphStyle('HeaderCI', parent=styles['Heading1'], fontName='Lora-Bold', fontSize=24, textColor=colors.HexColor(CI['dark_green']), spaceAfter=20)
                text_style = ParagraphStyle('TextCI', parent=styles['Normal'], fontName='PTSans-Regular', fontSize=12, textColor=colors.HexColor(CI['gray']))
                
                elements = [Paragraph(f"Oferta: {firma_n if firma_n else klient_imie}", header_style), Spacer(1, 12)]
                
                # Tabela Ofertowa
                t_data = [["Kategoria", "Usługa", "Ilość", "Suma"]]
                for _, row in edf.iterrows():
                    t_data.append([row["Kategoria"], row["Opis"], str(row["Ilość"]), f"{row['Suma']:.0f} zł"])
                
                t_data.append(["", "", "SUMA CAŁKOWITA:", f"{edf['Suma'].sum():.0f} zł"])
                
                table = Table(t_data, colWidths=[100, 250, 60, 100])
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(CI['dark_green'])),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('FONTNAME', (0, 0), (-1, 0), 'Lora-Bold'),
                    ('FONTNAME', (0, 1), (-1, -1), 'PTSans-Regular'),
                    ('GRID', (0, 0), (-1, -2), 0.5, colors.HexColor(CI['dark_green'])),
                    ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor(CI['light_green'])),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTSIZE', (0, 0), (-1, -1), 11),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
                ]))
                elements.append(table)
                
                doc.build(elements)
                buf.seek(0)
                merger.append(buf)

                final_pdf = io.BytesIO()
                merger.write(final_pdf)
                st.download_button("📥 POBIERZ KOMPLETNĄ OFERTĘ PDF", final_pdf.getvalue(), f"Oferta_{klient_imie}.pdf", "application/pdf")
