import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import os
import subprocess
import base64

# Biblioteki Google
from google.oauth2.service_account import Credentials as SACredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from pptx import Presentation
from pypdf import PdfWriter

# ReportLab - Generowanie eleganckiego PDF
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import reportlab.rl_config

# Wymuszenie poprawnego kodowania w ReportLab
reportlab.rl_config.warnOnMissingFontGlyphs = 0

# --- KONFIGURACJA ---
st.set_page_config(page_title="Generator Ofert - Dwór Dębogóra", layout="wide")

CI = {
    "dark_green": "#00622f",
    "light_green": "#e8ece6",
    "gray": "#333333",
    "white": "#ffffff"
}

ROOT_FOLDER_ID = "1tU6mo1YWpTep8vl5CRR5DhsZAINeWnHz"

# --- INTELIGENTNE ŁADOWANIE CZCIONEK ---
# Domyślnie używamy standardowych czcionek, jeśli TTF zawiodą
FONT_HEADER = 'Helvetica-Bold'
FONT_TEXT = 'Helvetica'
FONT_TEXT_BOLD = 'Helvetica-Bold'
fonts_loaded = False
font_error = ""

lora_path = 'Lora-Bold.ttf'
ptsans_path = 'PTSans-Regular.ttf'
ptsans_bold_path = 'PTSans-Bold.ttf'

if os.path.exists(lora_path) and os.path.exists(ptsans_path):
    try:
        # Rejestrujemy czcionki TTF (ReportLab automatycznie użyje kodowania UTF-8 dla TTF)
        pdfmetrics.registerFont(TTFont('Lora-Bold', lora_path))
        pdfmetrics.registerFont(TTFont('PTSans-Regular', ptsans_path))
        
        FONT_HEADER = 'Lora-Bold'
        FONT_TEXT = 'PTSans-Regular'
        
        if os.path.exists(ptsans_bold_path):
            pdfmetrics.registerFont(TTFont('PTSans-Bold', ptsans_bold_path))
            FONT_TEXT_BOLD = 'PTSans-Bold'
        else:
            FONT_TEXT_BOLD = 'PTSans-Regular'
            
        fonts_loaded = True
    except Exception as e:
        font_error = str(e)
else:
    font_error = "Brak plików Lora-Bold.ttf lub PTSans-Regular.ttf na serwerze."

# --- STYLE CSS INTERFEJSU ---
st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Lora:wght@400;700&family=PT+Sans:wght@400;700&display=swap');
    .stApp {{ background-color: {CI['white']}; font-family: 'PT Sans', sans-serif; }}
    h1, h2, h3, h4 {{ font-family: 'Lora', serif !important; color: {CI['dark_green']} !important; font-weight: 700 !important; }}
    div[data-testid="stVerticalBlock"] > div > div[data-testid="stVerticalBlock"] {{
        background-color: {CI['light_green']}; padding: 2rem; border-left: 5px solid {CI['dark_green']}; margin-bottom: 1.5rem;
    }}
    div.stButton > button {{
        background-color: {CI['dark_green']} !important; color: white !important;
        border-radius: 0px !important; font-family: 'Lora', serif !important; padding: 0.8rem 3rem !important;
        text-transform: uppercase; letter-spacing: 2px;
    }}
    div.stButton > button:hover {{ background-color: {CI['gray']} !important; }}
    </style>
""", unsafe_allow_html=True)

# --- GOOGLE DRIVE LOGIC ---
def get_drive_service():
    info = st.secrets["gcp_service_account"]
    creds = SACredentials.from_service_account_info(info)
    return build('drive', 'v3', credentials=creds)

def fetch_all_debogora_files(root_id):
    service = get_drive_service()
    all_files = []
    folders_to_search = [root_id]
    
    while folders_to_search:
        current_folder = folders_to_search.pop(0)
        query = f"'{current_folder}' in parents and trashed = false"
        try:
            request = service.files().list(q=query, fields="nextPageToken, files(id, name, mimeType)", pageSize=1000)
            while request is not None:
                results = request.execute()
                files = results.get('files', [])
                for f in files:
                    if f['mimeType'] == 'application/vnd.google-apps.folder':
                        folders_to_search.append(f['id'])
                    else:
                        all_files.append(f)
                request = service.files().list_next(request, results)
        except Exception:
            pass
    return all_files

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

def replace_text_in_pptx(prs, search_str, repl_str):
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if search_str in run.text:
                            run.text = run.text.replace(search_str, repl_str)

# --- PANEL BOCZNY ---
with st.sidebar:
    st.subheader("🛠 Diagnostyka systemu")
    st.markdown("**Status czcionek:**")
    if fonts_loaded:
        st.success("✅ Czcionki TTF załadowane.")
    else:
        st.error(f"❌ Aktywna czcionka zastępcza. Powód: {font_error}")
        
    try:
        wszystkie_pliki = fetch_all_debogora_files(ROOT_FOLDER_ID)
        st.success(f"✅ Połączono z Drive. Liczba plików: {len(wszystkie_pliki)}")
    except Exception as e:
        st.error(f"❌ Błąd połączenia Drive: {e}")
        wszystkie_pliki = []

# --- LOGIKA CENNIKA ---
CENNIK = {
    "nocleg_1_noc": 220, "nocleg_2_noce": 170, "doplata_domek": 40,
    "domki": {
        "Muuu 1": {"baza": 700, "max_os": 4, "pdf": "Krovacja"}, 
        "Muuu 2": {"baza": 700, "max_os": 4, "pdf": "Krovacja"},
        "Muuu 3": {"baza": 1050, "max_os": 6, "pdf": "Krovacja"}, 
        "Muuu 4": {"baza": 1050, "max_os": 6, "pdf": "Krovacja"},
        "Muuu 5": {"baza": 700, "max_os": 3, "pdf": "Krovacja"}, 
        "Muuu 6": {"baza": 700, "max_os": 3, "pdf": "Krovacja"}
    },
    "wyzywienie": {"Brak": 0, "Śniadanie": 50, "Śniadanie + Obiadokolacja": 120},
    "atrakcje": {
        "Kajaki": {"cena": 140, "typ": "osoba", "pdf": "Kajak"},
        "Sauna Olchowa": {"cena": 400, "typ": "grupa", "pdf": "Sauna"},
        "Balia": {"cena": 300, "typ": "grupa", "pdf": "Balia"},
        "Paintball": {"cena": 150, "typ": "osoba", "pdf": "Paintball"},
        "Skarby Dębogóry": {"cena": 200, "typ": "osoba", "pdf": "Skarby"},
        "Ognisko": {"cena": 150, "typ": "grupa", "pdf": "Ognisko"}
    }
}
POKOJE_DWOREK = {f"Pokój nr {i}": (1 if i==1 else 4 if i==11 else 3 if i in [7,9,10,12] else 2) for i in range(1,13)}

# --- INTERFEJS APLIKACJI ---
try:
    logo_b64 = base64.b64encode(open("logo.png", "rb").read()).decode()
    st.markdown(f'<div style="display: flex; justify-content: center; margin-bottom: 10px;"><img src="data:image/png;base64,{logo_b64}" width="120"></div>', unsafe_allow_html=True)
except: pass

st.markdown("<h1 style='text-align: center; margin-top:0;'>System Ofertowania</h1>", unsafe_allow_html=True)

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
    st.subheader("1. Dane Klienta i Termin")
    c1, c2 = st.columns(2)
    with c1:
        klient_imie = st.text_input("Imię i nazwisko *")
        firma_n = st.text_input("Firma (opcjonalnie)")
        nip_n = st.text_input("NIP (opcjonalnie)")
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
    with col_dw:
        p_sel = st.multiselect("Dworek", list(POKOJE_DWOREK.keys()), key="wybrane_p")
        for p in p_sel:
            ile = st.number_input(f"{p}", 1, POKOJE_DWOREK[p], key=f"os_{p}")
            pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"{p} (os: {ile})", "Ilość": ile, "Cena": stawka_dw*dni, "Suma": ile*stawka_dw*dni, "pdf_kw": "Dworek"})
    with col_dm:
        d_sel = st.multiselect("Domki", list(CENNIK["domki"].keys()), key="wybrane_d")
        for d in d_sel:
            ile = st.number_input(f"{d}", 1, CENNIK["domki"][d]["max_os"], key=f"os_{d}")
            cena_d = (CENNIK["domki"][d]["baza"] + (max(0, ile-1)*CENNIK["doplata_domek"]))*dni
            pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"{d} ({ile} os.)", "Ilość": 1, "Cena": cena_d, "Suma": cena_d, "pdf_kw": CENNIK["domki"][d]["pdf"]})

with st.container():
    st.subheader("3. Wyżywienie i Atrakcje")
    wyz_opt = st.selectbox("Wyżywienie", list(CENNIK["wyzywienie"].keys()))
    if wyz_opt != "Brak":
        pozycje_kosztowe.append({"Kategoria": "Gastronomia", "Opis": wyz_opt, "Ilość": st.session_state.l_osob_total*dni, "Cena": CENNIK["wyzywienie"][wyz_opt], "Suma": CENNIK["wyzywienie"][wyz_opt]*st.session_state.l_osob_total*dni, "pdf_kw": None})
    
    atr_sel = st.multiselect("Dodaj atrakcje", list(CENNIK["atrakcje"].keys()))
    for a in atr_sel:
        a_data = CENNIK["atrakcje"][a]
        ile = st.number_input(f"Ilość: {a}", 1, 100, st.session_state.l_osob_total if a_data["typ"]=="osoba" else 1)
        pozycje_kosztowe.append({"Kategoria": "Atrakcje", "Opis": a, "Ilość": ile, "Cena": a_data["cena"], "Suma": ile*a_data["cena"], "pdf_kw": a_data["pdf"]})

# --- FUNKCJA POMOCNICZA DLA ZNAKÓW ---
# Dla pewności usuwamy białe znaki i wymuszamy poprawne typy dla ReportLab
def safe_str(text):
    if pd.isna(text):
        return ""
    return str(text).strip()

# --- GENEROWANIE finalnego pliku ---
with st.container():
    st.subheader("4. Kosztorys i Eksport")
    df = pd.DataFrame(pozycje_kosztowe)
    if not df.empty:
        edf = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        razem = edf["Suma"].sum()
        st.markdown(f"<h3 style='color: {CI['dark_green']};'>RAZEM DO ZAPŁATY: {razem:,.2f} PLN</h3>".replace(",", " "), unsafe_allow_html=True)
        
        if st.button("GENERUJ FINALNY PDF"):
            if not klient_imie:
                st.error("Podaj imię i nazwisko klienta!")
            else:
                with st.spinner("Pobieranie plików z Dysku i budowanie oferty..."):
                    merger = PdfWriter()
                    
                    # 1. OKŁADKA
                    okladka_file = next((f for f in wszystkie_pliki if 'okładka' in f['name'].lower()), None)
                    if okladka_file:
                        try:
                            ppt_stream = download_file(okladka_file['id'])
                            prs = Presentation(ppt_stream)
                            nazwa_docelowa = firma_n if firma_n else klient_imie
                            replace_text_in_pptx(prs, "{{nazwa firmy}}", nazwa_docelowa)
                            prs.save("okladka_temp.pptx")
                            subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "okladka_temp.pptx"])
                            merger.append("okladka_temp.pdf")
                        except Exception as e:
                            st.error(f"Błąd przetwarzania okładki: {e}")

                    # 2. AUTOMATYCZNE KARTY PDF Z DYSKU GOOGLE
                    keywords = list(set([row["pdf_kw"] for _, row in edf.iterrows() if pd.notna(row["pdf_kw"]) and row["pdf_kw"]]))
                    for kw in keywords:
                        matched_files = [f for f in wszystkie_pliki if kw.lower() in f['name'].lower() and 'pdf' in f['mimeType'].lower()]
                        if matched_files:
                            prev_files = [f for f in matched_files if 'prev' in f['name'].lower()]
                            selected_pdf = prev_files[0] if prev_files else matched_files[0]
                            try:
                                merger.append(download_file(selected_pdf['id']))
                            except Exception:
                                pass

                    # 3. ZAPROJEKTOWANA STRONA OFERTOWA (REPORTLAB)
                    buf = io.BytesIO()
                    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=50, bottomMargin=50)
                    elements = []
                    styles = getSampleStyleSheet()

                    if os.path.exists("logo.png"):
                        elements.append(RLImage("logo.png", width=120, height=60, kind='proportional'))
                        elements.append(Spacer(1, 20))

                    header_style = ParagraphStyle('HeaderCI', parent=styles['Heading1'], fontName=FONT_HEADER, fontSize=20, textColor=colors.HexColor(CI['dark_green']), spaceAfter=20)
                    elements.append(Paragraph("PODSUMOWANIE KOSZTÓW", header_style))
                    
                    sub_style = ParagraphStyle('SubCI', parent=styles['Normal'], fontName=FONT_TEXT, fontSize=12, textColor=colors.HexColor(CI['gray']), spaceAfter=30)
                    nazwa = safe_str(firma_n) if firma_n else safe_str(klient_imie)
                    elements.append(Paragraph(f"Oferta przygotowana dla: <b>{nazwa}</b><br/>Data wygenerowania: {date.today().strftime('%d.%m.%Y')}", sub_style))
                    
                    t_data = [["Kategoria", "Opis usługi", "Ilość", "Suma"]]
                    for _, row in edf.iterrows():
                        kat = safe_str(row["Kategoria"])
                        opis = safe_str(row["Opis"])
                        ilosc = safe_str(row["Ilość"])
                        suma_str = f"{row['Suma']:,.0f} zł".replace(",", " ")
                        t_data.append([kat, opis, ilosc, suma_str])
                    
                    t_data.append(["", "", "RAZEM:", f"{razem:,.0f} zł".replace(",", " ")])
                    
                    table = Table(t_data, colWidths=[100, 230, 50, 90])
                    t_style = TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(CI['dark_green'])),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                        ('FONTNAME', (0, 0), (-1, 0), FONT_HEADER),
                        ('FONTSIZE', (0, 0), (-1, 0), 12),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('ALIGN', (1, 1), (1, -2), 'LEFT'),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
                        ('TOPPADDING', (0, 0), (-1, -1), 10),
                        ('FONTNAME', (0, 1), (-1, -1), FONT_TEXT),
                        ('FONTNAME', (2, -1), (-1, -1), FONT_TEXT_BOLD),
                        ('TEXTCOLOR', (2, -1), (-1, -1), colors.HexColor(CI['dark_green'])),
                        ('BACKGROUND', (2, -1), (-1, -1), colors.HexColor(CI['light_green'])),
                        ('LINEABOVE', (0, -1), (-1, -1), 1, colors.HexColor(CI['dark_green'])),
                    ])
                    
                    for i in range(1, len(t_data) - 1):
                        bg_color = colors.HexColor(CI['light_green']) if i % 2 == 0 else colors.white
                        t_style.add('BACKGROUND', (0, i), (-1, i), bg_color)
                        t_style.add('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#d1d9cf"))

                    table.setStyle(t_style)
                    elements.append(table)
                    
                    try:
                        doc.build(elements)
                        buf.seek(0)
                        merger.append(buf)
                    except Exception as e:
                        st.error(f"Wystąpił problem ze znakami podczas budowania PDF: {e}")

                    final_pdf = io.BytesIO()
                    merger.write(final_pdf)
                    
                    safe_filename = "".join([c for c in safe_str(klient_imie) if c.isalpha() or c.isdigit() or c==' ']).rstrip().replace(" ", "_")
                    st.download_button("📥 POBIERZ SCALONĄ OFERTĘ PDF", final_pdf.getvalue(), f"Oferta_{safe_filename}.pdf", "application/pdf")
