import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import os
import subprocess
import base64
import time

# Biblioteki Google
from google.oauth2.service_account import Credentials as SACredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from pptx import Presentation
from pypdf import PdfWriter, PdfReader, PageObject

# ReportLab
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import reportlab.rl_config

reportlab.rl_config.warnOnMissingFontGlyphs = 0

st.set_page_config(page_title="Generator Ofert - Dwór Dębogóra", layout="wide")

CI = {
    "dark_green": "#00622f",
    "light_green": "#e8ece6",
    "gray": "#333333",
    "white": "#ffffff"
}

ROOT_FOLDER_ID = "1tU6mo1YWpTep8vl5CRR5DhsZAINeWnHz"

# --- INTELIGENTNE ŁADOWANIE CZCIONEK ---
FONT_HEADER = 'Helvetica-Bold'
FONT_TEXT = 'Helvetica'
FONT_TEXT_BOLD = 'Helvetica-Bold'
fonts_loaded = False

lora_path = 'Lora-Bold.ttf'
text_font_path = None
text_font_bold_path = None

for f in ['Lato-Regular.ttf', 'PTSans-Regular.ttf']:
    if os.path.exists(f): text_font_path = f; break
for f in ['Lato-Bold.ttf', 'PTSans-Bold.ttf']:
    if os.path.exists(f): text_font_bold_path = f; break

if os.path.exists(lora_path) and text_font_path:
    try:
        pdfmetrics.registerFont(TTFont('Lora-Bold', lora_path))
        pdfmetrics.registerFont(TTFont('CI-Text', text_font_path))
        FONT_HEADER = 'Lora-Bold'
        FONT_TEXT = 'CI-Text'
        if text_font_bold_path:
            pdfmetrics.registerFont(TTFont('CI-Text-Bold', text_font_bold_path))
            FONT_TEXT_BOLD = 'CI-Text-Bold'
        else:
            FONT_TEXT_BOLD = 'CI-Text'
        fonts_loaded = True
    except: pass

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
@st.cache_resource
def get_drive_service():
    info = st.secrets["gcp_service_account"]
    creds = SACredentials.from_service_account_info(info)
    return build('drive', 'v3', credentials=creds)

@st.cache_data(ttl=600, show_spinner=False)
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
        except: pass
    return all_files

def download_file(file_id, retries=3):
    service = get_drive_service()
    for attempt in range(retries):
        try:
            request = service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request, chunksize=256*1024)
            done = False
            while not done:
                status, done = downloader.next_chunk()
            fh.seek(0)
            return fh
        except Exception as e:
            if attempt == retries - 1: raise e
            time.sleep(1)

def update_file_on_drive(file_id, df, file_name):
    service = get_drive_service()
    buffer = io.BytesIO()
    if 'xlsx' in file_name.lower():
        df.to_excel(buffer, index=False, engine='openpyxl')
        mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    else:
        df.to_csv(buffer, index=False, encoding='utf-8')
        mimetype = 'text/csv'
    buffer.seek(0)
    media = MediaIoBaseUpload(buffer, mimetype=mimetype, resumable=True)
    service.files().update(fileId=file_id, media_body=media).execute()

# --- PANCERNA ZAMIANA TEKSTU W PPTX ---
def process_shape(shape, replacements):
    if hasattr(shape, "shapes"):
        for subshape in shape.shapes:
            process_shape(subshape, replacements)
    if getattr(shape, "has_table", False):
        for row in shape.table.rows:
            for cell in row.cells:
                process_shape(cell, replacements)
    if getattr(shape, "has_text_frame", False):
        for paragraph in shape.text_frame.paragraphs:
            if not paragraph.runs: continue
            full_text = "".join(run.text for run in paragraph.runs)
            orig = full_text
            clean_text = full_text.replace("{{", "").replace("}}", "")
            for k, v in replacements.items():
                clean_text = clean_text.replace(k, str(v))
            if clean_text != orig:
                paragraph.runs[0].text = clean_text
                for i in range(1, len(paragraph.runs)):
                    paragraph.runs[i].text = ""

def replace_text_in_pptx(prs, replacements):
    for slide in prs.slides:
        for shape in slide.shapes:
            process_shape(shape, replacements)

def normalize_pl(text):
    rep = {'ą':'a', 'ć':'c', 'ę':'e', 'ł':'l', 'ń':'n', 'ó':'o', 'ś':'s', 'ź':'z', 'ż':'z'}
    res = str(text).lower()
    for k, v in rep.items(): res = res.replace(k, v)
    return res

# --- MAPOWANIE NAZW DO KART ---
SYNONYMS = {
    "okładka": "okładka_02",
    "powitalna": "karta powitalna",
    "dworek": "Zakwaterowanie", 
    "krovacja": "Zakwaterowanie",
    "wyżywienie": "wyżywienie",
    "atrakcje_wstęp": "atrakcje_",
    "wycena": "wycena",
    "agenda": "agenda",
    "kontakt": "kontakt"
}

def get_file_by_keyword(keyword, all_files):
    search_term = SYNONYMS.get(keyword, keyword)
    norm_search = normalize_pl(search_term)
    matches = [f for f in all_files if norm_search in normalize_pl(f['name']) and ('.pdf' in f['name'].lower() or 'pdf' in f['mimeType'].lower() or '.ppt' in f['name'].lower() or 'presentation' in f['mimeType'].lower())]
    if matches:
        return next((f for f in matches if 'prev' in f['name'].lower()), matches[0])
    return None

def process_file_to_pdf_stream(file_obj, replacements=None):
    fh = download_file(file_obj['id'])
    fname = file_obj['name'].lower()
    if 'ppt' in fname or 'presentation' in file_obj['mimeType']:
        temp_ppt = f"temp_{file_obj['id']}.pptx"
        temp_pdf = f"temp_{file_obj['id']}.pdf"
        with open(temp_ppt, "wb") as f:
            f.write(fh.getvalue())
        if replacements:
            prs = Presentation(temp_ppt)
            replace_text_in_pptx(prs, replacements)
            prs.save(temp_ppt)
        subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", temp_ppt], check=True)
        with open(temp_pdf, "rb") as f:
            pdf_bytes = f.read()
        return io.BytesIO(pdf_bytes)
    return fh

def add_file_to_merger(merger, keyword, all_files, open_streams, missing_cards, replacements=None):
    file_obj = get_file_by_keyword(keyword, all_files)
    if file_obj:
        try:
            pdf_stream = process_file_to_pdf_stream(file_obj, replacements)
            open_streams.append(pdf_stream)
            merger.append(PdfReader(pdf_stream, strict=False))
        except Exception as e:
            st.error(f"⚠️ Błąd przetwarzania pliku '{keyword}': {e}")
    elif keyword not in missing_cards:
        missing_cards.append(keyword)

def safe_str(text): return "" if pd.isna(text) else str(text).strip()

def get_price_from_df(usluga_name, df, default=0):
    if df is None or df.empty: return default
    try:
        col_name = next((c for c in df.columns if 'nazwa' in c.lower() or 'usługa' in c.lower() or 'usluga' in c.lower()), None)
        col_price = next((c for c in df.columns if 'cena' in c.lower()), None)
        if col_name and col_price:
            match = df[df[col_name].astype(str).str.strip().str.lower() == usluga_name.lower()]
            if not match.empty:
                val = match.iloc[0][col_price]
                if pd.notna(val): return float(val)
    except: pass
    return default

# --- POŁĄCZENIE Z DYSKIEM ---
wszystkie_pliki = []
cennik_file = None

try:
    with st.spinner("Skanowanie plików na Dysku Google..."):
        wszystkie_pliki = fetch_all_debogora_files(ROOT_FOLDER_ID)
    cennik_files = [f for f in wszystkie_pliki if 'cennik' in f['name'].lower() and ('xlsx' in f['name'].lower() or 'csv' in f['name'].lower())]
    cennik_files.sort(key=lambda f: 'xlsx' in f['name'].lower(), reverse=True)
    if cennik_files:
        cennik_file = cennik_files[0]
except Exception as e:
    st.sidebar.error(f"Błąd połączenia z Drive: {e}")

if cennik_file and 'df_cennik' not in st.session_state:
    file_stream = download_file(cennik_file['id'])
    if 'xlsx' in cennik_file['name'].lower():
        try: st.session_state.df_cennik = pd.read_excel(file_stream, engine='openpyxl')
        except: st.session_state.df_cennik = None
    else:
        try: st.session_state.df_cennik = pd.read_csv(file_stream, encoding='utf-8')
        except:
            file_stream.seek(0)
            st.session_state.df_cennik = pd.read_csv(file_stream, encoding='cp1250')
elif 'df_cennik' not in st.session_state:
    st.session_state.df_cennik = None

df_c = st.session_state.df_cennik

# --- DYNAMICZNY SŁOWNIK CENNIKA ---
CENNIK = {
    "nocleg_1_noc": get_price_from_df("Nocleg (1 noc)", df_c, 220),
    "nocleg_2_noce": get_price_from_df("Nocleg (2+ noce)", df_c, 170),
    "doplata_domek": 40,
    "domki": {
        "Muuu 1": {"baza": get_price_from_df("Muuu 1, 2", df_c, 700), "max_os": 4, "pdf": "Zakwaterowanie"}, 
        "Muuu 2": {"baza": get_price_from_df("Muuu 1, 2", df_c, 700), "max_os": 4, "pdf": "Zakwaterowanie"},
        "Muuu 3": {"baza": get_price_from_df("Muuu 3, 4", df_c, 1050), "max_os": 6, "pdf": "Zakwaterowanie"}, 
        "Muuu 4": {"baza": get_price_from_df("Muuu 3, 4", df_c, 1050), "max_os": 6, "pdf": "Zakwaterowanie"},
        "Muuu 5": {"baza": get_price_from_df("Muuu 5, 6", df_c, 700), "max_os": 3, "pdf": "Zakwaterowanie"}, 
        "Muuu 6": {"baza": get_price_from_df("Muuu 5, 6", df_c, 700), "max_os": 3, "pdf": "Zakwaterowanie"}
    },
    "wyzywienie": {
        "Śniadanie": {"cena": get_price_from_df("Śniadanie", df_c, 50), "pdf": "wyżywienie"},
        "Obiadokolacja": {"cena": get_price_from_df("Obiadokolacja", df_c, 70), "pdf": "wyżywienie"},
        "Śniadanie + Obiadokolacja": {"cena": get_price_from_df("Śniadanie", df_c, 50) + get_price_from_df("Obiadokolacja", df_c, 70), "pdf": "wyżywienie"},
        "Serwis kawowy": {"cena": get_price_from_df("Serwis kawowy", df_c, 35), "pdf": "wyżywienie"},
        "Wiejskie jadło": {"cena": get_price_from_df("Wiejskie jadło (Podstawowe)", df_c, 45), "pdf": "wyżywienie"},
        "Kolacja z rozszerzonym menu": {"cena": get_price_from_df("Biesiada wieczorna", df_c, 150), "pdf": "Biesiada"}
    },
    "SPAstwisko": {
        "Seans full experience": {"cena": get_price_from_df("Seans full experience", df_c, 0), "typ": "grupa", "pdf": "Seans saunowy"},
        "Sauna olchowa": {"cena": get_price_from_df("Sauna Olchowa", df_c, 400), "typ": "grupa", "pdf": "Sauna olchowa"},
        "Staw kąpielowy": {"cena": get_price_from_df("Staw kąpielowy", df_c, 0), "typ": "grupa", "pdf": "Staw"},
        "Balia opalana drewnem": {"cena": get_price_from_df("Balia opalana drewnem", df_c, 300), "typ": "grupa", "pdf": "Balia"},
        "Sauny": {"cena": get_price_from_df("Wynajem na wyłączność", df_c, 450), "typ": "grupa", "pdf": "Sauny"},
        "Masaż relaksacyjny": {"cena": get_price_from_df("Masaż relaksacyjny", df_c, 200), "typ": "osoba", "pdf": "Masaże"},
        "Masaż gorącą świecą": {"cena": get_price_from_df("Masaż gorącą świecą", df_c, 210), "typ": "osoba", "pdf": "Masaże"},
        "Masaż gorącymi kamieniami": {"cena": get_price_from_df("Masaż gorącymi kamieniami", df_c, 220), "typ": "osoba", "pdf": "Masaże"},
        "Masaż klasyczny częściowy": {"cena": get_price_from_df("Masaż klasyczny częściowy", df_c, 180), "typ": "osoba", "pdf": "Masaże"},
        "Masaż twarzy i dekoltu": {"cena": get_price_from_df("Masaż twarzy i dekoltu", df_c, 170), "typ": "osoba", "pdf": "Masaże"}
    },
    "Atrakcje": {
        "Łowcy krów": {"cena": get_price_from_df("Łowcy krów", df_c, 100), "typ": "osoba", "pdf": "ZłodziejKrów"},
        "Skarby Dębogóry": {"cena": get_price_from_df("Skarby Dębogóry", df_c, 200), "typ": "osoba", "pdf": "Skarby Dębogóry"},
        "Krowie Safari Standard": {"cena": get_price_from_df("Krowie Safari Standard", df_c, 100), "typ": "osoba", "pdf": "Krowie Safari dla grup_Standard"},
        "Krowie Safari Rozszerzone": {"cena": get_price_from_df("Krowie Safari Rozszerzone", df_c, 150), "typ": "osoba", "pdf": "Krowie Safari dla grup_Rozszerzona"},
        "Paintball": {"cena": get_price_from_df("Paintball", df_c, 150), "typ": "osoba", "pdf": "Paintball"},
        "Kajaki": {"cena": get_price_from_df("Kajaki", df_c, 140), "typ": "osoba", "pdf": "Spływ kajakowy"},
        "Rowery elektryczne krótka przejażdżka (2-3h)": {"cena": get_price_from_df("Rowery elektryczne krótka przejażdżka (2-3 godizny)", df_c, 0), "typ": "osoba", "pdf": "Rowery"},
        "Rowery elektryczne 1 dzień": {"cena": get_price_from_df("Rowery elektryczne 1 dzień", df_c, 0), "typ": "osoba", "pdf": "Rowery"},
        "Rowery elektryczne 2 dni": {"cena": get_price_from_df("Rowery elektryczne 2 dni", df_c, 0), "typ": "osoba", "pdf": "Rowery"},
        "Rowery elektryczne 3 dni": {"cena": get_price_from_df("Rowery elektryczne 3 dni", df_c, 0), "typ": "osoba", "pdf": "Rowery"},
        "Rowery MTB krótka przejażdżka (2-3h)": {"cena": get_price_from_df("Rowery MTB krótka przejażdżka (2-3 gozdiny)", df_c, 0), "typ": "osoba", "pdf": "Rowery"},
        "Rowery MTB 1 dzień": {"cena": get_price_from_df("Rowery MTB 1 dzień", df_c, 0), "typ": "osoba", "pdf": "Rowery"},
        "Rowery MTB 2 dni": {"cena": get_price_from_df("Rowery MTB 2 dni", df_c, 0), "typ": "osoba", "pdf": "Rowery"},
        "Rowery MTB 3 dni": {"cena": get_price_from_df("Rowery MTB 3 dni", df_c, 0), "typ": "osoba", "pdf": "Rowery"},
        "Ognisko": {"cena": get_price_from_df("Ognisko", df_c, 150), "typ": "grupa", "pdf": "Ognisko"},
        "Punkt widokowy": {"cena": get_price_from_df("Punkt widokowy", df_c, 0), "typ": "grupa", "pdf": "Punkt widokowy"},
        "Łączka cielaczków": {"cena": get_price_from_df("Łączka cielaczków", df_c, 0), "typ": "grupa", "pdf": "Łączka cielaczków"},
        "Atrakcje na wodzie": {"cena": get_price_from_df("Atrakcje na wodzie", df_c, 0), "typ": "grupa", "pdf": "Atrakcje na wodzie"},
        "Złów i wypuść": {"cena": get_price_from_df("Złow i wypuść", df_c, 0), "typ": "grupa", "pdf": "Złów i wypuść"},
        "Grzybobranie": {"cena": get_price_from_df("Grzybobranie", df_c, 0), "typ": "grupa", "pdf": "Grzybobranie"},
        "Roztańczony las": {"cena": get_price_from_df("Roztańczony las", df_c, 0), "typ": "grupa", "pdf": "Roztańczony las"},
        "Drawieński PN (3h)": {"cena": get_price_from_df("Drawieński PN (3h)", df_c, 0), "typ": "grupa", "pdf": "Drawieński PN"},
        "Drawieński PN (6h)": {"cena": get_price_from_df("Drawieński PN (6h)", df_c, 0), "typ": "grupa", "pdf": "Drawieński PN"}
    },
    "Biznes": {
        "Blok konferencyjny": {"cena": get_price_from_df("Blok konferencyjny", df_c, 0), "typ": "grupa", "pdf": "Blok konferencyjny"},
        "Wynajem sali": {"cena": get_price_from_df("Wynajem sali", df_c, 0), "typ": "grupa", "pdf": "Wynajem sali"},
        "Przejazd grupy (do 23 os.)": {"cena": get_price_from_df("Przejazd grupy (do 23 os.)", df_c, 0), "typ": "grupa", "pdf": "Przejazd"},
        "Przejazd grupy (do 50 os.)": {"cena": get_price_from_df("Przejazd grupy (do 50 os.)", df_c, 0), "typ": "grupa", "pdf": "Przejazd"}
    }
}
POKOJE_DWOREK = {f"Pokój nr {i}": (1 if i==1 else 4 if i==11 else 3 if i in [7,9,10,12] else 2) for i in range(1,13)}

# --- PANEL BOCZNY ---
with st.sidebar:
    st.subheader("🛠 Diagnostyka systemu")
    if wszystkie_pliki:
        st.success(f"✅ Połączono z Drive. Liczba plików: {len(wszystkie_pliki)}")
        if st.button("🔄 Wymuś ponowne skanowanie Dysku", type="primary"):
            fetch_all_debogora_files.clear()
            st.rerun()

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

tab1, tab2 = st.tabs(["📝 Kreator Ofert", "⚙️ Edycja Cennika Głównego"])

with tab2:
    st.subheader("Edycja pliku Cennika (Google Drive)")
    if df_c is not None:
        edited_df = st.data_editor(df_c, num_rows="dynamic", use_container_width=True)
        if st.button("💾 ZAPISZ ZMIANY NA DYSKU", type="primary"):
            with st.spinner("Zapisywanie na Dysku..."):
                try:
                    update_file_on_drive(cennik_file['id'], edited_df, cennik_file['name'])
                    st.session_state.df_cennik = edited_df
                    st.success("Zmiany zapisane!")
                except Exception as e: st.error(f"Błąd zapisu: {e}")

with tab1:
    pozycje_kosztowe = []

    with st.container():
        st.subheader("1. Główne Ustawienia i Dane Klienta")
        c1, c2 = st.columns(2)
        with c1:
            marka_oferty = st.selectbox("Marka wiodąca oferty *", ["Dwór Dębogóra", "Krovacja"])
            klient_imie = st.text_input("Imię i nazwisko osoby kontaktowej *")
            firma_n = st.text_input("Firma (opcjonalnie)")
            nip_n = st.text_input("NIP (opcjonalnie)")
            st.number_input("Liczba osób łącznie", 1, 100, key="l_osob_total")
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
        
        osoby_zadeklarowane = 0
        
        with col_dw:
            p_sel = st.multiselect("Dworek", list(POKOJE_DWOREK.keys()), key="wybrane_p")
            for p in p_sel:
                ile = st.number_input(f"{p} (max {POKOJE_DWOREK[p]} os.)", 1, POKOJE_DWOREK[p], key=f"os_{p}")
                osoby_zadeklarowane += ile
                pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"{p} (os: {ile})", "Ilość": ile, "Cena": stawka_dw*dni, "Suma": ile*stawka_dw*dni, "pdf_kw": "Zakwaterowanie"})
        with col_dm:
            d_sel = st.multiselect("Domki", list(CENNIK["domki"].keys()), key="wybrane_d")
            for d in d_sel:
                ile = st.number_input(f"{d} (max {CENNIK['domki'][d]['max_os']} os.)", 1, CENNIK["domki"][d]["max_os"], key=f"os_{d}")
                osoby_zadeklarowane += ile
                cena_d = (CENNIK["domki"][d]["baza"] + (max(0, ile-1)*CENNIK["doplata_domek"]))*dni
                pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"{d} ({ile} os.)", "Ilość": 1, "Cena": cena_d, "Suma": cena_d, "pdf_kw": CENNIK["domki"][d]["pdf"]})

        # BLOKADA OVERBOOKINGU
        overbooking_error = False
        if osoby_zadeklarowane > st.session_state.l_osob_total:
            st.error(f"⚠️ UWAGA: Przydzieliłeś miejsca dla {osoby_zadeklarowane} osób, a zadeklarowano łącznie {st.session_state.l_osob_total} miejsc! Musisz to poprawić.")
            overbooking_error = True
        elif osoby_zadeklarowane > 0 and osoby_zadeklarowane < st.session_state.l_osob_total:
            st.info(f"💡 Zostało {st.session_state.l_osob_total - osoby_zadeklarowane} osób bez przypisanego noclegu.")

    with st.container():
        st.subheader("3. Wyżywienie")
        wyz_sel = st.multiselect("Wybierz opcje wyżywienia", list(CENNIK["wyzywienie"].keys()))
        for w in wyz_sel:
            w_data = CENNIK["wyzywienie"][w]
            # Automatyczne liczenie posiłków: osoby x dni
            domyslna_ilosc_wyz = st.session_state.l_osob_total * dni
            ile = st.number_input(f"Ilość (domyślnie {st.session_state.l_osob_total} os. x {dni} dni): {w}", 1, 5000, domyslna_ilosc_wyz)
            pozycje_kosztowe.append({"Kategoria": "Gastronomia", "Opis": w, "Ilość": ile, "Cena": w_data["cena"], "Suma": ile*w_data["cena"], "pdf_kw": w_data["pdf"]})

    with st.container():
        st.subheader("4. Oferta Dodatkowa (SPAstwisko, Atrakcje, Biznes)")
        c_spa, c_atr, c_biz = st.columns(3)
        
        with c_spa:
            spa_sel = st.multiselect("SPAstwisko", list(CENNIK["SPAstwisko"].keys()))
            for a in spa_sel:
                a_data = CENNIK["SPAstwisko"][a]
                ile = st.number_input(f"Ilość: {a}", 1, 100, st.session_state.l_osob_total if a_data["typ"]=="osoba" else 1, key=f"spa_{a}")
                pozycje_kosztowe.append({"Kategoria": "SPAstwisko", "Opis": a, "Ilość": ile, "Cena": a_data["cena"], "Suma": ile*a_data["cena"], "pdf_kw": a_data["pdf"]})
                
        with c_atr:
            atr_sel = st.multiselect("Atrakcje", list(CENNIK["Atrakcje"].keys()))
            for a in atr_sel:
                a_data = CENNIK["Atrakcje"][a]
                ile = st.number_input(f"Ilość: {a}", 1, 100, st.session_state.l_osob_total if a_data["typ"]=="osoba" else 1, key=f"atr_{a}")
                pozycje_kosztowe.append({"Kategoria": "Atrakcje", "Opis": a, "Ilość": ile, "Cena": a_data["cena"], "Suma": ile*a_data["cena"], "pdf_kw": a_data["pdf"]})

        with c_biz:
            biz_sel = st.multiselect("Biznes", list(CENNIK["Biznes"].keys()))
            for a in biz_sel:
                a_data = CENNIK["Biznes"][a]
                ile = st.number_input(f"Ilość: {a}", 1, 100, st.session_state.l_osob_total if a_data["typ"]=="osoba" else 1, key=f"biz_{a}")
                pozycje_kosztowe.append({"Kategoria": "Biznes", "Opis": a, "Ilość": ile, "Cena": a_data["cena"], "Suma": ile*a_data["cena"], "pdf_kw": a_data["pdf"]})

    with st.container():
        st.subheader("5. Kosztorys i Eksport")
        df = pd.DataFrame(pozycje_kosztowe)
        if not df.empty:
            edf = st.data_editor(df, use_container_width=True, num_rows="dynamic")
            razem = edf["Suma"].sum()
            st.markdown(f"<h3 style='color: {CI['dark_green']};'>RAZEM DO ZAPŁATY: {razem:,.2f} PLN</h3>".replace(",", " "), unsafe_allow_html=True)
            
            if st.button("GENERUJ FINALNY PDF", disabled=overbooking_error):
                if not klient_imie:
                    st.error("Podaj imię i nazwisko klienta (Pole z gwiazdką)!")
                else:
                    with st.spinner("Pobieranie plików, zamiana danych i układanie oferty w idealnej kolejności..."):
                        merger = PdfWriter()
                        open_streams = []
                        missing_cards = []
                        
                        nazwa_docelowa = firma_n if firma_n else klient_imie
                        has_dworek = any("pokój" in str(row["Opis"]).lower() for _, row in edf.iterrows())
                        has_krovacja = any("muuu" in str(row["Opis"]).lower() for _, row in edf.iterrows())
                        if has_dworek and has_krovacja: zakwaterowanie_txt = "domkach i pokojach"
                        elif has_krovacja: zakwaterowanie_txt = "domkach"
                        else: zakwaterowanie_txt = "pokojach"
                            
                        wybrane_atrakcje = [row["Opis"] for _, row in edf.iterrows() if row["Kategoria"] in ["SPAstwisko", "Atrakcje", "Biznes"]]
                        if len(wybrane_atrakcje) >= 2: atrakcje_txt = f"{wybrane_atrakcje[0]} oraz {wybrane_atrakcje[1]}"
                        elif len(wybrane_atrakcje) == 1: atrakcje_txt = f"{wybrane_atrakcje[0]} oraz otaczającą nas naturę"
                        else: atrakcje_txt = "spokój, ciszę oraz bliskość natury"
                            
                        marka_wstawka = "Krovację" if marka_oferty == "Krovacja" else "Dwór Dębogóra"
                        
                        replacements = {
                            "nazwa firmy": nazwa_docelowa,
                            "Dwór Dębogóra/Krovację": marka_wstawka,
                            "Dwór Dębogóra / Krovację": marka_wstawka,
                            "domkach/pokojach": zakwaterowanie_txt,
                            "domkach / pokojach": zakwaterowanie_txt,
                            "atrakcja oraz atrakcja": atrakcje_txt,
                            "atrakcja i atrakcja": atrakcje_txt,
                            "Jest nam": "Jest nam",
                            "stada!": "stada!"
                        }
                        
                        # --- 1. OKŁADKA (Okładka 02) ---
                        add_file_to_merger(merger, "okładka", wszystkie_pliki, open_streams, missing_cards, replacements)

                        # --- 2. KARTA POWITALNA ---
                        add_file_to_merger(merger, "powitalna", wszystkie_pliki, open_streams, missing_cards, replacements)

                        # --- 3. ZAKWATEROWANIE ---
                        has_zakwaterowanie = any(row["pdf_kw"] == "Zakwaterowanie" for _, row in edf.iterrows())
                        if has_zakwaterowanie: 
                            add_file_to_merger(merger, "Zakwaterowanie", wszystkie_pliki, open_streams, missing_cards, replacements)

                        # --- 4. WYŻYWIENIE ---
                        has_wyzywienie = any(row["Kategoria"] == "Gastronomia" for _, row in edf.iterrows())
                        if has_wyzywienie: 
                            add_file_to_merger(merger, "wyżywienie", wszystkie_pliki, open_streams, missing_cards, replacements)

                        # --- 5. ATRAKCJE - WSTĘP ---
                        add_file_to_merger(merger, "atrakcje_wstęp", wszystkie_pliki, open_streams, missing_cards, replacements)

                        # --- 6. INDYWIDUALNE KARTY ATRAKCJI ---
                        atrakcje_kws = list(set([row["pdf_kw"] for _, row in edf.iterrows() if row["Kategoria"] in ["SPAstwisko", "Atrakcje", "Biznes"] and pd.notna(row["pdf_kw"])]))
                        for kw in atrakcje_kws:
                            add_file_to_merger(merger, kw, wszystkie_pliki, open_streams, missing_cards, replacements)

                        # --- 7. WYCENA (TABELA REPORTLAB + TŁO Z PPTX) ---
                        buf = io.BytesIO()
                        # Krajobrazowy format, żeby lepiej pasował do prezentacji (odsunięty od góry)
                        doc = SimpleDocTemplate(buf, pagesize=landscape(A4), rightMargin=40, leftMargin=40, topMargin=120, bottomMargin=50)
                        elements = []
                        styles = getSampleStyleSheet()
                        
                        t_data = [["Kategoria", "Opis usługi", "Ilość", "Suma"]]
                        for _, row in edf.iterrows():
                            kat = safe_str(row["Kategoria"])
                            opis = safe_str(row["Opis"])
                            ilosc = safe_str(row["Ilość"])
                            suma_str = f"{row['Suma']:,.0f} zł".replace(",", " ")
                            t_data.append([kat, opis, ilosc, suma_str])
                        t_data.append(["", "", "RAZEM:", f"{razem:,.0f} zł".replace(",", " ")])
                        
                        table = Table(t_data, colWidths=[150, 400, 70, 100])
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
                            open_streams.append(buf)
                            
                            # Logika podkładania tła z pliku Wycena
                            wycena_file = get_file_by_keyword("wycena", wszystkie_pliki)
                            if wycena_file:
                                wycena_bg_stream = process_file_to_pdf_stream(wycena_file, replacements)
                                open_streams.append(wycena_bg_stream)
                                bg_reader = PdfReader(wycena_bg_stream)
                                fg_reader = PdfReader(buf) # wygenerowana tabela
                                
                                for i, fg_page in enumerate(fg_reader.pages):
                                    bg_idx = min(i, len(bg_reader.pages) - 1)
                                    bg_page = bg_reader.pages[bg_idx]
                                    # Scalanie strony z tabelą ze stroną z tłem (PPTX)
                                    new_page = PageObject.create_blank_page(width=bg_page.mediabox.width, height=bg_page.mediabox.height)
                                    new_page.merge_page(bg_page)
                                    new_page.merge_page(fg_page)
                                    merger.append(new_page)
                            else:
                                merger.append(PdfReader(buf, strict=False))
                        except Exception as e:
                            st.error(f"Błąd generatora tabeli wyceny: {e}")

                        # --- 8. AGENDA ---
                        add_file_to_merger(merger, "agenda", wszystkie_pliki, open_streams, missing_cards, replacements)
                        
                        # --- 9. KONTAKT ---
                        add_file_to_merger(merger, "kontakt", wszystkie_pliki, open_streams, missing_cards, replacements)

                        if missing_cards:
                            st.warning(f"⚠️ Uwaga: Na Dysku Google nie odnaleziono następujących kart (pominęto w PDF): {', '.join(missing_cards)}")

                        # ZAPIS DO PLIKU KOŃCOWEGO
                        final_pdf = io.BytesIO()
                        merger.write(final_pdf)
                        
                        safe_filename = "".join([c for c in safe_str(klient_imie) if c.isalpha() or c.isdigit() or c==' ']).rstrip().replace(" ", "_")
                        st.download_button("📥 POBIERZ SCALONĄ OFERTĘ PDF", final_pdf.getvalue(), f"Oferta_{safe_filename}.pdf", "application/pdf", type="primary")
