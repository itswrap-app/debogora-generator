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
from pypdf import PdfWriter, PdfReader

# ReportLab
from reportlab.lib.pagesizes import A4
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
    except Exception: pass

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
        except Exception: pass
    return all_files

def download_file(file_id, retries=3):
    service = get_drive_service()
    for attempt in range(retries):
        try:
            request = service.files().get_media(fileId=file_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request, chunksize=256*1024)
            done = False
            while done is False:
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
    if hasattr(shape, "text_frame") and shape.text_frame is not None:
        for paragraph in shape.text_frame.paragraphs:
            if not paragraph.runs: continue
            full_text = "".join(run.text for run in paragraph.runs)
            orig = full_text
            for k, v in replacements.items():
                full_text = full_text.replace(k, str(v))
            if full_text != orig:
                paragraph.runs[0].text = full_text
                for i in range(1, len(paragraph.runs)):
                    paragraph.runs[i].text = ""
    if hasattr(shape, "shapes"):
        for subshape in shape.shapes:
            process_shape(subshape, replacements)
    if hasattr(shape, "has_table") and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                process_shape(cell, replacements)

def replace_text_in_pptx(prs, replacements):
    for slide in prs.slides:
        for shape in slide.shapes:
            process_shape(shape, replacements)

# --- MAPOWANIE NAZW I WYSZUKIWANIE PLIKÓW ---
def normalize_pl(text):
    rep = {'ą':'a', 'ć':'c', 'ę':'e', 'ł':'l', 'ń':'n', 'ó':'o', 'ś':'s', 'ź':'z', 'ż':'z'}
    res = str(text).lower()
    for k, v in rep.items(): res = res.replace(k, v)
    return res

SYNONYMS = {
    "okładka": "okładka_02", "powitalna": "karta powitalna", "dworek": "Zakwaterowanie_02", "krovacja": "Zakwaterowanie_02",
    "wyżywienie": "wyżywienie", "atrakcje_wstęp": "atrakcje", "wycena": "wycena", "agenda": "agenda", "kontakt": "kontakt",
    "ZłodziejKrów": "złodziej", "Łowcy krów": "złodziej", "Skarby": "skarby", "Krowie Safari": "safari",
    "Safari_Standard": "safari", "Safari_Rozszerzona": "rozszerzona", "Seans saunowy": "saunowy", "Sauna olchowa": "olchowa",
    "Staw": "staw", "Balia": "balia", "Sauny": "sauny", "Masaże": "masaż", "Paintball": "paintball", "Spływ kajakowy": "kajak",
    "Rowery": "rowery", "Ognisko": "ognisko", "Punkt widokowy": "widokowy", "Łączka cielaczków": "cielacz",
    "Atrakcje na wodzie": "wodzie", "Złów i wypuść": "złów", "Grzybobranie": "grzyb", "Roztańczony las": "roztańczony",
    "Drawieński PN": "drawieński", "Blok konferencyjny": "konferencyjny", "Wynajem sali": "sali", "Przejazd": "przejazd"
}

def get_file_by_keyword(keyword, all_files):
    norm_search = normalize_pl(SYNONYMS.get(keyword, keyword))
    matches = [f for f in all_files if norm_search in normalize_pl(f['name'])]
    if matches:
        prev_matches = [f for f in matches if 'prev' in f['name'].lower()]
        return prev_matches[0] if prev_matches else matches[0]
    return None

def add_file_to_merger(merger, keyword, all_files, open_streams, missing_cards, replacements=None):
    if not keyword: return
    file_obj = get_file_by_keyword(keyword, all_files)
    if file_obj:
        try:
            fh = download_file(file_obj['id'])
            fname = file_obj['name'].lower()
            if 'ppt' in fname or 'presentation' in file_obj['mimeType']:
                temp_ppt = f"temp_{file_obj['id']}.pptx"
                temp_pdf = f"temp_{file_obj['id']}.pdf"
                with open(temp_ppt, "wb") as f: f.write(fh.getvalue())
                if replacements:
                    prs = Presentation(temp_ppt)
                    replace_text_in_pptx(prs, replacements)
                    prs.save(temp_ppt)
                
                res = subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", temp_ppt], capture_output=True, text=True)
                if res.returncode != 0:
                    raise Exception(f"LibreOffice błąd: {res.stderr}")
                    
                with open(temp_pdf, "rb") as f: pdf_bytes = f.read()
                pdf_stream = io.BytesIO(pdf_bytes)
            else:
                pdf_stream = fh
            open_streams.append(pdf_stream)
            merger.append(PdfReader(pdf_stream, strict=False))
        except Exception as e:
            st.error(f"⚠️ Pominięto '{keyword}'. Błąd pliku '{file_obj['name']}': {e}")
            missing_cards.append(keyword)
    else:
        if keyword not in missing_cards: missing_cards.append(keyword)

def safe_str(text): return "" if pd.isna(text) else str(text).strip()

def get_price_data(usluga_name, df):
    if df is None or df.empty: return {"cena": 0, "czas": 0.0}
    try:
        col_name = next((c for c in df.columns if 'nazwa' in c.lower() or 'usługa' in c.lower() or 'usluga' in c.lower()), None)
        col_price = next((c for c in df.columns if 'cena' in c.lower()), None)
        col_czas = next((c for c in df.columns if 'długość' in c.lower() or 'dlugosc' in c.lower()), None)
        
        if col_name and col_price:
            match = df[df[col_name].astype(str).str.strip().str.lower() == usluga_name.lower()]
            if not match.empty:
                val = match.iloc[0][col_price]
                cena = float(val) if pd.notna(val) else 0.0
                
                czas_str = str(match.iloc[0][col_czas]) if col_czas and pd.notna(match.iloc[0][col_czas]) else "0"
                czas_str = czas_str.lower().replace('h', '').replace('godz', '').strip()
                if ',' in czas_str:
                    parts = czas_str.split(',')
                    if len(parts)>1 and parts[1] == '15': czas_num = float(parts[0]) + 0.25
                    elif len(parts)>1 and parts[1] == '30': czas_num = float(parts[0]) + 0.5
                    elif len(parts)>1 and parts[1] == '45': czas_num = float(parts[0]) + 0.75
                    else: czas_num = float(czas_str.replace(',', '.'))
                else:
                    try: czas_num = float(czas_str)
                    except: czas_num = 0.0
                
                return {"cena": cena, "czas": czas_num}
    except Exception: pass
    return {"cena": 0, "czas": 0.0}

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
        except Exception: st.session_state.df_cennik = None
    else:
        try: st.session_state.df_cennik = pd.read_csv(file_stream, encoding='utf-8')
        except: 
            file_stream.seek(0)
            st.session_state.df_cennik = pd.read_csv(file_stream, encoding='cp1250')
elif 'df_cennik' not in st.session_state:
    st.session_state.df_cennik = None

df_c = st.session_state.df_cennik

CENNIK = {
    "nocleg_1_noc": get_price_data("Nocleg (1 noc)", df_c)["cena"],
    "nocleg_2_noce": get_price_data("Nocleg (2+ noce)", df_c)["cena"],
    "doplata_domek": 40,
    "domki": {
        "Muuu 1": {"baza": get_price_data("Muuu 1, 2", df_c)["cena"], "max_os": 4, "pdf": "krovacja"}, 
        "Muuu 2": {"baza": get_price_data("Muuu 1, 2", df_c)["cena"], "max_os": 4, "pdf": "krovacja"},
        "Muuu 3": {"baza": get_price_data("Muuu 3, 4", df_c)["cena"], "max_os": 6, "pdf": "krovacja"}, 
        "Muuu 4": {"baza": get_price_data("Muuu 3, 4", df_c)["cena"], "max_os": 6, "pdf": "krovacja"},
        "Muuu 5": {"baza": get_price_data("Muuu 5, 6", df_c)["cena"], "max_os": 3, "pdf": "krovacja"}, 
        "Muuu 6": {"baza": get_price_data("Muuu 5, 6", df_c)["cena"], "max_os": 3, "pdf": "krovacja"}
    },
    "wyzywienie": {
        "Śniadanie": {"dane": get_price_data("Śniadanie", df_c), "pdf": "wyżywienie"},
        "Obiadokolacja": {"dane": get_price_data("Obiadokolacja", df_c), "pdf": "wyżywienie"},
        "Serwis kawowy": {"dane": get_price_data("Serwis kawowy", df_c), "pdf": "wyżywienie"},
        "Wiejskie jadło": {"dane": get_price_data("Wiejskie jadło (Podstawowe)", df_c), "pdf": "wyżywienie"},
        "Kolacja z rozszerzonym menu": {"dane": get_price_data("Biesiada wieczorna", df_c), "pdf": "wyżywienie"}
    },
    "SPAstwisko": {
        "Seans full experience": {"dane": get_price_data("Seans full experience", df_c), "typ": "grupa", "pdf": "Seans saunowy"},
        "Sauna olchowa": {"dane": get_price_data("Sauna Olchowa", df_c), "typ": "grupa", "pdf": "Sauna olchowa"},
        "Staw kąpielowy": {"dane": get_price_data("Staw kąpielowy", df_c), "typ": "grupa", "pdf": "Staw"},
        "Balia opalana drewnem": {"dane": get_price_data("Balia opalana drewnem", df_c), "typ": "grupa", "pdf": "Balia"},
        "Sauny": {"dane": get_price_data("Wynajem na wyłączność", df_c), "typ": "grupa", "pdf": "Sauny"},
        "Masaż relaksacyjny": {"dane": get_price_data("Masaż relaksacyjny", df_c), "typ": "osoba", "pdf": "Masaże"},
        "Masaż gorącą świecą": {"dane": get_price_data("Masaż gorącą świecą", df_c), "typ": "osoba", "pdf": "Masaże"},
    },
    "Atrakcje": {
        "Łowcy krów": {"dane": get_price_data("Łowcy krów", df_c), "typ": "osoba", "pdf": "Łowcy krów"},
        "Skarby Dębogóry": {"dane": get_price_data("Skarby Dębogóry", df_c), "typ": "osoba", "pdf": "Skarby"},
        "Krowie Safari Standard": {"dane": get_price_data("Krowie Safari Standard", df_c), "typ": "osoba", "pdf": "Krowie Safari"},
        "Krowie Safari Rozszerzone": {"dane": get_price_data("Krowie Safari Rozszerzone", df_c), "typ": "osoba", "pdf": "Krowie Safari"},
        "Paintball": {"dane": get_price_data("Paintball", df_c), "typ": "osoba", "pdf": "Paintball"},
        "Kajaki": {"dane": get_price_data("Kajaki", df_c), "typ": "osoba", "pdf": "Spływ kajakowy"},
        "Ognisko": {"dane": get_price_data("Ognisko", df_c), "typ": "grupa", "pdf": "Ognisko"},
        "Rowery elektryczne 1 dzień": {"dane": get_price_data("Rowery elektryczne 1 dzień", df_c), "typ": "osoba", "pdf": "Rowery"},
        "Punkt widokowy": {"dane": get_price_data("Punkt widokowy", df_c), "typ": "grupa", "pdf": "Punkt widokowy"},
    },
    "Biznes": {
        "Blok konferencyjny": {"dane": get_price_data("Blok konferencyjny", df_c), "typ": "grupa", "pdf": "Blok konferencyjny"},
        "Wynajem sali": {"dane": get_price_data("Wynajem sali", df_c), "typ": "grupa", "pdf": "Wynajem sali"},
    }
}
POKOJE_DWOREK = {f"Pokój nr {i}": (1 if i==1 else 4 if i==11 else 3 if i in [7,9,10,12] else 2) for i in range(1,13)}

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
            update_file_on_drive(cennik_file['id'], edited_df, cennik_file['name'])
            st.session_state.df_cennik = edited_df
            st.success("Zmiany zapisane!")

with tab1:
    pozycje_kosztowe = []
    wybrane_atrakcje_agenda = []

    with st.container():
        st.subheader("1. Główne Ustawienia i Dane Klienta")
        c1, c2 = st.columns(2)
        with c1:
            marka_oferty = st.selectbox("Marka wiodąca oferty *", ["Dwór Dębogóra", "Krovacja"])
            klient_imie = st.text_input("Imię i nazwisko osoby kontaktowej *")
            firma_n = st.text_input("Firma (opcjonalnie)")
            st.number_input("Liczba osób", 1, 100, key="l_osob_total")
            st.button("🤖 Automatycznie rozmieść gości", on_click=auto_alloc)
        with c2:
            email_n = st.text_input("Email")
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
                ile = st.number_input(f"{p}", 1, POKOJE_DWOREK[p], key=f"os_{p}")
                osoby_zadeklarowane += ile
                pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"{p}", "Ilość": ile, "Cena": stawka_dw*dni, "Suma": ile*stawka_dw*dni, "pdf_kw": "dworek"})
        with col_dm:
            d_sel = st.multiselect("Domki", list(CENNIK["domki"].keys()), key="wybrane_d")
            for d in d_sel:
                ile = st.number_input(f"{d}", 1, CENNIK["domki"][d]["max_os"], key=f"os_{d}")
                osoby_zadeklarowane += ile
                cena_d = (CENNIK["domki"][d]["baza"] + (max(0, ile-1)*CENNIK["doplata_domek"]))*dni
                pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"{d}", "Ilość": 1, "Cena": cena_d, "Suma": cena_d, "pdf_kw": CENNIK["domki"][d]["pdf"]})

        overbooking_error = False
        if osoby_zadeklarowane > st.session_state.l_osob_total:
            st.error(f"⚠️ UWAGA: Przydzieliłeś {osoby_zadeklarowane} miejsc dla {st.session_state.l_osob_total} zadeklarowanych gości. Zmniejsz ilość w noclegach!")
            overbooking_error = True

    with st.container():
        st.subheader("3. Wyżywienie")
        wyz_sel = st.multiselect("Wybierz opcje wyżywienia", list(CENNIK["wyzywienie"].keys()))
        for w in wyz_sel:
            w_data = CENNIK["wyzywienie"][w]["dane"]
            ile = st.number_input(f"Ilość porcji: {w}", 1, 5000, st.session_state.l_osob_total * dni)
            pozycje_kosztowe.append({"Kategoria": "Gastronomia", "Opis": w, "Ilość": ile, "Cena": w_data["cena"], "Suma": ile*w_data["cena"], "pdf_kw": CENNIK["wyzywienie"][w]["pdf"]})

    with st.container():
        st.subheader("4. Oferta Dodatkowa (SPAstwisko, Atrakcje, Biznes)")
        c_spa, c_atr, c_biz = st.columns(3)
        
        with c_spa:
            spa_sel = st.multiselect("SPAstwisko", list(CENNIK["SPAstwisko"].keys()))
            for a in spa_sel:
                a_data = CENNIK["SPAstwisko"][a]
                ile = st.number_input(f"Ilość: {a}", 1, 100, st.session_state.l_osob_total if a_data["typ"]=="osoba" else 1, key=f"spa_{a}")
                pozycje_kosztowe.append({"Kategoria": "SPAstwisko", "Opis": a, "Ilość": ile, "Cena": a_data["dane"]["cena"], "Suma": ile*a_data["dane"]["cena"], "pdf_kw": a_data["pdf"]})
                wybrane_atrakcje_agenda.append({"Nazwa": a, "Czas": a_data["dane"]["czas"]})
                
        with c_atr:
            atr_sel = st.multiselect("Atrakcje", list(CENNIK["Atrakcje"].keys()))
            for a in atr_sel:
                a_data = CENNIK["Atrakcje"][a]
                ile = st.number_input(f"Ilość: {a}", 1, 100, st.session_state.l_osob_total if a_data["typ"]=="osoba" else 1, key=f"atr_{a}")
                pozycje_kosztowe.append({"Kategoria": "Atrakcje", "Opis": a, "Ilość": ile, "Cena": a_data["dane"]["cena"], "Suma": ile*a_data["dane"]["cena"], "pdf_kw": a_data["pdf"]})
                wybrane_atrakcje_agenda.append({"Nazwa": a, "Czas": a_data["dane"]["czas"]})

        with c_biz:
            biz_sel = st.multiselect("Biznes", list(CENNIK["Biznes"].keys()))
            for a in biz_sel:
                a_data = CENNIK["Biznes"][a]
                ile = st.number_input(f"Ilość: {a}", 1, 100, st.session_state.l_osob_total if a_data["typ"]=="osoba" else 1, key=f"biz_{a}")
                pozycje_kosztowe.append({"Kategoria": "Biznes", "Opis": a, "Ilość": ile, "Cena": a_data["dane"]["cena"], "Suma": ile*a_data["dane"]["cena"], "pdf_kw": a_data["pdf"]})
                wybrane_atrakcje_agenda.append({"Nazwa": a, "Czas": a_data["dane"]["czas"]})

    with st.container():
        st.subheader("5. Generator Harmonogramu (Agenda)")
        
        # Obliczenie fizycznego czasu dostępnego podczas pobytu
        # Zakładamy 4h luzu pierwszego dnia po południu i 6h każdego następnego.
        dostepny_czas = 4.0 if dni == 1 else (4.0 + (dni-1)*6.0)
        zajety_czas = sum(item["Czas"] for item in wybrane_atrakcje_agenda)
        
        if zajety_czas > dostepny_czas:
            st.error(f"❌ WYKRYTO KONFLIKT CZASU: Wybrane atrakcje wymagają łącznie ok. {zajety_czas}h. Przewidywany dostępny czas aktywności podczas tego pobytu to ok. {dostepny_czas}h. Czasu może zabraknąć!")
            agenda_overload = True
        else:
            st.success(f"✅ Czas zaplanowany na atrakcje: {zajety_czas}h (W normie)")
            agenda_overload = False
        
        # Generowanie propozycji Agendy
        draft_agenda = f"TERMIN WYDARZENIA: {d_in.strftime('%d.%m.%Y')} - {d_out.strftime('%d.%m.%Y')}\n\n"
        
        if dni == 1:
            draft_agenda += "DZIEŃ 1\n"
            draft_agenda += "10:00 - Przyjazd i rozpoczęcie spotkania\n"
            for atr in wybrane_atrakcje_agenda:
                draft_agenda += f"- {atr['Nazwa']} (Czas: ~{atr['Czas']}h)\n"
            draft_agenda += "18:00 - Obiadokolacja / Zakończenie\n"
        else:
            draft_agenda += f"DZIEŃ 1 ({d_in.strftime('%d.%m')})\n"
            draft_agenda += "15:00 - Przyjazd i Zakwaterowanie\n"
            if wybrane_atrakcje_agenda: draft_agenda += f"16:00 - {wybrane_atrakcje_agenda[0]['Nazwa']}\n"
            draft_agenda += "18:00 - Obiadokolacja\n\n"
            
            draft_agenda += f"DZIEŃ 2 ({ (d_in + timedelta(days=1)).strftime('%d.%m') })\n"
            draft_agenda += "09:00 - Śniadanie\n"
            
            # Wrzucenie pozostałych atrakcji
            if len(wybrane_atrakcje_agenda) > 1:
                start_h = 10
                for atr in wybrane_atrakcje_agenda[1:]:
                    draft_agenda += f"{start_h}:00 - {atr['Nazwa']}\n"
                    start_h += max(1, int(atr['Czas']))
            
            draft_agenda += "18:00 - Obiadokolacja\n\n"
            
            draft_agenda += f"DZIEŃ {dni} (Wyjazd)\n"
            draft_agenda += "09:00 - Śniadanie\n"
            draft_agenda += "11:00 - Wymeldowanie i zakończenie pobytu\n"

        st.info("Poniższy harmonogram zostanie wklejony na stronę 'Agenda' w PDF (w miejsce klamry {{agenda}}). Możesz go w pełni zmodyfikować i dopisać dowolny tekst ręcznie!")
        final_agenda_text = st.text_area("Edytuj treść Harmonogramu:", value=draft_agenda, height=300)

    with st.container():
        st.subheader("6. Kosztorys i Eksport")
        df = pd.DataFrame(pozycje_kosztowe)
        if not df.empty:
            edf = st.data_editor(df, use_container_width=True, num_rows="dynamic")
            razem = edf["Suma"].sum()
            st.markdown(f"<h3 style='color: {CI['dark_green']};'>RAZEM DO ZAPŁATY: {razem:,.2f} PLN</h3>".replace(",", " "), unsafe_allow_html=True)
            
            if st.button("GENERUJ FINALNY PDF", disabled=(overbooking_error or agenda_overload)):
                if not klient_imie: st.error("Podaj imię i nazwisko klienta!")
                else:
                    with st.spinner("Pobieranie i kompilacja plików (To potrwa kilkanaście sekund)..."):
                        # MEGA KLAUZULA BEZPIECZEŃSTWA - ZAPOBIEGA BIAŁEMU EKRANOWI
                        try:
                            merger = PdfWriter()
                            open_streams = []
                            missing_cards = []
                            
                            nazwa_docelowa = firma_n if firma_n else klient_imie
                            has_dworek = any(row["pdf_kw"] == "dworek" for _, row in edf.iterrows())
                            has_krovacja = any(row["pdf_kw"] == "krovacja" for _, row in edf.iterrows())
                            if has_dworek and has_krovacja: zakwaterowanie_txt = "domkach i pokojach"
                            elif has_krovacja: zakwaterowanie_txt = "domkach"
                            else: zakwaterowanie_txt = "pokojach"
                                
                            atr_list = [row["Opis"] for _, row in edf.iterrows() if row["Kategoria"] in ["SPAstwisko", "Atrakcje", "Biznes"]]
                            if len(atr_list) >= 2: atrakcje_txt = f"{atr_list[0]} oraz {atr_list[1]}"
                            elif len(atr_list) == 1: atrakcje_txt = f"{atr_list[0]} oraz otaczającą nas naturę"
                            else: atrakcje_txt = "spokój, ciszę oraz bliskość natury"
                                
                            marka_wstawka = "Krovację" if marka_oferty == "Krovacja" else "Dwór Dębogóra"
                            
                            replacements = {
                                "{{nazwa firmy}}": nazwa_docelowa,
                                "{{Dwór Dębogóra/Krovację}}": marka_wstawka,
                                "{{Dwór Dębogóra / Krovację}}": marka_wstawka,
                                "{{domkach/pokojach}}": zakwaterowanie_txt,
                                "{{domkach / pokojach}}": zakwaterowanie_txt,
                                "{{atrakcja}} oraz {{atrakcja}}": atrakcje_txt,
                                "{{atrakcja}} i {{atrakcja}}": atrakcje_txt,
                                "{{Jest nam": "Jest nam",
                                "stada!}}": "stada!",
                                "{{agenda}}": final_agenda_text # TUTAJ WKLEJA ZBUDOWANY HARMONOGRAM
                            }
                            
                            add_file_to_merger(merger, "okładka", wszystkie_pliki, open_streams, missing_cards, replacements)
                            add_file_to_merger(merger, "powitalna", wszystkie_pliki, open_streams, missing_cards, replacements)
                            
                            if has_dworek: add_file_to_merger(merger, "dworek", wszystkie_pliki, open_streams, missing_cards, replacements)
                            if has_krovacja: add_file_to_merger(merger, "krovacja", wszystkie_pliki, open_streams, missing_cards, replacements)

                            if any(row["Kategoria"] == "Gastronomia" for _, row in edf.iterrows()): 
                                add_file_to_merger(merger, "wyżywienie", wszystkie_pliki, open_streams, missing_cards, replacements)

                            add_file_to_merger(merger, "atrakcje_wstęp", wszystkie_pliki, open_streams, missing_cards, replacements)

                            for kw in list(set([row["pdf_kw"] for _, row in edf.iterrows() if row["Kategoria"] in ["SPAstwisko", "Atrakcje", "Biznes"] and pd.notna(row["pdf_kw"])])):
                                add_file_to_merger(merger, kw, wszystkie_pliki, open_streams, missing_cards, replacements)

                            buf = io.BytesIO()
                            doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=120, bottomMargin=50)
                            elements = []
                            styles = getSampleStyleSheet()

                            t_data = [["Kategoria", "Opis usługi", "Ilość", "Suma"]]
                            for _, row in edf.iterrows():
                                t_data.append([safe_str(row["Kategoria"]), safe_str(row["Opis"]), safe_str(row["Ilość"]), f"{row['Suma']:,.0f} zł".replace(",", " ")])
                            t_data.append(["", "", "RAZEM:", f"{razem:,.0f} zł".replace(",", " ")])
                            
                            table = Table(t_data, colWidths=[100, 230, 50, 90])
                            t_style = TableStyle([
                                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(CI['dark_green'])), ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                                ('FONTNAME', (0, 0), (-1, 0), FONT_HEADER), ('FONTSIZE', (0, 0), (-1, 0), 12),
                                ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('ALIGN', (1, 1), (1, -2), 'LEFT'),
                                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('BOTTOMPADDING', (0, 0), (-1, -1), 10), ('TOPPADDING', (0, 0), (-1, -1), 10),
                                ('FONTNAME', (0, 1), (-1, -1), FONT_TEXT), ('FONTNAME', (2, -1), (-1, -1), FONT_TEXT_BOLD),
                                ('TEXTCOLOR', (2, -1), (-1, -1), colors.HexColor(CI['dark_green'])), ('BACKGROUND', (2, -1), (-1, -1), colors.HexColor(CI['light_green'])),
                                ('LINEABOVE', (0, -1), (-1, -1), 1, colors.HexColor(CI['dark_green'])),
                            ])
                            for i in range(1, len(t_data) - 1):
                                t_style.add('BACKGROUND', (0, i), (-1, i), colors.HexColor(CI['light_green']) if i % 2 == 0 else colors.white)
                                t_style.add('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#d1d9cf"))
                            table.setStyle(t_style)
                            elements.append(table)
                            doc.build(elements)
                            buf.seek(0)
                            open_streams.append(buf)
                            
                            wycena_file = get_file_by_keyword("wycena", wszystkie_pliki)
                            if wycena_file:
                                try:
                                    fh = download_file(wycena_file['id'])
                                    with open("temp_wycena.pptx", "wb") as f: f.write(fh.getvalue())
                                    prs = Presentation("temp_wycena.pptx")
                                    replace_text_in_pptx(prs, replacements)
                                    prs.save("temp_wycena.pptx")
                                    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "temp_wycena.pptx"], check=True)
                                    with open("temp_wycena.pdf", "rb") as f: bg_bytes = f.read()
                                    bg_stream = io.BytesIO(bg_bytes)
                                    open_streams.append(bg_stream)
                                    
                                    bg_reader = PdfReader(bg_stream)
                                    fg_reader = PdfReader(buf)
                                    for i, fg_page in enumerate(fg_reader.pages):
                                        bg_idx = min(i, len(bg_reader.pages) - 1)
                                        bg_page = bg_reader.pages[bg_idx]
                                        bg_page.merge_page(fg_page)
                                        merger.add_page(bg_page)
                                except Exception: merger.append(PdfReader(buf, strict=False))
                            else: merger.append(PdfReader(buf, strict=False))

                            add_file_to_merger(merger, "agenda", wszystkie_pliki, open_streams, missing_cards, replacements)
                            add_file_to_merger(merger, "kontakt", wszystkie_pliki, open_streams, missing_cards, replacements)

                            if missing_cards: st.warning(f"⚠️ Uwaga: Na Dysku Google nie odnaleziono kart: {', '.join(missing_cards)}")

                            final_pdf = io.BytesIO()
                            merger.write(final_pdf)
                            st.download_button("📥 POBIERZ SCALONĄ OFERTĘ PDF", final_pdf.getvalue(), f"Oferta_{safe_str(klient_imie).replace(' ', '_')}.pdf", "application/pdf", type="primary")

                        except Exception as global_error:
                            st.error("❌ KRYTYCZNY BŁĄD PODCZAS GENEROWANIA PDF!")
                            st.error(f"Treść błędu (zrób screena i wyślij mi go): {str(global_error)}")
                            st.info("Prawdopodobne przyczyny: 1) Brak pliku packages.txt na GitHubie z wpisanym 'libreoffice', 2) Uszkodzony plik PowerPoint, 3) Błąd pamięci serwera.")
