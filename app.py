import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import os
import subprocess
import base64
import time
import re
import glob
import requests
import json

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

st.set_page_config(page_title="PLANER OFERT - Dwór Dębogóra", layout="wide")

CI = {
    "dark_green": "#00622f",
    "light_green": "#e8ece6",
    "gray": "#333333",
    "white": "#ffffff"
}

# --- IDENTYFIKATORY DYSKU GOOGLE ---
ROOT_FOLDER_ID = "1tU6mo1YWpTep8vl5CRR5DhsZAINeWnHz"  
BAZA_OFERT_FOLDER_ID = "1i_a2UkK73ixyvMBe5l9SkE5vpqAu6he5" 

# =========================================================
# MAPOWANIE POKOI HOTRES
# =========================================================
HOTRES_ROOM_MAP = {
    29952: "Krovacja - cały kompleks",
    37951: "Muuu 6",
    37950: "Muuu 5",
    25074: "Muuu 4",
    25073: "Muuu 3",
    25072: "Muuu 2",
    25071: "Muuu 1",
    27589: "Ognisko #1",
    27588: "Strefa relaksu - balia #1",
    31294: "Strefa relaksu - balia #2",
    31230: "Łaźnia eventowa",
    23704: "Oranżeria",
    23698: "Sala kominkowa",
    31228: "Sala wielofunkcyjna",
    23705: "Sala bankietowa",
    22144: "Pokój nr 2",
    22146: "Pokój nr 3",
    22147: "Pokój nr 4",
    22148: "Pokój nr 5",
    22149: "Pokój nr 6",
    22150: "Pokój nr 7",
    22151: "Pokój nr 8",
    22152: "Pokój nr 9",
    22153: "Pokój nr 10",
    22154: "Pokój nr 11",
    22156: "Pokój nr 12",
    31229: "Dwór - cały obiekt"
}

# Inverse map to get IDs for Reservation push
HOTRES_ROOM_NAME_TO_ID = {v: k for k, v in HOTRES_ROOM_MAP.items()}

# --- INITIALIZE SESSION STATE FOR RELOADING OFFERS ---
if "klient_imie" not in st.session_state: st.session_state.klient_imie = ""
if "firma_n" not in st.session_state: st.session_state.firma_n = ""
if "nip_n" not in st.session_state: st.session_state.nip_n = ""
if "telefon_n" not in st.session_state: st.session_state.telefon_n = ""
if "email_n" not in st.session_state: st.session_state.email_n = ""
if "loaded_pozycje" not in st.session_state: st.session_state.loaded_pozycje = None
if "agenda_custom_text" not in st.session_state: st.session_state.agenda_custom_text = ""

# --- INTELIGENTNE ŁADOWANIE CZCIONEK Z DYSKU ---
FONT_HEADER = 'Helvetica-Bold'
FONT_TEXT = 'Helvetica'
FONT_TEXT_BOLD = 'Helvetica-Bold'

def register_custom_fonts():
    global FONT_HEADER, FONT_TEXT, FONT_TEXT_BOLD
    try:
        if os.path.exists('Lora-Bold.ttf'):
            pdfmetrics.registerFont(TTFont('Lora-Bold', 'Lora-Bold.ttf'))
            FONT_HEADER = 'Lora-Bold'
        if os.path.exists('PTSans-Regular.ttf'):
            pdfmetrics.registerFont(TTFont('PTSans-Regular', 'PTSans-Regular.ttf'))
            FONT_TEXT = 'PTSans-Regular'
        if os.path.exists('PTSans-Bold.ttf'):
            pdfmetrics.registerFont(TTFont('PTSans-Bold', 'PTSans-Bold.ttf'))
            FONT_TEXT_BOLD = 'PTSans-Bold'
    except Exception:
        pass

register_custom_fonts()

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

@st.cache_data(ttl=15, show_spinner=False)
def fetch_baza_files(baza_id):
    service = get_drive_service()
    query = f"'{baza_id}' in parents and trashed = false"
    results = service.files().list(q=query, fields="files(id, name, createdTime, webViewLink, mimeType)", orderBy="createdTime desc", supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    return results.get('files', [])

def upload_file_to_drive(file_bytes, filename, folder_id, mimetype='application/pdf'):
    service = get_drive_service()
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mimetype, resumable=True)
    file_metadata = {'name': filename, 'parents': [folder_id]}
    file = service.files().create(body=file_metadata, media_body=media, fields='id, webViewLink', supportsAllDrives=True).execute()
    try:
        service.permissions().create(fileId=file.get('id'), body={'type': 'anyone', 'role': 'reader'}, supportsAllDrives=True).execute()
    except Exception: pass
    return file

@st.cache_data(ttl=600, show_spinner=False)
def fetch_all_debogora_files(root_id):
    service = get_drive_service()
    all_files = []
    folders_to_search = [root_id]
    while folders_to_search:
        current_folder = folders_to_search.pop(0)
        query = f"'{current_folder}' in parents and trashed = false"
        try:
            request = service.files().list(q=query, fields="nextPageToken, files(id, name, mimeType, webViewLink)", pageSize=1000, supportsAllDrives=True, includeItemsFromAllDrives=True)
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
            request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
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
    service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()

# --- PANCERNA ZAMIANA TEKSTU W PPTX ---
def process_shape(shape, replacements):
    if hasattr(shape, "text_frame") and shape.text_frame is not None:
        for paragraph in shape.text_frame.paragraphs:
            if not paragraph.runs: continue
            full_text = "".join(run.text for run in paragraph.runs)
            replaced = False
            for k, v in replacements.items():
                if k in full_text:
                    full_text = full_text.replace(k, str(v))
                    replaced = True
            if replaced and paragraph.runs:
                paragraph.runs[0].text = full_text
                for i in range(1, len(paragraph.runs)):
                    paragraph.runs[i].text = ""

    if hasattr(shape, "shapes"):
        for subshape in shape.shapes: process_shape(subshape, replacements)
            
    if hasattr(shape, "has_table") and shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells: process_shape(cell, replacements)

def replace_text_in_pptx(prs, replacements):
    for slide in prs.slides:
        for shape in slide.shapes: process_shape(shape, replacements)

# --- MAPOWANIE NAZW I WYSZUKIWANIE PLIKÓW ---
def normalize_pl(text):
    rep = {'ą':'a', 'ć':'c', 'ę':'e', 'ł':'l', 'ń':'n', 'ó':'o', 'ś':'s', 'ź':'z', 'ż':'z'}
    res = str(text).lower()
    for k, v in rep.items(): res = res.replace(k, v)
    return re.sub(r'[\W_]+', '', res)

SYNONYMS = {
    "okładka": "okładka", 
    "powitalna": "karta powitalna", 
    "zakwaterowanie_dwor": "dwór 3",
    "zakwaterowanie_domki": "zakwaterowanie domki",
    "zakwaterowanie_oba": "dwór i domki",
    "wyżywienie": "wyżywienie", 
    "serwis kawowy": "serwis kawowy",
    "wiejskie jadło": "wiejskie jadło",
    "rozszerzonym menu": "kolacja z rozszerzonym menu",
    "atrakcje_wstęp": "atrakcje", 
    "wycena": "wycena", 
    "agenda": "agenda", 
    "kontakt": "AsystentAI_kontakt", 
    "uklad_debogora": "AsystentAI_układ łóżek_Dębogóra",
    "uklad_krovacja": "AsystentAI_układ łóżek_Krovacja",
    "Złodziej Krów": "złodziej", 
    "Złodziej krów": "złodziej",
    "Łowcy krów": "złodziej",
    "Skarby": "skarby", 
    "Safari_Standard": "safari dla grup_standard",
    "Safari_Rozszerzona": "safari dla grup_rozszerzona",
    "Seans saunowy": "saunowy", 
    "Sauna olchowa": "olchowa",
    "Staw": "staw", "Balia": "balia", "Sauny": "sauny", "Masaże": "masaż", "Paintball": "paintball", "Spływ kajakowy": "kajak",
    "Rowery": "rowery", "Ognisko": "ognisko", "Punkt widokowy": "widokowy", "Łączka cielaczków": "cielacz",
    "Atrakcje na wodzie": "wodzie", "Złów i wypuść": "złów", "Grzybobranie": "grzyb", "Roztańczony las": "roztańczony",
    "Drawieński PN": "drawieński", "Blok konferencyjny": "konferencyjny", "Wynajem sali": "sali", "Przejazd": "przejazd"
}

def get_file_by_keyword(keyword, all_files):
    norm_search = normalize_pl(SYNONYMS.get(keyword, keyword))
    matches = [f for f in all_files if norm_search in normalize_pl(f['name'])]
    if keyword == "atrakcje_wstęp":
        matches = [f for f in matches if "wodzi" not in normalize_pl(f['name'])]
    if matches:
        matches.sort(key=lambda x: len(x['name']))
        prev_matches = [f for f in matches if 'prev' in f['name'].lower()]
        return prev_matches[0] if prev_matches else matches[0]
    return None

def add_file_to_merger(merger, keyword, all_files, open_streams, missing_cards, added_file_ids, replacements=None):
    if not keyword: return
    file_obj = get_file_by_keyword(keyword, all_files)
    if file_obj:
        if file_obj['id'] in added_file_ids: return
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
                if res.returncode != 0: raise Exception(f"LibreOffice błąd: {res.stderr}")
                    
                with open(temp_pdf, "rb") as f: pdf_bytes = f.read()
                pdf_stream = io.BytesIO(pdf_bytes)
            else:
                pdf_stream = fh
            open_streams.append(pdf_stream)
            merger.append(PdfReader(pdf_stream, strict=False))
            added_file_ids.add(file_obj['id'])
        except Exception as e:
            st.error(f"⚠️ Pominięto '{keyword}'. Błąd pliku '{file_obj['name']}': {e}")
            missing_cards.append(keyword)
    else:
        if keyword not in missing_cards: missing_cards.append(keyword)

def safe_str(text): return "" if pd.isna(text) else str(text).strip()

def get_price_data(usluga_name, df):
    # Fallback mappings for rename safety
    search_names = [usluga_name.lower()]
    if "złodziej" in usluga_name.lower() or "łowcy" in usluga_name.lower():
        search_names.extend(["złodziej krów", "łowcy krów", "łowcy krów"])
        
    if df is None or df.empty: return {"cena": 0, "czas": 0.0, "min_start": 9.0, "max_start": 18.0}
    try:
        col_name = next((c for c in df.columns if 'nazwa' in c.lower() or 'usługa' in c.lower() or 'usluga' in c.lower()), None)
        col_price = next((c for c in df.columns if 'cena' in c.lower()), None)
        col_czas = next((c for c in df.columns if 'długość' in c.lower() or 'dlugosc' in c.lower()), None)
        col_kiedy = next((c for c in df.columns if 'kiedy' in c.lower() or 'zacząć' in c.lower() or 'zaczac' in c.lower()), None)
        
        if col_name and col_price:
            match = df[df[col_name].astype(str).str.strip().str.lower().isin(search_names)]
            if match.empty:
                match = df[df[col_name].astype(str).str.lower().str.contains(usluga_name.lower()[:5], na=False)]
                
            if not match.empty:
                val = match.iloc[0][col_price]
                cena = float(val) if pd.notna(val) else 0.0
                
                czas_num = 1.0
                if col_czas and pd.notna(match.iloc[0][col_czas]):
                    czas_str = str(match.iloc[0][col_czas]).lower().replace('h', '').replace('godz.', '').replace('godz', '').strip()
                    if czas_str:
                        if ',' in czas_str:
                            parts = czas_str.split(',')
                            if len(parts)>1 and parts[1] == '15': czas_num = float(parts[0]) + 0.25
                            elif len(parts)>1 and parts[1] == '30': czas_num = float(parts[0]) + 0.5
                            elif len(parts)>1 and parts[1] == '45': czas_num = float(parts[0]) + 0.75
                            else: czas_num = float(czas_str.replace(',', '.'))
                        else:
                            try: czas_num = float(czas_str)
                            except: czas_num = 1.0
                
                min_start = 9.0
                max_start = 18.0
                if col_kiedy and pd.notna(match.iloc[0][col_kiedy]):
                    kiedy_str = str(match.iloc[0][col_kiedy]).strip()
                    m = re.findall(r'(\d{1,2}):\d{2}', kiedy_str)
                    if len(m) >= 2:
                        min_start = float(m[0])
                        max_start = float(m[1])
                    elif len(m) == 1:
                        min_start = float(m[0])
                        max_start = 22.0
                return {"cena": cena, "czas": czas_num, "min_start": min_start, "max_start": max_start}
    except Exception: pass
    return {"cena": 0, "czas": 0.0, "min_start": 9.0, "max_start": 18.0}

def sprawdz_dostepnosc_hotres(data_od, data_do):
    try:
        api_key = st.secrets["hotres"]["api"]
        auth_key = st.secrets["hotres"]["auth"]
    except KeyError:
        return "Błąd: Zdefiniuj wpisy [hotres] api='...' oraz auth='...' w konfiguracji Secrets.", {}
    url = f"https://panel.hotres.pl/api_availability?auth={auth_key}&apikey={api_key}"
    try:
        response = requests.get(url, timeout=15)
        if response.status_code == 200:
            dane = response.json()
            szczegoly_zajetosci = {}
            liczba_nocy = max(1, (data_do - data_od).days)
            wymagane_daty = [(data_od + timedelta(days=i)).strftime("%Y-%m-%d") for i in range(liczba_nocy)]
            for room_data in dane:
                room_id = room_data.get("type_id")
                nazwa_pokoju = HOTRES_ROOM_MAP.get(int(room_id) if str(room_id).isdigit() else room_id, f"ID_{room_id}")
                for day_data in room_data.get("dates", []):
                    data_dnia = day_data.get("date")
                    dostepnosc = int(float(day_data.get("available", 0))) 
                    if data_dnia in wymagane_daty and dostepnosc <= 0:
                        if nazwa_pokoju not in szczegoly_zajetosci:
                            szczegoly_zajetosci[nazwa_pokoju] = []
                        szczegoly_zajetosci[nazwa_pokoju].append(data_dnia)
            return "", szczegoly_zajetosci
        else:
            return f"Hotres zwrócił kod: {response.status_code}", {}
    except Exception as e:
        return f"Błąd komunikacji z Hotres: {str(e)}", []

def utworz_rezerwacje_hotres(data_od, data_do, wybrane_pokoje_i_domki):
    try:
        api_key = st.secrets["hotres"]["api"]
        auth_key = st.secrets["hotres"]["auth"]
    except KeyError:
        return "Błąd kluczy autoryzacji Hotres w pliku konfiguracyjnym."
        
    url = f"https://panel.hotres.pl/api_reservations?auth={auth_key}&apikey={api_key}"
    
    pokoje_payload = []
    for rname in wybrane_pokoje_i_domki:
        tid = HOTRES_ROOM_NAME_TO_ID.get(rname)
        if tid:
            pokoje_payload.append({"type_id": str(tid), "count": "1"})

    payload = {
        "arrival": data_od.strftime("%Y-%m-%d"),
        "departure": data_do.strftime("%Y-%m-%d"),
        "rooms": pokoje_payload,
        "customer": {
            "name": st.session_state.klient_imie,
            "company": st.session_state.firma_n,
            "nip": st.session_state.nip_n,
            "phone": st.session_state.telefon_n,
            "email": st.session_state.email_n
        },
        "status": "1" # Status wstępnej rezerwacji / blokady block
    }
    
    try:
        res = requests.post(url, json=payload, timeout=15)
        if res.status_code in [200, 201]:
            return "OK"
        else:
            return f"Serwer Hotres zgłosił błąd: {res.text}"
    except Exception as e:
        return f"Nie udało się połączyć z API Hotres: {str(e)}"

def format_zajete_daty(lista_dat):
    return ", ".join([f"{d[8:10]}.{d[5:7]}" for d in sorted(lista_dat)])

# --- POŁĄCZENIE Z DYSKIEM ---
wszystkie_pliki = []
cennik_file = None

try:
    with st.spinner("Skanowanie plików na Dysku Google (Zabezpieczam czcionki)..."):
        wszystkie_pliki = fetch_all_debogora_files(ROOT_FOLDER_ID)
        czcionki_do_pobrania = ['Lora-Bold.ttf', 'PTSans-Regular.ttf', 'PTSans-Bold.ttf']
        pobrano_nowe = False
        for f in wszystkie_pliki:
            if f['name'] in czcionki_do_pobrania and not os.path.exists(f['name']):
                fh = download_file(f['id'])
                with open(f['name'], 'wb') as out: out.write(fh.getvalue())
                pobrano_nowe = True
        if pobrano_nowe: register_custom_fonts()

    cennik_files = [f for f in wszystkie_pliki if 'cennik' in f['name'].lower() and ('xlsx' in f['name'].lower() or 'csv' in f['name'].lower())]
    cennik_files.sort(key=lambda f: 'xlsx' in f['name'].lower(), reverse=True)
    if cennik_files: cennik_file = cennik_files[0]
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

# --- MENU BOCZNE Z KARTAMI ---
with st.sidebar:
    st.header("🗂️ Karty Produktów")
    with st.expander("Rozwiń pliki do pobrania"):
        if wszystkie_pliki:
            pliki_do_pobrania = [f for f in wszystkie_pliki if 'cennik' not in f['name'].lower() and f['mimeType'] != 'application/vnd.google-apps.folder' and '.ttf' not in f['name'].lower() and '.json' not in f['name'].lower()]
            for f in sorted(pliki_do_pobrania, key=lambda x: x['name']):
                link = f.get('webViewLink', '#')
                st.markdown(f"📄 [{f['name']}]({link})")
        else:
            st.info("Brak plików w bazie.")

CENNIK = {
    "nocleg_1_noc": get_price_data("Nocleg (1 noc)", df_c)["cena"],
    "nocleg_2_noce": get_price_data("Nocleg (2+ noce)", df_c)["cena"],
    "doplata_domek": 40,
    "domki": {
        "Muuu 1": {"baza": get_price_data("Muuu 1, 2", df_c)["cena"], "pdf": "krovacja"}, 
        "Muuu 2": {"baza": get_price_data("Muuu 1, 2", df_c)["cena"], "pdf": "krovacja"},
        "Muuu 3": {"baza": get_price_data("Muuu 3, 4", df_c)["cena"], "pdf": "krovacja"}, 
        "Muuu 4": {"baza": get_price_data("Muuu 3, 4", df_c)["cena"], "pdf": "krovacja"},
        "Muuu 5": {"baza": get_price_data("Muuu 5, 6", df_c)["cena"], "pdf": "krovacja"}, 
        "Muuu 6": {"baza": get_price_data("Muuu 5, 6", df_c)["cena"], "pdf": "krovacja"}
    },
    "wyzywienie": {
        "Śniadanie": {"dane": get_price_data("Śniadanie", df_c), "pdf": "wyżywienie"},
        "Obiadokolacja": {"dane": get_price_data("Obiadokolacja", df_c), "pdf": "wyżywienie"},
        "Serwis kawowy": {"dane": get_price_data("Serwis kawowy", df_c), "pdf": "serwis kawowy"},
        "Wiejskie jadło": {"dane": get_price_data("Wiejskie jadło (Podstawowe)", df_c), "pdf": "wiejskie jadło"},
        "Kolacja z rozszerzonym menu": {"dane": get_price_data("Biesiada wieczorna", df_c), "pdf": "rozszerzonym menu"}
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
        "Złodziej Krów": {"dane": get_price_data("Złodziej Krów", df_c), "typ": "osoba", "pdf": "Łowcy krów"},
        "Skarby Dębogóry": {"dane": get_price_data("Skarby Dębogóry", df_c), "typ": "osoba", "pdf": "Skarby"},
        "Krowie Safari Standard": {"dane": get_price_data("Krowie Safari Standard", df_c), "typ": "osoba", "pdf": "Safari_Standard"},
        "Krowie Safari Rozszerzone": {"dane": get_price_data("Krowie Safari Rozszerzone", df_c), "typ": "osoba", "pdf": "Safari_Rozszerzona"},
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
st.markdown("<h1 style='text-align: center; margin-top:0;'>PLANER OFERT</h1>", unsafe_allow_html=True)

if "l_osob_total" not in st.session_state: st.session_state.l_osob_total = 10
if "szczegoly_zajetosci" not in st.session_state: st.session_state.szczegoly_zajetosci = {}

def auto_alloc():
    total = st.session_state.l_osob_total
    typ = st.session_state.get("typ_klienta_radio", "Indywidualny")
    marka = st.session_state.get("marka_oferty_select", "Dwór Dębogóra")
    zajete_slownik = st.session_state.szczegoly_zajetosci
    zajete = list(zajete_slownik.keys())
    
    st.session_state.wybrane_p = []
    st.session_state.wybrane_d = []
            
    if typ == "Biznesowy":
        max_os_domki = {"Muuu 1": 2, "Muuu 2": 2, "Muuu 3": 4, "Muuu 4": 4, "Muuu 5": 1, "Muuu 6": 1}
    else:
        max_os_domki = {"Muuu 1": 4, "Muuu 2": 4, "Muuu 3": 6, "Muuu 4": 6, "Muuu 5": 3, "Muuu 6": 3}
        
    def przydziel_dworek(osoby):
        for p, cap in POKOJE_DWOREK.items():
            if osoby > 0 and p not in zajete:
                st.session_state.wybrane_p.append(p)
                val = min(cap, osoby)
                st.session_state[f"os_{p}"] = val
                osoby -= val
        return osoby
        
    def przydziel_domki(osoby):
        for d, cap in max_os_domki.items():
            if osoby > 0 and d not in zajete:
                st.session_state.wybrane_d.append(d)
                val = min(cap, osoby)
                st.session_state[f"os_{d}"] = val
                osoby -= val
        return osoby

    if marka == "Dwór Dębogóra":
        total = przydziel_dworek(total)
        total = przydziel_domki(total)
    else:
        total = przydziel_domki(total)
        total = przydziel_dworek(total)

# --- TABS: Kreator / Baza / Cennik ---
tab1, tab3, tab2 = st.tabs(["📝 Kreator Ofert", "📂 Baza Ofert", "⚙️ Edycja Cennika"])

with tab2:
    st.subheader("Edycja pliku Cennika (Google Drive)")
    if df_c又不 None:
        edited_df = st.data_editor(df_c, num_rows="dynamic", use_container_width=True)
        if st.button("💾 ZAPISZ ZMIANY NA DYSKU", type="primary"):
            update_file_on_drive(cennik_file['id'], edited_df, cennik_file['name'])
            st.session_state.df_cennik = edited_df
            st.success("Zmiany zapisane!")

with tab3:
    st.subheader("Baza i Historia Wygenerowanych Ofert")
    if st.button("🔄 Odśwież listę archiwalną"):
        fetch_baza_files.clear()
        
    try:
        baza_files = fetch_baza_files(BAZA_OFERT_FOLDER_ID)
        pdf_files = [f for f in baza_files if f['mimeType'] == 'application/pdf']
        json_files = {f['name'].replace('.json', ''): f['id'] for f in baza_files if '.json' in f['name']}
        
        if not pdf_files:
            st.info("Brak zapisanych ofert w bazie.")
        else:
            for f in pdf_files:
                base_key = f['name'].replace('.pdf', '')
                c_file, c_act = st.columns([3, 2])
                with c_file:
                    st.markdown(f"📄 **{f['name']}** - [🔗 Pobierz PDF]({f['webViewLink']})")
                with c_act:
                    if base_key in json_files:
                        if st.button("📂 Przywróć do Kreatora", key=f"load_{f['id']}"):
                            js_bytes = download_file(json_files[base_key])
                            meta = json.loads(js_bytes.getvalue().decode('utf-8'))
                            
                            st.session_state.klient_imie = meta.get("klient_imie", "")
                            st.session_state.firma_n = meta.get("firma_n", "")
                            st.session_state.nip_n = meta.get("nip_n", "")
                            st.session_state.telefon_n = meta.get("telefon_n", "")
                            st.session_state.email_n = meta.get("email_n", "")
                            st.session_state.loaded_pozycje = meta.get("pozycje", [])
                            st.session_state.agenda_custom_text = meta.get("final_agenda_text", "")
                            st.session_state.l_osob_total = meta.get("l_osob_total", 10)
                            st.success(f"Pomyślnie załadowano konfigurację dla: {meta.get('klient_imie')}. Przejdź do pierwszej zakładki!")
                            st.user_loaded = True
    except Exception as e:
        st.error(f"Nie można załadować bazy: {e}")

with tab1:
    pozycje_kosztowe = []
    wybrane_atrakcje_agenda = []

    with st.container():
        c_head1, c_head2 = st.columns([4, 1])
        with c_head1:
            st.subheader("1. Główne Ustawienia i Dane Klienta")
        with c_head2:
            if st.button("🧹 Resetuj formularz"):
                for key in list(st.session_state.keys()):
                    if key not in ['df_cennik']: del st.session_state[key]
                st.rerun()

        c1, c2 = st.columns(2)
        with c1:
            marka_oferty = st.selectbox("Marka wiodąca oferty *", ["Dwór Dębogóra", "Krovacja"], key="marka_oferty_select")
            typ_klienta = st.radio("Typ klienta", ["Indywidualny", "Biznesowy"], horizontal=True, key="typ_klienta_radio")
            
            st.session_state.klient_imie = st.text_input("Imię i nazwisko osoby kontaktowej *", value=st.session_state.klient_imie)
            st.session_state.firma_n = st.text_input("Firma (opcjonalnie)", value=st.session_state.firma_n)
            st.session_state.nip_n = st.text_input("NIP firmy (do integracji Hotres)", value=st.session_state.nip_n)
            
            st.number_input("Liczba osób", 1, 100, key="l_osob_total")
            st.button("🤖 Automatycznie rozmieść gości", on_click=auto_alloc)
        with c2:
            st.session_state.email_n = st.text_input("Email", value=st.session_state.email_n)
            st.session_state.telefon_n = st.text_input("Telefon kontaktowy (do integracji Hotres)", value=st.session_state.telefon_n)
            
            cd1, cd2 = st.columns(2)
            with cd1: d_in = st.date_input("Przyjazd", date.today())
            with cd2: d_out = st.date_input("Wyjazd", date.today()+timedelta(1))
            dni = max(1, (d_out - d_in).days)
            
            st.markdown("---")
            if st.button("🔍 Sprawdź dostępność na żywo (Hotres)"):
                with st.spinner("Łączenie z bazą rezerwacji..."):
                    err, zajete_szczegoly = sprawdz_dostepnosc_hotres(d_in, d_out)
                    if err: st.error(err)
                    else:
                        st.session_state.szczegoly_zajetosci = zajete_szczegoly
                        if zajete_szczegoly: st.warning(f"🚨 W wybranym terminie wykryto zablokowane obiekty. Szczegóły poniżej.")
                        else: st.success("✅ Wszystkie zmapowane obiekty są wolne w tym terminie!")

    with st.container():
        st.subheader("2. Zakwaterowanie")
        stawka_dw = CENNIK["nocleg_1_noc"] if dni == 1 else CENNIK["nocleg_2_noce"]
        col_dw, col_dm = st.columns(2)
        osoby_zadeklarowane = 0
        
        zajete_slownik = st.session_state.szczegoly_zajetosci
        zajete_nazwy = list(zajete_slownik.keys())
        
        if typ_klienta == "Biznesowy":
            max_os_domki = {"Muuu 1": 2, "Muuu 2": 2, "Muuu 3": 4, "Muuu 4": 4, "Muuu 5": 1, "Muuu 6": 1}
        else:
            max_os_domki = {"Muuu 1": 4, "Muuu 2": 4, "Muuu 3": 6, "Muuu 4": 6, "Muuu 5": 3, "Muuu 6": 3}
            
        with col_dw:
            dostepne_p = [p for p in POKOJE_DWOREK.keys() if p not in zajete_nazwy]
            p_sel = st.multiselect("Dworek (Dostępne pokoje)", dostepne_p, key="wybrane_p")
            
            zajete_p_do_wyswietlenia = {p: zajete_slownik[p] for p in POKOJE_DWOREK.keys() if p in zajete_slownik}
            if zajete_p_do_wyswietlenia:
                for p, daty in zajete_p_do_wyswietlenia.items():
                    st.markdown(f"<small style='color:gray;'>❌ {p} (Zajęte w: {format_zajete_daty(daty)})</small>", unsafe_allow_html=True)

            for p in p_sel:
                ile = st.number_input(f"{p} (Max: {POKOJE_DWOREK[p]})", 1, POKOJE_DWOREK[p], key=f"os_{p}")
                osoby_zadeklarowane += ile
                pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"{p}", "Ilość": ile, "Cena jednostkowa": stawka_dw*dni, "Suma": ile*stawka_dw*dni, "pdf_kw": "dworek"})
                
        with col_dm:
            dostepne_d = [d for d in CENNIK["domki"].keys() if d not in zajete_nazwy]
            d_sel = st.multiselect("Domki Krovacja (Dostępne)", dostepne_d, key="wybrane_d")
            
            zajete_d_do_wyswietlenia = {d: zajete_slownik[d] for d in CENNIK["domki"].keys() if d in zajete_slownik}
            if zajete_d_do_wyswietlenia:
                for d, daty in zajete_d_do_wyswietlenia.items():
                    st.markdown(f"<small style='color:gray;'>❌ {d} (Zajęte w: {format_zajete_daty(daty)})</small>", unsafe_allow_html=True)
                
            for d in d_sel:
                cap_domku = max_os_domki[d]
                ile = st.number_input(f"{d} (Max w tej opcji: {cap_domku})", 1, cap_domku, key=f"os_{d}")
                osoby_zadeklarowane += ile
                cena_d = (CENNIK["domki"][d]["baza"] + (max(0, ile-1)*CENNIK["doplata_domek"]))*dni
                pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"{d}", "Ilość": 1, "Cena jednostkowa": cena_d, "Suma": cena_d, "pdf_kw": "krovacja"})

        overbooking_error = False
        if osoby_zadeklarowane > st.session_state.l_osob_total:
            st.error(f"⚠️ Przydzieliłeś {osoby_zadeklarowane} miejsc dla {st.session_state.l_osob_total} gości.")
            overbooking_error = True

    with st.container():
        st.subheader("3. Wyżywienie")
        wyz_sel = st.multiselect("Wybierz opcje wyżywienia", list(CENNIK["wyzywienie"].keys()))
        for w in wyz_sel:
            w_data = CENNIK["wyzywienie"][w]["dane"]
            ile = st.number_input(f"Ilość porcji: {w}", 1, 5000, st.session_state.l_osob_total * dni)
            pozycje_kosztowe.append({"Kategoria": "Gastronomia", "Opis": w, "Ilość": ile, "Cena jednostkowa": w_data["cena"], "Suma": ile*w_data["cena"], "pdf_kw": CENNIK["wyzywienie"][w]["pdf"]})

    with st.container():
        st.subheader("4. Oferta Dodatkowa (SPAstwisko, Atrakcje, Biznes)")
        c_spa, c_atr, c_biz = st.columns(3)
        
        with c_spa:
            spa_sel = st.multiselect("SPAstwisko", list(CENNIK["SPAstwisko"].keys()))
            for a in spa_sel:
                a_data = CENNIK["SPAstwisko"][a]
                ile = st.number_input(f"Ilość: {a}", 1, 100, st.session_state.l_osob_total if a_data["typ"]=="osoba" else 1, key=f"spa_{a}")
                pozycje_kosztowe.append({"Kategoria": "SPAstwisko", "Opis": a, "Ilość": ile, "Cena jednostkowa": a_data["dane"]["cena"], "Suma": ile*a_data["dane"]["cena"], "pdf_kw": a_data["pdf"]})
                wybrane_atrakcje_agenda.append({"Nazwa": a, "Czas": a_data["dane"]["czas"], "MinStart": a_data["dane"]["min_start"], "MaxStart": a_data["dane"]["max_start"]})
                
        with c_atr:
            atr_sel = st.multiselect("Atrakcje", list(CENNIK["Atrakcje"].keys()))
            for a in atr_sel:
                a_data = CENNIK["Atrakcje"][a]
                ile = st.number_input(f"Ilość: {a}", 1, 100, st.session_state.l_osob_total if a_data["typ"]=="osoba" else 1, key=f"atr_{a}")
                pozycje_kosztowe.append({"Kategoria": "Atrakcje", "Opis": a, "Ilość": ile, "Cena jednostkowa": a_data["dane"]["cena"], "Suma": ile*a_data["dane"]["cena"], "pdf_kw": a_data["pdf"]})
                wybrane_atrakcje_agenda.append({"Nazwa": a, "Czas": a_data["dane"]["czas"], "MinStart": a_data["dane"]["min_start"], "MaxStart": a_data["dane"]["max_start"]})

        with c_biz:
            biz_sel = st.multiselect("Biznes", list(CENNIK["Biznes"].keys()))
            for a in biz_sel:
                a_data = CENNIK["Biznes"][a]
                ile = st.number_input(f"Ilość: {a}", 1, 100, st.session_state.l_osob_total if a_data["typ"]=="osoba" else 1, key=f"biz_{a}")
                pozycje_kosztowe.append({"Kategoria": "Biznes", "Opis": a, "Ilość": ile, "Cena jednostkowa": a_data["dane"]["cena"], "Suma": ile*a_data["dane"]["cena"], "pdf_kw": a_data["pdf"]})
                wybrane_atrakcje_agenda.append({"Nazwa": a, "Czas": a_data["dane"]["czas"], "MinStart": a_data["dane"]["min_start"], "MaxStart": a_data["dane"]["max_start"]})

    with st.container():
        st.subheader("5. Generator Harmonogramu (Agenda)")
        liczba_nocy = (d_out - d_in).days
        liczba_dni = liczba_nocy + 1 if liczba_nocy > 0 else 1
        
        def format_time(hours_float):
            h = int(hours_float)
            m = int(round((hours_float - h) * 60))
            return f"{h:02d}:{m:02d}"

        unassigned = wybrane_atrakcje_agenda.copy()
        unassigned.sort(key=lambda x: (x['MaxStart'] - x['MinStart'], x['MinStart']))
        
        def fill_slot(slot_start_h, slot_end_h, pending_events):
            events = []
            curr_h = slot_start_h
            while pending_events and curr_h < slot_end_h:
                best_idx, best_start, best_end = -1, -1, -1
                for i, atr in enumerate(pending_events):
                    prop_start = max(curr_h, atr['MinStart'])
                    prop_end = prop_start + (atr['Czas'] if atr['Czas']>0 else 1.0)
                    if prop_start <= atr['MaxStart'] and prop_end <= slot_end_h:
                        best_idx, best_start, best_end = i, prop_start, prop_end
                        break
                if best_idx != -1:
                    atr = pending_events.pop(best_idx)
                    if best_start > curr_h: events.append({"Nazwa": "Czas wolny", "Start": curr_h, "End": best_start})
                    events.append({"Nazwa": atr["Nazwa"], "Start": best_start, "End": best_end})
                    curr_h = best_end
                else: break 
            return events, pending_events
            
        def render_events(events):
            res = ""
            for e in events: res += f"• {format_time(e['Start'])} - {format_time(e['End'])} : {e['Nazwa']}\n"
            return res

        if st.session_state.agenda_custom_text:
            draft_agenda = st.session_state.agenda_custom_text
        else:
            draft_agenda = f"TERMIN WYDARZENIA: {d_in.strftime('%d.%m.%Y')} - {d_out.strftime('%d.%m.%Y')}\n\n"
            if liczba_dni == 1:
                draft_agenda += f"DZIEŃ 1 ({d_in.strftime('%d.%m')})\n• 10:00 - Przyjazd\n"
                evs, unassigned = fill_slot(10.5, 18.0, unassigned)
                draft_agenda += render_events(evs) + "• 18:00 - 19:00 : Obiadokolacja\n"
                evs_eve, unassigned = fill_slot(19.0, 23.0, unassigned)
                draft_agenda += render_events(evs_eve) + "• 23:00 - Zakończenie pobytu\n"
            else:
                draft_agenda += f"DZIEŃ 1 ({d_in.strftime('%d.%m')})\n• 15:00 - Przyjazd i Zakwaterowanie\n"
                evs, unassigned = fill_slot(16.0, 18.0, unassigned) 
                draft_agenda += render_events(evs) + "• 18:00 - 19:00 : Obiadokolacja\n"
                evs_eve, unassigned = fill_slot(19.0, 23.0, unassigned) 
                draft_agenda += render_events(evs_eve) + "\n"
                for d in range(2, liczba_dni):
                    draft_agenda += f"DZIEŃ {d} ({ (d_in + timedelta(days=d-1)).strftime('%d.%m') })\n• 09:00 - 10:00 : Śniadanie\n"
                    evs, unassigned = fill_slot(10.0, 18.0, unassigned) 
                    draft_agenda += render_events(evs) + "• 18:00 - 19:00 : Obiadokolacja\n"
                    evs_eve, unassigned = fill_slot(19.0, 23.0, unassigned)
                    draft_agenda += render_events(evs_eve) + "\n"
                draft_agenda += f"DZIEŃ {liczba_dni} ({d_out.strftime('%d.%m')}) (Wyjazd)\n• 09:00 - 10:00 : Śniadanie\n"
                evs, unassigned = fill_slot(10.0, 13.0, unassigned) 
                draft_agenda += render_events(evs) + "• 13:00 - Wymeldowanie\n"

        final_agenda_text = st.text_area("Szkic Harmonogramu (do edycji):", value=draft_agenda, height=350)
        st.session_state.agenda_custom_text = final_agenda_text

    with st.container():
        st.subheader("6. Kosztorys, Rezerwacja Hotres i Eksport PDF")
        
        df = pd.DataFrame(pozycje_kosztowe)
        if st.session_state.loaded_pozycje is not None:
            df = pd.DataFrame(st.session_state.loaded_pozycje)
            st.session_state.loaded_pozycje = None # clear buffer memory
            
        if not df.empty:
            df = df[["Kategoria", "Opis", "Ilość", "Cena jednostkowa", "Suma", "pdf_kw"]]
            edf = st.data_editor(df, use_container_width=True, num_rows="dynamic", column_config={
                "Suma": st.column_config.NumberColumn("Suma (Wyliczana)", disabled=True)
            })
            
            # AUTOMATIC SUM RECALCULATION
            edf["Ilość"] = pd.to_numeric(edf["Ilość"], errors='coerce').fillna(0)
            edf["Cena jednostkowa"] = pd.to_numeric(edf["Cena jednostkowa"], errors='coerce').fillna(0)
            edf["Suma"] = edf["Ilość"] * edf["Cena jednostkowa"]
            
            razem = edf["Suma"].sum()
            st.markdown(f"<h3 style='color: {CI['dark_green']};'>RAZEM DO ZAPŁATY: {razem:,.2f} PLN</h3>".replace(",", " "), unsafe_allow_html=True)
            
            c_actions1, c_actions2 = st.columns(2)
            
            with c_actions1:
                if st.button("GENERUJ FINALNY PDF Oferty", disabled=overbooking_error, type="primary"):
                    if not st.session_state.klient_imie: st.error("Podaj imię i nazwisko klienta!")
                    else:
                        with st.spinner("Kompilowanie oferty i zapisywanie archiwum..."):
                            try:
                                merger = PdfWriter()
                                open_streams, missing_cards, added_file_ids = [], [], set()
                                nazwa_docelowa = st.session_state.firma_n if st.session_state.firma_n else st.session_state.klient_imie
                                
                                has_dworek = any(row["pdf_kw"] == "dworek" for _, row in edf.iterrows())
                                has_krovacja = any(row["pdf_kw"] == "krovacja" for _, row in edf.iterrows())
                                
                                zakwaterowanie_txt = "domkach i pokojach" if (has_dworek and has_krovacja) else "domkach" if has_krovacja else "pokojach"
                                atr_list = [row["Opis"] for _, row in edf.iterrows() if row["Kategoria"] in ["SPAstwisko", "Atrakcje", "Biznes"]]
                                atrakcje_txt = f"{atr_list[0]} oraz {atr_list[1]}" if len(atr_list) >= 2 else f"{atr_list[0]} oraz naturę" if len(atr_list) == 1 else "spokój i bliskość natury"
                                marka_wstawka = "Krovację" if marka_oferty == "Krovacja" else "Dwór Dębogóra"
                                
                                replacements = {
                                    "{{nazwa firmy}}": nazwa_docelowa, "{{Dwór Dębogóra/Krovację}}": marka_wstawka, "{{domkach/pokojach}}": zakwaterowanie_txt,
                                    "{{atrakcja}} oraz {{atrakcja}}": atrakcje_txt, "{{przykład agendy}}": final_agenda_text, "{{tabela z wyceną dla b2b }}": "", "{{ tabela z wyceną dla b2b }}": ""
                                }
                                
                                add_file_to_merger(merger, "okładka", wszystkie_pliki, open_streams, missing_cards, added_file_ids, replacements)
                                add_file_to_merger(merger, "powitalna", wszystkie_pliki, open_streams, missing_cards, added_file_ids, replacements)
                                
                                if has_dworek and has_krovacja: 
                                    add_file_to_merger(merger, "zakwaterowanie_oba", wszystkie_pliki, open_streams, missing_cards, added_file_ids, replacements)
                                    add_file_to_merger(merger, "uklad_debogora", wszystkie_pliki, open_streams, missing_cards, added_file_ids, replacements)
                                    add_file_to_merger(merger, "uklad_krovacja", wszystkie_pliki, open_streams, missing_cards, added_file_ids, replacements)
                                elif has_dworek: 
                                    add_file_to_merger(merger, "zakwaterowanie_dwor", wszystkie_pliki, open_streams, missing_cards, added_file_ids, replacements)
                                    add_file_to_merger(merger, "uklad_debogora", wszystkie_pliki, open_streams, missing_cards, added_file_ids, replacements)
                                elif has_krovacja: 
                                    add_file_to_merger(merger, "zakwaterowanie_domki", wszystkie_pliki, open_streams, missing_cards, added_file_ids, replacements)
                                    add_file_to_merger(merger, "uklad_krovacja", wszystkie_pliki, open_streams, missing_cards, added_file_ids, replacements)

                                if any(row["Kategoria"] == "Gastronomia" for _, row in edf.iterrows()): 
                                    if any(row["Opis"] in ["Śniadanie", "Obiadokolacja"] for _, row in edf.iterrows()):
                                        add_file_to_merger(merger, "wyżywienie", wszystkie_pliki, open_streams, missing_cards, added_file_ids, replacements)
                                    gastronomia_kws = list(set([row["pdf_kw"] for _, row in edf.iterrows() if row["Kategoria"] == "Gastronomia" and pd.notna(row["pdf_kw"]) and row["pdf_kw"] != "wyżywienie"]))
                                    for g_kw in gastronomia_kws: add_file_to_merger(merger, g_kw, wszystkie_pliki, open_streams, missing_cards, added_file_ids, replacements)

                                add_file_to_merger(merger, "atrakcje_wstęp", wszystkie_pliki, open_streams, missing_cards, added_file_ids, replacements)
                                for kw in list(set([row["pdf_kw"] for _, row in edf.iterrows() if row["Kategoria"] in ["SPAstwisko", "Atrakcje", "Biznes"] and pd.notna(row["pdf_kw"])])):
                                    add_file_to_merger(merger, kw, wszystkie_pliki, open_streams, missing_cards, added_file_ids, replacements)

                                buf = io.BytesIO()
                                doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=120, bottomMargin=50)
                                elements = []

                                t_data = [["Kategoria", "Opis usługi", "Ilość", "Suma"]]
                                for _, row in edf.iterrows():
                                    t_data.append([safe_str(row["Kategoria"]), safe_str(row["Opis"]), safe_str(row["Ilość"]), f"{row['Suma']:,.0f} zł".replace(",", " ")])
                                t_data.append(["", "", "RAZEM:", f"{razem:,.0f} zł".replace(",", " ")])
                                
                                table = Table(t_data, colWidths=[110, 220, 50, 90])
                                table.setStyle(TableStyle([
                                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(CI['dark_green'])), ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                                    ('FONTNAME', (0, 0), (-1, 0), FONT_HEADER), ('FONTSIZE', (0, 0), (-1, 0), 12),
                                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('ALIGN', (1, 1), (1, -2), 'LEFT'),
                                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('BOTTOMPADDING', (0, 0), (-1, -1), 10), ('TOPPADDING', (0, 0), (-1, -1), 10),
                                    ('FONTNAME', (0, 1), (-1, -1), FONT_TEXT), ('FONTNAME', (2, -1), (-1, -1), FONT_TEXT_BOLD),
                                    ('TEXTCOLOR', (2, -1), (-1, -1), colors.HexColor(CI['dark_green'])), ('BACKGROUND', (2, -1), (-1, -1), colors.HexColor(CI['light_green'])),
                                    ('LINEABOVE', (0, -1), (-1, -1), 1, colors.HexColor(CI['dark_green'])),
                                ]))
                                elements.append(table)
                                doc.build(elements)
                                buf.seek(0)
                                
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
                                        fg_reader = PdfReader(buf)
                                        for i, fg_page in enumerate(fg_reader.pages):
                                            bg_stream_fresh = io.BytesIO(bg_bytes)
                                            open_streams.append(bg_stream_fresh)
                                            bg_reader = PdfReader(bg_stream_fresh)
                                            bg_page = bg_reader.pages[min(i, len(bg_reader.pages) - 1)]
                                            bg_page.merge_page(fg_page)
                                            merger.add_page(bg_page)
                                    except Exception: merger.append(PdfReader(buf, strict=False))
                                else: merger.append(PdfReader(buf, strict=False))

                                add_file_to_merger(merger, "agenda", wszystkie_pliki, open_streams, missing_cards, added_file_ids, replacements)
                                add_file_to_merger(merger, "kontakt", wszystkie_pliki, open_streams, missing_cards, added_file_ids, replacements)

                                final_pdf = io.BytesIO()
                                merger.write(final_pdf)
                                pdf_bytes_to_upload = final_pdf.getvalue()
                                timestamp = datetime.now().strftime('%d%m%H%M')
                                nazwa_pliku_pdf = f"Oferta_{safe_str(st.session_state.klient_imie).replace(' ', '_')}_{timestamp}.pdf"
                                nazwa_pliku_json = f"Oferta_{safe_str(st.session_state.klient_imie).replace(' ', '_')}_{timestamp}.json"
                                
                                st.success("✅ Oferta w formacie PDF została pomyślnie wygenerowana!")
                                st.download_button("📥 POBIERZ PDF NA DYSK LOKALNY", pdf_bytes_to_upload, nazwa_pliku_pdf, "application/pdf", type="primary")

                                # SAVE CONFIGURATION METADATA TO CLOUD
                                meta_payload = {
                                    "klient_imie": st.session_state.klient_imie, "firma_n": st.session_state.firma_n, "nip_n": st.session_state.nip_n,
                                    "telefon_n": st.session_state.telefon_n, "email_n": st.session_state.email_n, "marka_oferty": marka_oferty,
                                    "typ_klienta": typ_klienta, "l_osob_total": st.session_state.l_osob_total, "final_agenda_text": final_agenda_text,
                                    "pozycje": edf.to_dict(orient="records")
                                }
                                json_bytes = json.dumps(meta_payload, ensure_ascii=False, indent=4).encode('utf-8')
                                
                                upload_file_to_drive(pdf_bytes_to_upload, nazwa_pliku_pdf, BAZA_OFERT_FOLDER_ID, 'application/pdf')
                                upload_file_to_drive(json_bytes, nazwa_pliku_json, BAZA_OFERT_FOLDER_ID, 'application/json')
                                st.info("✅ Kopia oraz parametry konfiguracyjne zostały zapisane w chmurze.")
                                fetch_baza_files.clear()
                            except Exception as global_error:
                                st.error(f"❌ Błąd generatora: {str(global_error)}")
                            finally:
                                for f in glob.glob("temp_*"):
                                    try: os.remove(f)
                                    except: pass
            
            with c_actions2:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("⚡ PRZEŚLIJ REZERWACJĘ DO HOTRES", type="secondary", use_container_width=True):
                    wybrane_obiekty = list(st.session_state.get("wybrane_p", [])) + list(st.session_state.get("wybrane_d", []))
                    if not wybrane_obiekty:
                        st.error("Wybierz przynajmniej jeden pokój lub domek przed wysłaniem rezerwacji!")
                    elif not st.session_state.klient_imie or not st.session_state.email_n:
                        st.error("Dane klienta (Imię i nazwisko oraz Email) są wymagane do utworzenia rezerwacji!")
                    else:
                        with st.spinner("Wysyłanie rezerwacji i danych kontrahenta do systemu Hotres..."):
                            status_res = utworz_rezerwacje_hotres(d_in, d_out, wybrane_obiekty)
                            if status_res == "OK":
                                st.success("🎉 Sukces! Wszystkie pokoje/domki oraz dane klienta (w tym Telefon i NIP) zostały wprowadzone do Hotres.")
                            else:
                                st.error(status_res)
