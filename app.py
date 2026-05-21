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
FONT_HEADER = 'Helvetica-Bold'
FONT_TEXT = 'Helvetica'
FONT_TEXT_BOLD = 'Helvetica-Bold'
fonts_loaded = False
font_error = ""

lora_path = 'Lora-Bold.ttf'
text_font_path = None
text_font_bold_path = None

# Priorytet dla Lato (niezawodne polskie znaki w ReportLab), potem PT Sans
for f in ['Lato-Regular.ttf', 'PTSans-Regular.ttf']:
    if os.path.exists(f):
        text_font_path = f
        break

for f in ['Lato-Bold.ttf', 'PTSans-Bold.ttf']:
    if os.path.exists(f):
        text_font_bold_path = f
        break

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
    except Exception as e:
        font_error = str(e)
else:
    font_error = "Brak plików TTF na serwerze."

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

def add_pdf_to_merger(merger, keyword, all_files):
    """Funkcja szukająca pliku PDF na podstawie słowa kluczowego i doklejająca go do oferty."""
    matched_files = [f for f in all_files if keyword.lower() in f['name'].lower() and 'pdf' in f['mimeType'].lower()]
    if matched_files:
        prev_files = [f for f in matched_files if 'prev' in f['name'].lower()]
        selected_pdf = prev_files[0] if prev_files else matched_files[0]
        try:
            merger.append(download_file(selected_pdf['id']))
        except Exception:
            pass

def safe_str(text):
    if pd.isna(text):
        return ""
    return str(text).strip()

# --- PANEL BOCZNY ---
with st.sidebar:
    st.subheader("🛠 Diagnostyka systemu")
    st.markdown("**Status czcionek:**")
    if fonts_loaded:
        st.success(f"✅ Czcionki załadowane: {text_font_path}")
    else:
        st.error(f"❌ Aktywna czcionka zastępcza. Powód: {font_error}")
        
    try:
        wszystkie_pliki = fetch_all_debogora_files(ROOT_FOLDER_ID)
        st.success(f"✅ Połączono z Drive. Liczba plików: {len(wszystkie_pliki)}")
    except Exception as e:
        st.error(f"❌ Błąd połączenia Drive: {e}")
        wszystkie_pliki = []

# --- LOGIKA CENNIKA Z NOWYMI KATEGORIAMI ---
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
    "wyzywienie": {
        "Śniadanie": {"cena": 50, "pdf": "wyżywienie"},
        "Obiadokolacja": {"cena": 70, "pdf": "wyżywienie"},
        "Śniadanie + Obiadokolacja": {"cena": 120, "pdf": "wyżywienie"},
        "Serwis kawowy": {"cena": 0, "pdf": "wyżywienie"},
        "Wiejskie jadło": {"cena": 0, "pdf": "wyżywienie"},
        "Kolacja z rozszerzonym menu": {"cena": 0, "pdf": "wyżywienie"}
    },
    "SPAstwisko": {
        "Seans full experience": {"cena": 0, "typ": "grupa", "pdf": "Seans"},
        "Sauna olchowa": {"cena": 400, "typ": "grupa", "pdf": "Sauna"},
        "Staw kąpielowy": {"cena": 0, "typ": "grupa", "pdf": "Staw"},
        "Balia opalana drewnem": {"cena": 300, "typ": "grupa", "pdf": "Balia"},
        "Sauny": {"cena": 0, "typ": "grupa", "pdf": "Sauny"},
        "Masaż relaksacyjny": {"cena": 0, "typ": "osoba", "pdf": "Masaż"},
        "Masaż gorącą świecą": {"cena": 0, "typ": "osoba", "pdf": "Masaż"},
        "Masaż gorącymi kamieniami": {"cena": 0, "typ": "osoba", "pdf": "Masaż"},
        "Masaż klasyczny częściowy": {"cena": 0, "typ": "osoba", "pdf": "Masaż"},
        "Masaż twarzy i dekoltu": {"cena": 0, "typ": "osoba", "pdf": "Masaż"}
    },
    "Atrakcje": {
        "Łowcy krów": {"cena": 100, "typ": "osoba", "pdf": "Łowcy"},
        "Skarby Dębogóry": {"cena": 200, "typ": "osoba", "pdf": "Skarby"},
        "Krowie Safari Standard": {"cena": 100, "typ": "osoba", "pdf": "Safari"},
        "Krowie Safari Rozszerzone": {"cena": 150, "typ": "osoba", "pdf": "Safari"},
        "Paintball": {"cena": 150, "typ": "osoba", "pdf": "Paintball"},
        "Kajaki": {"cena": 140, "typ": "osoba", "pdf": "Kajaki"},
        "Rowery elektryczne krótka przejażdżka (2-3h)": {"cena": 0, "typ": "osoba", "pdf": "elektryczne"},
        "Rowery elektryczne 1 dzień": {"cena": 0, "typ": "osoba", "pdf": "elektryczne"},
        "Rowery elektryczne 2 dni": {"cena": 0, "typ": "osoba", "pdf": "elektryczne"},
        "Rowery elektryczne 3 dni": {"cena": 0, "typ": "osoba", "pdf": "elektryczne"},
        "Rowery MTB krótka przejażdżka (2-3h)": {"cena": 0, "typ": "osoba", "pdf": "MTB"},
        "Rowery MTB 1 dzień": {"cena": 0, "typ": "osoba", "pdf": "MTB"},
        "Rowery MTB 2 dni": {"cena": 0, "typ": "osoba", "pdf": "MTB"},
        "Rowery MTB 3 dni": {"cena": 0, "typ": "osoba", "pdf": "MTB"},
        "Ognisko": {"cena": 150, "typ": "grupa", "pdf": "Ognisko"},
        "Punkt widokowy": {"cena": 0, "typ": "grupa", "pdf": "widokowy"},
        "Łączka cielaczków": {"cena": 0, "typ": "grupa", "pdf": "cielaczków"},
        "Atrakcje na wodzie": {"cena": 0, "typ": "grupa", "pdf": "wodzie"},
        "Złów i wypuść": {"cena": 0, "typ": "grupa", "pdf": "Złów"},
        "Grzybobranie": {"cena": 0, "typ": "grupa", "pdf": "Grzybobranie"},
        "Roztańczony las": {"cena": 0, "typ": "grupa", "pdf": "Roztańczony"},
        "Drawieński PN (3h)": {"cena": 0, "typ": "grupa", "pdf": "Drawieński"},
        "Drawieński PN (6h)": {"cena": 0, "typ": "grupa", "pdf": "Drawieński"}
    },
    "Biznes": {
        "Blok konferencyjny": {"cena": 0, "typ": "grupa", "pdf": "Blok"},
        "Wynajem sali": {"cena": 0, "typ": "grupa", "pdf": "Wynajem"},
        "Przejazd grupy (do 23 os.)": {"cena": 0, "typ": "grupa", "pdf": "Przejazd"},
        "Przejazd grupy (do 50 os.)": {"cena": 0, "typ": "grupa", "pdf": "Przejazd"}
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
            pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"{d} ({ile} os.)", "Ilość": 1, "Cena": cena_d, "Suma": cena_d, "pdf_kw": "Krovacja"})

with st.container():
    st.subheader("3. Wyżywienie")
    wyz_sel = st.multiselect("Wybierz opcje wyżywienia", list(CENNIK["wyzywienie"].keys()))
    for w in wyz_sel:
        w_data = CENNIK["wyzywienie"][w]
        ile = st.number_input(f"Ilość porcji/osób: {w}", 1, 1000, st.session_state.l_osob_total)
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

# --- GENEROWANIE FINALNEGO PLIKU ---
with st.container():
    st.subheader("5. Kosztorys i Eksport")
    st.info("Ceny z wartością '0' możesz ręcznie wycenić bezpośrednio w poniższej tabeli przed wygenerowaniem PDF.")
    df = pd.DataFrame(pozycje_kosztowe)
    if not df.empty:
        edf = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        razem = edf["Suma"].sum()
        st.markdown(f"<h3 style='color: {CI['dark_green']};'>RAZEM DO ZAPŁATY: {razem:,.2f} PLN</h3>".replace(",", " "), unsafe_allow_html=True)
        
        if st.button("GENERUJ FINALNY PDF"):
            if not klient_imie:
                st.error("Podaj imię i nazwisko klienta!")
            else:
                with st.spinner("Składanie 8-etapowej oferty..."):
                    merger = PdfWriter()
                    
                    # ETAP 1: OKŁADKA (PPTX)
                    okladka_file = next((f for f in wszystkie_pliki if 'okładka' in f['name'].lower() and 'pdf' not in f['mimeType'].lower()), None)
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

                    # ETAP 2: Karta powitalna
                    add_pdf_to_merger(merger, "powitalna", wszystkie_pliki)

                    # ETAP 3: Zakwaterowanie (osobno Dworek i Krovacja)
                    has_dworek = any(row["pdf_kw"] == "Dworek" for _, row in edf.iterrows())
                    has_krovacja = any(row["pdf_kw"] == "Krovacja" for _, row in edf.iterrows())
                    if has_dworek: add_pdf_to_merger(merger, "dworek", wszystkie_pliki)
                    if has_krovacja: add_pdf_to_merger(merger, "krovacja", wszystkie_pliki)

                    # ETAP 4: Wyżywienie
                    has_wyzywienie = any(row["Kategoria"] == "Gastronomia" for _, row in edf.iterrows())
                    if has_wyzywienie: add_pdf_to_merger(merger, "wyżywienie", wszystkie_pliki)

                    # ETAP 5: Atrakcje (wstęp)
                    add_pdf_to_merger(merger, "atrakcje_wstęp", wszystkie_pliki)

                    # ETAP 6: Karty wybranych atrakcji (SPAstwisko, Atrakcje, Biznes)
                    atrakcje_kws = list(set([row["pdf_kw"] for _, row in edf.iterrows() if row["Kategoria"] in ["SPAstwisko", "Atrakcje", "Biznes"] and pd.notna(row["pdf_kw"])]))
                    for kw in atrakcje_kws:
                        add_pdf_to_merger(merger, kw, wszystkie_pliki)

                    # ETAP 7: Wycena (Tabela z systemu)
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
                        st.error(f"Problem w ReportLab: {e}")

                    # ETAP 8: Przykładowa agenda
                    add_pdf_to_merger(merger, "agenda", wszystkie_pliki)

                    # ETAP 9: Kontakt
                    add_pdf_to_merger(merger, "kontakt", wszystkie_pliki)

                    # ZAPIS FINALNY
                    final_pdf = io.BytesIO()
                    merger.write(final_pdf)
                    
                    safe_filename = "".join([c for c in safe_str(klient_imie) if c.isalpha() or c.isdigit() or c==' ']).rstrip().replace(" ", "_")
                    st.download_button("📥 POBIERZ SCALONĄ OFERTĘ PDF", final_pdf.getvalue(), f"Oferta_{safe_filename}.pdf", "application/pdf")
