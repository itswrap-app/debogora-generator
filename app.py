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
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# --- KONFIGURACJA ---
st.set_page_config(page_title="Generator Ofert - Dwór Dębogóra", layout="wide")

CI = {
    "dark_green": "#00622f",
    "light_green": "#e8ece6",
    "gray": "#333333",
    "white": "#ffffff"
}

FOLDER_ID = "1i_a2UkK73ixyvMBe5l9SkE5vpqAu6he5"

# --- STYLIZACJA CSS ---
st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Lora:ital,wght@0,400;0,700;1,400&family=PT+Sans:wght@400;700&display=swap');
    .stApp {{ background-color: {CI['white']}; font-family: 'PT Sans', sans-serif; color: {CI['gray']}; }}
    h1, h2, h3, h4 {{ font-family: 'Lora', serif !important; color: {CI['dark_green']} !important; font-weight: 700 !important; }}
    div[data-testid="stVerticalBlock"] > div > div[data-testid="stVerticalBlock"] {{
        background-color: {CI['light_green']}; padding: 2.5rem; border-radius: 0px; border-left: 5px solid {CI['dark_green']}; margin-bottom: 1.5rem;
    }}
    div.stButton > button {{
        background-color: {CI['dark_green']} !important; color: white !important;
        border-radius: 0px !important; border: none !important; font-family: 'Lora', serif !important; padding: 0.8rem 3rem !important;
        font-weight: bold !important; text-transform: uppercase; letter-spacing: 2px;
    }}
    div.stButton > button:hover {{ background-color: {CI['gray']} !important; }}
    </style>
""", unsafe_allow_html=True)

# --- POŁĄCZENIE Z GOOGLE DRIVE ---
@st.cache_resource
def get_drive_service():
    info = st.secrets["gcp_service_account"]
    creds = SACredentials.from_service_account_info(info)
    return build('drive', 'v3', credentials=creds)

def list_drive_files(folder_id):
    service = get_drive_service()
    query = f"'{folder_id}' in parents and trashed = false"
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
        "Muuu 1": {"baza": 700, "max_os": 4}, "Muuu 2": {"baza": 700, "max_os": 4},
        "Muuu 3": {"baza": 1050, "max_os": 6}, "Muuu 4": {"baza": 1050, "max_os": 6},
        "Muuu 5": {"baza": 700, "max_os": 3}, "Muuu 6": {"baza": 700, "max_os": 3}
    },
    "wyzywienie": {"Brak wyżywienia": 0, "Śniadanie": 50, "Śniadanie + Obiadokolacja": 120}
}

POKOJE_DWOREK = {
    "Pokój nr 1": 1, "Pokój nr 2": 2, "Pokój nr 3": 2, "Pokój nr 4": 2,
    "Pokój nr 5": 2, "Pokój nr 6": 2, "Pokój nr 7": 3, "Pokój nr 8": 2,
    "Pokój nr 9": 3, "Pokój nr 10": 3, "Pokój nr 11": 4, "Pokój nr 12": 3
}

# --- LOGO ---
try:
    logo_b64 = base64.b64encode(open("logo.png", "rb").read()).decode()
    st.markdown(f'<div style="display: flex; justify-content: center; margin-bottom: 20px;"><img src="data:image/png;base64,{logo_b64}" width="120"></div>', unsafe_allow_html=True)
except: pass

st.markdown("<h1 style='text-align: center;'>System Ofertowania</h1>", unsafe_allow_html=True)
st.markdown(f"<p style='text-align: center; color: {CI['gray']}; font-style: italic;'>Dwór Dębogóra & Domki Krovacja</p>", unsafe_allow_html=True)

# --- 1. PEŁNE DANE KLIENTA ---
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
        klient_imie = st.text_input("Imię i nazwisko (wymagane) *")
        firma_n = st.text_input("Nazwa Firmy")
        nip_n = st.text_input("NIP")
        st.number_input("Łączna liczba osób", 1, 100, key="l_osob_total")
        st.button("🤖 Automatycznie rozmieść gości", on_click=auto_alloc)
    with c2:
        email_n = st.text_input("Adres e-mail")
        tel_n = st.text_input("Telefon")
        cd1, cd2 = st.columns(2)
        with cd1: d_in = st.date_input("Przyjazd", date.today())
        with cd2: d_out = st.date_input("Wyjazd", date.today() + timedelta(days=1))
        l_dni = max(1, (d_out - d_in).days)
        st.info(f"Wyliczony pobyt: **{l_dni} dób**")

# --- 2. ZAKWATEROWANIE ---
with st.container():
    st.subheader("2. Konfiguracja Noclegów")
    stawka_dw = CENNIK["nocleg_1_noc"] if l_dni == 1 else CENNIK["nocleg_2_noce"]
    col_dw, col_dm = st.columns(2)
    sum_os = 0
    
    with col_dw:
        st.markdown("**Dworek (Pokoje)**")
        wybrane_p = st.multiselect("Wybierz pokoje", list(POKOJE_DWOREK.keys()), key="wybrane_p")
        for p in wybrane_p:
            cap = POKOJE_DWOREK[p]
            ile = st.number_input(f"{p} (max {cap} os.)", 1, cap, key=f"os_{p}")
            sum_os += ile
            pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"{p} (os: {ile})", "Ilość": ile, "Cena": stawka_dw*l_dni, "Suma": ile*stawka_dw*l_dni})

    with col_dm:
        st.markdown("**Domki Krovacja**")
        wybrane_d = st.multiselect("Wybierz domki", list(CENNIK["domki"].keys()), key="wybrane_d")
        for d in wybrane_d:
            cap = CENNIK["domki"][d]["max_os"]
            ile = st.number_input(f"{d} (max {cap} os.)", 1, cap, key=f"os_{d}")
            sum_os += ile
            cena_d = (CENNIK["domki"][d]["baza"] + (max(0, ile-1)*CENNIK["doplata_domek"]))*l_dni
            pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"{d} ({ile} os.)", "Ilość": 1, "Cena": cena_d, "Suma": cena_d})

    if sum_os != st.session_state.l_osob_total:
        st.warning(f"⚠️ Rozmieszczono {sum_os} z {st.session_state.l_osob_total} osób.")

# --- 3. DODATKI I ATRAKCJE (Z DRIVE) ---
with st.container():
    st.subheader("3. Wyżywienie i Atrakcje")
    c_w1, c_w2 = st.columns(2)
    with c_w1:
        wyz_opt = st.selectbox("Wariant wyżywienia", list(CENNIK["wyzywienie"].keys()))
        if wyz_opt != "Brak wyżywienia":
            c_j = CENNIK["wyzywienie"][wyz_opt]
            pozycje_kosztowe.append({"Kategoria": "Gastronomia", "Opis": wyz_opt, "Ilość": st.session_state.l_osob_total*l_dni, "Cena": c_j, "Suma": c_j*st.session_state.l_osob_total*l_dni})
    
    with c_w2:
        try:
            pliki_drive = list_drive_files(FOLDER_ID)
            karty_pdf = [f for f in pliki_drive if f['name'].lower().endswith('.pdf')]
            wybrane_karty = st.multiselect("Załącz karty produktów (PDF z Dysku)", [f['name'] for f in karty_pdf])
        except:
            st.error("Błąd połączenia z Dyskiem Google. Sprawdź Secrets.")
            wybrane_karty = []

# --- 4. PODSUMOWANIE EDYTOWALNE ---
with st.container():
    st.subheader("4. Kosztorys Ofertowy")
    df = pd.DataFrame(pozycje_kosztowe)
    if not df.empty:
        edf = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        razem = edf["Suma"].sum()
        st.markdown(f"### RAZEM: {razem:,.2f} PLN".replace(",", " "))
        
        if st.button("GENERUJ FINALNĄ OFERTĘ PDF"):
            if not klient_imie:
                st.error("Podaj imię i nazwisko klienta!")
            else:
                with st.spinner("Składanie oferty..."):
                    merger = PdfWriter()
                    # 1. Podmiana w PPTX na PDF (Okładka)
                    try:
                        okl_id = next(f['id'] for f in pliki_drive if 'okładka' in f['name'].lower())
                        ppt_fh = download_file(okl_id)
                        prs = Presentation(ppt_fh)
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if hasattr(shape, "text") and "{{nazwa firmy}}" in shape.text:
                                    shape.text = shape.text.replace("{{nazwa firmy}}", firma_n if firma_n else klient_imie)
                        prs.save("final_okladka.pptx")
                        subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "final_okladka.pptx"])
                        merger.append("final_okladka.pdf")
                    except Exception as e: st.error(f"Błąd okładki: {e}")

                    # 2. Dodawanie wybranych kart PDF
                    for k_name in wybrane_karty:
                        k_id = next(f['id'] for f in karty_pdf if f['name'] == k_name)
                        merger.append(download_file(k_id))
                    
                    # 3. Tabela kosztów
                    buf = io.BytesIO()
                    doc = SimpleDocTemplate(buf, pagesize=A4)
                    elements = [Paragraph(f"Specyfikacja finansowa dla: {klient_imie}", getSampleStyleSheet()['Heading1']), Spacer(1, 20)]
                    t_data = [["Kategoria", "Opis", "Ilość", "Suma"]] + edf[["Kategoria", "Opis", "Ilość", "Suma"]].values.tolist()
                    t_data.append(["", "", "SUMA:", f"{razem} zł"])
                    table = Table(t_data, colWidths=[100, 250, 50, 80])
                    table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor(CI['dark_green'])),('TEXTCOLOR',(0,0),(-1,0),colors.white),('GRID',(0,0),(-1,-2),0.5,colors.grey)]))
                    elements.append(table)
                    doc.build(elements)
                    buf.seek(0)
                    merger.append(buf)

                    final_out = io.BytesIO()
                    merger.write(final_out)
                    st.download_button("📥 POBIERZ OFERTĘ PDF", final_out.getvalue(), f"Oferta_{klient_imie}.pdf", "application/pdf")
