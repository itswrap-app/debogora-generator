import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import os
import subprocess
import base64

# Biblioteki Google
from google.oauth2.service_account import Credentials as SACredentials
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# Edycja plików
from pptx import Presentation
from pypdf import PdfWriter, PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# --- KONFIGURACJA ---
st.set_page_config(page_title="Generator Ofert - Dwór Dębogóra", layout="wide")

CI = {"dark_green": "#00622f", "light_green": "#e8ece6", "gray": "#333333", "white": "#ffffff"}
FOLDER_ID = "1i_a2UkK73ixyvMBe5l9SkE5vpqAu6he5"

st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Lora:ital,wght@0,400;0,700;1,400&family=PT+Sans:wght@400;700&display=swap');
    .stApp {{ background-color: {CI['white']}; font-family: 'PT Sans', sans-serif; }}
    h1, h2, h3 {{ font-family: 'Lora', serif !important; color: {CI['dark_green']} !important; }}
    div[data-testid="stVerticalBlock"] > div > div[data-testid="stVerticalBlock"] {{
        background-color: {CI['light_green']}; padding: 2rem; border-left: 5px solid {CI['dark_green']};
    }}
    div.stButton > button {{
        background-color: {CI['dark_green']} !important; color: white !important;
        border-radius: 0px !important; font-family: 'Lora', serif !important; padding: 0.8rem 2rem !important;
    }}
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

# --- LOGIKA POKOJÓW I CENNIKA ---
CENNIK = {
    "nocleg_1_noc": 220, "nocleg_2_noce": 170, "doplata_domek": 40,
    "domki": {
        "Muuu 1": {"baza": 700, "max_os": 4}, "Muuu 2": {"baza": 700, "max_os": 4},
        "Muuu 3": {"baza": 1050, "max_os": 6}, "Muuu 4": {"baza": 1050, "max_os": 6},
        "Muuu 5": {"baza": 700, "max_os": 3}, "Muuu 6": {"baza": 700, "max_os": 3}
    },
    "wyzywienie": {"Brak": 0, "Śniadanie": 50, "Śniadanie + Obiadokolacja": 120}
}
POKOJE_DWOREK = {f"Pokój nr {i}": (1 if i==1 else 4 if i==11 else 3 if i in [7,9,10,12] else 2) for i in range(1,13)}

# --- INTERFEJS ---
try:
    logo_data = base64.b64encode(open("logo.png", "rb").read()).decode()
    st.markdown(f'<div style="text-align: center;"><img src="data:image/png;base64,{logo_data}" width="120"></div>', unsafe_allow_html=True)
except: pass

st.markdown("<h1 style='text-align: center;'>System Ofertowania</h1>", unsafe_allow_html=True)

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

pozycje = []
with st.container():
    st.subheader("1. Dane Klienta")
    c1, c2 = st.columns(2)
    with c1:
        klient = st.text_input("Imię i Nazwisko / Firma *")
        st.number_input("Łączna liczba osób", 1, 100, key="l_osob_total")
        st.button("🤖 Autouzupełnianie", on_click=auto_alloc)
    with c2:
        cd1, cd2 = st.columns(2)
        with cd1: d_in = st.date_input("Przyjazd", date.today())
        with cd2: d_out = st.date_input("Wyjazd", date.today()+timedelta(1))
        dni = max(1, (d_out - d_in).days)

with st.container():
    st.subheader("2. Noclegi")
    stawka = CENNIK["nocleg_1_noc"] if dni == 1 else CENNIK["nocleg_2_noce"]
    col_dw, col_dm = st.columns(2)
    sum_os = 0
    with col_dw:
        p_sel = st.multiselect("Dworek", list(POKOJE_DWOREK.keys()), key="wybrane_p")
        for p in p_sel:
            ile = st.number_input(f"{p}", 1, POKOJE_DWOREK[p], key=f"os_{p}")
            sum_os += ile
            pozycje.append({"Kategoria": "Nocleg", "Opis": f"{p} (os: {ile})", "Ilość": ile, "Cena": stawka*dni, "Suma": ile*stawka*dni})
    with col_dm:
        d_sel = st.multiselect("Domki", list(CENNIK["domki"].keys()), key="wybrane_d")
        for d in d_sel:
            ile = st.number_input(f"{d}", 1, CENNIK["domki"][d]["max_os"], key=f"os_{d}")
            sum_os += ile
            cena_d = (CENNIK["domki"][d]["baza"] + (max(0, ile-1)*CENNIK["doplata_domek"]))*dni
            pozycje.append({"Kategoria": "Nocleg", "Opis": f"{d} ({ile} os.)", "Ilość": 1, "Cena": cena_d, "Suma": cena_d})

if sum_os != st.session_state.l_osob_total: st.warning(f"Różnica osób: {sum_os}/{st.session_state.l_osob_total}")

with st.container():
    st.subheader("3. Dodatki z Dysku Google")
    pliki_drive = list_drive_files(FOLDER_ID)
    karty_pdf = [f for f in pliki_drive if f['name'].endswith('.pdf')]
    wybrane_karty = st.multiselect("Wybierz karty produktów do załączenia", [f['name'] for f in karty_pdf])
    
    wyz = st.selectbox("Wyżywienie", list(CENNIK["wyzywienie"].keys()))
    if wyz != "Brak":
        c_w = CENNIK["wyzywienie"][wyz]
        pozycje.append({"Kategoria": "Gastronomia", "Opis": wyz, "Ilość": st.session_state.l_osob_total*dni, "Cena": c_w, "Suma": c_w*st.session_state.l_osob_total*dni})

# --- GENEROWANIE ---
with st.container():
    st.subheader("4. Kosztorys")
    df = pd.DataFrame(pozycje)
    if not df.empty:
        edf = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        razem = edf["Suma"].sum()
        if st.button("GENERUJ OFERTĘ PDF"):
            with st.spinner("Pobieranie i scalanie..."):
                merger = PdfWriter()
                
                # 1. Okładka PPTX
                okl_id = next(f['id'] for f in pliki_drive if 'okładka_02' in f['name'])
                ppt_stream = download_file(okl_id)
                prs = Presentation(ppt_stream)
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if hasattr(shape, "text") and "{{nazwa firmy}}" in shape.text:
                            shape.text = shape.text.replace("{{nazwa firmy}}", klient)
                prs.save("tmp.pptx")
                subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", "tmp.pptx"])
                merger.append("tmp.pdf")
                
                # 2. Karty produktów
                for k_name in wybrane_karty:
                    k_id = next(f['id'] for f in karty_pdf if f['name'] == k_name)
                    merger.append(download_file(k_id))
                
                # 3. Tabela kosztów (ReportLab)
                buf = io.BytesIO()
                doc = SimpleDocTemplate(buf, pagesize=A4)
                styles = getSampleStyleSheet()
                t_data = [["Kategoria", "Opis", "Ilość", "Suma"]] + edf[["Kategoria", "Opis", "Ilość", "Suma"]].values.tolist()
                t_data.append(["", "", "RAZEM", f"{razem} zł"])
                tab = Table(t_data, colWidths=[100, 250, 50, 100])
                tab.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),CI['dark_green']),('TEXTCOLOR',(0,0),(-1,0),colors.white)]))
                doc.build([Paragraph(f"Oferta dla: {klient}", styles['Heading1']), tab])
                buf.seek(0)
                merger.append(buf)
                
                final_buf = io.BytesIO()
                merger.write(final_buf)
                st.download_button("📥 POBIERZ GOTOWĄ OFERTĘ", final_buf.getvalue(), f"Oferta_{klient}.pdf", "application/pdf")
