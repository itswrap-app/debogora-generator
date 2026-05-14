import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import base64
import os
import io

from pptx import Presentation
from pypdf import PdfWriter, PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# --- KONFIGURACJA STRONY ---
st.set_page_config(page_title="Generator Ofert - Dwór Dębogóra", layout="wide")

CI = {"dark_green": "#00622f", "light_green": "#e8ece6", "gray": "#333333", "white": "#ffffff"}

st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Lora:ital,wght@0,400;0,700;1,400&family=PT+Sans:wght@400;700&display=swap');
    .stApp {{ background-color: {CI['white']}; font-family: 'PT Sans', sans-serif; color: {CI['gray']}; }}
    h1, h2, h3, h4 {{ font-family: 'Lora', serif !important; color: {CI['dark_green']} !important; font-weight: 700 !important; }}
    div[data-testid="stVerticalBlock"] > div > div[data-testid="stVerticalBlock"] {{ background-color: {CI['light_green']}; padding: 2rem; border-radius: 0px; border-left: 5px solid {CI['dark_green']}; margin-bottom: 1.5rem; }}
    div.stButton > button {{ background-color: {CI['dark_green']} !important; color: white !important; border-radius: 0px !important; border: none !important; font-family: 'Lora', serif !important; padding: 0.8rem 3rem !important; font-weight: bold !important; text-transform: uppercase; letter-spacing: 2px; }}
    div.stButton > button:hover {{ background-color: {CI['gray']} !important; }}
    </style>
""", unsafe_allow_html=True)

st.markdown(f"<h1 style='text-align: center;'>System Ofertowania</h1>", unsafe_allow_html=True)
st.markdown(f"<p style='text-align: center; color: {CI['gray']}; font-style: italic; margin-bottom: 30px;'>Dwór Dębogóra & Domki Krovacja</p>", unsafe_allow_html=True)

# --- BAZA CENNIKA I POKOJÓW ---
CENNIK = {
    "nocleg_1_noc": 220,
    "nocleg_2_noce": 170,
    "domki": {
        "Muuu 1": {"baza": 700, "max_os": 4},
        "Muuu 2": {"baza": 700, "max_os": 4},
        "Muuu 3": {"baza": 1050, "max_os": 6},
        "Muuu 4": {"baza": 1050, "max_os": 6},
        "Muuu 5": {"baza": 700, "max_os": 3}, # Zaktualizowano na 2+1 (max 3)
        "Muuu 6": {"baza": 700, "max_os": 3}, # Zaktualizowano na 2+1 (max 3)
    },
    "doplata_domek": 40,
    "wyzywienie": {"Brak wyżywienia": 0, "Śniadanie": 50, "Śniadanie + Obiadokolacja": 120},
    "atrakcje": {
        "Sauna Olchowa (do 12 osób)": {"cena": 400, "typ": "grupa"},
        "Balia opalana drewnem (do 6 osób)": {"cena": 300, "typ": "grupa"},
        "Skarby Dębogóry (gra dłuższa)": {"cena": 200, "typ": "osoba"},
        "Kajaki (min 4 os)": {"cena": 140, "typ": "osoba"},
        "Ognisko z drewnem": {"cena": 150, "typ": "grupa"},
    } # Skrócona lista dla czytelności testów
}

POKOJE_DWOREK = {
    "Pokój nr 1": 1, "Pokój nr 2": 2, "Pokój nr 3": 2, "Pokój nr 4": 2,
    "Pokój nr 5": 2, "Pokój nr 6": 2, "Pokój nr 7": 3, "Pokój nr 8": 2,
    "Pokój nr 9": 3, "Pokój nr 10": 3, "Pokój nr 11": 4, "Pokój nr 12": 3
}

# Inicjalizacja Session State
if "l_osob_total" not in st.session_state:
    st.session_state.l_osob_total = 10
    st.session_state.wybrane_p_keys = []
    st.session_state.wybrane_d_keys = []

def auto_rozmiesc():
    total = st.session_state.l_osob_total
    
    st.session_state.wybrane_p_keys = []
    for p in POKOJE_DWOREK.keys(): st.session_state[f"os_{p}"] = 1 # Zabezpieczenie minimalne
    
    st.session_state.wybrane_d_keys = []
    for d in CENNIK["domki"].keys(): st.session_state[f"os_{d}"] = 1
    
    # 1. Alokacja Dworek
    for p, cap in POKOJE_DWOREK.items():
        if total > 0:
            st.session_state.wybrane_p_keys.append(p)
            to_assign = min(cap, total)
            st.session_state[f"os_{p}"] = to_assign
            total -= to_assign
            
    # 2. Alokacja Domki (jeśli Dworek pełen)
    for d, param in CENNIK["domki"].items():
        if total > 0:
            st.session_state.wybrane_d_keys.append(d)
            to_assign = min(param["max_os"], total)
            st.session_state[f"os_{d}"] = to_assign
            total -= to_assign

pozycje_kosztowe = []

# --- 1. DANE KLIENTA ---
with st.container():
    st.subheader("1. Dane Klienta i Termin")
    c1, c2 = st.columns(2)
    with c1:
        klient_imie = st.text_input("Imię i nazwisko (wymagane) *")
        firma_n = st.text_input("Nazwa Firmy")
        st.number_input("Łączna liczba osób", min_value=1, value=10, key="l_osob_total")
        st.button("🤖 Automatycznie rozmieść gości", on_click=auto_rozmiesc, type="secondary")
    with c2:
        cd1, cd2 = st.columns(2)
        with cd1: d_in = st.date_input("Przyjazd", value=date.today())
        with cd2: d_out = st.date_input("Wyjazd", value=date.today() + timedelta(days=1))
        l_dni = max(1, (d_out - d_in).days)
        st.info(f"Pobyt: **{l_dni} dób**")

# --- 2. ZAKWATEROWANIE ---
with st.container():
    st.subheader("2. Konfiguracja Noclegów")
    stawka_dworek = CENNIK["nocleg_1_noc"] if l_dni == 1 else CENNIK["nocleg_2_noce"]
    col_dw, col_dm = st.columns(2)
    
    suma_osob_przypisanych = 0

    with col_dw:
        st.markdown("**Dworek (Konkretne pokoje)**")
        wybrane_p = st.multiselect("Wybierz pokoje (Dworek)", list(POKOJE_DWOREK.keys()), key="wybrane_p_keys")
        
        for p in wybrane_p:
            cap = POKOJE_DWOREK[p]
            ile = st.number_input(f"{p} (max {cap} os.)", min_value=1, max_value=cap, key=f"os_{p}")
            suma_osob_przypisanych += ile
            cena_pokoju = ile * stawka_dworek * l_dni
            pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"{p} (os: {ile})", "Ilość": ile, "Cena": stawka_dworek * l_dni, "Suma": cena_pokoju})

    with col_dm:
        st.markdown("**Domki Krovacja**")
        wybrane_d = st.multiselect("Wybierz domki", list(CENNIK["domki"].keys()), key="wybrane_d_keys")
        
        for d in wybrane_d:
            cap = CENNIK["domki"][d]["max_os"]
            ile = st.number_input(f"{d} (max {cap} os.)", min_value=1, max_value=cap, key=f"os_{d}")
            suma_osob_przypisanych += ile
            
            cena_baza = CENNIK["domki"][d]["baza"]
            doplata = (ile - 1) * CENNIK["doplata_domek"] if ile > 1 else 0
            suma_d = (cena_baza + doplata) * l_dni
            opis = f"{d} ({ile} os.)"
            if doplata > 0: opis += f" [+ dopłata x{ile-1}]"
            
            pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": opis, "Ilość": 1, "Cena": suma_d, "Suma": suma_d})

    # ALERTY WALIDACYJNE W OBIE STRONY
    if suma_osob_przypisanych < st.session_state.l_osob_total:
        st.error(f"⚠️ Za mało miejsc! Przypisano {suma_osob_przypisanych} os., a zadeklarowano {st.session_state.l_osob_total}. Wybierz dodatkowe pokoje.")
    elif suma_osob_przypisanych > st.session_state.l_osob_total:
        st.warning(f"⚠️ Uwaga! Do pokoi przypisano {suma_osob_przypisanych} osób, podczas gdy zadeklarowano przyjazd na {st.session_state.l_osob_total} osób.")
    else:
        st.success(f"✅ Przypisano idealnie: {suma_osob_przypisanych} z {st.session_state.l_osob_total} miejsc zakwaterowania.")

# --- 3. WYŻYWIENIE & ATRAKCJE ---
with st.container():
    st.subheader("3. Gastronomia i Atrakcje")
    wyz_opt = st.selectbox("Wariant posiłków", list(CENNIK["wyzywienie"].keys()))
    if wyz_opt != "Brak wyżywienia":
        c_jedn = CENNIK["wyzywienie"][wyz_opt]
        pozycje_kosztowe.append({"Kategoria": "Gastronomia", "Opis": wyz_opt, "Ilość": st.session_state.l_osob_total * l_dni, "Cena": c_jedn, "Suma": c_jedn * st.session_state.l_osob_total * l_dni})

    atr_sel = st.multiselect("Wybierz atrakcje", list(CENNIK["atrakcje"].keys()))
    for a in atr_sel:
        a_data = CENNIK["atrakcje"][a]
        if a_data["typ"] == "osoba":
            ile_a = st.number_input(f"Liczba osób na: {a}", 1, st.session_state.l_osob_total, st.session_state.l_osob_total)
            pozycje_kosztowe.append({"Kategoria": "Atrakcje", "Opis": a, "Ilość": ile_a, "Cena": a_data["cena"], "Suma": ile_a * a_data["cena"]})
        else:
            ile_g = st.number_input(f"Ilość (grupy/wynajmy) na: {a}", 1, 10, 1)
            pozycje_kosztowe.append({"Kategoria": "Atrakcje", "Opis": a, "Ilość": ile_g, "Cena": a_data["cena"], "Suma": ile_g * a_data["cena"]})

# --- 4. PODSUMOWANIE ---
with st.container():
    st.subheader("4. Kosztorys Ofertowy")
    df = pd.DataFrame(pozycje_kosztowe)
    if not df.empty:
        edytowany_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        razem = edytowany_df["Suma"].sum()
        st.markdown(f"### RAZEM DO ZAPŁATY: {razem:,.2f} PLN".replace(",", " "))
        
        st.info("System jest przygotowany na integrację z Dyskiem Google. Moduł Drive API zostanie wpięty pod poniższy przycisk.")
        if st.button("GENERUJ FINALNĄ OFERTĘ PDF"):
            # Tutaj podepniemy Google Drive API, np.:
            # 1. drive_service = get_drive_service()
            # 2. download_file_from_drive('1i_a2UkK73ixyvMBe5l9SkE5vpqAu6he5', 'AsystentAI_okładka_02.pptx')
            # 3. download karty atrakcji etc.
            st.success("W tym miejscu odpali się pobieranie z Dysku Google i łączenie plików PDF!")
