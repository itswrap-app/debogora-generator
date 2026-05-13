import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import base64

# --- KONFIGURACJA ---
st.set_page_config(page_title="Generator Ofert - Dwór Dębogóra", layout="wide")

# Barwy z Twojego CI
CI = {
    "dark_green": "#00622f",
    "light_green": "#e8ece6",
    "gray": "#333333",
    "white": "#ffffff"
}

# --- STYLIZACJA CSS (To zmienia wygląd na "ładny") ---
st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Lora:ital,wght@0,400;0,700;1,400&family=PT+Sans:wght@400;700&display=swap');

    /* Tło całej aplikacji */
    .stApp {{
        background-color: {CI['white']};
        font-family: 'PT Sans', sans-serif;
    }}

    /* Stylowanie nagłówków */
    h1, h2, h3 {{
        font-family: 'Lora', serif !important;
        color: {CI['dark_green']} !important;
        font-weight: 700 !important;
    }}

    /* Kontener dla sekcji (szare tło pod pola) */
    .st-emotion-cache-12w0qpk {{
        background-color: {CI['light_green']};
        padding: 2rem;
        border-radius: 10px;
        border: 1px solid #d1d9cf;
    }}

    /* Przyciski */
    div.stButton > button {{
        background-color: {CI['dark_green']} !important;
        color: white !important;
        border-radius: 0px !important;
        border: none !important;
        font-family: 'Lora', serif !important;
        padding: 0.6rem 2rem !important;
        font-weight: bold !important;
        text-transform: uppercase;
        letter-spacing: 1px;
    }}

    div.stButton > button:hover {{
        background-color: {CI['gray']} !important;
    }}

    /* Tabela edytowalna */
    .stDataFrame {{
        border: 1px solid {CI['light_green']};
    }}

    /* Customowe odstępy */
    .block-container {{
        padding-top: 2rem;
    }}
    </style>
""", unsafe_allow_html=True)

# --- WYŚWIETLANIE LOGO ---
# Zakładamy, że plik logo nazywa się 'logo.png' i jest w tym samym folderze na GitHub
try:
    st.image("logo.png", width=300)
except:
    st.warning("Wrzuć plik 'logo.png' do głównego folderu na GitHubie, aby wyświetlić logo.")

st.markdown(f"<h1 style='text-align: center;'>System Ofertowania</h1>", unsafe_allow_html=True)
st.markdown(f"<p style='text-align: center; color: {CI['gray']};'>Dwór Dębogóra & Domki Krovacja</p>", unsafe_allow_html=True)
st.write("---")

# --- LOGIKA BIZNESOWA (CENNIK) ---
CENNIK = {
    "nocleg_1_noc": 220,
    "nocleg_2_noce": 170,
    "domki": {
        "Muuu 1": {"baza": 700, "max_os": 4},
        "Muuu 2": {"baza": 700, "max_os": 4},
        "Muuu 3": {"baza": 1050, "max_os": 6},
        "Muuu 4": {"baza": 1050, "max_os": 6},
        "Muuu 5": {"baza": 700, "max_os": 4},
        "Muuu 6": {"baza": 700, "max_os": 4},
    },
    "doplata_domek": 40,
    "wyzywienie": {
        "Brak": 0, "Śniadanie": 50, "Śniadanie + Obiadokolacja": 120
    }
}

pozycje_kosztowe = []

# --- PANEL WEJŚCIOWY ---
with st.container():
    st.subheader("📅 Szczegóły Rezerwacji")
    col_k1, col_k2, col_k3 = st.columns([2,1,1])
    with col_k1:
        klient = st.text_input("Imię i Nazwisko / Firma *")
    with col_k2:
        d_in = st.date_input("Przyjazd", value=date.today())
    with col_k3:
        d_out = st.date_input("Wyjazd", value=date.today() + timedelta(days=1))
    
    l_dni = (d_out - d_in).days
    l_osob = st.number_input("Łączna liczba osób", min_value=1, value=1)

# --- ZAKWATEROWANIE ---
st.write("##")
col_dw, col_kr = st.columns(2)

with col_dw:
    st.subheader("🏰 Dworek")
    stawka = CENNIK["nocleg_1_noc"] if l_dni == 1 else CENNIK["nocleg_2_noce"]
    p1 = st.number_input("Pokoje 1-os", 0)
    p2 = st.number_input("Pokoje 2-os", 0)
    p3 = st.number_input("Pokoje 3-os", 0)
    
    os_dw = (p1*1) + (p2*2) + (p3*3)
    if os_dw > 0:
        pozycje_kosztowe.append({"Opis": f"Nocleg Dworek (osób: {os_dw})", "Ilość": l_dni, "Cena": os_dw * stawka, "Suma": os_dw * stawka * l_dni})

with col_kr:
    st.subheader("🐄 Domki Krovacja")
    domki_wybor = st.multiselect("Wybierz domki", list(CENNIK["domki"].keys()))
    os_kr = 0
    for d in domki_wybor:
        max_d = CENNIK["domki"][d]["max_os"]
        ile_os_d = st.number_input(f"Osób w {d} (max {max_d})", 1, max_d)
        os_kr += ile_os_d
        cena_d = CENNIK["domki"][d]["baza"] + (max(0, ile_os_d - 1) * CENNIK["doplata_domek"])
        pozycje_kosztowe.append({"Opis": f"Domek {d} ({ile_os_d} os.)", "Ilość": l_dni, "Cena": cena_d, "Suma": cena_d * l_dni})

# Alert walidacyjny
if (os_dw + os_kr) != l_osob:
    st.error(f"Niezgodność osób! Rozlokowano {os_dw + os_kr} z {l_osob} deklarowanych.")

# --- WYŻYWIENIE ---
st.write("##")
with st.container():
    st.subheader("🍴 Gastronomia")
    wyz = st.selectbox("Wariant", list(CENNIK["wyzywienie"].keys()))
    if wyz != "Brak":
        c_wyz = CENNIK["wyzywienie"][wyz]
        pozycje_kosztowe.append({"Opis": f"Wyżywienie: {wyz}", "Ilość": l_osob * l_dni, "Cena": c_wyz, "Suma": c_wyz * l_osob * l_dni})

# --- PODSUMOWANIE EDYTOWALNE ---
st.write("##")
st.subheader("💰 Kosztorys Ofertowy")
df = pd.DataFrame(pozycje_kosztowe)
if not df.empty:
    edytowalne_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
    suma = edytowalne_df["Suma"].sum()
    st.markdown(f"### Suma całkowita: {suma:,.2f} PLN".replace(",", " "))
    
    if st.button("GENERUJ OFERTĘ PDF"):
        st.balloons()
        st.success("Generowanie PDF w toku...")
else:
    st.info("Dodaj elementy zakwaterowania lub wyżywienia, aby zobaczyć tabelę.")
