import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta

# --- KONFIGURACJA STRONY ---
st.set_page_config(page_title="Generator Ofert - Dwór Dębogóra", layout="wide")

# Barwy z Twojego CI
CI = {
    "dark_green": "#00622f",
    "light_green": "#e8ece6",
    "gray": "#333333",
    "white": "#ffffff"
}

# --- ZAAWANSOWANA STYLIZACJA CSS ---
st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Lora:ital,wght@0,400;0,700;1,400&family=PT+Sans:wght@400;700&display=swap');

    /* Ogólne tło i fonty */
    .stApp {{
        background-color: {CI['white']};
        font-family: 'PT Sans', sans-serif;
        color: {CI['gray']};
    }}

    /* Nagłówki Lora */
    h1, h2, h3, h4 {{
        font-family: 'Lora', serif !important;
        color: {CI['dark_green']} !important;
        font-weight: 700 !important;
        margin-bottom: 1rem !important;
    }}

    /* Kontenery sekcji - jasna zieleń */
    div[data-testid="stVerticalBlock"] > div > div[data-testid="stVerticalBlock"] {{
        background-color: {CI['light_green']};
        padding: 2.5rem;
        border-radius: 0px;
        border-left: 5px solid {CI['dark_green']};
        margin-bottom: 2rem;
    }}

    /* Stylowanie przycisków */
    div.stButton > button {{
        background-color: {CI['dark_green']} !important;
        color: white !important;
        border-radius: 0px !important;
        border: none !important;
        font-family: 'Lora', serif !important;
        padding: 0.8rem 3rem !important;
        font-weight: bold !important;
        text-transform: uppercase;
        letter-spacing: 2px;
        transition: 0.3s;
    }}

    div.stButton > button:hover {{
        background-color: {CI['gray']} !important;
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }}

    /* Estetyka pól formularza */
    div[data-baseweb="input"] > div, div[data-baseweb="select"] > div, div[data-baseweb="datepicker"] > div {{ 
        border: 1px solid #ced4cd !important; 
        border-radius: 0px !important;
        background-color: white !important;
    }}
    </style>
""", unsafe_allow_html=True)

# --- WYŚRODKOWANE I MNIEJSZE LOGO ---
col_l1, col_l2, col_l3 = st.columns([1, 1, 1])
with col_l2:
    try:
        # Szerokość 150px (2x mniejsza niż poprzednio)
        st.image("logo.png", width=150)
    except:
        st.write(f"<p style='text-align:center; color:{CI['dark_green']}'>[ Tu pojawi się logo.png ]</p>", unsafe_allow_html=True)

st.markdown(f"<h1 style='text-align: center; margin-top:-20px;'>System Ofertowania</h1>", unsafe_allow_html=True)
st.markdown(f"<p style='text-align: center; color: {CI['gray']}; font-style: italic;'>Dwór Dębogóra & Domki Krovacja</p>", unsafe_allow_html=True)
st.write("---")

# --- BAZA DANYCH CENNIKA ---
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
        "Brak wyżywienia": 0,
        "Śniadanie": 50,
        "Śniadanie + Obiadokolacja": 120
    },
    "atrakcje": {
        "Sauna Olchowa (do 12 osób)": {"cena": 400, "typ": "grupa"},
        "Balia opalana drewnem (do 6 osób)": {"cena": 300, "typ": "grupa"},
        "Sauny Wynajem na wyłączność": {"cena": 450, "typ": "grupa"},
        "Łowcy krów (gra krótsza)": {"cena": 100, "typ": "osoba"},
        "Skarby Dębogóry (gra dłuższa)": {"cena": 200, "typ": "osoba"},
        "Krowie Safari Standard": {"cena": 100, "typ": "osoba"},
        "Paintball (min 10 os)": {"cena": 150, "typ": "osoba"},
        "Kajaki (min 4 os)": {"cena": 140, "typ": "osoba"},
        "Ognisko z drewnem": {"cena": 150, "typ": "grupa"},
        "Rower Elektryczny": {"cena": 120, "typ": "osoba"},
        "Wycieczka DPN (3h)": {"cena": 700, "typ": "grupa"}
    }
}

pozycje_kosztowe = []

# --- 1. PEŁNE DANE KLIENTA ---
with st.container():
    st.subheader("1. Dane Klienta i Termin")
    c1, c2 = st.columns(2)
    with c1:
        klient_imie = st.text_input("Imię i nazwisko (obligatoryjne) *")
        firma_n = st.text_input("Nazwa Firmy")
        nip_n = st.text_input("NIP")
        l_osob_total = st.number_input("Łączna liczba osób na imprezie", min_value=1, value=10)
    with c2:
        mail_n = st.text_input("E-mail")
        tel_n = st.text_input("Telefon")
        cd1, cd2 = st.columns(2)
        with cd1:
            d_in = st.date_input("Przyjazd", value=date.today())
        with cd2:
            d_out = st.date_input("Wyjazd", value=date.today() + timedelta(days=1))
        
        l_dni = (d_out - d_in).days
        if l_dni < 1:
            st.error("Data wyjazdu musi być późniejsza.")
            l_dni = 1
        else:
            st.info(f"Pobyt: **{l_dni} dób**")

# --- 2. ZAKWATEROWANIE ---
with st.container():
    st.subheader("2. Konfiguracja Noclegów")
    stawka_dworek = CENNIK["nocleg_1_noc"] if l_dni == 1 else CENNIK["nocleg_2_noce"]
    
    col_dw, col_dm = st.columns(2)
    
    with col_dw:
        st.markdown("**Dworek (Pokoje)**")
        p1 = st.number_input("Pokoje 1-os", 0)
        p2 = st.number_input("Pokoje 2-os", 0)
        p3 = st.number_input("Pokoje 3-os", 0)
        os_dw = (p1*1) + (p2*2) + (p3*3)
        st.write(f"Miejsca: {os_dw}")
        
        if os_dw > 0:
            if p1>0: pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"Dworek: Pokój 1-os (x{p1})", "Ilość": p1, "Cena": 1*stawka_dworek*l_dni, "Suma": p1*1*stawka_dworek*l_dni})
            if p2>0: pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"Dworek: Pokój 2-os (x{p2})", "Ilość": p2, "Cena": 2*stawka_dworek*l_dni, "Suma": p2*2*stawka_dworek*l_dni})
            if p3>0: pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"Dworek: Pokój 3-os (x{p3})", "Ilość": p3, "Cena": 3*stawka_dworek*l_dni, "Suma": p3*3*stawka_dworek*l_dni})

    with col_dm:
        st.markdown("**Domki Krovacja**")
        wybor_d = st.multiselect("Wybierz domki", list(CENNIK["domki"].keys()))
        os_dm = 0
        for d_name in wybor_d:
            par = CENNIK["domki"][d_name]
            ile_os = st.number_input(f"Osób w {d_name} (max {par['max_os']})", 1, par['max_os'])
            os_dm += ile_os
            cena_baza = par["baza"]
            doplata = (ile_os - 1) * CENNIK["doplata_domek"] if ile_os > 1 else 0
            suma_d = (cena_baza + doplata) * l_dni
            pozycje_kosztowe.append({"Kategoria": "Nocleg", "Opis": f"Domek {d_name} ({ile_os} os.)", "Ilość": 1, "Cena": suma_d, "Suma": suma_d})

    if (os_dw + os_dm) != l_osob_total:
        st.warning(f"⚠️ Uwaga: Rozlokowano {os_dw + os_dm} osób, a grupa liczy {l_osob_total}!")

# --- 3. WYŻYWIENIE ---
with st.container():
    st.subheader("3. Wyżywienie")
    wyz_opt = st.selectbox("Wariant posiłków", list(CENNIK["wyzywienie"].keys()))
    if wyz_opt != "Brak wyżywienia":
        c_jedn = CENNIK["wyzywienie"][wyz_opt]
        pozycje_kosztowe.append({"Kategoria": "Gastronomia", "Opis": wyz_opt, "Ilość": l_osob_total * l_dni, "Cena": c_jedn, "Suma": c_jedn * l_osob_total * l_dni})

# --- 4. ATRAKCJE ---
with st.container():
    st.subheader("4. Atrakcje dodatkowe")
    atr_sel = st.multiselect("Wybierz atrakcje", list(CENNIK["atrakcje"].keys()))
    for a in atr_sel:
        a_data = CENNIK["atrakcje"][a]
        if a_data["typ"] == "osoba":
            ile_a = st.number_input(f"Liczba osób na: {a}", 1, l_osob_total, l_osob_total)
            pozycje_kosztowe.append({"Kategoria": "Atrakcje", "Opis": a, "Ilość": ile_a, "Cena": a_data["cena"], "Suma": ile_a * a_data["cena"]})
        else:
            ile_g = st.number_input(f"Ilość (grupy/wynajmy) na: {a}", 1, 10, 1)
            pozycje_kosztowe.append({"Kategoria": "Atrakcje", "Opis": a, "Ilość": ile_g, "Cena": a_data["cena"], "Suma": ile_g * a_data["cena"]})

# --- 5. KOSZTORYS I GENEROWANIE ---
with st.container():
    st.subheader("5. Podsumowanie i Edycja Taryf")
    df = pd.DataFrame(pozycje_kosztowe)
    if not df.empty:
        edytowany_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        razem = edytowany_df["Suma"].sum()
        st.markdown(f"### RAZEM DO ZAPŁATY: {razem:,.2f} PLN".replace(",", " "))
        
        if st.button("GENERUJ FINALNĄ OFERTĘ PDF"):
            if not klient_imie:
                st.error("Błąd: Imię i nazwisko klienta jest wymagane!")
            else:
                st.balloons()
                st.success(f"Oferta dla {klient_imie} gotowa do pobrania (logika PDF w przygotowaniu).")
