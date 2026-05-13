import streamlit as st
import pandas as pd
from datetime import datetime

# --- KONFIGURACJA STRONY I CI ---
st.set_page_config(page_title="Generator Ofert - Dwór Dębogóra", layout="wide")

CI_COLORS = {
    "white": "#ffffff",
    "light_green": "#e8ece6",
    "dark_gray": "#333333",
    "dark_green": "#00622f"
}

st.markdown(f"""
    <style>
    .stApp {{ font-family: 'Lora', 'PT Sans Pro', sans-serif; }}
    .main-header {{ color: {CI_COLORS['dark_green']}; font-weight: bold; border-bottom: 2px solid {CI_COLORS['light_green']}; padding-bottom: 10px; margin-bottom: 20px; }}
    .section-header {{ color: {CI_COLORS['dark_gray']}; font-weight: bold; margin-top: 30px; }}
    div[data-baseweb="input"] > div {{ border: 1px solid {CI_COLORS['light_green']}; }}
    </style>
""", unsafe_allow_html=True)

# --- BAZA CENNIKA (Z CSV) ---
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
        "Śniadanie + Obiadokolacja": 120 # 50 + 70 z CSV
    },
    "atrakcje": {
        "Sauna Olchowa (do 12 osób)": {"cena": 400, "typ": "grupa"},
        "Balia opalana drewnem (do 6 osób)": {"cena": 300, "typ": "grupa"},
        "Sauny Wynajem na wyłączność": {"cena": 450, "typ": "grupa"},
        "Łowcy krów (gra krótsza)": {"cena": 100, "typ": "osoba"},
        "Skarby Dębogóry (gra dłuższa)": {"cena": 200, "typ": "osoba"},
        "Krowie Safari Standard": {"cena": 100, "typ": "osoba"},
        "Krowie Safari Rozszerzone": {"cena": 150, "typ": "osoba"},
        "Paintball (min 10 os)": {"cena": 150, "typ": "osoba"},
        "Kajaki (min 4 os)": {"cena": 140, "typ": "osoba"},
        "Ognisko z drewnem": {"cena": 150, "typ": "grupa"},
        "Rower Elektryczny": {"cena": 120, "typ": "osoba"}, # traktujemy ilość sztuk jako osoby
        "Rower MTB": {"cena": 60, "typ": "osoba"},
        "Wycieczka DPN (3h)": {"cena": 700, "typ": "grupa"}
    }
}

st.markdown('<h1 class="main-header">Generator Ofert - Dwór Dębogóra</h1>', unsafe_allow_html=True)

# --- 1. DANE KLIENTA ---
st.markdown('<h3 class="section-header">1. Dane Klienta i Wydarzenia</h3>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    klient_imie_nazwisko = st.text_input("Imię i nazwisko klienta (wymagane) *")
    firma = st.text_input("Firma (opcjonalnie)")
    nip = st.text_input("NIP (opcjonalnie)")
with col2:
    email = st.text_input("Adres e-mail (opcjonalnie)")
    telefon = st.text_input("Telefon (opcjonalnie)")
    l_osob_total = st.number_input("Łączna liczba uczestników wydarzenia", min_value=1, value=10, step=1)
    l_dni = st.number_input("Liczba dób noclegowych", min_value=1, value=1, step=1)

# --- 2. ZAKWATEROWANIE ---
st.markdown('<h3 class="section-header">2. Zakwaterowanie (Dworek & Domki)</h3>', unsafe_allow_html=True)
stawka_dworek = CENNIK["nocleg_1_noc"] if l_dni == 1 else CENNIK["nocleg_2_noce"]

col_dworek, col_domki = st.columns(2)

with col_dworek:
    st.subheader("Dworek / Budynek")
    pokoje_1 = st.number_input("Liczba pokoi 1-osobowych", min_value=0, value=0)
    pokoje_2 = st.number_input("Liczba pokoi 2-osobowych", min_value=0, value=0)
    pokoje_3 = st.number_input("Liczba pokoi 3-osobowych", min_value=0, value=0)
    osoby_dworek = (pokoje_1 * 1) + (pokoje_2 * 2) + (pokoje_3 * 3)
    st.write(f"Zadeklarowane miejsca w dworku: **{osoby_dworek}**")
    koszt_dworek = osoby_dworek * stawka_dworek * l_dni

with col_domki:
    st.subheader("Domki Krovacja")
    wybrane_domki = st.multiselect("Wybierz domki do rezerwacji", options=list(CENNIK["domki"].keys()))
    
    osoby_domki = 0
    koszt_domki = 0
    
    for domek in wybrane_domki:
        parametry = CENNIK["domki"][domek]
        l_osob_w_domku = st.number_input(f"Liczba osób w domku {domek} (max {parametry['max_os']})", min_value=1, max_value=parametry['max_os'], value=1)
        osoby_domki += l_osob_w_domku
        koszt_domek_pojedynczy = parametry["baza"]
        if l_osob_w_domku > 1:
            koszt_domek_pojedynczy += (l_osob_w_domku - 1) * CENNIK["doplata_domek"]
        koszt_domki += (koszt_domek_pojedynczy * l_dni)

# WALIDACJA LICZBY OSÓB
l_osob_przypisanych = osoby_dworek + osoby_domki
if l_osob_przypisanych != l_osob_total:
    st.error(f"⚠️ BŁĄD: Przypisano {l_osob_przypisanych} miejsc noclegowych, a deklarowana liczba uczestników to {l_osob_total}.")

# --- 3. WYŻYWIENIE ---
st.markdown('<h3 class="section-header">3. Wyżywienie</h3>', unsafe_allow_html=True)
opcja_wyzywienia = st.selectbox("Wybierz wariant wyżywienia", options=list(CENNIK["wyzywienia"].keys() if "wyzywienia" in CENNIK else CENNIK["wyzywienie"].keys()))
koszt_wyzywienie = CENNIK["wyzywienie"][opcja_wyzywienia] * l_osob_total * l_dni

# --- 4. ATRAKCJE DODATKOWE ---
st.markdown('<h3 class="section-header">4. Atrakcje Dodatkowe</h3>', unsafe_allow_html=True)
wybrane_atrakcje = st.multiselect("Wybierz atrakcje z listy", options=list(CENNIK["atrakcje"].keys()))

koszt_atrakcje = 0
szczegoly_atrakcji = []

for atrakcja in wybrane_atrakcje:
    dane_atrakcji = CENNIK["atrakcje"][atrakcja]
    if dane_atrakcji["typ"] == "osoba":
        l_chetnych = st.number_input(f"Ilu chętnych na: {atrakcja}?", min_value=1, max_value=l_osob_total, value=l_osob_total)
        cena_za_atrakcje = l_chetnych * dane_atrakcji["cena"]
        szczegoly_atrakcji.append({"Nazwa": atrakcja, "Ilość": l_chetnych, "Cena": cena_za_atrakcje})
    else:
        # Atrakcja grupowo ryczałtowa (np. wycieczka, wynajem sauny)
        ile_razy = st.number_input(f"Ile razy liczyć: {atrakcja}? (np. ilość wynajmów/grup)", min_value=1, value=1)
        cena_za_atrakcje = ile_razy * dane_atrakcji["cena"]
        szczegoly_atrakcji.append({"Nazwa": atrakcja, "Ilość": ile_razy, "Cena": cena_za_atrakcje})
    koszt_atrakcje += cena_za_atrakcje

# --- 5. PODSUMOWANIE I GENEROWANIE ---
st.markdown('<h3 class="section-header">5. Podsumowanie i PDF</h3>', unsafe_allow_html=True)

suma_calkowita = koszt_dworek + koszt_domki + koszt_wyzywienie + koszt_atrakcje

col_sum1, col_sum2 = st.columns(2)
with col_sum1:
    st.write(f"Koszt noclegów (Dworek): **{koszt_dworek} PLN**")
    st.write(f"Koszt noclegów (Domki): **{koszt_domki} PLN**")
    st.write(f"Koszt wyżywienia: **{koszt_wyzywienie} PLN**")
    st.write(f"Koszt atrakcji: **{koszt_atrakcje} PLN**")
with col_sum2:
    st.markdown(f"### RAZEM: {suma_calkowita} PLN")

if st.button("Generuj Ofertę PDF", type="primary"):
    if not klient_imie_nazwisko:
        st.error("Podaj imię i nazwisko klienta!")
    elif l_osob_przypisanych != l_osob_total:
        st.error("Popraw błędy w przypisaniu miejsc noclegowych przed wygenerowaniem oferty.")
    else:
        st.success("Dane poprawne. System jest gotowy na przyjęcie logiki łączącej PDF.")
        st.info("""
        W kolejnym kroku podepniemy tu bibliotekę `pypdf` / `reportlab`, która złoży:
        1. PDF Okładka
        2. PDF Wstęp (Daniel Heina)
        3. PDF Ośrodek (Dworek/Domki)
        4. PDF Atrakcje
        5. PDF Tabela cenowa (wygenerowana dynamicznie na bazie powyższych kwot)
        """)