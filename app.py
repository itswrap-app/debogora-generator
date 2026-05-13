import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta

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

# --- BAZA CENNIKA ---
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
        "Krowie Safari Rozszerzone": {"cena": 150, "typ": "osoba"},
        "Paintball (min 10 os)": {"cena": 150, "typ": "osoba"},
        "Kajaki (min 4 os)": {"cena": 140, "typ": "osoba"},
        "Ognisko z drewnem": {"cena": 150, "typ": "grupa"},
        "Rower Elektryczny": {"cena": 120, "typ": "osoba"},
        "Rower MTB": {"cena": 60, "typ": "osoba"},
        "Wycieczka DPN (3h)": {"cena": 700, "typ": "grupa"}
    }
}

st.markdown('<h1 class="main-header">Generator Ofert - Dwór Dębogóra</h1>', unsafe_allow_html=True)

# Lista do zbierania szczegółowych pozycji kosztowych
pozycje_kosztowe = []

# --- 1. DANE KLIENTA ---
st.markdown('<h3 class="section-header">1. Dane Klienta i Wydarzenia</h3>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    klient_imie_nazwisko = st.text_input("Imię i nazwisko klienta (wymagane) *")
    firma = st.text_input("Firma (opcjonalnie)")
    nip = st.text_input("NIP (opcjonalnie)")
    l_osob_total = st.number_input("Łączna liczba uczestników wydarzenia", min_value=1, value=10, step=1)
with col2:
    email = st.text_input("Adres e-mail (opcjonalnie)")
    telefon = st.text_input("Telefon (opcjonalnie)")
    
    col_data1, col_data2 = st.columns(2)
    with col_data1:
        data_przyjazdu = st.date_input("Data przyjazdu", value=date.today())
    with col_data2:
        data_wyjazdu = st.date_input("Data wyjazdu", value=date.today() + timedelta(days=1))
        
    l_dni = (data_wyjazdu - data_przyjazdu).days
    
    if l_dni < 1:
        st.error("Data wyjazdu musi być późniejsza niż data przyjazdu.")
        l_dni = 1 # Zabezpieczenie przed błędem wyliczeń
    else:
        st.info(f"Wyliczona liczba dób noclegowych: **{l_dni}**")

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
    
    if pokoje_1 > 0:
        pozycje_kosztowe.append({"Kategoria": "Zakwaterowanie", "Opis": f"Pokój 1-osobowy (x{pokoje_1}) x {l_dni} dób", "Ilość": pokoje_1, "Cena Jedn. (PLN)": 1 * stawka_dworek * l_dni, "Suma (PLN)": pokoje_1 * 1 * stawka_dworek * l_dni})
    if pokoje_2 > 0:
        pozycje_kosztowe.append({"Kategoria": "Zakwaterowanie", "Opis": f"Pokój 2-osobowy (x{pokoje_2}) x {l_dni} dób", "Ilość": pokoje_2, "Cena Jedn. (PLN)": 2 * stawka_dworek * l_dni, "Suma (PLN)": pokoje_2 * 2 * stawka_dworek * l_dni})
    if pokoje_3 > 0:
        pozycje_kosztowe.append({"Kategoria": "Zakwaterowanie", "Opis": f"Pokój 3-osobowy (x{pokoje_3}) x {l_dni} dób", "Ilość": pokoje_3, "Cena Jedn. (PLN)": 3 * stawka_dworek * l_dni, "Suma (PLN)": pokoje_3 * 3 * stawka_dworek * l_dni})

with col_domki:
    st.subheader("Domki Krovacja")
    wybrane_domki = st.multiselect("Wybierz domki do rezerwacji", options=list(CENNIK["domki"].keys()))
    
    osoby_domki = 0
    for domek in wybrane_domki:
        parametry = CENNIK["domki"][domek]
        l_osob_w_domku = st.number_input(f"Liczba osób w domku {domek} (max {parametry['max_os']})", min_value=1, max_value=parametry['max_os'], value=1)
        osoby_domki += l_osob_w_domku
        
        koszt_baza = parametry["baza"]
        doplata = (l_osob_w_domku - 1) * CENNIK["doplata_domek"] if l_osob_w_domku > 1 else 0
        suma_za_domek = (koszt_baza + doplata) * l_dni
        
        opis = f"Domek {domek} ({l_osob_w_domku} os.): Baza {koszt_baza} zł"
        if doplata > 0:
            opis += f" + {l_osob_w_domku - 1} x {CENNIK['doplata_domek']} zł (dopłata)"
        opis += f" x {l_dni} dób"
            
        pozycje_kosztowe.append({"Kategoria": "Zakwaterowanie", "Opis": opis, "Ilość": 1, "Cena Jedn. (PLN)": suma_za_domek, "Suma (PLN)": suma_za_domek})

# WALIDACJA LICZBY OSÓB
l_osob_przypisanych = osoby_dworek + osoby_domki
if l_osob_przypisanych != l_osob_total:
    st.error(f"⚠️ BŁĄD: Przypisano {l_osob_przypisanych} miejsc noclegowych, a deklarowana liczba uczestników to {l_osob_total}.")

# --- 3. WYŻYWIENIE ---
st.markdown('<h3 class="section-header">3. Wyżywienie</h3>', unsafe_allow_html=True)
opcja_wyzywienia = st.selectbox("Wybierz wariant wyżywienia", options=list(CENNIK["wyzywienie"].keys()))
if opcja_wyzywienia != "Brak wyżywienia":
    cena_wyz = CENNIK["wyzywienie"][opcja_wyzywienia]
    suma_wyz = cena_wyz * l_osob_total * l_dni
    pozycje_kosztowe.append({"Kategoria": "Wyżywienie", "Opis": f"{opcja_wyzywienia} ({l_dni} dób)", "Ilość": l_osob_total, "Cena Jedn. (PLN)": cena_wyz * l_dni, "Suma (PLN)": suma_wyz})

# --- 4. ATRAKCJE DODATKOWE ---
st.markdown('<h3 class="section-header">4. Atrakcje Dodatkowe</h3>', unsafe_allow_html=True)
wybrane_atrakcje = st.multiselect("Wybierz atrakcje z listy", options=list(CENNIK["atrakcje"].keys()))

for atrakcja in wybrane_atrakcje:
    dane_atrakcji = CENNIK["atrakcje"][atrakcja]
    if dane_atrakcji["typ"] == "osoba":
        l_chetnych = st.number_input(f"Ilu chętnych na: {atrakcja}?", min_value=1, max_value=l_osob_total, value=l_osob_total)
        suma_atr = l_chetnych * dane_atrakcji["cena"]
        pozycje_kosztowe.append({"Kategoria": "Atrakcje", "Opis": atrakcja, "Ilość": l_chetnych, "Cena Jedn. (PLN)": dane_atrakcji["cena"], "Suma (PLN)": suma_atr})
    else:
        ile_razy = st.number_input(f"Ile razy liczyć: {atrakcja}? (np. wynajmy/grupy)", min_value=1, value=1)
        suma_atr = ile_razy * dane_atrakcji["cena"]
        pozycje_kosztowe.append({"Kategoria": "Atrakcje", "Opis": atrakcja, "Ilość": ile_razy, "Cena Jedn. (PLN)": dane_atrakcji["cena"], "Suma (PLN)": suma_atr})

# --- 5. EDYTOWALNA TABELA KOSZTÓW ---
st.markdown('<h3 class="section-header">5. Edycja Kosztorysu przed generowaniem</h3>', unsafe_allow_html=True)
st.write("Tabela została wygenerowana automatycznie. Możesz ręcznie skorygować opisy, ceny lub usunąć wiersze, a suma całkowita zostanie przeliczona na nowo.")

df_pozycje = pd.DataFrame(pozycje_kosztowe)

if not df_pozycje.empty:
    edytowane_df = st.data_editor(
        df_pozycje,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Suma (PLN)": st.column_config.NumberColumn(
                "Suma (PLN)",
                min_value=0,
                format="%d zł"
            ),
            "Cena Jedn. (PLN)": st.column_config.NumberColumn(
                "Cena Jedn. (PLN)",
                min_value=0,
                format="%d zł"
            )
        }
    )
    suma_calkowita = edytowane_df["Suma (PLN)"].sum()
else:
    st.info("Kalkulator nie zawiera jeszcze żadnych pozycji.")
    suma_calkowita = 0

st.markdown(f"### RAZEM: {suma_calkowita} PLN")

if st.button("Generuj Ofertę PDF", type="primary"):
    if not klient_imie_nazwisko:
        st.error("Podaj imię i nazwisko klienta!")
    elif l_osob_przypisanych != l_osob_total:
        st.error("Popraw błędy w przypisaniu miejsc noclegowych przed wygenerowaniem oferty.")
    else:
        st.success("Dane gotowe do wygenerowania tabeli na pliku PDF.")
        # Oczekuje na skrypty pypdf / reportlab do wstawienia `edytowane_df` na stronę
