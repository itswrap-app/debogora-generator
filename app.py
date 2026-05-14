import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import base64
import os
import subprocess
import io

# Importy do generowania i edycji plików
from pptx import Presentation
from pypdf import PdfWriter, PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# --- KONFIGURACJA STRONY ---
st.set_page_config(page_title="Generator Ofert - Dwór Dębogóra", layout="wide")

# Barwy z CI Dębogóry
CI = {
    "dark_green": "#00622f",
    "light_green": "#e8ece6",
    "gray": "#333333",
    "white": "#ffffff"
}

# --- STYLIZACJA CSS ---
st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Lora:ital,wght@0,400;0,700;1,400&family=PT+Sans:wght@400;700&display=swap');

    .stApp {{
        background-color: {CI['white']};
        font-family: 'PT Sans', sans-serif;
        color: {CI['gray']};
    }}

    h1, h2, h3, h4 {{
        font-family: 'Lora', serif !important;
        color: {CI['dark_green']} !important;
        font-weight: 700 !important;
    }}

    div[data-testid="stVerticalBlock"] > div > div[data-testid="stVerticalBlock"] {{
        background-color: {CI['light_green']};
        padding: 2rem;
        border-radius: 0px;
        border-left: 5px solid {CI['dark_green']};
        margin-bottom: 1.5rem;
    }}

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
    }}

    div.stButton > button:hover {{
        background-color: {CI['gray']} !important;
    }}
    </style>
""", unsafe_allow_html=True)

# --- WYŚRODKOWANE LOGO ---
def get_base64_img(path):
    try:
        with open(path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except:
        return None

logo_b64 = get_base64_img("logo.png")

if logo_b64:
    st.markdown(
        f"""
        <div style="display: flex; justify-content: center; align-items: center; padding: 10px 0;">
            <img src="data:image/png;base64,{logo_b64}" width="120">
        </div>
        """,
        unsafe_allow_html=True
    )
else:
    st.markdown(f"<p style='text-align:center; color:{CI['dark_green']}'>[ Logo.png ]</p>", unsafe_allow_html=True)

st.markdown(f"<h1 style='text-align: center; margin-top: 0px;'>System Ofertowania</h1>", unsafe_allow_html=True)
st.markdown(f"<p style='text-align: center; color: {CI['gray']}; font-style: italic; margin-bottom: 30px;'>Dwór Dębogóra & Domki Krovacja</p>", unsafe_allow_html=True)

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
        "Paintball (min 10 os)": {"cena": 150, "typ": "osoba"},
        "Kajaki (min 4 os)": {"cena": 140, "typ": "osoba"},
        "Ognisko z drewnem": {"cena": 150, "typ": "grupa"},
        "Rower Elektryczny": {"cena": 120, "typ": "osoba"},
        "Wycieczka DPN (3h)": {"cena": 700, "typ": "grupa"}
    }
}

pozycje_kosztowe = []

# --- 1. DANE KLIENTA ---
with st.container():
    st.subheader("1. Dane Klienta i Termin")
    c1, c2 = st.columns(2)
    with c1:
        klient_imie = st.text_input("Imię i nazwisko (wymagane) *")
        firma_n = st.text_input("Nazwa Firmy")
        nip_n = st.text_input("NIP")
        l_osob_total = st.number_input("Łączna liczba osób", min_value=1, value=10)
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

# --- FUNKCJE POMOCNICZE (PPTX -> PDF -> MERGE) ---
def zamien_tekst_w_prezentacji(ppt, stara_wartosc, nowa_wartosc):
    for slide in ppt.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text") and stara_wartosc in shape.text:
                shape.text = shape.text.replace(stara_wartosc, nowa_wartosc)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if stara_wartosc in cell.text:
                            cell.text = cell.text.replace(stara_wartosc, nowa_wartosc)

def konwertuj_pptx_na_pdf(sciezka_pptx, folder_wyjsciowy):
    try:
        subprocess.run([
            "libreoffice", "--headless", "--convert-to", "pdf", 
            sciezka_pptx, "--outdir", folder_wyjsciowy
        ], check=True)
        return sciezka_pptx.replace(".pptx", ".pdf")
    except Exception as e:
        st.error(f"Błąd konwersji LibreOffice: {e}")
        return None

def generuj_pdf_tabela(df_dane, suma_calkowita, nazwa_klienta):
    sciezka = "tabela_kosztowa_temp.pdf"
    doc = SimpleDocTemplate(sciezka, pagesize=A4, rightMargin=40, leftMargin=40, topMargin=50, bottomMargin=50)
    elements = []
    
    ciemny_zielony = colors.HexColor("#00622f")
    jasny_zielony = colors.HexColor("#e8ece6")
    szary = colors.HexColor("#333333")
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], textColor=ciemny_zielony, fontSize=22, spaceAfter=20, alignment=1)
    subtitle_style = ParagraphStyle('SubTitleStyle', parent=styles['Normal'], textColor=szary, fontSize=12, spaceAfter=30, alignment=1)
    
    elements.append(Paragraph("Podsumowanie Kosztów Ofertowych", title_style))
    elements.append(Paragraph(f"Oferta przygotowana dla: <b>{nazwa_klienta}</b>", subtitle_style))
    
    tabela_dane = [["Kategoria", "Opis", "Ilość", "Cena Jedn.", "Suma"]]
    for _, row in df_dane.iterrows():
        tabela_dane.append([row["Kategoria"], row["Opis"], str(row["Ilość"]), f"{row['Cena']:.0f} zł", f"{row['Suma']:.0f} zł"])
    tabela_dane.append(["", "", "", "RAZEM:", f"{suma_calkowita:.0f} zł"])
    
    t = Table(tabela_dane, colWidths=[80, 200, 50, 80, 80])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), ciemny_zielony),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('ALIGN', (1,1), (1,-2), 'LEFT'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,0), 12),
        ('TOPPADDING', (0,0), (-1,0), 12),
        ('BACKGROUND', (0,1), (-1,-2), jasny_zielony),
        ('GRID', (0,0), (-1,-2), 1, colors.white),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        ('TEXTCOLOR', (3,-1), (-1,-1), ciemny_zielony),
        ('ALIGN', (3,-1), (4,-1), 'CENTER'),
        ('TOPPADDING', (0,-1), (-1,-1), 15),
    ]))
    
    elements.append(t)
    doc.build(elements)
    return sciezka

# --- 5. PODSUMOWANIE I GENEROWANIE ---
with st.container():
    st.subheader("5. Kosztorys Ofertowy")
    df = pd.DataFrame(pozycje_kosztowe)
    if not df.empty:
        edytowany_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        razem = edytowany_df["Suma"].sum()
        st.markdown(f"### RAZEM DO ZAPŁATY: {razem:,.2f} PLN".replace(",", " "))
        
        if st.button("GENERUJ FINALNĄ OFERTĘ PDF"):
            if not klient_imie:
                st.error("Błąd: Imię i nazwisko klienta jest wymagane!")
            else:
                with st.spinner("Przetwarzanie dokumentów..."):
                    # Ustalenie nazwy firmy z priorytetem na wpisaną nazwę lub imię klienta
                    nazwa_do_oferty = firma_n if firma_n else klient_imie

                    # KROK 1: Modyfikacja okładki PPTX
                    sciezka_okladka_pptx = "AsystentAI_okładka_02.pptx"
                    
                    if not os.path.exists(sciezka_okladka_pptx):
                        st.error(f"Brak pliku bazowego '{sciezka_okladka_pptx}' w folderze z aplikacją!")
                    else:
                        ppt = Presentation(sciezka_okladka_pptx)
                        # Reagujemy na różne warianty wpisania zmiennej
                        zamien_tekst_w_prezentacji(ppt, "{{nazwa firmy}}", nazwa_do_oferty)
                        zamien_tekst_w_prezentacji(ppt, "{{Nazwa_firmy}}", nazwa_do_oferty)
                        
                        zmodyfikowany_pptx = "okladka_temp.pptx"
                        ppt.save(zmodyfikowany_pptx)

                        # KROK 2: Konwersja PPTX -> PDF
                        okladka_pdf = konwertuj_pptx_na_pdf(zmodyfikowany_pptx, ".")
                        
                        if okladka_pdf:
                            # KROK 3: Generowanie tabeli cenowej do PDF
                            tabela_pdf = generuj_pdf_tabela(edytowany_df, razem, nazwa_do_oferty)

                            # KROK 4: Scalanie okładki i tabeli w jeden plik PDF
                            merger = PdfWriter()
                            merger.append(okladka_pdf)
                            merger.append(tabela_pdf)

                            output_filename = f"Oferta_Debogora_{nazwa_do_oferty.replace(' ', '_')}.pdf"
                            with open(output_filename, "wb") as f_out:
                                merger.write(f_out)

                            # Opcjonalne sprzątanie plików tymczasowych
                            try:
                                os.remove(zmodyfikowany_pptx)
                                os.remove(okladka_pdf)
                                os.remove(tabela_pdf)
                            except:
                                pass

                            # Przycisk pobierania gotowej połączonej oferty
                            with open(output_filename, "rb") as final_pdf:
                                st.download_button(
                                    label="📥 POBIERZ SCALONĄ OFERTĘ PDF",
                                    data=final_pdf,
                                    file_name=output_filename,
                                    mime="application/pdf",
                                    type="primary"
                                )
                            st.success("Oferta wygenerowana i pomyślnie scalona!")
