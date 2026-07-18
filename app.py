import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- KONFIGURACJA STRONY ---
st.set_page_config(page_title="System Wart Obozowych", page_icon="⛺", layout="wide")

# --- STYLE CSS ---
st.markdown("""
    <style>
        .stApp { background-color: #0e1117; color: #ffffff; }
        .warta-sekcja {
            background-color: #1f2937;
            padding: 15px;
            border-radius: 8px;
            margin-bottom: 10px;
            border: 1px solid #374151;
        }
        
        /* Klasyczny obozowy styl maszynopisu do druku */
        .rozkaz-kartka {
            background-color: white !important;
            color: black !important;
            padding: 30px;
            font-family: 'Courier New', Courier, monospace !important;
            border: 2px solid black;
            max-width: 800px;
            margin: 0 auto;
        }
        .rozkaz-kartka * { color: black !important; font-family: 'Courier New', Courier, monospace !important; }
        .tabela-rozkaz { width: 100%; border-collapse: collapse; margin-top: 15px; }
        .tabela-rozkaz th { border-bottom: 2px solid black; text-align: left; padding: 8px; font-weight: bold; }
        .tabela-rozkaz td { padding: 12px 8px; border-bottom: 1px dashed #666666; }

        @media print {
            body * { visibility: hidden; }
            .rozkaz-kartka, .rozkaz-kartka * { visibility: visible; }
            .rozkaz-kartka {
                position: absolute;
                left: 0;
                top: 0;
                width: 100%;
                border: none;
                padding: 0;
            }
        }
    </style>
""", unsafe_allow_html=True)

# --- INICJALIZACJA ZMIENNYCH W PAMIĘCI ---
if 'dane_uczestnikow' not in st.session_state: st.session_state.dane_uczestnikow = None
if 'harmonogram_wart' not in st.session_state: st.session_state.harmonogram_wart = {}
if 'liczba_wartowników' not in st.session_state: st.session_state.liczba_wartowników = {}
if 'lokalizacje_wart' not in st.session_state: st.session_state.lokalizacje_wart = {}

GODZINY_WART = {
    "22:00 - 23:00": ["Z"],
    "23:00 - 00:00": ["Z"],
    "00:00 - 02:00": ["H", "HS", "W", "I"],
    "02:00 - 04:00": ["W", "I", "HS"],
    "04:00 - 06:00": ["H", "HS", "W", "I"],
    "06:00 - 08:00": ["H", "HS", "W", "I"]
}

DNI = [f"{i:02d}.07" for i in range(19, 31)]
KOLORY_PIONOW = {'Z': '🟡', 'H': '🟢', 'HS': '🔵', 'W': '🔴', 'I': '⚪'}

# --- BOCZNY PANEL - WGÓROWANIE EXCELA ---
with st.sidebar:
    st.header("📥 Ładowanie Listy Obozowej")
    plik = st.file_uploader("Wgraj plik Excel (.xlsx)", type=["xlsx"])
    
    if plik and st.session_state.dane_uczestnikow is None:
        try:
            df = pd.read_excel(plik).fillna("")
            df.columns = [str(c).strip() for c in df.columns]
            
            mapowanie = {}
            for col in df.columns:
                cl = col.lower()
                if 'imi' in cl: mapowanie['Imię'] = col
                elif 'nazw' in cl: mapowanie['Nazwisko'] = col
                elif 'pion' in cl or 'grupa' in cl: mapowanie['Pion'] = col
                elif 'druż' in cl or 'druz' in cl: mapowanie['Drużyna'] = col
                elif 'wart' in cl or 'liczba' in cl: mapowanie['Liczba_Wart'] = col

            final_df = pd.DataFrame()
            final_df['Imię'] = df[mapowanie.get('Imię', df.columns[0])]
            final_df['Nazwisko'] = df[mapowanie.get('Nazwisko', df.columns[1])]
            final_df['Pion'] = df[mapowanie.get('Pion', df.columns[2])].astype(str).str.upper().str.strip()
            final_df['Drużyna'] = df[mapowanie.get('Drużyna', df.columns[3])] if 'Drużyna' in mapowanie else ""
            final_df['Liczba_Wart'] = pd.to_numeric(df[mapowanie.get('Liczba_Wart')], errors='coerce').fillna(0).astype(int) if 'Liczba_Wart' in mapowanie else 0
            
            final_df['Nazwa_Pelna'] = final_df['Imię'] + " " + final_df['Nazwisko'] + " (" + final_df['Pion'] + ")"
            st.session_state.dane_uczestnikow = final_df
            st.success("Lista wgrana pomyślnie!")
        except Exception as e:
            st.error(f"Błąd struktury pliku: {e}")

# --- PANEL GŁÓWNY ---
st.title("⛺ Kreator i Statystyki Wart Obozowych")

if st.session_state.dane_uczestnikow is None:
    st.info("Proszę wgrać plik Excel (.xlsx) w panelu bocznym, aby rozpocząć pracę.")
else:
    # --- SEKCJA 1: STATYSTYKI I RAPORTY DO POBRANIA ---
    st.header("📊 Analiza i Pobieranie Raportów")
    tab1, tab2, tab3 = st.tabs(["📈 Bieżący Licznik Wart", "🔍 Kto jeszcze nie stał?", "💾 Pobierz Końcowy Excel"])
    
    with tab1:
        st.dataframe(
            st.session_state.dane_uczestnikow[['Pion', 'Imię', 'Nazwisko', 'Drużyna', 'Liczba_Wart']].sort_values('Liczba_Wart'), 
            hide_index=True,
            use_container_width=True
        )
        
    with tab2:
        nie_byli = st.session_state.dane_uczestnikow[st.session_state.dane_uczestnikow['Liczba_Wart'] == 0]
        if not nie_byli.empty:
            st.warning(f"Liczba osób, które jeszcze nie odbyły żadnej warty: {len(nie_byli)}")
            st.dataframe(nie_byli[['Pion', 'Imię', 'Nazwisko', 'Drużyna']], hide_index=True, use_container_width=True)
        else:
            st.success("🎉 Wszyscy uczestnicy obozu pełnili już wartę chociaż raz!")

    with tab3:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            st.session_state.dane_uczestnikow[['Imię', 'Nazwisko', 'Pion', 'Drużyna', 'Liczba_Wart']].to_excel(writer, index=False, sheet_name='Raport Wart')
        buffer.seek(0)
        
        st.download_button(
            label="📥 POBIERZ UAKTUALNIONY ARKUSZ EXCEL (.XLSX)",
            data=buffer,
            file_name=f"raport_koncowy_wart_{datetime.now().strftime('%d_%m')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.markdown("---")
    
    # --- SEKCJA 2: KREATOR WART ---
    st.header("🛠️ Kreator Przydziału na Grafik")
    wybrany_dzien = st.selectbox("📅 Wybierz datę nocy:", DNI)

    if wybrany_dzien not in st.session_state.harmonogram_wart:
        st.session_state.harmonogram_wart[wybrany_dzien] = {g: [] for g in GODZINY_WART.keys()}
        st.session_state.liczba_wartowników[wybrany_dzien] = {g: 2 for g in GODZINY_WART.keys()}
        st.session_state.lokalizacje_wart[wybrany_dzien] = {g: [] for g in GODZINY_WART.keys()}

    dzisiejsze_warty = st.session_state.harmonogram_wart[wybrany_dzien]
    dzisiejsze_miejsca = st.session_state.lokalizacje_wart[wybrany_dzien]

    idx_dnia = DNI.index(wybrany_dzien)
    poprzedni_dzien = DNI[idx_dnia - 1] if idx_dnia > 0 else None

    # Zbieranie danych wejściowych w formularzu Streamlit
    with st.form(key=f"formularz_wart_{wybrany_dzien}"):
        for godzina, preferowane_piony in GODZINY_WART.items():
            st.markdown(f"<div class='warta-sekcja'>", unsafe_allow_html=True)
            c_info, c_wybory = st.columns([3, 5])
            
            with c_info:
                st.markdown(f"### ⏰ {godzina}")
                zuchy_ok = "Z" in preferowane_piony
                st.caption(f"Sugerowane: {', '.join(preferowane_piony)} " + ("" if zuchy_ok else "(ZAKAZ ZUCHÓW)"))
                wybrane_piony = st.multiselect("Filtruj pion wiekowy:", ["Z", "H", "HS", "W", "I"], default=preferowane_piony, key=f"pion_{godzina}_{wybrany_dzien}")

            ile_miejsc = st.session_state.liczba_wartowników[wybrany_dzien][godzina]
            while len(dzisiejsze_warty[godzina]) < ile_miejsc: dzisiejsze_warty[godzina].append("")
            while len(dzisiejsze_miejsca[godzina]) < ile_miejsc: dzisiejsze_miejsca[godzina].append("")
            dzisiejsze_warty[godzina] = dzisiejsze_warty[godzina][:ile_miejsc]
            dzisiejsze_miejsca[godzina] = dzisiejsze_miejsca[godzina][:ile_miejsc]

            with c_wybory:
                kolumny_slotow = st.columns(ile_miejsc)
                for i in range(ile_miejsc):
                    with kolumny_slotow[i]:
                        baza = st.session_state.dane_uczestnikow
                        if wybrane_piony:
                            baza = baza[baza['Pion'].isin(wybrane_piony)]
                        
                        baza_posortowana = baza.sort_values(by=['Liczba_Wart', 'Pion'])
                        lista_wyboru = ["-- Wybierz wartownika --", "⚠️ Ręczny wpis spoza listy"]
                        for _, row in baza_posortowana.iterrows():
                            ico = KOLORY_PIONOW.get(row['Pion'], '▪️')
                            druzyna = f" [{row['Drużyna']}]" if row['Drużyna'] else ""
                            lista_wyboru.append(f"{ico} {row['Nazwa_Pelna']}{druzyna} (Służb: {row['Liczba_Wart']})")

                        obecny = dzisiejsze_warty[godzina][i]
                        indeks_startowy = 0
                        reczny_tryb = False

                        if obecny:
                            dopasowania = [idx for idx, tekst in enumerate(lista_wyboru) if obecny in tekst]
                            if dopasowania: indeks_startowy = dopasowania[0]
                            else: indeks_startowy = 1; reczny_tryb = True

                        wybrana_pozycja = st.selectbox(f"Wartownik {i+1}", lista_wyboru, index=indeks_startowy, key=f"sb_{godzina}_{i}_{wybrany_dzien}")
                        
                        if wybrana_pozycja == "⚠️ Ręczny wpis spoza listy" or reczny_tryb:
                            dzisiejsze_warty[godzina][i] = st.text_input("Wpisz dane:", value=obecny, key=f"ti_{godzina}_{i}_{wybrany_dzien}")
                        elif wybrana_pozycja != "-- Wybierz wartownika --":
                            dzisiejsze_warty[godzina][i] = wybrana_pozycja.split(" (Służb:")[0][2:].split(" [")[0]
                        else:
                            dzisiejsze_warty[godzina][i] = ""

                        dzisiejsze_miejsca[godzina][i] = st.text_input("Posterunek:", value=dzisiejsze_miejsca[godzina][i], key=f"pos_{godzina}_{i}_{wybrany_dzien}")

                        osoba = dzisiejsze_warty[godzina][i]
                        if osoba:
                            if " (Z)" in osoba and not zuchy_ok:
                                st.error("🛑 ZAKAZ: Zuchy w nocy!")
                            if poprzedni_dzien and poprzedni_dzien in st.session_state.harmonogram_wart:
                                wszyscy_wczoraj = [l for podlista in st.session_state.harmonogram_wart[poprzedni_dzien].values() for l in podlista]
                                if osoba in wszyscy_wczoraj:
                                    st.warning("⚠️ Ta osoba stała wczoraj!")
            st.markdown("</div>", unsafe_allow_html=True)

        # Przycisk wysyłania formularza trwale zapisujący dane
        zapisz_klikniete = st.form_submit_button("💾 ZAPISZ GRAFIK I PRZELICZ STATYSTYKI", use_container_width=True, type="primary")
        
        if zapisz_klikniete:
            st.session_state.harmonogram_wart[wybrany_dzien] = dzisiejsze_warty
            st.session_state.lokalizacje_wart[wybrany_dzien] = dzisiejsze_miejsca
            
            # Przeliczenie sumy wszystkich wart od nowa
            st.session_state.dane_uczestnikow['Liczba_Wart'] = 0
            for d_klucz, warty_dniowe in st.session_state.harmonogram_wart.items():
                for g_klucz, spis_osob in warty_dniowe.items():
                    for os in spis_osob:
                        if os:
                            match = st.session_state.dane_uczestnikow['Nazwa_Pelna'] == os
                            if match.any():
                                st.session_state.dane_uczestnikow.loc[match, 'Liczba_Wart'] += 1
            st.success("Zapisano pomyślnie!")
            st.rerun()

    # Zmiana liczby wartowników poza głównym formularzem (aby uniknąć resetu wyboru)
    st.write("🔧 Szybka zmiana liczby stanowisk dla wybranej nocy:")
    c_godz, c_plus, c_minus = st.columns([4, 2, 2])
    with c_godz:
        g_zmiana = st.selectbox("Wybierz godzinę do modyfikacji:", list(GODZINY_WART.keys()), key="zmiana_godz")
    with c_plus:
        if st.button("➕ Dodaj stanowisko", use_container_width=True):
            st.session_state.liczba_wartowników[wybrany_dzien][g_zmiana] += 1
            st.rerun()
    with c_minus:
        if st.button("➖ Usuń stanowisko", use_container_width=True):
            if st.session_state.liczba_wartowników[wybrany_dzien][g_zmiana] > 1:
                st.session_state.liczba_wartowników[wybrany_dzien][g_zmiana] -= 1
                st.rerun()

    # --- SEKCJA 3: PODGLĄD ROZKAZU DO DRUKU ---
    st.markdown("---")
    st.subheader("🖨️ Generator Rozkazu (Podgląd A4)")
    
    nastepny_dzien = f"{(int(wybrany_dzien.split('.')[0]) + 1):02d}.07"
    wiersze_tabeli_html = ""
    for g, osoby in dzisiejsze_warty.items():
        lista_dla_godziny = []
        for j, o in enumerate(osoby):
            miejsce = dzisiejsze_miejsca[g][j] if j < len(dzisiejsze_miejsca[g]) else ""
            miejsce_str = f" ({miejsce})" if miejsce else ""
            wartownik_str = o if o else "........................................."
            lista_dla_godziny.append(f"{wartownik_str}{miejsce_str}")
        
        wiersze_tabeli_html += f"""
        <tr>
            <td style="padding: 12px 8px; border-bottom: 1px dashed #666666;"><b>{g}</b></td>
            <td style="padding: 12px 8px; border-bottom: 1px dashed #666666;">{', '.join(lista_dla_godziny)}</td>
        </tr>
        """

    st.html(f"""
    <div class="rozkaz-kartka">
        <div style="text-align: center; border-bottom: 2px solid black; padding-bottom: 10px; margin-bottom: 20px;">
            <h1 style="margin: 0; font-size: 24px; text-transform: uppercase;">ROZKAZ NA WARTĘ NOCNĄ</h1>
            <h3 style="margin: 5px 0 0 0; font-weight: normal;">Noc: {wybrany_dzien} / {nastepny_dzien} 2026 r.</h3>
        </div>
        <table class="tabela-rozkaz">
            <thead>
                <tr>
                    <th style="width: 30%;">GODZINY</th>
                    <th style="width: 70%;">DRUHNA / DRUH (POSTERUNEK)</th>
                </tr>
            </thead>
            <tbody>
                {wiersze_tabeli_html}
            </tbody>
        </table>
        <div style="margin-top: 40px; text-align: center; font-size: 14px;">
            <p>Czuwaj!</p>
            <p style="font-style: italic;">Odprawa wartowników odbędzie się podczas apelu wieczornego.</p>
            <br>
            <div style="text-align: right; font-weight: bold; padding-right: 20px;">Komendant Obozu</div>
        </div>
    </div>
    """)
    
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🖨️ WYDRUKUJ ROZKAZ (ZAPISZ JAKO PDF)", use_container_width=True):
        st.html("<script>window.print();</script>")
