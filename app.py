import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# --- KONFIGURACJA STRONY ---
st.set_page_config(page_title="System Wart Obozowych", page_icon="⛺", layout="wide")

# --- INICJALIZACJA BAZY DANYCH W PAMIĘCI (SESSION STATE) ---
# W pełnej wersji online zamiast st.session_state można podpiąć Google Sheets API
if 'db_uczestnicy' not in st.session_state:
    st.session_state.db_uczestnicy = None
if 'historia_wart' not in st.session_state:
    st.session_state.historia_wart = {}

# Słownik mapowania pionów dla czytelności interfejsu
PIONY_DICT = {'Z': 'Zuch', 'H': 'Harcerz', 'HS': 'Harcerz St.', 'W': 'Wędrownik', 'I': 'Instruktor'}

# Definicja wart i ich preferencji
WARTY_SPECYFIKACJA = {
    "22:00 - 23:00": {"preferencja": ["Z"], "opis": "Tylko Zuchy (Z)"},
    "23:00 - 00:00": {"preferencja": ["Z"], "opis": "Tylko Zuchy (Z)"},
    "00:00 - 02:00": {"preferencja": ["H", "HS", "W", "I"], "opis": "Starsze piony (Bez Zuchów)"},
    "02:00 - 04:00": {"preferencja": ["W", "I", "HS"], "opis": "Godziny ciężkie (Sugerowani W/I)"},
    "04:00 - 06:00": {"preferencja": ["H", "HS", "W", "I"], "opis": "Starsze piony"},
    "06:00 - 08:00": {"preferencja": ["H", "HS", "W", "I"], "opis": "Starsze piony"}
}

# Generowanie listy dni obozu (19.07 - 02.08.2026)
START_DATA = datetime(2026, 7, 19)
DNI_OBOZU = [(START_DATA + timedelta(days=i)).strftime("%d.%02m") for i in range(15)]

st.title("⛺ System Zarządzania Wartami Obosowymi")
st.write("Obóz: 19.07 - 02.08 | Automatyczne podpowiedzi i generowanie grafiku do druku.")

# --- PANEL BOCZNY: IMPORT EXCELA ---
with st.sidebar:
    st.header("📥 Import danych")
    uploaded_file = st.file_uploader("Wgraj plik Excel (.xlsx)", type=["xlsx"])
    
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file)
            # Standaryzacja nazw kolumn (małe litery, usunięcie spacji)
            df.columns = [c.strip().lower() for c in df.columns]
            
            if 'imię' in df.columns and 'nazwisko' in df.columns and 'pion' in df.columns:
                # Czyszczenie danych i dodanie liczników
                df['pion'] = df['pion'].astype(str).str.upper().str.strip()
                df['pelne_nazwisko'] = df['imię'] + " " + df['nazwisko'] + " (" + df['pion'] + ")"
                
                # Inicjalizacja liczników jeśli baza ładuje się pierwszy raz
                if st.session_state.db_uczestnicy is None:
                    df['liczba_wart'] = 0
                    df['ostatnia_warta'] = "-"
                    st.session_state.db_uczestnicy = df
                st.success("Pomyślnie załadowano bazę uczestników!")
            else:
                st.error("Błąd! Excel musi zawierać kolumny: 'Imię', 'Nazwisko', 'Pion'.")
        except Exception as e:
            st.error(f"Błąd odczytu pliku: {e}")

    if st.session_state.db_uczestnicy is not None:
        st.header("📊 Szybkie statystyki")
        st.dataframe(st.session_state.db_uczestnicy[['imię', 'nazwisko', 'pion', 'liczba_wart']], hide_index=True)

# --- PANEL GŁÓWNY: KREATOR WART ---
if st.session_state.db_uczestnicy is None:
    st.info("👋 Aby rozpocząć, wgraj plik Excel w panelu bocznym. Plik powinien mieć kolumny: Imię, Nazwisko, Pion (Z, H, HS, W, I).")
else:
    # Wybór dnia obozu
    wybrany_dzien = st.selectbox("📅 Wybierz noc na którą planujesz wartę:", DNI_OBOZU, 
                                 format_func=lambda x: f"Noc {x} / {(datetime.strptime(x+'.2026', '%d.%m.%Y') + timedelta(days=1)).strftime('%d.%02m')}")

    # Pobranie lub stworzenie czystego planu na dany dzień
    if wybrany_dzien not in st.session_state.historia_wart:
        st.session_state.historia_wart[wybrany_dzien] = {godzina: "" for godzina in WARTY_SPECYFIKACJA.keys()}
    
    plan_dnia = st.session_state.historia_wart[wybrany_dzien]

    st.subheader(f"🛠️ Układanie planu na noc {wybrany_dzien}")
    
    # Renderowanie formularza dla każdej godziny
    for godzina, info in WARTY_SPECYFIKACJA.items():
        col1, col2, col3 = st.columns([1, 2, 1])
        
        with col1:
            st.metric(label="Godziny", value=godzina)
            st.caption(f"Preferowane: {info['opis']}")
            
        with col2:
            # Filtrowanie i sortowanie ludzi pod kątem podpowiedzi
            db = st.session_state.db_uczestnicy
            
            # Sortujemy: najpierw preferowany pion, potem najmniejsza liczba wart
            db['jest_preferowany'] = db['pion'].isin(info['preferencja'])
            db_posortowana = db.sort_values(by=['jest_preferowany', 'liczba_wart'], ascending=[False, True])
            
            # Tworzenie opcji do wyboru z liczbą wart w nawiasie np. "Jan Kowalski (Z) [Warty: 0]"
            opcje_wyboru = ["-- Wybierz osobę --", "WYJĄTEK (Wpis ręczny)"] + list(
                db_posortowana['pelne_nazwisko'] + " [Warty: " + db_posortowana['liczba_wart'].astype(str) + "]"
            )
            
            # Szukanie co jest aktualnie wybrane
            aktualna_osoba = plan_dnia[godzina]
            index_domyslny = 0
            is_manual = False
            
            if aktualna_osoba:
                if any(aktualna_osoba in o for o in opcje_wyboru):
                    index_domyslny = [i for i, o in enumerate(opcje_wyboru) if aktualna_osoba in o][0]
                else:
                    index_domyslny = 1 # Wyjątek / wpis ręczny
                    is_manual = True

            wybor = st.selectbox(f"Obsada na {godzina}", opcje_wyboru, index=index_domyslny, key=f"sel_{godzina}_{wybrany_dzien}")
            
            if wybor == "WYJĄTEK (Wpis ręczny)" or is_manual:
                wartość_reczna = st.text_input(f"Wpisz kto stoi (wyjątek) na {godzina}:", value=aktualna_osoba if is_manual else "", key=f"txt_{godzina}_{wybrany_dzien}")
                plan_dnia[godzina] = wartość_reczna
            elif wybor != "-- Wybierz osobę --":
                # Wyciągamy czyste imię i nazwisko z ciągu opcji
                czyste_nazwisko = wybor.split(" [")[0]
                plan_dnia[godzina] = czyste_nazwisko
            else:
                plan_dnia[godzina] = ""
                
        with col3:
            # Walidacja w czasie rzeczywistym (Alerty)
            if plan_dnia[godzina] and not is_manual and " (Z)" not in plan_dnia[godzina] and "Z" in info['preferencja']:
                st.warning("⚠️ Na tej godzinie powinny stać Zuchy (Z)!")
            elif plan_dnia[godzina] and " (Z)" in plan_dnia[godzina] and "Z" not in info['preferencja']:
                st.error("🛑 Zuch nie może stać po północy!")
            elif plan_dnia[godzina]:
                st.success("✅ Przypisano poprawnie")

    # Przycisk zapisu i przeliczenia statystyk
    if st.button("💾 Zapisz plan i zaktualizuj liczniki wart", type="primary"):
        st.session_state.historia_wart[wybrany_dzien] = plan_dnia
        
        # Resetowanie liczników przed ponownym przeliczeniem z historii
        st.session_state.db_uczestnicy['liczba_wart'] = 0
        
        # Przeliczenie ile razy kto stał na podstawie całej historii
        for dzien, warty in st.session_state.historia_wart.items():
            for godz, osoba in warty.items():
                if osoba:
                    # Dopasowanie jeśli to standardowy wpis z bazy
                    maska = st.session_state.db_uczestnicy['pelne_nazwisko'] == osoba
                    if maska.any():
                        st.session_state.db_uczestnicy.loc[maska, 'liczba_wart'] += 1
                        st.session_state.db_uczestnicy.loc[maska, 'ostatnia_warta'] = dzien
                        
        st.success("Plan zapisany! Liczniki w bazie zostały zaktualizowane.")
        st.rerun()

    # --- SEKCOJA GENEROWANIA GRAFIKI DO DRUKU (A4) ---
    st.markdown("---")
    st.subheader("🖨️ Generator arkusza do wydruku (A4)")
    
    jutro_data = (datetime.strptime(wybrany_dzien+".2026", "%d.%m.%Y") + timedelta(days=1)).strftime("%d.%02m")
    
    # Tworzenie czystego kodu HTML stylizowanego pod wydruk A4
    html_layout = f"""
    <div style="font-family: 'Courier New', Courier, monospace; border: 4px double #000; padding: 40px; background-color: white; color: black; max-width: 650px; margin: 0 auto;">
        <div style="text-align: center; border-bottom: 2px solid #000; padding-bottom: 10px; margin-bottom: 30px;">
            <h1 style="margin: 0; font-size: 28px; text-transform: uppercase; letter-spacing: 2px;">Rozkaz na Wartę Nocną</h1>
            <h2 style="margin: 5px 0 0 0; font-size: 20px;">Noc: {wybrany_dzien} / {jutro_data} 2026 r.</h2>
        </div>
        
        <table style="width: 100%; border-collapse: collapse; font-size: 18px;">
            <thead>
                <tr style="border-bottom: 2px solid #000;">
                    <th style="text-align: left; padding: 10px; width: 40%;">GODZINY</th>
                    <th style="text-align: left; padding: 10px; width: 60%;">DRUHNA / DRUH</th>
                </tr>
            </thead>
            <tbody>
    """
    
    for g, o in plan_dnia.items():
        obsada = o if o else "........................................."
        html_layout += f"""
                <tr style="border-bottom: 1px dashed #666;">
                    <td style="padding: 20px 10px; font-weight: bold;">{g}</td>
                    <td style="padding: 20px 10px;">{obsada}</td>
                </tr>
        """
        
    html_layout += """
            </tbody>
        </table>
        <div style="text-align: center; margin-top: 5px; font-size: 14px; font-style: italic;">
            <br><br>
            Czuwaj!<br>
            Odprawa wart przy apelu wieczornym.
        </div>
    </div>
    """
    
    # Wyświetlenie podglądu wydruku na stronie
    st.components.v1.html(html_layout, height=650, scrolling=True)
    st.info("💡 Aby wydrukować powyższy grafik: kliknij na nim prawym przyciskiem myszy -> 'Drukuj' (lub Ctrl+P) i wybierz 'Zapisz jako PDF' lub bezpośrednio swoją drukarkę obozową.")