import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

# --- KONFIGURACJA STRONY ---
st.set_page_config(page_title="System Wart Obozowych", page_icon="⛺", layout="wide", initial_sidebar_state="expanded")

# --- STYLE CSS (Stylizacja aplikacji oraz reguły czystego druku PDF) ---
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;600;700&display=swap');
        * { font-family: 'Space Grotesk', sans-serif; }
        .stApp { background: radial-gradient(circle at 50% 50%, #0f172a, #020617); color: #f8fafc; }
        [data-testid="stSidebar"] { background-color: #0b0f19 !important; border-right: 2px solid #1e293b; }
        
        .main-title {
            font-size: 32px; font-weight: 700; letter-spacing: 1px; color: #3b82f6; margin-bottom: 20px;
        }
        .warta-card {
            background: rgba(30, 41, 59, 0.4); border: 1px solid #334155; 
            border-radius: 12px; padding: 15px; margin-bottom: 12px;
        }
        
        /* KLASYCZNY WYGLĄD STARODAWNEGO WYDRUKU MASZYNOPISU */
        .rozkaz-kontener {
            background-color: white !important;
            color: black !important;
            padding: 40px;
            max-width: 680px;
            margin: 30px auto;
            border: 4px double #000000;
            font-family: 'Courier New', Courier, monospace !important;
            box-shadow: 0 4px 15px rgba(0,0,0,0.5);
        }
        .rozkaz-kontener * { font-family: 'Courier New', Courier, monospace !important; color: black !important; }
        .rozkaz-tabela { width: 100%; border-collapse: collapse; margin-top: 20px; }
        .rozkaz-tabela th { text-align: left; padding: 10px; border-bottom: 2px solid #000000; font-size: 16px; font-weight: bold; }
        .rozkaz-tabela td { padding: 15px 10px; font-size: 16px; }
        .rozkaz-linia { border-bottom: 1px dashed #444444; }

        /* UKRYWANIE ELEMENTÓW INTERFEJSU STREAMLIT PODCZAS DRUKOWANIA PDF */
        @media print {
            html, body, .stApp, [data-testid="stReportBlock"], div { 
                background: white !important; 
                color: black !important; 
                box-shadow: none !important;
                border: none !important;
                padding: 0 !important;
                margin: 0 !important;
            }
            [data-testid="stSidebar"], button, .main-title, header, hr, .stMarkdown, .stInfo, .stCaption, .stSelectbox, .stTextInput { 
                display: none !important; 
            }
            body * { visibility: hidden; }
            .rozkaz-kontener, .rozkaz-kontener * { 
                visibility: visible; 
            }
            .rozkaz-kontener {
                position: absolute; 
                left: 50%;
                top: 0;
                transform: translateX(-50%);
                width: 100%;
                max-width: 100%;
                border: 4px double #000000 !important;
                box-shadow: none !important;
                padding: 20px !important;
            }
        }
    </style>
""", unsafe_allow_html=True)

# --- INICJALIZACJA BAZY W PAMIĘCI ---
if 'db_uczestnicy' not in st.session_state: st.session_state.db_uczestnicy = None
if 'historia_wart' not in st.session_state: st.session_state.historia_wart = {}
if 'liczba_straznikow' not in st.session_state: st.session_state.liczba_straznikow = {}
if 'lokalizacje_wart' not in st.session_state: st.session_state.lokalizacje_wart = {}

PION_ORDER = {'Z': 0, 'H': 1, 'HS': 2, 'W': 3, 'I': 4}
PION_COLORS = {'Z': '🟡', 'H': '🟢', 'HS': '🔵', 'W': '🔴', 'I': '⚪'}

WARTY_SPECYFIKACJA = {
    "22:00 - 23:00": {"preferencja": ["Z"], "opis": "Młodszy pion - zalecane Zuchy [Z]"},
    "23:00 - 00:00": {"preferencja": ["Z"], "opis": "Młodszy pion - zalecane Zuchy [Z]"},
    "00:00 - 02:00": {"preferencja": ["H", "HS", "W", "I"], "opis": "Starsze piony - zakaz dla [Z]"},
    "02:00 - 04:00": {"preferencja": ["W", "I", "HS"], "opis": "Godziny nocne - zalecane starsze piony"},
    "04:00 - 06:00": {"preferencja": ["H", "HS", "W", "I"], "opis": "Starsze piony - zakaz dla [Z]"},
    "06:00 - 08:00": {"preferencja": ["H", "HS", "W", "I"], "opis": "Starsze piony - zakaz dla [Z]"}
}

START_DATA = datetime(2026, 7, 19)
DNI_OBOZU = [(START_DATA + timedelta(days=i)).strftime("%d.%02m") for i in range(15)]

# --- PANEL BOCZNY: IMPORT EXCELA ---
with st.sidebar:
    st.header("📥 Import danych")
    uploaded_file = st.file_uploader("Wgraj plik Excel (.xlsx)", type=["xlsx"])
    
    if uploaded_file and st.session_state.db_uczestnicy is None:
        try:
            df = pd.read_excel(uploaded_file).fillna("")
            cols_map = {str(c).strip().lower(): str(c) for c in df.columns}
            
            def znajdz_kolumne(keywords, default_name):
                for k, v in cols_map.items():
                    if any(kw in k for kw in keywords): return v
                return default_name

            col_imie = znajdz_kolumne(['imie', 'imię'], 'Imię')
            col_nazwisko = znajdz_kolumne(['nazwisko'], 'Nazwisko')
            col_pion = znajdz_kolumne(['pion', 'grupa_wiekowa'], 'Pion')
            col_druzyna = znajdz_kolumne(['druzyna', 'drużyna', 'zastęp', 'zastep'], 'Drużyna')
            col_liczba = znajdz_kolumne(['liczba_wart', 'warty', 'ile razy'], 'Liczba_Wart')

            for c, orig in [('Imię', col_imie), ('Nazwisko', col_nazwisko), ('Pion', col_pion), ('Drużyna', col_druzyna)]:
                if orig in df.columns: df[c] = df[orig]
                else: df[c] = ""
            
            df['Liczba_Wart'] = pd.to_numeric(df[col_liczba], errors='coerce').fillna(0).astype(int) if col_liczba in df.columns else 0
            df['Pion'] = df['Pion'].astype(str).str.upper().str.strip()
            df['pion_waga'] = df['Pion'].map(PION_ORDER).fillna(99)
            df['pelne_nazwisko'] = df['Imię'].astype(str) + " " + df['Nazwisko'].astype(str) + " (" + df['Pion'] + ")"
            
            st.session_state.db_uczestnicy = df[['Imię', 'Nazwisko', 'Pion', 'Drużyna', 'Liczba_Wart', 'pion_waga', 'pelne_nazwisko']]
            st.success("Pomyślnie załadowano bazę!")
        except Exception as e:
            st.error(f"Błąd pliku: {e}")

    if st.session_state.db_uczestnicy is not None:
        st.header("📊 Statystyki wart")
        st.dataframe(st.session_state.db_uczestnicy[['Pion', 'Imię', 'Nazwisko', 'Liczba_Wart']], hide_index=True)

# --- PANEL GŁÓWNY ---
st.markdown("<div class='main-title'>⛺ System Zarządzania Wartami Obozowymi</div>", unsafe_allow_html=True)

if st.session_state.db_uczestnicy is None:
    st.info("👋 Aby rozpocząć, wgraj plik Excel w panelu bocznym. Plik powinien zawierać kolumny: Imię, Nazwisko, Pion.")
else:
    # Wybór dnia
    wybrany_dzien = st.selectbox("📅 Wybierz noc, na którą planujesz wartę:", DNI_OBOZU, 
                                 format_func=lambda x: f"Noc {x} / {(datetime.strptime(x+'.2026', '%d.%m.%Y') + timedelta(days=1)).strftime('%d.%02m')}")

    if wybrany_dzien not in st.session_state.historia_wart: st.session_state.historia_wart[wybrany_dzien] = {godzina: [] for godzina in WARTY_SPECYFIKACJA.keys()}
    if wybrany_dzien not in st.session_state.liczba_straznikow: st.session_state.liczba_straznikow[wybrany_dzien] = {godzina: 2 for godzina in WARTY_SPECYFIKACJA.keys()}
    if wybrany_dzien not in st.session_state.lokalizacje_wart: st.session_state.lokalizacje_wart[wybrany_dzien] = {godzina: [] for godzina in WARTY_SPECYFIKACJA.keys()}

    plan_dnia = st.session_state.historia_wart[wybrany_dzien]
    lokalizacje_dnia = st.session_state.lokalizacje_wart[wybrany_dzien]

    st.subheader(f"🛠️ Układanie planu obsady")
    
    idx_dzis = DNI_OBOZU.index(wybrany_dzien)
    wczorajszy_dzien = DNI_OBOZU[idx_dzis - 1] if idx_dzis > 0 else None

    # Renderowanie wierszy godzinowych
    for godzina, info in WARTY_SPECYFIKACJA.items():
        st.markdown(f"<div class='warta-card'>", unsafe_allow_html=True)
        col_meta, col_inputs, col_actions = st.columns([2.5, 4.5, 1])
        
        with col_meta:
            st.markdown(f"<div style='font-size:18px; font-weight:700; color:#3b82f6;'>{godzina}</div>", unsafe_allow_html=True)
            st.caption(f"{info['opis']}")
            filtr_pionu = st.multiselect("Filtruj pion:", ["Z", "H", "HS", "W", "I"], default=info['preferencja'], key=f"f_{godzina}_{wybrany_dzien}")
            
        with col_actions:
            st.markdown("<div style='text-align:center; font-size:12px;'>Obsada</div>", unsafe_allow_html=True)
            c1, c2 = st.columns(2)
            with c1:
                if st.button("➕", key=f"a_{godzina}_{wybrany_dzien}"): st.session_state.liczba_straznikow[wybrany_dzien][godzina] += 1; st.rerun()
            with c2:
                if st.button("➖", key=f"s_{godzina}_{wybrany_dzien}"):
                    if st.session_state.liczba_straznikow[wybrany_dzien][godzina] > 1: st.session_state.liczba_straznikow[wybrany_dzien][godzina] -= 1; st.rerun()

        current_slots = st.session_state.liczba_straznikow[wybrany_dzien][godzina]
        while len(plan_dnia[godzina]) < current_slots: plan_dnia[godzina].append("")
        while len(lokalizacje_dnia[godzina]) < current_slots: lokalizacje_dnia[godzina].append("")
        
        plan_dnia[godzina] = plan_dnia[godzina][:current_slots]
        lokalizacje_dnia[godzina] = lokalizacje_dnia[godzina][:current_slots]

        with col_inputs:
            grid_cols = st.columns(max(1, current_slots))
            for slot_idx in range(current_slots):
                with grid_cols[slot_idx]:
                    db = st.session_state.db_uczestnicy
                    db_filtered = db[db['Pion'].isin(filtr_pionu)] if filtr_pionu else db
                    db_sorted = db_filtered.sort_values(by=['pion_waga', 'Liczba_Wart'], ascending=[True, True])
                    
                    opcje = ["-- Wybierz osobę --", "WYJĄTEK (Wpis ręczny)"]
                    for _, row in db_sorted.iterrows():
                        emoji = PION_COLORS.get(row['Pion'], '▪️')
                        druzyna_str = f" [{row['Drużyna']}]" if row['Drużyna'] else ""
                        opcje.append(f"{emoji} {row['pelne_nazwisko']}{druzyna_str} [Warty: {row['Liczba_Wart']}]")
                    
                    aktualny_wybor = plan_dnia[godzina][slot_idx]
                    idx_default = 0
                    is_manual = False
                    
                    if aktualny_wybor:
                        pasujace = [i for i, o in enumerate(opcje) if aktualny_wybor in o]
                        if pasujace: idx_default = pasujace[0]
                        else: idx_default = 1; is_manual = True
                            
                    wybor = st.selectbox(f"Wartownik {slot_idx+1}", opcje, index=idx_default, key=f"sel_{godzina}_{slot_idx}_{wybrany_dzien}")
                    
                    if wybor == "WYJĄTEK (Wpis ręczny)" or is_manual:
                        plan_dnia[godzina][slot_idx] = st.text_input(f"Kto stoi:", value=aktualny_wybor, key=f"txt_{godzina}_{slot_idx}_{wybrany_dzien}")
                    elif wybor != "-- Wybierz osobę --":
                        plan_dnia[godzina][slot_idx] = wybor.split(" [Warty:")[0][2:].split(" [")[0]
                    else:
                        plan_dnia[godzina][slot_idx] = ""

                    lokalizacje_dnia[godzina][slot_idx] = st.text_input("Posterunek:", value=lokalizacje_dnia[godzina][slot_idx], key=f"loc_{godzina}_{slot_idx}_{wybrany_dzien}", placeholder="np. Brama")

                    # Alerty walidacyjne
                    if plan_dnia[godzina][slot_idx]:
                        st_name = plan_dnia[godzina][slot_idx]
                        if " (Z)" in st_name and "Z" not in info['preferencja']:
                            st.markdown("<span style='color:#ef4444; font-size:11px; font-weight:bold;'>🛑 Zakaz dla Zuchów po północy!</span>", unsafe_allow_html=True)
                        if wczorajszy_dzien and wczorajszy_dzien in st.session_state.historia_wart:
                            if st_name in [item for sublist in st.session_state.historia_wart[wczorajszy_dzien].values() for item in sublist]:
                                st.markdown("<span style='color:#ec4899; font-size:11px; font-weight:bold;'>⚠️ Służba noc po nocy!</span>", unsafe_allow_html=True)
        st.markdown(f"</div>", unsafe_allow_html=True)

    if st.button("💾 Zapisz plan i zaktualizuj liczniki", type="primary", use_container_width=True):
        st.session_state.historia_wart[wybrany_dzien] = plan_dnia
        st.session_state.lokalizacje_wart[wybrany_dzien] = lokalizacje_dnia
        st.session_state.db_uczestnicy['Liczba_Wart'] = 0
        
        for d, warty in st.session_state.historia_wart.items():
            for g, osoby in warty.items():
                for osoba in osoby:
                    if osoba:
                        maska = st.session_state.db_uczestnicy['pelne_nazwisko'] == osoba
                        if maska.any(): st.session_state.db_uczestnicy.loc[maska, 'Liczba_Wart'] += 1
        st.success("Zapisano pomyślnie!")
        st.rerun()

    # --- SEKJA GENEROWANIA PODGLĄDU I DRUKU A4 NA DOLE STRONY ---
    st.markdown("---")
    st.subheader("🖨️ Podgląd arkusza rozkazu komendanta (A4)")
    st.info("💡 Kliknij przycisk poniżej, aby wywołać systemowe drukowanie. Wybierz 'Zapisz jako PDF' w oknie przeglądarki.")
    
    jutro_data = (datetime.strptime(wybrany_dzien+".2026", "%d.%m.%Y") + timedelta(days=1)).strftime("%d.%02m")

    # Bezpieczne budowanie czystej struktury tabeli HTML (bez konfliktów markdown)
    tabela_wiersze_html = ""
    for g, osoby in plan_dnia.items():
        elementy_obsady = []
        for idx, o in enumerate(osoby):
            miejsce = lokalizacje_dnia[g][idx] if idx < len(lokalizacje_dnia[g]) else ""
            miejsce_str = f" ({miejsce})" if miejsce else ""
            straznik = o if o else "........................................."
            elementy_obsady.append(f"{straznik}{miejsce_str}")
        
        obsada_finalna = ", ".join(elementy_obsady)
        tabela_wiersze_html += f"""
        <tr class="rozkaz-linia">
            <td style="font-weight: bold; padding: 15px 10px;">{g}</td>
            <td style="padding: 15px 10px;">{obsada_finalna}</td>
        </tr>
        """

    rozkaz_pelny_html = f"""
    <div class="rozkaz-kontener">
        <div style="text-align: center; border-bottom: 2px solid #000000; padding-bottom: 10px; margin-bottom: 20px;">
            <h1 style="margin: 0; font-size: 26px; text-transform: uppercase; letter-spacing: 2px;">Rozkaz na Wartę Nocną</h1>
            <h2 style="margin: 5px 0 0 0; font-size: 18px; font-weight: normal;">Noc: {wybrany_dzien} / {jutro_data} 2026 r.</h2>
        </div>
        <table class="rozkaz-tabela">
            <thead>
                <tr>
                    <th style="width: 35%;">GODZINY</th>
                    <th style="width: 65%;">DRUHNA / DRUH (POSTERUNEK)</th>
                </tr>
            </thead>
            <tbody>
                {tabela_wiersze_html}
            </tbody>
        </table>
        <div style="text-align: center; margin-top: 40px; font-size: 14px; font-style: italic;">
            Czuwaj!<br>
            Odprawa służb przy apelu wieczornym.<br><br>
            <div style="text-align: right; font-weight: bold; font-style: normal; margin-top: 15px;">Komendant Obozu</div>
        </div>
    </div>
    """

    # Wyświetlenie wygenerowanego rozkazu na dole ekranu
    st.markdown(rozkaz_pelny_html, unsafe_allow_html=True)
    
    # Przycisk uruchomienia druku okna
    if st.button("🖨️ URUCHOM DRUKOWANIE (ZAPISZ JAKO PDF)", type="secondary", use_container_width=True):
        st.html("<script>window.print();</script>")
