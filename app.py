import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# --- NOWOCZESNA I DYNAMICZNA KONFIGURACJA INTERFEJSU ---
st.set_page_config(
    page_title="Warty Obozowe 2026", 
    page_icon="⛺", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Stylizacja CSS dla nowoczesnego wyglądu (Leśny Dark Mode)
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght=400;500;600;700&display=swap');
        * { font-family: 'Plus Jakarta Sans', sans-serif; }
        .stApp { background-color: #121614; color: #e2e8f0; }
        [data-testid="stSidebar"] { background-color: #1a221f !important; border-right: 1px solid #2d3a34; }
        .warta-card {
            background: linear-gradient(145deg, #1e2622, #161d1a);
            border: 1px solid #2d3a34;
            border-radius: 16px;
            padding: 24px;
            margin-bottom: 20px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.2);
            transition: transform 0.2s, border-color 0.2s;
        }
        .warta-card:hover { transform: translateY(-2px); border-color: #4ade80; }
        h1, h2, h3 { color: #ffffff !important; font-weight: 700 !important; }
        .main-title {
            background: linear-gradient(45deg, #4ade80, #22c55e);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-size: 32px;
            font-weight: 800;
            margin-bottom: 5px;
        }
        .stButton>button { border-radius: 8px !important; transition: all 0.2s; }
    </style>
""", unsafe_allow_html=True)

# --- INICJALIZACJA STANU APLIKACJI ---
if 'db_uczestnicy' not in st.session_state:
    st.session_state.db_uczestnicy = None
if 'historia_wart' not in st.session_state:
    st.session_state.historia_wart = {}
if 'liczba_straznikow' not in st.session_state:
    st.session_state.liczba_straznikow = {}

WARTY_SPECYFIKACJA = {
    "22:00 - 23:00": {"preferencja": ["Z"], "opis": "Preferowane Zuchy (Z)"},
    "23:00 - 00:00": {"preferencja": ["Z"], "opis": "Preferowane Zuchy (Z)"},
    "00:00 - 02:00": {"preferencja": ["H", "HS", "W", "I"], "opis": "Służą starsze piony (Bez Z)"},
    "02:00 - 04:00": {"preferencja": ["W", "I", "HS"], "opis": "Środek nocy (Sugerowane W / I)"},
    "04:00 - 06:00": {"preferencja": ["H", "HS", "W", "I"], "opis": "Służą starsze piony"},
    "06:00 - 08:00": {"preferencja": ["H", "HS", "W", "I"], "opis": "Służą starsze piony"}
}

START_DATA = datetime(2026, 7, 19)
DNI_OBOZU = [(START_DATA + timedelta(days=i)).strftime("%d.%02m") for i in range(15)]

# --- PANEL BOCZNY ---
with st.sidebar:
    st.markdown("<div style='text-align: center; padding: 20px 0;'><span style='font-size: 45px;'>⛺</span></div>", unsafe_allow_html=True)
    st.markdown("### 📥 Panel Importu Bazy")
    uploaded_file = st.file_uploader("Przeciągnij plik Excel (.xlsx)", type=["xlsx"], label_visibility="collapsed")
    
    if uploaded_file:
        try:
            # Wczytujemy plik i od razu wypełniamy puste komórki (NaN) pustym ciągiem tekstowym
            raw_df = pd.read_excel(uploaded_file, header=None).fillna("")
            
            header_row_index = 0
            for idx, row in raw_df.iterrows():
                # Bezpieczna konwersja na tekst z pominięciem błędów float
                row_str = [str(val).strip().lower() for val in row.tolist()]
                if any('imię' in s or 'imie' in s for s in row_str) and any('nazwisko' in s for s in row_str):
                    header_row_index = idx
                    break
            
            df = pd.read_excel(uploaded_file, header=header_row_index).fillna("")
            df.columns = [str(c).strip().lower() for c in df.columns]
            
            if 'imie' in df.columns and 'imię' not in df.columns:
                df.rename(columns={'imie': 'imię'}, inplace=True)
                
            if 'imię' in df.columns and 'nazwisko' in df.columns and 'pion' in df.columns:
                df['pion'] = df['pion'].astype(str).str.upper().str.strip()
                df['imię'] = df['imię'].astype(str).str.strip()
                df['nazwisko'] = df['nazwisko'].astype(str).str.strip()
                df['pelne_nazwisko'] = df['imię'] + " " + df['nazwisko'] + " (" + df['pion'] + ")"
                
                # Zabezpieczenie przed usunięciem danych przy ponownym ładowaniu
                if st.session_state.db_uczestnicy is None:
                    df['liczba_wart'] = 0
                    df['ostatnia_warta'] = "-"
                    st.session_state.db_uczestnicy = df
                else:
                    st.session_state.db_uczestnicy = df.merge(
                        st.session_state.db_uczestnicy[['pelne_nazwisko', 'liczba_wart', 'ostatnia_warta']], 
                        on='pelne_nazwisko', how='left'
                    ).fillna({'liczba_wart': 0, 'ostatnia_warta': "-"})
                
                st.success("Baza załadowana pomyślnie!")
            else:
                st.error("Błąd kolumn! Wymagane nagłówki: Imię, Nazwisko, Pion.")
        except Exception as e:
            st.error(f"Błąd odczytu: {e}")

    if st.session_state.db_uczestnicy is not None:
        st.markdown("---")
        st.markdown("### 📈 Licznik Służb")
        st.dataframe(
            st.session_state.db_uczestnicy[['pelne_nazwisko', 'liczba_wart']].sort_values(by="liczba_wart"), 
            hide_index=True, use_container_width=True
        )

# --- PANEL GŁÓWNY ---
st.markdown("<div class='main-title'>KREATOR WART NOCNYCH</div>", unsafe_allow_html=True)

if st.session_state.db_uczestnicy is None:
    st.info("👋 Witaj w systemie! Aby rozpocząć planowanie obozu, wgraj plik Excel w lewym panelu bocznym.")
else:
    wybrany_dzien = st.selectbox(
        "📅 Wybierz datę odprawy:", DNI_OBOZU, 
        format_func=lambda x: f"Noc {x} / {(datetime.strptime(x+'.2026', '%d.%m.%Y') + timedelta(days=1)).strftime('%d.%02m')} (Lipiec/Sierpień)",
        label_visibility="collapsed"
    )

    if wybrany_dzien not in st.session_state.historia_wart:
        st.session_state.historia_wart[wybrany_dzien] = {godzina: [] for godzina in WARTY_SPECYFIKACJA.keys()}
    if wybrany_dzien not in st.session_state.liczba_straznikow:
        st.session_state.liczba_straznikow[wybrany_dzien] = {godzina: 2 for godzina in WARTY_SPECYFIKACJA.keys()}

    plan_dnia = st.session_state.historia_wart[wybrany_dzien]
    
    st.markdown(f"### 🛠️ Konfiguracja stanowisk na noc: **{wybrany_dzien}**")

    for godzina, info in WARTY_SPECYFIKACJA.items():
        st.markdown(f"<div class='warta-card'>", unsafe_allow_html=True)
        col_meta, col_inputs, col_actions = st.columns([1.5, 4, 1])
        
        with col_meta:
            st.markdown(f"<span style='font-size:22px; font-weight:700; color:#4ade80;'>🕒 {godzina}</span>", unsafe_allow_html=True)
            st.markdown(f"<span style='color:#a1a1aa; font-size:13px;'>{info['opis']}</span>", unsafe_allow_html=True)
            
        with col_actions:
            c1, c2 = st.columns(2)
            with c1:
                if st.button("➕", key=f"add_{godzina}_{wybrany_dzien}"):
                    st.session_state.liczba_straznikow[wybrany_dzien][godzina] += 1
                    st.rerun()
            with c2:
                if st.button("➖", key=f"sub_{godzina}_{wybrany_dzien}"):
                    if st.session_state.liczba_straznikow[wybrany_dzien][godzina] > 0:
                        st.session_state.liczba_straznikow[wybrany_dzien][godzina] -= 1
                        st.rerun()
            st.markdown(f"<div style='text-align:center; font-size:12px; margin-top:5px; color:#a1a1aa;'>Obsada: {st.session_state.liczba_straznikow[wybrany_dzien][godzina]} os.</div>", unsafe_allow_html=True)

        current_slots = st.session_state.liczba_straznikow[wybrany_dzien][godzina]
        
        while len(plan_dnia[godzina]) < current_slots:
            plan_dnia[godzina].append("")
        if len(plan_dnia[godzina]) > current_slots:
            plan_dnia[godzina] = plan_dnia[godzina][:current_slots]

        with col_inputs:
            grid_cols = st.columns(max(1, current_slots))
            for slot_idx in range(current_slots):
                with grid_cols[slot_idx]:
                    db = st.session_state.db_uczestnicy
                    db['is_pref'] = db['pion'].isin(info['preferencja'])
                    db_sorted = db.sort_values(by=['is_pref', 'liczba_wart'], ascending=[False, True])
                    
                    opcje = ["-- Wybierz --", "Wpis ręczny (Wyjątek)"] + list(
                        db_sorted['pelne_nazwisko'] + " [Warty: " + db_sorted['liczba_wart'].astype(str) + "]"
                    )
                    
                    aktualny_wybor = plan_dnia[godzina][slot_idx]
                    
                    idx_default = 0
                    is_manual = False
                    
                    if aktualny_wybor:
                        pasujace_opcje = [i for i, o in enumerate(opcje) if aktualny_wybor in o]
                        if pasujace_opcje:
                            idx_default = pasujace_opcje[0]
                        else:
                            idx_default = 1
                            is_manual = True
                            
                    wybor = st.selectbox(
                        f"Osoba {slot_idx+1}", opcje, 
                        index=idx_default, 
                        key=f"slot_{godzina}_{slot_idx}_{wybrany_dzien}",
                        label_visibility="collapsed"
                    )
                    
                    if wybor == "Wpis ręczny (Wyjątek)" or is_manual:
                        tekst_reczny = st.text_input(
                            f"Wpisz ręcznie:", 
                            value=aktualny_wybor, 
                            key=f"man_{godzina}_{slot_idx}_{wybrany_dzien}", 
                            label_visibility="collapsed"
                        )
                        plan_dnia[godzina][slot_idx] = tekst_reczny
                    elif wybor != "-- Wybierz --":
                        plan_dnia[godzina][slot_idx] = wybor.split(" [")[0]
                    else:
                        plan_dnia[godzina][slot_idx] = ""

                    if plan_dnia[godzina][slot_idx]:
                        st_name = plan_dnia[godzina][slot_idx]
                        if " (Z)" in st_name and "Z" not in info['preferencja']:
                            st.markdown("<span style='color:#ef4444; font-size:11px;'>🛑 Zuch w środku nocy!</span>", unsafe_allow_html=True)
                        elif " (Z)" not in st_name and "Z" in info['preferencja'] and not is_manual:
                            st.markdown("<span style='color:#f59e0b; font-size:11px;'>⚠️ Sugerowany Zuch</span>", unsafe_allow_html=True)

        st.markdown(f"</div>", unsafe_allow_html=True)

    if st.button("💾 ZATWIERDŹ I ZAPISZ ROZKAZ NOCNY", type="primary", use_container_width=True):
        st.session_state.historia_wart[wybrany_dzien] = plan_dnia
        st.session_state.db_uczestnicy['liczba_wart'] = 0
        
        for d, warty in st.session_state.historia_wart.items():
            for g, osoby in warty.items():
                for osoba in osoby:
                    if osoba:
                        maska = st.session_state.db_uczestnicy['pelne_nazwisko'] == osoba
                        if maska.any():
                            st.session_state.db_uczestnicy.loc[maska, 'liczba_wart'] += 1
                            st.session_state.db_uczestnicy.loc[maska, 'ostatnia_warta'] = d
        st.success("Warty zapisane pomyślnie. Statystyki zaktualizowane!")
        st.rerun()

    # --- PODGLĄD WYDRUKU ---
    st.markdown("---")
    st.markdown("### 🖨️ Cyfrowy Podgląd Wydruku A4")
    
    jutro = (datetime.strptime(wybrany_dzien+".2026", "%d.%m.%Y") + timedelta(days=1)).strftime("%d.%02m")
    
    html_print = f"""
    <div style="font-family: 'Courier New', Courier, monospace; border: 4px solid #000; padding: 40px; background-color: white; color: black; max-width: 700px; margin: 0 auto; box-shadow: 0 0 10px rgba(0,0,0,0.1);">
        <div style="text-align: center; border-bottom: 3px double #000; padding-bottom: 15px; margin-bottom: 35px;">
            <h1 style="margin: 0; font-size: 30px; text-transform: uppercase; font-weight: bold; color: black;">ROZKAZ NA WARTĘ OBOSOWĄ</h1>
            <h2 style="margin: 8px 0 0 0; font-size: 20px; color: black;">Noc: {wybrany_dzien} / {jutro} 2026 r.</h2>
        </div>
        <table style="width: 100%; border-collapse: collapse; font-size: 18px;">
            <thead>
                <tr style="border-bottom: 2px solid #000;">
                    <th style="text-align: left; padding: 12px; width: 35%; color: black;">GODZINY</th>
                    <th style="text-align: left; padding: 12px; width: 65%; color: black;">OBSADA SŁUŻBOWA</th>
                </tr>
            </thead>
            <tbody>
    """
    for g, osoby in plan_dnia.items():
        czyste_osoby = [o for o in osoby if o]
        obsada_tekst = ", ".join(czyste_osoby) if czyste_osoby else "..................................................."
        html_print += f"""
                <tr style="border-bottom: 1px dashed #000;">
                    <td style="padding: 22px 12px; font-weight: bold; color: black;">{g}</td>
                    <td style="padding: 22px 12px; color: black; line-height: 1.4;">{obsada_tekst}</td>
                </tr>
        """
    html_print += """
            </tbody>
        </table>
        <div style="text-align: center; margin-top: 40px; font-size: 14px; font-style: italic; color: black;">
            Czuwaj!<br>Komendant Obozu
        </div>
    </div>
    """
    st.components.v1.html(html_print, height=750, scrolling=True)
