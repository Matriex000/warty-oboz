import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# --- NOWOCZESNA KONFIGURACJA INTERFEJSU (CYBER-GLOW) ---
st.set_page_config(
    page_title="System Wart Obozowych", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Dynamiczne style CSS z animacjami i kolorami pionów
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700&family=Space+Grotesk:wght@400;600;700&display=swap');
        
        /* Główne style strony */
        * { font-family: 'Space Grotesk', sans-serif; }
        .stApp { 
            background: radial-gradient(circle at 50% 50%, #0f172a, #020617); 
            color: #f8fafc; 
        }
        
        /* Sidebar */
        [data-testid="stSidebar"] { 
            background-color: #0b0f19 !important; 
            border-right: 2px solid #1e293b; 
        }
        
        /* Tytuł główny z animacją błysku */
        .main-title {
            font-family: 'Orbitron', sans-serif;
            background: linear-gradient(90deg, #3b82f6, #8b5cf6, #ec4899, #3b82f6);
            background-size: 300% start;
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-size: 38px;
            font-weight: 700;
            letter-spacing: 2px;
            animation: glow 8s linear infinite;
            margin-bottom: 25px;
        }
        
        @keyframes glow {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }
        
        /* Nowoczesne, animowane karty godzinowe */
        .warta-card {
            background: rgba(30, 41, 59, 0.4);
            backdrop-filter: blur(8px);
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 15px;
            position: relative;
            overflow: hidden;
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        }
        
        .warta-card:hover {
            transform: scale(1.01) translateY(-2px);
            border-color: #6366f1;
            box-shadow: 0 10px 25px rgba(99, 102, 241, 0.15);
            background: rgba(30, 41, 59, 0.6);
        }
        
        /* Kolorowe etykiety tekstowe dla alertów */
        .text-z { color: #facc15 !important; font-weight: bold; }
        .text-h { color: #22c55e !important; font-weight: bold; }
        .text-hs { color: #3b82f6 !important; font-weight: bold; }
        .text-w { color: #ef4444 !important; font-weight: bold; }
        .text-i { color: #ffffff !important; font-weight: bold; font-shadow: 0 0 5px rgba(255,255,255,0.5); }
        
        /* Przyciski operacyjne */
        .stButton>button {
            border-radius: 6px !important;
            font-weight: 600 !important;
            transition: all 0.3s ease !important;
        }
        .stButton>button:hover {
            transform: translateY(-1px);
            box-shadow: 0 5px 12px rgba(99, 102, 241, 0.3);
        }
    </style>
""", unsafe_allow_html=True)

# --- INICJALIZACJA STANU APLIKACJI ---
if 'db_uczestnicy' not in st.session_state:
    st.session_state.db_uczestnicy = None
if 'historia_wart' not in st.session_state:
    st.session_state.historia_wart = {}
if 'liczba_straznikow' not in st.session_state:
    st.session_state.liczba_straznikow = {}

# Definicja kolejności pionów i ich kolorów tekstowych w systemie
PION_ORDER = {'Z': 0, 'H': 1, 'HS': 2, 'W': 3, 'I': 4}

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

# --- PANEL BOCZNY ---
with st.sidebar:
    st.markdown("### PANEL IMPORTU BAZY")
    uploaded_file = st.file_uploader("Wgraj plik Excel", type=["xlsx"], label_visibility="collapsed")
    
    if uploaded_file:
        try:
            raw_df = pd.read_excel(uploaded_file, header=None).fillna("")
            
            header_row_index = 0
            for idx, row in raw_df.iterrows():
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
                
                # Dodajemy wagę sortowania na podstawie słownika PION_ORDER
                df['pion_waga'] = df['pion'].map(PION_ORDER).fillna(99)
                
                df['pelne_nazwisko'] = df['imię'] + " " + df['nazwisko'] + " (" + df['pion'] + ")"
                
                if st.session_state.db_uczestnicy is None:
                    df['liczba_wart'] = 0
                    df['ostatnia_warta'] = "-"
                    st.session_state.db_uczestnicy = df
                else:
                    st.session_state.db_uczestnicy = df.merge(
                        st.session_state.db_uczestnicy[['pelne_nazwisko', 'liczba_wart', 'ostatnia_warta']], 
                        on='pelne_nazwisko', how='left'
                    ).fillna({'liczba_wart': 0, 'ostatnia_warta': "-"})
                
                st.success("Baza załadowana poprawnie")
        except Exception as e:
            st.error(f"Błąd odczytu pliku: {e}")

    if st.session_state.db_uczestnicy is not None:
        st.markdown("---")
        st.markdown("### STATYSTYKI SŁUŻB")
        # Wyświetlanie posegregowane wg pionów i liczby wart
        st.dataframe(
            st.session_state.db_uczestnicy.sort_values(by=["pion_waga", "liczba_wart"])[['pion', 'imię', 'nazwisko', 'liczba_wart']], 
            hide_index=True, use_container_width=True
        )

# --- PANEL GŁÓWNY ---
st.markdown("<div class='main-title'>SYSTEM WART OBOZOWYCH</div>", unsafe_allow_html=True)

if st.session_state.db_uczestnicy is None:
    st.info("Wgraj bazę z pliku Excel w panelu bocznym, aby uruchomić aplikację.")
else:
    wybrany_dzien = st.selectbox(
        "Wybierz datę:", DNI_OBOZU, 
        format_func=lambda x: f"Noc {x} / {(datetime.strptime(x+'.2026', '%d.%m.%Y') + timedelta(days=1)).strftime('%d.%02m')}",
        label_visibility="collapsed"
    )

    if wybrany_dzien not in st.session_state.historia_wart:
        st.session_state.historia_wart[wybrany_dzien] = {godzina: [] for godzina in WARTY_SPECYFIKACJA.keys()}
    if wybrany_dzien not in st.session_state.liczba_straznikow:
        st.session_state.liczba_straznikow[wybrany_dzien] = {godzina: 2 for godzina in WARTY_SPECYFIKACJA.keys()}

    plan_dnia = st.session_state.historia_wart[wybrany_dzien]
    
    st.markdown(f"#### Konfiguracja stanowisk na datę: {wybrany_dzien}")

    for godzina, info in WARTY_SPECYFIKACJA.items():
        st.markdown(f"<div class='warta-card'>", unsafe_allow_html=True)
        col_meta, col_inputs, col_actions = st.columns([2, 4, 1])
        
        with col_meta:
            st.markdown(f"<div style='font-size:22px; font-weight:700; color:#6366f1; font-family:Orbitron;'>{godzina}</div>", unsafe_allow_html=True)
            st.markdown(f"<div style='color:#94a3b8; font-size:13px;'>{info['opis']}</div>", unsafe_allow_html=True)
            
        with col_actions:
            c1, c2 = st.columns(2)
            with c1:
                if st.button("+", key=f"add_{godzina}_{wybrany_dzien}"):
                    st.session_state.liczba_straznikow[wybrany_dzien][godzina] += 1
                    st.rerun()
            with c2:
                if st.button("-", key=f"sub_{godzina}_{wybrany_dzien}"):
                    if st.session_state.liczba_straznikow[wybrany_dzien][godzina] > 0:
                        st.session_state.liczba_straznikow[wybrany_dzien][godzina] -= 1
                        st.rerun()
            st.markdown(f"<div style='text-align:center; font-size:11px; margin-top:4px; color:#64748b;'>Obsada: {st.session_state.liczba_straznikow[wybrany_dzien][godzina]} os.</div>", unsafe_allow_html=True)

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
                    
                    # --- KLUCZOWE SORTOWANIE: NAJPIERW PIONY (Z, H, HS, W, I), POTEM LICZBA WART ---
                    db_sorted = db.sort_values(by=['pion_waga', 'liczba_wart'], ascending=[True, True])
                    
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
                            f"Ręcznie:", value=aktualny_wybor, 
                            key=f"man_{godzina}_{slot_idx}_{wybrany_dzien}", label_visibility="collapsed"
                        )
                        plan_dnia[godzina][slot_idx] = tekst_reczny
                    elif wybor != "-- Wybierz --":
                        plan_dnia[godzina][slot_idx] = wybor.split(" [")[0]
                    else:
                        plan_dnia[godzina][slot_idx] = ""

                    # Dynamiczne kolorowanie tekstu informacji zwrotnej na podstawie wybranego pionu
                    if plan_dnia[godzina][slot_idx]:
                        st_name = plan_dnia[godzina][slot_idx]
                        if " (Z)" in st_name:
                            if "Z" not in info['preferencja']:
                                st.markdown("<span class='text-z' style='font-size:11px;'>🛑 Zuch po północy!</span>", unsafe_allow_html=True)
                            else:
                                st.markdown("<span class='text-z' style='font-size:11px;'>Pion: Zuch</span>", unsafe_allow_html=True)
                        elif " (H)" in st_name:
                            st.markdown("<span class='text-h' style='font-size:11px;'>Pion: Harcerz</span>", unsafe_allow_html=True)
                        elif " (HS)" in st_name:
                            st.markdown("<span class='text-hs' style='font-size:11px;'>Pion: Harcerz St.</span>", unsafe_allow_html=True)
                        elif " (W)" in st_name:
                            st.markdown("<span class='text-w' style='font-size:11px;'>Pion: Wędrownik</span>", unsafe_allow_html=True)
                        elif " (I)" in st_name:
                            st.markdown("<span class='text-i' style='font-size:11px;'>Pion: Instruktor</span>", unsafe_allow_html=True)

        st.markdown(f"</div>", unsafe_allow_html=True)

    if st.button("ZAPISZ I ZAKTUALIZUJ STATYSTYKI", type="primary", use_container_width=True):
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
        st.success("Plan zapisany w chmurze.")
        st.rerun()

    # --- PRAWIDŁOWY CZARNO-BIAŁY WYDRUK A4 ---
    st.markdown("---")
    st.markdown("#### Podgląd arkusza do wydruku (Format A4)")
    
    jutro = (datetime.strptime(wybrany_dzien+".2026", "%d.%m.%Y") + timedelta(days=1)).strftime("%d.%02m")
    
    html_print = f"""
    <div style="font-family: 'Courier New', Courier, monospace; border: 4px solid #000; padding: 40px; background-color: white; color: black; max-width: 700px; margin: 0 auto;">
        <div style="text-align: center; border-bottom: 3px double #000; padding-bottom: 15px; margin-bottom: 35px;">
            <h1 style="margin: 0; font-size: 30px; text-transform: uppercase; font-weight: bold; color: black;">ROZKAZ NA WARTĘ OBOZOWĄ</h1>
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
            Czuwaj!<br>Oboźny Obozu
        </div>
    </div>
    """
    st.components.v1.html(html_print, height=750, scrolling=True)
