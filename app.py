import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# --- NOWOCZESNA KONFIGURACJA INTERFEJSU (CYBER-GLOW) ---
st.set_page_config(
    page_title="System Wart Obozowych Pro", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Dynamiczne style CSS z animacjami i neonowymi akcentami
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700&family=Space+Grotesk:wght@400;600;700&display=swap');
        
        * { font-family: 'Space Grotesk', sans-serif; }
        .stApp { 
            background: radial-gradient(circle at 50% 50%, #0f172a, #020617); 
            color: #f8fafc; 
        }
        
        [data-testid="stSidebar"] { 
            background-color: #0b0f19 !important; 
            border-right: 2px solid #1e293b; 
        }
        
        .main-title {
            font-family: 'Orbitron', sans-serif;
            background: linear-gradient(90deg, #3b82f6, #8b5cf6, #ec4899, #3b82f6);
            background-size: 300% auto;
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-size: 38px;
            font-weight: 700;
            letter-spacing: 2px;
            animation: glow 8s linear infinite;
            margin-bottom: 5px;
        }
        
        @keyframes glow {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }
        
        .warta-card {
            background: rgba(30, 41, 59, 0.4);
            backdrop-filter: blur(8px);
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 15px;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }
        
        .warta-card:hover {
            border-color: #6366f1;
            box-shadow: 0 10px 25px rgba(99, 102, 241, 0.15);
            background: rgba(30, 41, 59, 0.6);
        }
        
        /* Kolorowe etykiety tekstowe dla alertów i pionów */
        .badge-z { color: #facc15 !important; font-weight: bold; }
        .badge-h { color: #22c55e !important; font-weight: bold; }
        .badge-hs { color: #3b82f6 !important; font-weight: bold; }
        .badge-w { color: #ef4444 !important; font-weight: bold; }
        .badge-i { color: #ffffff !important; font-weight: bold; }
        
        .metric-card {
            background: rgba(15, 23, 42, 0.6);
            border: 1px solid #1e293b;
            border-radius: 10px;
            padding: 15px;
            text-align: center;
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

# Słowniki pomocnicze dla pionów
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

# --- PANEL GŁÓWNY Z ZAKŁADKAMI ---
st.markdown("<div class='main-title'>SYSTEM WART OBOZOWYCH</div>", unsafe_allow_html=True)

if st.session_state.db_uczestnicy is None:
    st.info("Wgraj bazę z pliku Excel w panelu bocznym, aby uruchomić aplikację.")
else:
    # Tworzenie nowoczesnych kart (Tabs)
    tab_kreator, tab_statystyki = st.tabs(["📝 Kreator Rozkazu", "📊 Statystyki i Analiza"])
    
    # ==========================================
    # ZAKŁADKA 1: KREATOR ROZKAZU
    # ==========================================
    with tab_kreator:
        wybrany_dzien = st.selectbox(
            "Wybierz datę:", DNI_OBOZU, 
            format_func=lambda x: f"Noc {x} / {(datetime.strptime(x+'.2026', '%d.%m.%Y') + timedelta(days=1)).strftime('%d.%02m')}"
        )

        # Ustalenie wczorajszego dnia, by sprawdzić zmęczenie
        idx_dzis = DNI_OBOZU.index(wybrany_dzien)
        wczorajszy_dzien = DNI_OBOZU[idx_dzis - 1] if idx_dzis > 0 else None

        if wybrany_dzien not in st.session_state.historia_wart:
            st.session_state.historia_wart[wybrany_dzien] = {godzina: [] for godzina in WARTY_SPECYFIKACJA.keys()}
        if wybrany_dzien not in st.session_state.liczba_straznikow:
            st.session_state.liczba_straznikow[wybrany_dzien] = {godzina: 2 for godzina in WARTY_SPECYFIKACJA.keys()}

        plan_dnia = st.session_state.historia_wart[wybrany_dzien]

        for godzina, info in WARTY_SPECYFIKACJA.items():
            st.markdown(f"<div class='warta-card'>", unsafe_allow_html=True)
            col_meta, col_inputs, col_actions = st.columns([2, 4, 1])
            
            with col_meta:
                st.markdown(f"<div style='font-size:20px; font-weight:700; color:#6366f1; font-family:Orbitron;'>{godzina}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='color:#94a3b8; font-size:12px; margin-bottom:8px;'>{info['opis']}</div>", unsafe_allow_html=True)
                
                # --- USPRAWNIENIE: Szybki filtr pionu per godzina ---
                filtr_pionu = st.multiselect(
                    "Filtruj listę do pionów:", ["Z", "H", "HS", "W", "I"], 
                    default=info['preferencja'], key=f"filter_{godzina}_{wybrany_dzien}"
                )
                
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
                st.markdown(f"<div style='text-align:center; font-size:11px; color:#64748b;'>Obsada: {st.session_state.liczba_straznikow[wybrany_dzien][godzina]} os.</div>", unsafe_allow_html=True)

            current_slots = st.session_state.liczba_straznikow[wybrany_dzien][godzina]
            while len(plan_dnia[godzina]) < current_slots: plan_dnia[godzina].append("")
            if len(plan_dnia[godzina]) > current_slots: plan_dnia[godzina] = plan_dnia[godzina][:current_slots]

            with col_inputs:
                grid_cols = st.columns(max(1, current_slots))
                for slot_idx in range(current_slots):
                    with grid_cols[slot_idx]:
                        db = st.session_state.db_uczestnicy
                        
                        # Filtrowanie bazy na podstawie wybranego widoku pionów
                        if filtr_pionu:
                            db_filtered = db[db['pion'].isin(filtr_pionu)]
                        else:
                            db_filtered = db
                            
                        db_sorted = db_filtered.sort_values(by=['pion_waga', 'liczba_wart'], ascending=[True, True])
                        
                        # --- KOLOROWANIE W LIŚCIE ROZWIJANEJ ZA POMOCĄ EMOJI PIONÓW ---
                        opcje = ["-- Wybierz --", "Wpis ręczny (Wyjątek)"]
                        for _, row in db_sorted.iterrows():
                            emoji = PION_COLORS.get(row['pion'], '▪️')
                            opcje.append(f"{emoji} {row['pelne_nazwisko']} [Warty: {row['liczba_wart']}]")
                        
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
                            f"Osoba {slot_idx+1}", opcje, index=idx_default, 
                            key=f"slot_{godzina}_{slot_idx}_{wybrany_dzien}", label_visibility="collapsed"
                        )
                        
                        if wybor == "Wpis ręczny (Wyjątek)" or is_manual:
                            tekst_reczny = st.text_input(f"Ręcznie:", value=aktualny_wybor, key=f"man_{godzina}_{slot_idx}_{wybrany_dzien}", label_visibility="collapsed")
                            plan_dnia[godzina][slot_idx] = tekst_reczny
                        elif wybor != "-- Wybierz --":
                            # Wyciągamy czyste imię i nazwisko bez emoji i licznika wart
                            czysty_tekst = wybor.split(" (")[0][2:] + " (" + wybor.split(" (")[1].split(")")[0] + ")"
                            plan_dnia[godzina][slot_idx] = czysty_tekst
                        else:
                            plan_dnia[godzina][slot_idx] = ""

                        # --- LOGIKA WALIDACJI BEZPIECZEŃSTWA OBOZOWEGO ---
                        if plan_dnia[godzina][slot_idx]:
                            st_name = plan_dnia[godzina][slot_idx]
                            
                            # 1. Alerty przynależności do pionu
                            if " (Z)" in st_name:
                                st.markdown("<span class='badge-z' style='font-size:11px;'>🟡 Pion: Zuch</span>", unsafe_allow_html=True)
                                if "Z" not in info['preferencja']:
                                    st.markdown("<span style='color:#ef4444; font-size:11px; font-weight:bold;'>🛑 Zuch w środku nocy!</span>", unsafe_allow_html=True)
                            elif " (H)" in st_name: st.markdown("<span class='badge-h' style='font-size:11px;'>🟢 Pion: Harcerz</span>", unsafe_allow_html=True)
                            elif " (HS)" in st_name: st.markdown("<span class='badge-hs' style='font-size:11px;'>🔵 Pion: Harcerz St.</span>", unsafe_allow_html=True)
                            elif " (W)" in st_name: st.markdown("<span class='badge-w' style='font-size:11px;'>🔴 Pion: Wędrownik</span>", unsafe_allow_html=True)
                            elif " (I)" in st_name: st.markdown("<span class='badge-i' style='font-size:11px;'>⚪ Pion: Instruktor</span>", unsafe_allow_html=True)
                            
                            # 2. USPRAWNIENIE: Sprawdzanie zmęczenia (czy osoba stała wczoraj)
                            if wczorajszy_dzien and wczorajszy_dzien in st.session_state.historia_wart:
                                wczorajsze_osoby = []
                                for g_wczoraj in st.session_state.historia_wart[wczorajszy_dzien].values():
                                    wczorajsze_osoby.extend(g_wczoraj)
                                if st_name in wczorajsze_osoby:
                                    st.markdown("<span style='color:#ec4899; font-size:11px; font-weight:bold;'>⚠️ Stał(a) poprzedniej nocy!</span>", unsafe_allow_html=True)

            st.markdown(f"</div>", unsafe_allow_html=True)

        if st.button("ZAPISZ I ZAKTUALIZUJ STATYSTYKI", type="primary", use_container_width=True):
            st.session_state.historia_wart[wybrany_dzien] = plan_dnia
            st.session_state.db_uczestnicy['liczba_wart'] = 0
            st.session_state.db_uczestnicy['ostatnia_warta'] = "-"
            
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

        # PODGLĄD WYDRUKU
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
        html_print += f"""
                </tbody>
            </table>
            <div style="text-align: center; margin-top: 40px; font-size: 14px; font-style: italic; color: black;">
                Czuwaj!<br>Oboźny Obozu
            </div>
        </div>
        """
        st.components.v1.html(html_print, height=750, scrolling=True)

    # ==========================================
    # ZAKŁADKA 2: STATYSTYKI I ANALIZA
    # ==========================================
    with tab_statystyki:
        st.markdown("#### 📈 Kluczowe Statystyki Akcji Letniej")
        
        db_stat = st.session_state.db_uczestnicy
        total_warty = db_stat['liczba_wart'].sum()
        srednia = db_stat['liczba_wart'].mean()
        
        # 3 kafelki podsumowujące
        c_tot, c_avg, c_unassigned = st.columns(3)
        with c_tot:
            st.markdown(f"<div class='metric-card'><h5 style='margin:0; color:#94a3b8;'>Suma wszystkich obsadzonych wart</h5><h2 style='margin:10px 0 0 0; color:#3b82f6; font-family:Orbitron;'>{total_warty}</h2></div>", unsafe_allow_html=True)
        with c_avg:
            st.markdown(f"<div class='metric-card'><h5 style='margin:0; color:#94a3b8;'>Średnia liczba służb na osobę</h5><h2 style='margin:10px 0 0 0; color:#22c55e; font-family:Orbitron;'>{srednia:.1f}</h2></div>", unsafe_allow_html=True)
        with c_unassigned:
            bez_warty = len(db_stat[db_stat['liczba_wart'] == 0])
            st.markdown(f"<div class='metric-card'><h5 style='margin:0; color:#94a3b8;'>Osoby, które jeszcze NIE stały</h5><h2 style='margin:10px 0 0 0; color:#ec4899; font-family:Orbitron;'>{bez_warty}</h2></div>", unsafe_allow_html=True)
            
        st.markdown("---")
        
        col_tables_1, col_tables_2 = st.columns(2)
        
        with col_tables_1:
            st.markdown("##### 🚀 Kto odpoczywa? (Najmniej wart w obozie)")
            st.dataframe(
                db_stat.sort_values(by="liczba_wart", ascending=True)[['pion', 'imię', 'nazwisko', 'liczba_wart']].head(10),
                hide_index=True, use_container_width=True
            )
            
        with col_tables_2:
            st.markdown("##### ⚔️ Obozowi Weterani (Najwięcej wart w obozie)")
            st.dataframe(
                db_stat.sort_values(by="liczba_wart", ascending=False)[['pion', 'imię', 'nazwisko', 'liczba_wart']].head(10),
                hide_index=True, use_container_width=True
            )
            
        st.markdown("---")
        st.markdown("##### 📊 Procentowe obciążenie służbami ze względu na Piony")
        
        pion_stats = db_stat.groupby('pion').agg(
            Liczba_Uczestników=('pion', 'count'),
            Suma_Wart=('liczba_wart', 'sum')
        ).reset_index()
        
        st.dataframe(pion_stats, hide_index=True, use_container_width=True)
