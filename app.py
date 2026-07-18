import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

# --- CONFIG INTERFEJSU ---
st.set_page_config(page_title="System Wart Obozowych Pro", layout="wide", initial_sidebar_state="expanded")

# Styl aplikacji pozostaje nowoczesny i ciemny, natomiast styl @media print wymusza w 100% czysty, pierwotny wydruk
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700&family=Space+Grotesk:wght@400;600;700&display=swap');
        * { font-family: 'Space Grotesk', sans-serif; }
        .stApp { background: radial-gradient(circle at 50% 50%, #0f172a, #020617); color: #f8fafc; }
        [data-testid="stSidebar"] { background-color: #0b0f19 !important; border-right: 2px solid #1e293b; }
        .main-title {
            font-family: 'Orbitron', sans-serif;
            background: linear-gradient(90deg, #3b82f6, #8b5cf6, #ec4899, #3b82f6);
            background-size: 300% auto;
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-size: 38px; font-weight: 700; letter-spacing: 2px;
            animation: glow 8s linear infinite; margin-bottom: 5px;
        }
        @keyframes glow { 0% { background-position: 0% 50%; } 50% { background-position: 100% 50%; } 100% { background-position: 0% 50%; } }
        .warta-card {
            background: rgba(30, 41, 59, 0.4); backdrop-filter: blur(8px);
            border: 1px solid #334155; border-radius: 12px; padding: 20px; margin-bottom: 15px;
        }
        .badge-z { color: #facc15 !important; font-weight: bold; }
        .badge-h { color: #22c55e !important; font-weight: bold; }
        .badge-hs { color: #3b82f6 !important; font-weight: bold; }
        .badge-w { color: #ef4444 !important; font-weight: bold; }
        .badge-i { color: #ffffff !important; font-weight: bold; }
        .metric-card { background: rgba(15, 23, 42, 0.6); border: 1px solid #1e293b; border-radius: 10px; padding: 15px; text-align: center; }
        
        /* --- CAŁKOWITE PRZYWRÓCENIE PIERWOTNEGO WYGLĄDU DLA WYDRUKU PDF --- */
        @media print {
            /* Ukrycie wszystkich elementów systemowych, kontenerów i ciemnych teł Streamlita */
            html, body, .stApp, [data-testid="stReportBlock"], div { 
                background: white !important; 
                color: black !important; 
                box-shadow: none !important;
                border: none !important;
                padding: 0 !important;
                margin: 0 !important;
            }
            
            /* Ukrycie elementów interfejsu aplikacji */
            [data-testid="stSidebar"], .stTabs, button, .main-title, header, hr, .stMarkdown { 
                display: none !important; 
            }
            
            /* Wymuszenie widoczności wyłącznie czystego obszaru druku */
            body * { visibility: hidden; }
            .print-area, .print-area * { visibility: visible; }
            
            .print-area {
                position: absolute; 
                left: 0; 
                top: 0; 
                width: 100%;
                background: white !important; 
                color: black !important;
                padding: 10px !important; 
                font-family: 'Courier New', monospace !important;
            }
            
            table { 
                width: 100% !important; 
                border-collapse: collapse !important; 
                background: white !important;
                color: black !important;
            }
            
            tr { 
                border-bottom: 1px solid black !important; 
                background: white !important;
            }
            
            td { 
                color: black !important; 
                background: white !important;
                padding: 12px !important;
            }
        }
    </style>
""", unsafe_allow_html=True)

# --- INICJALIZACJA STANU ---
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

# --- PANEL BOCZNY ---
with st.sidebar:
    st.markdown("### PANEL IMPORTU BAZY")
    uploaded_file = st.file_uploader("Wgraj plik Excel", type=["xlsx"], label_visibility="collapsed")
    
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
            col_ostatnia = znajdz_kolumne(['ostatnia_warta', 'kiedy'], 'Ostatnia_Warta')

            for c, orig in [('Imię', col_imie), ('Nazwisko', col_nazwisko), ('Pion', col_pion), ('Drużyna', col_druzyna)]:
                if orig in df.columns: df[c] = df[orig]
                else: df[c] = ""
            
            df['Liczba_Wart'] = pd.to_numeric(df[col_liczba], errors='coerce').fillna(0).astype(int) if col_liczba in df.columns else 0
            df['Ostatnia_Warta'] = df[col_ostatnia] if col_ostatnia in df.columns else "-"
            
            df['Pion'] = df['Pion'].astype(str).str.upper().str.strip()
            df['pion_waga'] = df['Pion'].map(PION_ORDER).fillna(99)
            df['pelne_nazwisko'] = df['Imię'].astype(str) + " " + df['Nazwisko'].astype(str) + " (" + df['Pion'] + ")"
            
            st.session_state.db_uczestnicy = df[['Imię', 'Nazwisko', 'Pion', 'Drużyna', 'Liczba_Wart', 'Ostatnia_Warta', 'pion_waga', 'pelne_nazwisko']]
            st.success("Baza załadowana pomyślnie!")
        except Exception as e:
            st.error(f"Błąd struktury pliku: {e}")

    if st.session_state.db_uczestnicy is not None:
        st.markdown("---")
        st.markdown("### AKTYWNA LISTA OBOZOWA")
        st.dataframe(
            st.session_state.db_uczestnicy.sort_values(by=["pion_waga", "Liczba_Wart"])[['Pion', 'Imię', 'Nazwisko', 'Drużyna', 'Liczba_Wart']], 
            hide_index=True, use_container_width=True
        )

# --- PANEL GŁÓWNY ---
st.markdown("<div class='main-title'>SYSTEM WART OBOZOWYCH</div>", unsafe_allow_html=True)

if st.session_state.db_uczestnicy is None:
    st.info("Wgraj bazę z pliku Excel w panelu bocznym, aby uruchomić aplikację.")
else:
    tab_kreator, tab_statystyki = st.tabs(["📝 Kreator Rozkazu", "📊 Statystyki i Eksport"])
    
    with tab_kreator:
        wybrany_dzien = st.selectbox("Wybierz datę:", DNI_OBOZU, format_func=lambda x: f"Noc {x} / {(datetime.strptime(x+'.2026', '%d.%m.%Y') + timedelta(days=1)).strftime('%d.%02m')}")

        idx_dzis = DNI_OBOZU.index(wybrany_dzien)
        wczorajszy_dzien = DNI_OBOZU[idx_dzis - 1] if idx_dzis > 0 else None

        if wybrany_dzien not in st.session_state.historia_wart: st.session_state.historia_wart[wybrany_dzien] = {godzina: [] for godzina in WARTY_SPECYFIKACJA.keys()}
        if wybrany_dzien not in st.session_state.liczba_straznikow: st.session_state.liczba_straznikow[wybrany_dzien] = {godzina: 2 for godzina in WARTY_SPECYFIKACJA.keys()}
        if wybrany_dzien not in st.session_state.lokalizacje_wart: st.session_state.lokalizacje_wart[wybrany_dzien] = {godzina: [] for godzina in WARTY_SPECYFIKACJA.keys()}

        plan_dnia = st.session_state.historia_wart[wybrany_dzien]
        lokalizacje_dnia = st.session_state.lokalizacje_wart[wybrany_dzien]

        for godzina, info in WARTY_SPECYFIKACJA.items():
            st.markdown(f"<div class='warta-card'>", unsafe_allow_html=True)
            col_meta, col_inputs, col_actions = st.columns([2, 4, 1])
            
            with col_meta:
                st.markdown(f"<div style='font-size:20px; font-weight:700; color:#6366f1; font-family:Orbitron;'>{godzina}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='color:#94a3b8; font-size:12px; margin-bottom:8px;'>{info['opis']}</div>", unsafe_allow_html=True)
                filtr_pionu = st.multiselect("Filtruj piony:", ["Z", "H", "HS", "W", "I"], default=info['preferencja'], key=f"filter_{godzina}_{wybrany_dzien}")
                
            with col_actions:
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("+", key=f"add_{godzina}_{wybrany_dzien}"): st.session_state.liczba_straznikow[wybrany_dzien][godzina] += 1; st.rerun()
                with c2:
                    if st.button("-", key=f"sub_{godzina}_{wybrany_dzien}"):
                        if st.session_state.liczba_straznikow[wybrany_dzien][godzina] > 0: st.session_state.liczba_straznikow[wybrany_dzien][godzina] -= 1; st.rerun()

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
                        
                        opcje = ["-- Wybierz --", "Wpis ręczny (Wyjątek)"]
                        for _, row in db_sorted.iterrows():
                            emoji = PION_COLORS.get(row['Pion'], '▪️')
                            druzyna_str = f" [{row['Drużyna']}]" if row['Drużyna'] else ""
                            opcje.append(f"{emoji} {row['pelne_nazwisko']}{druzyna_str} (Warty: {row['Liczba_Wart']})")
                        
                        aktualny_wybor = plan_dnia[godzina][slot_idx]
                        idx_default = 0
                        is_manual = False
                        
                        if aktualny_wybor:
                            pasujace_opcje = [i for i, o in enumerate(opcje) if aktualny_wybor in o]
                            if pasujace_opcje: idx_default = pasujace_opcje[0]
                            else: idx_default = 1; is_manual = True
                                
                        wybor = st.selectbox(f"Osoba {slot_idx+1}", opcje, index=idx_default, key=f"sl_{godzina}_{slot_idx}_{wybrany_dzien}", label_visibility="collapsed")
                        
                        if wybor == "Wpis ręczny (Wyjątek)" or is_manual:
                            plan_dnia[godzina][slot_idx] = st.text_input(f"Ręcznie:", value=aktualny_wybor, key=f"mn_{godzina}_{slot_idx}_{wybrany_dzien}", label_visibility="collapsed")
                        elif wybor != "-- Wybierz --":
                            plan_dnia[godzina][slot_idx] = wybor.split(" (Warty")[0][2:].split(" [")[0]
                        else:
                            plan_dnia[godzina][slot_idx] = ""

                        lokalizacje_dnia[godzina][slot_idx] = st.text_input(
                            "Miejsce warty:", value=lokalizacje_dnia[godzina][slot_idx], 
                            key=f"loc_{godzina}_{slot_idx}_{wybrany_dzien}", placeholder="np. Brama"
                        )

                        if plan_dnia[godzina][slot_idx]:
                            st_name = plan_dnia[godzina][slot_idx]
                            m_druzyna = db[db['pelne_nazwisko'] == st_name]['Drużyna'].values
                            if len(m_druzyna) > 0 and m_druzyna[0]:
                                st.markdown(f"<div style='font-size:11px; color:#a855f7;'>Drużyna: {m_druzyna[0]}</div>", unsafe_allow_html=True)
                            
                            if " (Z)" in st_name: st.markdown("<span class='badge-z' style='font-size:11px;'>🟡 Pion: Zuch</span>", unsafe_allow_html=True)
                            elif " (H)" in st_name: st.markdown("<span class='badge-h' style='font-size:11px;'>🟢 Pion: Harcerz</span>", unsafe_allow_html=True)
                            elif " (HS)" in st_name: st.markdown("<span class='badge-hs' style='font-size:11px;'>🔵 Pion: Harcerz St.</span>", unsafe_allow_html=True)
                            elif " (W)" in st_name: st.markdown("<span class='badge-w' style='font-size:11px;'>🔴 Pion: Wędrownik</span>", unsafe_allow_html=True)
                            elif " (I)" in st_name: st.markdown("<span class='badge-i' style='font-size:11px;'>⚪ Pion: Instruktor</span>", unsafe_allow_html=True)

                            if wczorajszy_dzien and wczorajszy_dzien in st.session_state.historia_wart:
                                if st_name in [item for sublist in st.session_state.historia_wart[wczorajszy_dzien].values() for item in sublist]:
                                    st.markdown("<span style='color:#ec4899; font-size:11px; font-weight:bold;'>⚠️ Stał(a) poprzedniej nocy!</span>", unsafe_allow_html=True)
            st.markdown(f"</div>", unsafe_allow_html=True)

        if st.button("ZAPISZ I ZAKTUALIZUJ STATYSTYKI W EXCELU", type="primary", use_container_width=True):
            st.session_state.historia_wart[wybrany_dzien] = plan_dnia
            st.session_state.lokalizacje_wart[wybrany_dzien] = lokalizacje_dnia
            
            st.session_state.db_uczestnicy['Liczba_Wart'] = 0
            st.session_state.db_uczestnicy['Ostatnia_Warta'] = "-"
            
            for d, warty in st.session_state.historia_wart.items():
                for g, osoby in warty.items():
                    for osoba in osoby:
                        if osoba:
                            maska = st.session_state.db_uczestnicy['pelne_nazwisko'] == osoba
                            if maska.any():
                                st.session_state.db_uczestnicy.loc[maska, 'Liczba_Wart'] += 1
                                st.session_state.db_uczestnicy.loc[maska, 'Ostatnia_Warta'] = d
            st.success("Zaktualizowano dane w systemie!")
            st.rerun()

        st.markdown("---")
        st.markdown("### 🖨️ Podgląd Rozkazu przed Wydrukiem")

        # Czysta budowa tabeli na oryginalny wzór
        straznicy_tabela = ""
        for g, osoby in plan_dnia.items():
            straznicy_str = []
            for i, o in enumerate(osoby):
                miejsce = lokalizacje_dnia[g][i] if i < len(lokalizacje_dnia[g]) else ""
                m_str = f" [MIEJSCE: {miejsce.upper()}]" if miejsce else ""
                osoba_str = o if o else "[Brak przypisania]"
                straznicy_str.append(f"{osoba_str}{m_str}")
            
            straznicy_html = "<br>".join(straznicy_str)
            straznicy_tabela += f"<tr style='border-bottom: 1px solid black;'><td style='padding: 12px; font-weight: bold; font-size: 16px; vertical-align: top; width: 30%; color: black !important; background: white !important;'>{g}</td><td style='padding: 12px; font-size: 16px; color: black !important; background: white !important;'>{straznicy_html}</td></tr>"

        kod_html_druku = f"""
        <div class="print-area" style="background-color: white !important; color: black !important; padding: 30px; border-radius: 0px; border: none; font-family: 'Courier New', monospace;">
            <h2 style="text-align: center; margin-bottom: 5px; font-weight: bold; color: black !important; letter-spacing: 2px;">ROZKAZ WART OBOZOWYCH</h2>
            <p style="text-align: center; margin-top: 0; font-size: 14px; color: black !important;">Noc: {wybrany_dzien} | Wygenerowano: {datetime.now().strftime('%d.%m.%Y')}</p>
            <hr style="border: 1px solid black; margin-bottom: 20px;">
            <table style="width: 100%; border-collapse: collapse; color: black !important; background: white !important;">
                <tbody>
                    {straznicy_tabela}
                </tbody>
            </table>
            <br><br><br>
            <p style="text-align: right; font-weight: bold; margin-right: 20px; color: black !important; background: white !important;">Podpisał: Komendant Obozu</p>
        </div>
        """
        
        st.html(kod_html_druku)
        st.markdown("<br>", unsafe_allow_html=True)

        if st.button("🖨️ URUCHOM DRUKOWANIE SYSTEMOWE (ZAPISZ JAKO PDF)", type="secondary", use_container_width=True):
            st.html("<script>window.print();</script>")

    with tab_statystyki:
        st.markdown("#### 📊 Podsumowanie Obciążeń Służbami")
        db_stat = st.session_state.db_uczestnicy
        
        c_tot, c_avg, c_un = st.columns(3)
        c_tot.markdown(f"<div class='metric-card'><h5>Suma wart</h5><h2>{db_stat['Liczba_Wart'].sum()}</h2></div>", unsafe_allow_html=True)
        c_avg.markdown(f"<div class='metric-card'><h5>Średnia na osobę</h5><h2>{db_stat['Liczba_Wart'].mean():.1f}</h2></div>", unsafe_allow_html=True)
        c_un.markdown(f"<div class='metric-card'><h5>Jeszcze nie stali</h5><h2>{len(db_stat[db_stat['Liczba_Wart'] == 0])}</h2></div>", unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("#### 📥 GENERATOR RAPORTU KOŃCOWEGO")
        
        export_df = db_stat[['Imię', 'Nazwisko', 'Pion', 'Drużyna', 'Liczba_Wart', 'Ostatnia_Warta']].copy()
        export_df['Czy_Był_Na_Warcie'] = export_df['Liczba_Wart'].apply(lambda x: "TAK" if x > 0 else "NIE")
        
        szczegoly_list = []
        for idx, row in db_stat.iterrows():
            pelne = row['pelne_nazwisko']
            historia_osoby = []
            for d, warty in st.session_state.historia_wart.items():
                for g, osoby in warty.items():
                    if pelne in osoby:
                        slot_idx = osoby.index(pelne)
                        miejsce = st.session_state.lokalizacje_wart[d][g][slot_idx] if slot_idx < len(st.session_state.lokalizacje_wart[d][g]) else ""
                        miejsce_str = f" -> {miejsce}" if miejsce else ""
                        historia_osoby.append(f"Noc {d} ({g}{miejsce_str})")
            szczegoly_list.append("; ".join(historia_osoby) if historia_osoby else "-")
            
        export_df['Dokładna_Historia_Służb'] = szczegoly_list

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            export_df.to_excel(writer, index=False, sheet_name='Raport Wart')
        
        st.download_button(
            label="POBIERZ ARKUSZ KOŃCOWY EXCEL (.XLSX)",
            data=buffer.getvalue(),
            file_name=f"Raport_Wart_Obóz_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        st.dataframe(export_df, hide_index=True, use_container_width=True)
