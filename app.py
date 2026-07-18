import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

# --- CONFIG INTERFEJSU ---
st.set_page_config(page_title="System Wart Obozowych Pro", page_icon="⛺", layout="wide", initial_sidebar_state="expanded")

# Zintegrowany, nowoczesny wygląd aplikacji z bezpiecznym dla wydruków PDF blokiem podglądu (.print-preview)
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
        
        /* KLASYCZNY WYGLĄD MASZYNOPISU Z PODWÓJNĄ RAMKĄ POD DRUK A4 */
        .print-preview {
            font-family: 'Courier New', Courier, monospace; 
            border: 4px double #000000; 
            padding: 40px; 
            background-color: #ffffff !important; 
            color: #000000 !important; 
            max-width: 650px; 
            margin: 20px auto;
            box-shadow: 0 4px 15px rgba(0,0,0,0.5);
        }
        .print-preview table { width: 100%; border-collapse: collapse; font-size: 18px; color: #000000 !important; }
        .print-preview th { text-align: left; padding: 10px; border-bottom: 2px solid #000000; color: #000000 !important; }
        .print-preview td { padding: 15px 10px; color: #000000 !important; }
        .print-preview tr { border-bottom: 1px dashed #666666; }

        /* Reguły wymuszające idealny czysty wydruk karty A4 bez ciemnego tła Streamlit */
        @media print {
            html, body, .stApp, [data-testid="stReportBlock"], div { 
                background: white !important; 
                color: black !important; 
                box-shadow: none !important;
                border: none !important;
                padding: 0 !important;
                margin: 0 !important;
            }
            [data-testid="stSidebar"], .stTabs, button, .main-title, header, hr, .stMarkdown, .stInfo, .stCaption { 
                display: none !important; 
            }
            body * { visibility: hidden; }
            .print-preview, .print-preview * { 
                visibility: visible; 
            }
            .print-preview {
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

# --- INICJALIZACJA STANU APLIKACJI ---
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

# --- PANEL BOCZNY: AUTODETEKCJA I IMPORT ---
with st.sidebar:
    st.markdown("### 📥 PANEL IMPORTU BAZY")
    uploaded_file = st.file_uploader("Wgraj plik Excel (.xlsx)", type=["xlsx"], label_visibility="collapsed")
    
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
        st.markdown("### 📊 AKTYWNA LISTA OBOZOWA")
        st.dataframe(
            st.session_state.db_uczestnicy.sort_values(by=["pion_waga", "Liczba_Wart"])[['Pion', 'Imię', 'Nazwisko', 'Drużyna', 'Liczba_Wart']], 
            hide_index=True, use_container_width=True
        )

# --- PANEL GŁÓWNY ---
st.markdown("<div class='main-title'>⛺ SYSTEM WART OBOZOWYCH</div>", unsafe_allow_html=True)

if st.session_state.db_uczestnicy is None:
    st.info("👋 Wgraj bazę z pliku Excel w panelu bocznym, aby uruchomić kreator i generowanie rozkazów A4.")
else:
    # Układ 3 funkcjonalnych zakładek
    tab_kreator, tab_druk, tab_statystyki = st.tabs(["📝 Kreator Rozkazu", "🖨️ Podgląd i Druk A4", "📊 Statystyki Globalne"])
    
    wybrany_dzien = st.selectbox("📅 Wybierz datę planowanej służby:", DNI_OBOZU, format_func=lambda x: f"Noc {x} / {(datetime.strptime(x+'.2026', '%d.%m.%Y') + timedelta(days=1)).strftime('%d.%02m')}")

    if wybrany_dzien not in st.session_state.historia_wart: st.session_state.historia_wart[wybrany_dzien] = {godzina: [] for godzina in WARTY_SPECYFIKACJA.keys()}
    if wybrany_dzien not in st.session_state.liczba_straznikow: st.session_state.liczba_straznikow[wybrany_dzien] = {godzina: 2 for godzina in WARTY_SPECYFIKACJA.keys()}
    if wybrany_dzien not in st.session_state.lokalizacje_wart: st.session_state.lokalizacje_wart[wybrany_dzien] = {godzina: [] for godzina in WARTY_SPECYFIKACJA.keys()}

    plan_dnia = st.session_state.historia_wart[wybrany_dzien]
    lokalizacje_dnia = st.session_state.lokalizacje_wart[wybrany_dzien]

    with tab_kreator:
        idx_dzis = DNI_OBOZU.index(wybrany_dzien)
        wczorajszy_dzien = DNI_OBOZU[idx_dzis - 1] if idx_dzis > 0 else None

        for godzina, info in WARTY_SPECYFIKACJA.items():
            st.markdown(f"<div class='warta-card'>", unsafe_allow_html=True)
            col_meta, col_inputs, col_actions = st.columns([2.5, 4.5, 1])
            
            with col_meta:
                st.markdown(f"<div style='font-size:19px; font-weight:700; color:#3b82f6; font-family:Orbitron;'>{godzina}</div>", unsafe_allow_html=True)
                st.markdown(f"<div style='color:#94a3b8; font-size:12px; margin-bottom:8px;'>{info['opis']}</div>", unsafe_allow_html=True)
                filtr_pionu = st.multiselect("Filtruj piony:", ["Z", "H", "HS", "W", "I"], default=info['preferencja'], key=f"filter_{godzina}_{wybrany_dzien}")
                
            with col_actions:
                st.markdown("<div style='text-align:center; font-size:11px; color:#94a3b8;'>Obsada</div>", unsafe_allow_html=True)
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("➕", key=f"add_{godzina}_{wybrany_dzien}"): st.session_state.liczba_straznikow[wybrany_dzien][godzina] += 1; st.rerun()
                with c2:
                    if st.button("➖", key=f"sub_{godzina}_{wybrany_dzien}"):
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
                                
                        wybor = st.selectbox(f"Wartownik {slot_idx+1}", opcje, index=idx_default, key=f"sl_{godzina}_{slot_idx}_{wybrany_dzien}")
                        
                        if wybor == "Wpis ręczny (Wyjątek)" or is_manual:
                            plan_dnia[godzina][slot_idx] = st.text_input(f"Kto (ręcznie):", value=aktualny_wybor, key=f"mn_{godzina}_{slot_idx}_{wybrany_dzien}")
                        elif wybor != "-- Wybierz --":
                            plan_dnia[godzina][slot_idx] = wybor.split(" (Warty")[0][2:].split(" [")[0]
                        else:
                            plan_dnia[godzina][slot_idx] = ""

                        lokalizacje_dnia[godzina][slot_idx] = st.text_input(
                            "Posterunek:", value=lokalizacje_dnia[godzina][slot_idx], 
                            key=f"loc_{godzina}_{slot_idx}_{wybrany_dzien}", placeholder="np. Brama"
                        )

                        # Walidacje i podpowiedzi w czasie rzeczywistym
                        if plan_dnia[godzina][slot_idx]:
                            st_name = plan_dnia[godzina][slot_idx]
                            
                            if " (Z)" in st_name and "Z" not in info['preferencja']:
                                st.markdown("<span style='color:#ef4444; font-size:11px; font-weight:bold;'>🛑 Zakaz wart nocnych dla Zuchów!</span>", unsafe_allow_html=True)
                            elif " (Z)" not in st_name and "Z" in info['preferencja']:
                                st.markdown("<span style='color:#facc15; font-size:11px; font-weight:bold;'>⚠️ Sugerowane Zuchy!</span>", unsafe_allow_html=True)
                            else:
                                st.markdown("<span style='color:#22c55e; font-size:11px;'>✅ Pion prawidłowy</span>", unsafe_allow_html=True)

                            if wczorajszy_dzien and wczorajszy_dzien in st.session_state.historia_wart:
                                if st_name in [item for sublist in st.session_state.historia_wart[wczorajszy_dzien].values() for item in sublist]:
                                    st.markdown("<span style='color:#ec4899; font-size:11px; font-weight:bold;'>⚠️ Służba noc po nocy!</span>", unsafe_allow_html=True)
            st.markdown(f"</div>", unsafe_allow_html=True)

        if st.button("💾 ZAPISZ GRAFIK WART I PRZELICZ LICZNIKI", type="primary", use_container_width=True):
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
            st.success("Zapisano! Liczniki w bazie zostały zaktualizowane pomyślnie.")
            st.rerun()

    # ZAKŁADKA DRUKU: ESTETYCZNY PODGLĄD Z BLOKIEM A4
    with tab_druk:
        st.markdown("### 🖨️ Podgląd druku rozkazu (Standard A4)")
        st.caption("Uruchom skrót Ctrl + P (lub Command + P) na klawiaturze, wybierając opcję 'Zapisz jako PDF'. Automatyczne style CSS ukryją zbędne elementy interfejsu.")
        
        jutro_data = (datetime.strptime(wybrany_dzien+".2026", "%d.%m.%Y") + timedelta(days=1)).strftime("%d.%02m")
        
        # Generowanie profesjonalnego układu rozkazu wewnątrz struktury tabeli z ramką
        rozkaz_html = f"""
        <div class="print-preview">
            <div style="text-align: center; border-bottom: 2px solid #000000; padding-bottom: 10px; margin-bottom: 25px;">
                <h1 style="margin: 0; font-size: 26px; text-transform: uppercase; letter-spacing: 2px;">Rozkaz na Wartę Nocną</h1>
                <h2 style="margin: 5px 0 0 0; font-size: 19px; font-weight: normal;">Noc: {wybrany_dzien} / {jutro_data} 2026 r.</h2>
            </div>
            
            <table>
                <thead>
                    <tr>
                        <th style="width: 35%;">GODZINY</th>
                        <th style="width: 65%;">DRUHNA / DRUH (POSTERUNEK)</th>
                    </tr>
                </thead>
                <tbody>
        """

        for g, osoby in plan_dnia.items():
            obsada_elementy = []
            for idx, o in enumerate(osoby):
                miejsce = lokalizacje_dnia[g][idx] if idx < len(lokalizacje_dnia[g]) else ""
                miejsce_str = f" [M: {miejsce}]" if miejsce else ""
                straznik = o if o else "........................................."
                obsada_elementy.append(f"{straznik}{miejsce_str}")
            
            obsada_finalna = "<br>".join(obsada_elementy)
            rozkaz_html += f"""
                <tr>
                    <td style="font-weight: bold; vertical-align: top;">{g}</td>
                    <td>{obsada_finalna}</td>
                </tr>
            """
            
        rozkaz_html += f"""
                </tbody>
            </table>
            <div style="text-align: center; margin-top: 35px; font-size: 14px; font-style: italic;">
                Czuwaj!<br>
                Odprawa służb przy apelu wieczornym.<br><br>
                <div style="text-align: right; font-weight: bold; font-style: normal; margin-top: 15px;">Komendant Obozu</div>
            </div>
        </div>
        """

        # Prezentacja graficznego podglądu na stronie
        st.markdown(rozkaz_html, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        
        if st.button("🖨️ URUCHOM SYSTEMOWE DRUKOWANIE OKNA", type="secondary", use_container_width=True):
            st.html("<script>window.print();</script>")

    # ZAKŁADKA STATYSTYK I EKSPORTU BAZY
    with tab_statystyki:
        st.markdown("#### 📊 Podsumowanie Obciążeń Służbami")
        db_stat = st.session_state.db_uczestnicy
        
        c_tot, c_avg, c_un = st.columns(3)
        c_tot.markdown(f"<div class='metric-card'><h5>Suma przydzielonych wart</h5><h2>{db_stat['Liczba_Wart'].sum()}</h2></div>", unsafe_allow_html=True)
        c_avg.markdown(f"<div class='metric-card'><h5>Średnia wart na osobę</h5><h2>{db_stat['Liczba_Wart'].mean():.1f}</h2></div>", unsafe_allow_html=True)
        c_un.markdown(f"<div class='metric-card'><h5>Uczestnicy bez służby</h5><h2>{len(db_stat[db_stat['Liczba_Wart'] == 0])}</h2></div>", unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("#### 📥 GENERATOR PEŁNEGO RAPORTU KOŃCOWEGO")
        
        export_df = db_stat[['Imię', 'Nazwisko', 'Pion', 'Drużyna', 'Liczba_Wart', 'Ostatnia_Warta']].copy()
        
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
            
        export_df['Szczegółowa_Historia_Wart'] = szczegoly_list

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            export_df.to_excel(writer, index=False, sheet_name='Raport Wart Obozu')
        
        st.download_button(
            label="Pobierz Kompletny Raport Excel (.xlsx)",
            data=buffer.getvalue(),
            file_name=f"Raport_Wart_Obóz_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.dataframe(export_df, hide_index=True, use_container_width=True)
