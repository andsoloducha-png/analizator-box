import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
import tempfile
import traceback

# Import z modułów
from processing import load_xlsx
import reports as rpt
from export_excel import write_report_xlsx

# Kolejność arkuszy
SHEET_ORDER = [
    "summary",
    "package_type_share",
    "hourly_dims_measured",
    "hourly_weight_measured",
    "loop_99",
    "nok_244",
    "overflow_243",
    "hourly_loop_nok_ovf",
    "chute_full",
    "problem_share_type",
    "bad_dims_pct",
    "bad_weight_pct",
    "top5_heaviest",
    "top5_lightest",
]

# Opisy (tekst + pozycja bloku)
DESCRIPTIONS = {
    "package_type_share": ("""Ta tabela przedstawia ilościowy i procentowy rozkład opakowań na instalacji wraz z ich wymiarami

Opis kolumn:

package_type - typ opakowania
avg_lenght - średnia długość paczki danego typu w mm
avg_width - średnia szerokość  paczki danego typu w mm
avg_height - średnia wysokość  paczki danego typu w mm
items_count_all - ile paczek danego typu wystąpiło na instalacji
pct_share - procentowy rozkład opakowań

Średnie liczone dla opakowań zmierzonych, dzięki czemu nieopomiarowane opakowanie nie zaniżają średniej. Ilości to wszystkie opakowania danego typu, w tym nieopomiarowane. To podejście zapewnia dużą precyzyjność danych

Tabela posortowana według opakowań najczęściej występujących, mających największy udział w rozkładzie
                           
avg_len to średnia długość wszystkich paczek na instalacji (średnia ważona)
predicted_eff to przewidywana wydajność sortera przy założeniu, że średnia długość paczek 400mm daje wydajność 8500 (zgodnie z dokumentacją)
                           
https://drive.google.com/file/d/1g8EU9LQgIKa3NrOvm24-8AwQVLDlzRYW/view?usp=sharing



""", "I4", "O22"),
    "hourly_dims_measured": ("""Ta tabela przedstawia średnie wymiary paczek w rozkładzie godzinowym oraz jakość pomiarów

Opis kolumn:
scan_hour - znacznik czasu
package_type - typ opakowania
avg_lenght - średnia długość paczki danego typu w mm
avg_width - średnia szerokość  paczki danego typu w mm
avg_height - średnia wysokość  paczki danego typu w mm
total_items - wszystkie paczki zarejestrowane na instalacji
unmensured_items - ilość paczek niezmierzonych
pct_unmeasured - procent paczek niezmierzonych

Średnie liczone dla opakowań zmierzonych, dzięki czemu nieopomiarowane opakowanie nie zaniżają średniej. Niezwymiarowanych jest niewiele, dzięki czemu dane są obarczone niskim błędem


""", "K4", "Q22"),
    "loop_99": ("""Ta tabela przedstawia wszystkie paczki wysłane do loop i ile razy

Opis kolumn:

scan_date - znacznik czasu
chunk - numer danej paczki zawarty na etykiecie wysyłkowej
package_type - typ opakowania
discharge - gdzie posortowano (loop)
items_count- ile razy dana paczka trafiła do loop

Jeśli dana paczka trafiła do loop więcej razy niż określa to system, wskazuje to na problem (np. krążenie paczek danego typu)

Jeśli pojawiły się paczki, które mają brak chunku (w kolumnie chunk) są one grupowane i zliczane po typie opakowania (nie musi być to jedna i ta sama paczka)

""", "H3", "N18"),
    "nok_244": ("""Ta tabela przedstawia wszystkie paczki posortowane do zrzutni nok 244 i ile razy

Opis kolumn:

scan_date - znacznik czasu
chunk - numer danej paczki zawarty na etykiecie wysyłkowej
package_type - typ opakowania
discharge - gdzie posortowano (nok 244)
items_count- ile razy dana paczka trafiła do nok 244

Jeśli dana paczka trafiła wielokrotnie do nok, wskazuje to na problem

Jeśli pojawiły się paczki, które mają brak chunku (w kolumnie chunk) są one grupowane i zliczane po typie opakowania (nie musi być to jedna i ta sama paczka)

""", "H4", "N19"),
    "overflow_243": ("""Ta tabela przedstawia wszystkie paczki posortowane do zrzutni overflow i ile razy

Opis kolumn:

scan_date - znacznik czasu
chunk - numer danej paczki zawarty na etykiecie wysyłkowej
package_type - typ opakowania
discharge - gdzie posortowano (overflow 243)
items_count- ile razy dana paczka trafiła do overflow 243

Jeśli dana paczka trafiła wielokrotnie do overflow, wskazuje to na problem
                     
Jeśli pojawiły się paczki, które mają brak chunku (w kolumnie chunk) są one grupowane i zliczane po typie opakowania (nie musi być to jedna i ta sama paczka)

""", "H4", "N19"),
    "hourly_loop_nok_ovf": ("""Ta tabela przedstawia, ile paczek w każdej godzinie trafia do loop, overflow, nok w odniesieniu do wszystkich paczek zarejestrowanych na instalacji

Opis kolumn:

scan_hour - znacznik czasu
total_items - wszystkie rzeczy zarejestrowane na instalacji
loop_99_count - ilość paczek posortowanych do loop
overflow_243_count - ilość paczek posortowana do zrzutni overflow 243
nok_count - ilość paczek posortowana do zrzutni nok 244

""", "H4", "N19"),
    "chute_full": ("""Ta tabela przedstawia, ile paczek z powodu chute full dana zrzutnia wysłała na loop lub - jeśli się zdarzy - do overflow i nok

Opis kolumn:

discharge - gdzie posortowano (loop, overflow, nok)
logic - zawiera numer zrzutni i powód (chute full)
items_count - ilość paczek
                   
""", "F4", "L19"),
    "problem_share_type": ("""Ta tabela przedstawia, jaki typ opakowania ma najwięcej procent wysyłania do loop bądź overflow czy nok

Opis kolumn:

package_type - zawiera kod opakowania
total_items - ilość paczek danego typu zarejestrowano na instalacji
discharge - gdzie posortowano (loop, overflow, nok) 
problem_items - ile paczek z danego typu posortowano do loop, overflow, nok
pct_of_type - ile paczek z danego typu procentowo posortowano do loop, overflow, nok
                           
Tabela posortowana według kolumny pct_of_type malejąco. 

""", "H4", "N19"),
    "bad_dims_pct": ("""Ta tabela przedstawia jakość wymiarowania danych opakowań

Opis kolumn:

type - typ opakowania
bad_meaasurements - ile razy paczki z danym typem opakowania nie było zwymiarowane
total_items - ile razy dany typ opakowania wystąpił na instalacji
pct_bad - ile procent opakowań danego typu nie jest wymiarowanych przez instalację

""", "G3", "M18"),
    "bad_weight_pct": ("""Ta tabela przedstawia jakość ważenia danych opakowań

Opis kolumn:

type - typ opakowania
bad_weight - ile razy paczki z danym typem opakowania nie było zważone
total_items - ile razy dany typ opakowania wystąpił na instalacji
pct_bad_weight - ile procent opakowań danego typu nie jest ważonych przez instalację

""", "G4", "M19"),

    "hourly_weight_measured": ("""Ta tabela przedstawia średnią masę paczek (kolumna Volume (waga) - masa w gramach) oraz skuteczność ważenia w ujęciu godzinowym.
avg_weight: średnia masa [g] 
measured_items: liczba paczek z masą 
unmeasured_items: niezważone paczki
pct_unmeasured: % paczek bez poprawnej masy""", "I2", "N6"),

    "top5_heaviest": ("""TOP 5 najcięższych paczek.
Kolumny:
chunk - Chunk Id
type - Package type Barcodes
weight - masa [g]""", "I2", "N6"),

    "top5_lightest": ("""TOP 5 najlżejszych paczek.
Kolumny:
chunk - Chunk Id
type - Package type Barcodes
weight - masa [g]""", "I2", "N6"),
}


def generate_report(uploaded_file):
    """Główna funkcja generująca raport"""
    try:
        # Zapisz tymczasowo plik
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_input:
            tmp_input.write(uploaded_file.getvalue())
            tmp_input_path = tmp_input.name

        # Wczytaj dane
        st.info(f"📂 Wczytuję plik: {uploaded_file.name}")
        loaded = load_xlsx(tmp_input_path)

        if loaded.min_scan is None or loaded.max_scan is None:
            raise RuntimeError("Nie udało się sparsować kolumny Scan (brak dat).")

        st.success(f"📅 Zakres czasu: {loaded.min_scan} → {loaded.max_scan}")
        st.info(f"📊 Wierszy w pliku: {len(loaded.df):,}")

        if len(loaded.df) == 0:
            raise RuntimeError("Plik po wczytaniu ma 0 wierszy.")

        # Liczenie raportów
        st.info("⚙️ Liczenie raportów...")
        progress_bar = st.progress(0)
        sheets = {}

        sheets["bad_dims_pct"] = rpt.report_bad_dims_pct(loaded.df)
        progress_bar.progress(10)
        
        sheets["bad_weight_pct"] = rpt.report_bad_weight_pct(loaded.df)
        progress_bar.progress(20)
        
        sheets["package_type_share"] = rpt.report_package_type_dims_share(loaded.df)
        progress_bar.progress(30)

        sheets["loop_99"] = rpt.report_discharge_detail(loaded.df, "99 Loop")
        progress_bar.progress(40)
        
        sheets["nok_244"] = rpt.report_discharge_detail(loaded.df, "Not Ok 244")
        progress_bar.progress(50)
        
        sheets["overflow_243"] = rpt.report_discharge_detail(loaded.df, "Overflow 243")
        progress_bar.progress(60)

        sheets["hourly_loop_nok_ovf"] = rpt.report_hourly_loop_nok_overflow(loaded.df)
        progress_bar.progress(70)
        
        sheets["hourly_dims_measured"] = rpt.report_hourly_dims_measured(loaded.df)
        progress_bar.progress(78)
        
        sheets["hourly_weight_measured"] = rpt.report_hourly_weight_measured(loaded.df)
        progress_bar.progress(82)
        sheets["chute_full"] = rpt.report_chute_full(loaded.df)
        progress_bar.progress(90)
        
        sheets["problem_share_type"] = rpt.report_problem_share_type(loaded.df, min_total=50)
        progress_bar.progress(94)

        top_heavy, top_light = rpt.report_top5_weight_extremes(loaded.df)
        sheets["top5_heaviest"] = top_heavy
        sheets["top5_lightest"] = top_light
        progress_bar.progress(95)
        # Policz średnią ważoną i prognozowaną wydajność
        wavg_len, pred_eff = rpt.compute_weighted_length_and_efficiency(
            sheets["package_type_share"],
            base_efficiency=8500.0,
            base_avg_length=400.0,
        )

        # Zapisz raport do tymczasowego pliku
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_output:
            tmp_output_path = Path(tmp_output.name)

        # --- SUMMARY sheet (do Excela jako pierwszy arkusz) ---
        total_rows = int(len(loaded.df))
        avg_length_mm = float(wavg_len) if pd.notna(wavg_len) else float("nan")
        total_length_km = (total_rows * avg_length_mm / 1_000_000) if pd.notna(avg_length_mm) else 0.0

        vol = pd.to_numeric(loaded.df.get("Volume"), errors="coerce")
        total_mass_g = vol[vol > 0].sum() if vol is not None else 0.0
        total_mass_t = float(total_mass_g) / 1_000_000

        if loaded.min_scan and loaded.max_scan:
            scan_min = loaded.min_scan.strftime("%Y-%m-%d")
            scan_max = loaded.max_scan.strftime("%Y-%m-%d")
            scan_label = scan_max if scan_min == scan_max else f"{scan_min} → {scan_max}"
        else:
            scan_label = "brak"

        sheets["summary"] = pd.DataFrame([{
            "scan": scan_label,
            "rows": total_rows,
            "avg_length_mm": round(avg_length_mm, 2) if pd.notna(avg_length_mm) else pd.NA,
            "total_length_km": round(total_length_km, 2),
            "total_mass_t": round(total_mass_t, 3),
        }])

        write_report_xlsx(
            tmp_output_path,
            sheets,
            sheet_order=SHEET_ORDER,
            descriptions=DESCRIPTIONS,
            package_type_share_summary=(wavg_len, pred_eff),
        )
        
        progress_bar.progress(100)
        st.success("✅ Raport wygenerowany!")

        # Zwróć plik do pobrania i dane do wizualizacji
        with open(tmp_output_path, 'rb') as f:
            return f.read(), tmp_output_path.name, sheets, (wavg_len, pred_eff), loaded

    except Exception as e:
        st.error(f"❌ Błąd: {e}")
        with st.expander("📋 Szczegóły błędu"):
            st.code(traceback.format_exc())
        return None, None, None, None, None



def show_visualizations(sheets, summary, loaded):
    """Wyświetl wizualizacje danych"""
    wavg_len, pred_eff = summary

    st.markdown("### 📊 Podsumowanie")

    # KPI bazowe
    rows = int(len(loaded.df))
    avg_length_mm = float(wavg_len) if pd.notna(wavg_len) else float("nan")
    total_length_km = (rows * avg_length_mm / 1_000_000) if pd.notna(avg_length_mm) else 0.0

    # Volume = masa w gramach
    vol = pd.to_numeric(loaded.df.get("Volume"), errors="coerce")
    total_mass_g = vol[vol > 0].sum() if vol is not None else 0.0
    total_mass_t = float(total_mass_g) / 1_000_000  # g -> t

    # Data (Scan)
    if loaded.min_scan and loaded.max_scan:
        if loaded.min_scan.date() == loaded.max_scan.date():
            scan_label = loaded.max_scan.strftime("%Y-%m-%d")
        else:
            scan_label = f"{loaded.min_scan.strftime('%Y-%m-%d')} → {loaded.max_scan.strftime('%Y-%m-%d')}"
    else:
        scan_label = "brak"

    # Skuteczności ważone (sum(measured)/sum(total))
    dims_eff = None
    if "hourly_dims_measured" in sheets:
        d = sheets["hourly_dims_measured"]
        if "measured_items" in d.columns and "total_items" in d.columns and d["total_items"].sum() > 0:
            dims_eff = float(d["measured_items"].sum() / d["total_items"].sum() * 100.0)

    weight_eff = None
    if "hourly_weight_measured" in sheets:
        w = sheets["hourly_weight_measured"]
        if "measured_items" in w.columns and "total_items" in w.columns and w["total_items"].sum() > 0:
            weight_eff = float(w["measured_items"].sum() / w["total_items"].sum() * 100.0)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Data", scan_label)
    with col2:
        st.metric("Przeprocesowane paczki", f"{rows}")
    with col3:
        st.metric("Średnia długość paczek", f"{avg_length_mm:.2f} mm" if pd.notna(avg_length_mm) else "brak")
    with col4:
        st.metric("Prognozowana wydajność sortera", f"{pred_eff}")

    col5, col6, col7, col8 = st.columns(4)
    with col5:
        st.metric("Przeprocesowana długość paczek", f"{total_length_km:.2f} km")
    with col6:
        st.metric("Przeprocesowana masa paczek", f"{total_mass_t:.3f} t")
    with col7:
        st.metric("Skuteczność mierzenia", f"{dims_eff:.2f}%" if dims_eff is not None else "brak")
    with col8:
        st.metric("Skuteczność ważenia", f"{weight_eff:.2f}%" if weight_eff is not None else "brak")

    # 1) Top 10 typów opakowań
    if "package_type_share" in sheets:
        st.markdown("### 📦 Top 10 typów opakowań")
        df = sheets["package_type_share"].head(10).copy()
        st.bar_chart(df.set_index("package_type")["items_count_all"])

    # 2) Wolumen całkowity w czasie (godzinowo)
    if "hourly_loop_nok_ovf" in sheets:
        st.markdown("### 📈 Wolumen całkowity w czasie")
        hourly_df = sheets["hourly_loop_nok_ovf"].copy()
        if "scan_hour" in hourly_df.columns and "total_items" in hourly_df.columns:
            hourly_df["scan_hour"] = pd.to_datetime(hourly_df["scan_hour"])
            hourly_df = hourly_df.set_index("scan_hour").sort_index()
            st.area_chart(hourly_df[["total_items"]])

    # 3) Skuteczność mierzenia i ważenia (godzinowo)
    st.markdown("### ✅ Skuteczność mierzenia i ważenia (godzinowo)")
    eff_df = None

    if "hourly_dims_measured" in sheets:
        d = sheets["hourly_dims_measured"].copy()
        if "scan_hour" in d.columns:
            d["scan_hour"] = pd.to_datetime(d["scan_hour"])
            d = d.set_index("scan_hour").sort_index()
            if "pct_unmeasured" in d.columns:
                eff_df = pd.DataFrame({"skuteczność_mierzenia_%": (100.0 - d["pct_unmeasured"]).round(2)})

    if "hourly_weight_measured" in sheets:
        w = sheets["hourly_weight_measured"].copy()
        if "scan_hour" in w.columns:
            w["scan_hour"] = pd.to_datetime(w["scan_hour"])
            w = w.set_index("scan_hour").sort_index()
            if "pct_unmeasured" in w.columns:
                w_eff = pd.DataFrame({"skuteczność_ważenia_%": (100.0 - w["pct_unmeasured"]).round(2)})
                eff_df = w_eff if eff_df is None else eff_df.join(w_eff, how="outer")

    if eff_df is not None and not eff_df.empty:
        st.line_chart(eff_df)
    else:
        st.info("Brak danych do wykresu skuteczności.")

    # 4) Problemy w czasie (Loop, NOK, Overflow) - liczby bezwzględne
    if "hourly_loop_nok_ovf" in sheets:
        st.markdown("### ⚠️ Loop, NOK, Overflow w czasie")
        prob_df = sheets["hourly_loop_nok_ovf"].copy()
        if "scan_hour" in prob_df.columns:
            prob_df["scan_hour"] = pd.to_datetime(prob_df["scan_hour"])
            prob_df = prob_df.set_index("scan_hour").sort_index()
            problem_cols = ["loop_99_count", "overflow_243_count", "nok_count"]
            available_cols = [c for c in problem_cols if c in prob_df.columns]
            if available_cols:
                st.line_chart(prob_df[available_cols])
            
                # 4a) Analiza problemów jako procent całości
                st.markdown("### ⚠️ Loop, NOK, Overflow jako procent wolumenu")
                if "total_items" in prob_df.columns:
                    pct = prob_df[available_cols].div(prob_df["total_items"], axis=0) * 100.0
                    st.line_chart(pct)

    # 5) Jakość pomiarów (tylko pojedyncze typy w aplikacji)
    if "bad_dims_pct" in sheets and "bad_weight_pct" in sheets:
        st.markdown("### 📏 Jakość pomiarów – pojedyncze typy")

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("**Wymiary - Top 5 problemowych**")
            bad_dims = sheets["bad_dims_pct"].copy()
            if "type" in bad_dims.columns:
                bad_dims = bad_dims[~bad_dims["type"].astype(str).str.contains(";", regex=False)]
            st.dataframe(bad_dims.head(5), hide_index=True, use_container_width=True)

        with col2:
            st.markdown("**Masa - Top 5 problemowych**")
            bad_weight = sheets["bad_weight_pct"].copy()
            if "type" in bad_weight.columns:
                bad_weight = bad_weight[~bad_weight["type"].astype(str).str.contains(";", regex=False)]
            st.dataframe(bad_weight.head(5), hide_index=True, use_container_width=True)

    # 6) Ekstrema masy (TOP 5) – tylko pojedyncze typy w aplikacji
    if "top5_heaviest" in sheets and "top5_lightest" in sheets:
        st.markdown("### 🏋️ Ekstrema masy (TOP 5) – pojedyncze typy ")

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("**TOP 5 najcięższych**")
            top_heavy = sheets["top5_heaviest"].copy()
            if "type" in top_heavy.columns:
                top_heavy = top_heavy[~top_heavy["type"].astype(str).str.contains(";", regex=False)]
            st.dataframe(top_heavy.head(5), hide_index=True, use_container_width=True)

        with col2:
            st.markdown("**TOP 5 najlżejszych**")
            top_light = sheets["top5_lightest"].copy()
            if "type" in top_light.columns:
                top_light = top_light[~top_light["type"].astype(str).str.contains(";", regex=False)]
            st.dataframe(top_light.head(5), hide_index=True, use_container_width=True)

def main():
    # Konfiguracja strony
    st.set_page_config(
        page_title="Analizator BOX 5000 Ultra",
        page_icon="📦",
        layout="wide"
    )

    # Tytuł i opis
    st.title("📦 Analizator BOX 5000 Ultra")
    st.markdown("*Profesjonalna analiza danych w kilka sekund*")
    st.markdown("---")
    
    # Sidebar z informacjami
    with st.sidebar:
        st.markdown("### ℹ️ Informacje")
        st.markdown("""
        **Raport zawiera:**
        - 📊 Rozkład typów opakowań
        - ⏱️ Analiza godzinowa
        - 🔄 Szczegóły Loop/NOK/Overflow
        - 📏 Jakość pomiarów
        - ⚡ Prognoza wydajności
        
        **Limity:**
        - Max: ~100k wierszy 
        - Formaty: XLSX
        - Czas: ~30-60 sek
        """)
        
        st.markdown("---")
        st.markdown("### 🎨 Opcje")
        show_preview = st.checkbox("Pokaż wizualizacje", value=True)
        show_data_preview = st.checkbox("Pokaż podgląd tabel", value=True)
    
    # Główna zawartość
    st.markdown("""
    ### 🚀 Jak używać:
    1. **Wgraj plik XLSX** z danymi MFC/Maintenace/Box sort detail
    2. **Kliknij "Generuj raport"** i poczekaj ~30-60 sekund
    3. **Obejrzyj** raport na stronie lub **pobierz** table z opisem w pliku Excel
    """)
    
    st.markdown("---")

    # Upload pliku
    uploaded_file = st.file_uploader(
        "📁 Wybierz plik XLSX do analizy",
        type=['xlsx'],
        help="Plik musi zawierać kolumnę 'Scan' z datami oraz dane logistyczne (Discharge, Package type, etc.)"
    )

    if uploaded_file is not None:
        # Wyświetl info o pliku
        file_size_mb = uploaded_file.size / 1024 / 1024
        st.info(f"📄 Wybrany plik: **{uploaded_file.name}** ({file_size_mb:.2f} MB)")
        
        # Przycisk generowania
        if st.button("🚀 Generuj raport", type="primary", use_container_width=True):
            with st.spinner("🔄 Przetwarzam dane... To może potrwać ~30-60 sekund"):
                report_data, _, sheets, summary, loaded = generate_report(uploaded_file)
                
                if report_data and sheets and summary and loaded:
                    # Zapisz w session state
                    st.session_state['report_data'] = report_data
                    st.session_state['sheets'] = sheets
                    st.session_state['summary'] = summary
                    st.session_state['loaded'] = loaded
                    st.session_state['uploaded_filename'] = uploaded_file.name
                    
                    st.balloons()
                    st.success("🎉 Raport gotowy do pobrania!")

    # Jeśli raport został wygenerowany, pokaż przycisk pobierania i wizualizacje
    if 'report_data' in st.session_state:
        st.markdown("---")
        
        # Przycisk pobierania
        stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        filename = f"BOX_raport_{stamp}.xlsx"
        
        st.download_button(
            label="⬇️ Pobierz raport Excel",
            data=st.session_state['report_data'],
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        
        # Wizualizacje
        if show_preview:
            st.markdown("---")
            show_visualizations(
                st.session_state['sheets'],
                st.session_state['summary'],
                st.session_state['loaded']
            )
        
        # Podgląd tabel
        if show_data_preview:
            st.markdown("---")
            st.markdown("### 📋 Podgląd danych")
            
            # Tabele w tej samej kolejności co arkusze w wygenerowanym XLSX
            ordered = [n for n in SHEET_ORDER if n in st.session_state['sheets']]
            for n in st.session_state['sheets'].keys():
                if n not in ordered:
                    ordered.append(n)

            tabs = st.tabs(ordered)

            for i, name in enumerate(ordered):
                df = st.session_state['sheets'][name]
                with tabs[i]:
                    st.markdown(f"**{name}** - Pokazuje pierwsze 50 wierszy")
                    st.dataframe(df.head(50), use_container_width=True, hide_index=True)

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #666; padding: 20px;'>
        <small>📦 Analizator BOX 5000 Ultra | Wersja webowa </small>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
