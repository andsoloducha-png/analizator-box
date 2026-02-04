import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
import tempfile
import traceback

# Import z Twoich modułów
from processing import load_xlsx
import reports as rpt
from export_excel import write_report_xlsx

# Kolejność arkuszy
SHEET_ORDER = [
    "package_type_share",
    "hourly_dims_measured",
    "loop_99",
    "nok_244",
    "overflow_243",
    "hourly_loop_nok_ovf",
    "chute_full",
    "problem_share_type",
    "bad_dims_pct",
    "bad_weight_pct",
]

# Opisy (tekst + pozycja bloku)
DESCRIPTIONS = {
    "package_type_share": ("""Ta tabela przedstawia ilościowy i procentowy rozkład opakowań na instalacji wraz z ich wymiarami

Opis kolumn:

package_type - typ opakowania
avg_lenght -średnia długość paczki danego typu w mm
avg_width  - średnia szerokość  paczki danego typu w mm
avg_height  - średnia wysokość  paczki danego typu w mm
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
avg_lenght -średnia długość paczki danego typu w mm
avg_width  - średnia szerokość  paczki danego typu w mm
avg_height  - średnia wysokość  paczki danego typu w mm
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
    "hourly_loop_nok_ovf": ("""Ta tabela przedstawia, ile paczek w każdej godzinie trafia do loop, overflow, nok i chute full w odniesieniu do wszystkich paczek zarejestrowanych na instalacji

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
        progress_bar.progress(80)
        
        sheets["chute_full"] = rpt.report_chute_full(loaded.df)
        progress_bar.progress(90)
        
        sheets["problem_share_type"] = rpt.report_problem_share_type(loaded.df, min_total=50)
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
    
    # Główne metryki
    st.markdown("### 📊 Podsumowanie")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Wierszy danych", f"{len(loaded.df):,}")
    with col2:
        st.metric("Średnia długość", f"{wavg_len:.2f} mm")
    with col3:
        st.metric("Prognozowana wydajność", f"{pred_eff:,}")
    with col4:
        period = (loaded.max_scan - loaded.min_scan).days if loaded.max_scan and loaded.min_scan else 0
        st.metric("Okres [dni]", f"{period}")
    
    # Top 10 typów opakowań
    if "package_type_share" in sheets:
        st.markdown("### 📦 Top 10 typów opakowań")
        df = sheets["package_type_share"].head(10).copy()
        
        # Stwórz wykres słupkowy
        st.bar_chart(df.set_index('package_type')['items_count_all'])
    
    # Problemy - loop, nok, overflow
    if "hourly_loop_nok_ovf" in sheets:
        st.markdown("### ⚠️ Analiza problemów (Loop, NOK, Overflow)")
        df = sheets["hourly_loop_nok_ovf"].copy()
        
        if 'scan_hour' in df.columns:
            df['scan_hour'] = pd.to_datetime(df['scan_hour'])
            df = df.set_index('scan_hour')
            
            # Wykres liniowy
            st.line_chart(df[['loop_99_count', 'overflow_243_count', 'nok_count']])
    
    # Jakość pomiarów
    if "bad_dims_pct" in sheets and "bad_weight_pct" in sheets:
        st.markdown("### 📏 Jakość pomiarów")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Wymiary - Top 5 problemowych**")
            bad_dims = sheets["bad_dims_pct"].head(5)
            st.dataframe(bad_dims, hide_index=True, use_container_width=True)
        
        with col2:
            st.markdown("**Waga - Top 5 problemowych**")
            bad_weight = sheets["bad_weight_pct"].head(5)
            st.dataframe(bad_weight, hide_index=True, use_container_width=True)


def main():
    # Konfiguracja strony
    st.set_page_config(
        page_title="Analizator BOX 5000 Ultra",
        page_icon="📦",
        layout="wide"
    )

    # Tytuł i opis
    st.title("📦 Analizator BOX 5000 Ultra")
    st.markdown("*Profesjonalna analiza danych logistycznych w kilka sekund*")
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
    1. **Wgraj plik XLSX** z danymi logistycznymi
    2. **Kliknij "Generuj raport"** i poczekaj ~30 sekund
    3. **Pobierz gotowy raport Excel** z 10 arkuszami analiz
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
            
            tab_names = list(st.session_state['sheets'].keys())
            tabs = st.tabs(tab_names)
            
            for i, (name, df) in enumerate(st.session_state['sheets'].items()):
                with tabs[i]:
                    st.markdown(f"**{name}** - Pokazuje pierwsze 50 wierszy")
                    st.dataframe(df.head(50), use_container_width=True, hide_index=True)

    # Footer
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: #666; padding: 20px;'>
        <small>📦 Analizator BOX 5000 Ultra v2.0 | Wersja webowa | Made with Streamlit</small>
    </div>
    """, unsafe_allow_html=True)


if __name__ == "__main__":
    main()
