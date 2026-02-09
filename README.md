## Analizator BOX 5000 Ultra

Aplikacja webowa (Streamlit) do analizy danych z sortera oraz generowania raportu Excel (XLSX) z opisami.

## Funkcjonalności
- dashboard podsumowujący (daty Scan, liczba paczek, długość, masa, skuteczności)
- analiza typów opakowań (Top 10)
- skuteczność wymiarowania i ważenia (godzinowo)
- analiza problemów: Loop / NOK / Overflow (liczby + %)
- TOP 5 najcięższych i najlżejszych paczek
- raport Excel z opisami na żółtym tle

## Kolumna Volume
W pliku XLSX kolumna `Volume` oznacza **wagę paczki w gramach**.  
W aplikacji i raporcie używane jest wyłącznie nazewnictwo **waga / masa**.

## Struktura projektu
```
.
├── streamlit_app_advanced.py
├── reports.py
├── processing.py
├── export_excel.py
├── requirements.txt
└── README.md
```

## Uruchomienie
```bash
pip install -r requirements.txt
streamlit run streamlit_app_advanced.py
```

## Raport Excel
- automatyczne formatowanie i opisy

## Przeznaczenie
Utrzymanie ruchu, inżynieria procesu, analiza jakości sortowania i raportowanie operacyjne.
