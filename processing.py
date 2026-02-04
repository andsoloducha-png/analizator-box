from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import Optional

import pandas as pd


@dataclass
class LoadedData:
    df: pd.DataFrame
    min_scan: Optional[datetime]
    max_scan: Optional[datetime]


def _fill_missing_text(df: pd.DataFrame) -> pd.DataFrame:
    """
    Zamienia braki (NaN/puste) na czytelne etykiety, żeby w Excelu nie było pustych pól.
    Dodatkowo usuwa końcówkę '.0' jeśli kolumna była liczbowa (typowy efekt XLSX->pandas).
    """
    mapping = {
        "Chunk Id": "brak chunku",
        "Package type Barcodes": "brak kodu",
        "Discharge": "brak discharge",
    }

    for col, label in mapping.items():
        if col not in df.columns:
            continue

        s = df[col].astype("string")

        # wyczyść spacje, puste stringi -> NA
        s = s.str.strip()
        s = s.replace({"": pd.NA, "nan": pd.NA, "NaN": pd.NA})

        # jeśli Excel zrobił z identyfikatora float (np. 12345.0), usuń ".0"
        s = s.str.replace(r"\.0$", "", regex=True)

        df[col] = s.fillna(label)

    return df


def load_xlsx(path: str) -> LoadedData:
    df = pd.read_excel(path)

    if "Scan" not in df.columns:
        raise RuntimeError("Brak kolumny 'Scan' w XLSX.")

    # Ujednolicenie czasu skanowania
    scan = pd.to_datetime(df["Scan"], errors="coerce")

    df = df.copy()
    df["Scan"] = scan

    # Kolumny wymagane przez reports.py
    df["scan_date"] = scan.dt.date
    df["scan_hour"] = scan.dt.floor("h")

    # zamień braki na czytelne teksty (żeby w Excelu nie było pustych pól)
    df = _fill_missing_text(df)

    min_scan = scan.min()
    max_scan = scan.max()

    if pd.isna(min_scan):
        min_scan = None
    if pd.isna(max_scan):
        max_scan = None

    return LoadedData(df=df, min_scan=min_scan, max_scan=max_scan)
