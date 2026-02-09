from __future__ import annotations

import pandas as pd


DISCHARGES = ["99 Loop", "Not Ok 244", "Overflow 243"]


def _get_package_type_col(df: pd.DataFrame) -> str:
    if "package_type" in df.columns:
        return "package_type"
    if "Package type Barcodes" in df.columns:
        return "Package type Barcodes"
    raise KeyError("Brak kolumny 'Package type Barcodes' lub 'package_type'.")


def report_bad_dims_pct(df: pd.DataFrame) -> pd.DataFrame:
    type_col = _get_package_type_col(df)
    g = df.groupby(type_col, dropna=False)

    bad = g.apply(
        lambda x: (
            (x["Length"].isna() | x["Width"].isna() | x["Height"].isna())
            | (x["Length"] <= 0) | (x["Width"] <= 0) | (x["Height"] <= 0)
        ).sum(),
        include_groups=False,
    )

    total = g.size()
    out = pd.DataFrame({
        "type": total.index,
        "bad_measurements": bad.values,
        "total_items": total.values,
        "pct_bad": (bad.values / total.values * 100.0)
    })
    out["pct_bad"] = out["pct_bad"].round(2)
    out = out.sort_values(["pct_bad", "bad_measurements"], ascending=[False, False]).reset_index(drop=True)
    return out


def report_bad_weight_pct(df: pd.DataFrame) -> pd.DataFrame:
    type_col = _get_package_type_col(df)
    g = df.groupby(type_col, dropna=False)

    # UWAGA: zostawiamy "Volume" (tak jest nazwane w systemie)
    bad = g.apply(
        lambda x: (x["Volume"].isna() | (x["Volume"] <= 0)).sum(),
        include_groups=False,
    )
    total = g.size()

    out = pd.DataFrame({
        "type": total.index,
        "bad_weight": bad.values,
        "total_items": total.values,
        "pct_bad_weight": (bad.values / total.values * 100.0)
    })
    out["pct_bad_weight"] = out["pct_bad_weight"].round(2)
    out = out.sort_values(["pct_bad_weight", "bad_weight"], ascending=[False, False]).reset_index(drop=True)
    return out


def report_package_type_dims_share(df: pd.DataFrame) -> pd.DataFrame:
    """
    Odpowiednik SQL z v_box_shift:
    - grupowanie po package_type
    - średnie wymiarów tylko z wartości > 0
    - items_count_all
    - pct_share w całym wolumenie
    - 'brak pomiaru' gdy średnia = NaN
    """
    type_col = _get_package_type_col(df)

    # bezpiecznie: konwersja na numery
    work = df.copy()
    for col in ["Length", "Width", "Height"]:
        if col in work.columns:
            work[col] = pd.to_numeric(work[col], errors="coerce")

    g = work.groupby(type_col, dropna=False)

    out = g.agg(
        avg_length=("Length", lambda s: s[s > 0].mean()),
        avg_width=("Width",  lambda s: s[s > 0].mean()),
        avg_height=("Height", lambda s: s[s > 0].mean()),
        items_count_all=(type_col, "size"),
    ).reset_index()

    total_count = len(work)
    out["pct_share"] = (100.0 * out["items_count_all"] / total_count).round(2) if total_count else pd.NA

    # -> tekst z przecinkiem + "brak pomiaru"
    for col in ["avg_length", "avg_width", "avg_height"]:
        out[col] = out[col].round(2)
        out[col] = out[col].map(
            lambda x: "brak pomiaru"
            if pd.isna(x)
            else f"{x:.2f}".replace(".", ",")
        )

    out = out.rename(columns={type_col: "package_type"})
    out = out.sort_values("items_count_all", ascending=False).reset_index(drop=True)
    return out


def report_hourly_weight(df: pd.DataFrame) -> pd.DataFrame:
    """
    Alias dla starego GUI.
    """
    return report_package_type_dims_share(df)


def report_discharge_detail(df: pd.DataFrame, discharge: str) -> pd.DataFrame:
    sub = df[df["Discharge"] == discharge].copy()
    out = (
        sub.groupby(["scan_date", "Chunk Id", "Package type Barcodes", "Discharge"], dropna=False)
        .size()
        .reset_index(name="items_count")
        .rename(columns={
            "Chunk Id": "chunk",
            "Package type Barcodes": "package_type",
            "Discharge": "discharge",
        })
        .sort_values(["discharge", "items_count"], ascending=[True, False])
        .reset_index(drop=True)
    )
    return out


def report_hourly_loop_nok_overflow(df: pd.DataFrame) -> pd.DataFrame:
    out = (
        df.groupby("scan_hour", dropna=False)
        .agg(
            total_items=("scan_hour", "size"),
            loop_99_count=("Discharge", lambda s: (s == "99 Loop").sum()),
            overflow_243_count=("Discharge", lambda s: (s == "Overflow 243").sum()),
            nok_count=("Discharge", lambda s: (s == "Not Ok 244").sum()),
        )
        .reset_index()
        .sort_values("scan_hour")
        .reset_index(drop=True)
    )
    return out



def report_hourly_weight_measured(df: pd.DataFrame) -> pd.DataFrame:
    """
    Godzinowa jakość ważenia na podstawie kolumny 'Volume' (masa w gramach).
    measured_items: Volume > 0
    unmeasured_items: Volume is NaN lub <= 0
    """
    if "Volume" not in df.columns:
        raise KeyError("Brak kolumny 'Volume' w danych.")

    g = df.groupby("scan_hour", dropna=False)

    out = g.agg(
        avg_weight_g=("Volume", lambda s: s[s > 0].mean()),
        total_items=("scan_hour", "size"),
        measured_items=("Volume", lambda s: (s > 0).sum()),
        unmeasured_items=("Volume", lambda s: (s.isna() | (s <= 0)).sum()),
    ).reset_index()

    out["pct_unmeasured"] = (out["unmeasured_items"] / out["total_items"] * 100.0).round(2)
    out = out.sort_values("scan_hour").reset_index(drop=True)
    return out


def report_top5_weight_extremes(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Zwraca (top5_najciezsze, top5_najlzejsze) dla Volume > 0.
    Kolumny: chunk, type, volume.
    """
    if "Volume" not in df.columns:
        raise KeyError("Brak kolumny 'Volume' w danych.")
    # używamy oryginalnych nazw, żeby działało na surowym df z load_xlsx
    cols = ["Chunk Id", "Package type Barcodes", "Volume"]
    for c in cols:
        if c not in df.columns:
            raise KeyError(f"Brak kolumny '{c}' w danych.")

    sub = df.loc[df["Volume"] > 0, cols].copy()

    sub = sub.rename(columns={
        "Chunk Id": "chunk",
        "Package type Barcodes": "type",
        "Volume": "weight_g",
    })

    # zabezpieczenie na brak danych
    if sub.empty:
        empty = pd.DataFrame(columns=["chunk", "type", "weight_g"])
        return empty, empty

    top_heavy = sub.sort_values("weight_g", ascending=False).head(5).reset_index(drop=True)
    top_light = sub.sort_values("weight_g", ascending=True).head(5).reset_index(drop=True)
    return top_heavy, top_light

def report_hourly_dims_measured(df: pd.DataFrame) -> pd.DataFrame:
    g = df.groupby("scan_hour", dropna=False)

    out = g.agg(
        avg_length=("Length", lambda s: s[s > 0].mean()),
        avg_width=("Width",  lambda s: s[s > 0].mean()),
        avg_height=("Height", lambda s: s[s > 0].mean()),
        total_items=("scan_hour", "size"),

        measured_items=(
            "Length",
            lambda s: ((s > 0)
                       & (df.loc[s.index, "Width"] > 0)
                       & (df.loc[s.index, "Height"] > 0)).sum()
        ),
        unmeasured_items=(
            "Length",
            lambda s: (
                s.isna()
                | (s <= 0)
                | df.loc[s.index, "Width"].isna()
                | (df.loc[s.index, "Width"] <= 0)
                | df.loc[s.index, "Height"].isna()
                | (df.loc[s.index, "Height"] <= 0)
            ).sum()
        ),
    ).reset_index()

    out["pct_unmeasured"] = (out["unmeasured_items"] / out["total_items"] * 100.0).round(2)
    out = out.sort_values("scan_hour").reset_index(drop=True)
    return out


def report_chute_full(df: pd.DataFrame) -> pd.DataFrame:
    sub = df[
        df["Logic"].astype("string").str.contains("Chute Full", na=False)
        & df["Discharge"].isin(DISCHARGES)
    ].copy()

    out = (
        sub.groupby(["Discharge", "Logic"], dropna=False)
        .size()
        .reset_index(name="items_count")
        .rename(columns={"Discharge": "discharge", "Logic": "logic"})
        .sort_values(["discharge", "items_count"], ascending=[True, False])
        .reset_index(drop=True)
    )
    return out


def report_problem_share_type(df: pd.DataFrame, min_total: int = 50) -> pd.DataFrame:
    totals = df.groupby("Package type Barcodes", dropna=False).size().rename("total_items").reset_index()
    probs = (
        df[df["Discharge"].isin(DISCHARGES)]
        .groupby(["Package type Barcodes", "Discharge"], dropna=False)
        .size()
        .rename("problem_items")
        .reset_index()
    )

    out = totals.merge(probs, on="Package type Barcodes", how="left")
    out["problem_items"] = out["problem_items"].fillna(0).astype(int)
    out["pct_of_type"] = (out["problem_items"] / out["total_items"] * 100.0).round(2)
    out = out.rename(columns={"Package type Barcodes": "package_type", "Discharge": "discharge"})
    out = out[out["total_items"] >= min_total].copy()
    out = out.sort_values(["pct_of_type", "problem_items"], ascending=[False, False]).reset_index(drop=True)
    return out


def compute_weighted_length_and_efficiency(
    package_type_share_df: pd.DataFrame,
    base_efficiency: float = 8500.0,
    base_avg_length: float = 400.0,
) -> tuple[float, int]:
    """
    Liczy:
    - średnią ważoną długość na podstawie avg_length (kol. B) i items_count_all (kol. E),
      z pominięciem 'brak pomiaru'
    - prognozowaną wydajność z proporcji odwrotnej:
        efficiency = base_efficiency * base_avg_length / weighted_avg_length

    Zwraca: (weighted_avg_length_2dp, efficiency_int)
    """
    df = package_type_share_df.copy()

    if "avg_length" not in df.columns or "items_count_all" not in df.columns:
        return (float("nan"), 0)

    # avg_length jest tekstem "409,15" albo "brak pomiaru" -> zamieniamy na float
    avg_len_num = pd.to_numeric(
        df["avg_length"].astype(str).str.replace(",", ".", regex=False),
        errors="coerce",
    )

    weights = pd.to_numeric(df["items_count_all"], errors="coerce").fillna(0)

    valid = avg_len_num.notna() & (weights > 0)
    if valid.sum() == 0:
        return (float("nan"), 0)

    w_sum = (avg_len_num[valid] * weights[valid]).sum()
    w = weights[valid].sum()
    if w == 0:
        return (float("nan"), 0)

    weighted_avg = float(w_sum / w)
    efficiency = float(base_efficiency) * float(base_avg_length) / weighted_avg if weighted_avg else 0.0

    return (round(weighted_avg, 2), int(round(efficiency)))
