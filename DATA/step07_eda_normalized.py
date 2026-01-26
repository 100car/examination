"""
STEP 7 | EDA для аркушів *-normalized (UA/KZ/UZ)

Вимога викладача:
- EDA: обрахунок агрегатних функцій
- побудова графіків (гістограми, "ящики з вусами")

Важливо:
- беремо НЕ всі SKU, а тільки ті, що починаються з префіксів із config.yaml (filters.include_prefixes),
  за замовчуванням ["PR-"].
- паралельно зберігаємо два списки:
  - включені (починаються з префіксів)
  - виключені (не починаються з префіксів)

Запуск у Colab:
!python DATA/step07_eda_normalized.py

Вихідні артефакти:
- <eda_dir>/normalized_long.csv
- <eda_dir>/normalized_summary_by_region.csv
- <eda_dir>/symbols_included_<REGION>.csv
- <eda_dir>/symbols_excluded_<REGION>.csv
- <eda_dir>/plots/*.png
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import yaml
import matplotlib.pyplot as plt


# =============================================================================
# ДЕФОЛТНИЙ CONFIG (як у ноутбуці) + завантаження config.yaml з кореня проєкту
# =============================================================================
default_cfg = {
    "paths": {
        "data_root": "DATA",
        "sku_master": "DATA/SKU/sku_master.xlsx",
        "sales_reports_dir": "DATA/SALES/SALES_REPORTS",
        "result_dir": "DATA/SALES/RESULT",
        "result_workbook": "DATA/SALES/RESULT/results.xlsx",
        "eda_dir": "DATA/EDA",
        "ml_dir": "DATA/ML",
    },
    "pipeline": {
        "debug": True,
        "llm_enabled": True,
    },
    "filters": {
        "include_prefixes": ["PR-"],
        "exclude_prefixes": [],
    },
    "analysis": {
        "sku_example": "",
    },
}


CFG_PATH = Path("config.yaml")


def load_cfg() -> Dict:
    """Завантажує config.yaml з поточної директорії. Якщо немає — бере default_cfg."""
    if CFG_PATH.exists():
        with open(CFG_PATH, "r", encoding="utf-8") as f:
            cfg = yaml.safe_load(f) or {}
        # Мʼякий мердж: default_cfg як база
        merged = default_cfg.copy()
        for k, v in (cfg or {}).items():
            if isinstance(v, dict) and isinstance(merged.get(k), dict):
                merged[k] = {**merged[k], **v}
            else:
                merged[k] = v
        return merged
    return default_cfg


# =============================================================================
# Константи
# =============================================================================
MONTHS = ["янв","фев","мар","апр","май","июн","июл","авг","сен","окт","ноя","дек"]

_MONTH_ALIASES = {
    "янв": ["янв", "январ", "january", "jan"],
    "фев": ["фев", "феврал", "february", "feb"],
    "мар": ["мар", "март", "march", "mar"],
    "апр": ["апр", "апрел", "april", "apr"],
    "май": ["май", "мая", "may"],
    "июн": ["июн", "июнь", "june", "jun"],
    "июл": ["июл", "июль", "july", "jul"],
    "авг": ["авг", "август", "august", "aug"],
    "сен": ["сен", "сентябр", "september", "sep", "sept"],
    "окт": ["окт", "октябр", "october", "oct"],
    "ноя": ["ноя", "ноябр", "november", "nov"],
    "дек": ["дек", "декабр", "december", "dec"],
}


def _canon_month(col_name: object) -> str | None:
    if col_name is None:
        return None
    h = str(col_name).strip().lower()
    for m, keys in _MONTH_ALIASES.items():
        if any(k in h for k in keys):
            return m
    return None


def _print_banner(title: str) -> None:
    line = "=" * 72
    print(line)
    print(title)
    print(line)


def _paths_from_cfg(cfg: Dict) -> Tuple[Path, Path]:
    paths = cfg.get("paths", {}) if isinstance(cfg.get("paths", {}), dict) else {}
    results_path = Path(paths.get("result_workbook", default_cfg["paths"]["result_workbook"]))
    eda_dir = Path(paths.get("eda_dir", default_cfg["paths"]["eda_dir"]))
    eda_dir.mkdir(parents=True, exist_ok=True)
    (eda_dir / "plots").mkdir(parents=True, exist_ok=True)
    return results_path, eda_dir


def _prefixes_from_cfg(cfg: Dict) -> Tuple[List[str], List[str]]:
    filt = cfg.get("filters", {}) if isinstance(cfg.get("filters", {}), dict) else {}
    include_prefixes = list(filt.get("include_prefixes", default_cfg["filters"]["include_prefixes"]))
    exclude_prefixes = list(filt.get("exclude_prefixes", default_cfg["filters"]["exclude_prefixes"]))
    return include_prefixes, exclude_prefixes


def _filter_by_prefix(symbols: pd.Series, include_prefixes: List[str]) -> pd.Series:
    s = symbols.astype(str).fillna("")
    mask = False
    for p in include_prefixes:
        mask = mask | s.str.startswith(str(p))
    return mask


def _read_normalized_sheets(results_path: Path) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(results_path, engine="openpyxl")
    out: Dict[str, pd.DataFrame] = {}
    for sh in xls.sheet_names:
        if sh.endswith("-normalized"):
            region = sh.split("-")[0].strip()
            out[region] = pd.read_excel(results_path, sheet_name=sh, engine="openpyxl")
    if not out:
        raise ValueError("У results.xlsx не знайдено жодного листа з суфіксом '-normalized'.")
    return out


def _normalize_month_columns(df: pd.DataFrame) -> pd.DataFrame:
    col_map = {}
    for c in df.columns:
        cm = _canon_month(c)
        if cm:
            col_map[c] = cm
    return df.rename(columns=col_map)


def _build_long_format(df: pd.DataFrame, region: str) -> pd.DataFrame:
    if "Symbol_normalized" not in df.columns:
        raise ValueError("Очікується колонка 'Symbol_normalized' у *-normalized.")

    df2 = _normalize_month_columns(df.copy())

    missing = [m for m in MONTHS if m not in df2.columns]
    if missing:
        raise ValueError(f"У *-normalized відсутні місячні колонки: {missing}. Перевір уніфікацію назв.")

    long = df2.melt(
        id_vars=["Symbol_normalized"],
        value_vars=MONTHS,
        var_name="month",
        value_name="value",
    )
    long["region"] = region
    long["value"] = pd.to_numeric(long["value"], errors="coerce").fillna(0.0)
    return long[["region", "Symbol_normalized", "month", "value"]]


def _save_plot(fig, out_path: Path) -> None:
    fig.tight_layout()
    fig.savefig(out_path, dpi=150)
    plt.close(fig)


def main() -> None:
    cfg = load_cfg()
    results_path, eda_dir = _paths_from_cfg(cfg)
    include_prefixes, _exclude_prefixes = _prefixes_from_cfg(cfg)

    _print_banner("STEP 7 | EDA (*-normalized) | long-format + агрегати + графіки")

    if not results_path.exists():
        raise FileNotFoundError(f"results.xlsx не знайдено: {results_path}")

    sheets = _read_normalized_sheets(results_path)

    long_parts = []
    summary_rows = []

    for region, df in sheets.items():
        if "Symbol_normalized" not in df.columns:
            continue

        mask_in = _filter_by_prefix(df["Symbol_normalized"], include_prefixes)
        df_in = df.loc[mask_in].copy()
        df_out = df.loc[~mask_in].copy()

        # 2 списки (вимога): включені / виключені
        (eda_dir / f"symbols_included_{region}.csv").write_text(
            "\n".join(df_in["Symbol_normalized"].astype(str).tolist()),
            encoding="utf-8",
        )
        (eda_dir / f"symbols_excluded_{region}.csv").write_text(
            "\n".join(df_out["Symbol_normalized"].astype(str).tolist()),
            encoding="utf-8",
        )

        # long формат
        long = _build_long_format(df_in, region)
        long_parts.append(long)

        # агрегати (TOTAL / Stock / Month_sales / Mediana) якщо є
        agg_cols = [c for c in ["TOTAL", "Stock", "Month_sales", "Mediana"] if c in df_in.columns]
        for c in agg_cols:
            s = pd.to_numeric(df_in[c], errors="coerce").fillna(0.0)
            summary_rows.append({
                "region": region,
                "metric": c,
                "count": int(s.count()),
                "sum": float(s.sum()),
                "mean": float(s.mean()),
                "median": float(s.median()),
                "min": float(s.min()),
                "max": float(s.max()),
                "std": float(s.std(ddof=0)),
            })

        # агрегати по місячним (value)
        v = long["value"]
        summary_rows.append({
            "region": region,
            "metric": "monthly_value",
            "count": int(v.count()),
            "sum": float(v.sum()),
            "mean": float(v.mean()),
            "median": float(v.median()),
            "min": float(v.min()),
            "max": float(v.max()),
            "std": float(v.std(ddof=0)),
        })

        # Гістограма місячних продажів по регіону
        fig = plt.figure()
        plt.hist(v.values, bins=30)
        plt.title(f"Гістограма місячних продажів (PR-), регіон {region}")
        plt.xlabel("Продажі за місяць")
        plt.ylabel("К-сть спостережень")
        _save_plot(fig, eda_dir / "plots" / f"hist_monthly_value_{region}.png")

        # Гістограма TOTAL (якщо є)
        if "TOTAL" in df_in.columns:
            t = pd.to_numeric(df_in["TOTAL"], errors="coerce").fillna(0.0)
            fig = plt.figure()
            plt.hist(t.values, bins=30)
            plt.title(f"Гістограма TOTAL (PR-), регіон {region}")
            plt.xlabel("TOTAL за 12 міс")
            plt.ylabel("К-сть SKU")
            _save_plot(fig, eda_dir / "plots" / f"hist_total_{region}.png")

    if not long_parts:
        raise ValueError("Після фільтру по префіксах не залишилось даних для EDA (перевір filters.include_prefixes).")

    long_all = pd.concat(long_parts, ignore_index=True)
    long_all.to_csv(eda_dir / "normalized_long.csv", index=False, encoding="utf-8")

    summary_df = pd.DataFrame(summary_rows)
    summary_df.to_csv(eda_dir / "normalized_summary_by_region.csv", index=False, encoding="utf-8")

    # Boxplot місячних продажів по регіонах
    fig = plt.figure()
    regions = sorted(long_all["region"].unique())
    data = [long_all.loc[long_all["region"] == r, "value"].values for r in regions]
    plt.boxplot(data, labels=regions, showfliers=True)
    plt.title("Ящик з вусами: місячні продажі (PR-) по регіонах")
    plt.xlabel("Регіон")
    plt.ylabel("Продажі за місяць")
    _save_plot(fig, eda_dir / "plots" / "box_monthly_value_by_region.png")

    print("OK | STEP 7 завершено")
    print(f"- long:      {eda_dir / 'normalized_long.csv'}")
    print(f"- summary:   {eda_dir / 'normalized_summary_by_region.csv'}")
    print(f"- plots:     {eda_dir / 'plots'}")


if __name__ == "__main__":
    main()
