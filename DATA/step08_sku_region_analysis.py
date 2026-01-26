"""
STEP 08 | Аналіз одного SKU по 3 регіонах (UA/KZ/UZ) на основі *-normalized

Мета:
- показово продемонструвати, що продажі можуть відрізнятися залежно від регіону.

Беремо SKU:
- або з config.yaml: analysis.sku_example
- або, якщо не задано, беремо перший SKU з префіксом filters.include_prefixes у UA-normalized.

Вихід:
- графік: <eda_dir>/plots/sku_<SKU>_3regions.png
- короткий висновок у консолі

Запуск:
!python DATA/step08_sku_region_analysis.py
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List

import numpy as np
import pandas as pd
import yaml
import matplotlib.pyplot as plt


# =============================================================================
# ДЕФОЛТНИЙ CONFIG
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
MONTHS = ["янв","фев","мар","апр","май","июн","июл","авг","сен","окт","ноя","дек"]


def load_cfg() -> Dict:
    if CFG_PATH.exists():
        with open(CFG_PATH, "r", encoding="utf-8") as f:
            cfg = yaml.safe_load(f) or {}
        merged = default_cfg.copy()
        for k, v in (cfg or {}).items():
            if isinstance(v, dict) and isinstance(merged.get(k), dict):
                merged[k] = {**merged[k], **v}
            else:
                merged[k] = v
        return merged
    return default_cfg


def _print_banner(title: str) -> None:
    line = "=" * 72
    print(line)
    print(title)
    print(line)


def _paths_from_cfg(cfg: Dict) -> tuple[Path, Path]:
    paths = cfg.get("paths", {}) if isinstance(cfg.get("paths", {}), dict) else {}
    results_path = Path(paths.get("result_workbook", default_cfg["paths"]["result_workbook"]))
    eda_dir = Path(paths.get("eda_dir", default_cfg["paths"]["eda_dir"]))
    eda_dir.mkdir(parents=True, exist_ok=True)
    (eda_dir / "plots").mkdir(parents=True, exist_ok=True)
    return results_path, eda_dir


def _include_prefixes(cfg: Dict) -> List[str]:
    filt = cfg.get("filters", {}) if isinstance(cfg.get("filters", {}), dict) else {}
    return list(filt.get("include_prefixes", default_cfg["filters"]["include_prefixes"]))


def _filter_by_prefix(symbols: pd.Series, include_prefixes: List[str]) -> pd.Series:
    s = symbols.astype(str).fillna("")
    mask = False
    for p in include_prefixes:
        mask = mask | s.str.startswith(str(p))
    return mask


def _pick_default_sku(results_path: Path, include_prefixes: List[str]) -> str:
    df = pd.read_excel(results_path, sheet_name="UA-normalized", engine="openpyxl")
    if "Symbol_normalized" not in df.columns:
        raise ValueError("UA-normalized не містить колонку Symbol_normalized.")
    mask = _filter_by_prefix(df["Symbol_normalized"], include_prefixes)
    pr = df.loc[mask, "Symbol_normalized"].astype(str)
    if pr.empty:
        raise ValueError("Не знайдено жодного SKU з потрібним префіксом у UA-normalized.")
    return pr.iloc[0].strip()


def _extract_row(results_path: Path, region: str, sku: str) -> pd.Series | None:
    df = pd.read_excel(results_path, sheet_name=f"{region}-normalized", engine="openpyxl")
    key = "Symbol_normalized" if "Symbol_normalized" in df.columns else "Symbol"
    sub = df[df[key].astype(str).str.strip() == sku]
    if sub.empty:
        return None
    return sub.iloc[0]


def main() -> None:
    cfg = load_cfg()
    results_path, eda_dir = _paths_from_cfg(cfg)
    include_prefixes = _include_prefixes(cfg)

    _print_banner("STEP 08 | Аналіз одного SKU по регіонах (UA/KZ/UZ)")

    if not results_path.exists():
        raise FileNotFoundError(f"results.xlsx не знайдено: {results_path}")

    sku = str(cfg.get("analysis", {}).get("sku_example", "")).strip()
    if not sku:
        sku = _pick_default_sku(results_path, include_prefixes)

    regions = ["UA", "KZ", "UZ"]
    series = {}
    info = []

    for r in regions:
        row = _extract_row(results_path, r, sku)
        if row is None:
            continue

        vals = []
        for m in MONTHS:
            v = pd.to_numeric(row.get(m, 0), errors="coerce")
            vals.append(0.0 if pd.isna(v) else float(v))
        series[r] = np.array(vals, dtype=float)

        info.append({
            "region": r,
            "TOTAL": float(pd.to_numeric(row.get("TOTAL", 0), errors="coerce") or 0),
            "Stock": float(pd.to_numeric(row.get("Stock", 0), errors="coerce") or 0),
            "Month_sales": float(pd.to_numeric(row.get("Month_sales", 0), errors="coerce") or 0),
        })

    if not series:
        raise ValueError(f"SKU '{sku}' не знайдено в UA/KZ/UZ-normalized.")

    info_df = pd.DataFrame(info)
    print("\nSKU:", sku)
    print(info_df.to_string(index=False))

    fig = plt.figure()
    x = np.arange(1, 13)
    for r, y in series.items():
        plt.plot(x, y, marker="o", label=r)
    plt.xticks(x, MONTHS, rotation=45)
    plt.title(f"Продажі SKU по регіонах: {sku}")
    plt.xlabel("Місяць")
    plt.ylabel("Продажі")
    plt.legend()

    out_plot = eda_dir / "plots" / f"sku_{sku}_3regions.png"
    fig.tight_layout()
    fig.savefig(out_plot, dpi=150)
    plt.close(fig)

    means = {r: float(np.mean(y)) for r, y in series.items()}
    max_r = max(means, key=means.get)
    min_r = min(means, key=means.get)
    ratio = (means[max_r] + 1e-9) / (means[min_r] + 1e-9)

    print("\nВисновок:")
    if ratio >= 1.5:
        print(f"- Середні місячні продажі суттєво різняться: max={max_r} vs min={min_r} (≈{ratio:.2f}x).")
        print("- Це підтримує гіпотезу, що попит залежить від регіону.")
    else:
        print(f"- Різниця між регіонами невелика (≈{ratio:.2f}x). Для більш показового прикладу зміни analysis.sku_example у config.yaml.")
    print(f"- Графік: {out_plot}")
    print("OK | STEP 08 завершено")


if __name__ == "__main__":
    main()
