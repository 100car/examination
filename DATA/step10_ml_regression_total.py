"""
STEP 10 | ML (supervised): прогноз TOTAL за профілем продажів (RandomForest)

Пояснення:
- маємо 1 рік даних (12 місяців), тому робимо показовий ML-приклад:
  модель вчиться на векторах продажів і прогнозує TOTAL.
- Це демонструє pipeline ML: split, метрики, графіки, збереження артефактів.

Результати:
- <ml_dir>/model_rf_total.joblib
- <ml_dir>/plots/pred_vs_true.png
- <ml_dir>/plots/feature_importance.png

Запуск:
!python DATA/step10_ml_regression_total.py
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List

import numpy as np
import pandas as pd
import yaml
import matplotlib.pyplot as plt

from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_absolute_error, r2_score
from sklearn.ensemble import RandomForestRegressor

import joblib


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
    ml_dir = Path(paths.get("ml_dir", default_cfg["paths"]["ml_dir"]))
    ml_dir.mkdir(parents=True, exist_ok=True)
    (ml_dir / "plots").mkdir(parents=True, exist_ok=True)
    return results_path, ml_dir


def _include_prefixes(cfg: Dict) -> List[str]:
    filt = cfg.get("filters", {}) if isinstance(cfg.get("filters", {}), dict) else {}
    return list(filt.get("include_prefixes", default_cfg["filters"]["include_prefixes"]))


def _filter_by_prefix(symbols: pd.Series, include_prefixes: List[str]) -> pd.Series:
    s = symbols.astype(str).fillna("")
    mask = False
    for p in include_prefixes:
        mask = mask | s.str.startswith(str(p))
    return mask


def _load_region(results_path: Path, region: str) -> pd.DataFrame:
    df = pd.read_excel(results_path, sheet_name=f"{region}-normalized", engine="openpyxl")
    if "Symbol_normalized" not in df.columns:
        raise ValueError(f"{region}-normalized не містить колонку Symbol_normalized.")
    cols = ["Symbol_normalized"] + [c for c in MONTHS if c in df.columns]
    for c in ["Stock", "TOTAL", "Month_sales", "Mediana"]:
        if c in df.columns:
            cols.append(c)
    df = df[cols].copy()
    for c in cols[1:]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return df


def _build_dataset(results_path: Path, include_prefixes: List[str]) -> pd.DataFrame:
    regions = ["UA", "KZ", "UZ"]
    dfs = []
    for r in regions:
        df = _load_region(results_path, r)
        mask = _filter_by_prefix(df["Symbol_normalized"], include_prefixes)
        dfs.append(df.loc[mask].copy())
    all_df = pd.concat(dfs, ignore_index=True)

    numeric_cols = [c for c in all_df.columns if c != "Symbol_normalized"]
    combined = all_df.groupby("Symbol_normalized", as_index=False)[numeric_cols].sum()
    return combined


def main() -> None:
    cfg = load_cfg()
    results_path, ml_dir = _paths_from_cfg(cfg)
    include_prefixes = _include_prefixes(cfg)

    _print_banner("STEP 10 | ML: RandomForestRegressor (прогноз TOTAL)")

    if not results_path.exists():
        raise FileNotFoundError(f"results.xlsx не знайдено: {results_path}")

    data = _build_dataset(results_path, include_prefixes)
    if "TOTAL" not in data.columns:
        raise ValueError("У датасеті немає TOTAL. Перевір, що *-normalized містять колонку TOTAL.")

    feature_cols = [c for c in MONTHS if c in data.columns]
    for c in ["Stock", "Month_sales"]:
        if c in data.columns:
            feature_cols.append(c)

    X = data[feature_cols].values.astype(float)
    y = data["TOTAL"].values.astype(float)

    if X.shape[0] < 30:
        print("WRN | STEP 10 | Дуже мало спостережень для стабільного ML. Результати можуть бути шумними.")
        # але все одно покажемо приклад

    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42
    )

    # Baseline: прогноз середнім
    y_pred_base = np.full_like(y_test, fill_value=float(np.mean(y_train)))
    mae_base = float(mean_absolute_error(y_test, y_pred_base))
    r2_base = float(r2_score(y_test, y_pred_base))

    # RandomForest
    model = RandomForestRegressor(
        n_estimators=400,
        random_state=42,
        n_jobs=-1,
        min_samples_leaf=1,
    )
    model.fit(X_train, y_train)
    y_pred = model.predict(X_test)

    mae = float(mean_absolute_error(y_test, y_pred))
    r2 = float(r2_score(y_test, y_pred))

    print("ML (RandomForestRegressor) — прогноз TOTAL за профілем продажів (PR-)")
    print(f"- Baseline (mean): MAE={mae_base:.2f}, R2={r2_base:.3f}")
    print(f"- Model (RF):      MAE={mae:.2f}, R2={r2:.3f}")

    # pred vs true
    fig = plt.figure()
    plt.scatter(y_test, y_pred)
    plt.title("Predicted vs True (TOTAL)")
    plt.xlabel("True TOTAL")
    plt.ylabel("Predicted TOTAL")
    fig.tight_layout()
    fig.savefig(ml_dir / "plots" / "pred_vs_true.png", dpi=150)
    plt.close(fig)

    # feature importance
    importances = model.feature_importances_
    idx = np.argsort(importances)[::-1]
    top_n = min(15, len(feature_cols))
    fig = plt.figure()
    plt.bar([feature_cols[i] for i in idx[:top_n]], importances[idx[:top_n]])
    plt.xticks(rotation=45, ha="right")
    plt.title("Feature importance (top)")
    fig.tight_layout()
    fig.savefig(ml_dir / "plots" / "feature_importance.png", dpi=150)
    plt.close(fig)

    out_model = ml_dir / "model_rf_total.joblib"
    joblib.dump({"model": model, "feature_cols": feature_cols}, out_model)

    print("OK | STEP 10 завершено")
    print(f"- model: {out_model}")
    print(f"- plots: {ml_dir / 'plots'}")


if __name__ == "__main__":
    main()
