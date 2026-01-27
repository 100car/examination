"""
STEP 09 | K-means кластеризація SKU (PR-) на основі *-normalized

Вимоги:
- визначити k методом "ліктя" (inertia) та silhouette
- навчити KMeans
- додати cluster_id у результат

Результати:
- <ml_dir>/clusters_pr.csv
- <ml_dir>/plots/elbow.png
- <ml_dir>/plots/silhouette.png
- лист PR_CLUSTERS у results.xlsx

Запуск:
!python DATA/step09_kmeans_clustering.py
"""

from __future__ import annotations

from pathlib import Path
from typing import Dict, List

import numpy as np
import pandas as pd
import yaml
import matplotlib.pyplot as plt

from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import silhouette_score


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


def _build_combined_dataset(results_path: Path, include_prefixes: List[str]) -> pd.DataFrame:
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

    _print_banner("STEP 09 | KMeans кластеризація SKU (PR-)")

    if not results_path.exists():
        raise FileNotFoundError(f"results.xlsx не знайдено: {results_path}")

    data = _build_combined_dataset(results_path, include_prefixes)

    feature_cols = [c for c in MONTHS if c in data.columns]
    for c in ["Stock", "Month_sales", "TOTAL"]:
        if c in data.columns:
            feature_cols.append(c)

    if data.shape[0] < 10:
        print("WRN | STEP 09 | Дуже мало SKU для кластеризації. Перевір filters.include_prefixes.")
        return

    X = data[feature_cols].values.astype(float)
    Xs = StandardScaler().fit_transform(X)

    kmin, kmax = 2, 10
    ks = list(range(kmin, kmax + 1))
    inertias = []
    silhouettes = []

    for k in ks:
        km = KMeans(n_clusters=k, random_state=42, n_init="auto")
        labels = km.fit_predict(Xs)
        inertias.append(float(km.inertia_))
        if len(set(labels)) > 1 and Xs.shape[0] > k:
            silhouettes.append(float(silhouette_score(Xs, labels)))
        else:
            silhouettes.append(float("nan"))

    sil_arr = np.array(silhouettes, dtype=float)
    valid = ~np.isnan(sil_arr)
    best_k = int(np.array(ks)[valid][np.argmax(sil_arr[valid])]) if valid.any() else 3
    print(f"Обрано k={best_k} (максимум silhouette).")

    # elbow
    fig = plt.figure()
    plt.plot(ks, inertias, marker="o")
    plt.title("Elbow: Inertia vs k")
    plt.xlabel("k")
    plt.ylabel("Inertia")
    fig.tight_layout()
    fig.savefig(ml_dir / "plots" / "elbow.png", dpi=150)
    plt.show()
    plt.close(fig)

    # silhouette
    fig = plt.figure()
    plt.plot(ks, silhouettes, marker="o")
    plt.title("Silhouette score vs k")
    plt.xlabel("k")
    plt.ylabel("Silhouette")
    fig.tight_layout()
    fig.savefig(ml_dir / "plots" / "silhouette.png", dpi=150)
    plt.show()
    plt.close(fig)

    # final kmeans
    km = KMeans(n_clusters=best_k, random_state=42, n_init="auto")
    data["cluster_id"] = km.fit_predict(Xs).astype(int)

    out_csv = ml_dir / "clusters_pr.csv"
    data.to_csv(out_csv, index=False, encoding="utf-8")

    with pd.ExcelWriter(results_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        data.to_excel(writer, sheet_name="PR_CLUSTERS", index=False)

    print("OK | STEP 09 завершено")
    print(f"- CSV:   {out_csv}")
    print(f"- sheet: PR_CLUSTERS у {results_path}")
    print(f"- plots: {ml_dir / 'plots'}")


if __name__ == "__main__":
    main()
