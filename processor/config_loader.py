from __future__ import annotations

import json
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CONFIG_DIR = PROJECT_ROOT / "config"


def load_json(name: str) -> dict:
    path = CONFIG_DIR / name
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def load_required_metrics() -> list[str]:
    data = load_json("requiredMetrics.json")
    return list(data)


def load_metric_dictionary() -> dict:
    return load_json("metricDictionary.json")


def load_excel_map() -> dict:
    return load_json("excelMap.json")


def load_article_map() -> dict:
    return load_json("articleMap.json")


def load_cross_checks() -> dict:
    return load_json("crossChecks.json")

