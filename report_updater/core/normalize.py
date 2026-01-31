from __future__ import annotations

import re
from datetime import date
from typing import Iterable

MONTHS = {
    "январь": 1,
    "января": 1,
    "февраль": 2,
    "февраля": 2,
    "март": 3,
    "марта": 3,
    "апрель": 4,
    "апреля": 4,
    "май": 5,
    "мая": 5,
    "июнь": 6,
    "июня": 6,
    "июль": 7,
    "июля": 7,
    "август": 8,
    "августа": 8,
    "сентябрь": 9,
    "сентября": 9,
    "октябрь": 10,
    "октября": 10,
    "ноябрь": 11,
    "ноября": 11,
    "декабрь": 12,
    "декабря": 12,
    "январь".upper(): 1,
    "февраль".upper(): 2,
    "март".upper(): 3,
    "апрель".upper(): 4,
    "май".upper(): 5,
    "июнь".upper(): 6,
    "июль".upper(): 7,
    "август".upper(): 8,
    "сентябрь".upper(): 9,
    "октябрь".upper(): 10,
    "ноябрь".upper(): 11,
    "декабрь".upper(): 12,
    "кыркүйөк": 9,
    "октябрь": 10,
    "ноябрь": 11,
    "декабрь": 12,
}

DATE_PATTERNS = [
    re.compile(r"\b(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})(?:\s*г\.?|\.?ж\.?)?\b"),
    re.compile(r"\b(\d{4})-(\d{2})-(\d{2})\b"),
    re.compile(r"\b(\d{4})-жыл(?:дын)?\s*(\d{1,2})[-\s]*([А-Яа-яӨөҮүҢңІіЁё]+)\s*күнү\b", re.IGNORECASE),
    re.compile(r"\b(\d{4})-жыл(?:дын)?\s*([А-Яа-яӨөҮүҢңІіЁё]+)\s*(\d{1,2})\s*күнү\b", re.IGNORECASE),
]

STOP_WORDS = [
    "токт",
    "токтог",
    "токтот",
    "токтоду",
    "токтогон",
    "кыскар",
    "кыскарт",
    "кыскарган",
    "кыскартылган",
]


def normalize_text(value: str) -> str:
    if value is None:
        return ""
    cleaned = str(value).replace("\u00a0", " ").replace("–", "-").replace("—", "-")
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.strip()


def normalize_label(value: str) -> str:
    cleaned = normalize_text(value).lower()
    cleaned = cleaned.replace("ё", "е")
    cleaned = cleaned.replace("таблица", "т").replace("табл", "т").replace("таб.", "т")
    cleaned = cleaned.replace("т.", "т")
    cleaned = cleaned.replace("ст.", "ст").replace("ст ", "ст")
    cleaned = re.sub(r"[\s\.:;,]+", "", cleaned)
    return cleaned


def normalize_article(value: str) -> str | None:
    if not value:
        return None
    text = normalize_text(value).lower()
    text = text.replace("ст.", "ст").replace("ст ", "ст").replace("бер", "бер")
    # patterns: ст.123, 123-бер 3-б.
    match = re.search(r"(?:ст|бер)?\s*(\d{1,3})(?:\s*[-–]\s*(\d{1,3}))?(?:\s*[-–]?\s*(\d+)\s*б\.)?", text)
    if not match:
        return None
    base = match.group(1)
    suffix = ""
    if match.group(2):
        suffix += f"-{match.group(2)}"
    if match.group(3):
        suffix += f" ч.{match.group(3)}"
    return f"ст.{base}{suffix}"


def parse_int(value: str | int | float | None) -> int | None:
    if value is None:
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value)
    cleaned = re.sub(r"[^0-9]", "", str(value))
    return int(cleaned) if cleaned else None


def parse_date_from_text(text: str) -> date | None:
    if not text:
        return None
    for pattern in DATE_PATTERNS:
        match = pattern.search(text)
        if not match:
            continue
        groups = match.groups()
        try:
            if pattern is DATE_PATTERNS[0]:
                day, month, year = groups
                if len(year) == 2:
                    year = f"20{year}"
                return date(int(year), int(month), int(day))
            if pattern is DATE_PATTERNS[1]:
                year, month, day = groups
                return date(int(year), int(month), int(day))
            if pattern in (DATE_PATTERNS[2], DATE_PATTERNS[3]):
                year = int(groups[0])
                if pattern is DATE_PATTERNS[2]:
                    day = int(groups[1])
                    month_name = groups[2]
                else:
                    month_name = groups[1]
                    day = int(groups[2])
                month = MONTHS.get(month_name.lower())
                if month:
                    return date(year, month, day)
        except ValueError:
            return None
    return None


def contains_any(text: str, terms: Iterable[str]) -> bool:
    lowered = normalize_text(text).lower()
    return any(term in lowered for term in terms)
