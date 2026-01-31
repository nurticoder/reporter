from core.normalize import normalize_article, normalize_label, parse_date_from_text


def test_normalize_label_variants() -> None:
    assert normalize_label("т. 2.1") == normalize_label("Т 2.1")
    assert normalize_label("таблица 2.1") == normalize_label("т.2.1")


def test_parse_date_from_text() -> None:
    parsed = parse_date_from_text("16.10.25.ж.")
    assert parsed is not None
    assert (parsed.year, parsed.month, parsed.day) == (2025, 10, 16)


def test_normalize_article() -> None:
    assert normalize_article("ст. 209") == "ст.209"
    assert normalize_article("209-бер 3-б.") == "ст.209"
