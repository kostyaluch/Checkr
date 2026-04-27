"""checkr.py — Автоматична валідація товарного фіду e-commerce.

Програма зчитує файл (CSV або Excel) зі списком товарів і знаходить
логічні розбіжності між назвою, описами та характеристиками одного й
того ж товару (наприклад, конфлікти об'єму SSD, оперативної пам'яті,
діагоналі екрана, ваги, роздільної здатності, типу матриці тощо).

Правила валідації зберігаються у validation_rules.py — їх легко
розширювати, не змінюючи основну логіку програми.

Використання з командного рядка:
    python checkr.py products.csv result.xlsx
    python checkr.py products.xlsx result.xlsx
"""

import argparse
import re
import sys
from pathlib import Path

import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from validation_rules import VALIDATION_RULES

# ---------------------------------------------------------------------------
# Константи: кольори для підсвітки у вихідному Excel-файлі
# ---------------------------------------------------------------------------

# Червоний (світлий) — для клітинок із конфліктними значеннями
CONFLICT_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

# Жовтий — для клітинки колонки "Помилки"
ERROR_FILL = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")


# ===========================================================================
# Модуль 1: Очищення тексту від HTML
# ===========================================================================


def clean_html(text: str) -> str:
    """Видаляє HTML-теги з рядка тексту.

    Використовує BeautifulSoup для надійного парсингу HTML.
    Якщо вхід не є рядком (None, число тощо) — повертає порожній рядок.

    Аргументи:
        text: Рядок, що може містити HTML-теги.

    Повертає:
        Рядок без HTML-тегів із збереженим текстовим вмістом.

    Приклад:
        clean_html("<p>Hello <b>World</b></p>")  →  "Hello World"
        clean_html(None)                          →  ""
    """
    if not isinstance(text, str):
        return ""
    if not text.strip():
        return text
    try:
        # lxml — швидший парсер, якщо встановлений
        return BeautifulSoup(text, "lxml").get_text(separator=" ").strip()
    except Exception:
        # Резервний варіант — вбудований парсер Python
        return BeautifulSoup(text, "html.parser").get_text(separator=" ").strip()


# ===========================================================================
# Модуль 2: Пошук значень пам'яті за допомогою регулярних виразів
# ===========================================================================

# Регулярний вираз для пошуку значень об'єму пам'яті.
# Підтримує:
#   Кириличні одиниці: ГБ (гігабайти), ТБ (терабайти), МБ (мегабайти)
#   Латинські одиниці: GB, TB, MB, GiB, TiB, MiB
#   Дробові значення:  1.5 ТБ, 2,5 GB
#   З пробілом та без: "512ГБ", "512 GB"
#
# Порядок альтернатив у групі одиниць важливий: більш специфічні — першими
# (TiB перед TB, GiB перед GB, MiB перед MB), щоб не захопити лише першу літеру.
_MEMORY_RE = re.compile(
    r"(?<!\d)"                              # не передує цифра (не частина більшого числа)
    r"(\d+(?:[.,]\d+)?)"                    # числова частина: ціле або дробове
    r"\s*"                                  # необов'язковий пробіл між числом та одиницею
    r"(ТБ|TiB|TB|ГБ|GiB|GB|МБ|MiB|MB)"    # одиниці виміру (порядок важливий!)
    r"(?!\w)",                              # не слідує символ слова (немає "GBps")
    re.IGNORECASE,
)

# Маппінг: нижній регістр одиниці → стандартне кириличне позначення.
# Щоб додати нову одиницю (наприклад, "ГБ/s"), додайте запис тут.
_UNIT_NORM: dict[str, str] = {
    "тб": "ТБ", "tb": "ТБ", "tib": "ТБ",
    "гб": "ГБ", "gb": "ГБ", "gib": "ГБ",
    "мб": "МБ", "mb": "МБ", "mib": "МБ",
}


def normalize_memory_value(raw: str) -> str:
    """Нормалізує одне значення пам'яті до стандартного формату «<число><одиниця>».

    Перетворює різні варіанти написання до єдиного стандарту:
      - Замінює латинські позначення на кириличні (GB → ГБ, TB → ТБ).
      - Прибирає пробіл між числом та одиницею.
      - Нормалізує роздільник дробу: кому замінює на крапку.
      - Прибирає зайві нулі після коми: 1.0 → 1, 1.50 → 1.5.
    Якщо рядок не відповідає шаблону — повертає оригінал у верхньому регістрі.

    Аргументи:
        raw: Рядок виду "512 GB", "1.5ТБ", "256мб", "2,5 TiB" тощо.

    Повертає:
        Нормалізований рядок, наприклад: "512ГБ", "1.5ТБ", "256МБ".

    Приклад:
        normalize_memory_value("512 GB")   →  "512ГБ"
        normalize_memory_value("1.5 TB")   →  "1.5ТБ"
        normalize_memory_value("2,5 ГБ")   →  "2.5ГБ"
        normalize_memory_value("unknown")  →  "UNKNOWN"
    """
    m = _MEMORY_RE.search(raw)
    if not m:
        return raw.strip().upper()

    # Нормалізуємо числову частину
    number = m.group(1).replace(",", ".")
    if "." in number:
        number = number.rstrip("0").rstrip(".")

    # Нормалізуємо одиницю: шукаємо в маппінгу, або використовуємо верхній регістр
    unit = _UNIT_NORM.get(m.group(2).lower(), m.group(2).upper())
    return f"{number}{unit}"


def extract_memory_matches(text: str) -> list[tuple[str, str]]:
    """Знаходить усі значення об'єму пам'яті у тексті.

    Повертає список пар (оригінал, нормалізоване), де:
      - оригінал      — рядок так, як знайдений у тексті (напр. "512 GB")
      - нормалізоване — стандартне представлення (напр. "512ГБ")

    Аргументи:
        text: Рядок, у якому шукаємо значення.

    Повертає:
        Список пар (str, str). Порожній список, якщо нічого не знайдено.

    Приклад:
        extract_memory_matches("Ноутбук 512 ГБ SSD та 32 GB RAM")
        →  [("512 ГБ", "512ГБ"), ("32 GB", "32ГБ")]
    """
    if not isinstance(text, str) or not text.strip():
        return []
    return [
        (f"{num}{unit}", normalize_memory_value(f"{num} {unit}"))
        for num, unit in _MEMORY_RE.findall(text)
    ]


def extract_memory_values(text: str) -> list[str]:
    """Зручна обгортка над extract_memory_matches: повертає лише нормалізовані значення.

    Аргументи:
        text: Рядок, у якому шукаємо значення.

    Повертає:
        Список нормалізованих значень, наприклад ["512ГБ", "32ГБ"].

    Приклад:
        extract_memory_values("Ноутбук 512 ГБ SSD та 32 GB RAM")  →  ["512ГБ", "32ГБ"]
        extract_memory_values("Просто текст")                      →  []
    """
    return [norm for _, norm in extract_memory_matches(text)]


# ===========================================================================
# Модуль 3: Пошук колонок у DataFrame
# ===========================================================================


def find_column(columns: list[str], hint: str) -> str | None:
    """Знаходить назву колонки за підрядком hint.

    Алгоритм пошуку (від найточнішого до найменш точного):
      1. Точний збіг повної назви колонки (нечутливий до регістру).
      2. Точний збіг базової назви — частини до крапки з комою
         (напр. знаходить "Описание;1" за точним збігом base="Описание").
      3. Базова назва починається з підказки з межею слова (\b):
         знаходить "Описание;1" за hint="Описание", але НЕ знаходить
         "Краткое описание", а також знаходить "Назва (ua)" за hint="Назва",
         але НЕ знаходить "Название".
      4. Підказка міститься у базовій назві (часткове входження, fallback):
         знаходить "Объём SSD;115411" за hint="SSD".

    Пошук нечутливий до регістру.

    Аргументи:
        columns: Список назв колонок DataFrame.
        hint:    Підрядок (або повна назва) для пошуку.

    Повертає:
        Назву першої знайденої колонки або None, якщо нічого не знайдено.

    Приклад:
        find_column(["Объём SSD;115411", "Название"], "Объём SSD")   →  "Объём SSD;115411"
        find_column(["Краткое описание", "Описание;1"], "Описание")   →  "Описание;1"
        find_column(["Название", "Назва (ua)"], "відеокарта")          →  None
    """
    hint_lower = hint.lower()

    # Крок 1: точний збіг повної назви колонки
    for col in columns:
        if col.lower() == hint_lower:
            return col

    # Крок 2: точний збіг базової назви (частина до ';')
    for col in columns:
        base = col.split(";")[0].strip().lower()
        if base == hint_lower:
            return col

    # Крок 3: базова назва починається з підказки + межа слова (\b).
    # Наприклад, hint="Описание" знаходить "Описание;1" (base="Описание"),
    # але НЕ знаходить "Краткое описание" (base="Краткое описание").
    start_pattern = re.compile(rf"^{re.escape(hint_lower)}\b", re.UNICODE)
    for col in columns:
        base = col.split(";")[0].strip().lower()
        if start_pattern.match(base):
            return col

    # Крок 4: підказка міститься у базовій назві (найширший fallback).
    # Наприклад, hint="SSD" знаходить "Объём SSD;115411" (base="Объём SSD").
    for col in columns:
        base = col.split(";")[0].strip().lower()
        if hint_lower in base:
            return col

    return None


# ===========================================================================
# Модуль 2b: Пошук значень додаткових типів характеристик
# ===========================================================================


# ---------------------------------------------------------------------------
# Діагональ екрана
# ---------------------------------------------------------------------------

# Патерн для значень діагоналі з одиницею виміру:
# Підтримує: 15.6", 15,6", 14 дюймів, 13.3 inch, 15.6 inches, 15.6″
_DIAGONAL_RE = re.compile(
    r"(?<!\d)"
    r"(\d+(?:[.,]\d+)?)"                                    # числова частина
    r"\s*"
    r"(?:\"|″|дюйм(?:ів|а|и)?|-?inch(?:es)?)"              # одиниця виміру
    r"(?!\d)",
    re.IGNORECASE | re.UNICODE,
)


def extract_screen_diagonal_matches(text: str) -> list[tuple[str, str]]:
    """Знаходить усі значення діагоналі екрана у тексті.

    Розпізнає форми: 15.6", 14 дюймів, 13.3-inch тощо.
    Повертає список пар (оригінал, нормалізоване) у форматі «15.6"».

    Аргументи:
        text: Рядок, у якому шукаємо значення.

    Повертає:
        Список пар (str, str). Порожній список, якщо нічого не знайдено.

    Приклад:
        extract_screen_diagonal_matches('Ноутбук 15.6" IPS')
        →  [('15.6"', '15.6"')]
    """
    if not isinstance(text, str) or not text.strip():
        return []
    results = []
    for m in _DIAGONAL_RE.finditer(text):
        original = m.group(0).strip()
        number = m.group(1).replace(",", ".")
        if "." in number:
            number = number.rstrip("0").rstrip(".")
        normalized = f'{number}"'
        results.append((original, normalized))
    return results


# ---------------------------------------------------------------------------
# Вага пристрою
# ---------------------------------------------------------------------------

# Патерн для значень ваги в кілограмах:
# Підтримує: 1.5 кг, 2,3 кг, 1.8kg, 2.1 KG
_WEIGHT_RE = re.compile(
    r"(?<!\d)"
    r"(\d+(?:[.,]\d+)?)"   # числова частина
    r"\s*"
    r"(кг|kg)"             # одиниця: кілограми (латиниця або кирилиця)
    r"(?!\w)",
    re.IGNORECASE | re.UNICODE,
)


def extract_weight_matches(text: str) -> list[tuple[str, str]]:
    """Знаходить усі значення ваги у тексті (кілограми).

    Розпізнає форми: 1.5 кг, 2,3 кг, 1.8kg тощо.
    Повертає список пар (оригінал, нормалізоване) у форматі «1.5кг».

    Аргументи:
        text: Рядок, у якому шукаємо значення.

    Повертає:
        Список пар (str, str). Порожній список, якщо нічого не знайдено.

    Приклад:
        extract_weight_matches("Легкий ноутбук вагою 1.5 кг")
        →  [('1.5 кг', '1.5кг')]
    """
    if not isinstance(text, str) or not text.strip():
        return []
    results = []
    for m in _WEIGHT_RE.finditer(text):
        original = m.group(0).strip()
        number = m.group(1).replace(",", ".")
        if "." in number:
            number = number.rstrip("0").rstrip(".")
        results.append((original, f"{number}кг"))
    return results


# ---------------------------------------------------------------------------
# Роздільна здатність екрана
# ---------------------------------------------------------------------------

# Патерн для числової роздільної здатності (NxM або N×M):
# Підтримує: 1920x1080, 1920×1080, 3840 x 2160
_RESOLUTION_NUM_RE = re.compile(
    r"(\d{3,4})\s*[xхх×]\s*(\d{3,4})",
    re.IGNORECASE,
)

# Псевдоніми роздільної здатності → нормалізована форма «ШxВ».
# Довші рядки мають бути раніше, щоб "Full HD" знайшлося раніше "HD".
_RESOLUTION_ALIASES: dict[str, str] = {
    "full hd": "1920x1080",
    "full-hd": "1920x1080",
    "fhd": "1920x1080",
    "ultra hd": "3840x2160",
    "ultra-hd": "3840x2160",
    "uhd": "3840x2160",
    "4k": "3840x2160",
    "qhd": "2560x1440",
    "hd+": "1600x900",
    "wxga+": "1600x900",
    "wxga": "1366x768",
    "hd": "1366x768",
}


def extract_resolution_matches(text: str) -> list[tuple[str, str]]:
    """Знаходить усі значення роздільної здатності у тексті.

    Розпізнає числові значення (1920x1080, 3840×2160) та псевдоніми
    (FHD, Full HD, 4K, QHD, HD+ тощо).
    Повертає список пар (оригінал, нормалізоване) у форматі «1920x1080».
    Довші псевдоніми мають пріоритет над коротшими (напр. "Full HD" над "HD").

    Аргументи:
        text: Рядок, у якому шукаємо значення.

    Повертає:
        Список пар (str, str). Порожній список, якщо нічого не знайдено.

    Приклад:
        extract_resolution_matches("Дисплей FHD 1920x1080")
        →  [('FHD', '1920x1080'), ('1920x1080', '1920x1080')]
    """
    if not isinstance(text, str) or not text.strip():
        return []

    results: list[tuple[str, str]] = []
    seen_norms: set[str] = set()
    matched_ranges: list[tuple[int, int]] = []  # Вже зайняті позиції
    text_lower = text.lower()

    def _overlaps(start: int, end: int) -> bool:
        return any(s < end and start < e for s, e in matched_ranges)

    # Перевіряємо псевдоніми (від довших до коротших, щоб "Full HD" знаходилось
    # раніше "HD" і попередні збіги не перекривалися)
    for alias in sorted(_RESOLUTION_ALIASES, key=len, reverse=True):
        idx = text_lower.find(alias)
        if idx >= 0:
            end = idx + len(alias)
            if not _overlaps(idx, end):
                norm = _RESOLUTION_ALIASES[alias]
                if norm not in seen_norms:
                    results.append((text[idx: end], norm))
                    seen_norms.add(norm)
                    matched_ranges.append((idx, end))

    # Шукаємо числові роздільні здатності
    for m in _RESOLUTION_NUM_RE.finditer(text):
        if not _overlaps(m.start(), m.end()):
            norm = f"{m.group(1)}x{m.group(2)}"
            if norm not in seen_norms:
                results.append((m.group(0), norm))
                seen_norms.add(norm)
                matched_ranges.append((m.start(), m.end()))

    return results


# ---------------------------------------------------------------------------
# Список допустимих значень (для текстових характеристик: матриця, GPU тощо)
# ---------------------------------------------------------------------------


def extract_value_list_matches(
    text: str, valid_values: list[str]
) -> list[tuple[str, str]]:
    """Знаходить у тексті будь-яке зі списку допустимих значень.

    Пошук нечутливий до регістру. Збіг визначається як входження рядка
    valid_values[i] у текст. Довші значення перевіряються першими.

    Аргументи:
        text:         Рядок, у якому шукаємо значення.
        valid_values: Список рядків для пошуку (наприклад, ["IPS", "TN", "VA"]).

    Повертає:
        Список пар (оригінал_у_тексті, нормалізоване_верхній_регістр).
        Порожній список, якщо нічого не знайдено.

    Приклад:
        extract_value_list_matches("Матриця IPS, тонкі рамки", ["IPS", "TN"])
        →  [('IPS', 'IPS')]
    """
    if not isinstance(text, str) or not text.strip():
        return []

    text_lower = text.lower()
    results: list[tuple[str, str]] = []
    seen: set[str] = set()

    # Довші значення перевіряємо першими (щоб "Mini-LED" знаходилось раніше "LED")
    for val in sorted(valid_values, key=len, reverse=True):
        val_lower = val.lower()
        if val_lower in seen:
            continue
        idx = text_lower.find(val_lower)
        if idx >= 0:
            results.append((text[idx: idx + len(val)], val.upper()))
            seen.add(val_lower)

    return results


# ===========================================================================
# Модуль 2c: Допоміжні функції витягування канонічного значення та диспетчер
# ===========================================================================


def _normalize_canonical(raw: str, rule: dict) -> str | None:
    """Витягує нормалізоване канонічне значення з колонки характеристики.

    Для кожного типу перевірки (checker_type) застосовує відповідний алгоритм.
    Якщо колонка характеристики містить значення без одиниці виміру (наприклад,
    "15.6" замість "15.6""), намагається витягти числову частину напряму.

    Аргументи:
        raw:  Рядок зі значенням колонки характеристики (після str().strip()).
        rule: Словник правила з VALIDATION_RULES.

    Повертає:
        Нормалізований рядок або None, якщо значення не вдалось розпізнати.
    """
    checker_type = rule.get("checker_type", "memory")

    if checker_type == "memory":
        matches = extract_memory_matches(raw)
        return matches[0][1] if matches else None

    if checker_type == "screen_diagonal":
        matches = extract_screen_diagonal_matches(raw)
        if matches:
            return matches[0][1]
        # Колонка характеристики може містити просте число: "15.6"
        bare = re.match(r"^\s*(\d+(?:[.,]\d+)?)\s*$", raw)
        if bare:
            number = bare.group(1).replace(",", ".")
            if "." in number:
                number = number.rstrip("0").rstrip(".")
            return f'{number}"'
        return None

    if checker_type == "weight":
        matches = extract_weight_matches(raw)
        if matches:
            return matches[0][1]
        # Колонка характеристики може містити просте число: "1.5" (вважаємо кг)
        bare = re.match(r"^\s*(\d+(?:[.,]\d+)?)\s*$", raw)
        if bare:
            number = bare.group(1).replace(",", ".")
            if "." in number:
                number = number.rstrip("0").rstrip(".")
            return f"{number}кг"
        return None

    if checker_type == "resolution":
        raw_lower = raw.strip().lower()
        # Перевіряємо псевдоніми
        if raw_lower in _RESOLUTION_ALIASES:
            return _RESOLUTION_ALIASES[raw_lower]
        # Перевіряємо числовий формат
        m = _RESOLUTION_NUM_RE.search(raw)
        if m:
            return f"{m.group(1)}x{m.group(2)}"
        return None

    if checker_type == "value_list":
        valid_values = rule.get("valid_values", [])
        raw_lower = raw.strip().lower()
        # Точний збіг (нечутливий до регістру)
        for val in valid_values:
            if val.lower() == raw_lower:
                return val.upper()
        # Часткове входження (для значень типу "Дискретна відеокарта")
        for val in sorted(valid_values, key=len, reverse=True):
            if val.lower() in raw_lower:
                return val.upper()
        return None

    return None


def _extract_for_rule(text: str, rule: dict) -> list[tuple[str, str]]:
    """Знаходить у тексті значення відповідного типу характеристики.

    Диспетчер: обирає функцію-екстрактор залежно від поля "checker_type" правила.

    Аргументи:
        text: Рядок для пошуку (вже очищений від HTML).
        rule: Словник правила з VALIDATION_RULES.

    Повертає:
        Список пар (оригінал, нормалізоване). Порожній список, якщо нічого не знайдено.
    """
    checker_type = rule.get("checker_type", "memory")

    if checker_type == "memory":
        return extract_memory_matches(text)
    if checker_type == "screen_diagonal":
        return extract_screen_diagonal_matches(text)
    if checker_type == "weight":
        return extract_weight_matches(text)
    if checker_type == "resolution":
        return extract_resolution_matches(text)
    if checker_type == "value_list":
        return extract_value_list_matches(text, rule.get("valid_values", []))

    return []


# ===========================================================================
# Модуль 4: Правила валідації (імпортуються з validation_rules.py)
# ===========================================================================
# VALIDATION_RULES імпортовано на початку файлу: from validation_rules import VALIDATION_RULES


# ===========================================================================
# Модуль 5: Логіка порівняння та виявлення конфліктів
# ===========================================================================


def check_conflicts(
    row: pd.Series,
    df_columns: list[str],
    rule: dict,
) -> tuple[str, list[str]]:
    """Перевіряє наявність конфліктів для одного правила в одному рядку товару.

    Алгоритм:
        1. Знаходить колонку з еталонним значенням характеристики.
        2. Витягує та нормалізує еталонне значення залежно від типу checker_type.
        3. Для кожного текстового поля (після очищення від HTML) знаходить
           значення того ж типу за допомогою відповідного екстрактора.
        4. Якщо у текстовому полі є значення, відмінні від еталонного —
           це конфлікт.
        5. Якщо текстове поле не містить жодного значення — пропускає
           (відсутність згадки не є конфліктом).

    Аргументи:
        row:        Рядок товару (pd.Series).
        df_columns: Список усіх назв колонок DataFrame.
        rule:       Словник правила з VALIDATION_RULES.

    Повертає:
        Кортеж (рядок_помилки, список_конфліктних_колонок).
        Якщо конфліктів немає — ("", []).

    Приклад:
        При Название="Ноутбук 128ГБ SSD", Характеристика SSD="512 ГБ":
        → ("Конфлікт SSD: Назва (128ГБ) != Характеристика ... (512 ГБ)",
           ["Название", "Объём SSD;115411"])
    """
    label = rule["label"]

    # Крок 1: Знаходимо колонку з еталонним значенням характеристики
    char_col: str | None = None
    for hint in rule["char_hints"]:
        char_col = find_column(df_columns, hint)
        if char_col:
            break
    if char_col is None:
        # Колонка характеристики відсутня у файлі — пропускаємо правило
        return "", []

    # Крок 2: Отримуємо та нормалізуємо еталонне значення
    raw_char = row.get(char_col, "")
    if pd.isna(raw_char) or not str(raw_char).strip():
        return "", []  # Порожня характеристика — немає з чим порівнювати

    canonical_raw = str(raw_char).strip()
    canonical_norm = _normalize_canonical(canonical_raw, rule)
    if canonical_norm is None:
        # Характеристика заповнена, але не містить розпізнаного значення
        return "", []

    # Крок 3: Перевіряємо кожне текстове поле
    conflict_parts: list[str] = []   # Частини тексту повідомлення про конфлікт
    conflict_cols: list[str] = []    # Назви колонок, де виявлено конфлікт
    seen_text_cols: set[str] = set() # Щоб не перевіряти одну й ту ж колонку двічі

    for hint in rule["text_hints"]:
        text_col = find_column(df_columns, hint)
        if text_col is None or text_col in seen_text_cols:
            continue  # Текстова колонка відсутня або вже оброблена
        seen_text_cols.add(text_col)

        raw_text = row.get(text_col, "")
        if pd.isna(raw_text) or not str(raw_text).strip():
            continue  # Порожнє поле — не конфлікт

        # Очищаємо HTML та шукаємо значення відповідного типу
        clean_text = clean_html(str(raw_text))
        matches = _extract_for_rule(clean_text, rule)

        if not matches:
            continue  # Поле не містить згадки про цю характеристику — не конфлікт

        # Крок 4: Конфлікт є, якщо канонічного значення серед знайдених немає
        found_norms = [norm for _, norm in matches]
        if canonical_norm not in found_norms:
            found_raws = [raw for raw, _ in matches]
            conflict_parts.append(f"{text_col} ({', '.join(found_raws)})")
            conflict_cols.append(text_col)

    # Крок 5: Формуємо повідомлення про конфлікт
    if conflict_parts:
        char_part = f"Характеристика {char_col} ({canonical_raw})"
        error_msg = (
            f"Конфлікт {label}: "
            f"{' != '.join(conflict_parts)} != {char_part}"
        )
        conflict_cols.append(char_col)
        return error_msg, conflict_cols

    return "", []


# ===========================================================================
# Модуль 6: Читання вхідного файлу
# ===========================================================================


def read_input_file(filepath: str | Path) -> pd.DataFrame:
    """Зчитує вхідний файл (CSV або Excel) у pandas DataFrame.

    Підтримує формати: .csv, .xlsx, .xls.
    Усі значення читаються як рядки (dtype=str), щоб уникнути небажаних
    автоматичних перетворень типів.
    Для CSV автоматично пробує роздільники: кома, крапка з комою, табуляція;
    обирає той, що дає найбільшу кількість колонок.

    Аргументи:
        filepath: Шлях до вхідного файлу.

    Повертає:
        pandas DataFrame із даними файлу.

    Виключення:
        FileNotFoundError: Якщо файл не знайдено за вказаним шляхом.
        ValueError:        Якщо формат файлу не підтримується.
    """
    path = Path(filepath)
    if not path.exists():
        raise FileNotFoundError(f"Файл не знайдено: {filepath}")

    suffix = path.suffix.lower()

    if suffix in (".xlsx", ".xls"):
        return pd.read_excel(path, dtype=str)

    if suffix == ".csv":
        best_df: pd.DataFrame | None = None
        for sep in (",", ";", "\t"):
            try:
                df = pd.read_csv(path, sep=sep, dtype=str, encoding="utf-8-sig")
                if best_df is None or len(df.columns) > len(best_df.columns):
                    best_df = df
            except Exception:
                continue
        if best_df is not None:
            return best_df
        # Останній варіант: автовизначення роздільника через pandas
        return pd.read_csv(
            path, sep=None, engine="python", dtype=str, encoding="utf-8-sig"
        )

    raise ValueError(
        f"Непідтримуваний формат файлу: '{suffix}'. Підтримуються: .csv, .xlsx, .xls"
    )


# ===========================================================================
# Модуль 7: Збереження результатів у Excel із підсвіткою
# ===========================================================================


def _save_with_highlights(
    df: pd.DataFrame,
    output_file: str,
    conflict_cells: dict[tuple[int, str], bool],
) -> None:
    """Зберігає DataFrame у .xlsx та підсвічує конфліктні клітинки.

    Алгоритм:
        1. Зберігає DataFrame через pandas (engine=openpyxl) без форматування.
        2. Відкриває збережений файл через openpyxl для ручного форматування.
        3. Задає червоний фон для клітинок із конфліктними значеннями.
        4. Задає жовтий фон для клітинок колонки "Помилки".
        5. Зберігає відформатований файл.

    Примітка:
        Звичайний pandas не підтримує форматування клітинок. Для підсвітки
        кольором необхідний openpyxl. Спочатку зберігаємо через pandas
        (швидко), потім відкриваємо через openpyxl лише для форматування.

    Аргументи:
        df:             DataFrame для збереження.
        output_file:    Шлях до вихідного .xlsx файлу.
        conflict_cells: Словник {(df_row_index, col_name): True} —
                        координати конфліктних клітинок.
    """
    # Крок 1: Зберігаємо без форматування
    df.to_excel(output_file, index=False, engine="openpyxl")

    # Крок 2: Відкриваємо для форматування
    wb = load_workbook(output_file)
    ws = wb.active

    # Будуємо маппінг: назва_колонки → номер_колонки Excel (1-based)
    header_to_col: dict[str, int] = {
        cell.value: cell.column for cell in ws[1] if cell.value is not None
    }

    # Будуємо маппінг: індекс рядка DataFrame → номер рядка Excel
    # (рядок 1 — заголовки, дані починаються з рядка 2)
    df_idx_to_excel_row: dict[int, int] = {
        df_idx: excel_row
        for excel_row, df_idx in enumerate(df.index, start=2)
    }

    # Крок 3: Підсвічуємо кожну конфліктну клітинку
    for (df_idx, col_name) in conflict_cells:
        excel_row = df_idx_to_excel_row.get(df_idx)
        excel_col = header_to_col.get(col_name)
        if excel_row is None or excel_col is None:
            continue

        cell = ws.cell(row=excel_row, column=excel_col)
        if col_name == "Помилки":
            cell.fill = ERROR_FILL    # Жовтий — для колонки з описом помилок
        else:
            cell.fill = CONFLICT_FILL  # Червоний — для полів із конфліктами

    # Крок 4: Зберігаємо відформатований файл
    wb.save(output_file)


# ===========================================================================
# Модуль 8: Головна функція валідації
# ===========================================================================


def validate_feed(input_file: str, output_file: str) -> pd.DataFrame:
    """Головна функція: зчитує товарний фід, валідує та зберігає результат.

    Алгоритм:
        1. Зчитує вхідний файл у DataFrame.
        2. Визначає, які правила застосовуються до цього файлу
           (яких колонок характеристик є у файлі).
        3. Для кожного рядка та кожного активного правила перевіряє
           наявність конфліктів між текстовими полями та характеристикою.
        4. Записує знайдені помилки у нову колонку "Помилки".
        5. Зберігає result.xlsx із підсвіченими конфліктними клітинками
           (червоний — конфліктні поля, жовтий — опис помилки).
        6. Виводить повний звіт: які характеристики перевірено,
           скільки конфліктів знайдено по кожній.

    Аргументи:
        input_file:  Шлях до вхідного файлу (.csv, .xlsx, .xls).
        output_file: Шлях до вихідного файлу (.xlsx).

    Повертає:
        DataFrame із доданою колонкою "Помилки".

    Виключення:
        FileNotFoundError: Якщо input_file не знайдено.
        ValueError:        Якщо формат файлу не підтримується.

    Приклад:
        df = validate_feed("products.csv", "result.xlsx")
    """
    print(f"Читаємо файл: {input_file}")
    df = read_input_file(input_file)
    print(f"Зчитано {len(df)} рядків, {len(df.columns)} колонок.")
    df_columns = list(df.columns)

    # Додаємо колонку для помилок, якщо її ще немає
    if "Помилки" not in df.columns:
        df["Помилки"] = ""

    # Словник координат конфліктних клітинок:
    # ключ = (df_row_index, column_name), значення = True (для унікальності)
    conflict_cells: dict[tuple[int, str], bool] = {}

    # Лічильники для повного звіту: label → {"applicable": bool, "conflicts": int}
    rule_stats: dict[str, dict] = {}
    for rule in VALIDATION_RULES:
        char_col = None
        for hint in rule["char_hints"]:
            char_col = find_column(df_columns, hint)
            if char_col:
                break
        rule_stats[rule["label"]] = {
            "applicable": char_col is not None,
            "conflicts": 0,
        }

    # Обробляємо кожен рядок товару.
    # Примітка: iterrows() зручний для складної рядкової логіки, але для дуже
    # великих фідів (>100k рядків) розгляньте можливість рефакторингу на df.apply().
    for idx, row in df.iterrows():
        row_errors: list[str] = []

        for rule in VALIDATION_RULES:
            try:
                error_msg, conflict_cols = check_conflicts(row, df_columns, rule)
            except Exception as exc:
                # Не зупиняємося при помилці в одному правилі — обробляємо далі
                print(
                    f"Попередження: помилка при перевірці '{rule['label']}' "
                    f"у рядку {idx}: {exc}"
                )
                continue

            if error_msg:
                row_errors.append(error_msg)
                rule_stats[rule["label"]]["conflicts"] += 1
                for col in conflict_cols:
                    conflict_cells[(idx, col)] = True

        if row_errors:
            df.at[idx, "Помилки"] = " | ".join(row_errors)
            conflict_cells[(idx, "Помилки")] = True

    # Підраховуємо та виводимо загальні результати
    errors_count = (df["Помилки"] != "").sum()
    print(f"\nЗнайдено конфліктів: {errors_count} рядків із {len(df)}.")

    # Виводимо повний звіт по характеристиках
    print("\nЗвіт по характеристиках:")
    for label, stats in rule_stats.items():
        if not stats["applicable"]:
            print(f"  —  {label}: колонка відсутня у файлі (пропущено)")
        elif stats["conflicts"] == 0:
            print(f"  ✓  {label}: конфліктів не знайдено")
        else:
            print(f"  ✗  {label}: знайдено {stats['conflicts']} конфліктів")

    if conflict_cells:
        print("\nКоординати конфліктних клітинок (рядок, колонка):")
        for row_idx, col_name in sorted(conflict_cells.keys()):
            print(f"  Рядок {row_idx + 2} (Excel), Колонка «{col_name}»")

    # Зберігаємо у Excel із підсвіткою
    _save_with_highlights(df, output_file, conflict_cells)
    print(f"\nРезультат збережено: {output_file}")

    return df


# ===========================================================================
# CLI-точка входу
# ===========================================================================


def main() -> None:
    """Запуск програми з командного рядка.

    Використання:
        python checkr.py <вхідний_файл> [<вихідний_файл>]

    Приклади:
        python checkr.py products.csv result.xlsx
        python checkr.py products.xlsx result.xlsx
        python checkr.py products.csv  # створить products_result.xlsx
    """
    parser = argparse.ArgumentParser(
        prog="checkr",
        description=(
            "Валідація товарного фіду e-commerce: "
            "пошук конфліктів між текстовими полями та характеристиками товару."
        ),
    )
    parser.add_argument("input", help="Вхідний файл (.csv, .xlsx або .xls)")
    parser.add_argument(
        "output",
        nargs="?",
        help="Вихідний файл (.xlsx). Якщо не вказано, створюється <input>_result.xlsx",
    )
    args = parser.parse_args()

    # Якщо output не вказано, створити автоматичне ім'я
    if args.output is None:
        input_path = Path(args.input)
        output_name = f"{input_path.stem}_result.xlsx"
        args.output = str(input_path.parent / output_name)

    try:
        validate_feed(args.input, args.output)
    except FileNotFoundError as exc:
        print(f"Помилка: {exc}", file=sys.stderr)
        sys.exit(1)
    except ValueError as exc:
        print(f"Помилка формату: {exc}", file=sys.stderr)
        sys.exit(1)
    except Exception as exc:
        print(f"Несподівана помилка: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
