"""Тести для checkr.py — автоматична валідація товарного фіду e-commerce.

Запуск:
    pytest tests/test_checkr.py -v
"""

import os
import tempfile
from pathlib import Path

import pandas as pd
import pytest

from checkr import (
    VALIDATION_RULES,
    check_conflicts,
    clean_html,
    extract_memory_matches,
    extract_memory_values,
    find_column,
    normalize_memory_value,
    read_input_file,
    validate_feed,
)


# ===========================================================================
# Тести Модуля 1: clean_html
# ===========================================================================


class TestCleanHtml:
    def test_removes_paragraph_tags(self):
        result = clean_html("<p>Hello <b>World</b></p>")
        assert "Hello" in result
        assert "World" in result
        assert "<" not in result

    def test_removes_all_html_tags(self):
        result = clean_html("<div><h1>Title</h1><p>Body</p></div>")
        assert "<" not in result
        assert "Title" in result
        assert "Body" in result

    def test_returns_empty_for_none(self):
        assert clean_html(None) == ""

    def test_returns_empty_for_integer(self):
        assert clean_html(42) == ""

    def test_returns_empty_for_float(self):
        assert clean_html(3.14) == ""

    def test_returns_unchanged_plain_text(self):
        assert clean_html("plain text") == "plain text"

    def test_returns_empty_for_empty_string(self):
        # Порожній рядок повертається як є (без змін)
        assert clean_html("") == ""

    def test_extracts_text_from_strong_tag(self):
        result = clean_html("Ноутбук із <strong>512 ГБ</strong> SSD")
        assert "512 ГБ" in result

    def test_handles_nested_tags(self):
        html = "<ul><li>512 ГБ SSD</li><li>16 GB RAM</li></ul>"
        result = clean_html(html)
        assert "512 ГБ SSD" in result
        assert "16 GB RAM" in result


# ===========================================================================
# Тести Модуля 2: normalize_memory_value
# ===========================================================================


class TestNormalizeMemoryValue:
    # Гігабайти: різні варіанти написання
    def test_gb_latin_uppercase(self):
        assert normalize_memory_value("512 GB") == "512ГБ"

    def test_gb_latin_no_space(self):
        assert normalize_memory_value("512GB") == "512ГБ"

    def test_gb_cyrillic_uppercase(self):
        assert normalize_memory_value("512 ГБ") == "512ГБ"

    def test_gb_cyrillic_no_space(self):
        assert normalize_memory_value("512ГБ") == "512ГБ"

    def test_gib(self):
        assert normalize_memory_value("512 GiB") == "512ГБ"

    # Терабайти: різні варіанти написання
    def test_tb_latin_uppercase(self):
        assert normalize_memory_value("1 TB") == "1ТБ"

    def test_tb_latin_no_space(self):
        assert normalize_memory_value("1TB") == "1ТБ"

    def test_tb_cyrillic_uppercase(self):
        assert normalize_memory_value("1 ТБ") == "1ТБ"

    def test_tb_cyrillic_no_space(self):
        assert normalize_memory_value("1ТБ") == "1ТБ"

    def test_tib(self):
        assert normalize_memory_value("2 TiB") == "2ТБ"

    # Мегабайти
    def test_mb_latin(self):
        assert normalize_memory_value("256 MB") == "256МБ"

    def test_mb_cyrillic(self):
        assert normalize_memory_value("256МБ") == "256МБ"

    def test_mib(self):
        assert normalize_memory_value("256 MiB") == "256МБ"

    # Дробові значення
    def test_decimal_dot(self):
        assert normalize_memory_value("1.5 TB") == "1.5ТБ"

    def test_decimal_comma(self):
        assert normalize_memory_value("2,5 ГБ") == "2.5ГБ"

    def test_decimal_trailing_zero_removed(self):
        assert normalize_memory_value("1.0 TB") == "1ТБ"

    def test_decimal_trailing_zeros_removed(self):
        assert normalize_memory_value("1.50 TB") == "1.5ТБ"

    # Регістр
    def test_lowercase_gb(self):
        assert normalize_memory_value("512 gb") == "512ГБ"

    def test_lowercase_tb(self):
        assert normalize_memory_value("1 tb") == "1ТБ"

    # Невідоме значення
    def test_unknown_returns_uppercase(self):
        result = normalize_memory_value("unknown")
        assert result == "UNKNOWN"


# ===========================================================================
# Тести Модуля 2: extract_memory_matches та extract_memory_values
# ===========================================================================


class TestExtractMemoryMatches:
    def test_single_value(self):
        matches = extract_memory_matches("Ноутбук з 512 ГБ SSD")
        assert len(matches) == 1
        raw, norm = matches[0]
        assert norm == "512ГБ"

    def test_multiple_values(self):
        matches = extract_memory_matches("512 ГБ SSD та 32 GB RAM")
        norms = [n for _, n in matches]
        assert "512ГБ" in norms
        assert "32ГБ" in norms

    def test_empty_string(self):
        assert extract_memory_matches("") == []

    def test_none_input(self):
        assert extract_memory_matches(None) == []

    def test_whitespace_only(self):
        assert extract_memory_matches("   ") == []

    def test_no_memory_values(self):
        assert extract_memory_matches("Просто текст без пам'яті") == []

    def test_returns_original_and_normalized(self):
        matches = extract_memory_matches("512GB")
        assert len(matches) == 1
        raw, norm = matches[0]
        assert norm == "512ГБ"

    def test_does_not_match_partial_number(self):
        # "51024GB" не повинно збігатися як "1024GB" після lookbehind
        matches = extract_memory_matches("51024GB")
        # Може збігатися як "51024GB" або не збігатися зовсім — головне, не "1024GB"
        norms = [n for _, n in matches]
        assert "1024ГБ" not in norms


class TestExtractMemoryValues:
    def test_returns_only_normalized(self):
        result = extract_memory_values("Ноутбук 512 ГБ SSD та 32 GB RAM")
        assert result == ["512ГБ", "32ГБ"]

    def test_empty_input(self):
        assert extract_memory_values("") == []

    def test_none_input(self):
        assert extract_memory_values(None) == []

    def test_no_match(self):
        assert extract_memory_values("Просто текст") == []


# ===========================================================================
# Тести Модуля 3: find_column
# ===========================================================================


class TestFindColumn:
    COLS = ["Название", "Назва (ua)", "Объём SSD;115411",
            "Объем установленной оперативной памяти;20863", "Описание;1",
            "Краткое описание"]

    def test_exact_match(self):
        assert find_column(self.COLS, "Название") == "Название"

    def test_partial_match_ssd(self):
        assert find_column(self.COLS, "Объём SSD") == "Объём SSD;115411"

    def test_partial_match_ram(self):
        result = find_column(self.COLS, "Объем установленной оперативной памяти")
        assert result == "Объем установленной оперативной памяти;20863"

    def test_finds_opisanie_not_kratkoe(self):
        # "Описание" має знаходити "Описание;1", а НЕ "Краткое описание"
        assert find_column(self.COLS, "Описание") == "Описание;1"

    def test_finds_nazva_ua_not_nazvanie(self):
        # "Назва" має знаходити "Назва (ua)", а НЕ "Название"
        assert find_column(self.COLS, "Назва") == "Назва (ua)"

    def test_not_found(self):
        assert find_column(self.COLS, "Відеокарта") is None

    def test_case_insensitive_exact(self):
        assert find_column(self.COLS, "название") == "Название"

    def test_case_insensitive_partial(self):
        assert find_column(self.COLS, "объём ssd") == "Объём SSD;115411"

    def test_empty_columns(self):
        assert find_column([], "Назва") is None

    def test_prefers_exact_over_partial(self):
        cols = ["SSD", "Объём SSD;115411"]
        # "SSD" — точний збіг, тому його повернути першим
        assert find_column(cols, "SSD") == "SSD"

    def test_ssd_found_via_content_match(self):
        # hint="SSD" знаходить "Объём SSD;115411" через часткове входження у base
        cols = ["Объём SSD;115411"]
        assert find_column(cols, "SSD") == "Объём SSD;115411"


# ===========================================================================
# Тести Модуля 5: check_conflicts
# ===========================================================================


class TestCheckConflicts:
    """Тести логіки виявлення конфліктів."""

    def _ssd_rule(self) -> dict:
        return next(r for r in VALIDATION_RULES if r["label"] == "SSD")

    def _ram_rule(self) -> dict:
        return next(r for r in VALIDATION_RULES if r["label"] == "RAM")

    # --- Тести без конфлікту ---

    def test_no_conflict_matching_values(self):
        row = pd.Series({
            "Название": "Ноутбук 512 ГБ SSD",
            "Краткое описание": "Ноутбук із 512 ГБ SSD",
            "Объём SSD;115411": "512 ГБ",
        })
        error, cols = check_conflicts(row, list(row.index), self._ssd_rule())
        assert error == ""
        assert cols == []

    def test_no_conflict_text_has_no_memory_value(self):
        # Текст не згадує пам'ять — це не конфлікт
        row = pd.Series({
            "Название": "Ноутбук ASUS VivoBook",
            "Краткое описание": "Потужний ноутбук",
            "Объём SSD;115411": "512 ГБ",
        })
        error, cols = check_conflicts(row, list(row.index), self._ssd_rule())
        assert error == ""
        assert cols == []

    def test_no_conflict_empty_characteristic(self):
        row = pd.Series({
            "Название": "Ноутбук 512 ГБ SSD",
            "Объём SSD;115411": "",
        })
        error, cols = check_conflicts(row, list(row.index), self._ssd_rule())
        assert error == ""
        assert cols == []

    def test_no_conflict_missing_char_column(self):
        # Колонка характеристики відсутня — не конфлікт
        row = pd.Series({"Название": "Ноутбук 512 ГБ SSD"})
        error, cols = check_conflicts(row, list(row.index), self._ssd_rule())
        assert error == ""
        assert cols == []

    def test_no_conflict_nan_characteristic(self):
        row = pd.Series({
            "Название": "Ноутбук 512 ГБ SSD",
            "Объём SSD;115411": float("nan"),
        })
        error, cols = check_conflicts(row, list(row.index), self._ssd_rule())
        assert error == ""
        assert cols == []

    # --- Тести з конфліктом ---

    def test_conflict_in_name(self):
        # Назва: 128ГБ, Характеристика: 512ГБ
        row = pd.Series({
            "Название": "Ноутбук 128 ГБ SSD",
            "Объём SSD;115411": "512 ГБ",
        })
        error, cols = check_conflicts(row, list(row.index), self._ssd_rule())
        assert error != ""
        assert "128ГБ" in error or "128 ГБ" in error
        assert "512ГБ" in error or "512 ГБ" in error
        assert "Название" in cols
        assert "Объём SSD;115411" in cols

    def test_conflict_in_description(self):
        # Опис: 256ГБ, Характеристика: 512ГБ
        row = pd.Series({
            "Краткое описание": "Ноутбук із 256 ГБ SSD",
            "Объём SSD;115411": "512 ГБ",
        })
        error, cols = check_conflicts(row, list(row.index), self._ssd_rule())
        assert error != ""
        assert "256ГБ" in error or "256 ГБ" in error

    def test_conflict_in_html_description(self):
        # Опис з HTML-тегами: виявити конфлікт після очищення
        row = pd.Series({
            "Описание;1": "<p>Ноутбук із <b>128 ГБ</b> SSD</p>",
            "Объём SSD;115411": "512 ГБ",
        })
        error, cols = check_conflicts(row, list(row.index), self._ssd_rule())
        assert error != ""
        assert "Конфлікт SSD" in error

    def test_conflict_mixed_notation(self):
        # Назва: "512 GB" (латиниця), Характеристика: "512 ГБ" (кирилиця) — НЕ конфлікт
        row = pd.Series({
            "Название": "Ноутбук 512 GB SSD",
            "Объём SSD;115411": "512 ГБ",
        })
        error, cols = check_conflicts(row, list(row.index), self._ssd_rule())
        # Обидва нормалізуються до "512ГБ" — конфлікту немає
        assert error == ""

    def test_conflict_error_message_format(self):
        row = pd.Series({
            "Название": "Ноутбук 128 ГБ SSD",
            "Объём SSD;115411": "512 ГБ",
        })
        error, _ = check_conflicts(row, list(row.index), self._ssd_rule())
        assert "Конфлікт SSD" in error
        assert "Характеристика" in error

    def test_conflict_returns_all_conflicting_columns(self):
        # Конфліктні значення в обох текстових полях
        row = pd.Series({
            "Название": "Ноутбук 128 ГБ SSD",
            "Краткое описание": "Ноутбук із 256 ГБ SSD",
            "Объём SSD;115411": "512 ГБ",
        })
        error, cols = check_conflicts(row, list(row.index), self._ssd_rule())
        assert "Название" in cols
        assert "Краткое описание" in cols
        assert "Объём SSD;115411" in cols

    def test_conflict_ram_rule(self):
        row = pd.Series({
            "Название": "Ноутбук 4 GB RAM",
            "Объем установленной оперативной памяти;20863": "32 GB",
        })
        error, cols = check_conflicts(row, list(row.index), self._ram_rule())
        assert error != ""
        assert "Конфлікт RAM" in error


# ===========================================================================
# Тести Модуля 6: read_input_file
# ===========================================================================


class TestReadInputFile:
    def test_reads_csv_with_comma_separator(self, tmp_path):
        csv_file = tmp_path / "test.csv"
        csv_file.write_text(
            "Название,Объём SSD;115411\n"
            "Ноутбук 512 ГБ,512 ГБ\n",
            encoding="utf-8"
        )
        df = read_input_file(str(csv_file))
        assert len(df) == 1
        assert "Название" in df.columns
        assert "Объём SSD;115411" in df.columns

    def test_reads_csv_with_semicolon_separator(self, tmp_path):
        csv_file = tmp_path / "test.csv"
        csv_file.write_text(
            "Название;Объём SSD\n"
            "Ноутбук 512 ГБ;512 ГБ\n",
            encoding="utf-8"
        )
        df = read_input_file(str(csv_file))
        assert len(df) == 1
        assert "Название" in df.columns

    def test_reads_excel(self, tmp_path):
        xlsx_file = tmp_path / "test.xlsx"
        df_orig = pd.DataFrame({
            "Название": ["Ноутбук"],
            "Объём SSD;115411": ["512 ГБ"],
        })
        df_orig.to_excel(str(xlsx_file), index=False)
        df = read_input_file(str(xlsx_file))
        assert len(df) == 1
        assert "Название" in df.columns

    def test_raises_file_not_found(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            read_input_file(str(tmp_path / "nonexistent.csv"))

    def test_raises_value_error_for_unsupported_format(self, tmp_path):
        txt_file = tmp_path / "test.txt"
        txt_file.write_text("test")
        with pytest.raises(ValueError, match="Непідтримуваний формат"):
            read_input_file(str(txt_file))

    def test_all_values_read_as_strings(self, tmp_path):
        csv_file = tmp_path / "test.csv"
        csv_file.write_text("Число\n42\n3.14\n", encoding="utf-8")
        df = read_input_file(str(csv_file))
        # pandas 3.x повертає StringDtype; перевіряємо, що значення — рядки
        assert pd.api.types.is_string_dtype(df["Число"])


# ===========================================================================
# Тести Модуля 8: validate_feed (інтеграційні)
# ===========================================================================


class TestValidateFeed:
    def make_csv(self, tmp_path: Path, rows: list[dict]) -> str:
        """Допоміжний метод: створює тимчасовий CSV-файл."""
        df = pd.DataFrame(rows)
        csv_path = str(tmp_path / "input.csv")
        df.to_csv(csv_path, index=False, encoding="utf-8")
        return csv_path

    def test_creates_output_file(self, tmp_path):
        csv_path = self.make_csv(tmp_path, [
            {"Название": "Ноутбук 512 ГБ SSD", "Объём SSD;115411": "512 ГБ"},
        ])
        out = str(tmp_path / "result.xlsx")
        validate_feed(csv_path, out)
        assert Path(out).exists()

    def test_no_errors_for_matching_values(self, tmp_path):
        csv_path = self.make_csv(tmp_path, [
            {"Название": "Ноутбук 512 ГБ SSD", "Объём SSD;115411": "512 ГБ"},
        ])
        out = str(tmp_path / "result.xlsx")
        df = validate_feed(csv_path, out)
        assert "Помилки" in df.columns
        assert df["Помилки"].iloc[0] == ""

    def test_detects_ssd_conflict(self, tmp_path):
        csv_path = self.make_csv(tmp_path, [
            {
                "Название": "Ноутбук 128 ГБ SSD",
                "Краткое описание": "Ноутбук із 256 ГБ SSD",
                "Объём SSD;115411": "512 ГБ",
            }
        ])
        out = str(tmp_path / "result.xlsx")
        df = validate_feed(csv_path, out)
        assert df["Помилки"].iloc[0] != ""
        assert "Конфлікт SSD" in df["Помилки"].iloc[0]

    def test_detects_ram_conflict(self, tmp_path):
        csv_path = self.make_csv(tmp_path, [
            {
                "Название": "Ноутбук 4 GB RAM",
                "Объем установленной оперативной памяти;20863": "32 GB",
            }
        ])
        out = str(tmp_path / "result.xlsx")
        df = validate_feed(csv_path, out)
        assert "Конфлікт RAM" in df["Помилки"].iloc[0]

    def test_detects_conflict_in_html_description(self, tmp_path):
        csv_path = self.make_csv(tmp_path, [
            {
                "Название": "Ноутбук ASUS",
                "Описание;1": "<p>Ноутбук із <b>128 ГБ</b> SSD</p>",
                "Объём SSD;115411": "512 ГБ",
            }
        ])
        out = str(tmp_path / "result.xlsx")
        df = validate_feed(csv_path, out)
        assert "Конфлікт SSD" in df["Помилки"].iloc[0]

    def test_multiple_conflicts_joined_by_pipe(self, tmp_path):
        csv_path = self.make_csv(tmp_path, [
            {
                "Название": "Ноутбук 128 ГБ SSD 4 GB RAM",
                "Объём SSD;115411": "512 ГБ",
                "Объем установленной оперативной памяти;20863": "32 GB",
            }
        ])
        out = str(tmp_path / "result.xlsx")
        df = validate_feed(csv_path, out)
        errors = df["Помилки"].iloc[0]
        assert "Конфлікт SSD" in errors
        assert "Конфлікт RAM" in errors
        assert " | " in errors

    def test_no_conflict_when_text_has_no_memory(self, tmp_path):
        csv_path = self.make_csv(tmp_path, [
            {
                "Название": "Ноутбук ASUS VivoBook",
                "Краткое описание": "Потужний тонкий ноутбук",
                "Объём SSD;115411": "512 ГБ",
            }
        ])
        out = str(tmp_path / "result.xlsx")
        df = validate_feed(csv_path, out)
        assert df["Помилки"].iloc[0] == ""

    def test_multiple_rows_only_conflicting_marked(self, tmp_path):
        csv_path = self.make_csv(tmp_path, [
            # Рядок 0: без конфлікту
            {
                "Название": "Ноутбук 512 ГБ SSD",
                "Объём SSD;115411": "512 ГБ",
            },
            # Рядок 1: з конфліктом
            {
                "Название": "Ноутбук 128 ГБ SSD",
                "Объём SSD;115411": "512 ГБ",
            },
        ])
        out = str(tmp_path / "result.xlsx")
        df = validate_feed(csv_path, out)
        assert df["Помилки"].iloc[0] == ""
        assert df["Помилки"].iloc[1] != ""

    def test_output_xlsx_has_pomylky_column(self, tmp_path):
        csv_path = self.make_csv(tmp_path, [
            {"Название": "Ноутбук", "Объём SSD;115411": "512 ГБ"},
        ])
        out = str(tmp_path / "result.xlsx")
        validate_feed(csv_path, out)
        df_out = pd.read_excel(out)
        assert "Помилки" in df_out.columns

    def test_sample_data_csv(self):
        """Перевірка на реальному файлі зразкових даних."""
        sample = Path(__file__).parent.parent / "sample_data.csv"
        if not sample.exists():
            pytest.skip("sample_data.csv не знайдено")
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
            out = f.name
        try:
            df = validate_feed(str(sample), out)
            # SKU002: назва каже 128ГБ, опис — 256ГБ, характеристика — 128ГБ → конфлікт в описі
            # SKU003: назва і характеристика 1ТБ, опис — 512ГБ → конфлікт
            # SKU004: назва 4ГБ RAM, опис і характеристика RAM різні → конфлікт
            errors_count = (df["Помилки"] != "").sum()
            assert errors_count > 0, "Зразкові дані повинні містити конфлікти"
        finally:
            os.unlink(out)


# ===========================================================================
# Тести командного рядка
# ===========================================================================


class TestCommandLine:
    """Тести для main() та аргументів командного рядка."""

    def test_main_with_optional_output(self, tmp_path):
        """Перевірка, що main() працює з опціональним аргументом output."""
        import subprocess
        import sys

        csv_path = tmp_path / "test_input.csv"
        df = pd.DataFrame([
            {"Название": "Ноутбук 512 ГБ SSD", "Объём SSD;115411": "512 ГБ"},
        ])
        df.to_csv(csv_path, index=False, encoding="utf-8")

        # Запускаємо без вказання вихідного файлу
        result = subprocess.run(
            [sys.executable, "checkr.py", str(csv_path)],
            capture_output=True,
            text=True,
            cwd=Path(__file__).parent.parent,
        )

        assert result.returncode == 0, f"Помилка: {result.stderr}"

        # Перевіряємо, що створився файл з автоматичною назвою
        # Файл має бути створений у тій же директорії, що й вхідний файл
        expected_output = tmp_path / "test_input_result.xlsx"
        assert expected_output.exists(), f"Вихідний файл не створено: {expected_output}"

        try:
            # Перевіряємо вміст
            df_out = pd.read_excel(expected_output)
            assert "Помилки" in df_out.columns
        finally:
            # Видаляємо створений файл після тесту
            if expected_output.exists():
                expected_output.unlink()
