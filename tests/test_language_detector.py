"""test_language_detector.py — Тести для модуля language_detector.py"""

import pytest
from language_detector import detect_language, check_language_consistency


class TestDetectLanguage:
    """Тести функції detect_language."""

    def test_ukrainian_unique_chars(self):
        """Текст з унікальними українськими символами."""
        assert detect_language("Ноутбук із SSD") == "uk"  # "із" має "і"
        assert detect_language("Це є ноутбук") == "uk"    # "є"
        assert detect_language("Комплект поставкі") == "uk"  # "і" в кінці
        assert detect_language("Пам'ять 16 GB") == "uk"  # апостроф
        assert detect_language("Ґрунт") == "uk"  # "ґ"

    def test_russian_unique_chars(self):
        """Текст з унікальними російськими символами."""
        assert detect_language("Ноутбук с SSD") == "ru"  # "с"
        assert detect_language("Это ноутбук") == "ru"    # "э"
        assert detect_language("Память 16 GB") == "ru"
        assert detect_language("Объём SSD") == "ru"      # "ъ"
        assert detect_language("Мы рады") == "ru"        # "ы"

    def test_ukrainian_common_words(self):
        """Визначення за частими словами (без унікальних символів)."""
        assert detect_language("Ноутбук та комп'ютер") == "uk"  # "та" - укр. слово
        assert detect_language("Комплект поставки з диска") == "uk"  # "з" - укр.

    def test_russian_common_words(self):
        """Визначення за частими словами (без унікальних символів)."""
        assert detect_language("Тонкий ноутбук для работы") == "ru"  # "работы" - рос.
        assert detect_language("Комплект поставкі") == "uk"  # "і" - укр. символ

    def test_english_text(self):
        """Англійський текст не визначається."""
        assert detect_language("Laptop with SSD") == "unknown"
        assert detect_language("RAM 16GB") == "unknown"

    def test_mixed_latin_cyrillic_ukrainian(self):
        """Змішаний текст (латиниця + кирилиця) з українськими маркерами."""
        assert detect_language("ASUS VivoBook із SSD") == "uk"
        assert detect_language("Dell XPS з RAM 16GB") == "uk"

    def test_mixed_latin_cyrillic_russian(self):
        """Змішаний текст (латиниця + кирилиця) з російськими маркерами."""
        assert detect_language("ASUS VivoBook с SSD") == "ru"
        assert detect_language("Dell XPS память 16GB") == "ru"

    def test_empty_string(self):
        """Порожній рядок."""
        assert detect_language("") == "unknown"
        assert detect_language("   ") == "unknown"

    def test_none_input(self):
        """None як вхідні дані."""
        assert detect_language(None) == "unknown"

    def test_numbers_only(self):
        """Тільки цифри."""
        assert detect_language("123456") == "unknown"
        assert detect_language("512 ГБ") == "unknown"  # немає достатньо тексту для визначення


class TestCheckLanguageConsistency:
    """Тести функції check_language_consistency."""

    def test_correct_pair_russian_ukrainian(self):
        """Правильна пара: російська та українська версії."""
        is_valid, error = check_language_consistency(
            "Ноутбук с SSD 512 ГБ",
            "Ноутбук із SSD 512 ГБ"
        )
        assert is_valid is True
        assert error == ""

    def test_correct_pair_with_latin(self):
        """Правильна пара з латинськими словами."""
        is_valid, error = check_language_consistency(
            "ASUS VivoBook с памятью 16GB",
            "ASUS VivoBook з пам'яттю 16GB"
        )
        assert is_valid is True
        assert error == ""

    def test_russian_in_ukrainian_field(self):
        """Російська мова в українському полі."""
        is_valid, error = check_language_consistency(
            "Ноутбук с SSD",
            "Ноутбук с SSD"  # тут має бути українська!
        )
        assert is_valid is False
        assert "Українська версія" in error
        assert "російську" in error

    def test_ukrainian_in_russian_field(self):
        """Українська мова в російському полі."""
        is_valid, error = check_language_consistency(
            "Ноутбук із SSD",  # тут має бути російська!
            "Ноутбук із SSD"
        )
        assert is_valid is False
        assert "Російська версія" in error
        assert "українську" in error

    def test_both_wrong(self):
        """Обидва поля містять невідповідну мову."""
        is_valid, error = check_language_consistency(
            "Ноутбук із SSD",  # має бути російська
            "Ноутбук с SSD"   # має бути українська
        )
        assert is_valid is False
        assert "Російська версія" in error
        assert "Українська версія" in error

    def test_empty_fields_allowed(self):
        """Порожні поля дозволені за замовчуванням."""
        is_valid, error = check_language_consistency("", "", allow_empty=True)
        assert is_valid is True
        assert error == ""

    def test_empty_fields_not_allowed(self):
        """Порожні поля заборонені."""
        is_valid, error = check_language_consistency("", "", allow_empty=False)
        assert is_valid is False
        assert "порожнє поле" in error

    def test_one_field_empty_allowed(self):
        """Одне поле порожнє (дозволено)."""
        is_valid, error = check_language_consistency(
            "Ноутбук с SSD",
            "",
            allow_empty=True
        )
        assert is_valid is True
        assert error == ""

    def test_latin_only_text(self):
        """Тільки латинський текст (мову не визначено, але це не помилка)."""
        is_valid, error = check_language_consistency(
            "Laptop SSD",
            "Notebook RAM"
        )
        # Мова не визначена, але це не вважається помилкою
        assert is_valid is True
        assert error == ""
