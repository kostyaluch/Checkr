"""Тести для нових функцій перевірки CPU, GPU та OS.

Запуск:
    pytest tests/test_cpu_gpu_os.py -v
"""

import pytest

from checkr import (
    VALIDATION_RULES,
    extract_cpu_model_matches,
    extract_gpu_model_matches,
    extract_os_matches,
    normalize_cpu_model,
    normalize_gpu_model,
    normalize_os,
)


# ===========================================================================
# Тести для CPU (процесор)
# ===========================================================================


class TestExtractCpuModelMatches:
    """Тести пошуку моделей процесорів."""

    def test_intel_core_i7(self):
        text = "Ноутбук з Intel Core i7-12700F процесором"
        matches = extract_cpu_model_matches(text)
        assert len(matches) == 1
        assert "Intel Core i7-12700F" in matches[0][1]

    def test_intel_core_i5(self):
        text = "Четырнадцатиядерный Intel Core i5-13600KF (3.5 - 5.1 ГГц)"
        matches = extract_cpu_model_matches(text)
        assert len(matches) == 1
        assert "Intel Core i5-13600KF" in matches[0][1]

    def test_intel_core_ultra(self):
        text = "Четырнадцатиядерный Intel Core Ultra 5 125H (1.2 - 4.5 ГГц)"
        matches = extract_cpu_model_matches(text)
        assert len(matches) == 1
        assert "Intel Core Ultra 5 125H" in matches[0][1]

    def test_amd_ryzen_5(self):
        text = "Четырехъядерный AMD Ryzen 5 7520U (2.8 - 4.3 ГГц)"
        matches = extract_cpu_model_matches(text)
        assert len(matches) == 1
        assert "AMD Ryzen 5 7520U" in matches[0][1]

    def test_amd_ryzen_7(self):
        text = "Восьмиядерный AMD Ryzen 7 9700X (3.8 - 5.5 ГГц)"
        matches = extract_cpu_model_matches(text)
        assert len(matches) == 1
        assert "AMD Ryzen 7 9700X" in matches[0][1]

    def test_amd_ryzen_9(self):
        text = "Двенадцатиядерный AMD Ryzen 9 9900X (4.4 - 5.6 ГГц)"
        matches = extract_cpu_model_matches(text)
        assert len(matches) == 1
        assert "AMD Ryzen 9 9900X" in matches[0][1]

    def test_intel_pentium(self):
        text = "Двухъядерный Intel Pentium Gold G7400 (3.7 ГГц)"
        matches = extract_cpu_model_matches(text)
        assert len(matches) == 1
        assert "Intel Pentium Gold G7400" in matches[0][1]

    def test_intel_core_processor(self):
        text = "Десятиядерный Intel Core 5 processor 120U (3.8 - 5 ГГц)"
        matches = extract_cpu_model_matches(text)
        assert len(matches) == 1
        assert "Intel Core 5 processor 120U" in matches[0][1]

    def test_empty_string(self):
        matches = extract_cpu_model_matches("")
        assert matches == []

    def test_no_cpu_in_text(self):
        matches = extract_cpu_model_matches("Просто текст без процесора")
        assert matches == []

    def test_short_description(self):
        text = "Intel Core i5-13600KF (3.5 - 5.1 ГГц) / RAM 64 ГБ / SSD 1 ТБ"
        matches = extract_cpu_model_matches(text)
        assert len(matches) == 1
        assert "Intel Core i5-13600KF" in matches[0][1]


class TestNormalizeCpuModel:
    """Тести нормалізації моделей процесорів."""

    def test_normalize_intel_lowercase(self):
        result = normalize_cpu_model("intel core i7-12700f")
        assert result.startswith("Intel")
        assert "Core" in result

    def test_normalize_amd_lowercase(self):
        result = normalize_cpu_model("amd ryzen 5 7520u")
        assert result.startswith("AMD")
        assert "Ryzen" in result

    def test_preserve_model_suffix(self):
        result = normalize_cpu_model("Intel Core i7-12700F")
        assert "12700F" in result


# ===========================================================================
# Тести для GPU (відеокарта)
# ===========================================================================


class TestExtractGpuModelMatches:
    """Тести пошуку моделей відеокарт."""

    def test_geforce_rtx_5060_ti(self):
        text = "nVidia GeForce RTX 5060 Ti, 16 ГБ"
        matches = extract_gpu_model_matches(text)
        assert len(matches) == 1
        assert "GeForce RTX 5060 Ti" in matches[0][1]

    def test_geforce_rtx_5070(self):
        text = "nVidia GeForce RTX 5070, 12 ГБ"
        matches = extract_gpu_model_matches(text)
        assert len(matches) == 1
        assert "GeForce RTX 5070" in matches[0][1]

    def test_geforce_rtx_5080(self):
        text = "GeForce RTX 5080"
        matches = extract_gpu_model_matches(text)
        assert len(matches) == 1
        assert "GeForce RTX 5080" in matches[0][1]

    def test_radeon_rx_9070_xt(self):
        text = "AMD Radeon RX 9070 XT"
        matches = extract_gpu_model_matches(text)
        assert len(matches) == 1
        assert "Radeon RX 9070 XT" in matches[0][1]

    def test_radeon_rx_9060(self):
        text = "Radeon RX 9060 XT"
        matches = extract_gpu_model_matches(text)
        assert len(matches) == 1
        assert "Radeon RX 9060 XT" in matches[0][1]

    def test_intel_arc_graphics(self):
        text = "Intel Arc Graphics"
        matches = extract_gpu_model_matches(text)
        assert len(matches) == 1
        assert "Intel Arc Graphics" in matches[0][1]

    def test_intel_hd_graphics(self):
        text = "Intel HD Graphics"
        matches = extract_gpu_model_matches(text)
        assert len(matches) == 1
        assert "Intel HD Graphics" in matches[0][1]

    def test_empty_string(self):
        matches = extract_gpu_model_matches("")
        assert matches == []

    def test_no_gpu_in_text(self):
        matches = extract_gpu_model_matches("Просто текст без відеокарти")
        assert matches == []

    def test_short_description(self):
        text = "SSD 1 ТБ / nVidia GeForce RTX 5070 Ti, 16 ГБ / без ОД"
        matches = extract_gpu_model_matches(text)
        assert len(matches) == 1
        assert "GeForce RTX 5070 Ti" in matches[0][1]


class TestNormalizeGpuModel:
    """Тести нормалізації моделей відеокарт."""

    def test_normalize_geforce_lowercase(self):
        result = normalize_gpu_model("geforce rtx 5060 ti")
        assert "GeForce" in result
        assert "RTX" in result
        assert "Ti" in result

    def test_normalize_radeon_lowercase(self):
        result = normalize_gpu_model("radeon rx 9070 xt")
        assert "Radeon" in result
        assert "RX" in result
        assert "XT" in result

    def test_preserve_model_number(self):
        result = normalize_gpu_model("GeForce RTX 5060")
        assert "5060" in result


# ===========================================================================
# Тести для OS (операційна система)
# ===========================================================================


class TestExtractOsMatches:
    """Тести пошуку операційних систем."""

    def test_windows_11_home(self):
        text = "Windows 11 Home"
        matches = extract_os_matches(text)
        assert len(matches) == 1
        assert "Windows 11 Home" in matches[0][1]

    def test_windows_11_pro(self):
        text = "Windows 11 Pro"
        matches = extract_os_matches(text)
        assert len(matches) == 1
        assert "Windows 11 Pro" in matches[0][1]

    def test_windows_10(self):
        text = "Windows 10"
        matches = extract_os_matches(text)
        assert len(matches) == 1
        assert "Windows 10" in matches[0][1]

    def test_macos(self):
        text = "macOS Ventura"
        matches = extract_os_matches(text)
        assert len(matches) == 1
        assert "macOS" in matches[0][1]

    def test_linux(self):
        text = "Ubuntu 22.04"
        matches = extract_os_matches(text)
        assert len(matches) == 1
        assert "Ubuntu" in matches[0][1]

    def test_empty_string(self):
        matches = extract_os_matches("")
        assert matches == []

    def test_no_os_in_text(self):
        matches = extract_os_matches("Просто текст без процесора")
        assert matches == []

    def test_short_description(self):
        text = "SSD 1 ТБ / nVidia GeForce RTX 5070, 12 ГБ / LAN / Windows 11 Home"
        matches = extract_os_matches(text)
        assert len(matches) == 1
        assert "Windows 11 Home" in matches[0][1]

    def test_no_os_status(self):
        # "без ОС" означає "без операційної системи" - це валідний статус
        text = "Ноутбук без ОС"
        matches = extract_os_matches(text)
        assert len(matches) == 1
        assert "без" in matches[0][1].lower()


class TestNormalizeOs:
    """Тести нормалізації назв операційних систем."""

    def test_normalize_windows_lowercase(self):
        result = normalize_os("windows 11 home")
        assert "Windows" in result
        assert "Home" in result

    def test_normalize_macos_lowercase(self):
        result = normalize_os("macos")
        assert result == "macOS"

    def test_preserve_version(self):
        result = normalize_os("Windows 11")
        assert "11" in result


# ===========================================================================
# Тести правил валідації
# ===========================================================================


class TestValidationRulesCpu:
    """Тести правил валідації: модель процесора."""

    def test_has_cpu_rule(self):
        labels = [r["label"] for r in VALIDATION_RULES]
        assert "Модель процесора" in labels

    def test_cpu_rule_has_correct_checker_type(self):
        cpu_rule = next(
            r for r in VALIDATION_RULES if r["label"] == "Модель процесора"
        )
        assert cpu_rule["checker_type"] == "cpu_model"


class TestValidationRulesGpu:
    """Тести правил валідації: модель відеокарти."""

    def test_has_gpu_rule(self):
        labels = [r["label"] for r in VALIDATION_RULES]
        assert "Модель відеокарти" in labels

    def test_gpu_rule_has_correct_checker_type(self):
        gpu_rule = next(
            r for r in VALIDATION_RULES if r["label"] == "Модель відеокарти"
        )
        assert gpu_rule["checker_type"] == "gpu_model"


class TestValidationRulesOs:
    """Тести правил валідації: операційна система."""

    def test_has_os_rule(self):
        labels = [r["label"] for r in VALIDATION_RULES]
        assert "Операційна система" in labels

    def test_os_rule_has_correct_checker_type(self):
        os_rule = next(
            r for r in VALIDATION_RULES if r["label"] == "Операційна система"
        )
        assert os_rule["checker_type"] == "os"
