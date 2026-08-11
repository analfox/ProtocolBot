"""
config_loader.py - хранение настроек программы в файле config.json.

Зачем нужен: чтобы программа работала на ЛЮБОМ компьютере без правки кода.
Все пути (к Tesseract и Poppler), язык распознавания и качество скана
хранятся в config.json рядом с программой. На новом компьютере достаточно
поправить только этот файл (или указать пути через окно программы).

Как это работает:
 - load_config() читает config.json;
 - если файла ещё нет - он создаётся автоматически с настройками по умолчанию;
 - если файл есть, но в нём не хватает каких-то полей - недостающие
   подставляются из значений по умолчанию (ничего не ломается при обновлении).

ВАЖНО: сам по себе этот файл НИЧЕГО не меняет в работе программы,
пока его никто не импортирует. Поэтому его безопасно добавить отдельным шагом.
"""

import os
import json


# Имя файла настроек (лежит рядом с программой).
CONFIG_FILENAME = "config.json"


# ---------------------------------------------------------------------
# Настройки по умолчанию.
# Они попадут в config.json при самом первом запуске.
# Пути ниже - это "заводские" значения; на другом компьютере их нужно
# будет заменить на свои (вручную в config.json или через окно программы).
# Слэши намеренно прямые (/) - так файл легче читать и править руками,
# а Windows такие пути прекрасно понимает.
# ---------------------------------------------------------------------
DEFAULT_CONFIG = {
    # Путь к исполняемому файлу Tesseract OCR.
    "tesseract_path": "C:/Program Files/Tesseract-OCR/tesseract.exe",

    # Путь к папке bin утилит Poppler (там лежат pdftoppm и т.п.).
    "poppler_path": (
        "C:/Users/xofla/AppData/Local/Microsoft/WinGet/Packages/"
        "oschwartz10612.Poppler_Microsoft.Winget.Source_8wekyb3d8bbwe/"
        "poppler-25.07.0/Library/bin"
    ),

    # Язык распознавания. "rus" - проверенный рабочий режим.
    # Поставь "rus+eng", только если в заявках много латиницы/марок.
    "ocr_lang": "rus",

    # Качество скана для OCR (dpi). Больше = чётче, но медленнее.
    "ocr_dpi": 300,
}


def _base_dir():
    """Папка, в которой лежит этот файл (туда же кладём config.json)."""
    return os.path.dirname(os.path.abspath(__file__))


def _config_path():
    return os.path.join(_base_dir(), CONFIG_FILENAME)


def save_config(config):
    """Сохраняет словарь настроек в config.json."""
    with open(_config_path(), "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)


def load_config():
    """
    Читает config.json и возвращает словарь настроек.
    - Если файла нет - создаёт его из DEFAULT_CONFIG.
    - Если файл есть, но неполный - дополняет недостающие поля дефолтами
      и перезаписывает файл, чтобы в нём всегда был полный набор полей.
    - Если файл повреждён (не читается как JSON) - молча возвращает дефолты,
      чужой файл при этом НЕ перезаписывает, чтобы не затереть данные.
    """
    config = dict(DEFAULT_CONFIG)  # начинаем с копии дефолтов
    path = _config_path()

    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                user = json.load(f)
            if isinstance(user, dict):
                config.update(user)  # значения пользователя поверх дефолтов
            # Перезаписываем файл, чтобы в нём оказался полный набор полей.
            save_config(config)
        except Exception:
            # Битый JSON - не падаем и не трогаем файл, работаем с дефолтами.
            pass
    else:
        # Файла ещё не было - создаём.
        save_config(config)

    return config


def get_config():
    """То же, что load_config() - короткое имя для удобства."""
    return load_config()


# ---------------------------------------------------------------------
# Проверка. Запускается ТОЛЬКО если открыть этот файл напрямую
# (python config_loader.py). При импорте из других файлов этот блок
# НЕ выполняется, поэтому побочных эффектов нет.
# ---------------------------------------------------------------------
if __name__ == "__main__":
    cfg = load_config()

    print("Файл настроек:", _config_path())
    print()
    print(json.dumps(cfg, ensure_ascii=False, indent=4))
    print()

    t_path = (cfg.get("tesseract_path") or "").strip()
    p_path = (cfg.get("poppler_path") or "").strip()

    t_ok = bool(t_path) and os.path.isfile(t_path)
    p_ok = bool(p_path) and os.path.isdir(p_path)

    print("Tesseract найден по указанному пути:", t_ok)
    print("Poppler   найден по указанному пути:", p_ok)
    print()
    print("Если где-то False - на ЭТОМ компьютере путь нужно поправить")
    print("в config.json (или позже - через окно программы).")