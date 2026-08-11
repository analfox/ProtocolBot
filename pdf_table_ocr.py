"""
pdf_table_ocr.py - Локальное (офлайн) извлечение таблиц участников
из сканированных PDF-заявок. Не использует никаких облачных API -
всё работает на компьютере пользователя через Tesseract OCR.

Установка (один раз):
    pip install pdf2image pytesseract opencv-python-headless numpy
Windows:
  - Poppler: https://github.com/oschwartz10612/poppler-windows/releases
    распаковать, путь к папке \\bin добавить в PATH (или в config.json)
  - Tesseract: https://github.com/UB-Mannheim/tesseract/wiki
    при установке ОБЯЗАТЕЛЬНО отметить галочку "Russian" в списке языков
Linux:
    sudo apt install poppler-utils tesseract-ocr tesseract-ocr-rus

НАСТРОЙКИ (пути, язык, dpi) берутся из config.json через config_loader.py.
Поддерживает выбор страниц (pages="1", "1-3", "" = все), два прохода поиска
сетки (сканы / цифровые PDF), вырезание клеток с отступом от линий сетки
и запасной поиск колонки ФИО по содержимому.
"""

import os
import re
import shutil
import tempfile

import cv2
import numpy as np
import pytesseract
from pdf2image import convert_from_path


# =====================================================================
# ЧТЕНИЕ НАСТРОЕК ИЗ config.json (НЕ ИЗМЕНЕНО).
# =====================================================================
_FALLBACK_TESSERACT = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
_FALLBACK_POPPLER = ""

try:
    from config_loader import get_config
    _CFG = get_config()
except Exception:
    _CFG = {}

DEFAULT_TESSERACT_PATH = (_CFG.get("tesseract_path") or "").strip() or _FALLBACK_TESSERACT
DEFAULT_POPPLER_PATH = (_CFG.get("poppler_path") or "").strip() or _FALLBACK_POPPLER
DEFAULT_OCR_LANG = (_CFG.get("ocr_lang") or "").strip() or "rus"

try:
    DEFAULT_OCR_DPI = int(_CFG.get("ocr_dpi") or 300)
except (TypeError, ValueError):
    DEFAULT_OCR_DPI = 300

if os.name == "nt" and os.path.exists(DEFAULT_TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = DEFAULT_TESSERACT_PATH

# Включается из __main__ для печати служебной информации.
_DEBUG = False

# Отступ внутрь клетки при вырезании: убирает линии сетки из кадра,
# чтобы они не сбивали распознавание (и не давали мусор вида "|").
CELL_INSET = 8


# =====================================================================
# Нормализация текста ПОСЛЕ распознавания (НЕ ИЗМЕНЕНА).
# =====================================================================
_LAT_TO_CYR = {
    'A': 'А', 'B': 'В', 'C': 'С', 'E': 'Е', 'H': 'Н', 'K': 'К',
    'M': 'М', 'O': 'О', 'P': 'Р', 'T': 'Т', 'X': 'Х', 'Y': 'У',
    'a': 'а', 'c': 'с', 'e': 'е', 'o': 'о', 'p': 'р', 'x': 'х', 'y': 'у',
}
_CYR_TO_LAT = {v: k for k, v in _LAT_TO_CYR.items()}


def _is_cyr_only(ch):
    return ('А' <= ch <= 'я' or ch in 'Ёё') and ch not in _CYR_TO_LAT


def _is_lat_only(ch):
    return ('A' <= ch <= 'Z' or 'a' <= ch <= 'z') and ch not in _LAT_TO_CYR


def normalize_ocr_text(text):
    if not text:
        return text

    out = []
    for word in re.split(r'(\s+)', text):
        if not word or word.isspace():
            out.append(word)
            continue

        cyr_only = sum(1 for ch in word if _is_cyr_only(ch))
        lat_only = sum(1 for ch in word if _is_lat_only(ch))

        if cyr_only > lat_only:
            word = ''.join(_LAT_TO_CYR.get(ch, ch) for ch in word)
        elif lat_only > cyr_only:
            word = ''.join(_CYR_TO_LAT.get(ch, ch) for ch in word)

        out.append(word)

    return ''.join(out)


# =====================================================================
# Поиск сетки таблиц (два прохода: adaptive для сканов, otsu для цифровых).
# =====================================================================
def _get_line_masks(gray, mode="adaptive"):
    inv = ~gray
    if mode == "otsu":
        _, thresh = cv2.threshold(inv, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    else:
        thresh = cv2.adaptiveThreshold(
            inv, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 15, -2
        )

    h, w = gray.shape
    horiz_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(w // 30, 10), 1))
    horiz = cv2.dilate(cv2.erode(thresh, horiz_kernel), horiz_kernel)
    vert_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(h // 30, 10)))
    vert = cv2.dilate(cv2.erode(thresh, vert_kernel), vert_kernel)
    return horiz, vert


def _find_table_regions(horiz, vert):
    grid = cv2.add(horiz, vert)
    bridge_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (25, 25))
    bridged = cv2.dilate(grid, bridge_kernel)
    contours, _ = cv2.findContours(bridged, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    boxes = [cv2.boundingRect(c) for c in contours]
    boxes = [b for b in boxes if b[2] > 150 and b[3] > 50]
    boxes.sort(key=lambda b: b[1])
    return boxes


def _line_positions(mask_1d_sum, min_len, gap=5):
    on = mask_1d_sum > min_len
    bands = []
    start = None
    for i, v in enumerate(on):
        if v and start is None:
            start = i
        elif not v and start is not None:
            bands.append((start, i))
            start = None
    if start is not None:
        bands.append((start, len(on)))
    merged = []
    for b in bands:
        if merged and b[0] - merged[-1][1] <= gap:
            merged[-1] = (merged[-1][0], b[1])
        else:
            merged.append(list(b))
    return [(s + e) // 2 for s, e in merged]


# =====================================================================
# Распознавание одной ячейки: белая рамка + повторная попытка (psm 7).
# =====================================================================
def _ocr_cell(img, lang=None):
    if lang is None:
        lang = DEFAULT_OCR_LANG

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY) if img.ndim == 3 else img

    # Белая кайма: текст не прилипает к краям кадра.
    gray = cv2.copyMakeBorder(gray, 8, 8, 8, 8, cv2.BORDER_CONSTANT, value=255)

    if gray.shape[0] < 150:
        gray = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)

    text = pytesseract.image_to_string(gray, lang=lang, config="--psm 6")
    text = re.sub(r"\s+", " ", text).strip()

    if not text:
        text = pytesseract.image_to_string(gray, lang=lang, config="--psm 7")
        text = re.sub(r"\s+", " ", text).strip()

    text = normalize_ocr_text(text)

    return text


def _classify_columns(header_texts):
    name_idx = pos_idx = org_idx = group_idx = None
    for i, h in enumerate(header_texts):
        hl = h.lower()
        if any(k in hl for k in ("фио", "фамилия")):
            name_idx = i
        elif any(k in hl for k in ("должность", "профессия")):
            pos_idx = i
        elif any(k in hl for k in ("подразделение", "организация", "наименование")):
            org_idx = i
        elif "групп" in hl:
            group_idx = i
    return name_idx, pos_idx, org_idx, group_idx


def _looks_like_name(text):
    """Похоже ли содержимое ячейки на ФИО (2-4 слова, из них >=2 с ЗАГЛАВНОЙ)."""
    t = (text or "").strip()
    if not t or len(t) < 5:
        return False
    words = t.split()
    if not (2 <= len(words) <= 4):
        return False
    # ВАЖНО: именно заглавные кириллические (А-Я, Ё), а не любые кириллические.
    cap = sum(
        1 for w in words
        if w and ('А' <= w[0] <= 'Я' or w[0] == 'Ё')
    )
    return cap >= 2


def _guess_name_column(grid_text, header, pos_idx, org_idx):
    """
    Запасной поиск колонки ФИО по содержимому строк данных.
    Включается, только если заголовок не подсказал ФИО, но таблица
    похожа на таблицу участников (есть Должность/СНИЛС/Группа).
    Колонки, уже опознанные как Должность/Организация, не рассматриваются.
    """
    header_low = [str(h).lower() for h in header]
    table_looks_like_participants = (
        pos_idx is not None
        or any("снилс" in h for h in header_low)
        or any("групп" in h for h in header_low)
    )
    if not table_looks_like_participants:
        return None

    if len(grid_text) < 2:
        return None

    n_cols = len(grid_text[0])
    best_ci, best_score = None, 0
    for ci in range(n_cols):
        if ci == pos_idx or ci == org_idx:
            continue
        score = sum(
            1 for row in grid_text[1:]
            if _looks_like_name(row[ci] if ci < len(row) else "")
        )
        if score > best_score:
            best_ci, best_score = ci, score

    need = 2 if len(grid_text) - 1 >= 3 else len(grid_text) - 1
    if best_ci is not None and best_score >= need:
        return best_ci
    return None


def guess_group(heading: str) -> str:
    h = heading.lower()
    h_ns = re.sub(r"\s+", "", h)

    if "3групп" in h_ns:
        return "3"
    if ("1и2" in h_ns) or ("1" in h_ns and "2" in h_ns and "групп" in h_ns):
        return "1и2"

    m = re.search(r"(\d)групп", h_ns)
    if m:
        return m.group(1)

    m = re.search(r"(\d)\s*групп", h)
    return m.group(1) if m else ""
    
def _clean_group_cell(text):
    """Приводит распознанную группу к цифрам: 'й'->2, 'з'->3 и т.п.
    Тессеракт на 600 dpi часто путает цифры с буквами."""
    if not text:
        return ""
    t = str(text).strip().lower()
    # Карта частых ошибок OCR для одиночных цифр группы.
    fix = {
        "й": "2", "и": "2", "z": "2",
        "з": "3", "s": "3",
        "1": "1", "2": "2", "3": "3",
    }
    out = []
    for ch in t:
        if ch in fix:
            out.append(fix[ch])
        elif ch in ",;.:/- ":
            out.append(",")  # разделители групп -> запятая
    # Убираем дубли и лишние запятые: "2,2" -> "2", ",2," -> "2"
    parts = [p for p in "".join(out).split(",") if p in ("1", "2", "3")]
    seen = []
    for p in parts:
        if p not in seen:
            seen.append(p)
    return ",".join(seen)

def _parse_pages(pages):
    """
    Разбирает строку со страницами: "1", "1-3", "2-", "-3", "" (все).
    Возвращает (first_page, last_page) для convert_from_path.
    """
    if not pages or not str(pages).strip():
        return None, None

    s = str(pages).strip()
    try:
        if "-" in s:
            a, _, b = s.partition("-")
            first = int(a) if a.strip().isdigit() else None
            last = int(b) if b.strip().isdigit() else None
            return first, last
        if s.isdigit():
            n = int(s)
            return n, n
    except ValueError:
        pass
    return None, None


def extract_tables_from_pdf(
    file_path: str,
    dpi: int = None,
    lang: str = None,
    tesseract_path: str = None,
    poppler_path: str = None,
    pages: str = None,
) -> list:
    """
    Извлекает все таблицы участников из скан-PDF заявки (офлайн, через OCR).

    dpi  : качество скана. Если None - из конфига (ocr_dpi).
    lang : язык распознавания. Если None - из конфига (ocr_lang).
    pages: какие страницы обрабатывать: "1", "1-3", "" или None = все.
    tesseract_path / poppler_path: переопределяют пути из конфига.

    Возвращает список словарей:
        {"heading": "<текст над таблицей>", "group_guess": "3"/"1и2"/"",
         "participants": [{"name":..., "position":..., "organization":...}, ...]}
    """
    if dpi is None:
        dpi = DEFAULT_OCR_DPI
    if lang is None:
        lang = DEFAULT_OCR_LANG

    if tesseract_path and os.path.exists(tesseract_path):
        pytesseract.pytesseract.tesseract_cmd = tesseract_path

    poppler_path = poppler_path or (
        DEFAULT_POPPLER_PATH if os.path.exists(DEFAULT_POPPLER_PATH) else None
    )

    first_page, last_page = _parse_pages(pages)

    # Poppler на Windows нестабильно открывает пути с кириллицей/пробелами,
    # а также файлы, открытые в другой программе. Поэтому копируем PDF во
    # временный файл с простым именем и распознаём копию; оригинал не трогаем.
    src_path = os.path.abspath(file_path)
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".pdf")
    os.close(tmp_fd)

    try:
        shutil.copyfile(src_path, tmp_path)
        images = convert_from_path(
            tmp_path,
            dpi=dpi,
            poppler_path=poppler_path,
            first_page=first_page,
            last_page=last_page,
        )
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass

    tables = []
    start_no = first_page or 1

    for page_no, page_img in enumerate(images, start=start_no):
        img = cv2.cvtColor(np.array(page_img), cv2.COLOR_RGB2BGR)
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        # Проход 1: адаптивный порог (сканы - проверенный режим).
        horiz, vert = _get_line_masks(gray, "adaptive")
        regions = _find_table_regions(horiz, vert)
        pass_used = "adaptive"

        # Проход 2: только если первый ничего не нашёл (цифровые PDF).
        if not regions:
            horiz, vert = _get_line_masks(gray, "otsu")
            regions = _find_table_regions(horiz, vert)
            pass_used = "otsu"

        if _DEBUG:
            print(f"--- страница {page_no}: {img.shape[1]}x{img.shape[0]}, "
                  f"регионов сетки: {len(regions)} (проход: {pass_used})")
            for (x, y, w, h) in regions:
                print(f"    регион: x={x} y={y} w={w} h={h}")

        page_prev_bottom = 0
        for (x, y, w, h) in regions:
            heading = ""
            if y > page_prev_bottom + 15:
                strip = img[page_prev_bottom:y, 0:img.shape[1]]
                heading = _ocr_cell(strip, lang=lang)

            sub_h = horiz[y:y + h, x:x + w]
            sub_v = vert[y:y + h, x:x + w]
            row_sum = (sub_h > 0).sum(axis=1)
            col_sum = (sub_v > 0).sum(axis=0)
            rows = _line_positions(row_sum, min_len=w * 0.2)
            cols = _line_positions(col_sum, min_len=h * 0.2)
            page_prev_bottom = y + h

            if _DEBUG:
                print(f"    таблица {w}x{h}: строк={len(rows)}, колонок={len(cols)}")

            if len(rows) < 2 or len(cols) < 2:
                continue

            # Вырезаем клетки С ОТСТУПОМ от линий сетки, чтобы линии
            # не попадали в кадр и не портили распознавание.
            grid_text = []
            for ri in range(len(rows) - 1):
                y0 = y + rows[ri] + CELL_INSET
                y1 = y + rows[ri + 1] - CELL_INSET
                if y1 <= y0:
                    y0, y1 = y + rows[ri], y + rows[ri + 1]

                row_cells = []
                for ci in range(len(cols) - 1):
                    x0 = x + cols[ci] + CELL_INSET
                    x1 = x + cols[ci + 1] - CELL_INSET
                    if x1 <= x0:
                        x0, x1 = x + cols[ci], x + cols[ci + 1]

                    cell_img = img[y0:y1, x0:x1]
                    row_cells.append(_ocr_cell(cell_img, lang=lang))
                grid_text.append(row_cells)

            if _DEBUG:
                for r_i, r in enumerate(grid_text):
                    print(f"    строка {r_i}: {r}")

            header = grid_text[0]
            name_idx, pos_idx, org_idx, group_idx = _classify_columns(header)

            # Запасной поиск колонки ФИО по содержимому строк.
            fallback_used = False
            if name_idx is None:
                guessed = _guess_name_column(grid_text, header, pos_idx, org_idx)
                if guessed is not None:
                    name_idx = guessed
                    fallback_used = True

            if _DEBUG:
                print(f"    name_idx={name_idx} (запасной поиск: {fallback_used}), "
                      f"pos_idx={pos_idx}, org_idx={org_idx}")

            if name_idx is None:
                continue

            participants = []
            for row in grid_text[1:]:
                if name_idx >= len(row) or not row[name_idx] or len(row[name_idx]) < 3:
                    continue
                participants.append({
                    "name": row[name_idx],
                    "position": row[pos_idx] if pos_idx is not None and pos_idx < len(row) else "",
                    "organization": row[org_idx] if org_idx is not None and org_idx < len(row) else "",
                    "group": _clean_group_cell(row[group_idx]) if group_idx is not None and group_idx < len(row) else "",
                })

            tables.append({
                "heading": heading,
                "group_guess": guess_group(heading),
                "participants": participants,
            })

    return tables


if __name__ == "__main__":
    import sys
    import json

    _DEBUG = True

    if len(sys.argv) < 2:
        print("Использование: python pdf_table_ocr.py <файл.pdf> [страницы]")
        print('Примеры: python pdf_table_ocr.py заявка.pdf 1')
        sys.exit(1)

    pdf_path = sys.argv[1]
    pages_arg = sys.argv[2] if len(sys.argv) > 2 else None

    result = extract_tables_from_pdf(pdf_path, pages=pages_arg)
    print("\n=== РЕЗУЛЬТАТ ===")
    print(json.dumps(result, ensure_ascii=False, indent=2))