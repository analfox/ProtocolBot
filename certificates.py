"""
certificates.py - удостоверения ПО ШАБЛОНУ Word с плейсхолдерами.

Вёрстка живёт в файле certificate_template.docx рядом с программой и
НЕ ТРОГАЕТСЯ программой. Размер удостоверения в шаблоне жёсткий:
  - строки таблицы точной высоты; справа подпись руководителя - отдельная
    нижняя строка, поэтому всегда внизу;
  - у строк Фамилия/Имя/Отчество точный межстрочный интервал (пустое
    отчество или уменьшенная фамилия не поднимают низ);
  - должность и организация - во вложенной таблице без рамок точной высоты
    справа от места под фото, текст прижат к низу (длинная должность растёт
    вверх и ничего ниже не сдвигает).
Программа только:
  1) подставляет данные вместо {плейсхолдеров};
  2) фиксирует ширины колонок;
  3) уменьшает шрифт, если текст не влезает в свою ячейку (должность,
     организация, допуски, длинные ФИО) - иначе Word обрезал бы его.

Плейсхолдеры: {number} {surname} {name} {patronymic} {position} {org}
{date_issue} {date_valid} {group} {permit_text} {protocol_number}
{protocol_date} {hours}

2 удостоверения на страницу. Группы 1 и 2 - на 3 года, группа 3 - на 5 лет.
Часы практического обучения - всегда 5.
"""
import copy
import os
import re
from datetime import datetime

from docx import Document
from docx.shared import Pt
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

TEMPLATE_NAME = "certificate_template.docx"
PRACTICE_HOURS = "5"

# Геометрия таблицы (как в эталоне) - фиксируем, чтобы длинный
# текст не расползался по ширине.
LEFT_WIDTH_CM = 8.93
RIGHT_WIDTH_CM = 8.75

# Место под текст в шаблоне (pt, замерено в Word). По нему подбирается шрифт,
# чтобы текст не вылез за ячейку точной высоты. Поменяли вёрстку шаблона -
# поправьте и эти числа.
BOX_WIDTH_PT = 143.85      # рамка «должность + организация» справа от фото
BOX_TEXT_PT = 33.75        # её высота минус «(профессия должность)», «(организация)» и отступ
PERMIT_WIDTH_PT = 237.25   # ширина текста правой ячейки
PERMIT_HEIGHT_PT = 75.75   # место под абзац «Может быть допущен...» до подписи
FIO_WIDTH_PT = 99.0        # от начала фамилии/имени/отчества до правого края
LINE_FACTOR = 1.2          # высота строки Times New Roman в Word ~1.2 кегля (с запасом)
MIN_PT = 4.5

_MONTHS = [
    "января", "февраля", "марта", "апреля", "мая", "июня",
    "июля", "августа", "сентября", "октября", "ноября", "декабря",
]

# Схемный порядок элементов - чтобы Word не считал файл битым.
_TBLPR_ORDER = [
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
    "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
    "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
    "tblCellMar", "tblLook", "caption", "tblDescription",
]
_TCPR_ORDER = [
    "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders",
    "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign",
    "hideMark",
]


def _tw(cm):
    return str(int(round(cm * 567)))


def _name(tag):
    return tag.split("}")[-1]


def _date_parts(date_str):
    m = re.match(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})$", str(date_str or "").strip())
    if not m:
        return None
    day, month, year = m.groups()
    if not 1 <= int(month) <= 12:
        return None
    return int(day), _MONTHS[int(month) - 1], int(year)


def next_number(current):
    m = re.match(r"^(\d+)(.*)$", str(current).strip(), re.S)
    if not m:
        return current
    return str(int(m.group(1)) + 1) + m.group(2)


def _upsert(parent, tag, order):
    el = parent.find(qn(tag))
    if el is not None:
        return el
    el = OxmlElement(tag)
    pos = order.index(_name(tag)) if _name(tag) in order else 999
    for i, child in enumerate(parent):
        cn = _name(child.tag)
        if cn in order and order.index(cn) > pos:
            parent.insert(i, el)
            return el
    parent.append(el)
    return el


def _strip_ids(element, seed=0):
    """Убирает дублирующиеся у копий paraId/textId и делает уникальными
    id фигур и docPr - иначе Word считает файл повреждённым."""
    n = 1000 + seed * 50
    for el in element.iter():
        for k in list(el.attrib):
            if k.split("}")[-1] in ("paraId", "textId"):
                del el.attrib[k]
        ln = _name(el.tag)
        if ln == "docPr":
            el.set("id", str(n))
            n += 1
        elif ln == "shape":
            el.set("id", f"_x0000_i{n}")
            n += 1


def _merge_split_placeholders(p):
    """Склеивает раны, в которых Word разбил плейсхолдер на куски."""
    runs = [r for r in p.findall(qn("w:r")) if r.find(qn("w:t")) is not None]
    i = 0
    while i < len(runs):
        t = runs[i].find(qn("w:t"))
        txt = t.text or ""
        j = i
        while txt.count("{") > txt.count("}") and j + 1 < len(runs):
            j += 1
            txt += (runs[j].find(qn("w:t")).text or "")
        if j > i:
            t.text = txt
            t.set(qn("xml:space"), "preserve")
            for k in range(i + 1, j + 1):
                p.remove(runs[k])
            runs = [r for r in p.findall(qn("w:r")) if r.find(qn("w:t")) is not None]
        i += 1


def _fill_element(element, mapping):
    """Заменяет {плейсхолдеры} во всех текстовых узлах таблицы."""
    for p in element.iter(qn("w:p")):
        _merge_split_placeholders(p)
    for t in element.iter(qn("w:t")):
        text = t.text
        if not text or "{" not in text:
            continue
        for key, value in mapping.items():
            text = text.replace("{" + key + "}", value)
        t.text = text


def _fix_geometry(tbl):
    """Фиксированная раскладка + точные ширины колонок (схемобезопасно)."""
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    lay = _upsert(tblPr, "w:tblLayout", _TBLPR_ORDER)
    lay.set(qn("w:type"), "fixed")
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        cols = grid.findall(qn("w:gridCol"))
        if len(cols) == 2:
            cols[0].set(qn("w:w"), _tw(LEFT_WIDTH_CM))
            cols[1].set(qn("w:w"), _tw(RIGHT_WIDTH_CM))
    # Только строки самого удостоверения, вложенную рамку должности не трогаем.
    for tr in tbl.findall(qn("w:tr")):
        for tc, w in zip(tr.findall(qn("w:tc")), (LEFT_WIDTH_CM, RIGHT_WIDTH_CM)):
            tcW = _upsert(tc.get_or_add_tcPr(), "w:tcW", _TCPR_ORDER)
            tcW.set(qn("w:type"), "dxa")
            tcW.set(qn("w:w"), _tw(w))


_PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd",
    "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
    "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
    "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
    "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
    "textDirection", "textAlignment", "textboxProperties", "outlineLvl",
    "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
]


def _suppress_auto_hyphens(p):
    pEl = p._p                      # внутренний XML-элемент абзаца
    pPr = pEl.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        pEl.insert(0, pPr)
    _upsert(pPr, "w:suppressAutoHyphens", _PPR_ORDER)


def suppress_hyphens_for(element, keys):
    """Запрет автопереноса в абзацах с этими плейсхолдерами: ФИО не рвём
    через дефис, длинное слово уходит на новую строку целиком."""
    for p in element.iter(qn("w:p")):
        text = "".join(t.text or "" for t in p.iter(qn("w:t")))
        if any("{" + k + "}" in text for k in keys):
            _suppress_auto_hyphens(Paragraph(p, None))

# --- Подбор шрифта: текст должен влезть в ячейку точной высоты ----------

_FONTS = {}


def _text_width(text, size, bold):
    """Ширина текста в pt по метрикам Times New Roman (шрифт шаблона)."""
    if bold not in _FONTS:
        try:
            from PIL import ImageFont
            _FONTS[bold] = ImageFont.truetype("timesbd.ttf" if bold else "times.ttf", 100)
        except Exception:
            _FONTS[bold] = None
    font = _FONTS[bold]
    if font is None:
        return len(text) * size * 0.6          # без шрифта - грубо и с запасом
    return font.getlength(text) * size / 100


def _line_count(text, width, size, bold):
    """Сколько строк займёт абзац при переносе по словам, как в Word."""
    space = _text_width(" ", size, bold)
    lines, cur = 1, 0.0
    for word in text.split():
        w = _text_width(word, size, bold)
        if cur and cur + space + w > width:
            lines += 1
            cur = 0.0
        cur += (space if cur else 0) + w
        while cur > width:                     # слово длиннее строки
            lines += 1
            cur -= width
    return lines


def _sizes_down(base):
    """base, base-0.5, ... до MIN_PT."""
    size = base
    while size > MIN_PT:
        yield size
        size -= 0.5
    yield MIN_PT


def _run_size(r, default):
    sz = r.find(qn("w:rPr") + "/" + qn("w:sz"))
    return int(sz.get(qn("w:val"))) / 2 if sz is not None else default


_RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
    "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz",
    "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign",
    "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath",
]


def _set_size(element, size):
    """Кегль для рана или знака абзаца (rPr)."""
    rPr = element if element.tag == qn("w:rPr") else element.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        element.insert(0, rPr)
    for tag in ("w:sz", "w:szCs"):
        _upsert(rPr, tag, _RPR_ORDER).set(qn("w:val"), str(int(round(size * 2))))


def _par_text(p):
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


class _Target:
    """Абзац с плейсхолдером: какие раны уменьшать и исходный кегль шаблона."""

    def __init__(self, p, key, whole_paragraph):
        self.p = p
        runs = p.findall(qn("w:r"))
        if not whole_paragraph:
            runs = [r for r in runs if "{" + key + "}" in _par_text(r)]
        self.runs = runs
        self.base = _run_size(runs[0], 7) if runs else 7
        self.bold = bool(runs) and runs[0].find(qn("w:rPr") + "/" + qn("w:b")) is not None

    def text(self):
        return "".join(_par_text(r) for r in self.runs)

    def resize(self, size):
        if size >= self.base:
            return
        for r in self.runs:
            _set_size(r, size)
        pPr = self.p.find(qn("w:pPr"))
        if pPr is not None and pPr.find(qn("w:rPr")) is not None:
            _set_size(pPr.find(qn("w:rPr")), size)   # знак абзаца тоже задаёт высоту строки


def _find_targets(tbl):
    """До подстановки: находит абзацы, размер которых может понадобиться уменьшить."""
    for p in tbl.iter(qn("w:p")):
        _merge_split_placeholders(p)
    found = {}
    for p in tbl.iter(qn("w:p")):
        text = _par_text(p)
        for key, whole in (("surname", False), ("name", False), ("patronymic", False),
                           ("position", True), ("org", True), ("permit_text", True)):
            if "{" + key + "}" in text and key not in found:
                found[key] = _Target(p, key, whole)
    return found


def _fit_text(found):
    """После подстановки: уменьшает шрифт, пока текст не влезет в свою ячейку."""
    for key in ("surname", "name", "patronymic"):
        t = found.get(key)
        if t:
            t.resize(next(s for s in _sizes_down(t.base)
                          if _text_width(t.text().strip(), s, t.bold) <= FIO_WIDTH_PT or s == MIN_PT))

    pos, org = found.get("position"), found.get("org")
    if pos and org:
        # Одно уменьшение на оба абзаца, организация не крупнее должности.
        for step in range(0, 20):
            ps = max(MIN_PT, pos.base - step * 0.5)
            os_ = max(MIN_PT, min(org.base, ps))
            height = (_line_count(pos.text(), BOX_WIDTH_PT, ps, pos.bold) * ps
                      + _line_count(org.text(), BOX_WIDTH_PT, os_, org.bold) * os_) * LINE_FACTOR
            if height <= BOX_TEXT_PT or ps == MIN_PT:
                break
        pos.resize(ps)
        org.resize(os_)

    permit = found.get("permit_text")
    if permit:
        permit.resize(next(
            s for s in _sizes_down(permit.base)
            if _line_count(permit.text(), PERMIT_WIDTH_PT, s, permit.bold) * s * LINE_FACTOR
            <= PERMIT_HEIGHT_PT or s == MIN_PT
        ))

def _ensure_trailing_p(tc):
    """Ячейка обязана заканчиваться абзацем, иначе Word считает файл битым."""
    if len(tc) == 0 or tc[-1].tag != qn("w:p"):
        tc.append(OxmlElement("w:p"))


def _append_before_sectpr(doc, element):
    body = doc.element.body
    sect = body.find(qn("w:sectPr"))
    if sect is not None:
        sect.addprevious(element)
    else:
        body.append(element)


def create_certificates_file(output_path, participants_with_groups, info):
    template_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), TEMPLATE_NAME
    )
    if not os.path.isfile(template_path):
        raise FileNotFoundError(
            f"Рядом с программой нет файла {TEMPLATE_NAME}.\n"
            "Создай шаблон с плейсхолдерами: одна таблица = одно удостоверение."
        )

    doc = Document(template_path)
    if not doc.tables:
        raise ValueError(f"В шаблоне {TEMPLATE_NAME} не найдена таблица удостоверения.")

    # Запоминаем "чистую" таблицу шаблона и снимаем обтекание "вокруг" -
    # иначе копии удостоверений накладываются друг на друга.
    clean_tbl = copy.deepcopy(doc.tables[0]._tbl)
    _tblPr = clean_tbl.find(qn("w:tblPr"))
    if _tblPr is not None:
        _ppr = _tblPr.find(qn("w:tblPPr"))
        if _ppr is not None:
            _tblPr.remove(_ppr)
    suppress_hyphens_for(clean_tbl, ("surname", "name", "patronymic", "position", "org"))

    # Полностью очищаем тело документа (кроме свойств страницы),
    # чтобы не оставалось лишних абзацев и пустых половин листа.
    body = doc.element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)

    parts = _date_parts(info.get("date"))
    if parts:
        day, month_name, year = parts
    else:
        now = datetime.now()
        day, month_name, year = now.day, _MONTHS[now.month - 1], now.year
    issue_date = f"«{day}» {month_name} {year} г."

    number = str(info.get("start_number", "")).strip()
    total = sum(len(groups) for _, groups in participants_with_groups)
    done = 0

    for participant, groups in participants_with_groups:
        words = (participant.get("name") or "").split()
        surname = words[0] if words else ""
        first_name = words[1] if len(words) > 1 else ""
        patronymic = " ".join(words[2:]) if len(words) > 2 else ""
        position = (participant.get("position") or "").strip()
        org = (participant.get("subdivision") or "").strip() or info.get("organization", "")
        texts = participant.get("permit_text_by_group", {}) or {}

        for grp in groups:
            try:
                valid_year = int(year) + (5 if grp == "3" else 3)
            except (TypeError, ValueError):
                valid_year = year

            mapping = {
                "number": str(number),
                "surname": surname,
                "name": first_name,
                "patronymic": patronymic,
                "position": position,
                "org": org,
                "date_issue": issue_date,
                "date_valid": f"«{day}» {month_name} {valid_year} г.",
                "group": grp,
                "permit_text": texts.get(grp) or texts.get("") or participant.get("permit_text", ""),
                "protocol_number": str(info.get("protocol_number", "")),
                "protocol_date": issue_date,
                "hours": PRACTICE_HOURS,
            }

            tbl = copy.deepcopy(clean_tbl)
            _strip_ids(tbl, done)
            targets = _find_targets(tbl)
            _fill_element(tbl, mapping)
            _fit_text(targets)
            _fix_geometry(tbl)
            for tc in tbl.iter(qn("w:tc")):
                _ensure_trailing_p(tc)

            _append_before_sectpr(doc, tbl)

            done += 1
            if done < total:
                if done % 2 == 0:
                    doc.add_page_break()          # 2 удостоверения на страницу
                else:
                    spacer = doc.add_paragraph("")
                    spacer.paragraph_format.space_after = Pt(6)
            number = next_number(number)          # у каждого удостоверения свой номер

    doc.save(output_path)