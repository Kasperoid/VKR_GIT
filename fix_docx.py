import glob
import os
import re
import sys
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph

# Основные настраиваемые параметры
FIGURE_SPACE_PT = 21  # 1 пустая строка (14 пт × 1,5)
FIRST_LINE_DXA = str(int(round(1.25 * 566.929)))  # 709 dxa = 1,25 см
TABLE_DASH = "–"  # среднее тире (замените на "-", если нужен обычный дефис)

# Ширины столбцов в % (сумма = 100) для таблиц, которые Pandoc сжимает слишком узко
CUSTOM_TABLE_WIDTHS = {
    "3.1": [18, 20, 38, 24],
    "3.2": [22, 20, 36, 22],
}

MATH_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"

CAPTION_STYLE_KEYS = {
    "imagecaption", "image caption",
    "caption", "название",
    "captionedfigure", "captioned figure",
    "figuretitle", "figure title",
}
TABLE_CAPTION_STYLES = {"tablecaption", "table caption"}
CODE_P_STYLES = {"sourcecode", "source code", "code", "verbatim"}
CODE_R_STYLES = {"verbatimchar", "verbatim char", "code", "source code"}

CAPTION_TEXT_RE = re.compile(
    r"^\s*(Рисунок|Рис\.?|Figure|Fig\.?)\s*\d+",
    re.IGNORECASE,
)
FIG_CAP_NORM_RE = re.compile(
    r"^(\s*(?:Рисунок|Рис\.?))[\s\u00a0]+([A-Za-zА-Яа-я]\.\d+|\d+(?:\.\d+)?)\s*[—–\-:]\s*(.+)$",
    re.IGNORECASE,
)
TABLE_CAPTION_PREFIX_RE = re.compile(
    r"^\s*Таблица[\s\u00a0]+([A-Za-zА-Яа-я]\.\d+|\d+(?:\.\d+)?)\s*[—–\-:.]?\s*(.+)$",
    re.IGNORECASE,
)
CHAPTER_RE = re.compile(r"^\s*(\d+)\.?\s+(.+)$")
CHAPTER_NUM_RE = re.compile(r"^\s*(?:ГЛАВА\s+)?(\d+)\.?\s+", re.IGNORECASE)
APP_H1_RE = re.compile(r"^\s*(?:Appendix|Приложение)\s+[AА]\s*[—–\-.:]?\s*(.+)$", re.IGNORECASE)
APP_H2_RE = re.compile(r"^\s*[AА]\.(\d+)\.?\s+(.+)$")
APP_REF_RE = re.compile(r"(?:раздел|section|приложение|appendix)([\s\u00a0]+)[AА]\.(\d+)", re.IGNORECASE)
CONCLUSION_HEADING_RE = re.compile(
    r"^\s*Вывод(?:ы)?\s+по\s+(?:параграфу|разделу|главе|\S+\s+главе)\b",
    re.IGNORECASE,
)


def get_style_name_map(doc):
    mapping = {}
    for s in doc.styles.element.findall(qn("w:style")):
        sid = s.get(qn("w:styleId"))
        name_el = s.find(qn("w:name"))
        if sid:
            sname = name_el.get(qn("w:val")) if name_el is not None else sid
            mapping[sid] = sname
    return mapping


def get_style_ids_by_name(doc, target_name):
    ids = set()
    for s in doc.styles.element.findall(qn("w:style")):
        sid = s.get(qn("w:styleId"))
        name_el = s.find(qn("w:name"))
        if sid and name_el is not None and name_el.get(qn("w:val"), "").lower() == target_name.lower():
            ids.add(sid)
    return ids


def get_p_text(p_el):
    return "".join(t.text for t in p_el.iter(qn("w:t")) if t.text)


def get_p_style(p_el, style_map):
    sid = ""
    for pPr in p_el.findall(qn("w:pPr")):
        pStyle = pPr.find(qn("w:pStyle"))
        if pStyle is not None and pStyle.get(qn("w:val")):
            sid = pStyle.get(qn("w:val"))
    return sid, style_map.get(sid, "")


def has_image(el):
    return (
        el.find(".//" + qn("w:drawing")) is not None
        or el.find(".//" + qn("w:pict")) is not None
        or el.find(".//" + qn("w:object")) is not None
    )


def is_caption_p(p_el, style_map):
    sid, sname = get_p_style(p_el, style_map)
    if sid.lower() in CAPTION_STYLE_KEYS or sname.lower() in CAPTION_STYLE_KEYS:
        return True
    return bool(CAPTION_TEXT_RE.match(get_p_text(p_el)))


def is_table_caption_p(p_el, style_map):
    sid, sname = get_p_style(p_el, style_map)
    return sid.lower() in TABLE_CAPTION_STYLES or sname.lower() in TABLE_CAPTION_STYLES


def is_code_run(r_el, style_map):
    rPr = r_el.find(qn("w:rPr"))
    if rPr is None:
        return False
    rStyle = rPr.find(qn("w:rStyle"))
    if rStyle is not None:
        rsid = rStyle.get(qn("w:val"), "")
        rsname = style_map.get(rsid, rsid)
        if rsid.lower() in CODE_R_STYLES or rsname.lower() in CODE_R_STYLES:
            return True
    return False


def is_display_formula_p(p_el):
    if p_el.find(f".//{MATH_NS}oMathPara") is not None:
        return True
    if p_el.find(f".//{MATH_NS}oMath") is not None:
        if not get_p_text(p_el).strip():
            return True
    return False


def reset_pPr_keeping_style(p_el):
    style_id = None
    for pPr in p_el.findall(qn("w:pPr")):
        pStyle = pPr.find(qn("w:pStyle"))
        if pStyle is not None and pStyle.get(qn("w:val")):
            style_id = pStyle.get(qn("w:val"))

    for pPr in list(p_el.findall(qn("w:pPr"))):
        p_el.remove(pPr)

    new_pPr = OxmlElement("w:pPr")
    if style_id:
        pStyle = OxmlElement("w:pStyle")
        pStyle.set(qn("w:val"), style_id)
        new_pPr.append(pStyle)
    p_el.insert(0, new_pPr)


def normalize_figure_caption_text(p_el):
    t_els = list(p_el.iter(qn("w:t")))
    if not t_els:
        return
    full_text = "".join(t.text or "" for t in t_els)
    m = FIG_CAP_NORM_RE.match(full_text)
    if not m:
        return
    prefix, num, title = m.group(1).strip(), m.group(2), m.group(3).strip()
    t_els[0].text = f"{prefix}\u00a0{num} \u2013 {title}"
    t_els[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for extra in t_els[1:]:
        extra.text = ""


def center_paragraph_el(p_el, keep_with_next=False, space_before_pt=0, space_after_pt=0):
    reset_pPr_keeping_style(p_el)
    pw = Paragraph(p_el, None)
    pw.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pw.paragraph_format.first_line_indent = Cm(0)
    pw.paragraph_format.left_indent = Cm(0)
    pw.paragraph_format.right_indent = Cm(0)
    pw.paragraph_format.line_spacing = 1.0
    pw.paragraph_format.space_before = Pt(space_before_pt)
    pw.paragraph_format.space_after = Pt(space_after_pt)
    if keep_with_next:
        pw.paragraph_format.keep_with_next = True


def set_tnr_rPr(parent_el):
    rPr = parent_el.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        parent_el.append(rPr)
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme", "hint"):
        rFonts.attrib.pop(qn(f"w:{attr}"), None)
    for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
        rFonts.set(qn(f"w:{attr}"), "Times New Roman")
    color = rPr.find(qn("w:color"))
    if color is None:
        color = OxmlElement("w:color")
        rPr.append(color)
    for attr in ("themeColor", "themeTint", "themeShade"):
        color.attrib.pop(qn(f"w:{attr}"), None)
    color.set(qn("w:val"), "000000")


def set_lvl_suff_space(lvl_el):
    for old in list(lvl_el.findall(qn("w:suff"))):
        lvl_el.remove(old)
    suff = OxmlElement("w:suff")
    suff.set(qn("w:val"), "space")
    before_tags = {
        qn("w:start"), qn("w:numFmt"), qn("w:lvlRestart"),
        qn("w:pStyle"), qn("w:isLgl"),
    }
    insert_pos = 0
    for i, ch in enumerate(lvl_el):
        if ch.tag in before_tags:
            insert_pos = i + 1
    lvl_el.insert(insert_pos, suff)


def fix_lists(doc):
    # 1. Стили списков в styles.xml: отступ слева 0 см, первая строка 1,25 см
    for st in doc.styles:
        if not st.name or not hasattr(st, "paragraph_format"):
            continue
        lname = st.name.lower()
        if lname == "compact" or lname.startswith("list"):
            pf = st.paragraph_format
            pf.left_indent = Cm(0)
            pf.right_indent = Cm(0)
            pf.first_line_indent = Cm(1.25)

    # 2. Все уровни списков в word/numbering.xml (убираем выступ Pandoc и меняем • на –)
    try:
        numbering_el = doc.part.numbering_part.numbering_definitions._numbering
    except NotImplementedError:
        numbering_el = None

    if numbering_el is not None:
        for lvl in numbering_el.findall(".//" + qn("w:lvl")):
            set_lvl_suff_space(lvl)
            numFmt = lvl.find(qn("w:numFmt"))
            fmt_val = numFmt.get(qn("w:val")) if numFmt is not None else ""
            lvlText = lvl.find(qn("w:lvlText"))

            if fmt_val == "bullet" and lvlText is not None and lvlText.get(qn("w:val"), "").strip():
                lvlText.set(qn("w:val"), "–")
                set_tnr_rPr(lvl)
            elif fmt_val == "decimal":
                set_tnr_rPr(lvl)

            pPr = lvl.find(qn("w:pPr"))
            if pPr is None:
                pPr = OxmlElement("w:pPr")
                rPr = lvl.find(qn("w:rPr"))
                if rPr is not None:
                    lvl.insert(list(lvl).index(rPr), pPr)
                else:
                    lvl.append(pPr)
            for tag in ("w:tabs", "w:ind"):
                for old in list(pPr.findall(qn(tag))):
                    pPr.remove(old)
            ind = OxmlElement("w:ind")
            ind.set(qn("w:left"), "0")
            ind.set(qn("w:firstLine"), FIRST_LINE_DXA)
            pPr.append(ind)

    # 3. Сами абзацы списков в document.xml
    for p_el in doc.element.body.iter(qn("w:p")):
        pPr = p_el.find(qn("w:pPr"))
        if pPr is not None and pPr.find(qn("w:numPr")) is not None:
            pw = Paragraph(p_el, None)
            pw.paragraph_format.left_indent = Cm(0)
            pw.paragraph_format.right_indent = Cm(0)
            pw.paragraph_format.first_line_indent = Cm(1.25)


def insert_tblPr_child(tblPr, new_el, before_tags):
    for old in list(tblPr.findall(new_el.tag)):
        tblPr.remove(old)
    before_qns = {qn(t) for t in before_tags}
    pos = len(tblPr)
    for i, ch in enumerate(tblPr):
        if ch.tag in before_qns:
            pos = i
            break
    tblPr.insert(pos, new_el)


def set_table_col_widths(tbl_el, col_pcts, text_width_dxa):
    total = sum(col_pcts)
    dxa_widths = [int(round(text_width_dxa * p / total)) for p in col_pcts]
    pct_widths = [int(round(5000 * p / total)) for p in col_pcts]

    tblPr = tbl_el.find(qn("w:tblPr"))
    if tblPr is not None:
        tblW = tblPr.find(qn("w:tblW"))
        if tblW is not None:
            tblW.set(qn("w:w"), "5000")
            tblW.set(qn("w:type"), "pct")

    tblGrid = tbl_el.find(qn("w:tblGrid"))
    if tblGrid is not None:
        grid_cols = tblGrid.findall(qn("w:gridCol"))
        if len(grid_cols) == len(dxa_widths):
            for gc, w in zip(grid_cols, dxa_widths):
                gc.set(qn("w:w"), str(w))

    for tr in tbl_el.findall(qn("w:tr")):
        tcs = tr.findall(qn("w:tc"))
        if len(tcs) != len(pct_widths):
            continue
        for tc, w_pct in zip(tcs, pct_widths):
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is None:
                tcPr = OxmlElement("w:tcPr")
                tc.insert(0, tcPr)
            tcW = tcPr.find(qn("w:tcW"))
            if tcW is None:
                tcW = OxmlElement("w:tcW")
                tcPr.insert(0, tcW)
            tcW.set(qn("w:w"), str(w_pct))
            tcW.set(qn("w:type"), "pct")


def apply_black_borders(tbl_el):
    tblPr = tbl_el.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl_el.insert(0, tblPr)
      
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    insert_tblPr_child(
        tblPr,
        jc,
        before_tags=(
            "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd",
            "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription",
        ),
    )

    borders_xml = f"""
    <w:tblBorders {nsdecls('w')}>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tblBorders>
    """
    insert_tblPr_child(
        tblPr,
        parse_xml(borders_xml),
        before_tags=("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription"),
    )

    for row_idx, tr in enumerate(tbl_el.findall(qn("w:tr"))):
        # Запрещаем разрыв строки пополам и закрепляем повторение шапки при переносе
        trPr = tr.find(qn("w:trPr"))
        if trPr is None:
            trPr = OxmlElement("w:trPr")
            tr.insert(0, trPr)
        if trPr.find(qn("w:cantSplit")) is None:
            trPr.append(OxmlElement("w:cantSplit"))
        if row_idx == 0 and trPr.find(qn("w:tblHeader")) is None:
            trPr.append(OxmlElement("w:tblHeader"))

        for p_el in tr.iter(qn("w:p")):
            pPr = p_el.find(qn("w:pPr"))
            is_list = pPr is not None and pPr.find(qn("w:numPr")) is not None
            pw = Paragraph(p_el, None)

            if not is_list:
                pw.paragraph_format.first_line_indent = Cm(0)
                pw.paragraph_format.left_indent = Cm(0)
                pw.paragraph_format.right_indent = Cm(0)

            # Шапка таблицы: по центру и полужирный шрифт
            if row_idx == 0:
                pw.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in pw.runs:
                    run.bold = True


def fix_tables(doc, style_map):
    # 1. Настраиваем стиль Table Caption в styles.xml
    for st in doc.styles:
        if st.name and st.name.lower() in TABLE_CAPTION_STYLES and hasattr(st, "paragraph_format"):
            pf = st.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf.first_line_indent = Cm(1.25)
            pf.left_indent = Cm(0)
            pf.right_indent = Cm(0)
            pf.line_spacing = 1.5
            pf.keep_with_next = True

    sec = doc.sections[0]
    text_width_dxa = int((sec.page_width - sec.left_margin - sec.right_margin) / 635)

    h1_ids = get_style_ids_by_name(doc, "heading 1") | {"Heading1", "1"}
    body = doc.element.body
    children = list(body)

    current_chapter = 0
    table_num_in_chapter = 0
    global_table_num = 0
    tables_fixed = 0

    for idx, ch in enumerate(children):
        if ch.tag == qn("w:p"):
            sid, _ = get_p_style(ch, style_map)
            if sid in h1_ids:
                m = CHAPTER_NUM_RE.match(get_p_text(ch))
                if m:
                    current_chapter = int(m.group(1))
                    table_num_in_chapter = 0
        elif ch.tag == qn("w:tbl"):
            apply_black_borders(ch)
            tables_fixed += 1
            global_table_num += 1
            table_num_in_chapter += 1

            num_str = f"{current_chapter}.{table_num_in_chapter}" if current_chapter > 0 else str(global_table_num)

            # Ищем абзац с названием таблицы прямо перед ней (пропуская закладки)
            j = idx - 1
            while j >= 0 and children[j].tag in (qn("w:bookmarkStart"), qn("w:bookmarkEnd")):
                j -= 1
            if j >= 0 and children[j].tag == qn("w:p") and is_table_caption_p(children[j], style_map):
                cap_p = children[j]
                t_els = list(cap_p.iter(qn("w:t")))
                full_text = "".join(t.text or "" for t in t_els).strip()
                if t_els and full_text:
                    m = TABLE_CAPTION_PREFIX_RE.match(full_text)
                    if m:
                        num_str, title = m.group(1), m.group(2).strip()
                    else:
                        title = full_text
                    t_els[0].text = f"Таблица\u00a0{num_str} {TABLE_DASH} {title}"
                    t_els[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    for extra in t_els[1:]:
                        extra.text = ""

                reset_pPr_keeping_style(cap_p)
                pw = Paragraph(cap_p, None)
                pw.alignment = WD_ALIGN_PARAGRAPH.LEFT
                pw.paragraph_format.first_line_indent = Cm(1.25)
                pw.paragraph_format.left_indent = Cm(0)
                pw.paragraph_format.right_indent = Cm(0)
                pw.paragraph_format.line_spacing = 1.5
                pw.paragraph_format.keep_with_next = True

            # Применяем заданные ширины столбцов (если указаны в CUSTOM_TABLE_WIDTHS)
            if num_str in CUSTOM_TABLE_WIDTHS:
                set_table_col_widths(ch, CUSTOM_TABLE_WIDTHS[num_str], text_width_dxa)

            # Отступ в 1 абзац после таблицы (вешаем space_before на следующий абзац)
            k = idx + 1
            while k < len(children) and children[k].tag in (qn("w:bookmarkStart"), qn("w:bookmarkEnd")):
                k += 1
            if k < len(children) and children[k].tag == qn("w:p"):
                next_p = children[k]
                next_sid, _ = get_p_style(next_p, style_map)
                if next_sid not in h1_ids:
                    next_pw = Paragraph(next_p, None)
                    target = Pt(FIGURE_SPACE_PT)
                    if next_pw.paragraph_format.space_before is None or next_pw.paragraph_format.space_before < target:
                        next_pw.paragraph_format.space_before = target

    return tables_fixed


def fix_formulas(doc):
    for p_el in doc.element.body.iter(qn("w:p")):
        if not is_display_formula_p(p_el):
            continue
        pw = Paragraph(p_el, None)
        pw.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pw.paragraph_format.first_line_indent = Cm(0)
        pw.paragraph_format.left_indent = Cm(0)
        pw.paragraph_format.right_indent = Cm(0)
        pw.paragraph_format.space_before = Pt(FIGURE_SPACE_PT)
        pw.paragraph_format.space_after = Pt(FIGURE_SPACE_PT)


def prefix_chapters(doc, prefix="ГЛАВА"):
    h1_ids = get_style_ids_by_name(doc, "heading 1") | {"Heading1", "1"}

    for p_el in doc.element.body.iter(qn("w:p")):
        pPr = p_el.find(qn("w:pPr"))
        if pPr is None:
            continue
        pStyle = pPr.find(qn("w:pStyle"))
        if pStyle is None or pStyle.get(qn("w:val")) not in h1_ids:
            continue

        t_els = list(p_el.iter(qn("w:t")))
        full_text = "".join(t.text or "" for t in t_els)
        if full_text.lstrip().upper().startswith(prefix.upper()):
            continue

        m = CHAPTER_RE.match(full_text)
        if not m:
            continue

        num, title = m.group(1), m.group(2)
        if t_els:
            t_els[0].text = f"{prefix} {num}. {title}"
            for extra in t_els[1:]:
                extra.text = ""


def format_appendices(doc):
    h1_ids = get_style_ids_by_name(doc, "heading 1") | {"Heading1", "1"}
    h2_ids = get_style_ids_by_name(doc, "heading 2") | {"Heading2", "21"}

    for p_el in doc.element.body.iter(qn("w:p")):
        pPr = p_el.find(qn("w:pPr"))
        sid = None
        if pPr is not None:
            pStyle = pPr.find(qn("w:pStyle"))
            if pStyle is not None:
                sid = pStyle.get(qn("w:val"))

        t_els = list(p_el.iter(qn("w:t")))
        if not t_els:
            continue
        full_text = "".join(t.text or "" for t in t_els)

        # 1. Заголовок 1-го уровня: убираем "Appendix A — ", оставляем только название
        if sid in h1_ids:
            m = APP_H1_RE.match(full_text)
            if m:
                t_els[0].text = m.group(1).strip()
                for extra in t_els[1:]:
                    extra.text = ""
                continue

        # 2. Заголовки 2-го уровня: "A.1 Название" -> "Приложение А.1 — Название"
        if sid in h2_ids:
            m = APP_H2_RE.match(full_text)
            if m:
                num, title = m.group(1), m.group(2).strip()
                t_els[0].text = f"Приложение А.{num} — {title}"
                for extra in t_els[1:]:
                    extra.text = ""
                continue

        # 3. Перекрестные ссылки в тексте и таблицах: "Приложение A.1" -> "Приложение А.1"
        for t in t_els:
            if t.text and ("A." in t.text or "А." in t.text):
                t.text = APP_REF_RE.sub(lambda m: f"Приложение{m.group(1)}А.{m.group(2)}", t.text)


def fix_conclusions(doc):
    target = Pt(FIGURE_SPACE_PT)
    for p_el in doc.element.body.findall(qn("w:p")):
        txt = get_p_text(p_el).strip()
        if CONCLUSION_HEADING_RE.match(txt) and len(txt) < 80:
            pw = Paragraph(p_el, None)
            if pw.paragraph_format.space_before is None or pw.paragraph_format.space_before < target:
                pw.paragraph_format.space_before = target
            pw.paragraph_format.keep_with_next = True


def fix_quotes(doc, style_map):
    for p_el in doc.element.body.iter(qn("w:p")):
        sid, sname = get_p_style(p_el, style_map)
        if sid.lower() in CODE_P_STYLES or sname.lower() in CODE_P_STYLES:
            continue

        in_straight_quote = False
        for r_el in p_el.iter(qn("w:r")):
            if is_code_run(r_el, style_map):
                continue
            for t_el in r_el.findall(qn("w:t")):
                txt = t_el.text
                if not txt or not any(q in txt for q in ('"', "“", "”", "„")):
                    continue

                out = []
                for idx, ch in enumerate(txt):
                    if ch in ("“", "„"):
                        out.append("«")
                    elif ch == "”":
                        out.append("»")
                    elif ch == '"':
                        prev_ch = txt[idx - 1] if idx > 0 else ""
                        next_ch = txt[idx + 1] if idx + 1 < len(txt) else ""
                        if not in_straight_quote and (not prev_ch or prev_ch in " \u00a0([{—–-\t\n"):
                            out.append("«")
                            in_straight_quote = True
                        elif in_straight_quote:
                            out.append("»")
                            in_straight_quote = False
                        elif next_ch and next_ch not in " \u00a0.,;:!?)]}—–-\t\n":
                            out.append("«")
                            in_straight_quote = True
                        else:
                            out.append("»")
                    else:
                        out.append(ch)

                new_txt = "".join(out)
                if new_txt != txt:
                    t_el.text = new_txt
                    t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def fix_hyphenation_and_lang(doc, locale="ru-RU", hyphenate_caps=False, consecutive_limit=3):
    styles_el = doc.styles.element
    for lang_el in styles_el.findall(".//" + qn("w:lang")):
        if lang_el.get(qn("w:val")) in ("ru", "en-US", None, ""):
            lang_el.set(qn("w:val"), locale)
        if lang_el.get(qn("w:eastAsia")) in ("ru", "en-US", None, ""):
            lang_el.set(qn("w:eastAsia"), locale)
        if lang_el.get(qn("w:bidi")) in ("ru", "en-US", None, ""):
            lang_el.set(qn("w:bidi"), locale)

    docDefaults = styles_el.find(qn("w:docDefaults"))
    if docDefaults is not None:
        rPrDefault = docDefaults.find(qn("w:rPrDefault"))
        if rPrDefault is not None:
            rPr = rPrDefault.find(qn("w:rPr"))
            if rPr is None:
                rPr = OxmlElement("w:rPr")
                rPrDefault.append(rPr)
            lang_el = rPr.find(qn("w:lang"))
            if lang_el is None:
                lang_el = OxmlElement("w:lang")
                rPr.append(lang_el)
            lang_el.set(qn("w:val"), locale)
            lang_el.set(qn("w:eastAsia"), locale)
            lang_el.set(qn("w:bidi"), locale)

    settings = doc.settings.element
    theme_lang = settings.find(qn("w:themeFontLang"))
    if theme_lang is not None:
        theme_lang.set(qn("w:val"), locale)

    for tag in ("w:autoHyphenation", "w:doNotHyphenateCaps", "w:consecutiveHyphenLimit"):
        for old in list(settings.findall(qn(tag))):
            settings.remove(old)

    anchor = settings.find(qn("w:defaultTabStop"))
    idx = list(settings).index(anchor) + 1 if anchor is not None else 0

    settings.insert(idx, OxmlElement("w:autoHyphenation"))
    idx += 1
    if not hyphenate_caps:
        settings.insert(idx, OxmlElement("w:doNotHyphenateCaps"))
        idx += 1
    if consecutive_limit is not None:
        lim = OxmlElement("w:consecutiveHyphenLimit")
        lim.set(qn("w:val"), str(consecutive_limit))
        settings.insert(idx, lim)


def collect_target_docx_files():
    candidates = []
    out_dir = os.environ.get("QUARTO_PROJECT_OUTPUT_DIR", "_book")

    for arg in sys.argv[1:]:
        candidates.append(arg)

    for line in os.environ.get("QUARTO_PROJECT_OUTPUT_FILES", "").splitlines():
        p = line.strip().strip('"').strip("'")
        if not p:
            continue
        candidates.append(p)
        candidates.append(os.path.join(out_dir, p))

    for folder in {out_dir, "_book"}:
        if folder and os.path.isdir(folder):
            candidates.extend(glob.glob(os.path.join(folder, "**", "*.docx"), recursive=True))

    seen = set()
    result = []
    for path in candidates:
        if not path.lower().endswith(".docx"):
            continue
        name = os.path.basename(path)
        if name.startswith("~$") or "reference" in name.lower():
            continue
        abs_path = os.path.abspath(path)
        if abs_path not in seen and os.path.isfile(abs_path):
            seen.add(abs_path)
            result.append(abs_path)
    return result


def fix_docx(path):
    doc = Document(path)
    style_map = get_style_name_map(doc)
    body = doc.element.body

    unwrapped = 0
    images_fixed = 0
    captions_fixed = 0

    for st in doc.styles:
        if st.name and st.name.lower() in CAPTION_STYLE_KEYS and hasattr(st, "paragraph_format"):
            st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            st.paragraph_format.first_line_indent = Cm(0)
            st.paragraph_format.left_indent = Cm(0)
            st.paragraph_format.right_indent = Cm(0)

    # 1. Разворачиваем 1-колоночные таблицы-обертки рисунков Quarto в обычные абзацы
    for tbl_el in list(body.findall(qn("w:tbl"))):
        rows = tbl_el.findall(qn("w:tr"))
        if not rows:
            continue
        if max(len(r.findall(qn("w:tc"))) for r in rows) != 1:
            continue
        p_els = list(tbl_el.iter(qn("w:p")))
        if any(has_image(p) or is_caption_p(p, style_map) for p in p_els):
            parent = tbl_el.getparent()
            idx = list(parent).index(tbl_el)
            offset = 0
            for r in rows:
                for tc in r.findall(qn("w:tc")):
                    for child in list(tc):
                        if child.tag in (qn("w:p"), qn("w:bookmarkStart"), qn("w:bookmarkEnd")):
                            parent.insert(idx + offset, child)
                            offset += 1
            parent.remove(tbl_el)
            unwrapped += 1

    # 2. Центрируем картинки и подписи, форматируем «Рисунок X.Y – Название» и отступы
    for p_el in body.iter(qn("w:p")):
        txt = get_p_text(p_el)
        if has_image(p_el) and not txt.strip():
            center_paragraph_el(
                p_el,
                keep_with_next=True,
                space_before_pt=FIGURE_SPACE_PT,
                space_after_pt=6,
            )
            images_fixed += 1
        elif is_caption_p(p_el, style_map) and not has_image(p_el):
            normalize_figure_caption_text(p_el)
            center_paragraph_el(
                p_el,
                keep_with_next=False,
                space_before_pt=0,
                space_after_pt=FIGURE_SPACE_PT,
            )
            captions_fixed += 1

    # 3. Списки, таблицы, формулы, главы, приложения, выводы, кавычки и автопереносы
    fix_lists(doc)
    tables_fixed = fix_tables(doc, style_map)
    fix_formulas(doc)
    prefix_chapters(doc)
    format_appendices(doc)
    fix_conclusions(doc)
    fix_quotes(doc, style_map)
    fix_hyphenation_and_lang(doc)

    doc.save(path)
    print(
        f"[fix_docx] OK: {path} | "
        f"оберток: {unwrapped}, картинок: {images_fixed}, "
        f"подписей рис.: {captions_fixed}, таблиц: {tables_fixed}"
    )


if __name__ == "__main__":
    files = collect_target_docx_files()
    if not files:
        print("[fix_docx] ВНИМАНИЕ: .docx файлы в _book/ не найдены!")
    else:
        for f in files:
            fix_docx(f)
