#!/usr/bin/env python3
"""
применить_стили.py

Просто положи этот скрипт в папку рядом с .docx файлом и запусти.
Скрипт сам найдёт документ и подготовит его для Calibre (FB2/EPUB).

Результат сохраняется как: имя_файла_calibre.docx
Оригинал не трогается.
"""

import sys
import os
import glob
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
# НАСТРОЙКИ — можно менять под себя
# ============================================================

BODY_FONT    = "Times New Roman"
BODY_SIZE_PT = 12
LINE_SPACING = 1.5      # межстрочный интервал
FIRST_LINE   = 1.25     # красная строка в см

# ============================================================

def cm_to_twips(cm):
    return int(cm * 567)

def pt_to_half(pt):
    return int(pt * 2)

def _get_or_add(parent, tag):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el

def _remove_if_exists(parent, tag):
    el = parent.find(qn(tag))
    if el is not None:
        parent.remove(el)


def set_page_layout(doc):
    for section in doc.sections:
        section.page_width    = Cm(21.0)
        section.page_height   = Cm(29.7)
        section.top_margin    = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin   = Cm(2.0)
        section.right_margin  = Cm(2.0)


def set_style_run(style, font, size_pt, bold=False, italic=False):
    s = style.element
    rPr = _get_or_add(s, "w:rPr")
    for tag in ["w:rFonts","w:sz","w:szCs","w:b","w:i","w:lang"]:
        _remove_if_exists(rPr, tag)

    rf = OxmlElement("w:rFonts")
    rf.set(qn("w:ascii"), font); rf.set(qn("w:hAnsi"), font); rf.set(qn("w:cs"), font)
    rPr.append(rf)

    sz = OxmlElement("w:sz");   sz.set(qn("w:val"),  str(pt_to_half(size_pt))); rPr.append(sz)
    sc = OxmlElement("w:szCs"); sc.set(qn("w:val"),  str(pt_to_half(size_pt))); rPr.append(sc)
    if bold:   rPr.append(OxmlElement("w:b"))
    if italic: rPr.append(OxmlElement("w:i"))

    lg = OxmlElement("w:lang")
    lg.set(qn("w:val"),  "ru-RU")
    lg.set(qn("w:bidi"), "ar-SA")
    rPr.append(lg)


def set_style_para(style, align="both", first_line=0, left=0, right=0,
                   space_before=0, space_after=0, line_spacing=288,
                   keep_with_next=False, page_break_before=False, outline_lvl=None):
    s = style.element
    pPr = _get_or_add(s, "w:pPr")
    for tag in ["w:jc","w:ind","w:spacing","w:keepNext","w:pageBreakBefore","w:outlineLvl"]:
        _remove_if_exists(pPr, tag)

    jc = OxmlElement("w:jc"); jc.set(qn("w:val"), align); pPr.append(jc)

    ind = OxmlElement("w:ind")
    if first_line: ind.set(qn("w:firstLine"), str(first_line))
    if left:       ind.set(qn("w:left"),      str(left))
    if right:      ind.set(qn("w:right"),     str(right))
    pPr.append(ind)

    sp = OxmlElement("w:spacing")
    sp.set(qn("w:before"),   str(space_before))
    sp.set(qn("w:after"),    str(space_after))
    sp.set(qn("w:line"),     str(line_spacing))
    sp.set(qn("w:lineRule"), "auto")
    pPr.append(sp)

    if keep_with_next:      pPr.append(OxmlElement("w:keepNext"))
    if page_break_before:   pPr.append(OxmlElement("w:pageBreakBefore"))
    if outline_lvl is not None:
        ol = OxmlElement("w:outlineLvl"); ol.set(qn("w:val"), str(outline_lvl)); pPr.append(ol)


def apply_styles(doc):
    LINE = int(240 * LINE_SPACING)
    FIRST = cm_to_twips(FIRST_LINE)

    # docDefaults
    styles_root = doc.styles.element
    dd  = _get_or_add(styles_root, "w:docDefaults")
    rpd = _get_or_add(dd,  "w:rPrDefault")
    rp  = _get_or_add(rpd, "w:rPr")
    for tag in ["w:rFonts","w:sz","w:szCs","w:lang"]:
        _remove_if_exists(rp, tag)
    rf = OxmlElement("w:rFonts")
    rf.set(qn("w:ascii"), BODY_FONT); rf.set(qn("w:hAnsi"), BODY_FONT); rf.set(qn("w:cs"), BODY_FONT)
    rp.append(rf)
    sz = OxmlElement("w:sz");   sz.set(qn("w:val"),  str(pt_to_half(BODY_SIZE_PT))); rp.append(sz)
    sc = OxmlElement("w:szCs"); sc.set(qn("w:val"),  str(pt_to_half(BODY_SIZE_PT))); rp.append(sc)
    lg = OxmlElement("w:lang"); lg.set(qn("w:val"), "ru-RU"); lg.set(qn("w:bidi"), "ar-SA"); rp.append(lg)

    # Normal
    try:    n = doc.styles["Normal"]
    except: n = doc.styles.add_style("Normal", 1)
    set_style_run(n, BODY_FONT, BODY_SIZE_PT)
    set_style_para(n, align="both", first_line=FIRST, space_before=0, space_after=0, line_spacing=LINE)
    print("  [+] Normal: по ширине, красная строка, интервал 1.5")

    # Heading 1 — глава, с разрывом страницы (ключевое для Calibre)
    try:    h1 = doc.styles["Heading 1"]
    except: h1 = doc.styles.add_style("Heading 1", 1)
    set_style_run(h1, BODY_FONT, 16, bold=True)
    set_style_para(h1, align="center",
                   space_before=cm_to_twips(0.5), space_after=cm_to_twips(0.3),
                   line_spacing=int(240*1.2),
                   keep_with_next=True, page_break_before=True, outline_lvl=0)
    print("  [+] Heading 1: 16pt жирный, по центру, разрыв перед главой")

    # Heading 2
    try:    h2 = doc.styles["Heading 2"]
    except: h2 = doc.styles.add_style("Heading 2", 1)
    set_style_run(h2, BODY_FONT, 14, bold=True)
    set_style_para(h2, align="center",
                   space_before=cm_to_twips(0.4), space_after=cm_to_twips(0.2),
                   line_spacing=int(240*1.2),
                   keep_with_next=True, outline_lvl=1)
    print("  [+] Heading 2: 14pt жирный, по центру")

    # Heading 3
    try:    h3 = doc.styles["Heading 3"]
    except: h3 = doc.styles.add_style("Heading 3", 1)
    set_style_run(h3, BODY_FONT, 13, bold=True, italic=True)
    set_style_para(h3, align="left",
                   space_before=cm_to_twips(0.3), space_after=cm_to_twips(0.15),
                   line_spacing=int(240*1.2),
                   keep_with_next=True, outline_lvl=2)
    print("  [+] Heading 3: 13pt жирный курсив")

    # Title / Subtitle
    for sname, size, bold, italic in [("Title",20,True,False),("Subtitle",14,False,True)]:
        try:
            s = doc.styles[sname]
            set_style_run(s, BODY_FONT, size, bold=bold, italic=italic)
            set_style_para(s, align="center",
                           space_before=cm_to_twips(0.5), space_after=cm_to_twips(0.3),
                           line_spacing=int(240*1.3))
        except KeyError: pass

    # Цитаты
    for sname in ["Quote","Intense Quote","Block Text"]:
        try:
            s = doc.styles[sname]
            set_style_run(s, BODY_FONT, BODY_SIZE_PT, italic=True)
            set_style_para(s, align="both",
                           left=cm_to_twips(2.0), right=cm_to_twips(2.0),
                           space_before=cm_to_twips(0.2), space_after=cm_to_twips(0.2),
                           line_spacing=LINE)
        except KeyError: pass

    # Сноски
    try:
        fn = doc.styles["Footnote Text"]
        set_style_run(fn, BODY_FONT, 10)
        set_style_para(fn, align="both", line_spacing=240)
    except KeyError: pass


def clean_paragraphs(doc):
    """Убирает мусор который мешает Calibre: пустые строки, Tab-отступы, прямое форматирование."""
    to_delete = []
    fixed = 0

    for para in doc.paragraphs:
        style_name = para.style.name if para.style else "Normal"
        is_heading = style_name.startswith("Heading") or "аголовок" in style_name

        # Пустые абзацы-разделители удаляем — в стилях уже есть отступы
        if not para.text.strip() and not is_heading:
            to_delete.append(para._element)
            continue

        if is_heading:
            continue

        pf = para.paragraph_format
        pf.alignment         = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.space_before      = Pt(0)
        pf.space_after       = Pt(0)
        pf.first_line_indent = None  # берётся из стиля Normal

        # Убираем Tab в начале абзаца (ручной отступ)
        for run in para.runs:
            if run.text.startswith('\t') or run.text.startswith('   '):
                run.text = run.text.lstrip('\t').lstrip(' ')
            break

        fixed += 1

    removed = 0
    for elem in to_delete:
        p = elem.getparent()
        if p is not None:
            p.remove(elem)
            removed += 1

    print(f"  [+] Абзацев обработано: {fixed}, пустых строк удалено: {removed}")


def find_docx(script_dir):
    """Ищет .docx в папке скрипта, исключая уже обработанные (_calibre.docx)."""
    all_docx = glob.glob(os.path.join(script_dir, "*.docx"))
    candidates = [
        f for f in all_docx
        if not os.path.basename(f).endswith("_calibre.docx")
        and not os.path.basename(f).startswith("~$")  # временные файлы Word
    ]
    return candidates


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"\n📁 Папка: {script_dir}")

    candidates = find_docx(script_dir)

    if not candidates:
        print("\n❌ .docx файлов не найдено в этой папке.")
        print("   Положи .docx рядом со скриптом и запусти снова.")
        input("\nНажми Enter для выхода...")
        sys.exit(1)

    if len(candidates) == 1:
        input_path = candidates[0]
        print(f"📖 Найден документ: {os.path.basename(input_path)}")
    else:
        print(f"\n📚 Найдено несколько документов:")
        for i, f in enumerate(candidates, 1):
            print(f"   {i}. {os.path.basename(f)}")
        print(f"   0. Обработать все")
        choice = input("\nКакой обработать? (номер или 0 для всех): ").strip()
        if choice == "0":
            for path in candidates:
                process_file(path, script_dir)
            print("\n✅ Все файлы обработаны!")
            input("Нажми Enter для выхода...")
            return
        else:
            try:
                input_path = candidates[int(choice) - 1]
            except (ValueError, IndexError):
                print("Неверный выбор.")
                input("Нажми Enter для выхода...")
                sys.exit(1)

    process_file(input_path, script_dir)
    input("\nНажми Enter для выхода...")


def process_file(input_path, script_dir):
    base = os.path.splitext(os.path.basename(input_path))[0]
    output_path = os.path.join(script_dir, f"{base}_calibre.docx")

    print(f"\n⚙️  Обработка: {os.path.basename(input_path)}")
    print(f"   → {os.path.basename(output_path)}")

    doc = Document(input_path)
    set_page_layout(doc)
    apply_styles(doc)
    clean_paragraphs(doc)
    doc.core_properties.language = "ru"
    doc.save(output_path)

    print(f"\n✅ Готово! Сохранён: {os.path.basename(output_path)}")
    print("   Следующий шаг: открой в Calibre → Конвертировать → FB2 или EPUB")


if __name__ == "__main__":
    main()
