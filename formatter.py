"""格式应用引擎 — 页面/段落/字符级格式统一"""

import os
import copy
import uuid
from docx import Document
from docx.shared import Pt, Mm, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree

from config import (
    PAGE_MARGINS, LINE_SPACING, PARA_SPACING_BEFORE, PARA_SPACING_AFTER,
    FONT_RULES,
)
from parser import parse_document
import utils

# 上传文件临时目录
UPLOAD_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uploads")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outputs")


def format_document(file_id: str) -> str:
    """
    对指定文件执行格式统一，返回修正后文件的 ID。
    原文件只读不写。
    """
    input_path = os.path.join(UPLOAD_DIR, f"{file_id}.docx")
    output_id = str(uuid.uuid4())
    output_path = os.path.join(OUTPUT_DIR, f"{output_id}.docx")

    doc = Document(input_path)
    structure = parse_document(input_path)

    # ── 0. 样式表默认间距归零（解决 Word 内置样式覆盖问题）──
    _reset_style_defaults(doc)

    # ── 1. 页面设置 ──
    _apply_page_setup(doc)

    # ── 1b. 全文段前段后间距归零 ──
    _reset_all_paragraph_spacing(doc)

    # ── 2. 逐段落格式设置 ──
    for info in structure["paragraphs"]:
        etype = info["element_type"]
        para = doc.paragraphs[info["index"]]

        if etype == "empty":
            # 检查是否包含图片，有图片则不处理，无图片则清空
            has_drawing = _para_has_drawing(para)
            if not has_drawing:
                _clear_paragraph(para)
            continue

        rules = FONT_RULES.get(etype)
        if not rules:
            continue

        _apply_paragraph_format(para, rules)

    # ── 2c. 落款前插入两个空行 ──
    _insert_signature_spacing(doc, structure["paragraphs"])

    # ── 2d. 小标题自动加粗 ──
    _apply_subtitle_bold(doc)

    # ── 4. 保存 ──
    doc.save(output_path)
    return output_id


def _apply_page_setup(doc: Document):
    """设置页面：边距、A4纸张"""
    for section in doc.sections:
        section.top_margin = PAGE_MARGINS["top"]
        section.bottom_margin = PAGE_MARGINS["bottom"]
        section.left_margin = PAGE_MARGINS["left"]
        section.right_margin = PAGE_MARGINS["right"]
        section.page_width = Mm(210)
        section.page_height = Mm(297)


def _clear_spacing_attrs(sp_element):
    """彻底清除一个 w:spacing 元素中的所有段间距属性"""
    sp_element.set(qn('w:before'), '0')
    sp_element.set(qn('w:after'), '0')
    # 清除行数单位的间距（Word 中显示为"X行"）
    sp_element.set(qn('w:beforeLines'), '0')
    sp_element.set(qn('w:afterLines'), '0')
    # 清除自动间距
    sp_element.set(qn('w:beforeAutospacing'), '0')
    sp_element.set(qn('w:afterAutospacing'), '0')


def _reset_style_defaults(doc: Document):
    """
    直接修改 styles.xml 中的 docDefaults 和 Normal 样式，
    从根源上消除 Word 内置默认段间距。
    """
    styles_element = doc.styles.element
    nsmap = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    }

    # ── 1. 修改 docDefaults/pPrDefault 确保所有段落的默认段间距为 0 ──
    docDefaults = styles_element.find(qn('w:docDefaults'))
    if docDefaults is None:
        docDefaults = etree.SubElement(styles_element, qn('w:docDefaults'))
    pPrDefault = docDefaults.find(qn('w:pPrDefault'))
    if pPrDefault is None:
        pPrDefault = etree.SubElement(docDefaults, qn('w:pPrDefault'))
    pPr = pPrDefault.find(qn('w:pPr'))
    if pPr is None:
        pPr = etree.SubElement(pPrDefault, qn('w:pPr'))
    sp_default = pPr.find(qn('w:spacing'))
    if sp_default is None:
        sp_default = etree.SubElement(pPr, qn('w:spacing'))
    _clear_spacing_attrs(sp_default)
    sp_default.set(qn('w:line'), str(int(LINE_SPACING.pt * 20)))
    sp_default.set(qn('w:lineRule'), 'exact')

    # ── 2. 修改 Normal 样式（默认段落样式）的段间距为 0 ──
    for style in styles_element.findall(qn('w:style')):
        is_default = style.get(qn('w:default'))
        style_type = style.get(qn('w:type'))
        if is_default == '1' and style_type == 'paragraph':
            style_pPr = style.find(qn('w:pPr'))
            if style_pPr is None:
                style_pPr = etree.SubElement(style, qn('w:pPr'))
            style_sp = style_pPr.find(qn('w:spacing'))
            if style_sp is None:
                style_sp = etree.SubElement(style_pPr, qn('w:spacing'))
            _clear_spacing_attrs(style_sp)
            break


def _reset_all_paragraph_spacing(doc: Document):
    """全文所有段落段前段后间距归零（XML 级别，确保写入）"""
    for para in doc.paragraphs:
        pPr = para._element.get_or_add_pPr()
        spacing = pPr.find(qn('w:spacing'))
        if spacing is None:
            spacing = etree.SubElement(pPr, qn('w:spacing'))
        _clear_spacing_attrs(spacing)
        # 同时通过 python-docx API 设置（双保险）
        pf = para.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)


def _para_has_drawing(para) -> bool:
    """检查段落中是否包含图片（drawing 元素）"""
    for run in para.runs:
        drawings = run._element.findall(qn('w:drawing'))
        if drawings:
            return True
    return False


def _clear_paragraph(para):
    """清空段落格式（空段落占位）"""
    pPr = para._element.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = etree.SubElement(pPr, qn('w:spacing'))
    _clear_spacing_attrs(spacing)
    # 清空内容
    for run in para.runs:
        run.text = ""


def _apply_paragraph_format(para, rules: dict):
    """对单个段落应用格式规则"""
    pf = para.paragraph_format

    # 对齐方式
    pf.alignment = rules.get("alignment", WD_ALIGN_PARAGRAPH.JUSTIFY)

    # 首行缩进 — 先清除旧缩进属性（包括 firstLineChars 避免冲突）
    _clear_indent(para)
    indent = rules.get("first_line_indent")
    if indent is not None:
        pf.first_line_indent = indent

    # 段前段后
    pf.space_before = PARA_SPACING_BEFORE
    pf.space_after = PARA_SPACING_AFTER

    # 固定行距 28 磅（需 XML 操作）
    _set_fixed_line_spacing(para, LINE_SPACING)

    # 字符级格式：CJK/ASCII 分割 + 字体设置
    _apply_run_format(para, rules)


def _clear_indent(para):
    """清除段落所有缩进属性，避免 firstLineChars 等残留属性与新值冲突"""
    pPr = para._element.get_or_add_pPr()
    ind_el = pPr.find(qn('w:ind'))
    if ind_el is not None:
        pPr.remove(ind_el)


def _set_fixed_line_spacing(para, spacing_pt):
    """通过 XML 设置固定行距（python-docx 高层 API 不完善）"""
    pPr = para._element.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = etree.SubElement(pPr, qn('w:spacing'))

    # 行距值单位：exact/atLeast 模式下为 1/20 磅 (ST_TwipsMeasure)
    line_value = int(spacing_pt.pt * 20)
    spacing.set(qn('w:line'), str(line_value))
    spacing.set(qn('w:lineRule'), 'exact')
    _clear_spacing_attrs(spacing)


def _apply_run_format(para, rules: dict):
    """
    字符级格式设置：
    1. 检查是否包含图片（drawing），如有则仅更新字体属性不重建 Run
    2. 无图片时：合并文本 → CJK/ASCII 分割 → 重建 Run
    """
    cn_font = rules["cn_font"]
    en_font = rules["en_font"]
    font_size = rules["size"]
    bold = rules["bold"]

    # 检查段落是否包含图片（drawing 元素），如有则保留 Run 结构仅更新字体
    has_drawing = False
    for run in para.runs:
        drawings = run._element.findall(qn('w:drawing'))
        if drawings:
            has_drawing = True
            break

    if has_drawing:
        # 仅更新现有 Run 的字体属性，不删除重建
        for run in para.runs:
            run.font.size = font_size
            run.font.bold = bold
            run.font.name = cn_font
            _set_east_asian_font(run, cn_font)
            _set_ascii_font(run, en_font)
        return

    # 无图片：合并文本，清空重建
    full_text = "".join(run.text for run in para.runs)

    for run in para.runs:
        run._element.getparent().remove(run._element)

    if not full_text.strip():
        return

    segments = utils.split_cjk_ascii(full_text)

    for seg_text, is_cjk in segments:
        if not seg_text:
            continue
        run = para.add_run(seg_text)
        run.font.size = font_size
        run.font.bold = bold

        if is_cjk:
            run.font.name = cn_font
            _set_east_asian_font(run, cn_font)
        else:
            run.font.name = en_font
            _set_east_asian_font(run, cn_font)

        _set_ascii_font(run, en_font)


def _set_east_asian_font(run, font_name: str):
    """设置东亚字体（w:eastAsia）"""
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = etree.SubElement(rPr, qn('w:rFonts'))
    rFonts.set(qn('w:eastAsia'), font_name)


def _set_ascii_font(run, font_name: str):
    """设置西文字体（w:ascii + w:hAnsi）"""
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = etree.SubElement(rPr, qn('w:rFonts'))
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)


def _apply_subtitle_bold(doc: Document):
    """自动检测小标题并加粗：
    1. 连续3段以上正文，首句（到第一个。）字数相同 → 每段首句加粗
    2. 段内出现3次以上"一是/二是/XX是"且各段到。字数相同 → 加粗这些分句（到。）
    """
    import re
    body_paras = [p for p in doc.paragraphs if p.text.strip()]

    # ── 模式1：检测连续段落首句长度相同的分组 ──
    first_sentences = []
    for p in body_paras:
        text = p.text.strip()
        dot_pos = text.find('。')
        if dot_pos == -1:
            first_sentences.append(None)
        else:
            first_sentences.append((p, dot_pos + 1, len(text[:dot_pos])))

    i = 0
    while i < len(first_sentences):
        if first_sentences[i] is None:
            i += 1
            continue
        _, _, length = first_sentences[i]
        j = i + 1
        while j < len(first_sentences):
            if first_sentences[j] is None or first_sentences[j][2] != length:
                break
            j += 1
        if j - i >= 3:
            for k in range(i, j):
                para, end_pos, _ = first_sentences[k]
                _bold_char_range(para, 0, end_pos)
            i = j
        else:
            i += 1

    # ── 模式2：段内"一是/二是/三是..."等分句加粗（到。）──
    branch_re = re.compile(r'[一二三四五六七八九十]+[、，,是]|\d+[、，,\.是]')
    for p in body_paras:
        text = p.text.strip()
        # 找出所有"XX是"或"XX、"标记的位置
        markers = []
        for m in branch_re.finditer(text):
            # 确保在句首或分句开头位置
            pos = m.start()
            if pos == 0 or text[pos-1] in '。；;，,、\n':
                markers.append((m.start(), m.end()))
        if len(markers) < 3:
            continue
        # 每段：从标记开头到下一个。或段落末尾
        segments = []
        for idx, (seg_start, _) in enumerate(markers):
            # 找这个分句的结束位置（下一个。或段落末尾）
            dot = text.find('。', seg_start)
            if dot == -1:
                seg_end = len(text)
            else:
                seg_end = dot + 1  # 包含。
            segments.append((seg_start, seg_end))
        if len(segments) < 3:
            continue
        # 检查各段字符数是否相同（±1）
        seg_lens = [end - start for start, end in segments]
        if max(seg_lens) - min(seg_lens) <= 1:
            for seg_start, seg_end in segments:
                _bold_char_range(p, seg_start, seg_end)


def _bold_char_range(para, start_char: int, end_char: int):
    """在段落中按字符位置精确加粗 [start_char, end_char) 范围。
    会拆分跨边界的 run，确保只加粗目标范围内的字符。"""
    if start_char >= end_char:
        return
    # 收集每个 run 的文本和字符范围
    runs_info = []
    pos = 0
    for run in para.runs:
        t = run.text if run.text else ''
        runs_info.append((run, t, pos, pos + len(t)))
        pos += len(t)

    # 找到 start_char 和 end_char 所在的 run 索引
    start_run_idx = end_run_idx = -1
    for idx, (_, _, s, e) in enumerate(runs_info):
        if start_run_idx == -1 and start_char < e:
            start_run_idx = idx
        if end_run_idx == -1 and end_char <= e:
            end_run_idx = idx
        if start_run_idx >= 0 and end_run_idx >= 0:
            break

    if start_run_idx < 0 or end_run_idx < 0:
        return

    # 从后往前处理，避免索引变化
    for idx in range(end_run_idx, start_run_idx - 1, -1):
        run, text, s, e = runs_info[idx]
        if idx == start_run_idx and idx == end_run_idx:
            # 起止在同一个 run 内：拆成 [s, start) [start, end) [end, e)，加粗中间
            if end_char < e:
                _split_run_at(para, run, end_char - s)
            if start_char > s:
                after_start = _split_run_at(para, run, start_char - s)
                _set_bold(after_start)
            else:
                _set_bold(run)
        elif idx == start_run_idx:
            if start_char > s:
                new_run = _split_run_at(para, run, start_char - s)
                _set_bold(new_run)
            else:
                _set_bold(run)
        elif idx == end_run_idx:
            if end_char < e:
                _split_run_at(para, run, end_char - s)
            _set_bold(run)
        else:
            _set_bold(run)


def _split_run_at(para, run, offset: int):
    """在 run 内 offset 字符处拆分，返回后半部分的新 Run。
    使用 para.add_run() 确保创建的是正确的 python-docx Run 对象。"""
    text = run.text if run.text else ''
    if offset <= 0 or offset >= len(text):
        return run
    left_text = text[:offset]
    right_text = text[offset:]

    # 保存原 run 的格式
    saved_name = run.font.name
    saved_size = run.font.size
    saved_bold = run.font.bold
    from docx.oxml.ns import qn as _qn
    rPr_orig = run._element.find(_qn('w:rPr'))

    # 用 para.add_run 创建新 run（python-docx 正确包装）
    new_run = para.add_run(right_text)
    if rPr_orig is not None:
        # 复制原格式属性到新 run
        new_rPr = new_run._element.find(_qn('w:rPr'))
        if new_rPr is not None:
            new_run._element.remove(new_rPr)
        new_run._element.insert(0, etree.fromstring(etree.tostring(rPr_orig)))
    new_run.font.bold = saved_bold

    # 修改原 run 文本为前半部分
    run.text = left_text
    _set_bold(run) if saved_bold else None

    # 将新 run 移动到原 run 之后
    run._element.addnext(new_run._element)

    return new_run


def _set_bold(run):
    """设置 run 为加粗"""
    run.font.bold = True


def _insert_signature_spacing(doc: Document, paragraphs_info: list):
    """在落款前插入两个空行（若前面已有空行则补足到两个）"""
    # 找到第一个签名或日期段落
    sig_idx = None
    for info in paragraphs_info:
        if info["element_type"] in ("signature", "signature_date"):
            sig_idx = info["index"]
            break
    if sig_idx is None:
        return

    # 统计签名段前已有的连续空段落数（从签名段向前数）
    body = doc.element.body
    body_children = list(body)
    # 找到签名段在 body 子元素中的位置
    sig_para = doc.paragraphs[sig_idx]
    sig_elem = sig_para._element
    sig_pos = None
    for i, child in enumerate(body_children):
        if child is sig_elem:
            sig_pos = i
            break
    if sig_pos is None:
        return

    # 向前统计已有空段落
    existing_empty = 0
    for j in range(sig_pos - 1, -1, -1):
        child = body_children[j]
        if child.tag == qn('w:p'):
            # 检查是否为空段落
            pPr = child.find(qn('w:pPr'))
            texts = child.findall('.//' + qn('w:t'))
            has_text = any(t.text and t.text.strip() for t in texts)
            if not has_text:
                existing_empty += 1
                continue
        break

    # 补足到两个空行
    need = max(0, 2 - existing_empty)
    for _ in range(need):
        empty_p = etree.SubElement(body, qn('w:p'))
        # 设置空行段间距为 0 + 固定行距 28 磅
        pPr = etree.SubElement(empty_p, qn('w:pPr'))
        sp = etree.SubElement(pPr, qn('w:spacing'))
        _clear_spacing_attrs(sp)
        sp.set(qn('w:line'), str(int(LINE_SPACING.pt * 20)))
        sp.set(qn('w:lineRule'), 'exact')
        # 在签名段之前插入
        body.remove(empty_p)
        sig_elem.addprevious(empty_p)


