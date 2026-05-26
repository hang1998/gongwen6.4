"""公文格式规则常量 — 所有格式参数集中定义"""

from docx.shared import Pt, Mm, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ── 页面设置 ──
PAGE_MARGINS = {
    "top": Mm(37),
    "bottom": Mm(35),
    "left": Mm(28),
    "right": Mm(26),
}
PAGE_SIZE = "A4"    # 默认 A4

# ── 段落格式 ──
LINE_SPACING = Pt(28)           # 固定行距 28 磅
PARA_SPACING_BEFORE = Pt(0)     # 段前 0 行
PARA_SPACING_AFTER = Pt(0)      # 段后 0 行

# ── 字体规则 ──
# key → { cn_font, en_font, size, bold, alignment, first_line_indent }
FONT_RULES = {
    "title": {
        "cn_font": "方正小标宋_GBK",
        "en_font": "Times New Roman",
        "size": Pt(22),          # 2号
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.CENTER,
        "first_line_indent": None,
    },
    "figure_caption": {
        "cn_font": "黑体",
        "en_font": "Times New Roman",
        "size": Pt(16),          # 3号
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.CENTER,
        "first_line_indent": None,    # 图表标题居中不缩进
    },
    "speaker": {
        "cn_font": "仿宋_GB2312",
        "en_font": "Times New Roman",
        "size": Pt(16),          # 3号
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "first_line_indent": None,    # 发言标签不缩进
    },
    "heading_1": {
        "cn_font": "黑体",
        "en_font": "Times New Roman",
        "size": Pt(16),          # 3号
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "first_line_indent": Cm(1.13),
    },
    "heading_2": {
        "cn_font": "楷体_GB2312",
        "en_font": "Times New Roman",
        "size": Pt(16),
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "first_line_indent": Cm(1.13),
    },
    "heading_3": {
        "cn_font": "仿宋_GB2312",
        "en_font": "Times New Roman",
        "size": Pt(16),
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "first_line_indent": Cm(1.13),
    },
    "body": {
        "cn_font": "仿宋_GB2312",
        "en_font": "Times New Roman",
        "size": Pt(16),
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "first_line_indent": Cm(1.13),  # 首行缩进约2字符
    },
    "attachment": {
        "cn_font": "仿宋_GB2312",
        "en_font": "Times New Roman",
        "size": Pt(16),
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "first_line_indent": Cm(1.13),
    },
    "attachment_title": {
        "cn_font": "黑体",
        "en_font": "Times New Roman",
        "size": Pt(16),
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.LEFT,
        "first_line_indent": None,
    },
    "note": {
        "cn_font": "楷体_GB2312",
        "en_font": "Times New Roman",
        "size": Pt(16),
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "first_line_indent": None,
    },
    "signature": {
        "cn_font": "仿宋_GB2312",
        "en_font": "Times New Roman",
        "size": Pt(16),
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.RIGHT,
        "first_line_indent": None,
    },
    "signature_date": {
        "cn_font": "仿宋_GB2312",
        "en_font": "Times New Roman",
        "size": Pt(16),
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.RIGHT,
        "first_line_indent": None,
    },
    "cc": {
        "cn_font": "仿宋_GB2312",
        "en_font": "Times New Roman",
        "size": Pt(16),
        "bold": False,
        "alignment": WD_ALIGN_PARAGRAPH.LEFT,
        "first_line_indent": None,
    },
}

# ── 页码 ──
PAGE_NUMBER_FONT = "宋体"
PAGE_NUMBER_SIZE = Pt(14)               # 4号
PAGE_NUMBER_ALIGNMENT = "CENTER"

# ── 字体文件映射（供下载） ──
FONT_FILES = {
    "FZXBS_GBK.ttf": "方正小标宋_GBK（公文标题）",
    "FS_GB2312.ttf": "仿宋_GB2312（正文/三级标题/落款）",
    "KT_GB2312.ttf": "楷体_GB2312（二级标题/附注）",
}
