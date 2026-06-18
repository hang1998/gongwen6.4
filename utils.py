"""工具函数：正则模式 + CJK/ASCII 字符分割"""

import re

# ── 结构检测正则 ──

# 一级标题：一、二、三、… 十、
HEADING_1 = re.compile(
    r'^\s*[一二三四五六七八九十]+[、．]\s*'
)

# 二级标题：（一）（二）…
HEADING_2 = re.compile(
    r'^\s*（[一二三四五六七八九十]+）\s*'
)

# 三级标题：1. 2. 3. …
HEADING_3 = re.compile(
    r'^\s*\d+[\.\、．]\s*'
)

# 附件
ATTACHMENT = re.compile(
    r'^\s*附件'
)

# 附件编号项 —— 上下文相关（与 HEADING_3 模式相同，但供 parser 和 formatter 语义明确引用）
ATTACHMENT_ITEM = re.compile(
    r'^\s*\d+[\.\、．]\s*'
)

# 附注
NOTE = re.compile(
    r'^\s*(附注|注)[：:]'
)

# 抄送
CC = re.compile(
    r'^\s*抄送[：:]'
)

# 日期：2024年1月1日 或 二〇二四年一月一日
DATE = re.compile(
    r'^\s*\d{4}年\d{1,2}月\d{1,2}日\s*$'
)

# 发文字号：XX发〔2024〕1号
DOC_NUMBER = re.compile(
    r'〔\d+〕'
)

# 演讲稿发言标签：仅匹配文章开头的称呼短行（如"张三："、"各位领导、同志们："）
# 排除正文中以冒号结尾的说明性语句（如"具体如下："、"会议指出："）
SPEAKER = re.compile(
    r'^[^一二三四五六七八九十（\d].*[：:]\s*$'
)

# 常见的正文冒号引导词，包含这些词的段落不应被当作发言标签
_BODY_COLON_WORDS = re.compile(
    r'(如下|指出|强调|认为|表示|通知|如下|方面|措施|要求|意见|建议|'
    r'安排|部署|决定|现将|特此|为此|据此|因此|所以|但是|然而|而且|'
    r'一是|二是|三是|首先|其次|最后|另外|此外|同时|总之|综上)'
)

# 图表标题：表X、图X、表X-X、图X-X 等
FIGURE_TABLE_CAPTION = re.compile(
    r'^\s*(表|图|Table|Figure|Fig)\s*[\d\-—–]+\s*[\.\s、．]?'
)

# ── 制度类检测正则 ──

# 第XX章（制度类一级标题）
REGULATION_CHAPTER = re.compile(
    r'^\s*第[一二三四五六七八九十百千\d]+章\s*'
)

# 第XX节（制度类二级标题）
REGULATION_SECTION = re.compile(
    r'^\s*第[一二三四五六七八九十百千\d]+节\s*'
)

# 第XX条（制度类条文）
REGULATION_ARTICLE = re.compile(
    r'^\s*第[一二三四五六七八九十百千\d]+条\s*'
)

# ── CJK/ASCII 字符分割 ──

CJK_PATTERN = re.compile(r'[一-鿿㐀-䶿豈-﫿　-〿＀-￯]')


def is_cjk(ch: str) -> bool:
    """判断单个字符是否为 CJK 字符（含中文及全角标点）"""
    return bool(CJK_PATTERN.match(ch))


def split_cjk_ascii(text: str):
    """
    按字符类型边界分割文本，返回 [(segment_text, is_cjk), ...]

    示例:
      "2024年GDP增长6.5%" →
      [("2024", False), ("年", True), ("GDP", False),
       ("增长", True), ("6.5", False), ("%", False)]
    """
    segments = []
    if not text:
        return segments

    current_chars = []
    current_is_cjk = None

    for ch in text:
        ch_is_cjk = is_cjk(ch)

        if current_is_cjk is None:
            current_is_cjk = ch_is_cjk
            current_chars.append(ch)
        elif ch_is_cjk == current_is_cjk:
            current_chars.append(ch)
        else:
            segments.append(("".join(current_chars), current_is_cjk))
            current_chars = [ch]
            current_is_cjk = ch_is_cjk

    if current_chars:
        segments.append(("".join(current_chars), current_is_cjk))

    return segments


def classify_paragraph(text: str, index: int, total: int, is_first_non_empty: bool) -> str:
    """
    对单个段落做初步分类。
    返回元素类型字符串。
    """
    if not text.strip():
        return "empty"

    # 附件开头优先识别，避免被误判为公文题目
    if ATTACHMENT.match(text):
        return "attachment_title"

    # 各级标题优先于大标题识别（首行也可能是标题序号开头）
    if HEADING_1.match(text):
        return "heading_1"
    if HEADING_2.match(text):
        return "heading_2"
    if HEADING_3.match(text):
        return "heading_3"

    # 制度类结构识别
    if REGULATION_CHAPTER.match(text):
        return "regulation_chapter"
    if REGULATION_SECTION.match(text):
        return "regulation_section"
    if REGULATION_ARTICLE.match(text):
        return "regulation_article"

    # 公文题目：第一个非空段落，需满足标题特征才识别
    if is_first_non_empty:
        if len(text) <= 50 and '。' not in text and '：' not in text and ':' not in text:
            return "title"
        # 不满足标题条件，继续后续判断（可能是无标题文档）

    # 图表标题：表X、图X 等
    if FIGURE_TABLE_CAPTION.match(text):
        return "figure_caption"

    # 演讲稿发言标签：以冒号结尾的称呼短行（≤15字），不含正文引导词
    if SPEAKER.match(text) and len(text) <= 15 and not _BODY_COLON_WORDS.search(text):
        return "speaker"

    return "body"  # 暂定，第二段精修
