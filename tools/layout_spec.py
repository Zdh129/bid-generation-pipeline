"""标书排版默认值与白名单。

阈值优先对齐阳性对照 V8 参考成品（见 standards/v8-stats.md），
并保留工程标书常见规范作为兼容档。检查器对 V8 本身不应误报阻断项。
"""
from __future__ import annotations

# 纸张：V8 以 A4 为主，含横向表节
PAGE_WIDTH_MM = 210.0
PAGE_HEIGHT_MM = 297.0
PAGE_SIZE_TOL_MM = 1.0
ALLOW_LANDSCAPE_A4 = True  # V8 有横向节

# 页边距：V8 共 8 节、多种边距。生成默认取 V8 首个纵向节。
# 检查允许范围覆盖 V8 全部节（最小约 1.69cm，最大约 2.79cm）。
MARGIN_TOP_CM = 2.52
MARGIN_BOTTOM_CM = 2.58
MARGIN_LEFT_CM = 2.79
MARGIN_RIGHT_CM = 2.73
MARGIN_TOL_CM = 0.20
MARGIN_MIN_CM = 1.50  # 低于此视为过窄（V8 最窄右边距 1.688，放宽到 1.50）
MARGIN_MAX_CM = 3.30
MARGIN_LR_ALT_CM = 2.54

# 字体：V8 正文/标题均为 PingFang SC。测试环境无苹方时用 Noto Sans CJK SC 等效。
FONT_BODY_EA = "Noto Sans CJK SC"
FONT_HEADING_EA = "Noto Sans CJK SC"
FONT_ASCII = "Times New Roman"  # V8 页脚出现 Times New Roman；环境无则 Liberation Serif
FONT_ASCII_FALLBACK = "Liberation Serif"

FONT_PINGFANG_ALIASES = {
    "PingFang SC", "PingFangSC", "苹方", "苹方-简", "PingFangSC-Regular", "PingFangSC-Semibold",
}
FONT_BODY_EA_ALIASES = FONT_PINGFANG_ALIASES | {
    "宋体", "SimSun", "NSimSun", "宋体-简",
    "Noto Serif CJK SC", "NotoSerifCJKsc-Regular", "NotoSerifCJKSC", "NotoSerifCJKsc",
    "Noto Sans CJK SC", "NotoSansCJKsc-Regular", "NotoSansCJKsc-Bold", "NotoSansCJKSC", "NotoSansCJKsc",
}
FONT_FANGSONG_ALIASES = {
    "仿宋", "仿宋_GB2312", "FangSong", "FangSong_GB2312", "仿宋-简", "仿宋;仿宋-简", "STFangsong",
}
FONT_HEADING_EA_ALIASES = FONT_BODY_EA_ALIASES | {
    "黑体", "SimHei", "FZHei-B01", "FZHTJW--GB1-0",
}
FONT_ASCII_ALIASES = {
    "Times New Roman", "Liberation Serif", "Times", "Nimbus Roman",
    "PingFang SC", "Noto Sans CJK SC", "Noto Serif CJK SC", "SimSun", "Calibri", "Arial",
}
FONT_BODY_ALLOWED = FONT_BODY_EA_ALIASES | FONT_FANGSONG_ALIASES
FONT_WHITELIST = FONT_BODY_ALLOWED | FONT_HEADING_EA_ALIASES | FONT_ASCII_ALIASES | {
    "楷体", "KaiTi", "STKaiti", "FZKai-Z03", "FZKTJW--GB1-0",
    "微软雅黑", "Microsoft YaHei", "MicrosoftYaHei",
    "Courier New", "CourierNew", "Liberation Sans",
}

# 字号：V8 正文、标题 1–3、表内、题注均为 14pt（四号）；文题约 47.5pt；目录 10.5pt；页脚 9pt
BODY_SIZE_PT = 14.0
BODY_SIZE_ALT_PT = 12.0
SIZE_TOL_PT = 0.5
TITLE_SIZE_PT = 48.0  # V8 文题约 47.5
TITLE_SIZE_TOL_PT = 6.0
H1_SIZE_PT = 14.0  # V8：与正文同号加粗居中，不是三号 16pt
H2_SIZE_PT = 14.0
H3_SIZE_PT = 14.0
CAPTION_SIZE_PT = 14.0  # V8 Bid Caption 为 14pt
TABLE_SIZE_PT = 14.0
TABLE_SIZE_MIN_PT = 9.0
TABLE_SIZE_MAX_PT = 14.5
HEADER_FOOTER_SIZE_PT = 9.0
HEADER_FOOTER_SIZE_TOL_PT = 1.0
TOC_SIZE_PT = 10.5

# 行距：V8 主峰 exact 20pt，其次 multiple 1.0；样式默认 line=288 auto ≈ 1.2 倍
LINE_SPACING_PT = 20.0
LINE_SPACING_ALT_PT = 28.0
LINE_SPACING_TOL_PT = 1.5
LINE_SPACING_ALLOWED_EXACT = (17.0, 18.0, 20.0, 25.0, 28.0)
LINE_SPACING_MULTIPLE_ALT = 1.5
LINE_SPACING_MULTIPLE_TOL = 0.15
LINE_SPACING_ALLOWED_MULTIPLE = (1.0, 1.2, 1.5)

# 缩进：V8 正文首行缩进均为 0。2 字缩进仅作规范兼容档，不强制。
FIRST_INDENT_CHARS = 0
FIRST_INDENT_ALT_CHARS = 2
ALLOW_ZERO_INDENT = True

def first_indent_pt(body_size_pt: float = BODY_SIZE_PT) -> float:
    return FIRST_INDENT_CHARS * body_size_pt

# 标题段间距：V8 样式 before=180twip=9pt, after=160twip=8pt
H1_SPACE_BEFORE_PT = 9.0
H1_SPACE_AFTER_PT = 8.0
H2_SPACE_BEFORE_PT = 9.0
H2_SPACE_AFTER_PT = 8.0
H3_SPACE_BEFORE_PT = 9.0
H3_SPACE_AFTER_PT = 8.0
H1_ALIGN = "center"  # V8 heading 1 居中
SPACING_TOL_PT = 4.0

# 表格：V8 全框线 sz=2（0.25pt），4/6 表有表头重复，单元格 14pt 居中为主
TABLE_BORDER_MIN_PT = 0.25
TABLE_BORDER_COLOR = "000000"

# 图片：V8 136 张有效 DPI 中位数 242.8，最低 65.1，仅 3 张≥300，107 张宽约 16cm，全部 inline。
# 不以 300 DPI 作为阻断，避免参考成品误报。推荐 300 仅作提示。
IMAGE_MIN_DPI = 60.0          # 通过下限（覆盖 V8 最低 65.1）
IMAGE_RECOMMENDED_DPI = 300.0
IMAGE_ALIGN_TOL_MM = 3.0
IMAGE_TYPICAL_WIDTH_CM = 16.0

# 扫描纠偏（文档类）
SKEW_IMAGE_MAX_DEG = 0.5
SKEW_PAGE_MAX_DEG = 0.3
SKEW_UNRELIABLE_DEG = 7.5

BLANK_INK_RATIO = 0.003
HF_GAP_MIN_MM = 4.0
REQUIRE_HEADER = False  # V8 无页眉，仅页脚页码
PAGE_NUMBER_PATTERN = r"[—\-－]\s*\d+\s*[—\-－]|第\s*\d+\s*页|\b\d+\b"

HEADER_OPTIONAL = True


def normalize_font(name: str | None) -> str:
    if not name:
        return ""
    return name.strip().replace(" ", "").replace("-", "").replace("_", "").lower()


def font_in_set(name: str | None, allowed: set[str]) -> bool:
    if not name or name in ("(unset)", "(unknown)"):
        return False
    allowed_n = {normalize_font(x) for x in allowed}
    n = normalize_font(name)
    if n in allowed_n:
        return True
    for a in allowed_n:
        if a and (a in n or n in a):
            return True
    return False
