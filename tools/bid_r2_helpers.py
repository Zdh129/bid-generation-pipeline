#!/usr/bin/env python3
"""封面、目录域、横向节、限价与偏离证据表等生成辅助。不读 V8 单价。"""
from __future__ import annotations

import csv
import json
import zipfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor
from docx.text.paragraph import Paragraph

ROOT = Path(__file__).resolve().parents[1]
CEILING_CSV = ROOT / "work" / "ceiling_prices.csv"
JUDGE_JSON = ROOT / "work" / "deviation_judgements.json"

# 官方招标文件（文件夹 1）非个人字段
PROJECT_NAME = "示例采购项目"
PROJECT_CODE = "EXAM-ZB-20260001"
PACKAGE_NO = "2"
# 文件夹 1：投标截止/开标 2026-11-03
BID_DATE = "2026年11月3日"  # r2 遗留常量；r3 起封面/投标函均改用 bid_r3_fixes.resolve_bid_date()
# 文件夹 3 投标主体文件名（非 V8）：示例投标人有限公司_QMS/EMS/OHSMS 主证
BIDDER_NAME = "示例投标人有限公司"
PACKAGE_CEILING = 100000.0  # 官方限价清单合计


def load_ceilings() -> dict[int, dict]:
    out: dict[int, dict] = {}
    with CEILING_CSV.open(encoding="utf-8") as f:
        rows = list(csv.reader(f))
    for row in rows[2:]:
        if not row or not str(row[0]).strip().isdigit():
            continue
        seq = int(str(row[0]).strip())
        if seq > 45:
            continue
        unit_p = row[5].replace(",", "").strip() if len(row) > 5 else ""
        tot_p = row[6].replace(",", "").strip() if len(row) > 6 else ""
        qty = row[4].replace(",", "").strip() if len(row) > 4 else "1"
        unit = (row[3] or "台").strip()
        try:
            up, tp, qn_ = float(unit_p or 0), float(tot_p or 0), float(qty or 1)
        except ValueError:
            up, tp, qn_ = 0.0, 0.0, 1.0
        out[seq] = {
            "name": (row[1] or "").strip(),
            "unit": unit or "台",
            "qty": int(qn_) if qn_ == int(qn_) else qn_,
            "ceiling_unit": up,
            "ceiling_total": tp,
        }
    return out


def load_judgements() -> dict[str, dict]:
    data = json.loads(JUDGE_JSON.read_text(encoding="utf-8"))
    return {j["id"]: j for j in data["judgements"]}


def fmt_money(v: float) -> str:
    if abs(v - round(v)) < 1e-9:
        return f"{int(round(v)):,}"
    return f"{v:,.2f}"


def set_para_text(p, text: str, size=14, bold=False, ea="宋体"):
    from build_bid import set_font, add_run
    xml = p._p
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    add_run(p, text, size=size, bold=bold)


def fill_cover(doc: Document) -> None:
    """填写封面项目/公司等非个人敏感字段，来源为文件夹 1 与 3 的公开名称。"""
    from build_bid import set_font, add_run
    for p in doc.paragraphs:
        raw = p.text or ""
        t = raw.replace("\u3000", " ").replace("\u200c", "").replace("\t", "")
        t_s = " ".join(t.split())
        if "包号" in t_s and "项目" not in t_s:
            set_para_text(p, f"（包号：{PACKAGE_NO}）", size=14, bold=False)
        elif t_s.startswith("项目名称") or "项目名称" in t_s[:8]:
            set_para_text(p, f"项目名称：{PROJECT_NAME}", size=14, bold=False)
        elif t_s.startswith("项目编号") or "项目编号" in t_s[:8]:
            set_para_text(p, f"项目编号：{PROJECT_CODE}", size=14, bold=False)
        elif "投标人" in t_s[:6] and "日期" not in t_s:
            set_para_text(p, f"投标人：{BIDDER_NAME}（加盖公章）", size=14, bold=False)
        elif "投标日期" in t_s or t_s.startswith("投标 日期") or "日期" in t_s[:8] and "投标" in raw:
            from bid_r3_fixes import resolve_bid_date
            set_para_text(p, f"投标日期：{resolve_bid_date()['cn']}", size=14, bold=False)  # r3：唯一日期源


def add_toc_field(doc: Document, page_break: bool = True) -> None:
    """插入 Word TOC 域（1–3 级标题）。LibreOffice 转 PDF 时应更新域。目录标题本身不用 Heading，以免写入目录。"""
    from build_bid import add_run, set_font
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.first_line_indent = Pt(0)
    title.paragraph_format.space_after = Pt(12)
    add_run(title, "目录", size=14, bold=True)

    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run()
    set_font(r, size=10)
    rPr_parent = r._r

    def fld(typ: str):
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), typ)
        if typ == "begin":
            el.set(qn("w:dirty"), "true")
        return el

    r1 = p.add_run(); set_font(r1, size=10)
    r1._r.append(fld("begin"))
    r2 = p.add_run(); set_font(r2, size=10)
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    r2._r.append(instr)
    r3 = p.add_run(); set_font(r3, size=10)
    r3._r.append(fld("separate"))
    r4 = p.add_run("（目录域，转 PDF 时由 LibreOffice 更新页码）")
    set_font(r4, size=10)
    r5 = p.add_run(); set_font(r5, size=10)
    r5._r.append(fld("end"))

    if not page_break:
        return  # r3：由调用方插入分节符
    pb = doc.add_paragraph()
    pb.paragraph_format.first_line_indent = Pt(0)
    br = pb.add_run()
    from docx.enum.text import WD_BREAK
    br.add_break(WD_BREAK.PAGE)


def configure_section(sec, *, landscape: bool) -> None:
    import layout_spec as spec
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width = Mm(297)
        sec.page_height = Mm(210)
        sec.top_margin = Cm(1.8)
        sec.bottom_margin = Cm(2.0)
        sec.left_margin = Cm(2.0)
        sec.right_margin = Cm(2.0)
    else:
        sec.orientation = WD_ORIENT.PORTRAIT
        sec.page_width = Mm(210)
        sec.page_height = Mm(297)
        sec.top_margin = Cm(spec.MARGIN_TOP_CM)
        sec.bottom_margin = Cm(spec.MARGIN_BOTTOM_CM)
        sec.left_margin = Cm(spec.MARGIN_LEFT_CM)
        sec.right_margin = Cm(spec.MARGIN_RIGHT_CM)


def add_oriented_section(doc: Document, *, landscape: bool):
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    configure_section(sec, landscape=landscape)
    return sec


def mark_pending(run, text: str = "【待填】") -> None:
    run.text = text
    run.bold = True
    run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    run.font.size = Pt(14)


def fill_pending_cells(table, col_indices: list[int], marker: str = "【待填】") -> None:
    from build_bid import set_font
    for i, row in enumerate(table.rows):
        if i == 0:
            continue
        for c in col_indices:
            cell = row.cells[c]
            # clear and mark first para
            p = cell.paragraphs[0]
            for r in list(p.runs):
                r._element.getparent().remove(r._element)
            run = p.add_run()
            mark_pending(run, marker)


def cite_evidence(j: dict | None) -> str:
    if not j:
        return "—"
    evs = j.get("evidence") or []
    if not evs:
        return (j.get("reason") or "")[:80]
    e = evs[0]
    fn = Path(e.get("file") or "").name
    page = e.get("page") or ""
    snip = " ".join((e.get("snippet") or "").split())[:60]
    return f"{fn} 第{page}页：「{snip}」" if fn else (j.get("reason") or "")[:80]


def patch_settings_update_fields(docx_path: Path) -> None:
    """在 word/settings.xml 打开时更新域，便于 Word/LibreOffice 刷新 TOC。"""
    tmp = docx_path.with_suffix(".tmpdocx")
    with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(tmp, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/settings.xml":
                xml = data.decode("utf-8")
                if "w:updateFields" not in xml:
                    if "</w:settings>" in xml:
                        xml = xml.replace(
                            "</w:settings>",
                            '<w:updateFields w:val="true"/></w:settings>',
                            1,
                        )
                    data = xml.encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(docx_path)


def collect_docx_headings(doc) -> list[dict]:
    out = []
    for p in doc.paragraphs:
        st = ((p.style.name if p.style else "") or "")
        if st in ("Heading 1", "Heading 2", "Heading 3") or st.startswith("标题"):
            text = (p.text or "").strip()
            if not text or text == "目录":
                continue
            level = 3
            if "1" in st or st.endswith("1") or st == "Heading 1" or "标题 1" in st:
                level = 1
            if "2" in st or st.endswith("2") or st == "Heading 2" or "标题 2" in st:
                level = 2
            if "3" in st or st.endswith("3") or st == "Heading 3" or "标题 3" in st:
                level = 3
            out.append({"text": text, "level": level})
    return out


def write_populated_toc(doc, entries: list[dict]) -> None:
    """把目录域的显示结果替换为带页码的条目（保留域结构供 Word 再更新）。"""
    from build_bid import set_font
    from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER
    # 找到含 TOC 域或占位说明的段，在其后插入条目；先删旧条目
    toc_para = None
    for p in doc.paragraphs:
        xml = p._p.xml
        if " TOC " in xml or "目录域" in (p.text or ""):
            toc_para = p
            break
    if toc_para is None:
        return
    # 删除占位文字 run（保留域 begin/instr/separate/end），在 separate 与 end 之间写入条目较复杂；
    # 改为：域后追加可见目录行（域自身若未展开也不影响 PDF 阅读）。
    # 清掉「目录域，转 PDF…」可见文字
    for r in toc_para.runs:
        if "目录域" in (r.text or "") or "更新页码" in (r.text or ""):
            r.text = ""
    # 删除旧的已写入目录行（普通段落，含右对齐点前导符且不是标题）
    body = doc.element.body
    # insert after toc_para
    anchor = toc_para._p
    # remove following toc-entry paras until page break
    nxt = anchor.getnext()
    while nxt is not None:
        if nxt.tag == qn("w:p"):
            pp = Paragraph(nxt, doc)
            raw = pp._p.xml
            if "w:type=\"page\"" in raw or 'w:type="page"' in raw:
                break
            if (pp.style and str(pp.style.name).startswith("Heading")):
                break
            if (pp.text or "").strip() == "目录":
                nxt = nxt.getnext(); continue
            # likely old toc line
            if "	" in (pp.text or "") or (pp.paragraph_format.tab_stops and len(list(pp.paragraph_format.tab_stops))):
                follow = nxt.getnext()
                nxt.getparent().remove(nxt)
                nxt = follow
                continue
        break
    # insert new lines after toc field para
    from docx.oxml import OxmlElement as _Oe
    content_w = 15.3  # r3：右对齐页码制表位放在版心内（纵向版心 15.48cm）
    for e in entries:
        p = Paragraph(_Oe("w:p"), toc_para._parent)
        # we will insert xml after anchor
        newp = _Oe("w:p")
        # build via a temp approach: add at end then move — easier use document add then relocate
    # simpler: add paragraphs at end of body then move xml nodes after toc
    created = []
    for e in entries:
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.left_indent = Pt(12 * (int(e.get("level") or 1) - 1))
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p.paragraph_format.line_spacing = Pt(20)
        try:
            p.paragraph_format.tab_stops.add_tab_stop(Cm(content_w), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        except Exception:
            pass
        page = e.get("page")
        r1 = p.add_run(e["text"])
        set_font(r1, size=10.5)
        r2 = p.add_run("	" + (str(page) if page else ""))
        set_font(r2, size=10.5)
        created.append(p)
    # move before the page-break after toc
    # find page break following toc
    pb = toc_para._p.getnext()
    for p in created:
        el = p._p
        el.getparent().remove(el)
        pb.addprevious(el)
