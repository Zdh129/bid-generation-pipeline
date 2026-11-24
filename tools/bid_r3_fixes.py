#!/usr/bin/env python3
"""r3 确定性修复：唯一日期源、图片按版心缩放、页码规则、日期全文校验。不调用任何模型/外部 API。"""
from __future__ import annotations

import json
import re
from pathlib import Path

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Emu

ROOT = Path(__file__).resolve().parents[1]
TENDER_TXT = ROOT / "work" / "tender_text" / "example-tender.txt"

# ---------------------------------------------------------------- d. 唯一日期源
_DATE_RE = re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")


def resolve_bid_date() -> dict:
    """唯一日期源：官方招标文件「投标文件递交的截止时间」中包2 所在行的日期。
    规则写死在代码中，可复核：找到「投标文件递交的截止时间」后 5 行内含「包2」的行，取其日期。"""
    lines = TENDER_TXT.read_text(encoding="utf-8").splitlines()
    for i, ln in enumerate(lines):
        if "投标文件递交的截止时间" in ln:
            for j in range(i, min(i + 6, len(lines))):
                if "包2" in lines[j]:
                    m = _DATE_RE.search(lines[j])
                    if m:
                        y, mo, d = map(int, m.groups())
                        return {
                            "iso": f"{y:04d}-{mo:02d}-{d:02d}",
                            "cn": f"{y}年{mo}月{d}日",
                            "source_file": TENDER_TXT.name,
                            "source_line": j + 1,
                            "source_text": lines[j].strip(),
                            "rule": "投标日期 = 官方招标文件第一章「投标文件递交的截止时间」包2 日期（开标同日）；封面与投标函统一使用",
                        }
    raise RuntimeError("未在招标文件中定位包2 投标截止日期")


def normalize_date(s: str) -> str | None:
    m = _DATE_RE.search(s)
    if not m:
        return None
    y, mo, d = map(int, m.groups())
    return f"{y:04d}-{mo:02d}-{d:02d}"


FIELD_LABEL_RE = re.compile(r"(投标日期|日\s*期)\s*[：:]")


def scan_dates_docx(doc) -> list[dict]:
    """扫描 DOCX 全文（正文段落 + 表格 + 页眉页脚）中的日期。返回每处日期及是否为「日期字段」。"""
    hits = []

    def scan_text(where: str, text: str):
        for m in _DATE_RE.finditer(text or ""):
            pre = (text or "")[max(0, m.start() - 12): m.start()]
            is_field = bool(FIELD_LABEL_RE.search(pre))
            hits.append({"where": where, "date": normalize_date(m.group(0)), "raw": m.group(0),
                         "is_field": is_field, "context": pre[-12:]})

    for i, p in enumerate(doc.paragraphs):
        scan_text(f"para#{i}", p.text)
    for ti, t in enumerate(doc.tables):
        for ri, row in enumerate(t.rows):
            for ci, c in enumerate(row.cells):
                scan_text(f"table#{ti}r{ri}c{ci}", c.text)
    for si, s in enumerate(doc.sections):
        for part, name in ((s.header, "header"), (s.footer, "footer")):
            for p in part.paragraphs:
                scan_text(f"sec{si}.{name}", p.text)
    return hits


def scan_dates_pdf_text(page_texts: list[str]) -> list[dict]:
    hits = []
    for i, t in enumerate(page_texts, 1):
        for m in _DATE_RE.finditer(t or ""):
            pre = (t or "")[max(0, m.start() - 12): m.start()]
            hits.append({"page": i, "date": normalize_date(m.group(0)),
                         "is_field": bool(FIELD_LABEL_RE.search(pre)), "context": pre.strip()[-10:]})
    return hits


def date_consistency_report(doc, pdf_texts: list[str] | None, src: dict) -> dict:
    dh = scan_dates_docx(doc)
    ph = scan_dates_pdf_text(pdf_texts) if pdf_texts is not None else []
    fields_docx = [h for h in dh if h["is_field"]]
    fields_pdf = [h for h in ph if h["is_field"]]
    bad_docx = [h for h in fields_docx if h["date"] != src["iso"]]
    bad_pdf = [h for h in fields_pdf if h["date"] != src["iso"]]
    # 生成器自撰正文中任何 2026-11-* 日期都必须等于唯一源（防漏网的「11月1日」）
    stray = [h for h in dh if not h["is_field"] and h["date"] and h["date"].startswith(src["iso"][:7]) and h["date"] != src["iso"]]
    return {
        "source": src,
        "docx_dates_total": len(dh),
        "docx_date_fields": len(fields_docx),
        "docx_field_values": sorted({h["date"] for h in fields_docx}),
        "pdf_date_fields": len(fields_pdf),
        "pdf_field_pages": sorted({h["page"] for h in fields_pdf}),
        "pdf_field_values": sorted({h["date"] for h in fields_pdf}),
        "non_field_dates": [{k: h[k] for k in ("where", "date", "context")} for h in dh if not h["is_field"]][:50],
        "mismatch_docx": bad_docx,
        "mismatch_pdf": bad_pdf,
        "stray_same_month": stray,
        "pass": not bad_docx and not bad_pdf and not stray and len(fields_docx) >= 2,
    }


# ---------------------------------------------------------------- b. 图片按版心缩放
def section_text_box_cm(sec) -> tuple[float, float]:
    w = (sec.page_width - sec.left_margin - sec.right_margin) / 360000.0
    h = (sec.page_height - sec.top_margin - sec.bottom_margin) / 360000.0
    return w, h


def fit_picture_width_cm(sec, img_path: Path, desired_cm: float, *, safety_cm=0.10, caption_reserve_cm=2.2) -> tuple[float, dict]:
    """返回不超过版心宽度、且图+题注不超过版心高度的显示宽度（等比，DPI 只会升不会降）。"""
    from PIL import Image
    with Image.open(img_path) as im:
        wpx, hpx = im.size
    tw, th = section_text_box_cm(sec)
    max_w = tw - safety_cm
    max_h = th - caption_reserve_cm
    w = min(desired_cm, max_w)
    if w * hpx / wpx > max_h:
        w = max_h * wpx / hpx
    w = round(w, 3)
    info = {"text_w_cm": round(tw, 3), "text_h_cm": round(th, 3), "desired_cm": desired_cm,
            "final_w_cm": w, "final_h_cm": round(w * hpx / wpx, 3),
            "eff_dpi": round(wpx / (w / 2.54), 1), "eff_dpi_at_desired": round(wpx / (desired_cm / 2.54), 1)}
    return w, info


# ---------------------------------------------------------------- c. 页码规则
def _page_field_runs(p, set_font):
    r1 = p.add_run("— "); set_font(r1, size=9)
    r = p.add_run(); set_font(r, size=9)
    f1 = OxmlElement("w:fldChar"); f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = " PAGE "
    f2 = OxmlElement("w:fldChar"); f2.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t"); t.text = "1"
    f3 = OxmlElement("w:fldChar"); f3.set(qn("w:fldCharType"), "end")
    r._r.append(f1); r._r.append(it); r._r.append(f2)
    r2 = p.add_run(); set_font(r2, size=9); r2._r.append(t)
    r3 = p.add_run(); set_font(r3, size=9); r3._r.append(f3)
    r4 = p.add_run(" —"); set_font(r4, size=9)


def _clear_part(part):
    el = part._element
    for child in list(el):
        el.remove(child)
    return part.add_paragraph()


def _set_pgnum(sec, fmt: str | None, start: int | None):
    sp = sec._sectPr
    old = sp.find(qn("w:pgNumType"))
    if old is not None:
        sp.remove(old)
    if fmt is None and start is None:
        return
    pg = OxmlElement("w:pgNumType")
    if fmt:
        pg.set(qn("w:fmt"), fmt)
    if start is not None:
        pg.set(qn("w:start"), str(start))
    # pgNumType 须位于 pgMar 之后、cols 之前；简单起见插在 pgMar 之后
    pgmar = sp.find(qn("w:pgMar"))
    if pgmar is not None:
        pgmar.addnext(pg)
    else:
        sp.append(pg)


def apply_page_numbering(doc, set_font, *, front_sections=(0, 1), body_start_section=2) -> dict:
    """统一页码规则（按标准 Q-PGN-001/003：封面可不显示，正文从 1 起编）：
    - 节0 封面：无页码（页脚清空，不链接）
    - 节1 目录：大写罗马数字「— I —」，从 I 起
    - 节2 起正文：阿拉伯数字「— N —」，节2 从 1 起，其后各节续接（不重设 start）
    所有节页脚格式一致（9pt 居中「— PAGE —」），页眉留空（与 V8 一致）。"""
    secs = list(doc.sections)
    plan = []
    for i, sec in enumerate(secs):
        sec.different_first_page_header_footer = False
        sec.footer.is_linked_to_previous = False
        sec.header.is_linked_to_previous = False
        _clear_part(sec.header)
        fp = _clear_part(sec.footer)
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.paragraph_format.first_line_indent = Pt(0)
        if i == 0:
            _set_pgnum(sec, None, None)
            plan.append({"section": i, "footer": "无页码（封面）"})
            continue
        _page_field_runs(fp, set_font)
        if i in front_sections:
            _set_pgnum(sec, "upperRoman", 1)
            plan.append({"section": i, "footer": "— I —", "fmt": "upperRoman", "start": 1})
        elif i == body_start_section:
            _set_pgnum(sec, "decimal", 1)
            plan.append({"section": i, "footer": "— N —", "fmt": "decimal", "start": 1})
        else:
            _set_pgnum(sec, "decimal", None)
            plan.append({"section": i, "footer": "— N —", "fmt": "decimal", "start": "续接"})
    return {"sections": len(secs), "plan": plan}


ROMAN = {"I": 1, "V": 5, "X": 10, "L": 50, "C": 100}


def roman_to_int(s: str) -> int | None:
    s = s.upper()
    if not s or any(c not in ROMAN for c in s):
        return None
    tot = 0
    for i, c in enumerate(s):
        v = ROMAN[c]
        if i + 1 < len(s) and ROMAN[s[i + 1]] > v:
            tot -= v
        else:
            tot += v
    return tot
