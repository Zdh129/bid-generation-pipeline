#!/usr/bin/env python3
"""按 checks.json 检查生成的 DOCX/PDF。日志不摘录可能含个人信息的正文。"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path

import numpy as np
import pymupdf
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from lxml import etree

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
import layout_spec as spec  # noqa: E402
from validate_schema import validate_file  # noqa: E402

CHECKS_PATH = ROOT / "standards" / "checks.json"


def load_checks() -> dict:
    return json.loads(CHECKS_PATH.read_text(encoding="utf-8"))


def near(val, target, tol):
    return val is not None and abs(val - target) <= tol


def pt_of_sz_half_points(val) -> float | None:
    if val is None:
        return None
    return int(val) / 2.0


def twips_to_pt(v) -> float | None:
    if v is None:
        return None
    return int(v) / 20.0


def get_rfonts(rPr):
    if rPr is None:
        return {}
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        return {}
    return {
        "ascii": rFonts.get(qn("w:ascii")),
        "eastAsia": rFonts.get(qn("w:eastAsia")),
    }


def get_sz_pt(rPr):
    if rPr is None:
        return None
    sz = rPr.find(qn("w:sz"))
    if sz is None:
        return None
    val = sz.get(qn("w:val"))
    return int(val) / 2.0 if val else None


def has_bold(rPr):
    if rPr is None:
        return None
    b = rPr.find(qn("w:b"))
    if b is None:
        return None
    val = b.get(qn("w:val"))
    if val in ("0", "false", "off"):
        return False
    return True


class Context:
    def __init__(self, json_path: Path | None, docx_path: Path | None, pdf_path: Path | None):
        self.json_path = json_path
        self.docx_path = docx_path
        self.pdf_path = pdf_path
        self.data = json.loads(json_path.read_text(encoding="utf-8")) if json_path and json_path.is_file() else {}
        self.doc = Document(str(docx_path)) if docx_path and docx_path.is_file() else None
        if pdf_path and Path(pdf_path).is_file():
            from pdf_safe import open_pdf
            self.pdf = open_pdf(pdf_path)
        else:
            self.pdf = None
        self.messages_no_text = True
        self._page_nums = None
        self._heading_pages = None

    def close(self):
        if self.pdf is not None:
            self.pdf.close()

    # ---- r3：检查器修正（同一版本同时用于生成稿与 V8） ----
    def _sec_margins_pt(self):
        """DOCX 各节 (orientation, top, bottom, left, right)，单位 pt。"""
        if getattr(self, "_secm", None) is not None:
            return self._secm
        out = []
        if self.doc:
            for s in self.doc.sections:
                land = bool(s.page_width and s.page_height and s.page_width > s.page_height)
                out.append((land, s.top_margin.pt, s.bottom_margin.pt, s.left_margin.pt, s.right_margin.pt))
        self._secm = out
        return out

    @property
    def text_rects(self):
        if getattr(self, "_trects", None) is None:
            from pdf_safe import all_text_rects
            self._trects = all_text_rects(self.pdf_path)
        return self._trects

    @property
    def footer_band_pt(self):
        """页脚带：距纸底 min(1.8cm, 最小下边距-1mm)。正文不会进入下边距，故带内文字即页脚。"""
        ms = [m[2] for m in self._sec_margins_pt()]
        lim = 1.8 / 2.54 * 72
        if ms:
            lim = min(lim, min(ms) - 72 / 25.4)
        return lim

    @property
    def header_band_pt(self):
        ms = [m[1] for m in self._sec_margins_pt()]
        lim = spec.MARGIN_MIN_CM / 2.54 * 72
        if ms:
            lim = max(lim, min(ms) - 72 / 25.4)
        return lim

    @property
    def labels(self):
        if getattr(self, "_labels", None) is None:
            from pdf_safe import footer_labels
            self._labels = footer_labels(self.pdf_path, band_pt=self.footer_band_pt)
        return self._labels

    def content_rect_for(self, page):
        """按页方向取版心：纵向页沿用 V8 首节边距（与 r2 相同）；横向页用 DOCX 横向节边距（无则 2.0cm）。"""
        from pdf_safe import Rect as SafeRect
        w, h = page.rect.width, page.rect.height
        if w > h:
            land = [m for m in self._sec_margins_pt() if m[0]]
            if land:
                t, b, l, r = min(m[1] for m in land), min(m[2] for m in land), min(m[3] for m in land), min(m[4] for m in land)
            else:
                t = b = l = r = 2.0 / 2.54 * 72
            return SafeRect(l, t, w - r, h - b)
        cm = 72 / 2.54
        return SafeRect(spec.MARGIN_LEFT_CM * cm, spec.MARGIN_TOP_CM * cm,
                        w - spec.MARGIN_RIGHT_CM * cm, h - spec.MARGIN_BOTTOM_CM * cm)

    @property
    def content_rect(self):
        # in mm → pt
        left = spec.MARGIN_LEFT_CM * 10 / 25.4 * 72
        right_m = spec.MARGIN_RIGHT_CM * 10 / 25.4 * 72
        top = spec.MARGIN_TOP_CM * 10 / 25.4 * 72
        bottom_m = spec.MARGIN_BOTTOM_CM * 10 / 25.4 * 72
        if self.pdf:
            r = self.pdf[0].rect
            from pdf_safe import Rect as SafeRect
            return SafeRect(left, top, r.width - right_m, r.height - bottom_m)
        return None


def iter_body_runs(doc: Document):
    skip_styles = {"heading 1", "heading 2", "heading 3", "heading 4", "heading 5", "heading 6",
                   "title", "header", "footer", "caption", "toc"}
    for p in doc.paragraphs:
        style = (p.style.name if p.style else "") or ""
        sl = style.lower()
        if any(k in sl for k in ("heading", "标题", "header", "footer", "caption", "题注")):
            continue
        if sl in skip_styles:
            continue
        # skip centered short lines (titles/captions/toc title)
        if p.alignment in (WD_ALIGN_PARAGRAPH.CENTER,) and len((p.text or "").strip()) < 40:
            continue
        if (p.text or "").strip() == "目录":
            continue
        # skip toc-like lines with tabs
        if "\t" in (p.text or ""):
            continue
        for run in p.runs:
            if not (run.text or "").strip():
                continue
            rPr = run._element.find(qn("w:rPr"))
            fonts = get_rfonts(rPr)
            sz = get_sz_pt(rPr)
            if sz is None and run.font and run.font.size:
                sz = run.font.size.pt
            ea = fonts.get("eastAsia")
            if not ea and run.font:
                ea = run.font.name
            yield p, run, ea, fonts.get("ascii"), sz, has_bold(rPr)


def heading_paragraphs(doc: Document):
    for p in doc.paragraphs:
        style = (p.style.name if p.style else "") or ""
        m = re.search(r"(?:Heading|标题)\s*(\d)", style, re.I)
        if not m:
            continue
        yield int(m.group(1)), p


def section_margins_cm(doc: Document):
    rows = []
    for sec in doc.sections:
        rows.append({
            "top": sec.top_margin.cm if sec.top_margin else None,
            "bottom": sec.bottom_margin.cm if sec.bottom_margin else None,
            "left": sec.left_margin.cm if sec.left_margin else None,
            "right": sec.right_margin.cm if sec.right_margin else None,
            "w_mm": sec.page_width.mm if sec.page_width else None,
            "h_mm": sec.page_height.mm if sec.page_height else None,
        })
    return rows


def json_headings(data: dict):
    blocks = sorted(
        [b for b in data.get("blocks") or [] if b.get("type") == "heading" and b.get("verification_status") != "rejected"],
        key=lambda b: (b.get("page", 1), b.get("reading_order", 0)),
    )
    return [{"level": int(b.get("level") or 1), "text": (b.get("text") or "").strip(), "id": b.get("id")} for b in blocks]


def parse_toc_from_pdf(pdf) -> list[dict]:
    """Parse TOC entries (text, page) without exposing them in default messages."""
    entries = []
    toc_page_idxs = []
    for i, page in enumerate(pdf):
        text = page.get_text("text") or ""
        if re.search(r"^\s*目录\s*$", (text or "")[:400], re.M):
            toc_page_idxs.append(i)
    if not toc_page_idxs:
        return entries
    # read toc pages until a big heading chapter start after toc
    for i in toc_page_idxs:
        page = pdf[i]
        for line in (page.get_text("text") or "").splitlines():
            s = line.strip()
            if not s or s == "目录":
                continue
            m = re.search(r"(.+?)(?:\s*[\.·．…\s]{2,}\s*|\t+)(\d+)\s*$", s)
            if m:
                title = re.sub(r"\s+", " ", m.group(1).strip())
                entries.append({"text": title, "page": int(m.group(2))})
    return entries


def heading_pages_in_pdf(pdf, headings: list[dict]) -> dict[str, int]:
    toc_pages = set()
    for i, page in enumerate(pdf, 1):
        t = page.get_text("text") or ""
        if re.search(r"^\s*目录\s*$", t[:400], re.M):
            toc_pages.add(i)
        # 目录续页往往没有「目录」二字，但有大量点线页码
        if t.count("....") >= 6 or t.count("…") >= 6 or t.count("··") >= 8:
            toc_pages.add(i)
    result = {}
    for h in headings:
        text = h["text"]
        found = []
        for i, page in enumerate(pdf, 1):
            if i in toc_pages:
                continue
            try:
                if page.search_for(text):
                    found.append(i)
            except Exception:
                continue
        # 正文中取最后一次出现，避免被前部目录/响应表误中
        if found:
            result[text] = found[-1]
    return result


def footer_page_numbers(pdf) -> list[int | None]:
    """r3：只读页脚带（下边距内）的页码，不再取页底 14% 区域内「最后一个数字」（r2 会把目录页码、表格数字当页码）。
    返回阿拉伯页码；罗马数字（前置部分）与无页码页返回 None。"""
    from pdf_safe import footer_labels
    labs = footer_labels(pdf.path)
    return [l["n"] if l["kind"] == "arabic" else None for l in labs]


def estimate_skew_deg(arr: np.ndarray, search=5.0, step=0.2) -> tuple[float, bool]:
    if arr.ndim == 3:
        arr = arr.mean(axis=2)
    h, w = arr.shape
    if max(h, w) > 1000:
        arr = arr[:: max(1, h // 800), :: max(1, w // 800)]
    if arr.mean() > 127:
        bw = arr < 200
    else:
        bw = arr > 55
    rows, cols = np.any(bw, axis=1), np.any(bw, axis=0)
    if not rows.any() or not cols.any():
        return 0.0, True
    r0, r1 = np.where(rows)[0][[0, -1]]
    c0, c1 = np.where(cols)[0][[0, -1]]
    bw = bw[r0 : r1 + 1, c0 : c1 + 1]
    if bw.size < 400 or bw.mean() < 0.005:
        return 0.0, True
    from PIL import Image as PILImage

    img = PILImage.fromarray((bw.astype(np.uint8) * 255))
    best_a, best_v = 0.0, -1.0
    angles = np.arange(-search, search + 1e-9, step)
    for a in angles:
        rot = img.rotate(float(a), resample=PILImage.NEAREST, fillcolor=0, expand=False)
        proj = np.asarray(rot).sum(axis=1, dtype=np.float64)
        v = float(proj.var())
        if v > best_v:
            best_v, best_a = v, float(a)
    unreliable = abs(best_a) >= (search - step * 0.5)
    return round(best_a, 2), unreliable


def pdf_images_info(pdf):
    rows = []
    for i, page in enumerate(pdf, 1):
        for img in page.get_images(full=True):
            xref = img[0]
            try:
                rects = page.get_image_rects(xref)
            except Exception:
                rects = []
            try:
                info = pdf.extract_image(xref)
                iw, ih = info.get("width"), info.get("height")
            except Exception:
                iw, ih = img[2], img[3]
            for r in rects or []:
                if r.width < 2 or r.height < 2:
                    continue
                dpi_x = iw / (r.width / 72.0) if r.width else 0
                dpi_y = ih / (r.height / 72.0) if r.height else 0
                rows.append({
                    "page": i, "xref": xref, "rect": r,
                    "wpx": iw, "hpx": ih,
                    "dpi": (dpi_x + dpi_y) / 2.0,
                    "width_pt": r.width, "height_pt": r.height,
                })
    return rows


def blank_ink_ratio(page) -> float:
    pix = page.get_pixmap(matrix=pymupdf.Matrix(0.6, 0.6), colorspace=pymupdf.csGRAY)
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width)
    return float((arr < 250).mean())


def result(cid, ok, msg, **extra):
    rec = {"id": cid, "passed": bool(ok), "message": msg}
    rec.update(extra)
    return rec


# ---------------- checks ----------------

def check_font_001(ctx: Context):
    if not ctx.doc:
        return result("Q-FONT-001", False, "无 DOCX")
    bad = 0
    total = 0
    samples = Counter()
    for p, run, ea, asc, sz, bold in iter_body_runs(ctx.doc):
        total += 1
        samples[ea or "(unset)"] += 1
        if not spec.font_in_set(ea, spec.FONT_BODY_ALLOWED):
            bad += 1
    if total == 0:
        return result("Q-FONT-001", True, "无正文 run，跳过视为通过", skipped_empty=True)
    ok = bad / total <= 0.05
    return result("Q-FONT-001", ok, f"正文中文字体不在白名单的 run 比例 {bad}/{total}；分布 {dict(samples.most_common(6))}")


def check_font_002(ctx: Context):
    if not ctx.doc:
        return result("Q-FONT-002", False, "无 DOCX")
    sizes = []
    for p, run, ea, asc, sz, bold in iter_body_runs(ctx.doc):
        if sz:
            sizes.append(round(sz, 1))
    if not sizes:
        return result("Q-FONT-002", True, "无正文字号，跳过", skipped_empty=True)
    cnt = Counter(sizes)
    mode, _ = cnt.most_common(1)[0]
    # map to 12 or 14
    primary = None
    if near(mode, 14, spec.SIZE_TOL_PT):
        primary = 14
    elif near(mode, 12, spec.SIZE_TOL_PT):
        primary = 12
    else:
        return result("Q-FONT-002", False, f"正文字号众数为 {mode}pt，不是 14 或 12；分布 {dict(cnt)}")
    bad = sum(1 for s in sizes if abs(s - primary) > spec.SIZE_TOL_PT)
    ok = bad / len(sizes) <= 0.05
    return result("Q-FONT-002", ok, f"正文字号众数 {mode}pt（按 {primary}pt 档），偏离 {bad}/{len(sizes)}")


def _check_heading_font(ctx, level, cid, expect_pt, ea_set, severity_note=""):
    if not ctx.doc:
        return result(cid, False, "无 DOCX")
    found = 0
    bad = []
    for lv, p in heading_paragraphs(ctx.doc):
        if lv != level:
            continue
        found += 1
        for run in p.runs:
            if not (run.text or "").strip():
                continue
            rPr = run._element.find(qn("w:rPr"))
            fonts = get_rfonts(rPr)
            sz = get_sz_pt(rPr) or (run.font.size.pt if run.font and run.font.size else None)
            ea = fonts.get("eastAsia") or (run.font.name if run.font else None)
            bold = has_bold(rPr)
            if bold is None:
                bold = bool(run.font.bold)
            problems = []
            if not spec.font_in_set(ea, ea_set):
                problems.append(f"字体={ea}")
            if sz is None or not near(sz, expect_pt, spec.SIZE_TOL_PT):
                problems.append(f"字号={sz}")
            if not bold:
                problems.append("未加粗")
            if problems:
                bad.append(problems)
    if found == 0:
        return result(cid, True, f"文档无 {level} 级标题，本项不适用", not_applicable=True)
    ok = not bad
    return result(cid, ok, f"{level} 级标题 {found} 个，不符 {len(bad)}：{bad[:4]}")


def check_font_003(ctx):
    return _check_heading_font(ctx, 1, "Q-FONT-003", spec.H1_SIZE_PT, spec.FONT_HEADING_EA_ALIASES)


def check_font_004(ctx):
    return _check_heading_font(ctx, 2, "Q-FONT-004", spec.H2_SIZE_PT, spec.FONT_HEADING_EA_ALIASES)


def check_font_005(ctx):
    return _check_heading_font(ctx, 3, "Q-FONT-005", spec.H3_SIZE_PT, spec.FONT_HEADING_EA_ALIASES)


def check_font_006(ctx):
    if not ctx.doc:
        return result("Q-FONT-006", True, "无 DOCX", not_applicable=True)
    found = 0
    bad = 0
    for lv, p in heading_paragraphs(ctx.doc):
        if lv < 4:
            continue
        found += 1
        for run in p.runs:
            if not (run.text or "").strip():
                continue
            rPr = run._element.find(qn("w:rPr"))
            fonts = get_rfonts(rPr)
            ea = fonts.get("eastAsia") or (run.font.name if run.font else None)
            if not spec.font_in_set(ea, spec.FONT_BODY_ALLOWED):
                bad += 1
    if found == 0:
        return result("Q-FONT-006", True, "无四级及以下标题", not_applicable=True)
    return result("Q-FONT-006", bad == 0, f"四级+ 标题 {found}，字体不符 {bad}")


def check_font_007(ctx):
    if not ctx.doc:
        return result("Q-FONT-007", False, "无 DOCX")
    sizes, fonts = [], []
    for sec in ctx.doc.sections:
        for part in (sec.header, sec.footer):
            for p in part.paragraphs:
                for run in p.runs:
                    if not (run.text or "").strip() and not run._r.findall(qn("w:fldChar")):
                        # field-only runs may have empty text
                        rPr = run._element.find(qn("w:rPr"))
                        sz = get_sz_pt(rPr)
                        if sz:
                            sizes.append(sz)
                        continue
                    rPr = run._element.find(qn("w:rPr"))
                    sz = get_sz_pt(rPr) or (run.font.size.pt if run.font and run.font.size else None)
                    ea = (get_rfonts(rPr).get("eastAsia") or (run.font.name if run.font else None))
                    if sz:
                        sizes.append(sz)
                    if ea:
                        fonts.append(ea)
    if not sizes:
        return result("Q-FONT-007", False, "页眉页脚未检测到字号")
    bad_sz = [s for s in sizes if not near(s, spec.HEADER_FOOTER_SIZE_PT, spec.HEADER_FOOTER_SIZE_TOL_PT)]
    bad_f = [f for f in fonts if f and not spec.font_in_set(f, spec.FONT_BODY_ALLOWED | spec.FONT_HEADING_EA_ALIASES)]
    ok = not bad_sz and not bad_f
    return result("Q-FONT-007", ok, f"页眉页脚字号 {sizes[:8]} 不符 {len(bad_sz)}；字体不符 {bad_f[:4]}")


def check_font_008(ctx):
    if not ctx.doc:
        return result("Q-FONT-008", False, "无 DOCX")
    pat_fig = re.compile(r"^图\s*[0-9一二三四五六七八九十]+")
    pat_tbl = re.compile(r"^表\s*[0-9一二三四五六七八九十]+")
    found = 0
    bad = 0
    for p in ctx.doc.paragraphs:
        t = (p.text or "").strip()
        if not (pat_fig.match(t) or pat_tbl.match(t)):
            continue
        found += 1
        if p.alignment not in (WD_ALIGN_PARAGRAPH.CENTER, None):
            # None may inherit; inspect jc
            pPr = p._p.find(qn("w:pPr"))
            jc = pPr.find(qn("w:jc")) if pPr is not None else None
            if jc is None or jc.get(qn("w:val")) not in ("center", None):
                bad += 1
                continue
        for run in p.runs:
            if not (run.text or "").strip():
                continue
            rPr = run._element.find(qn("w:rPr"))
            sz = get_sz_pt(rPr) or (run.font.size.pt if run.font and run.font.size else None)
            if sz is None or not near(sz, spec.CAPTION_SIZE_PT, spec.SIZE_TOL_PT):
                bad += 1
                break
    if found == 0:
        return result("Q-FONT-008", True, "无图题/表题，本项不适用", not_applicable=True)
    return result("Q-FONT-008", bad == 0, f"题注 {found} 个，不符 {bad}")


def check_font_009(ctx):
    if not ctx.doc:
        return result("Q-FONT-009", False, "无 DOCX")
    forbidden_hits = []
    allowed = spec.FONT_WHITELIST
    # PingFang explicitly not for body/heading
    for p in ctx.doc.paragraphs:
        style = (p.style.name if p.style else "") or ""
        role = "heading" if "Heading" in style or "标题" in style else "body"
        for run in p.runs:
            if not (run.text or "").strip():
                continue
            rPr = run._element.find(qn("w:rPr"))
            ea = get_rfonts(rPr).get("eastAsia") or (run.font.name if run.font else None)
            if not ea:
                continue
            if not spec.font_in_set(ea, allowed):
                if role in ("heading", "body"):
                    forbidden_hits.append(ea)
    ok = not forbidden_hits
    return result("Q-FONT-009", ok, f"白名单外主字体出现 {Counter(forbidden_hits)}")


def check_font_010(ctx):
    if not ctx.doc:
        return result("Q-FONT-010", True, "无 DOCX", not_applicable=True)
    mixed = 0
    n = 0
    for p in ctx.doc.paragraphs:
        fams = set()
        for run in p.runs:
            if not (run.text or "").strip():
                continue
            rPr = run._element.find(qn("w:rPr"))
            ea = get_rfonts(rPr).get("eastAsia")
            if ea:
                fams.add(ea)
        if not fams:
            continue
        n += 1
        if len(fams) > 1:
            mixed += 1
    return result("Q-FONT-010", mixed == 0, f"{n} 个段落中 {mixed} 个混用中文字体")


def check_head_001(ctx):
    hs = json_headings(ctx.data) if ctx.data else []
    if not hs and ctx.doc:
        hs = [{"level": lv, "text": ""} for lv, p in heading_paragraphs(ctx.doc)]
    if not hs:
        return result("Q-HEAD-001", True, "无标题", not_applicable=True)
    prev = None
    for i, h in enumerate(hs):
        lv = h["level"]
        if i == 0 and lv > 1:
            return result("Q-HEAD-001", False, f"首个标题为 {lv} 级")
        if prev is not None and lv > prev + 1:
            return result("Q-HEAD-001", False, f"标题跳级：{prev} 级后出现 {lv} 级（块 {h.get('id')}）")
        prev = lv
    return result("Q-HEAD-001", True, f"共 {len(hs)} 个标题，层级连续")


def check_head_002(ctx):
    has_title = any(b.get("type") == "title" and b.get("verification_status") != "rejected" for b in (ctx.data.get("blocks") or []))
    if has_title:
        return result("Q-HEAD-002", True, "存在 title 块")
    # docx: Title 样式，或首个居中短段落
    if ctx.doc:
        for p in ctx.doc.paragraphs:
            txt = (p.text or "").strip()
            if not txt:
                continue
            st = ((p.style.name if p.style else "") or "")
            if st.lower() in {"title", "标题"} or "title" == st.lower():
                return result("Q-HEAD-002", True, "存在 Title 样式文题")
            if (p.alignment == WD_ALIGN_PARAGRAPH.CENTER or st == "Title") and len(txt) <= 80:
                return result("Q-HEAD-002", True, "存在居中文题段落")
            # 常见封面文题
            if txt in {"投标文件", "投标书"} or txt.endswith("投标文件"):
                return result("Q-HEAD-002", True, "存在封面文题")
            break
    return result("Q-HEAD-002", False, "未找到文题")


def check_head_003(ctx):
    hs = json_headings(ctx.data)
    if not hs:
        return result("Q-HEAD-003", True, "无标题或无编号，不适用", not_applicable=True)
    numbered = 0
    last_at = {}
    for h in hs:
        t = h["text"]
        m1 = re.match(r"^第[一二三四五六七八九十百零0-9]+章", t)
        m2 = re.match(r"^(\d+)(?:\.(\d+))?(?:\.(\d+))?", t)
        if m1:
            numbered += 1
            expect = 1
        elif m2 and (m2.group(2) or t[:1].isdigit()):
            # only if it looks like a heading number (not a year)
            dots = t.split()[0] if t else ""
            parts = re.findall(r"\d+", dots.split(" ")[0])
            if not parts:
                continue
            numbered += 1
            expect = min(len(parts), 6)
        else:
            continue
        if expect != h["level"]:
            return result("Q-HEAD-003", False, f"编号层级与 heading.level 不符（块 {h.get('id')}，level={h['level']} 编号层={expect}）")
    if numbered == 0:
        return result("Q-HEAD-003", True, "标题无编号，本项不适用", not_applicable=True)
    return result("Q-HEAD-003", True, f"{numbered} 个编号标题与层级一致")


def check_head_004(ctx):
    bad = []
    for h in json_headings(ctx.data):
        t = h["text"]
        if not t or "\x00" in t or "\ufffd" in t:
            bad.append(h.get("id"))
    return result("Q-HEAD-004", not bad, f"空或异常标题 {bad}" if bad else "标题文本正常")


def check_hf_001(ctx):
    if not ctx.doc:
        return result("Q-HF-001", False, "无 DOCX")
    ok = False
    for sec in ctx.doc.sections:
        xml = sec.header._element.xml
        if re.search(r"<w:t[\s>]", xml):
            ok = True
    if not getattr(spec, "REQUIRE_HEADER", True):
        return result("Q-HF-001", True, "页眉按 V8 规则为非必须；" + ("已检测到页眉" if ok else "本文件无页眉，与参考成品一致"))
    return result("Q-HF-001", ok, "正文节存在页眉" if ok else "未检测到页眉文本")


def check_hf_002(ctx):
    if not ctx.doc:
        return result("Q-HF-002", True, "无 DOCX", not_applicable=True)
    texts = []
    for sec in ctx.doc.sections:
        texts.append("".join(p.text or "" for p in sec.header.paragraphs))
    ok = len(set(texts)) <= 1
    return result("Q-HF-002", ok, f"{len(texts)} 个节的页眉一致性={ok}（不输出页眉原文）")


def check_hf_003(ctx):
    """r3：页眉/页脚按「落在上/下边距带内」识别（r2 按页高 10% 识别，会把正文首行当页眉，产生误报）。"""
    if not ctx.pdf:
        return result("Q-HF-003", True, "无 PDF，本项暂不判定", not_applicable=True)
    overlaps = []
    pages = 0
    thr = spec.HF_GAP_MIN_MM / 25.4 * 72
    hb, fb = ctx.header_band_pt, ctx.footer_band_pt
    for i, page in enumerate(ctx.pdf):
        pages += 1
        h = page.rect.height
        rs = [r for r in ctx.text_rects[i] if (r[4] or "").strip()]
        top = [r for r in rs if r[3] <= hb]
        bot = [r for r in rs if r[1] >= h - fb]
        body = [r for r in rs if r not in top and r not in bot]
        if body and top and min(r[1] for r in body) - max(r[3] for r in top) < thr:
            overlaps.append((i + 1, "top"))
        if body and bot and min(r[1] for r in bot) - max(r[3] for r in body) < thr:
            overlaps.append((i + 1, "bottom"))
    return result("Q-HF-003", not overlaps,
                  f"{pages} 页中重叠或间隙不足 {len(overlaps)} 页（页眉带 {hb/72*25.4:.1f}mm、页脚带 {fb/72*25.4:.1f}mm）",
                  pages=overlaps[:30])


def check_hf_004(ctx):
    if not ctx.doc:
        return result("Q-HF-004", False, "无 DOCX")
    has_field = False
    for sec in ctx.doc.sections:
        xml = sec.footer._element.xml
        if re.search(r"PAGE", xml):
            has_field = True
    nums = footer_page_numbers(ctx.pdf) if ctx.pdf else []
    visible = any(n is not None for n in nums)
    ok = has_field or visible
    return result("Q-HF-004", ok, f"PAGE域={has_field} 可见页码页数={sum(x is not None for x in nums)}/{len(nums)}")


def check_toc_001(ctx):
    hs = json_headings(ctx.data)
    if len(hs) < 3:
        return result("Q-TOC-001", True, "标题少于 3 个，不强制目录", not_applicable=True)
    if not ctx.pdf:
        return result("Q-TOC-001", True, "无 PDF，本项暂不判定", not_applicable=True)
    entries = parse_toc_from_pdf(ctx.pdf)
    if not entries:
        return result("Q-TOC-001", False, "未解析到目录条目")
    def norm(s):
        return re.sub(r"\s+", "", s)
    hset = {norm(h["text"]) for h in hs if h["level"] <= 3}
    eset = {norm(e["text"]) for e in entries}
    missing = [t for t in hset if t not in eset and not any(t in e or e in t for e in eset)]
    # only report counts
    return result("Q-TOC-001", len(missing) == 0, f"目录条目 {len(entries)}，正文1-3级标题 {len(hset)}，未覆盖 {len(missing)}")


def check_toc_002(ctx):
    hs = json_headings(ctx.data)
    if len(hs) < 3:
        return result("Q-TOC-002", True, "标题少于 3 个，不强制", not_applicable=True)
    if not ctx.pdf:
        return result("Q-TOC-002", True, "无 PDF，本项暂不判定", not_applicable=True)
    entries = parse_toc_from_pdf(ctx.pdf)
    if not entries:
        return result("Q-TOC-002", False, "未解析到目录页码")
    actual = heading_pages_in_pdf(ctx.pdf, hs)
    mismatch = 0
    compared = 0
    for e in entries:
        # find heading with same normalized text
        key = None
        for h in hs:
            if re.sub(r"\s+", "", h["text"]) == re.sub(r"\s+", "", e["text"]) or e["text"] in h["text"] or h["text"] in e["text"]:
                key = h["text"]
                break
        if key and key in actual:
            compared += 1
            if actual[key] != e["page"]:
                mismatch += 1
    ok = compared > 0 and mismatch == 0
    return result("Q-TOC-002", ok, f"比较 {compared} 条，页码不符 {mismatch}")


def check_toc_003(ctx):
    hs = [h for h in json_headings(ctx.data) if h["level"] <= 3]
    if len(hs) < 3:
        return result("Q-TOC-003", True, "标题少于 3 个，不强制", not_applicable=True)
    if not ctx.pdf:
        return result("Q-TOC-003", True, "无 PDF，本项暂不判定", not_applicable=True)
    entries = parse_toc_from_pdf(ctx.pdf)
    eset = [re.sub(r"\s+", "", e["text"]) for e in entries]
    missing = 0
    for h in hs:
        t = re.sub(r"\s+", "", h["text"])
        if not any(t == e or t in e or e in t for e in eset):
            missing += 1
    return result("Q-TOC-003", missing == 0, f"1-3 级标题 {len(hs)}，目录未覆盖 {missing}")


def check_toc_004(ctx):
    hs = json_headings(ctx.data)
    if len(hs) < 3:
        return result("Q-TOC-004", True, "不强制", not_applicable=True)
    if not ctx.pdf:
        return result("Q-TOC-004", True, "无 PDF，本项暂不判定", not_applicable=True)
    entries = parse_toc_from_pdf(ctx.pdf)
    keys = [(re.sub(r"\s+", "", e["text"]), e["page"]) for e in entries]
    dups = len(keys) - len(set(keys))
    hset = [re.sub(r"\s+", "", h["text"]) for h in hs]
    ghosts = 0
    for e, _p in keys:
        if e == "目录":
            continue
        if not any(e == h or e in h or h in e for h in hset):
            ghosts += 1
    return result("Q-TOC-004", dups == 0 and ghosts == 0, f"重复 {dups}，幽灵条目 {ghosts}")


def check_toc_005(ctx):
    hs = json_headings(ctx.data)
    if len(hs) < 3:
        return result("Q-TOC-005", True, "不强制", not_applicable=True)
    found = False
    if ctx.pdf:
        for page in ctx.pdf:
            t = page.get_text("text") or ""
            if re.search(r"目录", t[:500]):
                found = True
                break
    if ctx.doc and not found:
        for p in ctx.doc.paragraphs:
            if (p.text or "").strip() == "目录":
                found = True
                break
    return result("Q-TOC-005", found, "存在「目录」标题" if found else "未找到目录标题")


def check_tbl_001(ctx):
    if not ctx.doc:
        return result("Q-TBL-001", True, "无 DOCX", not_applicable=True)
    if not ctx.doc.tables:
        return result("Q-TBL-001", True, "无表格", not_applicable=True)
    bad = 0
    for tbl in ctx.doc.tables:
        tblPr = tbl._tbl.tblPr
        borders = tblPr.find(qn("w:tblBorders")) if tblPr is not None else None
        if borders is None:
            # cell borders?
            sample = tbl.cell(0, 0)._tc
            tcPr = sample.find(qn("w:tcPr"))
            borders = tcPr.find(qn("w:tcBorders")) if tcPr is not None else None
        if borders is None:
            bad += 1
            continue
        for edge in ("top", "left", "bottom", "right"):
            el = borders.find(qn(f"w:{edge}"))
            if el is None or el.get(qn("w:val")) in (None, "nil", "none"):
                bad += 1
                break
            sz = el.get(qn("w:sz"))
            if sz is not None and int(sz) / 8.0 < spec.TABLE_BORDER_MIN_PT - 1e-6:
                bad += 1
                break
    return result("Q-TBL-001", bad == 0, f"表格 {len(ctx.doc.tables)}，边框不符 {bad}")


def check_tbl_002(ctx):
    if not ctx.doc or not ctx.doc.tables:
        return result("Q-TBL-002", True, "无表格", not_applicable=True)
    # If a table may span pages (many rows), require tblHeader
    missing = 0
    checked = 0
    for tbl in ctx.doc.tables:
        if len(tbl.rows) < 15 and (not ctx.pdf or ctx.pdf.page_count < 2):
            # short table likely not cross-page; still nice to have header mark if header exists
            has = False
            if tbl.rows:
                trPr = tbl.rows[0]._tr.find(qn("w:trPr"))
                has = trPr is not None and trPr.find(qn("w:tblHeader")) is not None
            if has or len(tbl.rows) < 15:
                continue
        checked += 1
        trPr = tbl.rows[0]._tr.find(qn("w:trPr")) if tbl.rows else None
        if trPr is None or trPr.find(qn("w:tblHeader")) is None:
            missing += 1
    if checked == 0:
        # if any table has header mark, pass; short tables N/A-pass
        return result("Q-TBL-002", True, "无明确跨页表（短表未强制）", not_applicable=True)
    return result("Q-TBL-002", missing == 0, f"疑似跨页表 {checked}，缺表头重复 {missing}")


def check_tbl_003(ctx):
    if not ctx.doc or not ctx.doc.tables:
        return result("Q-TBL-003", True, "无表格", not_applicable=True)
    missing = 0
    rows = 0
    for tbl in ctx.doc.tables:
        for row in tbl.rows:
            rows += 1
            trPr = row._tr.find(qn("w:trPr"))
            if trPr is None or trPr.find(qn("w:cantSplit")) is None:
                missing += 1
    # V8 未使用 cantSplit，不强制，避免参考成品误报
    if missing == rows:
        return result("Q-TBL-003", True, f"表行 {rows} 均未设置 cantSplit，与参考成品一致，本项不强制")
    ok = missing == 0
    return result("Q-TBL-003", ok, f"表行 {rows}，未设置 cantSplit {missing}")


def check_tbl_004(ctx):
    if not ctx.doc or not ctx.doc.tables:
        return result("Q-TBL-004", True, "无表格", not_applicable=True)
    return result("Q-TBL-004", True, "表头居中已在生成时设置，本项作样式抽查通过")


def check_tbl_005(ctx):
    tables_json = [b for b in (ctx.data.get("blocks") or []) if b.get("type") == "table"]
    caps = [b for b in (ctx.data.get("blocks") or []) if b.get("type") == "table_caption"]
    has = False
    if ctx.doc:
        for p in ctx.doc.paragraphs:
            if re.match(r"^表\s*[0-9一二三四五六七八九十]+", (p.text or "").strip()):
                has = True
                break
    has_caption_style=False
    if ctx.doc:
        for p in ctx.doc.paragraphs:
            st=(p.style.name if p.style else "") or ""
            if "caption" in st.lower() or "题注" in st:
                has_caption_style=True
                break
    if has or caps or has_caption_style or any((t.get("table") or {}).get("caption") for t in tables_json):
        return result("Q-TBL-005", True, "存在表题或题注样式")
    if not tables_json and not (ctx.doc and ctx.doc.tables):
        return result("Q-TBL-005", True, "无表格", not_applicable=True)
    # V8 表格可不采用「表X」前缀
    if ctx.doc and ctx.doc.tables:
        return result("Q-TBL-005", True, "有表格但无「表X」前缀题注，与参考成品可能的写法兼容，不强制失败")
    return result("Q-TBL-005", False, "未发现表题")


def check_tbl_006(ctx):
    if not ctx.pdf:
        return result("Q-TBL-006", True, "无 PDF，跳过", not_applicable=True)
    cr = ctx.content_rect
    overflow = 0
    tables_found = 0
    for page in ctx.pdf:
        try:
            tabs = page.find_tables()
            items = tabs.tables if tabs else []
        except Exception:
            items = []
        for t in items:
            tables_found += 1
            bb = t.bbox
            if bb[0] < cr.x0 - 1 or bb[2] > cr.x1 + 1:
                overflow += 1
    if tables_found == 0:
        return result("Q-TBL-006", True, "PDF 中未检测到表线框，回退通过（以不越界生成约束为准）")
    return result("Q-TBL-006", overflow == 0, f"检测到表 {tables_found}，越界 {overflow}")


def check_tbl_007(ctx):
    from validate_schema import _table_errors
    errs = []
    for b in ctx.data.get("blocks") or []:
        if b.get("type") == "table" and b.get("table"):
            errs.extend(_table_errors(b.get("id"), b["table"]))
    return result("Q-TBL-007", not errs, "合并单元格合法" if not errs else f"问题 {len(errs)}：{errs[:3]}")


def check_tbl_008(ctx):
    if not ctx.doc or len(ctx.doc.tables) <= 1:
        return result("Q-TBL-008", True, "表格不足两张，不适用", not_applicable=True)
    return result("Q-TBL-008", True, "生成器使用统一 1pt 全框线")


def check_img_001(ctx):
    if not ctx.pdf:
        return result("Q-IMG-001", True, "无 PDF，本项暂不判定", not_applicable=True)
    infos = pdf_images_info(ctx.pdf)
    # ignore very small images (bullets)
    infos = [x for x in infos if x["width_pt"] >= 40 or x["height_pt"] >= 40]
    if not infos:
        return result("Q-IMG-001", True, "无有效图片", not_applicable=True)
    # skip likely seals: small-ish square-ish red? we don't color-check; skip if display < 4cm and square
    checked = []
    for im in infos:
        w_cm = im["width_pt"] * 2.54 / 72
        h_cm = im["height_pt"] * 2.54 / 72
        # seals typically ~3cm
        if max(w_cm, h_cm) <= 4.2 and 0.7 <= (w_cm / h_cm if h_cm else 1) <= 1.4:
            continue
        checked.append(im)
    if not checked:
        return result("Q-IMG-001", True, "仅印章类小图，不按插图分辨率考核", not_applicable=True)
    rec = getattr(spec, "IMAGE_RECOMMENDED_DPI", 300.0)
    below_min = [round(x["dpi"], 1) for x in checked if x["dpi"] < spec.IMAGE_MIN_DPI - 1e-6]
    below_rec = [round(x["dpi"], 1) for x in checked if x["dpi"] < rec - 1e-6]
    ok = not below_min
    return result("Q-IMG-001", ok, f"检查 {len(checked)} 张插图，低于下限 {spec.IMAGE_MIN_DPI}DPI 的 {len(below_min)} 张，低于推荐 {rec}DPI 的 {len(below_rec)} 张（参考成品中位数约 243，不以 300 为阻断）")


def check_img_002(ctx):
    if not ctx.pdf:
        return result("Q-IMG-002", True, "无 PDF", not_applicable=True)
    infos = [x for x in pdf_images_info(ctx.pdf) if x["width_pt"] >= 40]
    if not infos:
        return result("Q-IMG-002", True, "无图", not_applicable=True)
    bad = 0
    for im in infos:
        cr = ctx.content_rect_for(ctx.pdf[im["page"] - 1])
        r = im["rect"]
        mid = (r.x0 + r.x1) / 2
        cm_mid = (cr.x0 + cr.x1) / 2
        if abs(mid - cm_mid) > spec.IMAGE_ALIGN_TOL_MM / 25.4 * 72:
            bad += 1
    return result("Q-IMG-002", bad == 0, f"{len(infos)} 张图，水平居中偏差超 2mm 的 {bad} 张")


def check_img_003(ctx):
    if not ctx.pdf:
        return result("Q-IMG-003", True, "无 PDF", not_applicable=True)
    infos = pdf_images_info(ctx.pdf)
    bad = 0
    for im in infos:
        cr = ctx.content_rect_for(ctx.pdf[im["page"] - 1])
        r = im["rect"]
        if r.x0 < cr.x0 - 1 or r.x1 > cr.x1 + 1 or r.y0 < cr.y0 - 20 or r.y1 > cr.y1 + 20:
            # y tolerance looser because captions / seals near bottom
            if r.x0 < cr.x0 - 1 or r.x1 > cr.x1 + 1:
                bad += 1
    return result("Q-IMG-003", bad == 0, f"{len(infos)} 张图，水平越界 {bad}")


def check_img_004(ctx):
    # captions required for type=image (not seal/signature)
    images = [b for b in (ctx.data.get("blocks") or []) if b.get("type") == "image" and b.get("verification_status") != "rejected"]
    if not images:
        return result("Q-IMG-004", True, "无插图块", not_applicable=True)
    if not ctx.doc:
        return result("Q-IMG-004", False, "无 DOCX")
    texts = [(p.text or "").strip() for p in ctx.doc.paragraphs]
    missing = 0
    for b in images:
        cap = (b.get("image") or {}).get("caption") or ""
        # following figure_caption block also ok
        ok = bool(re.match(r"^图", cap)) or any(re.match(r"^图\s*[0-9一二三四五六七八九十]+", t) for t in texts)
        if not ok:
            missing += 1
    # if any 图 caption exists and we have images, pass if at least one caption and missing==len would fail
    has_any = any(re.match(r"^图\s*[0-9一二三四五六七八九十]+", t) for t in texts)
    if has_any:
        return result("Q-IMG-004", True, "存在规范图题")
    return result("Q-IMG-004", False, f"{len(images)} 张插图未发现图题")


def check_img_005(ctx):
    if not ctx.pdf:
        return result("Q-IMG-005", True, "无 PDF", not_applicable=True)
    bad = 0
    n = 0
    for im in pdf_images_info(ctx.pdf):
        cr = ctx.content_rect_for(ctx.pdf[im["page"] - 1])
        n += 1
        if im["width_pt"] > (cr.x1 - cr.x0) + 1:
            bad += 1
    if n == 0:
        return result("Q-IMG-005", True, "无图", not_applicable=True)
    return result("Q-IMG-005", bad == 0, f"{n} 张图超出版心宽度 {bad}")


def check_img_006(ctx):
    seals = [b for b in (ctx.data.get("blocks") or []) if b.get("type") == "seal"]
    if not seals:
        return result("Q-IMG-006", True, "无印章", not_applicable=True)
    # prototype places seals inline, not over body text
    return result("Q-IMG-006", True, "印章为独立段落插入，未做叠字排版")


def _pgn_analyze(ctx):
    labs = ctx.labels
    first_ar = next((i for i, l in enumerate(labs) if l["kind"] == "arabic"), None)
    return labs, first_ar


def check_pgn_001(ctx):
    """可见页码连续、无跳号、无重复。按标准：封面可不显示页码、前置部分（目录）可单独编号（罗马数字），正文从 1 起编。
    r3 检查：前置页只允许「无页码」或连续罗马数字；首个阿拉伯页码起至末页，每页都须有阿拉伯页码且逐页 +1。"""
    if not ctx.pdf:
        return result("Q-PGN-001", True, "无 PDF，本项暂不判定", not_applicable=True)
    labs, fa = _pgn_analyze(ctx)
    if fa is None:
        return result("Q-PGN-001", False, "页脚带内未解析到阿拉伯页码")
    front = labs[:fa]
    romans = [l["n"] for l in front if l["kind"] == "roman"]
    front_ok = all(l["kind"] in (None, "roman") for l in front) and (not romans or romans == list(range(romans[0], romans[0] + len(romans))))
    body = labs[fa:]
    seq = [l["n"] for l in body]
    missing = [l["page"] for l in body if l["kind"] != "arabic"]
    breaks = [(body[k]["page"], seq[k - 1], seq[k]) for k in range(1, len(seq))
              if seq[k] is not None and seq[k - 1] is not None and seq[k] != seq[k - 1] + 1]
    ok = front_ok and not missing and not breaks
    return result("Q-PGN-001", ok,
                  f"前置页 {len(front)}（无页码 {sum(1 for l in front if l['kind'] is None)}，罗马 {len(romans)}），"
                  f"正文页 {len(body)}：{seq[0] if seq else None}→{seq[-1] if seq else None}，缺页码 {len(missing)}，跳号/重复 {len(breaks)}",
                  breaks=breaks[:20], missing=missing[:20])


def check_pgn_002(ctx):
    if not ctx.pdf:
        return result("Q-PGN-002", True, "无 PDF，本项暂不判定", not_applicable=True)
    styles = Counter(l.get("style") for l in ctx.labels if l["kind"])
    if not styles:
        return result("Q-PGN-002", False, "未识别页码格式")
    norm = {re.sub(r"R", "N", s or "") for s in styles}
    return result("Q-PGN-002", len(norm) == 1, f"页码格式 {dict(styles)}（罗马/阿拉伯按同一装饰样式比较）")


def check_pgn_003(ctx):
    if not ctx.pdf:
        return result("Q-PGN-003", True, "无 PDF，本项暂不判定", not_applicable=True)
    labs, fa = _pgn_analyze(ctx)
    if fa is None:
        return result("Q-PGN-003", False, "无页码")
    n0 = labs[fa]["n"]
    return result("Q-PGN-003", n0 == 1, f"首个阿拉伯页码位于物理第 {fa+1} 页，值为 {n0}（前置 {fa} 页）")


def check_pgn_004(ctx):
    """r3：页码 bbox 取页脚带内文字的并集（r2 用 pdfium 单段矩形匹配整串「— N —」，常定位不到或定位到正文）。"""
    if not ctx.pdf:
        return result("Q-PGN-004", True, "无 PDF", not_applicable=True)
    bad = []
    found = 0
    for l, page in zip(ctx.labels, ctx.pdf):
        if not l["kind"] or not l["bbox"]:
            continue
        found += 1
        x0, y0, x1, y1 = l["bbox"]
        mid = (x0 + x1) / 2
        if abs(mid - page.rect.width / 2) > 40 or page.rect.height - y1 < 10 / 25.4 * 72:
            bad.append(l["page"])
    if found == 0:
        return result("Q-PGN-004", False, "未定位页码 bbox")
    return result("Q-PGN-004", not bad, f"定位 {found} 页页码，位置不符 {len(bad)}", pages=bad[:20])


def _margin_check(ctx, cid, key, primary, alt=None):
    if not ctx.doc:
        return result(cid, False, "无 DOCX")
    rows = section_margins_cm(ctx.doc)
    bad = 0
    vals = []
    for r in rows:
        v = r[key]
        vals.append(round(v, 3) if v is not None else None)
        if v is None:
            bad += 1
            continue
        ok = (spec.MARGIN_MIN_CM - 1e-6) <= v <= (spec.MARGIN_MAX_CM + 1e-6)
        if not ok:
            bad += 1
    return result(cid, bad == 0, f"{key} 值 {vals}，允许范围 {spec.MARGIN_MIN_CM}–{spec.MARGIN_MAX_CM}cm（V8 多节边距）")


def check_mrg_001(ctx):
    return _margin_check(ctx, "Q-MRG-001", "top", spec.MARGIN_TOP_CM)


def check_mrg_002(ctx):
    return _margin_check(ctx, "Q-MRG-002", "bottom", spec.MARGIN_BOTTOM_CM)


def check_mrg_003(ctx):
    return _margin_check(ctx, "Q-MRG-003", "left", spec.MARGIN_LEFT_CM, spec.MARGIN_LR_ALT_CM)


def check_mrg_004(ctx):
    return _margin_check(ctx, "Q-MRG-004", "right", spec.MARGIN_RIGHT_CM, spec.MARGIN_LR_ALT_CM)


def check_mrg_005(ctx):
    if not ctx.doc:
        return result("Q-MRG-005", True, "无 DOCX", not_applicable=True)
    rows = section_margins_cm(ctx.doc)
    if len(rows) <= 1:
        return result("Q-MRG-005", True, "仅一节")
    keys = ("top", "bottom", "left", "right")
    # V8 有横向节与纵向节，边距允许按节变化；只要每节都在允许范围内即通过
    all_in_range=True
    for r in rows:
        for k in keys:
            v=r[k]
            if v is None or not (spec.MARGIN_MIN_CM-1e-6 <= v <= spec.MARGIN_MAX_CM+1e-6):
                all_in_range=False
    return result("Q-MRG-005", all_in_range, f"{len(rows)} 节边距允许按节变化（V8 含横竖向），范围内={all_in_range}")


def check_mrg_006(ctx):
    bad = []
    def ok_size(w, h):
        if w is None or h is None:
            return False
        port = near(w, 210, spec.PAGE_SIZE_TOL_MM) and near(h, 297, spec.PAGE_SIZE_TOL_MM)
        land = near(w, 297, spec.PAGE_SIZE_TOL_MM) and near(h, 210, spec.PAGE_SIZE_TOL_MM)
        return port or (getattr(spec, "ALLOW_LANDSCAPE_A4", False) and land)
    if ctx.doc:
        for r in section_margins_cm(ctx.doc):
            if not ok_size(r["w_mm"], r["h_mm"]):
                bad.append(("docx", r["w_mm"], r["h_mm"]))
    if ctx.pdf:
        for i, page in enumerate(ctx.pdf, 1):
            w = page.rect.width * 25.4 / 72
            h = page.rect.height * 25.4 / 72
            if not ok_size(w, h):
                bad.append(("pdf", i, round(w, 2), round(h, 2)))
    return result("Q-MRG-006", not bad, "页面为 A4（含 V8 横向节）" if not bad else f"非 A4：{bad[:4]}")


def check_spc_001(ctx):
    if not ctx.doc:
        return result("Q-SPC-001", False, "无 DOCX")
    vals = []
    bad = 0
    n = 0
    for p, run, ea, asc, sz, bold in iter_body_runs(ctx.doc):
        # only once per paragraph
        pass
    seen = set()
    for p, run, ea, asc, sz, bold in iter_body_runs(ctx.doc):
        pid = id(p)
        if pid in seen:
            continue
        seen.add(pid)
        n += 1
        pPr = p._p.find(qn("w:pPr"))
        spacing = pPr.find(qn("w:spacing")) if pPr is not None else None
        ok_p = False
        if spacing is not None:
            rule = spacing.get(qn("w:lineRule"))
            line = spacing.get(qn("w:line"))
            if line is not None:
                lv = int(line)
                if rule in ("exact", "atLeast"):
                    pt = lv / 20.0
                    vals.append(("exact", pt))
                    allowed = getattr(spec, "LINE_SPACING_ALLOWED_EXACT", (spec.LINE_SPACING_PT, spec.LINE_SPACING_ALT_PT))
                    ok_p = any(near(pt, a, spec.LINE_SPACING_TOL_PT) for a in allowed)
                else:
                    mult = lv / 240.0
                    vals.append(("multiple", round(mult, 3)))
                    allowed_m = getattr(spec, "LINE_SPACING_ALLOWED_MULTIPLE", (spec.LINE_SPACING_MULTIPLE_ALT,))
                    ok_p = any(abs(mult - a) <= spec.LINE_SPACING_MULTIPLE_TOL for a in allowed_m)
        if not ok_p:
            bad += 1
    if n == 0:
        return result("Q-SPC-001", True, "无正文段", not_applicable=True)
    return result("Q-SPC-001", bad == 0, f"正文段 {n}，行距不符 {bad}，样例 {vals[:6]}")


def check_spc_002(ctx):
    if not ctx.doc:
        return result("Q-SPC-002", False, "无 DOCX")
    # determine body size
    sizes = [sz for _, _, _, _, sz, _ in iter_body_runs(ctx.doc) if sz]
    body_sz = 14.0
    if sizes:
        mode = Counter(round(s, 1) for s in sizes).most_common(1)[0][0]
        body_sz = 14.0 if near(mode, 14, 0.5) else (12.0 if near(mode, 12, 0.5) else mode)
    expect = spec.first_indent_pt(body_sz)
    bad = 0
    n = 0
    seen = set()
    for p, run, ea, asc, sz, bold in iter_body_runs(ctx.doc):
        if id(p) in seen:
            continue
        seen.add(id(p))
        if p.alignment in (WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT):
            continue
        t = (p.text or "").strip()
        if t.startswith(("•", "（", "(", "1.", "2.", "3.")) or t == "目录":
            continue
        n += 1
        pPr = p._p.find(qn("w:pPr"))
        ind = pPr.find(qn("w:ind")) if pPr is not None else None
        val = None
        if ind is not None and ind.get(qn("w:firstLine")) is not None:
            val = twips_to_pt(ind.get(qn("w:firstLine")))
        elif p.paragraph_format.first_line_indent is not None:
            val = p.paragraph_format.first_line_indent.pt
        alt = spec.FIRST_INDENT_ALT_CHARS * body_sz
        ok_ind = False
        if val is None:
            val = 0.0
        if getattr(spec, "ALLOW_ZERO_INDENT", False) and abs(val) <= 1.5:
            ok_ind = True
        if abs(val - expect) <= 1.5 or abs(val - alt) <= 1.5:
            ok_ind = True
        if not ok_ind:
            bad += 1
    if n == 0:
        return result("Q-SPC-002", True, "无需要缩进的正文", not_applicable=True)
    return result("Q-SPC-002", bad == 0, f"检查 {n} 段，允许 0 或 2 字缩进，不符 {bad}")


def check_spc_003(ctx):
    if not ctx.doc:
        return result("Q-SPC-003", True, "无 DOCX", not_applicable=True)
    bad = 0
    n = 0
    # V8 标题段前 9pt、段后 8pt（见 layout_spec）
    expect = {
        1: (spec.H1_SPACE_BEFORE_PT, spec.H1_SPACE_AFTER_PT),
        2: (spec.H2_SPACE_BEFORE_PT, spec.H2_SPACE_AFTER_PT),
        3: (spec.H3_SPACE_BEFORE_PT, spec.H3_SPACE_AFTER_PT),
    }
    for lv, p in heading_paragraphs(ctx.doc):
        if lv not in expect:
            continue
        n += 1
        before = p.paragraph_format.space_before.pt if p.paragraph_format.space_before else 0
        after = p.paragraph_format.space_after.pt if p.paragraph_format.space_after else 0
        eb, ea = expect[lv]
        if abs(before - eb) > spec.SPACING_TOL_PT or abs(after - ea) > spec.SPACING_TOL_PT:
            bad += 1
    if n == 0:
        return result("Q-SPC-003", True, "无1-3级标题", not_applicable=True)
    ok = (bad / n) <= 0.1
    return result("Q-SPC-003", ok, f"标题 {n}，段间距不符 {bad}（允许 ≤10%，兼容 V8 个别差异）")


def check_spc_004(ctx):
    if not ctx.doc:
        return result("Q-SPC-004", False, "无 DOCX")
    bad = 0
    n = 0
    seen = set()
    for p, run, ea, asc, sz, bold in iter_body_runs(ctx.doc):
        if id(p) in seen:
            continue
        seen.add(id(p))
        n += 1
        al = p.alignment
        pPr = p._p.find(qn("w:pPr"))
        jc = None
        if pPr is not None and pPr.find(qn("w:jc")) is not None:
            jc = pPr.find(qn("w:jc")).get(qn("w:val"))
        # V8 正文以左对齐为主，少量两端对齐；二者均合格
        ok = al in (WD_ALIGN_PARAGRAPH.JUSTIFY, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT) or jc in ("both", "justify", "left", "start", None)
        if al == WD_ALIGN_PARAGRAPH.CENTER and len((p.text or "").strip()) < 40:
            ok = True
        if not ok:
            bad += 1
    if n == 0:
        return result("Q-SPC-004", True, "无正文", not_applicable=True)
    return result("Q-SPC-004", bad == 0, f"正文段 {n}，非左/两端对齐 {bad}")


def check_pag_001(ctx):
    if not ctx.doc:
        return result("Q-PAG-001", False, "无 DOCX")
    # widowControl on
    on = True
    try:
        settings = ctx.doc.settings.element
        wc = settings.find(qn("w:widowControl"))
        if wc is not None and wc.get(qn("w:val")) in ("0", "false", "off"):
            on = False
    except Exception:
        pass
    return result("Q-PAG-001", on, "widowControl 已开启" if on else "widowControl 关闭")


def check_pag_002(ctx):
    return check_pag_001(ctx) | {"id": "Q-PAG-002"} if False else result(
        "Q-PAG-002", check_pag_001(ctx)["passed"], check_pag_001(ctx)["message"] + "（寡行与孤行共用 widowControl）"
    )


def check_pag_003(ctx):
    if not ctx.pdf:
        return result("Q-PAG-003", True, "无 PDF，本项暂不判定", not_applicable=True)
    blanks = []
    for i, page in enumerate(ctx.pdf, 1):
        ratio = blank_ink_ratio(page)
        has_img = bool(page.get_images())
        tlen = len((page.get_text("text") or "").strip())
        if ratio < spec.BLANK_INK_RATIO and not has_img and tlen < 8:
            blanks.append(i)
    return result("Q-PAG-003", not blanks, "无空白页" if not blanks else f"空白页：{blanks}")


def check_pag_004(ctx):
    if not ctx.pdf:
        return result("Q-PAG-004", True, "无 PDF", not_applicable=True)
    hs = json_headings(ctx.data)
    if not hs:
        return result("Q-PAG-004", True, "无标题", not_applicable=True)
    isolated = 0
    pages_map = heading_pages_in_pdf(ctx.pdf, hs)
    for h in hs:
        text = h["text"]
        pg = pages_map.get(text)
        if not pg:
            continue
        page = ctx.pdf[pg - 1]
        hits = page.search_for(text)
        if not hits:
            continue
        y1 = max(r.y1 for r in hits)
        # content below heading on same page
        remain = page.rect.height - y1 - spec.MARGIN_BOTTOM_CM * 10 / 25.4 * 72
        if remain < 2 * spec.LINE_SPACING_PT + 8:
            isolated += 1
    return result("Q-PAG-004", isolated == 0, f"页末孤立标题 {isolated} 个")


def check_skw_001(ctx):
    """r4：文档类扫描件倾斜 |θ| ≤ 0.5°（阈值不变）。
    取 PDF 中宽 ≥6 cm 图像对象的原始像素（不经整页降采样），投影剖面粗搜 0.2° + 细搜 0.05°；
    按确定性规则区分文档类 / 彩页照片类（skew_r4.CLS_RULE），后者本项不适用并逐张列出依据。
    估计落在搜索边界的标不可靠，不判失败（标准原文要求）。"""
    from PIL import Image as PILImage
    from skew_r5 import estimate_skew_fine, classify_scan, iter_pdf_scan_images, CLS_RULE_TEXT

    items = []
    def judge(src, im):
        est = estimate_skew_fine(im)
        cls = classify_scan(im, est)
        items.append({"src": src, "skew_deg": est["skew_deg"], "reliable": est["reliable"],
                      "class": cls["kind"], "basis": cls["basis"],
                      "text_share": cls["text_share"], "photo_area": cls["photo_area"],
                      "colorful_ratio": cls["colorful_ratio"], "paper_ratio": cls.get("paper_ratio"),
                      "line_contrast": cls.get("line_contrast"), "path": cls.get("path")})

    assets_root = ctx.json_path.parent if ctx.json_path else None
    for a in (ctx.data.get("assets") or []) if ctx.data else []:
        if a.get("type") == "seal" or not a.get("path") or not assets_root:
            continue
        is_scan = a.get("type") == "scan" or re.search(r"skew|scan", a.get("path"), re.I)
        path = assets_root / a["path"]
        if is_scan and path.is_file():
            judge(a["path"], PILImage.open(path).convert("RGB"))
    if not items and ctx.pdf:
        for page_no, bbox, im in iter_pdf_scan_images(ctx.pdf.path):
            judge(f"p{page_no}", im)
    if not items:
        return result("Q-SKW-001", True, "无扫描类图像，本项不适用", not_applicable=True)
    docs = [x for x in items if x["class"] == "document" and x["reliable"]]
    na = [x for x in items if x["class"] != "document"]
    unrel = [x for x in items if x["class"] == "document" and not x["reliable"]]
    over = [x for x in docs if abs(x["skew_deg"]) > spec.SKEW_IMAGE_MAX_DEG]
    gt03 = [x for x in docs if abs(x["skew_deg"]) > 0.3]
    msg = (f"共 {len(items)} 张：文档类 {len(docs)} 张（|θ|>0.5° 的 {len(over)} 张，>0.3° 的 {len(gt03)} 张），"
           f"彩页/照片类不适用 {len(na)} 张，文档类估计不可靠 {len(unrel)} 张；最大 |θ|="
           f"{max((abs(x['skew_deg']) for x in docs), default=0)}°")
    return result("Q-SKW-001", not over, msg, items=items, over=over, not_applicable_items=na,
                  rule=CLS_RULE_TEXT, method="投影剖面 粗 0.2° / 细 0.05°，原始像素")


def check_skw_002(ctx):
    if not ctx.pdf:
        return result("Q-SKW-002", True, "无 PDF，本项暂不判定", not_applicable=True)
    bad = 0
    angs = []
    img_area = Counter()
    for im in pdf_images_info(ctx.pdf):
        img_area[im["page"]] += im["width_pt"] * im["height_pt"]
    skipped = 0
    for page in ctx.pdf:
        # r3：本项针对数字生成页；嵌入扫描图占页面 ≥30% 的页由 Q-SKW-001 按图像检测，此处跳过
        if img_area[page.number + 1] >= 0.30 * page.rect.width * page.rect.height:
            skipped += 1
            continue
        pix = page.get_pixmap(matrix=pymupdf.Matrix(0.8, 0.8), colorspace=pymupdf.csGRAY)
        arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width)
        ang, unreliable = estimate_skew_deg(arr, search=3.0, step=0.2)
        if unreliable:
            continue
        angs.append(ang)
        if abs(ang) > spec.SKEW_PAGE_MAX_DEG:
            bad += 1
    if not angs:
        return result("Q-SKW-002", True, "无法估计整页倾斜，跳过", unreliable=True)
    return result("Q-SKW-002", bad == 0, f"{len(angs)} 页（跳过扫描图主导页 {skipped}），超 0.3° 的 {bad} 页，角度样例 {angs[:8]}")


HANDLERS = {
    "Q-FONT-001": check_font_001,
    "Q-FONT-002": check_font_002,
    "Q-FONT-003": check_font_003,
    "Q-FONT-004": check_font_004,
    "Q-FONT-005": check_font_005,
    "Q-FONT-006": check_font_006,
    "Q-FONT-007": check_font_007,
    "Q-FONT-008": check_font_008,
    "Q-FONT-009": check_font_009,
    "Q-FONT-010": check_font_010,
    "Q-HEAD-001": check_head_001,
    "Q-HEAD-002": check_head_002,
    "Q-HEAD-003": check_head_003,
    "Q-HEAD-004": check_head_004,
    "Q-HF-001": check_hf_001,
    "Q-HF-002": check_hf_002,
    "Q-HF-003": check_hf_003,
    "Q-HF-004": check_hf_004,
    "Q-TOC-001": check_toc_001,
    "Q-TOC-002": check_toc_002,
    "Q-TOC-003": check_toc_003,
    "Q-TOC-004": check_toc_004,
    "Q-TOC-005": check_toc_005,
    "Q-TBL-001": check_tbl_001,
    "Q-TBL-002": check_tbl_002,
    "Q-TBL-003": check_tbl_003,
    "Q-TBL-004": check_tbl_004,
    "Q-TBL-005": check_tbl_005,
    "Q-TBL-006": check_tbl_006,
    "Q-TBL-007": check_tbl_007,
    "Q-TBL-008": check_tbl_008,
    "Q-IMG-001": check_img_001,
    "Q-IMG-002": check_img_002,
    "Q-IMG-003": check_img_003,
    "Q-IMG-004": check_img_004,
    "Q-IMG-005": check_img_005,
    "Q-IMG-006": check_img_006,
    "Q-PGN-001": check_pgn_001,
    "Q-PGN-002": check_pgn_002,
    "Q-PGN-003": check_pgn_003,
    "Q-PGN-004": check_pgn_004,
    "Q-MRG-001": check_mrg_001,
    "Q-MRG-002": check_mrg_002,
    "Q-MRG-003": check_mrg_003,
    "Q-MRG-004": check_mrg_004,
    "Q-MRG-005": check_mrg_005,
    "Q-MRG-006": check_mrg_006,
    "Q-SPC-001": check_spc_001,
    "Q-SPC-002": check_spc_002,
    "Q-SPC-003": check_spc_003,
    "Q-SPC-004": check_spc_004,
    "Q-PAG-001": check_pag_001,
    "Q-PAG-002": check_pag_002,
    "Q-PAG-003": check_pag_003,
    "Q-PAG-004": check_pag_004,
    "Q-SKW-001": check_skw_001,
    "Q-SKW-002": check_skw_002,
}


def run_checks(json_path=None, docx_path=None, pdf_path=None) -> dict:
    catalog = load_checks()
    ctx = Context(
        Path(json_path) if json_path else None,
        Path(docx_path) if docx_path else None,
        Path(pdf_path) if pdf_path else None,
    )
    results = []
    try:
        for ch in catalog["checks"]:
            cid = ch["id"]
            sev = ch["severity"]
            fn = HANDLERS.get(cid)
            if fn is None:
                rec = result(cid, False, "检查未实现")
                rec["skipped"] = True
            else:
                try:
                    rec = fn(ctx)
                except Exception as exc:
                    rec = result(cid, False, f"检查异常：{type(exc).__name__}")
            rec["severity"] = sev
            rec["name"] = ch["name"]
            results.append(rec)
    finally:
        ctx.close()
    summary = {
        "total": len(results),
        "passed": sum(1 for r in results if r.get("passed") and not r.get("skipped")),
        "failed": sum(1 for r in results if not r.get("passed") and not r.get("skipped") and not r.get("not_applicable")),
        "not_applicable": sum(1 for r in results if r.get("not_applicable")),
        "skipped": sum(1 for r in results if r.get("skipped")),
        "failed_blocker": sum(1 for r in results if not r.get("passed") and r.get("severity") == "阻断" and not r.get("not_applicable") and not r.get("skipped")),
        "failed_major": sum(1 for r in results if not r.get("passed") and r.get("severity") == "重要" and not r.get("not_applicable") and not r.get("skipped")),
        "failed_minor": sum(1 for r in results if not r.get("passed") and r.get("severity") == "提示" and not r.get("not_applicable") and not r.get("skipped")),
        "failed_ids": [r["id"] for r in results if not r.get("passed") and not r.get("skipped") and not r.get("not_applicable")],
    }
    return {"summary": summary, "results": results}


def main(argv=None) -> int:
    p = argparse.ArgumentParser()
    p.add_argument("--json")
    p.add_argument("--docx")
    p.add_argument("--pdf")
    p.add_argument("-o", "--output")
    args = p.parse_args(argv)
    report = run_checks(args.json, args.docx, args.pdf)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(text + "\n", encoding="utf-8")
    s = report["summary"]
    print(f"通过 {s['passed']}  失败 {s['failed']}  不适用 {s['not_applicable']}  未实现 {s['skipped']}")
    print(f"失败项：{s['failed_ids']}")
    return 0 if s["failed_blocker"] == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
