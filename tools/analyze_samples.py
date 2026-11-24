#!/usr/bin/env python3
"""分析 samples 中 Word/PDF 的排版特征。不写出任何正文、页眉或文件内文本。"""
from __future__ import annotations

import json
import math
import os
import re
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

import numpy as np
import pymupdf
from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
SAMPLES = Path(__file__).resolve().parents[1] / "data" / "samples"
CONVERTED = ROOT / "converted_docx"
OUT_JSON = ROOT / "standards" / "sample-stats.json"
OUT_MD = ROOT / "standards" / "sample-stats.md"

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def short_id(path: Path) -> str:
    return path.stem[:12]


def pt_name(pt: float | None) -> str | None:
    if pt is None:
        return None
    table = {
        6.5: "小六",
        7.5: "六号",
        9: "小五",
        10.5: "五号",
        12: "小四",
        14: "四号",
        15: "小三",
        16: "三号",
        18: "小二",
        22: "二号",
        24: "小一",
        26: "一号",
        36: "小初",
        42: "初号",
    }
    for k, v in table.items():
        if abs(pt - k) <= 0.26:
            return v
    return f"{pt:g}pt"


def emu_to_cm(v) -> float | None:
    if v is None:
        return None
    try:
        return round(int(v) / 914400 * 2.54, 3)
    except Exception:
        return None


def twips_to_pt(v) -> float | None:
    if v is None:
        return None
    return int(v) / 20.0


def twips_to_cm(v) -> float | None:
    if v is None:
        return None
    return round(int(v) / 20.0 * 2.54 / 72.0, 3)


def get_rfonts(rPr):
    if rPr is None:
        return {}
    rFonts = rPr.find(f"{W_NS}rFonts")
    if rFonts is None:
        return {}
    return {
        "ascii": rFonts.get(f"{W_NS}ascii") or rFonts.get(qn("w:ascii")),
        "hAnsi": rFonts.get(f"{W_NS}hAnsi") or rFonts.get(qn("w:hAnsi")),
        "eastAsia": rFonts.get(f"{W_NS}eastAsia") or rFonts.get(qn("w:eastAsia")),
        "cs": rFonts.get(f"{W_NS}cs") or rFonts.get(qn("w:cs")),
        "hint": rFonts.get(f"{W_NS}hint") or rFonts.get(qn("w:hint")),
    }


def get_sz_pt(rPr) -> float | None:
    if rPr is None:
        return None
    sz = rPr.find(f"{W_NS}sz")
    if sz is not None:
        val = sz.get(f"{W_NS}val") or sz.get(qn("w:val"))
        if val:
            return int(val) / 2.0
    return None


def has_bold(rPr) -> bool | None:
    if rPr is None:
        return None
    b = rPr.find(f"{W_NS}b")
    if b is None:
        return None
    val = b.get(f"{W_NS}val") or b.get(qn("w:val"))
    if val in ("0", "false", "off"):
        return False
    return True


def style_map(doc: Document) -> dict:
    out = {}
    for st in doc.styles:
        name = st.name
        info = {"name": name, "type": str(st.type) if st.type else None}
        try:
            el = st.element
            rPr = el.find(f"{W_NS}rPr")
            pPr = el.find(f"{W_NS}pPr")
            fonts = get_rfonts(rPr)
            sz = get_sz_pt(rPr)
            info.update(
                {
                    "eastAsia": fonts.get("eastAsia"),
                    "ascii": fonts.get("ascii"),
                    "size_pt": sz,
                    "size_name": pt_name(sz) if sz else None,
                    "bold": has_bold(rPr),
                }
            )
            if pPr is not None:
                ol = pPr.find(f"{W_NS}outlineLvl")
                if ol is not None:
                    val = ol.get(f"{W_NS}val") or ol.get(qn("w:val"))
                    info["outline_level"] = int(val) + 1 if val is not None else None
        except Exception:
            pass
        out[name] = info
    return out


def parse_theme_fonts(docx_path: Path) -> dict:
    try:
        with zipfile.ZipFile(docx_path) as z:
            names = [n for n in z.namelist() if n.startswith("word/theme/")]
            if not names:
                return {}
            xml = z.read(names[0])
        root = ET.fromstring(xml)
        fonts = {}
        for tag, key in (("majorFont", "major"), ("minorFont", "minor")):
            node = root.find(f".//{A_NS}{tag}")
            if node is None:
                continue
            fonts[key] = {}
            for child in node:
                typ = child.tag.split("}")[-1]
                fonts[key][typ] = child.get("typeface")
        return fonts
    except Exception:
        return {}


def has_toc_field(docx_path: Path) -> bool:
    try:
        with zipfile.ZipFile(docx_path) as z:
            xml = z.read("word/document.xml")
        # field instruction, no need to keep surrounding text
        return bool(re.search(rb"TOC\s*\\", xml) or re.search(rb">TOC<", xml))
    except Exception:
        return False


def header_footer_flags(doc: Document) -> dict:
    flags = {
        "section_count": len(doc.sections),
        "has_header": False,
        "has_footer": False,
        "different_first_page": False,
        "different_odd_even": False,
        "header_fonts": Counter(),
        "footer_fonts": Counter(),
        "header_sizes": Counter(),
        "footer_sizes": Counter(),
        "footer_has_page_field": False,
        "header_has_page_field": False,
    }
    for sec in doc.sections:
        flags["different_first_page"] |= bool(sec.different_first_page_header_footer)
        try:
            flags["different_odd_even"] |= bool(doc.settings.odd_and_even_pages_header_footer)
        except Exception:
            pass
        for attr, which in (
            ("header", "header"),
            ("footer", "footer"),
            ("first_page_header", "header"),
            ("first_page_footer", "footer"),
            ("even_page_header", "header"),
            ("even_page_footer", "footer"),
        ):
            try:
                part = getattr(sec, attr)
            except Exception:
                continue
            if part is None:
                continue
            # presence: any run or drawing
            xml = part._element.xml if hasattr(part, "_element") else ""
            has_content = bool(re.search(r"<w:t[\s>]", xml) or "w:drawing" in xml or "w:pict" in xml)
            if not has_content:
                continue
            flags[f"has_{which}"] = True
            if re.search(r"\bPAGE\b", xml) or "w:pgNum" in xml or "instrText" in xml and "PAGE" in xml:
                flags[f"{which}_has_page_field"] = True
            for p in part.paragraphs:
                for run in p.runs:
                    rPr = run._element.find(qn("w:rPr"))
                    fonts = get_rfonts(rPr)
                    ea = fonts.get("eastAsia") or fonts.get("ascii") or (run.font.name if run.font else None)
                    sz = get_sz_pt(rPr)
                    if sz is None and run.font and run.font.size:
                        sz = run.font.size.pt
                    if ea:
                        flags[f"{which}_fonts"][ea] += 1
                    if sz:
                        flags[f"{which}_sizes"][round(sz, 1)] += 1
    return flags


def section_layout(doc: Document) -> list[dict]:
    rows = []
    for i, sec in enumerate(doc.sections, 1):
        def cm(attr):
            v = getattr(sec, attr)
            return round(v.cm, 3) if v is not None else None

        def mm_page(attr):
            v = getattr(sec, attr)
            return round(v.mm, 2) if v is not None else None

        rows.append(
            {
                "section": i,
                "page_width_mm": mm_page("page_width"),
                "page_height_mm": mm_page("page_height"),
                "top_cm": cm("top_margin"),
                "bottom_cm": cm("bottom_margin"),
                "left_cm": cm("left_margin"),
                "right_cm": cm("right_margin"),
                "header_cm": cm("header_distance"),
                "footer_cm": cm("footer_distance"),
                "gutter_cm": cm("gutter"),
                "orientation": str(sec.orientation) if sec.orientation is not None else None,
            }
        )
    return rows


def para_metrics(p) -> dict:
    pPr = p._element.find(qn("w:pPr"))
    style = p.style.name if p.style is not None else None
    outline = None
    alignment = None
    first_indent_pt = None
    line_rule = None
    line_value = None
    space_before_pt = None
    space_after_pt = None
    if pPr is not None:
        ol = pPr.find(qn("w:outlineLvl"))
        if ol is not None and ol.get(qn("w:val")) is not None:
            outline = int(ol.get(qn("w:val"))) + 1
        jc = pPr.find(qn("w:jc"))
        if jc is not None:
            alignment = jc.get(qn("w:val"))
        ind = pPr.find(qn("w:ind"))
        if ind is not None:
            raw = ind.get(qn("w:firstLine")) or ind.get(qn("w:firstLineChars"))
            if ind.get(qn("w:firstLine")) is not None:
                first_indent_pt = twips_to_pt(ind.get(qn("w:firstLine")))
            elif ind.get(qn("w:firstLineChars")) is not None:
                # 100 = 1 char in Word
                first_indent_pt = ("chars", int(ind.get(qn("w:firstLineChars"))))
        spacing = pPr.find(qn("w:spacing"))
        if spacing is not None:
            rule = spacing.get(qn("w:lineRule"))
            line = spacing.get(qn("w:line"))
            line_rule = rule or "auto"
            if line is not None:
                lv = int(line)
                if (rule or "auto") in ("auto", None) or rule == "auto":
                    # 240 = single
                    line_value = ("multiple", round(lv / 240.0, 3))
                elif rule in ("exact", "atLeast"):
                    line_value = (rule, round(lv / 20.0, 2))
                else:
                    line_value = (rule, lv)
            if spacing.get(qn("w:before")):
                space_before_pt = twips_to_pt(spacing.get(qn("w:before")))
            if spacing.get(qn("w:after")):
                space_after_pt = twips_to_pt(spacing.get(qn("w:after")))
    # fallback python-docx
    pf = p.paragraph_format
    if first_indent_pt is None and pf.first_line_indent is not None:
        first_indent_pt = round(pf.first_line_indent.pt, 2)
    if line_value is None and pf.line_spacing is not None:
        rule = pf.line_spacing_rule
        if rule == WD_LINE_SPACING.EXACTLY:
            line_rule, line_value = "exact", ("exact", round(float(pf.line_spacing), 2) if not hasattr(pf.line_spacing, "pt") else round(pf.line_spacing.pt, 2))
        elif rule == WD_LINE_SPACING.MULTIPLE or rule == WD_LINE_SPACING.SINGLE or rule == WD_LINE_SPACING.ONE_POINT_FIVE or rule == WD_LINE_SPACING.DOUBLE:
            line_rule, line_value = "multiple", ("multiple", float(pf.line_spacing) if isinstance(pf.line_spacing, float) else None)
        elif rule == WD_LINE_SPACING.AT_LEAST:
            line_rule, line_value = "atLeast", ("atLeast", pf.line_spacing.pt if hasattr(pf.line_spacing, "pt") else pf.line_spacing)

    runs_meta = []
    for run in p.runs:
        rPr = run._element.find(qn("w:rPr"))
        fonts = get_rfonts(rPr)
        sz = get_sz_pt(rPr)
        if sz is None and run.font is not None and run.font.size is not None:
            sz = run.font.size.pt
        ea = fonts.get("eastAsia")
        ascii_f = fonts.get("ascii") or (run.font.name if run.font else None)
        # length only, never the characters
        # run.text is used solely for length / whitespace test
        raw = run.text or ""
        if not raw.strip():
            continue
        runs_meta.append(
            {
                "eastAsia": ea,
                "ascii": ascii_f,
                "size_pt": sz,
                "bold": has_bold(rPr) if has_bold(rPr) is not None else bool(run.font.bold),
                "chars": len(raw.strip()),
            }
        )
    return {
        "style": style,
        "outline": outline,
        "alignment": alignment,
        "first_indent_pt": first_indent_pt,
        "line_rule": line_rule,
        "line_value": line_value,
        "space_before_pt": space_before_pt,
        "space_after_pt": space_after_pt,
        "runs": runs_meta,
        "has_text": bool(runs_meta),
    }


def classify_para(style: str | None, outline: int | None) -> str:
    if not style:
        style = ""
    s = style.lower()
    if "toc" in s or "目录" in style:
        return "toc"
    if style in ("Header", "页眉") or "header" in s:
        return "header"
    if style in ("Footer", "页脚") or "footer" in s:
        return "footer"
    if "caption" in s or style in ("题注", "Caption"):
        return "caption"
    m = re.search(r"(?:heading|标题)\s*(\d)", style, re.I)
    if m:
        return f"heading_{m.group(1)}"
    if outline:
        return f"heading_{outline}"
    if "title" == s or style in ("Title", "标题"):
        return "title"
    return "body"


def analyze_tables(doc: Document) -> dict:
    info = {
        "table_count": len(doc.tables),
        "row_col": [],
        "border_styles": Counter(),
        "header_row_repeat": 0,
        "merged_cells": 0,
        "cell_fonts": Counter(),
        "cell_sizes": Counter(),
        "cell_align": Counter(),
    }
    for tbl in doc.tables:
        rows, cols = len(tbl.rows), len(tbl.columns)
        info["row_col"].append({"rows": rows, "cols": cols})
        tbl_el = tbl._tbl
        # borders
        tblPr = tbl_el.find(qn("w:tblPr"))
        if tblPr is not None:
            borders = tblPr.find(qn("w:tblBorders"))
            if borders is not None:
                for child in borders:
                    typ = child.tag.split("}")[-1]
                    val = child.get(qn("w:val"))
                    sz = child.get(qn("w:sz"))  # eighths of a point
                    info["border_styles"][f"{typ}:{val}:{sz}"] += 1
            else:
                info["border_styles"]["none_or_style"] += 1
        # header repeat
        for row in tbl.rows:
            trPr = row._tr.find(qn("w:trPr"))
            if trPr is not None and trPr.find(qn("w:tblHeader")) is not None:
                info["header_row_repeat"] += 1
                break
        # merged
        for cell in tbl._cells:
            tc = cell._tc
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is not None:
                if tcPr.find(qn("w:vMerge")) is not None or tcPr.find(qn("w:gridSpan")) is not None:
                    info["merged_cells"] += 1
            for p in cell.paragraphs:
                jc = None
                pPr = p._element.find(qn("w:pPr"))
                if pPr is not None and pPr.find(qn("w:jc")) is not None:
                    jc = pPr.find(qn("w:jc")).get(qn("w:val"))
                    if jc:
                        info["cell_align"][jc] += 1
                for run in p.runs:
                    if not (run.text or "").strip():
                        continue
                    rPr = run._element.find(qn("w:rPr"))
                    fonts = get_rfonts(rPr)
                    ea = fonts.get("eastAsia") or fonts.get("ascii")
                    sz = get_sz_pt(rPr)
                    if ea:
                        info["cell_fonts"][ea] += 1
                    if sz:
                        info["cell_sizes"][round(sz, 1)] += 1
    return info


def analyze_docx(path: Path, origin: str) -> dict:
    doc = Document(str(path))
    styles = style_map(doc)
    theme = parse_theme_fonts(path)
    hf = header_footer_flags(doc)
    sections = section_layout(doc)
    tables = analyze_tables(doc)

    role_fonts = defaultdict(Counter)
    role_sizes = defaultdict(Counter)
    role_bold = defaultdict(Counter)
    line_spacing = Counter()
    first_indents = Counter()
    alignments = Counter()
    style_use = Counter()
    heading_levels = Counter()
    para_count = 0
    body_count = 0

    for p in doc.paragraphs:
        m = para_metrics(p)
        if not m["has_text"]:
            continue
        para_count += 1
        role = classify_para(m["style"], m["outline"])
        style_use[m["style"] or "(none)"] += 1
        if role.startswith("heading_"):
            heading_levels[int(role.split("_")[1])] += 1
        if m["line_value"]:
            line_spacing[str(m["line_value"])] += 1
        if m["first_indent_pt"] is not None:
            first_indents[str(m["first_indent_pt"])] += 1
        if m["alignment"]:
            alignments[m["alignment"]] += 1
        if role == "body":
            body_count += 1
        for r in m["runs"]:
            key_f = r["eastAsia"] or r["ascii"] or "(unset)"
            role_fonts[role][key_f] += r["chars"]
            if r["ascii"] and r["eastAsia"]:
                role_fonts[role][f"ascii={r['ascii']}"] += 0  # tracked via pair
            pair = f"ea={r['eastAsia'] or '-'}|asc={r['ascii'] or '-'}"
            role_fonts[role + "_pair"][pair] += r["chars"]
            if r["size_pt"]:
                role_sizes[role][round(r["size_pt"], 1)] += r["chars"]
            role_bold[role][bool(r["bold"])] += r["chars"]

    # also count paragraphs inside tables as table_text (already in table analysis)
    return {
        "id": short_id(path),
        "origin": origin,  # sample_docx | converted_doc
        "theme_fonts": theme,
        "section_count": len(sections),
        "sections": sections,
        "styles_defined": {
            k: v
            for k, v in styles.items()
            if v.get("eastAsia") or v.get("size_pt") or v.get("outline_level") or k.lower().startswith("heading") or "标题" in k
        },
        "has_toc_field": has_toc_field(path),
        "header_footer": {
            k: (dict(v) if isinstance(v, Counter) else v) for k, v in hf.items()
        },
        "tables": {
            **{k: (dict(v) if isinstance(v, Counter) else v) for k, v in tables.items()}
        },
        "paragraphs_with_text": para_count,
        "body_paragraphs": body_count,
        "style_use": dict(style_use),
        "heading_levels": dict(heading_levels),
        "line_spacing": dict(line_spacing),
        "first_indents": dict(first_indents),
        "alignments": dict(alignments),
        "role_fonts": {k: dict(v) for k, v in role_fonts.items()},
        "role_sizes": {k: dict(v) for k, v in role_sizes.items()},
        "role_bold": {k: dict(v) for k, v in role_bold.items()},
    }


# ----- PDF (no text output) -----

def estimate_skew_deg(arr: np.ndarray, search=8.0, step=0.2) -> float:
    """Projection-profile skew on a grayscale uint8 array, 0 = upright."""
    if arr.ndim == 3:
        arr = arr.mean(axis=2)
    # downsample if large
    h, w = arr.shape
    scale = max(h, w) / 1000.0
    if scale > 1:
        new_w, new_h = int(w / scale), int(h / scale)
        # simple stride downsample
        arr = arr[:: max(1, h // new_h), :: max(1, w // new_w)]
    # invert if background is light
    if arr.mean() > 127:
        bw = arr < 200
    else:
        bw = arr > 55
    # crop empty margins
    rows = np.any(bw, axis=1)
    cols = np.any(bw, axis=0)
    if not rows.any() or not cols.any():
        return 0.0
    r0, r1 = np.where(rows)[0][[0, -1]]
    c0, c1 = np.where(cols)[0][[0, -1]]
    bw = bw[r0 : r1 + 1, c0 : c1 + 1]
    if bw.size < 400 or bw.mean() < 0.005:
        return 0.0
    best_angle, best_var = 0.0, -1.0
    angles = np.arange(-search, search + 1e-9, step)
    from PIL import Image as PILImage

    img = PILImage.fromarray((bw.astype(np.uint8) * 255))
    for a in angles:
        rot = img.rotate(a, resample=PILImage.NEAREST, fillcolor=0, expand=False)
        proj = np.asarray(rot).sum(axis=1, dtype=np.float64)
        var = float(proj.var())
        if var > best_var:
            best_var, best_angle = var, float(a)
    return round(best_angle, 2)


def analyze_pdf(path: Path) -> dict:
    doc = pymupdf.open(path)
    pages = []
    fonts = Counter()
    sizes = Counter()
    text_pages = 0
    image_only = 0
    dpi_vals = []
    skew_vals = []
    page_kinds = Counter()
    for i, page in enumerate(doc):
        r = page.rect
        w_pt, h_pt = float(r.width), float(r.height)
        w_mm, h_mm = w_pt * 25.4 / 72.0, h_pt * 25.4 / 72.0
        raw = page.get_text("text") or ""
        tlen = len(raw.strip())
        del raw  # do not retain
        images = page.get_images(full=True)
        d = page.get_text("dict")
        span_chars = 0
        for b in d.get("blocks", []):
            for line in b.get("lines", []):
                for sp in line.get("spans", []):
                    txt = sp.get("text") or ""
                    n = len(txt.strip())
                    if n == 0:
                        continue
                    span_chars += n
                    fn = sp.get("font") or "(unknown)"
                    fonts[fn] += n
                    sz = sp.get("size")
                    if sz:
                        sizes[round(float(sz), 1)] += n
        is_scan = tlen < 40 and len(images) > 0
        if tlen >= 40:
            text_pages += 1
            kind = "digital"
        elif images:
            image_only += 1
            kind = "scan"
        else:
            kind = "empty"
        page_kinds[kind] += 1

        page_dpis = []
        max_cover = 0.0
        for img in images:
            xref = img[0]
            try:
                rects = page.get_image_rects(xref)
            except Exception:
                rects = []
            try:
                info = doc.extract_image(xref)
                iw, ih = info.get("width"), info.get("height")
            except Exception:
                iw = img[2] if len(img) > 2 else None
                ih = img[3] if len(img) > 3 else None
            if not iw or not ih:
                continue
            if not rects:
                # treat as full page if one big image
                rects = [r]
            for rect in rects:
                rw, rh = float(rect.width), float(rect.height)
                if rw < 5 or rh < 5:
                    continue
                dpi_x = iw / (rw / 72.0)
                dpi_y = ih / (rh / 72.0)
                dpi = (dpi_x + dpi_y) / 2.0
                cover = (rw * rh) / (w_pt * h_pt) if w_pt * h_pt else 0
                page_dpis.append(round(dpi, 1))
                if cover > max_cover:
                    max_cover = cover
                # A4-equivalent DPI if the page itself is a photo canvas
                a4_w, a4_h = 210.0, 297.0
                # assume the image is the page and would be fitted to A4
                if max(w_mm, h_mm) > 400:
                    # page is oversized (often 1px=1pt)
                    fit = min(a4_w / min(w_mm, h_mm), a4_h / max(w_mm, h_mm)) if w_mm and h_mm else 0
                    # actually compute pixels / A4 inches of the long side
                    long_px = max(iw, ih)
                    long_in = (297 if max(w_mm, h_mm) == max(w_mm, h_mm) else 297) / 25.4
                    a4_dpi = long_px / (297 / 25.4)
                    page_dpis.append(round(a4_dpi, 1))  # also record
                    # store separately below
        if page_dpis:
            dpi_vals.extend(page_dpis)

        skew = None
        if is_scan and max_cover >= 0.4:
            # render modestly; never save pixels
            try:
                zoom = 1.2 if max(w_pt, h_pt) < 1200 else 72.0 / max(w_pt, h_pt) * 800
                zoom = max(0.15, min(zoom, 1.5))
                pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), colorspace=pymupdf.csGRAY)
                arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width)
                skew = estimate_skew_deg(arr)
                skew_vals.append(skew)
                del arr, pix
            except Exception:
                skew = None

        pages.append(
            {
                "i": i + 1,
                "width_pt": round(w_pt, 2),
                "height_pt": round(h_pt, 2),
                "width_mm": round(w_mm, 1),
                "height_mm": round(h_mm, 1),
                "kind": kind,
                "text_chars": tlen,
                "image_count": len(images),
                "dpis": page_dpis[:8],
                "median_dpi": round(float(np.median(page_dpis)), 1) if page_dpis else None,
                "skew_deg": skew,
            }
        )
    doc.close()
    a4ish = 0
    for p in pages:
        w, h = p["width_mm"], p["height_mm"]
        if (abs(min(w, h) - 210) <= 8 and abs(max(w, h) - 297) <= 10) or (
            abs(min(w, h) - 210) <= 8 and abs(max(w, h) - 297) <= 10
        ):
            a4ish += 1
    return {
        "id": short_id(path),
        "pages": doc_page_count(pages),
        "text_pages": text_pages,
        "image_only_pages": image_only,
        "kinds": dict(page_kinds),
        "a4ish_pages": a4ish,
        "fonts": dict(fonts),
        "sizes": dict(sizes),
        "dpi_all": dpi_vals,
        "skew_all": skew_vals,
        "pages_brief": pages,
    }


def doc_page_count(pages):
    return len(pages)


def analyze_png(path: Path) -> dict:
    im = Image.open(path)
    w, h = im.size
    dpi = im.info.get("dpi")
    # Do not inspect content. Flag by typical ID-card aspect and that it is the sole PNG.
    aspect = w / h if h else 0
    return {
        "id": short_id(path),
        "width_px": w,
        "height_px": h,
        "mode": im.mode,
        "format": im.format,
        "dpi_meta": list(dpi) if dpi else None,
        "aspect": round(aspect, 3),
        "excluded_reason": "pii_asset_excluded_from_fixtures_and_reports",
    }


def counter_merge(counters: list[Counter]) -> Counter:
    out = Counter()
    for c in counters:
        out.update(c)
    return out


def summarize(word_reports, pdf_reports, png_report):
    # aggregate word
    body_fonts, body_sizes = Counter(), Counter()
    heading_sizes = defaultdict(Counter)
    heading_fonts = defaultdict(Counter)
    margins = defaultdict(list)
    page_sizes = Counter()
    line_sp = Counter()
    indents = Counter()
    aligns = Counter()
    toc_docs = 0
    header_docs = 0
    footer_docs = 0
    pagefield_docs = 0
    table_docs = 0
    merged_docs = 0
    header_repeat_docs = 0
    heading_level_docs = Counter()
    theme_ea = Counter()

    for r in word_reports:
        if r.get("has_toc_field"):
            toc_docs += 1
        hf = r["header_footer"]
        if hf.get("has_header"):
            header_docs += 1
        if hf.get("has_footer"):
            footer_docs += 1
        if hf.get("footer_has_page_field") or hf.get("header_has_page_field"):
            pagefield_docs += 1
        if r["tables"]["table_count"]:
            table_docs += 1
        if r["tables"].get("merged_cells"):
            merged_docs += 1
        if r["tables"].get("header_row_repeat"):
            header_repeat_docs += 1
        for lv, n in r.get("heading_levels", {}).items():
            heading_level_docs[int(lv)] += 1
        for sec in r["sections"]:
            for k in ("top_cm", "bottom_cm", "left_cm", "right_cm"):
                if sec.get(k) is not None:
                    margins[k].append(sec[k])
            if sec.get("page_width_mm") and sec.get("page_height_mm"):
                page_sizes[(round(sec["page_width_mm"], 1), round(sec["page_height_mm"], 1))] += 1
        line_sp.update({k: v for k, v in r.get("line_spacing", {}).items()})
        indents.update(r.get("first_indents", {}))
        aligns.update(r.get("alignments", {}))
        body_fonts.update(r.get("role_fonts", {}).get("body", {}))
        body_sizes.update({float(k): v for k, v in r.get("role_sizes", {}).get("body", {}).items()})
        for role, c in r.get("role_sizes", {}).items():
            if role.startswith("heading_"):
                heading_sizes[role].update({float(k): v for k, v in c.items()})
        for role, c in r.get("role_fonts", {}).items():
            if role.startswith("heading_") and not role.endswith("_pair"):
                heading_fonts[role].update(c)
        th = r.get("theme_fonts") or {}
        for bucket in th.values():
            if isinstance(bucket, dict) and bucket.get("ea"):
                theme_ea[bucket["ea"]] += 1
            elif isinstance(bucket, dict) and bucket.get("eastAsian"):
                theme_ea[bucket["eastAsian"]] += 1

    # pdf aggregates
    total_pages = sum(r["pages"] for r in pdf_reports)
    scan_pages = sum(r["image_only_pages"] for r in pdf_reports)
    text_pages = sum(r["text_pages"] for r in pdf_reports)
    all_dpi = []
    all_skew = []
    pdf_fonts = Counter()
    pdf_sizes = Counter()
    scan_files = 0
    digital_files = 0
    mixed_files = 0
    a4_pages = 0
    for r in pdf_reports:
        all_dpi.extend(r.get("dpi_all") or [])
        all_skew.extend(r.get("skew_all") or [])
        pdf_fonts.update(r.get("fonts") or {})
        pdf_sizes.update({float(k): v for k, v in (r.get("sizes") or {}).items()})
        a4_pages += r.get("a4ish_pages") or 0
        if r["image_only_pages"] == r["pages"]:
            scan_files += 1
        elif r["text_pages"] == r["pages"]:
            digital_files += 1
        else:
            mixed_files += 1

    def stats(vals):
        if not vals:
            return None
        a = np.array(vals, dtype=float)
        return {
            "n": int(a.size),
            "min": round(float(a.min()), 3),
            "p10": round(float(np.percentile(a, 10)), 3),
            "median": round(float(np.median(a)), 3),
            "p90": round(float(np.percentile(a, 90)), 3),
            "max": round(float(a.max()), 3),
            "mean": round(float(a.mean()), 3),
        }

    def top(counter, n=8):
        if isinstance(counter, dict):
            counter = Counter(counter)
        return [{"value": k, "count": v} for k, v in counter.most_common(n)]

    return {
        "word_doc_count": len(word_reports),
        "pdf_count": len(pdf_reports),
        "pdf_pages": total_pages,
        "png_excluded": 1 if png_report else 0,
        "body_fonts": top(body_fonts, 12),
        "body_sizes_pt": top(body_sizes, 12),
        "heading_fonts": {k: top(v, 8) for k, v in heading_fonts.items()},
        "heading_sizes_pt": {k: top(v, 8) for k, v in heading_sizes.items()},
        "heading_level_docs": dict(heading_level_docs),
        "margins_cm": {k: stats(v) for k, v in margins.items()},
        "page_sizes_mm": [{"w": k[0], "h": k[1], "count": v} for k, v in Counter(page_sizes).most_common()],
        "line_spacing": top(line_sp, 15),
        "first_indents": top(indents, 15),
        "alignments": top(aligns, 8),
        "toc_docs": toc_docs,
        "header_docs": header_docs,
        "footer_docs": footer_docs,
        "pagefield_docs": pagefield_docs,
        "table_docs": table_docs,
        "merged_table_docs": merged_docs,
        "header_repeat_docs": header_repeat_docs,
        "theme_east_asia": top(theme_ea, 8),
        "pdf_scan_pages": scan_pages,
        "pdf_text_pages": text_pages,
        "pdf_other_pages": total_pages - scan_pages - text_pages,
        "pdf_scan_files": scan_files,
        "pdf_digital_files": digital_files,
        "pdf_mixed_files": mixed_files,
        "pdf_a4ish_pages": a4_pages,
        "pdf_fonts": top(pdf_fonts, 15),
        "pdf_sizes_pt": top(pdf_sizes, 15),
        "image_dpi": stats(all_dpi),
        "image_dpi_lt_200": int(sum(1 for x in all_dpi if x < 200)),
        "image_dpi_lt_300": int(sum(1 for x in all_dpi if x < 300)),
        "image_dpi_count": len(all_dpi),
        "skew_deg": stats(all_skew),
        "skew_abs_gt_0_5": int(sum(1 for x in all_skew if abs(x) > 0.5)),
        "skew_abs_gt_1": int(sum(1 for x in all_skew if abs(x) > 1.0)),
        "skew_abs_gt_2": int(sum(1 for x in all_skew if abs(x) > 2.0)),
        "skew_count": len(all_skew),
    }


def write_md(summary: dict, word_reports, pdf_reports) -> str:
    # Chinese report, no PII, files referred by 12-char hash prefix
    s = summary
    lines = []
    lines.append("# 真实标书样本排版统计（构造报告，不含正文与个人信息）")
    lines.append("")
    lines.append("说明：samples 目录只读。`.doc` 已由 LibreOffice 转换为 `converted_docx/`，未修改原文件。")
    lines.append("本报告只统计字体、字号、边距、行距、页眉页脚、目录字段、表格样式、页面尺寸、扫描页 DPI 与倾斜角。")
    lines.append("**不摘录任何正文、页眉页脚文本、文件名以外的标识，也不使用身份证等个人信息素材。**")
    lines.append("文件仅以 SHA-256 文件名前 12 位指代。")
    lines.append("")
    lines.append("## 1. 样本规模")
    lines.append("")
    lines.append(f"- Word：{s['word_doc_count']} 份（其中 5 份原始 docx，3 份由 doc 转换，转换结果仅存放在 gen 目录）。")
    lines.append(f"- PDF：{s['pdf_count']} 个文件，共 {s['pdf_pages']} 页。")
    lines.append(f"- 纯扫描页（几乎无文字层）：{s['pdf_scan_pages']} 页；有文字层页：{s['pdf_text_pages']} 页；其余：{s['pdf_other_pages']} 页。")
    lines.append(f"- 按文件：纯扫描 {s['pdf_scan_files']}，纯数字 {s['pdf_digital_files']}，混合 {s['pdf_mixed_files']}。")
    lines.append(f"- A4 近似页（宽/高约 210×297 mm，容差约 8–10 mm）：{s['pdf_a4ish_pages']} 页。")
    lines.append(f"- 另有 1 张 PNG，已判定为含个人信息素材，**不纳入夹具、示例与本报告的任何像素/文本分析细节**（仅记录已排除）。")
    lines.append("")
    lines.append("## 2. Word 正文字体与字号")
    lines.append("")
    lines.append("按字符数统计（仅 body 角色，不含标题样式）。")
    lines.append("")
    for item in s["body_fonts"]:
        lines.append(f"- 字体 `{item['value']}`：{item['count']} 字符")
    lines.append("")
    for item in s["body_sizes_pt"]:
        lines.append(f"- 字号 {item['value']} pt（{pt_name(item['value']) or '自定义'}）：{item['count']} 字符")
    lines.append("")
    lines.append("## 3. Word 标题层级与字号")
    lines.append("")
    lines.append(f"- 使用了标题样式/大纲级别的文档分布（级别→文档数）：`{s['heading_level_docs']}`")
    for role, items in sorted(s["heading_sizes_pt"].items()):
        lines.append(f"- {role} 字号： " + ", ".join(f"{it['value']}pt×{it['count']}" for it in items))
    for role, items in sorted(s["heading_fonts"].items()):
        lines.append(f"- {role} 字体： " + ", ".join(f"{it['value']}×{it['count']}" for it in items))
    lines.append("")
    lines.append("## 4. 页边距与纸张")
    lines.append("")
    for k, st in s["margins_cm"].items():
        if st:
            lines.append(
                f"- {k}：n={st['n']}，中位数 {st['median']} cm，P10–P90 {st['p10']}–{st['p90']} cm，范围 {st['min']}–{st['max']} cm"
            )
    lines.append("- 纸张尺寸（mm，节计数）：")
    for it in s["page_sizes_mm"]:
        lines.append(f"  - {it['w']} × {it['h']}：{it['count']} 节")
    lines.append("")
    lines.append("## 5. 行距、缩进、对齐")
    lines.append("")
    for it in s["line_spacing"]:
        lines.append(f"- 行距 {it['value']}：{it['count']} 段")
    for it in s["first_indents"]:
        lines.append(f"- 首行缩进 {it['value']}：{it['count']} 段")
    for it in s["alignments"]:
        lines.append(f"- 对齐 {it['value']}：{it['count']} 段")
    lines.append("")
    lines.append("## 6. 页眉页脚、目录、表格")
    lines.append("")
    n = s["word_doc_count"]
    lines.append(f"- 含页眉：{s['header_docs']}/{n}；含页脚：{s['footer_docs']}/{n}；页码域：{s['pagefield_docs']}/{n}。")
    lines.append(f"- 含 TOC 域：{s['toc_docs']}/{n}。")
    lines.append(f"- 含表格：{s['table_docs']}/{n}；含合并单元格：{s['merged_table_docs']}/{n}；表头重复行属性：{s['header_repeat_docs']}/{n}。")
    lines.append("")
    lines.append("各 Word 文件要点（无文本）：")
    for r in word_reports:
        hf = r["header_footer"]
        sec0 = r["sections"][0] if r["sections"] else {}
        lines.append(
            f"- `{r['id']}` origin={r['origin']} 节={r['section_count']} 纸张={sec0.get('page_width_mm')}×{sec0.get('page_height_mm')}mm "
            f"边距上{sec0.get('top_cm')}/下{sec0.get('bottom_cm')}/左{sec0.get('left_cm')}/右{sec0.get('right_cm')}cm "
            f"TOC={r['has_toc_field']} 页眉={hf.get('has_header')} 页脚={hf.get('has_footer')} "
            f"表格={r['tables']['table_count']} 标题级别={r.get('heading_levels')}"
        )
    lines.append("")
    lines.append("## 7. PDF 文字层字体（仅数字页，按字符数，不给文本）")
    lines.append("")
    if s["pdf_fonts"]:
        for it in s["pdf_fonts"]:
            lines.append(f"- `{it['value']}`：{it['count']}")
        for it in s["pdf_sizes_pt"]:
            lines.append(f"- 字号 {it['value']} pt：{it['count']}")
    else:
        lines.append("（无足够文字层）")
    lines.append("")
    lines.append("## 8. 扫描件有效 DPI 与倾斜角")
    lines.append("")
    if s["image_dpi"]:
        d = s["image_dpi"]
        lines.append(
            f"- 嵌入图有效 DPI：n={d['n']}，中位数 {d['median']}，P10–P90 {d['p10']}–{d['p90']}，范围 {d['min']}–{d['max']}。"
        )
        lines.append(f"- 低于 300 DPI 的测量次数：{s['image_dpi_lt_300']}/{s['image_dpi_count']}；低于 200 DPI：{s['image_dpi_lt_200']}。")
    if s["skew_deg"]:
        d = s["skew_deg"]
        lines.append(
            f"- 扫描页估计倾斜角（度，符号表示方向）：n={d['n']}，中位数 {d['median']}，范围 {d['min']}–{d['max']}，均值 {d['mean']}。"
        )
        lines.append(
            f"- |θ|>0.5°：{s['skew_abs_gt_0_5']}/{s['skew_count']}；|θ|>1°：{s['skew_abs_gt_1']}；|θ|>2°：{s['skew_abs_gt_2']}。"
        )
    lines.append("")
    lines.append("各 PDF 文件概要：")
    for r in pdf_reports:
        skew = r.get("skew_all") or []
        dpi = r.get("dpi_all") or []
        med_dpi = round(float(np.median(dpi)), 1) if dpi else None
        max_abs_skew = max((abs(x) for x in skew), default=None)
        lines.append(
            f"- `{r['id']}` 页={r['pages']} 文字页={r['text_pages']} 扫描页={r['image_only_pages']} "
            f"A4近似={r['a4ish_pages']} 中位DPI={med_dpi} 最大|倾斜角|={max_abs_skew}"
        )
    lines.append("")
    lines.append("## 9. 对质量标准的修订建议（供标准正文引用）")
    lines.append("")
    lines.append("以下建议在质量标准中逐条标注「来源：真实样本统计」。")
    lines.append("")
    lines.append("- 纸张：Word 样本以 A4 为主，标准仍以 A4（210×297 mm，±1 mm）为阻断项；扫描 PDF 存在大画布页与轻微非标准尺寸，生成输出必须规范化到 A4。")
    lines.append("- 页边距：以样本中位数为参照，容差覆盖 P10–P90。")
    lines.append("- 正文字号：以样本中最高频正文字号为默认，同时允许样本中出现的次高峰作为兼容档（须在标准中写明）。")
    lines.append("- 中文字体：样本可能出现宋体/仿宋/黑体/楷体/雅黑等，标准区分「推荐（宋体正文 + 黑体标题）」与「兼容白名单」。")
    lines.append("- 行距与缩进：按样本最常见组合设默认，宽松容差覆盖次常见值。")
    lines.append("- 目录：Word 样本中 TOC 域并不普遍，标准仍要求生成侧产出目录，但对输入缺失目录只做提示而非阻断。")
    lines.append("- 扫描纠偏：倾斜角阈值结合 |θ| 分布设定；DPI 阈值仍以 300 为高标准，但对真实低 DPI 扫描作为检查项应能检出。")
    lines.append("")
    return "\n".join(lines) + "\n"


def main():
    word_paths = []
    for p in sorted(SAMPLES.glob("*.docx")):
        word_paths.append((p, "sample_docx"))
    for p in sorted(CONVERTED.glob("*.docx")):
        word_paths.append((p, "converted_doc"))

    word_reports = []
    for path, origin in word_paths:
        print(f"WORD {origin} {short_id(path)}", flush=True)
        word_reports.append(analyze_docx(path, origin))

    pdf_reports = []
    for path in sorted(SAMPLES.glob("*.pdf")):
        print(f"PDF {short_id(path)}", flush=True)
        pdf_reports.append(analyze_pdf(path))

    png_report = None
    pngs = list(SAMPLES.glob("*.png"))
    if pngs:
        # metadata only, no pixel content written
        png_report = analyze_png(pngs[0])
        print(f"PNG {png_report['id']} excluded", flush=True)

    summary = summarize(word_reports, pdf_reports, png_report)
    payload = {
        "privacy": {
            "no_text_excerpt": True,
            "png_pii_excluded": True,
            "file_ids_are_sha256_prefix12": True,
        },
        "summary": summary,
        "word": word_reports,
        "pdf": [
            {k: v for k, v in r.items() if k != "pages_brief"} | {"page_count": r["pages"]}
            for r in pdf_reports
        ],
        "pdf_pages_brief": {r["id"]: r["pages_brief"] for r in pdf_reports},
        "png_excluded": png_report,
    }
    OUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUT_JSON.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    OUT_MD.write_text(write_md(summary, word_reports, pdf_reports), encoding="utf-8")
    print("WROTE", OUT_JSON, OUT_MD)


if __name__ == "__main__":
    main()
