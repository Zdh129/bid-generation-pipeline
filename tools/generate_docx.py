#!/usr/bin/env python3
"""将中间格式 JSON 生成标书 DOCX，并可选用 LibreOffice 转为 PDF。"""
from __future__ import annotations

import argparse
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_COLOR_INDEX, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Mm, Pt, Twips
from docx.oxml.ns import nsmap
from lxml import etree
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
import layout_spec as spec  # noqa: E402

# python-docx uses these
WD_ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "both": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "start": WD_ALIGN_PARAGRAPH.LEFT,
}


def _rpr(element):
    rPr = element.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        element.insert(0, rPr)
    return rPr


def set_run_font(run, ea=spec.FONT_BODY_EA, ascii_name=spec.FONT_ASCII, size_pt=None, bold=None, color=None):
    if size_pt is not None:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold
    run.font.name = ascii_name
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), ascii_name)
    rFonts.set(qn("w:hAnsi"), ascii_name)
    rFonts.set(qn("w:eastAsia"), ea)
    rFonts.set(qn("w:cs"), ascii_name)
    if color:
        run.font.color.rgb = None
        c = color[1:] if color.startswith("#") else color
        color_el = rPr.find(qn("w:color"))
        if color_el is None:
            color_el = OxmlElement("w:color")
            rPr.append(color_el)
        color_el.set(qn("w:val"), c)


def set_style_font(style, ea, ascii_name, size_pt, bold=False):
    style.font.name = ascii_name
    style.font.size = Pt(size_pt)
    style.font.bold = bold
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), ascii_name)
    rFonts.set(qn("w:hAnsi"), ascii_name)
    rFonts.set(qn("w:eastAsia"), ea)
    rFonts.set(qn("w:cs"), ascii_name)


def set_paragraph_east_asia(paragraph, ea, ascii_name):
    """Ensure paragraph mark also has fonts (for empty runs)."""
    pPr = paragraph._p.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), ea)
    rFonts.set(qn("w:ascii"), ascii_name)
    rFonts.set(qn("w:hAnsi"), ascii_name)


def _set_doc_defaults(doc: Document):
    # Normal
    normal = doc.styles["Normal"]
    set_style_font(normal, spec.FONT_BODY_EA, spec.FONT_ASCII, spec.BODY_SIZE_PT, False)
    pf = normal.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(spec.LINE_SPACING_PT)
    pf.first_line_indent = Pt(spec.first_indent_pt())
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.widow_control = True

    heading_cfg = {
        1: (spec.H1_SIZE_PT, spec.H1_SPACE_BEFORE_PT, spec.H1_SPACE_AFTER_PT),
        2: (spec.H2_SIZE_PT, spec.H2_SPACE_BEFORE_PT, spec.H2_SPACE_AFTER_PT),
        3: (spec.H3_SIZE_PT, spec.H3_SPACE_BEFORE_PT, spec.H3_SPACE_AFTER_PT),
        4: (spec.BODY_SIZE_PT, 6, 3),
        5: (spec.BODY_SIZE_PT, 6, 3),
        6: (spec.BODY_SIZE_PT, 6, 3),
    }
    for lv, (sz, before, after) in heading_cfg.items():
        st = doc.styles[f"Heading {lv}"]
        ea = spec.FONT_HEADING_EA if lv <= 3 else spec.FONT_BODY_EA
        set_style_font(st, ea, spec.FONT_ASCII, sz, True)
        st.paragraph_format.first_line_indent = Pt(0)
        st.paragraph_format.space_before = Pt(before)
        st.paragraph_format.space_after = Pt(after)
        st.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        st.paragraph_format.line_spacing = 1.0
        st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        st.paragraph_format.keep_with_next = True
        st.paragraph_format.widow_control = True
        # outline level already set by built-in styles

    for name, sz in (("Header", spec.HEADER_FOOTER_SIZE_PT), ("Footer", spec.HEADER_FOOTER_SIZE_PT)):
        try:
            st = doc.styles[name]
            set_style_font(st, spec.FONT_BODY_EA, spec.FONT_ASCII, sz, False)
            st.paragraph_format.first_line_indent = Pt(0)
            st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        except KeyError:
            pass

    # document default rPr in styles.xml
    styles = doc.styles.element
    doc_defaults = styles.find(qn("w:docDefaults"))
    if doc_defaults is None:
        doc_defaults = OxmlElement("w:docDefaults")
        styles.insert(0, doc_defaults)
    rpr_default = doc_defaults.find(qn("w:rPrDefault"))
    if rpr_default is None:
        rpr_default = OxmlElement("w:rPrDefault")
        doc_defaults.append(rpr_default)
    rPr = rpr_default.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        rpr_default.append(rPr)
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:ascii"), spec.FONT_ASCII)
    rFonts.set(qn("w:hAnsi"), spec.FONT_ASCII)
    rFonts.set(qn("w:eastAsia"), spec.FONT_BODY_EA)
    rFonts.set(qn("w:cs"), spec.FONT_ASCII)
    sz = rPr.find(qn("w:sz"))
    if sz is None:
        sz = OxmlElement("w:sz")
        rPr.append(sz)
    sz.set(qn("w:val"), str(int(spec.BODY_SIZE_PT * 2)))
    szCs = rPr.find(qn("w:szCs"))
    if szCs is None:
        szCs = OxmlElement("w:szCs")
        rPr.append(szCs)
    szCs.set(qn("w:val"), str(int(spec.BODY_SIZE_PT * 2)))

    # widow control in settings
    try:
        settings = doc.settings.element
        wc = settings.find(qn("w:widowControl"))
        if wc is None:
            wc = OxmlElement("w:widowControl")
            settings.append(wc)
        wc.set(qn("w:val"), "true")
    except Exception:
        pass


def _setup_section(doc: Document):
    sec = doc.sections[0]
    sec.page_width = Mm(spec.PAGE_WIDTH_MM)
    sec.page_height = Mm(spec.PAGE_HEIGHT_MM)
    # 边距取标准允许上限，使正文离开页眉/页脚检查带（页高 10%）并保留 ≥4mm 间隙
    # 同时仍落在 checks.json 允许的 1.50–3.30cm 范围内。
    edge = min(spec.MARGIN_MAX_CM, 3.20)
    sec.top_margin = Cm(edge)
    sec.bottom_margin = Cm(edge)
    sec.left_margin = Cm(spec.MARGIN_LEFT_CM)
    sec.right_margin = Cm(spec.MARGIN_RIGHT_CM)
    sec.header_distance = Cm(1.20)
    sec.footer_distance = Cm(1.20)
    sec.different_first_page_header_footer = False
    return sec


def _add_page_field(paragraph):
    """Insert '— {PAGE} —' using Word fields."""
    run1 = paragraph.add_run("— ")
    set_run_font(run1, size_pt=spec.HEADER_FOOTER_SIZE_PT, bold=False)
    run = paragraph.add_run()
    set_run_font(run, size_pt=spec.HEADER_FOOTER_SIZE_PT, bold=False)
    r = run._r
    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "end")
    r.append(fld1)
    r.append(instr)
    r.append(fld2)
    run3 = paragraph.add_run(" —")
    set_run_font(run3, size_pt=spec.HEADER_FOOTER_SIZE_PT, bold=False)


def _setup_header_footer(doc: Document, header_text: str):
    sec = doc.sections[0]
    header = sec.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    hp.text = ""
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.paragraph_format.first_line_indent = Pt(0)
    hp.paragraph_format.space_before = Pt(0)
    hp.paragraph_format.space_after = Pt(0)
    run = hp.add_run(header_text or "投标文件")
    set_run_font(run, size_pt=spec.HEADER_FOOTER_SIZE_PT, bold=False)
    set_paragraph_east_asia(hp, spec.FONT_BODY_EA, spec.FONT_ASCII)

    footer = sec.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    fp.text = ""
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.first_line_indent = Pt(0)
    fp.paragraph_format.space_before = Pt(0)
    fp.paragraph_format.space_after = Pt(0)
    _add_page_field(fp)
    set_paragraph_east_asia(fp, spec.FONT_BODY_EA, spec.FONT_ASCII)


def _apply_para_format(p, *, align=None, first_indent=None, space_before=None, space_after=None,
                       line_exact=None, keep_with_next=None, style=None):
    if style:
        p.style = style
    if align:
        p.alignment = WD_ALIGN.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    pf = p.paragraph_format
    if first_indent is not None:
        pf.first_line_indent = Pt(first_indent)
    if space_before is not None:
        pf.space_before = Pt(space_before)
    if space_after is not None:
        pf.space_after = Pt(space_after)
    if line_exact is not None:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(line_exact)
    if keep_with_next is not None:
        pf.keep_with_next = keep_with_next
    pf.widow_control = True


def _add_text_paragraph(doc, text, *, style=None, ea=spec.FONT_BODY_EA, size_pt=spec.BODY_SIZE_PT,
                        bold=False, align="justify", first_indent=None, space_before=None,
                        space_after=None, line_exact=spec.LINE_SPACING_PT, keep_with_next=None):
    p = doc.add_paragraph()
    if style:
        p.style = style
    _apply_para_format(
        p, align=align, first_indent=first_indent, space_before=space_before,
        space_after=space_after, line_exact=line_exact if style is None or style == "Normal" else None,
        keep_with_next=keep_with_next,
    )
    # heading styles already have spacing; don't force 25pt exact on headings
    if style and str(style).startswith("Heading"):
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.paragraph_format.first_line_indent = Pt(0)
    run = p.add_run(text or "")
    set_run_font(run, ea=ea, size_pt=size_pt, bold=bold)
    set_paragraph_east_asia(p, ea, spec.FONT_ASCII)
    return p


def _is_review(block: dict) -> bool:
    return block.get("verification_status") == "needs_review" or bool(block.get("needs_review"))


def _mark_review_run(run):
    try:
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    except Exception:
        pass


def _append_review_mark(p, ea=spec.FONT_BODY_EA, size_pt=spec.BODY_SIZE_PT):
    r = p.add_run(" 〔需复核〕")
    set_run_font(r, ea=ea, size_pt=size_pt, bold=False)
    _mark_review_run(r)
    return r


def _looks_numbered(text: str) -> bool:
    import re
    return bool(re.match(r"^(\d+[\.．、\)]|（\d+）|\(\d+\)|[•·\-—]|第[一二三四五六七八九十\d]+)", (text or "").strip()))


def _page_map(data: dict) -> dict:
    return {p.get("page_number"): p for p in (data.get("pages") or [])}


def _page_size_mm(page: dict) -> tuple[float, float]:
    """优先用 width_mm/height_mm；否则由 width/height 与 unit、render_dpi 换算。"""
    if not page:
        return (spec.PAGE_WIDTH_MM, spec.PAGE_HEIGHT_MM)
    if page.get("width_mm") and page.get("height_mm"):
        return float(page["width_mm"]), float(page["height_mm"])
    w, h = float(page.get("width") or 0), float(page.get("height") or 0)
    unit = page.get("unit") or "pt"
    dpi = float(page.get("render_dpi") or 72)
    if unit == "px" and dpi > 0:
        w, h = w * 25.4 / dpi, h * 25.4 / dpi
    else:
        w, h = w * 25.4 / 72.0, h * 25.4 / 72.0
    return w, h


def _page_flow_scale(page: dict | None, content_width_pt: float) -> float:
    """非 A4 大幅面页（如页尺寸设错的高分扫描件）按 A4 版心等比缩放（contain）。"""
    if not page:
        return 1.0
    w_mm, h_mm = _page_size_mm(page)
    if w_mm <= 0 or h_mm <= 0:
        return 1.0
    a4_w, a4_h = spec.PAGE_WIDTH_MM, spec.PAGE_HEIGHT_MM
    # 接近 A4（含略短的页面）不缩放
    if w_mm <= a4_w * 1.2 and h_mm <= a4_h * 1.2:
        return 1.0
    content_w_mm = _content_width_cm() * 10.0
    # 版心高度按允许的最大上下边距估算，与 _setup_section 一致
    edge_cm = min(spec.MARGIN_MAX_CM, 3.20)
    content_h_mm = a4_h - 20.0 * edge_cm
    if content_w_mm <= 0 or content_h_mm <= 0:
        return 1.0
    scale = min(content_w_mm / w_mm, content_h_mm / h_mm)
    return scale


def _deskew_notes(pages: list) -> list[dict]:
    """读取纠偏字段。angle_deg 是矩阵实际旋转角 θ，不是纠偏前倾角。

    纠偏前倾角只看 skew_angle_deg_before，残留角看 skew_angle_deg。
    deskew_applied=true 时 bbox 已在纠偏后坐标，生成侧不再旋转（即使 angle_deg 非 0）。
    """
    notes = []
    for p in pages or []:
        tr = p.get("deskew_transform") or {}
        rec = {
            "page": p.get("page_number"),
            "skew_angle_deg": p.get("skew_angle_deg"),  # 残留角
            "deskew_applied": p.get("deskew_applied"),
            "skew_angle_deg_before": p.get("skew_angle_deg_before"),  # 纠偏前倾角
            "has_transform": bool(p.get("deskew_transform")),
            "transform_angle_deg": tr.get("angle_deg"),  # 矩阵实际旋转角 θ
            "matrix_theta_deg": None,
            "angle_matrix_ok": None,
            "action": "none",
        }
        m = tr.get("matrix")
        ang_mat = tr.get("angle_deg")
        if isinstance(m, list) and len(m) == 6:
            try:
                a, b, c, d, tx, ty = (float(x) for x in m)
                mat = [[a, b, tx], [c, d, ty]]
                theta = math.degrees(math.atan2(mat[1][0], mat[0][0]))
                rec["matrix_theta_deg"] = theta
                if ang_mat is not None:
                    diff = abs((theta - float(ang_mat) + 180.0) % 360.0 - 180.0)
                    rec["angle_matrix_diff_deg"] = diff
                    rec["angle_matrix_ok"] = diff <= 0.01
            except (TypeError, ValueError):
                pass
        residual = p.get("skew_angle_deg")
        applied = p.get("deskew_applied")
        if applied is True:
            rec["action"] = "deskew_applied_true_do_not_rotate"
        elif applied or (residual is not None and abs(float(residual)) <= 0.5):
            rec["action"] = "respect_zero_or_applied_do_not_rotate"
        elif residual is not None and abs(float(residual)) > 0.5:
            rec["action"] = "residual_skew_recorded_not_rotated"
        else:
            rec["action"] = "no_skew_field"
        notes.append(rec)
    return notes


def _content_width_cm():
    return spec.PAGE_WIDTH_MM / 10.0 - spec.MARGIN_LEFT_CM - spec.MARGIN_RIGHT_CM


def _set_cell_border(cell, sz="8", color="000000", val="single"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = tcPr.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcPr.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        el = borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            borders.append(el)
        el.set(qn("w:val"), val)
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)


def _set_table_borders(table, sz="8"):
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblPr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            borders.append(el)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")


def _mark_header_row(row):
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:tblHeader")) is None:
        el = OxmlElement("w:tblHeader")
        el.set(qn("w:val"), "true")
        trPr.append(el)


def _mark_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn("w:cantSplit")) is None:
        el = OxmlElement("w:cantSplit")
        el.set(qn("w:val"), "true")
        trPr.append(el)


def _add_table(doc, table_obj):
    nr, nc = table_obj["n_rows"], table_obj["n_cols"]
    header_n = table_obj.get("header_row_count", 1) or 0
    cells = table_obj.get("cells") or []
    grid = [[None] * nc for _ in range(nr)]
    for cell in cells:
        grid[cell["row"]][cell["col"]] = cell

    caption = table_obj.get("caption")
    if caption:
        _add_text_paragraph(
            doc, caption, ea=spec.FONT_BODY_EA, size_pt=spec.CAPTION_SIZE_PT, bold=False,
            align="center", first_indent=0, line_exact=18, space_after=6, space_before=6,
        )

    table = doc.add_table(rows=nr, cols=nc)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    _set_table_borders(table, sz="8")  # 1pt >= 0.5pt

    # merge first (only anchors present in cells list); skip overlapping anchors
    merged = set()
    occupied = {}
    skipped_overlap = 0
    unique_cells = []
    for cell in cells:
        r, c = cell["row"], cell["col"]
        rs, cs = cell.get("rowspan", 1) or 1, cell.get("colspan", 1) or 1
        clash = False
        for rr in range(r, r + rs):
            for cc in range(c, c + cs):
                if (rr, cc) in occupied:
                    clash = True
        if clash:
            skipped_overlap += 1
            continue
        for rr in range(r, r + rs):
            for cc in range(c, c + cs):
                occupied[(rr, cc)] = (r, c)
        unique_cells.append(cell)
        if rs == 1 and cs == 1:
            continue
        a = table.cell(r, c)
        b = table.cell(min(r + rs - 1, nr - 1), min(c + cs - 1, nc - 1))
        a.merge(b)
        merged.add((r, c))

    for cell in unique_cells:
        r, c = cell["row"], cell["col"]
        tc = table.cell(r, c)
        text = cell.get("text") or ""
        is_header = r < header_n
        align = cell.get("alignment") or ("center" if is_header else "left")
        valign = cell.get("vertical_alignment") or ("middle" if is_header else "top")
        tc.vertical_alignment = {
            "top": WD_CELL_VERTICAL_ALIGNMENT.TOP,
            "middle": WD_CELL_VERTICAL_ALIGNMENT.CENTER,
            "bottom": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM,
        }.get(valign, WD_CELL_VERTICAL_ALIGNMENT.TOP)
        # clear default para
        p = tc.paragraphs[0]
        p.text = ""
        p.alignment = WD_ALIGN.get(align, WD_ALIGN_PARAGRAPH.LEFT)
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        run = p.add_run(text)
        set_run_font(
            run,
            ea=spec.FONT_BODY_EA,
            size_pt=spec.TABLE_SIZE_PT,
            bold=True if is_header or cell.get("bold") else False,
        )
        _set_cell_border(tc, sz="8")

    for i, row in enumerate(table.rows):
        _mark_cant_split(row)
        if i < header_n:
            _mark_header_row(row)
    hole_count = nr * nc - len(occupied)
    return {
        "table": table,
        "n_rows": nr,
        "n_cols": nc,
        "input_cells": len(cells),
        "written_cells": len(unique_cells),
        "merged_anchors": len(merged),
        "skipped_overlap": skipped_overlap,
        "holes": hole_count,
        "header_row_count": header_n,
    }


def _add_image(doc, image_path: Path, width_pt: float | None, max_width_cm: float):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    run = p.add_run()
    if width_pt:
        width_cm = min(width_pt * 2.54 / 72.0, max_width_cm)
    else:
        width_cm = max_width_cm
    # never exceed content width
    width_cm = min(width_cm, max_width_cm)
    run.add_picture(str(image_path), width=Cm(width_cm))
    set_paragraph_east_asia(p, spec.FONT_BODY_EA, spec.FONT_ASCII)
    return p, width_cm


def _sort_blocks(blocks):
    body_types_hf = {"header", "footer", "page_number"}
    body = [b for b in blocks if b.get("type") not in body_types_hf]
    body.sort(key=lambda b: (b.get("page", 1), b.get("reading_order", 0)))
    return body


def _header_text_from(data, options=None) -> str:
    options = options or {}
    if options.get("header_text"):
        return str(options["header_text"])
    for b in data.get("blocks") or []:
        if b.get("type") == "header" and (b.get("text") or "").strip() and not _is_review(b):
            return b["text"].strip()
    for b in data.get("blocks") or []:
        if b.get("type") == "header" and (b.get("text") or "").strip():
            return b["text"].strip()
    title = (data.get("document") or {}).get("title")
    if title:
        return title.strip()
    for b in data.get("blocks") or []:
        if b.get("type") == "title" and (b.get("text") or "").strip():
            return b["text"].strip()
    return "投标文件"


def _headings(blocks):
    out = []
    for b in _sort_blocks(blocks):
        if b.get("verification_status") == "rejected":
            continue
        if b.get("type") == "heading":
            out.append({"text": (b.get("text") or "").strip(), "level": int(b.get("level") or 1)})
    return out


def _write_toc(doc, headings, page_map: dict[str, int] | None):
    _add_text_paragraph(
        doc, "目录", style="Heading 1", ea=spec.FONT_HEADING_EA, size_pt=spec.H1_SIZE_PT,
        bold=True, align="center", first_indent=0, line_exact=None, space_before=0, space_after=12,
        keep_with_next=False,
    )
    # Override heading 1 "目录" to be a toc title not a numbered chapter — still Heading 1 for structure?
    # Better: use a dedicated centered title style without outline? Checks look for 「目录」.
    # Using Heading 1 would put "目录" itself into heading list if we collect from docx.
    # We'll use Normal centered 16pt bold instead, to avoid polluting heading outline.
    # Re-do: delete that and use Normal.
    # Actually we already added Heading 1. Let me change approach in caller.
    pass


def add_toc_paragraphs(doc, headings, page_map: dict[str, int] | None):
    p_title = _add_text_paragraph(
        doc, "目录", ea=spec.FONT_HEADING_EA, size_pt=spec.H1_SIZE_PT, bold=True,
        align="center", first_indent=0, line_exact=28, space_before=0, space_after=12,
    )
    content_w = _content_width_cm()
    for h in headings:
        text = h["text"]
        level = h["level"]
        page_no = ""
        if page_map is not None:
            page_no = str(page_map.get(text, ""))
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.left_indent = Pt(12 * (level - 1))
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p.paragraph_format.line_spacing = Pt(22)
        # right tab with dots
        tab_stops = p.paragraph_format.tab_stops
        # content width in inches
        tab_stops.add_tab_stop(Cm(content_w), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        run = p.add_run(text)
        set_run_font(run, ea=spec.FONT_BODY_EA, size_pt=12, bold=False)
        run2 = p.add_run("\t" + page_no)
        set_run_font(run2, ea=spec.FONT_BODY_EA, size_pt=12, bold=False)
        set_paragraph_east_asia(p, spec.FONT_BODY_EA, spec.FONT_ASCII)
    # page break after TOC
    p_break = doc.add_paragraph()
    p_break.paragraph_format.first_line_indent = Pt(0)
    run = p_break.add_run()
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run._r.append(br)
    return p_title


def resolve_asset_path(assets_root: Path, assets: list, asset_id: str, crop_bbox=None) -> Path:
    item = next((a for a in assets if a["id"] == asset_id), None)
    if not item:
        raise FileNotFoundError(f"asset not found: {asset_id}")
    path = assets_root / item["path"]
    if not path.is_file():
        raise FileNotFoundError(path)
    if crop_bbox:
        x0, y0, x1, y1 = [int(round(v)) for v in crop_bbox]
        im = Image.open(path)
        cropped = im.crop((x0, y0, x1, y1))
        tmp = path.parent / f".crop_{asset_id}.png"
        cropped.save(tmp)
        return tmp
    return path


def image_display_width_pt(block, asset, max_width_pt):
    img = block.get("image") or {}
    if img.get("display_width_pt"):
        return min(float(img["display_width_pt"]), max_width_pt)
    dpi = float(asset.get("original_dpi") or 300)
    wpx = float(asset.get("width_px") or 0)
    if wpx > 0 and dpi > 0:
        return min(wpx / dpi * 72.0, max_width_pt)
    return max_width_pt


def render_document(data: dict, assets_root: Path, options: dict | None = None,
                    toc_page_map: dict[str, int] | None = None) -> Document:
    options = options or {}
    doc = Document()
    _set_doc_defaults(doc)
    _setup_section(doc)
    _setup_header_footer(doc, _header_text_from(data, options))

    assets = data.get("assets") or []
    asset_map = {a["id"]: a for a in assets}
    blocks = _sort_blocks(data.get("blocks") or [])
    headings = _headings(data.get("blocks") or [])
    has_toc_block = any(b.get("type") == "toc" and b.get("verification_status") != "rejected" for b in blocks)
    want_toc = has_toc_block or (options.get("force_toc") or (len(headings) >= 3 and options.get("auto_toc", True)))

    max_w_cm = _content_width_cm()
    max_w_pt = max_w_cm * 72 / 2.54

    toc_inserted = False
    render_meta = {
        "rejected_skipped": [],
        "review_emitted": [],
        "empty_skipped": [],
        "tables": [],
        "images": [],
        "seals": [],
        "page_breaks": 0,
        "deskew": _deskew_notes(data.get("pages") or []),
    }
    pages_by_no = _page_map(data)
    current_page = None

    def _maybe_page_break(page_no):
        nonlocal current_page
        if current_page is None:
            current_page = page_no
            return
        if page_no != current_page and options.get("page_breaks_between_pages", True):
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Pt(0)
            run = p.add_run()
            br = OxmlElement("w:br")
            br.set(qn("w:type"), "page")
            run._r.append(br)
            render_meta["page_breaks"] += 1
        current_page = page_no

    for b in blocks:
        if b.get("verification_status") == "rejected":
            render_meta["rejected_skipped"].append(b.get("id"))
            continue
        typ = b.get("type")
        text = b.get("text") or ""
        st = b.get("style") or {}
        review = _is_review(b)
        if review:
            render_meta["review_emitted"].append({"id": b.get("id"), "type": typ})
        _maybe_page_break(b.get("page"))

        if typ == "toc":
            if not toc_inserted:
                add_toc_paragraphs(doc, headings, toc_page_map)
                toc_inserted = True
            continue

        if typ in ("header", "footer", "page_number"):
            continue

        if typ == "title":
            if not text.strip():
                render_meta["empty_skipped"].append(b.get("id"))
                continue
            p = _add_text_paragraph(
                doc, text, ea=spec.FONT_HEADING_EA,
                size_pt=st.get("font_size_pt") or spec.H1_SIZE_PT,
                bold=True, align=st.get("alignment") or "center", first_indent=0,
                space_before=12, space_after=10, line_exact=None,
            )
            if review:
                _append_review_mark(p)
            if want_toc and not toc_inserted:
                add_toc_paragraphs(doc, headings, toc_page_map)
                toc_inserted = True
            continue

        if typ == "heading":
            if want_toc and not toc_inserted:
                add_toc_paragraphs(doc, headings, toc_page_map)
                toc_inserted = True
            lv = int(b.get("level") or 1)
            lv = min(max(lv, 1), 6)
            sz = {1: spec.H1_SIZE_PT, 2: spec.H2_SIZE_PT, 3: spec.H3_SIZE_PT}.get(lv, spec.BODY_SIZE_PT)
            ea = spec.FONT_HEADING_EA if lv <= 3 else spec.FONT_BODY_EA
            if not text.strip():
                render_meta["empty_skipped"].append(b.get("id"))
                continue
            p = _add_text_paragraph(
                doc, text, style=f"Heading {lv}", ea=ea, size_pt=st.get("font_size_pt") or sz,
                bold=True, align=st.get("alignment") or ("center" if lv == 1 else "left"),
                first_indent=0,
                space_before=spec.H1_SPACE_BEFORE_PT,
                space_after=spec.H1_SPACE_AFTER_PT,
                keep_with_next=True, line_exact=None,
            )
            if review:
                _append_review_mark(p, size_pt=sz)
            for r in p.runs:
                r.font.bold = True
                set_run_font(r, ea=ea, size_pt=st.get("font_size_pt") or sz, bold=True)
            continue

        if typ in ("paragraph", "unknown", "handwriting"):
            if typ == "unknown" and b.get("image"):
                typ = "image"
            else:
                if not text.strip():
                    render_meta["empty_skipped"].append(b.get("id"))
                    continue
                align = st.get("alignment") or "left"
                indent = st.get("first_line_indent_pt")
                if indent is None:
                    indent = 0
                p = _add_text_paragraph(
                    doc, text, ea=spec.FONT_BODY_EA, size_pt=spec.BODY_SIZE_PT,
                    bold=bool(st.get("bold")), align=align, first_indent=indent,
                    line_exact=spec.LINE_SPACING_PT,
                )
                if review:
                    _append_review_mark(p)
                if typ != "image":
                    continue

        if typ == "list":
            lst = b.get("list") or {}
            items = lst.get("items") or [{"text": text}]
            for it in items:
                marker = it.get("marker")
                lv = int(it.get("level") or 1)
                body = it.get("text") or ""
                if marker:
                    line = f"{marker}{body}" if body.startswith(" ") or not marker else f"{marker} {body}"
                elif _looks_numbered(body):
                    line = body
                else:
                    prefix = "• " if lst.get("list_type") == "unordered" else "（1） "
                    line = prefix + body
                p = _add_text_paragraph(
                    doc, line, ea=spec.FONT_BODY_EA, size_pt=spec.BODY_SIZE_PT,
                    align="left", first_indent=0,
                    line_exact=spec.LINE_SPACING_PT,
                )
                p.paragraph_format.left_indent = Pt(21 * lv)
                p.paragraph_format.first_line_indent = Pt(-21)
                if review:
                    _append_review_mark(p)
            continue

        if typ == "table":
            if review:
                _add_text_paragraph(
                    doc, "（下表经识别侧标为需复核，已按 cells[] 还原合并单元格，请人工核对）",
                    ea=spec.FONT_BODY_EA, size_pt=spec.BODY_SIZE_PT, align="left", first_indent=0,
                    space_before=6, space_after=2, line_exact=spec.LINE_SPACING_PT,
                )
            stats = _add_table(doc, b.get("table") or {})
            stats["block_id"] = b.get("id")
            stats["review"] = review
            stats.pop("table", None)
            render_meta["tables"].append(stats)
            continue

        if typ in ("table_caption",):
            _add_text_paragraph(
                doc, text, ea=spec.FONT_BODY_EA, size_pt=spec.CAPTION_SIZE_PT,
                align="center", first_indent=0, line_exact=18, space_before=6, space_after=3,
            )
            continue

        if typ == "figure_caption":
            _add_text_paragraph(
                doc, text, ea=spec.FONT_BODY_EA, size_pt=spec.CAPTION_SIZE_PT,
                align="center", first_indent=0, line_exact=18, space_before=3, space_after=6,
            )
            continue

        if typ in ("image", "seal", "signature") or (typ == "formula" and b.get("image")):
            img = b.get("image") or {}
            asset_id = img.get("asset_id")
            if not asset_id:
                if text:
                    _add_text_paragraph(doc, text, align="center", first_indent=0)
                continue
            asset = asset_map[asset_id]
            path = resolve_asset_path(assets_root, assets, asset_id, img.get("crop_bbox"))
            page = pages_by_no.get(b.get("page"))
            scale = _page_flow_scale(page, max_w_pt)
            if typ == "seal" or asset.get("type") == "seal":
                raw_w = img.get("display_width_pt") or (3.0 * 72 / 2.54)
                width_pt = float(raw_w) * scale
                kind = "seals"
            else:
                width_pt = image_display_width_pt(b, asset, max_w_pt) * scale
                kind = "images"
            _, used_cm = _add_image(doc, path, width_pt, max_w_cm)
            render_meta[kind].append({
                "block_id": b.get("id"),
                "asset_id": asset_id,
                "page": b.get("page"),
                "page_scale": round(scale, 4),
                "width_pt_requested": round(float(width_pt), 2),
                "width_cm_used": round(used_cm, 2),
                "review": review,
                "deskew_applied_on_page": (page or {}).get("deskew_applied"),
                "page_skew_angle_deg": (page or {}).get("skew_angle_deg"),
            })
            if review:
                cap_r = "〔需复核：印章〕" if kind == "seals" else "〔需复核：图片〕"
                _add_text_paragraph(
                    doc, cap_r, ea=spec.FONT_BODY_EA, size_pt=spec.BODY_SIZE_PT, align="center",
                    first_indent=0, line_exact=spec.LINE_SPACING_PT, space_after=4,
                )
            cap = img.get("caption")
            if cap and typ == "image":
                _add_text_paragraph(
                    doc, cap, ea=spec.FONT_BODY_EA, size_pt=spec.CAPTION_SIZE_PT,
                    align="center", first_indent=0, line_exact=18, space_after=6,
                )
            continue

        if typ == "formula":
            _add_text_paragraph(
                doc, text, ea=spec.FONT_BODY_EA, size_pt=spec.BODY_SIZE_PT,
                align="center", first_indent=0, line_exact=spec.LINE_SPACING_PT, space_before=6, space_after=6,
            )
            continue

        if typ == "signature":
            _add_text_paragraph(
                doc, text or "（签字）", ea=spec.FONT_BODY_EA, size_pt=spec.BODY_SIZE_PT,
                align="right", first_indent=0, line_exact=spec.LINE_SPACING_PT, space_before=24,
            )
            continue

        # fallback
        if text:
            _add_text_paragraph(doc, text)

    if want_toc and not toc_inserted and headings:
        # toc at start: we already iterated; if no title/toc block, headings came first so should have inserted.
        # If document has no heading encountered before end and toc wanted, insert at beginning is too late.
        pass

    if options.get("review_appendix", True) and (render_meta["review_emitted"] or render_meta["rejected_skipped"]):
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Pt(0)
        run = p.add_run()
        br = OxmlElement("w:br")
        br.set(qn("w:type"), "page")
        run._r.append(br)
        _add_text_paragraph(
            doc, "复核清单（仅块编号与类型，不含正文）",
            ea=spec.FONT_HEADING_EA, size_pt=14, bold=True, align="center",
            first_indent=0, space_before=12, space_after=8, line_exact=20,
        )
        if render_meta["review_emitted"]:
            lines = "、".join(f"{x['id']}({x['type']})" for x in render_meta["review_emitted"])
            _add_text_paragraph(
                doc, "needs_review 已写入正文（黄底「需复核」标记）：" + lines,
                align="left", first_indent=0,
            )
        if render_meta["rejected_skipped"]:
            _add_text_paragraph(
                doc, "rejected 已跳过：" + "、".join(render_meta["rejected_skipped"]),
                align="left", first_indent=0,
            )
        if render_meta["empty_skipped"]:
            _add_text_paragraph(
                doc, "空文本块已跳过：" + "、".join(str(x) for x in render_meta["empty_skipped"]),
                align="left", first_indent=0,
            )

    doc._ocr_render_meta = render_meta

    if options.get("insert_blank_page"):
        # two page breaks create an empty page
        for _ in range(2):
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Pt(0)
            run = p.add_run()
            br = OxmlElement("w:br")
            br.set(qn("w:type"), "page")
            run._r.append(br)

    return doc


def convert_to_pdf(docx_path: Path, pdf_path: Path, timeout: int = 180) -> None:
    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    # LibreOffice 有时拒绝覆盖已有 PDF，先删除目标
    try:
        if Path(pdf_path).exists():
            Path(pdf_path).unlink()
    except Exception:
        pass
    outdir = pdf_path.parent
    # 优先走会更新 TOC 等域的宏导出
    try:
        from lo_export_pdf import convert_with_field_update
        ok, log = convert_with_field_update(Path(docx_path), Path(pdf_path), timeout=timeout)
        if ok:
            return
        print("LO_MACRO_EXPORT_FALLBACK", log[:500], flush=True)
    except Exception as exc:
        print("LO_MACRO_EXPORT_ERR", type(exc).__name__, exc, flush=True)
    profile = Path(tempfile.mkdtemp(prefix="lo_prof_"))
    cmd = [
        "soffice", "--headless", "--nologo", "--nofirststartwizard", "--norestore",
        f"-env:UserInstallation=file://{profile}",
        "--convert-to", "pdf:writer_pdf_Export",
        "--outdir", str(outdir),
        str(docx_path),
    ]
    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"
    try:
        proc = subprocess.run(cmd, check=False, timeout=timeout, env=env,
                              stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if proc.returncode != 0:
            raise RuntimeError(
                f"LibreOffice 转换失败 rc={proc.returncode} "
                f"stderr={(proc.stderr or b'')[:500]!r}"
            )
        produced = outdir / (docx_path.stem + ".pdf")
        if not produced.is_file():
            # sometimes soffice writes with slightly different name
            pdfs = list(outdir.glob("*.pdf"))
            raise FileNotFoundError(f"未生成 PDF: 期望 {produced}, 目录内 {pdfs}")
        if produced.resolve() != pdf_path.resolve():
            shutil.move(str(produced), str(pdf_path))
    finally:
        shutil.rmtree(profile, ignore_errors=True)


def collect_heading_pages(pdf_path: Path, headings: list[dict]) -> dict[str, int]:
    import pymupdf

    page_map: dict[str, int] = {}
    doc = pymupdf.open(pdf_path)
    try:
        for h in headings:
            text = h["text"]
            if not text or text in page_map:
                continue
            for i, page in enumerate(doc, 1):
                hits = page.search_for(text)
                if hits:
                    # skip if this looks like a TOC line (has dotted leaders nearby) on early pages
                    # Prefer the last occurrence if the first is on a TOC page.
                    page_map[text] = i
                    # continue to allow later occurrence overwrite
        # second pass: if a heading was only found once, keep it; if found on page 1 and later, use later
        # re-scan preferring non-TOC: TOC page typically contains "目录"
        toc_pages = set()
        for i, page in enumerate(doc, 1):
            t = page.get_text("text") or ""
            if "目录" in t[:400] or t.strip().startswith("目录"):
                toc_pages.add(i)
        for h in headings:
            text = h["text"]
            found = []
            for i, page in enumerate(doc, 1):
                if page.search_for(text):
                    found.append(i)
            body_hits = [i for i in found if i not in toc_pages]
            if body_hits:
                page_map[text] = body_hits[0]
            elif found:
                page_map[text] = found[0]
    finally:
        doc.close()
    return page_map


def generate(json_path: Path, out_docx: Path, out_pdf: Path | None = None,
             options: dict | None = None) -> dict:
    options = options or {}
    data = json.loads(json_path.read_text(encoding="utf-8"))
    assets_root = json_path.parent
    headings = _headings(data.get("blocks") or [])
    want_toc = any(b.get("type") == "toc" for b in data.get("blocks") or []) or len(headings) >= 3

    out_docx.parent.mkdir(parents=True, exist_ok=True)
    doc = render_document(data, assets_root, options, toc_page_map=None if want_toc else {})
    # if want toc, first pass uses empty page numbers
    if want_toc:
        doc = render_document(data, assets_root, options, toc_page_map={})
    doc.save(str(out_docx))

    render_meta = getattr(doc, "_ocr_render_meta", None)
    result = {"docx": str(out_docx), "pdf": None, "toc_updated": False, "render_meta": render_meta}

    if out_pdf is None:
        return result

    # first conversion
    tmp_pdf = out_pdf.with_name(out_pdf.stem + ".pass1.pdf")
    convert_to_pdf(out_docx, tmp_pdf)
    if want_toc and headings:
        page_map = collect_heading_pages(tmp_pdf, headings)
        doc = render_document(data, assets_root, options, toc_page_map=page_map)
        doc.save(str(out_docx))
        convert_to_pdf(out_docx, out_pdf)
        result["toc_updated"] = True
        try:
            tmp_pdf.unlink()
        except Exception:
            pass
    else:
        shutil.move(str(tmp_pdf), str(out_pdf))
    result["pdf"] = str(out_pdf)
    return result


def main(argv=None) -> int:
    p = argparse.ArgumentParser()
    p.add_argument("json_path")
    p.add_argument("--docx", required=True)
    p.add_argument("--pdf")
    p.add_argument("--blank-page", action="store_true")
    args = p.parse_args(argv)
    options = {"insert_blank_page": args.blank_page}
    generate(Path(args.json_path), Path(args.docx), Path(args.pdf) if args.pdf else None, options)
    print("OK", args.docx, args.pdf or "")
    return 0


if __name__ == "__main__":
    sys.exit(main())
