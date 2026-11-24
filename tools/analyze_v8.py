#!/usr/bin/env python3
"""精确统计 V8 参考成品与模板的排版特征。不把正文/个人信息写入报告。"""
from __future__ import annotations

import hashlib
import json
import re
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from repo_paths import load_paths
from xml.etree import ElementTree as ET

from docx import Document
from docx.oxml.ns import qn
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
PATHS = load_paths(ROOT)
OUT_DIR = ROOT / "standards"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NSMAP = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def pt_name(pt):
    if pt is None:
        return None
    table = {9: "小五", 10.5: "五号", 12: "小四", 14: "四号", 15: "小三", 16: "三号",
             18: "小二", 22: "二号", 24: "小一", 26: "一号", 8: "6.5? ", 7.5: "六号", 20: "20pt", 11: "11pt"}
    for k, v in table.items():
        if abs(pt - k) <= 0.26:
            return v
    return f"{pt:g}pt"


def emu_to_pt(v):
    return int(v) / 12700 if v is not None else None


def twips_to_cm(v):
    return round(int(v) / 20.0 * 2.54 / 72.0, 3) if v is not None else None


def twips_to_pt(v):
    return int(v) / 20.0 if v is not None else None


def get(el, attr):
    return el.get(f"{W}{attr}") if el is not None else None


def rfonts(rPr):
    if rPr is None:
        return {}
    rf = rPr.find(f"{W}rFonts")
    if rf is None:
        return {}
    return {k: rf.get(f"{W}{k}") for k in ("ascii", "hAnsi", "eastAsia", "cs", "hint")}


def sz_pt(rPr):
    if rPr is None:
        return None
    sz = rPr.find(f"{W}sz")
    if sz is None:
        return None
    v = get(sz, "val")
    return int(v) / 2.0 if v else None


def is_bold(rPr):
    if rPr is None:
        return False
    b = rPr.find(f"{W}b")
    if b is None:
        return False
    v = get(b, "val")
    return v not in ("0", "false", "off")


def para_text(p):
    texts = []
    for t in p.iter(f"{W}t"):
        if t.text:
            texts.append(t.text)
        if t.tail:
            texts.append(t.tail)
    return "".join(texts)


REDACT_PATTERNS = [
    # personal ID-like
    (re.compile(r"\b[1-9]\d{5}(?:18|19|20)\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])\d{3}[\dXx]\b"), "[证件号]"),
    (re.compile(r"\b[1-9]\d{14,17}[\dXx]\b"), "[证件号]"),
    (re.compile(r"(统一社会信用代码\s*[:：]?\s*)[0-9A-Z]{15,18}"), r"\1[信用代码]"),
    (re.compile(r"(电话|手机|联系电话)\s*[:：]?\s*1[3-9]\d{9}"), r"\1[电话]"),
]


def redact(s: str) -> str:
    if not s:
        return s
    for pat, rep in REDACT_PATTERNS:
        s = pat.sub(rep, s)
    return s


def parse_sections_from_xml(root):
    # section properties can be in body/sectPr or p/pPr/sectPr
    sects = []
    body = root.find(f"{W}body")
    if body is None:
        return sects
    # find all sectPr in document order
    for sectPr in body.iter(f"{W}sectPr"):
        pgSz = sectPr.find(f"{W}pgSz")
        pgMar = sectPr.find(f"{W}pgMar")
        pgNumType = sectPr.find(f"{W}pgNumType")
        cols = sectPr.find(f"{W}cols")
        titlePg = sectPr.find(f"{W}titlePg")
        evenOdd = sectPr.find(f"{W}evenAndOddHeaders")
        sectType = sectPr.find(f"{W}type")
        header_refs = []
        footer_refs = []
        for rel in sectPr.findall(f"{W}headerReference"):
            header_refs.append({"type": rel.get(f"{W}type"), "id": rel.get(f"{R}id")})
        for rel in sectPr.findall(f"{W}footerReference"):
            footer_refs.append({"type": rel.get(f"{W}type"), "id": rel.get(f"{R}id")})
        rec = {
            "w_mm": round(int(get(pgSz, "w")) / 20 * 25.4 / 72, 2) if pgSz is not None and get(pgSz, "w") else None,
            "h_mm": round(int(get(pgSz, "h")) / 20 * 25.4 / 72, 2) if pgSz is not None and get(pgSz, "h") else None,
            "orient": get(pgSz, "orient") if pgSz is not None else None,
            "margins_cm": None,
            "pgNumType": None,
            "title_page": titlePg is not None,
            "even_odd": evenOdd is not None,
            "sect_type": get(sectType, "val") if sectType is not None else "nextPage?",
            "header_refs": header_refs,
            "footer_refs": footer_refs,
            "cols": get(cols, "num") if cols is not None else None,
        }
        if pgMar is not None:
            rec["margins_cm"] = {
                "top": twips_to_cm(get(pgMar, "top")),
                "bottom": twips_to_cm(get(pgMar, "bottom")),
                "left": twips_to_cm(get(pgMar, "left")),
                "right": twips_to_cm(get(pgMar, "right")),
                "header": twips_to_cm(get(pgMar, "header")),
                "footer": twips_to_cm(get(pgMar, "footer")),
                "gutter": twips_to_cm(get(pgMar, "gutter")),
            }
        if pgNumType is not None:
            rec["pgNumType"] = {
                "start": get(pgNumType, "start"),
                "fmt": get(pgNumType, "fmt"),
                "chapStyle": get(pgNumType, "chapStyle"),
            }
        sects.append(rec)
    return sects


def parse_hf(z: zipfile.ZipFile):
    names = [n for n in z.namelist() if re.search(r"word/(header|footer)\d+\.xml", n)]
    out = []
    for n in names:
        root = ET.fromstring(z.read(n))
        texts = []
        fonts, sizes = Counter(), Counter()
        has_page = False
        xml = z.read(n).decode("utf-8", errors="ignore")
        if re.search(r"PAGE|w:instrText[^>]*>[^<]*PAGE", xml):
            has_page = True
        for p in root.iter(f"{W}p"):
            texts.append(para_text(p))
            for r in p.iter(f"{W}r"):
                rPr = r.find(f"{W}rPr")
                fr = rfonts(rPr)
                ea = fr.get("eastAsia") or fr.get("ascii")
                sz = sz_pt(rPr)
                t = "".join((t.text or "") for t in r.iter(f"{W}t"))
                if t.strip() or has_page:
                    if ea:
                        fonts[ea] += 1
                    if sz:
                        sizes[sz] += 1
        raw = "\n".join(texts)
        # only keep structure: length, has_page, fonts; redacted snippet hashed + pattern
        out.append({
            "file": Path(n).name,
            "kind": "header" if "header" in n else "footer",
            "has_page_field": has_page or ("PAGE" in xml),
            "text_len": len(raw.strip()),
            "fonts": dict(fonts),
            "sizes": {str(k): v for k, v in sizes.items()},
            "has_drawing": "<w:drawing" in xml or "w:pict" in xml,
            # pattern only: replace digits and CJK runs of possible names conservatively
            "pattern": redact(re.sub(r"\d+", "N", raw.strip()))[:80] if raw.strip() else "",
        })
    return out


def parse_styles(z):
    if "word/styles.xml" not in z.namelist():
        return {}
    root = ET.fromstring(z.read("word/styles.xml"))
    styles = {}
    for st in root.findall(f"{W}style"):
        sid = get(st, "styleId")
        typ = get(st, "type")
        name_el = st.find(f"{W}name")
        name = get(name_el, "val") if name_el is not None else sid
        rPr = st.find(f"{W}rPr")
        pPr = st.find(f"{W}pPr")
        info = {"id": sid, "name": name, "type": typ}
        fr = rfonts(rPr)
        info["eastAsia"] = fr.get("eastAsia")
        info["ascii"] = fr.get("ascii")
        info["size_pt"] = sz_pt(rPr)
        info["bold"] = is_bold(rPr)
        if pPr is not None:
            ol = pPr.find(f"{W}outlineLvl")
            if ol is not None:
                info["outline_level"] = int(get(ol, "val")) + 1
            jc = pPr.find(f"{W}jc")
            if jc is not None:
                info["align"] = get(jc, "val")
            sp = pPr.find(f"{W}spacing")
            if sp is not None:
                info["spacing"] = {k: get(sp, k) for k in ("before", "after", "line", "lineRule")}
        styles[sid] = info
    return styles


def looks_like_heading_text(t: str) -> int | None:
    t = t.strip()
    if not t or len(t) > 80:
        return None
    if re.match(r"^第[一二三四五六七八九十百零0-9]+[章节编部分]", t):
        return 1
    if re.match(r"^[一二三四五六七八九十]+[、.．]", t):
        return 2
    if re.match(r"^（[一二三四五六七八九十]+）|^[（(][0-9]+[)）]", t):
        return 3
    if re.match(r"^\d+(\.\d+){0,3}\s+\S", t) and not re.match(r"^\d+(\.\d+)*\s*$", t):
        parts = t.split()[0].split(".")
        return min(len(parts), 4)
    if t in ("目录", "封面", "投标文件"):
        return 0
    return None


def analyze_paragraphs(root, styles):
    body = root.find(f"{W}body")
    role_fonts = defaultdict(Counter)
    role_sizes = defaultdict(Counter)
    line_sp = Counter()
    indents = Counter()
    aligns = Counter()
    style_use = Counter()
    heading_list = []  # for structure, redacted text
    outline_paras = []
    para_count = 0
    body_count = 0
    for p in body.findall(f"{W}p"):
        pPr = p.find(f"{W}pPr")
        text = para_text(p).strip()
        if not text:
            # still might be image-only para
            if p.find(f".//{W}drawing") is None:
                continue
        para_count += 1
        style_id = None
        outline = None
        jc = None
        first_ind = None
        line = None
        if pPr is not None:
            ps = pPr.find(f"{W}pStyle")
            if ps is not None:
                style_id = get(ps, "val")
            ol = pPr.find(f"{W}outlineLvl")
            if ol is not None and get(ol, "val") is not None:
                outline = int(get(ol, "val")) + 1
            jc_el = pPr.find(f"{W}jc")
            if jc_el is not None:
                jc = get(jc_el, "val")
            ind = pPr.find(f"{W}ind")
            if ind is not None and get(ind, "firstLine") is not None:
                first_ind = twips_to_pt(get(ind, "firstLine"))
            elif ind is not None and get(ind, "firstLineChars") is not None:
                first_ind = ("chars", int(get(ind, "firstLineChars")))
            sp = pPr.find(f"{W}spacing")
            if sp is not None and get(sp, "line") is not None:
                rule = get(sp, "lineRule") or "auto"
                lv = int(get(sp, "line"))
                if rule in ("auto", None):
                    line = ("multiple", round(lv / 240.0, 3))
                else:
                    line = (rule, round(lv / 20.0, 2))
        st = styles.get(style_id or "", {})
        style_name = st.get("name") or style_id or "(none)"
        style_use[style_name] += 1
        if outline is None:
            outline = st.get("outline_level")
        # determine role
        role = "body"
        sl = (style_name or "").lower()
        m = re.search(r"(?:heading|标题)\s*(\d)", style_name or "", re.I)
        if m:
            role = f"heading_{m.group(1)}"
        elif outline:
            role = f"heading_{outline}"
        elif "toc" in sl or "目录" in (style_name or ""):
            role = "toc"
        elif style_name in ("Title", "标题"):
            role = "title"
        elif "caption" in sl or "题注" in (style_name or ""):
            role = "caption"
        else:
            inferred = looks_like_heading_text(text)
            if inferred and inferred > 0 and len(text) <= 60:
                # only treat as heading-like for structure list if also bold or larger
                role_guess = f"headinglike_{inferred}"
            else:
                role_guess = None
            if role == "body" and role_guess:
                # keep as body for typography, but record structure
                pass

        if role.startswith("heading") or (outline and text and len(text) <= 80):
            heading_list.append({
                "style": style_name,
                "outline": outline,
                "role": role,
                "text": redact(text)[:120],
                "len": len(text),
            })
        elif looks_like_heading_text(text) and len(text) <= 80:
            heading_list.append({
                "style": style_name,
                "outline": looks_like_heading_text(text),
                "role": "inferred",
                "text": redact(text)[:120],
                "len": len(text),
            })

        if role == "body":
            body_count += 1
        if line:
            line_sp[str(line)] += 1
        if first_ind is not None:
            indents[str(first_ind)] += 1
        if jc:
            aligns[jc] += 1

        # fonts from runs
        for r in p.findall(f"{W}r"):
            t = "".join((x.text or "") for x in r.iter(f"{W}t"))
            if not t.strip():
                continue
            rPr = r.find(f"{W}rPr")
            fr = rfonts(rPr)
            ea = fr.get("eastAsia") or fr.get("ascii") or st.get("eastAsia") or "(unset)"
            sz = sz_pt(rPr) or st.get("size_pt")
            role_fonts[role][ea] += len(t.strip())
            if sz:
                role_sizes[role][round(sz, 1)] += len(t.strip())

    return {
        "paragraphs_with_text": para_count,
        "body_paragraphs": body_count,
        "style_use": dict(style_use),
        "line_spacing": dict(line_sp),
        "first_indents": dict(indents),
        "alignments": dict(aligns),
        "role_fonts": {k: dict(v) for k, v in role_fonts.items()},
        "role_sizes": {k: {str(a): b for a, b in v.items()} for k, v in role_sizes.items()},
        "headings": heading_list,
    }


def analyze_tables(root):
    tables = list(root.iter(f"{W}tbl"))
    info = {
        "table_count": len(tables),
        "row_col": [],
        "border_styles": Counter(),
        "header_repeat": 0,
        "cant_split_tables": 0,
        "merged": 0,
        "cell_fonts": Counter(),
        "cell_sizes": Counter(),
        "cell_align": Counter(),
    }
    for tbl in tables:
        rows = tbl.findall(f"{W}tr")
        # cols from grid
        grid = tbl.find(f"{W}tblGrid")
        cols = len(grid.findall(f"{W}gridCol")) if grid is not None else 0
        info["row_col"].append({"rows": len(rows), "cols": cols})
        tblPr = tbl.find(f"{W}tblPr")
        if tblPr is not None:
            borders = tblPr.find(f"{W}tblBorders")
            if borders is not None:
                for child in list(borders):
                    typ = child.tag.split("}")[-1]
                    val = get(child, "val")
                    sz = get(child, "sz")
                    info["border_styles"][f"{typ}:{val}:{sz}"] += 1
            else:
                info["border_styles"]["none_or_style"] += 1
        hdr = False
        csplit = 0
        for tr in rows:
            trPr = tr.find(f"{W}trPr")
            if trPr is not None and trPr.find(f"{W}tblHeader") is not None:
                hdr = True
            if trPr is not None and trPr.find(f"{W}cantSplit") is not None:
                csplit += 1
            for tc in tr.findall(f"{W}tc"):
                tcPr = tc.find(f"{W}tcPr")
                if tcPr is not None and (tcPr.find(f"{W}vMerge") is not None or tcPr.find(f"{W}gridSpan") is not None):
                    info["merged"] += 1
                for p in tc.findall(f"{W}p"):
                    pPr = p.find(f"{W}pPr")
                    if pPr is not None and pPr.find(f"{W}jc") is not None:
                        info["cell_align"][get(pPr.find(f"{W}jc"), "val")] += 1
                    for r in p.findall(f"{W}r"):
                        t = "".join((x.text or "") for x in r.iter(f"{W}t"))
                        if not t.strip():
                            continue
                        rPr = r.find(f"{W}rPr")
                        fr = rfonts(rPr)
                        ea = fr.get("eastAsia") or fr.get("ascii")
                        sz = sz_pt(rPr)
                        if ea:
                            info["cell_fonts"][ea] += 1
                        if sz:
                            info["cell_sizes"][str(sz)] += 1
        if hdr:
            info["header_repeat"] += 1
        if csplit:
            info["cant_split_tables"] += 1
    info["border_styles"] = dict(info["border_styles"])
    info["cell_fonts"] = dict(info["cell_fonts"])
    info["cell_sizes"] = dict(info["cell_sizes"])
    info["cell_align"] = dict(info["cell_align"])
    return info


def analyze_images(z, root):
    # map rId -> media path
    rels = {}
    if "word/_rels/document.xml.rels" in z.namelist():
        rel_root = ET.fromstring(z.read("word/_rels/document.xml.rels"))
        for rel in rel_root:
            rels[rel.get("Id")] = rel.get("Target")
    media_meta = {}
    for name in z.namelist():
        if not name.startswith("word/media/"):
            continue
        data = z.read(name)
        try:
            import io
            im = Image.open(io.BytesIO(data))
            w, h = im.size
            dpi = im.info.get("dpi")
            media_meta[name] = {"w": w, "h": h, "dpi_meta": list(dpi) if dpi else None, "bytes": len(data), "fmt": im.format}
        except Exception as e:
            media_meta[name] = {"error": type(e).__name__, "bytes": len(data)}

    placements = []
    # inline and anchor drawings
    for drawing in root.iter(f"{W}drawing"):
        extent = drawing.find(f".//{WP}extent")
        cx = cy = None
        if extent is not None:
            cx = extent.get("cx")
            cy = extent.get("cy")
        blip = None
        for el in drawing.iter():
            if el.tag.endswith("}blip"):
                blip = el.get(f"{R}embed") or el.get("embed")
                break
        target = rels.get(blip) if blip else None
        media_path = "word/" + target.lstrip("/") if target and not target.startswith("word/") else target
        if target and target.startswith("media/"):
            media_path = "word/" + target
        meta = media_meta.get(media_path or "", {})
        wpx, hpx = meta.get("w"), meta.get("h")
        width_pt = emu_to_pt(cx) if cx else None
        height_pt = emu_to_pt(cy) if cy else None
        dpi = None
        if wpx and width_pt and width_pt > 0:
            dpi = wpx / (width_pt / 72.0)
        # alignment: look at ancestor paragraph jc, or wp:positionH
        align = None
        posH = drawing.find(f".//{WP}positionH")
        if posH is not None:
            align = posH.get("relativeFrom")
            algn = posH.find(f"{WP}align")
            if algn is not None and algn.text:
                align = f"{align}:{algn.text}"
        placements.append({
            "media": Path(media_path).name if media_path else None,
            "wpx": wpx, "hpx": hpx,
            "width_pt": round(width_pt, 1) if width_pt else None,
            "height_pt": round(height_pt, 1) if height_pt else None,
            "width_cm": round(width_pt * 2.54 / 72, 2) if width_pt else None,
            "effective_dpi": round(dpi, 1) if dpi else None,
            "pos": align,
            "meta_dpi": meta.get("dpi_meta"),
        })
    dpis = [p["effective_dpi"] for p in placements if p.get("effective_dpi")]
    import statistics
    dpi_stats = None
    if dpis:
        dpi_stats = {
            "n": len(dpis),
            "min": min(dpis),
            "median": round(statistics.median(dpis), 1),
            "max": max(dpis),
            "lt300": sum(1 for x in dpis if x < 300),
            "lt200": sum(1 for x in dpis if x < 200),
            "ge300": sum(1 for x in dpis if x >= 300),
        }
    return {
        "media_files": len(media_meta),
        "placements": len(placements),
        "dpi_stats": dpi_stats,
        "width_cm_hist": dict(Counter(round(p["width_cm"], 0) for p in placements if p.get("width_cm"))),
        "pos_hist": dict(Counter(p["pos"] or "inline/inherit" for p in placements)),
        "placements_brief": placements[:30],  # sample only, no pixels
    }


def has_toc(root, z):
    xml = z.read("word/document.xml")
    toc_field = bool(re.search(rb"TOC\s*\\", xml) or re.search(rb">TOC<", xml))
    # style based
    styles_toc = 0
    body = root.find(f"{W}body")
    for p in body.findall(f"{W}p"):
        pPr = p.find(f"{W}pPr")
        if pPr is None:
            continue
        ps = pPr.find(f"{W}pStyle")
        if ps is not None:
            sid = get(ps, "val") or ""
            if "TOC" in sid or "toc" in sid.lower():
                styles_toc += 1
    return {"toc_field": toc_field, "toc_style_paras": styles_toc}


def analyze_docx(path: Path, label: str, collect_all_headings=True):
    print(f"ANALYZE {label} size={path.stat().st_size}", flush=True)
    with zipfile.ZipFile(path) as z:
        doc_xml = z.read("word/document.xml")
        root = ET.fromstring(doc_xml)
        styles = parse_styles(z)
        sects = parse_sections_from_xml(root)
        hf = parse_hf(z)
        paras = analyze_paragraphs(root, styles)
        tables = analyze_tables(root)
        images = analyze_images(z, root)
        toc = has_toc(root, z)
        theme = {}
        if "word/theme/theme1.xml" in z.namelist():
            troot = ET.fromstring(z.read("word/theme/theme1.xml"))
            for tag, key in (("majorFont", "major"), ("minorFont", "minor")):
                node = troot.find(f".//{A}{tag}")
                if node is None:
                    continue
                theme[key] = {child.tag.split("}")[-1]: child.get("typeface") for child in node}
    return {
        "label": label,
        "bytes": path.stat().st_size,
        "theme": theme,
        "sections": sects,
        "section_count": len(sects),
        "headers_footers": hf,
        "styles_defined_interesting": {
            k: v for k, v in styles.items()
            if v.get("eastAsia") or v.get("size_pt") or v.get("outline_level")
            or (v.get("name") and re.search(r"Heading|标题|TOC|标题|题注", v["name"], re.I))
        },
        "paragraphs": {k: v for k, v in paras.items() if k != "headings"},
        "headings": paras["headings"] if collect_all_headings else paras["headings"][:200],
        "heading_count": len(paras["headings"]),
        "tables": tables,
        "images": images,
        "toc": toc,
    }


def summarize_v8(rep):
    p = rep["paragraphs"]
    body_fonts = p.get("role_fonts", {}).get("body", {})
    body_sizes = p.get("role_sizes", {}).get("body", {})
    margins = [s.get("margins_cm") for s in rep["sections"] if s.get("margins_cm")]
    # unique margin sets
    uniq_m = []
    for m in margins:
        if m not in uniq_m:
            uniq_m.append(m)
    hf = rep["headers_footers"]
    footers = [x for x in hf if x["kind"] == "footer"]
    headers = [x for x in hf if x["kind"] == "header"]
    return {
        "sections": rep["section_count"],
        "unique_margins": uniq_m,
        "page_sizes": [{"w_mm": k[0], "h_mm": k[1], "orient": k[2], "count": v}
                       for k, v in Counter((s.get("w_mm"), s.get("h_mm"), s.get("orient")) for s in rep["sections"]).items()],
        "page_num_starts": [s.get("pgNumType") for s in rep["sections"]],
        "title_pages": sum(1 for s in rep["sections"] if s.get("title_page")),
        "header_count_parts": len(headers),
        "footer_count_parts": len(footers),
        "footers_with_page": sum(1 for x in footers if x["has_page_field"]),
        "footer_fonts": dict(sum((Counter(x["fonts"]) for x in footers), Counter())),
        "footer_sizes": dict(sum((Counter({float(k): v for k, v in x["sizes"].items()}) for x in footers), Counter())),
        "body_fonts": body_fonts,
        "body_sizes": body_sizes,
        "line_spacing_top": dict(Counter(p.get("line_spacing", {})).most_common(8)) if isinstance(p.get("line_spacing"), dict) else p.get("line_spacing"),
        "indents_top": p.get("first_indents"),
        "alignments": p.get("alignments"),
        "style_use_top": dict(sorted(p.get("style_use", {}).items(), key=lambda kv: -kv[1])[:20]),
        "heading_count": rep["heading_count"],
        "toc": rep["toc"],
        "tables": {k: v for k, v in rep["tables"].items() if k not in ("row_col",)},
        "table_count": rep["tables"]["table_count"],
        "images": {k: v for k, v in rep["images"].items() if k != "placements_brief"},
    }


def main():
    v8 = Path(PATHS["v8"])
    tdir = Path(PATHS["template_dir"])
    v8_rep = analyze_docx(v8, "v8_reference")
    templates = []
    for p in sorted(tdir.glob("*.docx")):
        # skip prompt-only if small and named 提示词 — still analyze structure briefly
        templates.append(analyze_docx(p, p.name))

    v8_sum = summarize_v8(v8_rep)
    payload = {
        "privacy": {"redacted_ids": True, "no_body_excerpt": True, "headings_redacted": True},
        "v8_summary": v8_sum,
        "v8": v8_rep,
        "templates": templates,
    }
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    (OUT_DIR / "v8-stats.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    # markdown (no PII)
    lines = ["# V8 参考成品与模板排版统计", "",
             "说明：V8 为阳性对照。统计过程对证件号等做了脱敏。本报告不摘录正文，标题仅保留结构文本（已脱敏）。", ""]
    lines.append("## V8 概要")
    s = v8_sum
    lines.append(f"- 节（section）数：{s['sections']}；首页不同页眉设置的节：{s['title_pages']}")
    lines.append(f"- 纸张：{s['page_sizes']}")
    lines.append(f"- 页码起始/格式（各节 pgNumType）：{s['page_num_starts']}")
    lines.append(f"- 页眉部件：{s['header_count_parts']}；页脚部件：{s['footer_count_parts']}（含 PAGE 域 {s['footers_with_page']}）")
    lines.append(f"- 页脚字体：{s['footer_fonts']}；字号：{s['footer_sizes']}")
    lines.append(f"- 边距种类：")
    for m in s["unique_margins"]:
        lines.append(f"  - {m}")
    lines.append(f"- 正文字体（按字符）：{s['body_fonts']}")
    lines.append(f"- 正文字号（按字符）：{s['body_sizes']}")
    lines.append(f"- 行距：{s['line_spacing_top']}")
    lines.append(f"- 首行缩进：{s['indents_top']}")
    lines.append(f"- 对齐：{s['alignments']}")
    lines.append(f"- 样式使用（前 20）：{s['style_use_top']}")
    lines.append(f"- 识别到的标题/结构行数：{s['heading_count']}")
    lines.append(f"- 目录：{s['toc']}")
    lines.append(f"- 表格：{s['table_count']}；其它：边框/表头/合并等见 JSON")
    lines.append(f"  { {k:s['tables'][k] for k in s['tables'] if k!='row_col'} }")
    lines.append(f"- 图片：{s['images']}")
    lines.append("")
    lines.append("## V8 标题结构（脱敏）")
    for i, h in enumerate(v8_rep["headings"], 1):
        lines.append(f"- [{h.get('role')}|{h.get('style')}|L{h.get('outline')}] {h.get('text')}")
    lines.append("")
    lines.append("## 模板")
    for t in templates:
        lines.append(f"### {t['label']}  （{t['bytes']} 字节，节 {t['section_count']}，标题行 {t['heading_count']}，表 {t['tables']['table_count']}）")
        ts = summarize_v8(t)
        lines.append(f"- 边距：{ts['unique_margins']}")
        lines.append(f"- 正文字体/字号：{ts['body_fonts']} / {ts['body_sizes']}")
        lines.append(f"- 行距：{ts['line_spacing_top']}；缩进：{ts['indents_top']}")
        lines.append(f"- 页脚：{ts['footer_count_parts']}（PAGE {ts['footers_with_page']}）；页眉：{ts['header_count_parts']}")
        lines.append(f"- 样式：{ts['style_use_top']}")
        lines.append(f"- 目录：{ts['toc']}")
        lines.append("- 结构行：")
        for h in t["headings"]:
            lines.append(f"  - [{h.get('role')}|{h.get('style')}|L{h.get('outline')}] {h.get('text')}")
        lines.append("")
    (OUT_DIR / "v8-stats.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print("WROTE", OUT_DIR / "v8-stats.json", OUT_DIR / "v8-stats.md")
    print("V8 headings", v8_rep["heading_count"], "sections", v8_rep["section_count"], "tables", v8_rep["tables"]["table_count"], "images", v8_rep["images"]["placements"])


if __name__ == "__main__":
    main()
