#!/usr/bin/env python3
"""稳定的全文 PDF 访问：pypdfium2 渲染 + poppler 文本/图像列表，避免 PyMuPDF 在长 PDF 上 SIGSEGV。"""
from __future__ import annotations

import re
import subprocess
from pathlib import Path


class Rect:
    def __init__(self, x0, y0, x1, y1):
        self.x0, self.y0, self.x1, self.y1 = float(x0), float(y0), float(x1), float(y1)
        self.width = self.x1 - self.x0
        self.height = self.y1 - self.y0


class Pixmap:
    def __init__(self, width, height, samples: bytes):
        self.width = width
        self.height = height
        self.samples = samples


class SafePage:
    def __init__(self, doc: "SafePdf", index: int):
        self._doc = doc
        self.number = index  # 0-based
        w, h = doc._sizes[index]
        self.rect = Rect(0, 0, w, h)

    def get_text(self, mode: str = "text", clip=None):
        # 需要 clip 或 dict（带 bbox）时走 pypdfium2，避免全文串进页脚检查
        if mode == "dict" or clip is not None:
            return self._extract_pdfium(mode, clip)
        return self._doc._texts[self.number]

    def _extract_pdfium(self, mode: str, clip=None):
        import pypdfium2 as pdfium
        pdf = pdfium.PdfDocument(str(self._doc.path))
        try:
            page = pdf[self.number]
            w, h = page.get_size()
            tp = page.get_textpage()
            def bounded(l, b, r, tp_):
                return tp.get_text_bounded(left=l, bottom=b, right=r, top=tp_) or ""
            if clip is not None:
                left, right = float(clip.x0), float(clip.x1)
                bottom = max(0.0, h - float(clip.y1))
                top = min(h, h - float(clip.y0))
                text = bounded(left, bottom, right, top)
            else:
                text = tp.get_text_bounded() or ""
            if mode != "dict":
                tp.close()
                return text
            blocks = []
            for i in range(tp.count_rects()):
                box = tp.get_rect(i)
                x0, y0, x1, y1 = box
                fx0, fy0, fx1, fy1 = x0, h - y1, x1, h - y0
                t = bounded(x0, y0, x1, y1)
                blocks.append({
                    "type": 0,
                    "bbox": (fx0, fy0, fx1, fy1),
                    "lines": [{"spans": [{"text": t}]}],
                })
            tp.close()
            return {"blocks": blocks}
        finally:
            pdf.close()

    def search_for(self, needle: str):
        if not needle:
            return []
        t = self._doc._texts[self.number]
        if needle in t or needle.replace(" ", "") in t.replace(" ", ""):
            return [True]
        return []

    def get_images(self, full=True):
        imgs = [im for im in self._doc._images if im["page"] == self.number + 1]
        # fitz-like tuple, xref = synthetic
        return [(im["xref"], 0, im["wpx"], im["hpx"]) for im in imgs]

    def get_image_rects(self, xref):
        for im in self._doc._images:
            if im["xref"] == xref and im["page"] == self.number + 1:
                return [im["rect"]]
        return []

    def get_pixmap(self, matrix=None, colorspace=None, clip=None):
        scale = 1.0
        if matrix is not None:
            scale = float(getattr(matrix, "a", None) or getattr(matrix, "xx", None) or 1.0)
        import pypdfium2 as pdfium
        # render this page only
        pdf = pdfium.PdfDocument(str(self._doc.path))
        try:
            page = pdf[self.number]
            # clip in pt → bitmap
            w_pt, h_pt = page.get_size()
            if clip is not None:
                # render full then crop
                bmp = page.render(scale=scale, rev_byteorder=True)
                pil = bmp.to_pil().convert("L")
                # pdfium origin bottom-left; pil top-left. Use clip as top-origin (fitz style).
                x0 = int(max(0, clip.x0) * scale)
                y0 = int(max(0, clip.y0) * scale)
                x1 = int(min(w_pt, clip.x1) * scale)
                y1 = int(min(h_pt, clip.y1) * scale)
                pil = pil.crop((x0, y0, max(x0 + 1, x1), max(y0 + 1, y1)))
            else:
                bmp = page.render(scale=scale, rev_byteorder=True)
                pil = bmp.to_pil().convert("L")
            import numpy as np
            arr = np.array(pil, dtype=np.uint8)
            return Pixmap(arr.shape[1], arr.shape[0], arr.tobytes())
        finally:
            pdf.close()


class SafePdf:
    def __init__(self, path: str | Path):
        self.path = Path(path)
        self._sizes = _page_sizes(self.path)
        self.page_count = len(self._sizes)
        self._texts = _page_texts(self.path, self.page_count)
        self._images = _image_list(self.path, self._sizes)
        self._pdfium = None  # opened lazily, not kept (avoid long-lived handle)

    def __len__(self):
        return self.page_count

    def __iter__(self):
        for i in range(self.page_count):
            yield SafePage(self, i)

    def __getitem__(self, i):
        if i < 0:
            i += self.page_count
        return SafePage(self, i)

    def extract_image(self, xref):
        for im in self._images:
            if im["xref"] == xref:
                return {"width": im["wpx"], "height": im["hpx"]}
        return {"width": 0, "height": 0}

    def close(self):
        self._pdfium = None

    @property
    def landscape_pages(self) -> list[int]:
        return [i + 1 for i, (w, h) in enumerate(self._sizes) if w > h + 1]


def _page_sizes(path: Path) -> list[tuple[float, float]]:
    out = subprocess.check_output(
        ["pdfinfo", "-f", "1", "-l", "10000", str(path)],
        text=True, errors="replace",
    )
    sizes = []
    for m in re.finditer(r"Page\s+(\d+)\s+size:\s+([\d.]+)\s+x\s+([\d.]+)", out):
        sizes.append((float(m.group(2)), float(m.group(3))))
    if not sizes:
        # fallback pypdfium2
        import pypdfium2 as pdfium
        pdf = pdfium.PdfDocument(str(path))
        try:
            for i in range(len(pdf)):
                sizes.append(tuple(pdf[i].get_size()))
        finally:
            pdf.close()
    return sizes


def _page_texts(path: Path, n_pages: int) -> list[str]:
    try:
        raw = subprocess.check_output(
            ["pdftotext", "-layout", str(path), "-"],
            text=True, errors="replace",
        )
        parts = raw.split("\f")
        # pdftotext often has trailing empty after last ff
        if parts and parts[-1].strip() == "":
            parts = parts[:-1]
        if len(parts) < n_pages:
            parts += [""] * (n_pages - len(parts))
        return parts[:n_pages]
    except Exception:
        import pypdfium2 as pdfium
        pdf = pdfium.PdfDocument(str(path))
        texts = []
        try:
            for i in range(len(pdf)):
                page = pdf[i]
                tp = page.get_textpage()
                texts.append(tp.get_text_bounded() or "")
                tp.close()
        finally:
            pdf.close()
        return texts


def _image_list(path: Path, sizes: list[tuple[float, float]]) -> list[dict]:
    """r3：优先用 pypdfium2 页面对象读取图像真实放置位置（bounds），不再按页居中估计。
    失败时回退到 r2 的 pdfimages -list + 居中估计。"""
    try:
        rows = _image_list_pdfium(path)
        if rows:
            return rows
    except Exception:
        pass
    return _image_list_pdfimages(path, sizes)


def _image_list_pdfium(path: Path) -> list[dict]:
    import pypdfium2 as pdfium
    import pypdfium2.raw as pdfium_c
    rows = []
    pdf = pdfium.PdfDocument(str(path))
    xref = 1
    try:
        for i in range(len(pdf)):
            page = pdf[i]
            w, h = page.get_size()
            for obj in page.get_objects(filter=[pdfium_c.FPDF_PAGEOBJ_IMAGE], max_depth=3):
                l, b, r, t = obj.get_bounds()
                wpx, hpx = obj.get_px_size()
                wpt, hpt = r - l, t - b
                if wpt < 1 or hpt < 1:
                    continue
                dpi = ((wpx / (wpt / 72.0)) + (hpx / (hpt / 72.0))) / 2.0
                rows.append({
                    "page": i + 1, "xref": xref, "wpx": wpx, "hpx": hpx, "dpi": dpi,
                    "rect": Rect(l, h - t, r, h - b), "width_pt": wpt, "height_pt": hpt,
                    "rect_source": "pdfium_bounds",
                })
                xref += 1
    finally:
        pdf.close()
    return rows


def text_rects(path: Path, index: int) -> list[tuple[float, float, float, float, str]]:
    """单页文字段矩形（top-left 原点，pt）及文字。"""
    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(str(path))
    try:
        page = pdf[index]
        w, h = page.get_size()
        tp = page.get_textpage()
        out = []
        for k in range(tp.count_rects()):
            x0, y0, x1, y1 = tp.get_rect(k)
            t = tp.get_text_bounded(x0, y0, x1, y1) or ""
            out.append((x0, h - y1, x1, h - y0, t))
        tp.close()
        return out
    finally:
        pdf.close()


def all_text_rects(path: Path) -> list[list[tuple[float, float, float, float, str]]]:
    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(str(path))
    res = []
    try:
        for i in range(len(pdf)):
            page = pdf[i]
            w, h = page.get_size()
            tp = page.get_textpage()
            out = []
            for k in range(tp.count_rects()):
                x0, y0, x1, y1 = tp.get_rect(k)
                t = tp.get_text_bounded(x0, y0, x1, y1) or ""
                out.append((x0, h - y1, x1, h - y0, t))
            tp.close()
            res.append(out)
    finally:
        pdf.close()
    return res


_FOOT_RE = re.compile(r"[—\-－]\s*([0-9]+|[IVXLC]+)\s*[—\-－]|第\s*(\d+)\s*页")


def footer_labels(path: Path, band_pt: float = 1.8 / 2.54 * 72) -> list[dict]:
    """逐页只读页脚带（距纸底 band_pt 内）的文字，解析页码标签：arabic / roman / None。"""
    rects = all_text_rects(path)
    import pypdfium2 as pdfium
    pdf = pdfium.PdfDocument(str(path))
    sizes = [pdf[i].get_size() for i in range(len(pdf))]
    pdf.close()
    out = []
    for i, rs in enumerate(rects):
        w, h = sizes[i]
        band = [r for r in rs if r[1] >= h - band_pt]
        band.sort(key=lambda r: r[0])
        txt = "".join(r[4] for r in band).replace("\r", "").replace("\n", " ")
        m = _FOOT_RE.search(txt)
        lab = {"page": i + 1, "text": txt.strip()[:40], "kind": None, "n": None, "bbox": None}
        if band:
            lab["bbox"] = (min(r[0] for r in band), min(r[1] for r in band), max(r[2] for r in band), max(r[3] for r in band))
        if m:
            if m.group(2):
                lab["kind"], lab["n"] = "arabic", int(m.group(2))
            elif m.group(1).isdigit():
                lab["kind"], lab["n"] = "arabic", int(m.group(1))
            else:
                from bid_r3_fixes import roman_to_int
                lab["kind"], lab["n"] = "roman", roman_to_int(m.group(1))
            lab["style"] = re.sub(r"[0-9]+", "N", re.sub(r"[IVXLC]+", "R", m.group(0)))
        elif re.fullmatch(r"\s*\d{1,4}\s*", txt or ""):
            lab["kind"], lab["n"], lab["style"] = "arabic", int(txt.strip()), "N"
        elif re.fullmatch(r"\s*[IVXLC]{1,6}\s*", txt or ""):
            from bid_r3_fixes import roman_to_int
            lab["kind"], lab["n"], lab["style"] = "roman", roman_to_int(txt.strip()), "R"
        out.append(lab)
    return out


def _image_list_pdfimages(path: Path, sizes: list[tuple[float, float]]) -> list[dict]:
    rows = []
    try:
        out = subprocess.check_output(
            ["pdfimages", "-list", str(path)],
            text=True, errors="replace",
        )
    except Exception:
        return rows
    # skip header lines
    xref = 1
    for line in out.splitlines():
        m = re.match(
            r"\s*(\d+)\s+(\d+)\s+\S+\s+(\d+)\s+(\d+)\s+\S+\s+\d+\s+\d+\s+\S+\s+\S+\s+\S+\s+\S+\s+(\d+)\s+(\d+)",
            line,
        )
        # fallback looser: page num width height ... xppi yppi
        if not m:
            parts = line.split()
            if len(parts) >= 5 and parts[0].isdigit() and parts[1].isdigit():
                try:
                    page = int(parts[0]); wpx = int(parts[2]); hpx = int(parts[3])
                    # xppi yppi often last numeric pair before size
                    nums = [p for p in parts if re.fullmatch(r"\d+(\.\d+)?", p)]
                    xppi = yppi = 0.0
                    if len(nums) >= 8:
                        xppi, yppi = float(nums[-4]), float(nums[-3])
                    elif len(parts) >= 14:
                        xppi, yppi = float(parts[-4]), float(parts[-3])
                except Exception:
                    continue
            else:
                continue
        else:
            page = int(m.group(1)); wpx = int(m.group(3)); hpx = int(m.group(4))
            xppi, yppi = float(m.group(5)), float(m.group(6))
        if page < 1 or page > len(sizes):
            continue
        pw, ph = sizes[page - 1]
        # estimate display rect from ppi if present, else assume 16cm width
        if xppi > 1 and yppi > 1:
            wpt = wpx / xppi * 72.0
            hpt = hpx / yppi * 72.0
            dpi = (xppi + yppi) / 2.0
        else:
            wpt = min(pw * 0.9, 16 / 2.54 * 72)
            hpt = wpt * (hpx / max(wpx, 1))
            dpi = wpx / (wpt / 72.0) if wpt else 0
        # centered guess
        x0 = max(0, (pw - wpt) / 2)
        y0 = max(0, (ph - hpt) / 2)
        rows.append({
            "page": page, "xref": xref, "wpx": wpx, "hpx": hpx,
            "dpi": dpi, "rect": Rect(x0, y0, x0 + wpt, y0 + hpt),
            "width_pt": wpt, "height_pt": hpt,
        })
        xref += 1
    return rows


def open_pdf(path: str | Path) -> SafePdf:
    return SafePdf(path)


def summarize(path: str | Path) -> dict:
    pdf = SafePdf(path)
    dpsis = [im["dpi"] for im in pdf._images if im["dpi"] > 1]
    dpsis.sort()
    def pct(a, p):
        if not a:
            return None
        i = int(round((p / 100) * (len(a) - 1)))
        return a[min(max(i, 0), len(a) - 1)]
    return {
        "path": str(path),
        "pages": pdf.page_count,
        "landscape_pages": len(pdf.landscape_pages),
        "landscape_runs": _runs(pdf.landscape_pages),
        "images": len(pdf._images),
        "dpi_n": len(dpsis),
        "dpi_min": dpsis[0] if dpsis else None,
        "dpi_median": pct(dpsis, 50),
        "dpi_p25": pct(dpsis, 25),
        "dpi_p75": pct(dpsis, 75),
        "dpi_max": dpsis[-1] if dpsis else None,
        "bytes": Path(path).stat().st_size,
    }


def _runs(nums: list[int]) -> list[tuple[int, int, int]]:
    if not nums:
        return []
    runs = []
    s = prev = nums[0]
    for n in nums[1:]:
        if n == prev + 1:
            prev = n
        else:
            runs.append((s, prev, prev - s + 1))
            s = prev = n
    runs.append((s, prev, prev - s + 1))
    return runs
