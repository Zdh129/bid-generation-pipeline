#!/usr/bin/env python3
"""对本机扫描件附件做离线 OCR（RapidOCR/ONNX），只下载模型权重，不上传数据。"""
from __future__ import annotations

import json
import re
import time
from pathlib import Path

import numpy as np
import pymupdf
from rapidocr_onnxruntime import RapidOCR

ROOT = Path(__file__).resolve().parents[1]
INDEX = ROOT / "work" / "att_corpus" / "index.json"
OUT_DIR = ROOT / "work" / "att_corpus"
DPI = 220


def render_pdf_page(doc, i, dpi=DPI):
    page = doc[i]
    zoom = dpi / 72.0
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
    h, w = pix.height, pix.width
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(h, w, pix.n)
    return arr


def render_image(path: Path):
    import pymupdf
    # reuse pymupdf / PIL
    from PIL import Image
    im = Image.open(path).convert("RGB")
    return np.array(im)


def ocr_image(engine, arr) -> tuple[str, float]:
    result, elapse = engine(arr)
    lines = []
    if result:
        for item in result:
            # item: [box, text, conf]
            if len(item) >= 2 and item[1]:
                lines.append(str(item[1]))
    return "\n".join(lines), float(elapse or 0)


def should_ocr(rec: dict) -> bool:
    if rec.get("error"):
        return False
    if rec.get("likely_scan"):
        return True
    # 低文本确认函也补做
    if rec.get("kind") in {"confirm"} and (rec.get("total_chars") or 0) < 80:
        return True
    return False


def main():
    idx = json.loads(INDEX.read_text(encoding="utf-8"))
    targets = [r for r in idx["index"] if should_ocr(r)]
    # 优先确认函 / 偏离表，其次短彩页，最后长彩页
    def pri(r):
        kind = 0 if r["kind"] == "confirm" or "偏离" in (r.get("file") or "") else (1 if (r.get("pages") or 1) <= 4 else 2)
        return (kind, r.get("pages") or 1, r["seq"])
    targets.sort(key=pri)
    print(f"ocr targets={len(targets)} pages≈{sum(r.get('pages') or 1 for r in targets)}")
    t0 = time.time()
    engine = RapidOCR()
    print(f"engine loaded in {time.time()-t0:.1f}s")
    results = []
    for rec in targets:
        path = Path(json.loads(Path(rec["corpus_file"]).read_text(encoding="utf-8"))["path"])
        kind = rec["kind"]
        seq = rec["seq"]
        page_texts = []
        t1 = time.time()
        try:
            if path.suffix.lower() == ".pdf":
                doc = pymupdf.open(path)
                for i in range(doc.page_count):
                    # 若该页已有足够文字则跳过
                    raw = json.loads(Path(rec["corpus_file"]).read_text(encoding="utf-8"))
                    existing = ""
                    if i < len(raw.get("page_texts") or []):
                        existing = raw["page_texts"][i].get("text") or ""
                    compact = re.sub(r"\s+", "", existing)
                    if len(compact) >= 80:
                        page_texts.append({"page": i + 1, "chars": len(compact), "text": existing, "source": "embedded_text"})
                        continue
                    arr = render_pdf_page(doc, i)
                    text, el = ocr_image(engine, arr)
                    compact = re.sub(r"\s+", "", text)
                    page_texts.append({"page": i + 1, "chars": len(compact), "text": text, "source": "rapidocr", "ocr_s": round(el, 3)})
                doc.close()
            else:
                arr = render_image(path)
                text, el = ocr_image(engine, arr)
                compact = re.sub(r"\s+", "", text)
                page_texts.append({"page": 1, "chars": len(compact), "text": text, "source": "rapidocr", "ocr_s": round(el, 3)})
        except Exception as exc:
            rec_err = dict(rec)
            rec_err["ocr_error"] = f"{type(exc).__name__}: {exc}"
            results.append(rec_err)
            print(f"FAIL eq{seq:02d} {path.name}: {type(exc).__name__}: {exc}")
            continue
        total = sum(p["chars"] for p in page_texts)
        # 回写 corpus
        raw = json.loads(Path(rec["corpus_file"]).read_text(encoding="utf-8"))
        raw["page_texts"] = page_texts
        raw["total_chars"] = total
        raw["avg_chars"] = round(total / max(len(page_texts), 1), 1)
        raw["ocr_applied"] = True
        raw["likely_scan"] = total < 80
        raw["scan_pages"] = [p["page"] for p in page_texts if p["chars"] < 40]
        Path(rec["corpus_file"]).write_text(json.dumps(raw, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        rec2 = dict(rec)
        rec2.update({"total_chars": total, "ocr_applied": True, "elapsed_s": round(time.time() - t1, 2), "likely_scan": raw["likely_scan"]})
        results.append(rec2)
        print(f"OK eq{seq:02d} {kind} p={len(page_texts)} chars={total} {time.time()-t1:.1f}s {path.name[:50]}")
    (OUT_DIR / "ocr_results.json").write_text(json.dumps({"elapsed_s": round(time.time()-t0, 2), "results": results}, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    ok = [r for r in results if r.get("ocr_applied") and not r.get("ocr_error")]
    print(f"done {len(ok)}/{len(targets)} elapsed={time.time()-t0:.1f}s residual_lowtext={sum(1 for r in ok if r.get('likely_scan'))}")


if __name__ == "__main__":
    main()
