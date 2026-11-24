#!/usr/bin/env python3
"""从设备附件（彩页/确认函/说明书等）本机抽取文字，供偏离判断。不外发、不调外部 API。"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

import pymupdf
from docx import Document

ROOT = Path(__file__).resolve().parents[1]
ATT = Path(
    "data/example-package"
    "/3.生成初稿标书/生成初版标书所需资料/3、投标设备信息（JS）"
)
OUT = ROOT / "work" / "att_corpus"
OUT.mkdir(parents=True, exist_ok=True)

SCAN_CHAR_THRESHOLD = 40  # 每页平均字符少于此视为可能扫描页


def classify_file(path: Path) -> str:
    n = path.name
    if re.search(r"确认|偏离", n):
        return "confirm"
    if re.search(r"彩页|说明书|白皮书", n):
        return "brochure"
    if re.search(r"承诺", n):
        return "commitment"
    if re.search(r"检测报告|校准|证书", n):
        return "report"
    if re.search(r"方案", n):
        return "scheme"
    return "other"


def extract_pdf(path: Path) -> dict:
    doc = pymupdf.open(path)
    pages = []
    for i, page in enumerate(doc):
        text = page.get_text("text") or ""
        text = text.replace("\u00a0", " ").strip()
        pages.append(
            {
                "page": i + 1,
                "chars": len(re.sub(r"\s+", "", text)),
                "width": page.rect.width,
                "height": page.rect.height,
                "text": text,
            }
        )
    n = doc.page_count
    doc.close()
    total_chars = sum(p["chars"] for p in pages)
    scan_pages = [p["page"] for p in pages if p["chars"] < SCAN_CHAR_THRESHOLD]
    return {
        "pages": n,
        "total_chars": total_chars,
        "avg_chars": round(total_chars / n, 1) if n else 0,
        "scan_pages": scan_pages,
        "likely_scan": bool(n and (len(scan_pages) / n >= 0.5 or total_chars < 80 * n and total_chars < 400)),
        "page_texts": pages,
    }


def extract_docx(path: Path) -> dict:
    d = Document(str(path))
    paras = [p.text.strip() for p in d.paragraphs if p.text.strip()]
    tables = []
    for ti, tbl in enumerate(d.tables):
        rows = []
        for row in tbl.rows:
            rows.append([c.text.strip() for c in row.cells])
        tables.append(rows)
    text = "\n".join(paras)
    for rows in tables:
        text += "\n" + "\n".join(" | ".join(r) for r in rows)
    compact = re.sub(r"\s+", "", text)
    return {
        "pages": None,
        "total_chars": len(compact),
        "avg_chars": len(compact),
        "scan_pages": [],
        "likely_scan": False,
        "page_texts": [{"page": 1, "chars": len(compact), "text": text}],
        "tables": tables,
    }


def extract_image_placeholder(path: Path) -> dict:
    return {
        "pages": 1,
        "total_chars": 0,
        "avg_chars": 0,
        "scan_pages": [1],
        "likely_scan": True,
        "page_texts": [{"page": 1, "chars": 0, "text": "", "note": "raster image, needs local OCR"}],
    }


def main():
    by_dev = defaultdict(list)
    index = []
    for d in sorted(ATT.iterdir()):
        if not d.is_dir():
            continue
        m = re.search(r"(?:示例|ex|pkg|包2)-0*(\d+)", d.name)
        if not m:
            continue
        seq = int(m.group(1))
        files = []
        for p in sorted(d.iterdir()):
            if not p.is_file():
                continue
            kind = classify_file(p)
            ext = p.suffix.lower()
            rec = {
                "seq": seq,
                "dir": d.name,
                "file": p.name,
                "path": str(p),
                "ext": ext,
                "kind": kind,
                "bytes": p.stat().st_size,
            }
            try:
                if ext == ".pdf":
                    rec.update(extract_pdf(p))
                elif ext in {".docx", ".doc"}:
                    rec.update(extract_docx(p))
                elif ext in {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}:
                    rec.update(extract_image_placeholder(p))
                else:
                    rec["error"] = f"unsupported {ext}"
            except Exception as exc:
                rec["error"] = f"{type(exc).__name__}: {exc}"
            # 写单文件文本，不把路径以外的敏感信息打到 stdout
            safe_stem = re.sub(r"[^\w.#（）()-]+", "_", p.stem)[:80]
            outp = OUT / f"eq{seq:02d}_{kind}_{safe_stem}.json"
            # 精简落盘
            outp.write_text(json.dumps(rec, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
            rec["corpus_file"] = str(outp)
            files.append(rec)
            index.append({k: rec.get(k) for k in (
                "seq", "file", "kind", "ext", "bytes", "pages", "total_chars",
                "avg_chars", "likely_scan", "scan_pages", "error", "corpus_file"
            )})
        by_dev[seq] = files

    # 摘要：不列出全文
    n_files = len(index)
    n_scan = sum(1 for r in index if r.get("likely_scan"))
    n_err = sum(1 for r in index if r.get("error"))
    summary = {
        "n_files": n_files,
        "n_scan_or_lowtext": n_scan,
        "n_errors": n_err,
        "by_kind": {},
        "scan_files": [r for r in index if r.get("likely_scan")],
        "index": index,
    }
    from collections import Counter
    summary["by_kind"] = dict(Counter(r["kind"] for r in index))
    (OUT / "index.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"files={n_files} scan_or_lowtext={n_scan} errors={n_err} kinds={summary['by_kind']}")
    # 确认函扫描比例最关键
    conf = [r for r in index if r["kind"] == "confirm"]
    print(f"confirm={len(conf)} lowtext={sum(1 for r in conf if r.get('likely_scan'))} "
          f"median_chars={sorted(r.get('total_chars') or 0 for r in conf)[len(conf)//2] if conf else 0}")
    bro = [r for r in index if r["kind"] == "brochure"]
    print(f"brochure={len(bro)} lowtext={sum(1 for r in bro if r.get('likely_scan'))}")


if __name__ == "__main__":
    main()
