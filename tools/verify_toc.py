#!/usr/bin/env python3
"""核对 PDF 目录页码与正文标题的印刷页码（页脚页码）。r3：含目录续页；按页脚印刷页码而非物理页序比较。"""
from __future__ import annotations
import json, re, sys
from pathlib import Path
from pdf_safe import open_pdf, footer_labels
from check_layout import heading_pages_in_pdf

LINE = re.compile(r"(.+?)(?:\s*[\.·．…\s]{2,}\s*|\t+)(\d+)\s*$")


def verify(pdf_path: Path, out: Path | None = None) -> dict:
    pdf = open_pdf(pdf_path)
    labs = footer_labels(pdf_path)
    first_ar = next((i for i, l in enumerate(labs) if l["kind"] == "arabic"), 0)
    toc_pages = []
    for i in range(first_ar if first_ar else pdf.page_count):
        t = pdf[i].get_text("text") or ""
        if re.search(r"^\s*目录\s*$", t[:400], re.M) or t.count("....") >= 6 or t.count("…") >= 6:
            toc_pages.append(i)
    entries = []
    for i in toc_pages:
        for line in (pdf[i].get_text("text") or "").splitlines():
            s = line.strip()
            if not s or s == "目录":
                continue
            m = LINE.search(s)
            if m:
                entries.append({"text": re.sub(r"\s+", " ", m.group(1).strip()), "page": int(m.group(2)), "toc_pdf_page": i + 1})
    phys = heading_pages_in_pdf(pdf, [{"text": e["text"]} for e in entries])
    label_of = {l["page"]: (l["n"] if l["kind"] == "arabic" else None) for l in labs}
    rows, ok = [], 0
    for e in entries:
        p = phys.get(e["text"])
        printed = label_of.get(p) if p else None
        match = printed == e["page"]
        ok += match
        rows.append({"title": e["text"], "toc_page": e["page"], "actual_printed_page": printed, "actual_pdf_page": p, "match": match})
    report = {
        "pdf": str(pdf_path),
        "toc_pdf_pages": [i + 1 for i in toc_pages],
        "front_matter_pages": first_ar,
        "n_toc": len(entries), "n_match": ok,
        "accuracy": (ok / len(entries)) if entries else None,
        "pages": pdf.page_count, "landscape_pages": len(pdf.landscape_pages),
        "rows": rows,
    }
    pdf.close()
    out = out or Path(pdf_path).with_suffix(".toc_verify.json")
    Path(out).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return report


def main(pdf_path: Path):
    r = verify(pdf_path)
    print(json.dumps({k: v for k, v in r.items() if k != "rows"}, ensure_ascii=False, indent=2))
    bad = [x for x in r["rows"] if not x["match"]]
    if bad:
        print("mismatches", json.dumps(bad[:20], ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main(Path(sys.argv[1]))
