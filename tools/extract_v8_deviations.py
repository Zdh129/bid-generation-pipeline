#!/usr/bin/env python3
"""从 V8 DOCX 抽取技术参数响应/偏离列。本机处理，不调外部 API。"""
from __future__ import annotations

import json
import re
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]
V8 = Path(
    "data/example-package"
    "/3.生成初稿标书/example-v8-reference.docx"
)
OUT = ROOT / "work" / "v8_deviations.json"

HEADER_HINTS = ("招标文件要求", "投标设备", "偏离")


def cell_text(cell) -> str:
    t = "\n".join(p.text for p in cell.paragraphs)
    return re.sub(r"[ \t]+", " ", t).replace("\u00a0", " ").strip()


def normalize_header(s: str) -> str:
    return re.sub(r"\s+", "", s)


def map_headers(row) -> dict | None:
    texts = [normalize_header(cell_text(c)) for c in row.cells]
    joined = "".join(texts)
    if "偏离" not in joined or "招标" not in joined:
        return None
    mapping = {}
    for i, t in enumerate(texts):
        if "序号" == t or t in {"号", "序号"}:
            mapping.setdefault("seq_in_table", i)
        elif "货物名称" in t or t == "名称":
            mapping.setdefault("name", i)
        elif "型号" in t:
            mapping.setdefault("model", i)
        elif "品牌" in t or "生产厂家" in t:
            mapping.setdefault("brand", i)
        elif "招标文件要求" in t or "招标" in t and "参数" in t:
            mapping.setdefault("req", i)
        elif "投标设备" in t or ("投标" in t and "参数" in t):
            mapping.setdefault("offered", i)
        elif t == "偏离情况" or t.endswith("偏离情况") or t == "偏离":
            mapping.setdefault("deviation", i)
        elif "单价" in t:
            mapping.setdefault("unit_price", i)
        elif t == "小计" or "小计" in t:
            mapping.setdefault("subtotal", i)
        elif "证明材料" in t or "页码" in t:
            mapping.setdefault("evidence_pages", i)
    if "req" in mapping and "deviation" in mapping:
        mapping["_headers"] = texts
        return mapping
    return None


def classify_dev(s: str) -> str | None:
    s = re.sub(r"\s+", "", s or "")
    if not s:
        return None
    if "负偏离" in s:
        return "负偏离"
    if "正偏离" in s:
        return "正偏离"
    if "无偏离" in s or s in {"响应", "满足", "符合"}:
        return "响应"  # 无偏离视为响应
    if "偏离" in s:
        return s
    return s


def main():
    print("opening V8 docx...")
    doc = Document(str(V8))
    print(f"tables={len(doc.tables)}")
    rows_out = []
    table_hits = 0
    for ti, tbl in enumerate(doc.tables):
        if not tbl.rows:
            continue
        # header may be first or first two rows
        mapping = map_headers(tbl.rows[0])
        header_row = 0
        if mapping is None and len(tbl.rows) > 1:
            mapping = map_headers(tbl.rows[1])
            header_row = 1
        if mapping is None:
            continue
        table_hits += 1
        current_name = None
        current_model = None
        current_brand = None
        for ri, row in enumerate(tbl.rows):
            if ri <= header_row:
                continue
            cells = [cell_text(c) for c in row.cells]
            def get(k):
                i = mapping.get(k)
                return cells[i] if i is not None and i < len(cells) else ""
            name = get("name") or current_name
            model = get("model") or current_model
            brand = get("brand") or current_brand
            if name:
                current_name = name
            if get("model"):
                current_model = get("model")
            if get("brand"):
                current_brand = get("brand")
            req = get("req")
            offered = get("offered")
            dev = get("deviation")
            if not (req or dev):
                continue
            # skip leftover header repeats
            if "招标文件要求" in re.sub(r"\s+", "", req) and "偏离" in re.sub(r"\s+", "", dev):
                continue
            rows_out.append(
                {
                    "table_index": ti,
                    "row_index": ri,
                    "device_name": current_name,
                    "model": current_model,
                    "brand": current_brand,
                    "req_text": req,
                    "offered_text": offered,
                    "deviation_raw": dev,
                    "deviation": classify_dev(dev),
                    "unit_price": get("unit_price"),
                    "subtotal": get("subtotal"),
                    "evidence_pages": get("evidence_pages"),
                }
            )
    # counts
    from collections import Counter
    c = Counter(r["deviation"] for r in rows_out)
    out = {
        "source": str(V8),
        "n_tables_matched": table_hits,
        "n_rows": len(rows_out),
        "deviation_counts": dict(c),
        "rows": rows_out,
    }
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"matched_tables={table_hits} rows={len(rows_out)} counts={dict(c)}")
    # sample
    for r in rows_out[:5]:
        print(" sample:", r["device_name"], r["deviation"], (r["req_text"] or "")[:40])


if __name__ == "__main__":
    main()
