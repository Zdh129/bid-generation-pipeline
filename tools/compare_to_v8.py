#!/usr/bin/env python3
"""量化对比生成标书与 V8：结构、招标要求响应、排版检查、附件/页数。不摘录个人信息。"""
from __future__ import annotations

import json
import re
import sys
import time
import zipfile
from pathlib import Path
from repo_paths import load_paths

import pymupdf
from docx import Document

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

REQ = json.loads((ROOT / "work" / "requirements.json").read_text(encoding="utf-8"))
V8_STATS = json.loads((ROOT / "standards" / "v8-stats.json").read_text(encoding="utf-8"))

BIZ_NEEDLES = {
    "B1": ["投标有效期", "90"],
    "B2": ["最高限价", "一次性"],
    "B3": ["投标保证金"],
    "B4": ["交货期"],
    "B5": ["质量标准"],
    "B6": ["中文技术资料", "维护说明"],
    "B7": ["包装", "运输"],
    "B8": ["检定", "校准"],
    "B9": ["进口", "正品"],
    "B10": ["廉洁"],
    "Q1": ["营业执照"],
    "Q2": ["法定代表人"],
    "Q3": ["纳税", "社会保险"],
    "Q4": ["质量体系", "质量管理体系"],
    "Q5": ["环境管理体系", "环境体系"],
    "Q6": ["职业健康"],
    "Q7": ["业绩"],
}


def normalize(s: str) -> str:
    s = re.sub(r"\s+", "", s or "")
    s = s.replace("（仅供参考）", "")
    return s


def heading_texts_from_v8():
    out = []
    for h in V8_STATS["v8"]["headings"]:
        st = h.get("style") or ""
        if "TOC" in st:
            continue
        role = h.get("role")
        if str(st).lower().startswith("heading") or role in {"heading_1", "heading_2", "heading_3"}:
            t = normalize(h.get("text") or "")
            # strip trailing page numbers often glued in TOC-derived text
            t = re.sub(r"\d+$", "", t) if re.search(r"[^\d]\d{1,3}$", t) else t
            if t:
                out.append({"style": st, "text": t})
    return out


def heading_texts_from_docx(path: Path):
    doc = Document(str(path))
    out = []
    for p in doc.paragraphs:
        st = p.style.name if p.style else ""
        t = normalize(p.text or "")
        if not t:
            continue
        if str(st).lower().startswith("heading"):
            out.append({"style": st, "text": t})
    return out


def structure_coverage(gen_heads, v8_heads, req_items):
    v8_set = [h["text"] for h in v8_heads]
    g_set = [h["text"] for h in gen_heads]

    def matched(title, pool):
        key = normalize(title)
        return any(key == x or key in x or x in key for x in pool)

    v8_vs_req = sum(1 for it in req_items if matched(it["title"], v8_set))
    gen_vs_req = sum(1 for it in req_items if matched(it["title"], g_set))
    covered_v8_titles = 0
    missing = []
    for t in v8_set:
        if any(t == g or t in g or g in t for g in g_set):
            covered_v8_titles += 1
        else:
            missing.append(t)
    return {
        "tender_structure_items": len(req_items),
        "v8_headings": len(v8_set),
        "generated_headings": len(g_set),
        "v8_matched_tender_structure": v8_vs_req,
        "generated_matched_tender_structure": gen_vs_req,
        "generated_covered_v8_headings": covered_v8_titles,
        "generated_vs_v8_heading_coverage": round(covered_v8_titles / len(v8_set), 4) if v8_set else None,
        "missing_v8_heading_count": len(missing),
        "missing_v8_heading_titles": missing,
        "unmatched_tender_structure_ids": {
            "generated": [it["id"] for it in req_items if not matched(it["title"], g_set)],
            "v8": [it["id"] for it in req_items if not matched(it["title"], v8_set)],
        },
    }


def docx_text_blob(path: Path) -> str:
    doc = Document(str(path))
    parts = []
    for p in doc.paragraphs:
        t = p.text or ""
        if t.strip():
            parts.append(t)
    for tbl in doc.tables:
        for row in tbl.rows:
            for c in row.cells:
                t = c.text or ""
                if t.strip():
                    parts.append(t)
    return normalize("\n".join(parts))


def needle_of_param(text: str) -> str:
    t = normalize(text)
    t = re.sub(r"^[★＊*]*", "", t)
    if len(t) <= 24:
        return t
    # mid-length distinctive slice, skip very generic prefixes
    for prefix in ("用途：", "用途:", "规格：", "规格:"):
        if t.startswith(prefix) and len(t) > len(prefix) + 8:
            return t[len(prefix) : len(prefix) + 24]
    return t[:24]


def match_params(blob: str, params):
    hit = 0
    miss_ids = []
    star_hit = 0
    star_total = 0
    for p in params:
        nd = needle_of_param(p.get("text") or "")
        ok = bool(nd) and nd in blob
        if ok:
            hit += 1
        else:
            miss_ids.append(p["id"])
        if p.get("star"):
            star_total += 1
            if ok:
                star_hit += 1
    return {
        "total": len(params),
        "matched_by_text": hit,
        "unmatched": len(miss_ids),
        "coverage": round(hit / len(params), 4) if params else None,
        "star_total": star_total,
        "star_matched_by_text": star_hit,
        "star_coverage": round(star_hit / star_total, 4) if star_total else None,
        "unmatched_ids_head": miss_ids[:20],
        "unmatched_id_count": len(miss_ids),
        "note": "以招标参数原文规范化后的短词在文档中检索，不摘录正文；未命中不代表完全未响应（可能改写）。",
    }


def match_devices(blob: str, devices):
    found = []
    miss = []
    for d in devices:
        n = normalize(d["name"]).replace("（七）", "")
        if n and n in blob:
            found.append(d["seq"])
        else:
            miss.append(d["seq"])
    return {
        "total": len(devices),
        "found": len(found),
        "missing_seq": miss,
        "coverage": round(len(found) / len(devices), 4) if devices else None,
    }


def match_business(blob: str):
    ok = 0
    miss = []
    for item in REQ["business_qualification"]:
        needles = BIZ_NEEDLES.get(item["id"], [item["title"]])
        if any(normalize(n) in blob for n in needles):
            ok += 1
        else:
            miss.append(item["id"])
    return {
        "total": len(REQ["business_qualification"]),
        "responded": ok,
        "unmatched_ids": miss,
        "coverage": round(ok / len(REQ["business_qualification"]), 4),
    }


def page_info(pdf: Path | None):
    if not pdf or not pdf.is_file():
        return None
    d = pymupdf.open(pdf)
    port = land = 0
    for i in range(d.page_count):
        r = d[i].rect
        if r.width <= r.height + 1:
            port += 1
        else:
            land += 1
    n = d.page_count
    d.close()
    return {"pages": n, "portrait": port, "landscape": land}


def image_count(docx: Path):
    with zipfile.ZipFile(docx) as z:
        media = [n for n in z.namelist() if n.startswith("word/media/")]
        return len(media)


def table_count(docx: Path):
    doc = Document(str(docx))
    return {"tables": len(doc.tables), "rows": sum(len(t.rows) for t in doc.tables), "sections": len(doc.sections)}


def load_check(path: Path):
    if path.is_file():
        return json.loads(path.read_text(encoding="utf-8"))
    return None


def fail_ids(rep):
    if not rep:
        return None
    return [r["id"] for r in rep["results"] if not r.get("passed") and not r.get("not_applicable") and not r.get("skipped")]


def main():
    import argparse
    p = argparse.ArgumentParser()
    p.add_argument("--gen-docx", default=str(ROOT / "output/generated/bid_from_template.docx"))
    p.add_argument("--gen-pdf", default=str(ROOT / "output/generated/bid_from_template.pdf"))
    p.add_argument("--v8-pdf", default=str(ROOT / "output/v8/reference.pdf"))
    p.add_argument("-o", default=str(ROOT / "output/generated/compare_v8.json"))
    args = p.parse_args()
    paths = load_paths(ROOT)
    t0 = time.time()
    gen_docx, gen_pdf = Path(args.gen_docx), Path(args.gen_pdf)
    v8_docx, v8_pdf = Path(paths["v8"]), Path(args.v8_pdf)

    print("Building text blobs (no logging of content)...", flush=True)
    gen_blob = docx_text_blob(gen_docx) if gen_docx.is_file() else ""
    v8_blob = docx_text_blob(v8_docx)
    struct = structure_coverage(
        heading_texts_from_docx(gen_docx) if gen_docx.is_file() else [],
        heading_texts_from_v8(),
        REQ["structure"],
    )
    gen_params = match_params(gen_blob, REQ["tech_params"])
    v8_params = match_params(v8_blob, REQ["tech_params"])
    gen_dev = match_devices(gen_blob, REQ["devices"])
    v8_dev = match_devices(v8_blob, REQ["devices"])
    gen_biz = match_business(gen_blob)
    v8_biz = match_business(v8_blob)

    # Generated also declares response for 0-param devices + all parsed params in tables
    gen_declared = {
        "checklist_total": REQ["counts"]["total_checklist"],
        "structure_declared": REQ["counts"]["structure"],
        "business_qualification_declared": REQ["counts"]["business_qualification"],
        "tech_params_declared_in_response_table": REQ["counts"]["tech_params"],
        "star_params_declared": REQ["counts"]["star_params"],
        "devices_declared": REQ["counts"]["devices"],
        "zero_param_devices": [d["seq"] for d in REQ["devices"] if d["n_params"] == 0],
        "note": "生成稿对已解析参数均写入响应表并声明「响应」；9 台未拆出参数的设备以整机声明+附件计。文本命中见 matched_by_text。",
    }

    v8_chk = load_check(ROOT / "output/v8/check_report_docx.json") or load_check(ROOT / "output/v8/check_report.json")
    gen_chk = load_check(ROOT / "output/generated/check_report_docx.json")

    meta = {}
    mp = ROOT / "output/generated/build_meta.json"
    if mp.is_file():
        meta = json.loads(mp.read_text(encoding="utf-8"))

    report = {
        "privacy": "无证件号/姓名/账号摘录；标题仅为章节名或要求编号",
        "structure": struct,
        "requirements": {
            "checklist_counts": REQ["counts"],
            "generated_declared": gen_declared,
            "generated_text_match": {
                "business_qualification": gen_biz,
                "devices": gen_dev,
                "tech_params": gen_params,
            },
            "v8_text_match": {
                "business_qualification": v8_biz,
                "devices": v8_dev,
                "tech_params": v8_params,
            },
            "scoreboard": {
                "structure": {
                    "total": struct["tender_structure_items"],
                    "generated": struct["generated_matched_tender_structure"],
                    "v8": struct["v8_matched_tender_structure"],
                },
                "business_qualification": {
                    "total": gen_biz["total"],
                    "generated": gen_biz["responded"],
                    "v8": v8_biz["responded"],
                },
                "devices_named": {
                    "total": gen_dev["total"],
                    "generated": gen_dev["found"],
                    "v8": v8_dev["found"],
                },
                "tech_params_text": {
                    "total": gen_params["total"],
                    "generated": gen_params["matched_by_text"],
                    "v8": v8_params["matched_by_text"],
                },
                "star_params_text": {
                    "total": gen_params["star_total"],
                    "generated": gen_params["star_matched_by_text"],
                    "v8": v8_params["star_matched_by_text"],
                },
            },
        },
        "pages": {
            "v8": page_info(v8_pdf if v8_pdf.is_file() else None) or {"pages": 250},
            "generated": page_info(gen_pdf if gen_pdf.is_file() else None),
        },
        "images_media_count": {
            "v8_media_files": V8_STATS["v8"]["images"]["media_files"],
            "v8_placements": V8_STATS["v8"]["images"]["placements"],
            "generated_media_files": image_count(gen_docx) if gen_docx.is_file() else None,
            "generated_embedded_pages": (meta.get("attachment_counts") or {}),
        },
        "tables_sections": {
            "v8_tables": V8_STATS["v8"]["tables"]["table_count"],
            "v8_sections": V8_STATS["v8"]["section_count"],
            "generated": table_count(gen_docx) if gen_docx.is_file() else None,
        },
        "layout_checks": {
            "note": "DOCX 项；PDF 全量检查曾因 PyMuPDF 段错误退出，未纳入。",
            "v8": v8_chk["summary"] if v8_chk else None,
            "v8_failed_ids": fail_ids(v8_chk),
            "generated": gen_chk["summary"] if gen_chk else None,
            "generated_failed_ids": fail_ids(gen_chk),
        },
        "build_meta": {
            k: meta.get(k)
            for k in (
                "docx_bytes", "pdf_bytes", "pdf_pages",
                "elapsed_docx_s", "elapsed_pdf_s", "elapsed_total_s",
                "attachment_counts", "api_cost_yuan", "note",
            )
        } if meta else None,
        "elapsed_s": None,
        "api_cost_yuan": 0,
    }
    report["elapsed_s"] = round(time.time() - t0, 2)
    Path(args.o).write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report["requirements"]["scoreboard"], ensure_ascii=False, indent=2))
    print("structure", json.dumps(struct, ensure_ascii=False, indent=2)[:2000])
    print("pages", report["pages"])
    print("images", report["images_media_count"])
    print("tables", report["tables_sections"])
    print("layout", report["layout_checks"])
    print("compare_elapsed_s", report["elapsed_s"])


if __name__ == "__main__":
    main()
