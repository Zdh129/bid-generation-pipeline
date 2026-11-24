#!/usr/bin/env python3
"""只读消费识别侧 v1.0.0 四页样例：校验、生成 DOCX/PDF、排版检查。不调用外部 API。"""
from __future__ import annotations

import json
import shutil
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
from validate_schema import validate_file, load_schema  # noqa: E402
from generate_docx import generate  # noqa: E402
from check_layout import run_checks  # noqa: E402

OCR_DIR = Path(__file__).resolve().parents[1] / "data" / "ocr-samples"
OUT = ROOT / "output" / "ocr-samples"
SAMPLES = [
    ("p01", "p01-text_layer-b7aeaca205b9-p20.json", "文字层页"),
    ("p02", "p02-scan-0cd9150d8db0-p2.json", "扫描页"),
    ("p03", "p03-table-0cd9150d8db0-p17.json", "表格页"),
    ("p04", "p04-seal-3d7c31d236b9-p12.json", "扫描件+印章"),
]


def _now():
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def validate_one(path: Path) -> dict:
    t0 = time.time()
    errs = validate_file(path)
    schema_errs = [e for e in errs if e.startswith("schema:")]
    semantic = [e for e in errs if not e.startswith("schema:")]
    return {
        "file": path.name,
        "ok": not errs,
        "error_count": len(errs),
        "schema_errors": schema_errs,
        "semantic_errors": semantic,
        "elapsed_s": round(time.time() - t0, 3),
    }


def merge_samples(dest_json: Path, assets_dest: Path) -> dict:
    assets_dest.mkdir(parents=True, exist_ok=True)
    pages, blocks, assets = [], [], []
    seen_assets = set()
    for key, name, role in SAMPLES:
        src = OCR_DIR / name
        data = json.loads(src.read_text(encoding="utf-8"))
        for pg in data.get("pages") or []:
            pages.append(pg)
        for b in data.get("blocks") or []:
            blocks.append(b)
        for a in data.get("assets") or []:
            if a["id"] in seen_assets:
                continue
            seen_assets.add(a["id"])
            src_path = src.parent / a["path"]
            rel = a["path"]
            dest = assets_dest.parent / rel
            dest.parent.mkdir(parents=True, exist_ok=True)
            if src_path.is_file():
                shutil.copy2(src_path, dest)
            assets.append(a)
    merged = {
        "schema_version": "1.0.0",
        "data_label": "example",
        "needs_review": True,
        "document": {
            "source_file": "ocr-schema-v1-sample-combined.json",
            "source_type": "mixed",
            "language": "zh-CN",
            "page_count": len(pages),
            "generated_at": _now(),
            "engine_name": "pdf-ocr-test-pipeline+gen-merge",
            "engine_version": "0.1.0",
            "schema_version": "1.0.0",
            "render_dpi_default": 300,
            "notes": "由识别侧四份摘页样例合并，仅供生成侧消费测试。保留原始 page_number。",
            "title": "识别侧中间格式样例合集",
        },
        "coordinate_system": {
            "origin": "top-left",
            "x_axis": "rightward",
            "y_axis": "downward",
            "default_unit": "pt",
            "description": "页面左上原点，x 向右、y 向下，默认单位 pt。换算：pt = px × 72 / render_dpi。",
        },
        "pages": pages,
        "blocks": blocks,
        "assets": assets,
        "warnings": [],
    }
    dest_json.write_text(json.dumps(merged, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return merged


def fail_ids(rep):
    if not rep:
        return []
    return [r["id"] for r in rep["results"] if not r.get("passed") and not r.get("not_applicable") and not r.get("skipped")]


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    t_all = time.time()
    # 1) validate four originals
    t0 = time.time()
    val = [validate_one(OCR_DIR / name) for _, name, _ in SAMPLES]
    val_elapsed = round(time.time() - t0, 3)

    # 2) merge + generate
    merged_path = OUT / "combined.json"
    t1 = time.time()
    merge_samples(merged_path, OUT / "assets")
    merge_val = validate_one(merged_path)
    options = {
        "header_text": "识别侧中间格式样例（生成测试）",
        "auto_toc": False,
        "force_toc": False,
        "page_breaks_between_pages": True,
        "review_appendix": True,
    }
    # disable auto toc: generate() still uses headings>=3 internally; combined may have 0 headings
    gen = generate(merged_path, OUT / "combined.docx", OUT / "combined.pdf", options)
    gen_elapsed = round(time.time() - t1, 3)

    # per-page optional (no pdf to save time unless useful) — also generate pdf for p03/p04 focus
    per_page = []
    t2 = time.time()
    for key, name, role in SAMPLES:
        src = OCR_DIR / name
        # generate from original path so assets resolve next to OCR json
        # copy json+assets already in OUT; use original for path relative assets
        out_d = OUT / f"{key}.docx"
        out_p = OUT / f"{key}.pdf"
        g = generate(src, out_d, out_p, options)
        per_page.append({"key": key, "role": role, "docx": str(out_d), "pdf": str(out_p), "render_meta": g.get("render_meta")})
    per_elapsed = round(time.time() - t2, 3)

    # 3) layout checks
    t3 = time.time()
    checks = {}
    # combined: docx + pdf (only 4+1 pages, safer than 250)
    for label, docx, pdf in (
        ("combined", OUT / "combined.docx", OUT / "combined.pdf"),
        ("p03_table", OUT / "p03.docx", OUT / "p03.pdf"),
        ("p04_seal", OUT / "p04.docx", OUT / "p04.pdf"),
    ):
        try:
            rep = run_checks(None, str(docx), str(pdf) if pdf.is_file() else None)
        except Exception as exc:
            checks[label] = {"error": type(exc).__name__, "message": "检查异常（未写入正文）"}
            continue
        checks[label] = {
            "summary": rep["summary"],
            "failed": [
                {"id": r["id"], "severity": r.get("severity"), "message": r.get("message")}
                for r in rep["results"]
                if not r.get("passed") and not r.get("not_applicable") and not r.get("skipped")
            ],
        }
        (OUT / f"check_{label}.json").write_text(json.dumps(rep, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    check_elapsed = round(time.time() - t3, 3)

    report = {
        "privacy": "不摘录样例正文、印章文字或源文件人名",
        "api_cost_yuan": 0,
        "validation_after_rule_fix": val,
        "validation_elapsed_s": val_elapsed,
        "merged_validation": merge_val,
        "generation": {
            "combined_docx": str(OUT / "combined.docx"),
            "combined_pdf": str(OUT / "combined.pdf"),
            "combined_docx_bytes": (OUT / "combined.docx").stat().st_size if (OUT / "combined.docx").is_file() else None,
            "combined_pdf_bytes": (OUT / "combined.pdf").stat().st_size if (OUT / "combined.pdf").is_file() else None,
            "elapsed_s": gen_elapsed,
            "render_meta": gen.get("render_meta"),
        },
        "per_page": [{"key": x["key"], "role": x["role"], "docx": x["docx"], "pdf": x["pdf"], "render_meta": x["render_meta"]} for x in per_page],
        "per_page_elapsed_s": per_elapsed,
        "layout_checks": checks,
        "check_elapsed_s": check_elapsed,
        "elapsed_total_s": round(time.time() - t_all, 3),
    }
    (OUT / "consumption_meta.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (OUT / "validation-ours.json").write_text(
        json.dumps({"baseline_note": "修订页码连续性规则后的结果；p03 单元格重叠仍为语义错误", "results": val, "merged": merge_val}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps({
        "validation": [{k: r[k] for k in ("file", "ok", "error_count", "semantic_errors")} for r in val],
        "merged_ok": merge_val["ok"],
        "gen_elapsed_s": gen_elapsed,
        "render_tables": (gen.get("render_meta") or {}).get("tables"),
        "render_seals": (gen.get("render_meta") or {}).get("seals"),
        "render_images": (gen.get("render_meta") or {}).get("images"),
        "review": (gen.get("render_meta") or {}).get("review_emitted"),
        "deskew": (gen.get("render_meta") or {}).get("deskew"),
        "checks": {k: (v.get("summary") if isinstance(v, dict) else v) for k, v in checks.items()},
        "failed": {k: v.get("failed") if isinstance(v, dict) else v for k, v in checks.items()},
        "elapsed_total_s": report["elapsed_total_s"],
        "check_elapsed_s": check_elapsed,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
