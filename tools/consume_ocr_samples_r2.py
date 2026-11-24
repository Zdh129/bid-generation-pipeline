#!/usr/bin/env python3
"""第二轮识别侧样例：独立校验、核表格与纠偏矩阵、生成、排版检查。不调用外部 API。"""
from __future__ import annotations

import json
import math
import shutil
import sys
import time
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
from validate_schema import (  # noqa: E402
    DESKEW_ANGLE_TOL_DEG,
    angle_diff_deg,
    matrix_theta_deg,
    validate_file_full,
)
from generate_docx import generate  # noqa: E402
from check_layout import run_checks  # noqa: E402

OCR_DIR = Path(__file__).resolve().parents[1] / "data" / "ocr-samples-r2"
OUT = ROOT / "output" / "ocr-samples-r2"
# 按源页码排序消费
SAMPLES = [
    ("p02", "p02-scan-0cd9150d8db0-p2.json", "扫描页"),
    ("p04", "p04-scan-table-0cd9150d8db0-p4.json", "扫描表"),
    ("p12", "p12-scan-seal-3d7c31d236b9-p12.json", "扫描件+印章"),
    ("p17", "p17-text_layer-table-0cd9150d8db0-p17.json", "文字层表格页"),
    ("p20", "p20-text_layer-b7aeaca205b9-p20.json", "文字层页"),
]

THEIR_TABLE = {
    "p17": {"block": "b-017-009", "n_rows": 27, "n_cols": 9, "n_cells": 237, "overlap": 0, "holes": 0, "rowspan_gt1": 0, "colspan_gt1": 6},
    "p04": {"block": "b-004-002", "n_rows": 5, "n_cols": 9, "n_cells": 45, "overlap": 0, "holes": 0, "rowspan_gt1": 0, "colspan_gt1": 0},
}


def _now():
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def table_stats(t: dict) -> dict:
    nr, nc = t.get("n_rows"), t.get("n_cols")
    cells = t.get("cells") or []
    cov = [[0] * nc for _ in range(nr)] if nr and nc else []
    oob = rs_gt = cs_gt = 0
    for c in cells:
        r, col = c["row"], c["col"]
        rs, cs = c.get("rowspan") or 1, c.get("colspan") or 1
        if rs > 1:
            rs_gt += 1
        if cs > 1:
            cs_gt += 1
        if r + rs > nr or col + cs > nc or r < 0 or col < 0:
            oob += 1
            continue
        for i in range(r, r + rs):
            for j in range(col, col + cs):
                cov[i][j] += 1
    overlap = sum(1 for row in cov for v in row if v > 1)
    holes = sum(1 for row in cov for v in row if v == 0)
    coverage = (sum(1 for row in cov for v in row if v >= 1) / (nr * nc)) if nr and nc else None
    return {
        "n_rows": nr, "n_cols": nc, "n_cells": len(cells),
        "overlap": overlap, "holes": holes, "oob": oob,
        "rowspan_gt1": rs_gt, "colspan_gt1": cs_gt,
        "coverage": coverage,
    }


def rotation_matrix(theta_deg: float, cx: float, cy: float) -> list[float]:
    """[a,b,c,d,tx,ty] for x'=ax+by+tx, y'=cx+dy+ty, CCW θ about (cx,cy)."""
    th = math.radians(theta_deg)
    co, si = math.cos(th), math.sin(th)
    a, b, c, d = co, -si, si, co
    tx = -co * cx + si * cy + cx
    ty = -si * cx - co * cy + cy
    return [a, b, c, d, tx, ty]


def check_deskew(page: dict) -> dict:
    """按已确认约定核验：angle_deg = 矩阵实际旋转角 θ，残留角看 skew_angle_deg。"""
    tr = page.get("deskew_transform") or {}
    m = tr.get("matrix")
    center = tr.get("center")
    ang = tr.get("angle_deg")
    before = page.get("skew_angle_deg_before")
    after = page.get("skew_angle_deg")
    applied = page.get("deskew_applied")
    rec = {
        "page": page.get("page_number"),
        "deskew_applied": applied,
        "skew_angle_deg": after,
        "skew_angle_deg_before": before,
        "transform_type": tr.get("type"),
        "transform_angle_deg": ang,
        "center": center,
        "matrix": m,
        "matrix_theta_deg": None,
        "angle_vs_matrix_diff_deg": None,
        "angle_vs_matrix_ok": None,
        "consistent": None,
        "note": None,
    }
    if not m:
        rec["consistent"] = applied in (False, None) and not tr
        rec["note"] = "无 deskew_transform（本页未声明已应用旋转矩阵）"
        return rec
    if not (isinstance(m, list) and len(m) == 6):
        rec["consistent"] = False
        rec["note"] = "matrix 字段不完整，无法按 [[a,b,tx],[c,d,ty]] 求 θ"
        return rec
    try:
        theta = matrix_theta_deg(m)
    except (TypeError, ValueError):
        rec["consistent"] = False
        rec["note"] = "matrix 无法计算 atan2(matrix[1][0], matrix[0][0])"
        return rec
    rec["matrix_theta_deg"] = theta
    if ang is not None:
        diff = angle_diff_deg(theta, float(ang))
        rec["angle_vs_matrix_diff_deg"] = diff
        rec["angle_vs_matrix_ok"] = diff <= DESKEW_ANGLE_TOL_DEG
        rec["consistent"] = rec["angle_vs_matrix_ok"]
        rec["note"] = (
            f"angle_deg 按矩阵实际旋转角 θ 解读；"
            f"atan2(matrix[1][0], matrix[0][0])={theta:.6f}°，"
            f"angle_deg={float(ang):.6f}°，偏差 {diff:.3e}° "
            f"{'≤' if rec['angle_vs_matrix_ok'] else '>'} {DESKEW_ANGLE_TOL_DEG}°。"
        )
    else:
        rec["consistent"] = False
        rec["note"] = f"有 matrix，但缺少 angle_deg；矩阵 θ={theta:.6f}°。"
    if center and len(center) == 2 and ang is not None:
        cx, cy = float(center[0]), float(center[1])
        exp = rotation_matrix(float(ang), cx, cy)
        diffs = [abs(float(x) - y) for x, y in zip(m, exp)]
        rec["matrix_vs_rotation"] = {
            "theta_deg": float(ang),
            "max_abs_diff": max(diffs),
        }
        w, h = float(page.get("width") or 0), float(page.get("height") or 0)
        rec["center_vs_page_mid"] = {
            "page_mid": [round(w / 2, 2), round(h / 2, 2)] if w and h else None,
            "delta": [round(cx - w / 2, 3), round(cy - h / 2, 3)] if w and h else None,
        }
    if applied and after is not None and abs(float(after)) > 0.5:
        rec["consistent"] = False
        rec["note"] = (rec.get("note") or "") + " deskew_applied=true 但残留角大于 0.5°。"
    return rec


def validate_one(path: Path) -> dict:
    t0 = time.time()
    result = validate_file_full(path)
    errs = result["errors"]
    warns = result["warnings"]
    return {
        "file": path.name,
        "ok": not errs,
        "error_count": len(errs),
        "warning_count": len(warns),
        "schema_errors": [e for e in errs if e.startswith("schema:")],
        "semantic_errors": [e for e in errs if not e.startswith("schema:")],
        "semantic_warnings": warns,
        "elapsed_s": round(time.time() - t0, 3),
    }


def inspect_json(path: Path, key: str) -> dict:
    data = json.loads(path.read_text(encoding="utf-8"))
    pages = data.get("pages") or []
    blocks = data.get("blocks") or []
    tables = []
    for b in blocks:
        if b.get("type") == "table":
            st = table_stats(b.get("table") or {})
            st["block_id"] = b.get("id")
            st["verification_status"] = b.get("verification_status")
            expected = THEIR_TABLE.get(key)
            match = None
            if expected and expected.get("block") == b.get("id"):
                match = all(st.get(k) == expected[k] for k in expected if k != "block")
            st["matches_their_report"] = match
            st["their_report"] = expected
            tables.append(st)
    ro = [(b.get("reading_order"), b.get("type")) for b in sorted(blocks, key=lambda x: x.get("reading_order", 0))]
    headers = [r for r in ro if r[1] in ("header",)]
    footers = [r for r in ro if r[1] in ("footer", "page_number")]
    body = [r for r in ro if r[1] not in ("header", "footer", "page_number")]
    ro_ok = True
    note = None
    if headers and body and headers[-1][0] > body[0][0]:
        ro_ok = False
        note = "页眉 reading_order 未全部位于正文之前"
    if footers and body and min(r[0] for r in footers) < max(r[0] for r in body):
        ro_ok = False
        note = (note or "") + " 页脚/页码未全部位于正文之后"
    return {
        "page_count": data.get("document", {}).get("page_count"),
        "pages": [
            {
                "page_number": pg.get("page_number"),
                "width": pg.get("width"),
                "height": pg.get("height"),
                "unit": pg.get("unit"),
                "width_mm": pg.get("width_mm"),
                "height_mm": pg.get("height_mm"),
                "mm_from_pt": [
                    round(float(pg["width"]) * 25.4 / 72, 2) if pg.get("width") else None,
                    round(float(pg["height"]) * 25.4 / 72, 2) if pg.get("height") else None,
                ] if (pg.get("unit") or "pt") == "pt" else None,
                "render_dpi": pg.get("render_dpi"),
                "has_coordinate_space": bool(pg.get("coordinate_space")),
                "deskew": check_deskew(pg),
            }
            for pg in pages
        ],
        "block_types": dict(Counter(b.get("type") for b in blocks)),
        "verification": dict(Counter(b.get("verification_status") for b in blocks)),
        "n_assets": len(data.get("assets") or []),
        "tables": tables,
        "reading_order_header_footer_ok": ro_ok,
        "reading_order_note": note,
    }


def merge_samples(dest_json: Path) -> dict:
    pages, blocks, assets = [], [], []
    seen = set()
    for key, name, role in SAMPLES:
        src = OCR_DIR / name
        data = json.loads(src.read_text(encoding="utf-8"))
        pages.extend(data.get("pages") or [])
        blocks.extend(data.get("blocks") or [])
        for a in data.get("assets") or []:
            if a["id"] in seen:
                continue
            seen.add(a["id"])
            src_path = src.parent / a["path"]
            dest = dest_json.parent / a["path"]
            dest.parent.mkdir(parents=True, exist_ok=True)
            if src_path.is_file():
                shutil.copy2(src_path, dest)
            assets.append(a)
    merged = {
        "schema_version": "1.0.0",
        "data_label": "example",
        "needs_review": True,
        "document": {
            "source_file": "ocr-schema-v1-sample-r2-combined.json",
            "source_type": "mixed",
            "language": "zh-CN",
            "page_count": len(pages),
            "generated_at": _now(),
            "engine_name": "pdf-ocr-test-pipeline+gen-merge-r2",
            "engine_version": "0.2.0",
            "schema_version": "1.0.0",
            "render_dpi_default": 300,
            "notes": "第二轮五份摘页合并，保留原始 page_number。",
            "title": "识别侧中间格式样例合集（第二轮）",
        },
        "coordinate_system": {
            "origin": "top-left",
            "x_axis": "rightward",
            "y_axis": "downward",
            "default_unit": "pt",
        },
        "pages": pages,
        "blocks": blocks,
        "assets": assets,
        "warnings": [],
    }
    dest_json.write_text(json.dumps(merged, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return merged


def summarize_check(rep: dict) -> dict:
    s = rep["summary"]
    applicable = s["total"] - s.get("not_applicable", 0) - s.get("skipped", 0)
    failed = s.get("failed", 0)
    passed_applicable = applicable - failed
    return {
        "summary": s,
        "applicable": applicable,
        "passed_applicable": passed_applicable,
        "pass_rate_applicable": round(passed_applicable / applicable, 4) if applicable else None,
        "failed": [
            {"id": r["id"], "severity": r.get("severity"), "message": r.get("message")}
            for r in rep["results"]
            if not r.get("passed") and not r.get("not_applicable") and not r.get("skipped")
        ],
    }


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    t_all = time.time()
    inspections = {}
    validations = []
    t0 = time.time()
    for key, name, role in SAMPLES:
        path = OCR_DIR / name
        val = validate_one(path)
        val["role"] = role
        validations.append(val)
        inspections[key] = inspect_json(path, key)
    val_elapsed = round(time.time() - t0, 3)

    merged_path = OUT / "combined.json"
    t1 = time.time()
    merge_samples(merged_path)
    merge_val = validate_one(merged_path)
    options = {
        "header_text": "识别侧中间格式样例（第二轮）",
        "auto_toc": False,
        "force_toc": False,
        "page_breaks_between_pages": True,
        "review_appendix": True,
    }
    gen = generate(merged_path, OUT / "combined.docx", OUT / "combined.pdf", options)
    gen_elapsed = round(time.time() - t1, 3)

    per_page = []
    t2 = time.time()
    for key, name, role in SAMPLES:
        src = OCR_DIR / name
        g = generate(src, OUT / f"{key}.docx", OUT / f"{key}.pdf", options)
        per_page.append({"key": key, "role": role, "docx": str(OUT / f"{key}.docx"), "pdf": str(OUT / f"{key}.pdf"), "render_meta": g.get("render_meta")})
    per_elapsed = round(time.time() - t2, 3)

    t3 = time.time()
    checks = {}
    targets = [
        ("combined", OUT / "combined.docx", OUT / "combined.pdf"),
        ("p04_table", OUT / "p04.docx", OUT / "p04.pdf"),
        ("p17_table", OUT / "p17.docx", OUT / "p17.pdf"),
        ("p12_seal", OUT / "p12.docx", OUT / "p12.pdf"),
        ("p02_scan", OUT / "p02.docx", OUT / "p02.pdf"),
        ("p20_text", OUT / "p20.docx", OUT / "p20.pdf"),
    ]
    for label, docx, pdf in targets:
        try:
            rep = run_checks(None, str(docx), str(pdf) if pdf.is_file() else None)
        except Exception as exc:
            checks[label] = {"error": type(exc).__name__}
            continue
        checks[label] = summarize_check(rep)
        (OUT / f"check_{label}.json").write_text(json.dumps(rep, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    check_elapsed = round(time.time() - t3, 3)

    r1_ref = {
        "note": "第一轮合集/表格页/印章页（见 output/ocr-samples 与 docs/ocr-sample-consumption-report.md）",
        "combined": {"failed_ids": ["Q-FONT-002", "Q-HF-003"], "applicable_pass": "38/40", "rate": 0.95},
        "p03_table": {"failed_ids": ["Q-FONT-002", "Q-HEAD-002", "Q-HF-003", "Q-PGN-004"], "applicable_pass": "32/36", "rate": 0.89},
        "p04_seal": {"failed_ids": ["Q-FONT-002", "Q-HF-003"], "applicable_pass": "31/33", "rate": 0.94},
    }

    report = {
        "privacy": "不摘录样例正文、印章文字或源文件人名",
        "api_cost_yuan": 0,
        "validation": validations,
        "validation_elapsed_s": val_elapsed,
        "inspections": inspections,
        "merged_validation": merge_val,
        "generation": {
            "combined_docx": str(OUT / "combined.docx"),
            "combined_pdf": str(OUT / "combined.pdf"),
            "combined_docx_bytes": (OUT / "combined.docx").stat().st_size if (OUT / "combined.docx").is_file() else None,
            "combined_pdf_bytes": (OUT / "combined.pdf").stat().st_size if (OUT / "combined.pdf").is_file() else None,
            "elapsed_s": gen_elapsed,
            "render_meta": gen.get("render_meta"),
        },
        "per_page": per_page,
        "per_page_elapsed_s": per_elapsed,
        "layout_checks": checks,
        "check_elapsed_s": check_elapsed,
        "round1_reference": r1_ref,
        "elapsed_total_s": round(time.time() - t_all, 3),
    }
    (OUT / "consumption_meta.json").write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    (OUT / "validation-ours.json").write_text(
        json.dumps({"results": validations, "merged": merge_val, "tables_and_deskew": {
            k: {"tables": inspections[k]["tables"], "deskew": inspections[k]["pages"][0]["deskew"],
                "mm": [inspections[k]["pages"][0]["width_mm"], inspections[k]["pages"][0]["height_mm"]],
                "mm_from_pt": inspections[k]["pages"][0]["mm_from_pt"]}
            for k in inspections
        }}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    # compact stdout
    out = {
        "validation": [{k: r[k] for k in ("file", "ok", "error_count", "semantic_errors")} for r in validations],
        "merged_ok": merge_val["ok"],
        "tables": {k: inspections[k]["tables"] for k in inspections},
        "deskew": {k: inspections[k]["pages"][0]["deskew"] for k in inspections},
        "mm": {k: {"mm": [inspections[k]["pages"][0]["width_mm"], inspections[k]["pages"][0]["height_mm"]],
                   "from_pt": inspections[k]["pages"][0]["mm_from_pt"]} for k in inspections},
        "ro_ok": {k: inspections[k]["reading_order_header_footer_ok"] for k in inspections},
        "gen_elapsed_s": gen_elapsed,
        "render_tables": (gen.get("render_meta") or {}).get("tables"),
        "render_seals": (gen.get("render_meta") or {}).get("seals"),
        "render_images": (gen.get("render_meta") or {}).get("images"),
        "deskew_actions": (gen.get("render_meta") or {}).get("deskew"),
        "checks": {k: {"rate": v.get("pass_rate_applicable"), "app": f"{v.get('passed_applicable')}/{v.get('applicable')}",
                       "failed": [x["id"] for x in v.get("failed") or []]} if "error" not in v else v
                   for k, v in checks.items()},
        "elapsed_total_s": report["elapsed_total_s"],
    }
    print(json.dumps(out, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
