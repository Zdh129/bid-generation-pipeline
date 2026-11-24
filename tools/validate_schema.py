#!/usr/bin/env python3
"""校验中间格式 JSON 是否符合 intermediate.schema.json，并做语义检查。"""
from __future__ import annotations

import argparse
import json
import math
import sys
from collections import Counter
from pathlib import Path

from jsonschema import Draft202012Validator, FormatChecker

ROOT = Path(__file__).resolve().parents[1]
SCHEMA_PATH = ROOT / "schema" / "intermediate.schema.json"

DESKEW_ANGLE_TOL_DEG = 0.01


def angle_diff_deg(a: float, b: float) -> float:
    """最短角差（度），处理 ±180/360 回绕。"""
    return abs((float(a) - float(b) + 180.0) % 360.0 - 180.0)


def matrix_theta_deg(matrix) -> float:
    """从打包矩阵求实际旋转角 θ（度）。

    matrix 约定为 [a,b,c,d,tx,ty]，对应 2×3 形式 [[a,b,tx],[c,d,ty]]，
    即 x' = a*x + b*y + tx，y' = c*x + d*y + ty。
    标准旋转：[[cosθ,-sinθ,tx],[sinθ,cosθ,ty]]。
    θ = atan2(matrix[1][0], matrix[0][0]) = atan2(c, a)，逆时针为正（左上原点，y 向下）。
    """
    a, b, c, d, tx, ty = (float(x) for x in matrix)
    mat = [[a, b, tx], [c, d, ty]]
    return math.degrees(math.atan2(mat[1][0], mat[0][0]))


def deskew_transform_warnings(tr: dict, page_number=None) -> list[str]:
    """angle_deg 必须等于矩阵实际旋转角；偏差 > 0.01° 记 warning，不判 schema 失败。"""
    if not isinstance(tr, dict):
        return []
    m = tr.get("matrix")
    ang = tr.get("angle_deg")
    if m is None or ang is None:
        return []
    if not (isinstance(m, list) and len(m) == 6):
        return []
    try:
        theta = matrix_theta_deg(m)
        diff = angle_diff_deg(theta, float(ang))
    except (TypeError, ValueError):
        return []
    if diff <= DESKEW_ANGLE_TOL_DEG:
        return []
    prefix = f"page {page_number} " if page_number is not None else ""
    return [
        f"{prefix}deskew_transform: atan2(matrix[1][0], matrix[0][0])={theta:.6f}° "
        f"与 angle_deg={float(ang):.6f}° 偏差 {diff:.6f}°，超过 {DESKEW_ANGLE_TOL_DEG}°"
    ]


def semantic_warnings(data: dict) -> list[str]:
    warns: list[str] = []
    for p in data.get("pages") or []:
        tr = p.get("deskew_transform")
        if tr:
            warns.extend(deskew_transform_warnings(tr, p.get("page_number")))
    return warns


def load_schema() -> dict:
    return json.loads(SCHEMA_PATH.read_text(encoding="utf-8"))


def semantic_errors(data: dict, json_path: Path | None = None) -> list[str]:
    errs: list[str] = []
    doc = data.get("document") or {}
    pages = data.get("pages") or []
    blocks = data.get("blocks") or []
    assets = data.get("assets") or []

    if doc.get("schema_version") and doc.get("schema_version") != data.get("schema_version"):
        errs.append("document.schema_version 与根级 schema_version 不一致")
    if doc.get("page_count") != len(pages):
        errs.append(f"document.page_count={doc.get('page_count')} 与 pages 长度 {len(pages)} 不一致")

    page_nums = [p.get("page_number") for p in pages]
    if any(n is None or not isinstance(n, int) or n < 1 for n in page_nums):
        errs.append("pages[].page_number 必须是从 1 起的正整数（可与源 PDF 页码一致，摘页不必从 1 重编）")
    if len(page_nums) != len(set(page_nums)):
        errs.append("pages[].page_number 有重复")
    # 摘页可保留原始页码，不再要求 1..N 连续。
    for p in pages:
        if p.get("render_dpi") is None:
            errs.append(f"page {p.get('page_number')} 缺少 render_dpi")
        if p.get("width", 1) <= 0 or p.get("height", 1) <= 0:
            errs.append(f"page {p.get('page_number')} 宽高必须为正")

    page_set = set(page_nums)
    bid = [b.get("id") for b in blocks]
    if len(bid) != len(set(bid)):
        dups = [k for k, v in Counter(bid).items() if v > 1]
        errs.append(f"blocks[].id 重复: {dups[:8]}")

    aid = [a.get("id") for a in assets]
    if len(aid) != len(set(aid)):
        errs.append("assets[].id 有重复")
    asset_map = {a.get("id"): a for a in assets}

    for a in assets:
        path = a.get("path")
        if json_path is not None and path:
            ap = (json_path.parent / path).resolve()
            if not ap.is_file():
                errs.append(f"asset {a.get('id')} 文件不存在: {path}")

    for b in blocks:
        prefix = f"block {b.get('id')}"
        if b.get("page") not in page_set:
            errs.append(f"{prefix} 引用了不存在的 page={b.get('page')}")
        if "verification_status" not in b:
            errs.append(f"{prefix} 缺少 verification_status")
        bbox = b.get("bbox")
        if isinstance(bbox, list) and len(bbox) == 4:
            x0, y0, x1, y1 = bbox
            if not (x1 > x0 and y1 > y0):
                errs.append(f"{prefix} bbox 须满足 x1>x0 且 y1>y0")
        if b.get("bbox_unit") == "px" and not b.get("bbox_dpi"):
            errs.append(f"{prefix} bbox_unit=px 时必须提供 bbox_dpi")
        if b.get("type") == "heading" and b.get("level") is None:
            errs.append(f"{prefix} heading 缺少 level")
        if b.get("type") == "table":
            t = b.get("table")
            if not t:
                errs.append(f"{prefix} table 类型缺少 table 对象")
            else:
                errs.extend(_table_errors(prefix, t))
        if b.get("type") in ("image", "seal"):
            img = b.get("image") or {}
            aid_ = img.get("asset_id")
            if not aid_:
                errs.append(f"{prefix} 缺少 image.asset_id")
            elif aid_ not in asset_map:
                errs.append(f"{prefix} asset_id={aid_} 不在 assets 中")
    return errs


def _table_errors(prefix: str, t: dict) -> list[str]:
    errs = []
    nr, nc = t.get("n_rows"), t.get("n_cols")
    if not nr or not nc:
        return [f"{prefix} n_rows/n_cols 无效"]
    occ = [[None] * nc for _ in range(nr)]
    for i, cell in enumerate(t.get("cells") or []):
        r, c = cell.get("row"), cell.get("col")
        rs, cs = cell.get("rowspan", 1) or 1, cell.get("colspan", 1) or 1
        if r is None or c is None or r < 0 or c < 0 or r >= nr or c >= nc:
            errs.append(f"{prefix} 单元格 {i} 行列越界 row={r} col={c}")
            continue
        if r + rs > nr or c + cs > nc:
            errs.append(f"{prefix} 单元格 ({r},{c}) rowspan/colspan 越界")
            continue
        for rr in range(r, r + rs):
            for cc in range(c, c + cs):
                if occ[rr][cc] is not None:
                    errs.append(f"{prefix} 单元格重叠于 ({rr},{cc})")
                else:
                    occ[rr][cc] = (r, c)
    return errs


def validate_instance(data: dict, json_path: Path | None = None, schema: dict | None = None) -> list[str]:
    schema = schema or load_schema()
    validator = Draft202012Validator(schema, format_checker=FormatChecker())
    errs = []
    for e in sorted(validator.iter_errors(data), key=lambda x: list(x.path)):
        loc = "/".join(str(p) for p in e.path) or "(root)"
        errs.append(f"schema: {loc}: {e.message}")
    errs.extend(semantic_errors(data, json_path))
    return errs


def validate_instance_full(data: dict, json_path: Path | None = None, schema: dict | None = None) -> dict:
    """errors 导致失败；warnings（如 angle_deg 与矩阵 θ 偏差）仅提示。"""
    return {
        "errors": validate_instance(data, json_path, schema),
        "warnings": semantic_warnings(data),
    }


def validate_file(path: Path, schema: dict | None = None) -> list[str]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        return [f"无法解析 JSON: {exc}"]
    return validate_instance(data, path, schema)


def validate_file_full(path: Path, schema: dict | None = None) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        return {"errors": [f"无法解析 JSON: {exc}"], "warnings": []}
    return validate_instance_full(data, path, schema)


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description="校验中间格式 JSON")
    p.add_argument("paths", nargs="+", help="JSON 文件或目录")
    args = p.parse_args(argv)
    schema = load_schema()
    files: list[Path] = []
    for raw in args.paths:
        path = Path(raw)
        if path.is_dir():
            files.extend(sorted(path.rglob("*.json")))
        else:
            files.append(path)
    # skip checks.json etc. if someone passes a tree: only validate intended files
    rc = 0
    total = 0
    warn_files = 0
    for f in files:
        result = validate_file_full(f, schema)
        errs, warns = result["errors"], result["warnings"]
        total += 1
        if errs:
            rc = 1
            print(f"FAIL {f}  ({len(errs)} 个问题)")
            for e in errs:
                print(f"  - {e}")
        else:
            print(f"PASS {f}")
        if warns:
            warn_files += 1
            print(f"WARN {f}  ({len(warns)} 条警告，不判失败)")
            for w in warns:
                print(f"  - {w}")
    print(f"共 {total} 个文件，状态={'通过' if rc==0 else '失败'}，警告文件 {warn_files} 个")
    return rc


if __name__ == "__main__":
    sys.exit(main())
