"""仓库内路径与项目信息加载。优先 config/*.json，否则用 example 文件。"""
from __future__ import annotations

import json
from pathlib import Path


def _load(root: Path, name: str) -> dict:
    for cand in (root / "config" / f"{name}.json", root / "config" / f"{name}.example.json"):
        if cand.exists():
            return json.loads(cand.read_text(encoding="utf-8"))
    return {}


def load_paths(root: Path) -> dict:
    data = _load(root, "paths")
    out = {}
    for k, v in data.items():
        if k in {"comment", "$comment"} or not isinstance(v, str) or not v:
            continue
        p = Path(v)
        out[k] = str(p if p.is_absolute() else (root / p).resolve())
    return out


def load_project(root: Path) -> dict:
    defaults = {
        "project_name": "示例采购项目",
        "project_code": "EXAM-ZB-20260001",
        "package_no": "1",
        "bidder_name": "示例投标人有限公司",
        "addressee": "示例招标人有限公司",
        "package_ceiling": 100000.0,
        "delivery": "合同签订之后，接招标人通知之日起 30 日历日内完成供货，并通过验收",
        "warranty": "验收合格之日起 1 年",
    }
    data = _load(root, "project")
    defaults.update({k: v for k, v in data.items() if not str(k).startswith("$")})
    return defaults
