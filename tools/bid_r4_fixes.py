"""r4 修复：偏离结论按「设备序号 + 行序」稳定唯一键对位，并在生成前、生成后各做一次逐行断言。

r3 问题根因：生成器用 f"P{seq:02d}-{n}" 当键，而判定文件的 id 是 P{seq:02d}-{行序 index}；
n 是招标原文条款号（含 "12.1" 这类子号，也会在同一设备内重复，如 P07 有两个 3），二者在 77 行上不同，
其中 45 行取到了错误判定或取不到（回退为「无证据待确认」）。
r4 键：(device_seq, index)，index = 该设备参数表中的行序（1 起）。判定规则与判定文件均未改动。
"""
from __future__ import annotations

import json
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
JUDGE_JSON = ROOT / "work" / "deviation_judgements.json"
TEXT_MAX = 220  # 与 r3 表格截断长度一致


class VerdictJoinError(AssertionError):
    pass


def _norm(s: str) -> str:
    return " ".join((s or "").split())


def load_judgement_index() -> dict[tuple[int, int], dict]:
    data = json.loads(JUDGE_JSON.read_text(encoding="utf-8"))
    by_dev = defaultdict(list)
    for j in data["judgements"]:
        by_dev[int(j["device_seq"])].append(j)
    idx = {}
    for seq, js in by_dev.items():
        for k, j in enumerate(js, 1):
            key = (seq, k)
            if key in idx:
                raise VerdictJoinError(f"判定键重复 {key}")
            idx[key] = j
    return idx


def stable_key(seq: int, p: dict, row_no: int) -> tuple[int, int]:
    """稳定唯一键：设备序号 + 参数在设备内的行序。devices.json 自带 index 时必须与行序一致。"""
    if p.get("index") is not None and int(p["index"]) != row_no:
        raise VerdictJoinError(f"设备 {seq} 第 {row_no} 行 index={p['index']} 与行序不一致")
    return (int(seq), row_no)


def build_param_rows(devices: list[dict], cite_evidence) -> tuple[dict[int, list[tuple]], list[dict]]:
    """返回 {seq: rows}，rows=(序号, 参数文本, 结论, 证据)；同时返回对位记录 join_log。
    生成前断言：每行参数文本与判定文件 req_text 完全一致，否则抛 VerdictJoinError。"""
    idx = load_judgement_index()
    used = set()
    out, join_log, errors = {}, [], []
    for d in devices:
        seq = int(d["seq"])
        rows = []
        for row_no, p in enumerate(d.get("params") or [], 1):
            key = stable_key(seq, p, row_no)
            j = idx.get(key)
            if j is None:
                errors.append(f"{key} 无判定")
                continue
            if _norm(j["req_text"]) != _norm(p["text"]) or str(j.get("param_n")) != str(p.get("n")):
                errors.append(f"{key} 参数文本/条款号与判定不一致：{p.get('n')} vs {j.get('param_n')}")
                continue
            used.add(key)
            star = "★" if p.get("star") else ""
            verdict = j.get("verdict") or "无证据待确认"
            ev = cite_evidence(j)
            rows.append((f"{star}{p['n']}", p["text"][:TEXT_MAX], verdict, ev))
            join_log.append({"key": f"{seq}#{row_no}", "judgement_id": j["id"], "clause_n": str(p.get("n")),
                             "r3_key": f"P{seq:02d}-{p.get('n')}", "verdict": verdict,
                             "v8": j.get("v8_deviation")})
        out[seq] = rows
    missing = set(idx) - used
    if missing:
        errors.append(f"判定文件中有 {len(missing)} 条未被使用：{sorted(missing)[:5]}")
    if errors:
        raise VerdictJoinError("偏离结论对位失败：" + "；".join(errors[:10]))
    return out, join_log


def audit_docx_param_tables(docx_path: Path, devices: list[dict], cite_evidence) -> dict:
    """生成后断言：从 DOCX 读回 45 张参数表，逐行核对 参数文本 / 结论 / 证据 与判定文件一一对应。
    任一行不对应即抛 VerdictJoinError。同时统计与 V8 的一致率（正文实际值）。"""
    from docx import Document
    idx = load_judgement_index()
    doc = Document(str(docx_path))
    tables = [t for t in doc.tables
              if len(t.rows) and [c.text.strip() for c in t.rows[0].cells][:4] == ["序号", "招标参数要求", "偏离结论", "引用证据"]]
    devs = [d for d in devices if d.get("params")]
    if len(tables) != len(devs):
        raise VerdictJoinError(f"参数表数量 {len(tables)} ≠ 有参数的设备数 {len(devs)}")
    rows_checked, bad, agree = 0, [], 0
    dist, v8dist = Counter(), Counter()
    disagree = []
    for t, d in zip(tables, devs):
        seq = int(d["seq"])
        body = t.rows[1:]
        if len(body) != len(d["params"]):
            raise VerdictJoinError(f"设备 {seq} 表行数 {len(body)} ≠ 参数数 {len(d['params'])}")
        for row_no, (r, p) in enumerate(zip(body, d["params"]), 1):
            cells = [c.text for c in r.cells]
            j = idx[(seq, row_no)]
            exp = (_norm(p["text"][:TEXT_MAX]), j.get("verdict") or "无证据待确认", _norm(cite_evidence(j)))
            got = (_norm(cells[1]), cells[2].strip(), _norm(cells[3]))
            rows_checked += 1
            if exp != got or _norm(j["req_text"]) != _norm(p["text"]):
                bad.append({"key": f"{seq}#{row_no}", "expected": exp, "got": got})
            dist[got[1]] += 1
            v8 = j.get("v8_deviation")
            v8dist[v8] += 1
            if got[1] == v8:
                agree += 1
            else:
                disagree.append({"key": f"{seq}#{row_no}", "judgement_id": j["id"], "clause_n": str(p.get("n")),
                                 "star": bool(p.get("star")), "docx_verdict": got[1], "v8": v8,
                                 "req_text": p["text"][:60]})
    rep = {"tables": len(tables), "rows_checked": rows_checked, "row_mismatch": len(bad), "bad_rows": bad[:20],
           "docx_agree_v8": agree, "rate": round(agree / rows_checked, 4) if rows_checked else None,
           "docx_verdict_dist": dict(dist), "v8_dist": dict(v8dist), "disagreements": disagree}
    if bad:
        raise VerdictJoinError(f"生成后断言失败：{len(bad)} 行结论/证据与原参数不对应，例 {bad[:2]}")
    return rep
