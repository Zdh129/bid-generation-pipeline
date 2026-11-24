#!/usr/bin/env python3
"""从招标附件 3 PDF 重抽并解析 45 台设备技术参数。

修复第一轮因折行/跨页、以及无编号条款（「★键：值」或「键：值」）导致
设备 16、25、32、36、38、39、41、42、44 解析出 0 条参数的问题。
仅本地 PyMuPDF，不调用外部 API。
"""
from __future__ import annotations

import json
import re
import shutil
from pathlib import Path

import pymupdf

ROOT = Path(__file__).resolve().parents[1]
ATT3 = Path(
    "data/example-package"
    "/1. 官方招标文件/annex3-example-params.pdf"
)
WORK = ROOT / "work"
REQ_PATH = WORK / "requirements.json"
DEV_PATH = WORK / "devices.json"

CN = {"零": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}


def cn_to_int(s: str) -> int:
    s = s.strip()
    if s == "十":
        return 10
    if s.startswith("十"):
        return 10 + CN[s[1:]]
    if "十" in s:
        a, b = s.split("十", 1)
        return CN[a] * 10 + (CN[b] if b else 0)
    return CN[s]


HEADER_RE = re.compile(r"^（([一二三四五六七八九十百零]+)）\s*(.*)$")
# ★1、 text / 1. text / ★13 text (号与正文仅空格) / 12.1 text
NUM_ITEM_RE = re.compile(
    r"^(★?)\s*(\d+)(?:\.(\d+))?(?:\s*[、.．]\s*|\s+)(.*)$"
)
# 页边界丢失数字后残行：、压盘直径：120mm
LOST_NUM_RE = re.compile(r"^[、.．]\s*(\S.*)$")
# 无编号星标条款
STAR_UNNUM_RE = re.compile(r"^★\s*(.+)$")
# 短键名 + 冒号
KV_RE = re.compile(r"^([^：:]{1,40})[：:]\s*(.*)$")

SKIP_LABELS = {
    "技术参数",
    "技术参数：",
    "技术参数及性能",
    "技术参数及性能：",
    "配置",
    "配置：",
    "配置清单",
    "配制清单",
    "配置清单：",
    "配制清单：",
    "参数及功能",
    "参数及功能：",
    "仪器配置",
    "仪器最低配置",
    "仪器最低配置：",
}

# 单独成行的子序号，如「4.2」下一行才是条文
SUB_ONLY_RE = re.compile(r"^(★?)\s*(\d+)\.(\d+)\s*(.*)$")
COMPLETE_UNITS = (
    "℃", "°C", "℉", "%", "％", "L", "mL", "ml", "mm", "nm", "cm", "kg", "Kg", "g",
    "bar", "psi", "Hz", "W", "kW", "kw", "min", "s", "h", "J", "kN", "NTU", "Abs",
    "LX", "lx", "V", "A", "Pa", "KPa", "kPa", "MPa",
)
CONT_PREFIX = ("、", "，", ",", "；", ";", "）", ")", "的", "和", "及", "或", "与", "等", "醇", "子水")


def extract_lines(pdf_path: Path) -> list[tuple[int, str]]:
    """按页抽出非空行，保留页码以便跨页拼接，不在页边界丢行首编号。"""
    doc = pymupdf.open(pdf_path)
    rows: list[tuple[int, str]] = []
    for i, page in enumerate(doc):
        for raw in page.get_text("text").splitlines():
            line = raw.replace("\u00a0", " ").replace("\ufffd", "").strip()
            if line:
                rows.append((i + 1, line))
    doc.close()
    return rows


def split_device_sections(rows: list[tuple[int, str]]):
    """从第一个（一）起按中文序号切段。"""
    sections = []
    cur = None
    started = False
    for page, line in rows:
        # 修正「（七）（七）智能…」这类重复序号
        line_n = re.sub(r"^(（[一二三四五六七八九十百零]+）)\1+", r"\1", line)
        m = HEADER_RE.match(line_n)
        if m:
            seq = cn_to_int(m.group(1))
            if 1 <= seq <= 45:
                started = True
                if cur is not None:
                    sections.append(cur)
                name = re.sub(r"\s+", "", m.group(2)).strip()
                cur = {"seq": seq, "name": name, "page_start": page, "lines": []}
                continue
        if not started:
            continue
        if cur is not None:
            cur["lines"].append((page, line))
    if cur is not None:
        sections.append(cur)
    return sections


def _is_skip_label(line: str) -> bool:
    s = line.strip()
    s2 = s.replace(" ", "")
    return s in SKIP_LABELS or s2 in {x.replace(" ", "") for x in SKIP_LABELS}


def _looks_complete(text: str) -> bool:
    t = text.rstrip()
    if not t:
        return False
    return t[-1] in "。；;．."


def _value_complete(text: str) -> bool:
    t = text.rstrip()
    if not t:
        return False
    if t[-1] in "。；;．.":
        return True
    if any(t.endswith(u) for u in COMPLETE_UNITS):
        return True
    if re.search(r"[≤≧≥≦±]\s*[\d.]+(?:\s*[/～~\-]\s*[\d.]+)?\s*$", t):
        return True
    return False


def _is_label_only(text: str) -> bool:
    t = text.strip()
    return bool(re.fullmatch(r"[^：:]{1,20}[：:]", t))


def _is_unnumbered_new_item(prev: str, curr: str, curr_star: bool) -> bool:
    if curr_star:
        return True
    if curr.startswith(CONT_PREFIX):
        return False
    if prev.endswith(("：", ":", "（", "(", "包括", "如", "、", "，", ",")):
        return False
    if KV_RE.match(curr):
        return True
    if re.match(r"^(定值运行|波动度|显示精度|温度均匀性|温度范围|使用温度|容积|准确性|重复性|检测通道|工作温度)", curr):
        return True
    if re.match(r"^.{1,24}[≤≧≥≦±]", curr):
        return True
    if _value_complete(prev):
        return True
    return False


def parse_section_lines(lines: list[tuple[int, str]]) -> list[dict]:
    nonempty = [(p, ln) for p, ln in lines if ln.strip()]
    has_numbered = any(
        (NUM_ITEM_RE.match(ln) or SUB_ONLY_RE.match(ln)) and not _is_skip_label(ln)
        for _, ln in nonempty
    )

    items: list[dict] = []

    def start_item(star: bool, n, text: str, page: int, kind: str):
        items.append(
            {
                "n": n,
                "star": star,
                "text": text.strip(),
                "page": page,
                "kind": kind,
            }
        )

    def append_text(extra: str):
        if not items:
            return
        extra = extra.strip()
        if not extra:
            return
        prev = items[-1]["text"]
        join = ""
        items[-1]["text"] = prev + join + extra

    last_major = 0
    last_minor_of = {}

    if has_numbered:
        for page, ln in nonempty:
            if _is_skip_label(ln):
                continue
            # 优先匹配 4.2 / 4.2 正文（避免把 4.2 当成「4」+「.2」失败）
            msub = SUB_ONLY_RE.match(ln)
            m = NUM_ITEM_RE.match(ln)
            # 「0.15％」这类续行：上一行以 ±、.、≤ 结尾
            if items and items[-1]["text"].endswith(("±", ".", "≤", "≥", "~", "～", "-")):
                if re.match(r"^\d", ln) and not (msub and (msub.group(4) or "").strip()):
                    # 纯数值续行，且不是「4.2 电导率…」这种有正文的子项
                    if not (msub and msub.group(4).strip() and not re.match(r"^[\d.]+(?:％|%|μl|ul|mm|nm)", msub.group(4).strip())):
                        if re.match(r"^[\d.]+(?:％|%|μl|ul|ml|mL|mm|nm)", ln) or ln.startswith("0."):
                            append_text(ln)
                            continue
            if msub and (msub.group(3) is not None):
                star, major, minor, rest = msub.group(1), int(msub.group(2)), int(msub.group(3)), msub.group(4)
                # 0.15％ 是上一行「≤±」的跨页小数续写，不是第 0.15 条
                if major == 0 or (items and items[-1]["text"].endswith(("±", "≤", "≥", ".", "-", "～", "~"))):
                    append_text(ln)
                    continue
                n = f"{major}.{minor}"
                last_major = major
                last_minor_of[major] = minor
                start_item(bool(star), n, rest, page, "sub")
                continue
            if m:
                star, major, minor, rest = m.group(1), int(m.group(2)), m.group(3), m.group(4)
                if minor is not None:
                    n = f"{major}.{int(minor)}"
                    last_major = major
                    last_minor_of[major] = int(minor)
                    start_item(bool(star), n, rest, page, "sub")
                else:
                    n = major
                    last_major = major
                    start_item(bool(star), n, rest, page, "numbered")
                continue
            m = LOST_NUM_RE.match(ln)
            if m:
                guess = last_major + 1 if last_major else None
                start_item(False, guess if guess else "?", m.group(1), page, "lost_num")
                if guess:
                    last_major = guess
                continue
            # 编号项若只是「分辨率：」「测量范围：」这类标签，后续键值收入其子项
            if items and (_is_label_only(items[-1]["text"]) or items[-1]["kind"] in ("sub", "label_sub")):
                mkv = KV_RE.match(ln)
                if mkv or items[-1]["kind"] in ("sub", "label_sub") and not NUM_ITEM_RE.match(ln):
                    parent = last_major
                    if _is_label_only(items[-1]["text"]) or items[-1]["kind"] == "numbered" and _is_label_only(items[-1]["text"]):
                        nxt = last_minor_of.get(parent, 0) + 1
                        last_minor_of[parent] = nxt
                        start_item(items[-1]["star"], f"{parent}.{nxt}", ln, page, "label_sub")
                        # 标签行本身若无实质内容，可保留为父项标题
                    else:
                        # 已是子项：若当前行是同级键值，再开新子项；否则续写
                        if mkv and _value_complete(items[-1]["text"]):
                            nxt = last_minor_of.get(parent, 0) + 1
                            last_minor_of[parent] = nxt
                            start_item(items[-1]["star"], f"{parent}.{nxt}", ln, page, "label_sub")
                        else:
                            append_text(ln)
                    continue
            if items:
                append_text(ln)
            else:
                mkv = KV_RE.match(ln)
                start_item(ln.startswith("★"), 1, re.sub(r"^★\s*", "", ln), page, "kv" if mkv else "plain")
                last_major = 1
        return _normalize_item_numbers(items)

    # 纯无编号条款
    for page, ln in nonempty:
        if _is_skip_label(ln):
            continue
        is_star = False
        body = ln
        mstar = STAR_UNNUM_RE.match(ln)
        if mstar:
            is_star = True
            body = mstar.group(1).strip()
        if not items:
            start_item(is_star, 1, body, page, "star_unnum" if is_star else ("kv" if KV_RE.match(body) else "plain"))
            continue
        if _is_unnumbered_new_item(items[-1]["text"], body, is_star):
            start_item(is_star, len(items) + 1, body, page, "star_unnum" if is_star else ("kv" if KV_RE.match(body) else "plain"))
        else:
            if is_star:
                items[-1]["star"] = True
            append_text(body)
    for i, it in enumerate(items, 1):
        it["n"] = i
    return items


def _normalize_item_numbers(items: list[dict]) -> list[dict]:
    """把 u* / ? 编号改成不冲突的可读编号，保留已有数字和 12.1。"""
    used = set()
    for it in items:
        n = it["n"]
        if isinstance(n, int) or (isinstance(n, str) and re.fullmatch(r"\d+(\.\d+)?", str(n))):
            used.add(str(n))
    auto = 1
    for it in items:
        n = it["n"]
        if n == "?" or (isinstance(n, str) and str(n).startswith("u")):
            while str(auto) in used:
                auto += 1
            it["n"] = auto
            used.add(str(auto))
            auto += 1
    return items


def clean_text(s: str) -> str:
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\s+([，,。；;：:）)])", r"\1", s)
    return s.strip()



def drop_empty_parents(items: list[dict]) -> list[dict]:
    """去掉「测量范围：」「精度：」这类空标签父项（其子项已展开）。

    若父项带 ★，星标继承到子项（如 ★7 精度 → 7.1/7.2/7.3）。
    """
    labels = {str(it["n"]) for it in items if _is_label_only(it["text"])}
    starred_parents = {str(it["n"]) for it in items if _is_label_only(it["text"]) and it.get("star")}
    if not labels:
        return items
    keep = []
    for it in items:
        n = str(it["n"])
        if n in labels:
            has_child = any(str(x["n"]).startswith(n + ".") for x in items)
            if has_child and _is_label_only(it["text"]):
                continue
        if "." in n and n.split(".")[0] in starred_parents:
            it = dict(it)
            it["star"] = True
        keep.append(it)
    return keep


def parse_att3(pdf_path: Path = ATT3):
    rows = extract_lines(pdf_path)
    sections = split_device_sections(rows)
    devices = []
    tech_params = []
    star_clauses = []
    raw_blocks = []
    for sec in sections:
        items = drop_empty_parents(parse_section_lines(sec["lines"]))
        params = []
        for i, it in enumerate(items, 1):
            text = clean_text(it["text"])
            if not text:
                continue
            # 丢掉空壳
            rec = {"n": it["n"], "index": i, "star": bool(it["star"]), "text": text, "page": it.get("page")}
            params.append(rec)
            pid = f"P{sec['seq']:02d}-{i}"
            tech_params.append(
                {
                    "id": pid,
                    "star": bool(it["star"]),
                    "device_seq": sec["seq"],
                    "device": sec["name"],
                    "param_n": it["n"],
                    "index": i,
                    "text": text,
                    "page": it.get("page"),
                }
            )
            if it["star"]:
                star_clauses.append(
                    {
                        "id": f"T{sec['seq']:02d}-{i}",
                        "cat": "技术★",
                        "device_seq": sec["seq"],
                        "device": sec["name"],
                        "param_n": it["n"],
                        "text": text,
                        "page": it.get("page"),
                    }
                )
        devices.append(
            {
                "seq": sec["seq"],
                "name": sec["name"],
                "page_start": sec["page_start"],
                "n_params": len(params),
                "n_star": sum(1 for p in params if p["star"]),
                "params": params,
            }
        )
        raw_blocks.append(
            {
                "seq": sec["seq"],
                "name": sec["name"],
                "n_lines": len(sec["lines"]),
                "n_params": len(params),
            }
        )
    return devices, tech_params, star_clauses, raw_blocks, rows


def main():
    # 备份上一轮
    for p in (DEV_PATH, REQ_PATH):
        bak = p.with_suffix(p.suffix + ".r1bak")
        if p.is_file() and not bak.exists():
            shutil.copy2(p, bak)
    devices, tech_params, star_clauses, raw_blocks, rows = parse_att3()
    # 写出新正文便于人工核对
    body_path = WORK / "params_raw_body_r2.txt"
    # 仅设备段
    started = False
    out_lines = []
    for page, line in rows:
        if HEADER_RE.match(re.sub(r"^(（[一二三四五六七八九十百零]+）)\1+", r"\1", line)):
            started = True
        if started:
            out_lines.append(f"[p{page}] {line}")
    body_path.write_text("\n".join(out_lines) + "\n", encoding="utf-8")

    DEV_PATH.write_text(json.dumps(devices, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    req = json.loads(REQ_PATH.read_text(encoding="utf-8")) if REQ_PATH.is_file() else {}
    req["source"] = "official tender + annex3, local PyMuPDF re-extract r2, no external API"
    req["devices"] = [{"seq": d["seq"], "name": d["name"], "n_params": d["n_params"], "n_star": d["n_star"]} for d in devices]
    req["tech_params"] = tech_params
    req["star_clauses"] = star_clauses
    req.setdefault("counts", {})
    req["counts"]["devices"] = len(devices)
    req["counts"]["tech_params"] = len(tech_params)
    req["counts"]["star_params"] = len(star_clauses)
    if "structure" in req and "business_qualification" in req:
        req["counts"]["structure"] = len(req["structure"])
        req["counts"]["business_qualification"] = len(req["business_qualification"])
        req["counts"]["total_checklist"] = (
            req["counts"]["structure"]
            + req["counts"]["business_qualification"]
            + req["counts"]["tech_params"]
        )
    REQ_PATH.write_text(json.dumps(req, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    # 对照第一轮备份
    old = []
    bak = DEV_PATH.with_suffix(".json.r1bak")
    if bak.exists():
        old = json.loads(bak.read_text(encoding="utf-8"))
    old_map = {d["seq"]: d for d in old}
    print(f"devices={len(devices)} tech_params={len(tech_params)} star_params={len(star_clauses)}")
    zero = [d["seq"] for d in devices if d["n_params"] == 0]
    print(f"zero_param_devices={zero}")
    print("seq\tname\told_n\told_star\tnew_n\tnew_star\tdelta_n")
    for d in devices:
        o = old_map.get(d["seq"], {})
        on, os_ = o.get("n_params"), o.get("n_star")
        print(f"{d['seq']}\t{d['name']}\t{on}\t{os_}\t{d['n_params']}\t{d['n_star']}\t{(d['n_params']-(on or 0))}")
    # 重点 9 台
    focus = [16, 25, 32, 36, 38, 39, 41, 42, 44]
    print("\n=== focus devices ===")
    fmap = {d["seq"]: d for d in devices}
    for s in focus:
        d = fmap[s]
        print(f"\n({s}) {d['name']}  n={d['n_params']} star={d['n_star']}")
        for p in d["params"]:
            mark = "★" if p["star"] else " "
            print(f"  {mark}{p['n']}. {p['text'][:120]}")


if __name__ == "__main__":
    main()
