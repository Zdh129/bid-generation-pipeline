#!/usr/bin/env python3
"""对每条技术参数做本机偏离判断：规则 + 数值比较 + 附件证据。

结论仅四类：响应 / 正偏离 / 负偏离 / 无证据待确认。不得一律写「响应」。
不调用外部 API。
"""
from __future__ import annotations

import json
import math
import random
import re
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WORK = ROOT / "work"
DEVICES = json.loads((WORK / "devices.json").read_text(encoding="utf-8"))
REQ = json.loads((WORK / "requirements.json").read_text(encoding="utf-8"))
V8 = json.loads((WORK / "v8_deviations.json").read_text(encoding="utf-8"))
CORPUS_DIR = WORK / "att_corpus"
OUT = WORK / "deviation_judgements.json"

# --------------- corpus ---------------

def load_corpus() -> dict[int, list[dict]]:
    by: dict[int, list[dict]] = defaultdict(list)
    for p in sorted(CORPUS_DIR.glob("eq*.json")):
        d = json.loads(p.read_text(encoding="utf-8"))
        seq = d.get("seq")
        if not seq:
            continue
        for pg in d.get("page_texts") or []:
            text = pg.get("text") or ""
            if not re.sub(r"\s+", "", text):
                continue
            by[seq].append(
                {
                    "file": d.get("file"),
                    "kind": d.get("kind"),
                    "page": pg.get("page"),
                    "source": pg.get("source") or ("rapidocr" if d.get("ocr_applied") else "text"),
                    "text": text,
                }
            )
    return by


def norm(s: str) -> str:
    s = s or ""
    s = s.replace("０", "0").replace("１", "1").replace("２", "2").replace("３", "3")
    s = s.replace("４", "4").replace("５", "5").replace("６", "6").replace("７", "7")
    s = s.replace("８", "8").replace("９", "9")
    s = s.replace("％", "%").replace("～", "-").replace("—", "-").replace("–", "-")
    s = s.replace("≦", "≤").replace("≧", "≥").replace("≤", "≤").replace("≥", "≥")
    s = s.replace("μ", "u").replace("μ", "u")
    s = re.sub(r"\s+", "", s)
    return s.lower()


# --------------- numbers / evidence ---------------

NUM = r"[-+]?\d+(?:\.\d+)?(?:[eE][-+]?\d+)?"

OCR_FIXES = (
    ("秤量", "称量"),
    ("土", "±"),
    ("睿积", "容积"),
    ("大平", "天平"),
    ("μ1", "μl"),
    ("u1", "μl"),
    ("Ⅲg", "mg"),
    ("Ⅲ", "mg"),
    ("网客", "网络"),
)

def ocr_fix(s: str) -> str:
    for a, b in OCR_FIXES:
        s = s.replace(a, b)
    for a, b in (("－", "-"), ("−", "-"), ("–", "-"), ("—", "-"), ("～", "-")):
        s = s.replace(a, b)
    # RT+10_300℃ / RT+10, 300℃ → RT+10-300℃
    s = re.sub(r"(RT\s*\+\s*\d+)\s*[_ ,，]\s*(\d+(?:\.\d+)?\s*(?:℃|°C))", r"\1-\2", s, flags=re.I)
    return s


def to_float(s: str) -> float | None:
    s = s.replace(",", "").replace(" ", "")
    try:
        return float(s)
    except ValueError:
        return None


KNOWN_UNITS = {
    "g", "mg", "kg", "ml", "l", "ul", "w", "kw", "v", "hz", "c",
    "mm", "cm", "nm", "bar", "psi", "mpa", "kpa", "l/min", "ml/min",
    "ppm", "lx", "ntu", "abs", "ph", "%", "s", "min", "h", "kn", "n",
    "wei", "rpm", "brix", "mv",
}


def unit_norm(u: str) -> str:
    u = (u or "").strip().lower()
    u = u.replace("μ", "u").replace("µ", "u").replace("％", "%")
    u = u.rstrip("/;；,，")
    aliases = {
        "毫升": "ml", "ml": "ml", "ul": "ul", "μl": "ul",
        "l": "l", "g": "g", "mg": "mg", "kg": "kg",
        "w": "w", "kw": "kw", "v": "v", "vdc": "v", "vac": "v", "dc": "v",
        "hz": "hz", "mv": "mv", "millivolt": "mv",
        "℃": "c", "°c": "c", "c": "c",
        "mm": "mm", "cm": "cm", "nm": "nm",
        "bar": "bar", "psi": "psi", "mv": "mv", "milli volt": "mv", "mpa": "mpa", "kpa": "kpa",
        "l/min": "l/min", "ml/min": "ml/min", "cc/min": "ml/min",
        "ppm": "ppm", "lx": "lx", "ntu": "ntu", "abs": "abs",
        "ph": "ph", "%": "%", "s": "s", "min": "min", "h": "h",
        "kn": "kn", "n": "n", "位": "wei", "wei": "wei",
        "rpm": "rpm", "brix": "brix", "brix%": "brix",
    }
    return aliases.get(u, u)


def units_compatible(u1: str, u2: str) -> bool:
    a, b = unit_norm(u1), unit_norm(u2)
    if not a or not b:
        return False
    if a not in KNOWN_UNITS or b not in KNOWN_UNITS:
        return False
    SYN = [
        {"量程", "称量", "称量范围", "称量值", "最大称量", "最大称量值"},
        {"流量", "出口流量", "最大流量"},
        {"温度范围", "使用温度范围", "控温范围", "工作温度"},
        {"测量范围", "测定范围", "电位测量范围"},
    ]
    if a == b:
        return True
    for g in SYN:
        if any(s in a for s in g) and any(s in b for s in g):
            return True
    families = [
        {"l", "ml"},
        {"l/min", "ml/min"},
        {"w", "kw"},
    ]
    return any(a in f and b in f for f in families)


def convert(v: float, src: str, dst: str) -> float:
    s, d = unit_norm(src), unit_norm(dst)
    if s == d:
        return v
    vol = {("ml", "l"): 0.001, ("l", "ml"): 1000}
    flow = {("ml/min", "l/min"): 0.001, ("l/min", "ml/min"): 1000}
    power = {("w", "kw"): 0.001, ("kw", "w"): 1000}
    for table in (vol, flow, power):
        if (s, d) in table:
            return v * table[(s, d)]
    return v


def _label_before(text: str, pos: int) -> str:
    pre = text[max(0, pos - 18):pos]
    pre = pre.rstrip("≤≥<>±:：= \t")
    # 取最近的中文键名
    m = re.search(r"([\u4e00-\u9fff]{2,8}|[A-Za-z]{2,12})\s*[:：]?\s*$", pre)
    if m:
        return m.group(1)
    keys = ("分辨率", "准确度", "重复性", "线性误差", "称量值", "称量", "可读性", "样品位",
            "滴定杯", "流量", "纯度", "露点", "压力", "容积", "容量", "波动", "精度",
            "电源", "电压", "尺寸", "位移", "频率")
    for k in keys:
        if k in pre:
            return k
    return ""


def extract_constraints(text: str) -> list[dict]:
    raw = ocr_fix(text)
    t = raw.replace("≦", "≤").replace("≧", "≥").replace("＜", "<").replace("＞", ">")
    t = t.replace("−", "-").replace("–", "-").replace("—", "-")
    t = t.replace("不小于", "≥").replace("不低于", "≥").replace("至少覆盖", "≥").replace("至少", "≥")
    t = t.replace("不超过", "≤").replace("不高于", "≤").replace("不大于", "≤")
    t = re.sub(r"(?<=\d),(?=\d{3})", "", t)
    t = re.sub(r"(\d)\s+\.(\d)", r"\1.\2", t)  # "0 .005Abs" -> 0.005Abs
    cons: list[dict] = []
    occupied = [False] * (len(t) + 1)

    def mark(a, b):
        for i in range(max(0, a), min(len(occupied), b)):
            occupied[i] = True

    def overlaps(a, b):
        return any(occupied[i] for i in range(max(0, a), min(len(occupied), b)))

    def add(d, a, b):
        if overlaps(a, b):
            return
        d["label"] = _label_before(t, a)
        d["pos"] = a
        cons.append(d)
        mark(a, b)

    # 外形尺寸 600×700×1500mm
    for m in re.finditer(r"((?:大于|小于|不小于|≥|>)?)?\s*(\d+(?:\.\d+)?)\s*[×xX*]\s*(\d+(?:\.\d+)?)\s*[×xX*]\s*(\d+(?:\.\d+)?)\s*(mm|cm|m)?", t):
        dims = [float(m.group(2)), float(m.group(3)), float(m.group(4))]
        how = m.group(1) or ""
        op = "dim_gt" if how in ("大于", ">", "≥", "不小于") else "dim"
        add({"op": op, "dims": dims, "unit": unit_norm(m.group(5) or "mm"), "span": m.group(0)}, m.start(), m.end())

    # 带符号区间，必须一端有正负号，避免吞掉 1-40％、0-12000LX
    for m in re.finditer(r"([+-]\d+(?:\.\d+)?)\s*[-~～至到]\s*[+]?(-?\d+(?:\.\d+)?)\s*(mV|mv|V|℃|°C|C|LX|lx|％|%)?", t):
        a, b = to_float(m.group(1)), to_float(m.group(2))
        if a is None or b is None:
            continue
        unit = unit_norm(m.group(3) or "")
        if unit == "v" and (abs(a) + abs(b)) > 250:
            unit = "mv" if "mv" in m.group(0).lower() or (abs(a) >= 100) else unit
        trail = t[m.end():m.end()+3]
        if "mV" in m.group(0) or "mv" in m.group(0).lower() or trail.lower().startswith("mv"):
            unit = "mv"
        elif trail.lower().startswith("ph") or "pH" in trail:
            unit = "ph"
        add({"op": "range", "lo": min(a, b), "hi": max(a, b), "unit": unit, "span": m.group(0)+trail[:2]}, m.start(), m.end())

    # RT+10-300℃
    for m in re.finditer(r"RT\s*\+\s*(\d+(?:\.\d+)?)\s*[-~～至到]\s*(\d+(?:\.\d+)?)\s*(?:℃|°C)", t, flags=re.I):
        a, b = to_float(m.group(1)), to_float(m.group(2))
        if a is not None and b is not None:
            add({"op": "range", "lo": min(a, b), "hi": max(a, b), "unit": "c", "span": m.group(0)}, m.start(), m.end())

    # 0℃-60℃
    for m in re.finditer(r"([-+]?\d+(?:\.\d+)?)\s*(?:℃|°C)\s*[-~～至到]\s*([-+]?\d+(?:\.\d+)?)\s*(?:℃|°C)?", t):
        a, b = to_float(m.group(1)), to_float(m.group(2))
        if a is not None and b is not None:
            add({"op": "range", "lo": min(a, b), "hi": max(a, b), "unit": "c", "span": m.group(0)}, m.start(), m.end())

    # 0.1 g - 200 g（单位夹在中间）
    for m in re.finditer(rf"({NUM})\s*(g|mg|kg|ml|l|mm|cm|nm|℃|°C)\s*[-~～至到]\s*({NUM})\s*(g|mg|kg|ml|l|mm|cm|nm|℃|°C)", t, flags=re.I):
        a, b = to_float(m.group(1)), to_float(m.group(3))
        u1, u2 = unit_norm(m.group(2)), unit_norm(m.group(4))
        if a is None or b is None:
            continue
        # 同一表达式内允许 mg↔g 等换算（全局 units_compatible 仍禁止跨条款用 mg 去比 g）
        intra = {("mg","g"): 0.001, ("g","mg"): 1000, ("kg","g"): 1000, ("g","kg"): 0.001,
                 ("ml","l"): 0.001, ("l","ml"): 1000, ("cm","mm"): 10, ("mm","cm"): 0.1}
        if u1 != u2 and u1 and u2:
            if (u1, u2) in intra:
                a = a * intra[(u1, u2)]
                u1 = u2
            elif not units_compatible(u1, u2):
                continue
        add({"op": "range", "lo": min(a, b), "hi": max(a, b), "unit": u1 or u2, "span": m.group(0)}, m.start(), m.end())

    # 普通区间 15-100psi / 100-240 V（避免把 100-240 的后半当成 -240）
    for m in re.finditer(rf"({NUM})\s*[-~～至到]\s*({NUM})\s*([%％°℃A-Za-zμµΩ/]+|毫升|位)?", t):
        a, b = to_float(m.group(1)), to_float(m.group(2))
        if a is None or b is None:
            continue
        unit = unit_norm(m.group(3) or "")
        trail = t[m.end():m.end()+4]
        if not unit:
            if trail.lower().startswith("lx") or m.group(0).lower().endswith("lx"):
                unit = "lx"
            elif trail.startswith(("%", "％")) or m.group(0).endswith(("%", "％")):
                unit = "%"
        add({"op": "range", "lo": min(a, b), "hi": max(a, b), "unit": unit, "span": m.group(0)+trail[:2]}, m.start(), m.end()+(2 if unit in ("lx","%") and not m.group(3) else 0))

    for m in re.finditer(rf"(≥|≤|>|<|±)\s*({NUM})\s*([%％°℃A-Za-zμµΩ/]+|毫升|位)?", t):
        v = to_float(m.group(2))
        if v is None:
            continue
        add({"op": m.group(1), "value": v, "unit": unit_norm(m.group(3) or ""), "span": m.group(0)}, m.start(), m.end())

    # 裸数值+单位（白名单单位），不与已占用区间重叠
    for m in re.finditer(rf"({NUM})\s*([A-Za-zμµΩ℃°/%]{{1,8}}|毫升|位)", t):
        raw_n = m.group(1)
        # 跳过 003pH 这类被截断的前导零片段（本应是 0.003pH）
        if re.match(r"^0\d", raw_n):
            continue
        v = to_float(raw_n)
        unit = unit_norm(m.group(2))
        if v is None or unit not in KNOWN_UNITS:
            continue
        add({"op": "value", "value": v, "unit": unit, "span": m.group(0)}, m.start(), m.end())

    # 样品位 64位 / 64 位
    for m in re.finditer(r"(\d+(?:\.\d+)?)\s*位", t):
        v = to_float(m.group(1))
        if v is not None:
            add({"op": "value", "value": v, "unit": "wei", "span": m.group(0), "label": "样品位"}, m.start(), m.end())

    return cons


def leading_keys(param_text: str) -> list[str]:
    t = ocr_fix(param_text)
    keys = []
    m = re.match(r"^(?:★)?\s*(?:[\d.]+[、.．]\s*)?([^：:]{1,16})[：:]", t)
    if m:
        keys.append(m.group(1).strip())
    for k in (
        "功率消耗", "最大称量值", "最大称量", "样品位", "滴定杯体积", "氢气流量",
        "示例流量A", "出口流量", "示例容积", "使用温度范围", "空载最低温度",
        "温度稳定性", "光照强度", "显示精度", "测量范围", "工作温度范围",
        "重复性", "线性误差", "稳定性", "分辨率", "准确度", "精度", "流量", "容积", "容量",
        "露点", "输出压力", "排污阀", "干燥机", "真空度", "真空压力", "最大压力", "转速",
        "外形尺寸", "电源要求",
    ):
        if k in t:
            keys.append(k)
    seen, out = set(), []
    for k in keys:
        if k and k not in seen:
            seen.add(k)
            out.append(k)
    return out or keywords(t)[:3]


def local_windows(param_text: str, pages: list[dict], width=180) -> list[dict]:
    """只截取本条款窗口，不把上一编号条款的数字带进来。"""
    keys = leading_keys(param_text)
    syn = list(keys)
    for k in keys:
        if "称量" in k:
            syn.append(k.replace("称量", "秤量"))
    hits = []
    item_m = re.match(r"^(?:★)?\s*(\d+(?:\.\d+)?)[、.．]", param_text.strip())
    item_no = item_m.group(1) if item_m else None

    def clip_from(raw: str, start: int) -> str:
        # 向前最多 16 字保留「变压吸附」等修饰，但不跨过上一条编号
        a = max(0, start - 16)
        head = raw[a:start]
        prev = None
        for mm in re.finditer(r"★?\s*\d+[、.．]", head):
            prev = mm.end()
        if prev is not None:
            a = a + prev
        tail = raw[start:]
        nxt = re.search(r"(?:\n|。|；)\s*★?\s*\d+[、.．]", tail[2:] if len(tail) > 2 else "")
        end = start + 2 + nxt.start() if nxt else min(len(raw), start + width)
        return raw[a:end]

    for pg in pages:
        raw = ocr_fix(pg["text"])
        positions = []
        key_hits = []
        for k in syn:
            pos = 0
            while True:
                i = raw.find(k, pos)
                if i < 0:
                    break
                key_hits.append(i)
                pos = i + max(len(k), 1)
        if key_hits:
            # 招标序号与确认函序号常不一致，有键名时不要用「4、」去切到邻项
            positions.extend(key_hits)
        elif item_no:
            for mm in re.finditer(rf"★?\s*{re.escape(item_no)}\s*[、.．]", raw):
                positions.append(mm.start())
        # 去重、排序
        used = set()
        for i in sorted(positions):
            keyp = (pg.get("file"), pg.get("page"), i)
            if keyp in used:
                continue
            used.add(keyp)
            win = clip_from(raw, i)
            if len(re.sub(r"\s+", "", win)) < 2:
                continue
            hits.append({**pg, "text": win, "anchor_pos": i})
    return hits


def _labels_similar(a: str | None, b: str | None) -> bool:
    a = (a or "").strip()
    b = (b or "").strip()
    if not a or not b:
        return False
    SYN = [
        {"量程", "称量", "称量范围", "称量值", "最大称量", "最大称量值"},
        {"流量", "出口流量", "最大流量"},
        {"温度范围", "使用温度范围", "控温范围", "工作温度"},
        {"测量范围", "测定范围", "电位测量范围", "水分"},
    ]
    if a == b:
        return True
    for g in SYN:
        if any(s in a for s in g) and any(s in b for s in g):
            return True
    dist = ("环境", "工作", "存放", "存储", "室温", "电源", "输出", "输入", "出口", "入口", "外形", "外观")
    da, db = {d for d in dist if d in a}, {d for d in dist if d in b}
    if da != db:
        # fall through to core compare only if no distinguishing prefixes
        pass
    elif a in b or b in a:
        return True
    def core(s: str) -> str:
        return re.sub(r"(温度|使用|测定|测量|波长|光度|最大|最小|额定|典型|电位)", "", s)
    ca, cb = core(a), core(b)
    if ca and cb and (ca == cb or ca in cb or cb in ca):
        # 环境温度 vs 温度 等：前缀改变量的含义则不算相近
        dist = ("环境", "工作", "存放", "存储", "室温", "电源", "输出", "输入", "出口", "入口", "外形", "外观")
        da, db = {d for d in dist if d in a}, {d for d in dist if d in b}
        if da != db:
            return False
        return True
    return False


def _pair_offered(r, oc, used: set) -> tuple[dict | None, str]:
    cands = []
    for i, o in enumerate(oc):
        if i in used:
            continue
        if r["op"] in ("dim", "dim_gt"):
            if o.get("op") in ("dim", "dim_gt") and units_compatible(r.get("unit") or "mm", o.get("unit") or "mm"):
                cands.append((i, o, 3))
            continue
        if not units_compatible(r.get("unit"), o.get("unit")):
            continue
        # 区间对裸值、公差对公差：必须标签相近，避免邻项串比
        if r["op"] == "range" and o.get("op") == "±":
            continue
        if r["op"] == "range" and o.get("op") == "range" and r.get("label") and o.get("label") and not _labels_similar(r.get("label"), o.get("label")):
            continue
        if r["op"] == "range" and o.get("op") == "value" and not _labels_similar(r.get("label"), o.get("label")):
            continue
        if r["op"] == "range" and o.get("op") == "range":
            score_bonus_range = 4
        else:
            score_bonus_range = 0
        if r["op"] == "±" and o.get("label") and r.get("label") and not _labels_similar(r.get("label"), o.get("label")):
            continue
        score = 1
        if _labels_similar(r.get("label"), o.get("label")):
            score += 5
        if r.get("label") and o.get("label") and r["label"] == o["label"]:
            score += 1
        if r["op"] == o.get("op"):
            score += 2
        if r["op"] in ("≥", ">", "≤", "<") and o.get("op") in {r["op"], "value", "≥", "≤"}:
            score += 1
        if r["op"] == "range" and o.get("op") == "range":
            score += 1 + score_bonus_range
        if r["op"] == "±" and o.get("op") in {"±", "value"}:
            score += 1
        cands.append((i, o, score))
    if not cands:
        return None, ""
    # 标签一致优先；再按分数；避免抢用明显更远的数值
    cands.sort(key=lambda x: (-x[2], abs((x[1].get("value") or x[1].get("lo") or 0) - (r.get("value") or r.get("lo") or 0))))
    # 若最高分中有标签匹配，只用它们
    top = cands[0][2]
    cands = [c for c in cands if c[2] == top or (r.get("label") and c[1].get("label") == r["label"])]
    i, o, sc = cands[0]
    used.add(i)
    how = "同标签" if r.get("label") and r.get("label") == o.get("label") else "同单位配对"
    return o, how


def compare_numeric(req_text: str, offered_text: str) -> tuple[str | None, str]:
    rc = [c for c in extract_constraints(req_text) if c["op"] != "value" or c.get("label")]
    # 需求侧优先用带算子的约束；裸值只在没有算子时使用
    primary = [c for c in extract_constraints(req_text) if c["op"] not in {"value"}]
    if not primary:
        primary = extract_constraints(req_text)
    oc = extract_constraints(offered_text)
    if not primary or not oc:
        return None, "无足够数值约束"
    notes = []
    verdicts = []
    used: set[int] = set()

    for r in primary:
        if r["op"] in ("dim", "dim_gt"):
            o, how = _pair_offered(r, oc, used)
            if not o or "dims" not in o:
                continue
            rd, od = list(r["dims"]), list(o["dims"])
            # 按大小排序后逐项比，避免长宽顺序不同
            rd_s, od_s = sorted(rd), sorted(od)
            # 「大于 A×B×C」要求投标三维均不小于，且对应排序后每一维 ≥
            worse = any(a < b - 1e-6 for a, b in zip(od_s, rd_s))
            better = all(a + 1e-6 >= b for a, b in zip(od_s, rd_s)) and any(a > b + 1e-6 for a, b in zip(od_s, rd_s))
            if worse:
                verdicts.append("负偏离")
                notes.append(f"外形尺寸{od}未满足{rd}")
            elif better:
                verdicts.append("正偏离")
                notes.append(f"外形尺寸{od}优于{rd}")
            else:
                verdicts.append("响应")
                notes.append(f"外形尺寸{od}符合{rd}")
            continue

        o, how = _pair_offered(r, oc, used)
        if not o:
            continue
        if r["op"] == "range" and o["op"] == "range":
            if o["lo"] <= r["lo"] + 1e-9 and o["hi"] >= r["hi"] - 1e-9:
                if o["lo"] < r["lo"] - 1e-9 or o["hi"] > r["hi"] + 1e-9:
                    verdicts.append("正偏离")
                    notes.append(f"范围{o['lo']}-{o['hi']}覆盖并宽于要求{r['lo']}-{r['hi']}{r.get('unit')}")
                else:
                    verdicts.append("响应")
                    notes.append(f"范围一致{r['lo']}-{r['hi']}{r.get('unit')}")
            elif o["lo"] > r["lo"] + 1e-6 or o["hi"] < r["hi"] - 1e-6:
                verdicts.append("负偏离")
                notes.append(f"范围{o['lo']}-{o['hi']}未覆盖要求{r['lo']}-{r['hi']}{r.get('unit')}")
            continue
        if r["op"] == "range" and o.get("value") is not None:
            # 仅当标签相近时把单点当作区间上的点；且单点不在区间内不直接判负偏离（常见是邻项分辨率/精度）
            if not _labels_similar(r.get("label"), o.get("label")):
                continue
            ov = convert(o["value"], o.get("unit") or r.get("unit"), r.get("unit") or o.get("unit"))
            if r["lo"] - 1e-6 <= ov <= r["hi"] + 1e-6:
                verdicts.append("响应")
                notes.append(f"投标值{ov}落在要求区间{r['lo']}-{r['hi']}{r.get('unit')}")
            # 否则忽略，等待区间对区间
            continue
        if r["op"] == "±":
            ov = o.get("value")
            if ov is None:
                continue
            ov = convert(ov, o.get("unit") or r.get("unit"), r.get("unit") or o.get("unit"))
            if abs(ov - r["value"]) <= max(1e-9, 0.002 * abs(r["value"])):
                verdicts.append("响应")
                notes.append(f"±{r['value']}{r.get('unit')}一致")
            elif ov < r["value"] - 1e-12:
                verdicts.append("正偏离")
                notes.append(f"公差更严：要求±{r['value']}{r.get('unit')}，投标±{ov}")
            else:
                verdicts.append("负偏离")
                notes.append(f"公差更松：要求±{r['value']}{r.get('unit')}，投标±{ov}")
            continue
        rv = r.get("value")
        if rv is None:
            continue
        if o["op"] == "range":
            # 以能代表该算子的端点比较
            offered_v = o["hi"] if r["op"] in ("≥", ">") else o["lo"] if r["op"] in ("≤", "<") else o["hi"]
        else:
            if o.get("value") is None:
                continue
            offered_v = convert(o["value"], o.get("unit") or r.get("unit"), r.get("unit") or o.get("unit"))
        tol = max(1e-9, 0.002 * abs(rv) if rv else 1e-9)
        if r["op"] in ("≥", ">"):
            worse = offered_v < rv - tol
            better = offered_v > rv + tol
        elif r["op"] in ("≤", "<"):
            worse = offered_v > rv + tol
            better = offered_v < rv - tol
        elif r["op"] == "value" and _labels_similar(r.get("label"), o.get("label") or r.get("label")):
            lab = (r.get("label") or "") + (o.get("label") or "")
            higher_better = any(k in lab for k in ("流量", "称量", "容积", "容量", "功率", "最大称量", "出口"))
            lower_better = any(k in lab for k in ("可读性", "感量", "线性误差", "重复性"))
            if not higher_better and not lower_better:
                # 相同则响应，不同不擅自定方向
                if abs(offered_v - rv) <= tol:
                    verdicts.append("响应")
                    notes.append(f"{how}: {rv}{r.get('unit')}一致")
                continue
            if higher_better:
                worse = offered_v < rv - tol
                better = offered_v > rv + tol
            else:
                worse = offered_v > rv + tol
                better = offered_v < rv - tol
        else:
            continue
        if worse:
            verdicts.append("负偏离")
            notes.append(f"{how}: 要求{r['op']}{rv}{r.get('unit')}，投标{offered_v}")
        elif better:
            verdicts.append("正偏离")
            notes.append(f"{how}: 要求{r['op']}{rv}{r.get('unit')}，投标{offered_v}更优")
        else:
            verdicts.append("响应")
            notes.append(f"{how}: 要求{r['op']}{rv}{r.get('unit')}，投标{offered_v}符合（含等于界限）")
    if not verdicts:
        return None, "未能对齐同单位可比数值"
    if "负偏离" in verdicts:
        return "负偏离", "；".join(notes)
    if "正偏离" in verdicts:
        return "正偏离", "；".join(notes)
    return "响应", "；".join(notes)


TOKEN_RE = re.compile(r"[\u4e00-\u9fff]{2,}|[A-Za-z]{2,}|[%％°℃]+|" + NUM)


def keywords(text: str) -> list[str]:
    stop = {"要求", "具有", "采用", "配置", "功能", "可以", "能够", "进行", "以及", "或者", "包括", "通过", "用于", "参数", "技术", "仪器"}
    toks = []
    for t in TOKEN_RE.findall(ocr_fix(text)):
        if t in stop or len(t) == 1:
            continue
        toks.append(t)
    return toks


def snippet(text: str, keys: list[str], width=90) -> str:
    ntext = ocr_fix(text)
    pos = -1
    for k in keys:
        i = ntext.find(k)
        if i >= 0:
            pos = i
            break
    if pos < 0:
        compact = re.sub(r"\s+", " ", text).strip()
        return compact[: width * 2]
    a = max(0, pos - width)
    b = min(len(text), pos + width)
    s = re.sub(r"\s+", " ", text[a:b]).strip()
    if a > 0:
        s = "…" + s
    if b < len(text):
        s = s + "…"
    return s


def find_evidences(param_text: str, pages: list[dict], topn=3) -> list[dict]:
    keys = leading_keys(param_text) + keywords(param_text)
    windows = local_windows(param_text, pages)
    pool = windows or pages
    nt_param = norm(ocr_fix(param_text))
    scored = []
    for pg in pool:
        nt = norm(ocr_fix(pg["text"]))
        if not nt:
            continue
        score = 0
        if len(nt_param) >= 8 and nt_param[:18] in nt:
            score += 10
        hit = 0
        for k in keys:
            kn = norm(k)
            if kn and kn in nt:
                hit += 1
                score += 2 if len(k) >= 3 else 1
        if hit == 0 and score == 0:
            continue
        if pg.get("kind") == "confirm":
            score += 3
        scored.append((score, hit, pg))
    scored.sort(key=lambda x: (-x[0], 0 if x[2].get("kind") == "confirm" else 1))
    evs, seen = [], set()
    for score, hit, pg in scored:
        sig = (pg["file"], pg.get("page"), pg["text"][:40])
        if sig in seen:
            continue
        seen.add(sig)
        evs.append(
            {
                "file": pg["file"],
                "page": pg["page"],
                "kind": pg["kind"],
                "source": pg["source"],
                "score": score,
                "snippet": snippet(pg["text"], keys),
            }
        )
        if len(evs) >= topn:
            break
    return evs


DEV_LABEL_RE = re.compile(r"(无偏离|正偏离|负偏离)")


def parse_explicit_deviations(pages: list[dict]) -> list[dict]:
    found = []
    for pg in pages:
        text = ocr_fix(pg["text"])
        for m in DEV_LABEL_RE.finditer(text):
            label = m.group(1)
            start = max(0, m.start() - 180)
            ctx = text[start:m.end()]
            mapped = "响应" if label == "无偏离" else label
            found.append(
                {
                    "label": mapped,
                    "raw": label,
                    "context": re.sub(r"\s+", " ", ctx).strip(),
                    "file": pg["file"],
                    "page": pg["page"],
                    "source": pg["source"],
                }
            )
    return found


def match_explicit(param_text: str, explicit: list[dict]) -> dict | None:
    if not explicit:
        return None
    keys = leading_keys(param_text) + keywords(param_text)
    nt = norm(ocr_fix(param_text))
    best, best_s = None, 0
    for ex in explicit:
        ctx = norm(ex["context"])
        s = 0
        if len(nt) >= 8 and nt[:14] in ctx:
            s += 12
        for k in keys[:8]:
            if norm(k) and norm(k) in ctx:
                s += 3
        if s > best_s:
            best_s, best = s, ex
    if best and best_s >= 6:
        return {**best, "match_score": best_s}
    return None


def qualitative_hints(param_text: str, offered: str) -> tuple[str | None, str]:
    pt, ot = ocr_fix(param_text), ocr_fix(offered)
    if "排污阀" in pt:
        if "手动" in pt and ("全自动" in ot or "自动排污" in ot) and "手动" not in ot:
            return "正偏离", "招标为手动排污阀，附件为自动/全自动示例阀"
        if "手动" in pt and "手动" in ot:
            return "响应", "附件同样为手动排污阀"
    if "干燥机" in pt and "非耗材" in pt:
        if "变压吸附" in ot and "非耗材" not in ot:
            return "负偏离", "招标要求吸附式干燥机（非耗材），附件表述为变压吸附干燥机、未写明非耗材"
    if "活塞压缩机" in pt and "涡旋" in ot:
        return "负偏离", "招标为活塞压缩机，附件为涡旋式压缩机"
    # 直流电压明确不同：9 V DC vs 12V DC
    m1 = re.search(r"(\d+(?:\.\d+)?)\s*V\s*DC", pt, re.I)
    m2 = re.search(r"(\d+(?:\.\d+)?)\s*V\s*DC", ot, re.I)
    if m1 and m2:
        a, b = float(m1.group(1)), float(m2.group(1))
        if abs(a - b) > 0.05:
            return "负偏离", f"直流电压招标 {a} V DC，附件为 {b} V DC"
    return None, ""


def offered_window(param_text: str, pages: list[dict]) -> str:
    wins = local_windows(param_text, pages)
    if not wins:
        evs = find_evidences(param_text, pages, topn=2)
        return "\n".join(e["snippet"] for e in evs)
    # 确认函优先，避免彩页规格表里其它型号数字串进来
    conf = [w for w in wins if w.get("kind") == "confirm"]
    use = conf or wins
    # 取最短且含键名的窗口，减少邻项
    use.sort(key=lambda w: (0 if w.get("kind") == "confirm" else 1, len(w["text"])))
    return "\n".join(w["text"] for w in use[:3])


def confirm_restates(param_text: str, pages: list[dict]) -> bool:
    distinctive = [k for k in keywords(param_text) if len(k) >= 2 and not re.fullmatch(NUM, k)]
    if not distinctive:
        return False
    keys = leading_keys(param_text)
    for pg in pages:
        if pg.get("kind") not in {"confirm", "other"}:
            continue
        nt = norm(ocr_fix(pg["text"]))
        hit = sum(1 for k in distinctive if norm(k) in nt)
        if hit / max(len(distinctive), 1) >= 0.5:
            return True
        if keys and norm(keys[0]) in nt and hit >= min(3, len(distinctive)):
            return True
    return False


def judge_one(param: dict, pages: list[dict], explicit: list[dict]) -> dict:
    text = param["text"]
    evs = find_evidences(text, pages, topn=3)
    ex = match_explicit(text, explicit)
    offered = offered_window(text, pages)
    q_v, q_note = qualitative_hints(text, offered)
    num_v, num_note = compare_numeric(text, offered) if offered else (None, "无附件片段")

    restated = False
    if evs:
        evn = norm(ocr_fix(evs[0]["snippet"]))
        tn = norm(ocr_fix(text))
        if tn[:16] in evn or evn[:16] in tn:
            restated = True
        elif leading_keys(text) and all(norm(k) in evn for k in leading_keys(text)[:1]):
            # 键名命中且数值约束一致时后面再判断
            restated = False

    reason_parts = []
    verdict = None
    # 1) 定性高信号
    if q_v in {"正偏离", "负偏离"}:
        verdict = q_v
        reason_parts.append(q_note)
    # 2) 同单位数值
    elif num_v:
        verdict = num_v
        reason_parts.append("数值比较：" + num_note)
        if ex and ex["label"] != num_v:
            reason_parts.append(f"附件偏离表写「{ex['raw']}」，与数值比较不一致，以数值为准")
    # 3) 偏离表
    elif ex:
        verdict = ex["label"]
        reason_parts.append(f"附件偏离表写「{ex['raw']}」")
    elif q_v == "响应":
        verdict = "响应"
        reason_parts.append(q_note)
    elif confirm_restates(text, pages) and num_v not in {"正偏离", "负偏离"}:
        verdict = "响应"
        reason_parts.append("参数确认函/偏离表覆盖该条款，且未见同单位数值差异")
    elif evs and (tn := norm(ocr_fix(text))) and (tn[:16] in norm(ocr_fix(evs[0]["snippet"])) or restated):
        verdict = "响应"
        reason_parts.append("附件复述该条款且未见同单位数值差异")
    elif evs:
        verdict = "无证据待确认"
        reason_parts.append("附件有相关表述，但无法可靠判断是否完全满足或存在偏离")
    else:
        verdict = "无证据待确认"
        reason_parts.append("本机附件中未检索到可引用证据")

    cites = evs
    if ex:
        cites = [
            {
                "file": ex["file"],
                "page": ex["page"],
                "kind": "confirm",
                "source": ex["source"],
                "score": ex.get("match_score"),
                "snippet": ex["context"][:180],
            }
        ] + [e for e in evs if not (e["file"] == ex["file"] and e["page"] == ex["page"])]

    return {
        "id": param["id"],
        "device_seq": param["device_seq"],
        "device": param["device"],
        "param_n": param["param_n"],
        "star": param["star"],
        "req_text": text,
        "verdict": verdict,
        "reason": "；".join(reason_parts),
        "numeric_note": num_note,
        "evidence": cites[:3],
    }


# --------------- V8 match ---------------

def v8_clean_rows():
    rows = []
    for r in V8["rows"]:
        if r.get("deviation") == "合计（元）" or "合计" in (r.get("deviation_raw") or ""):
            continue
        rows.append(r)
    return rows


def assign_v8_to_params(tech_params: list[dict]):
    """把 V8 行配到本轮参数。同名多台按文本最佳匹配。"""
    by_name: dict[str, list[dict]] = defaultdict(list)
    for p in tech_params:
        by_name[re.sub(r"\s+", "", p["device"])].append(p)
    used = set()
    matches = {}  # param id -> v8 row
    unmatched_v8 = []
    for r in v8_clean_rows():
        name = re.sub(r"\s+", "", r.get("device_name") or "")
        cands = by_name.get(name) or []
        # 宽松：名称包含
        if not cands:
            for n, ps in by_name.items():
                if name and (name in n or n in name):
                    cands = ps
                    break
        if not cands:
            unmatched_v8.append(r)
            continue
        nt = norm(r.get("req_text") or "")
        best, bs = None, -1
        for p in cands:
            if p["id"] in used:
                continue
            pn = norm(p["text"])
            s = 0
            if nt and pn:
                if nt[:16] in pn or pn[:16] in nt:
                    s += 10
                # token overlap
                tk = set(keywords(r.get("req_text") or "")) & set(keywords(p["text"]))
                s += len(tk)
            if s > bs:
                bs, best = s, p
        if best and bs >= 3:
            used.add(best["id"])
            matches[best["id"]] = r
        else:
            unmatched_v8.append(r)
    return matches, unmatched_v8


def main():
    corpus = load_corpus()
    tech = REQ["tech_params"]
    judgements = []
    for p in tech:
        pages = corpus.get(p["device_seq"], [])
        explicit = parse_explicit_deviations(pages)
        judgements.append(judge_one(p, pages, explicit))

    counts = Counter(j["verdict"] for j in judgements)
    v8map, unmatched_v8 = assign_v8_to_params(tech)
    agree = 0
    compared = 0
    disagreements = []
    for j in judgements:
        r = v8map.get(j["id"])
        if not r:
            j["v8_deviation"] = None
            j["v8_agree"] = None
            continue
        compared += 1
        v8v = r["deviation"]
        j["v8_deviation"] = v8v
        j["v8_raw"] = r.get("deviation_raw")
        j["v8_offered"] = r.get("offered_text")
        j["v8_agree"] = j["verdict"] == v8v
        if j["v8_agree"]:
            agree += 1
        else:
            disagreements.append(
                {
                    "id": j["id"],
                    "device": j["device"],
                    "star": j["star"],
                    "ours": j["verdict"],
                    "v8": v8v,
                    "req": j["req_text"][:120],
                    "reason": j["reason"],
                    "v8_offered": (r.get("offered_text") or "")[:120],
                }
            )

    rate = round(agree / compared, 4) if compared else None

    # 抽样：全部正/负偏离 + V8 正/负 + 部分★ + 补足 ≥30
    sample_ids = []
    def add(jid):
        if jid not in sample_ids:
            sample_ids.append(jid)

    for j in judgements:
        if j["verdict"] in {"正偏离", "负偏离"}:
            add(j["id"])
    for j in judgements:
        if j.get("v8_deviation") in {"正偏离", "负偏离"}:
            add(j["id"])
    stars = [j for j in judgements if j["star"]]
    # 全部★中的非「简单响应」+ 一部分响应
    for j in stars:
        if j["verdict"] != "响应":
            add(j["id"])
    random.Random(20261117).shuffle(stars)
    for j in stars:
        if len(sample_ids) >= 36:
            break
        add(j["id"])
    # 补无证据
    pending = [j for j in judgements if j["verdict"] == "无证据待确认"]
    random.Random(20261117).shuffle(pending)
    for j in pending:
        if len(sample_ids) >= 40:
            break
        add(j["id"])

    sample = [j for j in judgements if j["id"] in sample_ids]
    # 人工式核对在单独步骤填写 audit 字段占位
    out = {
        "n_params": len(judgements),
        "n_star": sum(1 for j in judgements if j["star"]),
        "verdict_counts": dict(counts),
        "v8_compared": compared,
        "v8_agree": agree,
        "v8_agreement_rate": rate,
        "v8_unmatched_rows": len(unmatched_v8),
        "disagreements": disagreements,
        "sample_ids": sample_ids,
        "judgements": judgements,
    }
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"params={len(judgements)} verdicts={dict(counts)}")
    print(f"v8 compared={compared} agree={agree} rate={rate} unmatched_v8={len(unmatched_v8)} disagreements={len(disagreements)}")
    print(f"sample={len(sample_ids)} stars_in_sample={sum(1 for j in sample if j['star'])}")
    print("disagreements (ours vs V8):")
    for d in disagreements[:25]:
        print(f"  {d['id']} {d['device']} ours={d['ours']} v8={d['v8']} {d['req'][:60]}")


if __name__ == "__main__":
    main()
