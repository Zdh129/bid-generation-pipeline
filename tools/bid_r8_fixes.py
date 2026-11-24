"""r8：在 r7 叠加层之上再套 r8 口径（donghao t15u + 上级保守口径），仍不改 deviation_judgements.json。
- A、B 照旧；P28-3 继续「待确认」（r7 守卫）。
- 规则 C 不作用于 P18-1、P36-5★：回退为「待确认」，等人工核对（不是正偏离）。
- 严格不等号核对（修正后的通用规则 S）：原判定把「要求>X / <X、投标等于 X」一律记为「符合（含等于界限）」。
  r8 逐条核对：招标原文确为严格不等号，且附件同样写严格不等号（如「＞100 单位」）→ 满足，结论不变；
  附件只写等于临界值（无严格不等号）→ 改为负偏离。每条都记入日志。
- 结论栏仍用「响应」（不改「无偏离」）。
"""
from __future__ import annotations

import json
import re
from collections import Counter, defaultdict
from pathlib import Path

import bid_r7_fixes as R7

ROOT = Path(__file__).resolve().parents[1]
TENDER_ATT3 = ROOT / "work" / "tender_text" / "annex3-example-params.txt"
C_EXCLUDE = {"P18-1": "规则C 不作用于本行（r8 口径），回退待人工核对：附件数值优于要求，V8 判正偏离",
             "P36-5": "规则C 不作用于本行（r8 口径），回退待人工核对：附件范围宽于要求，V8 判正偏离"}
STRICT_FR = re.compile(r"要求([<>＜＞])(-?[\d.]+)([^，；]*?)，投标(-?[\d.]+)符合（含等于界限）")
LOG: list[dict] = []
STRICT_LOG: list[dict] = []


def _num_forms(x: str):
    v = float(x)
    return {x, str(int(v))} if v == int(v) else {x, str(v)}


def strict_check(j: dict) -> list[dict]:
    out = []
    tender = re.sub(r"\s", "", TENDER_ATT3.read_text(encoding="utf-8"))
    ev = " ".join(e.get("snippet") or "" for e in j.get("evidence") or [])
    for m in STRICT_FR.finditer(j.get("reason") or ""):
        op, req, unit, off = m.groups()
        in_tender = re.sub(r"\s", "", j["req_text"]) in tender
        tender_strict = bool(re.search(r"[<>＜＞]\s*" + re.escape(req.rstrip("0").rstrip(".") if "." in req else req), j["req_text"])) or \
            any(re.search(r"[<>＜＞]\s*" + re.escape(n) + r"(?![\d.])", j["req_text"]) for n in _num_forms(req))
        offer_strict = any(re.search(r"[<>＜＞]\s*" + re.escape(n) + r"(?![\d.])", ev) for n in _num_forms(off))
        out.append({"id": j["id"], "op": op, "req": req, "unit": unit.strip(), "offer": off, "req_text": j["req_text"],
                    "tender_original_found": in_tender, "tender_strict": tender_strict, "offer_strict": offer_strict,
                    "satisfied": (not tender_strict) or offer_strict,
                    "evidence_file": (j.get("evidence") or [{}])[0].get("file"), "evidence_page": (j.get("evidence") or [{}])[0].get("page")})
    return out


def apply_overlay(judgements: list[dict]) -> list[dict]:
    js = R7.apply_overlay(judgements)
    r7log = {x["id"]: x for x in R7.OVERLAY_LOG}
    LOG.clear(); STRICT_LOG.clear()
    for j in js:
        x = dict(r7log[j["id"]])
        x["r7_after"] = x["after"]
        if j["id"] in C_EXCLUDE and j.get("r7_rule") == "C":
            j["verdict"], j["r7_rule"], j["r7_note"] = R7.TBD, "C排除→待确认", C_EXCLUDE[j["id"]] + "，需人工确认"
            x.update(after=R7.TBD, rule="C排除→待确认", note=j["r7_note"])
        sc = strict_check(j)
        if sc:
            STRICT_LOG.extend(sc)
            bad = [s for s in sc if not s["satisfied"]]
            if bad:
                b = bad[0]
                j["verdict"], j["r7_rule"] = "负偏离", "S严格不等号"
                j["r7_note"] = f"招标要求{b['op']}{b['req']}{b['unit']}（严格），附件为等于临界值 {b['offer']}"
                x.update(after="负偏离", rule="S严格不等号", note=j["r7_note"])
        LOG.append(x)
    return js


def cite_with_note(base_cite):
    def _c(j):
        s = base_cite(j)
        n = (j or {}).get("r7_note")
        if not n:
            return s
        r = (j or {}).get("r7_rule") or ""
        if r == "人工指定" or r.endswith("守卫") or r.endswith("待确认"):
            return f"【待确认】{n}"
        if r.startswith("S"):
            return f"【负偏离】{n} {s}"
        return f"{n} {s}"
    return _c


def install(bid_r4_fixes_mod, bid_r2_helpers_mod):
    if getattr(bid_r4_fixes_mod.load_judgement_index, "_r8", False):
        return

    def load_with_overlay():
        data = json.loads(bid_r4_fixes_mod.JUDGE_JSON.read_text(encoding="utf-8"))
        js = apply_overlay(data["judgements"])
        by_dev = defaultdict(list)
        for j in js:
            by_dev[int(j["device_seq"])].append(j)
        return {(seq, k): j for seq, lst in by_dev.items() for k, j in enumerate(lst, 1)}
    load_with_overlay._r8 = True
    bid_r4_fixes_mod.load_judgement_index = load_with_overlay
    bid_r2_helpers_mod.cite_evidence = cite_with_note(bid_r2_helpers_mod.cite_evidence)


def device_summary() -> dict[int, str]:
    """偏离情况（设备汇总）：负偏离 > 待确认（含无证据待确认）> 正偏离 > 响应；非全部响应时写明各类条数。"""
    by = defaultdict(Counter)
    for x in LOG:
        a = x["after"]
        by[x["device_seq"]]["待确认" if a in ("待确认", "无证据待确认") else a] += 1
    out = {}
    for seq, c in by.items():
        parts = [f"{k}{c[k]}项" for k in ("负偏离", "待确认", "正偏离") if c.get(k)]
        out[seq] = "响应" if not parts else "；".join(parts)
    return out


def summary() -> dict:
    L = LOG
    b, a = Counter(x["before"] for x in L), Counter(x["after"] for x in L)
    r7a = Counter(x["r7_after"] for x in L)
    agree = lambda key: sum(1 for x in L if x[key] == x["v8"])
    return {"n": len(L), "dist_before": dict(b), "dist_r7": dict(r7a), "dist_after": dict(a),
            "by_rule": dict(Counter(x["rule"] for x in L if x["rule"])),
            "v8_agree_before": agree("before"), "v8_agree_r7": agree("r7_after"), "v8_agree_after": agree("after"),
            "changed": [x for x in L if x["rule"]], "still_disagree": [x for x in L if x["after"] != x["v8"]],
            "strict_check": STRICT_LOG}


# ---- 招标文件 3.3「资格承诺函（仅供参考）」原文（work/r7/tender_main_raw.txt 第 1737–1818 行），照录；
#      本标书补入的条目标注「（待公司确认属实）」
QUAL_TENDER_LINES = (1742, 1818)
ADDED_MARK = "（待公司确认属实）"


def qualification_letter() -> list[tuple[str, bool]]:
    """返回 [(段落文本, 是否r8补入)]。原文按条目合并换行；条目 1（履约能力）与 18（廉洁）r7 已有，不加标注。"""
    raw = (ROOT / "work" / "r7" / "tender_main_raw.txt").read_text(encoding="utf-8").split("\n")
    seg = raw[QUAL_TENDER_LINES[0] - 1:QUAL_TENDER_LINES[1]]
    seg = [l.replace("\f", "") for l in seg]
    seg = [l for l in seg if l.strip() and not re.fullmatch(r"\s*\d{1,3}\s*", l)]  # 去掉页码行
    text = "".join(seg)
    text = text.split("如上述承诺不实")[0]
    items = re.split(r"(?=(?<![\d.（(])(?:1[0-9]|[1-9])\.(?!\d))|(?=（[1-6]）)", text)
    items = [i.strip() for i in items if i.strip()]
    out = []
    for it in items:
        num = re.match(r"(\d+)\.", it)
        existing = bool(num and num.group(1) in ("1", "18")) or (it.startswith("（") and False)
        if it.startswith("我单位郑重作出以下承诺"):
            out.append((it, False)); continue
        out.append((it + ("" if existing else ADDED_MARK), not existing))
    return out


STAR_COMMIT = "我方承诺：完全满足招标文件第三章招标需求第一节商务要求及第二节技术要求的★条款的所有要求。"
