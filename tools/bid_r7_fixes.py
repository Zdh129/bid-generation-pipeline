"""r7：判定规则 A/B/C 叠加层（overlay）+ 明确矛盾守卫 + 人工指定的「待确认」行。

不改 work/deviation_judgements.json：读入原判定后在内存里逐行套规则，产出新结论，
每行都记录 before/after/rule/依据，写到 verdict_overlay_r7.json，前后可追溯。
规则口径（r5 报告第 6 节，donghao 已批 A+B+C 全上）：
  A：原结论「无证据待确认」且理由为「附件有相关表述，但无法可靠判断」→ 响应〔规则A·需人工复核〕
  B：原结论「无证据待确认」且理由为「本机附件中未检索到可引用证据」，且该设备有厂家签章的参数确认函
     → 响应〔规则B·依据厂家参数确认函整体声明，需人工复核〕；没有参数确认函的设备不套 B
  C：原结论「正偏离」且来自数值比较（更优/覆盖并宽于/公差更严）→ 响应〔规则C·优于要求〕；
     附件偏离表自己写「正偏离」的行保留厂家原口径，不套 C
  守卫（明确矛盾优先于 A/B）：招标条款的互斥特征词（手动/自动、活塞/涡旋……）在该设备附件中以对立词
     修饰同一部件出现时，不套 A/B，改标「待确认」并写明矛盾。
  人工指定（donghao 批）：排污阀行 → 待确认，理由「V8 写全自动，厂家函写手动」。
判定时不读 V8；V8 只在事后统计一致率时使用。
"""
from __future__ import annotations

import copy
import glob
import json
import re
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TBD = "待确认"
MANUAL_TBD = {
    "P04-5": "V8 写全自动，厂家函写手动（排污阀），需人工确认",
    "P06-5": "V8 写全自动，厂家函写手动（排污阀），需人工确认",
}
EXCLUSIVE = [("手动", ["全自动", "自动"]), ("自动", ["手动"]), ("活塞", ["涡旋", "螺杆"]), ("涡旋", ["活塞", "螺杆"]),
             ("风冷", ["水冷"]), ("水冷", ["风冷"]), ("单相", ["三相"]), ("三相", ["单相"]), ("无油", ["有油", "含油"]),
             ("立式", ["卧式"]), ("卧式", ["立式"])]
OVERLAY_LOG: list[dict] = []


def _att_text(seq: int, kinds=("confirm", "other", "commitment")) -> dict[str, str]:
    out = {}
    for f in sorted(glob.glob(str(ROOT / "work" / "att_corpus" / f"eq{seq:02d}_*.json"))):
        d = json.loads(Path(f).read_text(encoding="utf-8"))
        if d.get("kind") not in kinds:
            continue
        pts = d.get("page_texts") or []
        out[d["file"]] = " ".join(p if isinstance(p, str) else (p.get("text") or "") for p in pts)
    return out


def contradiction(seq: int, req: str) -> str | None:
    """招标条款里的互斥特征词 a + 部件名 head，附件中出现「对立词 b …… head」且附件该部件处不含 a → 返回矛盾说明。"""
    texts = _att_text(seq)
    for a, bs in EXCLUSIVE:
        for m in re.finditer(a, req):
            if a == "自动" and req[max(0, m.start() - 1):m.start()] in ("全", "手"):
                continue
            tail = re.sub(r"^[式型的]", "", req[m.end():])
            head = re.match(r"[\u4e00-\u9fff]{2,3}", tail)
            if not head:
                continue
            h = head.group(0)
            for fn, t in texts.items():
                t2 = re.sub(r"\s+", "", t)
                if re.search(re.escape(a) + r"[\u4e00-\u9fff式]{0,6}?" + re.escape(h[:2]), t2):
                    continue  # 附件同部件也写了 a，不算矛盾
                for b in bs:
                    mb = re.search(re.escape(b) + r"[\u4e00-\u9fff式]{0,6}?" + re.escape(h[:2]), t2)
                    if mb:
                        return f"招标要求「{a}…{h}」，附件《{fn[:28]}》写「{mb.group(0)}」"
    return None


def has_confirm_letter(seq: int) -> str | None:
    t = _att_text(seq, kinds=("confirm",))
    return next(iter(t), None)


def apply_overlay(judgements: list[dict]) -> list[dict]:
    OVERLAY_LOG.clear()
    out = []
    for j0 in judgements:
        j = copy.deepcopy(j0)
        before = j0.get("verdict") or "无证据待确认"
        reason = j0.get("reason") or ""
        rule, note = None, None
        seq = int(j0["device_seq"])
        if j0["id"] in MANUAL_TBD:
            rule, j["verdict"], note = "人工指定", TBD, MANUAL_TBD[j0["id"]]
        elif before == "无证据待确认" and (reason.startswith("附件有相关表述") or reason.startswith("本机附件中未检索到")):
            isA = reason.startswith("附件有相关表述")
            c = contradiction(seq, j0["req_text"])
            if c:
                rule, j["verdict"], note = ("A" if isA else "B") + "→守卫", TBD, f"明确矛盾：{c}，需人工确认"
            elif isA:
                rule, j["verdict"], note = "A", "响应", "〔规则A·附件有相关表述，需人工复核〕"
            else:
                cf = has_confirm_letter(seq)
                if cf:
                    rule, j["verdict"], note = "B", "响应", f"〔规则B·依据厂家参数确认函整体声明，需人工复核〕《{cf[:30]}》"
        elif before == "正偏离" and reason.startswith("数值比较") and re.search(r"更优|覆盖并宽于|公差更严", reason):
            rule, j["verdict"], note = "C", "响应", "〔规则C·优于要求〕"
        if rule:
            j["r7_rule"], j["r7_note"] = rule, note
        OVERLAY_LOG.append({"id": j0["id"], "device_seq": seq, "star": bool(j0.get("star")), "req_text": j0["req_text"][:80],
                            "before": before, "after": j.get("verdict") or "无证据待确认", "rule": rule, "note": note,
                            "orig_reason": reason[:120], "v8": j0.get("v8_deviation")})
        out.append(j)
    return out


def cite_with_note(base_cite):
    def _c(j):
        s = base_cite(j)
        n = (j or {}).get("r7_note")
        if not n:
            return s
        r = (j or {}).get("r7_rule") or ""
        if r == "人工指定" or r.endswith("守卫"):
            return f"【待确认】{n}"
        return f"{n} {s}"
    return _c


def install(bid_r4_fixes_mod, bid_r2_helpers_mod):
    """把 overlay 装到 r4 对位加载器与证据引用函数上：生成前和生成后断言用同一份 overlay 结论。"""
    orig_load = bid_r4_fixes_mod.load_judgement_index
    if getattr(orig_load, "_r7", False):
        return

    def load_with_overlay():
        data = json.loads(bid_r4_fixes_mod.JUDGE_JSON.read_text(encoding="utf-8"))
        js = apply_overlay(data["judgements"])
        from collections import defaultdict
        by_dev = defaultdict(list)
        for j in js:
            by_dev[int(j["device_seq"])].append(j)
        idx = {}
        for seq, lst in by_dev.items():
            for k, j in enumerate(lst, 1):
                idx[(seq, k)] = j
        return idx
    load_with_overlay._r7 = True
    bid_r4_fixes_mod.load_judgement_index = load_with_overlay
    bid_r2_helpers_mod.cite_evidence = cite_with_note(bid_r2_helpers_mod.cite_evidence)


def overlay_summary() -> dict:
    L = OVERLAY_LOG
    b, a = Counter(x["before"] for x in L), Counter(x["after"] for x in L)
    agree_b = sum(1 for x in L if x["before"] == x["v8"]); agree_a = sum(1 for x in L if x["after"] == x["v8"])
    return {"n": len(L), "dist_before": dict(b), "dist_after": dict(a), "by_rule": dict(Counter(x["rule"] for x in L if x["rule"])),
            "v8_agree_before": agree_b, "v8_agree_after": agree_a,
            "rate_before": round(agree_b / len(L), 4), "rate_after": round(agree_a / len(L), 4),
            "changed": [x for x in L if x["rule"]],
            "still_disagree": [x for x in L if x["after"] != x["v8"]],
            "rule_flips_v8_negative": [x for x in L if x["rule"] in ("B", "B→守卫", "A", "A→守卫") and x["v8"] == "负偏离"]}
