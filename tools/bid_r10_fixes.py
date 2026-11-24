"""r10 叠加层（在 r9b 之上；deviation_judgements.json 不改）。
donghao t19u：「单价、承诺的最终确认和信用查询截图、品牌缺 7 项、型号缺 4 项，以及设备 34 和 40 的参数确认函。
这些没有就先标记为未提供，不计入一致性判定」。
- 标书中：上述资料统一写「未提供」；示例厂家甲/示例厂家乙品牌冲突（设备 23/33/39/41）不属于「缺」，保持【待填】原样。
- 一致性判定：依赖未提供资料的参数行（设备 34、40 全部行：其证据只有彩页/参数页，招标要求的厂家参数确认函未提供）
  单独记「未提供，不计入」，从 V8 一致率与结论分布的分母剔除；结论文字本身不改（不算负偏离、不算不一致）。
- 排污阀 P04-5、P06-5 仍为「待确认」（r9b），不属于未提供。
"""
from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

import bid_r9b_fixes as R9B

NP = "未提供"
NP_COMMIT = "【未提供：公司最终确认】"
NOT_PROVIDED = {
    "unit_price": {"what": "45 项投标单价、合价及投标总价", "where": "投标函附录-分项报价表「单价」「合价」栏、表后「投标总价」行",
                   "basis": "V8 报价表单价/小计栏全为 ____，资料包仅有招标方《最高限价.xls》（r9 检索记录 work/r9/price_search.json）"},
    "commit_final": {"what": "28 条承诺的公司最终确认（3.3 第 2–19 条及分项 27 条 + 3.4「完全满足★条款」1 条）",
                     "where": "3.3 资格承诺函、3.4 承诺书各条末尾标记",
                     "basis": "r9 核对：6 条有资料包证据、22 条无证据但不矛盾，均缺公司最终确认"},
    "credit_screens": {"what": "信用查询截图 3 张", "items": ["国家企业信用信息公示系统（www.gsxt.gov.cn）", "信用中国（www.creditchina.gov.cn）", "中国执行信息公开网（zxgk.court.gov.cn）"],
                       "where": "3.2 投标人资格相关资料末尾 3 个截图位置", "basis": "资料包与 V8 均无查询截图"},
    "brand": {"what": "品牌（缺）", "devices": [20, 34, 40], "where": "分项报价表「品牌」栏",
              "basis": "r9 交叉校验无 ≥2 个一致来源；设备 23/33/39/41 为确认函「示例厂家甲」与 V8「示例厂家乙」冲突，按要求不列入，保持【待填】"},
    "model": {"what": "型号（缺）", "devices": [18, 20, 23, 29], "where": "分项报价表「型号」栏", "basis": "r9 交叉校验无 ≥2 个一致来源"},
    "confirm_letter": {"what": "厂家参数确认函", "devices": [34, 40], "where": "6.1 附件 6.1.34、6.1.40 证明材料位置；6.1.34、6.1.40 参数表证据栏",
                       "basis": "资料包设备 34 只有彩页，设备 40 只有参数页与彩页"},
}
CONFLICT_BRAND = [23, 33, 39, 41]
EXCL_DEVICES = set(NOT_PROVIDED["confirm_letter"]["devices"])
EXCL_TAG = "〔未提供：厂家参数确认函（本行不计入一致性判定）〕"
EXCL_REASON = "依赖未提供的厂家参数确认函：本行证据只有彩页/参数页，示例招标文件相应条款要求非★参数以厂家参数确认函等作证明"
LOG10: list[dict] = []

# 继承 r9b
FILLED, cite, device_summary, COMMIT, EV, VOCAB = R9B.FILLED, R9B.cite, R9B.device_summary, R9B.COMMIT, R9B.EV, R9B.VOCAB
COMMIT_LOG: list[dict] = []


def apply_overlay(judgements):
    js = R9B.apply_overlay(judgements)
    LOG10.clear()
    for j in js:
        ex = int(j["device_seq"]) in EXCL_DEVICES
        j["r10_excluded"] = ex
        LOG10.append({"id": j["id"], "device_seq": int(j["device_seq"]), "verdict": j["verdict"],
                      "v8": j.get("v8_deviation"), "excluded": ex, "reason": EXCL_REASON if ex else None})
    return js


def cite_with_note(base_cite):
    c9 = R9B.cite_with_note(base_cite)

    def _c(j):
        s = c9(j)
        return f"{EXCL_TAG} {s}" if (j or {}).get("r10_excluded") else s
    return _c


def install(bid_r4_fixes_mod, bid_r2_helpers_mod):
    if getattr(bid_r4_fixes_mod.load_judgement_index, "_r10", False):
        return
    base = bid_r2_helpers_mod.cite_evidence
    from collections import defaultdict

    def load_with_overlay():
        data = json.loads(bid_r4_fixes_mod.JUDGE_JSON.read_text(encoding="utf-8"))
        js = apply_overlay(data["judgements"])
        by_dev = defaultdict(list)
        for j in js:
            by_dev[int(j["device_seq"])].append(j)
        return {(seq, k): j for seq, lst in by_dev.items() for k, j in enumerate(lst, 1)}
    load_with_overlay._r10 = True
    bid_r4_fixes_mod.load_judgement_index = load_with_overlay
    bid_r2_helpers_mod.cite_evidence = cite_with_note(base)


def consistency():
    L = LOG10
    keep = [x for x in L if not x["excluded"]]
    exc = [x for x in L if x["excluded"]]
    ag = lambda xs: sum(1 for x in xs if x["verdict"] == x["v8"])
    return {
        "before": {"n": len(L), "agree_v8": ag(L), "rate": round(ag(L) / len(L), 4), "dist": dict(Counter(x["verdict"] for x in L))},
        "after": {"n": len(keep), "agree_v8": ag(keep), "rate": round(ag(keep) / len(keep), 4), "dist": dict(Counter(x["verdict"] for x in keep)),
                  "未提供，不计入": len(exc)},
        "excluded": [{k: x[k] for k in ("id", "device_seq", "verdict", "v8", "reason")} for x in exc],
        "still_disagree_after": [{k: x[k] for k in ("id", "verdict", "v8")} for x in keep if x["verdict"] != x["v8"]],
        "note": "被剔除行的结论文字保持 r9b 原样（不改判、不算负偏离或不一致），只在证据栏加注并从分母剔除；P04-5、P06-5 仍为待确认并计入分母",
    }


def summary():
    s = R9B.summary()
    s["r10_consistency"] = consistency()
    s["r10_not_provided"] = NOT_PROVIDED
    return s


def _mark(t):
    t = t.replace(R9B.R8.ADDED_MARK, NP_COMMIT)
    return re.sub(r"（已对照(.+?)证据，待公司最终确认）", lambda m: f"（已对照{m.group(1)}证据）{NP_COMMIT}", t)


import re  # noqa: E402


def qualification_letter_r10():
    out = []
    COMMIT_LOG.clear()
    for t, cls in R9B.qualification_letter_r9():
        out.append((_mark(t), cls))
    for x in R9B.COMMIT_LOG:
        COMMIT_LOG.append({**x, "r10_final_confirmation": NP, "text": _mark(x["text"])})
    return out


def star_commit_text():
    t, log = R9B.star_commit_text()
    return _mark(t), {**log, "r10_final_confirmation": NP, "text": _mark(log["text"])}
