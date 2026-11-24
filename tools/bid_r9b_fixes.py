"""r9：在 r8 叠加层上执行 donghao t17u「这几条你想办法解决」。仍不改 deviation_judgements.json。
- 人工核对行改判（证据写进证据栏，单列清单）；
- 待确认行全部定案后，按招标要求把结论用语「响应」改为「无偏离」（仅叠加层副本，V8 对照同步归一）；
- 资格承诺函逐条核对证据并改标注；
- 投标函抬头、交货期、质保期按招标文件填值并标出处。"""
from __future__ import annotations

import json
import re
from collections import Counter, defaultdict
from pathlib import Path

import bid_r7_fixes as R7
import bid_r8_fixes as R8

ROOT = Path(__file__).resolve().parents[1]
TENDER_PDF = "example-tender.pdf"
ATT3 = "annex3-example-params.pdf/.txt"

# ---------------- 招标文件取值（出处：招标文件 PDF 物理页 = 页脚印刷页码；行号为 work/r7/tender_main_raw.txt）
FILLED = {
    "addressee": {"value": "示例招标人有限公司", "page": 8, "where": "第二章投标人须知前附表第 1 项「招标人 名称：示例招标人有限公司」",
                  "lines": "218-222", "also": "招标公告第 1 页「招 标 人 ： 示例招标人有限公司」"},
    "delivery": {"value": "合同签订之后，接招标人通知之日起 30 日历日内完成供货，并通过验收", "page": 33,
                 "where": "第三章招标需求第一节商务要求第 1 条", "lines": "1070-1071", "also": "招标公告第 3 页第 3 条（行 88-89）"},
    "warranty": {"value": "验收合格之日起 1 年", "page": 33, "where": "第三章招标需求第一节商务要求第 3 条", "lines": "1073",
                 "also": "招标公告第 3 页第 5 条（行 91）；评分办法示例页「质保期可按招标文件规则加分」"},
}


def cite(key):
    f = FILLED[key]
    return f"（依据招标文件第 {f['page']} 页）"


# ---------------- 人工核对行改判（证据均为本机资料包原文）
R9B_REVERT = {"P04-5", "P06-5"}  # r9b：donghao 决定排污阀两行改回待确认
R9B_NOTE = "V8 写全自动，厂家函写手动，待人工确认"
R9_CHANGES = {
    "P18-1": ("正偏离", "招标「最大量程：≥100 单位」（附件3 设备18 第1条）；厂家函《示例-18#示例设备18-技术参数确认函》第1页「最大量程：120 单位」，120＞100，"
              "V8 同判正偏离（V8 写「最大量程：120 单位」）"),
    "P36-5": ("正偏离", "招标「★示例强度:0-1,000 六级可调」；厂家函《示例-36#示例设备36-参数确认函》第1页「★示例强度:0-1500（100级调控）」，"
              "彩页《示例-36#示例设备36-彩页资料》第2页「额定示例强度: 1500（100级调控）」，量程覆盖且调节级数更细；★条款允许正偏离，V8 同判正偏离"),
    "P28-3": ("负偏离", "招标「示例A型压缩机」；厂家函《示例-28#示例设备28-参数确认函》第1页「3.示例B型压缩机」，压缩机型式与要求不同；"
              "彩页只写「示例压缩机」未写活塞式；V8 同判负偏离（非★条款）"),
    "P04-5": ("响应", "招标「设计有示例手动阀」；厂家函《4#示例设备04-MODEL-A-技术参数确认函》第1页逐字为「5 设计有示例手动阀」，满足要求；"
              "V8 写「全自动示例阀」并判正偏离，但资料包内厂家函无此说法（彩页第3页「自动执行，无须人工干预」不否定有手动排污阀），按厂家盖章函判定"),
    "P06-5": ("响应", "招标「主机设计有示例手动阀」；厂家函《示例-6#示例设备06-MODEL-B-技术参数确认函》第1页逐字为「5 主机设计有示例手动阀」，满足要求；"
              "V8 写「全自动示例阀」并判正偏离，资料包内厂家函无此说法，按厂家盖章函判定"),
    "P44-9": ("负偏离", "招标原文「最大示例压力：＞100 单位」（示例参数表第 N 行）；同一份厂家函《示例-44#示例设备44-技术参数确认函》第1页另写「示例压力：0 - 20（并联）、0 - 25（串联）可调」，"
              "实际真空度最高 95 kPa＜400 kPa，按招标原文字面不满足；V8 同判负偏离。注：真空度物理上限约 101 kPa，＞400 kPa 不可能达到，招标单位疑有误（若为 mbar/hPa 即 40 kPa、"
              "或 mmHg 即 53.3 kPa，则 95 kPa 满足），建议向招标人澄清"),
    "P40-11": ("响应", "招标配置清单「主机1 台」；设备40 资料《示例-40#示例设备40-MODEL-C-参数》第1页与彩页第1页均为 MODEL-C 主机本体参数，"
               "同页附属品「隔板 2 件、隔板架 4 件」与配置清单第 2、3 项一致（P40-12、P40-13 已判响应）；所投即主机 1 台，V8 同判无偏离"),
}
LOG: list[dict] = []
VOCAB = {"changed": False}


def apply_overlay(judgements: list[dict]) -> list[dict]:
    js = R8.apply_overlay(judgements)
    r8log = {x["id"]: x for x in R8.LOG}
    LOG.clear()
    for j in js:
        x = dict(r8log[j["id"]]); x["r8_after"] = x["after"]
        if j["id"] in R9_CHANGES:
            v, why = R9_CHANGES[j["id"]]
            j["verdict"], j["r7_rule"], j["r7_note"] = v, "R9改判", why
            x.update(after=v, rule="R9改判", note=why)
        if j["id"] in R9B_REVERT:
            j["verdict"], j["r7_rule"], j["r7_note"] = "待确认", "R9b改回待确认", R9B_NOTE
            x.update(after="待确认", rule="R9b改回待确认", note=R9B_NOTE)
        LOG.append(x)
    pending = [j["id"] for j in js if j["verdict"] in ("待确认", "无证据待确认")]
    VOCAB["pending"] = pending
    VOCAB["changed"] = True  # r9b：保持 r9 的「无偏离」用语，仅 P04-5/P06-5 两行为待确认
    if VOCAB["changed"]:  # 全部定案：结论用语按招标「正偏离/负偏离/无偏离」
        for j, x in zip(js, LOG):
            if j["verdict"] == "响应":
                j["verdict"] = "无偏离"; x["after"] = "无偏离"
            if j.get("v8_deviation") == "响应":
                j["v8_deviation"] = "无偏离"
            x["v8"] = "无偏离" if x["v8"] == "响应" else x["v8"]
            for k in ("before", "r7_after", "r8_after"):
                x[k + "_vocab"] = "无偏离" if x[k] == "响应" else x[k]
    return js


def cite_with_note(base_cite):
    r8c = R8.cite_with_note(base_cite)

    def _c(j):
        if (j or {}).get("r7_rule") == "R9b改回待确认":
            return R9B_NOTE
        if (j or {}).get("r7_rule") == "R9改判":
            return f"〔r9人工核对改判〕{j['r7_note']}；" + base_cite(j)
        return r8c(j)
    return _c


def install(bid_r4_fixes_mod, bid_r2_helpers_mod):
    if getattr(bid_r4_fixes_mod.load_judgement_index, "_r9b", False):
        return
    base = bid_r2_helpers_mod.cite_evidence

    def load_with_overlay():
        data = json.loads(bid_r4_fixes_mod.JUDGE_JSON.read_text(encoding="utf-8"))
        js = apply_overlay(data["judgements"])
        by_dev = defaultdict(list)
        for j in js:
            by_dev[int(j["device_seq"])].append(j)
        return {(seq, k): j for seq, lst in by_dev.items() for k, j in enumerate(lst, 1)}
    load_with_overlay._r9b = True
    bid_r4_fixes_mod.load_judgement_index = load_with_overlay
    bid_r2_helpers_mod.cite_evidence = cite_with_note(base)


def device_summary() -> dict[int, str]:
    ok = "无偏离" if VOCAB["changed"] else "响应"
    by = defaultdict(Counter)
    for x in LOG:
        a = x["after"]
        by[x["device_seq"]]["待确认" if a in ("待确认", "无证据待确认") else a] += 1
    out = {}
    for seq, c in by.items():
        parts = [f"{k}{c[k]}项" for k in ("负偏离", "待确认", "正偏离") if c.get(k)]
        out[seq] = ok if not parts else "；".join(parts)
    return out


def summary() -> dict:
    L = LOG
    key = lambda x, k: x.get(k + "_vocab", x[k])
    agree = lambda k: sum(1 for x in L if (x["after"] if k == "after" else key(x, k)) == x["v8"])
    return {"n": len(L), "vocab_changed_to_无偏离": VOCAB["changed"], "pending_after": VOCAB.get("pending"),
            "dist_before": dict(Counter(x["before"] for x in L)), "dist_r7": dict(Counter(x["r7_after"] for x in L)),
            "dist_r8": dict(Counter(x["r8_after"] for x in L)), "dist_after": dict(Counter(x["after"] for x in L)),
            "by_rule": dict(Counter(x["rule"] for x in L if x["rule"])),
            "v8_agree_before": agree("before"), "v8_agree_r7": agree("r7_after"), "v8_agree_r8": agree("r8_after"), "v8_agree_after": agree("after"),
            "r9_changes": [{**{k: x[k] for k in ("id", "device_seq", "star", "req_text", "r8_after", "after", "v8", "note")}} for x in L if x["rule"] in ("R9改判", "R9b改回待确认")],
            "changed": [x for x in L if x["rule"]], "still_disagree": [x for x in L if x["after"] != x["v8"]],
            "strict_check": R8.STRICT_LOG}


# ---------------- 资格承诺函逐条核对（证据只在本机读取；报告中不出现证件号码、个人姓名）
EV = {
    "LIC": "营业执照（资料包 2.投标主体/example-business-license.pdf 第1页，本机 400dpi OCR：有限责任公司，成立日期 2000年01月01日，注册资本 1000万元，执照核发 2024年01月01日）",
    "TAX": "2026年纳税和社保缴纳凭证（example-tax-social-security.pdf 第1-5页：所属期示例年4-7月，转账日期示例区间，税种含增值税、城建税及社会保险）",
    "ISO": "ISO 9001/14001/45001 认证证书（QMS/EMS/OHSMS 主证，各第1页：有效期至 2099年12月31日）",
    "PERF": "业绩合同及发票（example-performance-contracts.pdf，共若干页；示例合同金额均达到招标要求）",
    "SOLE": "本投标文件与 V8 均以示例投标人有限公司单一主体投标，资料包内无联合体协议",
    "STAR": "6.1 节 111 条★参数：负偏离 0 条（r9 定案后），均附厂家参数确认函/偏离表逐条对应；其中 5 条（P08-7、P19-13、P20-8、P22-4、P22-5）依规则A/B归并，需人工复核",
    "BrandP": "BrandP（示例厂家丙）5 台设备（序号3、5、9、10、17）的厂家承诺函（落款为示例厂家丙（示例地）有限公司，OCR 读数）承诺「到货时提供安装使用、维修所需的全部技术资料，验收时提供正品验证渠道或证明」",
}
COMMIT = {  # key: (分类, [证据键], 说明)
    "2": ("无证据但不矛盾", [], "关联单位是否同时投标需查股权关系，资料包无相关材料"),
    "3": ("有证据支持", ["LIC", "TAX"], "执照在有效期内，示例期间持续纳税缴社保，说明未停业；是否被暂扣/吊销仍需在公示系统复核"),
    "4": ("无证据但不矛盾", ["TAX"], "持续纳税缴社保与「未清算/破产」不矛盾，但资料包无清算/破产查询结果"),
    "5": ("无证据但不矛盾", [], "需「国家企业信用信息公示系统」查询截图，资料包与 V8 均无"),
    "6": ("无证据但不矛盾", [], "需「信用中国」/中国执行信息公开网查询截图，资料包与 V8 均无"),
    "7": ("无证据但不矛盾", [], "需无行贿犯罪记录查询或声明，资料包与 V8 均无"),
    "8": ("无证据但不矛盾", [], "招标人「暂停合作名单」本地无法核对"),
    "9": ("无证据但不矛盾", [], "中标后行为承诺"),
    "10": ("有证据支持", ["SOLE"], "单一主体投标"),
    "11": ("无证据但不矛盾", [], "中标后行为承诺"),
    "12": ("无证据但不矛盾", [], "接受核查的承诺"),
    "13": ("有证据支持", ["PERF", "ISO", "LIC"], "5 个百万以上业绩、三体系认证有效、注册资本 1000 万元"),
    "14": ("无证据但不矛盾", [], "中标后行为承诺"),
    "15": ("无证据但不矛盾", [], "资料包无保密制度文件（三体系证书不含信息安全）"),
    "16": ("无证据但不矛盾", [], "所投均为厂家品牌产品并附厂家确认函，但无知识产权声明"),
    "17": ("有证据支持", ["STAR"], "★条款无负偏离"),
    "18(1)": ("无证据但不矛盾", [], "廉洁行为承诺"), "18(2)": ("无证据但不矛盾", [], "廉洁行为承诺"),
    "18(3)": ("无证据但不矛盾", [], "廉洁行为承诺"), "18(4)": ("无证据但不矛盾", [], "廉洁行为承诺"),
    "19": ("无证据但不矛盾", [], "总述，见各分项"),
    "19(1)": ("无证据但不矛盾", [], "供货质量承诺，需中标后兑现"),
    "19(2)": ("无证据但不矛盾", [], "供货资料承诺，需中标后兑现；r8 曾把设备42 文件判为外文，r9 用 400dpi OCR 核实为中文（文字层字体编码错乱所致），不构成矛盾"),
    "19(3)": ("无证据但不矛盾", [], "三包承诺；V8 第 5.2 节亦写包修、包换、包退（属同类承诺，不是证据）"),
    "19(4)": ("无证据但不矛盾", [], "包装责任承诺"),
    "19(5)": ("无证据但不矛盾", [], "中标后检定校准承诺"),
    "19(6)": ("有证据支持", ["BrandP"], "仅 BrandP 5 台有厂家承诺函；其他进口设备（如 示例厂家丁、BrandH、BrandI、BrandA、BrandK、BrandL、示例厂家戊）无同类函"),
    "3.4": ("有证据支持", ["STAR"], "同第 17 条"),
}
COMMIT_LOG: list[dict] = []


def qualification_letter_r9() -> list[tuple[str, str | None]]:
    """返回 [(段落文本, 分类)]；有证据支持的改标注「（已对照…证据，待公司最终确认）」。"""
    COMMIT_LOG.clear()
    out, main = [], None
    for t, added in R8.qualification_letter():
        m = re.match(r"(\d+)\.", t)
        sub = re.match(r"（(\d)）", t)
        if m:
            main = m.group(1); key = main
        elif sub and main:
            key = f"{main}({sub.group(1)})"
        else:
            key = None
        if not added or key not in COMMIT:
            out.append((t, None)); continue
        cls, evk, why = COMMIT[key]
        if cls == "有证据支持":
            names = "、".join({"LIC": "营业执照", "TAX": "2026年纳税社保凭证", "ISO": "三体系认证证书", "PERF": "业绩合同",
                              "SOLE": "投标主体资料", "STAR": "6.1 厂家参数确认函", "BrandP": "BrandP 厂家承诺函"}[k] for k in evk)
            t = t.replace(R8.ADDED_MARK, f"（已对照{names}证据，待公司最终确认）")
        COMMIT_LOG.append({"item": key, "class": cls, "evidence": [EV[k] for k in evk], "note": why, "text": t[:60]})
        out.append((t, cls))
    return out


def star_commit_text() -> tuple[str, dict]:
    cls, evk, why = COMMIT["3.4"]
    t = R8.STAR_COMMIT + "（已对照6.1 厂家参数确认函证据，待公司最终确认）"
    return t, {"item": "3.4（完全满足★条款）", "class": cls, "evidence": [EV[k] for k in evk], "note": why, "text": t[:60]}
