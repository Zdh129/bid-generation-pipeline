"""r10 废标项：r9b 口径之上，依赖未提供资料（donghao t19u：单价、承诺最终确认与信用查询截图、品牌/型号缺项、设备34/40 参数确认函）的项
记「未提供，不计入」，从统计分母剔除，同时保留按 r9b 口径的原状态（status_r9b_rule）以便对照。
r9 废标项（r8 口径不变；「已对照…证据，待公司最终确认」仍记待人工确认；结论用语、O3 外文判断按 r9 核实结果）。
r8：（在 r7 映射基础上：补入的承诺标注「待公司确认属实」者不算已核实；响应值为【待填】者不算已响应）。
r7 原说明：废标项 / 无效投标 / ★ / 实质性条款 → 生成标书位置映射（纯本地规则，不调用模型）。
输入：官方招标文件文本（pdftotext 抽取，work/r7/tender_main_raw.txt，行号即出处）、生成标书 DOCX/PDF、
      r7 判定叠加结果（verdict_overlay_r9.json）、附件语料（work/att_corpus）。
状态：已响应 = 标书中找到全部要件；缺失 = 一个要件都找不到；待人工确认 = 部分找到，或要件本身需盖章/签字/凭证/平台操作/外部核查，
      本地无法判定。
"""
from __future__ import annotations

import glob
import json
import re
import sys
from pathlib import Path

import fitz
from docx import Document

ROOT = Path(__file__).resolve().parents[1]
OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "output" / "generated-r10"
TENDER = (ROOT / "work" / "r7" / "tender_main_raw.txt").read_text(encoding="utf-8").splitlines()

# (编号, 类别, 出处行, 条款摘要, 后果, 检查方式)
#   检查方式 doc：[正则组…] 全部命中=已响应、全不中=缺失、部分=待人工确认；manual：本地不可判定；derived：由其他结果推出
C = [
    ("Z1", "资格性审查", "98-100", "营业执照（或其他组织证照）复印件/扫描件，加盖公章", "资格不符→否决（5.4.4-12）", ("doc", [r"附件：营业执照"])),
    ("Z2", "资格性审查", "101-102", "2026年1月至今任意三个月纳税和社保缴纳凭证，加盖公章", "同上", ("doc", [r"纳税及社会保险缴纳凭证"])),
    ("Z3", "资格性审查", "103-104", "具有履行合同所必备的能力（自行承诺，加盖公章）", "同上", ("doc", [r"履约能力|履行合同所必备的能力"])),
    ("Z4-1", "资格性审查", "106-111", "承诺：单位负责人非同一人/无控股、管理关系的投标人同时投标", "查实→否决", ("doc", [r"单位负责人为同一人|控股、?管理关系"])),
    ("Z4-2", "资格性审查", "112-114", "承诺：未被责令停业、暂扣或吊销执照/资质", "资格不符→否决", ("doc", [r"停业|吊销"])),
    ("Z4-3", "资格性审查", "115-116", "承诺：未进入清算程序、未被宣告破产", "同上", ("doc", [r"清算|破产"])),
    ("Z4-4", "资格性审查", "117-121", "承诺：未被列入严重违法失信企业名单", "同上", ("doc", [r"严重违法失信"])),
    ("Z4-5", "资格性审查", "125-128", "承诺：未被列入失信被执行人名单", "同上", ("doc", [r"失信被执行人"])),
    ("Z4-6", "资格性审查", "129-131", "承诺：投标人及法定代表人2023年1月1日起无行贿犯罪行为", "同上", ("doc", [r"行贿犯罪"])),
    ("Z4-7", "资格性审查", "132-135", "承诺：未被列入招标人「暂停合作名单」", "同上", ("doc", [r"暂停合作"])),
    ("Z5", "资格性审查", "136-137", "承诺：不分包、不转包", "同上", ("doc", [r"分包|转包"])),
    ("Z6", "资格性审查", "138", "不接受联合体投标（资格承诺函第10条）", "同上", ("doc", [r"联合体"])),
    ("H1", "符合性审查/★", "629-631;1098-1100;1186-1187", "承诺「完全满足」第三章第一节商务要求及第二节技术要求的★条款所有要求（加盖公章）",
     "未提供视为无效标/可否决", ("doc", [r"★条款", r"完全满足"])),
    ("H2-1", "符合性审查", "605-607", "承诺按《示例招标人有限公司第三方管理办法》开展业务（前附表33-1）", "未提供承诺可否决", ("doc", [r"第三方管理办法"])),
    ("H2-2", "符合性审查", "608-611", "承诺接受评标后实地考察核实，弄虚作假取消资格（33-2）", "同上", ("doc", [r"实地考察|考察和核实|考察核实"])),
    ("H2-3", "符合性审查", "612", "承诺在人员、设备、资金等方面具有相应服务能力（33-3）", "同上", ("doc", [r"人员、设备、资金|服务能力"])),
    ("H2-4", "符合性审查", "615-616", "承诺不减少或降低招标人拟定的工作内容、工作量和质量（33-4）", "同上", ("doc", [r"不.{0,4}(减少|降低).{0,12}工作"])),
    ("H2-5", "符合性审查", "617-619", "承诺具有完善保密制度和措施（33-5）", "同上", ("doc", [r"保密"])),
    ("H2-6", "符合性审查", "620-623", "承诺不侵犯第三方知识产权并承担赔偿责任（33-6）", "同上", ("doc", [r"知识产权"])),
    ("H3", "符合性审查", "632-648", "廉洁响应承诺（不宴请送礼、不安排旅游娱乐、不购置物品、不行贿）（前附表34）", "未提供承诺可否决", ("doc", [r"宴请", r"娱乐", r"行贿"])),
    ("H4-1", "符合性审查", "650-657", "承诺所供货物全新、未使用、合格，质保期内对缺陷负责（前附表35-1）", "未提供承诺可否决", ("doc", [r"全新|未使用过"])),
    ("H4-2", "符合性审查", "658-659", "承诺供货时以中文附主要图样、合格证明、使用维护说明（35-2）", "同上", ("doc", [r"中文", r"合格证明"])),
    ("H4-3", "符合性审查", "663-666", "承诺质保期内包修、包换、包退并承担费用，符合「三包」（35-3）", "同上", ("doc", [r"维修|包修", r"更换|包换", r"退货|包退"])),
    ("H4-4", "符合性审查", "667-668", "承诺包装方式和费用由中标人负责，运输损坏丢失由投标人负责（35-4）", "同上", ("doc", [r"包装"])),
    ("H4-5", "符合性审查", "669-670", "承诺聘请有资质第三方对计量器具检定校准并提供证书（35-5）", "同上", ("doc", [r"检定校准"])),
    ("H4-6", "符合性审查", "671-673", "承诺进口设备提供安装维修技术资料及正品验证渠道（35-6）", "同上", ("doc", [r"正品验证"])),
    ("H5", "符合性审查", "676-687;1621-1625", "分项报价表按清单填写每项品牌（或生产厂家）及型号，不得宽泛表述，同一标的唯一品牌唯一报价", "未按要求填写的作无效标", ("quote_cols", [r"品牌|生产厂家", r"型号"])),
    ("H6", "符合性审查", "1195-1201", "投标函由法定代表人或委托代理人签字（盖章）并加盖公章；附法定代表人身份证明或授权委托书", "符合性不通过", ("manual_doc", [r"法定代表人身份证明"])),
    ("W1", "无效标5.4.4", "895;前附表17(L330-360)", "按要求提交投标保证金（包2：80,000.00元，银行转账或电子保函）", "无效标", ("manual_doc", [r"投标保证金"])),
    ("W2", "无效标5.4.4", "896;前附表15", "总价不超过包2最高限价4,328,261.00元，单价不超过最高限价单价", "无效标", ("manual_doc", [r"最高限价"])),
    ("W3", "无效标5.4.4", "897", "不得提供虚假材料", "无效标", ("manual_doc", [r"真实有效"])),
    ("W4", "无效标5.4.4", "898;808-809", "投标有效期：自投标截止时间起90日历天", "无效标/否决", ("doc", [r"有效期", r"90\s*个?日历[天日]"])),
    ("W5", "无效标5.4.4", "899-900", "只提交一份投标文件，同一标的同一要求不得有两个及以上报价", "无效标", ("doc", [r"一次性(固定)?报价|唯一报价"])),
    ("W6", "无效标5.4.4", "901-902", "投标文件与获取招标文件的单位名称和组织结构一致", "无效标", ("manual", [])),
    ("W7", "无效标5.4.4", "903-904", "报价不得明显低于成本", "无效标", ("manual", [])),
    ("W8", "无效标5.4.4", "905-906", "不得以他人名义投标、串通投标、行贿、弄虚作假", "无效标", ("manual_doc", [r"行贿"])),
    ("W9", "无效标5.4.4", "907", "按招标文件要求提供有关承诺", "无效标", ("derived", ["Z3", "Z4-1", "Z4-2", "Z4-3", "Z4-4", "Z4-5", "Z4-6", "Z4-7", "Z5", "Z6", "H1", "H2-1", "H2-2", "H2-3", "H2-4", "H2-5", "H2-6", "H3", "H4-1", "H4-2", "H4-3", "H4-4", "H4-5", "H4-6"])),
    ("W10", "无效标5.4.4", "908-909;881-889", "对招标文件要求和条件作出实质性响应，无实质性偏离", "无效标", ("derived", ["T1", "T3", "O5", "O6"])),
    ("W12", "无效标5.4.4", "910-914", "资格条件符合要求", "无效标", ("derived", ["Z1", "Z2", "Z3", "Z4-1", "Z4-2", "Z4-3", "Z4-4", "Z4-5", "Z4-6", "Z4-7", "Z5", "Z6"])),
    ("T1", "★技术参数", "1101-1103;1630-1631", "★条款设备（技术）参数不允许负偏离", "该项不满足→视为未提供承诺→无效标", ("star_rows", [])),
    ("T2", "★技术参数", "1101-1103;1630-1631", "★条款中要求提供其他证明材料的，须按要求提供", "同上", ("star_proof", [])),
    ("T3", "非★技术参数", "1632-1634", "非★参数须提供厂家参数确认函/检测报告/技术白皮书/产品说明书作为证明", "该项视为不满足", ("device_proof", [])),
    ("T4", "报价表格式", "1635", "偏离情况栏填写「正偏离」「负偏离」或「无偏离」", "格式不符（可能被认定未按要求填写）", ("verdict_vocab", [])),
    ("T5", "报价表格式", "1590-1627", "分项报价表含「招标文件要求参数 / 投标设备参数 / 偏离情况 / 证明材料页码范围」栏", "未按要求填写的作无效标（注3/注5）", ("quote_cols", [r"偏离", r"页码"])),
    ("O1", "递交/开标", "487-492;841-845", "平台上传加密电子投标文件(.TBJ)，开标后30分钟内解密", "解密失败视为无效投标", ("manual", [])),
    ("O2", "签字盖章", "前附表19", "要求签字盖章处均签字盖章；行间插字、涂改处加盖单位章或法定代表人印章", "符合性不通过", ("manual", [])),
    ("O3", "语言", "768-770", "投标文件使用中文；非中文资料须附中文译本", "可能被认定未实质性响应", ("lang", [])),
    ("O5", "商务要求(★承诺覆盖)", "1068-1070", "交货期：合同签订后接通知之日起30日历日内完成供货并通过验收", "实质性条款", ("doc", [r"交货期", r"30\s*个?日历[日天]"])),
    ("O6", "商务要求(★承诺覆盖)", "1070-1071", "质保期：验收合格之日起2年", "实质性条款", ("doc", [r"质保期", r"(2|两|二)\s*年"])),
    ("O7", "商务要求(★承诺覆盖)", "1070", "交货地点：招标人指定地点", "实质性条款", ("doc", [r"指定地点"])),
    ("O8", "商务要求(★承诺覆盖)", "前附表26-27", "付款方式（验收后付95%，5%质保金）与履约保证金（中标价5%）", "实质性条款", ("doc", [r"付款|95%", r"履约保证金"])),
]


NPX = "未提供，不计入"
NP_COMMIT = "【未提供：公司最终确认】"
CREDIT_ITEMS = {"Z4-4", "Z4-5"}
PRICE_ITEMS = {"W2", "W7"}
NP_CONFIRM = {34, 40}
PRE: dict = {}  # 被归为「未提供」前按 r9b 口径的状态（仅当它不是「待人工确认」时记录）


def derive(subs, res):
    vals = {x: res.get(x) for x in subs}
    live = {x: v for x, v in vals.items() if v != NPX}
    if not live:
        return NPX, "相关子项均为未提供，不计入"
    sub = list(live.values())
    st = "缺失" if "缺失" in sub else ("待人工确认" if "待人工确认" in sub else "已响应")
    exc = [x for x, v in vals.items() if v == NPX]
    if st == "已响应" and exc:  # 结论依赖未提供的子项：整项归「未提供，不计入」，不当作已响应
        return NPX, f"计入的子项均已响应，但结论依赖未提供的子项 {exc}，整项不计入"
    nt = ("由 " + "、".join(f"{x}={v}" for x, v in live.items() if v != "已响应")) if st != "已响应" else "计入的子项均已响应"
    return st, nt + (f"（未提供不计入的子项：{exc}）" if exc else "")


def load_bid():
    doc = Document(str(OUT / "bid_from_template.docx"))
    paras = []
    head = ""
    for p in doc.paragraphs:
        t = p.text.strip()
        st = (p.style.name if p.style is not None else "") or ""
        if not t:
            continue
        if "\t" in p.text and re.search(r"\t\d+$", p.text):
            continue  # 目录行
        if st.lower().startswith("heading") or re.match(r"^\d+(\.\d+)*[、 ]", t) and len(t) < 30:
            head = t
        paras.append((head, t))
    quote = doc.tables[0]
    qhead = [c.text for c in quote.rows[0].cells]
    other_tables = []
    for t in doc.tables[1:]:
        h = [c.text for c in t.rows[0].cells]
        if h[:2] == ["序号", "招标参数"] or any("结论" in x for x in h):
            continue
        for r in t.rows:
            other_tables.append(("表：" + "/".join(h[:2]), " | ".join(c.text for c in r.cells)))
    return paras + other_tables, qhead


_PT = {}


def pdf_page(pdf, snippet):
    """在去空白的页文本中找片段；跳过封面与目录（前 3 页）。"""
    if not _PT:
        for pno, pg in enumerate(pdf, 1):
            _PT[pno] = re.sub(r"\s+", "", pg.get_text())
    s = re.sub(r"\s+", "", snippet)[:16]
    for pno in sorted(_PT):
        if pno > 3 and s and s in _PT[pno]:
            return pno
    return None


def main():
    units, qhead = load_bid()
    pdf = fitz.open(str(OUT / "bid_from_template.pdf"))
    ov = json.loads((OUT / "verdict_overlay_r10.json").read_text(encoding="utf-8"))
    rows_after = {x["id"]: x for x in ov["changed"]}
    judg = json.loads((ROOT / "work" / "deviation_judgements.json").read_text(encoding="utf-8"))["judgements"]
    res = {}
    out = []
    for cid, cat, src, text, conseq, (mode, arg) in C:
        st, where, ev, note = None, [], [], ""
        if mode in ("doc", "manual_doc"):
            found = []
            for rx in arg:
                # r8：优先取未标「待公司确认属实」、不含【待填】的表述；只有标注/待填的才退而取之
                hit = next(((h, t) for h, t in units if re.search(rx, t) and "待公司确认属实" not in t and "待公司最终确认" not in t and "【待填】" not in t and NP_COMMIT not in t and not re.search(r"：未提供", t)), None) or \
                    next(((h, t) for h, t in units if re.search(rx, t)), None)
                found.append(hit)
                if hit:
                    m = re.search(rx, hit[1])
                    ev.append(hit[1][max(0, m.start() - 20):m.end() + 30])
                    pg = pdf_page(pdf, hit[1][max(0, m.start() - 8):m.end() + 8]) or pdf_page(pdf, hit[1][m.start():m.end() + 12])
                    where.append(f"{hit[0][:24]}（PDF 第{pg}页）" if pg else hit[0][:24])
            n = sum(1 for f in found if f)
            if mode == "doc":
                st = "已响应" if n == len(arg) else ("缺失" if n == 0 else "待人工确认")
                if st == "待人工确认":
                    note = "只找到部分要件：缺 " + "、".join(a for a, f in zip(arg, found) if not f)
                hits = [f[1] for f in found if f]
                if st == "已响应" and not any(NP_COMMIT in h for h in hits) and re.match(r"(Z4|Z5|Z6|H1|H2|H3|H4)", cid) and \
                        any(NP_COMMIT in t and any(re.search(rx, t) for rx in arg) for _h, t in units):
                    PRE[cid] = st
                    st = NPX
                    note = "标书另有未加标注的同义承诺表述，但 3.3/3.4 对应承诺条目的公司最终确认未提供（标「【未提供：公司最终确认】」）"
                if st == "已响应" and any(NP_COMMIT in h for h in hits):
                    st = NPX
                    note = "承诺已按招标原文写入；公司最终确认未提供（标书标「【未提供：公司最终确认】」）" + \
                        ("；对应信用查询截图亦未提供" if cid in CREDIT_ITEMS else "")
                if st == "已响应" and any("待公司最终确认" in h for h in hits):
                    st, note = "待人工确认", "已按招标原文写入，并已对照资料包证据（见 commitment_check_r10.json），标注「待公司最终确认」"
                elif st == "已响应" and any("待公司确认属实" in h for h in hits):
                    st, note = "待人工确认", "已按招标原文写入标书，资料包无直接证据（不矛盾），标注「待公司确认属实」，须公司核实后才算有效承诺"
                elif st == "已响应" and any("【待填】" in h for h in hits):
                    st, note = "待人工确认", "标书已列出招标要求，但投标响应值为【待填】，待 donghao 提供"

            else:
                st = "待人工确认"
                note = ("标书中有相关表述，" if n else "标书中无相关表述，") + "但需盖章/签字/凭证/金额或外部核查，本地无法判定"
            if cid.startswith("O") and st == "缺失" and res.get("H1") in ("已响应", "待人工确认"):
                st, note = "待人工确认", "标书无逐项表述，仅由 3.4「全部商务要求……作出实质性响应」概括承诺覆盖，建议逐项写明"
            if cid == "H1" and st == "待人工确认":
                note += "；3.3 第 17 条与 3.4 均已改为招标原文「完全满足……★条款的所有要求」措辞"

        elif mode == "manual":
            st, note = "待人工确认", "平台操作/身份一致性/成本判断，标书文本无法体现"
        elif mode == "quote_cols":
            miss = [a for a in arg if not any(re.search(a, h) for h in qhead)]
            st = "已响应" if not miss else ("缺失" if len(miss) == len(arg) else "待人工确认")
            where = [f"投标函附录-分项报价表（表头：{'/'.join(qhead)}）"]
            note = ("缺列：" + "、".join(miss)) if miss else ""
            qt = Document(str(OUT / "bid_from_template.docx")).tables[0]
            body = [[c.text.strip() for c in r.cells] for r in qt.rows[1:]]
            col = {h: i for i, h in enumerate(qhead)}
            if not miss:
                if cid == "H5":
                    nb = sum(1 for r in body if r[col["品牌"]] == "【待填】")
                    nm = sum(1 for r in body if r[col["型号"]] == "【待填】")
                    npb = [r[0] for r in body if r[col["品牌"]] == "未提供"]
                    npm = [r[0] for r in body if r[col["型号"]] == "未提供"]
                    cfl = [r[0] for r in body if r[col["品牌"]] == "【待填】" or r[col["型号"]] == "【待填】"]
                    if cfl:
                        st = "待人工确认"
                        note = f"品牌/型号【待填】{len(cfl)} 项（设备 {cfl}，确认函「示例厂家甲」与 V8「示例厂家乙」冲突，不属于未提供）；另品牌未提供 {npb}、型号未提供 {npm}"
                    elif npb or npm:
                        st = NPX
                        note = f"品牌未提供：设备 {npb}；型号未提供：设备 {npm}"
                if cid == "T5":
                    npg = sum(1 for r in body if r[col["证明材料页码范围"]] in ("", "【待填】"))
                    st = "待人工确认"
                    note = (f"偏离情况、证明材料页码范围栏已设（页码 {len(body) - npg}/{len(body)} 项已从 PDF 回填）；"
                            "偏离情况栏已按招标用语（见 T4）；"
                            "招标格式中的「招标文件要求参数 / 投标设备参数」两栏未并入报价表，逐条对照在 6.1 各设备参数表")
            if cid == "T5" and miss:
                note += "；参数对照与偏离结论在 6.1 各设备表中，但报价表本身无对应栏"
        elif mode == "star_rows":
            stars = [j for j in judg if j.get("star")]
            after = {j["id"]: rows_after.get(j["id"], {}).get("after", j.get("verdict") or "无证据待确认") for j in stars}
            neg = [k for k, v in after.items() if v == "负偏离"]
            tbd = [k for k, v in after.items() if v in ("待确认", "无证据待确认")]
            ruled = [k for k in after if rows_after.get(k, {}).get("rule") in ("A", "B")]
            st = "缺失" if neg else ("待人工确认" if tbd or ruled else "已响应")
            note = f"★条款 {len(stars)} 条：负偏离 {len(neg)} 条{neg}；待确认 {len(tbd)} 条{tbd}；经规则A/B改为响应、需人工复核 {len(ruled)} 条{ruled}"
            where = ["6.1 技术参数响应及证明材料（各设备参数表）"]
        elif mode == "star_proof":
            need = [j["id"] for j in judg if j.get("star") and re.search(r"提供.{0,12}(证明|报告|证书|检测)|检测报告|证书|截图", j["req_text"])]
            st = "待人工确认" if need else "已响应"
            note = f"★条款中另要求证明材料的 {len(need)} 条：{need}（需逐条核对附件是否为所要求的证明形式）"
            where = ["6.1 各设备证明材料"]
        elif mode == "device_proof":
            devs = sorted({int(j["device_seq"]) for j in judg})
            none, weak = [], []
            for s in devs:
                kinds = {json.loads(Path(f).read_text(encoding="utf-8"))["kind"] for f in glob.glob(str(ROOT / "work" / "att_corpus" / f"eq{s:02d}_*.json"))}
                if not kinds:
                    none.append(s)
                elif "confirm" not in kinds:
                    weak.append((s, sorted(kinds)))
            st = "缺失" if none else ("待人工确认" if weak else "已响应")
            if not none and weak and {w[0] for w in weak} <= NP_CONFIRM:
                st = NPX
                note += "；设备 34、40 厂家参数确认函未提供（标书附件位置标「未提供」，两台设备 21 行参数不计入一致性判定）"
            note = f"无任何证明材料的设备 {len(none)} 台：{none}；没有参数确认函、只有彩页/其他资料的设备 {len(weak)} 台：{weak}"
            where = ["6.1 附件 6.1.N"]
        elif mode == "verdict_vocab":
            st = "待人工确认"
            vocab = set()
            for t in Document(str(OUT / "bid_from_template.docx")).tables:
                if [c.text.strip() for c in t.rows[0].cells][:3] == ["序号", "招标参数要求", "偏离结论"]:
                    vocab |= {r.cells[2].text.strip() for r in t.rows[1:]}
            bad = vocab - {"正偏离", "负偏离", "无偏离"}
            st = "已响应" if not bad else "待人工确认"
            note = f"6.1 结论栏用语：{sorted(vocab)}" + (f"；不合招标用语：{sorted(bad)}（待确认行：{[x['id'] for x in ov.get('still_disagree', []) if x.get('after') == '待确认']}，V8 写全自动、厂家函写手动，待人工确认）" if bad else "；全部为招标要求的「正偏离/负偏离/无偏离」（r9 待确认行全部定案后改用）")
            where = ["6.1 各设备参数表「结论」栏"]
        elif mode == "lang":
            eng = []
            for f in glob.glob(str(ROOT / "work" / "att_corpus" / "*.json")):
                d = json.loads(Path(f).read_text(encoding="utf-8"))
                txt = " ".join(p if isinstance(p, str) else (p.get("text") or "") for p in d.get("page_texts") or [])
                fn = d.get("file") or Path(d.get("path", f)).name
                ocr = sorted((ROOT / "work" / "r9" / "ocr400").glob(glob.escape(fn) + "_p*.json")) if fn else []
                if ocr:  # r9：文字层字体编码错乱的文件以 400dpi OCR 为准
                    txt = " ".join(r["text"] for f2 in ocr for r in json.loads(f2.read_text(encoding="utf-8"))["full"])
                if len(txt) > 200:
                    cjk = len(re.findall(r"[\u4e00-\u9fff]", txt))
                    if cjk / max(1, len(re.findall(r"\S", txt))) < 0.15:
                        eng.append(d.get("file") or Path(f).stem)
            st = "待人工确认" if eng else "已响应"
            note = f"附件中以外文为主的文件 {len(eng)} 个：{eng}（扫描件无文字层的未计入）"
        elif mode == "derived":
            sub = [res.get(x) for x in arg]
            st = "缺失" if "缺失" in sub else ("待人工确认" if "待人工确认" in sub else "已响应")
            note = "由 " + "、".join(f"{x}={res.get(x)}" for x in arg if res.get(x) != "已响应") if st != "已响应" else "相关子项均已响应"
        if cid in PRICE_ITEMS:
            st, note = NPX, "依赖投标单价/总价：45 项单价、合价及总价未提供（报价表与表后总价行标「未提供」）"
        if mode == "derived":
            st, note = derive(arg, res)
        res[cid] = st
        out.append({"id": cid, "category": cat, "tender_lines": src, "clause": text, "consequence": conseq,
                    "status": st, "bid_location": "；".join(where) or "—", "evidence": ev[:3], "note": note})
    # derived 项依赖后面的 T/O 项：再算一遍
    for o in out:
        spec = next(c for c in C if c[0] == o["id"])
        if spec[5][0] == "derived":
            st, nt = derive(spec[5][1], res)
            o["status"] = res[o["id"]] = st
            o["note"] = nt
    raw = {}
    for o in out:  # r9b 口径（未提供项原本都记「待人工确认」；derived 项按该口径重算）
        raw[o["id"]] = PRE.get(o["id"], "待人工确认") if o["status"] == NPX else o["status"]
    for o in out:
        spec = next(c for c in C if c[0] == o["id"])
        if spec[5][0] == "derived":
            sub = [raw.get(x) for x in spec[5][1]]
            raw[o["id"]] = "缺失" if "缺失" in sub else ("待人工确认" if "待人工确认" in sub else "已响应")
        o["status_r9b_rule"] = raw[o["id"]]
    from collections import Counter
    summ = {"before_未提供不剔除": dict(Counter(o["status_r9b_rule"] for o in out)), "n_before": len(out),
            "after": dict(Counter(o["status"] for o in out if o["status"] != NPX)), "n_after": sum(1 for o in out if o["status"] != NPX),
            "未提供，不计入": [o["id"] for o in out if o["status"] == NPX]}
    (OUT / "disqualification_map_r10.json").write_text(json.dumps({"summary": summ, "n": len(out), "items": out}, ensure_ascii=False, indent=1), encoding="utf-8")
    md = ["| 编号 | 类别 | 招标出处(行) | 条款 | 后果 | 状态 | 标书位置 | 说明 |", "|---|---|---|---|---|---|---|---|"]
    for o in out:
        md.append(f"| {o['id']} | {o['category']} | {o['tender_lines']} | {o['clause']} | {o['consequence']} | **{o['status']}** | {o['bid_location']} | {o['note']} |")
    (OUT / "disqualification_map_r10.md").write_text("\n".join(md) + "\n", encoding="utf-8")
    print(summ)
    for o in out:
        print(o["id"], o["status"], o["bid_location"][:50], "|", o["note"][:160])


if __name__ == "__main__":
    main()
