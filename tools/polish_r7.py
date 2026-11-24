"""r7 润色（纯本地规则/词表，不调用任何模型）。
lint_document(doc)  → 问题清单（只给建议，不改字）
apply_safe_fixes(doc) → 只做格式级安全修改并逐处记录：
  F1 中文与数字/拉丁字母之间的半角空格删除（「不少于 90 个」→「不少于90个」）
  F2 中文语境内的半角标点改全角（, ; : ? ! 前后都是中文时）
  F3 连续重复的中文标点合并（「。。」→「。」）
安全边界：不改表格（参数表文本来自招标文件、结论列受 461 行断言约束）、不改标题与目录、不改图题、
不改含【待填】/日期/金额/编号的 run 中的数字本身；只在单个 run 内替换，不跨 run，不改变任何承诺语义。
"""
from __future__ import annotations

import re
from collections import Counter

CJK = "\u4e00-\u9fff"
CJKP = "\u4e00-\u9fff，。、；：！？（）《》「」“”"
class _F1:
    """中文与数字/字母之间的单个半角空格。例外（保留空格）：章节号后的空格（「6.1 技术」「3.2 投标人」），
    以及空格前是英文单词/型号（避免把型号与中文粘连后难以辨认）。"""
    rx = re.compile(rf"(?<=[{CJKP}]) (?=[0-9])|(?<=[0-9%]) (?=[{CJKP}])")

    def _ok(self, s, m):
        left, right = s[:m.start()], s[m.end():]
        if re.search(r"\d+(\.\d+)+$", left) and not right[:1] in "节条款页章":
            return False  # 章节号后的空格保留（「6.1 技术参数」「附件6.1.1 示例设备01」）
        return True

    def finditer(self, s):
        return [m for m in self.rx.finditer(s) if self._ok(s, m)]

    def findall(self, s):
        return self.finditer(s)

    def sub(self, fn, s):
        out, last = [], 0
        for m in self.finditer(s):
            out.append(s[last:m.start()]); last = m.end()
        out.append(s[last:])
        return "".join(out)


F_RULES = [
    ("F1", _F1(), lambda m: ""),
    ("F2", re.compile(rf"(?<=[{CJK}])([,;:?!])(?=[{CJK}])"), lambda m: {",": "，", ";": "；", ":": "：", "?": "？", "!": "！"}[m.group(1)]),
    ("F3", re.compile(r"([，。、；：！？])\1+"), lambda m: m.group(1)),
]
TYPO = {"登陆": "登录", "帐号": "账号", "帐户": "账户", "既使": "即使", "按装": "安装", "调式": "调试", "部份": "部分",
        "做为": "作为", "交纳": "缴纳", "较准": "校准", "检订": "检定", "维户": "维护", "形号": "型号", "规隔": "规格",
        "保正": "保证", "质保其": "质保期", "截止至": "截至", "附和要求": "符合要求", "符和": "符合", "须知前付表": "须知前附表",
        "项且": "项目", "签定": "签订", "定货": "订货", "坚定校准": "检定校准", "既日": "即日"}
REDUP_OK = set("常常 刚刚 渐渐 仅仅 稍稍 往往 偏偏 默默 慢慢 看看 试试 层层 处处 时时 人人 各个 个个 次次 天天 年年 一一 谢谢 妈妈 爸爸 哈哈 匆匆 纷纷 轻轻 久久 明明 好好 多多".split())
TERMS = [  # (我方用语正则, 招标文件用语, 说明)
    (r"供应商", "投标人", "招标文件统一称「投标人」"),
    (r"采购人|甲方", "招标人", "招标文件同义词条款：招标阶段按「招标人」理解"),
    (r"投标方(?!案)", "投标人", "招标文件统一称「投标人」"),
    (r"标书", "投标文件", "招标文件统一称「投标文件」"),
    (r"保修期|保证期", "质保期", "招标文件用「质保期」"),
    (r"技术规格中的★条款", "第二节技术要求的★条款", "招标文件原文为「第三章招标需求第一节商务要求及第二节技术要求的★条款」"),
    (r"致：招标人 / 招标代理机构", "致：示例招标人有限公司（招标人）", "招标人名称见前附表第 1 条；代理机构为示例招标代理有限公司"),
]
LONG = 100


def _skip_para(p) -> str | None:
    st = (p.style.name if p.style is not None else "") or ""
    t = p.text.strip()
    if st.lower().startswith(("heading", "toc")) or "标题" in st or st.lower().startswith("title"):
        return "标题/目录"
    if t.startswith("图 ") and re.match(r"图 \d+ ", t):
        return "图题"
    if "\t" in p.text and re.search(r"\t\d+$", p.text):
        return "目录行"
    return None


def iter_paragraphs(doc):
    """正文段落（不含表格），带全局序号。"""
    for i, p in enumerate(doc.paragraphs):
        yield i, p


def iter_table_cells(doc):
    for ti, t in enumerate(doc.tables, 1):
        head = [c.text for c in t.rows[0].cells]
        kind = "参数表" if head[:2] == ["序号", "招标参数"] or "偏离结论" in "".join(head) or "结论" in "".join(head) else "其他表"
        for ri, r in enumerate(t.rows):
            seen = set()
            for ci, c in enumerate(r.cells):
                if id(c._tc) in seen:
                    continue
                seen.add(id(c._tc))
                yield ti, ri, ci, kind, c.text


def lint_text(text: str, loc: dict, out: list):
    def add(cat, m, sug, note=""):
        s, e = m.start(), m.end()
        out.append({**loc, "cat": cat, "hit": m.group(0), "suggest": sug, "note": note,
                    "context": text[max(0, s - 18):e + 18]})
    for k, v in TYPO.items():
        for m in re.finditer(re.escape(k), text):
            add("错别字", m, v)
    for m in re.finditer(rf"([{CJK}]{{1,4}})\1", text):
        w = m.group(0)
        if len(m.group(1)) == 1 and (w in REDUP_OK or m.group(1) in "的了是和在与及或"):
            if m.group(1) in "的了是和在与及或":
                add("重复词", m, m.group(1))
            continue
        if w in REDUP_OK:
            continue
        if len(m.group(1)) >= 2:
            add("重复词", m, m.group(1))
    for pat, tender, note in TERMS:
        for m in re.finditer(pat, text):
            add("术语不一致", m, tender, note)
    for code, rx, fn in F_RULES:
        for m in rx.finditer(text):
            add({"F1": "全半角/空格", "F2": "全半角/标点", "F3": "重复标点"}[code], m, fn(m), code)
    for m in re.finditer(r"[０-９Ａ-Ｚａ-ｚ]", text):
        add("全半角/全角字母数字", m, chr(ord(m.group(0)) - 0xFEE0))
    for m in re.finditer(r"\(([^()]*[" + CJK + r"][^()]*)\)", text):
        add("全半角/括号", m, f"（{m.group(1)}）")
    pos = 0
    for seg in re.split(r"(?<=[。！？；])", text):
        n = len(re.sub(r"\s", "", seg))
        if n > LONG:
            out.append({**loc, "cat": "句式过长", "hit": f"{n} 字", "suggest": "拆分为 2–3 句", "note": f">{LONG} 字",
                        "context": seg[:40] + "…"})
        pos += len(seg)


def lint_document(doc) -> list[dict]:
    out = []
    for i, p in iter_paragraphs(doc):
        if not p.text.strip():
            continue
        sk = _skip_para(p)
        lint_text(p.text, {"where": "段落" if not sk else f"段落·{sk}", "para": i, "skip": sk}, out)
    for ti, ri, ci, kind, text in iter_table_cells(doc):
        if text.strip():
            w = kind if kind != "参数表" else {1: "参数表·招标原文", 3: "参数表·证据引文"}.get(ci, "参数表·其他列")
            lint_text(text, {"where": w, "table": ti, "row": ri, "col": ci}, out)
    return out


def apply_safe_fixes(doc) -> list[dict]:
    changes = []
    for i, p in iter_paragraphs(doc):
        if not p.text.strip() or _skip_para(p):
            continue
        for ri, r in enumerate(p.runs):
            t0 = r.text
            if not t0 or "【待填】" in t0:
                continue
            t = t0
            applied = []
            for code, rx, fn in F_RULES:
                t2 = rx.sub(fn, t)
                if t2 != t:
                    applied.append(f"{code}×{len(rx.findall(t)) if code != 'F1' else sum(1 for _ in rx.finditer(t))}")
                    t = t2
            if t != t0:
                r.text = t
                changes.append({"para": i, "run": ri, "rules": applied, "before": t0[:160], "after": t[:160]})
    return changes


def summarize(issues: list[dict], changes: list[dict]) -> dict:
    by = Counter((x["cat"], x["where"]) for x in issues)
    return {"issues_total": len(issues),
            "by_category": dict(Counter(x["cat"] for x in issues)),
            "by_category_where": {f"{a}|{b}": n for (a, b), n in sorted(by.items())},
            "safe_fixes_runs": len(changes),
            "safe_fixes_by_rule": dict(Counter(r.split("×")[0] for c in changes for r in c["rules"]))}
