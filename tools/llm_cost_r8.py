"""r8：下一轮「大模型语义复核」估费（只在本机统计字符数，不调用任何模型/API）。
复核对象：r7 规则 A/B 改为「响应」的 45 行（其中★ 5 行）。每行拟发送内容：
  ① 固定指令（系统提示，全批共用但每次请求都要带）；② 设备名 + 招标参数原文；③ 原判定理由；
  ④ 证据：文件名 + 页码 + 片段（方案 S，最小）或 证据所在页的全文（方案 F，含上下文）。
token 估算：中文/全角字符按 1 字≈1 token（偏保守；常见中文分词器约 0.6–1.0 token/字），
其他字符按 4 字符≈1 token；输出按每行 150 token（结论 + 一句理由）。"""
import glob, json, re, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output" / "generated-r8"
J = {j["id"]: j for j in json.loads((ROOT / "work" / "deviation_judgements.json").read_text(encoding="utf-8"))["judgements"]}
ov = json.loads((OUT / "verdict_overlay_r8.json").read_text(encoding="utf-8"))
rows = [x for x in ov["changed"] if x["rule"] in ("A", "B")]
DEV = {d["seq"]: d["name"] for d in json.loads((ROOT / "work" / "devices.json").read_text(encoding="utf-8"))}
SYS = ("你是招标技术参数复核员。给定招标参数原文、投标方证据片段，判断证据是否满足该参数：只能回答「满足/不满足/无法判断」，"
       "并用一句话说明依据（引用证据原文）。严格不等号、数值范围、单位换算须逐项核对；证据未提及的视为无法判断。输出 JSON：{verdict, reason}。")
PAGE = {}
for f in glob.glob(str(ROOT / "work" / "att_corpus" / "*.json")):
    d = json.loads(Path(f).read_text(encoding="utf-8"))
    for p in d.get("page_texts") or []:
        PAGE[(d["file"], p.get("page"))] = p.get("text") or ""

def tok(s):
    cjk = len(re.findall(r"[\u3000-\u303f\u4e00-\u9fff\uff00-\uffef]", s))
    return cjk + (len(s) - cjk) / 4

PII = {"手机号": r"(?<!\d)1[3-9]\d{9}(?!\d)", "身份证号": r"(?<![\dA-Za-z])\d{17}[\dXx](?![\dA-Za-z])", "固话": r"0\d{2,3}-\d{7,8}",
       "邮箱": r"[\w.+-]+@[\w-]+\.[\w.]+", "人名字段": r"(联系人|法定代表人|授权代表|经办人|负责人)\s*[:：]\s*\S{2,4}", "统一社会信用代码": r"(?<![0-9A-Z])[0-9A-HJ-NPQRTUWXY]{2}\d{6}[0-9A-HJ-NPQRTUWXY]{10}(?![0-9A-Z])"}
res = {"rows": [], "sys_chars": len(SYS), "sys_tokens": tok(SYS)}
for x in rows:
    j = J[x["id"]]
    head = f"设备：{DEV.get(int(j['device_seq']), '')}\n参数{'（★）' if j.get('star') else ''}：{j['req_text']}\n原判定理由：{j.get('reason','')}\n"
    evs = j.get("evidence") or []
    s_txt = head + "".join(f"证据：{e.get('file')} 第{e.get('page')}页：{e.get('snippet','')}\n" for e in evs)
    if not evs and x["rule"] == "B":
        s_txt += "证据：该设备厂家参数确认函整体声明满足招标参数（无逐条片段）\n"
    f_pages = {(e.get("file"), e.get("page")) for e in evs}
    if x["rule"] == "B":
        seq = int(j["device_seq"])
        for f in glob.glob(str(ROOT / "work" / "att_corpus" / f"eq{seq:02d}_confirm_*.json")):
            d = json.loads(Path(f).read_text(encoding="utf-8"))
            f_pages |= {(d["file"], p.get("page")) for p in d.get("page_texts") or []}
    f_txt = head + "".join(f"证据页全文：{fn} 第{pg}页：{PAGE.get((fn, pg), '')}\n" for fn, pg in sorted(f_pages, key=str))
    pii = {k: len(re.findall(v, f_txt)) for k, v in PII.items()}
    res["rows"].append({"id": x["id"], "rule": x["rule"], "star": bool(j.get("star")), "S_chars": len(s_txt), "S_tokens": round(tok(s_txt)),
                        "F_chars": len(f_txt), "F_tokens": round(tok(f_txt)), "pii_hits_F": {k: v for k, v in pii.items() if v},
                        "pii_hits_S": {k: len(re.findall(v, s_txt)) for k, v in PII.items() if re.search(v, s_txt)}})
OUT_TOK = 150
def agg(sel):
    n = len(sel)
    sin = sum(r["S_tokens"] for r in sel) + n * res["sys_tokens"]
    fin = sum(r["F_tokens"] for r in sel) + n * res["sys_tokens"]
    return {"n": n, "S_chars": sum(r["S_chars"] for r in sel), "F_chars": sum(r["F_chars"] for r in sel),
            "S_in_tokens": round(sin), "F_in_tokens": round(fin), "out_tokens": n * OUT_TOK}
res["star5"] = agg([r for r in res["rows"] if r["star"]])
res["all45"] = agg(res["rows"])
# 价格为假设档位（未联网查价），仅示意量级：元/百万 token（输入, 输出）
TIERS = {"经济档(假设 2/8)": (2, 8), "中档(假设 4/16)": (4, 16), "高档(假设 20/80)": (20, 80)}
for k in ("star5", "all45"):
    a = res[k]
    a["cost_yuan"] = {t: {"S": round((a["S_in_tokens"] * i + a["out_tokens"] * o) / 1e6, 4), "F": round((a["F_in_tokens"] * i + a["out_tokens"] * o) / 1e6, 4)} for t, (i, o) in TIERS.items()}
(OUT / "llm_cost_estimate_r8.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
print(json.dumps({k: res[k] for k in ("sys_chars", "sys_tokens", "star5", "all45")}, ensure_ascii=False, indent=1))
print("PII rows F:", [(r["id"], r["pii_hits_F"]) for r in res["rows"] if r["pii_hits_F"]])
print("PII rows S:", [(r["id"], r["pii_hits_S"]) for r in res["rows"] if r["pii_hits_S"]])
print("star ids", [r["id"] for r in res["rows"] if r["star"]])
