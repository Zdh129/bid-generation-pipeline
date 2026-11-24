"""r9：品牌/型号多来源交叉校验（全部本机处理）。
来源（互相独立的才算两个来源）：
  L  厂家确认函/偏离表（同一份文件的文字层或 400dpi 本地 RapidOCR 读数，算一个来源）
  F  附件文件名
  V  V8 分项报价表（品牌、型号栏）
  B  彩页/产品介绍/参数说明（文字层、原有 OCR 或 400dpi OCR）
  C  同设备的厂家承诺函/其他厂家文件
规则：候选值在 ≥2 个独立来源中一致（规范化后相等或被包含）才填；多个互相矛盾的候选都 ≥2 → 不填，列候选；否则【待填】。"""
from __future__ import annotations
import glob, json, re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PENDING = "【待填】"
OCR_DIR = ROOT / "work" / "r9" / "ocr400"
V8 = json.loads((ROOT / "work" / "r9" / "v8_brand_model.json").read_text(encoding="utf-8"))
INV = json.loads((ROOT / "work" / "r9" / "attachments_inventory.json").read_text(encoding="utf-8"))


def norm(s: str) -> str:
    s = (s or "").upper()
    s = s.translate(str.maketrans("（）：，　／", "():, /"))
    return re.sub(r"[\s\-_()（）·.\"'“”]", "", s)


def ocr_lines(js: dict, region="full") -> str:
    rows = sorted(js.get(region) or [], key=lambda r: (round(r["box"][0][1] / 40), r["box"][0][0]))
    return "\n".join(r["text"] for r in rows if r["score"] >= 0.5)


def device_files(seq: int):
    out = []
    for rel, _ in INV:
        if "设备信息" in rel and re.search(rf"(?:示例|ex|pkg|包2)-0*{seq}#|(?:EX|B2)-0*{seq}#|/2-0*{seq} ", "/" + rel):
            out.append(rel)
    return out


def pools(seq: int) -> dict:
    """返回 {来源类别: [(文件名, 页, 读法, 文本)]}"""
    P = {"L": [], "B": [], "C": []}
    corpus = {}
    for f in glob.glob(str(ROOT / "work" / "att_corpus" / f"eq{seq:02d}_*.json")):
        d = json.loads(Path(f).read_text(encoding="utf-8")); corpus[d["file"]] = d
    for name, d in corpus.items():
        kind = d["kind"]
        cat = "L" if kind == "confirm" else ("B" if kind in ("brochure", "other") else "C")
        for p in d.get("page_texts") or []:
            t = p.get("text") or ""
            if "source" not in p and len(t.strip()) >= 20:
                P[cat].append((name, p.get("page"), "文字层", t))
            o = OCR_DIR / f"{re.sub(r'[^0-9A-Za-z一-龥#._-]', '_', name)}_p{p.get('page')}.json"
            if o.exists():
                js = json.loads(o.read_text(encoding="utf-8"))
                P[cat].append((name, p.get("page"), "OCR400", ocr_lines(js)))
            o6 = ROOT / "work" / "r9" / "ocr_fields" / o.name
            if o6.exists():
                js6 = json.loads(o6.read_text(encoding="utf-8"))
                P[cat].append((name, p.get("page"), "OCR600局部", "\n".join(x["line600"] for x in js6["lines"])))
    return P


BRAND_RX = re.compile(r"(?:品\s*牌|生产厂家|制造商|生产商)\s*[:：]?\s*([^\n:：]{1,30})")
MODEL_RX = re.compile(r"(?:投标型号|规格型号|产品型号|设备型号|型\s*号)\s*[:：]?\s*([A-Za-z0-9][A-Za-z0-9\-/\. ]{1,30})")
SEAL_RX = re.compile(r"([一-龥（）()A-Za-z]{2,30}(?:有限公司|有限责任公司|股份有限公司))")


def field_values(texts, rx):
    vals = []
    for name, page, how, t in texts:
        for line in t.split("\n"):
            for m in rx.finditer(line):
                v = m.group(1).strip(" ;；,，。")
                v = re.split(r"\s{2,}|[，,；;。]", v)[0].strip()
                if v and not re.fullmatch(r"[\d\s.]+", v):
                    vals.append((v, name, page, how))
    return vals


if __name__ == "__main__":
    for seq in range(1, 46):
        P = pools(seq)
        print(f"\n### {seq} V8={V8.get(str(seq),{}).get('pairs')}")
        print("  files:", [Path(r).name[:50] for r in device_files(seq)])
        print("  L brand:", [(v, h) for v, n, p, h in field_values(P["L"], BRAND_RX)][:6])
        print("  L model:", [(v, h) for v, n, p, h in field_values(P["L"], MODEL_RX)][:6])
        print("  L seal :", sorted({m for _, _, h, t in P["L"] for m in SEAL_RX.findall(t)})[:4])
        print("  B model:", [(v, h) for v, n, p, h in field_values(P["B"], MODEL_RX)][:4])


def _fn_model_tokens(files):
    toks = set()
    for rel in files:
        name = Path(rel).stem
        for m in re.finditer(r"[A-Za-z][A-Za-z0-9]*[-/]?[A-Za-z0-9\-/]*\d[A-Za-z0-9\-/]*", name):
            t = m.group(0).strip("-/")
            if not re.match(r"(?i)^B?2-\d+$", t) and len(t) >= 3:
                toks.add(t)
    return toks


def _brand_tokens(b: str):
    parts = re.split(r"[（）()\s/]+", b or "")
    return [p for p in parts if len(p) >= 2]


def cross_check(seq: int) -> dict:
    P = pools(seq)
    files = device_files(seq)
    fn = norm(" ".join(Path(r).name for r in files))
    txt = {k: [(n, p, h, norm(t)) for n, p, h, t in v] for k, v in P.items()}
    v8 = V8.get(str(seq), {})
    v8_model, v8_brand = (v8.get("pairs") or [[None, None]])[0]

    def support(c: str):
        nc = norm(c)
        s = {}
        if not nc or len(nc) < 2:
            return s
        for k in ("L", "B", "C"):
            hit = next(((n, p, h) for n, p, h, t in txt[k] if nc in t), None)
            if hit:
                s[k] = {"file": hit[0], "page": hit[1], "how": hit[2]}
        if nc in fn:
            s["F"] = {"file": next(Path(r).name for r in files if nc in norm(Path(r).name))}
        if v8_model and nc == norm(v8_model) or v8_brand and nc in norm(v8_brand):
            s["V"] = {"file": "V8 分项报价表（表 0）", "value": v8_model if nc == norm(v8_model or "") else v8_brand}
        return s

    # ---- 型号
    mc = {}
    for v, n, p, h in field_values(P["L"], MODEL_RX):
        mc.setdefault(norm(v), v)
    for t in _fn_model_tokens(files):
        mc.setdefault(norm(t), t)
    if v8_model:
        mc[norm(v8_model)] = v8_model  # V8 写法覆盖（与 OCR/文件名规范化相等时用 V8 的规范写法）
    mres = []
    for nv, v in mc.items():
        s = support(v)
        mres.append({"value": v, "sources": s, "n": len(s)})
    mres.sort(key=lambda r: -r["n"])
    ok = sorted([r for r in mres if r["n"] >= 2], key=lambda r: -len(norm(r["value"])))
    if ok:  # 被最长候选包含的短候选（如「Vortex3」⊂「Vortex 3 S025」）视为同一值
        ok = [r for r in ok if r is ok[0] or norm(r["value"]) not in norm(ok[0]["value"])]
    distinct = {norm(r["value"]) for r in ok}
    # 被包含的（如 MODEL-D 与 MODEL-D-xx）视为同一值，取最长
    model, mreason = PENDING, None
    if len(distinct) == 1:
        model = ok[0]["value"]
    elif len(distinct) > 1:
        mreason = "多个候选各有 ≥2 来源，互相矛盾：" + "、".join(r["value"] for r in ok)
    else:
        mreason = "没有任何候选值得到 ≥2 个独立来源"
    # ---- 品牌（按 V8 品牌的各组成词、确认函品牌栏、确认函落款公司分别核对）
    bc = []
    if v8_brand:
        bc.append(v8_brand)
        bc += _brand_tokens(v8_brand)
    for v, n, p, h in field_values(P["L"], BRAND_RX):
        if len(v) >= 2 and not re.search(r"确认函|名称|产地|或生|（或", v):
            bc.append(v)
    bres = []
    seen = set()
    for v in bc:
        if norm(v) in seen:
            continue
        seen.add(norm(v))
        s = support(v)
        bres.append({"value": v, "sources": s, "n": len(s)})
    bres.sort(key=lambda r: (-r["n"], -len(r["value"])))
    # V8 品牌整体未直接命中、但其每个组成词都各有 ≥2 来源时，按 V8 写法计为得到支持（来源取各组成词来源之交集外的并集）
    if v8_brand:
        toks = _brand_tokens(v8_brand)
        full = next((r for r in bres if r["value"] == v8_brand), None)
        if full and full["n"] < 2 and len(toks) >= 2:
            ts = [next((r for r in bres if r["value"] == t), {"n": 0}) for t in toks]
            if all(t["n"] >= 2 for t in ts):
                full["sources"] = {k: v for t in ts for k, v in t["sources"].items()}
                full["n"] = len(full["sources"]); full["note"] = "各组成词分别得到 ≥2 来源"
    okb = sorted([r for r in bres if r["n"] >= 2], key=lambda r: -len(norm(r["value"])))
    brand, breason = PENDING, None
    if okb:
        # 取得到 ≥2 来源的最长写法；其余 ≥2 来源候选须被它包含，否则视为矛盾
        best = okb[0]
        clash = [r for r in okb if norm(r["value"]) not in norm(best["value"]) and norm(best["value"]) not in norm(r["value"])]
        if clash:
            breason = "品牌候选矛盾：" + "、".join(r["value"] for r in okb)
        else:
            brand = best["value"]
    else:
        breason = "没有任何候选值得到 ≥2 个独立来源"
    return {"seq": seq, "brand": brand, "brand_reason": breason, "brand_candidates": bres,
            "model": model, "model_reason": mreason, "model_candidates": mres, "files": [Path(r).name for r in files],
            "v8": {"brand": v8_brand, "model": v8_model}}


def cross_check_all():
    return {s: cross_check(s) for s in range(1, 46)}
