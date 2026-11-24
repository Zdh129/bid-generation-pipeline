"""r9：汇总所有新填值的来源（文件、页、行），写 output/generated-r9/sources_r9.json。"""
import json, re, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
import brand_model_r9 as BM
import bid_r9_fixes as R9
ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "output" / "generated-r9"
bm = json.loads((ROOT / "work" / "r9" / "brand_model_r9.json").read_text(encoding="utf-8"))


def line_of(seq, cat, file, page, value):
    nv = BM.norm(value)
    for n, p, h, t in BM.pools(int(seq)).get(cat, []):
        if n == file and p == page:
            for i, l in enumerate(t.split("\n"), 1):
                if nv and nv in BM.norm(l):
                    return {"line_no": i, "line": l.strip()[:120], "reading": h}
            # 跨行时给出所在页文本中第一处命中的上下文
            nt = t.replace("\n", " ")
            return {"line_no": None, "line": nt[:120], "reading": h}
    return None


def enrich(seq, field):
    r = bm[str(seq)]
    val = r[field]
    if val == BM.PENDING:
        return {"value": val, "reason": r[field + "_reason"],
                "candidates": [{"value": c["value"], "sources": sorted(c["sources"])} for c in r[field + "_candidates"]]}
    cand = next(c for c in r[field + "_candidates"] if c["value"] == val)
    src = []
    for k, s in cand["sources"].items():
        e = {"type": {"L": "厂家确认函/偏离表", "F": "附件文件名", "V": "V8 分项报价表", "B": "彩页/产品资料", "C": "厂家承诺函"}[k], **s}
        if k in ("L", "B", "C"):
            e.update(line_of(seq, k, s["file"], s["page"], val) or {})
        src.append(e)
    return {"value": val, "n_sources": len(src), "sources": src, "note": cand.get("note")}


res = {
    "principle": "每个新填值至少有一个可复核出处；品牌/型号须 ≥2 个独立来源一致；无可靠来源者留【待填】",
    "tender_values": {k: {**v, "file": R9.TENDER_PDF, "line_ref": "work/r7/tender_main_raw.txt 行 " + v["lines"]} for k, v in R9.FILLED.items()},
    "brand_model": {s: {"brand": enrich(s, "brand"), "model": enrich(s, "model")} for s in bm},
    "verdict_changes": {k: {"verdict": v, "evidence": e} for k, (v, e) in R9.R9_CHANGES.items()},
    "commitments": json.loads((OUT / "commitment_check_r9.json").read_text(encoding="utf-8")),
    "unit_prices": json.loads((ROOT / "work" / "r9" / "price_search.json").read_text(encoding="utf-8")),
}
(OUT / "sources_r9.json").write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
print("brand filled", sum(1 for s in bm.values() if s["brand"] != BM.PENDING), "model filled", sum(1 for s in bm.values() if s["model"] != BM.PENDING))
