"""r8：从厂家参数确认函（本地语料 work/att_corpus，kind=confirm）提取品牌、型号。不猜：
- 只认确认函正文里「品牌/生产厂家/制造商：」「型号/投标型号/规格型号：」后面的值；
- 来自文字层（PDF 自带文字）的值直接采用；
- 来自本地 OCR 的值：型号须与文件名中的型号一致才采用（OCR 易错），品牌一律不采用（OCR 易把「示例厂家丙」读成「串克」等）；
- 文件名里有型号而与正文提取值不一致 → 有歧义；
- 同一设备多份确认函提取到不同值 → 有歧义；
以上不采用的情况一律填【待填】，并在清单中给出原因和读到的原始值，供人工核对。"""
from __future__ import annotations

import glob
import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
PENDING = "【待填】"
BRAND_RX = re.compile(r"(?:品\s*牌|生产厂家|制造商)\s*[:：]\s*(\S{1,30})")
# 型号：连续的 ASCII 字母数字记号（允许中间空格，如「Intura H2 500」「RPX -150B」），遇到中文即停
# 续接记号须含字母、或以「-」开头、或为 ≥2 位纯数字且后面不是条目序号标点（避免把下一条的「1.」「4」吞进来）
MODEL_RX = re.compile(r"(?:投标型号|规格型号|型\s*号)\s*[:：]\s*([A-Za-z0-9][A-Za-z0-9\-/\.]*?(?:\s+(?:-?[A-Za-z0-9\-/]*[A-Za-z][A-Za-z0-9\-/]*|-[A-Za-z0-9\-/]+|\d{2,}(?![\.、．\d])))*)(?=[\s\u4e00-\u9fff，,；;。]|$)")
FN_MODEL_RX = re.compile(r"[A-Za-z]{1,8}[-A-Za-z]*\d[A-Za-z0-9\-/]*")


def _norm(s: str) -> str:
    return re.sub(r"[\s\-_/（）()]", "", s or "").upper()


def _clean(v: str) -> str:
    v = re.split(r"[，,；;。]", v)[0]
    return re.sub(r"\s+", " ", v).strip(" ;；,，。:：")


_ALL = None


def _all_brand_readings():
    global _ALL
    if _ALL is None:
        _ALL = []
        for f in glob.glob(str(ROOT / "work" / "att_corpus" / "*.json")):
            d = json.loads(Path(f).read_text(encoding="utf-8"))
            for p in d.get("page_texts") or []:
                for m in BRAND_RX.finditer(re.sub(r"\s+", " ", p.get("text") or "")):
                    if "source" not in p:  # 只认文字层读数作为印证
                        _ALL.append((_norm(_clean(m.group(1))), d["file"]))
    return _ALL


def _brand_corroborated(val: str, own_file: str) -> bool:
    n = _norm(val)
    if not n:
        return False
    if n in _norm(own_file):
        return True
    return any(b == n and f != own_file for b, f in _all_brand_readings())


def extract(seq: int) -> dict:
    letters = []
    for f in sorted(glob.glob(str(ROOT / "work" / "att_corpus" / f"eq{seq:02d}_confirm_*.json"))):
        d = json.loads(Path(f).read_text(encoding="utf-8"))
        for p in d.get("page_texts") or []:
            t = re.sub(r"\s+", " ", p.get("text") or "")
            src = "ocr" if "source" in p else "text"
            for m in BRAND_RX.finditer(t):
                letters.append(("brand", _clean(m.group(1)), src, d["file"], p.get("page")))
            for m in MODEL_RX.finditer(t):
                letters.append(("model", _clean(m.group(1)), src, d["file"], p.get("page")))
    files = sorted({Path(f).name for f in glob.glob(str(ROOT / "work" / "att_corpus" / f"eq{seq:02d}_confirm_*.json"))})
    fn_models = sorted({m.group(0).strip("-/") for fn in files for m in FN_MODEL_RX.finditer(fn.split("_", 2)[-1])
                        if not re.fullmatch(r"\d+", m.group(0)) and not re.match(r"(?i)^B?2-\d+", m.group(0))})
    out = {"seq": seq, "confirm_files": files, "filename_model_tokens": fn_models, "raw": letters}
    for key in ("brand", "model"):
        vals = [x for x in letters if x[0] == key]
        reason, value = None, PENDING
        if not files:
            reason = "无厂家参数确认函"
        elif not vals:
            reason = "确认函正文中未出现该字段"
        else:
            distinct = sorted({_norm(v[1]) for v in vals})
            if len(distinct) > 1:
                reason = f"多处取值不一致：{sorted({v[1] for v in vals})}"
            else:
                v = vals[0]
                if key == "brand" and v[2] == "ocr" and not _brand_corroborated(v[1], v[3]):
                    reason = f"仅本地 OCR 读到「{v[1]}」，且无独立来源（文件名或其他文件文字层的相同品牌）印证，不采用"
                elif key == "model" and fn_models and not any(_norm(v[1]) == _norm(t) or _norm(t) in _norm(v[1]) for t in fn_models):
                    reason = f"正文型号「{v[1]}」与文件名型号 {fn_models} 不一致"
                elif key == "model" and v[2] == "ocr" and not fn_models:
                    reason = f"仅本地 OCR 读到「{v[1]}」，文件名无型号可核对，不采用"
                else:
                    value = v[1]
        out[key] = value
        out[key + "_reason"] = reason
        out[key + "_source"] = None if reason else f"{vals[0][3]} 第{vals[0][4]}页（{'文字层' if vals[0][2] == 'text' else '本地OCR，已有独立来源印证'}）"
    return out


def extract_all(seqs) -> dict[int, dict]:
    return {s: extract(s) for s in seqs}


if __name__ == "__main__":
    res = extract_all(range(1, 46))
    for s, r in res.items():
        print(s, r["brand"], "|", r["model"], "|", r["brand_reason"] or "", "|", r["model_reason"] or "")
