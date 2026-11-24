"""r9：对确认函中「品牌/型号」所在行做 600 dpi 局部重识别（本机 RapidOCR）。
行位置取自 400 dpi 全页识别结果；裁出整行（左右放宽到整页宽、上下各放宽半行高）后再识别。结果 work/r9/ocr_fields/*.json。"""
import json, re
from pathlib import Path
from repo_paths import load_paths
import numpy as np, pypdfium2 as pdfium
from rapidocr_onnxruntime import RapidOCR
ROOT = Path(__file__).resolve().parents[1]
P = load_paths(ROOT)
ATT = Path(P["attachments_dir"])
OUT = ROOT / "work" / "r9" / "ocr_fields"; OUT.mkdir(parents=True, exist_ok=True)
ocr = RapidOCR()
for f in sorted((ROOT / "work" / "r9" / "ocr400").glob("*.json")):
    js = json.loads(f.read_text(encoding="utf-8"))
    if js["kind"] not in ("confirm",) or not js["rel"].lower().endswith(".pdf"):
        continue
    hits = [r for r in js["full"] if re.search(r"型\s*号|品\s*牌|制造商|生产厂家|名称", r["text"])]
    if not hits:
        continue
    pdf = pdfium.PdfDocument(str(ATT / js["rel"]))
    img = pdf[js["page"] - 1].render(scale=600 / 72).to_pil().convert("RGB"); pdf.close()
    k = 600 / 400
    res = []
    for r in hits:
        ys = [pt[1] for pt in r["box"]]; h = max(ys) - min(ys)
        y0, y1 = max(0, (min(ys) - h * 0.6) * k), (max(ys) + h * 0.6) * k
        crop = img.crop((0, int(y0), img.size[0], int(y1)))
        out, _ = ocr(np.array(crop))
        line = " ".join(t for _, t, s in sorted(out or [], key=lambda x: x[0][0][0]) if float(s) >= 0.5)
        res.append({"line400": r["text"], "line600": line, "y": [int(y0), int(y1)]})
    (OUT / f.name).write_text(json.dumps({"file": js["file"], "rel": js["rel"], "page": js["page"], "dpi": 600, "lines": res}, ensure_ascii=False, indent=0), encoding="utf-8")
    print(js["file"][:40], js["page"], [x["line600"][:40] for x in res])
