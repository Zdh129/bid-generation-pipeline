"""r9：400 dpi 重新渲染 + 本地 RapidOCR 局部识别（只在本机运行，不调用任何外部服务）。
对象：各设备确认函/偏离表/承诺函/参数文件中无文字层的页（局部 = 页面上半部 55%，品牌/型号字段所在区域；同时保留整页结果），
以及无文字层的彩页页（整页）。结果写 work/r9/ocr400/*.json（只留在本机）。"""
import json, re, sys, time
from pathlib import Path
from repo_paths import load_paths
import pypdfium2 as pdfium
import numpy as np
from rapidocr_onnxruntime import RapidOCR

ROOT = Path(__file__).resolve().parents[1]
P = load_paths(ROOT)
ATT = Path(P["attachments_dir"])
OUT = ROOT / "work" / "r9" / "ocr400"; OUT.mkdir(parents=True, exist_ok=True)
inv = json.loads((ROOT / "work" / "r9" / "attachments_inventory.json").read_text(encoding="utf-8"))
corpus = {}
for f in (ROOT / "work" / "att_corpus").glob("eq*.json"):
    d = json.loads(f.read_text(encoding="utf-8")); corpus[d["file"]] = d
ocr = RapidOCR()
todo = []
for rel, _ in inv:
    if "设备信息" not in rel: continue
    name = Path(rel).name
    d = corpus.get(name)
    if not d: continue
    for p in d.get("page_texts") or []:
        no_text = ("source" in p) or len((p.get("text") or "").strip()) < 20
        if no_text:
            todo.append((rel, name, d["kind"], p.get("page")))
print("pages", len(todo), flush=True)
t0 = time.time()
for rel, name, kind, page in todo:
    outf = OUT / f"{re.sub(r'[^0-9A-Za-z一-龥#._-]', '_', name)}_p{page}.json"
    if outf.exists(): continue
    src = ATT / rel
    if src.suffix.lower() == ".pdf":
        pdf = pdfium.PdfDocument(str(src)); img = pdf[page - 1].render(scale=400 / 72).to_pil(); pdf.close()
    else:
        from PIL import Image
        img = Image.open(src); img.load()
        img = img.convert("RGB")
    img = img.convert("RGB")
    W, H = img.size
    if max(W, H) > 4800:  # 400dpi A4 约 3300x4700；超大图片按比例缩到 4800
        s = 4800 / max(W, H); img = img.resize((int(W * s), int(H * s))); W, H = img.size
    res = {}
    regions = {"full": (0, 0, W, H)}
    if kind in ("confirm", "commitment", "other"):
        regions["top55"] = (0, 0, W, int(H * 0.55))
    for k, box in regions.items():
        r, _ = ocr(np.array(img.crop(box)))
        res[k] = [{"text": t, "score": round(float(s), 3), "box": [[round(x) for x in pt] for pt in b]} for b, t, s in (r or [])]
    outf.write_text(json.dumps({"file": name, "rel": rel, "page": page, "kind": kind, "dpi": 400, "size": [W, H], **res}, ensure_ascii=False), encoding="utf-8")
    print(name, page, len(res["full"]), f"{time.time()-t0:.0f}s", flush=True)
print("done", time.time() - t0)
