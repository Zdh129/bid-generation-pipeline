"""按文档顺序导出 DOCX 内嵌图（含其后题注），计算 r4 倾斜/分类特征。仅本地。"""
import sys, io, json
sys.path.insert(0, __file__.rsplit('/',1)[0])
from docx import Document
from docx.oxml.ns import qn
from PIL import Image
from skew_r4 import estimate_skew_fine, classify_scan

def iter_pics(docx_path):
    doc = Document(docx_path)
    body = doc.element.body
    paras = [p for p in body.iter(qn('w:p'))]
    out = []
    for i, p in enumerate(paras):
        blips = p.findall('.//' + qn('a:blip'))
        for b in blips:
            rid = b.get(qn('r:embed'))
            part = doc.part.related_parts[rid]
            cap = ''
            for q in paras[i+1:i+3]:
                t = ''.join(x.text or '' for x in q.iter(qn('w:t')))
                if t.startswith('图'):
                    cap = t; break
            out.append((cap, part.blob))
    return out

if __name__ == '__main__':
    res = []
    for k, (cap, blob) in enumerate(iter_pics(sys.argv[1]), 1):
        im = Image.open(io.BytesIO(blob)); im.load()
        e = estimate_skew_fine(im); c = classify_scan(im, e)
        res.append({"k": k, "caption": cap, "skew_deg": e["skew_deg"], "reliable": e["reliable"], **c})
    json.dump(res, open(sys.argv[2], 'w'), ensure_ascii=False, indent=0)
    print(len(res))
