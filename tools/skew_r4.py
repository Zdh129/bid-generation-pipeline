"""r4：扫描件倾斜精估 + 文档/彩页照片确定性分类 + 纠偏（生成器与检查器共用同一实现）。

角度约定（本模块返回值）：corr_deg = 使内容水平所需的 PIL.Image.rotate 角（逆时针为正）。
纠偏即 im.rotate(corr_deg)。「倾斜角」skew_deg = -corr_deg（内容相对水平的实际倾角，逆时针为正）。
r3 生成器的 bug：估计得到的就是 corr，却执行了 rotate(-corr)，触发后倾斜反而加倍。
"""
from __future__ import annotations

import numpy as np
from PIL import Image

DESKEW_TRIGGER_DEG = 0.3      # r4：纠偏触发阈值（r3 为 0.5）
COARSE_SEARCH_DEG = 5.0
COARSE_STEP = 0.2
FINE_HALF = 0.3
FINE_STEP = 0.05
MAX_DIM = 1600

# 文档 / 彩页照片 区分规则（确定性，阈值固定，r4 首次设定后未再调整）
CLS_RULE = {
    "text_share_min": 0.40,       # 字符尺寸连通域墨迹 / 全部非线条墨迹（排除表格线、引线）
    "photo_area_max": 0.12,       # 非白连续大块（照片/色块/产品图）面积占比
    "colorful_ratio_max": 0.20,   # 高彩度像素（max-min>60）占比
}
CLS_RULE_TEXT = ("文档类 ⇔ 文字墨迹占比≥0.40 且 照片块面积≤0.12 且 高彩度占比≤0.20；"
                 "否则为彩页/照片类：不纠偏，检查器 Q-SKW-001 对其标不适用并逐张列出依据")


def _to_gray(im: Image.Image) -> np.ndarray:
    return np.asarray(im.convert("L"), dtype=np.uint8)


def _downscale(im: Image.Image, max_dim=MAX_DIM) -> Image.Image:
    w, h = im.size
    s = max(w, h) / max_dim
    if s > 1:
        im = im.resize((max(1, int(w / s)), max(1, int(h / s))), Image.BOX)
    return im


def _otsu(gray: np.ndarray) -> int:
    hist = np.bincount(gray.ravel(), minlength=256).astype(np.float64)
    p = hist / max(1.0, hist.sum())
    omega = np.cumsum(p)
    mu = np.cumsum(p * np.arange(256))
    mu_t = mu[-1]
    with np.errstate(divide="ignore", invalid="ignore"):
        sb = (mu_t * omega - mu) ** 2 / (omega * (1 - omega))
    sb = np.nan_to_num(sb)
    return int(np.argmax(sb))


def _ink_coords(gray: np.ndarray):
    thr = min(200, max(80, _otsu(gray)))
    bw = gray < thr
    ys, xs = np.nonzero(bw)
    n = ys.size
    if n > 400_000:
        step = int(np.ceil(n / 400_000))
        ys, xs = ys[::step], xs[::step]
    return xs.astype(np.float64), ys.astype(np.float64), float(bw.mean())


def _score(xs, ys, corr_deg, h):
    # 对 PIL rotate(corr)（视觉逆时针，y 轴向下）后点的纵坐标：y' = -x sin + y cos
    t = np.deg2rad(corr_deg)
    yp = -xs * np.sin(t) + ys * np.cos(t)
    yi = np.round(yp - yp.min()).astype(np.int64)
    c = np.bincount(yi).astype(np.float64)
    return float((c * c).sum())


def estimate_skew_fine(im: Image.Image | np.ndarray) -> dict:
    """投影剖面法：粗搜 ±5°/0.2°，再在最佳附近 ±0.3° 以 0.05° 细搜。返回 corr_deg（rotate 角）。"""
    if isinstance(im, np.ndarray):
        im = Image.fromarray(im)
    g = _to_gray(_downscale(im))
    h, w = g.shape
    xs, ys, ink = _ink_coords(g)
    if xs.size < 500 or ink < 0.003:
        return {"corr_deg": 0.0, "skew_deg": 0.0, "reliable": False, "why": "墨迹过少", "ink": round(ink, 4), "line_contrast": None}
    xs = xs - w / 2.0
    ys = ys - h / 2.0
    coarse = np.round(np.arange(-COARSE_SEARCH_DEG, COARSE_SEARCH_DEG + 1e-9, COARSE_STEP), 3)
    cs = np.array([_score(xs, ys, a, h) for a in coarse])
    b = float(coarse[int(np.argmax(cs))])
    fine = np.round(np.arange(b - FINE_HALF, b + FINE_HALF + 1e-9, FINE_STEP), 3)
    fs = np.array([_score(xs, ys, a, h) for a in fine])
    best = float(fine[int(np.argmax(fs))])
    contrast = float(fs.max() / max(1.0, np.median(cs)))
    reliable = abs(best) < COARSE_SEARCH_DEG - COARSE_STEP
    corr = round(best, 2) + 0.0
    return {"corr_deg": corr, "skew_deg": -corr + 0.0, "reliable": bool(reliable),
            "why": "" if reliable else "落在搜索边界", "ink": round(ink, 4), "line_contrast": round(contrast, 3)}


def _cc_text_share(im: Image.Image) -> float:
    from scipy import ndimage as ndi
    g = np.asarray(_downscale(im.convert("L"), 1200), dtype=np.uint8)
    H, W = g.shape
    ink = g < min(200, max(80, _otsu(g)))
    lab, n = ndi.label(ink, structure=np.ones((3, 3)))
    if n == 0:
        return 0.0
    areas = np.bincount(lab.ravel())[1:]
    text = tot = 0
    for i, s in enumerate(ndi.find_objects(lab)):
        h = s[0].stop - s[0].start
        w = s[1].stop - s[1].start
        a = int(areas[i])
        if h <= 0.035 * H and w <= 0.08 * W:
            text += a
            tot += a
        elif a / (h * w) < 0.08:
            continue  # 细线（表格线、引线、边框）
        else:
            tot += a
    return text / tot if tot else 0.0


def _photo_area(im: Image.Image) -> float:
    from scipy import ndimage as ndi
    g = np.asarray(_downscale(im.convert("L"), 800), dtype=np.int16)
    bg = np.percentile(g, 90)
    nw = g < bg - 20
    H, W = nw.shape
    bh, bw = H // 16 * 16, W // 16 * 16
    if not bh or not bw:
        return 0.0
    full = nw[:bh, :bw].reshape(bh // 16, 16, bw // 16, 16).mean(axis=(1, 3)) >= 0.85
    lab, n = ndi.label(full)
    if n == 0:
        return 0.0
    sz = np.bincount(lab.ravel())[1:]
    return float(sz[sz >= 6].sum() / full.size)


def classify_scan(im: Image.Image, est: dict | None = None) -> dict:
    """确定性区分：文档类扫描件 vs 彩页/照片类。"""
    small = _downscale(im.convert("RGB"), 800)
    a = np.asarray(small, dtype=np.int16)
    colorful = float(((a.max(axis=2) - a.min(axis=2)) > 60).mean())
    ts = _cc_text_share(im)
    pa = _photo_area(im)
    r = CLS_RULE
    is_doc = ts >= r["text_share_min"] and pa <= r["photo_area_max"] and colorful <= r["colorful_ratio_max"]
    reasons = []
    if ts < r["text_share_min"]:
        reasons.append(f"文字墨迹占比{ts:.2f}<{r['text_share_min']}")
    if pa > r["photo_area_max"]:
        reasons.append(f"照片块面积{pa:.3f}>{r['photo_area_max']}")
    if colorful > r["colorful_ratio_max"]:
        reasons.append(f"高彩度占比{colorful:.3f}>{r['colorful_ratio_max']}")
    return {"kind": "document" if is_doc else "photo_brochure",
            "text_share": round(ts, 3), "photo_area": round(pa, 3), "colorful_ratio": round(colorful, 3),
            "basis": "满足文档类全部条件" if is_doc else "；".join(reasons)}


def deskew(im: Image.Image) -> tuple[Image.Image, dict]:
    """仅文档类且估计可靠、|θ|>0.3° 时纠偏；纠偏后复测残余角。"""
    est = estimate_skew_fine(im)
    cls = classify_scan(im, est)
    rec = {"before_skew_deg": est["skew_deg"], "reliable": est["reliable"], "class": cls, "applied": False}
    if cls["kind"] == "document" and est["reliable"] and abs(est["corr_deg"]) > DESKEW_TRIGGER_DEG:
        fill = (255, 255, 255) if im.mode == "RGB" else 255
        im = im.rotate(est["corr_deg"], resample=Image.BICUBIC, expand=True, fillcolor=fill)
        rec["applied"] = True
        rec["rotate_deg"] = est["corr_deg"]
        after = estimate_skew_fine(im)
        rec["after_skew_deg"] = after["skew_deg"]
    else:
        rec["after_skew_deg"] = est["skew_deg"]
        if cls["kind"] != "document":
            rec["skip_reason"] = "彩页/照片类，不纠偏"
        elif not est["reliable"]:
            rec["skip_reason"] = "估计不可靠"
        else:
            rec["skip_reason"] = f"|θ|≤{DESKEW_TRIGGER_DEG}°"
    return im, rec


def iter_pdf_scan_images(pdf_path, min_w_cm: float = 6.0):
    """逐张取 PDF 中放置宽度 ≥6 cm 的图像对象的原始像素（pypdfium2，不经整页渲染降采样）。"""
    import pypdfium2 as pdfium
    import pypdfium2.raw as pdfium_c
    pdf = pdfium.PdfDocument(str(pdf_path))
    try:
        for i in range(len(pdf)):
            page = pdf[i]
            for obj in page.get_objects(filter=[pdfium_c.FPDF_PAGEOBJ_IMAGE], max_depth=3):
                l, b, r, t = obj.get_bounds()
                if (r - l) * 2.54 / 72 < min_w_cm:
                    continue
                try:
                    im = obj.get_bitmap(render=False).to_pil()
                except Exception:
                    try:
                        im = obj.get_bitmap(render=True).to_pil()
                    except Exception:
                        continue
                yield i + 1, (round(l, 1), round(b, 1), round(r, 1), round(t, 1)), im.convert("RGB")
    finally:
        pdf.close()
