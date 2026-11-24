"""r6：单页内容级归类（在 r5 按文件归类之上，确定性覆盖规则）+ 图片压缩参数。

覆盖规则（只在「页内容分类」与「文件类别」冲突时才检查，其余页一律沿用文件类别）：
- 规则 P1（确认函 → 彩页）：文件类别=参数确认函，且 页序≥2，且 skew_r5 内容分类=彩页/照片类，
  且 本地 OCR 文本不含任何函件用语（确认/偏离/响应/承诺/声明/特此/盖章/公章/招标/投标/致），
  且 页内无红色印章（红色像素占比 < 0.2%）。四条全部满足才改为「产品彩页/说明书」。
  理由：确认函首页、带印章或带函件用语的页属于厂家签章声明的一部分，必须留在确认函名下；
  只有不含声明内容、未盖章的纯产品图片页才按内容归为彩页。图片位置不移动（仍紧随原文件），只改图题类别。
- 规则 P2（彩页/其他 → 确认函）：内容分类=文档类，且 OCR 文本含确认函标题用语
  （参数确认函/技术参数确认/技术偏离表/偏离表）时改为「参数确认函」。
OCR 只用本机 RapidOCR；日志只记关键词命中与字数，不记原文。
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import numpy as np
from PIL import Image

LETTER_KW = re.compile(r"确认|偏离|响应|承诺|声明|特此|盖章|公章|招标|投标|致[:：]")
CONFIRM_TITLE_KW = re.compile(r"参数确认函|技术参数确认|技术偏离表|偏离表")
SEAL_RED_MAX = 0.002

_ENG = None


def _ocr_text(img_path: Path) -> str:
    global _ENG
    if _ENG is None:
        from rapidocr_onnxruntime import RapidOCR
        _ENG = RapidOCR()
    im = Image.open(img_path).convert("RGB")
    res, _ = _ENG(np.asarray(im))
    return "".join(r[1] for r in (res or []))


def red_seal_ratio(img_path: Path) -> float:
    a = np.asarray(Image.open(img_path).convert("RGB").resize((600, 850)), dtype=np.int16)
    r, g, b = a[..., 0], a[..., 1], a[..., 2]
    return float(((r > 150) & (g < 110) & (b < 110) & (r - g > 70)).mean())


def page_features(img_path: Path) -> dict:
    """OCR 关键词命中（缓存到同名 .page.json）。"""
    side = Path(img_path).with_suffix(".page.json")
    if side.is_file():
        return json.loads(side.read_text(encoding="utf-8"))
    t = _ocr_text(Path(img_path))
    rec = {"ocr_chars": len(t),
           "letter_kw": sorted(set(LETTER_KW.findall(t))),
           "confirm_title_kw": sorted(set(CONFIRM_TITLE_KW.findall(t))),
           "red_seal_ratio": round(red_seal_ratio(Path(img_path)), 4)}
    side.write_text(json.dumps(rec, ensure_ascii=False), encoding="utf-8")
    return rec


def page_kind(file_kind: str, page_no: int, content_kind: str, img_path: Path) -> tuple[str, dict]:
    """返回 (页类别, 依据)。只有冲突页才跑 OCR。"""
    info = {"file_kind": file_kind, "page_no": page_no, "content_kind": content_kind, "rule": None}
    conflict = (file_kind == "confirm" and content_kind != "document") or \
               (file_kind != "confirm" and content_kind == "document")
    if not conflict:
        info["rule"] = "无冲突，沿用文件类别"
        return file_kind, info
    f = page_features(img_path)
    info.update(f)
    if file_kind == "confirm":
        ok = page_no >= 2 and not f["letter_kw"] and f["red_seal_ratio"] < SEAL_RED_MAX
        if ok:
            info["rule"] = "P1：确认函第≥2页、彩页/照片类、无函件用语、无印章 → 产品彩页/说明书"
            return "brochure", info
        why = []
        if page_no < 2: why.append("为首页")
        if f["letter_kw"]: why.append(f"含函件用语{f['letter_kw']}")
        if f["red_seal_ratio"] >= SEAL_RED_MAX: why.append(f"有印章(红色{f['red_seal_ratio']})")
        info["rule"] = "P1 不满足（" + "、".join(why) + "），沿用确认函"
        return file_kind, info
    if f["confirm_title_kw"]:
        info["rule"] = f"P2：文档类且含确认函标题用语{f['confirm_title_kw']} → 参数确认函"
        return "confirm", info
    info["rule"] = "P2 不满足（无确认函标题用语），沿用文件类别"
    return file_kind, info


# 图片压缩（r6 选定参数，见 docs/r6-report.md 体积-质量对照）
IMG_TARGET_DPI = None  # 由 build_bid 设置
IMG_JPEG_Q = None


def compress_for_embed(src: Path, w_cm: float, out_dir: Path, dpi: int, q: int) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    dest = out_dir / f"{Path(src).stem}_d{dpi if dpi and dpi > 0 else 'orig'}_q{q}.jpg"
    if dest.is_file():
        return dest
    im = Image.open(src).convert("RGB")
    # r7：dpi<=0 表示「原分辨率」——不主动降采样；只把超过 300 dpi 的图压到略低于 300 dpi，
    # 因为 LibreOffice 导出 PDF 时本来就会把 >300 dpi 的图重采样到 300 dpi（r5/r6 PDF 最大 299.9 dpi），
    # 先在生成器侧做掉，才能让 JPEG 原字节直通匹配上，避免 LO 二次编码。
    tw = int(round(w_cm / 2.54 * dpi)) if dpi and dpi > 0 else min(im.width, int(w_cm / 2.54 * 300) - 2)
    if im.width > tw:
        im = im.resize((tw, int(round(im.height * tw / im.width))), Image.LANCZOS)
    im.save(dest, "JPEG", quality=q, optimize=True)
    return dest


def passthrough_jpegs(pdf_path, embed_dir, log=None):
    """r6：LibreOffice 导出 PDF 时会把 DOCX 内的 JPEG 以其默认质量(≈90)重新编码，体积反而变大、且多一次有损。
    这里把 PDF 中每个 DCT 图像流按「像素尺寸 + 缩略图差异最小」匹配回生成器嵌入的那张 JPEG，原字节直通替换。
    只替换图像流本身（尺寸不变、位置/版面/文字层不变），匹配不上的保持原样；返回统计。"""
    import io
    import os
    import pikepdf
    import numpy as np
    from pathlib import Path
    from PIL import Image

    def thumb(im):
        return np.asarray(im.convert("L").resize((48, 48), Image.BILINEAR), dtype=np.float32)

    cands = {}
    for p in sorted(Path(embed_dir).glob("*.jpg")):
        with Image.open(p) as im:
            cands.setdefault((im.width, im.height), []).append((p, im.mode, thumb(im)))
    pdf = pikepdf.open(str(pdf_path), allow_overwriting_input=True)
    st = {"pdf_images": 0, "replaced": 0, "unmatched": 0, "bytes_before": 0, "bytes_after": 0, "max_thumb_diff": 0.0}
    for obj in list(pdf.objects):
        if not (isinstance(obj, pikepdf.Stream) and obj.get("/Subtype") == "/Image"):
            continue
        st["pdf_images"] += 1
        raw = obj.read_raw_bytes()
        st["bytes_before"] += len(raw)
        key = (int(obj.Width), int(obj.Height))
        best = None
        if str(obj.get("/Filter")) == "/DCTDecode" and key in cands and "/SMask" not in obj:
            try:
                t = thumb(Image.open(io.BytesIO(raw)))
                best = min(((float(np.abs(t - c[2]).mean()), c) for c in cands[key]), key=lambda x: x[0])
            except Exception:
                best = None
        if best is None or best[0] > 6.0:
            st["unmatched"] += 1
            st["bytes_after"] += len(raw)
            continue
        diff, (p, mode, _) = best
        data = p.read_bytes()
        obj.write(data, filter=pikepdf.Name.DCTDecode)
        obj.ColorSpace = pikepdf.Name.DeviceGray if mode == "L" else pikepdf.Name.DeviceRGB
        obj.BitsPerComponent = 8
        if "/DecodeParms" in obj:
            del obj["/DecodeParms"]
        st["replaced"] += 1
        st["bytes_after"] += len(data)
        st["max_thumb_diff"] = max(st["max_thumb_diff"], round(diff, 3))
    # 无损：对象流打包 + flate 重压（只改容器结构，不改任何内容流/图像字节）
    pdf.save(str(pdf_path), object_stream_mode=pikepdf.ObjectStreamMode.generate, compress_streams=True, recompress_flate=True)
    pdf.close()
    st["pdf_bytes"] = os.path.getsize(pdf_path)
    if log:
        log(f"JPEG_PASSTHROUGH {st}")
    return st
