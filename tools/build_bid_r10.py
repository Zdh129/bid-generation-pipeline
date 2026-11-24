#!/usr/bin/env python3
"""按「官方招标文件 + 制作标书模版」生成本地标书，附件取自「生成初版标书所需资料」。
不调用任何外部 LLM/OCR API。不把证件号、姓名等写入日志。
"""
from __future__ import annotations

import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from collections import defaultdict
from pathlib import Path
from repo_paths import load_paths

import numpy as np
import pymupdf
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, Mm
from docx.table import Table
from docx.text.paragraph import Paragraph
from PIL import Image

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
import layout_spec as spec  # noqa: E402
from generate_docx import convert_to_pdf, set_run_font  # noqa: E402

PATHS = load_paths(ROOT)
OUT_DIR = Path(os.environ.get("BID_OUT_DIR") or (ROOT / "output" / "generated-r10"))  # r9：r2–r8 产物保持原样
PIC_LOG: list = []
from bid_r3_fixes import resolve_bid_date, fit_picture_width_cm, apply_page_numbering, date_consistency_report  # noqa: E402
BID_DATE_SRC = resolve_bid_date()
CACHE = ROOT / "work" / "att_cache_r10"  # r9：独立缓存，r3–r8 缓存不动
from bid_r6_fixes import page_kind, compress_for_embed  # noqa: E402
EMBED_DPI, EMBED_Q = int(os.environ.get("BID_EMBED_DPI", "0")), int(os.environ.get("BID_EMBED_Q", "62"))  # r7：donghao 选原分辨率（0=不降采样），JPEG q62
import bid_r2_helpers as _r2h, bid_r4_fixes as _r4f  # noqa: E402
import bid_r7_fixes  # noqa: E402
import bid_r8_fixes  # noqa: E402
import bid_r10_fixes as bid_r9_fixes  # noqa: E402  r10：r9b 叠加层 + 「未提供，不计入」
sys.modules.setdefault("build_bid", sys.modules[__name__])  # r9b：bid_r2_helpers 内部 `from build_bid import …` 指向本模块，避免再执行 r9 版 build_bid 覆盖叠加层
bid_r9_fixes.install(_r4f, _r2h)  # r9：r8 叠加层 + 人工核对改判 + 全部定案后用语改「无偏离」，不改判定文件
import polish_r7  # noqa: E402
PAGE_CLASS_LOG: list = []
from skew_r5 import deskew as r4_deskew  # noqa: E402
from bid_r5_fixes import classify_equip_file, CaptionRegistry, KIND_LABEL, audit_captions  # noqa: E402
CAPS = CaptionRegistry()
EQUIP_CLASS_LOG: list = []
DESKEW_LOG: dict = {}
DEVICES = json.loads((ROOT / "work" / "devices.json").read_text(encoding="utf-8"))
REQ = json.loads((ROOT / "work" / "requirements.json").read_text(encoding="utf-8"))

EA = spec.FONT_BODY_EA
ASC = spec.FONT_ASCII_FALLBACK if True else spec.FONT_ASCII


def log(msg: str):
    print(msg, flush=True)


def set_font(run, size=14, bold=False, ea=EA):
    set_run_font(run, ea=ea, ascii_name=ASC, size_pt=size, bold=bold)


def add_run(p, text, size=14, bold=False, ea=EA):
    r = p.add_run(text)
    set_font(r, size=size, bold=bold, ea=ea)
    return r


def style_or_normal(doc, name):
    try:
        doc.styles[name]
        return name
    except KeyError:
        return "Normal"


def insert_p_after(paragraph: Paragraph, text="", style=None) -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if style:
        try:
            new_para.style = style
        except Exception:
            pass
    new_para.paragraph_format.first_line_indent = Pt(0)
    new_para.paragraph_format.space_after = Pt(2)
    new_para.paragraph_format.space_before = Pt(0)
    if text:
        add_run(new_para, text)
    return new_para


def insert_table_after(paragraph: Paragraph, rows, cols) -> Table:
    tbl = OxmlElement("w:tbl")
    paragraph._p.addnext(tbl)
    # create via a temp table at end then move? easier: use doc.add_table then move xml
    # We'll construct minimal tbl and wrap
    doc = paragraph.part.document if hasattr(paragraph, "part") else None
    # Use body
    table = Table(tbl, paragraph._parent)
    # python-docx Table needs tblGrid and rows – use document.add_table then relocate
    return None


def style_table(table: Table, *, header_repeat=None):
    """V8 风格：全框线 sz=2（0.25pt），长表重复表头，行不可分割。"""
    tblPr = table._tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        table._tbl.insert(0, tblPr)
    old = tblPr.find(qn("w:tblBorders"))
    if old is not None:
        tblPr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "2")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tblPr.append(borders)
    # r3：所有带表头的表（≥2 行）首行设 w:tblHeader，跨页即重复表头（r2 仅 ≥15 行的表）
    repeat = header_repeat if header_repeat is not None else (len(table.rows) >= 2)
    for i, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        if trPr.find(qn("w:cantSplit")) is None:
            trPr.append(OxmlElement("w:cantSplit"))
        if i == 0 and repeat and trPr.find(qn("w:tblHeader")) is None:
            trPr.append(OxmlElement("w:tblHeader"))
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.first_line_indent = Pt(0)
                for r in p.runs:
                    set_font(r, size=14, bold=bool(r.bold))


def add_table_at_end(doc: Document, headers, rows, col_w=None) -> Table:
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_run(p, h, size=14, bold=True)
    for r_i, row in enumerate(rows, 1):
        for c_i, val in enumerate(row):
            cell = table.rows[r_i].cells[c_i]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_i >= 2 else WD_ALIGN_PARAGRAPH.CENTER
            add_run(p, str(val), size=14, bold=False)
    style_table(table)
    return table


def add_heading(doc, text, level=1):
    style = f"Heading {level}"
    try:
        p = doc.add_paragraph(style=style)
    except Exception:
        p = doc.add_paragraph()
    p.clear() if hasattr(p, "clear") else None
    # 避免模板样式未含显式字体：先空段再写入 run
    if p.runs:
        for r in list(p.runs):
            r._element.getparent().remove(r._element)
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(spec.H1_SPACE_BEFORE_PT)
    p.paragraph_format.space_after = Pt(spec.H1_SPACE_AFTER_PT)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(20)
    if level == 1:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(p, text, size=14, bold=True)
    return p


def _np_hl(p):
    """r10：「未提供」段落醒目标注（红字、加粗、黄底），字体字号沿用正文。"""
    from docx.enum.text import WD_COLOR_INDEX
    from docx.shared import RGBColor
    for r in p.runs:
        r.bold = True
        r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return p


def add_body(doc, text, *, center=False, bold=False):
    p = doc.add_paragraph()
    p.style = style_or_normal(doc, "Normal")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if center else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(20)
    add_run(p, text, size=14, bold=bold)
    return p


def estimate_skew(arr, search=5.0, step=0.25):
    if arr.ndim == 3:
        arr = arr.mean(axis=2)
    h, w = arr.shape
    if max(h, w) > 900:
        arr = arr[:: max(1, h // 800), :: max(1, w // 800)]
    bw = arr < 200 if arr.mean() > 127 else arr > 55
    rows, cols = np.any(bw, axis=1), np.any(bw, axis=0)
    if not rows.any() or not cols.any():
        return 0.0, True
    r0, r1 = np.where(rows)[0][[0, -1]]
    c0, c1 = np.where(cols)[0][[0, -1]]
    bw = bw[r0 : r1 + 1, c0 : c1 + 1]
    if bw.size < 400 or bw.mean() < 0.005:
        return 0.0, True
    img = Image.fromarray((bw.astype(np.uint8) * 255))
    best_a, best_v = 0.0, -1.0
    for a in np.arange(-search, search + 1e-9, step):
        rot = img.rotate(float(a), resample=Image.NEAREST, fillcolor=0, expand=False)
        v = float(np.asarray(rot).sum(axis=1, dtype=np.float64).var())
        if v > best_v:
            best_v, best_a = v, float(a)
    return round(best_a, 2), abs(best_a) >= (search - step)


def render_pdf_pages(pdf_path: Path, out_dir: Path, max_pages=None, dpi=185) -> list[Path]:
    """Render pages to JPEG in cache. Deskew if reliable |θ|>0.5. No text extracted to logs.
    dpi=185 且版心宽 16cm 时，A4 页有效显示 DPI ≈ 185*21/16 ≈ 243，贴近 V8 中位。
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    key = hashlib.sha256(f"{pdf_path}|dpi={dpi}".encode()).hexdigest()[:16]
    doc = pymupdf.open(pdf_path)
    paths = []
    n = doc.page_count if max_pages is None else min(doc.page_count, max_pages)
    for i in range(n):
        dest = out_dir / f"{key}_d{dpi}_p{i+1:03d}.jpg"
        side = dest.with_suffix(".deskew.json")
        if dest.is_file() and dest.stat().st_size > 1000 and side.is_file():
            DESKEW_LOG[dest.name] = json.loads(side.read_text(encoding="utf-8"))
            paths.append(dest)
            continue
        page = doc[i]
        pix = page.get_pixmap(matrix=pymupdf.Matrix(dpi / 72, dpi / 72), alpha=False)
        img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        # r4 deskew：仅文档类、可靠且 |θ|>0.3° 时按正确方向旋转（r3 为 rotate(-ang)，方向反）
        img, rec = r4_deskew(img.convert("RGB"))
        rec.update({"src": pdf_path.name, "src_page": i + 1})
        DESKEW_LOG[dest.name] = rec
        side.write_text(json.dumps(rec, ensure_ascii=False), encoding="utf-8")
        img.convert("RGB").save(dest, "JPEG", quality=82, optimize=True)
        paths.append(dest)
    doc.close()
    return paths


def render_image_file(path: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    dest = out_dir / (hashlib.sha256(str(path).encode()).hexdigest()[:16] + path.suffix.lower())
    side = dest.with_suffix(".deskew.json")
    if dest.is_file() and side.is_file():
        DESKEW_LOG[dest.name] = json.loads(side.read_text(encoding="utf-8"))
        return dest
    im = Image.open(path)
    im = im.convert("RGB")
    im, rec = r4_deskew(im)
    rec.update({"src": path.name, "src_page": 1})
    DESKEW_LOG[dest.name] = rec
    side.write_text(json.dumps(rec, ensure_ascii=False), encoding="utf-8")
    if dest.suffix in (".jpg", ".jpeg"):
        im.convert("RGB").save(dest, "JPEG", quality=80)
    else:
        im.save(dest)
    return dest


def add_picture(doc, img_path: Path, width_cm=16.0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    # 图片段落不可使用固定行距，否则会被裁切到 20pt
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.line_spacing = 1.0
    run = p.add_run()
    # r3：按当前节版心等比缩放（宽 ≤ 版心宽-0.1cm；图+题注高 ≤ 版心高）
    w_fit, info = fit_picture_width_cm(doc.sections[-1], Path(img_path), width_cm)
    info["img"] = Path(img_path).name
    # r6/r7：按显示宽度下采样到 EMBED_DPI（r7 为 0=不降采样）、JPEG 质量 EMBED_Q 后嵌入
    if EMBED_Q > 0:
        emb = compress_for_embed(Path(img_path), w_fit, CACHE / "embed", EMBED_DPI, EMBED_Q)
        with Image.open(emb) as _im:
            info["embed_wpx"] = _im.width
        info["embed_img"] = emb.name
        info["eff_dpi_before_compress"] = info["eff_dpi"]
        info["eff_dpi"] = round(info["embed_wpx"] / (w_fit / 2.54), 1)
        info["embed_q"] = EMBED_Q
        img_path = emb
    PIC_LOG.append(info)
    shape = run.add_picture(str(img_path), width=Cm(w_fit))
    # r3：显式 dist*=0。python-docx 不写该属性时 LibreOffice 按 0.125in(≈0.32cm) 默认环绕间距处理，
    # 图宽 >版心-0.64cm 时不再居中而右溢（r2 实测 x0 恒为左边距+9pt）
    for k in ("distT", "distB", "distL", "distR"):
        shape._inline.set(k, "0")
    return p


def add_caption(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Pt(0)
    add_run(p, text, size=14, bold=False)
    if PIC_LOG and "caption" not in PIC_LOG[-1]:
        PIC_LOG[-1]["caption"] = text
    return p


def classify_entity(path: Path) -> str:
    name = path.name
    mapping = [
        ("营业执照", "营业执照"),
        ("法定代表人", "法定代表人证明"),
        ("纳税", "纳税及社会保险缴纳凭证"),
        ("保险", "纳税及社会保险缴纳凭证"),
        ("业绩", "类似项目业绩证明"),
        ("QMS", "质量管理体系证书"),
        ("OHSMS", "职业健康安全管理体系证书"),
        ("EMS", "环境管理体系证书"),
    ]
    for k, v in mapping:
        if k in name:
            return v
    return "资格证明文件"


def collect_attachments():
    att = Path(PATHS["attachments_dir"])
    entity = []
    equip = defaultdict(lambda: {"confirm": [], "brochure": [], "other": []})
    for p in sorted(att.rglob("*")):
        if not p.is_file():
            continue
        rel = p.relative_to(att).as_posix()
        if "投标主体" in rel:
            entity.append(p)
            continue
        if "设备信息" not in rel:
            continue
        m = re.search(r"(?:示例|ex|pkg|包2)-0*(\d+)", p.as_posix())
        if not m:
            continue
        seq = int(m.group(1))
        if p.suffix.lower() in {".pdf", ".jpg", ".jpeg", ".png"}:
            # r5：去掉设备名后按关键词判定真实类别（r4 前：文件名含「参数」即归确认函，设备名含「参数」时误归）
            kind, why = classify_equip_file(p)
            if p.suffix.lower() == ".pdf":
                old = "confirm" if re.search(r"确认|参数|偏离", p.name) else ("brochure" if re.search(r"彩页", p.name) else "other")
            else:
                old = "brochure"
            EQUIP_CLASS_LOG.append({"seq": seq, "file": p.name, "kind_r4": old, "kind_r5": kind, "basis": why})
            equip[seq][kind].append(p)
    return entity, equip


LETTER = """致：示例招标人有限公司{CITE_ADDR}

我方已仔细研究了本项目招标文件的全部内容，包括澄清、修改及有关附件，我方兹以本投标文件对本项目（示例包：示例仪器采购）作出实质性响应，并承诺：

一、我方同意在招标文件规定的投标有效期内（自投标截止之日起不少于 90 个日历天）遵守本投标文件，该文件对我方具有约束力，并可随时被接受。

二、我方的投标报价为一次性固定报价，以人民币计价并含应向国家缴纳的税费，在投标有效期及合同有效期内不予调整；报价不超过招标文件公布的最高投标限价。

三、如我方中标，将按照招标文件、本投标文件及合同约定组织供货、安装、调试、培训和售后服务，保证货物及服务满足招标需求（含技术规格中的★条款）及质量标准，并在规定交货期内完成交付。

四、我方已按招标文件要求准备投标保证金及相关资格、商务、技术证明材料，详见后续章节及附件。

五、与本投标文件一并提交的证明材料如为复印件，其内容与原件一致，并承担相应法律责任。

六、我方理解，招标文件未提供规定格式的内容由投标人自拟；本文件已编制目录及页码。

七、交货期：合同签订之后，接招标人通知之日起 30 日历日内完成供货，并通过验收{CITE_DEL}。

八、质保期：验收合格之日起 1 年{CITE_WAR}。

（投标人名称以营业执照附件为准）
日期：{BID_DATE}
"""

COMMITMENT_STAR = """商务要求及技术★条款响应承诺书

{STAR_COMMIT}

我方承诺：对招标文件及附件 3《示例采购参数》中列明的全部商务要求和全部技术★条款作出实质性响应；★条款如未按要求提供厂家参数确认函、检测报告、技术白皮书或产品说明书等证明材料的，该项视为不满足。证明材料按设备分项附于「6.1 技术参数响应及证明材料」。若参数另行要求提供其他证明的，我方已按对应要求附入投标文件。
"""

WARRANTY = """5.2 质保期承诺

我方承诺：中标后所供仪器、装置及随机软件的质保期为验收合格之日起 1 年（依据招标文件示例页）。质保期内因产品质量、安装调试或随机附件原因产生的故障，我方负责免费维修或更换，并承担由此发生的合理费用。质保期满后，我方继续提供维修、备件及技术支持服务，见「6.3 售后服务方案」。
"""

BIZ_ITEMS = [
    ("投标有效期", "自投标截止之日起不少于 90 个日历天", "响应"),
    ("投标报价", "一次性固定人民币报价、含税，不超过最高限价", "响应"),
    ("投标保证金", "按投标人须知前附表规定提交", "响应（缴付凭证如有则见本章附件，否则按招标要求在递交前补齐）"),
    ("交货期", "合同签订之后，接招标人通知之日起 30 日历日内完成供货，并通过验收。", "合同签订之后，接招标人通知之日起 30 日历日内完成供货，并通过验收（依据招标文件示例页）"),
    ("交货地点", "招标人指定地点。", "响应：招标人指定地点。"),
    ("质保期", "验收合格之日起 1 年。", "验收合格之日起 1 年（依据招标文件示例页）"),
    ("付款方式", "本项目无预付款，货物安装调试完毕，经招标人验收合格后，按实际验收合格数量付至货款的 95%。剩余 5%作为质保金，质保期满后无质量问题,一次性无息支付（若投标人延长质保期的，以延长的质保期为准，质保金退还时间也相应延长）。付款时需提供税率为 13%的增值税专用发票。（如遇国家税率政策调整，按相关文件执行。）", "响应：按招标文件投标人须知前附表第 26 项执行。"),
    ("履约保证金", "中标人接到中标通知书后 30 日内（签订合同前），向招标人递交对应标包中标价的 5%履约保证金。履约保证金形式：银行转账或银行保函。", "响应：按招标文件投标人须知前附表第 27 项执行。"),
    ("质量标准", "满足招标需求、技术规格及国家/行业强制标准", "响应"),
    ("技术资料与合格证明", "供货时以中文附主要图样、合格证明、使用维护说明等", "响应"),
    ("包装运输", "包装方式及费用由中标人负责，运输损坏、丢失由投标人负责", "响应"),
    ("计量器具检定校准", "按招标人要求聘请具备资质的第三方检定校准并提供证书", "响应"),
    ("进口设备证明", "如有进口设备，到货时提供安装维修全部技术资料，验收时提供正品验证渠道或证明", "响应"),
    ("廉洁投标", "不宴请、不送礼、不安排旅游娱乐、不购置物品、不以其他形式行贿", "响应，承诺见 3.3"),
]


def scheme_implementation():
    return [
        ("6.2.1 供货范围与项目组织",
         "供货范围覆盖招标包 2 附件 3 所列全部仪器装置及随机软件、标准附件、必备耗材启动量和中文技术资料。项目实行「项目负责人—技术负责人—现场实施组—厂家协同组」四级组织：项目负责人统筹进度与接口，技术负责人审核技术参数与★条款证明材料的符合性，现场实施组负责到货开箱、安装环境确认与配合，厂家协同组负责原厂或授权服务机构的安装调试与培训。实施前向招标人提交供货清单、到货计划及人员联系方式（联系方式在正式递交文件中载明，本测试稿不写入个人电话与姓名）。"),
        ("6.2.2 供货进度计划与控制",
         "我方按照招标文件前附表规定的交货期倒排计划，一般划分为：合同签订及技术交底、生产/备货与出厂检验、运输与到货验收、安装调试、培训与初步验收、资料归档六个阶段。关键路径为进口件或长周期装置的备货、现场水电及环境条件确认、计量器具的第三方校准。进度以周为周期向招标人书面通报，延误风险提前预警并提供替代方案，确保合同交货期不受影响。"),
        ("6.2.3 全过程质量与成品保护",
         "质量活动覆盖供方评价、进货检验、出厂检验、到货开箱检验、安装过程检验和验收检验。开箱时对照装箱单与配置清单清点主机、电极、传感器、软件介质及随机文件；发现短缺或外观损伤立即记录并补救。安装期间对已开箱仪器采取防尘、防潮、防撞击保护，避免与土建交叉作业损坏设备。软件与方法文件建立版本记录，保证检测数据可追溯。"),
        ("6.2.4 分类安装与调试",
         "按仪器类别分类实施：天平、水分仪等计量类设备需在稳固、无风、温湿度适宜的环境安装并完成校准；发生器、冷干机等气源类设备核对气体管路、纯度与排气；培养箱、烘箱、气候箱等温度类设备进行空载与负载升温试验；力学与包装检测类设备核对夹具、防护与安全连锁；滴定、电化学、光谱类设备完成电极/光路检查与标准物质验证。调试记录随验收资料一并移交。"),
        ("6.2.5 计量试运行与验收",
         "对招标文件要求检定或校准的计量器具，我方聘请具备相应资质的第三方进行检定校准并提供证书。试运行按招标人指定样品或标准物质进行功能演示，确认★条款和主要指标达到投标响应值。验收时提交：配置清单、合格证明、中文说明书、软件授权、校准证书、安装调试记录和培训记录。进口设备同时提供正品验证渠道或证明。"),
        ("6.2.6 厂家协同与实施风险处置",
         "对需原厂或授权服务商配合的设备，我方提前落实服务档期与备件。主要风险及处置：运输破损——保险与原厂更换；现场条件不符——书面告知并协助整改后再安装；进口清关延误——安全库存或同等满足招标要求的应急安排（须事先征得招标人同意）；调试不达标——限时整改，直至满足合同与★条款。"),
    ]


def scheme_service():
    return [
        ("6.3.1 服务组织与故障受理",
         "设立售后服务接口，受理招标人的故障申报、备件需求和技术咨询。受理后登记设备名称、现象与时间，在承诺时限内给出初步判断和处置安排。服务过程形成工单，便于追溯。"),
        ("6.3.2 维修响应与现场服务",
         "质保期内产品质量问题免费维修或更换。接到报修后，我方在合理且不劣于招标文件要求的时限内响应：先远程诊断，必要时安排技术人员到场。关键设备优先保障，重大故障提供备用方案以减少检测中断。"),
        ("6.3.3 回访维护与备品备件",
         "质保期内按约定进行回访或预防性维护，检查易耗件、气路、温控、传感器等状态。常用易损件和耗材建立备件清单，保证合理库存或可在承诺周期内到货，避免因备件短缺延长停机时间。"),
        ("6.3.4 质保期满后服务与资料交接",
         "质保期满后继续提供有偿维修、校准协助、软件升级咨询和操作培训复习。移交完整的设备档案，包括说明书、合格证、校准证书、维修记录和软件介质。"),
    ]


def scheme_training():
    return [
        ("6.4.1 培训目标与组织",
         "培训目标是使招标人指定的使用人员能够独立完成开机、日常检测、质量控制、一般维护和安全操作，并了解数据记录与溯源要求。培训由我方技术负责人和/或厂家应用工程师组织，可按仪器类别分批进行。"),
        ("6.4.2 培训次数与时间安排",
         "安装调试完成后组织现场培训；验收前可根据招标人需要进行补充培训。培训时长按设备复杂程度安排，满足招标文件对培训次数与学时的要求（如招标文件有明确最低学时，按就高执行）。培训时间与招标人协商确定，避免影响生产检测。"),
        ("6.4.3 分类操作与维护重点",
         "培训内容分类实施：计量类——校准、最小称样与环境要求；滴定与电化学类——电极保养、试剂更换与方法设置；气源发生器——纯度监控、滤芯与排水；温箱类——程序设定与超温保护；力学检测类——夹具安装与安全防护；软件类——方法调用、数据导出与备份。同时讲解日常清洁、耗材更换和常见故障判别。"),
        ("6.4.4 考核与培训资料交接",
         "培训结束进行操作复核或问答考核，对未掌握的内容补训。向招标人移交中文培训讲义、操作速查、注意事项及签到/考核记录。"),
    ]



def clear_body_keep_cover(doc: Document):
    """保留封面（首个 Heading 之前），删除模板占位章节，避免与生成正文重复。"""
    body = doc.element.body
    start = None
    for child in list(body.iterchildren()):
        if child.tag == qn("w:p"):
            p = Paragraph(child, doc)
            st = ((p.style.name if p.style else "") or "").lower()
            if st.startswith("heading") or "标题" in st:
                start = child
                break
    if start is None:
        return
    el = start
    while el is not None:
        nxt = el.getnext()
        if el.tag == qn("w:sectPr"):
            break
        parent = el.getparent()
        parent.remove(el)
        el = nxt


def apply_document_styles(doc: Document):
    from generate_docx import set_style_font
    mapping = [
        ("Normal", 14, False),
        ("Title", 48, True),
        ("Heading 1", 14, True),
        ("Heading 2", 14, True),
        ("Heading 3", 14, True),
    ]
    for name, size, bold in mapping:
        try:
            st = doc.styles[name]
        except KeyError:
            continue
        set_style_font(st, EA, ASC, size, bold=bold)
        pf = st.paragraph_format
        pf.first_line_indent = Pt(0)
        if name == "Normal":
            pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            pf.line_spacing = Pt(20)
            pf.space_after = Pt(4)
        elif name.startswith("Heading"):
            pf.space_before = Pt(spec.H1_SPACE_BEFORE_PT)
            pf.space_after = Pt(spec.H1_SPACE_AFTER_PT)
            pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            pf.line_spacing = Pt(20)
            if name == "Heading 1":
                pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif name == "Title":
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf.space_after = Pt(12)
    for st in doc.styles:
        name = st.name or ""
        if name.startswith("投标"):
            set_style_font(st, EA, ASC, 14, bold=False)
    # 封面：文题居中并写字体；不改写封面栏位中的空格占位
    for p in doc.paragraphs:
        st = (p.style.name if p.style else "") or ""
        if st == "Title" or (p.text or "").strip() == "投标文件":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Pt(0)
            if not p.runs:
                add_run(p, "投标文件", size=48, bold=True)
            else:
                for r in p.runs:
                    set_font(r, size=48, bold=True)
        elif st.startswith("投标"):
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            p.paragraph_format.line_spacing = Pt(20)
            for r in p.runs:
                if (r.text or "").strip() or True:
                    set_font(r, size=14, bold=bool(r.bold))


def polish_document(doc: Document):
    """生成后补齐字体、行距、表格边框，覆盖残留封面 run。"""
    for p in doc.paragraphs:
        st = ((p.style.name if p.style else "") or "")
        if st.lower().startswith("heading") or "标题" in st:
            p.paragraph_format.space_before = Pt(spec.H1_SPACE_BEFORE_PT)
            p.paragraph_format.space_after = Pt(spec.H1_SPACE_AFTER_PT)
            p.paragraph_format.first_line_indent = Pt(0)
            for r in p.runs:
                set_font(r, size=14, bold=True)
        elif st == "Title":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                set_font(r, size=48, bold=True)
        else:
            has_drawing = bool(p._p.findall(".//" + qn("w:drawing"))) or "w:drawing" in p._p.xml
            if has_drawing:
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
                p.paragraph_format.line_spacing = 1.0
            elif p.paragraph_format.line_spacing is None:
                p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                p.paragraph_format.line_spacing = Pt(20)
            p.paragraph_format.first_line_indent = Pt(0)
            for r in p.runs:
                rPr = r._element.find(qn("w:rPr"))
                fonts = rPr.find(qn("w:rFonts")) if rPr is not None else None
                if fonts is None or not fonts.get(qn("w:eastAsia")):
                    set_font(r, size=14 if (r.font.size is None) else (r.font.size.pt or 14), bold=bool(r.bold))
    for tbl in doc.tables:
        style_table(tbl)


def build():
    t0 = time.time()
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    CACHE.mkdir(parents=True, exist_ok=True)
    tmpl_dir = Path(PATHS["template_dir"])
    # Prefer 2.3 (AI-refined counterpart of the template); fallback 2.1
    tmpl = None
    for p in sorted(tmpl_dir.glob("*.docx")):
        if "2.3" in p.name or "成品对照" in p.name:
            tmpl = p
            break
    if tmpl is None:
        tmpl = next(p for p in tmpl_dir.glob("*.docx") if "提示词" not in p.name)
    dest = OUT_DIR / "bid_from_template.docx"
    shutil.copy(tmpl, dest)
    log(f"BASE_TEMPLATE size={tmpl.stat().st_size}")

    doc = Document(str(dest))
    # Standardize section margins toward V8 first-section defaults (don't destroy landscape if any)
    for sec in doc.sections:
        w, h = sec.page_width, sec.page_height
        # only touch portrait-like
        if w and h and w <= h:
            sec.top_margin = Cm(spec.MARGIN_TOP_CM)
            sec.bottom_margin = Cm(spec.MARGIN_BOTTOM_CM)
            sec.left_margin = Cm(spec.MARGIN_LEFT_CM)
            sec.right_margin = Cm(spec.MARGIN_RIGHT_CM)
            sec.page_width = Mm(210)
            sec.page_height = Mm(297)
        # footer page field: keep if already present

    apply_document_styles(doc)
    clear_body_keep_cover(doc)

    # 封面末段若已带分节符则不必再插入分页符，避免空白页
    has_section_break = False
    for p in doc.paragraphs:
        pPr = p._p.find(qn("w:pPr"))
        if pPr is not None and pPr.find(qn("w:sectPr")) is not None:
            has_section_break = True
            break
    if not has_section_break:
        pb = doc.add_paragraph()
        pb.paragraph_format.first_line_indent = Pt(0)
        br = pb.add_run()
        set_font(br, size=14)
        br.add_break(WD_BREAK.PAGE)

    # 封面非个人字段（项目/公司/包号/日期）来自文件夹 1 与 3，不取自 V8
    from bid_r2_helpers import (
        fill_cover, add_toc_field, add_oriented_section, load_ceilings, load_judgements,
        cite_evidence, fill_pending_cells, fmt_money, PACKAGE_CEILING, PROJECT_NAME,
        PROJECT_CODE, BIDDER_NAME, patch_settings_update_fields,
        collect_docx_headings, write_populated_toc,
    )
    fill_cover(doc)

    # Title page already kept from template. TOC 使用域，不用静态页码列表。
    add_toc_field(doc, page_break=False)
    # r3：目录单独成节（前置部分，罗马页码），正文节从 1 起编
    add_oriented_section(doc, landscape=False)

    add_heading(doc, "1、投标函", 1)
    for para in (LETTER.strip().replace("{BID_DATE}", BID_DATE_SRC["cn"]).replace("{CITE_ADDR}", bid_r9_fixes.cite("addressee"))
                 .replace("{CITE_DEL}", bid_r9_fixes.cite("delivery")).replace("{CITE_WAR}", bid_r9_fixes.cite("warranty"))).split("\n"):
        add_body(doc, para)

    add_oriented_section(doc, landscape=True)
    add_heading(doc, "投标函附录-分项报价表", 2)
    add_body(doc, "说明：数量、单位及「最高限价」取自官方招标文件《最高限价》清单（文件夹 1），不是对照稿 V8 的单价。本投标单价、合价及投标总价在资料包与 V8 中均无投标人报价依据，标为醒目的「未提供」，正式递交前由投标人填写；各项投标单价不得超过对应最高限价单价。计价货币为人民币，一次性固定含税价。包 2 最高限价合计 "+fmt_money(PACKAGE_CEILING)+" 元。")
    ceilings = load_ceilings()
    # r8：先套用判定叠加层得到设备汇总；品牌/型号仅从厂家参数确认函本地提取，读不准的留【待填】
    _r4f.load_judgement_index()
    DEV_SUMMARY = bid_r9_fixes.device_summary()
    BRAND_MODEL = json.loads((ROOT / "work" / "r9" / "brand_model_r9.json").read_text(encoding="utf-8"))
    BRAND_MODEL = {int(k): v for k, v in BRAND_MODEL.items()}
    (OUT_DIR / "brand_model_r10.json").write_text(json.dumps(BRAND_MODEL, ensure_ascii=False, indent=1), encoding="utf-8")
    from bid_r2_helpers import mark_pending

    def fill_pending_cells_one(cell, marker="【待填】"):
        p = cell.paragraphs[0]
        for r in list(p.runs):
            r._element.getparent().remove(r._element)
        mark_pending(p.add_run(), marker)
    quote_rows = []
    for d in DEVICES:
        c = ceilings.get(d["seq"], {})
        qty = c.get("qty", 1)
        unit = c.get("unit", "台")
        cu = c.get("ceiling_unit")
        ceil_s = fmt_money(cu) if cu else "见官方限价清单"
        bm = dict(BRAND_MODEL[d["seq"]])
        if d["seq"] in bid_r9_fixes.NOT_PROVIDED["brand"]["devices"] and bm["brand"] == "【待填】":
            bm["brand"] = "未提供"
        if d["seq"] in bid_r9_fixes.NOT_PROVIDED["model"]["devices"] and bm["model"] == "【待填】":
            bm["model"] = "未提供"
        quote_rows.append((d["seq"], d["name"].replace("（七）",""), bm["brand"], bm["model"], unit, qty, "未提供", "未提供", ceil_s,
                           DEV_SUMMARY.get(d["seq"], "【待填】"), "【待填】"))
    qtable = add_table_at_end(
        doc,
        ["序号", "货物名称", "品牌", "型号", "单位", "数量", "单价（人民币）", "合价（人民币）", "最高限价单价（元）", "偏离情况", "证明材料页码范围"],
        quote_rows,
    )
    fill_pending_cells(qtable, [6, 7], marker="未提供")
    # r8：11 栏按内容分配列宽（等宽时货物名称/型号折行过多）
    _sec = doc.sections[-1]
    _avail = (_sec.page_width - _sec.left_margin - _sec.right_margin)
    _rel = [1.0, 3.0, 2.6, 3.6, 1.0, 1.0, 1.9, 1.9, 2.0, 2.6, 2.4]
    from docx.shared import Emu
    qtable.autofit = False
    for row in qtable.rows:
        for c_i, cell in enumerate(row.cells):
            cell.width = Emu(int(_avail * _rel[c_i] / sum(_rel)))
    # LibreOffice 按 tblGrid/gridCol 排版：同步写入网格列宽（twips）与固定布局
    _grid = qtable._tbl.tblGrid
    for c_i, gc in enumerate(_grid.findall(qn("w:gridCol"))):
        gc.set(qn("w:w"), str(int(_avail * _rel[c_i] / sum(_rel) / 635)))
    _tblPr = qtable._tbl.tblPr
    _lay = OxmlElement("w:tblLayout"); _lay.set(qn("w:type"), "fixed"); _tblPr.append(_lay)
    for r_i, row in enumerate(qtable.rows[1:], 1):
        for c in (2, 3, 10):
            if row.cells[c].text.strip() in ("【待填】", "未提供"):
                fill_pending_cells_one(row.cells[c], row.cells[c].text.strip())
    build._qtable = qtable
    _np_hl(add_body(doc, "投标总价（45 项合价合计）：未提供（资料包与 V8 均无报价，正式递交前由投标人填写，且不得超过包 2 最高限价 "+fmt_money(PACKAGE_CEILING)+" 元）"))
    add_body(doc, "注：品牌、型号经多来源交叉校验，须在厂家确认函（文字层或本机 400dpi OCR）、附件文件名、彩页/产品资料、厂家承诺函、V8 分项报价表中至少两个独立来源一致才填写，来源逐项记录于 sources_r9.json；资料缺失的标「未提供」（品牌：设备20、34、40；型号：设备18、20、23、29）；设备23、33、39、41 的品牌在厂家确认函（示例厂家甲）与 V8（示例厂家乙）之间互相矛盾，不属于缺失，仍留【待填】由投标人核定。偏离情况为该设备 6.1 节各条结论的汇总（全部无偏差写「"+("无偏离" if bid_r9_fixes.VOCAB.get("changed") else "响应")+"」，否则列出负偏离/待确认/正偏离条数；设备4、6 的排污阀条款按投标人决定保留「待确认」）。证明材料页码范围为该设备 6.1 节证明材料图像所在的印刷页码，由生成后的 PDF 回填。")
    add_oriented_section(doc, landscape=False)

    add_heading(doc, "2、投标保证金", 1)
    add_body(doc, "我方承诺按照招标文件投标人须知前附表规定的金额、形式和时限提交投标保证金，并将其作为投标的一部分。如在正式递交前已取得缴付凭证，凭证附于本章；本测试稿如未附缴付凭证，不代表免除招标文件规定的保证金义务。")

    add_heading(doc, "3、投标人基本情况", 1)
    add_heading(doc, "3.1 投标人基本情况表", 2)
    add_body(doc, "下表按招标常用栏目列出。涉及投标人名称、地址、法定代表人等具体信息，以所附营业执照及法定代表人证明扫描件为准，本表不重复抄录证件号码与个人姓名。")
    add_table_at_end(
        doc,
        ["栏目", "内容"],
        [
            ("投标人名称", "见附件：营业执照"),
            ("注册地址 / 经营场所", "见附件：营业执照"),
            ("统一社会信用代码", "见附件：营业执照（本文件不抄录号码）"),
            ("法定代表人", "见附件：法定代表人证明（本文件不抄录姓名与证件号）"),
            ("企业性质 / 成立日期", "见附件：营业执照"),
            ("经营范围与本项目相关性", "示例仪器相关业务，详见营业执照经营范围"),
            ("联系方式", "正式递交文件中载明；本测试稿不写入个人电话"),
        ],
    )

    add_heading(doc, "3.2 投标人资格相关资料", 2)
    add_body(doc, "本章按招标资格要求附具证明材料。下列扫描件仅在本机生成件中嵌入，报告与日志不摘录证件号、姓名等个人信息。")
    entity, equip = collect_attachments()
    att_counts = {"entity_files": 0, "entity_pages": 0, "equip_files": 0, "equip_pages": 0}
    for p in entity:
        role = classify_entity(p)
        add_heading(doc, role, 3) if False else add_body(doc, f"附件：{role}", bold=True)
        add_body(doc, "（原件扫描，内容以图像为准；本处不转写证件文字。）")
        try:
            if p.suffix.lower() == ".pdf":
                pages = render_pdf_pages(p, CACHE / "entity")
            else:
                pages = [render_image_file(p, CACHE / "entity")]
        except Exception as exc:
            log(f"ENTITY_RENDER_FAIL role={role} err={type(exc).__name__}")
            pages = []
        att_counts["entity_files"] += 1
        att_counts["entity_pages"] += len(pages)
        for i, img in enumerate(pages, 1):
            add_picture(doc, img, 16.0)
            add_caption(doc, CAPS.next("3.2", role, "（附件页）", src=p.name, src_page=i))

    add_body(doc, "附件：信用查询截图", bold=True)
    for _site in bid_r9_fixes.NOT_PROVIDED["credit_screens"]["items"]:
        _np_hl(add_body(doc, f"{_site} 查询截图：未提供"))
    add_heading(doc, "3.3 投标人资格承诺函", 2)
    add_body(doc, "致：示例招标人有限公司" + bid_r9_fixes.cite("addressee"))
    add_body(doc, "（以下按招标文件第八章「3.3 投标人资格承诺函（仅供参考）」原文照录；标注「【未提供：公司最终确认】」的条目尚未取得公司最终确认（其中注明「已对照…证据」者已对照资料包证据），须由投标人逐条确认属实后方可保留。）")
    for _t, _cls in bid_r9_fixes.qualification_letter_r10():
        add_body(doc, _t)
    add_body(doc, "如上述承诺不实，我单位愿承担由此产生的全部责任。")
    add_body(doc, "我方承诺：具备承担本项目的法定资格与履约能力；所附营业执照、体系证书、纳税或社会保险缴纳凭证、业绩证明真实有效；不以任何理由宴请相关方或赠送礼品、礼金及有价证券；不以洽谈、签约等为借口邀请相关方旅游或进入营业性娱乐场所；不为相关方购置通讯、交通、家电或高档办公用品；不以其他任何形式向相关方行贿。如有不实，我方承担招标文件规定的责任。")

    add_heading(doc, "3.4 商务要求及技术★条款响应承诺书", 2)
    _star_t, _star_log = bid_r9_fixes.star_commit_text()
    bid_r9_fixes.COMMIT_LOG.append(_star_log)
    (OUT_DIR / "commitment_check_r10.json").write_text(json.dumps(bid_r9_fixes.COMMIT_LOG, ensure_ascii=False, indent=1), encoding="utf-8")
    for para in COMMITMENT_STAR.strip().replace("{STAR_COMMIT}", _star_t).split("\n"):
        add_body(doc, para)
    add_body(doc, f"本包技术参数共解析 {REQ['counts']['tech_params']} 条，其中★条款 {REQ['counts']['star_params']} 条，分项响应见第 6.1 节。")

    add_heading(doc, "4、投标人代表身份证明", 1)
    add_heading(doc, "法定代表人身份证明书", 2)
    add_body(doc, "法定代表人身份证明以附件扫描件为准。本测试生成稿只嵌入原件图像，不在正文中抄录姓名、职务与证件号码。若正式投标需同时提交授权委托书，授权委托书按招标第八章格式另行签署后替换或增补本章。")
    # legal_rep already in entity loop if present

    add_heading(doc, "5、商务文件", 1)
    add_heading(doc, "附表1：企业业绩", 2)
    add_body(doc, "类似项目业绩证明见 3.2 所附业绩资料扫描件。业绩合同金额、委托方等信息以扫描件为准，本处不摘录可能涉及的单位内部编号或个人姓名。")

    add_heading(doc, "5.2 质保期承诺", 2)
    for para in WARRANTY.strip().split("\n"):
        add_body(doc, para)

    add_heading(doc, "5.3 商务响应表", 2)
    add_table_at_end(
        doc,
        ["序号", "商务要求", "招标要求摘要", "投标响应"],
        [(i + 1, a, b, c) for i, (a, b, c) in enumerate(BIZ_ITEMS)],
    )

    add_heading(doc, "6、技术文件", 1)
    add_heading(doc, "6.1 技术参数响应及证明材料", 2)
    add_body(doc, "下列各表依据招标附件 3 的设备参数编制，共 "+str(REQ['counts']['tech_params'])+" 条（其中★ "+str(REQ['counts']['star_params'])+" 条）。每一条给出偏离结论（"+("无偏离/正偏离/负偏离，按招标文件「偏离情况栏填写正偏离或负偏离或无偏离」；其中 P04-5、P06-5 两条（示例手动阀）V8 写全自动、厂家函写手动，暂标「待确认」，须人工确认后改为招标用语" if bid_r9_fixes.VOCAB.get("changed") else "响应/正偏离/负偏离/待确认/无证据待确认")+"）及引用证据（文件名+页+片段）。证据栏标〔规则A/B/C〕者为按统一口径归并的结论（A：附件有相关表述；B：依据厂家参数确认函整体声明；C：优于要求），需人工复核；标〔r9人工核对改判〕者为逐条核对招标原文、厂家函、彩页与 V8 后的定案，证据写在该栏；标〔未提供：厂家参数确认函〕者（设备34、40）厂家参数确认函未提供，结论暂依彩页/参数页，不计入一致性判定。★条款必须满足；证明材料以所附图像为准。参数文本来自招标文件。本表置于横向节，以便宽表阅读。")
    from bid_r4_fixes import build_param_rows
    PARAM_ROWS, join_log = build_param_rows(DEVICES, cite_evidence)
    (OUT_DIR / "verdict_join_log.json").write_text(json.dumps(join_log, ensure_ascii=False, indent=0), encoding="utf-8")
    (OUT_DIR / "verdict_overlay_r10.json").write_text(json.dumps(bid_r9_fixes.summary(), ensure_ascii=False, indent=1), encoding="utf-8")
    add_oriented_section(doc, landscape=True)

    # 横向节只放宽表（参照 V8：报价表与参数表横向；证明图像回到纵向）
    for d in DEVICES:
        seq = d["seq"]
        name = d["name"].replace("（七）", "")
        # 不用 Heading 3，避免 45 台设备名挤进自动目录（目录只收 1–3 级正式章节）
        add_body(doc, f"6.1.{seq} {name}", bold=True)
        params = d.get("params") or []
        if params:
            # r4：按（设备序号, 行序）稳定唯一键对位，生成前已断言参数文本一致
            rows = PARAM_ROWS[seq]
            add_table_at_end(doc, ["序号", "招标参数要求", "偏离结论", "引用证据"], rows)
        else:
            add_body(doc, "该设备在招标扫描/转写文本中未能完整拆出逐条参数。证明材料见所附确认函及彩页。")

    add_oriented_section(doc, landscape=False)
    add_body(doc, "技术证明材料（参数确认函、彩页等）按设备顺序附后，以图像为准。")
    for d in DEVICES:
        seq = d["seq"]
        name = d["name"].replace("（七）", "")
        files = equip.get(seq, {})
        added = 0
        add_body(doc, f"附件 6.1.{seq} {name}", bold=True)
        for kind, cap in KIND_LABEL.items():
            for f in files.get(kind, []):
                try:
                    if f.suffix.lower() == ".pdf":
                        mx = None if kind == "confirm" else 4
                        pages = render_pdf_pages(f, CACHE / f"eq{seq:02d}", max_pages=mx)
                    else:
                        pages = [render_image_file(f, CACHE / f"eq{seq:02d}")]
                except Exception as exc:
                    log(f"EQUIP_RENDER_FAIL seq={seq} kind={kind} err={type(exc).__name__}")
                    pages = []
                att_counts["equip_files"] += 1
                att_counts["equip_pages"] += len(pages)
                for j, img in enumerate(pages, 1):
                    add_picture(doc, img, 16.0)
                    # r6：单页内容级归类（只在页内容与文件类别冲突时检查，规则见 bid_r6_fixes）
                    ck = (DESKEW_LOG.get(Path(img).name) or {}).get("class", {}).get("kind", "document")
                    pk, pinfo = page_kind(kind, j, ck, Path(img))
                    PAGE_CLASS_LOG.append({"seq": seq, "file": f.name, "page": j, "file_kind": kind, "page_kind": pk, **pinfo})
                    add_caption(doc, CAPS.next(f"6.1.{seq}", KIND_LABEL[pk], src=f.name, src_page=j))
                    added += 1
        if seq in bid_r9_fixes.NOT_PROVIDED["confirm_letter"]["devices"]:
            _np_hl(add_body(doc, f"厂家参数确认函（设备 {seq}）：未提供"))
        if added == 0:
            add_body(doc, "本测试资料包中未检索到该设备的确认函或彩页文件，技术响应以本表声明为准，正式递交时应补齐证明。")

    add_heading(doc, "6.2 实施方案", 2)
    for title, body in scheme_implementation():
        add_heading(doc, title, 3)
        add_body(doc, body)

    add_heading(doc, "6.3 售后服务方案", 2)
    for title, body in scheme_service():
        add_heading(doc, title, 3)
        add_body(doc, body)

    add_heading(doc, "6.4 培训方案", 2)
    for title, body in scheme_training():
        add_heading(doc, title, 3)
        add_body(doc, body)

    add_heading(doc, "7、投标人需提交的其它材料", 1)
    add_body(doc, "（1）招标文件要求的其他材料：如招标人后续澄清要求补充的文件，我方将按期提交。")
    add_body(doc, "（2）投标人认为有必要的其他材料：本测试生成稿已将资料包中的资格证明与分项技术证明作为附件嵌入；资料包未提供的保证金缴付凭证等，正式投标时按招标文件补齐。")

    # r3：统一页码规则（替代 r2「缺 PAGE 才补」的逻辑，r2 中节1 起页脚只有裸数字）
    t_layout = time.time()
    pgn_plan = apply_page_numbering(doc, set_font, front_sections=(1,), body_start_section=2)
    log(f"PGN_PLAN {json.dumps(pgn_plan, ensure_ascii=False)}")

    # r7 润色：先出问题清单，再做格式级安全修改（逐处记录），改后再出一次清单
    _lint_before = polish_r7.lint_document(doc)
    POLISH_CHANGES = polish_r7.apply_safe_fixes(doc)
    _lint_after = polish_r7.lint_document(doc)
    (OUT_DIR / "polish_r7.json").write_text(json.dumps({"before": polish_r7.summarize(_lint_before, []), "after": polish_r7.summarize(_lint_after, POLISH_CHANGES),
        "changes": POLISH_CHANGES, "issues_after": _lint_after}, ensure_ascii=False, indent=0), encoding="utf-8")
    log(f"POLISH issues_before={len(_lint_before)} safe_fix_runs={len(POLISH_CHANGES)} issues_after={len(_lint_after)}")
    polish_document(doc)
    doc.save(str(dest))
    patch_settings_update_fields(dest)
    elapsed_docx = time.time() - t0
    log(f"DOCX_SAVED {dest} bytes={dest.stat().st_size} elapsed={elapsed_docx:.1f}s att={att_counts}")
    # r5：图题断言（重复/不连续即抛错中止）
    cap_rep = audit_captions(dest)
    (OUT_DIR / "caption_audit.json").write_text(json.dumps({**cap_rep, "rows": CAPS.rows, "equip_classification": EQUIP_CLASS_LOG, "page_classification": PAGE_CLASS_LOG}, ensure_ascii=False, indent=0), encoding="utf-8")
    log(f"CAPTION_AUDIT {json.dumps({k: v for k, v in cap_rep.items()}, ensure_ascii=False)}")
    # r4：生成后逐行断言（不对应即抛错中止）
    from bid_r4_fixes import audit_docx_param_tables
    from bid_r2_helpers import cite_evidence as _cite
    t_a = time.time()
    vj = audit_docx_param_tables(dest, DEVICES, _cite)
    (OUT_DIR / "verdict_join_audit.json").write_text(json.dumps(vj, ensure_ascii=False, indent=1), encoding="utf-8")
    log(f"VERDICT_AUDIT rows={vj['rows_checked']} mismatch={vj['row_mismatch']} agree_v8={vj['docx_agree_v8']} rate={vj['rate']} dist={vj['docx_verdict_dist']} s={time.time()-t_a:.1f}")
    deskew_rows = [{**{k: x.get(k) for k in ("caption", "img")}, **DESKEW_LOG.get(x.get("img"), {})} for x in PIC_LOG]
    (OUT_DIR / "deskew_log.json").write_text(json.dumps(deskew_rows, ensure_ascii=False, indent=0), encoding="utf-8")

    from pdf_safe import open_pdf, footer_labels
    from check_layout import heading_pages_in_pdf
    pdf_path = OUT_DIR / "bid_from_template.pdf"
    stage = {"docx_build_s": round(elapsed_docx, 2)}

    def printed_pages():
        labs = footer_labels(pdf_path)
        return {l["page"]: (l["n"] if l["kind"] == "arabic" else None) for l in labs}, labs

    def heading_printed(hs):
        pdf = open_pdf(pdf_path)
        phys = heading_pages_in_pdf(pdf, hs)
        n_pages = pdf.page_count
        pdf.close()
        lab, _ = printed_pages()
        return {k: lab.get(v) for k, v in phys.items()}, phys, n_pages

    # 两遍目录：第 1 次转 PDF 取各标题「印刷页码」（页脚页码，非物理页序）写回目录，再转第 2 次核对；不一致则再写一遍
    # r8：分项报价表「证明材料页码范围」= 该设备 6.1 节证明图像题注（「图 N 6.1.{序号}-…」）所在页的印刷页码范围
    def device_page_ranges():
        import re as _re
        pdf = open_pdf(pdf_path)
        texts = list(pdf._texts)
        pdf.close()
        lab, _ = printed_pages()
        pages = {}
        for i, t in enumerate(texts, 1):
            for m in _re.finditer(r"图\s*\d+\s*6\s*\.\s*1\s*\.\s*(\d+)\s*-", t or ""):
                if lab.get(i) is not None:
                    pages.setdefault(int(m.group(1)), set()).add(lab[i])
        out = {}
        for seq, ps in pages.items():
            lo, hi = min(ps), max(ps)
            out[seq] = f"第{lo}页" if lo == hi else f"第{lo}–{hi}页"
        return out

    def write_page_ranges(rng):
        for row in build._qtable.rows[1:]:
            seq = int(row.cells[0].text.strip())
            cell = row.cells[10]
            p = cell.paragraphs[0]
            for r in list(p.runs):
                r._element.getparent().remove(r._element)
            if seq in rng:
                add_run(p, rng[seq], size=14, bold=False)
            else:
                mark_pending(p.add_run(), "【待填】")

    hs = collect_docx_headings(doc)
    passes = []
    build._page_ranges = None
    for pass_no in (1, 2, 3, 4, 5):
        t1 = time.time()
        convert_to_pdf(dest, pdf_path, timeout=1800)
        conv_s = time.time() - t1
        printed, phys, n_pages = heading_printed(hs)
        written = {e["text"]: e.get("page") for e in getattr(build, "_toc_entries", [])}
        match = sum(1 for h in hs if written.get(h["text"]) is not None and written.get(h["text"]) == printed.get(h["text"]))
        rng = device_page_ranges()
        rng_ok = build._page_ranges == rng
        passes.append({"pass": pass_no, "convert_s": round(conv_s, 2), "pages": n_pages,
                       "toc_written": len(written), "toc_match": match, "page_ranges_found": len(rng), "page_ranges_stable": rng_ok})
        log(f"TOC_PASS {passes[-1]}")
        if written and match == len(hs) and rng_ok:
            break
        entries = [{**h, "page": printed.get(h["text"])} for h in hs]
        build._toc_entries = entries
        write_populated_toc(doc, entries)
        build._page_ranges = rng
        write_page_ranges(rng)
        doc.save(str(dest))
        patch_settings_update_fields(dest)
    stage["pdf_convert_passes"] = passes
    stage["pdf_total_s"] = round(sum(p["convert_s"] for p in passes), 2)
    # r6：JPEG 原字节直通（避免 LibreOffice 二次编码）+ 无损对象流打包
    if EMBED_Q > 0:
        from bid_r6_fixes import passthrough_jpegs
        t_pt = time.time()
        stage["jpeg_passthrough"] = {**passthrough_jpegs(pdf_path, CACHE / "embed", log), "s": round(time.time() - t_pt, 2)}

    # 目录核对：解析 PDF 目录页（含续页）条目 vs 正文标题印刷页码
    from verify_toc import verify
    toc_verify = verify(pdf_path, out=OUT_DIR / "toc_verify.json")
    log(f"TOC_VERIFY {json.dumps({k: v for k, v in toc_verify.items() if k != 'rows'}, ensure_ascii=False)}")

    # d. 日期单点校验（DOCX 全文 + PDF 文字层）
    pdf = open_pdf(pdf_path)
    texts = list(pdf._texts)
    n_pages = pdf.page_count
    landscape = len(pdf.landscape_pages)
    pdf.close()
    date_rep = date_consistency_report(Document(str(dest)), texts, BID_DATE_SRC)
    (OUT_DIR / "date_check.json").write_text(json.dumps(date_rep, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"DATE_CHECK pass={date_rep['pass']} fields_docx={date_rep['docx_date_fields']} values={date_rep['docx_field_values']} pdf_pages={date_rep['pdf_field_pages']}")
    (OUT_DIR / "picture_fit.json").write_text(json.dumps(PIC_LOG, ensure_ascii=False, indent=1), encoding="utf-8")

    meta = {
        "round": "r10",
        "docx": str(dest),
        "pdf": str(pdf_path),
        "docx_bytes": dest.stat().st_size,
        "pdf_bytes": pdf_path.stat().st_size if pdf_path.is_file() else None,
        "pdf_pages": n_pages,
        "landscape_pages": landscape,
        "stage_s": stage,
        "elapsed_total_s": round(time.time() - t0, 2),
        "attachment_counts": att_counts,
        "devices": len(DEVICES),
        "tech_params": REQ["counts"]["tech_params"],
        "star_params": REQ["counts"]["star_params"],
        "api_cost_yuan": 0,
        "bid_date": BID_DATE_SRC,
        "page_numbering": pgn_plan,
        "pictures": {"n": len(PIC_LOG), "max_w_cm": max((x["final_w_cm"] for x in PIC_LOG), default=None),
                     "min_w_cm": min((x["final_w_cm"] for x in PIC_LOG), default=None),
                     "text_w_cm": sorted({x["text_w_cm"] for x in PIC_LOG})},
        "toc_verify": {k: v for k, v in toc_verify.items() if k != "rows"},
        "date_check_pass": date_rep["pass"],
        "embed": {"dpi": EMBED_DPI, "jpeg_q": EMBED_Q},
        "page_class_changed": [x for x in PAGE_CLASS_LOG if x["page_kind"] != x["file_kind"]],
        "verdict_overlay": {k: v for k, v in bid_r9_fixes.summary().items() if k in ("vocab_changed_to_无偏离", "pending_after", "dist_before", "dist_r7", "dist_r8", "dist_after", "by_rule", "v8_agree_before", "v8_agree_r7", "v8_agree_r8", "v8_agree_after", "r10_consistency")},
        "quote_page_ranges": getattr(build, "_page_ranges", None),
        "note": "r10（r9b 基础上：缺失资料标「未提供」并从一致性判定剔除）：正文本机撰写，未调用外部 LLM/OCR API；附件仅本地渲染嵌入。",
        "verdict_audit": {k: v for k, v in vj.items() if k not in ("disagreements", "bad_rows")},
        "captions": cap_rep,
        "deskew": {"applied": sum(1 for r in deskew_rows if r.get("applied")), "n": len(deskew_rows)},
    }
    (OUT_DIR / "build_meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"DONE {json.dumps(meta, ensure_ascii=False)}")
    return meta


if __name__ == "__main__":
    build()
