"""r5 修复：附件真实类别（确定性规则）+ 全局唯一、按出现顺序连续编号的图题 + 断言。

r4 问题根因：
1) 设备附件按「文件名含 确认|参数|偏离」归入参数确认函。设备名本身含「参数」（22# 示例设备22），
   于是「…-彩页资料.pdf」被归为参数确认函。
2) 图题序号按「单个文件内页序」从 1 起编；同一设备同一类别有多个文件时（22#、28#、3#）序号重复。
r5 规则：
- 类别：先从文件名中去掉所在文件夹里的设备名，再按关键词判定，优先级 彩页类 > 确认函类 > 其他。
  彩页类关键词：彩页|说明书|产品介绍|样本|画册；确认函类：确认函|确认书|偏离|参数。
- 图题：「图 {全文序号} {章节}-{类别}-{类内序号}」。全文序号 1..N 按出现顺序连续；
  类内序号在（章节, 类别）内跨文件连续。
"""
from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path

BROCHURE_KW = re.compile(r"彩页|说明书|产品介绍|样本|画册")
CONFIRM_KW = re.compile(r"确认函|确认书|偏离|参数")
KIND_LABEL = {"confirm": "参数确认函", "brochure": "产品彩页/说明书", "other": "其他技术证明"}


def device_name_from_folder(folder: str) -> str:
    return re.sub(r"^(?:示例|ex|pkg|包2)-0*\d+#?-?", "", folder).strip()


def classify_equip_file(path: Path) -> tuple[str, str]:
    """返回 (kind, 判定依据)。"""
    dev = device_name_from_folder(path.parent.name)
    residual = path.stem.replace(dev, "") if dev else path.stem
    m = BROCHURE_KW.search(residual)
    if m:
        return "brochure", f"去掉设备名「{dev}」后含「{m.group(0)}」"
    m = CONFIRM_KW.search(residual)
    if m:
        return "confirm", f"去掉设备名「{dev}」后含「{m.group(0)}」"
    return "other", f"去掉设备名「{dev}」后无彩页/确认函关键词"


class CaptionRegistry:
    def __init__(self):
        self.n = 0
        self.group = defaultdict(int)
        self.rows = []

    def next(self, section: str, label: str, suffix: str = "", src: str = "", src_page=None) -> str:
        self.n += 1
        self.group[(section, label)] += 1
        k = self.group[(section, label)]
        cap = f"图 {self.n} {section}-{label}-{k}{suffix}"
        self.rows.append({"n": self.n, "section": section, "label": label, "k": k, "caption": cap,
                          "src": src, "src_page": src_page})
        return cap


CAP_RE = re.compile(r"^图 (\d+) (\S+?)-(.+)-(\d+)(（附件页）)?$")


class CaptionError(AssertionError):
    pass


def audit_captions(docx_path: Path) -> dict:
    """断言：全文图题无重复；全文序号 1..N 连续；（章节, 类别）内序号 1..k 连续；图题数 = 内嵌图数。"""
    from docx import Document
    doc = Document(str(docx_path))
    caps = [p.text.strip() for p in doc.paragraphs if re.match(r"^图\s*\d", p.text.strip())]
    n_pics = len(doc.inline_shapes)
    errs = []
    if len(set(caps)) != len(caps):
        dup = sorted({c for c in caps if caps.count(c) > 1})
        errs.append(f"图题重复 {len(dup)} 个：{dup[:5]}")
    nums, groups = [], defaultdict(list)
    for c in caps:
        m = CAP_RE.match(c)
        if not m:
            errs.append(f"图题格式不符：{c}")
            continue
        nums.append(int(m.group(1)))
        groups[(m.group(2), m.group(3))].append(int(m.group(4)))
    if nums != list(range(1, len(nums) + 1)):
        errs.append("全文图号不连续")
    for g, ks in groups.items():
        if ks != list(range(1, len(ks) + 1)):
            errs.append(f"{g} 类内序号不连续：{ks}")
    if len(caps) != n_pics:
        errs.append(f"图题数 {len(caps)} ≠ 内嵌图数 {n_pics}")
    rep = {"captions": len(caps), "pictures": n_pics, "unique": len(set(caps)), "max_no": max(nums) if nums else 0,
           "groups": len(groups), "errors": errs, "pass": not errs}
    if errs:
        raise CaptionError("图题断言失败：" + "；".join(errs[:5]))
    return rep
