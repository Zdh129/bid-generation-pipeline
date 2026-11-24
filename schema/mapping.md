# 识别侧输出到中间格式的映射说明

- 目标 schema：`schema/intermediate.schema.json`（draft 2020-12，版本 `1.0.0`）
- 坐标系：**页面左上角为原点，x 向右、y 向下**。这与 PyMuPDF `page.get_text("dict")` 及渲染位图像素坐标一致，与 PDF 原生用户空间（左下原点）相反。
- 推荐单位：页面 `unit` 与块 `bbox_unit` 使用 **`pt`**（1 pt = 1/72 inch）。若沿用 PaddleOCR 的像素框，必须设 `bbox_unit="px"`，并填写 `bbox_dpi`（通常等于该页 `render_dpi`）。
- 换算：`pt = px * 72 / dpi`，`px = pt * dpi / 72`。识别侧约定渲染 `render_dpi ≥ 216`（`zoom ≥ 3`），**推荐 300**。
- 所有块必须带 `verification_status`，缺省请填 `unverified`。页面必须带 `render_dpi`。

## 1. 根对象与 document

| 中间格式字段 | 填写建议 |
| --- | --- |
| `schema_version` | 固定 `"1.0.0"`，并与 `document.schema_version` 保持一致 |
| `data_label` | 真实识别填 `production`；示例 `example`；构造测试 `constructed` |
| `document.source_file` | 原始 PDF/图片文件名或相对路径 |
| `document.source_type` | 数字 PDF 为 `pdf`；扫描 PDF 为 `scan`；单页图为 `image`；混合为 `mixed` |
| `document.page_count` | 必须等于 `pages` 长度，且等于最大 `page_number` |
| `document.generated_at` | ISO 8601，建议带时区 |
| `document.engine_name` / `engine_version` | 流水线名称与版本，例如 `paddleocr-pp-structure` / `2.x.x` |
| `document.render_dpi_default` | 与各页 `render_dpi` 对齐，推荐 `300` |
| `coordinate_system` | `origin=top-left`，`x_axis=rightward`，`y_axis=downward`，`default_unit` 推荐 `pt` |
| `needs_review` | 任一块为 `needs_review` 或低置信度时建议为 `true` |

## 2. pages[]

| 中间格式 | PyMuPDF | PaddleOCR / 渲染 |
| --- | --- | --- |
| `page_number` | `page.number + 1` | 与源 PDF 页序一致（从 1 开始）。摘页保留原始页码，不必重编为 1..N |
| `width` / `height` / `unit` | `page.rect.width/height`，单位 `pt` | 若只用渲染图尺寸，则 `unit=px`，数值为像素宽高 |
| `render_dpi` **必填** | `zoom * 72`；约定 `page.get_pixmap(matrix=fitz.Matrix(zoom,zoom))` 且 `zoom≥3` | 送入检测/识别的图像 DPI，推荐 300 |
| `rotation` | `page.rotation`（0/90/180/270） | bbox 一律给旋转后的视觉页面坐标 |
| `skew_angle_deg` | 可选，纠偏**后残留角**（已纠偏则填 `0`） | 识别侧在渲染后、版面分析前纠偏；生成侧不再次旋转 |
| `skew_angle_deg_before` | 可选，纠偏**前**倾角（度，逆时针为正） | 只在此字段记录倾角，不要写入 `deskew_transform.angle_deg` |
| `deskew_applied` | 可选 | `true` 时 bbox 已在纠偏后坐标，生成侧不旋转 |
| `deskew_transform` | 可选 | 见下方约定 |

**纠偏角度约定（双方已确认，schema 1.0.0 描述级约定）：**

- 坐标系：页面左上原点，y 向下；角度逆时针为正。
- `skew_angle_deg_before`：纠偏前倾角。
- `skew_angle_deg`：纠偏后残留角。
- `deskew_transform.angle_deg`：**矩阵实际旋转角 θ**，不是纠偏前倾角。转正时常为 `θ = -skew_angle_deg_before`。
- `deskew_transform.matrix`：2×3 矩阵 `[[cosθ,-sinθ,tx],[sinθ,cosθ,ty]]`，JSON 打包为 `[a,b,c,d,tx,ty]`（即 `[[a,b,tx],[c,d,ty]]`）。
- 语义校验：`atan2(matrix[1][0], matrix[0][0])`（即 `atan2(c, a)`）与 `angle_deg` 偏差应 ≤ 0.01°；超出则 warning，不作为 schema 失败。
| `media_box` | `page.mediabox` 转换为左上原点（如需要） | 可省略 |

注意：部分真实扫描 PDF 的页面尺寸并非标准 A4（甚至以「1 px = 1 pt」的大画布保存）。`pages` 应如实记录源页面；生成侧会把输出规范到 A4，不要求识别侧改写源尺寸。

## 3. blocks[] 通用字段

| 字段 | 必填 | 映射 |
| --- | --- | --- |
| `id` | 是 | 文档内唯一。建议 `b-{页码:03d}-{序号:03d}` |
| `page` | 是 | 页码，从 1 起 |
| `bbox` | 是 | `[x0,y0,x1,y1]`，须 `x1>x0` 且 `y1>y0` |
| `bbox_unit` / `bbox_dpi` | 像素时必填 dpi | Paddle 检测框多为渲染图像素；PyMuPDF 文本块已是 pt |
| `type` | 是 | 见下节枚举映射 |
| `level` | `heading` 时必填 | 1–6。PP-Structure 常只有 `title`，需按字号/编号规则推断 |
| `reading_order` | 是 | 先左栏后右栏、先上后下。可用版面阅读序或按 `y0,x0` 排序后编号 |
| `text` | 建议 | OCR / VLM 文本。图片、印章可为空串 |
| `confidence` | 建议 | `[0,1]`。多片段取平均或最小值，并在 `warnings` 说明 |
| `source` | 建议 | 普通文本 `ocr`；表格/印章/手写/混排 `vlm` 或 `mixed`；后处理 `rule` |
| `verification_status` | **必填** | 识别侧默认 `unverified`；自动质检通过 `auto_passed`；低置信或冲突 `needs_review`；人工通过 `human_verified`；确认错误 `rejected`（生成侧将跳过） |
| `style` | 可选 | 可从 PyMuPDF span 的 `font/size/flags` 或 VLM 描述填写 |
| `needs_review` | 可选 | 与 `verification_status=needs_review` 对齐 |

页眉、页脚、页码块仍要输出（不要丢），但不要把它们编进正文 `reading_order` 的中间；建议单独用较小或较大的序号，并正确设置 `type`。

## 4. 版面类型映射

### 4.1 PaddleOCR PP-Structure（2.x layout）

| PP-Structure `type` | 中间格式 `type` | 备注 |
| --- | --- | --- |
| `title` | `title` 或 `heading` | 封面/文题用 `title`；章节标题用 `heading` 并给 `level` |
| `text` | `paragraph` 或 `list` | 若文本匹配「（1）/1./·」等列表模式，改为 `list` |
| `figure` / `image` | `image` | 必须写 `image.asset_id`，并把裁块写入 `assets` |
| `figure_caption` / `table_caption` | `figure_caption` / `table_caption` | 不要并进表格或图片的 `text` 后丢失结构 |
| `table` | `table` | 必须结构化为 `table.cells`（见第 5 节），不要只给 HTML |
| `header` | `header` | |
| `footer` | `footer` | 若仅为页码，可用 `page_number` |
| `equation` / `formula` | `formula` | 文本公式放 `text`；公式图放 `image` |
| `reference` | `paragraph` 或 `list` | 可在 `warnings` 标注 reference |
| `caption` | 视位置判定图题或表题 | |

### 4.2 PaddleX / PP-Structure 近年布局类别

| 检测类别（常见名） | 中间格式 `type` |
| --- | --- |
| `paragraph_title` | `heading`（推断 `level`） |
| `text` / `content` | `paragraph` |
| `image` / `chart` / `figure` | `image`（`assets[].type` 可为 `drawing`） |
| `table` | `table` |
| `table_title` / `chart_title` / `figure_title` | `table_caption` / `figure_caption` |
| `header` / `footer` / `number` / `page_number` | `header` / `footer` / `page_number` |
| `formula` / `equation` | `formula` |
| `seal` / `stamp` | `seal` |
| `handwriting` / `handwritten` | `handwriting` |
| `footnote` | `paragraph`（`warnings` 标注 footnote） |
| 无法归类 | `unknown` |

印章、手写、复杂表、图文混排约定由视觉模型产出，`source` 填 `vlm` 或 `mixed`。

### 4.3 PyMuPDF

| 来源 | 映射 |
| --- | --- |
| `get_text("dict")` 中 `type=0` 的 block | 合并同一段落的 line/span，生成 `paragraph` 或 `heading` |
| span `size` + `flags`（bold）+ 位置 | 推断 `heading.level` 或写入 `style` |
| `type=1` 图像块或 `page.get_images()` | `image` / `seal`（若 VLM/规则判定为章） |
| `page.find_tables()` 或表格模型 | `table` |
| 位于页面上 8% / 下 8% 且重复出现的短文本 | `header` / `footer` / `page_number` |
| 字体名含 `SimSun`/`宋体` 等 | `style.font_family` |

PyMuPDF 文本框已是 pt、左上原点，直接作为 `bbox`，`bbox_unit="pt"`，不必填 `bbox_dpi`。

由像素框转到 pt 时，必须使用**同一页**的 `render_dpi`，并考虑 `rotation`。不要混用 72 DPI 页面坐标与 300 DPI 检测框。

## 5. 表格结构

中间格式要求 `table.n_rows`、`table.n_cols`、`table.cells[]`。合并区域只输出左上角锚点单元格，使用 `rowspan`/`colspan`，**不要再输出被覆盖的格子**。

从 PP-Structure 表格 HTML 映射：

- 解析 `<table><tr><td rowspan colspan>`；行、列从 0 计数。
- 单元格文本取 `td` 纯文本，保留换行意义时可把 `<br>` 转成 `\n`。
- 若同时给出 `cell_bbox` / 单元格四点坐标，写入 `cells[].bbox`（单位与所属 table 块一致）。
- `<th>` 行计入 `header_row_count`（缺省 1）。跨页表必须填写准确的表头行数，供生成侧重复表头。
- 不要只交 HTML 字符串；生成侧以 `cells` 为准。HTML 如需保留可放在块级 `warnings` 或扩展字段（schema 允许块上的额外字段）。

从 PyMuPDF `Table` 映射：`rows/cols`、`header_row_count`，每个 cell 的 `text` 与 `bbox`。

## 6. 图片、印章与 assets[]

裁出的每个图像资源都必须进入 `assets`：

| 字段 | 说明 |
| --- | --- |
| `id` | 唯一，例如 `a-001` |
| `type` | `image` 插图；`scan` 扫描裁块（参与倾斜检查）；`seal` 印章；`drawing` 线框图；`signature` 签字图 |
| `path` | **相对本 JSON 文件所在目录** 的路径，例如 `assets/a-001.png` |
| `hash_sha256` | 文件 SHA-256 十六进制 |
| `width_px` / `height_px` | 资源像素尺寸 |
| `original_dpi` | 固有分辨率。从 300 DPI 渲染页裁出则填 300；扫描件按其像素/页面英寸计算。未知时填渲染 DPI 并警告 |
| `crop_bbox` | 可选，相对更大原图的像素框 |
| `source_page` | 可选 |

块内 `image.asset_id` 指向 `assets[].id`。`image.crop_bbox` 是对该资源再裁一次（像素坐标，左上原点）。印章用 `type=seal` 的块 + `assets[].type=seal`。

**不要**把身份证等个人敏感图像写入将随测试报告分发的示例；若识别流水线遇到此类图像，应在 `warnings` 中标记且不要把号码、姓名等原文写入 `text`。

## 7. 标题层级推断（识别侧待对齐）

PP-Structure 常常只给出 `title`，没有 1–6 级。建议规则（`source=rule`）：

1. 文本匹配「第[一二三四五六七八九十百零0-9]+章」或「^[0-9]+(\s|、)」→ `level=1`
2. 「^[0-9]+\.[0-9]+」→ `level=2`，「再多一节数字」→ `level=3`，以此类推
3. 否则用字号分档：与正文中位数相比，≥16pt 倾向 1 级，14–15.5pt 倾向 2 级，接近正文但加粗居中倾向 3 级
4. 无法判定时：`type=heading`、`level=1`、`verification_status=needs_review`

## 8. 阅读顺序与多栏

推荐：版面模型给出的 reading order；若无，则按栏（x 聚类）再按 `y0,x0` 排序。双栏页面不要整页按纯 y 排序。页眉/页脚/页码不参与正文阅读序。

## 9. 最小可运行示例（字段级）

识别侧对一页数字 PDF 上的一段正文，用 PyMuPDF 文本框时可输出：

```json
{
  "id": "b-001-004",
  "page": 1,
  "bbox": [90.0, 160.0, 505.0, 220.0],
  "bbox_unit": "pt",
  "type": "paragraph",
  "reading_order": 4,
  "text": "……",
  "confidence": 0.97,
  "source": "ocr",
  "verification_status": "unverified",
  "style": {"font_family": "宋体", "font_size_pt": 14, "alignment": "justify"}
}
```

同一段若来自 300 DPI 渲染图上的 Paddle 检测框：

```json
{
  "bbox": [375, 667, 2104, 917],
  "bbox_unit": "px",
  "bbox_dpi": 300
}
```

`375 px * 72 / 300 = 90 pt`，与上例一致。

## 10. 生成侧忽略与跳过规则

- `verification_status=rejected` 的块：不写入正文，只记日志。
- `unknown`：有 `image` 按图处理，否则按段落。
- 额外字段会被忽略，不视为错误，便于识别侧试验。
- `assets.path` 必须是相对路径且文件存在，否则生成失败。
