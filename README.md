# 标书生成流水线（公开测试版）

从中间格式 JSON 生成投标文件（DOCX，可选 PDF），并按 57 项排版质量标准做检查。本仓库是脱敏后的公开测试版，**不含**真实招标原文、设备清单、标书成品或个人信息。

完整的项目说明、规则叠加层设计和各轮结论见 [HANDOVER.md](HANDOVER.md)。

## 能做什么

- 校验中间格式（`schema/intermediate.schema.json`）
- 由中间格式生成符合排版标准的 DOCX，可用 LibreOffice 导出 PDF
- 按 57 项标准检查 DOCX / PDF
- 查阅偏离判定叠加层（`deviation_judgements` + overlay，可追溯）的示例与历史脚本

完整的「按招标模版灌装 45 类设备、嵌入证明材料、回填页码」流水线脚本也在 `tools/` 中（如 `build_bid_r10.py`），但运行它们需要你自己的招标资料和 Word 模版，本仓库不提供这些文件。请先阅读 `examples/template.md` 和 `config/*.example.json`。

## 环境

建议使用 Python 3.11+ 和虚拟环境：

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

还需要：

| 组件 | 用途 | 是否必需 |
| --- | --- | --- |
| LibreOffice（`soffice`） | DOCX 转 PDF | 仅在需要 PDF 或 PDF 侧检查时 |
| poppler（`pdfinfo` 等） | PDF 文本与页面信息 | PDF 检查 |
| pypdfium2 / pikepdf | PDF 渲染、原图回写 | 已写入 requirements，完整流水线需要 |
| RapidOCR（`rapidocr-onnxruntime`） | 本机离线识别扫描件 | 可选；默认不安装。只下载模型，不上传数据 |

本流水线**不调用**外部大模型或云识别接口。需要大模型复核时，应先做费用估计并获得批准，再在隔离环境运行。

## 快速开始（示例）

```bash
mkdir -p examples/output
python tools/validate_schema.py examples/intermediate.json
python tools/generate_docx.py examples/intermediate.json --docx examples/output/example.docx
python tools/check_layout.py --docx examples/output/example.docx -o examples/output/check_report.json
```

导出 PDF：

```bash
python tools/generate_docx.py examples/intermediate.json \
  --docx examples/output/example.docx \
  --pdf examples/output/example.pdf
python tools/check_layout.py \
  --docx examples/output/example.docx \
  --pdf examples/output/example.pdf \
  -o examples/output/check_report.json
```

示例只有少量页面，检查结果里会出现「不适用」项，这是预期行为。

## 目录

| 路径 | 说明 |
| --- | --- |
| `schema/` | 中间格式 JSON Schema、识别侧映射说明、schema 示例 |
| `standards/` | 57 项排版质量标准与 `checks.json` |
| `tools/` | 生成、检查、偏离判定与各轮叠加层脚本 |
| `examples/` | 虚构的小示例（设备、参数、偏离判定、中间格式） |
| `config/` | 路径与项目信息的示例配置 |
| `HANDOVER.md` | 交接文档 |

## 安全约定

- 不要把真实招标文件、对照成品、证件扫描件或含个人信息的数据推送到本仓库。
- 本机路径写在 `config/paths.json` / `config/project.json`（已列入 `.gitignore`）。
- 身份证件等个人材料只在本机处理，不要外发，也不要调用外部 API。
