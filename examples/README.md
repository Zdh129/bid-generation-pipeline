# 示例数据

全部为虚构数据，用于公开测试，不是任何真实招标项目的摘录。

| 文件 | 内容 |
| --- | --- |
| `intermediate.json` | 中间格式示例（含构造图片） |
| `devices.json` | 3 台示例设备及参数 |
| `tender_params.json` | 招标参数摘录（虚构） |
| `deviation_judgements.json` | 偏离判定示例（含正偏离、负偏离、待确认、未提供不计入） |
| `template.md` | 完整 `build_bid` 所需模版的说明 |
| `assets/` | 构造用图，非扫描件 |

## 一键演示

在仓库根目录、已安装依赖的虚拟环境中：

```bash
mkdir -p examples/output
python tools/validate_schema.py examples/intermediate.json
python tools/generate_docx.py examples/intermediate.json --docx examples/output/example.docx
python tools/check_layout.py --docx examples/output/example.docx -o examples/output/check_report.json
```

导出 PDF（需要 LibreOffice）：

```bash
python tools/generate_docx.py examples/intermediate.json   --docx examples/output/example.docx   --pdf examples/output/example.pdf
python tools/check_layout.py   --docx examples/output/example.docx   --pdf examples/output/example.pdf   -o examples/output/check_report.json
```

在测试机上用本示例生成了 DOCX 和 PDF（LibreOffice），57 项检查结果为：失败 0，不适用 12。小样页数少，部分项不适用是预期行为。
