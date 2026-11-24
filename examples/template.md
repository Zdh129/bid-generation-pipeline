# 示例模板说明

本仓库**不包含**真实的 Word 标书模版或对照成品。

完整流水线（`tools/build_bid_r10.py`）需要你在本机准备：

1. 招标文件与附件参数表（文件夹 1 一类的官方文件）；
2. 制作标书用的 Word 模版（文件夹 2）；
3. 资格证明、厂家参数确认函、彩页等附件素材（仅作嵌入材料）；
4. 可选的对照成品，用于与生成稿做一致性比较。

把本机路径写入 `config/paths.json`（先复制 `config/paths.example.json`），项目名称、编号、招标人等写入 `config/project.json`。

公开仓库提供的可运行示例是另一条更短的路径，不依赖真实模版：

- 输入：`examples/intermediate.json`（中间格式，虚构数据）
- 生成：`tools/generate_docx.py` → `examples/output/example.docx`
- 检查：`tools/check_layout.py`（57 项排版标准）

若要导出 PDF 并做 PDF 侧检查，需本机安装 LibreOffice（`soffice`）。详见仓库根目录 `README.md` 与 `HANDOVER.md`。
