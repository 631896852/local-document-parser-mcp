# Local Document Parser MCP

<p align="center">
  <strong>让本地文档直接成为 AI 可调用的结构化知识。</strong><br>
  Word、PDF、Excel 一套接口；章节、表格、检索、OCR 一次打通。
</p>

<p align="center">
  <a href="README_EN.md">English</a> · <a href="#5-分钟开始使用">快速开始</a> · <a href="docs/local_mcp_server_usage.md">完整文档</a>
</p>

<p align="center">
  <img alt="Python 3.10+" src="https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white">
  <img alt="MCP" src="https://img.shields.io/badge/MCP-stdio-6C47FF">
  <img alt="Local first" src="https://img.shields.io/badge/Data-local--first-18A558">
  <img alt="License" src="https://img.shields.io/badge/License-MIT-yellow.svg">
</p>

Local Document Parser MCP 是一个本地优先的文档解析器，也是一台可直接接入 AI 客户端的 stdio MCP Server。传入文件路径，就能读取文档总览、定位章节、检索自然语言问题、抽取表格，并处理扫描 PDF。

它解决的是一个很实际的问题：**不上传文件，不先搭建向量数据库，也能让模型可靠地读懂手边的报告、合同、方案和表格。**

## 为什么实用

| 你面对的文件 | 它能做什么 |
| --- | --- |
| 长篇 Word 报告 | 识别章节层级，按章节摘要或提取正文，保留合并单元格 |
| 原生或复杂 PDF | 判断 PDF 类型，检测多栏、表格页和编码异常，输出结构化 Markdown |
| 扫描 PDF | 自动定位需要 OCR 的页面，在 macOS 上调用本机 Vision OCR |
| Excel 台账 | 按工作表、行关键词和列关键词精确找到单元格 |
| 模糊问题 | 用 `queryContent` 从章节、段落和表格中返回相关上下文 |
| AI 客户端 | 通过 stdio MCP 暴露 11 个工具，直接传入本地 `docPath` |

### 核心特点

- **本地优先**：解析过程在本机完成；macOS OCR 使用系统 Vision 框架。
- **一个入口，多种格式**：支持 `.docx`、`.doc`、`.pdf`、`.xlsx`、`.xlsm`、`.xls`。
- **不是只“转文字”**：保留章节、行号、表格标题、工作表和单元格定位信息。
- **对扫描件有准备**：自动判断哪些页面需要 OCR，并把识别结果合并回章节结构。
- **故障可降级**：PDF 增强引擎不可用时自动回退到 PyMuPDF，不让一次解析整体失败。
- **方便接入现有工作流**：既能作为 Python 模块使用，也能作为 MCP Server 使用。

## 5 分钟开始使用

### 1. 安装

```bash
git clone https://github.com/631896852/local-document-parser-mcp.git
cd local-document-parser-mcp
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Python 需要 3.10 或更高版本。解析旧版 `.doc`、`.xls` 时还需要 [LibreOffice](https://www.libreoffice.org/)：

```bash
brew install --cask libreoffice
```

### 2. 直接解析文件

```bash
python examples/local_parse_docx.py "/absolute/path/to/report.pdf" --chapter 11
```

这个命令会先输出文档总览，再输出指定章节摘要。Word、PDF 和 Excel 都使用同一个入口。

### 3. 启动 MCP Server

```bash
python scripts/local_mcp_server.py
```

在 Cherry Studio 或其他支持 stdio MCP 的客户端中加入：

```json
{
  "mcpServers": {
    "local-document-parser": {
      "command": "/absolute/path/to/local-document-parser-mcp/.venv/bin/python",
      "args": [
        "/absolute/path/to/local-document-parser-mcp/scripts/local_mcp_server.py"
      ]
    }
  }
}
```

随后可以直接对模型说：

> 解析 `/absolute/path/to/report.pdf`，找到“投资估算”相关内容，并给出所在章节和上下文。

## MCP 工具

| 工具 | 用途 |
| --- | --- |
| `extractFileSummaryText` | 获取封面、目录、章节和字符统计 |
| `extractChapterSummary` | 快速查看一个或多个章节摘要 |
| `extractChapterContent` | 提取指定章节正文 |
| `readLines` | 按全局行号读取内容 |
| `searchContent` | 用明确关键词检索 |
| `queryContent` | 用自然语言问题检索相关上下文 |
| `extractTableList` | 列出章节内的表格 |
| `extractTableContent` | 提取表格；Excel 可按行列关键词精确取数 |
| `inspectPdf` | 判断 PDF 类型、复杂版面及需要 OCR 的页面 |
| `extractPdfMarkdown` | 输出合并原生文字与 OCR 的 Markdown |
| `ocrPdf` | 返回扫描页的逐页 OCR 文本 |

调用参数以本地绝对路径 `docPath` 为主，例如：

```json
{
  "docPath": "/absolute/path/to/report.pdf",
  "query": "项目建设的主要风险是什么？",
  "pageSize": 5
}
```

## PDF 与 OCR

PDF 处理由三层组成：

1. `pdf-inspector` 判断原生文字、扫描、图片型或混合 PDF，并生成结构化 Markdown。
2. PyMuPDF 提取文字、章节和表格，同时承担降级处理。
3. macOS Vision OCR 只识别确实需要 OCR 的页面，减少无效计算。

OCR 支持简体中文、繁体中文和英文，默认以 200 DPI 渲染。可通过环境变量调整：

```bash
LOCAL_PDF_OCR=off          # 关闭自动 OCR
LOCAL_PDF_OCR_DPI=200      # 120~300，数值越大通常越清晰也越慢
```

非 macOS 系统仍可使用 Word、Excel 和原生文字 PDF 解析；Vision OCR 会自动跳过。

## Python 调用

```python
from src.local_docx_parser import LocalDocxParseService

parser = LocalDocxParseService("/absolute/path/to/report.docx")

print(parser.extract_file_summary_text())
print(parser.query_content("项目总投资是多少？", page_size=5))
print(parser.extract_chapter_content("3.2"))
```

## 验证

```bash
python -m unittest discover -s tests
```

当前测试覆盖 Word、PDF、Excel、本地 OCR 适配层和 MCP 协议处理。

## 已知边界

- `.doc` 和 `.xls` 依赖 LibreOffice 做临时格式转换。
- macOS Vision OCR 需要本机可用的 `clang`；其他系统不会执行这部分 OCR。
- PDF 跨页表格、流程图、复杂图表和照片语义暂不能完整还原。
- Excel 读取工作簿中已保存的公式缓存值，不执行宏，也不重新计算公式。

## 更多文档

- [MCP Server 使用说明](docs/local_mcp_server_usage.md)
- [解析能力与边界](docs/local_parser_capabilities.md)
- [Python 本地调用说明](docs/local_service_usage.md)

## 参与贡献

欢迎提交 Issue 和 Pull Request。特别期待这些方向：跨平台 OCR、更稳健的跨页表格还原、更多真实版式测试，以及英文文档的章节识别。

## 许可证

[MIT License](LICENSE)
