# 本地 MCP Server 使用说明

本地 MCP Server 暴露远端 `file_parse_service` 兼容工具，并新增 PDF 检查、OCR 与 Markdown 工具。参数以 `docPath` 为主，直接读取本地 `.docx`、`.doc`、`.pdf`、`.xlsx`、`.xlsm`、`.xls` 文件。

## 启动方式

```bash
.venv/bin/python scripts/local_mcp_server.py
```

工具调用时传入本地绝对路径：

```json
{
  "docPath": "/absolute/path/to/example.pdf"
}
```

`.docx`、`.pdf`、`.xlsx`、`.xlsm` 可直接解析。`.doc` 会自动转换成临时 `.docx` 后解析，`.xls` 会自动转换成临时 `.xlsx` 后解析；解析 `.doc` 或 `.xls` 前请安装 LibreOffice：

```bash
brew install --cask libreoffice
```

## 预注册文档

预注册不是必需的。它只用于提前登记常用文件，首次调用时仍以 `docPath` 定位文档。

```bash
python3 scripts/local_mcp_server.py \
  --doc "/absolute/path/to/report.docx" \
  --doc "/absolute/path/to/report.pdf" \
  --doc "/absolute/path/to/template.xlsx"
```

## Cherry Studio 配置

```json
{
  "mcpServers": {
    "local_file_parse_service": {
      "command": "/absolute/path/to/local-document-parser-mcp/.venv/bin/python",
      "args": [
        "/absolute/path/to/local-document-parser-mcp/scripts/local_mcp_server.py"
      ]
    }
  }
}
```

如果希望预注册文档：

```json
{
  "mcpServers": {
    "local_file_parse_service": {
      "command": "/absolute/path/to/local-document-parser-mcp/.venv/bin/python",
      "args": [
        "/absolute/path/to/local-document-parser-mcp/scripts/local_mcp_server.py",
        "--doc",
        "/absolute/path/to/report.docx",
        "--doc",
        "/absolute/path/to/report.pdf",
        "--doc",
        "/absolute/path/to/template.xlsx"
      ]
    }
  }
}
```

## 已支持工具

```text
inspectPdf
extractPdfMarkdown
ocrPdf
extractFileSummaryText
extractChapterSummary
extractChapterContent
readLines
searchContent
queryContent
extractTableList
extractTableContent
```

## 参数说明

所有工具都支持兼容字段 `fileId`，但本地推荐使用 `docPath`。

| 工具 | 必填参数 |
| --- | --- |
| `inspectPdf` | `docPath`；可选 `includeImages` |
| `extractPdfMarkdown` | `docPath` |
| `ocrPdf` | `docPath` |
| `extractFileSummaryText` | `docPath` |
| `extractChapterSummary` | `docPath`, `chapters` |
| `extractChapterContent` | `docPath`, `chapters` |
| `readLines` | `docPath`, `startLine`, `endLine` |
| `searchContent` | `docPath`, `keywords` |
| `queryContent` | `docPath`, `query` |
| `extractTableList` | `docPath`, `chapters` |
| `extractTableContent` | `docPath`, `tableTitle` |

`searchContent` 可选：

```text
pageNo
pageSize
```

`extractTableContent` 可选：

```text
rowScope
rowKeywords
colKeywords
```

## 本地验证

列出工具：

```bash
printf '%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{}}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/list"}' \
  | python3 scripts/local_mcp_server.py
```

搜索 PDF：

```bash
printf '%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{}}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"searchContent","arguments":{"docPath":"/absolute/path/to/example.pdf","keywords":"投资估算","pageNo":1,"pageSize":5}}}' \
  | python3 scripts/local_mcp_server.py
```

提取章节摘要：

```bash
printf '%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{}}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"extractChapterSummary","arguments":{"docPath":"/absolute/path/to/example.pdf","chapters":"11"}}}' \
  | python3 scripts/local_mcp_server.py
```

提取 Excel 工作表：

```bash
printf '%s\n' \
  '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{}}' \
  '{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"extractChapterContent","arguments":{"docPath":"/absolute/path/to/template.xlsx","chapters":"项目基本信息"}}}' \
  | python3 scripts/local_mcp_server.py
```

## PDF 说明

PDF 解析已支持：

- 检测 `text_based`、`scanned`、`image_based`、`mixed`。
- 返回置信度、需要 OCR 的页面及原因、编码异常、表格页和多栏页。
- 扫描页自动使用本机 macOS Vision OCR，支持简体中文、繁体中文和英文，不上传文件。
- `ocrPdf` 返回逐页 OCR 文字；`extractPdfMarkdown` 输出 pdf-inspector 与 OCR 合并后的结构化 Markdown。
- 返回图片对象的位置和尺寸，不导出图片文件。
- 总览中区分封面、目录、正文。
- 章节标题识别。
- 行读取和关键词搜索。
- 表格标题粗提取。
- 按表格标题截取表格附近内容。

PDF 行号和字符数不保证与远端完全一致。不同 PDF 解析引擎对换行、图片、上下标和表格单元格的处理会有差异。

OCR 默认以 200 DPI 渲染。可用 `LOCAL_PDF_OCR=off` 关闭，或用 `LOCAL_PDF_OCR_DPI=120~300` 调整清晰度与速度。

## Excel 说明

Excel 解析已支持：

- 总览中列出工作表。
- 工作表名称作为章节名称。
- 按工作表提取摘要和正文。
- 行读取和关键词搜索。
- 合并单元格输出为 HTML `colspan`、`rowspan`。
- 空单元格保留为 `<td></td>`。

远端样本中 `extractTableList` 和 `extractTableContent` 对 Excel 表格标题不可用，因此本地第一版保持兼容：表格列表返回空，表格标题提取返回未命中提示。需要读取 Excel 内容时，优先使用 `extractChapterContent(sheetName)` 或 `searchContent`。
