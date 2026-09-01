# 本地解析器能力说明

本文说明 `src/local_docx_parser.py` 当前支持的本地解析能力和已知边界。

## 支持格式

| 格式 | 支持情况 | 说明 |
| --- | --- | --- |
| `.docx` | 支持 | 使用 `python-docx` 解析正文、章节和表格。 |
| `.doc` | 支持 | 自动调用 LibreOffice 转为临时 `.docx` 后解析。 |
| `.pdf` | 增强支持 | pdf-inspector 负责类型/版面检测和 Markdown，macOS Vision 负责扫描页 OCR，PyMuPDF 负责中文章节、表格辅助和降级。 |
| `.xlsx` | 支持 | 使用 openpyxl 解析工作表、合并单元格和搜索内容。 |
| `.xlsm` | 支持 | 使用 openpyxl 解析，宏不执行。 |
| `.xls` | 支持 | 自动调用 LibreOffice 转为临时 `.xlsx` 后解析。 |

## Word 解析

Word 文档重点对齐远端工具的输出形态：

- 章节标题识别。
- 章节摘要递归返回子章节。
- 行读取。
- 关键词搜索。
- 表格标题列表。
- 表格内容提取。
- 合并单元格尽量输出为 HTML 表格，保留 `colspan`、部分 `rowspan`。

`.doc` 文件会转换后再走 Word 解析逻辑。转换缓存目录位于系统临时目录：

```text
document_parser_doc_cache
```

缓存键包含文件路径、修改时间和文件大小。

旧版 `.doc` 转换后可能出现目录文字丢失为大量空行的情况。本地解析器会识别 `目录` 行，并用已识别章节标题回填目录区，便于 `searchContent` 命中目录中的章节和投资类关键词。

## PDF 解析

PDF 解析优先保证数据提取可用：

- `inspectPdf` 返回 PDF 类型、置信度、复杂版面、表格页、多栏页、需要 OCR 的页面、编码异常和图片位置。
- `ocrPdf` 返回扫描页的逐页 OCR 文字和成功、失败页信息。
- `extractPdfMarkdown` 返回 pdf-inspector 与 OCR 合并后的结构化 Markdown；不可用时降级为 PyMuPDF 文本。
- `extractFileSummaryText` 会尽量分出封面、目录、正文。
- `extractChapterSummary` 可识别常见可研报告的 `1~11` 章结构。
- `extractChapterContent` 可返回章节正文。
- `readLines` 可按本地行号读取。
- `searchContent` 可分页搜索关键词。
- `extractTableList` 可粗提取 `表11-1`、`表2.2-1` 等表格标题。
- `extractTableContent` 可按表格标题截取后续表格区域。

PDF 表格会合并 PyMuPDF `find_tables()` 和 pdf-inspector Markdown 表格结果。现有中文表格标题及章节归属规则保持不变。

PDF 增强处理流程：

1. pdf-inspector 预检并生成结构化 Markdown。
2. PyMuPDF 提取原有文本行和表格。
3. 对扫描、图片型、混合 PDF 的必要页面调用本机 macOS Vision OCR。
4. OCR 标题经过换行修复后继续运行原有中文章节规则，因此章节摘要、章节正文和检索接口仍可使用。
5. pdf-inspector 识别到的 Markdown 表格用于补充 PyMuPDF 结果；增强引擎不可用时自动降级。

## Excel 解析

Excel 文档重点对齐远端工具的工作表级行为：

- `extractFileSummaryText` 列出 sheet 列表。
- sheet 名称作为章节名称。
- `extractChapterSummary` 返回字符数、表格字符数、行范围、表格摘要和前两列预览。
- `extractChapterContent` 返回 sheet 标记和工作表内容。
- `readLines` 可按本地行号读取。
- `searchContent` 可分页搜索关键词。
- 合并单元格输出为 HTML `colspan`、`rowspan`。
- 空单元格保留。

当前 Excel 的 `extractTableList` 和 `extractTableContent` 保持远端样本行为：不按表格标题建立索引。需要读取 Excel 内容时，推荐使用工作表名称调用 `extractChapterContent`。

## 与远端的主要差异

Word 文档：

- 章节和表格结构已尽量接近远端。
- 行号可能仍有小幅差异，主要来自空行、目录和合并单元格处理。

PDF 文档：

- 行号和字符数不保证与远端一致。
- 远端会保留更多 `<image>`、`<sub>`、`<sup>` 和版式信息。
- 本地表格标题提取在部分 PDF 上可能比远端更可用，但表格二维结构不完全还原。

## 推荐调用顺序

1. `extractFileSummaryText` 获取目录和章节树。
2. `extractChapterSummary` 预览目标章节。
3. `searchContent` 搜索关键词，例如 `投资估算`、`建设投资`、`新增产能`。
4. `extractChapterContent` 或 `readLines` 获取上下文。
5. `extractTableList` 和 `extractTableContent` 读取表格区域。

## 已知边界

- OCR 当前依赖 macOS 的 Vision、AppKit 和本机 `clang`；非 macOS 环境会保留原生文字/PyMuPDF 降级结果。
- OCR 默认 200 DPI，可通过 `LOCAL_PDF_OCR_DPI=120~300` 调整；可通过 `LOCAL_PDF_OCR=off` 关闭。
- PDF 表格跨页时，可能需要通过章节正文或行读取补充上下文。
- `inspectPdf` 能返回图片的位置和尺寸；OCR 可识别图片中的文字，但暂不理解图表、流程图和照片语义。
- 英文 PDF 只做基础标题支持，本项目主要面向中文可研报告。
- Excel 公式读取的是工作簿缓存值；如果文件没有保存过计算结果，公式值可能为空或保留公式文本。
- Excel 暂不执行宏，也不计算公式。
