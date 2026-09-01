# 本地统一调用层使用说明

`src/file_parse_service.py` 提供与远端 `FileParseMCPClient` 接近的方法形态，适合在 Python 代码中直接调用本地解析器。它保留 `file_id` 入参，以兼容旧调用方式；底层实际读取的是初始化时登记的本地路径。

如果要给模型或 Cherry Studio 使用，请优先使用本地 MCP Server，并通过 `docPath` 传文件路径。

## 支持格式

```text
.docx
.doc
.pdf
.xlsx
.xlsm
.xls
```

`.doc` 和 `.xls` 需要安装 LibreOffice，解析时会自动转换成临时 `.docx` 或 `.xlsx`。

## 示例

```python
from pathlib import Path
from src.file_parse_service import LocalFileParseService

service = LocalFileParseService(
    {
        "report": Path("/absolute/path/to/report.pdf"),
    }
)

summary = service.extract_file_summary_text("report")
chapter = service.extract_chapter_summary("report", "11")
search = service.search_content("report", "投资估算", page_no=1, page_size=10)
```

命令行示例：

```bash
python3 examples/local_service_call.py --chapter 1.1.3 --keywords 投资估算
```

## 已支持方法

```text
extract_file_summary_text(file_id)
extract_chapter_summary(file_id, chapters)
extract_chapter_content(file_id, chapters)
read_lines(file_id, start_line, end_line)
search_content(file_id, keywords, page_no=None, page_size=None)
extract_table_list(file_id, chapters)
extract_table_content(file_id, table_title, row_keywords=None, col_keywords=None, row_scope=None)
```

## 返回结构

所有方法返回统一的 `ParsedToolResult`：

```text
ok
kind
text
raw_text
meta
is_error
```

常见 `kind`：

```text
result
tip
tip_result
error
text
```

缺失文件会返回业务错误：

```text
ok=False
kind="error"
text="文件不存在"
```

## 与 MCP Server 的区别

- 本地统一调用层：Python 代码内直接调用，仍使用 `file_id` 映射本地路径。
- 本地 MCP Server：给模型或 MCP Client 调用，推荐直接使用 `docPath`。
