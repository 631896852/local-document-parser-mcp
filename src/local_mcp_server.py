from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from src.file_parse_mcp_client import ParsedToolResult
from src.file_parse_service import LocalFileParseService
from src.local_docx_parser import LocalDocxParseService


PROTOCOL_VERSION = "2025-03-26"


ToolHandler = Callable[[dict[str, Any]], str]


@dataclass(frozen=True)
class ToolSpec:
    name: str
    description: str
    input_schema: dict[str, Any]


TOOLS: list[ToolSpec] = [
    ToolSpec(
        name="inspectPdf",
        description="检查本地 PDF 类型、置信度、复杂版面、表格页、多栏页和图片位置；扫描页会自动使用本机 OCR，并返回识别状态",
        input_schema={
            "type": "object",
            "additionalProperties": False,
            "required": ["docPath"],
            "properties": {
                "docPath": {"type": "string", "description": "本地 PDF 绝对路径"},
                "includeImages": {
                    "type": "boolean",
                    "description": "是否额外提取图片位置，默认为 true",
                },
            },
        },
    ),
    ToolSpec(
        name="extractPdfMarkdown",
        description="将本地 PDF 转换为结构化 Markdown；原生文字使用 pdf-inspector，扫描页自动使用本机 macOS Vision OCR，并保留章节结构",
        input_schema={
            "type": "object",
            "additionalProperties": False,
            "required": ["docPath"],
            "properties": {
                "docPath": {"type": "string", "description": "本地 PDF 绝对路径"},
            },
        },
    ),
    ToolSpec(
        name="ocrPdf",
        description="使用本机 macOS Vision OCR 识别扫描 PDF 中需要 OCR 的页面，支持简体中文和英文，文件不会上传",
        input_schema={
            "type": "object",
            "additionalProperties": False,
            "required": ["docPath"],
            "properties": {
                "docPath": {"type": "string", "description": "本地 PDF 绝对路径"},
            },
        },
    ),
    ToolSpec(
        name="extractTableList",
        description="根据本地文档路径，提取特定章节的表格标题信息列表，返回数据没有表格具体内容，只列举标题信息",
        input_schema={
            "type": "object",
            "additionalProperties": False,
            "required": ["docPath", "chapters"],
            "properties": {
                "docPath": {"type": "string", "description": "本地 doc、docx、pdf、xlsx、xlsm 或 xls 绝对路径"},
                "chapters": {
                    "type": "string",
                    "description": "章节序号，或者章节完整标题，多个章节使用英文','隔开",
                },
                "fileId": {"type": "string", "description": "兼容字段，已废弃"},
            },
        },
    ),
    ToolSpec(
        name="readLines",
        description="根据本地文档路径，起始及终止行号，获取特定区间行内容",
        input_schema={
            "type": "object",
            "additionalProperties": False,
            "required": ["docPath", "startLine", "endLine"],
            "properties": {
                "docPath": {"type": "string", "description": "本地 doc、docx、pdf、xlsx、xlsm 或 xls 绝对路径"},
                "startLine": {"type": "integer", "description": "开始行号"},
                "endLine": {"type": "integer", "description": "结束行号"},
                "fileId": {"type": "string", "description": "兼容字段，已废弃"},
            },
        },
    ),
    ToolSpec(
        name="extractChapterSummary",
        description="根据本地文档路径，章节编号如'5'、'6.2.1'，提取章节概要数据，包含子章节内容、表格列表、总行数等，快速预览章节摘要内容",
        input_schema={
            "type": "object",
            "additionalProperties": False,
            "required": ["docPath", "chapters"],
            "properties": {
                "docPath": {"type": "string", "description": "本地 doc、docx、pdf、xlsx、xlsm 或 xls 绝对路径"},
                "chapters": {
                    "type": "string",
                    "description": "章节编号或名称,多个章节使用','隔开",
                },
                "fileId": {"type": "string", "description": "兼容字段，已废弃"},
            },
        },
    ),
    ToolSpec(
        name="extractFileSummaryText",
        description="根据本地文档路径，提取文档内容总结信息，包括封面、目录、正文章节信息，正文章节包含总字符数、表格字符数，主要用于总览当前文件内容",
        input_schema={
            "type": "object",
            "additionalProperties": False,
            "required": ["docPath"],
            "properties": {
                "docPath": {"type": "string", "description": "本地 doc、docx、pdf、xlsx、xlsm 或 xls 绝对路径"},
                "fileId": {"type": "string", "description": "兼容字段，已废弃"},
            },
        },
    ),
    ToolSpec(
        name="searchContent",
        description="关键词检索工具。仅适合输入很短的明确关键词，如'设计审查费'、'投资计划'。如果用户是在提问、描述需求、询问是否包含某类要求，必须优先使用 queryContent；需要 Excel 单元格精确取数时优先使用 extractTableContent。",
        input_schema={
            "type": "object",
            "additionalProperties": False,
            "required": ["docPath", "keywords"],
            "properties": {
                "docPath": {"type": "string", "description": "本地 doc、docx、pdf、xlsx、xlsm 或 xls 绝对路径"},
                "keywords": {
                    "type": "string",
                    "description": "搜索关键词，如果同时搜索多个，使用英文','隔开",
                },
                "pageNo": {"type": "integer", "description": "当前浏览页码,默认为1"},
                "pageSize": {"type": "integer", "description": "每页展示结果条数,默认为20"},
                "fileId": {"type": "string", "description": "兼容字段，已废弃"},
            },
        },
    ),
    ToolSpec(
        name="queryContent",
        description="自然语言问答和语义检索首选工具。用户提出问题、描述一段需求、问'有没有/是否/怎么/如何/哪些/多少/要求/规定/依据/原因'时必须优先使用本工具。根据本地文档路径和用户问题，检索相关章节、段落和表格片段。适合 doc/docx/pdf/xlsx 的随口提问，如'有没有技术设计方案完整性先进性的要求'、'介绍设计审查费'、'工程监理费怎么算'、'4月份股票持仓是多少'。",
        input_schema={
            "type": "object",
            "additionalProperties": False,
            "required": ["docPath", "query"],
            "properties": {
                "docPath": {"type": "string", "description": "本地 doc、docx、pdf、xlsx、xlsm 或 xls 绝对路径"},
                "query": {"type": "string", "description": "自然语言问题或检索描述"},
                "pageNo": {"type": "integer", "description": "当前浏览页码,默认为1"},
                "pageSize": {"type": "integer", "description": "每页展示结果条数,默认为10"},
                "fileId": {"type": "string", "description": "兼容字段，已废弃"},
            },
        },
    ),
    ToolSpec(
        name="extractChapterContent",
        description='章节内容提取工具。仅当用户明确要求某个章节编号/章节标题时使用，如"1.1"、"3.2.1"。普通问答优先使用 queryContent；Excel 取数优先使用 extractTableContent。',
        input_schema={
            "type": "object",
            "additionalProperties": False,
            "required": ["docPath", "chapters"],
            "properties": {
                "docPath": {"type": "string", "description": "本地 doc、docx、pdf、xlsx、xlsm 或 xls 绝对路径"},
                "chapters": {
                    "type": "string",
                    "description": "章节序号，或者章节完整标题，多个章节使用','隔开",
                },
                "fileId": {"type": "string", "description": "兼容字段，已废弃"},
            },
        },
    ),
    ToolSpec(
        name="extractTableContent",
        description="表格精确取数工具。doc/docx/pdf 用表格标题提取表格；xlsx/xlsm/xls 用 tableTitle 定位 sheet，rowKeywords 定位行，colKeywords 定位列，返回具体单元格。用户问'某月某字段是多少'这类 Excel 问题时使用本工具；普通自然语言问答优先 queryContent。",
        input_schema={
            "type": "object",
            "additionalProperties": False,
            "required": ["docPath", "tableTitle"],
            "properties": {
                "docPath": {"type": "string", "description": "本地 doc、docx、pdf、xlsx、xlsm 或 xls 绝对路径"},
                "tableTitle": {"type": "string", "description": "doc/docx/pdf 为单个表格标题；Excel 为工作表名称/sheet 名，如'总收入'"},
                "rowKeywords": {
                    "type": "string",
                    "description": "行搜索关键词。Excel 中可传月份/日期/行标签，如'4月'、'2026-04-01'；如果只想查某字段整列，也可传字段名如'股票期末余额'",
                },
                "colKeywords": {
                    "type": "string",
                    "description": "列搜索关键词。Excel 中传字段名，如'股票期末余额'、'盈亏合计'；多个关键词使用','隔开",
                },
                "rowScope": {
                    "type": "string",
                    "description": "筛选行范围，只返回前几行与筛选行范围的数据。如获取10-15行，则值为'10-15'，如果需要多个区域，可用','隔开，如：'10~15,30-40'",
                },
                "fileId": {"type": "string", "description": "兼容字段，已废弃"},
            },
        },
    ),
]


class LocalMCPServer:
    def __init__(
        self,
        documents: dict[str, Path] | LocalFileParseService | None = None,
    ) -> None:
        self._legacy_service = documents if isinstance(documents, LocalFileParseService) else None
        self._documents = {
            str(Path(path).expanduser().resolve()): Path(path).expanduser()
            for path in (documents if isinstance(documents, dict) else {}).values()
        }
        self._services: dict[str, LocalDocxParseService] = {}
        self.handlers: dict[str, ToolHandler] = {
            "inspectPdf": self._inspect_pdf,
            "extractPdfMarkdown": self._extract_pdf_markdown,
            "ocrPdf": self._ocr_pdf,
            "extractFileSummaryText": self._extract_file_summary_text,
            "extractChapterSummary": self._extract_chapter_summary,
            "extractChapterContent": self._extract_chapter_content,
            "readLines": self._read_lines,
            "searchContent": self._search_content,
            "queryContent": self._query_content,
            "extractTableList": self._extract_table_list,
            "extractTableContent": self._extract_table_content,
        }

    def handle(self, request: dict[str, Any]) -> dict[str, Any] | None:
        method = request.get("method")
        if method == "notifications/initialized":
            return None
        request_id = request.get("id")
        try:
            if method == "initialize":
                result = {
                    "protocolVersion": PROTOCOL_VERSION,
                    "capabilities": {"tools": {}},
                    "serverInfo": {"name": "local-file-parse-service", "version": "0.3.0"},
                }
            elif method == "tools/list":
                result = {
                    "tools": [
                        {
                            "name": tool.name,
                            "description": tool.description,
                            "inputSchema": tool.input_schema,
                        }
                        for tool in TOOLS
                    ]
                }
            elif method == "tools/call":
                result = self._call_tool(request.get("params", {}))
            else:
                return _error_response(request_id, -32601, f"Method not found: {method}")
            return {"jsonrpc": "2.0", "id": request_id, "result": result}
        except Exception as exc:  # noqa: BLE001 - JSON-RPC should convert all handler errors.
            return _error_response(request_id, -32000, str(exc))

    def _call_tool(self, params: object) -> dict[str, Any]:
        if not isinstance(params, dict):
            raise ValueError("params must be an object")
        name = params.get("name")
        arguments = params.get("arguments") or {}
        if not isinstance(name, str):
            raise ValueError("params.name must be a string")
        if not isinstance(arguments, dict):
            raise ValueError("params.arguments must be an object")
        handler = self.handlers.get(name)
        if handler is None:
            raise ValueError(f"unknown tool: {name}")
        text = handler(arguments)
        return {"content": [{"type": "text", "text": text}], "isError": False}

    def _extract_file_summary_text(self, args: dict[str, Any]) -> str:
        if self._legacy_service is not None and "docPath" not in args:
            return _legacy_text(
                self._legacy_service.extract_file_summary_text(_required_str(args, "fileId"))
            )
        service = self._service_for_args(args)
        return service.extract_file_summary_text()

    def _inspect_pdf(self, args: dict[str, Any]) -> str:
        service = self._service_for_args(args)
        return service.inspect_pdf(include_images=_optional_bool(args, "includeImages", default=True))

    def _extract_pdf_markdown(self, args: dict[str, Any]) -> str:
        service = self._service_for_args(args)
        return service.extract_pdf_markdown()

    def _ocr_pdf(self, args: dict[str, Any]) -> str:
        service = self._service_for_args(args)
        return service.ocr_pdf()

    def _extract_chapter_summary(self, args: dict[str, Any]) -> str:
        if self._legacy_service is not None and "docPath" not in args:
            return _legacy_text(
                self._legacy_service.extract_chapter_summary(
                    _required_str(args, "fileId"),
                    _required_str(args, "chapters"),
                )
            )
        service = self._service_for_args(args)
        return service.extract_chapter_summary(_required_str(args, "chapters"))

    def _extract_chapter_content(self, args: dict[str, Any]) -> str:
        if self._legacy_service is not None and "docPath" not in args:
            return _legacy_text(
                self._legacy_service.extract_chapter_content(
                    _required_str(args, "fileId"),
                    _required_str(args, "chapters"),
                )
            )
        service = self._service_for_args(args)
        return service.extract_chapter_content(_required_str(args, "chapters"))

    def _read_lines(self, args: dict[str, Any]) -> str:
        if self._legacy_service is not None and "docPath" not in args:
            return _legacy_text(
                self._legacy_service.read_lines(
                    _required_str(args, "fileId"),
                    _required_int(args, "startLine"),
                    _required_int(args, "endLine"),
                )
            )
        service = self._service_for_args(args)
        return service.read_lines(_required_int(args, "startLine"), _required_int(args, "endLine"))

    def _search_content(self, args: dict[str, Any]) -> str:
        if self._legacy_service is not None and "docPath" not in args:
            return _legacy_text(
                self._legacy_service.search_content(
                    _required_str(args, "fileId"),
                    _required_str(args, "keywords"),
                    page_no=_optional_int(args, "pageNo"),
                    page_size=_optional_int(args, "pageSize"),
                )
            )
        service = self._service_for_args(args)
        keywords = _required_str(args, "keywords")
        if _looks_like_natural_language_query(keywords):
            return service.query_content(
                keywords,
                page_no=_optional_int(args, "pageNo"),
                page_size=_optional_int(args, "pageSize"),
            )
        return service.search_content(
            keywords,
            page_no=_optional_int(args, "pageNo"),
            page_size=_optional_int(args, "pageSize"),
        )

    def _query_content(self, args: dict[str, Any]) -> str:
        service = self._service_for_args(args)
        return service.query_content(
            _required_str(args, "query"),
            page_no=_optional_int(args, "pageNo"),
            page_size=_optional_int(args, "pageSize"),
        )

    def _extract_table_list(self, args: dict[str, Any]) -> str:
        if self._legacy_service is not None and "docPath" not in args:
            return _legacy_text(
                self._legacy_service.extract_table_list(
                    _required_str(args, "fileId"),
                    _required_str(args, "chapters"),
                )
            )
        service = self._service_for_args(args)
        return service.extract_table_list(_required_str(args, "chapters"))

    def _extract_table_content(self, args: dict[str, Any]) -> str:
        if self._legacy_service is not None and "docPath" not in args:
            return _legacy_text(
                self._legacy_service.extract_table_content(
                    _required_str(args, "fileId"),
                    _required_str(args, "tableTitle"),
                    row_keywords=_optional_str(args, "rowKeywords"),
                    col_keywords=_optional_str(args, "colKeywords"),
                    row_scope=_optional_str(args, "rowScope"),
                )
            )
        service = self._service_for_args(args)
        return service.extract_table_content(
            _required_str(args, "tableTitle"),
            row_keywords=_optional_str(args, "rowKeywords"),
            col_keywords=_optional_str(args, "colKeywords"),
            row_scope=_optional_str(args, "rowScope"),
        )

    def _service_for_args(self, args: dict[str, Any]) -> LocalDocxParseService:
        doc_path = _required_str(args, "docPath")
        resolved = str(Path(doc_path).expanduser().resolve())
        if resolved not in self._documents:
            self._documents[resolved] = Path(resolved)
        if resolved not in self._services:
            self._services[resolved] = LocalDocxParseService(self._documents[resolved], file_id=resolved)
        return self._services[resolved]


def run_stdio(server: LocalMCPServer) -> None:
    for line in sys.stdin:
        if not line.strip():
            continue
        try:
            request = json.loads(line)
            if not isinstance(request, dict):
                response = _error_response(None, -32600, "Request must be an object")
            else:
                response = server.handle(request)
        except json.JSONDecodeError as exc:
            response = _error_response(None, -32700, f"Parse error: {exc.msg}")
        if response is None:
            continue
        sys.stdout.write(json.dumps(response, ensure_ascii=False, separators=(",", ":")) + "\n")
        sys.stdout.flush()


def build_service_from_args(args: argparse.Namespace) -> dict[str, Path]:
    documents: dict[str, Path] = {}
    for item in args.doc or []:
        raw_path = item.split("=", 1)[1] if "=" in item else item
        path = Path(raw_path).expanduser()
        documents[str(path.resolve())] = path
    return documents


def _legacy_text(result: ParsedToolResult) -> str:
    return result.raw_text


def _looks_like_natural_language_query(text: str) -> bool:
    normalized = text.strip()
    compact = normalized.replace(" ", "")
    if not compact:
        return False
    if "," in normalized or "，" in normalized:
        return False
    if len(compact) >= 12:
        return True
    return any(
        marker in compact
        for marker in (
            "有没有",
            "是否",
            "请问",
            "怎么",
            "如何",
            "哪些",
            "多少",
            "要求",
            "规定",
            "依据",
            "原因",
        )
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Run local document parser as an MCP stdio server.")
    parser.add_argument(
        "--doc",
        action="append",
        help="Register a local document as /path/to/file.docx, /path/to/file.doc, /path/to/file.pdf, /path/to/file.xlsx, or fileId=/path/to/file. Can be repeated.",
    )
    args = parser.parse_args()
    run_stdio(LocalMCPServer(build_service_from_args(args)))
    return 0


def _required_str(args: dict[str, Any], key: str) -> str:
    value = args.get(key)
    if not isinstance(value, str) or not value:
        raise ValueError(f"{key} must be a non-empty string")
    return value


def _optional_str(args: dict[str, Any], key: str) -> str | None:
    value = args.get(key)
    if value is None:
        return None
    if not isinstance(value, str):
        raise ValueError(f"{key} must be a string")
    return value


def _required_int(args: dict[str, Any], key: str) -> int:
    value = args.get(key)
    if not isinstance(value, int):
        raise ValueError(f"{key} must be an integer")
    return value


def _optional_int(args: dict[str, Any], key: str) -> int | None:
    value = args.get(key)
    if value is None:
        return None
    if not isinstance(value, int):
        raise ValueError(f"{key} must be an integer")
    return value


def _optional_bool(args: dict[str, Any], key: str, *, default: bool) -> bool:
    value = args.get(key)
    if value is None:
        return default
    if not isinstance(value, bool):
        raise ValueError(f"{key} must be a boolean")
    return value


def _error_response(request_id: object, code: int, message: str) -> dict[str, Any]:
    return {"jsonrpc": "2.0", "id": request_id, "error": {"code": code, "message": message}}


if __name__ == "__main__":
    raise SystemExit(main())
