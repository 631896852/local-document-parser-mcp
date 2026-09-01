from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Protocol

from src.file_parse_mcp_client import ParsedToolResult, parse_tool_result
from src.local_docx_parser import LocalDocxParseService


class FileParseService(Protocol):
    def extract_file_summary_text(self, file_id: str) -> ParsedToolResult: ...

    def extract_chapter_summary(self, file_id: str, chapters: str) -> ParsedToolResult: ...

    def extract_chapter_content(self, file_id: str, chapters: str) -> ParsedToolResult: ...

    def read_lines(self, file_id: str, start_line: int, end_line: int) -> ParsedToolResult: ...

    def search_content(
        self,
        file_id: str,
        keywords: str,
        page_no: int | None = None,
        page_size: int | None = None,
    ) -> ParsedToolResult: ...

    def query_content(
        self,
        file_id: str,
        query: str,
        page_no: int | None = None,
        page_size: int | None = None,
    ) -> ParsedToolResult: ...

    def extract_table_list(self, file_id: str, chapters: str) -> ParsedToolResult: ...

    def extract_table_content(
        self,
        file_id: str,
        table_title: str,
        row_keywords: str | None = None,
        col_keywords: str | None = None,
        row_scope: str | None = None,
    ) -> ParsedToolResult: ...


@dataclass(frozen=True)
class LocalDocument:
    file_id: str
    path: Path


class LocalFileParseService:
    def __init__(self, documents: list[LocalDocument] | dict[str, str | Path]) -> None:
        if isinstance(documents, dict):
            documents = [
                LocalDocument(file_id=file_id, path=Path(path))
                for file_id, path in documents.items()
            ]
        self._paths = {doc.file_id: Path(doc.path) for doc in documents}
        self._services: dict[str, LocalDocxParseService] = {}

    def extract_file_summary_text(self, file_id: str) -> ParsedToolResult:
        return self._call(file_id, lambda service: service.extract_file_summary_text())

    def extract_chapter_summary(self, file_id: str, chapters: str) -> ParsedToolResult:
        return self._call(file_id, lambda service: service.extract_chapter_summary(chapters))

    def extract_chapter_content(self, file_id: str, chapters: str) -> ParsedToolResult:
        return self._call(file_id, lambda service: service.extract_chapter_content(chapters))

    def read_lines(self, file_id: str, start_line: int, end_line: int) -> ParsedToolResult:
        return self._call(file_id, lambda service: service.read_lines(start_line, end_line))

    def search_content(
        self,
        file_id: str,
        keywords: str,
        page_no: int | None = None,
        page_size: int | None = None,
    ) -> ParsedToolResult:
        return self._call(
            file_id,
            lambda service: service.search_content(
                keywords,
                page_no=page_no or 1,
                page_size=page_size or 20,
            ),
        )

    def query_content(
        self,
        file_id: str,
        query: str,
        page_no: int | None = None,
        page_size: int | None = None,
    ) -> ParsedToolResult:
        return self._call(
            file_id,
            lambda service: service.query_content(
                query,
                page_no=page_no or 1,
                page_size=page_size or 10,
            ),
        )

    def extract_table_list(self, file_id: str, chapters: str) -> ParsedToolResult:
        return self._call(file_id, lambda service: service.extract_table_list(chapters))

    def extract_table_content(
        self,
        file_id: str,
        table_title: str,
        row_keywords: str | None = None,
        col_keywords: str | None = None,
        row_scope: str | None = None,
    ) -> ParsedToolResult:
        return self._call(
            file_id,
            lambda service: service.extract_table_content(
                table_title,
                row_scope=row_scope,
                row_keywords=row_keywords,
                col_keywords=col_keywords,
            ),
        )

    def _call(self, file_id: str, func: object) -> ParsedToolResult:
        if file_id not in self._paths:
            return _parse_local_text(f"<error>文件不存在</error>")
        service = self._get_service(file_id)
        text = func(service)  # type: ignore[operator]
        return _parse_local_text(text)

    def _get_service(self, file_id: str) -> LocalDocxParseService:
        if file_id not in self._services:
            self._services[file_id] = LocalDocxParseService(
                self._paths[file_id],
                file_id=file_id,
            )
        return self._services[file_id]


def _parse_local_text(text: str) -> ParsedToolResult:
    return parse_tool_result(
        {
            "jsonrpc": "2.0",
            "id": 1,
            "result": {
                "content": [{"type": "text", "text": text}],
                "isError": False,
            },
        }
    )


def build_default_local_service(
    doc_path: str | Path,
    *,
    file_id: str = "local-document",
) -> LocalFileParseService:
    """Build the compatibility service for one caller-provided local document."""
    return LocalFileParseService({file_id: doc_path})


def tool_name_to_method(name: str) -> str:
    snake = re.sub(r"(?<!^)(?=[A-Z])", "_", name).lower()
    return snake
