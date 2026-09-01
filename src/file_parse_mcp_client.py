from __future__ import annotations

import json
import os
import re
import time
from dataclasses import dataclass
from typing import Any

import httpx


DEFAULT_URL = os.environ.get("FILE_PARSE_MCP_URL", "http://127.0.0.1:8000/mcp")
AUTH_ENV = "FILE_PARSE_MCP_AUTH"


class FileParseMCPError(RuntimeError):
    """Raised for transport, protocol, or business-level MCP failures."""


@dataclass(frozen=True)
class ParsedToolResult:
    ok: bool
    kind: str
    text: str
    raw_text: str
    meta: dict[str, str]
    is_error: bool


def _extract_json_response(resp: httpx.Response) -> Any:
    content_type = resp.headers.get("content-type", "")
    text = resp.text
    if "text/event-stream" in content_type:
        events: list[Any] = []
        for line in text.splitlines():
            if not line.startswith("data:"):
                continue
            payload = line.removeprefix("data:").strip()
            if not payload:
                continue
            try:
                events.append(json.loads(payload))
            except json.JSONDecodeError:
                events.append({"raw": payload})
        return events
    if not text.strip():
        return None
    return resp.json()


def _decode_text_payload(value: str) -> str:
    try:
        decoded = json.loads(value)
    except json.JSONDecodeError:
        return value
    return decoded if isinstance(decoded, str) else value


def _tag_text(raw: str, tag: str) -> str | None:
    match = re.search(rf"<{tag}>(.*?)</{tag}>", raw, re.DOTALL)
    return match.group(1).strip() if match else None


def parse_tool_result(response: Any) -> ParsedToolResult:
    if isinstance(response, list):
        if not response:
            raise FileParseMCPError("empty event-stream response")
        item = response[0]
    elif isinstance(response, dict):
        item = response
    else:
        raise FileParseMCPError(f"unsupported response type: {type(response).__name__}")

    if "error" in item:
        raise FileParseMCPError(json.dumps(item["error"], ensure_ascii=False))

    result = item.get("result")
    if not isinstance(result, dict):
        raise FileParseMCPError("missing JSON-RPC result")

    content = result.get("content")
    if not isinstance(content, list) or not content:
        raise FileParseMCPError("missing MCP content")

    first = content[0]
    if not isinstance(first, dict):
        raise FileParseMCPError("invalid MCP content item")

    raw_text = _decode_text_payload(str(first.get("text", "")))
    meta = {
        "fileId": _tag_text(raw_text, "file-id") or "",
        "fileName": _tag_text(raw_text, "file-name") or "",
    }
    error_text = _tag_text(raw_text, "error")
    tip_text = _tag_text(raw_text, "tip")
    result_text = _tag_text(raw_text, "result")

    if error_text is not None:
        return ParsedToolResult(
            ok=False,
            kind="error",
            text=error_text,
            raw_text=raw_text,
            meta=meta,
            is_error=bool(result.get("isError")),
        )

    if tip_text is not None and result_text is None:
        return ParsedToolResult(
            ok=False,
            kind="tip",
            text=tip_text,
            raw_text=raw_text,
            meta=meta,
            is_error=bool(result.get("isError")),
        )

    if tip_text is not None and result_text is not None:
        text = f"{tip_text}\n{result_text}".strip()
        kind = "tip_result"
    elif result_text is not None:
        text = result_text
        kind = "result"
    else:
        text = raw_text
        kind = "text"

    return ParsedToolResult(
        ok=True,
        kind=kind,
        text=text,
        raw_text=raw_text,
        meta=meta,
        is_error=bool(result.get("isError")),
    )


class FileParseMCPClient:
    def __init__(
        self,
        url: str = DEFAULT_URL,
        authorization: str | None = None,
        timeout: float = 60.0,
        delay: float = 1.0,
    ) -> None:
        self.url = url
        self.authorization = authorization or os.environ.get(AUTH_ENV)
        if not self.authorization:
            raise FileParseMCPError(f"missing authorization; set {AUTH_ENV}")
        self.timeout = timeout
        self.delay = delay
        self._client: httpx.Client | None = None
        self._session_id: str | None = None

    def __enter__(self) -> "FileParseMCPClient":
        self.connect()
        return self

    def __exit__(self, exc_type: object, exc: object, tb: object) -> None:
        self.close()

    def close(self) -> None:
        if self._client is not None:
            self._client.close()
        self._client = None
        self._session_id = None

    def connect(self) -> None:
        if self._client is not None:
            return
        self._client = httpx.Client(timeout=self.timeout, trust_env=True)
        response = self._post(
            {
                "jsonrpc": "2.0",
                "id": 1,
                "method": "initialize",
                "params": {
                    "protocolVersion": "2025-03-26",
                    "capabilities": {},
                    "clientInfo": {"name": "file-parse-mcp-client", "version": "0.1"},
                },
            },
            include_session=False,
        )
        self._session_id = (
            response.headers.get("Mcp-Session-Id") or response.headers.get("mcp-session-id")
        )
        time.sleep(self.delay)
        self._post(
            {
                "jsonrpc": "2.0",
                "method": "notifications/initialized",
                "params": {},
            }
        )

    def call_tool(self, name: str, arguments: dict[str, Any]) -> ParsedToolResult:
        self.connect()
        time.sleep(self.delay)
        resp = self._post(
            {
                "jsonrpc": "2.0",
                "id": 2,
                "method": "tools/call",
                "params": {"name": name, "arguments": arguments},
            }
        )
        return parse_tool_result(_extract_json_response(resp))

    def extract_file_summary_text(self, file_id: str) -> ParsedToolResult:
        return self.call_tool("extractFileSummaryText", {"fileId": file_id})

    def extract_chapter_summary(self, file_id: str, chapters: str) -> ParsedToolResult:
        return self.call_tool("extractChapterSummary", {"fileId": file_id, "chapters": chapters})

    def extract_chapter_content(self, file_id: str, chapters: str) -> ParsedToolResult:
        return self.call_tool("extractChapterContent", {"fileId": file_id, "chapters": chapters})

    def read_lines(self, file_id: str, start_line: int, end_line: int) -> ParsedToolResult:
        return self.call_tool(
            "readLines",
            {"fileId": file_id, "startLine": start_line, "endLine": end_line},
        )

    def search_content(
        self,
        file_id: str,
        keywords: str,
        page_no: int | None = None,
        page_size: int | None = None,
    ) -> ParsedToolResult:
        args: dict[str, Any] = {"fileId": file_id, "keywords": keywords}
        if page_no is not None:
            args["pageNo"] = page_no
        if page_size is not None:
            args["pageSize"] = page_size
        return self.call_tool("searchContent", args)

    def extract_table_list(self, file_id: str, chapters: str) -> ParsedToolResult:
        return self.call_tool("extractTableList", {"fileId": file_id, "chapters": chapters})

    def extract_table_content(
        self,
        file_id: str,
        table_title: str,
        row_keywords: str | None = None,
        col_keywords: str | None = None,
        row_scope: str | None = None,
    ) -> ParsedToolResult:
        args: dict[str, Any] = {"fileId": file_id, "tableTitle": table_title}
        if row_keywords:
            args["rowKeywords"] = row_keywords
        if col_keywords:
            args["colKeywords"] = col_keywords
        if row_scope:
            args["rowScope"] = row_scope
        return self.call_tool("extractTableContent", args)

    def _headers(self, include_session: bool = True) -> dict[str, str]:
        headers = {
            "Accept": "application/json, text/event-stream",
            "Content-Type": "application/json",
            "User-Agent": "file-parse-mcp-client/0.1",
            "Authorization": self.authorization or "",
        }
        if include_session and self._session_id:
            headers["Mcp-Session-Id"] = self._session_id
        return headers

    def _post(self, payload: dict[str, Any], include_session: bool = True) -> httpx.Response:
        if self._client is None:
            raise FileParseMCPError("client is not connected")
        resp = self._client.post(
            self.url,
            headers=self._headers(include_session=include_session),
            json=payload,
        )
        resp.raise_for_status()
        return resp
