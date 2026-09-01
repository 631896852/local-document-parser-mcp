import tempfile
import unittest
from pathlib import Path

from tests.test_local_docx_parser import build_sample_docx, build_sample_pdf
from src.file_parse_service import LocalFileParseService
from src.local_mcp_server import LocalMCPServer, build_service_from_args


class LocalMCPServerTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / "sample.docx"
        self.pdf_path = Path(self.tmp.name) / "sample.pdf"
        build_sample_docx(self.path)
        build_sample_pdf(self.pdf_path)
        service = LocalFileParseService({"sample-id": self.path})
        self.server = LocalMCPServer(service)

    def tearDown(self):
        self.tmp.cleanup()

    def test_initialize(self):
        response = self.server.handle(
            {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {}}
        )
        self.assertEqual(response["result"]["serverInfo"]["name"], "local-file-parse-service")
        self.assertIn("tools", response["result"]["capabilities"])

    def test_tools_list(self):
        response = self.server.handle({"jsonrpc": "2.0", "id": 2, "method": "tools/list"})
        tools = response["result"]["tools"]
        self.assertEqual(len(tools), 11)
        self.assertIn("extractFileSummaryText", {tool["name"] for tool in tools})
        self.assertIn("queryContent", {tool["name"] for tool in tools})
        self.assertIn("inspectPdf", {tool["name"] for tool in tools})
        self.assertIn("extractPdfMarkdown", {tool["name"] for tool in tools})
        self.assertIn("ocrPdf", {tool["name"] for tool in tools})

    def test_tools_call(self):
        response = self.server.handle(
            {
                "jsonrpc": "2.0",
                "id": 3,
                "method": "tools/call",
                "params": {
                    "name": "extractChapterSummary",
                    "arguments": {"fileId": "sample-id", "chapters": "1"},
                },
            }
        )
        text = response["result"]["content"][0]["text"]
        self.assertIn("<result>", text)
        self.assertIn("1.总论", text)

    def test_tools_call_with_doc_path(self):
        server = LocalMCPServer()
        response = server.handle(
            {
                "jsonrpc": "2.0",
                "id": 5,
                "method": "tools/call",
                "params": {
                    "name": "extractChapterSummary",
                    "arguments": {"docPath": str(self.path), "chapters": "1"},
                },
            }
        )
        text = response["result"]["content"][0]["text"]
        self.assertIn("<result>", text)
        self.assertIn("1.总论", text)

    def test_search_content_natural_language_falls_back_to_query_content(self):
        server = LocalMCPServer()
        response = server.handle(
            {
                "jsonrpc": "2.0",
                "id": 6,
                "method": "tools/call",
                "params": {
                    "name": "searchContent",
                    "arguments": {
                        "docPath": str(self.path),
                        "keywords": "有没有投资估算相关要求",
                    },
                },
            }
        )
        text = response["result"]["content"][0]["text"]
        self.assertIn("自然语言检索结果总条数", text)
        self.assertIn("投资估算", text)

    def test_inspect_pdf_returns_routing_metadata(self):
        server = LocalMCPServer()
        response = server.handle(
            {
                "jsonrpc": "2.0",
                "id": 7,
                "method": "tools/call",
                "params": {
                    "name": "inspectPdf",
                    "arguments": {"docPath": str(self.pdf_path), "includeImages": False},
                },
            }
        )
        text = response["result"]["content"][0]["text"]
        self.assertIn('"pdfType": "text_based"', text)
        self.assertIn('"pagesNeedingOcr": []', text)

    def test_extract_pdf_markdown_keeps_structured_output(self):
        server = LocalMCPServer()
        response = server.handle(
            {
                "jsonrpc": "2.0",
                "id": 8,
                "method": "tools/call",
                "params": {
                    "name": "extractPdfMarkdown",
                    "arguments": {"docPath": str(self.pdf_path)},
                },
            }
        )
        text = response["result"]["content"][0]["text"]
        self.assertIn("Contents", text)
        self.assertIn("Project investment table", text)

    def test_build_service_from_args_accepts_legacy_doc_mapping(self):
        args = type("Args", (), {"doc": [f"sample-id={self.path}"]})()
        documents = build_service_from_args(args)
        self.assertEqual(list(documents.values()), [self.path])

    def test_unknown_tool_returns_jsonrpc_error(self):
        response = self.server.handle(
            {
                "jsonrpc": "2.0",
                "id": 4,
                "method": "tools/call",
                "params": {"name": "missing", "arguments": {}},
            }
        )
        self.assertEqual(response["error"]["code"], -32000)
        self.assertIn("unknown tool", response["error"]["message"])


if __name__ == "__main__":
    unittest.main()
