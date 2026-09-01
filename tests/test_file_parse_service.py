import tempfile
import unittest
from pathlib import Path

from tests.test_local_docx_parser import build_sample_docx
from src.file_parse_service import LocalFileParseService


class LocalFileParseServiceTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / "sample.docx"
        build_sample_docx(self.path)
        self.service = LocalFileParseService({"sample-id": self.path})

    def tearDown(self):
        self.tmp.cleanup()

    def test_returns_parsed_tool_result(self):
        result = self.service.extract_chapter_summary("sample-id", "1")
        self.assertTrue(result.ok)
        self.assertEqual(result.kind, "result")
        self.assertEqual(result.meta["fileId"], "sample-id")
        self.assertIn("1.总论", result.text)

    def test_missing_file_is_business_error(self):
        result = self.service.extract_file_summary_text("missing")
        self.assertFalse(result.ok)
        self.assertEqual(result.kind, "error")
        self.assertEqual(result.text, "文件不存在")

    def test_search_uses_same_result_shape(self):
        result = self.service.search_content("sample-id", "投资估算", page_no=1, page_size=5)
        self.assertTrue(result.ok)
        self.assertEqual(result.kind, "tip_result")
        self.assertIn("搜索结果总条数", result.text)

    def test_query_content_uses_same_result_shape(self):
        result = self.service.query_content("sample-id", "介绍投资估算", page_no=1, page_size=5)
        self.assertTrue(result.ok)
        self.assertEqual(result.kind, "tip_result")
        self.assertIn("自然语言检索结果总条数", result.text)


if __name__ == "__main__":
    unittest.main()
