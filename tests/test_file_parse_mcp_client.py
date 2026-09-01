import json
import unittest

from src.file_parse_mcp_client import parse_tool_result


def event_with_text(text: str, is_error: bool = False):
    return [
        {
            "id": 2,
            "jsonrpc": "2.0",
            "result": {
                "content": [{"type": "text", "text": json.dumps(text, ensure_ascii=False)}],
                "isError": is_error,
            },
        }
    ]


class ParseToolResultTest(unittest.TestCase):
    def test_result_with_meta(self):
        parsed = parse_tool_result(
            event_with_text(
                "<meta><file-id>abc</file-id><file-name>测试.docx</file-name></meta>"
                "<result>正文</result>"
            )
        )
        self.assertTrue(parsed.ok)
        self.assertEqual(parsed.kind, "result")
        self.assertEqual(parsed.text, "正文")
        self.assertEqual(parsed.meta["fileId"], "abc")
        self.assertEqual(parsed.meta["fileName"], "测试.docx")

    def test_business_error(self):
        parsed = parse_tool_result(event_with_text("<error>文件不存在</error>"))
        self.assertFalse(parsed.ok)
        self.assertEqual(parsed.kind, "error")
        self.assertEqual(parsed.text, "文件不存在")

    def test_tip_without_result_is_soft_failure(self):
        parsed = parse_tool_result(event_with_text("<tip>未找到章节</tip>"))
        self.assertFalse(parsed.ok)
        self.assertEqual(parsed.kind, "tip")
        self.assertEqual(parsed.text, "未找到章节")

    def test_tip_with_result_is_success(self):
        parsed = parse_tool_result(event_with_text("<tip>总数:1</tip><result>命中</result>"))
        self.assertTrue(parsed.ok)
        self.assertEqual(parsed.kind, "tip_result")
        self.assertIn("总数:1", parsed.text)
        self.assertIn("命中", parsed.text)


if __name__ == "__main__":
    unittest.main()
