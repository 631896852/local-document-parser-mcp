import unittest

from src.local_ocr import _parse_helper_output


class LocalOcrTest(unittest.TestCase):
    def test_parses_page_text_and_confidence(self):
        result = _parse_helper_output(
            '{"path":"page-1.png","lines":[{"text":"第一章 总则","confidence":0.98}],"error":null}\n',
            [1],
        )
        self.assertEqual(result.successful_pages, [1])
        self.assertEqual(result.pages[1].text, "第一章 总则")
        self.assertAlmostEqual(result.pages[1].lines[0].confidence, 0.98)

    def test_marks_missing_page_result_as_failed(self):
        result = _parse_helper_output("", [1, 2])
        self.assertEqual(result.failed_pages, [1, 2])
        self.assertIn("未返回", result.pages[2].error)


if __name__ == "__main__":
    unittest.main()
