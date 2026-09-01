import unittest

from src.pdf_inspector_backend import PdfInspectorBackend


class _Reason:
    page = 2
    reasons = ["no_text"]


class _Result:
    pdf_type = "mixed"
    confidence = 0.91
    page_count = 3
    processing_time_ms = 12
    pages_needing_ocr = [2]
    ocr_reasons_by_page = [_Reason()]
    pages_with_tables = [1]
    pages_with_columns = [3]
    is_complex_layout = True
    has_encoding_issues = False
    title = "Sample"
    markdown = "# Sample"


class _ImageItem:
    item_type = "image"
    text = "[Image: Im0]"
    page = 1
    x = 10.0
    y = 20.0
    width = 30.0
    height = 40.0


class _TextItem:
    item_type = "text"


class _FakeModule:
    @staticmethod
    def process_pdf(path: str):
        return _Result()

    @staticmethod
    def extract_text_with_positions(path: str):
        return [_ImageItem(), _TextItem()]


class PdfInspectorBackendTest(unittest.TestCase):
    def test_maps_process_result(self):
        inspection = PdfInspectorBackend(_FakeModule()).inspect("sample.pdf")
        self.assertTrue(inspection.available)
        self.assertEqual(inspection.pdf_type, "mixed")
        self.assertEqual(inspection.pages_needing_ocr, [2])
        self.assertEqual(inspection.ocr_reasons_by_page, {2: ["no_text"]})
        self.assertEqual(inspection.pages_with_columns, [3])

    def test_extracts_only_image_regions(self):
        regions = PdfInspectorBackend(_FakeModule()).extract_image_regions("sample.pdf")
        self.assertEqual(len(regions), 1)
        self.assertEqual(regions[0].name, "Im0")
        self.assertEqual(regions[0].page, 1)


if __name__ == "__main__":
    unittest.main()
