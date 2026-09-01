# Local Document Parser MCP

<p align="center">
  <strong>Turn local documents into structured knowledge your AI can actually use.</strong><br>
  One interface for Word, PDF, Excel, chapters, tables, search, and OCR.
</p>

<p align="center">
  <a href="README.md">中文</a> · <a href="#get-started-in-5-minutes">Quick start</a> · <a href="docs/local_mcp_server_usage.md">Documentation</a>
</p>

<p align="center">
  <img alt="Python 3.10+" src="https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white">
  <img alt="MCP" src="https://img.shields.io/badge/MCP-stdio-6C47FF">
  <img alt="Local first" src="https://img.shields.io/badge/Data-local--first-18A558">
  <img alt="License" src="https://img.shields.io/badge/License-MIT-yellow.svg">
</p>

Local Document Parser MCP is a local-first document engine and a ready-to-use stdio MCP server. Give it a file path and it can map the document, locate chapters, answer natural-language searches, extract tables, and process scanned PDFs.

It is built for a practical goal: **let an AI work with reports, contracts, plans, and spreadsheets without uploading the files or setting up a vector database first.**

## What it handles

| Your document | What it does |
| --- | --- |
| Long Word reports | Detects chapter hierarchy, extracts sections, and preserves merged table cells |
| Native or complex PDFs | Classifies the PDF, detects columns and table pages, and produces structured Markdown |
| Scanned PDFs | Locates pages that need OCR and uses on-device macOS Vision OCR |
| Excel workbooks | Finds exact cells by worksheet, row keywords, and column keywords |
| Vague questions | Returns relevant passages and table context with `queryContent` |
| AI clients | Exposes 11 tools over stdio MCP and reads a local `docPath` directly |

### Highlights

- **Local first:** parsing stays on your machine; OCR uses the macOS Vision framework.
- **One entry point, six formats:** `.docx`, `.doc`, `.pdf`, `.xlsx`, `.xlsm`, and `.xls`.
- **More than plain text:** keeps chapters, line numbers, table titles, worksheets, and cell locations.
- **Scan-aware:** identifies only the pages that need OCR and merges the result back into the document structure.
- **Graceful fallback:** falls back to PyMuPDF when the enhanced PDF engine is unavailable.
- **Easy to integrate:** use it as a Python module or connect it to any stdio MCP client.

## Get started in 5 minutes

### 1. Install

```bash
git clone https://github.com/631896852/local-document-parser-mcp.git
cd local-document-parser-mcp
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Python 3.10 or newer is required. Legacy `.doc` and `.xls` files also require [LibreOffice](https://www.libreoffice.org/):

```bash
brew install --cask libreoffice
```

### 2. Parse a file directly

```bash
python examples/local_parse_docx.py "/absolute/path/to/report.pdf" --chapter 11
```

The command prints a document overview followed by the requested chapter summary. Word, PDF, and Excel use the same entry point.

### 3. Run the MCP server

```bash
python scripts/local_mcp_server.py
```

Add it to Cherry Studio or another stdio MCP client:

```json
{
  "mcpServers": {
    "local-document-parser": {
      "command": "/absolute/path/to/local-document-parser-mcp/.venv/bin/python",
      "args": [
        "/absolute/path/to/local-document-parser-mcp/scripts/local_mcp_server.py"
      ]
    }
  }
}
```

Then ask your model:

> Parse `/absolute/path/to/report.pdf`, find the content related to “investment estimate,” and return its chapter and surrounding context.

## MCP tools

| Tool | Purpose |
| --- | --- |
| `extractFileSummaryText` | Get cover, table of contents, chapter map, and character counts |
| `extractChapterSummary` | Preview one or more chapters |
| `extractChapterContent` | Extract a specific chapter |
| `readLines` | Read content by global line range |
| `searchContent` | Search for explicit keywords |
| `queryContent` | Search with a natural-language question |
| `extractTableList` | List tables inside a chapter |
| `extractTableContent` | Extract a table or locate Excel cells by row and column keywords |
| `inspectPdf` | Classify a PDF and identify complex or OCR-required pages |
| `extractPdfMarkdown` | Export Markdown that combines native text and OCR |
| `ocrPdf` | Return page-by-page OCR text for scanned pages |

Tools accept an absolute local `docPath`:

```json
{
  "docPath": "/absolute/path/to/report.pdf",
  "query": "What are the main project risks?",
  "pageSize": 5
}
```

## PDF and OCR pipeline

1. `pdf-inspector` classifies text-based, scanned, image-based, or mixed PDFs and produces structured Markdown.
2. PyMuPDF extracts text, chapters, and tables and provides the fallback path.
3. macOS Vision OCR processes only pages that actually require recognition.

OCR supports Simplified Chinese, Traditional Chinese, and English. Pages render at 200 DPI by default:

```bash
LOCAL_PDF_OCR=off          # Disable automatic OCR
LOCAL_PDF_OCR_DPI=200      # 120-300; higher is usually clearer but slower
```

Word, Excel, and text-based PDF parsing still work on non-macOS systems; the Vision OCR step is skipped.

## Python API

```python
from src.local_docx_parser import LocalDocxParseService

parser = LocalDocxParseService("/absolute/path/to/report.docx")

print(parser.extract_file_summary_text())
print(parser.query_content("What is the total project investment?", page_size=5))
print(parser.extract_chapter_content("3.2"))
```

## Test

```bash
python -m unittest discover -s tests
```

The suite covers Word, PDF, Excel, the local OCR adapter, and MCP protocol handling.

## Current limits

- Legacy `.doc` and `.xls` conversion depends on LibreOffice.
- macOS Vision OCR requires a working local `clang`; it is skipped on other platforms.
- Cross-page PDF tables, flowcharts, complex charts, and image semantics are not fully reconstructed yet.
- Excel reads saved formula cache values; it does not execute macros or recalculate formulas.

## Documentation

- [MCP server guide (Chinese)](docs/local_mcp_server_usage.md)
- [Parser capabilities and limits (Chinese)](docs/local_parser_capabilities.md)
- [Local Python API guide (Chinese)](docs/local_service_usage.md)

## Contributing

Issues and pull requests are welcome. High-impact areas include cross-platform OCR, better cross-page table reconstruction, more real-world layout tests, and stronger English heading detection.

## License

[MIT License](LICENSE)
