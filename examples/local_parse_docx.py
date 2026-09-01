#!/usr/bin/env python3
from __future__ import annotations

import argparse

from src.local_docx_parser import LocalDocxParseService


def main() -> int:
    parser = argparse.ArgumentParser(description="Parse a local doc/docx/pdf with the compatibility parser.")
    parser.add_argument("doc_path")
    parser.add_argument("--file-id")
    parser.add_argument("--chapter", default="1")
    args = parser.parse_args()

    service = LocalDocxParseService(args.doc_path, file_id=args.file_id)
    print(service.extract_file_summary_text()[:1500])
    print()
    print(service.extract_chapter_summary(args.chapter)[:1500])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
