#!/usr/bin/env python3
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.file_parse_service import build_default_local_service


def main() -> int:
    parser = argparse.ArgumentParser(description="Call the local file parse compatibility service.")
    parser.add_argument("doc_path", help="Absolute path to a local document")
    parser.add_argument("--file-id", default="local-document")
    parser.add_argument("--chapter", default="1")
    parser.add_argument("--keywords", default="投资估算")
    args = parser.parse_args()

    service = build_default_local_service(args.doc_path, file_id=args.file_id)

    summary = service.extract_file_summary_text(args.file_id)
    print("summary:", summary.ok, summary.kind, summary.meta)
    print(summary.text[:800])

    chapter = service.extract_chapter_summary(args.file_id, args.chapter)
    print("\nchapter:", chapter.ok, chapter.kind)
    print(chapter.text[:800])

    search = service.search_content(args.file_id, args.keywords, page_no=1, page_size=5)
    print("\nsearch:", search.ok, search.kind)
    print(search.text[:800])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
