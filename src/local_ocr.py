from __future__ import annotations

import hashlib
import json
import os
import platform
import shutil
import subprocess
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path

import fitz


@dataclass(frozen=True)
class OcrLine:
    text: str
    confidence: float


@dataclass
class OcrPage:
    page: int
    lines: list[OcrLine] = field(default_factory=list)
    error: str | None = None

    @property
    def text(self) -> str:
        return "\n".join(line.text for line in self.lines if line.text.strip())


@dataclass
class OcrDocument:
    available: bool
    engine: str = "macos-vision"
    pages: dict[int, OcrPage] = field(default_factory=dict)
    processing_time_ms: int = 0
    error: str | None = None

    @property
    def successful_pages(self) -> list[int]:
        return sorted(page for page, result in self.pages.items() if result.text.strip())

    @property
    def failed_pages(self) -> list[int]:
        return sorted(page for page, result in self.pages.items() if not result.text.strip())


class LocalOcrBackend:
    def __init__(self, helper_source: str | Path | None = None) -> None:
        self.helper_source = Path(helper_source) if helper_source else _default_helper_source()

    @property
    def available(self) -> bool:
        return (
            platform.system() == "Darwin"
            and shutil.which("clang") is not None
            and self.helper_source.exists()
        )

    @property
    def enabled(self) -> bool:
        return os.environ.get("LOCAL_PDF_OCR", "auto").strip().lower() not in {
            "0",
            "false",
            "no",
            "off",
        }

    def recognize_pdf(
        self,
        path: str | Path,
        pages: list[int],
        *,
        dpi: int | None = None,
    ) -> OcrDocument:
        requested_pages = sorted({page for page in pages if page >= 1})
        if not requested_pages:
            return OcrDocument(available=self.available)
        if not self.enabled:
            return OcrDocument(
                available=False,
                error="本地 OCR 已通过 LOCAL_PDF_OCR=off 禁用",
            )
        if not self.available:
            return OcrDocument(
                available=False,
                error="macOS Vision OCR 不可用，未找到 clang 或辅助程序源码",
            )

        started = time.perf_counter()
        try:
            executable = self._ensure_helper()
            resolution = dpi or _ocr_dpi()
            with tempfile.TemporaryDirectory(prefix="document-parser-ocr-") as temp_dir:
                image_paths = _render_pdf_pages(
                    Path(path),
                    requested_pages,
                    Path(temp_dir),
                    resolution,
                )
                timeout = max(120, len(image_paths) * 30)
                completed = subprocess.run(
                    [str(executable), *(str(image_paths[page]) for page in requested_pages)],
                    check=False,
                    capture_output=True,
                    text=True,
                    timeout=timeout,
                )
                if completed.returncode != 0:
                    message = completed.stderr.strip() or f"Vision OCR 返回状态 {completed.returncode}"
                    return OcrDocument(
                        available=True,
                        processing_time_ms=_elapsed_ms(started),
                        error=message,
                    )
                parsed = _parse_helper_output(completed.stdout, requested_pages)
                parsed.processing_time_ms = _elapsed_ms(started)
                return parsed
        except Exception as exc:  # noqa: BLE001 - OCR failures must preserve PDF fallback.
            return OcrDocument(
                available=True,
                processing_time_ms=_elapsed_ms(started),
                error=f"Vision OCR 失败: {exc}",
            )

    def _ensure_helper(self) -> Path:
        source_bytes = self.helper_source.read_bytes()
        digest = hashlib.sha256(source_bytes).hexdigest()[:16]
        cache_dir = Path(tempfile.gettempdir()) / "document_parser_vision_ocr" / digest
        executable = cache_dir / "vision-ocr"
        if executable.exists():
            return executable
        cache_dir.mkdir(parents=True, exist_ok=True)
        completed = subprocess.run(
            [
                shutil.which("clang") or "clang",
                "-fobjc-arc",
                "-fblocks",
                str(self.helper_source),
                "-framework",
                "Vision",
                "-framework",
                "AppKit",
                "-framework",
                "Foundation",
                "-O",
                "-o",
                str(executable),
            ],
            check=False,
            capture_output=True,
            text=True,
            timeout=120,
        )
        if completed.returncode != 0 or not executable.exists():
            raise RuntimeError(completed.stderr.strip() or "无法编译 Vision OCR 辅助程序")
        return executable


def _default_helper_source() -> Path:
    return Path(__file__).resolve().parents[1] / "scripts" / "vision_ocr.m"


def _ocr_dpi() -> int:
    raw = os.environ.get("LOCAL_PDF_OCR_DPI", "200")
    try:
        return min(300, max(120, int(raw)))
    except ValueError:
        return 200


def _render_pdf_pages(
    path: Path,
    pages: list[int],
    output_dir: Path,
    dpi: int,
) -> dict[int, Path]:
    document = fitz.open(str(path))
    rendered: dict[int, Path] = {}
    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    for page_number in pages:
        if page_number > document.page_count:
            raise ValueError(f"OCR 页码超出范围: {page_number}")
        page = document.load_page(page_number - 1)
        image_path = output_dir / f"page-{page_number:05d}.png"
        page.get_pixmap(matrix=matrix, alpha=False).save(str(image_path))
        rendered[page_number] = image_path
    document.close()
    return rendered


def _parse_helper_output(stdout: str, pages: list[int]) -> OcrDocument:
    results = [line for line in stdout.splitlines() if line.strip()]
    document = OcrDocument(available=True)
    for page_number, raw in zip(pages, results):
        payload = json.loads(raw)
        lines = [
            OcrLine(
                text=str(item.get("text", "")).strip(),
                confidence=float(item.get("confidence", 0.0) or 0.0),
            )
            for item in payload.get("lines", [])
            if str(item.get("text", "")).strip()
        ]
        document.pages[page_number] = OcrPage(
            page=page_number,
            lines=lines,
            error=payload.get("error"),
        )
    for page_number in pages:
        document.pages.setdefault(
            page_number,
            OcrPage(page=page_number, error="OCR 辅助程序未返回该页结果"),
        )
    return document


def _elapsed_ms(started: float) -> int:
    return round((time.perf_counter() - started) * 1000)
