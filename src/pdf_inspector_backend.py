from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

try:
    import pdf_inspector as _pdf_inspector
except ImportError:  # pragma: no cover - exercised when the optional wheel is unavailable.
    _pdf_inspector = None  # type: ignore[assignment]


@dataclass(frozen=True)
class PdfImageRegion:
    page: int
    name: str
    x: float
    y: float
    width: float
    height: float

    def as_dict(self) -> dict[str, object]:
        return {
            "page": self.page,
            "name": self.name,
            "x": round(self.x, 2),
            "y": round(self.y, 2),
            "width": round(self.width, 2),
            "height": round(self.height, 2),
        }


@dataclass
class PdfInspection:
    available: bool
    pdf_type: str = "unknown"
    confidence: float = 0.0
    page_count: int = 0
    processing_time_ms: int = 0
    pages_needing_ocr: list[int] = field(default_factory=list)
    ocr_reasons_by_page: dict[int, list[str]] = field(default_factory=dict)
    pages_with_tables: list[int] = field(default_factory=list)
    pages_with_columns: list[int] = field(default_factory=list)
    is_complex_layout: bool = False
    has_encoding_issues: bool = False
    title: str | None = None
    markdown: str | None = None
    image_regions: list[PdfImageRegion] = field(default_factory=list)
    error: str | None = None

    def as_dict(self, *, include_images: bool = True) -> dict[str, object]:
        data: dict[str, object] = {
            "engine": "pdf-inspector" if self.available else "pymupdf-fallback",
            "available": self.available,
            "pdfType": self.pdf_type,
            "confidence": round(self.confidence, 4),
            "pageCount": self.page_count,
            "processingTimeMs": self.processing_time_ms,
            "pagesNeedingOcr": self.pages_needing_ocr,
            "ocrReasonsByPage": self.ocr_reasons_by_page,
            "pagesWithTables": self.pages_with_tables,
            "pagesWithColumns": self.pages_with_columns,
            "isComplexLayout": self.is_complex_layout,
            "hasEncodingIssues": self.has_encoding_issues,
            "title": self.title,
            "hasMarkdown": bool(self.markdown and self.markdown.strip()),
        }
        if include_images:
            data["imageRegions"] = [region.as_dict() for region in self.image_regions]
        if self.error:
            data["fallbackReason"] = self.error
        return data


class PdfInspectorBackend:
    """Small fault-tolerant adapter around the optional Rust extension."""

    def __init__(self, module: Any | None = None) -> None:
        self._module = _pdf_inspector if module is None else module

    @property
    def available(self) -> bool:
        return self._module is not None

    def inspect(self, path: str | Path) -> PdfInspection:
        if self._module is None:
            return PdfInspection(
                available=False,
                error="未安装 pdf-inspector，已使用 PyMuPDF 降级解析",
            )
        try:
            result = self._module.process_pdf(str(Path(path)))
        except Exception as exc:  # noqa: BLE001 - parser failures must not break the fallback.
            return PdfInspection(
                available=False,
                error=f"pdf-inspector 解析失败，已使用 PyMuPDF 降级: {exc}",
            )

        reasons: dict[int, list[str]] = {}
        for item in getattr(result, "ocr_reasons_by_page", []) or []:
            page = int(getattr(item, "page", 0) or 0)
            if page <= 0:
                continue
            reasons[page] = [str(reason) for reason in getattr(item, "reasons", []) or []]

        return PdfInspection(
            available=True,
            pdf_type=str(getattr(result, "pdf_type", "unknown")),
            confidence=float(getattr(result, "confidence", 0.0) or 0.0),
            page_count=int(getattr(result, "page_count", 0) or 0),
            processing_time_ms=int(getattr(result, "processing_time_ms", 0) or 0),
            pages_needing_ocr=_int_list(getattr(result, "pages_needing_ocr", [])),
            ocr_reasons_by_page=reasons,
            pages_with_tables=_int_list(getattr(result, "pages_with_tables", [])),
            pages_with_columns=_int_list(getattr(result, "pages_with_columns", [])),
            is_complex_layout=bool(getattr(result, "is_complex_layout", False)),
            has_encoding_issues=bool(getattr(result, "has_encoding_issues", False)),
            title=_optional_text(getattr(result, "title", None)),
            markdown=_optional_text(getattr(result, "markdown", None)),
        )

    def extract_image_regions(self, path: str | Path) -> list[PdfImageRegion]:
        if self._module is None:
            return []
        try:
            items = self._module.extract_text_with_positions(str(Path(path)))
        except Exception:  # noqa: BLE001 - image metadata is an optional enhancement.
            return []
        regions: list[PdfImageRegion] = []
        for item in items:
            if str(getattr(item, "item_type", "")).lower() != "image":
                continue
            text = str(getattr(item, "text", "Image"))
            name = text.removeprefix("[Image: ").removesuffix("]") or "Image"
            regions.append(
                PdfImageRegion(
                    page=int(getattr(item, "page", 0) or 0),
                    name=name,
                    x=float(getattr(item, "x", 0.0) or 0.0),
                    y=float(getattr(item, "y", 0.0) or 0.0),
                    width=float(getattr(item, "width", 0.0) or 0.0),
                    height=float(getattr(item, "height", 0.0) or 0.0),
                )
            )
        return regions


def _int_list(value: object) -> list[int]:
    if not isinstance(value, (list, tuple)):
        return []
    return [int(item) for item in value]


def _optional_text(value: object) -> str | None:
    if value is None:
        return None
    text = str(value)
    return text if text.strip() else None
