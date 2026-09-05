#!/usr/bin/env python3
"""Convert PDF and DOCX documents to structured Markdown.

Native PDFs are processed locally from text spans, fonts, coordinates,
bookmarks, vector table lines, links, and embedded raster images.  Scanned or
non-native PDFs are routed to MinerU OCR.  Every input document is written to
its own directory containing exactly one Markdown document and an ``images``
directory (plus an optional sibling JSON report when requested).
"""

from __future__ import annotations

import argparse
import hashlib
import json
import logging
import os
import re
import shutil
import sys
import uuid
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Iterator, Sequence
from urllib.parse import unquote, urlparse

from word_to_md import WORD_SUFFIXES, convert_word_document

try:
    import pymupdf as fitz
except ImportError:  # PyMuPDF < 1.24
    import fitz  # type: ignore[no-redef]


LOGGER = logging.getLogger("native_pdf_to_md")

HEADING_SIZES = ((22.0, 1), (15.0, 2), (12.2, 3))
CALLOUT_LABELS = {"note", "tip", "warning", "caution", "important", "example"}
BULLETS = {"•": "-", "●": "-", "▪": "-", "·": "-", "–": "-", "◦": "  -", "○": "  -"}
URL_RE = re.compile(r"(?<![\[(])https?://[^\s<>*]+", re.IGNORECASE)
HEADING_NUMBER_RE = re.compile(r"^(?:\d+(?:\.\d+){0,5}|[A-Z](?:\.\d+)*)[.)]?\s+")
FIGURE_CAPTION_RE = re.compile(
    r"^(?:Figure|Fig\.)\s+[A-Z0-9]+(?:[-.][A-Z0-9]+)*\s*[:.]", re.IGNORECASE
)
TABLE_CAPTION_RE = re.compile(
    r"^Table\s+[A-Z0-9]+(?:[-.][A-Z0-9]+)*(?:\s*[:.]|\s+)", re.IGNORECASE
)
PROMPT_RE = re.compile(r"^(?:SETUP|ANALYSIS|ATPG|FAULT|PATTERN|SIMULATION)>\s", re.IGNORECASE)
DOT_LEADER_RE = re.compile(r"\.{5,}\s*(?:\d+|[A-Z])-\d+\s*$", re.IGNORECASE)
ROMAN_SECTION_RE = re.compile(r"^[IVXLCDM]+\.\s+[A-Z][A-Z0-9 &/(),'-]+$")
LETTER_SECTION_RE = re.compile(r"^[A-Z]\.\s+\S.+$")
REFERENCE_START_RE = re.compile(r"^\[\d+\]\s*")
SUPPORTED_INPUT_SUFFIXES = {".pdf"} | WORD_SUFFIXES


class ConversionError(RuntimeError):
    """Raised when a PDF cannot be converted by the selected route."""


@dataclass(frozen=True)
class Bookmark:
    level: int
    title: str
    page: int
    key: str
    y: float | None = None


@dataclass
class Element:
    kind: str
    bbox: fitz.Rect
    data: Any
    order: int = 0


@dataclass
class DocumentReport:
    source: str
    output: str
    pages: int
    text_characters: int = 0
    headings: dict[str, int] = field(default_factory=lambda: defaultdict(int))
    images: int = 0
    tables: int = 0
    code_blocks: int = 0
    bookmarks_total: int = 0
    bookmarks_matched: int = 0
    unmatched_bookmarks: list[dict[str, Any]] = field(default_factory=list)
    ocr_used: bool = False
    conversion_engine: str = "native"
    native_text_rejection: str | None = None
    chunks: int = 1

    def as_dict(self) -> dict[str, Any]:
        return {
            "source": self.source,
            "output": self.output,
            "pages": self.pages,
            "text_characters": self.text_characters,
            "headings": dict(self.headings),
            "images": self.images,
            "tables": self.tables,
            "code_blocks": self.code_blocks,
            "bookmarks_total": self.bookmarks_total,
            "bookmarks_matched": self.bookmarks_matched,
            "unmatched_bookmarks": self.unmatched_bookmarks,
            "ocr_used": self.ocr_used,
            "conversion_engine": self.conversion_engine,
            "native_text_rejection": self.native_text_rejection,
            "chunks": self.chunks,
        }


def normalized_key(text: str) -> str:
    text = text.replace("\u00ad", "").replace("\u00a0", " ")
    text = text.replace("™", "").replace("®", "")
    text = re.sub(r"[\s.·•…]+", " ", text)
    return text.strip().casefold()


def heading_match_key(text: str) -> str:
    """Normalize a heading while tolerating a separately drawn chapter number."""
    text = clean_text(text)
    text = re.sub(r"^[•●▪·–◦○]\s*", "", text)
    text = re.sub(r"^Chapter\s+\d+(?:\.\d+)*\s*[:.)-]?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"^\d+(?:\.\d+)*\s*[:.)-]?\s+", "", text)
    return normalized_key(text)


def clean_text(text: str) -> str:
    replacements = {
        "\u00ad": "",
        "\u00a0": " ",
        "\uf0b7": "•",
        "\uf0a7": "▪",
        "\uf0d8": "◦",
        "\ufb01": "fi",
        "\ufb02": "fl",
        "\u200b": "",
    }
    for old, new in replacements.items():
        text = text.replace(old, new)
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def safe_name(value: str, fallback: str = "document") -> str:
    value = unquote(value).strip().rstrip(". ")
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value)
    value = re.sub(r"\s+", "_", value)
    return value[:120] or fallback


def rect_intersection_ratio(a: fitz.Rect, b: fitz.Rect) -> float:
    intersection = a & b
    if intersection.is_empty or a.get_area() <= 0:
        return 0.0
    return intersection.get_area() / a.get_area()


def inside_any(rect: fitz.Rect, regions: Sequence[fitz.Rect], threshold: float = 0.55) -> bool:
    return any(rect_intersection_ratio(rect, region) >= threshold for region in regions)


def make_bookmarks(doc: fitz.Document) -> tuple[list[Bookmark], dict[int, list[Bookmark]]]:
    bookmarks: list[Bookmark] = []
    by_page: dict[int, list[Bookmark]] = defaultdict(list)
    for row in doc.get_toc(simple=False):
        if len(row) < 3:
            continue
        level, title, page = int(row[0]), clean_text(str(row[1])), int(row[2])
        if not title or page < 1 or page > doc.page_count:
            continue
        destination = row[3] if len(row) > 3 and isinstance(row[3], dict) else {}
        point = destination.get("to")
        try:
            destination_y = float(point.y if hasattr(point, "y") else point[1])
        except (IndexError, TypeError, ValueError):
            destination_y = None
        item = Bookmark(
            max(1, min(level, 6)),
            title,
            page,
            normalized_key(title),
            destination_y,
        )
        bookmarks.append(item)
        by_page[page].append(item)
    return bookmarks, by_page


def block_text(block: dict[str, Any], preserve_lines: bool = False) -> str:
    lines: list[str] = []
    for line in block.get("lines", []):
        value = clean_text("".join(span.get("text", "") for span in line.get("spans", [])))
        if value:
            lines.append(value)
    if preserve_lines:
        return "\n".join(lines)
    merged = ""
    for line in lines:
        if not merged:
            merged = line
        elif merged.endswith("-") and line[:1].islower():
            merged = merged[:-1] + line
        else:
            merged += " " + line
    return clean_text(merged)


def block_font_info(block: dict[str, Any]) -> tuple[float, bool, bool, bool, int | None]:
    spans = [span for line in block.get("lines", []) for span in line.get("spans", []) if span.get("text", "").strip()]
    if not spans:
        return 0.0, False, False, False, None
    weighted = sorted(
        ((float(s.get("size", 0)), len(s.get("text", ""))) for s in spans),
        key=lambda item: item[0],
    )
    total = sum(weight for _, weight in weighted) or 1
    midpoint = total / 2
    cumulative = 0
    median_size = weighted[-1][0]
    for size, weight in weighted:
        cumulative += weight
        if cumulative >= midpoint:
            median_size = size
            break
    font_names = " ".join(str(s.get("font", "")).casefold() for s in spans)
    flags = 0
    for span in spans:
        flags |= int(span.get("flags", 0))
    colors = [int(s.get("color", 0)) for s in spans]
    return (
        median_size,
        any(weight in font_names for weight in ("bold", "black", "semibold", "demi")) or bool(flags & 16),
        "italic" in font_names or "oblique" in font_names or bool(flags & 2),
        "courier" in font_names or "mono" in font_names or bool(flags & 8),
        max(set(colors), key=colors.count) if colors else None,
    )


MARGIN_PAGE_NUMBER_RE = re.compile(
    r"^(?:page\s*)?(?:\d+|[ivxlcdm]+)(?:\s*(?:of|/)\s*\d+)?$",
    re.IGNORECASE,
)


def margin_text_key(text: str) -> str:
    return re.sub(r"\s+", " ", clean_text(text)).strip().casefold()


def detect_repeated_margin_texts(doc: fitz.Document) -> set[str]:
    """Find running heads and footers repeated in outer page bands."""
    counts: Counter[str] = Counter()
    for page in doc:
        seen_on_page: set[str] = set()
        for block in page.get_text("dict", sort=True).get("blocks", []):
            if block.get("type") != 0 or not block.get("lines"):
                continue
            rect = fitz.Rect(block["bbox"])
            in_margin = (
                rect.y1 <= page.rect.height * 0.12
                or rect.y0 >= page.rect.height * 0.88
            )
            if not in_margin:
                continue
            text = block_text(block)
            key = margin_text_key(text)
            if key and len(key) <= 260:
                seen_on_page.add(key)
        counts.update(seen_on_page)
    minimum = 2 if doc.page_count <= 12 else 3
    return {key for key, count in counts.items() if count >= minimum}


def is_header_or_footer(
    rect: fitz.Rect,
    page_rect: fitz.Rect,
    page_number: int,
    text: str = "",
    repeated_margin_texts: set[str] | None = None,
) -> bool:
    # The first page may use the full canvas for the cover.  On regular pages,
    # headers and footers occupy stable outer bands.  Wider candidate bands
    # are used only for text proven to repeat across pages.
    if page_number == 1:
        return rect.y0 >= page_rect.height * 0.94
    top_extreme = rect.y1 <= page_rect.height * 0.075
    bottom_extreme = rect.y0 >= page_rect.height * 0.89
    if top_extreme or bottom_extreme:
        return True
    top_candidate = rect.y1 <= page_rect.height * 0.12
    bottom_candidate = rect.y0 >= page_rect.height * 0.88
    key = margin_text_key(text)
    if bottom_candidate and MARGIN_PAGE_NUMBER_RE.fullmatch(key):
        return True
    return bool(
        (top_candidate or bottom_candidate)
        and key
        and repeated_margin_texts
        and key in repeated_margin_texts
    )


def detect_heading(
    text: str,
    block: dict[str, Any],
    page_bookmarks: Sequence[Bookmark],
    matched: set[tuple[int, str]],
    page_number: int,
    suppress_fallback: bool = False,
) -> tuple[int | None, Bookmark | None]:
    key = normalized_key(text)
    candidates = [
        bookmark
        for bookmark in page_bookmarks
        if bookmark.key == key
        and (bookmark.page, bookmark.key) not in matched
    ]
    if not candidates:
        match_key = heading_match_key(text)
        candidates = [
            bookmark
            for bookmark in page_bookmarks
            if heading_match_key(bookmark.title) == match_key
            and (bookmark.page, bookmark.key) not in matched
        ]
    if candidates:
        bookmark = candidates[0]
        return bookmark.level, bookmark

    if suppress_fallback:
        return None, None

    if ROMAN_SECTION_RE.fullmatch(text):
        return 2, None
    if LETTER_SECTION_RE.fullmatch(text) and len(text) <= 100:
        return 3, None
    if text.strip().upper() == "REFERENCES":
        return 2, None

    size, bold, _, _, color = block_font_info(block)
    # Native-font fallback for PDFs whose visible heading was omitted from the
    # bookmark tree.  It uses exact document styling, never OCR inference.
    if bold and len(text) <= 180 and not FIGURE_CAPTION_RE.match(text) and not TABLE_CAPTION_RE.match(text):
        for minimum_size, level in HEADING_SIZES:
            if size >= minimum_size:
                return level, None
        # A colored 13pt bold span is the fourth visible heading tier in many
        # technical manuals.  Only number-prefixed or short title-like blocks
        # are accepted to avoid turning emphasized prose into headings.
        if size >= 11.8 and color not in (None, 0) and (HEADING_NUMBER_RE.match(text) or len(text.split()) <= 14):
            return 4, None

    if page_number == 1 and bold and size >= 20:
        return 1, None
    return None, None


def markdown_escape(text: str) -> str:
    return text.replace("\\", "\\\\").replace("|", "\\|")


def apply_links(text: str) -> str:
    def replace(match: re.Match[str]) -> str:
        url = match.group(0)
        trailing = ""
        while url and url[-1] in '.,;:)':
            trailing = url[-1] + trailing
            url = url[:-1]
        return f"[{url}]({url}){trailing}"

    return URL_RE.sub(replace, text)


def is_math_font(font: str) -> bool:
    name = font.upper().replace("+", "")
    return name.startswith(("CMMI", "CMSY", "CMEX", "CMR"))


def math_text(value: str, font: str) -> str:
    value = value.strip()
    if font.upper().startswith("CMEX") and value == "P":
        return r"\sum"
    replacements = {
        "π": r"\pi",
        "λ": r"\lambda",
        "∈": "\\in ",
        "√": r"\sqrt",
        "−": "-",
    }
    # Discard spaces originating in PDF positioning, but retain the delimiter
    # appended to LaTeX control words such as ``\in ``. Without it, KaTeX
    # reads ``\in S`` as the undefined command ``\inS``.
    return "".join(
        "" if character.isspace() else replacements.get(character, character)
        for character in value
    )


def figure_alt_text(caption: str, page_number: int, order: int) -> str:
    """Use a short image alt label that cannot be broken by caption citations."""
    match = re.match(
        r"^((?:Figure|Fig\.)\s+[A-Z0-9]+(?:[-.][A-Z0-9]+)*)",
        caption,
        re.IGNORECASE,
    )
    if match:
        return match.group(1).rstrip(".")
    return f"Page {page_number} image {order}"


def math_sequence_markdown(spans: Sequence[dict[str, Any]]) -> str:
    raw_sequence = "".join(str(span.get("text", "")) for span in spans)
    leading = raw_sequence[: len(raw_sequence) - len(raw_sequence.lstrip())]
    trailing = raw_sequence[len(raw_sequence.rstrip()) :]
    meaningful = [span for span in spans if str(span.get("text", "")).strip()]
    if not meaningful:
        return ""
    base_size = max(float(span.get("size", 0)) for span in meaningful)
    base_origins = [
        float(span.get("origin", (0, span.get("bbox", (0, 0, 0, 0))[3]))[1])
        for span in meaningful
        if float(span.get("size", 0)) >= base_size * 0.9
    ]
    baseline = sorted(base_origins)[len(base_origins) // 2] if base_origins else 0.0
    pieces: list[str] = []
    pending_marker = ""
    pending_script: list[str] = []

    def flush_script() -> None:
        nonlocal pending_marker
        if pending_marker and pending_script:
            pieces.append(f"{pending_marker}{{{''.join(pending_script)}}}")
        pending_marker = ""
        pending_script.clear()

    for span in meaningful:
        value = math_text(str(span.get("text", "")), str(span.get("font", "")))
        if not value:
            continue
        size = float(span.get("size", 0))
        origin_y = float(span.get("origin", (0, span.get("bbox", (0, 0, 0, 0))[3]))[1])
        if size < base_size * 0.85:
            marker = "^" if origin_y < baseline - 0.7 else "_"
            if pending_marker and pending_marker != marker:
                flush_script()
            pending_marker = marker
            pending_script.append(value)
        else:
            flush_script()
            pieces.append(value)
    flush_script()
    return leading + "$" + "".join(pieces) + "$" + trailing


def inline_markdown(block: dict[str, Any]) -> str:
    raw_lines = [
        clean_text("".join(str(span.get("text", "")) for span in line.get("spans", [])))
        for line in block.get("lines", [])
    ]
    if any("@" in line for line in raw_lines):
        contact_lines: list[str] = []
        for line in raw_lines:
            grouped = re.fullmatch(r"\{([^{}]+)\}@([A-Z0-9.-]+\.[A-Z]{2,})", line, re.IGNORECASE)
            if grouped:
                users = [item.strip() for item in grouped.group(1).split(",") if item.strip()]
                line = ", ".join(f"{user}@{grouped.group(2)}" for user in users)
            contact_lines.append(line)
        return "  \n".join(line for line in contact_lines if line)

    normal_spans = [
        span
        for line in block.get("lines", [])
        for span in line.get("spans", [])
        if not is_math_font(str(span.get("font", ""))) and str(span.get("text", "")).strip()
    ]
    normal_characters = sum(len(str(span.get("text", "")).strip()) for span in normal_spans)
    bold_characters = sum(
        len(str(span.get("text", "")).strip())
        for span in normal_spans
        if "bold" in str(span.get("font", "")).casefold() or bool(int(span.get("flags", 0)) & 16)
    )
    # In abstracts and similar regions, a Medium/Bold font is the paragraph's
    # baseline design rather than dozens of separate semantic emphases.
    suppress_baseline_bold = normal_characters >= 80 and bold_characters >= normal_characters * 0.7

    pieces: list[str] = []
    for line in block.get("lines", []):
        line_parts: list[str] = []
        spans = list(line.get("spans", []))
        index = 0
        while index < len(spans):
            span = spans[index]
            raw = span.get("text", "").replace("\u00ad", "").replace("\u00a0", " ")
            if not raw:
                index += 1
                continue
            font = str(span.get("font", "")).casefold()
            if is_math_font(font):
                math_spans: list[dict[str, Any]] = []
                while index < len(spans) and is_math_font(str(spans[index].get("font", ""))):
                    math_spans.append(spans[index])
                    index += 1
                value = math_sequence_markdown(math_spans)
                if value:
                    line_parts.append(value)
                continue
            flags = int(span.get("flags", 0))
            bold = "bold" in font or bool(flags & 16)
            italic = "italic" in font or "oblique" in font or bool(flags & 2)
            if suppress_baseline_bold:
                bold = False
            run = [span]
            index += 1
            while index < len(spans):
                candidate = spans[index]
                candidate_font = str(candidate.get("font", "")).casefold()
                if is_math_font(candidate_font):
                    break
                candidate_flags = int(candidate.get("flags", 0))
                candidate_bold = "bold" in candidate_font or bool(candidate_flags & 16)
                candidate_italic = (
                    "italic" in candidate_font or "oblique" in candidate_font or bool(candidate_flags & 2)
                )
                if (candidate_bold, candidate_italic) != (bold, italic):
                    break
                run.append(candidate)
                index += 1
            value = "".join(str(item.get("text", "")) for item in run).replace("\u00ad", "").replace("\u00a0", " ")
            if value.strip():
                prefix = value[: len(value) - len(value.lstrip())]
                suffix = value[len(value.rstrip()) :]
                core = value.strip().replace("*", "\\*").replace("_", "\\_")
                if bold and italic:
                    core = f"***{core}***"
                elif bold:
                    core = f"**{core}**"
                elif italic:
                    core = f"*{core}*"
                value = prefix + core + suffix
            line_parts.append(value)
        line_value = "".join(line_parts).strip()
        if line_value:
            pieces.append(line_value)
    merged = ""
    for line in pieces:
        if not merged:
            merged = line
        elif merged.endswith("-") and re.match(r"^[a-z]", re.sub(r"^[*_]+", "", line)):
            merged = merged[:-1] + line
        else:
            merged += " " + line
    # PDF generators commonly create one run per visual line. Collapse equal
    # Markdown emphasis across line boundaries into a single semantic run.
    previous = None
    while previous != merged:
        previous = merged
        merged = re.sub(r"\*\*\*([^*]+)\*\*\*\s+\*\*\*([^*]+)\*\*\*", r"***\1 \2***", merged)
        merged = re.sub(r"\*\*([^*]+)\*\*\s+\*\*([^*]+)\*\*", r"**\1 \2**", merged)
        merged = re.sub(r"(?<!\*)\*([^*]+)\*\s+\*([^*]+)\*(?!\*)", r"*\1 \2*", merged)
    merged = re.sub(r"[ \t]+", " ", merged).strip()
    return apply_links(merged)


def plain_lines(block: dict[str, Any]) -> list[str]:
    return [
        clean_text("".join(span.get("text", "") for span in line.get("spans", [])))
        for line in block.get("lines", [])
        if clean_text("".join(span.get("text", "") for span in line.get("spans", [])))
    ]


def block_from_lines(block: dict[str, Any], lines: Sequence[dict[str, Any]]) -> dict[str, Any]:
    result = dict(block)
    result["lines"] = list(lines)
    rect = fitz.Rect(lines[0]["bbox"])
    for line in lines[1:]:
        rect |= fitz.Rect(line["bbox"])
    result["bbox"] = tuple(rect)
    return result


def display_equation_markdown(block: dict[str, Any]) -> str | None:
    """Rebuild a native stacked fraction as display LaTeX without OCR."""
    lines = list(block.get("lines", []))
    if len(lines) < 5:
        return None
    all_spans = [span for line in lines for span in line.get("spans", [])]
    if not any(str(span.get("font", "")).upper().startswith("CMEX") for span in all_spans):
        return None
    number_match = re.fullmatch(r"\((\d+)\)", clean_text("".join(str(span.get("text", "")) for span in lines[-1].get("spans", []))))
    if not number_match:
        return None

    left = math_sequence_markdown(lines[0].get("spans", [])).strip().strip("$")
    numerator = math_sequence_markdown(lines[1].get("spans", [])).strip().strip("$")
    operator_spans = [span for span in lines[2].get("spans", []) if str(span.get("text", "")).strip()]
    denominator_spans = [span for span in lines[3].get("spans", []) if str(span.get("text", "")).strip()]
    if not left.endswith("=") or not numerator or not operator_spans or not denominator_spans:
        return None
    upper = "".join(
        math_text(str(span.get("text", "")), str(span.get("font", "")))
        for span in operator_spans
        if not str(span.get("font", "")).upper().startswith("CMEX")
    )
    lambda_index = next(
        (index for index, span in enumerate(denominator_spans) if "λ" in str(span.get("text", ""))),
        None,
    )
    if lambda_index is None:
        return None
    lower = "".join(
        math_text(str(span.get("text", "")), str(span.get("font", "")))
        for span in denominator_spans[:lambda_index]
    )
    term = math_sequence_markdown(denominator_spans[lambda_index:]).strip().strip("$")
    if not upper or not lower or not term:
        return None
    return (
        "$$\n"
        f"{left} \\frac{{{numerator}}}{{\\sum_{{{lower}}}^{{{upper}}} {term}}} "
        f"\\tag{{{number_match.group(1)}}}\n"
        "$$"
    )


def reference_markdown_parts(block: dict[str, Any]) -> list[str]:
    """Split a whole-column references block into one Markdown paragraph per citation."""
    groups: list[list[dict[str, Any]]] = []
    for line in block.get("lines", []):
        text = clean_text("".join(str(span.get("text", "")) for span in line.get("spans", [])))
        if REFERENCE_START_RE.match(text):
            groups.append([])
        if groups:
            groups[-1].append(line)
    if not groups or not REFERENCE_START_RE.match(block_text(block)):
        return []
    return [inline_markdown(block_from_lines(block, group)) for group in groups if group]


def academic_numbered_list_markdown(block: dict[str, Any]) -> str | None:
    """Split multiple IEEE numbered items stored in one native text block."""
    groups: list[list[dict[str, Any]]] = []
    for line in block.get("lines", []):
        text = clean_text("".join(str(span.get("text", "")) for span in line.get("spans", [])))
        if re.match(r"^\d+\)\s+", text):
            groups.append([])
        if groups:
            groups[-1].append(line)
    if len(groups) < 2:
        return None
    rendered: list[str] = []
    for group in groups:
        value = inline_markdown(block_from_lines(block, group))
        value = re.sub(r"^(\d+)\)\s*", r"\1. ", value)
        rendered.append(value)
    return "\n".join(rendered)


def paragraph_markdown_parts(
    block: dict[str, Any], page_rect: fitz.Rect, two_column: bool
) -> list[str]:
    """Split native first-line-indented paragraphs and preserve their indentation."""
    lines = list(block.get("lines", []))
    if not lines:
        return []
    line_x = [float(line.get("bbox", block.get("bbox", (0, 0, 0, 0)))[0]) for line in lines]
    if two_column:
        mid = page_rect.x0 + page_rect.width / 2
        base = page_rect.x0 + page_rect.width * 0.08 if fitz.Rect(block["bbox"]).x1 <= mid + 8 else mid + 6
    else:
        base = min(line_x)

    groups: list[list[dict[str, Any]]] = []
    indented: list[bool] = []
    for index, (line, x0) in enumerate(zip(lines, line_x)):
        delta = x0 - base
        starts_paragraph = 5 <= delta <= 16
        if index == 0 or starts_paragraph:
            groups.append([])
            indented.append(starts_paragraph)
        groups[-1].append(line)

    parts: list[str] = []
    for group, has_indent in zip(groups, indented):
        value = inline_markdown(block_from_lines(block, group))
        if value:
            parts.append(("&emsp;&emsp;" if has_indent else "") + value)
    return parts


def code_markdown(block: dict[str, Any]) -> str:
    lines: list[str] = []
    for line in block.get("lines", []):
        spans = line.get("spans", [])
        if not spans:
            continue
        raw = "".join(str(span.get("text", "")) for span in spans).replace("\u00a0", " ").rstrip()
        x0 = float(line.get("bbox", block.get("bbox", (0, 0, 0, 0)))[0])
        # Some PDFs encode indentation as real space characters while others
        # position every line separately.  Preserve whichever representation
        # carries the larger indent instead of stripping both away.
        text_indent = len(raw) - len(raw.lstrip(" \t"))
        position_indent = max(0, round((x0 - float(block["bbox"][0])) / 6.2))
        indent = max(text_indent, position_indent)
        lines.append((" " * indent) + raw.lstrip(" \t"))
    while lines and not lines[-1].strip():
        lines.pop()
    language = "tcl" if any(PROMPT_RE.match(line.lstrip()) or "set_context" in line for line in lines) else "text"
    return f"```{language}\n" + "\n".join(lines) + "\n```"


def listing_block_markdown(block: dict[str, Any], page_rect: fitz.Rect) -> str:
    """Render TOC / list-of-figures lines as visibly separate nested items."""
    rendered: list[str] = []
    # This document family uses a 61pt content margin and 30pt per outline
    # level.  Deriving it from page width keeps the same proportions for pages
    # that are scaled versions of Letter size.
    base_x = page_rect.width * (61.0 / 612.0)
    step = max(18.0, page_rect.width * (30.0 / 612.0))
    for line in block.get("lines", []):
        value = clean_text("".join(str(span.get("text", "")) for span in line.get("spans", [])))
        if not value:
            continue
        x0 = float(line.get("bbox", block.get("bbox", (base_x, 0, 0, 0)))[0])
        level = max(0, min(5, round((x0 - base_x) / step)))
        rendered.append(("  " * level) + "- " + apply_links(value))
    return "\n".join(rendered)


def table_markdown(rows: Sequence[Sequence[Any]]) -> str:
    normalized: list[list[str]] = []
    width = max((len(row) for row in rows), default=0)
    if width == 0:
        return ""
    for row in rows:
        values = []
        for cell in row:
            value = clean_text(str(cell or "").replace("\n", " "))
            values.append(markdown_escape(value))
        values.extend([""] * (width - len(values)))
        normalized.append(values)
    if not normalized:
        return ""
    header = normalized[0]
    body = normalized[1:]
    rendered = ["| " + " | ".join(header) + " |", "| " + " | ".join(["---"] * width) + " |"]
    rendered.extend("| " + " | ".join(row) + " |" for row in body)
    return "\n".join(rendered)


def native_cell_text(page: fitz.Page, cell: Sequence[float] | None) -> str:
    """Read a table cell from native PDF words instead of reordered glyphs."""
    if cell is None:
        return ""
    rect = fitz.Rect(cell)
    words = page.get_text("words", clip=rect, sort=True)
    return clean_text(" ".join(str(word[4]) for word in words if len(word) > 4))


def native_table_rows(page: fitz.Page, table: Any) -> list[list[str]]:
    """Extract detected table geometry using the PDF's intact word tokens."""
    rows: list[list[str]] = []
    table_rows = list(getattr(table, "rows", []))
    for row in table_rows:
        rows.append([native_cell_text(page, cell) for cell in getattr(row, "cells", [])])
    # IEEE tables often use a merged super-header above ordinary column names.
    # Explicit grid recovery deliberately uses the major row rules, so remove
    # the clipped super-header fragments from numeric column headers.
    if rows and table_rows and len(rows[0]) >= 4:
        first_cells = list(getattr(table_rows[0], "cells", []))
        valid_cells = [fitz.Rect(cell) for cell in first_cells if cell is not None]
        if valid_cells:
            header_rect = fitz.Rect(valid_cells[0])
            for rect in valid_cells[1:]:
                header_rect |= rect
            header_words = page.get_text("words", clip=header_rect, sort=True)
            header_text = clean_text(" ".join(str(word[4]) for word in header_words if len(word) > 4)).casefold()
            if "backtracks" in header_text and "guidance" in header_text and "data" in header_text:
                for index, cell in enumerate(first_cells[1:], start=1):
                    if cell is None:
                        continue
                    words = page.get_text("words", clip=fitz.Rect(cell), sort=True)
                    if not words:
                        continue
                    lowest = max(float(word[1]) for word in words)
                    bottom_words = [str(word[4]) for word in words if abs(float(word[1]) - lowest) <= 2]
                    if bottom_words:
                        rows[0][index] = clean_text(" ".join(bottom_words))
    return rows


def caption_guided_ruled_tables(page: fitz.Page) -> list[Element]:
    """Recover ruled tables whose merged headers confuse automatic detection."""
    captions: list[fitz.Rect] = []
    for block in page.get_text("dict", sort=True).get("blocks", []):
        if block.get("type") != 0 or not block.get("lines"):
            continue
        if TABLE_CAPTION_RE.match(block_text(block)):
            captions.append(fitz.Rect(block["bbox"]))
    if not captions:
        return []

    horizontal: list[fitz.Rect] = []
    vertical: list[fitz.Rect] = []
    for drawing in page.get_drawings():
        rect = fitz.Rect(drawing["rect"])
        if rect.width >= page.rect.width * 0.16 and rect.height <= 1.5:
            horizontal.append(rect)
        elif rect.height >= 3 and rect.width <= 1.5:
            vertical.append(rect)

    # Full-width row rules share practically identical endpoints. Split equal
    # endpoint groups at large vertical gaps so stacked tables stay separate.
    endpoint_groups: dict[tuple[int, int], list[float]] = defaultdict(list)
    for rect in horizontal:
        endpoint_groups[(round(rect.x0), round(rect.x1))].append((rect.y0 + rect.y1) / 2)
    grids: list[tuple[float, float, list[float]]] = []
    for (x0, x1), values in endpoint_groups.items():
        ys = sorted(set(round(value, 2) for value in values))
        segment: list[float] = []
        for value in ys:
            if segment and value - segment[-1] > 35:
                if len(segment) >= 3:
                    grids.append((float(x0), float(x1), segment))
                segment = []
            segment.append(value)
        if len(segment) >= 3:
            grids.append((float(x0), float(x1), segment))

    results: list[Element] = []
    for caption in captions:
        choices = [
            grid
            for grid in grids
            if grid[2][0] >= caption.y1 - 4
            and grid[2][0] - caption.y1 <= 80
            and grid[0] < caption.x1
            and grid[1] > caption.x0
        ]
        if not choices:
            continue
        x0, x1, ys = min(choices, key=lambda grid: grid[2][0] - caption.y1)
        y0, y1 = ys[0], ys[-1]
        xs_by_position: dict[int, int] = defaultdict(int)
        for rect in vertical:
            x = (rect.x0 + rect.x1) / 2
            overlap = max(0.0, min(rect.y1, y1) - max(rect.y0, y0))
            if x0 - 2 <= x <= x1 + 2 and overlap > 1:
                xs_by_position[round(x)] += 1
        xs = sorted(float(x) for x, count in xs_by_position.items() if count >= 2 or abs(x - x0) <= 2 or abs(x - x1) <= 2)
        if len(xs) < 3:
            continue
        try:
            finder = page.find_tables(
                vertical_strategy="explicit",
                horizontal_strategy="explicit",
                vertical_lines=xs,
                horizontal_lines=ys,
            )
        except Exception as exc:
            LOGGER.debug("Page %d explicit table recovery failed: %s", page.number + 1, exc)
            continue
        candidates = [table for table in finder.tables if fitz.Rect(table.bbox).intersects(fitz.Rect(x0, y0, x1, y1))]
        if not candidates:
            continue
        table = max(candidates, key=lambda item: fitz.Rect(item.bbox).get_area())
        rows = native_table_rows(page, table)
        if len(rows) >= 2 and max((len(row) for row in rows), default=0) >= 2:
            results.append(Element("table", fitz.Rect(table.bbox), rows, len(results)))
    return results


def detect_tables(page: fitz.Page) -> list[Element]:
    elements: list[Element] = caption_guided_ruled_tables(page)
    try:
        finder = page.find_tables()
    except Exception as exc:
        LOGGER.debug("Page %d table detection failed: %s", page.number + 1, exc)
        return elements
    for index, table in enumerate(finder.tables):
        # table.extract() rebuilds text from individual glyphs and can turn
        # identifiers such as add_control_points into "add control points _ _"
        # or transpose ligatures.  Cell geometry is reliable, so read native
        # word tokens inside each cell instead.
        bbox = fitz.Rect(table.bbox)
        if any((bbox & existing.bbox).get_area() >= bbox.get_area() * 0.35 for existing in elements):
            continue
        rows = native_table_rows(page, table)
        if len(rows) < 2 or max((len(row) for row in rows), default=0) < 2:
            continue
        elements.append(Element("table", bbox, rows, len(elements) + index))
    return elements


def combine_soft_mask(color: fitz.Pixmap, mask: fitz.Pixmap) -> fitz.Pixmap:
    """Attach a PDF soft mask, scaling it when encodings use different DPIs."""
    if color.width != mask.width or color.height != mask.height:
        mask = fitz.Pixmap(mask, color.width, color.height)
    if mask.colorspace is not None and mask.colorspace.n != 1:
        mask = fitz.Pixmap(fitz.csGRAY, mask)
    return fitz.Pixmap(color, mask)


def pixmap_png_bytes(pix: fitz.Pixmap) -> bytes:
    """Encode any color PDF pixmap through a PNG-compatible color space."""
    if pix.colorspace is None:
        raise ValueError("image pixmap has no color space")
    # Indexed and ICCBased gray-looking pixmaps can report one component but
    # MuPDF still refuses to encode them as PNG.  DeviceGray and DeviceRGB are
    # the only native PNG color spaces; normalize every other space to RGB.
    if pix.colorspace.name not in {"DeviceGray", "DeviceRGB"}:
        pix = fitz.Pixmap(fitz.csRGB, pix)
    return pix.tobytes("png")


def render_image_region(page: fitz.Page, bbox: Sequence[float]) -> bytes:
    """Render an image's page rectangle when its embedded stream is malformed."""
    clip = fitz.Rect(bbox) & page.rect
    if clip.is_empty or clip.width <= 0 or clip.height <= 0:
        raise ValueError("image rectangle is outside the page")
    pix = page.get_pixmap(
        matrix=fitz.Matrix(2, 2),
        clip=clip,
        colorspace=fitz.csRGB,
        alpha=False,
    )
    return pix.tobytes("png")


def save_image(
    doc: fitz.Document,
    page: fitz.Page,
    info: dict[str, Any],
    image_dir: Path,
    page_number: int,
    image_number: int,
) -> Path:
    xref = int(info.get("xref", 0))
    filename = f"page-{page_number:03d}-image-{image_number:02d}.png"
    destination = image_dir / filename
    if xref > 0:
        try:
            pix = fitz.Pixmap(doc, xref)
            smask = 0
            for row in page.get_images(full=True):
                if int(row[0]) == xref:
                    smask = int(row[1])
                    break
            if smask > 0 and not pix.alpha:
                pix = combine_soft_mask(pix, fitz.Pixmap(doc, smask))
            payload = pixmap_png_bytes(pix)
        except Exception as exc:
            LOGGER.debug(
                "Page %d image xref %d requires rendered fallback: %s",
                page_number,
                xref,
                exc,
            )
            payload = render_image_region(page, info["bbox"])
        destination.write_bytes(payload)
        return destination

    # Inline images without an xref are rendered from their exact PDF bounds.
    destination.write_bytes(render_image_region(page, info["bbox"]))
    return destination


def image_elements(doc: fitz.Document, page: fitz.Page, image_dir: Path, text_blocks: Sequence[dict[str, Any]]) -> list[Element]:
    elements: list[Element] = []
    candidates = sorted(page.get_image_info(xrefs=True), key=lambda info: (info["bbox"][1], info["bbox"][0]))
    number = 0
    for info in candidates:
        rect = fitz.Rect(info["bbox"])
        if rect.width < 24 or rect.height < 24 or rect.get_area() < 1200:
            continue
        number += 1
        try:
            path = save_image(doc, page, info, image_dir, page.number + 1, number)
        except Exception as exc:
            LOGGER.warning(
                "第 %d 页图片 %d 无法保存，已跳过：%s",
                page.number + 1,
                number,
                exc,
            )
            number -= 1
            continue
        caption = ""
        preceding: list[tuple[float, str]] = []
        following: list[tuple[float, str]] = []
        for block in text_blocks:
            bbox = fitz.Rect(block["bbox"])
            text = block_text(block)
            if FIGURE_CAPTION_RE.match(text):
                if 0 <= rect.y0 - bbox.y1 <= 55:
                    preceding.append((rect.y0 - bbox.y1, text))
                elif 0 <= bbox.y0 - rect.y1 <= 40:
                    following.append((bbox.y0 - rect.y1, text))
        near = preceding or following
        if near:
            caption = min(near, key=lambda item: item[0])[1]
        elements.append(Element("image", rect, {"path": path, "caption": caption}, number))
    return elements


def vector_figure_elements(
    page: fitz.Page,
    image_dir: Path,
    text_blocks: Sequence[dict[str, Any]],
    occupied_rects: Sequence[fitz.Rect],
) -> list[Element]:
    """Render captioned vector figures that have no embedded raster image."""
    drawings = page.get_drawings()
    if len(drawings) < 5:
        return []
    drawing_rects = [fitz.Rect(item["rect"]) for item in drawings if not fitz.Rect(item["rect"]).is_empty]
    elements: list[Element] = []
    figure_number = 0
    mid = page.rect.x0 + page.rect.width / 2

    for block in text_blocks:
        caption = block_text(block)
        if not FIGURE_CAPTION_RE.match(caption) or not caption.casefold().startswith("fig."):
            continue
        caption_rect = fitz.Rect(block["bbox"])
        if caption_rect.x1 <= mid + 5:
            column = fitz.Rect(page.rect.x0 + page.rect.width * 0.07, 0, mid - 6, caption_rect.y0)
        elif caption_rect.x0 >= mid - 5:
            column = fitz.Rect(mid + 6, 0, page.rect.x1 - page.rect.width * 0.07, caption_rect.y0)
        else:
            column = fitz.Rect(page.rect.x0 + page.rect.width * 0.06, 0, page.rect.x1 - page.rect.width * 0.06, caption_rect.y0)

        candidates = [
            rect
            for rect in drawing_rects
            if rect.y1 <= caption_rect.y0 + 3
            and 0 <= caption_rect.y0 - rect.y1 <= page.rect.height * 0.48
            and (rect & column).get_area() > 0
        ]
        if len(candidates) < 5:
            continue

        # Split drawings into vertically connected groups and select the group
        # nearest the caption. This keeps two stacked figures separate.
        groups: list[list[fitz.Rect]] = []
        for rect in sorted(candidates, key=lambda value: (value.y0, value.y1)):
            if not groups:
                groups.append([rect])
                continue
            group_end = max(item.y1 for item in groups[-1])
            if rect.y0 <= group_end + 20:
                groups[-1].append(rect)
            else:
                groups.append([rect])
        viable = [group for group in groups if len(group) >= 5]
        if not viable:
            continue
        group = min(viable, key=lambda values: caption_rect.y0 - max(item.y1 for item in values))
        union = fitz.Rect(group[0])
        for rect in group[1:]:
            union |= rect
        crop = fitz.Rect(column.x0, max(0, union.y0 - 4), column.x1, caption_rect.y0 - 4)
        if crop.height < 35 or crop.width < 80 or inside_any(crop, occupied_rects, threshold=0.5):
            continue

        figure_number += 1
        destination = image_dir / f"page-{page.number + 1:03d}-figure-{figure_number:02d}.png"
        pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5), clip=crop, alpha=False)
        pix.save(destination)
        elements.append(
            Element(
                "image",
                crop,
                {"path": destination, "caption": caption},
                1000 + figure_number,
            )
        )
    return elements


def dot_leader_block_markdown(block: dict[str, Any]) -> str | None:
    """Render a chapter's inline mini-TOC as one Markdown row per PDF line."""
    lines = plain_lines(block)
    matching = [line for line in lines if DOT_LEADER_RE.search(line)]
    if len(lines) < 2 or len(matching) < 2 or len(matching) / len(lines) < 0.7:
        return None
    return "\n".join(f"- **{apply_links(line)}**" for line in lines)


def list_markdown(text: str, level: int = 0) -> str | None:
    stripped = text.lstrip()
    if not stripped:
        return None
    indent = "    " * max(0, level)
    bullet = stripped[0]
    if bullet in BULLETS:
        return f"{indent}{BULLETS[bullet]} {stripped[1:].strip()}"
    if re.match(r"^\d+[.)]\s+", stripped):
        return indent + re.sub(r"^(\d+)[.)]\s+", r"\1. ", stripped)
    if re.match(r"^[a-zA-Z][.)]\s+", stripped):
        return f"{indent}- {stripped}"
    return None


def split_leading_academic_heading(block: dict[str, Any]) -> list[dict[str, Any]]:
    """Separate an IEEE-style section line from prose in the same PDF block."""
    lines = block.get("lines", [])
    if len(lines) < 2:
        return [block]
    first_text = clean_text("".join(str(span.get("text", "")) for span in lines[0].get("spans", [])))
    if not (ROMAN_SECTION_RE.fullmatch(first_text) or LETTER_SECTION_RE.fullmatch(first_text)):
        return [block]

    def make_block(selected_lines: list[dict[str, Any]]) -> dict[str, Any]:
        result = dict(block)
        result["lines"] = selected_lines
        rect = fitz.Rect(selected_lines[0]["bbox"])
        for line in selected_lines[1:]:
            rect |= fitz.Rect(line["bbox"])
        result["bbox"] = tuple(rect)
        return result

    return [make_block([lines[0]]), make_block(list(lines[1:]))]


def is_horizontal_text_block(block: dict[str, Any]) -> bool:
    lines = block.get("lines", [])
    if not lines:
        return True
    horizontal = 0
    for line in lines:
        direction = line.get("dir", (1.0, 0.0))
        if abs(float(direction[0])) >= 0.8:
            horizontal += 1
    return horizontal >= max(1, len(lines) / 2)


def has_two_columns(elements: Sequence[Element], page_rect: fitz.Rect) -> bool:
    mid = page_rect.x0 + page_rect.width / 2
    left_rects: list[fitz.Rect] = []
    right_rects: list[fitz.Rect] = []
    for element in elements:
        rect = element.bbox
        if rect.width > page_rect.width * 0.52 or rect.height < 3:
            continue
        if rect.x1 <= mid + page_rect.width * 0.015 and rect.x0 < mid - page_rect.width * 0.08:
            left_rects.append(rect)
        elif rect.x0 >= mid - page_rect.width * 0.015 and rect.x1 > mid + page_rect.width * 0.08:
            right_rects.append(rect)
    if not left_rects or not right_rects:
        return False
    # A whole column is often emitted as one tall PDF text block (references
    # are a common example), so element count alone cannot decide the layout.
    left_is_column = len(left_rects) >= 2 or max(rect.height for rect in left_rects) >= page_rect.height * 0.22
    right_is_column = len(right_rects) >= 2 or max(rect.height for rect in right_rects) >= page_rect.height * 0.22
    if not left_is_column or not right_is_column:
        return False
    left_edges = [rect.x1 for rect in left_rects]
    right_edges = [rect.x0 for rect in right_rects]
    left_boundary = sorted(left_edges)[len(left_edges) // 2]
    right_boundary = sorted(right_edges)[len(right_edges) // 2]
    return right_boundary - left_boundary >= page_rect.width * 0.008


def sort_elements_reading_order(elements: Sequence[Element], page_rect: fitz.Rect) -> list[Element]:
    """Sort single-column pages normally and double-column pages by reading flow."""
    normal = sorted(elements, key=lambda item: (round(item.bbox.y0, 1), round(item.bbox.x0, 1), item.kind, item.order))
    if not has_two_columns(normal, page_rect):
        return normal

    mid = page_rect.x0 + page_rect.width / 2
    spanning: list[Element] = []
    column_items: list[Element] = []
    for element in normal:
        rect = element.bbox
        crosses_gutter = rect.x0 < mid - page_rect.width * 0.025 and rect.x1 > mid + page_rect.width * 0.025
        if rect.width >= page_rect.width * 0.62 or crosses_gutter:
            spanning.append(element)
        else:
            column_items.append(element)

    ordered: list[Element] = []
    remaining = set(range(len(column_items)))

    def emit_band(limit_y: float) -> None:
        selected = [index for index in remaining if column_items[index].bbox.y0 < limit_y]
        left = [column_items[index] for index in selected if column_items[index].bbox.x0 + column_items[index].bbox.x1 < 2 * mid]
        right = [column_items[index] for index in selected if column_items[index].bbox.x0 + column_items[index].bbox.x1 >= 2 * mid]
        ordered.extend(sorted(left, key=lambda item: (item.bbox.y0, item.bbox.x0, item.order)))
        ordered.extend(sorted(right, key=lambda item: (item.bbox.y0, item.bbox.x0, item.order)))
        remaining.difference_update(selected)

    for element in sorted(spanning, key=lambda item: (item.bbox.y0, item.bbox.x0, item.order)):
        emit_band(element.bbox.y0)
        ordered.append(element)
    emit_band(float("inf"))
    return ordered


def page_elements(
    doc: fitz.Document,
    page: fitz.Page,
    image_dir: Path,
    repeated_margin_texts: set[str] | None = None,
) -> tuple[list[Element], list[dict[str, Any]]]:
    raw = page.get_text("dict", sort=True)
    text_blocks: list[dict[str, Any]] = []
    for block in raw.get("blocks", []):
        if block.get("type") != 0 or not block.get("lines") or not is_horizontal_text_block(block):
            continue
        text_blocks.extend(split_leading_academic_heading(block))
    images = image_elements(doc, page, image_dir, text_blocks)
    raster_rects = [element.bbox for element in images]
    vector_figures = vector_figure_elements(page, image_dir, text_blocks, raster_rects)
    images.extend(vector_figures)
    image_rects = [element.bbox for element in images]
    # Decorative cover lines can look like a large empty table.  Figures with
    # chart grids can too, so discard table candidates inside detected figures.
    tables = [] if page.number == 0 else detect_tables(page)
    tables = [element for element in tables if not inside_any(element.bbox, image_rects, threshold=0.35)]
    table_rects = [element.bbox for element in tables]
    elements: list[Element] = [*tables, *images]

    for index, block in enumerate(text_blocks):
        rect = fitz.Rect(block["bbox"])
        text = block_text(block)
        if is_header_or_footer(
            rect,
            page.rect,
            page.number + 1,
            text,
            repeated_margin_texts,
        ):
            continue
        if inside_any(rect, table_rects):
            continue
        # Text inside a raster image is normally a duplicate accessibility
        # layer.  Keep captions outside the image, drop overlapping duplicates.
        if inside_any(rect, image_rects, threshold=0.75):
            continue
        if text:
            elements.append(Element("text", rect, block, index))

    return sort_elements_reading_order(elements, page.rect), text_blocks


def render_page(
    doc: fitz.Document,
    page: fitz.Page,
    image_dir: Path,
    page_bookmarks: Sequence[Bookmark],
    matched: set[tuple[int, str]],
    report: DocumentReport,
    suppress_heading_fallback: bool = False,
    metadata_title: str = "",
    repeated_margin_texts: set[str] | None = None,
) -> list[str]:
    elements, _ = page_elements(doc, page, image_dir, repeated_margin_texts)
    two_column_layout = has_two_columns(elements, page.rect)
    output: list[str] = [f"<!-- PDF page {page.number + 1} -->"]
    pending_callout = False
    list_buffer: list[str] = []
    list_base_x: float | None = None
    ordered_bookmarks = sorted(
        page_bookmarks,
        key=lambda item: (float("inf") if item.y is None else item.y, item.level),
    )

    def flush_list() -> None:
        nonlocal list_base_x
        if list_buffer:
            output.append("\n".join(list_buffer))
            list_buffer.clear()
        list_base_x = None

    def emit_unrepresented_bookmarks_before(element: Element) -> None:
        """Keep bookmark headings even when table detection swallowed their text block."""
        current_text = ""
        if element.kind == "text":
            current_text = block_text(element.data)
        for bookmark in ordered_bookmarks:
            identity = (bookmark.page, bookmark.key)
            if identity in matched or bookmark.y is None:
                continue
            # Let the ordinary text-heading path preserve the actual element
            # when the bookmark destination lands on its matching block.
            if current_text and heading_match_key(current_text) == heading_match_key(bookmark.title):
                continue
            due_before = bookmark.y <= element.bbox.y0 + 4.0
            swallowed_by_table = (
                element.kind == "table"
                and element.bbox.y0 - 4.0 <= bookmark.y <= element.bbox.y1 + 4.0
            )
            if not due_before and not swallowed_by_table:
                continue
            flush_list()
            output.append(f"{'#' * bookmark.level} {bookmark.title}")
            report.headings[str(bookmark.level)] += 1
            matched.add(identity)

    for element in elements:
        emit_unrepresented_bookmarks_before(element)
        if element.kind == "table":
            flush_list()
            value = table_markdown(element.data)
            if value:
                output.append(value)
                report.tables += 1
            pending_callout = False
            continue

        if element.kind == "image":
            flush_list()
            relative = Path("images") / element.data["path"].name
            alt = figure_alt_text(element.data["caption"], page.number + 1, element.order)
            output.append(f"![{alt}]({relative.as_posix()})")
            report.images += 1
            pending_callout = False
            continue

        block = element.data
        text = block_text(block)
        if not text:
            continue
        # A bookmark may already have been emitted from its destination
        # coordinate because the PDF draws the chapter number and title as
        # separate blocks.  Suppress those nearby source fragments so the
        # canonical bookmark heading is not followed by a duplicate fallback
        # heading or a standalone chapter number.
        represented_here = False
        for represented in page_bookmarks:
            identity = (represented.page, represented.key)
            if identity not in matched:
                continue
            nearby = represented.y is None or abs(represented.y - element.bbox.y0) <= 48.0
            same_title = heading_match_key(represented.title) == heading_match_key(text)
            leading_number = re.match(r"^(\d+(?:\.\d+)*)\b", represented.title)
            same_number = bool(leading_number and text.strip() == leading_number.group(1))
            if same_title or (nearby and same_number):
                represented_here = True
                break
        if represented_here:
            continue
        if (
            page.number == 0
            and metadata_title
            and element.bbox.y0 < page.rect.height * 0.16
            and normalized_key(text) in normalized_key(metadata_title)
            and len(text) >= 12
        ):
            continue
        size, bold, _, monospaced, _ = block_font_info(block)
        heading_level, bookmark = detect_heading(
            text,
            block,
            page_bookmarks,
            matched,
            page.number + 1,
            suppress_fallback=suppress_heading_fallback,
        )
        if heading_level:
            flush_list()
            heading_source = bookmark.title if bookmark else text
            clean_heading = re.sub(r"\.{3,}\s*[A-Z0-9-]+$", "", heading_source).strip()
            output.append(f"{'#' * heading_level} {clean_heading}")
            report.headings[str(heading_level)] += 1
            if bookmark:
                matched.add((bookmark.page, bookmark.key))
            pending_callout = False
            continue

        if suppress_heading_fallback:
            flush_list()
            listing = listing_block_markdown(block, page.rect)
            if listing:
                output.append(listing)
                report.text_characters += len(text)
            pending_callout = False
            continue

        lines = plain_lines(block)
        is_code = monospaced or (bold and len(lines) == 1 and PROMPT_RE.match(lines[0]))
        if is_code:
            flush_list()
            output.append(code_markdown(block))
            report.code_blocks += 1
            pending_callout = False
            continue

        equation = display_equation_markdown(block)
        if equation:
            flush_list()
            output.append(equation)
            report.text_characters += len(text)
            pending_callout = False
            continue

        label = text.rstrip(":").casefold()
        if label in CALLOUT_LABELS and bold and len(text.split()) <= 2:
            flush_list()
            output.append(f"> **{text.rstrip(':')}**")
            pending_callout = True
            continue

        formatted = inline_markdown(block)
        if not formatted:
            continue
        report.text_characters += len(text)

        mini_toc = dot_leader_block_markdown(block)
        if mini_toc:
            flush_list()
            output.append(mini_toc)
            pending_callout = False
            continue

        references = reference_markdown_parts(block)
        if references:
            flush_list()
            output.extend(references)
            pending_callout = False
            continue

        academic_list = academic_numbered_list_markdown(block)
        if academic_list:
            flush_list()
            output.append(academic_list)
            pending_callout = False
            continue

        unindented_list_value = list_markdown(text)
        list_value = None
        if unindented_list_value:
            if list_base_x is None:
                list_base_x = element.bbox.x0
            indent_step = max(14.0, page.rect.width * (22.0 / 612.0))
            level = max(0, min(6, round((element.bbox.x0 - list_base_x) / indent_step)))
            list_value = list_markdown(text, level=level)
        if list_value:
            list_buffer.append(apply_links(list_value))
            pending_callout = False
        elif pending_callout:
            flush_list()
            output.append("> " + formatted.replace("\n", "\n> "))
            pending_callout = False
        elif FIGURE_CAPTION_RE.match(text) or TABLE_CAPTION_RE.match(text):
            flush_list()
            output.append(f"**{apply_links(text)}**")
        else:
            flush_list()
            output.extend(paragraph_markdown_parts(block, page.rect, two_column_layout) or [formatted])

    flush_list()
    # A bookmark can target blank space or content consumed by a page-wide
    # vector/table region.  Preserve it at the page boundary rather than lose
    # the PDF's explicit hierarchy.
    for bookmark in ordered_bookmarks:
        identity = (bookmark.page, bookmark.key)
        if identity in matched:
            continue
        output.append(f"{'#' * bookmark.level} {bookmark.title}")
        report.headings[str(bookmark.level)] += 1
        matched.add(identity)
    return output


def page_image_coverage(page: fitz.Page) -> float:
    """Estimate how much of a page is covered by embedded raster images."""
    page_area = page.rect.width * page.rect.height
    if page_area <= 0:
        return 0.0

    areas: list[float] = []
    try:
        image_infos = page.get_image_info()
    except (AttributeError, RuntimeError, ValueError):
        image_infos = []
    for info in image_infos:
        try:
            bbox = fitz.Rect(info["bbox"]) & page.rect
        except (KeyError, TypeError, ValueError):
            continue
        if not bbox.is_empty:
            areas.append(bbox.width * bbox.height / page_area)

    # A scan is usually one full-page image, but some producers split it into
    # tiles.  The capped sum catches tiled pages; the maximum catches the usual
    # single-image case.  Exact overlap is unnecessary for this guard because
    # it is combined with the invisible-text test below.
    return min(1.0, max(areas, default=0.0, ) if len(areas) == 1 else sum(areas))


def page_text_layer_counts(page: fitz.Page) -> tuple[int, int]:
    """Return non-space visible and PDF ``ignore-text`` character counts."""
    visible = 0
    ignored = 0
    try:
        traces = page.get_texttrace()
    except (AttributeError, RuntimeError, ValueError):
        traces = []
    for span in traces:
        target = "ignored" if span.get("type") == 3 else "visible"
        count = 0
        for char in span.get("chars", []):
            try:
                value = chr(char[0])
            except (IndexError, TypeError, ValueError):
                value = ""
            if value and not value.isspace():
                count += 1
        if target == "ignored":
            ignored += count
        else:
            visible += count
    return visible, ignored


def validate_text_layer(doc: fitz.Document, minimum_chars_per_page: int = 20) -> int:
    total = 0
    pages_with_text = 0
    hidden_ocr_pages: list[int] = []
    scanned_image_pages: list[int] = []
    for page in doc:
        extracted_count = len(re.sub(r"\s+", "", page.get_text("text")))
        visible_count, ignored_count = page_text_layer_counts(page)
        image_coverage = page_image_coverage(page)

        # A full-page raster image with nearly all text set to PDF rendering
        # mode 3 (invisible / ignore-text) is a scanned page carrying a hidden
        # OCR layer.  PyMuPDF's normal get_text() includes that layer, so merely
        # checking extracted_count would wrongly accept OCR output as native
        # document text.
        mostly_ignored = ignored_count >= minimum_chars_per_page and (
            visible_count < minimum_chars_per_page
            or visible_count <= ignored_count * 0.05
        )
        if image_coverage >= 0.8 and mostly_ignored:
            hidden_ocr_pages.append(page.number + 1)
            continue

        # Do not count invisible text toward the native-text threshold.  For
        # conventional PDFs get_texttrace() reports the same visible text that
        # get_text() returns; the fallback keeps compatibility with rare files
        # whose producer prevents text tracing.
        count = min(extracted_count, visible_count) if visible_count else 0
        if not ignored_count and not visible_count and extracted_count:
            count = extracted_count
        # Routing remains PDF-granular.  Detecting even one high-confidence
        # full-page scan therefore routes the complete PDF to MinerU instead
        # of mixing native and OCR output inside one Markdown document.
        if image_coverage >= 0.8 and count < minimum_chars_per_page:
            scanned_image_pages.append(page.number + 1)
            continue
        total += count
        if count >= minimum_chars_per_page:
            pages_with_text += 1

    if hidden_ocr_pages:
        page_list = ", ".join(str(number) for number in hidden_ocr_pages[:12])
        suffix = "……" if len(hidden_ocr_pages) > 12 else ""
        raise ConversionError(
            "PDF 的以下页面是整页扫描图，并且只有隐藏 OCR 文字层："
            f"{page_list}{suffix}。这些页面需要重新 OCR，不能把隐藏 OCR 结果冒充为原生文字。"
        )
    if scanned_image_pages:
        page_list = ", ".join(str(number) for number in scanned_image_pages[:12])
        suffix = "……" if len(scanned_image_pages) > 12 else ""
        raise ConversionError(
            "PDF 的以下页面检测为整页扫描图："
            f"{page_list}{suffix}。处理以 PDF 为单位，因此整份 PDF 需要使用 OCR。"
        )
    required = max(1, int(doc.page_count * 0.2))
    if total < 100 or pages_with_text < required:
        raise ConversionError(
            "PDF 没有足够的可见原生文字层，需要 OCR 才能转换扫描图片中的文字。"
        )
    return total


def read_dotenv_value(path: Path, keys: Sequence[str]) -> str | None:
    """Read selected values from a dotenv file without changing the process environment."""
    try:
        lines = path.read_text(encoding="utf-8-sig").splitlines()
    except OSError:
        return None
    accepted = set(keys)
    assignment = re.compile(r"^\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$")
    for raw_line in lines:
        if not raw_line.strip() or raw_line.lstrip().startswith("#"):
            continue
        match = assignment.match(raw_line)
        if not match or match.group(1) not in accepted:
            continue
        value = match.group(2).strip()
        if value.startswith(('"', "'")):
            closing_quote = value.find(value[0], 1)
            if closing_quote > 0:
                value = value[1:closing_quote]
        else:
            value = re.split(r"\s+#", value, maxsplit=1)[0].strip()
        value = value.removeprefix("Bearer ").strip()
        if value:
            return value
    return None


def find_mineru_token(token_file: Path | None, input_path: Path) -> str | None:
    """Load a MinerU token without putting it in source code or command logs."""
    candidates: list[Path] = []
    if token_file is not None:
        explicit = token_file.expanduser().resolve()
        if not explicit.is_file():
            raise ConversionError(f"MinerU Token 文件不存在：{explicit}")
        try:
            value = explicit.read_text(encoding="utf-8").strip()
        except OSError as exc:
            raise ConversionError(f"无法读取 MinerU Token 文件：{explicit}：{exc}") from exc
        if not value:
            raise ConversionError(f"MinerU Token 文件为空：{explicit}")
        return value.removeprefix("Bearer ").strip()

    for key in ("MINERU_TOKEN", "MINERU_API_TOKEN"):
        value = os.environ.get(key, "").strip()
        if value:
            return value.removeprefix("Bearer ").strip()

    script_dir = Path(__file__).resolve().parent
    input_dir = input_path if input_path.is_dir() else input_path.parent

    # The launchers keep the virtual environment isolated but previously did
    # not load the project's .env file.  Read it here so Windows batch,
    # PowerShell, Linux/macOS, single-file and directory runs behave alike.
    dotenv_candidates = [
        script_dir / ".env",
        script_dir.parent / ".env",
        Path.cwd() / ".env",
        input_dir / ".env",
    ]
    seen_dotenv: set[Path] = set()
    for dotenv_path in dotenv_candidates:
        dotenv_path = dotenv_path.expanduser().resolve()
        if dotenv_path in seen_dotenv:
            continue
        seen_dotenv.add(dotenv_path)
        if not dotenv_path.is_file():
            continue
        value = read_dotenv_value(dotenv_path, ("MINERU_TOKEN", "MINERU_API_TOKEN"))
        if value:
            return value

    candidates.extend(
        [
            script_dir / "mineru_token.txt",
            script_dir.parent / "mineru_token.txt",
            input_dir / "mineru_token.txt",
        ]
    )
    seen: set[Path] = set()
    for candidate in candidates:
        candidate = candidate.resolve()
        if candidate in seen:
            continue
        seen.add(candidate)
        if not candidate.is_file():
            continue
        try:
            value = candidate.read_text(encoding="utf-8").strip()
        except OSError as exc:
            continue
        if value:
            return value.removeprefix("Bearer ").strip()
    return None


def locate_mineru_markdown(
    output_dir: Path,
    expected_name: str,
    metadata: dict[str, Any] | None = None,
) -> Path:
    """Locate MinerU's merged Markdown and normalize it to the final name."""
    expected = output_dir / expected_name
    if expected.is_file():
        return expected

    candidates: list[Path] = []
    metadata_name = str((metadata or {}).get("markdown", "")).strip()
    if metadata_name:
        metadata_path = output_dir / Path(metadata_name).name
        if metadata_path.is_file():
            candidates.append(metadata_path)
    candidates.extend(path for path in output_dir.glob("*.md") if path.is_file())

    unique_candidates: list[Path] = []
    seen: set[Path] = set()
    for candidate in candidates:
        resolved = candidate.resolve()
        if resolved in seen:
            continue
        seen.add(resolved)
        unique_candidates.append(candidate)
    if not unique_candidates:
        raise ConversionError(
            f"MinerU OCR 已完成，但结果目录中没有 Markdown 文件：{output_dir}"
        )
    if len(unique_candidates) > 1:
        names = ", ".join(path.name for path in unique_candidates[:6])
        raise ConversionError(
            f"MinerU OCR 结果中存在多个 Markdown 文件，无法确定主文件：{names}"
        )

    unique_candidates[0].replace(expected)
    return expected


def convert_pdf_with_mineru(
    source: Path,
    output_root: Path,
    *,
    token: str,
    overwrite: bool,
    write_report: bool,
    native_text_rejection: str,
    model_version: str = "vlm",
    language: str = "en",
    max_api_file_mb: float = 190.0,
    max_pages_per_chunk: int = 180,
) -> DocumentReport:
    """Convert a scanned / non-native PDF through MinerU's precise OCR API."""
    try:
        from mineru_pdf_to_md import (
            IMAGE_SUFFIXES,
            MineruApi,
            MineruError,
            PdfItem,
            PdfProcessor,
        )
    except ImportError as exc:
        raise ConversionError(
            "需要 MinerU OCR，但缺少 requests 或 pypdf；请重新安装 requirements.txt。"
        ) from exc

    folder_name = safe_name(source.stem)
    destination = output_root.expanduser().resolve() / folder_name
    if destination.exists() and not overwrite:
        raise ConversionError(f"输出目录已存在：{destination}；使用 --overwrite 可覆盖。")
    destination.parent.mkdir(parents=True, exist_ok=True)
    if not 0 < max_api_file_mb < 200:
        raise ConversionError("MinerU 单分片大小必须在 0 到 200 MiB 之间。")
    if not 0 < max_pages_per_chunk <= 200:
        raise ConversionError("MinerU 单分片页数必须在 1 到 200 页之间。")

    # Keep resumable state outside the final per-PDF folder.  The cache key
    # changes when the source size or modification time changes, preventing a
    # stale completed result from being reused for a replaced PDF.
    stat = source.stat()
    cache_key = hashlib.sha256(
        f"{source.resolve()}::{stat.st_size}::{stat.st_mtime_ns}".encode("utf-8", errors="replace")
    ).hexdigest()[:16]
    cache_root = destination.parent / ".mineru_cache"
    temp_output = cache_root / f"{folder_name}-{cache_key}"
    temp_output.mkdir(parents=True, exist_ok=True)

    try:
        api = MineruApi(
            token,
            model_version=model_version,
            language=language,
            enable_table=True,
            enable_formula=True,
            is_ocr=True,
        )
        processor = PdfProcessor(
            api,
            max_bytes=int(max_api_file_mb * 1024 * 1024),
            max_pages=max_pages_per_chunk,
            poll_interval=5.0,
            poll_timeout=21600.0,
            task_retries=3,
            retry_base_seconds=2.0,
        )
        item = PdfItem(source=source, relative_path=Path(source.name), output_dir=temp_output)
        processor.process(item, overwrite=False)

        meta_path = temp_output / ".mineru_meta.json"
        try:
            mineru_meta = json.loads(meta_path.read_text(encoding="utf-8")) if meta_path.is_file() else {}
        except (OSError, json.JSONDecodeError):
            mineru_meta = {}
        chunk_count = max(1, int(mineru_meta.get("chunks", 1)))

        # PdfProcessor writes resumable metadata while working.  A successful
        # final PDF folder follows this tool's stricter one-md-plus-images
        # contract, so operational files are removed before the atomic move.
        for operational in (".mineru_meta.json", ".mineru_state.json"):
            (temp_output / operational).unlink(missing_ok=True)
        shutil.rmtree(temp_output / ".mineru_work", ignore_errors=True)
        shutil.rmtree(temp_output / ".mineru_extract", ignore_errors=True)

        markdown_path = locate_mineru_markdown(
            temp_output,
            f"{folder_name}.md",
            mineru_meta,
        )
        markdown_text = markdown_path.read_text(encoding="utf-8", errors="replace")
        with fitz.open(source) as doc:
            page_count = doc.page_count
        image_count = sum(
            1
            for path in (temp_output / "images").glob("*")
            if path.is_file() and path.suffix.casefold() in IMAGE_SUFFIXES
        )
        heading_counts: dict[str, int] = defaultdict(int)
        for match in re.finditer(r"^(#{1,6})\s+\S", markdown_text, flags=re.MULTILINE):
            heading_counts[str(len(match.group(1)))] += 1
        report = DocumentReport(
            source=str(source),
            output=str(destination / f"{folder_name}.md"),
            pages=page_count,
            text_characters=len(re.sub(r"\s+", "", markdown_text)),
            headings=heading_counts,
            images=image_count,
            tables=sum(1 for line in markdown_text.splitlines() if re.match(r"^\|?\s*:?-{3,}", line)),
            ocr_used=True,
            conversion_engine=f"mineru-{model_version}",
            native_text_rejection=native_text_rejection,
            chunks=chunk_count,
        )
        atomic_replace_dir(temp_output, destination, overwrite=overwrite)
        if write_report:
            report_dir = destination.parent / "_reports"
            report_dir.mkdir(parents=True, exist_ok=True)
            (report_dir / f"{folder_name}.json").write_text(
                json.dumps(report.as_dict(), ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
        if cache_root.exists() and not any(cache_root.iterdir()):
            cache_root.rmdir()
        return report
    except MineruError as exc:
        raise ConversionError(
            f"MinerU OCR 转换失败：{exc}。断点保存在：{temp_output}，下次运行会继续。"
        ) from exc


FENCED_CODE_RE = re.compile(r"^```([^\n]*)\n([\s\S]*?)\n```$", re.MULTILINE)
PAGE_COMMENT_RE = re.compile(r"^<!-- PDF page \d+ -->$")


def merge_adjacent_code_parts(parts: Iterable[str]) -> list[str]:
    """Merge consecutive fenced blocks, including continuations across pages."""
    merged: list[str] = []
    pending_page_comments: list[str] = []

    for original in parts:
        part = original.strip("\r\n")
        if not part.strip():
            continue
        if PAGE_COMMENT_RE.fullmatch(part):
            if merged and FENCED_CODE_RE.fullmatch(merged[-1]):
                pending_page_comments.append(part)
            else:
                merged.append(part)
            continue

        current_code = FENCED_CODE_RE.fullmatch(part)
        previous_code = FENCED_CODE_RE.fullmatch(merged[-1]) if merged else None
        if current_code and previous_code:
            previous_language, previous_body = previous_code.groups()
            current_language, current_body = current_code.groups()
            language = previous_language if previous_language == current_language else "text"
            merged[-1] = (
                f"```{language}\n"
                f"{previous_body.rstrip()}\n"
                f"{current_body}\n"
                "```"
            )
            # A hidden page marker between two pieces of one code listing is
            # less useful than a valid uninterrupted fenced block.
            pending_page_comments.clear()
            continue

        if pending_page_comments:
            merged.extend(pending_page_comments)
            pending_page_comments.clear()
        merged.append(part)

    if pending_page_comments:
        merged.extend(pending_page_comments)
    return merged


def join_markdown_parts(parts: Iterable[str]) -> str:
    """Join blocks without deleting meaningful list/code indentation."""
    cleaned = merge_adjacent_code_parts(parts)
    markdown = "\n\n".join(cleaned)
    # PyMuPDF may split an inline summation at a column block boundary: the
    # operator remains at the end of one block and its lower limit begins the
    # next. Rejoin that native math sequence before final output.
    markdown = re.sub(
        r"\$\\sum\$\n\n\$_\{([^}]+)\}(\\pi_\{[^}]+\})\$\.",
        r"$\\sum_{\1} \2$.",
        markdown,
    )
    markdown = re.sub(
        r"\$\\sqrt\$\n\n\$([^$\n]+)\$",
        r"$\\sqrt{\1}$",
        markdown,
    )
    return markdown.rstrip() + "\n"


def heading_fallback_suppressed_pages(bookmarks: Sequence[Bookmark], page_count: int) -> set[int]:
    """Return front-matter listing pages where styled entries are not headings."""
    top_level = [item for item in bookmarks if item.level == 1]
    suppressed: set[int] = set()
    listing_titles = {"table of contents", "list of figures", "list of tables"}
    for index, item in enumerate(top_level):
        if normalized_key(item.title) not in listing_titles:
            continue
        next_page = top_level[index + 1].page if index + 1 < len(top_level) else page_count + 1
        suppressed.update(range(item.page, max(item.page + 1, next_page)))
    return suppressed


def atomic_replace_dir(source: Path, destination: Path, overwrite: bool) -> None:
    if destination.exists():
        if not overwrite:
            raise ConversionError(f"输出目录已存在：{destination}；使用 --overwrite 可覆盖。")
        shutil.rmtree(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    source.replace(destination)


def convert_pdf(
    source: Path,
    output_root: Path,
    overwrite: bool = False,
    write_report: bool = False,
    *,
    mineru_token: str | None = None,
    mineru_model: str = "vlm",
    mineru_language: str = "en",
    route: str = "auto",
    max_api_file_mb: float = 190.0,
    max_pages_per_chunk: int = 180,
) -> DocumentReport:
    source = source.expanduser().resolve()
    if not source.is_file() or source.suffix.casefold() != ".pdf":
        raise ConversionError(f"不是有效的 PDF 文件：{source}")

    if route not in {"auto", "native", "mineru-ocr"}:
        raise ConversionError(f"未知处理路由：{route}")

    native_rejection: str | None = None
    if route != "mineru-ocr":
        # Classify before creating an output directory. Native visible text
        # stays local; scanned pages and hidden OCR layers are sent to MinerU
        # only in auto mode.
        with fitz.open(source) as probe_doc:
            try:
                validate_text_layer(probe_doc)
            except ConversionError as native_error:
                native_rejection = str(native_error)
                if route == "native":
                    raise ConversionError(
                        f"已强制使用 native 路由，但该 PDF 不满足原生文字要求：{native_error}"
                    ) from native_error

    if route == "mineru-ocr" or native_rejection is not None:
        if not mineru_token:
            reason = native_rejection or "已强制使用 mineru-ocr 路由"
            raise ConversionError(
                f"{reason.rstrip('。.!！？')}。若要启用 MinerU OCR，请设置 MINERU_TOKEN，"
                "或使用 --mineru-token-file。"
            )
        return convert_pdf_with_mineru(
            source,
            output_root,
            token=mineru_token,
            overwrite=overwrite,
            write_report=write_report,
            native_text_rejection=native_rejection or "用户强制使用 mineru-ocr 路由",
            model_version=mineru_model,
            language=mineru_language,
            max_api_file_mb=max_api_file_mb,
            max_pages_per_chunk=max_pages_per_chunk,
        )

    folder_name = safe_name(source.stem)
    destination = output_root.expanduser().resolve() / folder_name
    destination.parent.mkdir(parents=True, exist_ok=True)
    # ``tempfile.mkdtemp`` may inherit restrictive ACLs in managed Windows
    # environments.  A normal workspace directory keeps the parent's ACLs.
    temp_dir = destination.parent / f".{folder_name}-{uuid.uuid4().hex}"
    temp_dir.mkdir(parents=False, exist_ok=False)
    final_temp = temp_dir / folder_name
    image_dir = final_temp / "images"
    final_temp.mkdir(parents=True, exist_ok=True)
    image_dir.mkdir(parents=True, exist_ok=True)

    try:
        with fitz.open(source) as doc:
            text_count = validate_text_layer(doc)
            bookmarks, bookmarks_by_page = make_bookmarks(doc)
            suppress_fallback_pages = heading_fallback_suppressed_pages(bookmarks, doc.page_count)
            report = DocumentReport(
                source=str(source),
                output=str(destination / f"{folder_name}.md"),
                pages=doc.page_count,
                text_characters=text_count,
                bookmarks_total=len(bookmarks),
            )
            matched: set[tuple[int, str]] = set()
            markdown: list[str] = []
            metadata_title = clean_text(doc.metadata.get("title", "") if doc.metadata else "")
            repeated_margin_texts = detect_repeated_margin_texts(doc)
            if metadata_title and not bookmarks:
                markdown.append(f"# {metadata_title}")

            for page in doc:
                LOGGER.info("%s: page %d/%d", source.name, page.number + 1, doc.page_count)
                markdown.extend(
                    render_page(
                        doc,
                        page,
                        image_dir,
                        bookmarks_by_page.get(page.number + 1, []),
                        matched,
                        report,
                        suppress_heading_fallback=(page.number + 1 in suppress_fallback_pages),
                        metadata_title=metadata_title,
                        repeated_margin_texts=repeated_margin_texts,
                    )
                )

            report.bookmarks_matched = len(matched)
            report.unmatched_bookmarks = [
                {"level": item.level, "title": item.title, "page": item.page}
                for item in bookmarks
                if (item.page, item.key) not in matched
            ]

        markdown_text = join_markdown_parts(markdown)
        markdown_path = final_temp / f"{folder_name}.md"
        markdown_path.write_text(markdown_text, encoding="utf-8", newline="\n")
        atomic_replace_dir(final_temp, destination, overwrite=overwrite)
        if write_report:
            # Keep the per-PDF folder contract exact: one Markdown file plus
            # images/.  Diagnostics live in a sibling root-level directory.
            report_dir = destination.parent / "_reports"
            report_dir.mkdir(parents=True, exist_ok=True)
            (report_dir / f"{folder_name}.json").write_text(
                json.dumps(report.as_dict(), ensure_ascii=False, indent=2) + "\n",
                encoding="utf-8",
            )
        shutil.rmtree(temp_dir, ignore_errors=True)
        return report
    except Exception:
        shutil.rmtree(temp_dir, ignore_errors=True)
        raise


def iter_pdfs(input_path: Path, recursive: bool = True) -> Iterator[Path]:
    """Yield every PDF; directory inputs are always traversed recursively."""
    if input_path.is_file():
        if input_path.suffix.casefold() == ".pdf":
            yield input_path
        return
    # ``recursive`` remains in the signature for API compatibility.  Folder
    # mode is intentionally recursive regardless of the old flag.
    yield from sorted(
        (path for path in input_path.rglob("*") if path.is_file() and path.suffix.casefold() == ".pdf"),
        key=lambda path: str(path).casefold(),
    )


def iter_documents(input_path: Path) -> Iterator[Path]:
    """Yield supported PDF and DOCX files recursively."""
    if input_path.is_file():
        if input_path.suffix.casefold() in SUPPORTED_INPUT_SUFFIXES:
            yield input_path
        return
    yield from sorted(
        (
            path
            for path in input_path.rglob("*")
            if path.is_file() and path.suffix.casefold() in SUPPORTED_INPUT_SUFFIXES
        ),
        key=lambda path: str(path).casefold(),
    )


def default_output_root(input_path: Path) -> Path:
    resolved = input_path.expanduser().resolve()
    if resolved.is_file():
        # Single-file mode: convert beside the PDF. ``convert_pdf`` adds the
        # PDF stem as its own folder, e.g. a/manual.pdf -> a/manual/manual.md.
        return resolved.parent
    # Folder mode: keep source files untouched and build a sibling mirror.
    return resolved.parent / f"{resolved.name}_md"


def output_root_for_pdf(input_path: Path, pdf: Path, batch_output_root: Path) -> Path:
    """Map one PDF to its parent output root under the required path rules."""
    input_path = input_path.expanduser().resolve()
    pdf = pdf.expanduser().resolve()
    if input_path.is_file():
        return batch_output_root.expanduser().resolve()
    relative_parent = pdf.relative_to(input_path).parent
    return batch_output_root.expanduser().resolve() / relative_parent


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "把 PDF 和 DOCX 转为结构化 Markdown；"
            "非原生文字 PDF 自动调用 MinerU OCR。"
        )
    )
    parser.add_argument("input", type=Path, help="单个受支持文档或包含文档的目录")
    parser.add_argument(
        "-o",
        "--output-dir",
        type=Path,
        help="兼容旧命令；单文件输出到 PDF 同目录，目录输入输出到同级 <目录名>_md",
    )
    parser.add_argument(
        "--route",
        choices=("auto", "native", "mineru-ocr"),
        default="auto",
        help="处理路由：自动判断、强制原生解析或强制 MinerU OCR",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="兼容旧命令；目录输入现在始终自动递归",
    )
    parser.add_argument("--overwrite", action="store_true", help="覆盖已存在的同名输出文件夹")
    parser.add_argument(
        "--report",
        action="store_true",
        help="在输出根目录的 _reports 中写入诊断报告",
    )
    parser.add_argument(
        "--mineru-token-file",
        type=Path,
        help="MinerU Token 文本文件；也可设置 MINERU_TOKEN 环境变量",
    )
    parser.add_argument(
        "--mineru-model",
        choices=("pipeline", "vlm"),
        default="vlm",
        help="OCR 使用的 MinerU 模型（默认 vlm）",
    )
    parser.add_argument(
        "--mineru-language",
        default="en",
        help="MinerU 文档语言；英文用 en，中文用 ch",
    )
    parser.add_argument(
        "--max-api-file-mb",
        type=float,
        default=190.0,
        help="MinerU 单分片安全大小上限，必须小于 200 MiB",
    )
    parser.add_argument(
        "--max-pages-per-chunk",
        type=int,
        default=180,
        help="MinerU 单分片页数上限，最大 200 页",
    )
    parser.add_argument(
        "--disable-mineru-ocr",
        action="store_true",
        help="禁用 MinerU 自动 OCR；非原生文字 PDF 将直接报错",
    )
    parser.add_argument("--verbose", action="store_true", help="显示逐页进度")
    return parser


def main(argv: Iterable[str] | None = None) -> int:
    # Pipes created by Windows launchers may otherwise inherit a legacy code
    # page and turn Chinese status messages into mojibake.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except (AttributeError, ValueError):
                pass
    args = build_parser().parse_args(argv)
    logging.basicConfig(
        level=logging.INFO if args.verbose else logging.WARNING,
        format="%(levelname)s: %(message)s",
    )
    input_path = args.input.expanduser().resolve()
    if not input_path.exists():
        print(f"错误：输入路径不存在：{input_path}", file=sys.stderr)
        return 2
    output_root = default_output_root(input_path)
    if args.output_dir and args.output_dir.expanduser().resolve() != output_root:
        print(
            "提示：-o 已忽略；输出固定写入："
            f"{output_root}"
        )
    documents = list(iter_documents(input_path))
    if not documents:
        print("错误：没有找到支持的 PDF 或 DOCX 文件。", file=sys.stderr)
        return 2

    if args.disable_mineru_ocr and args.route == "mineru-ocr":
        print("错误：--disable-mineru-ocr 与 --route mineru-ocr 不能同时使用。", file=sys.stderr)
        return 2
    effective_route = "native" if args.disable_mineru_ocr and args.route == "auto" else args.route
    mineru_token = None
    if effective_route != "native" and any(item.suffix.casefold() == ".pdf" for item in documents):
        mineru_token = find_mineru_token(args.mineru_token_file, input_path)

    failures = 0
    for index, document in enumerate(documents, 1):
        try:
            display_name = (
                str(document.relative_to(input_path)) if input_path.is_dir() else document.name
            )
            document_output_root = output_root_for_pdf(input_path, document, output_root)
            expected_markdown = (
                document_output_root
                / safe_name(document.stem)
                / f"{safe_name(document.stem)}.md"
            )
            if input_path.is_dir() and not args.overwrite and expected_markdown.is_file():
                print(
                    f"[{index}/{len(documents)}] 跳过 {display_name}"
                    "（输出已存在）"
                )
                continue
            print(f"[{index}/{len(documents)}] 转换 {display_name}")
            if document.suffix.casefold() == ".pdf":
                report = convert_pdf(
                    document,
                    document_output_root,
                    overwrite=args.overwrite,
                    write_report=args.report,
                    mineru_token=mineru_token,
                    mineru_model=args.mineru_model,
                    mineru_language=args.mineru_language,
                    route=effective_route,
                    max_api_file_mb=args.max_api_file_mb,
                    max_pages_per_chunk=args.max_pages_per_chunk,
                )
            else:
                report = convert_word_document(
                    document,
                    document_output_root,
                    overwrite=args.overwrite,
                    write_report=args.report,
                )
            page_label = f"{report.pages} 页" if report.pages else "页数未知"
            if str(report.conversion_engine).startswith("pandoc-"):
                heading_total = sum(int(value) for value in report.headings.values())
                print(
                    f"  完成：{report.output} | {page_label}，"
                    f"{report.images} 图，{report.tables} 表，"
                    f"Word 标题 {heading_total} 个，代码块 {report.code_blocks} 个，"
                    f"处理方式 {report.conversion_engine}"
                )
            else:
                print(
                    f"  完成：{report.output} | {page_label}，"
                    f"{report.images} 图，{report.tables} 表，"
                    f"书签标题 {report.bookmarks_matched}/{report.bookmarks_total} 匹配，"
                    f"处理方式 {report.conversion_engine}，分片 {report.chunks}"
                )
        except Exception as exc:
            failures += 1
            print(f"  失败：{exc}", file=sys.stderr)
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
