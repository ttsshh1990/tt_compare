#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import difflib
import json
import math
import os
import re
import shutil
import subprocess
import tempfile
import textwrap
import zipfile
from collections import Counter
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET

try:
    from playwright.sync_api import sync_playwright
except ImportError:  # pragma: no cover - optional dependency
    sync_playwright = None

try:
    from pypdf import PdfReader, PdfWriter
    from pypdf.annotations import Popup, Text
except ImportError:  # pragma: no cover - optional dependency
    PdfReader = None
    PdfWriter = None
    Popup = None
    Text = None

try:
    import fitz
except ImportError:  # pragma: no cover - optional dependency
    fitz = None


WORD_NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
BLOCK_TAGS = {
    "p",
    "li",
    "td",
    "th",
    "h1",
    "h2",
    "h3",
    "h4",
    "h5",
    "h6",
    "caption",
    "blockquote",
    "pre",
    "figcaption",
}
SKIP_TAGS = {
    "script",
    "style",
    "noscript",
    "template",
    "svg",
    "canvas",
    "head",
    "meta",
    "link",
    "iframe",
    "object",
}
INLINE_BOLD_TAGS = {"b", "strong"}
INLINE_ITALIC_TAGS = {"i", "em"}
INLINE_UNDERLINE_TAGS = {"u"}
DIFF_TOKEN_RE = re.compile(r"\d[\d,]*(?:\.\d+)?%?|\.\d+%?|[A-Za-z]+(?:[’'\-][A-Za-z]+)*")
CURRENCY_SYMBOL_RE = re.compile(r"([$€£¥])\s*(?=\(?\d)")
LEADING_FOOTNOTE_MARKER_RE = re.compile(r"^\s*\(?\d{1,2}\)?(?=\s+[A-Za-z])\s*")
TRAILING_FOOTNOTE_MARKER_RE = re.compile(r"(?<=[A-Za-z])\d{1,2}(?=\s*$)")
LEADING_ROW_PUNCTUATION_RE = re.compile(r"^[\s'\"`‘’“”\-–—•]+")
LEADING_MARKER_LINE_RE = re.compile(r"^\s*[-–—•]+\s*")
LEADING_LABEL_SYMBOL_RE = re.compile(r"^\s*[%\-–—•]+\s*(?=[A-Za-z])")
DATE_PHRASE_RE = re.compile(
    r"\b(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|"
    r"Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+\d{1,2},\s+\d{4}\b",
    flags=re.I,
)
WINDOWS_TESSERACT_CANDIDATES = [
    Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
    Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
]
OCR_RUNTIME_CONFIG = Path(__file__).resolve().with_name("ocr_runtime.json")


@dataclass
class Block:
    id: str
    source: str
    order: int
    text: str
    normalized: str
    heading: bool = False
    heading_level: int | None = None
    bold: bool = False
    italic: bool = False
    underline: bool = False
    list_item: bool = False
    table_cell: bool = False
    kind: str = ""
    table_pos: tuple[int, int, int] | None = None
    row_key: str | None = None
    row_slot: int | None = None
    numeric_slot: int | None = None


@dataclass
class Match:
    docx_index: int
    html_index: int
    match_type: str
    score: float
    formatting_diffs: list[str]


@dataclass
class DiffToken:
    text: str
    normalized: str
    kind: str
    start: int = 0
    end: int = 0
    spaces_before: int = 0
    prefix_symbol: str | None = None


@dataclass
class TokenRect:
    text: str
    normalized: str
    kind: str
    x: float
    y: float
    width: float
    height: float


@dataclass
class HtmlComment:
    order: int
    contents: str
    token_index: int | None = None


@dataclass
class BrowserRenderResult:
    blocks: list[Block]
    width_px: float
    height_px: float
    rects_by_order: dict[int, tuple[float, float, float, float]]
    token_rects_by_order: dict[int, list[TokenRect]]
    page_numbers_by_order: dict[int, int]
    coordinate_space: str = "browser_px"


def table_pos_key(block: Block) -> tuple[int, int, int] | None:
    return block.table_pos if block.table_cell and block.table_pos is not None else None


def table_index_key(block: Block) -> int | None:
    return block.table_pos[0] if block.table_cell and block.table_pos is not None else None


def row_context_key(block: Block) -> tuple[int, str, int] | None:
    if not block.table_cell or block.table_pos is None or not block.row_key or block.row_slot is None:
        return None
    table_idx, _row_idx, _col_idx = block.table_pos
    return (table_idx, block.row_key, block.row_slot)


def global_row_context_key(block: Block) -> tuple[str, int] | None:
    if not block.table_cell or not block.row_key or block.row_slot is None:
        return None
    return (block.row_key, block.row_slot)


def global_numeric_context_key(block: Block) -> tuple[str, int] | None:
    token = single_value_token(block)
    if not block.table_cell or not block.row_key or block.numeric_slot is None or token is None or token.kind != "number":
        return None
    return (block.row_key, block.numeric_slot)


def single_value_token(block: Block) -> DiffToken | None:
    tokens = diff_tokens(block.text)
    return tokens[0] if len(tokens) == 1 else None


def parse_numeric_token(token_text: str) -> tuple[float, bool] | None:
    text = normalize_text(token_text).strip()
    if not text:
        return None
    is_percent = "%" in text
    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]
    text = text.replace("$", "").replace("€", "").replace("£", "").replace("¥", "")
    text = text.replace("%", "").replace(",", "").replace("~", "").replace(" ", "")
    if text.startswith("-"):
        negative = True
        text = text[1:]
    if text.startswith("+"):
        text = text[1:]
    if not text:
        return None
    try:
        value = float(text)
    except ValueError:
        return None
    if negative:
        value = -value
    return value, is_percent


def numeric_value_compatibility(doc_token: DiffToken, target_token: DiffToken) -> float:
    doc_value = parse_numeric_token(doc_token.text)
    target_value = parse_numeric_token(target_token.text)
    if doc_value is None or target_value is None:
        return 0.0
    doc_number, doc_percent = doc_value
    target_number, target_percent = target_value
    if doc_percent != target_percent:
        return 0.0
    if doc_number == target_number:
        return 1.0
    if (doc_number < 0) != (target_number < 0):
        return 0.0
    doc_abs = abs(doc_number)
    target_abs = abs(target_number)
    if doc_abs > 0 and target_abs > 0:
        order_gap = abs(math.log10(doc_abs) - math.log10(target_abs))
        if order_gap > 1.0:
            return 0.0
    scale = max(doc_abs, target_abs, 1e-9)
    relative_gap = abs(doc_number - target_number) / max(scale, 1.0)
    return max(0.0, 1.0 - min(1.0, relative_gap * 0.5))


def table_context_match_score(doc_block: Block, html_block: Block) -> float:
    if normalize_for_compare(strip_leading_markers(doc_block.text)) == normalize_for_compare(strip_leading_markers(html_block.text)):
        return 1.0
    score = similarity(doc_block.normalized, html_block.normalized)
    doc_token = single_value_token(doc_block)
    html_token = single_value_token(html_block)
    if doc_token and html_token and doc_token.kind == html_token.kind:
        if doc_token.normalized == html_token.normalized:
            return 1.0
        if doc_token.kind == "number":
            compatibility = numeric_value_compatibility(doc_token, html_token)
            if compatibility <= 0.0:
                return score
            return max(score, compatibility)
        return max(score, 0.8)
    return score


def split_lead_label_text(text: str, *, table_cell: bool = False) -> tuple[str, str] | None:
    if table_cell:
        return None
    lines = [part.strip() for part in str(text or "").splitlines() if part.strip()]
    if len(lines) < 2:
        return None
    lead = lines[0]
    rest = " ".join(lines[1:]).strip()
    if not rest:
        return None
    lead_words = len([word for word in lead.split() if word])
    looks_like_label = len(lead) <= 64 and lead_words <= 8 and not re.search(r"[.!?:;]$", lead)
    body_substantial = len(rest) >= 80
    if looks_like_label and body_substantial:
        return lead, rest
    return None


def normalize_text(text: str) -> str:
    return (
        str(text or "")
        .replace("\u00A0", " ")
        .replace("\u2018", "'")
        .replace("\u2019", "'")
        .replace("\u201C", '"')
        .replace("\u201D", '"')
        .replace("\u2013", "-")
        .replace("\u2014", "-")
        .replace("\u2022", "-")
        .replace("&nbsp;", " ")
        .replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("$ ", "$")
        .replace("€ ", "€")
        .replace("£ ", "£")
        .replace("¥ ", "¥")
        .replace("~ ", "~")
    )


def strip_footnote_markers(text: str) -> str:
    text = normalize_text(text)
    text = LEADING_FOOTNOTE_MARKER_RE.sub("", text)
    if re.search(r"[A-Za-z]", text):
        text = re.sub(r"\(\s*\d{1,2}\s*\)\s*$", "", text)
        text = TRAILING_FOOTNOTE_MARKER_RE.sub("", text)
    return text.strip()


def normalize_for_compare(text: str) -> str:
    text = strip_footnote_markers(text)
    text = re.sub(r"\bwww\.", "", text, flags=re.I)
    text = re.sub(r"\(\s+(?=\d)", "(", text)
    text = re.sub(r"(\d)\s+%", r"\1%", text)
    text = re.sub(r"(\d)\s+\)", r"\1)", text)
    text = re.sub(r"([+-])\s+(?=\d)", r"\1", text)
    return re.sub(r"\s+", " ", text).strip().lower()


def normalize_row_key(text: str) -> str:
    text = strip_footnote_markers(text)
    text = LEADING_ROW_PUNCTUATION_RE.sub("", text)
    text = re.sub(r"\(\s*\d+\s*\)", "", text)
    text = re.sub(r"\[\s*\d+\s*\]", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return normalize_for_compare(text)


def strip_leading_markers(text: str) -> str:
    lines = normalize_text(text).splitlines() or [normalize_text(text)]
    cleaned: list[str] = []
    for line in lines:
        line = LEADING_MARKER_LINE_RE.sub("", line)
        if re.fullmatch(r"\s*%?\s*of total\s*", line, flags=re.I):
            line = re.sub(r"^\s*%+\s*", "", line)
        else:
            line = LEADING_LABEL_SYMBOL_RE.sub("", line)
        cleaned.append(line)
    return "\n".join(cleaned).strip()


def normalize_without_punctuation(text: str) -> str:
    cleaned = normalize_for_compare(strip_leading_markers(text))
    cleaned = re.sub(r"[^0-9a-z]+", " ", cleaned)
    return re.sub(r"\s+", " ", cleaned).strip()


def normalize_without_footnote_refs(text: str) -> str:
    cleaned = normalize_text(text)
    cleaned = re.sub(r"\(\s*\d+\s*\)", "", cleaned)
    cleaned = re.sub(r"(?<=[A-Za-z])\)", "", cleaned)
    cleaned = re.sub(r"\(\s*$", "", cleaned)
    return normalize_for_compare(cleaned)


def normalize_pdf_paragraph_artifacts(text: str) -> str:
    cleaned = normalize_text(text)
    cleaned = re.sub(r"([A-Za-z])-\s*\n\s*([A-Za-z])", r"\1-\2", cleaned)
    cleaned = re.sub(r"(?:\n\s*)+\d{1,2}\s*$", "", cleaned)
    return normalize_for_compare(cleaned)


def detect_dash_like_gap(
    left_rect: tuple[float, float, float, float],
    right_rect: tuple[float, float, float, float],
    *,
    page_width: float,
    page_height: float,
    pixmap: Any | None,
) -> bool:
    if pixmap is None:
        return False
    gap_width = right_rect[0] - left_rect[2]
    if gap_width < 2.0 or gap_width > 18.0:
        return False
    word_height = min(left_rect[3] - left_rect[1], right_rect[3] - right_rect[1])
    if word_height <= 0:
        return False
    center_y = ((left_rect[1] + left_rect[3]) / 2 + (right_rect[1] + right_rect[3]) / 2) / 2
    band_height = max(1.5, min(5.0, word_height * 0.35))
    x0 = max(0.0, left_rect[2] + 0.6)
    x1 = min(page_width, right_rect[0] - 0.6)
    y0 = max(0.0, center_y - band_height / 2)
    y1 = min(page_height, center_y + band_height / 2)
    if x1 <= x0 or y1 <= y0:
        return False

    scale_x = pixmap.width / max(page_width, 1.0)
    scale_y = pixmap.height / max(page_height, 1.0)
    px0 = max(0, min(pixmap.width - 1, int(x0 * scale_x)))
    px1 = max(0, min(pixmap.width, int(x1 * scale_x)))
    py0 = max(0, min(pixmap.height - 1, int(y0 * scale_y)))
    py1 = max(0, min(pixmap.height, int(y1 * scale_y)))
    if px1 - px0 < 2 or py1 - py0 < 1:
        return False

    samples = pixmap.samples
    stride = pixmap.stride
    threshold = 180
    longest_run = 0
    dark_rows = 0
    for py in range(py0, py1):
        row = samples[py * stride : py * stride + stride]
        current_run = 0
        row_longest = 0
        dark_pixels = 0
        for px in range(px0, px1):
            value = row[px]
            if value < threshold:
                dark_pixels += 1
                current_run += 1
                row_longest = max(row_longest, current_run)
            else:
                current_run = 0
        if dark_pixels:
            dark_rows += 1
        longest_run = max(longest_run, row_longest)

    width_px = px1 - px0
    height_px = py1 - py0
    if dark_rows == 0:
        return False
    return (
        longest_run >= max(3, int(width_px * 0.35))
        and dark_rows <= max(4, int(height_px * 0.75))
    )


def join_words_for_pdf_cluster(
    words: list[dict[str, Any]],
    *,
    page_width: float,
    page_height: float,
    pixmap: Any | None = None,
) -> str:
    if not words:
        return ""
    parts = [str(words[0]["text"]).strip()]
    prev = words[0]
    for word in words[1:]:
        separator = " "
        if detect_dash_like_gap(prev["rect"], word["rect"], page_width=page_width, page_height=page_height, pixmap=pixmap):
            separator = " - "
        parts.append(separator + str(word["text"]).strip())
        prev = word
    return "".join(parts).strip()


def diff_token_kind(token: str) -> str:
    return "number" if re.fullmatch(r"[+-]?(?:\d[\d,]*(?:\.\d+)?%?|\.\d+%?)", token) else "word"


def normalize_diff_token(token: str) -> str:
    token = normalize_text(token).strip()
    if diff_token_kind(token) == "number":
        return token.replace(",", "").lower()
    return token.lower()


def diff_tokens(text: str) -> list[DiffToken]:
    normalized_text = normalize_text(text)
    tokens: list[DiffToken] = []
    for match in DIFF_TOKEN_RE.finditer(normalized_text):
        cursor = match.start() - 1
        spaces_before = 0
        while cursor >= 0 and normalized_text[cursor].isspace():
            spaces_before += 1
            cursor -= 1
        prefix_symbol = None
        if cursor >= 0:
            candidate = normalized_text[cursor]
            if not candidate.isalnum():
                prefix_symbol = candidate
        token_text = match.group(0)
        token_end = match.end()
        if diff_token_kind(token_text) == "number" and not token_text.endswith("%"):
            suffix = normalized_text[token_end:]
            percent_match = re.match(r"(\s+)%(?=\s|$|[),.;:])", suffix)
            if percent_match:
                token_text = token_text + "%"
                token_end += percent_match.end()
        tokens.append(
            DiffToken(
                text=token_text,
                normalized=normalize_diff_token(token_text),
                kind=diff_token_kind(token_text),
                start=match.start(),
                end=token_end,
                spaces_before=spaces_before,
                prefix_symbol=prefix_symbol,
            )
        )
    return tokens


def extract_currency_symbol(text: str) -> str | None:
    match = CURRENCY_SYMBOL_RE.search(normalize_text(text))
    return match.group(1) if match else None


def semantic_currency_hint(block: Block) -> str | None:
    row_key = block.row_key or ""
    if not row_key:
        return None
    non_currency_row_markers = (
        "shares used in computing",
        "outstanding shares",
        "shares outstanding",
        "common stock",
    )
    if any(marker in row_key for marker in non_currency_row_markers):
        return None
    currency_row_markers = (
        " eps",
        "gaap eps",
        "non-gaap eps",
        "per diluted share",
        "earnings per diluted share",
        "net income per diluted share",
    )
    return "$" if any(marker in row_key for marker in currency_row_markers) else None


def table_context_currency_hint(block: Block, blocks: list[Block]) -> str | None:
    if not block.table_cell or block.table_pos is None:
        return None
    row_key = block.row_key or ""
    if not row_key:
        return None
    token = single_value_token(block)
    if token is None or token.kind != "number":
        return None
    parsed = parse_numeric_token(token.text)
    if parsed is not None:
        numeric_value, _is_percent = parsed
        if abs(numeric_value) < 100 and "," not in token.text:
            return None
    non_currency_row_markers = (
        "shares used in computing",
        "outstanding shares",
        "shares outstanding",
        "tax rate",
        "effective tax rate",
        "% of total",
        "operating margin",
    )
    if any(marker in row_key for marker in non_currency_row_markers):
        return None
    page_index, row_index, _col_index = block.table_pos
    nearby_percent = 0
    nearby_explicit = 0
    for peer in blocks:
        if not peer.table_cell or peer.table_pos is None:
            continue
        peer_page, peer_row, _ = peer.table_pos
        if peer_page != page_index or abs(peer_row - row_index) > 6:
            continue
        peer_token = single_value_token(peer)
        if peer_token is not None and peer_token.kind == "number" and peer_token.text.endswith("%"):
            nearby_percent += 1
        if extract_currency_symbol(peer.text):
            nearby_explicit += 1
    if nearby_percent >= 2:
        return None
    for peer in blocks:
        if not peer.table_cell or peer.table_pos is None:
            continue
        peer_page, peer_row, _ = peer.table_pos
        if peer_page != page_index or abs(peer_row - row_index) > 6:
            continue
        if extract_currency_symbol(peer.text):
            if nearby_explicit >= 1 and (
                semantic_currency_hint(block)
                or any(
                    marker in row_key
                    for marker in (
                        "revenue",
                        "expense",
                        "expenses",
                        "income",
                        "loss",
                        "amortization",
                        "compensation",
                        "charges",
                        "acquisition/divestiture",
                        "adjustment",
                        "cash flow",
                        "capital expenditures",
                        "debt",
                        "liabilities",
                        "assets",
                        "equity",
                    )
                )
            ):
                return "$"
    return None


def tesseract_from_runtime_config() -> Path | None:
    if not OCR_RUNTIME_CONFIG.exists():
        return None
    try:
        payload = json.loads(OCR_RUNTIME_CONFIG.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    configured = payload.get("tesseract_path")
    if not configured:
        return None
    configured_path = Path(configured)
    return configured_path if configured_path.exists() else None


def resolve_tesseract_executable() -> Path | None:
    configured_path = tesseract_from_runtime_config()
    if configured_path is not None:
        return configured_path
    command_path = shutil.which("tesseract")
    if command_path:
        return Path(command_path)
    for candidate in WINDOWS_TESSERACT_CANDIDATES:
        if candidate.exists():
            return candidate
    return None


def prepare_tesseract_environment(tesseract_path: Path) -> None:
    tesseract_dir = str(tesseract_path.parent)
    current_path = os.environ.get("PATH", "")
    path_parts = current_path.split(os.pathsep) if current_path else []
    if tesseract_dir not in path_parts:
        os.environ["PATH"] = tesseract_dir + (os.pathsep + current_path if current_path else "")

    tessdata_dir = tesseract_path.parent / "tessdata"
    if tessdata_dir.exists() and not os.environ.get("TESSDATA_PREFIX"):
        os.environ["TESSDATA_PREFIX"] = str(tessdata_dir)


def row_peer_blocks(block: Block, blocks: list[Block]) -> list[Block]:
    if not block.table_cell or block.table_pos is None:
        return []
    table_idx, row_idx, _col_idx = block.table_pos
    return [
        candidate
        for candidate in blocks
        if candidate is not block
        and candidate.table_cell
        and candidate.table_pos is not None
        and candidate.table_pos[0] == table_idx
        and candidate.table_pos[1] == row_idx
    ]


def inferred_row_currency_symbol(block: Block, blocks: list[Block]) -> str | None:
    for peer in [block] + row_peer_blocks(block, blocks):
        token = single_value_token(peer)
        if token is None or token.kind != "number":
            continue
        symbol = extract_currency_symbol(peer.text)
        if symbol:
            return symbol
    return None


def visible_meaningful(text: str) -> bool:
    normalized = normalize_for_compare(text)
    return bool(normalized and re.search(r"[a-z0-9]", normalized))


def tokenize(text: str) -> list[str]:
    cleaned = re.sub(r"[^0-9a-z]+", " ", normalize_for_compare(text))
    return [token for token in cleaned.split() if token]


def bigrams(text: str) -> list[str]:
    cleaned = normalize_for_compare(text).replace(" ", "")
    return [cleaned[i : i + 2] for i in range(len(cleaned) - 1)]


def jaccard_tokens(a: str, b: str) -> float:
    set_a = set(tokenize(a))
    set_b = set(tokenize(b))
    if not set_a and not set_b:
        return 1.0
    if not set_a or not set_b:
        return 0.0
    inter = len(set_a & set_b)
    union = len(set_a | set_b)
    return inter / union if union else 0.0


def dice_bigrams(a: str, b: str) -> float:
    grams_a = bigrams(a)
    grams_b = bigrams(b)
    if not grams_a and not grams_b:
        return 1.0
    if not grams_a or not grams_b:
        return 0.0
    counts: dict[str, int] = {}
    for gram in grams_a:
        counts[gram] = counts.get(gram, 0) + 1
    inter = 0
    for gram in grams_b:
        count = counts.get(gram, 0)
        if count:
            inter += 1
            counts[gram] = count - 1
    return (2 * inter) / (len(grams_a) + len(grams_b))


def similarity(a: str, b: str) -> float:
    len_ratio = min(len(a), len(b)) / max(len(a), len(b), 1)
    return (0.5 * jaccard_tokens(a, b)) + (0.35 * dice_bigrams(a, b)) + (0.15 * len_ratio)


def token_overlap_ratio(fragment: str, container: str) -> float:
    fragment_tokens = tokenize(fragment)
    container_tokens = tokenize(container)
    if not fragment_tokens:
        return 0.0
    overlap = Counter(fragment_tokens) & Counter(container_tokens)
    return sum(overlap.values()) / len(fragment_tokens)


def token_subsequence_ratio(fragment: str, container: str) -> float:
    fragment_tokens = tokenize(fragment)
    container_tokens = tokenize(container)
    if not fragment_tokens:
        return 0.0
    cursor = 0
    matched = 0
    for token in container_tokens:
        if cursor < len(fragment_tokens) and token == fragment_tokens[cursor]:
            matched += 1
            cursor += 1
    return matched / len(fragment_tokens)


def pdf_block_has_docx_anchor(block: Block, docx_blocks: list[Block]) -> bool:
    if not visible_meaningful(block.text):
        return False
    if len(tokenize(block.text)) < 3 and len(block.normalized) < 24:
        return False
    block_norm = normalize_without_footnote_refs(block.text)
    for doc_block in docx_blocks:
        doc_norm = normalize_without_footnote_refs(doc_block.text)
        if not doc_norm:
            continue
        if doc_norm == block_norm:
            return True
        if similarity(block_norm, doc_norm) >= 0.9:
            return True
        overlap = max(token_overlap_ratio(block_norm, doc_norm), token_overlap_ratio(doc_norm, block_norm))
        if overlap >= 0.82:
            return True
    return False


def short_fragment_subsequence_similarity(fragment_tokens: list[str], container_tokens: list[str]) -> float:
    if not fragment_tokens or not container_tokens:
        return 0.0
    fragment_text = " ".join(fragment_tokens)
    best_score = 0.0
    for window_len in range(max(1, len(fragment_tokens) - 1), min(len(container_tokens), len(fragment_tokens) + 2) + 1):
        for start in range(0, len(container_tokens) - window_len + 1):
            candidate_text = " ".join(container_tokens[start:start + window_len])
            score = difflib.SequenceMatcher(None, fragment_text, candidate_text).ratio()
            if score > best_score:
                best_score = score
                if best_score >= 0.999:
                    return best_score
    return best_score


def containment_match_score(fragment: Block, container: Block) -> float:
    if not fragment.normalized or not container.normalized:
        return 0.0
    if fragment.table_cell:
        return 0.0

    fragment_tokens = tokenize(fragment.normalized)
    container_tokens = tokenize(container.normalized)
    if len(fragment_tokens) <= 4:
        if fragment.normalized in container.normalized:
            return 0.99
        short_score = short_fragment_subsequence_similarity(fragment_tokens, container_tokens)
        if short_score >= 0.9:
            return short_score
        return 0.0

    if fragment.normalized in container.normalized:
        return 0.99
    if len(fragment_tokens) < 5:
        return 0.0

    overlap_ratio = token_overlap_ratio(fragment.normalized, container.normalized)
    if len(fragment_tokens) >= 8 and overlap_ratio >= 0.94:
        return 0.9 + min(0.09, (overlap_ratio - 0.94) * 1.5)
    if len(fragment_tokens) >= 20 and overlap_ratio >= 0.88:
        return 0.84 + min(0.05, (overlap_ratio - 0.88) * 1.2)
    return 0.0


def summarize_formatting_diff(docx: Block, html: Block) -> list[str]:
    diffs: list[str] = []
    checks = [
        ("heading", docx.heading, html.heading),
        ("bold", docx.bold, html.bold),
        ("italic", docx.italic, html.italic),
        ("underline", docx.underline, html.underline),
        ("list item", docx.list_item, html.list_item),
        ("table cell", docx.table_cell, html.table_cell),
    ]
    for label, doc_value, html_value in checks:
        if doc_value != html_value:
            diffs.append(
                f"DOCX {'has' if doc_value else 'does not have'} {label}; "
                f"HTML {'has' if html_value else 'does not have'} it."
            )
    if docx.heading and html.heading and docx.heading_level != html.heading_level:
        diffs.append(
            f"DOCX heading level is {docx.heading_level or '?'}; "
            f"HTML heading level is {html.heading_level or '?'}."
        )
    return diffs


def xml_local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def style_flag(style_text: str, needle: str) -> bool:
    return needle in style_text.lower()


def collect_word_text(paragraph: ET.Element) -> tuple[str, bool, bool, bool]:
    text_parts: list[str] = []
    bold = False
    italic = False
    underline = False
    for run in paragraph.findall(".//w:r", WORD_NS):
        for child in list(run):
            local = xml_local_name(child.tag) if isinstance(child.tag, str) else ""
            if local == "t":
                text_parts.append(child.text or "")
            elif local == "tab":
                text_parts.append("\t")
            elif local == "br":
                text_parts.append("\n")
        run_props = run.find("w:rPr", WORD_NS)
        if run_props is None:
            continue
        if run_props.find("w:b", WORD_NS) is not None:
            bold = True
        if run_props.find("w:i", WORD_NS) is not None:
            italic = True
        if run_props.find("w:u", WORD_NS) is not None:
            underline = True
    return "".join(text_parts), bold, italic, underline


def extract_docx_footnote_blocks(archive: zipfile.ZipFile, order_start: int) -> tuple[list[Block], int]:
    try:
        xml_text = archive.read("word/footnotes.xml")
    except KeyError:
        return [], order_start

    root = ET.fromstring(xml_text)
    blocks: list[Block] = []
    order = order_start
    for footnote in root.findall("w:footnote", WORD_NS):
        footnote_type = (
            footnote.get(f"{{{WORD_NS['w']}}}type")
            or footnote.get("w:type")
            or footnote.get("type")
            or ""
        )
        if footnote_type in {"separator", "continuationSeparator", "continuationNotice"}:
            continue
        for paragraph in footnote.findall(".//w:p", WORD_NS):
            text, bold, italic, underline = collect_word_text(paragraph)
            if not visible_meaningful(text):
                continue
            blocks.append(
                Block(
                    id=f"docx-{order}",
                    source="docx",
                    order=order,
                    text=normalize_text(text).strip(),
                    normalized=normalize_for_compare(text),
                    bold=bold,
                    italic=italic,
                    underline=underline,
                    kind="footnote",
                )
            )
            order += 1
    return blocks, order


def extract_docx_blocks(path: Path) -> list[Block]:
    with zipfile.ZipFile(path) as archive:
        xml_text = archive.read("word/document.xml")
        root = ET.fromstring(xml_text)
        body = root.find("w:body", WORD_NS)
        if body is None:
            return []

        blocks: list[Block] = []
        order = 0
        table_ordinal = 0
        for child in body:
            local = xml_local_name(child.tag)
            if local == "p":
                para_props = child.find("w:pPr", WORD_NS)
                style_val = ""
                list_item = False
                if para_props is not None:
                    style = para_props.find("w:pStyle", WORD_NS)
                    if style is not None:
                        style_val = (
                            style.get(f"{{{WORD_NS['w']}}}val", "")
                            or style.get("w:val", "")
                            or style.get("val", "")
                        )
                    list_item = para_props.find("w:numPr", WORD_NS) is not None
                text, bold, italic, underline = collect_word_text(child)
                if not visible_meaningful(text):
                    continue
                heading_match = re.search(r"heading\s*([1-6])?", style_val, flags=re.I)
                is_title = re.search(r"title", style_val, flags=re.I)
                heading = bool(heading_match or is_title)
                heading_level = int(heading_match.group(1)) if heading_match and heading_match.group(1) else 1 if is_title else None
                blocks.append(
                    Block(
                        id=f"docx-{order}",
                        source="docx",
                        order=order,
                        text=normalize_text(text).strip(),
                        normalized=normalize_for_compare(text),
                        heading=heading,
                        heading_level=heading_level,
                        bold=bold,
                        italic=italic,
                        underline=underline,
                        list_item=list_item,
                        kind="p",
                    )
                )
                order += 1
            elif local == "tbl":
                rows = child.findall(".//w:tr", WORD_NS)
                for row_index, row in enumerate(rows):
                    row_cells: list[tuple[int, str, bool, bool, bool]] = []
                    for col_index, cell in enumerate(row.findall("w:tc", WORD_NS)):
                        cell_parts: list[str] = []
                        bold = False
                        italic = False
                        underline = False
                        for paragraph in cell.findall(".//w:p", WORD_NS):
                            text, p_bold, p_italic, p_underline = collect_word_text(paragraph)
                            if visible_meaningful(text):
                                cell_parts.append(normalize_text(text).strip())
                            bold = bold or p_bold
                            italic = italic or p_italic
                            underline = underline or p_underline
                        cell_text = "\n".join(part for part in cell_parts if part)
                        row_cells.append((col_index, cell_text, bold, italic, underline))
                    row_key = next(
                        (normalize_row_key(text) for _col_index, text, _bold, _italic, _underline in row_cells if visible_meaningful(text)),
                        None,
                    )
                    row_slot = 0
                    numeric_slot = 0
                    for col_index, cell_text, bold, italic, underline in row_cells:
                        if not visible_meaningful(cell_text):
                            continue
                        token = single_value_token(
                            Block(
                                id="",
                                source="docx",
                                order=0,
                                text=cell_text,
                                normalized=normalize_for_compare(cell_text),
                            )
                        )
                        cell_numeric_slot = None
                        if token is not None and token.kind == "number":
                            cell_numeric_slot = numeric_slot
                            numeric_slot += 1
                        blocks.append(
                            Block(
                                id=f"docx-{order}",
                                source="docx",
                                order=order,
                                text=cell_text,
                                normalized=normalize_for_compare(cell_text),
                                bold=bold,
                                italic=italic,
                                underline=underline,
                                table_cell=True,
                                kind="td",
                                table_pos=(table_ordinal, row_index, col_index),
                                row_key=row_key,
                                row_slot=row_slot,
                                numeric_slot=cell_numeric_slot,
                            )
                        )
                        order += 1
                        row_slot += 1
                table_ordinal += 1

        footnote_blocks, order = extract_docx_footnote_blocks(archive, order)
        blocks.extend(footnote_blocks)
        return blocks


def get_descendant_text(element: ET.Element) -> str:
    parts: list[str] = []
    for text in element.itertext():
        parts.append(text)
    return normalize_text("".join(parts))


def node_has_descendant_block(element: ET.Element) -> bool:
    for child in list(element):
        if not isinstance(child.tag, str):
            continue
        if xml_local_name(child.tag).lower() in BLOCK_TAGS:
            return True
        if node_has_descendant_block(child):
            return True
    return False


def detect_inline_flags(element: ET.Element) -> tuple[bool, bool, bool]:
    bold = False
    italic = False
    underline = False
    for node in element.iter():
        if not isinstance(node.tag, str):
            continue
        tag = xml_local_name(node.tag).lower()
        style = (node.get("style") or "").lower()
        if tag in INLINE_BOLD_TAGS or style_flag(style, "font-weight: bold"):
            bold = True
        if tag in INLINE_ITALIC_TAGS or style_flag(style, "font-style: italic"):
            italic = True
        if tag in INLINE_UNDERLINE_TAGS or style_flag(style, "text-decoration: underline"):
            underline = True
    return bold, italic, underline


def is_hidden(element: ET.Element) -> bool:
    style = (element.get("style") or "").lower()
    hidden_attr = (element.get("hidden") or "").lower()
    return "display:none" in style or "visibility:hidden" in style or hidden_attr in {"hidden", "true"}


def annotate_html_parents(element: ET.Element, parent: ET.Element | None = None) -> None:
    setattr(element, "_parent", parent)
    for child in list(element):
        annotate_html_parents(child, element)


def nearest_html_ancestor(element: ET.Element, tag_name: str) -> ET.Element | None:
    current = getattr(element, "_parent", None)
    while current is not None:
        if isinstance(current.tag, str) and xml_local_name(current.tag).lower() == tag_name:
            return current
        current = getattr(current, "_parent", None)
    return None


def html_table_position(element: ET.Element, root: ET.Element) -> tuple[int, int, int] | None:
    tag = xml_local_name(element.tag).lower() if isinstance(element.tag, str) else ""
    if tag not in {"td", "th"}:
        return None
    tr = nearest_html_ancestor(element, "tr")
    table = nearest_html_ancestor(element, "table")
    if tr is None or table is None:
        return None
    tables = [node for node in root.iter() if isinstance(node.tag, str) and xml_local_name(node.tag).lower() == "table"]
    rows = [node for node in table.iter() if isinstance(node.tag, str) and xml_local_name(node.tag).lower() == "tr"]
    cells = [node for node in list(tr) if isinstance(node.tag, str) and xml_local_name(node.tag).lower() in {"td", "th"}]
    return (
        max(0, tables.index(table)),
        max(0, rows.index(tr)),
        max(0, cells.index(element)),
    )


def walk_html_blocks(element: ET.Element, blocks: list[Block], order_ref: list[int], hidden_ancestor: bool = False) -> None:
    if not isinstance(element.tag, str):
        return
    tag = xml_local_name(element.tag).lower()
    if tag in SKIP_TAGS:
        return

    hidden_here = hidden_ancestor or is_hidden(element)
    if hidden_here:
        return

    if tag in BLOCK_TAGS:
        capture_self = tag in {"td", "th", "li", "pre", "blockquote", "caption", "figcaption"}
        if not capture_self:
            capture_self = not node_has_descendant_block(element)
        if capture_self:
            text = get_descendant_text(element).strip()
            if visible_meaningful(text):
                bold, italic, underline = detect_inline_flags(element)
                heading_level = int(tag[1]) if re.fullmatch(r"h[1-6]", tag) else None
                table_cell = tag in {"td", "th"}
                split = split_lead_label_text(text, table_cell=table_cell)
                base_kwargs: dict[str, Any] = {
                    "source": "html",
                    "heading": heading_level is not None,
                    "heading_level": heading_level,
                    "bold": bold,
                    "italic": italic,
                    "underline": underline,
                    "list_item": tag == "li",
                    "table_cell": table_cell,
                    "kind": tag,
                    "table_pos": html_table_position(element, getattr(element, "_root", element)),
                }
                if split:
                    lead, rest = split
                    order = order_ref[0]
                    blocks.append(
                        Block(
                            id=f"html-{order}",
                            order=order,
                            text=lead,
                            normalized=normalize_for_compare(lead),
                            **base_kwargs,
                        )
                    )
                    order_ref[0] += 1
                    order = order_ref[0]
                    blocks.append(
                        Block(
                            id=f"html-{order}",
                            order=order,
                            text=rest,
                            normalized=normalize_for_compare(rest),
                            **{
                                **base_kwargs,
                                "heading": False,
                                "heading_level": None,
                                "bold": False,
                                "italic": False,
                                "underline": False,
                            },
                        )
                    )
                    order_ref[0] += 1
                else:
                    order = order_ref[0]
                    blocks.append(
                        Block(
                            id=f"html-{order}",
                            order=order,
                            text=text,
                            normalized=normalize_for_compare(text),
                            **base_kwargs,
                        )
                    )
                    order_ref[0] += 1
            return

    for child in list(element):
        walk_html_blocks(child, blocks, order_ref, hidden_here)


def extract_html_blocks(path: Path) -> list[Block]:
    root = ET.fromstring(path.read_text(encoding="utf-8"))
    annotate_html_parents(root)
    for node in root.iter():
        setattr(node, "_root", root)
    body = None
    if xml_local_name(root.tag).lower() == "html":
        for child in root:
            if isinstance(child.tag, str) and xml_local_name(child.tag).lower() == "body":
                body = child
                break
    if body is None:
        body = root
    blocks: list[Block] = []
    walk_html_blocks(body, blocks, [0])
    return blocks


def browser_render_and_extract(html_path: Path, output_pdf: Path) -> BrowserRenderResult:
    if sync_playwright is None:
        raise RuntimeError(
            "Playwright is not installed. Install it with `pip install playwright` and `playwright install chromium`."
        )

    js = """
() => {
  const BLOCK_TAGS = new Set(['P','LI','TD','TH','H1','H2','H3','H4','H5','H6','CAPTION','BLOCKQUOTE','PRE','FIGCAPTION']);
  const SKIP_TAGS = new Set(['SCRIPT','STYLE','NOSCRIPT','TEMPLATE','SVG','CANVAS','HEAD','META','LINK','IFRAME','OBJECT']);
  const boldTags = new Set(['B','STRONG']);
  const italicTags = new Set(['I','EM']);
  const underlineTags = new Set(['U']);

  const normalizeText = (text) => String(text || '')
    .replace(/\\u00A0/g, ' ')
    .replace(/\\bwww\\./gi, '')
    .replace(/[ \\t\\r\\f\\v]+/g, ' ')
    .trim();
  const DIFF_TOKEN_RE = /\\d[\\d,]*(?:\\.\\d+)?%?|\\.\\d+%?|[A-Za-z]+(?:[’'\\-][A-Za-z]+)*/g;

  const diffTokenKind = (token) => /^[+-]?(?:\\d[\\d,]*(?:\\.\\d+)?%?|\\.\\d+%?)$/.test(token) ? 'number' : 'word';
  const normalizeDiffToken = (token) => {
    const cleaned = normalizeText(token);
    return diffTokenKind(cleaned) === 'number'
      ? cleaned.replace(/,/g, '').toLowerCase()
      : cleaned.toLowerCase();
  };
  const diffTokensText = (text) => {
    const tokens = [];
    const source = normalizeText(text);
    DIFF_TOKEN_RE.lastIndex = 0;
    let match;
    while ((match = DIFF_TOKEN_RE.exec(source)) !== null) {
      tokens.push({
        text: match[0],
        normalized: normalizeDiffToken(match[0]),
        kind: diffTokenKind(match[0]),
      });
    }
    return tokens;
  };

  const splitLeadLabelText = (text, tableCell = false) => {
    if (tableCell) return null;
    const lines = String(text || '').split(/\\n+/).map(s => s.trim()).filter(Boolean);
    if (lines.length < 2) return null;
    const lead = lines[0];
    const rest = lines.slice(1).join(' ').trim();
    if (!rest) return null;
    const leadWords = lead.split(/\\s+/).filter(Boolean).length;
    const looksLikeLabel = lead.length <= 64 && leadWords <= 8 && !/[.!?:;]$/.test(lead);
    const bodySubstantial = rest.length >= 80;
    return looksLikeLabel && bodySubstantial ? { lead, rest } : null;
  };

  const visibleMeaningful = (text) => /[A-Za-z0-9]/.test(normalizeText(text));

  function isHidden(el) {
    const style = window.getComputedStyle(el);
    return style.display === 'none' || style.visibility === 'hidden';
  }

  function hasDescendantBlock(el) {
    for (const child of el.children) {
      if (BLOCK_TAGS.has(child.tagName)) return true;
      if (hasDescendantBlock(child)) return true;
    }
    return false;
  }

  function detectInlineFlags(el) {
    let bold = false;
    let italic = false;
    let underline = false;
    for (const node of el.querySelectorAll('*')) {
      const style = window.getComputedStyle(node);
      if (boldTags.has(node.tagName) || parseInt(style.fontWeight || '400', 10) >= 600) bold = true;
      if (italicTags.has(node.tagName) || style.fontStyle === 'italic') italic = true;
      if (underlineTags.has(node.tagName) || (style.textDecorationLine || '').includes('underline')) underline = true;
    }
    const selfStyle = window.getComputedStyle(el);
    if (boldTags.has(el.tagName) || parseInt(selfStyle.fontWeight || '400', 10) >= 600) bold = true;
    if (italicTags.has(el.tagName) || selfStyle.fontStyle === 'italic') italic = true;
    if (underlineTags.has(el.tagName) || (selfStyle.textDecorationLine || '').includes('underline')) underline = true;
    return { bold, italic, underline };
  }

  function collectTokenRects(el) {
    const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
    const tokens = [];
    let node;
    while ((node = walker.nextNode())) {
      const raw = node.nodeValue || '';
      if (!raw.trim()) continue;
      DIFF_TOKEN_RE.lastIndex = 0;
      let match;
      while ((match = DIFF_TOKEN_RE.exec(raw)) !== null) {
        const range = document.createRange();
        range.setStart(node, match.index);
        range.setEnd(node, match.index + match[0].length);
        const rects = Array.from(range.getClientRects());
        if (!rects.length) continue;
        const left = Math.min(...rects.map(rect => rect.left));
        const top = Math.min(...rects.map(rect => rect.top));
        const right = Math.max(...rects.map(rect => rect.right));
        const bottom = Math.max(...rects.map(rect => rect.bottom));
        tokens.push({
          text: match[0],
          normalized: normalizeDiffToken(match[0]),
          kind: diffTokenKind(match[0]),
          x: left + window.scrollX,
          y: top + window.scrollY,
          width: Math.max(1, right - left),
          height: Math.max(1, bottom - top),
        });
      }
    }
    return tokens;
  }

  const blocks = [];
  let order = 0;

  function walk(el, hiddenAncestor = false) {
    if (!(el instanceof HTMLElement)) return;
    if (SKIP_TAGS.has(el.tagName)) return;
    const hidden = hiddenAncestor || isHidden(el);
    if (hidden) return;

    if (BLOCK_TAGS.has(el.tagName)) {
      let captureSelf = ['TD','TH','LI','PRE','BLOCKQUOTE','CAPTION','FIGCAPTION'].includes(el.tagName);
      if (!captureSelf) captureSelf = !hasDescendantBlock(el);
      if (captureSelf) {
        const text = normalizeText(el.innerText || el.textContent || '');
        if (visibleMeaningful(text)) {
          const rect = el.getBoundingClientRect();
          const flags = detectInlineFlags(el);
          const tokenRects = collectTokenRects(el);
          const headingMatch = /^H([1-6])$/.exec(el.tagName);
          const tableEl = ['TD', 'TH'].includes(el.tagName) ? el.closest('table') : null;
          const trEl = ['TD', 'TH'].includes(el.tagName) ? el.closest('tr') : null;
          const tableRows = tableEl ? Array.from(tableEl.querySelectorAll('tr')) : [];
          const cellSiblings = trEl ? Array.from(trEl.children).filter(child => child.tagName === 'TD' || child.tagName === 'TH') : [];
          const meaningfulCellSiblings = cellSiblings.filter(cell => visibleMeaningful(normalizeText(cell.innerText || cell.textContent || '')));
          const numericMeaningfulCellSiblings = meaningfulCellSiblings.filter(cell => {
            const tokens = diffTokensText(cell.innerText || cell.textContent || '');
            return tokens.length === 1 && tokens[0].kind === 'number';
          });
          const rowKeyCell = cellSiblings.find(cell => visibleMeaningful(normalizeText(cell.innerText || cell.textContent || '')));
          const normalizeRowKey = (text) => normalizeText(text)
            .replace(/\\(\\s*\\d+\\s*\\)/g, '')
            .replace(/\\[\\s*\\d+\\s*\\]/g, '')
            .replace(/\\s+/g, ' ')
            .trim()
            .toLowerCase();
          const rowKey = rowKeyCell ? normalizeRowKey(rowKeyCell.innerText || rowKeyCell.textContent || '') : null;
          const rowSlot = meaningfulCellSiblings.length ? meaningfulCellSiblings.indexOf(el) : -1;
          const numericSlot = numericMeaningfulCellSiblings.length ? numericMeaningfulCellSiblings.indexOf(el) : -1;
          const block = {
            id: `html-${order}`,
            source: 'html',
            order,
            text,
            normalized: '',
            heading: Boolean(headingMatch),
            heading_level: headingMatch ? Number(headingMatch[1]) : null,
            bold: flags.bold,
            italic: flags.italic,
            underline: flags.underline,
            list_item: el.tagName === 'LI',
            table_cell: el.tagName === 'TD' || el.tagName === 'TH',
            kind: el.tagName.toLowerCase(),
            table_pos: tableEl && trEl ? {
              table: Math.max(0, Array.from(document.querySelectorAll('table')).indexOf(tableEl)),
              row: Math.max(0, tableRows.indexOf(trEl)),
              col: Math.max(0, cellSiblings.indexOf(el))
            } : null,
            row_key: rowKey,
            row_slot: rowSlot >= 0 ? rowSlot : null,
            numeric_slot: numericSlot >= 0 ? numericSlot : null,
            x: rect.left + window.scrollX,
            y: rect.top + window.scrollY,
            width: rect.width,
            height: rect.height,
            tokens: tokenRects,
          };
          const split = splitLeadLabelText(text, block.table_cell);
          if (split) {
            const leadTokens = diffTokensText(split.lead);
            const restTokens = diffTokensText(split.rest);
            const leadCount = leadTokens.length;
            const restCount = restTokens.length;
            el.setAttribute('data-docx-compare-order', String(order));
            blocks.push({
              ...block,
              id: `html-${order}`,
              order,
              text: split.lead,
              tokens: tokenRects.slice(0, leadCount),
            });
            order += 1;
            blocks.push({
              ...block,
              id: `html-${order}`,
              order,
              text: split.rest,
              heading: false,
              heading_level: null,
              bold: false,
              italic: false,
              underline: false,
              tokens: tokenRects.slice(leadCount, leadCount + restCount),
            });
            order += 1;
          } else {
            el.setAttribute('data-docx-compare-order', String(order));
            blocks.push(block);
            order += 1;
          }
          return;
        }
      }
    }
    for (const child of el.children) walk(child, hidden);
  }

  document.documentElement.style.setProperty('-webkit-print-color-adjust', 'exact');
  document.documentElement.style.setProperty('print-color-adjust', 'exact');
  walk(document.body || document.documentElement);
  return {
    blocks,
    width: Math.max(
      document.documentElement.scrollWidth,
      document.body ? document.body.scrollWidth : 0
    ),
    height: Math.max(
      document.documentElement.scrollHeight,
      document.body ? document.body.scrollHeight : 0
    ),
  };
}
"""

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch()
        page = browser.new_page(viewport={"width": 1280, "height": 900})
        page.goto(html_path.resolve().as_uri(), wait_until="load")
        page.wait_for_load_state("networkidle")
        page.emulate_media(media="screen")
        result = page.evaluate(js)
        width_px = max(float(result["width"]), 1.0)
        height_px = max(float(result["height"]), 1.0)
        page.pdf(
            path=str(output_pdf),
            width=f"{math.ceil(width_px)}px",
            height=f"{math.ceil(height_px)}px",
            print_background=True,
            margin={"top": "0px", "right": "0px", "bottom": "0px", "left": "0px"},
            prefer_css_page_size=False,
        )
        browser.close()

    blocks: list[Block] = []
    rects_by_order: dict[int, tuple[float, float, float, float]] = {}
    token_rects_by_order: dict[int, list[TokenRect]] = {}
    page_numbers_by_order: dict[int, int] = {}
    for item in result["blocks"]:
        order = int(item["order"])
        rects_by_order[order] = (
            float(item["x"]),
            float(item["y"]),
            float(item["width"]),
            float(item["height"]),
        )
        page_numbers_by_order[order] = 0
        token_rects_by_order[order] = [
            TokenRect(
                text=str(token["text"]),
                normalized=str(token["normalized"]),
                kind=str(token["kind"]),
                x=float(token["x"]),
                y=float(token["y"]),
                width=float(token["width"]),
                height=float(token["height"]),
            )
            for token in item.get("tokens", [])
        ]
        blocks.append(
            Block(
                id=item["id"],
                source="html",
                order=order,
                text=str(item["text"]),
                normalized=normalize_for_compare(str(item["text"])),
                heading=bool(item["heading"]),
                heading_level=int(item["heading_level"]) if item["heading_level"] is not None else None,
                bold=bool(item["bold"]),
                italic=bool(item["italic"]),
                underline=bool(item["underline"]),
                list_item=bool(item["list_item"]),
                table_cell=bool(item["table_cell"]),
                kind=str(item["kind"]),
                table_pos=(
                    int(item["table_pos"]["table"]),
                    int(item["table_pos"]["row"]),
                    int(item["table_pos"]["col"]),
                ) if item.get("table_pos") else None,
                row_key=normalize_row_key(str(item["row_key"])) if item.get("row_key") else None,
                row_slot=int(item["row_slot"]) if item.get("row_slot") is not None else None,
                numeric_slot=int(item["numeric_slot"]) if item.get("numeric_slot") is not None else None,
            )
        )
    return BrowserRenderResult(
        blocks=blocks,
        width_px=width_px,
        height_px=height_px,
        rects_by_order=rects_by_order,
        token_rects_by_order=token_rects_by_order,
        page_numbers_by_order=page_numbers_by_order,
        coordinate_space="browser_px",
    )


def rect_contains_point(rect: tuple[float, float, float, float], x: float, y: float) -> bool:
    x0, y0, x1, y1 = rect
    return x0 <= x <= x1 and y0 <= y <= y1


def overlap_area(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> float:
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    ix0 = max(ax0, bx0)
    iy0 = max(ay0, by0)
    ix1 = min(ax1, bx1)
    iy1 = min(ay1, by1)
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    return (ix1 - ix0) * (iy1 - iy0)


def is_pdf_chrome_text(text: str) -> bool:
    normalized = normalize_for_compare(text)
    compact = re.sub(r"\s+", " ", normalize_text(text)).strip()
    if not normalized:
        return True
    if "donnelley financial" in normalized:
        return True
    if "form 8-k" in normalized and "synopsys, inc." in normalized:
        return True
    chrome_patterns = [
        r"^\d{1,2}/\d{1,2}/\d{2,4},? \d{1,2}:\d{2} ?[ap]m$",
        r"^\d+/\d+$",
        r"^page \d+ of \d+$",
        r"^8-k$",
        r"^ex-?99\.1$",
        r"^ex-?99\.1 .*ex-?99\.1$",
        r"^exhibit 99\.1$",
        r"^[a-z]{2,}-w\d+-pf-\d+$",
        r"^\d{2}\.\d{2}\.\d{2}\.\d$",
        r"^[a-z0-9_-]{2,} \d{2}\.\d{2}\.\d{2}\.\d$",
        r"^[a-z0-9-]{4,} \d{2}\.\d{2}\.\d{2}\.\d$",
        r"^https?[:/].*sec\.gov/archives/edgar/.*$",
    ]
    return any(re.fullmatch(pattern, normalized) for pattern in chrome_patterns) or compact in {"1", "2", "3", "4", "5", "6", "7"}


def should_merge_ocr_paragraph(
    prev_rect: tuple[float, float, float, float],
    prev_text: str,
    curr_rect: tuple[float, float, float, float],
    curr_text: str,
    page_width: float,
) -> bool:
    if is_pdf_chrome_text(prev_text) or is_pdf_chrome_text(curr_text):
        return False
    prev_x0, prev_y0, prev_x1, prev_y1 = prev_rect
    curr_x0, curr_y0, curr_x1, curr_y1 = curr_rect
    prev_h = max(1.0, prev_y1 - prev_y0)
    curr_h = max(1.0, curr_y1 - curr_y0)
    vertical_gap = curr_y0 - prev_y1
    if vertical_gap > max(prev_h, curr_h) * 1.6 + 6:
        return False
    left_delta = abs(curr_x0 - prev_x0)
    if left_delta > max(26.0, prev_h * 2.2):
        return False
    prev_width = prev_x1 - prev_x0
    curr_width = curr_x1 - curr_x0
    prev_text_clean = normalize_text(prev_text).strip()
    curr_text_clean = normalize_text(curr_text).strip()
    if (
        re.search(r'[.!?]["”\']?$', prev_text_clean)
        and re.match(r'^[A-Z(“"]', curr_text_clean)
        and prev_width < page_width * 0.82
        and curr_width < page_width * 0.82
    ):
        return False
    return True


def cluster_words_into_blocks(
    words: list[dict[str, Any]],
    *,
    page_number: int,
    page_width: float,
    page_height: float,
    order_start: int,
    pixmap: Any | None = None,
) -> tuple[list[Block], dict[int, tuple[float, float, float, float]], dict[int, list[TokenRect]], dict[int, int], int]:
    blocks: list[Block] = []
    rects_by_order: dict[int, tuple[float, float, float, float]] = {}
    token_rects_by_order: dict[int, list[TokenRect]] = {}
    page_numbers_by_order: dict[int, int] = {}
    order = order_start

    words_sorted = sorted(words, key=lambda word: (round(word["rect"][1], 1), round(word["rect"][0], 1)))
    row_groups: list[list[dict[str, Any]]] = []
    for word in words_sorted:
        rect = word["rect"]
        center_y = (rect[1] + rect[3]) / 2
        if not row_groups:
            row_groups.append([word])
            continue
        last_row = row_groups[-1]
        last_centers = [((item["rect"][1] + item["rect"][3]) / 2) for item in last_row]
        last_heights = [(item["rect"][3] - item["rect"][1]) for item in last_row]
        threshold = max(4.0, (sum(last_heights) / len(last_heights)) * 0.6)
        if abs(center_y - (sum(last_centers) / len(last_centers))) <= threshold:
            last_row.append(word)
        else:
            row_groups.append([word])

    paragraph_words: list[dict[str, Any]] = []
    paragraph_rect: tuple[float, float, float, float] | None = None
    paragraph_text_parts: list[str] = []

    def flush_paragraph() -> None:
        nonlocal order, paragraph_words, paragraph_rect, paragraph_text_parts
        if not paragraph_words or paragraph_rect is None:
            paragraph_words = []
            paragraph_rect = None
            paragraph_text_parts = []
            return
        block_text = " ".join(part.strip() for part in paragraph_text_parts if part.strip()).strip()
        if not visible_meaningful(block_text):
            paragraph_words = []
            paragraph_rect = None
            paragraph_text_parts = []
            return
        block = Block(
            id=f"pdf-{order}",
            source="pdf",
            order=order,
            text=block_text,
            normalized=normalize_for_compare(block_text),
            table_cell=False,
            kind="pdf",
        )
        blocks.append(block)
        x0, y0, x1, y1 = paragraph_rect
        rects_by_order[order] = (x0, y0, x1 - x0, y1 - y0)
        page_numbers_by_order[order] = page_number
        token_rects: list[TokenRect] = []
        for word in paragraph_words:
            word_text = word["text"]
            word_rect = word["rect"]
            for match in DIFF_TOKEN_RE.finditer(word_text):
                token_text = match.group(0)
                token_rects.append(
                    TokenRect(
                        text=token_text,
                        normalized=normalize_diff_token(token_text),
                        kind=diff_token_kind(token_text),
                        x=word_rect[0],
                        y=word_rect[1],
                        width=max(1.0, word_rect[2] - word_rect[0]),
                        height=max(1.0, word_rect[3] - word_rect[1]),
                    )
                )
        token_rects_by_order[order] = token_rects
        order += 1
        paragraph_words = []
        paragraph_rect = None
        paragraph_text_parts = []

    for row_index, row_words in enumerate(row_groups):
        row_words.sort(key=lambda word: word["rect"][0])
        avg_height = sum(word["rect"][3] - word["rect"][1] for word in row_words) / max(len(row_words), 1)
        cluster_gap = max(22.0, avg_height * 2.6)
        clusters: list[list[dict[str, Any]]] = []
        for word in row_words:
            if not clusters:
                clusters.append([word])
                continue
            prev_word = clusters[-1][-1]
            gap = word["rect"][0] - prev_word["rect"][2]
            if gap > cluster_gap:
                clusters.append([word])
            else:
                clusters[-1].append(word)

        meaningful_clusters = [cluster for cluster in clusters if visible_meaningful(" ".join(item["text"] for item in cluster))]
        table_like = len(meaningful_clusters) > 1
        if not table_like and len(meaningful_clusters) == 1:
            cluster = meaningful_clusters[0]
            cluster_text = join_words_for_pdf_cluster(
                cluster,
                page_width=page_width,
                page_height=page_height,
                pixmap=pixmap,
            )
            x0 = min(item["rect"][0] for item in cluster)
            y0 = min(item["rect"][1] for item in cluster)
            x1 = max(item["rect"][2] for item in cluster)
            y1 = max(item["rect"][3] for item in cluster)
            cluster_rect = (x0, y0, x1, y1)
            if paragraph_rect is None:
                paragraph_words = list(cluster)
                paragraph_rect = cluster_rect
                paragraph_text_parts = [cluster_text]
            else:
                if should_merge_ocr_paragraph(paragraph_rect, " ".join(paragraph_text_parts), cluster_rect, cluster_text, page_width):
                    paragraph_words.extend(cluster)
                    px0, py0, px1, py1 = paragraph_rect
                    paragraph_rect = (min(px0, x0), min(py0, y0), max(px1, x1), max(py1, y1))
                    paragraph_text_parts.append(cluster_text)
                else:
                    flush_paragraph()
                    paragraph_words = list(cluster)
                    paragraph_rect = cluster_rect
                    paragraph_text_parts = [cluster_text]
            continue

        flush_paragraph()
        row_key = (
            normalize_row_key(
                join_words_for_pdf_cluster(
                    meaningful_clusters[0],
                    page_width=page_width,
                    page_height=page_height,
                    pixmap=pixmap,
                )
            )
            if table_like
            else None
        )
        row_slot = 0
        numeric_slot = 0

        for cluster in meaningful_clusters:
            cluster.sort(key=lambda word: word["rect"][0])
            block_text = join_words_for_pdf_cluster(
                cluster,
                page_width=page_width,
                page_height=page_height,
                pixmap=pixmap,
            )
            x0 = min(item["rect"][0] for item in cluster)
            y0 = min(item["rect"][1] for item in cluster)
            x1 = max(item["rect"][2] for item in cluster)
            y1 = max(item["rect"][3] for item in cluster)
            temp_block = Block(
                id="",
                source="pdf",
                order=0,
                text=block_text,
                normalized=normalize_for_compare(block_text),
            )
            token = single_value_token(temp_block)
            block_numeric_slot = None
            if table_like and token is not None and token.kind == "number":
                block_numeric_slot = numeric_slot
                numeric_slot += 1

            block = Block(
                id=f"pdf-{order}",
                source="pdf",
                order=order,
                text=block_text,
                normalized=normalize_for_compare(block_text),
                table_cell=table_like,
                kind="pdf",
                table_pos=(page_number, row_index, row_slot) if table_like else None,
                row_key=row_key,
                row_slot=row_slot if table_like else None,
                numeric_slot=block_numeric_slot,
            )
            blocks.append(block)
            rects_by_order[order] = (x0, y0, x1 - x0, y1 - y0)
            page_numbers_by_order[order] = page_number

            token_rects: list[TokenRect] = []
            for word in cluster:
                word_text = word["text"]
                word_rect = word["rect"]
                for match in DIFF_TOKEN_RE.finditer(word_text):
                    token_text = match.group(0)
                    token_rects.append(
                        TokenRect(
                            text=token_text,
                            normalized=normalize_diff_token(token_text),
                            kind=diff_token_kind(token_text),
                            x=word_rect[0],
                            y=word_rect[1],
                            width=max(1.0, word_rect[2] - word_rect[0]),
                            height=max(1.0, word_rect[3] - word_rect[1]),
                        )
                    )
            token_rects_by_order[order] = token_rects
            order += 1
            if table_like:
                row_slot += 1

    flush_paragraph()

    return blocks, rects_by_order, token_rects_by_order, page_numbers_by_order, order


def extract_pdf_blocks_via_ocr(path: Path) -> BrowserRenderResult:
    if fitz is None:
        raise RuntimeError("PyMuPDF is required for OCR PDF extraction.")
    tesseract_path = resolve_tesseract_executable()
    if tesseract_path is None:
        raise RuntimeError(
            "This PDF has no extractable text layer, and Tesseract OCR is not installed or not detectable. "
            "Install `tesseract` to compare DOCX against image/vector-outline PDFs. "
            "On Windows, run `winget install -e --id UB-Mannheim.TesseractOCR` and rerun setup."
        )
    prepare_tesseract_environment(tesseract_path)

    document = fitz.open(path)
    blocks: list[Block] = []
    rects_by_order: dict[int, tuple[float, float, float, float]] = {}
    token_rects_by_order: dict[int, list[TokenRect]] = {}
    page_numbers_by_order: dict[int, int] = {}
    max_width = 1.0
    max_height = 1.0
    order = 0

    for page_number, page in enumerate(document):
        page_rect = page.rect
        max_width = max(max_width, float(page_rect.width))
        max_height = max(max_height, float(page_rect.height))
        text_page = page.get_textpage_ocr(language="eng", dpi=300, full=True)
        dash_pixmap = page.get_pixmap(matrix=fitz.Matrix(4, 4), colorspace=fitz.csGRAY, alpha=False)
        words_raw = page.get_text("words", textpage=text_page)
        words: list[dict[str, Any]] = []
        for item in words_raw:
            x0, y0, x1, y1, text = float(item[0]), float(item[1]), float(item[2]), float(item[3]), str(item[4] or "")
            cleaned = normalize_text(text).strip()
            if not visible_meaningful(cleaned):
                continue
            words.append({"text": cleaned, "rect": (x0, y0, x1, y1)})

        (
            page_blocks,
            page_rects,
            page_token_rects,
            page_numbers,
            order,
        ) = cluster_words_into_blocks(
            words,
            page_number=page_number,
            page_width=float(page_rect.width),
            page_height=float(page_rect.height),
            order_start=order,
            pixmap=dash_pixmap,
        )
        blocks.extend(page_blocks)
        rects_by_order.update(page_rects)
        token_rects_by_order.update(page_token_rects)
        page_numbers_by_order.update(page_numbers)

    document.close()
    return BrowserRenderResult(
        blocks=blocks,
        width_px=max_width,
        height_px=max_height,
        rects_by_order=rects_by_order,
        token_rects_by_order=token_rects_by_order,
        page_numbers_by_order=page_numbers_by_order,
        coordinate_space="pdf_pt",
    )


def extract_pdf_blocks(path: Path) -> BrowserRenderResult:
    if fitz is None:
        raise RuntimeError(
            "PyMuPDF is not installed. Install it with `pip install PyMuPDF` to compare a DOCX against a PDF."
        )

    document = fitz.open(path)
    blocks: list[Block] = []
    rects_by_order: dict[int, tuple[float, float, float, float]] = {}
    token_rects_by_order: dict[int, list[TokenRect]] = {}
    page_numbers_by_order: dict[int, int] = {}
    max_width = 1.0
    max_height = 1.0
    order = 0
    total_text_blocks = 0

    for page_number, page in enumerate(document):
        page_rect = page.rect
        max_width = max(max_width, float(page_rect.width))
        max_height = max(max_height, float(page_rect.height))

        raw_blocks = [item for item in page.get_text("blocks") if len(item) >= 7 and int(item[6]) == 0]
        total_text_blocks += len(raw_blocks)
        text_blocks: list[dict[str, Any]] = []
        for item in raw_blocks:
            x0, y0, x1, y1, text = float(item[0]), float(item[1]), float(item[2]), float(item[3]), str(item[4] or "")
            cleaned = normalize_text(text).strip()
            if not visible_meaningful(cleaned):
                continue
            text_blocks.append(
                {
                    "rect": (x0, y0, x1, y1),
                    "text": cleaned,
                }
            )

        text_blocks.sort(key=lambda entry: (round(entry["rect"][1], 1), round(entry["rect"][0], 1)))
        row_groups: list[list[dict[str, Any]]] = []
        for entry in text_blocks:
            rect = entry["rect"]
            center_y = (rect[1] + rect[3]) / 2
            if not row_groups:
                row_groups.append([entry])
                continue
            last_group = row_groups[-1]
            last_rects = [item["rect"] for item in last_group]
            last_center = sum((item[1] + item[3]) / 2 for item in last_rects) / len(last_rects)
            last_height = sum(item[3] - item[1] for item in last_rects) / len(last_rects)
            threshold = max(6.0, last_height * 0.55)
            if abs(center_y - last_center) <= threshold:
                last_group.append(entry)
            else:
                row_groups.append([entry])

        words_raw = page.get_text("words")
        page_words: list[dict[str, Any]] = []
        for item in words_raw:
            x0, y0, x1, y1, text = float(item[0]), float(item[1]), float(item[2]), float(item[3]), str(item[4] or "")
            cleaned = normalize_text(text).strip()
            if not visible_meaningful(cleaned):
                continue
            page_words.append(
                {
                    "rect": (x0, y0, x1, y1),
                    "text": cleaned,
                }
            )

        for row_index, row in enumerate(row_groups):
            row.sort(key=lambda entry: entry["rect"][0])
            table_like = len(row) > 1
            row_key = normalize_row_key(row[0]["text"]) if table_like else None
            row_slot = 0
            numeric_slot = 0
            for entry in row:
                block_text = entry["text"]
                block_rect = entry["rect"]
                token = single_value_token(
                    Block(
                        id="",
                        source="pdf",
                        order=0,
                        text=block_text,
                        normalized=normalize_for_compare(block_text),
                    )
                )
                block_numeric_slot = None
                if table_like and token is not None and token.kind == "number":
                    block_numeric_slot = numeric_slot
                    numeric_slot += 1

                block = Block(
                    id=f"pdf-{order}",
                    source="pdf",
                    order=order,
                    text=block_text,
                    normalized=normalize_for_compare(block_text),
                    table_cell=table_like,
                    kind="pdf",
                    table_pos=(page_number, row_index, row_slot) if table_like else None,
                    row_key=row_key,
                    row_slot=row_slot if table_like else None,
                    numeric_slot=block_numeric_slot,
                )
                blocks.append(block)
                rects_by_order[order] = (
                    block_rect[0],
                    block_rect[1],
                    block_rect[2] - block_rect[0],
                    block_rect[3] - block_rect[1],
                )
                page_numbers_by_order[order] = page_number

                token_rects: list[TokenRect] = []
                matched_words = [
                    word for word in page_words
                    if overlap_area(block_rect, word["rect"]) > 0
                    or rect_contains_point(block_rect, (word["rect"][0] + word["rect"][2]) / 2, (word["rect"][1] + word["rect"][3]) / 2)
                ]
                matched_words.sort(key=lambda word: (word["rect"][1], word["rect"][0]))
                for word in matched_words:
                    word_text = word["text"]
                    word_rect = word["rect"]
                    for match in DIFF_TOKEN_RE.finditer(word_text):
                        token_text = match.group(0)
                        token_rects.append(
                            TokenRect(
                                text=token_text,
                                normalized=normalize_diff_token(token_text),
                                kind=diff_token_kind(token_text),
                                x=word_rect[0],
                                y=word_rect[1],
                                width=max(1.0, word_rect[2] - word_rect[0]),
                                height=max(1.0, word_rect[3] - word_rect[1]),
                            )
                        )
                token_rects_by_order[order] = token_rects
                order += 1
                if table_like:
                    row_slot += 1

    document.close()
    if total_text_blocks == 0 or not blocks:
        return extract_pdf_blocks_via_ocr(path)
    return BrowserRenderResult(
        blocks=blocks,
        width_px=max_width,
        height_px=max_height,
        rects_by_order=rects_by_order,
        token_rects_by_order=token_rects_by_order,
        page_numbers_by_order=page_numbers_by_order,
        coordinate_space="pdf_pt",
    )


def build_html_indices(
    html_blocks: list[Block],
) -> tuple[
    dict[str, list[int]],
    dict[str, list[int]],
    dict[int, list[int]],
    dict[tuple[int, int, int], list[int]],
    dict[int, list[int]],
    dict[tuple[int, str, int], list[int]],
    dict[tuple[str, int], list[int]],
    dict[tuple[str, int], list[int]],
]:
    exact_map: dict[str, list[int]] = {}
    token_index: dict[str, list[int]] = {}
    length_buckets: dict[int, list[int]] = {}
    table_pos_map: dict[tuple[int, int, int], list[int]] = {}
    table_index_map: dict[int, list[int]] = {}
    row_context_map: dict[tuple[int, str, int], list[int]] = {}
    global_row_context_map: dict[tuple[str, int], list[int]] = {}
    global_numeric_context_map: dict[tuple[str, int], list[int]] = {}
    for index, block in enumerate(html_blocks):
        exact_map.setdefault(block.normalized, []).append(index)
        pos_key = table_pos_key(block)
        if pos_key is not None:
            table_pos_map.setdefault(pos_key, []).append(index)
        table_idx = table_index_key(block)
        if table_idx is not None:
            table_index_map.setdefault(table_idx, []).append(index)
        row_key = row_context_key(block)
        if row_key is not None:
            row_context_map.setdefault(row_key, []).append(index)
        global_row_key = global_row_context_key(block)
        if global_row_key is not None:
            global_row_context_map.setdefault(global_row_key, []).append(index)
        global_numeric_key = global_numeric_context_key(block)
        if global_numeric_key is not None:
            global_numeric_context_map.setdefault(global_numeric_key, []).append(index)
        for token in dict.fromkeys(token for token in tokenize(block.normalized) if len(token) >= 3):
            token_index.setdefault(token, []).append(index)
        bucket = max(1, round(len(block.normalized) / 30))
        length_buckets.setdefault(bucket, []).append(index)
    return (
        exact_map,
        token_index,
        length_buckets,
        table_pos_map,
        table_index_map,
        row_context_map,
        global_row_context_map,
        global_numeric_context_map,
    )


def get_approx_candidates(
    doc_block: Block,
    html_blocks: list[Block],
    token_index: dict[str, list[int]],
    length_buckets: dict[int, list[int]],
    table_pos_map: dict[tuple[int, int, int], list[int]],
    table_index_map: dict[int, list[int]],
    used_html: set[int],
) -> list[int]:
    pos_key = table_pos_key(doc_block)
    if pos_key is not None and pos_key in table_pos_map:
        return [index for index in table_pos_map[pos_key] if index not in used_html]
    table_idx = table_index_key(doc_block)
    if table_idx is not None and table_idx in table_index_map:
        return [index for index in table_index_map[table_idx] if index not in used_html]
    candidates: set[int] = set()
    tokens = list(dict.fromkeys(token for token in tokenize(doc_block.normalized) if len(token) >= 3))
    for token in tokens[:3]:
        for index in token_index.get(token, []):
            if len(candidates) >= 120:
                break
            candidates.add(index)
    bucket = max(1, round(len(doc_block.normalized) / 30))
    for probe in range(bucket - 1, bucket + 2):
        for index in length_buckets.get(probe, []):
            if len(candidates) >= 200:
                break
            candidates.add(index)

    filtered: list[int] = []
    doc_len = max(len(doc_block.normalized), 1)
    for index in candidates:
        if index in used_html:
            continue
        html_len = max(len(html_blocks[index].normalized), 1)
        ratio = min(doc_len, html_len) / max(doc_len, html_len)
        if ratio >= 0.45:
            filtered.append(index)
    return filtered


def containment_cover_threshold(html_block: Block, doc_blocks: list[Block]) -> bool:
    if not doc_blocks:
        return False
    combined_length = sum(len(block.normalized) for block in doc_blocks)
    html_length = max(len(html_block.normalized), 1)
    if len(doc_blocks) >= 3:
        return True
    if combined_length / html_length >= 0.42:
        return True
    if any(len(block.normalized) / html_length >= 0.6 for block in doc_blocks):
        return True
    return False


def match_embedded_pdf_blocks(
    docx_blocks: list[Block],
    html_blocks: list[Block],
    unmatched_doc_indices: list[int],
    unmatched_html_indices: list[int],
) -> tuple[list[Match], set[int], set[int]]:
    if not unmatched_doc_indices or not unmatched_html_indices:
        return [], set(), set()

    candidate_matches_by_html: dict[int, list[tuple[int, float]]] = {}
    for html_index in unmatched_html_indices:
        html_block = html_blocks[html_index]
        if html_block.table_cell:
            continue
        if html_block.kind != "pdf" and len(html_block.normalized) < 40 and "\n" not in html_block.text:
            continue
        candidates: list[tuple[int, float]] = []
        for doc_index in unmatched_doc_indices:
            doc_block = docx_blocks[doc_index]
            score = containment_match_score(doc_block, html_block)
            if score >= 0.88:
                candidates.append((doc_index, score))
        if candidates:
            candidate_matches_by_html[html_index] = candidates

    consumed_doc_indices: set[int] = set()
    embedded_matches: list[Match] = []
    covered_html_indices: set[int] = set()

    for html_index, candidates in sorted(candidate_matches_by_html.items()):
        filtered_candidates = [(doc_index, score) for doc_index, score in candidates if doc_index not in consumed_doc_indices]
        if not filtered_candidates:
            continue
        filtered_candidates.sort(key=lambda item: docx_blocks[item[0]].order)
        covered_doc_blocks = [docx_blocks[doc_index] for doc_index, _score in filtered_candidates]
        if not containment_cover_threshold(html_blocks[html_index], covered_doc_blocks):
            continue
        covered_html_indices.add(html_index)
        for doc_index, score in filtered_candidates:
            consumed_doc_indices.add(doc_index)
            embedded_matches.append(
                Match(
                    docx_index=doc_index,
                    html_index=html_index,
                    match_type="contained",
                    score=score,
                    formatting_diffs=[],
                )
            )

    return embedded_matches, consumed_doc_indices, covered_html_indices


def compare_blocks(docx_blocks: list[Block], html_blocks: list[Block]) -> tuple[list[Match], list[Block], list[Block]]:
    (
        exact_map,
        token_index,
        length_buckets,
        table_pos_map,
        table_index_map,
        row_context_map,
        global_row_context_map,
        global_numeric_context_map,
    ) = build_html_indices(html_blocks)
    used_html: set[int] = set()
    matches: list[Match] = []
    unmatched_doc_indices: list[int] = []

    for doc_index, doc_block in enumerate(docx_blocks):
        pos_key = table_pos_key(doc_block)
        table_idx = table_index_key(doc_block)
        row_key = row_context_key(doc_block)
        global_row_key = global_row_context_key(doc_block)
        global_numeric_key = global_numeric_context_key(doc_block)
        pos_candidates: list[int] = []
        if pos_key is not None:
            pos_candidates = [index for index in table_pos_map.get(pos_key, []) if index not in used_html]
            exact_pos_index = next(
                (index for index in pos_candidates if html_blocks[index].normalized == doc_block.normalized),
                None,
            )
            if exact_pos_index is not None:
                used_html.add(exact_pos_index)
                matches.append(
                    Match(
                        docx_index=doc_index,
                        html_index=exact_pos_index,
                        match_type="exact",
                        score=1.0,
                        formatting_diffs=summarize_formatting_diff(doc_block, html_blocks[exact_pos_index]),
                    )
                )
                continue
        if global_numeric_key is not None:
            numeric_candidates = [index for index in global_numeric_context_map.get(global_numeric_key, []) if index not in used_html]
            exact_numeric_index = next(
                (index for index in numeric_candidates if html_blocks[index].normalized == doc_block.normalized),
                None,
            )
            if exact_numeric_index is not None:
                used_html.add(exact_numeric_index)
                matches.append(
                    Match(
                        docx_index=doc_index,
                        html_index=exact_numeric_index,
                        match_type="exact",
                        score=1.0,
                        formatting_diffs=summarize_formatting_diff(doc_block, html_blocks[exact_numeric_index]),
                    )
                )
                continue
            if numeric_candidates:
                best_numeric_index = max(
                    numeric_candidates,
                    key=lambda index: table_context_match_score(doc_block, html_blocks[index]),
                )
                best_numeric_score = table_context_match_score(doc_block, html_blocks[best_numeric_index])
                if best_numeric_score >= 0.8:
                    used_html.add(best_numeric_index)
                    matches.append(
                        Match(
                            docx_index=doc_index,
                            html_index=best_numeric_index,
                            match_type="approx",
                            score=best_numeric_score,
                            formatting_diffs=summarize_formatting_diff(doc_block, html_blocks[best_numeric_index]),
                        )
                    )
                    continue
        if row_key is not None:
            row_candidates = [index for index in row_context_map.get(row_key, []) if index not in used_html]
            exact_row_index = next(
                (index for index in row_candidates if html_blocks[index].normalized == doc_block.normalized),
                None,
            )
            if exact_row_index is not None:
                used_html.add(exact_row_index)
                matches.append(
                    Match(
                        docx_index=doc_index,
                        html_index=exact_row_index,
                        match_type="exact",
                        score=1.0,
                        formatting_diffs=summarize_formatting_diff(doc_block, html_blocks[exact_row_index]),
                    )
                )
                continue
            if row_candidates:
                best_row_index = max(
                    row_candidates,
                    key=lambda index: table_context_match_score(doc_block, html_blocks[index]),
                )
                best_row_score = table_context_match_score(doc_block, html_blocks[best_row_index])
                if best_row_score >= 0.8:
                    used_html.add(best_row_index)
                    matches.append(
                        Match(
                            docx_index=doc_index,
                            html_index=best_row_index,
                            match_type="approx",
                            score=best_row_score,
                            formatting_diffs=summarize_formatting_diff(doc_block, html_blocks[best_row_index]),
                        )
                    )
                    continue
        if global_row_key is not None:
            global_row_candidates = [index for index in global_row_context_map.get(global_row_key, []) if index not in used_html]
            exact_global_row_index = next(
                (index for index in global_row_candidates if html_blocks[index].normalized == doc_block.normalized),
                None,
            )
            if exact_global_row_index is not None:
                used_html.add(exact_global_row_index)
                matches.append(
                    Match(
                        docx_index=doc_index,
                        html_index=exact_global_row_index,
                        match_type="exact",
                        score=1.0,
                        formatting_diffs=summarize_formatting_diff(doc_block, html_blocks[exact_global_row_index]),
                    )
                )
                continue
            if global_row_candidates:
                best_global_row_index = max(
                    global_row_candidates,
                    key=lambda index: table_context_match_score(doc_block, html_blocks[index]),
                )
                best_global_row_score = table_context_match_score(doc_block, html_blocks[best_global_row_index])
                if best_global_row_score >= 0.8:
                    used_html.add(best_global_row_index)
                    matches.append(
                        Match(
                            docx_index=doc_index,
                            html_index=best_global_row_index,
                            match_type="approx",
                            score=best_global_row_score,
                            formatting_diffs=summarize_formatting_diff(doc_block, html_blocks[best_global_row_index]),
                        )
                    )
                    continue
        if doc_block.table_cell and doc_block.row_key:
            fuzzy_row_candidates = [
                index
                for index, candidate in enumerate(html_blocks)
                if index not in used_html
                and candidate.table_cell
                and candidate.row_key
                and similarity(doc_block.row_key, candidate.row_key) >= 0.94
                and (
                    (
                        doc_block.numeric_slot is not None
                        and candidate.numeric_slot == doc_block.numeric_slot
                    )
                    or (
                        doc_block.numeric_slot is None
                        and doc_block.row_slot is not None
                        and candidate.row_slot == doc_block.row_slot
                    )
                )
            ]
            if fuzzy_row_candidates:
                best_fuzzy_row_index = max(
                    fuzzy_row_candidates,
                    key=lambda index: table_context_match_score(doc_block, html_blocks[index]),
                )
                best_fuzzy_row_score = table_context_match_score(doc_block, html_blocks[best_fuzzy_row_index])
                if best_fuzzy_row_score >= 0.8:
                    used_html.add(best_fuzzy_row_index)
                    matches.append(
                        Match(
                            docx_index=doc_index,
                            html_index=best_fuzzy_row_index,
                            match_type="approx",
                            score=best_fuzzy_row_score,
                            formatting_diffs=summarize_formatting_diff(doc_block, html_blocks[best_fuzzy_row_index]),
                        )
                    )
                    continue
        if table_idx is not None and not pos_candidates:
            exact_table_candidates = table_index_map.get(table_idx, [])
            exact_table_index = next(
                (index for index in exact_table_candidates if index not in used_html and html_blocks[index].normalized == doc_block.normalized),
                None,
            )
            if exact_table_index is not None:
                used_html.add(exact_table_index)
                matches.append(
                    Match(
                        docx_index=doc_index,
                        html_index=exact_table_index,
                        match_type="exact",
                        score=1.0,
                        formatting_diffs=summarize_formatting_diff(doc_block, html_blocks[exact_table_index]),
                    )
                )
                continue
        allow_global_exact_for_table = (
            table_idx is None
            or (
                single_value_token(doc_block) is None
                and len(doc_block.normalized) >= 12
            )
        )
        exact_candidates = exact_map.get(doc_block.normalized, []) if allow_global_exact_for_table else []
        exact_index = next((index for index in exact_candidates if index not in used_html), None)
        if exact_index is not None:
            used_html.add(exact_index)
            matches.append(
                Match(
                    docx_index=doc_index,
                    html_index=exact_index,
                    match_type="exact",
                    score=1.0,
                    formatting_diffs=summarize_formatting_diff(doc_block, html_blocks[exact_index]),
                )
            )
            continue

        best_index = None
        best_score = 0.0
        for html_index in get_approx_candidates(doc_block, html_blocks, token_index, length_buckets, table_pos_map, table_index_map, used_html):
            score = similarity(doc_block.normalized, html_blocks[html_index].normalized)
            if score > best_score:
                best_score = score
                best_index = html_index
        if best_index is not None and best_score >= 0.73:
            used_html.add(best_index)
            matches.append(
                Match(
                    docx_index=doc_index,
                    html_index=best_index,
                    match_type="approx",
                    score=best_score,
                    formatting_diffs=summarize_formatting_diff(doc_block, html_blocks[best_index]),
                )
            )
        else:
            unmatched_doc_indices.append(doc_index)

    covered_html = set(used_html)
    embedded_matches, embedded_doc_indices, embedded_html_indices = match_embedded_pdf_blocks(
        docx_blocks,
        html_blocks,
        unmatched_doc_indices,
        [index for index in range(len(html_blocks)) if index not in covered_html],
    )
    matches.extend(embedded_matches)
    unmatched_doc_indices = [index for index in unmatched_doc_indices if index not in embedded_doc_indices]
    covered_html.update(embedded_html_indices)

    unmatched_docx = [docx_blocks[index] for index in unmatched_doc_indices]
    unmatched_html = [block for index, block in enumerate(html_blocks) if index not in covered_html]
    matches.sort(key=lambda match: match.html_index)
    return matches, unmatched_docx, unmatched_html


def shorten(text: str, limit: int = 220) -> str:
    text = re.sub(r"\s+", " ", normalize_text(text)).strip()
    if len(text) <= limit:
        return text
    return text[: limit - 1].rstrip() + "…"


def prnewswire_only_difference(doc_text: str, html_text: str) -> bool:
    if "/PRNewswire/" not in html_text or "/PRNewswire/" in doc_text:
        return False
    html_without_wire = normalize_for_compare(html_text.replace("/PRNewswire/", ""))
    doc_normalized = normalize_for_compare(doc_text)
    html_without_wire = html_without_wire.replace(", --", " -").replace("--", "-")
    doc_normalized = doc_normalized.replace(" - ", " -")
    return (
        "synopsys, inc." in doc_normalized
        and "synopsys, inc." in html_without_wire
        and similarity(doc_normalized, html_without_wire) >= 0.985
    )


def difference_subject(doc_kind: str, html_kind: str) -> str:
    if doc_kind == html_kind == "number":
        return "number"
    if doc_kind == html_kind == "word":
        return "word"
    return "token"


def date_phrase_for_token(text: str, token: DiffToken) -> str | None:
    source = normalize_text(text)
    for match in DATE_PHRASE_RE.finditer(source):
        if token.start >= match.start() and token.end <= match.end():
            return match.group(0)
    return None


def contextual_replacement_comment(
    doc_text: str,
    target_text: str,
    doc_token: DiffToken,
    target_token: DiffToken,
    *,
    target_name: str,
) -> str:
    doc_date = date_phrase_for_token(doc_text, doc_token)
    target_date = date_phrase_for_token(target_text, target_token)
    if doc_date and target_date and normalize_for_compare(doc_date) != normalize_for_compare(target_date):
        return f"The date is different, {target_date} in {target_name} while {doc_date} in word."
    return replacement_comment(doc_token, target_token, target_name=target_name)


def comment_token_text(token: DiffToken) -> str:
    if token.kind == "number":
        return token_with_prefix(token)
    return token.text


def replacement_comment(doc_token: DiffToken, target_token: DiffToken, *, target_name: str) -> str:
    subject = difference_subject(doc_token.kind, target_token.kind)
    return (
        f"The {subject} is different, {comment_token_text(target_token)} in {target_name} "
        f"while {comment_token_text(doc_token)} in word."
    )


def insertion_comment(target_token: DiffToken, *, target_name: str) -> str:
    subject = "number" if target_token.kind == "number" else "word"
    return f"The {subject} is extra in {target_name}, {comment_token_text(target_token)}. It is not present in word."


def deletion_comment(doc_token: DiffToken, *, target_name: str) -> str:
    subject = "number" if doc_token.kind == "number" else "word"
    return f"The {subject} is missing in {target_name}, {comment_token_text(doc_token)} in word."


def currency_symbol_comment(
    doc_block: Block,
    target_block: Block,
    *,
    target_name: str,
    docx_blocks: list[Block] | None = None,
    target_blocks: list[Block] | None = None,
) -> HtmlComment | None:
    doc_token = single_value_token(doc_block)
    target_token = single_value_token(target_block)
    if doc_token is None or target_token is None:
        return None
    if doc_token.kind != "number" or target_token.kind != "number":
        return None
    if doc_token.normalized != target_token.normalized:
        return None

    doc_currency = extract_currency_symbol(doc_block.text)
    target_currency = extract_currency_symbol(target_block.text)
    inferred_doc_currency = inferred_row_currency_symbol(doc_block, docx_blocks or [doc_block])
    inferred_target_currency = inferred_row_currency_symbol(target_block, target_blocks or [target_block])
    if target_name == "pdf":
        effective_target_currency = target_currency or (
            (inferred_target_currency or semantic_currency_hint(target_block)) if not doc_currency else None
        )
        if doc_currency and not target_currency:
            return None
        if not effective_target_currency:
            return None
        if doc_currency == effective_target_currency:
            return None
        contents = (
            f"The currency symbol is different, {target_name} has {effective_target_currency}{target_token.text} "
            f"while word has {doc_currency}{doc_token.text}."
            if doc_currency
            else f"The currency symbol is different, {target_name} has {effective_target_currency}{target_token.text} "
            f"while word has {doc_token.text}."
        )
        return HtmlComment(
            order=target_block.order,
            contents=contents,
            token_index=0,
        )
    if doc_currency == target_currency:
        if not (
            target_name == "pdf"
            and not doc_currency
            and inferred_doc_currency
        ):
            return None

    # OCR PDF table extraction often drops a leading currency symbol even when it
    # is visibly present in the rendered PDF. Do not treat that as a reliable
    # difference unless the PDF block explicitly contains a conflicting symbol.
    if (
        target_name == "pdf"
        and doc_block.table_cell
        and target_block.table_cell
        and doc_currency
        and not target_currency
    ):
        return None

    if target_currency and not doc_currency:
        contents = (
            f"The currency symbol is different, {target_name} has {target_currency}{target_token.text} "
            f"while word has {doc_token.text}."
        )
    elif (
        target_name == "pdf"
        and not doc_currency
        and not target_currency
        and inferred_doc_currency
    ):
        contents = (
            f"The currency symbol is different, {target_name} has {inferred_doc_currency}{target_token.text} "
            f"while word has {doc_token.text}."
        )
    elif doc_currency and not target_currency:
        contents = (
            f"The currency symbol is different, {target_name} has {target_token.text} "
            f"while word has {doc_currency}{doc_token.text}."
        )
    else:
        contents = (
            f"The currency symbol is different, {target_name} has {target_currency}{target_token.text} "
            f"while word has {doc_currency}{doc_token.text}."
        )

    return HtmlComment(
        order=target_block.order,
        contents=contents,
        token_index=0,
    )


def effective_currency_for_comment(
    block: Block,
    token: DiffToken,
    *,
    peer_blocks: list[Block] | None = None,
) -> str | None:
    peers = peer_blocks or [block]
    explicit = extract_currency_symbol(block.text)
    if explicit:
        return explicit
    if token.prefix_symbol and token.prefix_symbol in "$€£¥":
        return token.prefix_symbol
    inferred = inferred_row_currency_symbol(block, peers)
    if inferred:
        return inferred
    table_hint = table_context_currency_hint(block, peers)
    if table_hint:
        return table_hint
    return semantic_currency_hint(block)


def numeric_block_difference_comment(
    doc_block: Block,
    target_block: Block,
    *,
    target_name: str,
    docx_blocks: list[Block] | None = None,
    target_blocks: list[Block] | None = None,
) -> HtmlComment | None:
    doc_token = single_value_token(doc_block)
    target_token = single_value_token(target_block)
    if doc_token is None or target_token is None:
        return None
    if doc_token.kind != "number" or target_token.kind != "number":
        return None
    if doc_token.normalized == target_token.normalized:
        return None

    doc_currency = effective_currency_for_comment(doc_block, doc_token, peer_blocks=docx_blocks)
    target_currency = effective_currency_for_comment(target_block, target_token, peer_blocks=target_blocks)
    doc_display = token_with_prefix(doc_token)
    target_display = token_with_prefix(target_token)
    if doc_currency and not (doc_token.prefix_symbol and doc_token.prefix_symbol in "$€£¥"):
        doc_display = f"{doc_currency}{doc_token.text}"
    if target_currency and not (target_token.prefix_symbol and target_token.prefix_symbol in "$€£¥"):
        target_display = f"{target_currency}{target_token.text}"
    return HtmlComment(
        order=target_block.order,
        contents=f"The number is different, {target_display} in {target_name} while {doc_display} in word.",
        token_index=0,
    )


def append_word_level_comments(
    comments: list[HtmlComment],
    *,
    order: int,
    doc_text: str,
    target_text: str,
    target_tokens: list[DiffToken],
    doc_slice: list[DiffToken],
    target_slice: list[DiffToken],
    target_start: int,
    target_name: str,
) -> None:
    pair_count = min(len(doc_slice), len(target_slice))
    for offset in range(pair_count):
        doc_token = doc_slice[offset]
        target_token = target_slice[offset]
        if doc_token.normalized == target_token.normalized:
            continue
        comments.append(
            HtmlComment(
                order=order,
                contents=contextual_replacement_comment(
                    doc_text,
                    target_text,
                    doc_token,
                    target_token,
                    target_name=target_name,
                ),
                token_index=target_start + offset,
            )
        )

    for offset, target_token in enumerate(target_slice[pair_count:], start=pair_count):
        comments.append(
            HtmlComment(
                order=order,
                contents=insertion_comment(target_token, target_name=target_name),
                token_index=target_start + offset,
            )
        )

    for doc_token in doc_slice[pair_count:]:
        anchor_index = None
        if target_tokens:
            anchor_index = min(max(target_start + pair_count, 0), len(target_tokens) - 1)
        comments.append(
            HtmlComment(
                order=order,
                contents=deletion_comment(doc_token, target_name=target_name),
                token_index=anchor_index,
            )
        )


def trim_contained_token_alignment(
    doc_tokens: list[DiffToken],
    target_tokens: list[DiffToken],
) -> tuple[list[DiffToken], list[DiffToken], int]:
    if not doc_tokens or not target_tokens:
        return doc_tokens, target_tokens, 0
    matcher = difflib.SequenceMatcher(
        a=[token.normalized for token in doc_tokens],
        b=[token.normalized for token in target_tokens],
        autojunk=False,
    )
    matching_blocks = [block for block in matcher.get_matching_blocks() if block.size > 0]
    if not matching_blocks:
        return doc_tokens, target_tokens, 0
    anchor = max(matching_blocks, key=lambda block: block.size)
    expected_start = max(0, anchor.b - anchor.a)
    min_start = max(0, expected_start - 6)
    max_start = min(max(0, len(target_tokens) - 1), expected_start + 6)
    min_window = max(1, len(doc_tokens) - 2)
    max_window = min(len(target_tokens), len(doc_tokens) + 8)
    best_start = expected_start
    best_end = min(len(target_tokens), expected_start + len(doc_tokens))
    best_score = -1.0
    doc_text = " ".join(token.normalized for token in doc_tokens)
    for start in range(min_start, max_start + 1):
        for window_len in range(min_window, max_window + 1):
            end = start + window_len
            if end > len(target_tokens):
                break
            target_text = " ".join(token.normalized for token in target_tokens[start:end])
            score = difflib.SequenceMatcher(None, doc_text, target_text).ratio() - (0.015 * abs(window_len - len(doc_tokens)))
            if score > best_score:
                best_score = score
                best_start = start
                best_end = end
    return doc_tokens, target_tokens[best_start:best_end], best_start


def token_with_prefix(token: DiffToken) -> str:
    prefix = token.prefix_symbol or ""
    spaces = " " * token.spaces_before
    return f"{prefix}{spaces}{token.text}" if prefix or spaces else token.text


def inter_token_separator(text: str, left: DiffToken, right: DiffToken) -> str:
    source = normalize_text(text)
    return source[left.end:right.start]


def normalized_separator_symbol(separator: str) -> str:
    compact = re.sub(r"\s+", "", normalize_text(separator))
    return compact


def contextual_equal_token_comments(
    *,
    order: int,
    target_name: str,
    doc_text: str,
    target_text: str,
    doc_tokens: list[DiffToken],
    target_tokens: list[DiffToken],
    target_offset: int = 0,
) -> list[HtmlComment]:
    comments: list[HtmlComment] = []
    matcher = difflib.SequenceMatcher(
        a=[token.normalized for token in doc_tokens],
        b=[token.normalized for token in target_tokens],
        autojunk=False,
    )
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag != "equal":
            continue
        for offset in range(min(i2 - i1, j2 - j1)):
            doc_token = doc_tokens[i1 + offset]
            target_token = target_tokens[j1 + offset]
            token_index = target_offset + j1 + offset
            doc_currency = bool(doc_token.prefix_symbol) and doc_token.prefix_symbol in "$€£¥"
            target_currency = bool(target_token.prefix_symbol) and target_token.prefix_symbol in "$€£¥"
            if (
                doc_token.kind == "number"
                and doc_token.prefix_symbol != target_token.prefix_symbol
                and (doc_currency or target_currency)
            ):
                if target_name == "pdf" and not target_currency:
                    continue
                comments.append(
                    HtmlComment(
                        order=order,
                        contents=(
                            f"The currency symbol is different, {target_name} has {token_with_prefix(target_token)} "
                            f"while word has {token_with_prefix(doc_token)}."
                        ),
                        token_index=token_index,
                    )
                )
                continue
            if (
                doc_token.kind == "number"
                and doc_token.prefix_symbol == target_token.prefix_symbol
                and doc_token.spaces_before != target_token.spaces_before
                and max(doc_token.spaces_before, target_token.spaces_before) >= 2
            ):
                comments.append(
                    HtmlComment(
                        order=order,
                        contents=(
                            f"The spacing is different, {target_name} has {token_with_prefix(target_token)} "
                            f"while word has {token_with_prefix(doc_token)}."
                        ),
                        token_index=token_index,
                    )
                )
        for offset in range(min(i2 - i1, j2 - j1) - 1):
            doc_left = doc_tokens[i1 + offset]
            doc_right = doc_tokens[i1 + offset + 1]
            target_left = target_tokens[j1 + offset]
            target_right = target_tokens[j1 + offset + 1]
            doc_sep = normalized_separator_symbol(inter_token_separator(doc_text, doc_left, doc_right))
            target_sep = normalized_separator_symbol(inter_token_separator(target_text, target_left, target_right))
            if doc_sep == target_sep:
                continue
            if not doc_sep and not target_sep:
                continue
            if doc_left.kind == "number" or doc_right.kind == "number":
                continue
            doc_desc = doc_sep or "no symbol"
            target_desc = target_sep or "no symbol"
            if target_name == "pdf" and target_desc == "no symbol":
                continue
            comments.append(
                HtmlComment(
                    order=order,
                    contents=(
                        f"The symbol is different, {target_name} has {target_desc} "
                        f"between {target_left.text} and {target_right.text} while word has {doc_desc}."
                    ),
                    token_index=target_offset + j1 + offset + 1,
                )
            )
    return comments


def dedupe_html_comments(comments: list[HtmlComment]) -> list[HtmlComment]:
    seen: set[tuple[int, int | None, str]] = set()
    deduped: list[HtmlComment] = []
    for comment in comments:
        key = (comment.order, comment.token_index, comment.contents)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(comment)
    return deduped


def single_token_target_comments(
    doc_block: Block,
    target_block: Block,
    *,
    target_name: str,
) -> list[HtmlComment] | None:
    doc_tokens = diff_tokens(doc_block.text)
    if len(doc_tokens) != 1:
        return None
    target_tokens = diff_tokens(target_block.text)
    if len(target_tokens) <= 1:
        return None

    doc_token = doc_tokens[0]
    best_index: int | None = None
    best_score = 0.0
    for index, target_token in enumerate(target_tokens):
        score = 0.0
        if doc_token.kind == target_token.kind:
            if doc_token.normalized == target_token.normalized:
                score = 1.0
            elif doc_token.kind == "number":
                score = numeric_value_compatibility(doc_token, target_token)
            else:
                score = difflib.SequenceMatcher(None, doc_token.normalized, target_token.normalized).ratio()
        if score > best_score:
            best_score = score
            best_index = index

    if best_index is None or best_score < 0.72:
        return None

    target_token = target_tokens[best_index]
    if doc_token.normalized == target_token.normalized:
        doc_currency = bool(doc_token.prefix_symbol) and doc_token.prefix_symbol in "$€£¥"
        target_currency = bool(target_token.prefix_symbol) and target_token.prefix_symbol in "$€£¥"
        if doc_token.prefix_symbol != target_token.prefix_symbol and (doc_currency or target_currency):
            if target_name == "pdf" and not target_currency:
                return []
            return [
                HtmlComment(
                    order=target_block.order,
                    contents=(
                        f"The currency symbol is different, {target_name} has {token_with_prefix(target_token)} "
                        f"while word has {token_with_prefix(doc_token)}."
                    ),
                    token_index=best_index,
                )
            ]
        if (
            doc_token.kind == "number"
            and doc_token.prefix_symbol == target_token.prefix_symbol
            and doc_token.spaces_before != target_token.spaces_before
            and max(doc_token.spaces_before, target_token.spaces_before) >= 2
        ):
            return [
                HtmlComment(
                    order=target_block.order,
                    contents=(
                        f"The spacing is different, {target_name} has {token_with_prefix(target_token)} "
                        f"while word has {token_with_prefix(doc_token)}."
                    ),
                    token_index=best_index,
                )
            ]
        return []

    return [
        HtmlComment(
            order=target_block.order,
            contents=contextual_replacement_comment(
                doc_block.text,
                target_block.text,
                doc_token,
                target_token,
                target_name=target_name,
            ),
            token_index=best_index,
        )
    ]


def text_difference_comments(
    doc_block: Block,
    target_block: Block,
    score: float,
    *,
    target_name: str,
    docx_blocks: list[Block] | None = None,
    target_blocks: list[Block] | None = None,
    match_type: str = "approx",
) -> list[HtmlComment]:
    contained_like = match_type == "contained" or (
        target_name == "pdf"
        and not doc_block.table_cell
        and not target_block.table_cell
        and len(target_block.normalized) > max(len(doc_block.normalized) * 2, 160)
    )
    currency_comment = currency_symbol_comment(
        doc_block,
        target_block,
        target_name=target_name,
        docx_blocks=docx_blocks,
        target_blocks=target_blocks,
    )
    if currency_comment is not None:
        return [currency_comment]
    numeric_comment = numeric_block_difference_comment(
        doc_block,
        target_block,
        target_name=target_name,
        docx_blocks=docx_blocks,
        target_blocks=target_blocks,
    )
    if numeric_comment is not None:
        return [numeric_comment]
    single_token_comments = single_token_target_comments(
        doc_block,
        target_block,
        target_name=target_name,
    )
    if single_token_comments is not None:
        return single_token_comments
    if target_name == "pdf" and normalize_without_footnote_refs(doc_block.text) == normalize_without_footnote_refs(target_block.text):
        return []
    if (
        target_name == "pdf"
        and not doc_block.table_cell
        and not target_block.table_cell
        and normalize_pdf_paragraph_artifacts(doc_block.text) == normalize_pdf_paragraph_artifacts(target_block.text)
    ):
        return []
    if normalize_for_compare(strip_leading_markers(doc_block.text)) == normalize_for_compare(strip_leading_markers(target_block.text)):
        return []
    if target_name == "pdf" and not contained_like and not doc_block.table_cell and not target_block.table_cell and score < 0.9:
        return []

    doc_tokens = diff_tokens(doc_block.text)
    target_tokens = diff_tokens(target_block.text)
    if not doc_tokens and not target_tokens:
        return []
    target_offset = 0
    if contained_like:
        doc_tokens, target_tokens, target_offset = trim_contained_token_alignment(doc_tokens, target_tokens)
        if not doc_tokens and not target_tokens:
            return []

    contextual_comments = contextual_equal_token_comments(
        order=target_block.order,
        target_name=target_name,
        doc_text=doc_block.text,
        target_text=target_block.text,
        doc_tokens=doc_tokens,
        target_tokens=target_tokens,
        target_offset=target_offset,
    )
    if doc_block.normalized == target_block.normalized or score >= 0.999:
        return contextual_comments

    comments: list[HtmlComment] = []
    matcher = difflib.SequenceMatcher(
        a=[token.normalized for token in doc_tokens],
        b=[token.normalized for token in target_tokens],
        autojunk=False,
    )
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        if tag == "replace":
            append_word_level_comments(
                comments,
                order=target_block.order,
                doc_text=doc_block.text,
                target_text=target_block.text,
                target_tokens=target_tokens,
                doc_slice=doc_tokens[i1:i2],
                target_slice=target_tokens[j1:j2],
                target_start=target_offset + j1,
                target_name=target_name,
            )
            continue
        if tag == "insert":
            append_word_level_comments(
                comments,
                order=target_block.order,
                doc_text=doc_block.text,
                target_text=target_block.text,
                target_tokens=target_tokens,
                doc_slice=[],
                target_slice=target_tokens[j1:j2],
                target_start=target_offset + j1,
                target_name=target_name,
            )
            continue
        if tag == "delete":
            append_word_level_comments(
                comments,
                order=target_block.order,
                doc_text=doc_block.text,
                target_text=target_block.text,
                target_tokens=target_tokens,
                doc_slice=doc_tokens[i1:i2],
                target_slice=[],
                target_start=target_offset + j1,
                target_name=target_name,
            )

    comments.extend(contextual_comments)
    comments = dedupe_html_comments(comments)

    if comments:
        if target_name == "pdf" and contained_like and not doc_block.table_cell and not target_block.table_cell:
            focused_comments = [
                comment
                for comment in comments
                if (
                    "spacing is different" in comment.contents
                    or "currency symbol is different" in comment.contents
                    or comment.contents.startswith("The date ")
                    or comment.contents.startswith("The number ")
                )
            ]
            if focused_comments:
                return focused_comments
            if len(doc_tokens) <= 6:
                return comments
            return []
        if (
            target_name == "pdf"
            and comments
            and all(comment.contents == "The number is missing in pdf, 1 in word." for comment in comments)
            and re.search(r"\(\s*1\s*\)", doc_block.text)
        ):
            return []
        if (
            target_name == "pdf"
            and not doc_block.table_cell
            and not target_block.table_cell
            and len(comments) <= 2
            and len(doc_block.text) <= 80
            and len(target_block.text) <= 80
            and all(comment.contents.startswith("The word is extra in pdf,") for comment in comments)
            and token_subsequence_ratio(doc_block.normalized, target_block.normalized) >= 0.99
        ):
            return []
        if target_name == "pdf" and not doc_block.table_cell and not target_block.table_cell and len(comments) > 4:
            return [
                HtmlComment(
                    order=target_block.order,
                    contents=(
                        "The paragraph text differs between the PDF and Word. "
                        f"PDF: {shorten(target_block.text, 140)} "
                        f"Word: {shorten(doc_block.text, 140)}"
                    ),
                    token_index=0 if target_tokens else None,
                )
            ]
        return comments

    if target_name == "html" and prnewswire_only_difference(doc_block.text, target_block.text):
        pr_tokens = diff_tokens(target_block.text)
        pr_index = next((index for index, token in enumerate(pr_tokens) if token.normalized == "prnewswire"), None)
        return [
            HtmlComment(
                order=target_block.order,
                contents="The word is extra in html, PRNewswire. It is not present in word.",
                token_index=pr_index,
            )
        ]

    if target_name == "pdf" and normalize_without_punctuation(doc_block.text) == normalize_without_punctuation(target_block.text):
        return []

    if target_name == "pdf" and contained_like:
        return []

    return [
        HtmlComment(
            order=target_block.order,
            contents=(
                "The paragraph text is different. "
                f"{target_name.upper()}: {shorten(target_block.text, 140)} "
                f"Word: {shorten(doc_block.text, 140)}"
            ),
            token_index=0 if target_tokens else None,
        )
    ]


def appendix_summary_blocks(docx_blocks: list[Block], unmatched_docx: list[Block]) -> list[Block]:
    summary_blocks: list[Block] = []
    seen_ids: set[str] = set()

    def add(block: Block) -> None:
        if block.id in seen_ids:
            return
        seen_ids.add(block.id)
        summary_blocks.append(block)

    for block in unmatched_docx:
        if normalize_for_compare(block.text) == "synopsys, inc.":
            continue
        add(block)

    if len(summary_blocks) >= 28:
        return summary_blocks

    label_texts = {
        "three months ended",
        "january 31,",
        "adjustments:",
        "amortization of acquired intangible assets",
        "stock-based compensation",
    }
    for block in docx_blocks:
        normalized = block.normalized
        if not block.table_cell or block.table_pos is None:
            continue
        table_idx, row_idx, _col_idx = block.table_pos
        if "gaap to non-gaap reconciliation" in normalized:
            add(block)
        if table_idx != 1 or row_idx < 19:
            continue
        numeric_summary_value = False
        if re.fullmatch(r"\$?\d\.\d{2}", block.text.strip()):
            if row_idx in {22, 25}:
                numeric_summary_value = True
            elif block.table_pos[2] == 1:
                numeric_summary_value = True
        if (
            "per diluted share attributed to synopsys" in normalized
            or normalized in label_texts
            or numeric_summary_value
        ):
            add(block)
        if len(summary_blocks) >= 28:
            break

    return summary_blocks


def build_comments(
    docx_blocks: list[Block],
    html_blocks: list[Block],
    matches: list[Match],
    unmatched_docx: list[Block],
    unmatched_html: list[Block],
    *,
    target_label: str = "HTML",
) -> tuple[list[HtmlComment], list[tuple[Block, str]]]:
    html_comments: list[HtmlComment] = []
    target_name = target_label.lower()
    for match in matches:
        doc_block = docx_blocks[match.docx_index]
        html_block = html_blocks[match.html_index]
        html_comments.extend(
            text_difference_comments(
                doc_block,
                html_block,
                match.score,
                target_name=target_name,
                docx_blocks=docx_blocks,
                target_blocks=html_blocks,
                match_type=match.match_type,
            )
        )

    def unmatched_html_fallback_comments(block: Block) -> list[HtmlComment]:
        best_doc_block: Block | None = None
        best_score = 0.0
        candidates = unmatched_docx if unmatched_docx else docx_blocks
        for doc_block in candidates:
            if doc_block.table_cell != block.table_cell:
                continue
            score = similarity(doc_block.normalized, block.normalized)
            if block.table_cell and doc_block.row_key and block.row_key:
                score = max(score, table_context_match_score(doc_block, block))
            if score > best_score:
                best_doc_block = doc_block
                best_score = score

        if best_doc_block is None:
            return []

        overlap = token_overlap_ratio(best_doc_block.normalized, block.normalized)
        strong_near_match = (
            best_score >= 0.78
            or (
                overlap >= 0.5
                and best_score >= 0.58
                and len(diff_tokens(best_doc_block.text)) <= 6
                and len(diff_tokens(block.text)) <= 6
            )
        )
        if not strong_near_match:
            return []
        return text_difference_comments(
            best_doc_block,
            block,
            max(best_score, similarity(best_doc_block.normalized, block.normalized)),
            target_name=target_name,
            docx_blocks=docx_blocks,
            target_blocks=html_blocks,
            match_type="approx",
        )

    def unmatched_docx_fallback_comments(block: Block) -> list[HtmlComment]:
        best_target_block: Block | None = None
        best_score = 0.0
        target_candidates = unmatched_html if unmatched_html else html_blocks
        doc_token = single_value_token(block)
        doc_token_norm = doc_token.normalized if doc_token is not None else None
        for target_block in target_candidates:
            if target_block.table_cell != block.table_cell:
                continue
            if (
                block.table_cell
                and block.row_key
                and target_block.row_key
                and similarity(block.row_key, target_block.row_key) < 0.82
            ):
                continue
            score = similarity(block.normalized, target_block.normalized)
            if block.table_cell and block.row_key and target_block.row_key:
                score = max(score, table_context_match_score(block, target_block))
            if doc_token_norm is not None:
                target_token_norms = [token.normalized for token in diff_tokens(target_block.text)]
                if doc_token_norm in target_token_norms:
                    score = max(score, 0.72 if not block.table_cell else 0.82)
            if score > best_score:
                best_target_block = target_block
                best_score = score

        if best_target_block is None:
            return []

        overlap = token_overlap_ratio(block.normalized, best_target_block.normalized)
        strong_near_match = (
            best_score >= 0.8
            or (
                block.table_cell
                and best_score >= 0.58
                and len(diff_tokens(block.text)) <= 6
            )
            or (
                overlap >= 0.5
                and best_score >= 0.6
                and len(diff_tokens(block.text)) <= 8
            )
        )
        if not strong_near_match:
            return []

        match_type = "contained" if len(best_target_block.normalized) > max(len(block.normalized) * 1.25, 24) else "approx"
        return text_difference_comments(
            block,
            best_target_block,
            max(best_score, similarity(block.normalized, best_target_block.normalized)),
            target_name=target_name,
            docx_blocks=docx_blocks,
            target_blocks=html_blocks,
            match_type=match_type,
        )

    for block in unmatched_html:
        if target_label == "PDF":
            continue
        fallback_comments = unmatched_html_fallback_comments(block)
        if fallback_comments:
            html_comments.extend(fallback_comments)
            continue
        html_comments.append(
            HtmlComment(
                order=block.order,
                contents=f"This {target_label} block has no corresponding content in the DOCX.",
                token_index=0 if diff_tokens(block.text) else None,
            )
        )

    appendix_comments = [
        (block, f"This DOCX content was not found in the {target_label}.")
        for block in appendix_summary_blocks(docx_blocks, unmatched_docx)
    ]
    if target_label == "PDF":
        for block in unmatched_docx:
            html_comments.extend(unmatched_docx_fallback_comments(block))
        appendix_comments = []
    return html_comments, appendix_comments


def group_html_comments(html_comments: list[HtmlComment]) -> dict[int, list[str]]:
    grouped: dict[int, list[str]] = {}
    for comment in html_comments:
        grouped.setdefault(comment.order, []).append(comment.contents)
    return grouped


def pdf_page_summary_comments(
    *,
    docx_blocks: list[Block],
    pdf_blocks: list[Block],
    unmatched_pdf: list[Block],
    matches: list[Match],
    render_result: BrowserRenderResult,
) -> list[HtmlComment]:
    matched_pages: Counter[int] = Counter()
    unmatched_by_page: dict[int, list[Block]] = {}
    order_by_page: dict[int, int] = {}

    for match in matches:
        block = pdf_blocks[match.html_index]
        page_number = render_result.page_numbers_by_order.get(block.order)
        if page_number is not None:
            matched_pages[page_number] += 1

    for block in unmatched_pdf:
        page_number = render_result.page_numbers_by_order.get(block.order)
        if page_number is None:
            continue
        unmatched_by_page.setdefault(page_number, []).append(block)
        order_by_page.setdefault(page_number, block.order)

    comments: list[HtmlComment] = []
    for page_number, blocks in sorted(unmatched_by_page.items()):
        meaningful_blocks = [block for block in blocks if visible_meaningful(block.text) and not is_pdf_chrome_text(block.text)]
        if len(meaningful_blocks) < 3:
            continue
        top_blocks = meaningful_blocks[:3]
        if any(pdf_block_has_docx_anchor(block, docx_blocks) for block in top_blocks):
            continue
        matched_count = matched_pages.get(page_number, 0)
        if matched_count > 0 and not (matched_count <= 1 and len(meaningful_blocks) >= 6):
            continue
        anchor_block = meaningful_blocks[0]
        comments.append(
            HtmlComment(
                order=anchor_block.order,
                contents="This PDF page contains content with no corresponding content in the Word document.",
                token_index=0 if diff_tokens(anchor_block.text) else None,
            )
        )
    return comments


def pdf_safe_text(text: str) -> str:
    text = normalize_text(text)
    text = text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    encoded = text.encode("latin-1", errors="replace")
    return encoded.decode("latin-1")


class PdfBuilder:
    def __init__(self, page_width: int = 612, page_height: int = 792) -> None:
        self.page_width = page_width
        self.page_height = page_height
        self.margin_left = 54
        self.margin_right = 54
        self.margin_top = 54
        self.margin_bottom = 54
        self.font_size = 11
        self.leading = 14
        self.title_size = 16
        self.subtitle_size = 10
        self.pages: list[dict[str, object]] = []
        self.current_page: dict[str, object] | None = None
        self.current_y = 0

    def new_page(self) -> None:
        page: dict[str, object] = {"ops": [], "annots": []}
        self.pages.append(page)
        self.current_page = page
        self.current_y = self.page_height - self.margin_top

    def ensure_page(self) -> None:
        if self.current_page is None:
            self.new_page()

    def available_height(self) -> int:
        return int(self.current_y - self.margin_bottom)

    def wrap_text(self, text: str, font_size: int | None = None, indent: int = 0) -> list[str]:
        size = font_size or self.font_size
        usable_width = self.page_width - self.margin_left - self.margin_right - indent
        chars_per_line = max(30, int(usable_width / (size * 0.54)))
        text = re.sub(r"[ \t]+", " ", normalize_text(text)).strip()
        if not text:
            return [""]
        wrapped: list[str] = []
        for paragraph in text.splitlines() or [""]:
            if not paragraph.strip():
                wrapped.append("")
                continue
            wrapped.extend(textwrap.wrap(paragraph, width=chars_per_line, break_long_words=False, break_on_hyphens=False))
        return wrapped or [""]

    def _append_text_op(self, x: int, y: int, text: str, font_size: int) -> None:
        assert self.current_page is not None
        op = f"BT /F1 {font_size} Tf 1 0 0 1 {x} {y} Tm ({pdf_safe_text(text)}) Tj ET"
        self.current_page["ops"].append(op)

    def add_wrapped_text(
        self,
        text: str,
        *,
        font_size: int | None = None,
        indent: int = 0,
        gap_after: int = 0,
    ) -> tuple[int, int]:
        self.ensure_page()
        size = font_size or self.font_size
        lines = self.wrap_text(text, font_size=size, indent=indent)
        required = len(lines) * self.leading + gap_after
        if self.available_height() < required:
            self.new_page()
        assert self.current_page is not None
        first_y = self.current_y
        x = self.margin_left + indent
        for line in lines:
            self._append_text_op(x, self.current_y, line, size)
            self.current_y -= self.leading
        self.current_y -= gap_after
        return x, first_y

    def add_annotation(self, x: int, y: int, contents: str) -> None:
        self.ensure_page()
        assert self.current_page is not None
        rect = [x, y - 4, x + 16, y + 12]
        self.current_page["annots"].append({"rect": rect, "contents": contents})

    def add_block(self, label: str, body: str, comment: str | None = None) -> None:
        header_x, header_y = self.add_wrapped_text(label, font_size=self.subtitle_size, gap_after=0)
        if comment:
            self.add_annotation(header_x - 20, header_y, comment)
        self.add_wrapped_text(body, indent=12, gap_after=8)

    def build(self) -> bytes:
        objects: list[bytes] = []

        def add_object(data: bytes) -> int:
            objects.append(data)
            return len(objects)

        font_id = add_object(b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>")

        page_ids: list[int] = []
        page_object_indices: list[int] = []
        content_ids: list[int] = []
        annot_ids_per_page: list[list[int]] = []

        for page in self.pages:
            stream_text = "\n".join(page["ops"]) + "\n"
            stream_bytes = stream_text.encode("latin-1", errors="replace")
            content_id = add_object(
                b"<< /Length "
                + str(len(stream_bytes)).encode("ascii")
                + b" >>\nstream\n"
                + stream_bytes
                + b"endstream"
            )
            content_ids.append(content_id)

            page_annot_ids: list[int] = []
            for annot in page["annots"]:
                rect = annot["rect"]
                contents = pdf_safe_text(str(annot["contents"]))
                date_text = datetime.now(timezone.utc).strftime("D:%Y%m%d%H%M%SZ")
                annot_id = add_object(
                    (
                        "<< /Type /Annot /Subtype /Text "
                        f"/Rect [{' '.join(f'{value:.2f}' for value in rect)}] "
                        f"/Contents ({contents}) "
                        "/Name /Comment "
                        "/T (DOCX Compare) "
                        f"/M ({date_text}) "
                        "/Open false "
                        "/C [1 0.94 0.35] >>"
                    ).encode("latin-1", errors="replace")
                )
                page_annot_ids.append(annot_id)
            annot_ids_per_page.append(page_annot_ids)

            page_ids.append(add_object(b""))
            page_object_indices.append(len(objects) - 1)

        kids = " ".join(f"{page_id} 0 R" for page_id in page_ids)
        pages_id = add_object(
            f"<< /Type /Pages /Count {len(page_ids)} /Kids [{kids}] >>".encode("ascii")
        )
        catalog_id = add_object(f"<< /Type /Catalog /Pages {pages_id} 0 R >>".encode("ascii"))

        for page_id, obj_index, content_id, annot_ids in zip(page_ids, page_object_indices, content_ids, annot_ids_per_page):
            annots = ""
            if annot_ids:
                annots = " /Annots [" + " ".join(f"{annot_id} 0 R" for annot_id in annot_ids) + "]"
            page_obj = (
                f"<< /Type /Page /Parent {pages_id} 0 R "
                f"/MediaBox [0 0 {self.page_width} {self.page_height}] "
                f"/Resources << /Font << /F1 {font_id} 0 R >> >> "
                f"/Contents {content_id} 0 R{annots} >>"
            ).encode("ascii")
            objects[obj_index] = page_obj

        output = bytearray(b"%PDF-1.4\n%\xE2\xE3\xCF\xD3\n")
        offsets = [0]
        for number, obj in enumerate(objects, start=1):
            offsets.append(len(output))
            output.extend(f"{number} 0 obj\n".encode("ascii"))
            output.extend(obj)
            output.extend(b"\nendobj\n")
        xref_offset = len(output)
        output.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
        output.extend(b"0000000000 65535 f \n")
        for offset in offsets[1:]:
            output.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
        output.extend(
            (
                f"trailer\n<< /Size {len(objects) + 1} /Root {catalog_id} 0 R >>\n"
                f"startxref\n{xref_offset}\n%%EOF\n"
            ).encode("ascii")
        )
        return bytes(output)


def annotate_existing_pdf(
    pdf_path: Path,
    html_comments: list[HtmlComment],
    appendix_comments: list[tuple[Block, str]],
    render_result: BrowserRenderResult,
    *,
    target_label: str = "HTML",
) -> None:
    if PdfReader is None or PdfWriter is None or Text is None:
        raise RuntimeError(
            "pypdf is not installed. Install it with `pip install pypdf` to embed PDF comments."
        )

    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    if not writer.pages:
        raise RuntimeError("The rendered PDF has no pages.")

    for comment in html_comments:
        page_number = render_result.page_numbers_by_order.get(comment.order, 0)
        if page_number >= len(writer.pages):
            continue
        page = writer.pages[page_number]
        page_width_pt = float(page.mediabox.right) - float(page.mediabox.left)
        page_height_pt = float(page.mediabox.top) - float(page.mediabox.bottom)
        rect = None
        if comment.token_index is not None:
            token_rects = render_result.token_rects_by_order.get(comment.order, [])
            if 0 <= comment.token_index < len(token_rects):
                token_rect = token_rects[comment.token_index]
                rect = (token_rect.x, token_rect.y, token_rect.width, token_rect.height)
        if rect is None:
            rect = render_result.rects_by_order.get(comment.order)
        if not rect:
            continue
        x_val, y_val, width_val, height_val = rect
        if render_result.coordinate_space == "browser_px":
            scale_x = page_width_pt / max(render_result.width_px, 1.0)
            scale_y = page_height_pt / max(render_result.height_px, 1.0)
            x_pt = x_val * scale_x
            width_pt = width_val * scale_x
            y_top_pt = page_height_pt - (y_val * scale_y)
        else:
            x_pt = x_val
            width_pt = width_val
            y_top_pt = page_height_pt - y_val
        note_w = 16
        note_h = 16
        note_rect = (
            max(6, x_pt - note_w - 2),
            max(6, y_top_pt - note_h),
            max(22, x_pt - 2),
            max(22, y_top_pt),
        )
        annotation = Text(
            text=comment.contents,
            rect=note_rect,
            open=False,
        )
        text_annotation = writer.add_annotation(page_number=page_number, annotation=annotation)
        popup_rect = (
            min(page_width_pt - 20, x_pt + max(width_pt, 180)),
            max(20, y_top_pt - 140),
            min(page_width_pt - 10, x_pt + max(width_pt, 360)),
            max(80, y_top_pt - 20),
        )
        writer.add_annotation(
            page_number=page_number,
            annotation=Popup(
                rect=popup_rect,
                open=False,
                parent=text_annotation,
            ),
        )

    if appendix_comments:
        first_page = writer.pages[0]
        first_page_height_pt = float(first_page.mediabox.top) - float(first_page.mediabox.bottom)
        lines = [f"- {shorten(block.text, 160)}" for block, _comment in appendix_comments[:20]]
        if len(appendix_comments) > 20:
            lines.append(f"- ... and {len(appendix_comments) - 20} more DOCX-only blocks")
        summary_text = (
            f"DOCX content with no corresponding {target_label} block:\n"
            + "\n".join(lines)
        )
        writer.add_annotation(
            page_number=0,
            annotation=Text(
                text=summary_text,
                rect=(18, first_page_height_pt - 26, 34, first_page_height_pt - 10),
                open=False,
            ),
        )

    with pdf_path.open("wb") as handle:
        writer.write(handle)


def render_pdf(
    html_blocks: list[Block],
    html_comments: list[HtmlComment],
    appendix_comments: list[tuple[Block, str]],
    output_path: Path,
    docx_path: Path,
    html_path: Path,
    *,
    target_label: str = "HTML",
) -> None:
    grouped_comments = group_html_comments(html_comments)
    pdf = PdfBuilder()
    title = "Webpage PDF With DOCX Difference Comments" if target_label == "HTML" else "PDF With DOCX Difference Comments"
    pdf.add_wrapped_text(title, font_size=16, gap_after=6)
    pdf.add_wrapped_text(
        f"{target_label} source: {html_path.name} | DOCX source: {docx_path.name}",
        font_size=10,
        gap_after=14,
    )
    pdf.add_wrapped_text(
        "Each yellow note icon is an embedded PDF comment that flags text, formatting, or presence differences versus the DOCX source.",
        font_size=10,
        gap_after=14,
    )

    for block in html_blocks:
        tag_bits = []
        if block.heading:
            tag_bits.append(f"heading{block.heading_level or ''}")
        if block.list_item:
            tag_bits.append("list")
        if block.table_cell:
            tag_bits.append("table")
        if block.bold:
            tag_bits.append("bold")
        if block.italic:
            tag_bits.append("italic")
        if block.underline:
            tag_bits.append("underline")
        meta = f"[{target_label} {block.order + 1}]"
        if tag_bits:
            meta += " " + ", ".join(tag_bits)
        comments = grouped_comments.get(block.order, [])
        pdf.add_block(meta, block.text, "\n\n".join(comments) if comments else None)

    if appendix_comments:
        pdf.new_page()
        pdf.add_wrapped_text("DOCX-Only Content Appendix", font_size=16, gap_after=8)
        pdf.add_wrapped_text(
            f"These blocks were present in the DOCX but had no match in the {target_label.lower()} content. Each block below also carries an embedded PDF comment.",
            font_size=10,
            gap_after=14,
        )
        for index, (block, comment) in enumerate(appendix_comments, start=1):
            pdf.add_block(f"[DOCX only {index}]", block.text, comment)

    output_path.write_bytes(pdf.build())


def build_summary(
    docx_blocks: list[Block],
    html_blocks: list[Block],
    matches: list[Match],
    unmatched_docx: list[Block],
    unmatched_html: list[Block],
) -> dict[str, object]:
    exact = sum(1 for match in matches if match.match_type == "exact")
    approx = sum(1 for match in matches if match.match_type == "approx")
    return {
        "docx_blocks": len(docx_blocks),
        "html_blocks": len(html_blocks),
        "exact_matches": exact,
        "approx_matches": approx,
        "docx_only": len(unmatched_docx),
        "html_only": len(unmatched_html),
        "matches": [asdict(match) for match in matches],
        "docx_only_blocks": [asdict(block) for block in unmatched_docx],
        "html_only_blocks": [asdict(block) for block in unmatched_html],
    }


def default_output_name(html_path: Path) -> str:
    return f"{html_path.stem}__docx_diff_comments.pdf"


def default_pdf_output_name(pdf_path: Path) -> str:
    return f"{pdf_path.stem}__docx_diff_comments.pdf"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Render a saved HTML webpage into a PDF and embed comments for differences versus a DOCX file."
    )
    parser.add_argument(
        "--mode",
        choices=("html", "pdf"),
        default="html",
        help="Compare a DOCX against either a saved HTML page or an existing PDF.",
    )
    parser.add_argument(
        "--docx",
        type=Path,
        default=Path("2026-02-25 SNPS_Q1'26_EarningsRelease_Final.docx"),
        help="Source DOCX file.",
    )
    parser.add_argument(
        "--html",
        type=Path,
        default=Path("Q1'26 Earnings Release Proof_022426 4pm.html"),
        help="Source saved HTML webpage file.",
    )
    parser.add_argument(
        "--pdf",
        type=Path,
        default=None,
        help="Source PDF file for DOCX-vs-PDF compare mode.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Destination PDF file. Defaults to <html-stem>__docx_diff_comments.pdf.",
    )
    parser.add_argument(
        "--summary-json",
        type=Path,
        default=None,
        help="Optional JSON summary path.",
    )
    parser.add_argument(
        "--renderer",
        choices=("auto", "playwright", "simple"),
        default="auto",
        help="PDF rendering mode. Use `playwright` for browser-faithful rendering on Windows/macOS/Linux.",
    )
    return parser.parse_args()


def run_compare(
    *,
    docx_path: Path,
    html_path: Path,
    output_path: Path | None = None,
    summary_json_path: Path | None = None,
    renderer: str = "auto",
) -> dict[str, object]:
    output_path = output_path or Path(default_output_name(html_path))
    docx_blocks = extract_docx_blocks(docx_path)
    use_playwright = renderer == "playwright" or (
        renderer == "auto" and sync_playwright is not None and PdfReader is not None
    )

    render_result: BrowserRenderResult | None = None
    if use_playwright:
        render_result = browser_render_and_extract(html_path, output_path)
        html_blocks = render_result.blocks
    else:
        html_blocks = extract_html_blocks(html_path)

    matches, unmatched_docx, unmatched_html = compare_blocks(docx_blocks, html_blocks)
    html_comments, appendix_comments = build_comments(
        docx_blocks,
        html_blocks,
        matches,
        unmatched_docx,
        unmatched_html,
        target_label="HTML",
    )

    if use_playwright:
        annotate_existing_pdf(
            pdf_path=output_path,
            html_comments=html_comments,
            appendix_comments=appendix_comments,
            render_result=render_result,
            target_label="HTML",
        )
    else:
        render_pdf(
            html_blocks=html_blocks,
            html_comments=html_comments,
            appendix_comments=appendix_comments,
            output_path=output_path,
            docx_path=docx_path,
            html_path=html_path,
            target_label="HTML",
        )

    summary = build_summary(docx_blocks, html_blocks, matches, unmatched_docx, unmatched_html)
    summary["renderer"] = "playwright" if use_playwright else "simple"
    summary["target_kind"] = "html"
    summary["output_pdf"] = str(output_path)
    if summary_json_path:
        summary_json_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")
        summary["summary_json"] = str(summary_json_path)
    return summary


def run_compare_pdf(
    *,
    docx_path: Path,
    pdf_path: Path,
    output_path: Path | None = None,
    summary_json_path: Path | None = None,
) -> dict[str, object]:
    output_path = output_path or Path(default_pdf_output_name(pdf_path))
    docx_blocks = extract_docx_blocks(docx_path)
    render_result = extract_pdf_blocks(pdf_path)
    pdf_blocks = render_result.blocks

    matches, unmatched_docx, unmatched_pdf = compare_blocks(docx_blocks, pdf_blocks)
    pdf_comments, appendix_comments = build_comments(
        docx_blocks,
        pdf_blocks,
        matches,
        unmatched_docx,
        unmatched_pdf,
        target_label="PDF",
    )
    pdf_comments.extend(
        pdf_page_summary_comments(
            docx_blocks=docx_blocks,
            pdf_blocks=pdf_blocks,
            unmatched_pdf=unmatched_pdf,
            matches=matches,
            render_result=render_result,
        )
    )
    output_path.write_bytes(pdf_path.read_bytes())
    annotate_existing_pdf(
        pdf_path=output_path,
        html_comments=pdf_comments,
        appendix_comments=appendix_comments,
        render_result=render_result,
        target_label="PDF",
    )

    summary = build_summary(docx_blocks, pdf_blocks, matches, unmatched_docx, unmatched_pdf)
    summary["renderer"] = "pdf"
    summary["target_kind"] = "pdf"
    summary["output_pdf"] = str(output_path)
    if summary_json_path:
        summary_json_path.write_text(json.dumps(summary, indent=2), encoding="utf-8")
        summary["summary_json"] = str(summary_json_path)
    return summary


def main() -> int:
    args = parse_args()
    if args.mode == "pdf":
        if args.pdf is None:
            raise SystemExit("--pdf is required when --mode pdf is used.")
        summary = run_compare_pdf(
            docx_path=args.docx,
            pdf_path=args.pdf,
            output_path=args.output,
            summary_json_path=args.summary_json,
        )
    else:
        summary = run_compare(
            docx_path=args.docx,
            html_path=args.html,
            output_path=args.output,
            summary_json_path=args.summary_json,
            renderer=args.renderer,
        )

    print(f"PDF written: {summary['output_pdf']}")
    print(
        "Summary: "
        f"{summary['docx_blocks']} DOCX blocks, "
        f"{summary['html_blocks']} HTML blocks, "
        f"{summary['exact_matches']} exact matches, "
        f"{summary['approx_matches']} approximate matches, "
        f"{summary['docx_only']} DOCX-only, "
        f"{summary['html_only']} HTML-only."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
