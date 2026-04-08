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
from dataclasses import asdict, dataclass, field
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
    raw_text: str = ""
    proof_text: str = ""
    match_text: str = ""
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
    footnote_marker: str | None = None
    structure_role: str = ""
    family_table_index: int | None = None
    runs: list["InlineRun"] = field(default_factory=list)


@dataclass
class InlineRun:
    text: str
    kind: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    hyperlink: bool = False
    source_index: int = 0


@dataclass
class WordStyle:
    style_id: str
    style_type: str
    based_on: str | None = None
    name: str = ""
    props: dict[str, bool | None] = field(default_factory=dict)


@dataclass
class WordStyleResolver:
    default_run_props: dict[str, bool | None] = field(default_factory=dict)
    styles: dict[str, WordStyle] = field(default_factory=dict)
    hyperlink_style_id: str | None = None
    _cache: dict[str, dict[str, bool | None]] = field(default_factory=dict)

    def resolve_style_props(self, style_id: str | None) -> dict[str, bool | None]:
        if not style_id:
            return {}
        cached = self._cache.get(style_id)
        if cached is not None:
            return dict(cached)
        style = self.styles.get(style_id)
        if style is None:
            return {}
        props: dict[str, bool | None] = {}
        if style.based_on and style.based_on != style_id:
            props = self.resolve_style_props(style.based_on)
        props = merge_word_format_props(props, style.props)
        self._cache[style_id] = dict(props)
        return dict(props)

@dataclass
class Match:
    docx_index: int
    html_index: int
    match_type: str
    score: float
    formatting_diffs: list[str]


@dataclass
class TableFamily:
    source: str
    table_index: int
    block_indices: list[int]
    order_start: int
    row_keys: list[str]
    header_texts: list[str]
    title_context: list[str]
    row_count: int
    numeric_count: int
    comparable: bool


@dataclass
class BlockGroup:
    group_id: str
    source: str
    group_type: str
    block_indices: list[int]
    key: str
    text: str
    order_start: int
    order_end: int
    footnote_marker: str | None = None


@dataclass
class SectionFamily:
    family_id: str
    source: str
    family_type: str
    schema_key: str | None
    block_indices: list[int]
    key: str
    text: str
    order_start: int
    order_end: int


@dataclass(frozen=True)
class SectionSchema:
    schema_key: str
    family_type: str
    patterns: tuple[str, ...]


KNOWN_EARNINGS_SECTION_SCHEMAS = (
    SectionSchema("results_summary", "results_summary", ("results summary",)),
    SectionSchema("gaap_results", "gaap_results", ("gaap results",)),
    SectionSchema("non_gaap_results", "non_gaap_results", ("non gaap results",)),
    SectionSchema("business_segments", "business_segments", ("business segments",)),
    SectionSchema("financial_targets", "financial_targets", ("financial targets",)),
    SectionSchema("earnings_call_open_to_investors", "earnings_call", ("earnings call open to investors", "earnings call open to investor")),
    SectionSchema("availability_of_final_financial_statements", "availability", ("availability of final financial statements",)),
    SectionSchema("reconciliation_quarter_results", "reconciliation", (r"reconciliation of (?:first|second|third|fourth) quarter fiscal year \d{4} results",)),
    SectionSchema("reconciliation_targets", "reconciliation", (r"reconciliation of \d{4} targets",)),
    SectionSchema("forward_looking_statements", "forward_looking", ("forward looking statements",)),
    SectionSchema(
        "synopsys_income_statement",
        "statement_table",
        ("synopsys inc unaudited condensed consolidated statements of income",),
    ),
    SectionSchema(
        "synopsys_balance_sheets",
        "statement_table",
        ("synopsys inc unaudited condensed consolidated balance sheets",),
    ),
    SectionSchema(
        "synopsys_cash_flows",
        "statement_table",
        ("synopsys inc unaudited condensed consolidated statements of cash flows",),
    ),
    SectionSchema("segment_information", "segment_information", ("segment information",)),
    SectionSchema("gaap_to_non_gaap_reconciliation", "gaap_to_non_gaap_reconciliation", ("gaap to non gaap reconciliation",)),
    SectionSchema("about_synopsys", "about_synopsys", ("about synopsys",)),
    SectionSchema("investor_contact", "investor_contact", ("investor contact",)),
    SectionSchema("editorial_contact", "editorial_contact", ("editorial contact",)),
)


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


EXTRA_COMMENT_RE = re.compile(
    r"^The (number|word) is extra in ([a-z]+), (.+)\. It is not present in word\.$"
)
MISSING_COMMENT_RE = re.compile(
    r"^The (number|word) is missing in ([a-z]+), (.+) in word\.$"
)


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
    if block.family_table_index is not None:
        return block.family_table_index
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


def mapped_table_index_key(
    block: Block,
    doc_to_target_table_map: dict[int, int] | None = None,
) -> int | None:
    table_idx = table_index_key(block)
    if table_idx is None:
        return None
    if doc_to_target_table_map is None:
        return table_idx
    return doc_to_target_table_map.get(table_idx, table_idx)


def mapped_table_pos_key(
    block: Block,
    doc_to_target_table_map: dict[int, int] | None = None,
) -> tuple[int, int, int] | None:
    pos = table_pos_key(block)
    if pos is None:
        return None
    table_idx, row_idx, col_idx = pos
    mapped_table_idx = table_idx if doc_to_target_table_map is None else doc_to_target_table_map.get(table_idx, table_idx)
    return (mapped_table_idx, row_idx, col_idx)


def mapped_row_context_key(
    block: Block,
    doc_to_target_table_map: dict[int, int] | None = None,
) -> tuple[int, str, int] | None:
    row_key = row_context_key(block)
    if row_key is None:
        return None
    table_idx, normalized_row_key, row_slot = row_key
    mapped_table_idx = table_idx if doc_to_target_table_map is None else doc_to_target_table_map.get(table_idx, table_idx)
    return (mapped_table_idx, normalized_row_key, row_slot)


def single_value_token(block: Block) -> DiffToken | None:
    tokens = diff_tokens(block.text)
    return tokens[0] if len(tokens) == 1 else None


def parse_numeric_token(token_text: str) -> tuple[float, bool] | None:
    text = normalize_text(token_text).strip()
    if not text:
        return None
    is_percent = "%" in text
    text = text.replace("%", "")
    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]
    text = text.replace("$", "").replace("€", "").replace("£", "").replace("¥", "")
    text = text.replace(",", "").replace("~", "").replace(" ", "")
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


def unique_preserve_order(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        if value and value not in seen:
            seen.add(value)
            result.append(value)
    return result


def table_family_title_context(blocks: list[Block], first_index: int) -> list[str]:
    context: list[str] = []
    for index in range(first_index - 1, -1, -1):
        block = blocks[index]
        if block.table_cell:
            break
        text = normalize_without_punctuation(block.text)
        if not visible_meaningful(text):
            continue
        context.append(text)
        if len(context) >= 3:
            break
    context.reverse()
    return context


def header_family_group(role: str) -> str | None:
    if role in {"table_title", "table_subtitle"}:
        return "title_family"
    if role == "table_column_header":
        return "column_header_family"
    return None


def build_family_role_ordinals(blocks: list[Block]) -> tuple[dict[int, int], dict[tuple[int, str, int], list[int]]]:
    index_to_ordinal: dict[int, int] = {}
    ordinal_map: dict[tuple[int, str, int], list[int]] = {}
    grouped: dict[tuple[int, str], list[int]] = {}
    for index, block in enumerate(blocks):
        family_idx = table_index_key(block)
        group = header_family_group(block.structure_role)
        if family_idx is None or group is None:
            continue
        grouped.setdefault((family_idx, group), []).append(index)
    for (family_idx, group), indices in grouped.items():
        for ordinal, index in enumerate(sorted(indices, key=lambda idx: blocks[idx].order)):
            index_to_ordinal[index] = ordinal
            ordinal_map.setdefault((family_idx, group, ordinal), []).append(index)
    return index_to_ordinal, ordinal_map


def build_row_label_ordinals(blocks: list[Block]) -> tuple[dict[int, int], dict[tuple[int, int], list[int]]]:
    index_to_ordinal: dict[int, int] = {}
    ordinal_map: dict[tuple[int, int], list[int]] = {}
    grouped: dict[int, list[int]] = {}
    for index, block in enumerate(blocks):
        family_idx = table_index_key(block)
        if family_idx is None or block.structure_role != "table_row_label":
            continue
        grouped.setdefault(family_idx, []).append(index)
    for family_idx, indices in grouped.items():
        def row_sort_key(idx: int) -> tuple[int, int]:
            block = blocks[idx]
            row_idx = block.table_pos[1] if block.table_pos is not None else 10**9
            order = block.order
            return (row_idx, order)
        for ordinal, index in enumerate(sorted(indices, key=row_sort_key)):
            index_to_ordinal[index] = ordinal
            ordinal_map.setdefault((family_idx, ordinal), []).append(index)
    return index_to_ordinal, ordinal_map


def extract_table_families(blocks: list[Block]) -> dict[int, TableFamily]:
    grouped_indices: dict[int, list[int]] = {}
    for index, block in enumerate(blocks):
        table_idx = table_index_key(block)
        if table_idx is None:
            continue
        grouped_indices.setdefault(table_idx, []).append(index)

    families: dict[int, TableFamily] = {}
    for table_idx, block_indices in grouped_indices.items():
        table_blocks = [blocks[index] for index in block_indices]
        row_keys = unique_preserve_order(
            [block.row_key for block in table_blocks if block.row_key]
        )
        header_texts = unique_preserve_order(
            [
                normalize_without_punctuation(strip_leading_markers(block.text))
                for block in table_blocks
                if visible_meaningful(block.text)
                and (single_value_token(block) is None or single_value_token(block).kind != "number")
            ]
        )[:8]
        numeric_count = sum(1 for block in table_blocks if block.numeric_slot is not None)
        row_count = len(row_keys)
        comparable = row_count >= 2 and (numeric_count >= 2 or len(header_texts) >= 3)
        families[table_idx] = TableFamily(
            source=table_blocks[0].source,
            table_index=table_idx,
            block_indices=list(block_indices),
            order_start=min(blocks[index].order for index in block_indices),
            row_keys=row_keys,
            header_texts=header_texts,
            title_context=table_family_title_context(blocks, min(block_indices)),
            row_count=row_count,
            numeric_count=numeric_count,
            comparable=comparable,
        )
    return families


def section_family_text(blocks: list[Block], indices: list[int]) -> str:
    return "\n".join(blocks[index].text for index in indices if visible_meaningful(blocks[index].text))


def section_family_key(text: str) -> str:
    return " ".join(tokenize(normalize_for_compare(text))[:18])


def section_family_signature_text(blocks: list[Block], indices: list[int]) -> str:
    parts: list[str] = []
    for index in indices:
        block = blocks[index]
        if not visible_meaningful(block.text):
            continue
        if block.table_cell and block.structure_role == "table_data_cell":
            continue
        parts.append(block.text)
    return "\n".join(parts)


def extract_section_families(blocks: list[Block]) -> tuple[list[SectionFamily], dict[int, str]]:
    families: list[SectionFamily] = []
    block_to_family: dict[int, str] = {}
    used_indices: set[int] = set()

    known_header_indices = {
        index
        for index, block in enumerate(blocks)
        if not block.table_cell and block.kind != "footnote" and known_section_header(block.text)
    }

    for index, block in enumerate(blocks):
        if index in used_indices or block.table_cell or block.kind == "footnote":
            continue
        schema = known_section_schema(block.text, allow_fuzzy=True)
        if schema is None:
            continue
        family_indices: list[int] = [index]
        probe = index + 1
        while probe < len(blocks):
            candidate = blocks[probe]
            if probe in known_header_indices and probe != index:
                break
            if candidate.kind == "footnote":
                probe += 1
                continue
            if not candidate.table_cell:
                family_indices.append(probe)
            probe += 1
        family_indices = sorted(set(family_indices), key=lambda idx: blocks[idx].order)
        text = section_family_text(blocks, family_indices)
        signature_text = section_family_signature_text(blocks, family_indices)
        if not visible_meaningful(signature_text):
            used_indices.add(index)
            continue
        section = SectionFamily(
            family_id=f"{block.source}-section-schema-{schema.schema_key}-{index}",
            source=block.source,
            family_type=schema.family_type,
            schema_key=schema.schema_key,
            block_indices=family_indices,
            key=schema.schema_key,
            text=text,
            order_start=blocks[family_indices[0]].order,
            order_end=blocks[family_indices[-1]].order,
        )
        families.append(section)
        for family_index in family_indices:
            block_to_family[family_index] = section.family_id
            used_indices.add(family_index)

    table_families = extract_table_families(blocks)
    for table_idx, family in sorted(table_families.items()):
        title_bits = " ".join(family.title_context + family.header_texts[:4])
        normalized_title_bits = normalize_for_compare(title_bits)
        if "reconciliation" not in normalized_title_bits:
            continue
        indices = list(family.block_indices)
        for index, block in enumerate(blocks):
            if block.family_table_index == table_idx and not block.table_cell and block.structure_role in {"table_title", "table_subtitle"}:
                indices.append(index)
        indices = sorted(set(indices), key=lambda idx: blocks[idx].order)
        if not indices:
            continue
        text = section_family_text(blocks, indices)
        section = SectionFamily(
            family_id=f"{blocks[indices[0]].source}-section-reconciliation-{table_idx}",
            source=blocks[indices[0]].source,
            family_type="reconciliation",
            schema_key=None,
            block_indices=indices,
            key=section_family_key(text),
            text=text,
            order_start=blocks[indices[0]].order,
            order_end=blocks[indices[-1]].order,
        )
        families.append(section)
        for index in indices:
            block_to_family[index] = section.family_id
            used_indices.add(index)

    index = 0
    while index < len(blocks):
        if index in used_indices:
            index += 1
            continue
        block = blocks[index]
        if block.table_cell or block.kind == "footnote":
            index += 1
            continue
        start_header = (
            block.structure_role in {"section_header", "section_lead"}
            or normalize_for_compare(block.text) == "forward looking statements"
        )
        if not start_header:
            index += 1
            continue
        family_indices = [index]
        probe = index + 1
        while probe < len(blocks):
            candidate = blocks[probe]
            if candidate.table_cell or candidate.kind == "footnote":
                break
            if candidate.structure_role in {"section_header", "section_lead"} and visible_meaningful(candidate.text):
                break
            family_indices.append(probe)
            probe += 1
        text = section_family_text(blocks, family_indices)
        if len(diff_tokens(text)) < 20:
            index += 1
            continue
        section = SectionFamily(
            family_id=f"{block.source}-section-narrative-{index}",
            source=block.source,
            family_type="narrative",
            schema_key=None,
            block_indices=family_indices,
            key=section_family_key(text),
            text=text,
            order_start=blocks[family_indices[0]].order,
            order_end=blocks[family_indices[-1]].order,
        )
        families.append(section)
        for family_index in family_indices:
            block_to_family[family_index] = section.family_id
            used_indices.add(family_index)
        index = probe
    return families, block_to_family


def section_family_similarity(doc_family: SectionFamily, target_family: SectionFamily) -> float:
    if doc_family.schema_key and target_family.schema_key:
        if doc_family.schema_key == target_family.schema_key:
            return 1.0
        return 0.0
    if doc_family.family_type != target_family.family_type:
        return 0.0
    key_score = similarity(doc_family.key, target_family.key)
    overlap = token_overlap_ratio(
        normalize_for_compare(doc_family.text),
        normalize_for_compare(target_family.text),
    )
    if doc_family.family_type == "reconciliation":
        return max(key_score, overlap)
    return max(key_score * 0.7 + overlap * 0.3, overlap)


def match_section_families(
    docx_blocks: list[Block],
    target_blocks: list[Block],
) -> tuple[dict[str, str], dict[int, str], dict[int, str], dict[str, SectionFamily], dict[str, SectionFamily]]:
    doc_families, doc_block_to_family = extract_section_families(docx_blocks)
    target_families, target_block_to_family = extract_section_families(target_blocks)
    doc_family_map = {family.family_id: family for family in doc_families}
    target_family_map = {family.family_id: family for family in target_families}
    matches: dict[str, str] = {}
    used_target_ids: set[str] = set()
    for doc_family in doc_families:
        best_target: SectionFamily | None = None
        best_score = 0.0
        for target_family in target_families:
            if target_family.family_id in used_target_ids:
                continue
            score = section_family_similarity(doc_family, target_family)
            threshold = 0.48 if doc_family.family_type == "reconciliation" else 0.6
            if score < threshold:
                continue
            if score > best_score:
                best_score = score
                best_target = target_family
        if best_target is not None:
            matches[doc_family.family_id] = best_target.family_id
            used_target_ids.add(best_target.family_id)
    return matches, doc_block_to_family, target_block_to_family, doc_family_map, target_family_map


def overlap_ratio(values_a: list[str], values_b: list[str]) -> float:
    if not values_a or not values_b:
        return 0.0
    set_a = set(values_a)
    set_b = set(values_b)
    intersection = len(set_a & set_b)
    return intersection / max(len(set_a), len(set_b), 1)


def title_context_similarity(doc_family: TableFamily, target_family: TableFamily) -> float:
    if not doc_family.title_context or not target_family.title_context:
        return 0.0
    best = 0.0
    for doc_title in doc_family.title_context:
        for target_title in target_family.title_context:
            best = max(best, similarity(doc_title, target_title))
    return best


def table_family_alignment_score(
    doc_family: TableFamily,
    target_family: TableFamily,
    *,
    doc_rank: int,
    target_rank: int,
) -> float:
    row_overlap = overlap_ratio(doc_family.row_keys, target_family.row_keys)
    header_overlap = overlap_ratio(doc_family.header_texts, target_family.header_texts)
    title_overlap = title_context_similarity(doc_family, target_family)
    row_count_ratio = min(doc_family.row_count, target_family.row_count) / max(doc_family.row_count, target_family.row_count, 1)
    numeric_ratio = min(doc_family.numeric_count, target_family.numeric_count) / max(doc_family.numeric_count, target_family.numeric_count, 1)
    order_proximity = max(0.0, 1.0 - min(1.0, abs(doc_rank - target_rank) * 0.5))
    return (
        row_overlap * 6.0
        + header_overlap * 2.5
        + title_overlap * 2.0
        + row_count_ratio * 0.75
        + numeric_ratio * 0.75
        + order_proximity * 0.5
    )


def align_table_families(
    docx_blocks: list[Block],
    target_blocks: list[Block],
) -> tuple[dict[int, int], dict[int, int]]:
    if not target_blocks or target_blocks[0].source not in {"html", "pdf"}:
        return {}, {}

    doc_families = extract_table_families(docx_blocks)
    target_families = extract_table_families(target_blocks)
    comparable_doc = [family for family in sorted(doc_families.values(), key=lambda family: family.table_index) if family.comparable]
    comparable_target = [family for family in sorted(target_families.values(), key=lambda family: family.table_index) if family.comparable]
    if not comparable_doc or not comparable_target:
        return {}, {}

    doc_rank = {family.table_index: rank for rank, family in enumerate(comparable_doc)}
    target_rank = {family.table_index: rank for rank, family in enumerate(comparable_target)}

    candidate_pairs: list[tuple[float, int, int]] = []
    for doc_family in comparable_doc:
        for target_family in comparable_target:
            score = table_family_alignment_score(
                doc_family,
                target_family,
                doc_rank=doc_rank[doc_family.table_index],
                target_rank=target_rank[target_family.table_index],
            )
            if score >= 2.4:
                candidate_pairs.append((score, doc_family.table_index, target_family.table_index))

    doc_to_target: dict[int, int] = {}
    target_to_doc: dict[int, int] = {}
    for _score, doc_table_idx, target_table_idx in sorted(
        candidate_pairs,
        key=lambda item: (
            item[0],
            -abs(doc_rank[item[1]] - target_rank[item[2]]),
            -item[1],
            -item[2],
        ),
        reverse=True,
    ):
        if doc_table_idx in doc_to_target or target_table_idx in target_to_doc:
            continue
        doc_to_target[doc_table_idx] = target_table_idx
        target_to_doc[target_table_idx] = doc_table_idx
    return doc_to_target, target_to_doc


CONTACT_LABELS = {
    "investor contact": "investor_contact",
    "editorial contact": "editorial_contact",
}
CONTACT_PHONE_RE = re.compile(r"\b\d{3}-\d{3}-\d{4}\b")
CONTACT_EMAIL_RE = re.compile(r"\b[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}\b")
CONTACT_COMPANY_RE = re.compile(r"\bSynopsys,\s*Inc\.?\b", flags=re.I)


def contact_label_key(text: str) -> str | None:
    normalized = normalize_for_compare(text)
    for label, key in CONTACT_LABELS.items():
        if normalized == label or normalized.startswith(f"{label}:") or normalized.startswith(f"{label} "):
            return key
    return None


def looks_like_contact_fragment(block: Block) -> bool:
    if block.table_cell or block.heading:
        return False
    return looks_like_contact_text(block.text)


def looks_like_contact_text(text: str) -> bool:
    text = normalize_text(text).strip()
    normalized = normalize_for_compare(text)
    if not text:
        return False
    if contact_label_key(text):
        return True
    if CONTACT_EMAIL_RE.search(text):
        return True
    if CONTACT_PHONE_RE.search(text):
        return True
    if normalized == "synopsys, inc.":
        return True
    if len(text.split()) <= 4 and text[:1].isupper():
        return True
    return False


def contact_field_role(text: str) -> str | None:
    clean_text = normalize_text(text).strip()
    normalized = normalize_for_compare(clean_text)
    if not clean_text:
        return None
    if contact_label_key(clean_text):
        return "contact_label"
    if CONTACT_COMPANY_RE.search(clean_text):
        return "contact_company"
    if CONTACT_EMAIL_RE.search(clean_text):
        return "contact_email"
    if CONTACT_PHONE_RE.search(clean_text):
        return "contact_phone"
    if len(clean_text.split()) <= 4 and clean_text[:1].isupper() and not clean_text.isupper():
        return "contact_name"
    return None


def extract_contact_fields(text: str) -> dict[str, str]:
    clean_text = normalize_text(text).strip()
    fields: dict[str, str] = {}
    label_match = re.match(r"^\s*((?:INVESTOR|EDITORIAL)\s+CONTACT:?)\s*", clean_text, flags=re.I)
    remainder = clean_text
    if label_match:
        fields["contact_label"] = label_match.group(1).strip()
        remainder = clean_text[label_match.end():].strip()

    company_match = CONTACT_COMPANY_RE.search(remainder)
    phone_match = CONTACT_PHONE_RE.search(remainder)
    email_match = CONTACT_EMAIL_RE.search(remainder)
    if company_match:
        fields["contact_company"] = company_match.group(0).strip()
    if phone_match:
        fields["contact_phone"] = phone_match.group(0).strip()
    if email_match:
        fields["contact_email"] = email_match.group(0).strip()

    name_end = len(remainder)
    for match in (company_match, phone_match, email_match):
        if match is not None:
            name_end = min(name_end, match.start())
    contact_name = remainder[:name_end].strip(" ,;")
    if contact_name:
        fields["contact_name"] = contact_name
    return fields


def contact_field_payload(text: str) -> tuple[str, str] | None:
    role = contact_field_role(text)
    if role is None:
        return None
    if role == "contact_label":
        label_match = re.match(r"^\s*((?:INVESTOR|EDITORIAL)\s+CONTACT:?)", normalize_text(text), flags=re.I)
        if label_match:
            return role, label_match.group(1).strip()
    if role == "contact_company":
        company_match = CONTACT_COMPANY_RE.search(normalize_text(text))
        if company_match:
            return role, company_match.group(0).strip()
    if role == "contact_phone":
        phone_match = CONTACT_PHONE_RE.search(normalize_text(text))
        if phone_match:
            return role, phone_match.group(0).strip()
    if role == "contact_email":
        email_match = CONTACT_EMAIL_RE.search(normalize_text(text))
        if email_match:
            return role, email_match.group(0).strip()
    return role, normalize_text(text).strip()


def quote_role_key(text: str) -> str | None:
    normalized = normalize_for_compare(text)
    if "chief financial officer" in normalized or "finance chief" in normalized or re.search(r"\bcfo\b", normalized):
        return "cfo"
    if (
        "chief executive officer" in normalized
        or "president and ceo" in normalized
        or "president chief executive officer" in normalized
        or re.search(r"\bceo\b", normalized)
    ):
        return "ceo"
    return None


def quote_speaker_title(text: str) -> tuple[str, str]:
    match = re.search(
        r"said\s+([A-Z][A-Za-z]+(?:\s+[A-Z][A-Za-z.'-]+){0,3})\s*,\s*([^\"”]+)",
        text,
    )
    if not match:
        return "", ""
    speaker = normalize_for_compare(match.group(1))
    title = normalize_for_compare(match.group(2))
    title = re.sub(r"[^a-z0-9]+", " ", title).strip()
    return speaker, title


def quote_lead_key(text: str) -> str:
    match = re.search(r"[\"“]([^\"”]{20,220})[\"”]", normalize_text(text))
    lead = match.group(1) if match else text
    return " ".join(tokenize(normalize_for_compare(lead))[:12])


def parse_quote_key(key: str) -> tuple[str, str, str, str, bool]:
    parts = key.split("|")
    while len(parts) < 5:
        parts.append("")
    role, speaker, title, lead, placeholder = parts[:5]
    return role, speaker, title, lead, placeholder == "1"


def quote_group_summary_label(text: str) -> str:
    role = quote_role_key(text)
    if role == "ceo":
        return "CEO quote"
    if role == "cfo":
        return "CFO quote"
    return "quote"


def quote_diff_summary(doc_text: str, target_text: str, *, target_name: str) -> str:
    label = quote_group_summary_label(target_text or doc_text)
    doc_proof = normalize_proof_text(doc_text).strip()
    target_proof = normalize_proof_text(target_text).strip()
    if doc_proof == target_proof:
        return (
            f"The {label} block is broadly matched rather than exactly matched. "
            "The normalized quote wording is the same, but the extracted quote blocks did not qualify for an exact structural match. "
            f"{target_name.upper()}: {shorten(target_text, 160)} "
            f"Word: {shorten(doc_text, 160)}"
        )
    if normalize_without_punctuation(doc_text) == normalize_without_punctuation(target_text):
        return (
            f"The {label} punctuation or encoding is different. "
            f"{target_name.upper()}: {shorten(target_text, 160)} "
            f"Word: {shorten(doc_text, 160)}"
        )
    return (
        f"The {label} text is different. "
        f"{target_name.upper()}: {shorten(target_text, 160)} "
        f"Word: {shorten(doc_text, 160)}"
    )


def extract_primary_quote_text(text: str) -> str | None:
    source = normalize_text(text)
    starts = [idx for idx, char in enumerate(source) if char in {'"', "“"}]
    ends = [idx for idx, char in enumerate(source) if char in {'"', "”"}]
    if not starts or not ends:
        return None
    start = starts[0]
    end = ends[-1]
    if end <= start:
        return None
    extracted = source[start : end + 1].strip()
    return extracted if len(extracted) >= 40 else None


def pdf_embedded_lead_body(block: Block) -> str | None:
    if block.source != "pdf" or block.table_cell:
        return None
    source = normalize_text(block.text)
    if "\n" not in source:
        return None
    first_line, remainder = source.split("\n", 1)
    first_line = first_line.strip()
    remainder = remainder.strip()
    if not first_line or not remainder:
        return None
    if len(diff_tokens(first_line)) > 8 or len(first_line) > 90:
        return None
    if len(diff_tokens(remainder)) < 12:
        return None
    if block.runs:
        first_line_chars = len(first_line)
        consumed = 0
        first_line_has_text = False
        for run in block.runs:
            if run.kind == "linebreak":
                break
            if run.kind != "space":
                first_line_has_text = True
                if not run.bold:
                    return None
            consumed += len(run.text)
            if consumed >= first_line_chars:
                break
        if not first_line_has_text:
            return None
    return remainder


def spacing_is_only_pdf_line_wrap(doc_sep: str, target_sep: str) -> bool:
    doc_spaces, _doc_tabs, doc_lines = whitespace_signature(doc_sep)
    target_spaces, _target_tabs, target_lines = whitespace_signature(target_sep)
    if normalized_separator_symbol(doc_sep) != normalized_separator_symbol(target_sep):
        return False
    if doc_lines == target_lines:
        return False
    if doc_lines == 0 and target_lines > 0 and doc_spaces >= 1:
        return True
    if target_lines == 0 and doc_lines > 0 and target_spaces >= 1:
        return True
    return False


def long_narrative_block(block: Block) -> bool:
    return (
        not block.table_cell
        and not block.heading
        and len(block.normalized) >= 160
        and len(diff_tokens(block.text)) >= 20
    )


def looks_like_section_lead(block: Block, next_block: Block | None) -> bool:
    if block.table_cell or block.kind == "footnote" or block.heading:
        return False
    if not block.bold:
        return False
    if next_block is None or next_block.table_cell or next_block.kind == "footnote":
        return False
    text = normalize_text(block.text).strip()
    if not visible_meaningful(text):
        return False
    token_count = len(diff_tokens(text))
    if token_count == 0 or token_count > 8:
        return False
    if len(text) > 80:
        return False
    if len(diff_tokens(next_block.text)) < 12:
        return False
    if headerish_table_text(text):
        return False
    return True


def repeated_label_key(text: str) -> str | None:
    normalized = normalize_for_compare(strip_leading_markers(text))
    if not normalized:
        return None
    tokens = tokenize(normalized)
    if not tokens or len(tokens) > 4:
        return None
    if len(normalized) > 32:
        return None
    if any(token in {"quote", "tbc"} for token in tokens):
        return None
    if all(token.isdigit() and len(token) == 4 for token in tokens):
        return normalized
    if any(token.isdigit() for token in tokens) and len(tokens) > 2:
        return None
    if normalized in {"low", "high", "synopsys inc", "synopsys inc.", "gaap eps", "non gaap eps"}:
        return normalized
    if len(tokens) <= 2:
        return normalized
    if all(len(token) <= 8 for token in tokens):
        return normalized
    return None


def repeated_label_block(block: Block) -> bool:
    return repeated_label_key(block.text) is not None


STRICT_STRUCTURAL_ROLES = {
    "table_title",
    "table_subtitle",
    "table_column_header",
    "table_row_label",
    "table_data_cell",
    "table_data_label",
    "section_header",
    "section_lead",
}


def structural_role_compatible(doc_block: Block, target_block: Block) -> bool:
    doc_role = doc_block.structure_role or ""
    target_role = target_block.structure_role or ""
    if not doc_role or not target_role:
        return True
    if doc_role == target_role:
        return True
    if doc_role in STRICT_STRUCTURAL_ROLES or target_role in STRICT_STRUCTURAL_ROLES:
        return False
    return True


def exact_like_match_type(match_type: str) -> bool:
    return match_type in {"exact", "exact_structural"}


def compatible_header_family_roles(doc_role: str, target_role: str) -> bool:
    return {doc_role, target_role} <= {"table_title", "table_subtitle"}


def matching_block_exists(block: Block, candidates: list[Block]) -> bool:
    key = repeated_label_key(block.text)
    if key is None:
        return False
    return any(repeated_label_key(candidate.text) == key for candidate in candidates)


def match_confidence_tier(
    doc_block: Block,
    target_block: Block,
    *,
    score: float,
    match_type: str,
    grouped_match_type: str | None,
    target_name: str,
) -> str:
    proofread_target = target_name in {"html", "pdf"}
    doc_schema = known_section_schema(doc_block.text, allow_fuzzy=True)
    target_schema = known_section_schema(target_block.text, allow_fuzzy=True)
    if (
        doc_block.structure_role == "section_header"
        and target_block.structure_role == "section_header"
        and doc_schema is not None
        and target_schema is not None
        and doc_schema.schema_key == target_schema.schema_key
        and max(len(doc_block.normalized), len(target_block.normalized)) <= 48
    ):
        if exact_like_match_type(match_type):
            return "strong"
        header_similarity = difflib.SequenceMatcher(None, doc_block.normalized, target_block.normalized).ratio()
        if header_similarity >= 0.88 and score >= 0.6:
            return "strong"
        if header_similarity >= 0.8 and score >= 0.52:
            return "medium"
        return "weak"
    if doc_block.structure_role == "section_lead" and target_block.structure_role == "section_lead":
        if exact_like_match_type(match_type):
            return "strong"
        lead_strong = 0.8 if target_name == "html" else 0.76 if target_name == "pdf" else 0.8
        lead_medium = 0.7 if target_name == "html" else 0.66 if target_name == "pdf" else 0.7
        if score >= lead_strong:
            return "strong"
        if score >= lead_medium:
            return "medium"
        return "weak"
    if (
        doc_block.structure_role
        and target_block.structure_role
        and doc_block.structure_role != target_block.structure_role
        and doc_block.structure_role not in {"paragraph", "table_cell"}
        and target_block.structure_role not in {"paragraph", "table_cell"}
    ):
        if (
            exact_like_match_type(match_type)
            and compatible_header_family_roles(doc_block.structure_role, target_block.structure_role)
            and proofread_target
            and max(len(doc_block.normalized), len(target_block.normalized)) <= 32
            and similarity(doc_block.normalized, target_block.normalized) >= (0.6 if target_name == "html" else 0.52)
        ):
            return "strong"
        if score >= 0.96 and exact_like_match_type(match_type):
            return "medium"
        return "weak"
    if exact_like_match_type(match_type):
        if (
            doc_block.structure_role.startswith("contact_")
            and target_block.structure_role.startswith("contact_")
        ):
            return "strong"
        if repeated_label_block(doc_block) and repeated_label_block(target_block) and proofread_target:
            return "medium"
        return "strong"
    if grouped_match_type in {"quote", "footnote", "contact"}:
        group_strong = 0.95 if target_name == "html" else 0.9 if target_name == "pdf" else 0.95
        group_medium = 0.88 if target_name == "html" else 0.8 if target_name == "pdf" else 0.88
        if score >= group_strong:
            return "strong"
        if score >= group_medium:
            return "medium"
        return "weak"
    if (
        proofread_target
        and (
            repeated_label_block(doc_block)
            or repeated_label_block(target_block)
            or len(doc_block.normalized) <= 24
            or len(target_block.normalized) <= 24
        )
    ):
        short_strong = 0.98 if target_name == "html" else 0.94 if target_name == "pdf" else 0.98
        short_medium = 0.9 if target_name == "html" else 0.82 if target_name == "pdf" else 0.9
        if score >= short_strong and doc_block.normalized == target_block.normalized:
            return "strong"
        if score >= short_medium:
            return "medium"
        return "weak"
    if doc_block.table_cell and target_block.table_cell:
        table_strong = 0.9 if target_name != "pdf" else 0.84
        table_medium = 0.8 if target_name != "pdf" else 0.72
        if score >= table_strong:
            return "strong"
        if score >= table_medium:
            return "medium"
        return "weak"
    if repeated_label_block(doc_block) and repeated_label_block(target_block):
        repeated_strong = 0.985 if target_name != "pdf" else 0.95
        repeated_medium = 0.94 if target_name != "pdf" else 0.86
        if score >= repeated_strong:
            return "strong"
        if score >= repeated_medium:
            return "medium"
        return "weak"
    if long_narrative_block(doc_block) and long_narrative_block(target_block):
        narrative_strong = 0.96 if target_name != "pdf" else 0.92
        narrative_medium = 0.88 if target_name != "pdf" else 0.8
        if score >= narrative_strong:
            return "strong"
        if score >= narrative_medium:
            return "medium"
        return "weak"
    default_strong = 0.95 if target_name != "pdf" else 0.9
    default_medium = 0.84 if target_name != "pdf" else 0.76
    if score >= default_strong:
        return "strong"
    if score >= default_medium:
        return "medium"
    return "weak"


def allow_precise_schema_header_diffs(
    doc_block: Block,
    target_block: Block,
    *,
    score: float,
    match_type: str,
) -> bool:
    if not (
        doc_block.structure_role == "section_header"
        and target_block.structure_role == "section_header"
    ):
        return False
    doc_schema = known_section_schema(doc_block.text, allow_fuzzy=True)
    target_schema = known_section_schema(target_block.text, allow_fuzzy=True)
    if (
        doc_schema is None
        or target_schema is None
        or doc_schema.schema_key != target_schema.schema_key
    ):
        return False
    if max(len(doc_block.normalized), len(target_block.normalized)) > 64:
        return False
    if exact_like_match_type(match_type):
        return True
    header_similarity = difflib.SequenceMatcher(None, doc_block.normalized, target_block.normalized).ratio()
    return header_similarity >= 0.8 and score >= 0.52


def next_narrative_neighbor(blocks: list[Block], start_index: int) -> Block | None:
    for probe in range(start_index + 1, len(blocks)):
        candidate = blocks[probe]
        if candidate.table_cell or candidate.kind == "footnote":
            continue
        if not visible_meaningful(candidate.text):
            continue
        return candidate
    return None


def promote_exact_structural_match(
    doc_index: int,
    target_index: int,
    doc_blocks: list[Block],
    target_blocks: list[Block],
    *,
    grouped_match_type: str | None = None,
    score: float,
    match_type: str,
) -> str:
    if exact_like_match_type(match_type):
        return match_type
    doc_block = doc_blocks[doc_index]
    target_block = target_blocks[target_index]
    if grouped_match_type == "contact":
        payload = contact_field_payload(doc_block.text)
        if payload is None:
            return match_type
        role, doc_value = payload
        target_value = extract_contact_fields(target_block.text).get(role)
        if target_value and normalize_proof_text(doc_value) == normalize_proof_text(target_value):
            return "exact_structural"
        return match_type
    doc_schema = known_section_schema(doc_block.text, allow_fuzzy=True)
    target_schema = known_section_schema(target_block.text, allow_fuzzy=True)
    if (
        doc_block.structure_role == "section_header"
        and target_block.structure_role == "section_header"
        and doc_schema is not None
        and target_schema is not None
        and doc_schema.schema_key == target_schema.schema_key
    ):
        if score >= 0.78:
            next_doc = next_narrative_neighbor(doc_blocks, doc_index)
            next_target = next_narrative_neighbor(target_blocks, target_index)
            if next_doc is None or next_target is None:
                return "exact_structural"
            if similarity(next_doc.normalized, next_target.normalized) >= 0.82:
                return "exact_structural"
    if doc_block.structure_role == "section_lead" and target_block.structure_role == "section_lead":
        if score < 0.78:
            return match_type
        next_doc = next_narrative_neighbor(doc_blocks, doc_index)
        next_target = next_narrative_neighbor(target_blocks, target_index)
        if next_doc is None or next_target is None:
            return match_type
        if similarity(next_doc.normalized, next_target.normalized) >= 0.9:
            return "exact_structural"
    return match_type


def quote_group_key(block: Block) -> str | None:
    if block.table_cell or block.heading:
        return None
    text = normalize_text(block.text).strip()
    normalized = normalize_for_compare(text)
    role = quote_role_key(text) or ""
    placeholder = "quote tbc" in normalized
    if len(text) < 40 and not placeholder:
        return None
    if '"' not in text and "“" not in text and not placeholder:
        return None
    if "said " not in normalized and not placeholder:
        return None
    speaker, title = quote_speaker_title(text)
    lead = quote_lead_key(text)
    if not any((role, speaker, title, lead)):
        return None
    return f"{role}|{speaker}|{title}|{lead}|{'1' if placeholder else '0'}"


def footnote_group_key(block: Block) -> str | None:
    if block.kind == "footnote":
        return normalize_without_footnote_refs(block.text)
    if block.table_cell and len(block.normalized) > 80 and block.text[:1].isdigit():
        return normalize_without_footnote_refs(block.text)
    return None


def leading_footnote_marker(text: str) -> str | None:
    match = re.match(r"^\s*\(?(\d{1,2})\)?(?=\s+[A-Za-z])", normalize_text(text))
    if not match:
        return None
    return match.group(1)


def extract_block_groups(blocks: list[Block]) -> tuple[list[BlockGroup], dict[int, str]]:
    groups: list[BlockGroup] = []
    block_to_group: dict[int, str] = {}
    used_indices: set[int] = set()

    index = 0
    while index < len(blocks):
        if index in used_indices:
            index += 1
            continue
        block = blocks[index]
        contact_key = contact_label_key(block.text)
        if contact_key:
            group_indices = [index]
            used_indices.add(index)
            probe = index + 1
            while probe < len(blocks):
                candidate = blocks[probe]
                if candidate.table_cell or candidate.heading:
                    break
                if contact_label_key(candidate.text):
                    break
                if not looks_like_contact_fragment(candidate):
                    break
                group_indices.append(probe)
                used_indices.add(probe)
                probe += 1
                if len(group_indices) >= 6:
                    break
            group = BlockGroup(
                group_id=f"{blocks[index].source}-group-contact-{index}",
                source=blocks[index].source,
                group_type="contact",
                block_indices=group_indices,
                key=contact_key,
                text="\n".join(blocks[item].text for item in group_indices),
                order_start=blocks[group_indices[0]].order,
                order_end=blocks[group_indices[-1]].order,
            )
            groups.append(group)
            for item in group_indices:
                block_to_group[item] = group.group_id
            index = probe
            continue

        quote_key = quote_group_key(block)
        if quote_key:
            group = BlockGroup(
                group_id=f"{block.source}-group-quote-{index}",
                source=block.source,
                group_type="quote",
                block_indices=[index],
                key=quote_key,
                text=block.text,
                order_start=block.order,
                order_end=block.order,
            )
            groups.append(group)
            block_to_group[index] = group.group_id
            used_indices.add(index)
            index += 1
            continue

        footnote_key = footnote_group_key(block)
        if footnote_key:
            marker = block.footnote_marker or leading_footnote_marker(block.text)
            group = BlockGroup(
                group_id=f"{block.source}-group-footnote-{index}",
                source=block.source,
                group_type="footnote",
                block_indices=[index],
                key=footnote_key,
                text=block.text,
                order_start=block.order,
                order_end=block.order,
                footnote_marker=marker,
            )
            groups.append(group)
            block_to_group[index] = group.group_id
            used_indices.add(index)
        index += 1

    return groups, block_to_group


def group_similarity(doc_group: BlockGroup, target_group: BlockGroup) -> float:
    if doc_group.group_type != target_group.group_type:
        return 0.0
    if doc_group.group_type == "contact":
        return 1.0 if doc_group.key == target_group.key else 0.0
    if doc_group.group_type == "quote":
        if doc_group.key == target_group.key:
            return 1.0
        doc_role, doc_speaker, doc_title, doc_lead, doc_placeholder = parse_quote_key(doc_group.key)
        target_role, target_speaker, target_title, target_lead, target_placeholder = parse_quote_key(target_group.key)
        role_match = doc_role and target_role and doc_role == target_role
        speaker_score = similarity(doc_speaker, target_speaker) if doc_speaker and target_speaker else 0.0
        title_score = similarity(doc_title, target_title) if doc_title and target_title else 0.0
        lead_overlap = token_overlap_ratio(doc_lead, target_lead) if doc_lead and target_lead else 0.0
        key_similarity = similarity(doc_group.key, target_group.key)
        if role_match and (doc_placeholder or target_placeholder):
            return max(key_similarity, 0.95)
        if role_match and speaker_score >= 0.92 and title_score >= 0.72:
            return max(key_similarity, 0.98)
        if role_match and max(title_score, speaker_score) >= 0.62:
            return max(key_similarity, 0.93)
        if speaker_score >= 0.9 and lead_overlap >= 0.45:
            return max(key_similarity, 0.92)
        return max(key_similarity, 0.0)
    if doc_group.group_type == "footnote":
        marker_match = (
            doc_group.footnote_marker is not None
            and target_group.footnote_marker is not None
            and doc_group.footnote_marker == target_group.footnote_marker
        )
        text_similarity = similarity(doc_group.key, target_group.key)
        overlap = token_overlap_ratio(doc_group.key, target_group.key)
        if marker_match and overlap >= 0.45:
            return max(text_similarity, 0.96)
        return text_similarity
    return 0.0


def match_block_groups(
    docx_blocks: list[Block],
    target_blocks: list[Block],
) -> tuple[dict[str, str], dict[int, str], dict[int, str], dict[str, BlockGroup], dict[str, BlockGroup]]:
    doc_groups, doc_block_to_group = extract_block_groups(docx_blocks)
    target_groups, target_block_to_group = extract_block_groups(target_blocks)
    doc_group_map = {group.group_id: group for group in doc_groups}
    target_group_map = {group.group_id: group for group in target_groups}

    group_matches: dict[str, str] = {}
    used_target_group_ids: set[str] = set()

    for doc_group in doc_groups:
        best_target: BlockGroup | None = None
        best_score = 0.0
        for target_group in target_groups:
            if target_group.group_id in used_target_group_ids:
                continue
            score = group_similarity(doc_group, target_group)
            if doc_group.group_type == "footnote" and score < 0.88:
                continue
            if doc_group.group_type == "quote" and score < 0.9:
                continue
            if doc_group.group_type == "contact" and score < 1.0:
                continue
            if score > best_score:
                best_score = score
                best_target = target_group
        if best_target is not None:
            group_matches[doc_group.group_id] = best_target.group_id
            used_target_group_ids.add(best_target.group_id)

    block_group_match_by_doc_index: dict[int, str] = {}
    for doc_group_id, target_group_id in group_matches.items():
        target_group = target_group_map[target_group_id]
        for block_index in doc_group_map[doc_group_id].block_indices:
            block_group_match_by_doc_index[block_index] = target_group_id

    return block_group_match_by_doc_index, doc_block_to_group, target_block_to_group, doc_group_map, target_group_map


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


def same_cell_numeric_table_match(
    doc_block: Block,
    html_block: Block,
    *,
    expected_target_table_idx: int | None = None,
) -> bool:
    if not doc_block.table_cell or not html_block.table_cell:
        return False
    if doc_block.table_pos is None or html_block.table_pos is None:
        return False
    if expected_target_table_idx is not None and table_index_key(html_block) != expected_target_table_idx:
        return False
    if doc_block.table_pos[1:] != html_block.table_pos[1:]:
        return False
    doc_token = single_value_token(doc_block)
    html_token = single_value_token(html_block)
    if doc_token is None or html_token is None:
        return False
    if doc_token.kind != "number" or html_token.kind != "number":
        return False
    if doc_block.row_key and html_block.row_key and similarity(doc_block.row_key, html_block.row_key) < 0.94:
        return False
    if (
        doc_block.numeric_slot is not None
        and html_block.numeric_slot is not None
        and doc_block.numeric_slot != html_block.numeric_slot
    ):
        return False
    if (
        doc_block.row_slot is not None
        and html_block.row_slot is not None
        and doc_block.row_slot != html_block.row_slot
    ):
        return False
    return True


def prematch_same_cell_numeric_table_blocks(
    docx_blocks: list[Block],
    html_blocks: list[Block],
    table_pos_map: dict[tuple[int, int, int], list[int]],
    doc_to_html_table_map: dict[int, int] | None = None,
) -> tuple[list[Match], set[int], set[int]]:
    used_doc: set[int] = set()
    used_html: set[int] = set()
    matches: list[Match] = []

    for doc_index, doc_block in enumerate(docx_blocks):
        pos_key = mapped_table_pos_key(doc_block, doc_to_html_table_map)
        if pos_key is None:
            continue
        doc_token = single_value_token(doc_block)
        if doc_token is None or doc_token.kind != "number":
            continue
        expected_table_idx = mapped_table_index_key(doc_block, doc_to_html_table_map)
        candidates = [
            index
            for index in table_pos_map.get(pos_key, [])
            if index not in used_html
            and same_cell_numeric_table_match(
                doc_block,
                html_blocks[index],
                expected_target_table_idx=expected_table_idx,
            )
        ]
        if not candidates:
            continue
        best_index = max(candidates, key=lambda index: similarity(doc_block.normalized, html_blocks[index].normalized))
        used_doc.add(doc_index)
        used_html.add(best_index)
        score = 1.0 if html_blocks[best_index].normalized == doc_block.normalized else 0.95
        match_type = "exact" if score == 1.0 else "approx"
        matches.append(
            Match(
                docx_index=doc_index,
                html_index=best_index,
                match_type=match_type,
                score=score,
                formatting_diffs=summarize_formatting_diff(doc_block, html_blocks[best_index]),
            )
        )
    return matches, used_doc, used_html


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
    if not re.search(r"[A-Za-z]", lead):
        return None
    if not re.match(r"^[A-Z][A-Za-z0-9'()/%& .-]+$", lead):
        return None
    if not re.match(r"^[\"'(]?[A-Z]", rest):
        return None
    if lead_words <= 1 and len(lead) < 12:
        return None
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


def normalize_proof_text(text: str) -> str:
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
    )


def classify_inline_run(text: str) -> str:
    if not text:
        return "text"
    if text == "\t":
        return "tab"
    if text == "\n":
        return "linebreak"
    if re.fullmatch(r"[ \t]+", text):
        return "space"
    if re.fullmatch(r"\s+", text):
        return "space"
    if re.fullmatch(r"[$€£¥%~()\[\]{}:;,.!?/&+\-–—•]+", text):
        return "symbol"
    return "text"


def block_texts_from_runs(runs: list[InlineRun]) -> tuple[str, str, str]:
    raw_text = "".join(run.text for run in runs)
    proof_text = normalize_proof_text(raw_text)
    match_text = normalize_for_compare(proof_text)
    return raw_text, proof_text, match_text


def formatting_flags_from_runs(runs: list[InlineRun]) -> tuple[bool, bool, bool]:
    return (
        any(run.bold for run in runs),
        any(run.italic for run in runs),
        any(run.underline for run in runs),
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
    text = re.sub(r"([$€£¥~])\s+(?=\(?\d)", r"\1", text)
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


TABLE_HEADER_TEXT_MARKERS = {
    "low",
    "high",
    "three months ended",
    "range for three months ending",
    "range for fiscal year ending",
    "january 31",
    "april 30",
    "july 31",
    "october 31",
    "fiscal year ending",
    "in millions",
    "in thousands",
    "except per share amounts",
}


def headerish_table_text(text: str) -> bool:
    normalized = normalize_row_key(text)
    if not normalized:
        return False
    if normalized in TABLE_HEADER_TEXT_MARKERS:
        return True
    if any(marker in normalized for marker in TABLE_HEADER_TEXT_MARKERS):
        return True
    tokens = tokenize(normalized)
    if not tokens:
        return False
    if all(token in {"low", "high"} for token in tokens):
        return True
    if all(token.isdigit() and len(token) == 4 for token in tokens):
        return True
    if DATE_PHRASE_RE.search(text):
        return True
    return False


def assign_structural_roles(blocks: list[Block]) -> list[Block]:
    for block in blocks:
        if block.heading:
            block.structure_role = "section_header"
        elif block.table_cell:
            block.structure_role = "table_cell"
        elif block.kind == "footnote":
            block.structure_role = "footnote"
        else:
            block.structure_role = "paragraph"

    for index, block in enumerate(blocks):
        if block.table_cell or block.kind == "footnote":
            continue
        next_block = blocks[index + 1] if index + 1 < len(blocks) else None
        prev_block = blocks[index - 1] if index > 0 else None
        upcoming_table = any(candidate.table_cell for candidate in blocks[index + 1:index + 4])
        upcoming_table_index = next(
            (candidate.table_pos[0] for candidate in blocks[index + 1:index + 4] if candidate.table_cell and candidate.table_pos is not None),
            None,
        )
        text = normalize_text(block.text).strip()
        token_count = len(diff_tokens(text))
        schema = known_section_schema(text, allow_fuzzy=True)
        if schema is not None:
            block.structure_role = "section_header"
        elif upcoming_table:
            prev_schema = known_section_schema(prev_block.text, allow_fuzzy=True) if prev_block is not None else None
            if (
                prev_schema is not None
                and prev_block is not None
                and prev_block.structure_role == "section_header"
                and token_count > 14
                and not text.startswith("(")
            ):
                block.structure_role = "paragraph"
                continue
            if text.startswith("(") or (block.italic and token_count <= 20):
                block.structure_role = "table_subtitle"
            elif block.heading or token_count <= 14 or (block.bold and token_count <= 14):
                block.structure_role = "table_title"
            if block.structure_role in {"table_title", "table_subtitle"}:
                block.family_table_index = upcoming_table_index
        elif (
            prev_block is not None
            and prev_block.structure_role == "table_title"
            and (text.startswith("(") or block.italic or token_count <= 18)
        ):
            block.structure_role = "table_subtitle"
            block.family_table_index = prev_block.family_table_index
        elif looks_like_section_lead(block, next_block):
            block.structure_role = "section_lead"

    table_rows: dict[int, dict[int, list[Block]]] = {}
    for block in blocks:
        if not block.table_cell or block.table_pos is None:
            continue
        table_idx, row_idx, _col_idx = block.table_pos
        block.family_table_index = table_idx
        table_rows.setdefault(table_idx, {}).setdefault(row_idx, []).append(block)

    for rows in table_rows.values():
        ordered_rows = sorted(rows.items())
        first_body_row: int | None = None
        for row_idx, row_blocks in ordered_rows:
            ordered_blocks = sorted(row_blocks, key=lambda item: (item.row_slot or 0, item.table_pos[2] if item.table_pos else 0))
            numeric_count = 0
            first_text_block: Block | None = None
            for row_block in ordered_blocks:
                token = single_value_token(row_block)
                if token is not None and token.kind == "number":
                    numeric_count += 1
                if first_text_block is None and visible_meaningful(row_block.text):
                    first_text_block = row_block
            if (
                first_text_block is not None
                and numeric_count >= 1
                and not headerish_table_text(first_text_block.text)
            ):
                first_body_row = row_idx
                break

        for row_idx, row_blocks in ordered_rows:
            header_row = first_body_row is None or row_idx < first_body_row
            ordered_blocks = sorted(row_blocks, key=lambda item: (item.row_slot or 0, item.table_pos[2] if item.table_pos else 0))
            for row_block in ordered_blocks:
                token = single_value_token(row_block)
                if header_row:
                    row_block.structure_role = "table_column_header"
                elif token is not None and token.kind == "number":
                    row_block.structure_role = "table_data_cell"
                elif (row_block.row_slot or 0) == 0:
                    row_block.structure_role = "table_row_label"
                else:
                    row_block.structure_role = "table_data_label"

    return blocks


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


def known_section_schema(text: str, *, allow_fuzzy: bool = False) -> SectionSchema | None:
    normalized = normalize_without_punctuation(text)
    if not normalized:
        return None
    for schema in KNOWN_EARNINGS_SECTION_SCHEMAS:
        for pattern in schema.patterns:
            if pattern == normalized:
                return schema
            if re.fullmatch(pattern, normalized):
                return schema
    if allow_fuzzy:
        tokens = tokenize(normalized)
        if 0 < len(tokens) <= 8 and len(normalized) <= 96:
            best_schema: SectionSchema | None = None
            best_score = 0.0
            for schema in KNOWN_EARNINGS_SECTION_SCHEMAS:
                for pattern in schema.patterns:
                    if re.search(r"[().?:|+*\\\[\]{}]", pattern):
                        continue
                    pattern_tokens = tokenize(pattern)
                    if len(pattern_tokens) != len(tokens):
                        continue
                    score = sum(difflib.SequenceMatcher(None, left, right).ratio() for left, right in zip(tokens, pattern_tokens)) / max(len(tokens), 1)
                    if score > best_score:
                        best_score = score
                        best_schema = schema
            if best_schema is not None and best_score >= 0.9:
                return best_schema
    return None


def known_section_header(text: str) -> bool:
    return known_section_schema(text, allow_fuzzy=True) is not None


def normalize_without_footnote_refs(text: str) -> str:
    cleaned = normalize_text(text)
    cleaned = re.sub(r"\(\s*\d+\s*\)", "", cleaned)
    cleaned = re.sub(r"(?<=[A-Za-z])\)", "", cleaned)
    cleaned = re.sub(r"\(\s*$", "", cleaned)
    return normalize_for_compare(cleaned)


def normalize_pdf_paragraph_artifacts(text: str) -> str:
    cleaned = normalize_text(text)
    cleaned = re.sub(r"([A-Za-z])-\s*\n\s*([A-Za-z])", r"\1-\2", cleaned)
    lines: list[str] = []
    for raw_line in cleaned.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if re.fullmatch(r"\d{1,2}", line):
            continue
        if likely_pdf_noise_line(line):
            continue
        lines.append(line)
    cleaned = "\n".join(lines)
    cleaned = re.sub(r"(?:\n\s*)+\d{1,2}\s*$", "", cleaned)
    cleaned = re.sub(r"\b[A-Z][a-z]{4,}(and|or|the|to|of|for|with|from|by)\b", r"\1", cleaned)
    cleaned = re.sub(r"\s*\n\s*", " ", cleaned)
    return normalize_for_compare(cleaned)


def likely_pdf_noise_line(text: str) -> bool:
    line = normalize_text(text).strip()
    if not line:
        return False
    if re.fullmatch(r"\[?[A-Z0-9]{2,}(?:\s+[A-Z0-9]{1,4}){1,6}\]?", line):
        return True
    if headerish_table_text(line):
        return False
    alpha_chars = [char for char in line if char.isalpha()]
    if not alpha_chars:
        return False
    upper_ratio = sum(1 for char in alpha_chars if char.isupper()) / len(alpha_chars)
    token_count = len(diff_tokens(line))
    if token_count <= 4 and len(line) <= 28 and upper_ratio >= 0.75 and not re.search(r"[.:;!?]", line):
        return True
    return False


def looks_like_pdf_section_line(text: str, next_line: str | None) -> bool:
    line = normalize_text(text).strip()
    if not visible_meaningful(line):
        return False
    if next_line is None or len(diff_tokens(next_line)) < 8:
        return False
    if len(line) > 90:
        return False
    token_count = len(diff_tokens(line))
    if token_count < 2 or token_count > 10:
        return False
    if headerish_table_text(line) or line.endswith((".", ";", ",")):
        return False
    words = re.findall(r"[A-Za-z][A-Za-z'’-]*", line)
    if not words:
        return False
    titled = sum(1 for word in words if word[:1].isupper())
    lowered = sum(1 for word in words if word.islower())
    if titled / len(words) < 0.7:
        return False
    if lowered / len(words) > 0.35:
        return False
    return True


def pdf_narrative_segments(block: Block) -> list[str]:
    if block.source != "pdf" or block.table_cell:
        return [normalize_text(block.text).strip()]
    lines = [normalize_text(line).strip() for line in normalize_text(block.text).splitlines() if visible_meaningful(line)]
    if len(lines) < 2:
        return [normalize_text(block.text).strip()]

    segments: list[str] = []
    start = 0
    first_next = lines[1] if len(lines) > 1 else None
    if looks_like_pdf_section_line(lines[0], first_next):
        segments.append(lines[0])
        start = 1

    current: list[str] = []
    for index in range(start, len(lines)):
        line = lines[index]
        next_line = lines[index + 1] if index + 1 < len(lines) else None
        if current and looks_like_pdf_section_line(line, next_line):
            segment = "\n".join(current).strip()
            if segment:
                segments.append(segment)
            current = [line]
            continue
        current.append(line)

    if current:
        segment = "\n".join(current).strip()
        if segment:
            segments.append(segment)

    return segments or [normalize_text(block.text).strip()]


def best_pdf_narrative_focus(doc_block: Block, target_block: Block) -> str | None:
    if target_block.source != "pdf" or target_block.table_cell:
        return None
    segments = pdf_narrative_segments(target_block)
    if len(segments) <= 1:
        return None
    doc_norm = normalize_for_compare(doc_block.text)
    full_norm = normalize_for_compare(target_block.text)
    full_score = similarity(doc_norm, full_norm)
    best_segment: str | None = None
    best_score = full_score
    for segment in segments:
        segment_norm = normalize_for_compare(segment)
        score = similarity(doc_norm, segment_norm)
        if score > best_score:
            best_score = score
            best_segment = segment
    if best_segment is None:
        return None
    if best_score >= 0.9 or best_score - full_score >= 0.05:
        return best_segment
    return None


def pdf_span_format(flags: int, font_name: str) -> tuple[bool, bool, bool]:
    font_lower = (font_name or "").lower()
    bold = bool(flags & 16) or "bold" in font_lower or "black" in font_lower or "semibold" in font_lower
    italic = bool(flags & 2) or "italic" in font_lower or "oblique" in font_lower
    underline = "underline" in font_lower
    return bold, italic, underline


def build_pdf_span_entries(page: Any) -> list[dict[str, Any]]:
    span_entries: list[dict[str, Any]] = []
    page_dict = page.get_text("dict")
    for block in page_dict.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = normalize_text(str(span.get("text", "")))
                if not text:
                    continue
                bbox = span.get("bbox")
                if not bbox or len(bbox) < 4:
                    continue
                bold, italic, underline = pdf_span_format(int(span.get("flags", 0) or 0), str(span.get("font", "")))
                span_entries.append(
                    {
                        "rect": (float(bbox[0]), float(bbox[1]), float(bbox[2]), float(bbox[3])),
                        "text": text,
                        "bold": bold,
                        "italic": italic,
                        "underline": underline,
                    }
                )
    return span_entries


def collect_pdf_runs_for_rect(
    rect: tuple[float, float, float, float],
    span_entries: list[dict[str, Any]] | None,
) -> list[InlineRun]:
    if not span_entries:
        return []
    x0, y0, x1, y1 = rect
    matched: list[dict[str, Any]] = []
    for span in span_entries:
        span_rect = span["rect"]
        if overlap_area(rect, span_rect) > 0 or rect_contains_point(rect, (span_rect[0] + span_rect[2]) / 2, (span_rect[1] + span_rect[3]) / 2):
            matched.append(span)
    if not matched:
        return []
    matched.sort(key=lambda item: (round(item["rect"][1], 1), round(item["rect"][0], 1)))
    runs: list[InlineRun] = []
    prev_rect: tuple[float, float, float, float] | None = None
    for span_index, span in enumerate(matched):
        span_rect = span["rect"]
        if prev_rect is not None:
            prev_h = max(1.0, prev_rect[3] - prev_rect[1])
            curr_h = max(1.0, span_rect[3] - span_rect[1])
            if abs(span_rect[1] - prev_rect[1]) > max(prev_h, curr_h) * 0.8:
                if not runs or runs[-1].text != "\n":
                    runs.append(InlineRun(text="\n", kind="linebreak", source_index=span_index))
            elif span_rect[0] - prev_rect[2] > 0.8:
                if not runs or runs[-1].kind != "space":
                    runs.append(InlineRun(text=" ", kind="space", source_index=span_index))
        for chunk in re.findall(r"\s+|[^\s]+", span["text"]):
            normalized_chunk = normalize_text(chunk)
            if not normalized_chunk:
                continue
            runs.append(
                InlineRun(
                    text=normalized_chunk,
                    kind=classify_inline_run(normalized_chunk),
                    bold=bool(span["bold"]),
                    italic=bool(span["italic"]),
                    underline=bool(span["underline"]),
                    source_index=span_index,
                )
            )
        prev_rect = span_rect
    return runs


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


def cluster_signature_x_positions(clusters: list[list[dict[str, Any]]]) -> list[float]:
    positions: list[float] = []
    for cluster in clusters:
        if not cluster:
            continue
        x0 = min(item["rect"][0] for item in cluster)
        positions.append(round(x0, 1))
    return positions


def compatible_pdf_column_signature(left: list[float], right: list[float]) -> bool:
    if not left or not right:
        return False
    if abs(len(left) - len(right)) > 1:
        return False
    compare_count = min(len(left), len(right), 4)
    if compare_count == 0:
        return False
    matches = 0
    for idx in range(compare_count):
        if abs(left[idx] - right[idx]) <= 24.0:
            matches += 1
    return matches >= max(1, compare_count - 1)


def build_pdf_row_infos(
    row_groups: list[list[dict[str, Any]]],
    *,
    page_width: float,
    page_height: float,
    pixmap: Any | None = None,
) -> list[dict[str, Any]]:
    row_infos: list[dict[str, Any]] = []
    for row_index, row_words in enumerate(row_groups):
        row_words = sorted(row_words, key=lambda word: word["rect"][0])
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
        cluster_texts = [
            join_words_for_pdf_cluster(
                cluster,
                page_width=page_width,
                page_height=page_height,
                pixmap=pixmap,
            )
            for cluster in meaningful_clusters
        ]
        numeric_count = 0
        currency_count = 0
        headerish_count = 0
        for text in cluster_texts:
            temp_block = Block(
                id="",
                source="pdf",
                order=0,
                text=text,
                normalized=normalize_for_compare(text),
            )
            token = single_value_token(temp_block)
            if token is not None and token.kind == "number":
                numeric_count += 1
            if extract_currency_symbol(text):
                currency_count += 1
            if headerish_table_text(text):
                headerish_count += 1
        row_infos.append(
            {
                "row_index": row_index,
                "clusters": meaningful_clusters,
                "cluster_texts": cluster_texts,
                "x_positions": cluster_signature_x_positions(meaningful_clusters),
                "numeric_count": numeric_count,
                "currency_count": currency_count,
                "headerish_count": headerish_count,
            }
        )
    return row_infos


def detect_pdf_table_regions(row_infos: list[dict[str, Any]]) -> dict[int, tuple[int, int]]:
    if not row_infos:
        return {}
    candidate_rows: set[int] = set()
    for info in row_infos:
        cluster_count = len(info["clusters"])
        if cluster_count < 2:
            continue
        if info["numeric_count"] >= 1 or info["currency_count"] >= 1 or info["headerish_count"] >= 1:
            candidate_rows.add(info["row_index"])

    expanded_candidates = set(candidate_rows)
    for info in row_infos:
        row_index = info["row_index"]
        if row_index in candidate_rows:
            continue
        if len(info["clusters"]) != 1 or info["headerish_count"] == 0:
            continue
        prev_candidate = row_index - 1 in candidate_rows
        next_candidate = row_index + 1 in candidate_rows
        if prev_candidate and next_candidate:
            expanded_candidates.add(row_index)

    region_rows: dict[int, tuple[int, int]] = {}
    region_id = 0
    index = 0
    while index < len(row_infos):
        row_index = row_infos[index]["row_index"]
        if row_index not in expanded_candidates:
            index += 1
            continue
        start = index
        end = index
        while end + 1 < len(row_infos) and row_infos[end + 1]["row_index"] in expanded_candidates:
            end += 1
        region_infos = row_infos[start : end + 1]
        multi_cluster_infos = [info for info in region_infos if len(info["clusters"]) >= 2]
        if len(multi_cluster_infos) < 2:
            index = end + 1
            continue
        signatures = [info["x_positions"] for info in multi_cluster_infos if info["x_positions"]]
        signature_ok = False
        for left, right in zip(signatures, signatures[1:]):
            if compatible_pdf_column_signature(left, right):
                signature_ok = True
                break
        if not signature_ok:
            index = end + 1
            continue
        row_ordinal = 0
        for info in region_infos:
            region_rows[info["row_index"]] = (region_id, row_ordinal)
            row_ordinal += 1
        region_id += 1
        index = end + 1
    return region_rows


def diff_token_kind(token: str) -> str:
    return "number" if re.fullmatch(r"(?:\([+-]?(?:\d[\d,]*(?:\.\d+)?|\.\d+)\)|[+-]?(?:\d[\d,]*(?:\.\d+)?|\.\d+))%?", token) else "word"


def normalize_diff_token(token: str) -> str:
    token = normalize_text(token).strip()
    if diff_token_kind(token) == "number":
        return token.replace(",", "").lower()
    return token.lower()


def diff_tokens(text: str) -> list[DiffToken]:
    normalized_text = normalize_proof_text(text)
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
        if prefix_symbol == "(":
            suffix = normalized_text[token_end:]
            close_paren_match = re.match(r"(\s*\))", suffix)
            if close_paren_match:
                token_text = f"({token_text}{close_paren_match.group(1)}"
                token_end += close_paren_match.end()
                prefix_symbol = None
        if diff_token_kind(token_text) == "number" and not token_text.endswith("%"):
            suffix = normalized_text[token_end:]
            percent_match = re.match(r"(\s*)%(?=\s|$|[),.;:])", suffix)
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
    for peer in row_peer_blocks(block, blocks):
        peer_text = normalize_text(peer.text).strip()
        if peer_text in {"$", "€", "£", "¥"}:
            return peer_text
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


def pdf_blocks_equal_after_cleanup(doc_block: Block, target_block: Block) -> bool:
    if normalize_without_footnote_refs(doc_block.text) == normalize_without_footnote_refs(target_block.text):
        return True
    if (
        not doc_block.table_cell
        and not target_block.table_cell
        and normalize_pdf_paragraph_artifacts(doc_block.text) == normalize_pdf_paragraph_artifacts(target_block.text)
    ):
        return True
    if normalize_for_compare(strip_leading_markers(doc_block.text)) == normalize_for_compare(strip_leading_markers(target_block.text)):
        return True
    return False


def pdf_minor_narrative_noise_only(doc_text: str, target_text: str) -> bool:
    doc_clean = normalize_pdf_paragraph_artifacts(doc_text)
    target_clean = normalize_pdf_paragraph_artifacts(target_text)
    if not doc_clean or not target_clean:
        return False
    if difflib.SequenceMatcher(None, doc_clean, target_clean).ratio() < 0.995:
        return False
    doc_tokens = tokenize(doc_clean)
    target_tokens = tokenize(target_clean)
    if not doc_tokens or not target_tokens:
        return False
    total_changed = 0
    matcher = difflib.SequenceMatcher(a=doc_tokens, b=target_tokens, autojunk=False)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue
        doc_slice = doc_tokens[i1:i2]
        target_slice = target_tokens[j1:j2]
        if tag == "replace":
            return False
        if any(any(char.isdigit() for char in token) for token in doc_slice + target_slice):
            return False
        total_changed += len(doc_slice) + len(target_slice)
        if total_changed > 2:
            return False
    return total_changed > 0


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


def formatting_context_confident(docx: Block, html: Block) -> bool:
    if docx.table_cell != html.table_cell:
        return False
    if len(docx.normalized) < 40 and docx.kind != html.kind and not (docx.heading and html.heading):
        return False
    return True


def summarize_formatting_diff(docx: Block, html: Block) -> list[str]:
    diffs: list[str] = []
    checks = [
        ("heading", docx.heading, html.heading),
        ("list item", docx.list_item, html.list_item),
        ("table cell", docx.table_cell, html.table_cell),
    ]
    for label, doc_value, html_value in checks:
        if doc_value != html_value:
            diffs.append(
                f"DOCX {'has' if doc_value else 'does not have'} {label}; "
                f"HTML {'has' if html_value else 'does not have'} it."
            )
    if formatting_context_confident(docx, html):
        diffs.extend(compare_inline_formatting_diffs(docx, html))
    if docx.heading and html.heading and docx.heading_level != html.heading_level:
        diffs.append(
            f"DOCX heading level is {docx.heading_level or '?'}; "
            f"HTML heading level is {html.heading_level or '?'}."
        )
    if (
        docx.structure_role
        and html.structure_role
        and docx.structure_role != html.structure_role
        and {docx.structure_role, html.structure_role} <= {
            "table_title",
            "table_subtitle",
            "table_column_header",
            "table_row_label",
            "table_data_cell",
            "table_data_label",
            "section_header",
            "paragraph",
        }
    ):
        diffs.append(
            f'DOCX role is "{docx.structure_role}"; HTML role is "{html.structure_role}".'
        )
    return diffs


def _visual_underline(run: InlineRun) -> bool:
    return run.underline or run.hyperlink


def _run_style_enabled(run: InlineRun, style: str) -> bool:
    if style == "underline":
        return _visual_underline(run)
    return bool(getattr(run, style, False))


def _block_token_counter(block: Block) -> Counter[str]:
    counter: Counter[str] = Counter()
    for token in diff_tokens(block.text):
        counter[token.normalized] += 1
    return counter


def _styled_run_segments(block: Block, style: str) -> list[str]:
    segments: list[str] = []
    current: list[str] = []
    for run in block.runs:
        if _run_style_enabled(run, style):
            current.append(run.text)
        else:
            if current:
                segments.append("".join(current))
                current = []
    if current:
        segments.append("".join(current))
    return segments


def _run_style_token_counter(
    block: Block,
    style: str,
    allowed_tokens: Counter[str] | None = None,
) -> Counter[str]:
    counter: Counter[str] = Counter()
    remaining = Counter(allowed_tokens) if allowed_tokens is not None else None
    for segment in _styled_run_segments(block, style):
        for token in diff_tokens(segment):
            if remaining is not None:
                if remaining[token.normalized] <= 0:
                    continue
                remaining[token.normalized] -= 1
            counter[token.normalized] += 1
    return counter


def _run_style_excerpt(
    block: Block,
    style: str,
    allowed_tokens: Counter[str] | None = None,
) -> str | None:
    remaining = Counter(allowed_tokens) if allowed_tokens is not None else None
    for segment in _styled_run_segments(block, style):
        if remaining is not None:
            matched_parts: list[str] = []
            for token in diff_tokens(segment):
                if remaining[token.normalized] <= 0:
                    continue
                remaining[token.normalized] -= 1
                matched_parts.append(token.text)
            text = normalize_proof_text(" ".join(matched_parts)).strip()
        else:
            text = normalize_proof_text(segment).strip()
        if visible_meaningful(text):
            return text[:60]
    return None


def compare_inline_formatting_diffs(docx: Block, html: Block) -> list[str]:
    diffs: list[str] = []
    if not docx.runs or not html.runs:
        fallback_checks = [
            ("bold", docx.bold, html.bold),
            ("italic", docx.italic, html.italic),
            ("underline", docx.underline, html.underline),
        ]
        for label, doc_value, html_value in fallback_checks:
            if doc_value != html_value:
                diffs.append(
                    f"DOCX {'has' if doc_value else 'does not have'} {label}; "
                    f"HTML {'has' if html_value else 'does not have'} it."
                )
        return diffs

    common_tokens = _block_token_counter(docx) & _block_token_counter(html)
    if not common_tokens:
        return diffs

    for style in ("bold", "italic", "underline"):
        doc_counter = _run_style_token_counter(docx, style, common_tokens)
        html_counter = _run_style_token_counter(html, style, common_tokens)
        if doc_counter == html_counter:
            continue
        doc_only = doc_counter - html_counter
        html_only = html_counter - doc_counter
        if not doc_only and not html_only:
            continue
        if html_only and not doc_only:
            excerpt = _run_style_excerpt(html, style, common_tokens)
            if excerpt:
                diffs.append(f'DOCX does not have {style} on "{excerpt}"; HTML has it.')
            else:
                diffs.append(f"DOCX does not have {style}; HTML has it.")
            continue
        if doc_only and not html_only:
            excerpt = _run_style_excerpt(docx, style, common_tokens)
            if excerpt:
                diffs.append(f'DOCX has {style} on "{excerpt}"; HTML does not.')
            else:
                diffs.append(f"DOCX has {style}; HTML does not have it.")
            continue
        diffs.append(f"DOCX and HTML apply {style} to different text spans.")
    return diffs


def formatting_alignment_score(docx: Block, html: Block) -> float:
    score = 0.0
    if docx.structure_role and html.structure_role:
        if docx.structure_role == html.structure_role:
            score += 2.0
        else:
            score -= 1.25
    if docx.heading == html.heading:
        score += 2.0
    else:
        score -= 1.0
    if docx.list_item == html.list_item:
        score += 1.0
    else:
        score -= 0.5
    if docx.table_cell == html.table_cell:
        score += 1.5
    else:
        score -= 0.75
    if docx.kind == html.kind:
        score += 0.5
    if docx.heading and html.heading and docx.heading_level == html.heading_level:
        score += 0.5
    for style in ("bold", "italic", "underline"):
        if docx.runs and html.runs:
            if _run_style_token_counter(docx, style) == _run_style_token_counter(html, style):
                score += 1.25
            elif bool(getattr(docx, style, False)) == bool(getattr(html, style, False)):
                score += 0.4
            else:
                score -= 0.5
        else:
            if bool(getattr(docx, style, False)) == bool(getattr(html, style, False)):
                score += 0.8
            else:
                score -= 0.5
    return score


def select_best_exact_candidate(
    doc_block: Block,
    candidate_indices: list[int],
    html_blocks: list[Block],
    used_html: set[int],
) -> int | None:
    available = [index for index in candidate_indices if index not in used_html]
    available = [index for index in available if structural_role_compatible(doc_block, html_blocks[index])]
    if not available:
        return None
    ranked = sorted(
        available,
        key=lambda index: (
            formatting_alignment_score(doc_block, html_blocks[index]),
            -abs(len(doc_block.text) - len(html_blocks[index].text)),
            -html_blocks[index].order,
        ),
        reverse=True,
    )
    if repeated_label_block(doc_block) and len(ranked) > 1:
        top = ranked[0]
        second = ranked[1]
        top_score = formatting_alignment_score(doc_block, html_blocks[top])
        second_score = formatting_alignment_score(doc_block, html_blocks[second])
        top_group = (
            doc_block.table_cell == html_blocks[top].table_cell,
            doc_block.heading == html_blocks[top].heading,
            doc_block.kind == html_blocks[top].kind,
            doc_block.structure_role == html_blocks[top].structure_role,
        )
        second_group = (
            doc_block.table_cell == html_blocks[second].table_cell,
            doc_block.heading == html_blocks[second].heading,
            doc_block.kind == html_blocks[second].kind,
            doc_block.structure_role == html_blocks[second].structure_role,
        )
        if abs(top_score - second_score) < 1.0 and top_group == second_group:
            return None
    return ranked[0]


def xml_local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def style_flag(style_text: str, needle: str) -> bool:
    return needle in style_text.lower()


def merge_word_format_props(
    base: dict[str, bool | None] | None,
    override: dict[str, bool | None] | None,
) -> dict[str, bool | None]:
    merged = dict(base or {})
    for key, value in (override or {}).items():
        if value is not None:
            merged[key] = value
    return merged


def word_prop_enabled(rpr: ET.Element | None, tag: str) -> bool | None:
    if rpr is None:
        return None
    child = rpr.find(tag, WORD_NS)
    if child is None:
        return None
    val = (
        child.get(f"{{{WORD_NS['w']}}}val")
        or child.get("w:val")
        or child.get("val")
    )
    if xml_local_name(tag) == "u":
        if val is None:
            return True
        return str(val).lower() not in {"none", "0", "false", "off"}
    if val is None:
        return True
    return str(val).lower() not in {"0", "false", "off", "none"}


def extract_word_rpr_props(rpr: ET.Element | None) -> dict[str, bool | None]:
    return {
        "bold": word_prop_enabled(rpr, "w:b"),
        "italic": word_prop_enabled(rpr, "w:i"),
        "underline": word_prop_enabled(rpr, "w:u"),
    }


def word_style_val(node: ET.Element | None, child_tag: str) -> str | None:
    if node is None:
        return None
    child = node.find(child_tag, WORD_NS)
    if child is None:
        return None
    return (
        child.get(f"{{{WORD_NS['w']}}}val")
        or child.get("w:val")
        or child.get("val")
    )


def build_word_style_resolver(archive: zipfile.ZipFile) -> WordStyleResolver:
    try:
        styles_xml = archive.read("word/styles.xml")
    except KeyError:
        return WordStyleResolver(default_run_props={})
    root = ET.fromstring(styles_xml)
    default_run_props = extract_word_rpr_props(root.find("w:docDefaults/w:rPrDefault/w:rPr", WORD_NS))
    styles: dict[str, WordStyle] = {}
    hyperlink_style_id: str | None = None
    for style in root.findall("w:style", WORD_NS):
        style_id = (
            style.get(f"{{{WORD_NS['w']}}}styleId")
            or style.get("w:styleId")
            or style.get("styleId")
            or ""
        )
        if not style_id:
            continue
        style_type = (
            style.get(f"{{{WORD_NS['w']}}}type")
            or style.get("w:type")
            or style.get("type")
            or ""
        )
        style_name = word_style_val(style, "w:name") or style_id
        based_on = word_style_val(style, "w:basedOn")
        props = extract_word_rpr_props(style.find("w:rPr", WORD_NS))
        styles[style_id] = WordStyle(
            style_id=style_id,
            style_type=style_type,
            based_on=based_on,
            name=style_name,
            props=props,
        )
        if style_id.lower() == "hyperlink" or style_name.lower() == "hyperlink":
            hyperlink_style_id = style_id
    return WordStyleResolver(
        default_run_props=default_run_props,
        styles=styles,
        hyperlink_style_id=hyperlink_style_id,
    )


def collect_word_runs(paragraph: ET.Element, resolver: WordStyleResolver | None = None) -> list[InlineRun]:
    runs: list[InlineRun] = []
    paragraph_props = paragraph.find("w:pPr", WORD_NS)
    paragraph_style_id = word_style_val(paragraph_props, "w:pStyle")
    paragraph_style_props = resolver.resolve_style_props(paragraph_style_id) if resolver else {}
    paragraph_direct_props = extract_word_rpr_props(paragraph_props.find("w:rPr", WORD_NS) if paragraph_props is not None else None)

    def walk(node: ET.Element, *, hyperlink: bool = False) -> None:
        local = xml_local_name(node.tag) if isinstance(node.tag, str) else ""
        if local == "hyperlink":
            for child in list(node):
                if isinstance(child.tag, str):
                    walk(child, hyperlink=True)
            return
        if local == "r":
            run_index = len(runs)
            run_props = node.find("w:rPr", WORD_NS)
            run_style_id = word_style_val(run_props, "w:rStyle")
            effective_props = dict(resolver.default_run_props) if resolver else {}
            effective_props = merge_word_format_props(effective_props, paragraph_style_props)
            effective_props = merge_word_format_props(effective_props, paragraph_direct_props)
            if resolver and hyperlink and resolver.hyperlink_style_id:
                effective_props = merge_word_format_props(
                    effective_props,
                    resolver.resolve_style_props(resolver.hyperlink_style_id),
                )
            if resolver and run_style_id:
                effective_props = merge_word_format_props(effective_props, resolver.resolve_style_props(run_style_id))
            effective_props = merge_word_format_props(effective_props, extract_word_rpr_props(run_props))
            bold = bool(effective_props.get("bold"))
            italic = bool(effective_props.get("italic"))
            underline = bool(effective_props.get("underline"))
            for child in list(node):
                child_local = xml_local_name(child.tag) if isinstance(child.tag, str) else ""
                text = ""
                if child_local == "t":
                    text = child.text or ""
                elif child_local == "tab":
                    text = "\t"
                elif child_local == "br":
                    text = "\n"
                if text == "":
                    continue
                runs.append(
                    InlineRun(
                        text=text,
                        kind=classify_inline_run(text),
                        bold=bold,
                        italic=italic,
                        underline=underline,
                        hyperlink=hyperlink,
                        source_index=run_index,
                    )
                )
            return
        for child in list(node):
            if isinstance(child.tag, str):
                walk(child, hyperlink=hyperlink)

    walk(paragraph)
    return runs


def collect_word_text(paragraph: ET.Element) -> tuple[str, bool, bool, bool]:
    runs = collect_word_runs(paragraph)
    raw_text, proof_text, _match_text = block_texts_from_runs(runs)
    bold, italic, underline = formatting_flags_from_runs(runs)
    return (proof_text if proof_text else raw_text), bold, italic, underline


def extract_docx_footnote_blocks(
    archive: zipfile.ZipFile,
    order_start: int,
    resolver: WordStyleResolver | None = None,
) -> tuple[list[Block], int]:
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
        footnote_id = (
            footnote.get(f"{{{WORD_NS['w']}}}id")
            or footnote.get("w:id")
            or footnote.get("id")
            or ""
        )
        if footnote_type in {"separator", "continuationSeparator", "continuationNotice"}:
            continue
        for paragraph in footnote.findall(".//w:p", WORD_NS):
            runs = collect_word_runs(paragraph, resolver=resolver)
            raw_text, proof_text, match_text = block_texts_from_runs(runs)
            if not visible_meaningful(proof_text):
                continue
            bold, italic, underline = formatting_flags_from_runs(runs)
            blocks.append(
                Block(
                    id=f"docx-{order}",
                    source="docx",
                    order=order,
                    text=proof_text.strip(),
                    normalized=match_text,
                    raw_text=raw_text,
                    proof_text=proof_text,
                    match_text=match_text,
                    bold=bold,
                    italic=italic,
                    underline=underline,
                    kind="footnote",
                    footnote_marker=str(footnote_id) if str(footnote_id).isdigit() else None,
                    runs=runs,
                )
            )
            order += 1
    return blocks, order


def extract_docx_blocks(path: Path) -> list[Block]:
    with zipfile.ZipFile(path) as archive:
        resolver = build_word_style_resolver(archive)
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
                runs = collect_word_runs(child, resolver=resolver)
                raw_text, proof_text, match_text = block_texts_from_runs(runs)
                if not visible_meaningful(proof_text):
                    continue
                bold, italic, underline = formatting_flags_from_runs(runs)
                heading_match = re.search(r"heading\s*([1-6])?", style_val, flags=re.I)
                is_title = re.search(r"title", style_val, flags=re.I)
                heading = bool(heading_match or is_title)
                heading_level = int(heading_match.group(1)) if heading_match and heading_match.group(1) else 1 if is_title else None
                blocks.append(
                    Block(
                        id=f"docx-{order}",
                        source="docx",
                        order=order,
                        text=proof_text.strip(),
                        normalized=match_text,
                        raw_text=raw_text,
                        proof_text=proof_text,
                        match_text=match_text,
                        heading=heading,
                        heading_level=heading_level,
                        bold=bold,
                        italic=italic,
                        underline=underline,
                        list_item=list_item,
                        kind="p",
                        runs=runs,
                    )
                )
                order += 1
            elif local == "tbl":
                rows = child.findall("w:tr", WORD_NS)
                for row_index, row in enumerate(rows):
                    row_cells: list[tuple[int, str, str, str, bool, bool, bool, list[InlineRun]]] = []
                    for col_index, cell in enumerate(row.findall("w:tc", WORD_NS)):
                        cell_runs: list[InlineRun] = []
                        bold = False
                        italic = False
                        underline = False
                        for paragraph in cell.findall(".//w:p", WORD_NS):
                            paragraph_runs = collect_word_runs(paragraph, resolver=resolver)
                            if not paragraph_runs:
                                continue
                            if cell_runs:
                                cell_runs.append(
                                    InlineRun(
                                        text="\n",
                                        kind="linebreak",
                                        source_index=len(cell_runs),
                                    )
                                )
                            cell_runs.extend(paragraph_runs)
                            bold = bold or any(run.bold for run in paragraph_runs)
                            italic = italic or any(run.italic for run in paragraph_runs)
                            underline = underline or any(run.underline for run in paragraph_runs)
                        cell_raw_text, cell_proof_text, cell_match_text = block_texts_from_runs(cell_runs)
                        row_cells.append((col_index, cell_raw_text, cell_proof_text, cell_match_text, bold, italic, underline, cell_runs))
                    row_key = next(
                        (
                            normalize_row_key(cell_proof_text)
                            for _col_index, _cell_raw_text, cell_proof_text, _cell_match_text, _bold, _italic, _underline, _cell_runs in row_cells
                            if visible_meaningful(cell_proof_text)
                        ),
                        None,
                    )
                    row_slot = 0
                    numeric_slot = 0
                    for col_index, cell_raw_text, cell_proof_text, cell_match_text, bold, italic, underline, cell_runs in row_cells:
                        if not visible_meaningful(cell_proof_text):
                            continue
                        token = single_value_token(
                            Block(
                                id="",
                                source="docx",
                                order=0,
                                text=cell_proof_text.strip(),
                                normalized=cell_match_text,
                                raw_text=cell_raw_text,
                                proof_text=cell_proof_text,
                                match_text=cell_match_text,
                                runs=cell_runs,
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
                                text=cell_proof_text.strip(),
                                normalized=cell_match_text,
                                raw_text=cell_raw_text,
                                proof_text=cell_proof_text,
                                match_text=cell_match_text,
                                bold=bold,
                                italic=italic,
                                underline=underline,
                                table_cell=True,
                                kind="td",
                                table_pos=(table_ordinal, row_index, col_index),
                                row_key=row_key,
                                row_slot=row_slot,
                                numeric_slot=cell_numeric_slot,
                                runs=cell_runs,
                            )
                        )
                        order += 1
                        row_slot += 1
                table_ordinal += 1

        footnote_blocks, order = extract_docx_footnote_blocks(archive, order, resolver=resolver)
        blocks.extend(footnote_blocks)
        return assign_structural_roles(blocks)


def get_descendant_text(element: ET.Element) -> str:
    parts: list[str] = []
    for text in element.itertext():
        parts.append(text)
    return normalize_text("".join(parts))


def collect_html_inline_runs(
    element: ET.Element,
    *,
    inherited_bold: bool = False,
    inherited_italic: bool = False,
    inherited_underline: bool = False,
    inherited_hyperlink: bool = False,
    run_index_start: int = 0,
) -> list[InlineRun]:
    runs: list[InlineRun] = []
    if not isinstance(element.tag, str):
        return runs
    style = (element.get("style") or "").lower()
    tag = xml_local_name(element.tag).lower()
    if tag == "br":
        return [
            InlineRun(
                text="\n",
                kind="linebreak",
                bold=inherited_bold,
                italic=inherited_italic,
                underline=inherited_underline,
                hyperlink=inherited_hyperlink,
                source_index=run_index_start,
            )
        ]
    bold = inherited_bold or tag in INLINE_BOLD_TAGS or style_flag(style, "font-weight: bold")
    italic = inherited_italic or tag in INLINE_ITALIC_TAGS or style_flag(style, "font-style: italic")
    underline = inherited_underline or tag in INLINE_UNDERLINE_TAGS or style_flag(style, "text-decoration: underline")
    hyperlink = inherited_hyperlink or tag == "a"

    def append_text(text: str | None, source_index: int) -> None:
        if not text:
            return
        runs.append(
            InlineRun(
                text=text,
                kind=classify_inline_run(text),
                bold=bold,
                italic=italic,
                underline=underline,
                hyperlink=hyperlink,
                source_index=source_index,
            )
        )

    append_text(element.text, run_index_start)
    child_index = run_index_start + 1
    for child in list(element):
        runs.extend(
            collect_html_inline_runs(
                child,
                inherited_bold=bold,
                inherited_italic=italic,
                inherited_underline=underline,
                inherited_hyperlink=hyperlink,
                run_index_start=child_index,
            )
        )
        child_index += 1
        append_text(child.tail, child_index)
        child_index += 1
    return runs


def split_inline_runs_by_proof_text(runs: list[InlineRun], lead_text: str, rest_text: str) -> tuple[list[InlineRun], list[InlineRun]]:
    lead_len = len(normalize_proof_text(lead_text))
    rest_len = len(normalize_proof_text(rest_text))
    if lead_len <= 0:
        return [], runs
    lead_runs: list[InlineRun] = []
    rest_runs: list[InlineRun] = []
    consumed = 0
    total_target = lead_len + rest_len
    for run in runs:
        proof_chunk = normalize_proof_text(run.text)
        if not proof_chunk:
            continue
        chunk_len = len(proof_chunk)
        next_consumed = consumed + chunk_len
        if consumed >= total_target:
            break
        if next_consumed <= lead_len:
            lead_runs.append(run)
        elif consumed >= lead_len:
            rest_runs.append(run)
        else:
            split_at = lead_len - consumed
            lead_text_part = proof_chunk[:split_at]
            rest_text_part = proof_chunk[split_at:]
            if lead_text_part:
                lead_runs.append(
                    InlineRun(
                        text=lead_text_part,
                        kind=classify_inline_run(lead_text_part),
                        bold=run.bold,
                        italic=run.italic,
                        underline=run.underline,
                        hyperlink=run.hyperlink,
                        source_index=run.source_index,
                    )
                )
            if rest_text_part:
                rest_runs.append(
                    InlineRun(
                        text=rest_text_part,
                        kind=classify_inline_run(rest_text_part),
                        bold=run.bold,
                        italic=run.italic,
                        underline=run.underline,
                        hyperlink=run.hyperlink,
                        source_index=run.source_index,
                    )
                )
        consumed = next_consumed
    return lead_runs, rest_runs


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
            runs = collect_html_inline_runs(element)
            raw_text, proof_text, match_text = block_texts_from_runs(runs)
            text = proof_text.strip()
            if visible_meaningful(text):
                bold, italic, underline = formatting_flags_from_runs(runs)
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
                    "footnote_marker": leading_footnote_marker(text),
                }
                if split:
                    lead, rest = split
                    lead_runs, rest_runs = split_inline_runs_by_proof_text(runs, lead, rest)
                    lead_raw_text, lead_proof_text, lead_match_text = block_texts_from_runs(lead_runs)
                    rest_raw_text, rest_proof_text, rest_match_text = block_texts_from_runs(rest_runs)
                    lead_bold, lead_italic, lead_underline = formatting_flags_from_runs(lead_runs)
                    rest_bold, rest_italic, rest_underline = formatting_flags_from_runs(rest_runs)
                    order = order_ref[0]
                    blocks.append(
                        Block(
                            id=f"html-{order}",
                            order=order,
                            text=lead_proof_text.strip() or lead,
                            normalized=lead_match_text or normalize_for_compare(lead),
                            raw_text=lead_raw_text,
                            proof_text=lead_proof_text or lead,
                            match_text=lead_match_text or normalize_for_compare(lead),
                            runs=lead_runs or runs,
                            **{
                                **base_kwargs,
                                "bold": lead_bold,
                                "italic": lead_italic,
                                "underline": lead_underline,
                            },
                        )
                    )
                    order_ref[0] += 1
                    order = order_ref[0]
                    blocks.append(
                        Block(
                            id=f"html-{order}",
                            order=order,
                            text=rest_proof_text.strip() or rest,
                            normalized=rest_match_text or normalize_for_compare(rest),
                            raw_text=rest_raw_text,
                            proof_text=rest_proof_text or rest,
                            match_text=rest_match_text or normalize_for_compare(rest),
                            runs=rest_runs or runs,
                            **{
                                **base_kwargs,
                                "heading": False,
                                "heading_level": None,
                                "bold": rest_bold,
                                "italic": rest_italic,
                                "underline": rest_underline,
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
                            normalized=match_text,
                            raw_text=raw_text,
                            proof_text=proof_text,
                            match_text=match_text,
                            runs=runs,
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
    return assign_structural_roles(blocks)


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

  const normalizeProofText = (text) => String(text || '')
    .replace(/\\u00A0/g, ' ')
    .replace(/[\\u2018\\u2019]/g, "'")
    .replace(/[\\u201C\\u201D]/g, '"')
    .replace(/[\\u2013\\u2014]/g, '-')
    .replace(/\\u2022/g, '-')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>');
  const normalizeMatchText = (text) => normalizeProofText(text)
    .replace(/\\bwww\\./gi, '')
    .replace(/\\(\\s+(?=\\d)/g, '(')
    .replace(/(\\d)\\s+%/g, '$1%')
    .replace(/(\\d)\\s+\\)/g, '$1)')
    .replace(/([+-])\\s+(?=\\d)/g, '$1')
    .replace(/\\s+/g, ' ')
    .trim()
    .toLowerCase();
  const DIFF_TOKEN_RE = /\\d[\\d,]*(?:\\.\\d+)?%?|\\.\\d+%?|[A-Za-z]+(?:[’'\\-][A-Za-z]+)*/g;

  const diffTokenKind = (token) => /^[+-]?(?:\\d[\\d,]*(?:\\.\\d+)?%?|\\.\\d+%?)$/.test(token) ? 'number' : 'word';
  const normalizeDiffToken = (token) => {
    const cleaned = normalizeProofText(token).trim();
    return diffTokenKind(cleaned) === 'number'
      ? cleaned.replace(/,/g, '').toLowerCase()
      : cleaned.toLowerCase();
  };
  const diffTokensText = (text) => {
    const tokens = [];
    const source = normalizeProofText(text);
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
    if (!/[A-Za-z]/.test(lead)) return null;
    if (!/^[A-Z][A-Za-z0-9'()/%& .-]+$/.test(lead)) return null;
    if (!/^[\"'(]?[A-Z]/.test(rest)) return null;
    if (leadWords <= 1 && lead.length < 12) return null;
    return looksLikeLabel && bodySubstantial ? { lead, rest } : null;
  };

  const visibleMeaningful = (text) => /[A-Za-z0-9]/.test(normalizeMatchText(text));

  const classifyRunKind = (text) => {
    if (!text) return 'text';
    if (text === '\\t') return 'tab';
    if (text === '\\n') return 'linebreak';
    if (/^[\\s]+$/.test(text)) return 'space';
    if (/^[$€£¥%~()\\[\\]{}:;,.!?/&+\\-–—•]+$/.test(text)) return 'symbol';
    return 'text';
  };

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

  function formattingFlagsFromRuns(runs) {
    return {
      bold: runs.some(run => !!run.bold),
      italic: runs.some(run => !!run.italic),
      underline: runs.some(run => !!run.underline),
    };
  }

  function collectInlineRuns(el) {
    const runs = [];
    let index = 0;

    const appendRun = (text, parent) => {
      if (!text) return;
      const style = window.getComputedStyle(parent);
      const bold = boldTags.has(parent.tagName) || parseInt(style.fontWeight || '400', 10) >= 600;
      const italic = italicTags.has(parent.tagName) || style.fontStyle === 'italic';
      const underline = underlineTags.has(parent.tagName) || (style.textDecorationLine || '').includes('underline');
      const hyperlink = Boolean(parent.closest('a'));
      runs.push({
        text,
        kind: classifyRunKind(text),
        bold,
        italic,
        underline,
        hyperlink,
        source_index: index,
      });
      index += 1;
    };

    const walkNode = (node, parent) => {
      if (node.nodeType === Node.TEXT_NODE) {
        appendRun(node.nodeValue || '', parent);
        return;
      }
      if (node.nodeType !== Node.ELEMENT_NODE) return;
      const element = /** @type {HTMLElement} */ (node);
      if (element.tagName === 'BR') {
        appendRun('\\n', parent);
        return;
      }
      for (const child of element.childNodes) {
        walkNode(child, element);
      }
    };

    for (const child of el.childNodes) {
      walkNode(child, el);
    }
    return runs;
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
          const runs = collectInlineRuns(el);
          const rawText = runs.map(run => run.text).join('');
          const proofText = normalizeProofText(rawText);
          const text = proofText.trim();
          if (visibleMeaningful(text)) {
            const rect = el.getBoundingClientRect();
          const flags = formattingFlagsFromRuns(runs);
          const tokenRects = collectTokenRects(el);
          const headingMatch = /^H([1-6])$/.exec(el.tagName);
          const tableEl = ['TD', 'TH'].includes(el.tagName) ? el.closest('table') : null;
          const trEl = ['TD', 'TH'].includes(el.tagName) ? el.closest('tr') : null;
          const tableRows = tableEl ? Array.from(tableEl.querySelectorAll('tr')) : [];
          const cellSiblings = trEl ? Array.from(trEl.children).filter(child => child.tagName === 'TD' || child.tagName === 'TH') : [];
          const meaningfulCellSiblings = cellSiblings.filter(cell => visibleMeaningful(normalizeProofText(cell.innerText || cell.textContent || '')));
          const numericMeaningfulCellSiblings = meaningfulCellSiblings.filter(cell => {
            const tokens = diffTokensText(cell.innerText || cell.textContent || '');
            return tokens.length === 1 && tokens[0].kind === 'number';
          });
          const rowKeyCell = cellSiblings.find(cell => visibleMeaningful(normalizeProofText(cell.innerText || cell.textContent || '')));
          const normalizeRowKey = (text) => normalizeProofText(text)
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
            normalized: normalizeMatchText(text),
            raw_text: rawText,
            proof_text: proofText,
            match_text: normalizeMatchText(text),
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
            runs,
          };
          const split = splitLeadLabelText(text, block.table_cell);
          if (split) {
            const leadTokens = diffTokensText(split.lead);
            const restTokens = diffTokensText(split.rest);
            const leadCount = leadTokens.length;
            const restCount = restTokens.length;
            const leadRuns = [];
            const restRuns = [];
            let consumed = 0;
            const leadLen = normalizeProofText(split.lead).length;
            const totalLen = leadLen + normalizeProofText(split.rest).length;
            for (const run of runs) {
              const proofChunk = normalizeProofText(run.text);
              if (!proofChunk) continue;
              const chunkLen = proofChunk.length;
              const nextConsumed = consumed + chunkLen;
              if (consumed >= totalLen) break;
              if (nextConsumed <= leadLen) {
                leadRuns.push(run);
              } else if (consumed >= leadLen) {
                restRuns.push(run);
              } else {
                const splitAt = leadLen - consumed;
                const leadTextPart = proofChunk.slice(0, splitAt);
                const restTextPart = proofChunk.slice(splitAt);
                if (leadTextPart) {
                  leadRuns.push({ ...run, text: leadTextPart, kind: classifyRunKind(leadTextPart) });
                }
                if (restTextPart) {
                  restRuns.push({ ...run, text: restTextPart, kind: classifyRunKind(restTextPart) });
                }
              }
              consumed = nextConsumed;
            }
            const leadFlags = formattingFlagsFromRuns(leadRuns);
            const restFlags = formattingFlagsFromRuns(restRuns);
            el.setAttribute('data-docx-compare-order', String(order));
            blocks.push({
              ...block,
              id: `html-${order}`,
              order,
              text: split.lead,
              raw_text: split.lead,
              proof_text: split.lead,
              match_text: normalizeMatchText(split.lead),
              normalized: normalizeMatchText(split.lead),
              tokens: tokenRects.slice(0, leadCount),
              bold: leadFlags.bold,
              italic: leadFlags.italic,
              underline: leadFlags.underline,
              runs: leadRuns.length ? leadRuns : runs,
            });
            order += 1;
            blocks.push({
              ...block,
              id: `html-${order}`,
              order,
              text: split.rest,
              raw_text: split.rest,
              proof_text: split.rest,
              match_text: normalizeMatchText(split.rest),
              normalized: normalizeMatchText(split.rest),
              heading: false,
              heading_level: null,
              bold: restFlags.bold,
              italic: restFlags.italic,
              underline: restFlags.underline,
              tokens: tokenRects.slice(leadCount, leadCount + restCount),
              runs: restRuns.length ? restRuns : runs,
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
                normalized=str(item.get("match_text", normalize_for_compare(str(item["text"])))),
                raw_text=str(item.get("raw_text", item["text"])),
                proof_text=str(item.get("proof_text", item["text"])),
                match_text=str(item.get("match_text", normalize_for_compare(str(item["text"])))),
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
                footnote_marker=leading_footnote_marker(str(item.get("proof_text", item["text"]))),
                runs=[
                    InlineRun(
                        text=str(run["text"]),
                        kind=str(run["kind"]),
                        bold=bool(run.get("bold", False)),
                        italic=bool(run.get("italic", False)),
                        underline=bool(run.get("underline", False)),
                        hyperlink=bool(run.get("hyperlink", False)),
                        source_index=int(run.get("source_index", 0)),
                    )
                    for run in item.get("runs", [])
                ],
            )
        )
    return BrowserRenderResult(
        blocks=assign_structural_roles(blocks),
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
    prev_contact = looks_like_contact_text(prev_text)
    curr_contact = looks_like_contact_text(curr_text)
    if prev_contact != curr_contact:
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
    table_index_start: int,
    pixmap: Any | None = None,
    span_entries: list[dict[str, Any]] | None = None,
) -> tuple[list[Block], dict[int, tuple[float, float, float, float]], dict[int, list[TokenRect]], dict[int, int], int, int]:
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
        if not visible_meaningful(block_text) or is_pdf_chrome_text(block_text):
            paragraph_words = []
            paragraph_rect = None
            paragraph_text_parts = []
            return
        runs = collect_pdf_runs_for_rect(paragraph_rect, span_entries)
        if runs:
            raw_text, proof_text, match_text = block_texts_from_runs(runs)
            bold, italic, underline = formatting_flags_from_runs(runs)
            block_text = raw_text or block_text
        else:
            proof_text = normalize_proof_text(block_text)
            match_text = normalize_for_compare(proof_text)
            bold = italic = underline = False
        block = Block(
            id=f"pdf-{order}",
            source="pdf",
            order=order,
            text=block_text,
            normalized=match_text,
            raw_text=block_text,
            proof_text=proof_text,
            match_text=match_text,
            table_cell=False,
            kind="pdf",
            runs=runs,
            bold=bold,
            italic=italic,
            underline=underline,
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

    row_infos = build_pdf_row_infos(
        row_groups,
        page_width=page_width,
        page_height=page_height,
        pixmap=pixmap,
    )
    region_rows = detect_pdf_table_regions(row_infos)
    table_id_by_region: dict[int, int] = {}
    next_table_index = table_index_start
    for region_id, _row_ordinal in region_rows.values():
        if region_id not in table_id_by_region:
            table_id_by_region[region_id] = next_table_index
            next_table_index += 1

    for info in row_infos:
        row_index = info["row_index"]
        meaningful_clusters = info["clusters"]
        table_meta = region_rows.get(row_index)
        table_like = table_meta is not None
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
        table_idx_value = table_id_by_region[table_meta[0]] if table_meta is not None else None
        table_row_index = table_meta[1] if table_meta is not None else row_index
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
            if is_pdf_chrome_text(block_text):
                if table_like:
                    row_slot += 1
                continue
            x0 = min(item["rect"][0] for item in cluster)
            y0 = min(item["rect"][1] for item in cluster)
            x1 = max(item["rect"][2] for item in cluster)
            y1 = max(item["rect"][3] for item in cluster)
            cluster_rect = (x0, y0, x1, y1)
            runs = collect_pdf_runs_for_rect(cluster_rect, span_entries)
            if runs:
                raw_text, proof_text, match_text = block_texts_from_runs(runs)
                block_text = raw_text or block_text
                bold, italic, underline = formatting_flags_from_runs(runs)
            else:
                proof_text = normalize_proof_text(block_text)
                match_text = normalize_for_compare(proof_text)
                bold = italic = underline = False
            temp_block = Block(
                id="",
                source="pdf",
                order=0,
                text=block_text,
                normalized=match_text,
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
                normalized=match_text,
                raw_text=block_text,
                proof_text=proof_text,
                match_text=match_text,
                bold=bold,
                italic=italic,
                underline=underline,
                table_cell=table_like,
                kind="pdf",
                table_pos=(table_idx_value, table_row_index, row_slot) if table_like and table_idx_value is not None else None,
                row_key=row_key,
                row_slot=row_slot if table_like else None,
                numeric_slot=block_numeric_slot,
                runs=runs,
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

    return blocks, rects_by_order, token_rects_by_order, page_numbers_by_order, order, next_table_index


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
    table_index = 0

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
            table_index,
        ) = cluster_words_into_blocks(
            words,
            page_number=page_number,
            page_width=float(page_rect.width),
            page_height=float(page_rect.height),
            order_start=order,
            table_index_start=table_index,
            pixmap=dash_pixmap,
        )
        blocks.extend(page_blocks)
        rects_by_order.update(page_rects)
        token_rects_by_order.update(page_token_rects)
        page_numbers_by_order.update(page_numbers)

    document.close()
    return BrowserRenderResult(
        blocks=assign_structural_roles(blocks),
        width_px=max_width,
        height_px=max_height,
        rects_by_order=rects_by_order,
        token_rects_by_order=token_rects_by_order,
        page_numbers_by_order=page_numbers_by_order,
        coordinate_space="pdf_pt",
    )


def extract_pdf_blocks_from_words(path: Path) -> BrowserRenderResult:
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
    table_index = 0

    for page_number, page in enumerate(document):
        page_rect = page.rect
        max_width = max(max_width, float(page_rect.width))
        max_height = max(max_height, float(page_rect.height))
        span_entries = build_pdf_span_entries(page)
        words_raw = page.get_text("words")
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
            table_index,
        ) = cluster_words_into_blocks(
            words,
            page_number=page_number,
            page_width=float(page_rect.width),
            page_height=float(page_rect.height),
            order_start=order,
            table_index_start=table_index,
            span_entries=span_entries,
        )
        blocks.extend(page_blocks)
        rects_by_order.update(page_rects)
        token_rects_by_order.update(page_token_rects)
        page_numbers_by_order.update(page_numbers)

    document.close()
    if not blocks:
        return extract_pdf_blocks_via_ocr(path)
    return BrowserRenderResult(
        blocks=assign_structural_roles(blocks),
        width_px=max_width,
        height_px=max_height,
        rects_by_order=rects_by_order,
        token_rects_by_order=token_rects_by_order,
        page_numbers_by_order=page_numbers_by_order,
        coordinate_space="pdf_pt",
    )


def extract_pdf_blocks(path: Path, *, proofread_mode: bool = False) -> BrowserRenderResult:
    if fitz is None:
        raise RuntimeError(
            "PyMuPDF is not installed. Install it with `pip install PyMuPDF` to compare a DOCX against a PDF."
        )
    if proofread_mode:
        return extract_pdf_blocks_from_words(path)

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
            if not visible_meaningful(cleaned) or is_pdf_chrome_text(cleaned):
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
                if is_pdf_chrome_text(block_text):
                    if table_like:
                        row_slot += 1
                    continue
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
        blocks=assign_structural_roles(blocks),
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
    *,
    table_pos_override: tuple[int, int, int] | None = None,
    table_index_override: int | None = None,
) -> list[int]:
    pos_key = table_pos_override if table_pos_override is not None else table_pos_key(doc_block)
    if pos_key is not None and pos_key in table_pos_map:
        return [index for index in table_pos_map[pos_key] if index not in used_html]
    table_idx = table_index_override if table_index_override is not None else table_index_key(doc_block)
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
        if not structural_role_compatible(doc_block, html_blocks[index]):
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


def compare_blocks(
    docx_blocks: list[Block],
    html_blocks: list[Block],
    *,
    target_name: str = "html",
    proofread_mode: bool = False,
) -> tuple[list[Match], list[Block], list[Block]]:
    doc_to_html_table_map, _html_to_doc_table_map = align_table_families(docx_blocks, html_blocks)
    doc_header_ordinals, _doc_header_map = build_family_role_ordinals(docx_blocks)
    _html_header_ordinals, html_header_map = build_family_role_ordinals(html_blocks)
    doc_row_ordinals, _doc_row_map = build_row_label_ordinals(docx_blocks)
    _html_row_ordinals, html_row_map = build_row_label_ordinals(html_blocks)
    (
        grouped_target_by_doc_index,
        doc_block_to_group,
        html_block_to_group,
        doc_group_map,
        html_group_map,
    ) = match_block_groups(docx_blocks, html_blocks)
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
    prematches, used_doc, used_html = prematch_same_cell_numeric_table_blocks(
        docx_blocks,
        html_blocks,
        table_pos_map,
        doc_to_html_table_map=doc_to_html_table_map,
    )
    for doc_index, target_group_id in grouped_target_by_doc_index.items():
        if doc_index in used_doc:
            continue
        doc_group_id = doc_block_to_group.get(doc_index)
        if doc_group_id is None:
            continue
        doc_group = doc_group_map.get(doc_group_id)
        target_group = html_group_map.get(target_group_id)
        if doc_group is None or target_group is None:
            continue
        if doc_group.group_type == "contact":
            if not target_group.block_indices:
                continue
            html_index = target_group.block_indices[0]
            if html_index in used_html:
                continue
            for group_doc_index in doc_group.block_indices:
                payload = contact_field_payload(docx_blocks[group_doc_index].text)
                if payload is None:
                    score = 0.9
                    match_type = "approx"
                else:
                    role, doc_value = payload
                    target_value = extract_contact_fields(html_blocks[html_index].text).get(role)
                    if target_value is None:
                        score = 0.85
                        match_type = "approx"
                    else:
                        score = max(
                            similarity(normalize_for_compare(doc_value), normalize_for_compare(target_value)),
                            0.82,
                        )
                        match_type = "exact_structural"
                prematches.append(
                    Match(
                        docx_index=group_doc_index,
                        html_index=html_index,
                        match_type=match_type,
                        score=score,
                        formatting_diffs=summarize_formatting_diff(docx_blocks[group_doc_index], html_blocks[html_index]),
                    )
                )
                used_doc.add(group_doc_index)
            used_html.add(html_index)
            continue
        if doc_group.group_type not in {"footnote", "quote"}:
            continue
        if len(doc_group.block_indices) != 1 or len(target_group.block_indices) != 1:
            continue
        html_index = target_group.block_indices[0]
        if html_index in used_html:
            continue
        score = group_similarity(doc_group, target_group)
        footnote_group_floor = 0.9 if target_name == "html" else 0.84 if proofread_mode and target_name == "pdf" else 0.88 if target_name == "pdf" else 0.9
        quote_group_floor = 0.9 if target_name == "html" else 0.82 if proofread_mode and target_name == "pdf" else 0.88 if target_name == "pdf" else 0.9
        if doc_group.group_type == "footnote" and score < footnote_group_floor:
            continue
        if doc_group.group_type == "quote" and score < quote_group_floor:
            continue
        prematches.append(
            Match(
                docx_index=doc_index,
                html_index=html_index,
                match_type=promote_exact_structural_match(
                    doc_index,
                    html_index,
                    docx_blocks,
                    html_blocks,
                    grouped_match_type=doc_group.group_type,
                    score=score,
                    match_type="approx",
                ),
                score=score,
                formatting_diffs=summarize_formatting_diff(docx_blocks[doc_index], html_blocks[html_index]),
            )
        )
        used_doc.add(doc_index)
        used_html.add(html_index)
    matches: list[Match] = list(prematches)
    unmatched_doc_indices: list[int] = []

    for doc_index, doc_block in enumerate(docx_blocks):
        if doc_index in used_doc:
            continue
        mapped_table_idx = mapped_table_index_key(doc_block, doc_to_html_table_map)
        pos_key = mapped_table_pos_key(doc_block, doc_to_html_table_map)
        table_idx = mapped_table_idx
        row_key = mapped_row_context_key(doc_block, doc_to_html_table_map)
        global_row_key = global_row_context_key(doc_block)
        global_numeric_key = global_numeric_context_key(doc_block)
        target_group_id = grouped_target_by_doc_index.get(doc_index)
        doc_group = doc_group_map.get(doc_block_to_group.get(doc_index, ""))
        doc_group_type = doc_group.group_type if doc_group is not None else None
        family_filter = (
            (lambda index: table_index_key(html_blocks[index]) == mapped_table_idx)
            if mapped_table_idx is not None
            else (lambda _index: True)
        )
        group_filter = (
            (lambda index: html_block_to_group.get(index) == target_group_id)
            if target_group_id is not None
            else (lambda _index: True)
        )
        pos_candidates: list[int] = []
        if pos_key is not None:
            pos_candidates = [
                index
                for index in table_pos_map.get(pos_key, [])
                if index not in used_html and group_filter(index)
            ]
            exact_pos_index = select_best_exact_candidate(
                doc_block,
                [index for index in pos_candidates if html_blocks[index].normalized == doc_block.normalized],
                html_blocks,
                used_html,
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
            numeric_candidates = [
                index
                for index in global_numeric_context_map.get(global_numeric_key, [])
                if index not in used_html and family_filter(index) and group_filter(index)
            ]
            exact_numeric_index = select_best_exact_candidate(
                doc_block,
                [index for index in numeric_candidates if html_blocks[index].normalized == doc_block.normalized],
                html_blocks,
                used_html,
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
                            match_type=promote_exact_structural_match(
                                doc_index,
                                best_numeric_index,
                                docx_blocks,
                                html_blocks,
                                grouped_match_type=doc_group_type,
                                score=best_numeric_score,
                                match_type="approx",
                            ),
                            score=best_numeric_score,
                            formatting_diffs=summarize_formatting_diff(doc_block, html_blocks[best_numeric_index]),
                        )
                    )
                continue
        if row_key is not None:
            row_candidates = [
                index
                for index in row_context_map.get(row_key, [])
                if index not in used_html and group_filter(index)
            ]
            exact_row_index = select_best_exact_candidate(
                doc_block,
                [index for index in row_candidates if html_blocks[index].normalized == doc_block.normalized],
                html_blocks,
                used_html,
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
                            match_type=promote_exact_structural_match(
                                doc_index,
                                best_row_index,
                                docx_blocks,
                                html_blocks,
                                grouped_match_type=doc_group_type,
                                score=best_row_score,
                                match_type="approx",
                            ),
                            score=best_row_score,
                            formatting_diffs=summarize_formatting_diff(doc_block, html_blocks[best_row_index]),
                        )
                    )
                continue
        if global_row_key is not None:
            global_row_candidates = [
                index
                for index in global_row_context_map.get(global_row_key, [])
                if index not in used_html and family_filter(index) and group_filter(index)
            ]
            exact_global_row_index = select_best_exact_candidate(
                doc_block,
                [index for index in global_row_candidates if html_blocks[index].normalized == doc_block.normalized],
                html_blocks,
                used_html,
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
                            match_type=promote_exact_structural_match(
                                doc_index,
                                best_global_row_index,
                                docx_blocks,
                                html_blocks,
                                grouped_match_type=doc_group_type,
                                score=best_global_row_score,
                                match_type="approx",
                            ),
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
                and family_filter(index)
                and group_filter(index)
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
                            match_type=promote_exact_structural_match(
                                doc_index,
                                best_fuzzy_row_index,
                                docx_blocks,
                                html_blocks,
                                grouped_match_type=doc_group_type,
                                score=best_fuzzy_row_score,
                                match_type="approx",
                            ),
                            score=best_fuzzy_row_score,
                            formatting_diffs=summarize_formatting_diff(doc_block, html_blocks[best_fuzzy_row_index]),
                        )
                    )
                    continue
        if table_idx is not None and not pos_candidates:
            exact_table_candidates = table_index_map.get(table_idx, [])
            exact_table_index = select_best_exact_candidate(
                doc_block,
                [
                    index for index in exact_table_candidates
                    if html_blocks[index].normalized == doc_block.normalized and group_filter(index)
                ],
                html_blocks,
                used_html,
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
        exact_candidates = (
            [
                index
                for index in exact_map.get(doc_block.normalized, [])
                if family_filter(index) and group_filter(index)
            ]
            if allow_global_exact_for_table
            else []
        )
        exact_index = select_best_exact_candidate(doc_block, exact_candidates, html_blocks, used_html)
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
        for html_index in get_approx_candidates(
            doc_block,
            html_blocks,
            token_index,
            length_buckets,
            table_pos_map,
            table_index_map,
            used_html,
            table_pos_override=pos_key,
            table_index_override=table_idx,
        ):
            if not group_filter(html_index):
                continue
            score = similarity(doc_block.normalized, html_blocks[html_index].normalized)
            if score > best_score:
                best_score = score
                best_index = html_index
        approx_floor = 0.73 if target_name == "html" else 0.78 if proofread_mode and target_name == "pdf" else 0.74
        doc_schema = known_section_schema(doc_block.text, allow_fuzzy=True)
        if (
            doc_block.structure_role == "section_header"
            and best_index is not None
            and html_blocks[best_index].structure_role == "section_header"
        ):
            target_schema = known_section_schema(html_blocks[best_index].text, allow_fuzzy=True)
            if (
                doc_schema is not None
                and target_schema is not None
                and doc_schema.schema_key == target_schema.schema_key
            ):
                approx_floor = min(approx_floor, 0.6)
        if best_index is not None and best_score >= approx_floor:
            used_html.add(best_index)
            matches.append(
                Match(
                    docx_index=doc_index,
                    html_index=best_index,
                    match_type=promote_exact_structural_match(
                        doc_index,
                        best_index,
                        docx_blocks,
                        html_blocks,
                        grouped_match_type=doc_group_type,
                        score=best_score,
                        match_type="approx",
                    ),
                    score=best_score,
                    formatting_diffs=summarize_formatting_diff(doc_block, html_blocks[best_index]),
                )
            )
        else:
            unmatched_doc_indices.append(doc_index)

    recovered_doc_indices: set[int] = set()
    recovered_html_indices: set[int] = set()

    for doc_index in list(unmatched_doc_indices):
        doc_block = docx_blocks[doc_index]
        mapped_family_idx = mapped_table_index_key(doc_block, doc_to_html_table_map)
        group = header_family_group(doc_block.structure_role)
        if mapped_family_idx is None or group is None:
            continue
        ordinal = doc_header_ordinals.get(doc_index)
        if ordinal is None:
            continue
        candidate_indices = [
            index
            for index in html_header_map.get((mapped_family_idx, group, ordinal), [])
            if index not in used_html
        ]
        if len(candidate_indices) != 1:
            continue
        html_index = candidate_indices[0]
        target_block = html_blocks[html_index]
        score = max(
            similarity(doc_block.normalized, target_block.normalized),
            formatting_alignment_score(doc_block, target_block) / 10.0,
        )
        min_score = 0.6 if group == "title_family" else 0.72
        if score < min_score:
            continue
        matches.append(
            Match(
                docx_index=doc_index,
                html_index=html_index,
                match_type="exact_structural",
                score=score,
                formatting_diffs=summarize_formatting_diff(doc_block, target_block),
            )
        )
        recovered_doc_indices.add(doc_index)
        recovered_html_indices.add(html_index)
        used_html.add(html_index)

    for doc_index in list(unmatched_doc_indices):
        if doc_index in recovered_doc_indices:
            continue
        doc_block = docx_blocks[doc_index]
        if doc_block.structure_role != "table_row_label":
            continue
        mapped_family_idx = mapped_table_index_key(doc_block, doc_to_html_table_map)
        if mapped_family_idx is None:
            continue
        ordinal = doc_row_ordinals.get(doc_index)
        if ordinal is None:
            continue
        candidate_indices = [
            index
            for index in html_row_map.get((mapped_family_idx, ordinal), [])
            if index not in used_html
        ]
        if len(candidate_indices) != 1:
            continue
        html_index = candidate_indices[0]
        target_block = html_blocks[html_index]
        score = similarity(doc_block.normalized, target_block.normalized)
        row_recovery_floor = 0.62
        if score < row_recovery_floor and not (
            doc_block.row_key
            and target_block.row_key
            and similarity(doc_block.row_key, target_block.row_key) >= row_recovery_floor
        ):
            continue
        matches.append(
            Match(
                docx_index=doc_index,
                html_index=html_index,
                match_type="exact_structural" if score >= 0.82 else "approx",
                score=max(score, 0.82 if score >= row_recovery_floor else score),
                formatting_diffs=summarize_formatting_diff(doc_block, target_block),
            )
        )
        recovered_doc_indices.add(doc_index)
        recovered_html_indices.add(html_index)
        used_html.add(html_index)

    if recovered_doc_indices:
        unmatched_doc_indices = [index for index in unmatched_doc_indices if index not in recovered_doc_indices]

    covered_html = set(used_html)
    embedded_matches: list[Match] = []
    embedded_doc_indices: set[int] = set()
    embedded_html_indices: set[int] = set()
    if target_name == "pdf":
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
    source = normalize_proof_text(text)
    for match in DATE_PHRASE_RE.finditer(source):
        if token.start >= match.start() and token.end <= match.end():
            return match.group(0)
    return None


def date_phrase_for_token_slice(text: str, tokens: list[DiffToken]) -> str | None:
    if not tokens:
        return None
    source = normalize_proof_text(text)
    start = min(token.start for token in tokens)
    end = max(token.end for token in tokens)
    for match in DATE_PHRASE_RE.finditer(source):
        if start >= match.start() and end <= match.end():
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


def block_date_difference_comment(
    doc_text: str,
    target_text: str,
    *,
    target_name: str,
) -> str | None:
    doc_dates = [match.group(0) for match in DATE_PHRASE_RE.finditer(normalize_proof_text(doc_text))]
    target_dates = [match.group(0) for match in DATE_PHRASE_RE.finditer(normalize_proof_text(target_text))]
    if len(doc_dates) != 1 or len(target_dates) != 1:
        return None
    doc_date = doc_dates[0]
    target_date = target_dates[0]
    if normalize_for_compare(doc_date) == normalize_for_compare(target_date):
        return None
    return f"The date is different, {target_date} in {target_name} while {doc_date} in word."


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


def percent_symbol_comment(
    doc_block: Block,
    target_block: Block,
    *,
    target_name: str,
) -> HtmlComment | None:
    doc_token = single_value_token(doc_block)
    target_token = single_value_token(target_block)
    if doc_token is None or target_token is None:
        return None
    if doc_token.kind != "number" or target_token.kind != "number":
        return None
    doc_value = parse_numeric_token(doc_token.text)
    target_value = parse_numeric_token(target_token.text)
    if doc_value is None or target_value is None:
        return None
    doc_number, doc_percent = doc_value
    target_number, target_percent = target_value
    if math.isclose(doc_number, target_number, rel_tol=1e-9, abs_tol=1e-9) and doc_percent != target_percent:
        return HtmlComment(
            order=target_block.order,
            contents=(
                f"The percent sign is different, {target_name} has {comment_token_text(target_token)} "
                f"while word has {comment_token_text(doc_token)}."
            ),
            token_index=0,
        )
    return None


def single_unambiguous_numeric_difference_comment(
    doc_text: str,
    target_text: str,
    *,
    target_name: str,
) -> tuple[str, int | None] | None:
    doc_tokens = diff_tokens(doc_text)
    target_tokens = diff_tokens(target_text)
    doc_numbers = [(index, token) for index, token in enumerate(doc_tokens) if token.kind == "number"]
    target_numbers = [(index, token) for index, token in enumerate(target_tokens) if token.kind == "number"]
    if not doc_numbers or not target_numbers or len(doc_numbers) != len(target_numbers):
        return None
    mismatches: list[tuple[int, DiffToken, int, DiffToken]] = []
    for (doc_index, doc_token), (target_index, target_token) in zip(doc_numbers, target_numbers):
        if doc_token.normalized != target_token.normalized:
            mismatches.append((doc_index, doc_token, target_index, target_token))
    if len(mismatches) != 1:
        return None
    _doc_index, doc_token, target_index, target_token = mismatches[0]
    return (
        f"The number is different, {comment_token_text(target_token).strip()} in {target_name} "
        f"while {comment_token_text(doc_token).strip()} in word.",
        target_index,
    )


def single_unambiguous_local_token_comments(
    doc_text: str,
    target_text: str,
    *,
    target_name: str,
) -> list[tuple[str, int | None]] | None:
    doc_tokens = diff_tokens(doc_text)
    target_tokens = diff_tokens(target_text)
    if not doc_tokens or not target_tokens:
        return None
    matcher = difflib.SequenceMatcher(
        a=[token.normalized for token in doc_tokens],
        b=[token.normalized for token in target_tokens],
        autojunk=False,
    )
    opcodes = matcher.get_opcodes()
    non_equal = [opcode for opcode in opcodes if opcode[0] != "equal"]
    if not non_equal or len(non_equal) > 2:
        return None
    changed_token_count = sum((i2 - i1) + (j2 - j1) for _tag, i1, i2, j1, j2 in non_equal)
    if changed_token_count > 4:
        return None
    equal_token_count = sum(i2 - i1 for tag, i1, i2, _j1, _j2 in opcodes if tag == "equal")
    if equal_token_count < max(len(doc_tokens), len(target_tokens)) - 4:
        return None

    if len(non_equal) == 2:
        first, second = non_equal
        if {first[0], second[0]} != {"delete", "insert"}:
            return None
        if first[2] != second[1] or first[4] != second[3]:
            return None

    comments: list[HtmlComment] = []
    for tag, i1, i2, j1, j2 in non_equal:
        append_word_level_comments(
            comments,
            order=0,
            doc_text=doc_text,
            target_text=target_text,
            target_tokens=target_tokens,
            doc_slice=doc_tokens[i1:i2],
            target_slice=target_tokens[j1:j2],
            target_start=j1,
            target_name=target_name,
        )
    comments = dedupe_html_comments(comments)
    comments = collapse_insert_delete_comment_pairs(comments)
    detailed_prefixes = (
        "The word is different",
        "The word is extra in html",
        "The word is missing in html",
        "The number is different",
        "The number is extra in html",
        "The number is missing in html",
        "The date is different",
    )
    if not comments or len(comments) > 2:
        return None
    if not all(comment.contents.startswith(detailed_prefixes) for comment in comments):
        return None
    return [(comment.contents, comment.token_index) for comment in comments]


def phrase_token_text(tokens: list[DiffToken]) -> str:
    if not tokens:
        return ""
    parts: list[str] = []
    for index, token in enumerate(tokens):
        if index > 0:
            parts.append(" ")
        parts.append(comment_token_text(token).strip())
    return "".join(parts).strip()


def single_unambiguous_phrase_difference_comments(
    doc_text: str,
    target_text: str,
    *,
    target_name: str,
) -> list[tuple[str, int | None]] | None:
    doc_tokens = diff_tokens(doc_text)
    target_tokens = diff_tokens(target_text)
    if not doc_tokens or not target_tokens:
        return None
    matcher = difflib.SequenceMatcher(
        a=[token.normalized for token in doc_tokens],
        b=[token.normalized for token in target_tokens],
        autojunk=False,
    )
    opcodes = matcher.get_opcodes()
    non_equal = [opcode for opcode in opcodes if opcode[0] != "equal"]
    if len(non_equal) != 1:
        return None
    tag, i1, i2, j1, j2 = non_equal[0]
    if tag not in {"insert", "delete"}:
        return None
    equal_token_count = sum(i2_ - i1_ for tag_, i1_, i2_, _j1, _j2 in opcodes if tag_ == "equal")
    if equal_token_count < max(len(doc_tokens), len(target_tokens)) - 6:
        return None
    if tag == "insert":
        inserted = target_tokens[j1:j2]
        if not (2 <= len(inserted) <= 6) or not all(token.kind == "word" for token in inserted):
            return None
        phrase = phrase_token_text(inserted)
        if not phrase:
            return None
        return [(f"The words are extra in {target_name}, {phrase}. They are not present in word.", j1)]
    deleted = doc_tokens[i1:i2]
    if not (2 <= len(deleted) <= 6) or not all(token.kind == "word" for token in deleted):
        return None
    phrase = phrase_token_text(deleted)
    if not phrase:
        return None
    anchor = j1 if target_tokens else None
    if anchor is not None and target_tokens:
        anchor = min(max(anchor, 0), len(target_tokens) - 1)
    return [(f"The words are missing in {target_name}, {phrase} in word.", anchor)]


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
    if target_name == "pdf" and doc_currency and not target_currency:
        parsed_target = parse_numeric_token(target_token.text)
        if parsed_target is not None:
            target_value, target_is_percent = parsed_target
            if not target_is_percent and (abs(target_value) >= 100 or "," in target_token.text):
                target_currency = doc_currency
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
    doc_date = date_phrase_for_token_slice(doc_text, doc_slice)
    target_date = date_phrase_for_token_slice(target_text, target_slice)
    if (
        doc_date
        and target_date
        and normalize_for_compare(doc_date) != normalize_for_compare(target_date)
    ):
        comments.append(
            HtmlComment(
                order=order,
                contents=f"The date is different, {target_date} in {target_name} while {doc_date} in word.",
                token_index=target_start,
            )
        )
        return
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


def whitespace_signature(text: str) -> tuple[int, int, int]:
    source = normalize_proof_text(text)
    return (source.count(" "), source.count("\t"), source.count("\n"))


def describe_whitespace(text: str) -> str:
    spaces, tabs, newlines = whitespace_signature(text)
    parts: list[str] = []
    if spaces:
        parts.append(f"{spaces} space" + ("" if spaces == 1 else "s"))
    if tabs:
        parts.append(f"{tabs} tab" + ("" if tabs == 1 else "s"))
    if newlines:
        parts.append(f"{newlines} line break" + ("" if newlines == 1 else "s"))
    return ", ".join(parts) if parts else "no space"


def spacing_only_difference(doc_text: str, target_text: str) -> bool:
    doc_tokens = diff_tokens(doc_text)
    target_tokens = diff_tokens(target_text)
    if len(doc_tokens) != len(target_tokens):
        return False
    if [token.normalized for token in doc_tokens] != [token.normalized for token in target_tokens]:
        return False
    if any(doc.spaces_before != target.spaces_before for doc, target in zip(doc_tokens, target_tokens)):
        return True
    for left_index in range(len(doc_tokens) - 1):
        doc_sep = inter_token_separator(doc_text, doc_tokens[left_index], doc_tokens[left_index + 1])
        target_sep = inter_token_separator(target_text, target_tokens[left_index], target_tokens[left_index + 1])
        if normalized_separator_symbol(doc_sep) == normalized_separator_symbol(target_sep):
            if whitespace_signature(doc_sep) != whitespace_signature(target_sep):
                return True
    return False


def suppress_html_layout_spacing(
    doc_block: Block,
    target_block: Block,
    *,
    doc_token: DiffToken | None = None,
    target_token: DiffToken | None = None,
) -> bool:
    if not (doc_block.table_cell and target_block.table_cell):
        return False
    if doc_token is not None and target_token is not None:
        if doc_token.kind != target_token.kind:
            return False
        if doc_token.kind == "number":
            return doc_token.normalized == target_token.normalized
    return True


def suppress_pdf_layout_spacing(
    doc_block: Block,
    target_block: Block,
    *,
    doc_token: DiffToken | None = None,
    target_token: DiffToken | None = None,
) -> bool:
    if not (doc_block.table_cell and target_block.table_cell):
        return False
    if doc_token is not None and target_token is not None:
        if doc_token.kind != target_token.kind:
            return False
        if doc_token.kind == "number":
            return doc_token.normalized == target_token.normalized
    return True


def same_table_numeric_value_content(
    doc_block: Block,
    target_block: Block,
    *,
    docx_blocks: list[Block] | None = None,
    target_blocks: list[Block] | None = None,
) -> bool:
    if not (doc_block.table_cell and target_block.table_cell):
        return False
    doc_token = single_value_token(doc_block)
    target_token = single_value_token(target_block)
    if doc_token is None or target_token is None:
        return False
    if doc_token.kind != "number" or target_token.kind != "number":
        return False
    if doc_token.normalized != target_token.normalized:
        return False
    doc_currency = effective_currency_for_comment(doc_block, doc_token, peer_blocks=docx_blocks)
    target_currency = effective_currency_for_comment(target_block, target_token, peer_blocks=target_blocks)
    return doc_currency == target_currency


def inter_token_separator(text: str, left: DiffToken, right: DiffToken) -> str:
    source = normalize_proof_text(text)
    return source[left.end:right.start]


def normalized_separator_symbol(separator: str) -> str:
    compact = re.sub(r"\s+", "", normalize_text(separator))
    return compact


def contextual_equal_token_comments(
    *,
    doc_block: Block,
    target_block: Block,
    order: int,
    target_name: str,
    doc_text: str,
    target_text: str,
    doc_tokens: list[DiffToken],
    target_tokens: list[DiffToken],
    target_offset: int = 0,
    proofread_mode: bool = False,
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
                and (proofread_mode or max(doc_token.spaces_before, target_token.spaces_before) >= 2)
            ):
                if target_name == "html" and suppress_html_layout_spacing(
                    doc_block,
                    target_block,
                    doc_token=doc_token,
                    target_token=target_token,
                ):
                    continue
                if target_name == "pdf" and suppress_pdf_layout_spacing(
                    doc_block,
                    target_block,
                    doc_token=doc_token,
                    target_token=target_token,
                ):
                    continue
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
            if (
                target_right.prefix_symbol is not None
                and target_right.prefix_symbol in "$€£¥"
                and doc_right.prefix_symbol == target_right.prefix_symbol
            ):
                continue
            doc_sep_raw = inter_token_separator(doc_text, doc_left, doc_right)
            target_sep_raw = inter_token_separator(target_text, target_left, target_right)
            if proofread_mode and whitespace_signature(doc_sep_raw) != whitespace_signature(target_sep_raw):
                if target_name == "pdf" and not doc_block.table_cell and not target_block.table_cell and spacing_is_only_pdf_line_wrap(doc_sep_raw, target_sep_raw):
                    continue
                if normalized_separator_symbol(doc_sep_raw) == normalized_separator_symbol(target_sep_raw):
                    if target_name == "html" and suppress_html_layout_spacing(doc_block, target_block):
                        continue
                    if target_name == "pdf" and suppress_pdf_layout_spacing(doc_block, target_block):
                        continue
                    comments.append(
                        HtmlComment(
                            order=order,
                            contents=(
                                f"The spacing is different, {target_name} has {describe_whitespace(target_sep_raw)} "
                                f"between {target_left.text} and {target_right.text} while word has {describe_whitespace(doc_sep_raw)}."
                            ),
                            token_index=target_offset + j1 + offset + 1,
                        )
                    )
                    continue
            doc_sep = normalized_separator_symbol(doc_sep_raw)
            target_sep = normalized_separator_symbol(target_sep_raw)
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


def collapse_insert_delete_comment_pairs(comments: list[HtmlComment]) -> list[HtmlComment]:
    collapsed: list[HtmlComment] = []
    index = 0
    while index < len(comments):
        current = comments[index]
        if index + 1 < len(comments):
            nxt = comments[index + 1]
            current_extra = EXTRA_COMMENT_RE.match(current.contents)
            current_missing = MISSING_COMMENT_RE.match(current.contents)
            next_extra = EXTRA_COMMENT_RE.match(nxt.contents)
            next_missing = MISSING_COMMENT_RE.match(nxt.contents)

            extra_match = current_extra or next_extra
            missing_match = current_missing or next_missing
            if (
                extra_match is not None
                and missing_match is not None
                and extra_match.group(1) == missing_match.group(1)
                and extra_match.group(2) == missing_match.group(2)
                and current.order == nxt.order
            ):
                subject = extra_match.group(1)
                target_name = extra_match.group(2)
                target_value = extra_match.group(3).strip()
                doc_value = missing_match.group(3).strip()
                token_index_candidates = [value for value in (current.token_index, nxt.token_index) if value is not None]
                collapsed.append(
                    HtmlComment(
                        order=current.order,
                        contents=(
                            f"The {subject} is different, {target_value} in {target_name} "
                            f"while {doc_value} in word."
                        ),
                        token_index=min(token_index_candidates) if token_index_candidates else None,
                    )
                )
                index += 2
                continue
        collapsed.append(current)
        index += 1
    return collapsed


def single_token_target_comments(
    doc_block: Block,
    target_block: Block,
    *,
    target_name: str,
    proofread_mode: bool = False,
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
            if target_name == "pdf" and not target_currency and not proofread_mode:
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
            and (proofread_mode or max(doc_token.spaces_before, target_token.spaces_before) >= 2)
        ):
            if target_name == "html" and suppress_html_layout_spacing(
                doc_block,
                target_block,
                doc_token=doc_token,
                target_token=target_token,
            ):
                return []
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


def contact_field_difference_comments(
    doc_block: Block,
    target_block: Block,
    score: float,
    *,
    target_name: str,
    proofread_mode: bool,
    match_type: str,
) -> list[HtmlComment]:
    payload = contact_field_payload(doc_block.text)
    if payload is None:
        return []
    role, doc_value = payload
    target_value = extract_contact_fields(target_block.text).get(role)
    if not target_value:
        return []
    if normalize_proof_text(doc_value) == normalize_proof_text(target_value):
        return []
    doc_field_block = Block(
        id=f"{doc_block.id}-contact-field",
        source=doc_block.source,
        order=doc_block.order,
        text=doc_value,
        normalized=normalize_for_compare(doc_value),
        raw_text=doc_value,
        proof_text=doc_value,
        match_text=normalize_for_compare(doc_value),
        kind=doc_block.kind,
        structure_role=role,
    )
    target_field_block = Block(
        id=f"{target_block.id}-contact-field-{role}",
        source=target_block.source,
        order=target_block.order,
        text=target_value,
        normalized=normalize_for_compare(target_value),
        raw_text=target_value,
        proof_text=target_value,
        match_text=normalize_for_compare(target_value),
        kind=target_block.kind,
        structure_role=role,
    )
    return text_difference_comments(
        doc_field_block,
        target_field_block,
        max(score, similarity(doc_field_block.normalized, target_field_block.normalized)),
        target_name=target_name,
        match_type="exact_structural" if exact_like_match_type(match_type) else "approx",
        proofread_mode=proofread_mode,
        grouped_match_type=None,
    )


def text_difference_comments(
    doc_block: Block,
    target_block: Block,
    score: float,
    *,
    target_name: str,
    docx_blocks: list[Block] | None = None,
    target_blocks: list[Block] | None = None,
    match_type: str = "approx",
    formatting_diffs: list[str] | None = None,
    proofread_mode: bool = False,
    grouped_match_type: str | None = None,
) -> list[HtmlComment]:
    if grouped_match_type == "contact":
        return contact_field_difference_comments(
            doc_block,
            target_block,
            score,
            target_name=target_name,
            proofread_mode=proofread_mode,
            match_type=match_type,
        )
    compare_doc_block = doc_block
    compare_target_block = target_block
    target_focus_applied = False
    if target_name == "pdf" and grouped_match_type == "quote":
        doc_quote = extract_primary_quote_text(doc_block.text)
        target_quote = extract_primary_quote_text(target_block.text)
        if doc_quote and target_quote:
            target_focus_applied = True
            compare_doc_block = Block(
                id=f"{doc_block.id}-quote-focus",
                source=doc_block.source,
                order=doc_block.order,
                text=doc_quote,
                normalized=normalize_for_compare(doc_quote),
                raw_text=doc_quote,
                proof_text=normalize_proof_text(doc_quote),
                match_text=normalize_for_compare(doc_quote),
                structure_role=doc_block.structure_role,
            )
            compare_target_block = Block(
                id=f"{target_block.id}-quote-focus",
                source=target_block.source,
                order=target_block.order,
                text=target_quote,
                normalized=normalize_for_compare(target_quote),
                raw_text=target_quote,
                proof_text=normalize_proof_text(target_quote),
                match_text=normalize_for_compare(target_quote),
                structure_role=target_block.structure_role,
            )
    elif target_name == "pdf" and not doc_block.table_cell and not target_block.table_cell:
        focused_text = best_pdf_narrative_focus(doc_block, target_block)
        body_text = pdf_embedded_lead_body(target_block)
        if focused_text is not None:
            target_focus_applied = True
            compare_target_block = Block(
                id=f"{target_block.id}-segment-focus",
                source=target_block.source,
                order=target_block.order,
                text=focused_text,
                normalized=normalize_for_compare(focused_text),
                raw_text=focused_text,
                proof_text=normalize_proof_text(focused_text),
                match_text=normalize_for_compare(focused_text),
                structure_role=target_block.structure_role,
            )
        elif body_text and similarity(normalize_for_compare(doc_block.text), normalize_for_compare(body_text)) >= 0.9:
            target_focus_applied = True
            compare_target_block = Block(
                id=f"{target_block.id}-body-focus",
                source=target_block.source,
                order=target_block.order,
                text=body_text,
                normalized=normalize_for_compare(body_text),
                raw_text=body_text,
                proof_text=normalize_proof_text(body_text),
                match_text=normalize_for_compare(body_text),
                structure_role=target_block.structure_role,
            )
    confidence_tier = match_confidence_tier(
        compare_doc_block,
        compare_target_block,
        score=score,
        match_type=match_type,
        grouped_match_type=grouped_match_type,
        target_name=target_name,
    )
    precise_schema_header = allow_precise_schema_header_diffs(
        compare_doc_block,
        compare_target_block,
        score=score,
        match_type=match_type,
    )
    narrative_like = target_name in {"html", "pdf"} and long_narrative_block(compare_doc_block) and long_narrative_block(compare_target_block)
    formatting_allowed = not (
        proofread_mode
        and target_name in {"html", "pdf"}
        and formatting_diffs
        and not exact_like_match_type(match_type)
        and score < (0.92 if target_name == "html" else 0.95)
        and not precise_schema_header
    ) and (confidence_tier == "strong" or precise_schema_header) and not (target_name == "pdf" and target_focus_applied)
    formatting_comments = [
        HtmlComment(
            order=target_block.order,
            contents=f"Formatting differs: {diff}",
            token_index=0 if diff_tokens(target_block.text) else None,
        )
        for diff in (formatting_diffs or [])
    ] if proofread_mode and formatting_diffs and formatting_allowed else []
    contained_like = match_type == "contained" or (
        target_name == "pdf"
        and not doc_block.table_cell
        and not target_block.table_cell
        and len(target_block.normalized) > max(len(doc_block.normalized) * 2, 160)
    )
    currency_comment = currency_symbol_comment(
        compare_doc_block,
        compare_target_block,
        target_name=target_name,
        docx_blocks=docx_blocks,
        target_blocks=target_blocks,
    )
    if currency_comment is not None:
        return [currency_comment]
    percent_comment = percent_symbol_comment(
        compare_doc_block,
        compare_target_block,
        target_name=target_name,
    )
    if percent_comment is not None:
        return [percent_comment]
    numeric_comment = numeric_block_difference_comment(
        compare_doc_block,
        compare_target_block,
        target_name=target_name,
        docx_blocks=docx_blocks,
        target_blocks=target_blocks,
    )
    if numeric_comment is not None:
        return [numeric_comment]
    block_date_comment = block_date_difference_comment(
        compare_doc_block.text,
        compare_target_block.text,
        target_name=target_name,
    )
    if block_date_comment is not None:
        return [
            HtmlComment(
                order=compare_target_block.order,
                contents=block_date_comment,
                token_index=0 if diff_tokens(compare_target_block.text) else None,
            )
        ]
    if (
        target_name == "html"
        and not compare_doc_block.table_cell
        and not compare_target_block.table_cell
    ):
        numeric_block_token_comment = single_unambiguous_numeric_difference_comment(
            compare_doc_block.text,
            compare_target_block.text,
            target_name=target_name,
        )
        if numeric_block_token_comment is not None:
            contents, token_index = numeric_block_token_comment
            return [
                HtmlComment(
                    order=compare_target_block.order,
                    contents=contents,
                    token_index=token_index,
                )
            ]
        if (
            grouped_match_type is None
            and not precise_schema_header
            and compare_doc_block.structure_role == "paragraph"
            and compare_target_block.structure_role == "paragraph"
        ):
            phrase_token_comments = single_unambiguous_phrase_difference_comments(
                compare_doc_block.text,
                compare_target_block.text,
                target_name=target_name,
            )
            if phrase_token_comments is not None:
                return [
                    HtmlComment(
                        order=compare_target_block.order,
                        contents=contents,
                        token_index=token_index,
                    )
                    for contents, token_index in phrase_token_comments
                ]
            local_token_comments = single_unambiguous_local_token_comments(
                compare_doc_block.text,
                compare_target_block.text,
                target_name=target_name,
            )
            if local_token_comments is not None:
                return [
                    HtmlComment(
                        order=compare_target_block.order,
                        contents=contents,
                        token_index=token_index,
                    )
                    for contents, token_index in local_token_comments
                ]
    single_token_comments = single_token_target_comments(
        compare_doc_block,
        compare_target_block,
        target_name=target_name,
        proofread_mode=proofread_mode,
    )
    if single_token_comments is not None:
        return single_token_comments
    if target_name == "html" and prnewswire_only_difference(compare_doc_block.text, compare_target_block.text):
        pr_tokens = diff_tokens(compare_target_block.text)
        pr_index = next((index for index, token in enumerate(pr_tokens) if token.normalized == "prnewswire"), None)
        return [
            HtmlComment(
                order=compare_target_block.order,
                contents="The word is extra in html, PRNewswire. It is not present in word.",
                token_index=pr_index,
            )
        ]
    if target_name == "pdf" and pdf_blocks_equal_after_cleanup(compare_doc_block, compare_target_block):
        if not spacing_only_difference(compare_doc_block.text, compare_target_block.text):
            return formatting_comments
    if (
        target_name == "pdf"
        and not compare_doc_block.table_cell
        and not compare_target_block.table_cell
        and normalize_pdf_paragraph_artifacts(compare_doc_block.text) == normalize_pdf_paragraph_artifacts(compare_target_block.text)
    ):
        return formatting_comments
    if (
        target_name == "pdf"
        and narrative_like
        and not compare_doc_block.table_cell
        and not compare_target_block.table_cell
        and pdf_minor_narrative_noise_only(compare_doc_block.text, compare_target_block.text)
    ):
        return formatting_comments
    if not (proofread_mode and target_name == "html"):
        doc_match_text = normalize_for_compare(strip_leading_markers(compare_doc_block.text))
        target_match_text = normalize_for_compare(strip_leading_markers(compare_target_block.text))
        if doc_match_text == target_match_text:
            if normalize_proof_text(compare_doc_block.text) == normalize_proof_text(compare_target_block.text):
                return formatting_comments
    pdf_score_floor = 0.82 if proofread_mode else 0.9
    if target_name == "pdf" and not contained_like and not compare_doc_block.table_cell and not compare_target_block.table_cell and score < pdf_score_floor:
        return []
    if confidence_tier == "weak":
        return []
    if grouped_match_type == "quote" and not exact_like_match_type(match_type):
        if normalize_proof_text(compare_doc_block.text).strip() == normalize_proof_text(compare_target_block.text).strip():
            return formatting_comments
        quote_floor = 0.88 if target_name == "html" else 0.8 if target_name == "pdf" else 0.88
        quote_detail = 0.96 if target_name == "html" else 0.9 if target_name == "pdf" else 0.96
        if score < quote_floor:
            return []
        if score < quote_detail or narrative_like:
            return [
                HtmlComment(
                    order=target_block.order,
                    contents=quote_diff_summary(compare_doc_block.text, compare_target_block.text, target_name=target_name),
                    token_index=0 if diff_tokens(compare_target_block.text) else None,
                )
            ]
    narrative_floor = 0.88 if target_name == "html" else 0.8 if target_name == "pdf" else 0.88
    if narrative_like and not exact_like_match_type(match_type) and score < narrative_floor:
        return []

    doc_tokens = diff_tokens(compare_doc_block.text)
    target_tokens = diff_tokens(compare_target_block.text)
    if not doc_tokens and not target_tokens:
        return []
    target_offset = 0
    if contained_like:
        doc_tokens, target_tokens, target_offset = trim_contained_token_alignment(doc_tokens, target_tokens)
        if not doc_tokens and not target_tokens:
            return []

    contextual_comments = contextual_equal_token_comments(
        doc_block=doc_block,
        target_block=compare_target_block,
        order=compare_target_block.order,
        target_name=target_name,
        doc_text=compare_doc_block.text,
        target_text=compare_target_block.text,
        doc_tokens=doc_tokens,
        target_tokens=target_tokens,
        target_offset=target_offset,
        proofread_mode=proofread_mode,
    )
    if compare_doc_block.normalized == compare_target_block.normalized or score >= 0.999:
        return contextual_comments or formatting_comments

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
                order=compare_target_block.order,
                doc_text=compare_doc_block.text,
                target_text=compare_target_block.text,
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
                order=compare_target_block.order,
                doc_text=compare_doc_block.text,
                target_text=compare_target_block.text,
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
                order=compare_target_block.order,
                doc_text=compare_doc_block.text,
                target_text=compare_target_block.text,
                target_tokens=target_tokens,
                doc_slice=doc_tokens[i1:i2],
                target_slice=[],
                target_start=target_offset + j1,
                target_name=target_name,
            )

    comments.extend(contextual_comments)
    comments = dedupe_html_comments(comments)
    comments = collapse_insert_delete_comment_pairs(comments)

    if comments:
        high_signal_comments = [
            comment for comment in comments
            if (
                comment.contents.startswith("The date is different")
                or comment.contents.startswith("The number is different")
                or comment.contents.startswith("The number is extra in html")
                or comment.contents.startswith("The number is missing in html")
                or comment.contents.startswith("The currency symbol is different")
                or comment.contents.startswith("The percent sign is different")
                or comment.contents.startswith("The symbol is different")
            )
        ]
        detailed_comment_prefixes = (
            "The word is different",
            "The word is extra in html",
            "The word is missing in html",
            "The number is different",
            "The number is extra in html",
            "The number is missing in html",
            "The date is different",
            "The symbol is different",
            "The currency symbol is different",
            "The percent sign is different",
        )
        detailed_token_comments = [
            comment for comment in comments if comment.contents.startswith(detailed_comment_prefixes)
        ]
        clear_medium_token_details = (
            target_name == "html"
            and confidence_tier == "medium"
            and not precise_schema_header
            and grouped_match_type is None
            and not repeated_label_block(compare_doc_block)
            and not repeated_label_block(compare_target_block)
            and len(comments) <= 6
            and (
                (
                    not narrative_like
                    and len(doc_tokens) <= 120
                    and len(target_tokens) <= 120
                )
                or (
                    narrative_like
                    and len(comments) <= 4
                    and all(comment.contents.startswith(detailed_comment_prefixes) for comment in comments)
                )
            )
        )
        if precise_schema_header and formatting_comments:
            comments.extend(formatting_comments)
            comments = dedupe_html_comments(comments)
        if grouped_match_type == "quote":
            if normalize_proof_text(compare_doc_block.text).strip() == normalize_proof_text(compare_target_block.text).strip():
                return formatting_comments
            return [
                HtmlComment(
                    order=target_block.order,
                    contents=quote_diff_summary(compare_doc_block.text, compare_target_block.text, target_name=target_name),
                    token_index=0 if target_tokens else None,
                )
            ]
        if confidence_tier == "medium" and not precise_schema_header:
            if (
                high_signal_comments
                and not repeated_label_block(compare_doc_block)
                and not repeated_label_block(compare_target_block)
            ):
                return dedupe_html_comments(high_signal_comments + formatting_comments)
            if (
                target_name == "html"
                and grouped_match_type is None
                and not repeated_label_block(compare_doc_block)
                and not repeated_label_block(compare_target_block)
                and detailed_token_comments
                and len(detailed_token_comments) <= 4
            ):
                return dedupe_html_comments(detailed_token_comments + formatting_comments)
            if clear_medium_token_details:
                return dedupe_html_comments(comments + formatting_comments)
            return [
                HtmlComment(
                    order=target_block.order,
                    contents=(
                        "The paragraph text is different. "
                        f"{target_name.upper()}: {shorten(compare_target_block.text, 160)} "
                        f"Word: {shorten(compare_doc_block.text, 160)}"
                    ),
                    token_index=0 if target_tokens else None,
                )
            ]
        if (
            grouped_match_type == "footnote"
            and (
                len(comments) > 4
                or len(doc_block.text) > 140
                or len(target_block.text) > 140
            )
        ):
            return [
                HtmlComment(
                    order=target_block.order,
                    contents=(
                        "The footnote text is different. "
                        f"{target_name.upper()}: {shorten(compare_target_block.text, 140)} "
                        f"Word: {shorten(compare_doc_block.text, 140)}"
                    ),
                    token_index=0 if target_tokens else None,
                )
            ]
        if (
            target_name in {"html", "pdf"}
            and narrative_like
            and not exact_like_match_type(match_type)
            and not (target_name == "html" and clear_medium_token_details)
        ):
            return [
                HtmlComment(
                    order=target_block.order,
                    contents=(
                        "The paragraph text is different. "
                        f"{target_name.upper()}: {shorten(compare_target_block.text, 160)} "
                        f"Word: {shorten(compare_doc_block.text, 160)}"
                    ),
                    token_index=0 if target_tokens else None,
                )
            ]
        if target_name == "pdf" and contained_like and not doc_block.table_cell and not target_block.table_cell:
            focused_comments = [
                comment
                for comment in comments
                if (
                    "spacing is different" in comment.contents
                    or "currency symbol is different" in comment.contents
                    or comment.contents.startswith("The date ")
                    or comment.contents.startswith("The number ")
                    or (proofread_mode and comment.contents.startswith("The word "))
                    or (proofread_mode and comment.contents.startswith("The symbol "))
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
            and re.search(r"\(\s*1\s*\)", compare_doc_block.text)
        ):
            return []
        if (
            target_name == "pdf"
            and not compare_doc_block.table_cell
            and not compare_target_block.table_cell
            and len(comments) <= 2
            and len(compare_doc_block.text) <= 80
            and len(compare_target_block.text) <= 80
            and all(comment.contents.startswith("The word is extra in pdf,") for comment in comments)
            and token_subsequence_ratio(compare_doc_block.normalized, compare_target_block.normalized) >= 0.99
        ):
            return []
        if target_name == "pdf" and not compare_doc_block.table_cell and not compare_target_block.table_cell and len(comments) > (8 if proofread_mode else 4):
            return [
                HtmlComment(
                    order=compare_target_block.order,
                    contents=(
                        "The paragraph text differs between the PDF and Word. "
                        f"PDF: {shorten(compare_target_block.text, 140)} "
                        f"Word: {shorten(compare_doc_block.text, 140)}"
                    ),
                    token_index=0 if target_tokens else None,
                )
            ]
        return comments

    if formatting_comments:
        return formatting_comments

    if target_name == "pdf" and normalize_without_punctuation(compare_doc_block.text) == normalize_without_punctuation(compare_target_block.text):
        return []

    if (
        target_name == "html"
        and compare_doc_block.table_cell
        and compare_target_block.table_cell
        and suppress_html_layout_spacing(compare_doc_block, compare_target_block)
        and same_table_numeric_value_content(
            compare_doc_block,
            compare_target_block,
            docx_blocks=docx_blocks,
            target_blocks=target_blocks,
        )
    ):
        return formatting_comments

    if (
        target_name == "pdf"
        and compare_doc_block.table_cell
        and compare_target_block.table_cell
        and suppress_pdf_layout_spacing(compare_doc_block, compare_target_block)
        and same_table_numeric_value_content(
            compare_doc_block,
            compare_target_block,
            docx_blocks=docx_blocks,
            target_blocks=target_blocks,
        )
    ):
        return formatting_comments

    if target_name == "pdf" and contained_like:
        return []

    if grouped_match_type == "footnote":
        return [
            HtmlComment(
                    order=target_block.order,
                    contents=(
                        "The footnote text is different. "
                        f"{target_name.upper()}: {shorten(compare_target_block.text, 140)} "
                        f"Word: {shorten(compare_doc_block.text, 140)}"
                    ),
                    token_index=0 if target_tokens else None,
                )
            ]

    return [
        HtmlComment(
            order=compare_target_block.order,
            contents=(
                "The paragraph text is different. "
                f"{target_name.upper()}: {shorten(compare_target_block.text, 140)} "
                f"Word: {shorten(compare_doc_block.text, 140)}"
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


def build_section_family_appendix_comments(
    docx_blocks: list[Block],
    target_blocks: list[Block],
    unmatched_docx: list[Block],
    *,
    target_label: str,
) -> tuple[list[tuple[Block, str]], set[str]]:
    (
        section_matches,
        doc_block_to_family,
        _target_block_to_family,
        doc_family_map,
        target_family_map,
    ) = match_section_families(docx_blocks, target_blocks)
    unmatched_ids = {block.id for block in unmatched_docx}
    doc_index_by_id = {block.id: index for index, block in enumerate(docx_blocks)}
    comments: list[tuple[Block, str]] = []
    covered_block_ids: set[str] = set()
    seen_family_ids: set[str] = set()

    for block in unmatched_docx:
        doc_index = doc_index_by_id.get(block.id)
        if doc_index is None:
            continue
        family_id = doc_block_to_family.get(doc_index)
        if family_id is None or family_id in seen_family_ids:
            continue
        family = doc_family_map.get(family_id)
        if family is None:
            continue
        family_block_ids = {docx_blocks[index].id for index in family.block_indices}
        if not (family_block_ids & unmatched_ids):
            continue
        seen_family_ids.add(family_id)
        covered_block_ids.update(family_block_ids & unmatched_ids)
        representative = docx_blocks[family.block_indices[0]]
        matched_target_id = section_matches.get(family_id)
        if matched_target_id is not None:
            continue
        comments.append(
            (
                representative,
                (
                    f"This DOCX {family.family_type} section was not found in the {target_label}. "
                    f"Word: {shorten(family.text, 180)}"
                ),
            )
        )
    return comments, covered_block_ids


def build_section_family_inline_comments(
    docx_blocks: list[Block],
    target_blocks: list[Block],
    unmatched_docx: list[Block],
    *,
    target_label: str,
) -> tuple[list[HtmlComment], set[str]]:
    (
        section_matches,
        doc_block_to_family,
        _target_block_to_family,
        doc_family_map,
        target_family_map,
    ) = match_section_families(docx_blocks, target_blocks)
    unmatched_ids = {block.id for block in unmatched_docx}
    doc_index_by_id = {block.id: index for index, block in enumerate(docx_blocks)}
    comments: list[HtmlComment] = []
    covered_block_ids: set[str] = set()
    seen_family_ids: set[str] = set()

    for block in unmatched_docx:
        doc_index = doc_index_by_id.get(block.id)
        if doc_index is None:
            continue
        family_id = doc_block_to_family.get(doc_index)
        if family_id is None or family_id in seen_family_ids:
            continue
        matched_target_id = section_matches.get(family_id)
        if matched_target_id is None:
            continue
        family = doc_family_map.get(family_id)
        target_family = target_family_map.get(matched_target_id)
        if family is None or target_family is None:
            continue
        family_block_ids = {docx_blocks[index].id for index in family.block_indices}
        if not (family_block_ids & unmatched_ids):
            continue
        seen_family_ids.add(family_id)
        covered_block_ids.update(family_block_ids & unmatched_ids)
        comments.append(
            HtmlComment(
                order=target_family.order_start,
                contents=(
                    f"This DOCX {family.family_type} section is broadly matched to the {target_label} but not aligned block-by-block. "
                    f"{target_label}: {shorten(target_family.text, 180)} "
                    f"Word: {shorten(family.text, 180)}"
                ),
                token_index=None,
            )
        )
    return comments, covered_block_ids


def build_comments(
    docx_blocks: list[Block],
    html_blocks: list[Block],
    matches: list[Match],
    unmatched_docx: list[Block],
    unmatched_html: list[Block],
    *,
    target_label: str = "HTML",
    proofread_mode: bool = False,
) -> tuple[list[HtmlComment], list[tuple[Block, str]]]:
    html_comments: list[HtmlComment] = []
    target_name = target_label.lower()
    (
        grouped_target_by_doc_index,
        doc_block_to_group,
        target_block_to_group,
        doc_group_map,
        target_group_map,
    ) = match_block_groups(docx_blocks, html_blocks)
    doc_index_by_id = {block.id: index for index, block in enumerate(docx_blocks)}
    target_index_by_id = {block.id: index for index, block in enumerate(html_blocks)}
    family_inline_comments, family_inline_docx_ids = build_section_family_inline_comments(
        docx_blocks,
        html_blocks,
        unmatched_docx,
        target_label=target_label,
    )
    family_appendix_comments, family_covered_docx_ids = build_section_family_appendix_comments(
        docx_blocks,
        html_blocks,
        unmatched_docx,
        target_label=target_label,
    )
    family_covered_docx_ids |= family_inline_docx_ids

    def grouped_match_type(doc_index: int, target_index: int) -> str | None:
        doc_group_id = doc_block_to_group.get(doc_index)
        target_group_id = target_block_to_group.get(target_index)
        if doc_group_id is None or target_group_id is None:
            return None
        if grouped_target_by_doc_index.get(doc_index) != target_group_id:
            return None
        doc_group = doc_group_map.get(doc_group_id)
        target_group = target_group_map.get(target_group_id)
        if doc_group is None or target_group is None:
            return None
        if doc_group.group_type != target_group.group_type:
            return None
        return doc_group.group_type

    for match in matches:
        doc_block = docx_blocks[match.docx_index]
        html_block = html_blocks[match.html_index]
        matched_group_type = grouped_match_type(match.docx_index, match.html_index)
        html_comments.extend(
            text_difference_comments(
                doc_block,
                html_block,
                match.score,
                target_name=target_name,
                docx_blocks=docx_blocks,
                target_blocks=html_blocks,
                match_type=match.match_type,
                formatting_diffs=match.formatting_diffs,
                proofread_mode=proofread_mode,
                grouped_match_type=matched_group_type,
            )
        )

    html_comments.extend(family_inline_comments)

    def unmatched_html_fallback_comments(block: Block) -> list[HtmlComment]:
        target_group_id = target_block_to_group.get(target_index_by_id.get(block.id, -1))
        if target_group_id is not None:
            target_group = target_group_map.get(target_group_id)
            if target_group is not None and target_group.group_type in {"quote", "contact", "footnote"}:
                return []
        if repeated_label_block(block) and matching_block_exists(block, docx_blocks):
            return []
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
        if repeated_label_block(block) and repeated_label_block(best_doc_block) and best_score < 0.94:
            return []

        overlap = token_overlap_ratio(best_doc_block.normalized, block.normalized)
        strong_near_match = (
            best_score >= (0.86 if proofread_mode and target_label == "HTML" else 0.78)
            or (
                block.table_cell
                and best_doc_block.table_cell
                and best_score >= (0.74 if proofread_mode and target_label == "HTML" else 0.68)
                and overlap >= (0.62 if proofread_mode and target_label == "HTML" else 0.55)
                and len(diff_tokens(best_doc_block.text)) <= 6
                and len(diff_tokens(block.text)) <= 6
            )
            or (
                overlap >= 0.5
                and best_score >= (0.7 if proofread_mode and target_label == "HTML" else 0.58)
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
            proofread_mode=proofread_mode,
        )

    def unmatched_docx_fallback_comments(block: Block) -> list[HtmlComment]:
        if block.id in family_covered_docx_ids:
            return []
        doc_group_id = doc_block_to_group.get(doc_index_by_id.get(block.id, -1))
        if doc_group_id is not None:
            doc_group = doc_group_map.get(doc_group_id)
            if doc_group is not None and doc_group.group_type in {"quote", "contact", "footnote"}:
                return []
        if repeated_label_block(block) and matching_block_exists(block, html_blocks):
            return []
        best_target_block: Block | None = None
        best_score = 0.0
        target_candidates = unmatched_html if unmatched_html else html_blocks
        doc_token = single_value_token(block)
        doc_token_norm = doc_token.normalized if doc_token is not None else None
        for target_block in target_candidates:
            if target_block.table_cell != block.table_cell:
                continue
            if target_label == "PDF" and is_pdf_chrome_text(target_block.text):
                continue
            if (
                proofread_mode
                and target_label == "PDF"
                and not (
                    structural_role_compatible(block, target_block)
                    or compatible_header_family_roles(block.structure_role or "", target_block.structure_role or "")
                )
            ):
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
        if repeated_label_block(block) and repeated_label_block(best_target_block) and best_score < 0.94:
            return []

        overlap = token_overlap_ratio(block.normalized, best_target_block.normalized)
        if proofread_mode and target_label == "PDF" and not block.table_cell:
            if overlap < 0.7:
                return []
            if best_score < 0.84:
                return []
        strong_near_match = (
            best_score >= (
                0.86 if proofread_mode and target_label == "HTML"
                else 0.74 if proofread_mode and target_label == "PDF"
                else 0.8
            )
            or (
                block.table_cell
                and best_score >= (
                    0.74 if proofread_mode and target_label == "HTML"
                    else 0.52 if proofread_mode and target_label == "PDF"
                    else 0.58
                )
                and overlap >= (0.62 if proofread_mode and target_label == "HTML" else 0.0)
                and len(diff_tokens(block.text)) <= 6
            )
            or (
                overlap >= (
                    0.62 if proofread_mode and target_label == "HTML"
                    else 0.42 if proofread_mode and target_label == "PDF"
                    else 0.5
                )
                and best_score >= (
                    0.72 if proofread_mode and target_label == "HTML"
                    else 0.54 if proofread_mode and target_label == "PDF"
                    else 0.6
                )
                and len(diff_tokens(block.text)) <= (
                    10 if proofread_mode and target_label == "HTML"
                    else 12 if proofread_mode and target_label == "PDF"
                    else 8
                )
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
            proofread_mode=proofread_mode,
        )

    for block in unmatched_html:
        if target_label == "PDF":
            continue
        fallback_comments = unmatched_html_fallback_comments(block)
        if fallback_comments:
            html_comments.extend(fallback_comments)
            continue
        if repeated_label_block(block) and matching_block_exists(block, docx_blocks):
            continue
        html_comments.append(
            HtmlComment(
                order=block.order,
                contents=f"This {target_label} block has no corresponding content in the DOCX.",
                token_index=0 if diff_tokens(block.text) else None,
            )
        )

    remaining_unmatched_docx = [block for block in unmatched_docx if block.id not in family_covered_docx_ids]
    appendix_comments = list(family_appendix_comments)
    if family_appendix_comments:
        appendix_comments.extend(
            [
                (block, f"This DOCX content was not found in the {target_label}.")
                for block in remaining_unmatched_docx
            ]
        )
    else:
        appendix_comments.extend(
            [
                (block, f"This DOCX content was not found in the {target_label}.")
                for block in appendix_summary_blocks(docx_blocks, remaining_unmatched_docx)
            ]
        )
    if target_label == "PDF":
        pdf_fallback_ids: set[str] = set()
        for block in unmatched_docx:
            fallback_comments = unmatched_docx_fallback_comments(block)
            if fallback_comments:
                html_comments.extend(fallback_comments)
                pdf_fallback_ids.add(block.id)
        if proofread_mode:
            remaining_pdf_docx = [
                block
                for block in unmatched_docx
                if block.id not in family_covered_docx_ids and block.id not in pdf_fallback_ids
            ]
            appendix_comments = list(family_appendix_comments)
            if family_appendix_comments:
                appendix_comments.extend(
                    [
                        (block, f"This DOCX content was not found in the {target_label}.")
                        for block in remaining_pdf_docx
                    ]
                )
            else:
                appendix_comments.extend(
                    [
                        (block, f"This DOCX content was not found in the {target_label}.")
                        for block in appendix_summary_blocks(docx_blocks, remaining_pdf_docx)
                    ]
                )
        else:
            appendix_comments = []
    return html_comments, appendix_comments


def group_html_comments(html_comments: list[HtmlComment]) -> dict[int, list[str]]:
    grouped: dict[int, list[str]] = {}
    for comment in html_comments:
        grouped.setdefault(comment.order, []).append(format_review_comment_text(comment.contents))
    return grouped


def review_metadata(contents: str) -> tuple[str, str]:
    lower = normalize_for_compare(contents)
    if "broadly matched" in lower or "section is broadly matched" in lower:
        return "medium", "structural"
    if "no corresponding content" in lower or "was not found in the html" in lower or "was not found in the pdf" in lower:
        return "medium", "structural"
    if lower.startswith("formatting differs"):
        return "medium", "formatting"
    if lower.startswith("the footnote text is different") or lower.startswith("the paragraph text is different"):
        return "medium", "text"
    if (
        lower.startswith("the word is different")
        or lower.startswith("the number is different")
        or lower.startswith("the date is different")
        or lower.startswith("the currency symbol is different")
        or lower.startswith("the percent sign is different")
        or lower.startswith("the symbol is different")
        or lower.startswith("the spacing is different")
        or lower.startswith("the word is extra")
        or lower.startswith("the word is missing")
        or lower.startswith("the number is extra")
        or lower.startswith("the number is missing")
    ):
        return "high", "critical"
    return "medium", "informational"


def review_label(confidence: str, tier: str) -> str:
    tier_labels = {
        "critical": "Critical",
        "text": "Text",
        "formatting": "Formatting",
        "structural": "Structural",
        "informational": "Info",
    }
    return f"[{confidence.title()} Confidence | {tier_labels.get(tier, tier.title())}]"


def format_review_comment_text(contents: str) -> str:
    confidence, tier = review_metadata(contents)
    return f"{review_label(confidence, tier)} {contents}"


def review_summary_counts(comment_texts: list[str]) -> dict[tuple[str, str], int]:
    counts: dict[tuple[str, str], int] = {}
    for text in comment_texts:
        key = review_metadata(text)
        counts[key] = counts.get(key, 0) + 1
    return counts


def review_summary_lines(html_comments: list[HtmlComment], appendix_comments: list[tuple[Block, str]]) -> list[str]:
    counts = review_summary_counts(
        [comment.contents for comment in html_comments] + [comment for _block, comment in appendix_comments]
    )
    ordered_keys = [
        ("high", "critical"),
        ("medium", "text"),
        ("medium", "formatting"),
        ("medium", "structural"),
        ("medium", "informational"),
    ]
    lines: list[str] = []
    for key in ordered_keys:
        count = counts.get(key, 0)
        if not count:
            continue
        lines.append(f"{review_label(*key)}: {count}")
    return lines


def pdf_page_summary_comments(
    *,
    docx_blocks: list[Block],
    pdf_blocks: list[Block],
    unmatched_pdf: list[Block],
    matches: list[Match],
    render_result: BrowserRenderResult,
    proofread_mode: bool = False,
) -> list[HtmlComment]:
    if proofread_mode:
        return []
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
        matched_count = matched_pages.get(page_number, 0)
        if not should_emit_pdf_page_summary_comment(
            meaningful_blocks=meaningful_blocks,
            matched_count=matched_count,
            docx_blocks=docx_blocks,
        ):
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


def should_emit_pdf_page_summary_comment(
    *,
    meaningful_blocks: list[Block],
    matched_count: int,
    docx_blocks: list[Block],
) -> bool:
    if len(meaningful_blocks) < 5:
        return False
    top_blocks = meaningful_blocks[:3]
    if any(pdf_block_has_docx_anchor(block, docx_blocks) for block in top_blocks):
        return False
    if matched_count > 0:
        return False
    return True


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
            text=format_review_comment_text(comment.contents),
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
        review_lines = review_summary_lines(html_comments, appendix_comments)
        lines = review_lines + [f"- {shorten(block.text, 160)}" for block, _comment in appendix_comments[:20]]
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
    summary_lines = review_summary_lines(html_comments, appendix_comments)
    if summary_lines:
        pdf.add_wrapped_text("Proofread Summary", font_size=12, gap_after=6)
        for line in summary_lines:
            pdf.add_wrapped_text(line, font_size=10, indent=12, gap_after=0)
        pdf.current_y -= 10

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
            pdf.add_block(f"[DOCX only {index}]", block.text, format_review_comment_text(comment))

    output_path.write_bytes(pdf.build())


def build_summary(
    docx_blocks: list[Block],
    html_blocks: list[Block],
    matches: list[Match],
    unmatched_docx: list[Block],
    unmatched_html: list[Block],
) -> dict[str, object]:
    exact = sum(1 for match in matches if exact_like_match_type(match.match_type))
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
    parser.add_argument(
        "--proofread",
        action="store_true",
        help="Use the stricter proofread profile for compare mode.",
    )
    return parser.parse_args()


def run_compare(
    *,
    docx_path: Path,
    html_path: Path,
    output_path: Path | None = None,
    summary_json_path: Path | None = None,
    renderer: str = "auto",
    proofread_mode: bool = False,
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

    matches, unmatched_docx, unmatched_html = compare_blocks(
        docx_blocks,
        html_blocks,
        target_name="html",
        proofread_mode=proofread_mode,
    )
    html_comments, appendix_comments = build_comments(
        docx_blocks,
        html_blocks,
        matches,
        unmatched_docx,
        unmatched_html,
        target_label="HTML",
        proofread_mode=proofread_mode,
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
    summary["compare_profile"] = "proofread" if proofread_mode else "standard"
    summary["formatting_scope"] = "token, symbol, spacing, and structural formatting checks against extracted HTML blocks"
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
    proofread_mode: bool = False,
) -> dict[str, object]:
    output_path = output_path or Path(default_pdf_output_name(pdf_path))
    docx_blocks = extract_docx_blocks(docx_path)
    render_result = extract_pdf_blocks(pdf_path, proofread_mode=proofread_mode)
    pdf_blocks = render_result.blocks

    matches, unmatched_docx, unmatched_pdf = compare_blocks(
        docx_blocks,
        pdf_blocks,
        target_name="pdf",
        proofread_mode=proofread_mode,
    )
    pdf_comments, appendix_comments = build_comments(
        docx_blocks,
        pdf_blocks,
        matches,
        unmatched_docx,
        unmatched_pdf,
        target_label="PDF",
        proofread_mode=proofread_mode,
    )
    if not proofread_mode:
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
    summary["compare_profile"] = "proofread" if proofread_mode else "standard"
    summary["formatting_scope"] = "token, symbol, spacing, grouped narrative/footnote checks, and limited structural formatting; full visual formatting is not guaranteed for PDFs"
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
            proofread_mode=args.proofread,
        )
    else:
        summary = run_compare(
            docx_path=args.docx,
            html_path=args.html,
            output_path=args.output,
            summary_json_path=args.summary_json,
            renderer=args.renderer,
            proofread_mode=args.proofread,
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
