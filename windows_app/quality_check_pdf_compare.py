#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from pypdf import PdfReader

import generate_diff_pdf as comparator


REQUIRED_COMMENTS = [
    "This PDF page contains content with no corresponding content in the Word document.",
    "The currency symbol is different, pdf has $0.23 while word has 0.23.",
    "The number is different, 0.08 in pdf while 0.06 in word.",
]

FORBIDDEN_COMMENTS = [
    "The currency symbol is different, pdf has $1,915 while word has 1,915.",
    "The currency symbol is different, pdf has $1,860 while word has 1,860.",
    "The currency symbol is different, pdf has 2,225 while word has $2,225.",
    "The currency symbol is different, pdf has 2,020 while word has $2,020.",
    "The currency symbol is different, pdf has $189,593 while word has 189,593.",
    "The currency symbol is different, pdf has $154,408 while word has 154,408.",
    "The currency symbol is different, pdf has $190,762 while word has 190,762.",
    "The currency symbol is different, pdf has $156,189 while word has 156,189.",
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run the DOCX vs PDF compare on the sample pair and validate expected annotations.",
    )
    parser.add_argument(
        "--docx",
        type=Path,
        default=Path("local_test/sample_inputs/2026-02-25 SNPS_Q1'26_EarningsRelease_Final - Test.docx"),
    )
    parser.add_argument(
        "--pdf",
        type=Path,
        default=Path("local_test/sample_inputs/SNPS Form 8-K - Q1'26 Earnings & Stock Replenishment_Final.pdf"),
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("quality_check_output.pdf"),
    )
    parser.add_argument(
        "--summary-json",
        type=Path,
        default=Path("quality_check_output.json"),
    )
    parser.add_argument(
        "--report-json",
        type=Path,
        default=Path("quality_check_report.json"),
    )
    return parser.parse_args()


def extract_annotation_texts(pdf_path: Path) -> list[dict[str, object]]:
    reader = PdfReader(str(pdf_path))
    annotations: list[dict[str, object]] = []
    for page_number, page in enumerate(reader.pages, start=1):
        for annot_ref in (page.get("/Annots") or []):
            annot = annot_ref.get_object()
            contents = str(annot.get("/Contents") or "").strip()
            if not contents:
                continue
            annotations.append({"page": page_number, "contents": contents})
    return annotations


def main() -> int:
    args = parse_args()
    summary = comparator.run_compare_pdf(
        docx_path=args.docx,
        pdf_path=args.pdf,
        output_path=args.output,
        summary_json_path=args.summary_json,
    )
    annotations = extract_annotation_texts(args.output)
    annotation_texts = [item["contents"] for item in annotations]

    required_results = []
    for needle in REQUIRED_COMMENTS:
        count = sum(needle in text for text in annotation_texts)
        required_results.append({"comment": needle, "count": count, "ok": count > 0})

    forbidden_results = []
    for needle in FORBIDDEN_COMMENTS:
        count = sum(needle in text for text in annotation_texts)
        forbidden_results.append({"comment": needle, "count": count, "ok": count == 0})

    page_comment_counts: dict[str, int] = {}
    for annotation in annotations:
        page_comment_counts[str(annotation["page"])] = page_comment_counts.get(str(annotation["page"]), 0) + 1

    passed = all(item["ok"] for item in required_results) and all(item["ok"] for item in forbidden_results)
    report = {
        "passed": passed,
        "docx": str(args.docx),
        "pdf": str(args.pdf),
        "output_pdf": str(args.output),
        "summary_json": str(args.summary_json),
        "summary": summary,
        "annotation_count": len(annotations),
        "page_comment_counts": page_comment_counts,
        "required_results": required_results,
        "forbidden_results": forbidden_results,
        "annotations": annotations,
        "limitations": [
            "The PDF path relies on OCR because this sample PDF does not expose a text layer.",
            "The automated quality check verifies text, number, symbol, and extra-page findings for this sample pair.",
            "Reliable format-by-format PDF validation is still limited because OCR does not preserve formatting semantics.",
        ],
    }
    args.report_json.write_text(json.dumps(report, indent=2), encoding="utf-8")

    print(json.dumps(report, indent=2))
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
