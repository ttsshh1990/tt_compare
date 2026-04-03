from __future__ import annotations

import unittest
from collections import Counter
from pathlib import Path

import generate_diff_pdf as g


ROOT = Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "local_test" / "sample_inputs"
Q1_DOCX = SAMPLES / "2026-02-25 SNPS_Q1'26_EarningsRelease_Final - Test.docx"
Q1_PDF = SAMPLES / "SNPS Form 8-K - Q1'26 Earnings & Stock Replenishment_Final.pdf"
Q3_DOCX = SAMPLES / "2025-09-09 SNPS_Q3'25_EarningsRelease_Draft_0908_1006AM.docx"
Q3_PDF = SAMPLES / "SNPS - Q3'25 Earnings 8-K (Final Proof).pdf"


class CompareHelperTests(unittest.TestCase):
    @staticmethod
    def _table_block(
        text: str,
        *,
        source: str,
        row_key: str,
        row_slot: int,
        numeric_slot: int | None = None,
        table_pos: tuple[int, int, int] = (0, 0, 0),
    ) -> g.Block:
        return g.Block(
            id=f"{source}-{text}-{table_pos}",
            source=source,
            order=table_pos[1] * 10 + table_pos[2],
            text=text,
            normalized=g.normalize_for_compare(text),
            table_cell=True,
            row_key=row_key,
            row_slot=row_slot,
            numeric_slot=numeric_slot,
            table_pos=table_pos,
        )

    def test_diff_tokens_keeps_percent_suffix_when_spaced(self) -> None:
        tokens = g.diff_tokens("18\t%")
        self.assertEqual(len(tokens), 1)
        self.assertEqual(tokens[0].text, "18%")
        self.assertEqual(tokens[0].normalized, "18%")

    def test_pdf_artifact_normalization_handles_hyphenated_wrap_and_page_number(self) -> None:
        text = "our short-\nterm financial targets\n \n8"
        self.assertEqual(
            g.normalize_pdf_paragraph_artifacts(text),
            g.normalize_for_compare("our short-term financial targets"),
        )

    def test_pdf_blocks_equal_after_cleanup_ignores_footnote_marker_loss(self) -> None:
        doc = g.Block(id="d", source="docx", order=1, text="Free Cash Flow(1)", normalized=g.normalize_for_compare("Free Cash Flow(1)"))
        pdf = g.Block(id="p", source="pdf", order=1, text="Free Cash Flow)", normalized=g.normalize_for_compare("Free Cash Flow)"))
        self.assertTrue(g.pdf_blocks_equal_after_cleanup(doc, pdf))

    def test_pdf_blocks_equal_after_cleanup_ignores_hyphenated_line_breaks(self) -> None:
        doc = g.Block(
            id="d",
            source="docx",
            order=1,
            text="our short-term financial targets",
            normalized=g.normalize_for_compare("our short-term financial targets"),
        )
        pdf = g.Block(
            id="p",
            source="pdf",
            order=1,
            text="our short-\nterm financial targets\n8",
            normalized=g.normalize_for_compare("our short-\nterm financial targets\n8"),
        )
        self.assertTrue(g.pdf_blocks_equal_after_cleanup(doc, pdf))

    def test_large_amount_difference_keeps_currency_context_from_row_marker(self) -> None:
        row_key = "acquisition/divestiture related items"
        doc_label = self._table_block(
            "Acquisition/divestiture related items",
            source="docx",
            row_key=row_key,
            row_slot=0,
            table_pos=(1, 11, 0),
        )
        doc_currency = self._table_block(
            "$",
            source="docx",
            row_key=row_key,
            row_slot=1,
            table_pos=(1, 11, 1),
        )
        doc_value = self._table_block(
            "15,59",
            source="docx",
            row_key=row_key,
            row_slot=1,
            numeric_slot=0,
            table_pos=(1, 11, 2),
        )
        pdf_label = self._table_block(
            "Acquisition/divestiture related items",
            source="pdf",
            row_key=row_key,
            row_slot=0,
            table_pos=(6, 25, 0),
        )
        pdf_value = self._table_block(
            "15,592",
            source="pdf",
            row_key=row_key,
            row_slot=1,
            numeric_slot=0,
            table_pos=(6, 25, 1),
        )
        comments = g.text_difference_comments(
            doc_value,
            pdf_value,
            0.95,
            target_name="pdf",
            docx_blocks=[doc_label, doc_currency, doc_value],
            target_blocks=[pdf_label, pdf_value],
            match_type="approx",
        )
        self.assertEqual(
            [comment.contents for comment in comments],
            ["The number is different, $15,592 in pdf while $15,59 in word."],
        )

    def test_spacing_difference_comment_preserves_currency_symbol(self) -> None:
        row_key = "gaap eps"
        doc = self._table_block(
            "$    0.34",
            source="docx",
            row_key=row_key,
            row_slot=1,
            numeric_slot=0,
            table_pos=(1, 9, 2),
        )
        pdf = self._table_block(
            "$0.34",
            source="pdf",
            row_key=row_key,
            row_slot=1,
            numeric_slot=0,
            table_pos=(6, 17, 1),
        )
        comments = g.text_difference_comments(
            doc,
            pdf,
            0.99,
            target_name="pdf",
            docx_blocks=[doc],
            target_blocks=[pdf],
            match_type="approx",
        )
        self.assertEqual(
            [comment.contents for comment in comments],
            ["The spacing is different, pdf has $0.34 while word has $   0.34."],
        )

    def test_date_difference_comment_uses_full_dates(self) -> None:
        doc = g.Block(
            id="d",
            source="docx",
            order=1,
            text="The filing will be filed on or before March 13, 2026.",
            normalized=g.normalize_for_compare("The filing will be filed on or before March 13, 2026."),
        )
        pdf = g.Block(
            id="p",
            source="pdf",
            order=1,
            text="The filing will be filed on or before March 12, 2026.",
            normalized=g.normalize_for_compare("The filing will be filed on or before March 12, 2026."),
        )
        comments = g.text_difference_comments(
            doc,
            pdf,
            0.95,
            target_name="pdf",
            docx_blocks=[doc],
            target_blocks=[pdf],
            match_type="approx",
        )
        self.assertEqual(
            [comment.contents for comment in comments],
            ["The date is different, March 12, 2026 in pdf while March 13, 2026 in word."],
        )

    def test_pdf_internal_dash_ocr_loss_does_not_raise_false_comment(self) -> None:
        doc = g.Block(
            id="d",
            source="docx",
            order=1,
            text="GAAP total operating income - as reported",
            normalized=g.normalize_for_compare("GAAP total operating income - as reported"),
        )
        pdf = g.Block(
            id="p",
            source="pdf",
            order=1,
            text="GAAP total operating income as reported",
            normalized=g.normalize_for_compare("GAAP total operating income as reported"),
        )
        comments = g.text_difference_comments(
            doc,
            pdf,
            0.99,
            target_name="pdf",
            docx_blocks=[doc],
            target_blocks=[pdf],
            match_type="approx",
        )
        self.assertEqual(comments, [])


class Q1PdfRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.docx_blocks = g.extract_docx_blocks(Q1_DOCX)
        cls.render_result = g.extract_pdf_blocks(Q1_PDF)
        cls.pdf_blocks = cls.render_result.blocks
        cls.matches, cls.unmatched_docx, cls.unmatched_pdf = g.compare_blocks(cls.docx_blocks, cls.pdf_blocks)

    def _find_doc_block(self, predicate) -> g.Block:
        return next(block for block in self.docx_blocks if predicate(block))

    def _find_pdf_block(self, predicate) -> g.Block:
        return next(block for block in self.pdf_blocks if predicate(block))

    def test_currency_comment_for_gaap_eps_missing_dollar(self) -> None:
        doc = self._find_doc_block(lambda b: b.row_key == "gaap eps" and b.row_slot == 1 and b.text.strip() == "0.23")
        pdf = self._find_pdf_block(lambda b: b.row_key == "gaap eps" and b.row_slot == 1 and b.text.strip() == "0.23")
        comments = g.text_difference_comments(
            doc,
            pdf,
            1.0,
            target_name="pdf",
            docx_blocks=self.docx_blocks,
            target_blocks=self.pdf_blocks,
        )
        self.assertEqual(
            [comment.contents for comment in comments],
            ["The currency symbol is different, pdf has $0.23 while word has 0.23."],
        )

    def test_small_decimal_difference_does_not_infer_currency(self) -> None:
        doc = self._find_doc_block(lambda b: b.text == "0.06" and b.row_key == "acquisition/divestiture related items")
        pdf = self._find_pdf_block(lambda b: b.text == "0.08" and b.row_key == "acquisition/divestiture related items")
        comments = g.text_difference_comments(
            doc,
            pdf,
            0.95,
            target_name="pdf",
            docx_blocks=self.docx_blocks,
            target_blocks=self.pdf_blocks,
            match_type="approx",
        )
        self.assertEqual(
            [comment.contents for comment in comments],
            ["The number is different, 0.08 in pdf while 0.06 in word."],
        )

    def test_percent_rows_do_not_generate_false_percent_differences(self) -> None:
        bad_comments: list[str] = []
        for match in self.matches:
            doc = self.docx_blocks[match.docx_index]
            pdf = self.pdf_blocks[match.html_index]
            comments = g.text_difference_comments(
                doc,
                pdf,
                match.score,
                target_name="pdf",
                docx_blocks=self.docx_blocks,
                target_blocks=self.pdf_blocks,
                match_type=match.match_type,
            )
            for comment in comments:
                if "18% in pdf while 18 in word" in comment.contents or "18.1% in pdf while 18.1 in word" in comment.contents:
                    bad_comments.append(comment.contents)
        self.assertEqual(bad_comments, [])

    def test_footnote_marker_loss_does_not_generate_missing_number_comment(self) -> None:
        doc = self._find_doc_block(lambda b: b.text == "Free Cash Flow(1)")
        pdf = self._find_pdf_block(lambda b: b.text == "Free Cash Flow)")
        comments = g.text_difference_comments(
            doc,
            pdf,
            0.9747826086956523,
            target_name="pdf",
            docx_blocks=self.docx_blocks,
            target_blocks=self.pdf_blocks,
            match_type="approx",
        )
        self.assertEqual(comments, [])


class Q3PdfRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.docx_blocks = g.extract_docx_blocks(Q3_DOCX)
        cls.render_result = g.extract_pdf_blocks(Q3_PDF)
        cls.pdf_blocks = cls.render_result.blocks
        cls.matches, cls.unmatched_docx, cls.unmatched_pdf = g.compare_blocks(cls.docx_blocks, cls.pdf_blocks)

    def test_page_summary_comments_are_suppressed_when_docx_has_anchor_sections(self) -> None:
        comments = g.pdf_page_summary_comments(
            docx_blocks=self.docx_blocks,
            pdf_blocks=self.pdf_blocks,
            unmatched_pdf=self.unmatched_pdf,
            matches=self.matches,
            render_result=self.render_result,
        )
        self.assertEqual(comments, [])

    def test_forward_looking_paragraph_ignores_hyphenated_line_break_and_page_number_artifacts(self) -> None:
        doc = next(block for block in self.docx_blocks if "short-term and long-term financial targets" in block.text)
        pdf = next(block for block in self.pdf_blocks if "short-\nterm and long-term financial targets" in block.text)
        comments = g.text_difference_comments(
            doc,
            pdf,
            0.99,
            target_name="pdf",
            docx_blocks=self.docx_blocks,
            target_blocks=self.pdf_blocks,
            match_type="approx",
        )
        self.assertEqual(comments, [])


if __name__ == "__main__":
    unittest.main()
