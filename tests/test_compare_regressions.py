from __future__ import annotations

import unittest
from collections import Counter
from pathlib import Path
import tempfile
import xml.etree.ElementTree as ET
import zipfile

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

    def test_diff_tokens_preserve_internal_spacing_after_currency_symbol(self) -> None:
        tokens = g.diff_tokens("$    0.34")
        self.assertEqual(len(tokens), 1)
        self.assertEqual(tokens[0].text, "0.34")
        self.assertEqual(tokens[0].prefix_symbol, "$")
        self.assertEqual(tokens[0].spaces_before, 4)

    def test_diff_tokens_keep_percent_on_parenthesized_number(self) -> None:
        tokens = g.diff_tokens("(2.1) %")
        self.assertEqual(len(tokens), 1)
        self.assertEqual(tokens[0].text, "(2.1)%")
        self.assertIsNone(tokens[0].prefix_symbol)
        self.assertEqual(tokens[0].normalized, "(2.1)%")

    def test_collect_word_runs_preserve_symbol_and_whitespace(self) -> None:
        paragraph = ET.fromstring(
            """
            <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:r><w:t>Net income was </w:t></w:r>
              <w:r><w:t>$</w:t></w:r>
              <w:r><w:t xml:space="preserve">    </w:t></w:r>
              <w:r><w:t>0.34</w:t></w:r>
            </w:p>
            """
        )
        runs = g.collect_word_runs(paragraph)
        raw_text, proof_text, match_text = g.block_texts_from_runs(runs)
        self.assertEqual(raw_text, "Net income was $    0.34")
        self.assertEqual(proof_text, "Net income was $    0.34")
        self.assertEqual(match_text, "net income was $0.34")
        self.assertEqual([run.kind for run in runs], ["text", "symbol", "space", "text"])

    def test_collect_word_runs_resolve_paragraph_and_character_styles(self) -> None:
        styles_xml = """
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:docDefaults>
            <w:rPrDefault><w:rPr><w:i/></w:rPr></w:rPrDefault>
          </w:docDefaults>
          <w:style w:type="paragraph" w:styleId="ParaStrong">
            <w:rPr><w:b/></w:rPr>
          </w:style>
          <w:style w:type="character" w:styleId="CharUnderline">
            <w:rPr><w:u w:val="single"/></w:rPr>
          </w:style>
        </w:styles>
        """
        paragraph = ET.fromstring(
            """
            <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:pPr><w:pStyle w:val="ParaStrong"/></w:pPr>
              <w:r>
                <w:rPr><w:rStyle w:val="CharUnderline"/></w:rPr>
                <w:t>Styled text</w:t>
              </w:r>
            </w:p>
            """
        )
        with tempfile.NamedTemporaryFile(suffix=".docx") as tmp:
            with zipfile.ZipFile(tmp.name, "w") as archive:
                archive.writestr("word/styles.xml", styles_xml)
            with zipfile.ZipFile(tmp.name) as archive:
                resolver = g.build_word_style_resolver(archive)
            runs = g.collect_word_runs(paragraph, resolver=resolver)
        self.assertEqual(len(runs), 1)
        self.assertTrue(runs[0].bold)
        self.assertTrue(runs[0].italic)
        self.assertTrue(runs[0].underline)

    def test_collect_word_runs_apply_hyperlink_style(self) -> None:
        styles_xml = """
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:style w:type="character" w:styleId="Hyperlink">
            <w:name w:val="Hyperlink"/>
            <w:rPr><w:u w:val="single"/></w:rPr>
          </w:style>
        </w:styles>
        """
        paragraph = ET.fromstring(
            """
            <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <w:hyperlink r:id="rId1">
                <w:r><w:t>www.synopsys.com</w:t></w:r>
              </w:hyperlink>
            </w:p>
            """
        )
        with tempfile.NamedTemporaryFile(suffix=".docx") as tmp:
            with zipfile.ZipFile(tmp.name, "w") as archive:
                archive.writestr("word/styles.xml", styles_xml)
            with zipfile.ZipFile(tmp.name) as archive:
                resolver = g.build_word_style_resolver(archive)
            runs = g.collect_word_runs(paragraph, resolver=resolver)
        self.assertEqual(len(runs), 1)
        self.assertTrue(runs[0].hyperlink)
        self.assertTrue(runs[0].underline)

    def test_collect_html_inline_runs_preserve_symbol_and_whitespace(self) -> None:
        element = ET.fromstring("<p>Net income was <span>$</span>    <span>0.34</span></p>")
        runs = g.collect_html_inline_runs(element)
        raw_text, proof_text, match_text = g.block_texts_from_runs(runs)
        self.assertEqual(raw_text, "Net income was $    0.34")
        self.assertEqual(proof_text, "Net income was $    0.34")
        self.assertEqual(match_text, "net income was $0.34")
        self.assertEqual([run.kind for run in runs], ["text", "symbol", "space", "text"])

    def test_collect_html_inline_runs_preserve_br_linebreak_for_lead_label_split(self) -> None:
        element = ET.fromstring("<p><i>GAAP Results<br /></i>On a GAAP basis.</p>")
        runs = g.collect_html_inline_runs(element)
        raw_text, proof_text, match_text = g.block_texts_from_runs(runs)
        self.assertEqual(raw_text, "GAAP Results\nOn a GAAP basis.")
        self.assertEqual(proof_text, "GAAP Results\nOn a GAAP basis.")
        self.assertEqual(match_text, "gaap results on a gaap basis.")
        self.assertEqual([run.kind for run in runs], ["text", "linebreak", "text"])

    def test_compare_inline_formatting_uses_only_shared_text_span(self) -> None:
        doc = g.Block(
            id="d",
            source="docx",
            order=1,
            text="Synopsys, Inc.",
            normalized=g.normalize_for_compare("Synopsys, Inc."),
            runs=[g.InlineRun(text="Synopsys, Inc.", kind="text", bold=False)],
        )
        html = g.Block(
            id="h",
            source="html",
            order=1,
            text="INVESTOR CONTACT:\nSynopsys, Inc.",
            normalized=g.normalize_for_compare("INVESTOR CONTACT:\nSynopsys, Inc."),
            runs=[
                g.InlineRun(text="INVESTOR CONTACT", kind="text", bold=True),
                g.InlineRun(text=":\n", kind="text", bold=False),
                g.InlineRun(text="Synopsys, Inc.", kind="text", bold=False),
            ],
        )
        self.assertEqual(g.compare_inline_formatting_diffs(doc, html), [])

    def test_summarize_formatting_diff_skips_span_style_on_table_cell_mismatch(self) -> None:
        doc = g.Block(
            id="d",
            source="docx",
            order=1,
            text="SYNOPSYS, INC.",
            normalized=g.normalize_for_compare("SYNOPSYS, INC."),
            kind="p",
            table_cell=False,
            runs=[g.InlineRun(text="SYNOPSYS, INC.", kind="text", bold=False)],
        )
        html = g.Block(
            id="h",
            source="html",
            order=1,
            text="SYNOPSYS, INC.",
            normalized=g.normalize_for_compare("SYNOPSYS, INC."),
            kind="td",
            table_cell=True,
            runs=[g.InlineRun(text="SYNOPSYS, INC.", kind="text", bold=True)],
        )
        self.assertEqual(
            g.summarize_formatting_diff(doc, html),
            ["DOCX does not have table cell; HTML has it."],
        )

    def test_compare_blocks_prematches_same_cell_numeric_table_typo(self) -> None:
        doc = g.Block(
            id="doc",
            source="docx",
            order=1,
            text="$15,59",
            normalized=g.normalize_for_compare("$15,59"),
            raw_text="\t$15,59\t",
            proof_text="\t$15,59\t",
            match_text="$15,59",
            table_cell=True,
            kind="td",
            table_pos=(1, 11, 1),
            row_key="acquisition/divestiture related items",
            row_slot=1,
            numeric_slot=0,
            runs=[
                g.InlineRun(text="\t", kind="tab"),
                g.InlineRun(text="$", kind="symbol"),
                g.InlineRun(text="15,59", kind="text"),
                g.InlineRun(text="\t", kind="tab"),
            ],
        )
        html = g.Block(
            id="html",
            source="html",
            order=1,
            text="15,592",
            normalized=g.normalize_for_compare("15,592"),
            raw_text="15,592",
            proof_text="15,592",
            match_text="15,592",
            table_cell=True,
            kind="td",
            table_pos=(1, 11, 1),
            row_key="acquisition/divestiture related items",
            row_slot=1,
            numeric_slot=0,
            runs=[g.InlineRun(text="15,592", kind="text")],
        )
        matches, docx_only, html_only = g.compare_blocks([doc], [html])
        self.assertEqual(len(matches), 1)
        self.assertEqual(matches[0].html_index, 0)
        self.assertEqual(matches[0].match_type, "approx")
        self.assertEqual(matches[0].score, 0.95)
        self.assertEqual(docx_only, [])
        self.assertEqual(html_only, [])

    def test_align_table_families_ignores_tiny_html_table(self) -> None:
        doc_blocks = [
            self._table_block("Revenue", source="docx", row_key="revenue", row_slot=0, table_pos=(0, 0, 0)),
            self._table_block("2,230", source="docx", row_key="revenue", row_slot=1, numeric_slot=0, table_pos=(0, 0, 1)),
            self._table_block("GAAP Expenses", source="docx", row_key="gaap expenses", row_slot=2, table_pos=(0, 1, 0)),
            self._table_block("2,115,000", source="docx", row_key="gaap expenses", row_slot=3, numeric_slot=1, table_pos=(0, 1, 1)),
            self._table_block("Target Revenue", source="docx", row_key="target revenue", row_slot=0, table_pos=(1, 0, 0)),
            self._table_block("7,030", source="docx", row_key="target revenue", row_slot=1, numeric_slot=0, table_pos=(1, 0, 1)),
            self._table_block("Target GAAP Expenses", source="docx", row_key="target gaap expenses", row_slot=2, table_pos=(1, 1, 0)),
            self._table_block("6,078,598", source="docx", row_key="target gaap expenses", row_slot=3, numeric_slot=1, table_pos=(1, 1, 1)),
        ]
        html_blocks = [
            self._table_block("1 The operating results of Ansys", source="html", row_key="ansys footnote", row_slot=0, table_pos=(0, 0, 0)),
            self._table_block("Revenue", source="html", row_key="revenue", row_slot=0, table_pos=(1, 0, 0)),
            self._table_block("2,230", source="html", row_key="revenue", row_slot=1, numeric_slot=0, table_pos=(1, 0, 1)),
            self._table_block("GAAP Expenses", source="html", row_key="gaap expenses", row_slot=2, table_pos=(1, 1, 0)),
            self._table_block("2,115,000", source="html", row_key="gaap expenses", row_slot=3, numeric_slot=1, table_pos=(1, 1, 1)),
            self._table_block("Target Revenue", source="html", row_key="target revenue", row_slot=0, table_pos=(2, 0, 0)),
            self._table_block("7,030", source="html", row_key="target revenue", row_slot=1, numeric_slot=0, table_pos=(2, 0, 1)),
            self._table_block("Target GAAP Expenses", source="html", row_key="target gaap expenses", row_slot=2, table_pos=(2, 1, 0)),
            self._table_block("6,078,598", source="html", row_key="target gaap expenses", row_slot=3, numeric_slot=1, table_pos=(2, 1, 1)),
        ]
        doc_to_html, html_to_doc = g.align_table_families(doc_blocks, html_blocks)
        self.assertEqual(doc_to_html, {0: 1, 1: 2})
        self.assertEqual(html_to_doc, {1: 0, 2: 1})

    def test_compare_blocks_restricts_table_matching_to_aligned_family(self) -> None:
        doc_blocks = [
            self._table_block("Revenue", source="docx", row_key="revenue", row_slot=0, table_pos=(0, 0, 0)),
            self._table_block("2,230", source="docx", row_key="revenue", row_slot=1, numeric_slot=0, table_pos=(0, 0, 1)),
            self._table_block("GAAP Expenses", source="docx", row_key="gaap expenses", row_slot=2, table_pos=(0, 1, 0)),
            self._table_block("2,115,000", source="docx", row_key="gaap expenses", row_slot=3, numeric_slot=1, table_pos=(0, 1, 1)),
            self._table_block("Target Revenue", source="docx", row_key="target revenue", row_slot=0, table_pos=(1, 0, 0)),
            self._table_block("7,030", source="docx", row_key="target revenue", row_slot=1, numeric_slot=0, table_pos=(1, 0, 1)),
            self._table_block("Target GAAP Expenses", source="docx", row_key="target gaap expenses", row_slot=2, table_pos=(1, 1, 0)),
            self._table_block("6,078,598", source="docx", row_key="target gaap expenses", row_slot=3, numeric_slot=1, table_pos=(1, 1, 1)),
        ]
        html_blocks = [
            self._table_block("1 The operating results of Ansys", source="html", row_key="ansys footnote", row_slot=0, table_pos=(0, 0, 0)),
            self._table_block("Revenue", source="html", row_key="revenue", row_slot=0, table_pos=(1, 0, 0)),
            self._table_block("2,230", source="html", row_key="revenue", row_slot=1, numeric_slot=0, table_pos=(1, 0, 1)),
            self._table_block("GAAP Expenses", source="html", row_key="gaap expenses", row_slot=2, table_pos=(1, 1, 0)),
            self._table_block("2,115,000", source="html", row_key="gaap expenses", row_slot=3, numeric_slot=1, table_pos=(1, 1, 1)),
            self._table_block("Target Revenue", source="html", row_key="target revenue", row_slot=0, table_pos=(2, 0, 0)),
            self._table_block("7,030", source="html", row_key="target revenue", row_slot=1, numeric_slot=0, table_pos=(2, 0, 1)),
            self._table_block("Target GAAP Expenses", source="html", row_key="target gaap expenses", row_slot=2, table_pos=(2, 1, 0)),
            self._table_block("6,102,598", source="html", row_key="target gaap expenses", row_slot=3, numeric_slot=1, table_pos=(2, 1, 1)),
        ]
        matches, docx_only, html_only = g.compare_blocks(doc_blocks, html_blocks)
        target_match = next(match for match in matches if doc_blocks[match.docx_index].text == "6,078,598")
        self.assertEqual(target_match.html_index, 8)
        self.assertEqual(target_match.match_type, "approx")
        self.assertEqual(docx_only, [])
        self.assertEqual({block.id for block in html_only}, {html_blocks[0].id})

    def test_extract_block_groups_collects_docx_contact_lines(self) -> None:
        blocks = [
            g.Block(id="d0", source="docx", order=0, text="INVESTOR CONTACT:", normalized=g.normalize_for_compare("INVESTOR CONTACT:"), kind="p"),
            g.Block(id="d1", source="docx", order=1, text="Tushar Jain", normalized=g.normalize_for_compare("Tushar Jain"), kind="p"),
            g.Block(id="d2", source="docx", order=2, text="Synopsys, Inc.", normalized=g.normalize_for_compare("Synopsys, Inc."), kind="p"),
            g.Block(id="d3", source="docx", order=3, text="650-584-4289", normalized=g.normalize_for_compare("650-584-4289"), kind="p"),
            g.Block(id="d4", source="docx", order=4, text="Synopsys-ir@synopsys.com", normalized=g.normalize_for_compare("Synopsys-ir@synopsys.com"), kind="p"),
        ]
        groups, block_to_group = g.extract_block_groups(blocks)
        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0].group_type, "contact")
        self.assertEqual(groups[0].block_indices, [0, 1, 2, 3, 4])
        self.assertEqual({block_to_group[index] for index in range(5)}, {groups[0].group_id})

    def test_match_block_groups_matches_contact_group_to_merged_html_contact(self) -> None:
        doc_blocks = [
            g.Block(id="d0", source="docx", order=0, text="INVESTOR CONTACT:", normalized=g.normalize_for_compare("INVESTOR CONTACT:"), kind="p"),
            g.Block(id="d1", source="docx", order=1, text="Tushar Jain", normalized=g.normalize_for_compare("Tushar Jain"), kind="p"),
            g.Block(id="d2", source="docx", order=2, text="Synopsys, Inc.", normalized=g.normalize_for_compare("Synopsys, Inc."), kind="p"),
            g.Block(id="d3", source="docx", order=3, text="650-584-4289", normalized=g.normalize_for_compare("650-584-4289"), kind="p"),
            g.Block(id="d4", source="docx", order=4, text="Synopsys-ir@synopsys.com", normalized=g.normalize_for_compare("Synopsys-ir@synopsys.com"), kind="p"),
        ]
        html_blocks = [
            g.Block(
                id="h0",
                source="html",
                order=0,
                text="INVESTOR CONTACT: Tushar Jain Synopsys, Inc. 650-584-4289 Synopsys-ir@synopsys.com",
                normalized=g.normalize_for_compare("INVESTOR CONTACT: Tushar Jain Synopsys, Inc. 650-584-4289 Synopsys-ir@synopsys.com"),
                kind="p",
            )
        ]
        grouped_target_by_doc_index, _doc_map, target_map, _doc_groups, _target_groups = g.match_block_groups(doc_blocks, html_blocks)
        self.assertEqual({grouped_target_by_doc_index[index] for index in range(5)}, {target_map[0]})

    def test_compare_blocks_contact_field_match_emits_precise_name_difference(self) -> None:
        doc_blocks = [
            g.Block(id="d0", source="docx", order=0, text="INVESTOR CONTACT:", normalized=g.normalize_for_compare("INVESTOR CONTACT:"), kind="p"),
            g.Block(id="d1", source="docx", order=1, text="Tushar Jain", normalized=g.normalize_for_compare("Tushar Jain"), kind="p"),
            g.Block(id="d2", source="docx", order=2, text="Synopsys, Inc.", normalized=g.normalize_for_compare("Synopsys, Inc."), kind="p"),
            g.Block(id="d3", source="docx", order=3, text="650-584-4289", normalized=g.normalize_for_compare("650-584-4289"), kind="p"),
            g.Block(id="d4", source="docx", order=4, text="Synopsys-ir@synopsys.com", normalized=g.normalize_for_compare("Synopsys-ir@synopsys.com"), kind="p"),
        ]
        html_blocks = [
            g.Block(
                id="h0",
                source="html",
                order=0,
                text="INVESTOR CONTACT: Tushar Jans Synopsys, Inc. 650-584-4289 Synopsys-ir@synopsys.com",
                normalized=g.normalize_for_compare("INVESTOR CONTACT: Tushar Jans Synopsys, Inc. 650-584-4289 Synopsys-ir@synopsys.com"),
                kind="p",
            )
        ]
        matches, docx_only, html_only = g.compare_blocks(doc_blocks, html_blocks)
        self.assertEqual(docx_only, [])
        self.assertEqual(html_only, [])
        comments, _appendix = g.build_comments(doc_blocks, html_blocks, matches, [], [], target_label="HTML", proofread_mode=True)
        self.assertEqual(
            [comment.contents for comment in comments],
            ["The word is different, Jans in html while Jain in word."],
        )

    def test_build_comments_suppresses_contact_line_vs_paragraph_noise(self) -> None:
        doc_blocks = [
            g.Block(id="d0", source="docx", order=0, text="INVESTOR CONTACT:", normalized=g.normalize_for_compare("INVESTOR CONTACT:"), kind="p"),
            g.Block(id="d1", source="docx", order=1, text="650-584-4289", normalized=g.normalize_for_compare("650-584-4289"), kind="p"),
            g.Block(id="d2", source="docx", order=2, text="Synopsys-ir@synopsys.com", normalized=g.normalize_for_compare("Synopsys-ir@synopsys.com"), kind="p"),
        ]
        html_blocks = [
            g.Block(
                id="h0",
                source="html",
                order=0,
                text="INVESTOR CONTACT: Tushar Jain Synopsys, Inc. 650-584-4289 Synopsys-ir@synopsys.com",
                normalized=g.normalize_for_compare("INVESTOR CONTACT: Tushar Jain Synopsys, Inc. 650-584-4289 Synopsys-ir@synopsys.com"),
                kind="p",
            )
        ]
        matches = [
            g.Match(docx_index=0, html_index=0, match_type="approx", score=0.9, formatting_diffs=[]),
            g.Match(docx_index=1, html_index=0, match_type="approx", score=0.9, formatting_diffs=[]),
            g.Match(docx_index=2, html_index=0, match_type="approx", score=0.9, formatting_diffs=[]),
        ]
        comments, appendix = g.build_comments(doc_blocks, html_blocks, matches, [], [], target_label="HTML", proofread_mode=True)
        self.assertEqual(comments, [])
        self.assertEqual(appendix, [])

    def test_unmatched_html_quote_block_does_not_emit_word_level_near_match_comments(self) -> None:
        doc_blocks = [
            g.Block(id="d0", source="docx", order=0, text='"CEO Quote TBC"', normalized=g.normalize_for_compare('"CEO Quote TBC"'), kind="p"),
        ]
        html_blocks = [
            g.Block(
                id="h0",
                source="html",
                order=0,
                text='"Q3 was a transformational quarter," said Sassine Ghazi, president and CEO of Synopsys. "We are taking action to enhance our competitive advantage."',
                normalized=g.normalize_for_compare('"Q3 was a transformational quarter," said Sassine Ghazi, president and CEO of Synopsys. "We are taking action to enhance our competitive advantage."'),
                kind="p",
            )
        ]
        comments, _appendix = g.build_comments(doc_blocks, html_blocks, [], doc_blocks, html_blocks, target_label="HTML", proofread_mode=True)
        self.assertEqual(
            [comment.contents for comment in comments],
            ["This HTML block has no corresponding content in the DOCX."],
        )

    def test_match_block_groups_matches_quote_placeholder_by_role(self) -> None:
        doc_blocks = [
            g.Block(
                id="d0",
                source="docx",
                order=0,
                text='"CEO Quote TBC"',
                normalized=g.normalize_for_compare('"CEO Quote TBC"'),
                kind="p",
            ),
        ]
        html_blocks = [
            g.Block(
                id="h0",
                source="html",
                order=0,
                text='"Q3 was a transformational quarter," said Sassine Ghazi, president and CEO of Synopsys. "We are taking action to enhance our competitive advantage."',
                normalized=g.normalize_for_compare('"Q3 was a transformational quarter," said Sassine Ghazi, president and CEO of Synopsys. "We are taking action to enhance our competitive advantage."'),
                kind="p",
            ),
        ]
        grouped_target_by_doc_index, _doc_map, target_map, _doc_groups, _target_groups = g.match_block_groups(doc_blocks, html_blocks)
        self.assertEqual(grouped_target_by_doc_index[0], target_map[0])

    def test_text_difference_comments_summarizes_quote_group(self) -> None:
        doc = g.Block(
            id="d0",
            source="docx",
            order=0,
            text='"CEO Quote TBC"',
            normalized=g.normalize_for_compare('"CEO Quote TBC"'),
            kind="p",
        )
        html = g.Block(
            id="h0",
            source="html",
            order=0,
            text='"Q3 was a transformational quarter," said Sassine Ghazi, president and CEO of Synopsys. "We are taking action to enhance our competitive advantage and drive resilient, long-term growth."',
            normalized=g.normalize_for_compare('"Q3 was a transformational quarter," said Sassine Ghazi, president and CEO of Synopsys. "We are taking action to enhance our competitive advantage and drive resilient, long-term growth."'),
            kind="p",
        )
        comments = g.text_difference_comments(
            doc,
            html,
            0.95,
            target_name="html",
            docx_blocks=[doc],
            target_blocks=[html],
            match_type="approx",
            proofread_mode=True,
            grouped_match_type="quote",
        )
        self.assertEqual(len(comments), 1)
        self.assertTrue(comments[0].contents.startswith("The CEO quote text is different."))

    def test_quote_diff_summary_marks_broad_match_when_normalized_quote_text_is_same(self) -> None:
        text = '"Synopsys enters 2026 with an expanded portfolio," said Sassine Ghazi, president and CEO of Synopsys.'
        summary = g.quote_diff_summary(text, text, target_name="html")
        self.assertTrue(summary.startswith("The CEO quote block is broadly matched rather than exactly matched."))

    def test_text_difference_comments_summarizes_medium_confidence_repeated_label(self) -> None:
        doc = g.Block(
            id="d0",
            source="docx",
            order=0,
            text="SYNOPSYS, INC.",
            normalized=g.normalize_for_compare("SYNOPSYS, INC."),
            kind="p",
        )
        html = g.Block(
            id="h0",
            source="html",
            order=0,
            text="SYNOPSYS, INC. FORM 8-K",
            normalized=g.normalize_for_compare("SYNOPSYS, INC. FORM 8-K"),
            kind="p",
        )
        comments = g.text_difference_comments(
            doc,
            html,
            0.95,
            target_name="html",
            docx_blocks=[doc],
            target_blocks=[html],
            match_type="approx",
            proofread_mode=True,
        )
        self.assertEqual(len(comments), 1)
        self.assertTrue(comments[0].contents.startswith("The paragraph text is different."))

    def test_section_lead_word_difference_is_not_suppressed(self) -> None:
        doc = g.Block(
            id="d0",
            source="docx",
            order=0,
            text="Earnings Call Open to Investor",
            normalized=g.normalize_for_compare("Earnings Call Open to Investor"),
            kind="p",
            bold=True,
            structure_role="section_lead",
        )
        html = g.Block(
            id="h0",
            source="html",
            order=0,
            text="Earnings Call Open to Investors",
            normalized=g.normalize_for_compare("Earnings Call Open to Investors"),
            kind="p",
            bold=True,
            structure_role="section_lead",
        )
        comments = g.text_difference_comments(
            doc,
            html,
            0.82,
            target_name="html",
            docx_blocks=[doc],
            target_blocks=[html],
            match_type="approx",
            proofread_mode=True,
        )
        self.assertTrue(comments)
        self.assertTrue(any("Investors" in comment.contents and "Investor" in comment.contents for comment in comments))

    def test_exact_structural_short_header_family_still_emits_word_diff(self) -> None:
        doc = g.Block(
            id="d0",
            source="docx",
            order=0,
            text="Financial Targetss",
            normalized=g.normalize_for_compare("Financial Targetss"),
            kind="p",
            bold=True,
            structure_role="table_subtitle",
            family_table_index=0,
        )
        html = g.Block(
            id="h0",
            source="html",
            order=0,
            text="Financial Targets",
            normalized=g.normalize_for_compare("Financial Targets"),
            kind="p",
            bold=True,
            structure_role="table_title",
            family_table_index=0,
        )
        comments = g.text_difference_comments(
            doc,
            html,
            0.6470430107526881,
            target_name="html",
            docx_blocks=[doc],
            target_blocks=[html],
            match_type="exact_structural",
            formatting_diffs=[
                'DOCX has italic on "Financial"; HTML does not.',
                'DOCX role is "table_subtitle"; HTML role is "table_title".',
            ],
            proofread_mode=True,
        )
        self.assertTrue(any("Targets" in comment.contents and "Targetss" in comment.contents for comment in comments))

    def test_compare_blocks_promotes_section_lead_to_exact_structural(self) -> None:
        doc_blocks = [
            g.Block(
                id="d0",
                source="docx",
                order=0,
                text="Earnings Call Open to Investor",
                normalized=g.normalize_for_compare("Earnings Call Open to Investor"),
                kind="p",
                bold=True,
            ),
            g.Block(
                id="d1",
                source="docx",
                order=1,
                text="Synopsys will hold a conference call for investors today at 2:00 p.m. Pacific Time and the webcast replay will be available afterward.",
                normalized=g.normalize_for_compare("Synopsys will hold a conference call for investors today at 2:00 p.m. Pacific Time and the webcast replay will be available afterward."),
                kind="p",
            ),
        ]
        html_blocks = [
            g.Block(
                id="h0",
                source="html",
                order=0,
                text="Earnings Call Open to Investors",
                normalized=g.normalize_for_compare("Earnings Call Open to Investors"),
                kind="p",
                bold=True,
            ),
            g.Block(
                id="h1",
                source="html",
                order=1,
                text="Synopsys will hold a conference call for investors today at 2:00 p.m. Pacific Time and the webcast replay will be available afterward.",
                normalized=g.normalize_for_compare("Synopsys will hold a conference call for investors today at 2:00 p.m. Pacific Time and the webcast replay will be available afterward."),
                kind="p",
            ),
        ]
        g.assign_structural_roles(doc_blocks)
        g.assign_structural_roles(html_blocks)
        matches, _docx_only, _html_only = g.compare_blocks(doc_blocks, html_blocks)
        lead_match = next(match for match in matches if match.docx_index == 0)
        self.assertEqual(lead_match.match_type, "exact_structural")

    def test_compare_blocks_recovers_table_title_family_by_ordinal(self) -> None:
        doc_blocks = [
            g.Block(id="d0", source="docx", order=0, text="Financial Targetss", normalized=g.normalize_for_compare("Financial Targetss"), kind="p", bold=True, structure_role="table_subtitle", family_table_index=0),
            g.Block(id="d1", source="docx", order=1, text="Second Quarter and Full Fiscal Year 2026 Financial Targets", normalized=g.normalize_for_compare("Second Quarter and Full Fiscal Year 2026 Financial Targets"), table_cell=True, kind="td", table_pos=(0, 0, 0), structure_role="table_column_header", family_table_index=0),
        ]
        html_blocks = [
            g.Block(id="h0", source="html", order=0, text="Financial Targets", normalized=g.normalize_for_compare("Financial Targets"), kind="p", bold=True, structure_role="table_title", family_table_index=0),
            g.Block(id="h1", source="html", order=1, text="Second Quarter and Full Fiscal Year 2026 Financial Targets", normalized=g.normalize_for_compare("Second Quarter and Full Fiscal Year 2026 Financial Targets"), table_cell=True, kind="td", table_pos=(0, 0, 0), structure_role="table_column_header", family_table_index=0),
        ]
        matches, docx_only, html_only = g.compare_blocks(doc_blocks, html_blocks)
        self.assertEqual(docx_only, [])
        self.assertEqual(html_only, [])
        recovered = next(match for match in matches if match.docx_index == 0)
        self.assertEqual(recovered.html_index, 0)
        self.assertEqual(recovered.match_type, "exact_structural")

    def test_compare_blocks_recovers_row_label_by_lineage_ordinal(self) -> None:
        doc_blocks = [
            self._table_block("Stock-based compensation", source="docx", row_key="stock-based compensation", row_slot=0, table_pos=(2, 6, 0)),
            self._table_block("0.5%", source="docx", row_key="stock-based compensation", row_slot=1, numeric_slot=0, table_pos=(2, 6, 1)),
            self._table_block("Acquisition/divestiture related items", source="docx", row_key="acquisition/divestiture related items", row_slot=0, table_pos=(2, 7, 0)),
            self._table_block("Tax adjustments (2)", source="docx", row_key="tax adjustments", row_slot=0, table_pos=(2, 8, 0)),
            self._table_block("43.5%", source="docx", row_key="tax adjustments", row_slot=1, numeric_slot=0, table_pos=(2, 8, 1)),
        ]
        html_blocks = [
            self._table_block("Stock-based compensation", source="html", row_key="stock-based compensation", row_slot=0, table_pos=(9, 6, 0)),
            self._table_block("0.5%", source="html", row_key="stock-based compensation", row_slot=1, numeric_slot=0, table_pos=(9, 6, 1)),
            self._table_block("Acquisition/divestiture related items (1)", source="html", row_key="acquisition/divestiture related items", row_slot=0, table_pos=(9, 7, 0)),
            self._table_block("Tax adjustments", source="html", row_key="tax adjustments", row_slot=0, table_pos=(9, 8, 0)),
            self._table_block("43.5%", source="html", row_key="tax adjustments", row_slot=1, numeric_slot=0, table_pos=(9, 8, 1)),
        ]
        for block in doc_blocks:
            block.structure_role = "table_data_cell" if block.numeric_slot is not None else "table_row_label"
            block.family_table_index = 2
        for block in html_blocks:
            block.structure_role = "table_data_cell" if block.numeric_slot is not None else "table_row_label"
            block.family_table_index = 9
        matches, docx_only, html_only = g.compare_blocks(doc_blocks, html_blocks)
        self.assertEqual(docx_only, [])
        self.assertEqual(html_only, [])
        recovered = next(match for match in matches if match.docx_index == 2)
        self.assertEqual(recovered.html_index, 2)

    def test_match_section_families_matches_reconciliation_family(self) -> None:
        doc_blocks = [
            g.Block(id="d0", source="docx", order=0, text="GAAP to Non-GAAP Reconciliation of Third Quarter Fiscal Year 2025 Results", normalized=g.normalize_for_compare("GAAP to Non-GAAP Reconciliation of Third Quarter Fiscal Year 2025 Results"), kind="p", bold=True, structure_role="table_title", family_table_index=2),
            self._table_block("GAAP net income from continuing operations per diluted share attributed to Synopsys", source="docx", row_key="gaap net income from continuing operations per diluted share attributed to synopsys", row_slot=0, table_pos=(2, 0, 0)),
            self._table_block("0.46", source="docx", row_key="gaap net income from continuing operations per diluted share attributed to synopsys", row_slot=1, numeric_slot=0, table_pos=(2, 0, 1)),
        ]
        html_blocks = [
            g.Block(id="h0", source="html", order=0, text="GAAP to Non-GAAP Reconciliation of Third Quarter Fiscal Year 2025 Results", normalized=g.normalize_for_compare("GAAP to Non-GAAP Reconciliation of Third Quarter Fiscal Year 2025 Results"), kind="p", bold=True, structure_role="table_title", family_table_index=9),
            self._table_block("GAAP net income from continuing operations per diluted share attributed to Synopsys", source="html", row_key="gaap net income from continuing operations per diluted share attributed to synopsys", row_slot=0, table_pos=(9, 0, 0)),
            self._table_block("0.46", source="html", row_key="gaap net income from continuing operations per diluted share attributed to synopsys", row_slot=1, numeric_slot=0, table_pos=(9, 0, 1)),
        ]
        for block in doc_blocks[1:]:
            block.structure_role = "table_row_label" if block.numeric_slot is None else "table_data_cell"
            block.family_table_index = 2
        for block in html_blocks[1:]:
            block.structure_role = "table_row_label" if block.numeric_slot is None else "table_data_cell"
            block.family_table_index = 9
        matches, doc_map, _target_map, doc_family_map, target_family_map = g.match_section_families(doc_blocks, html_blocks)
        self.assertEqual(len(matches), 1)
        doc_family_id, target_family_id = next(iter(matches.items()))
        self.assertEqual(doc_family_map[doc_family_id].family_type, "reconciliation")
        self.assertEqual(target_family_map[target_family_id].family_type, "reconciliation")

    def test_build_section_family_appendix_comments_skips_broad_reconciliation_match(self) -> None:
        doc_blocks = [
            g.Block(id="d0", source="docx", order=0, text="GAAP to Non-GAAP Reconciliation of Third Quarter Fiscal Year 2025 Results", normalized=g.normalize_for_compare("GAAP to Non-GAAP Reconciliation of Third Quarter Fiscal Year 2025 Results"), kind="p", bold=True, structure_role="table_title", family_table_index=2),
            self._table_block("Acquisition/divestiture related items", source="docx", row_key="acquisition/divestiture related items", row_slot=0, table_pos=(2, 7, 0)),
        ]
        html_blocks = [
            g.Block(id="h0", source="html", order=0, text="GAAP to Non-GAAP Reconciliation of Third Quarter Fiscal Year 2025 Results", normalized=g.normalize_for_compare("GAAP to Non-GAAP Reconciliation of Third Quarter Fiscal Year 2025 Results"), kind="p", bold=True, structure_role="table_title", family_table_index=9),
            self._table_block("Acquisition/divestiture related items (1)", source="html", row_key="acquisition/divestiture related items", row_slot=0, table_pos=(9, 7, 0)),
        ]
        doc_blocks[1].structure_role = "table_row_label"
        doc_blocks[1].family_table_index = 2
        html_blocks[1].structure_role = "table_row_label"
        html_blocks[1].family_table_index = 9
        comments, covered = g.build_section_family_appendix_comments(doc_blocks, html_blocks, [doc_blocks[1]], target_label="HTML")
        self.assertEqual(covered, {doc_blocks[1].id})
        self.assertEqual(comments, [])

    def test_build_section_family_inline_comments_emits_structural_summary(self) -> None:
        doc_blocks = [
            g.Block(id="d0", source="docx", order=0, text="GAAP to Non-GAAP Tax Rate Reconciliation (1)", normalized=g.normalize_for_compare("GAAP to Non-GAAP Tax Rate Reconciliation (1)"), kind="p", bold=True, structure_role="table_title", family_table_index=2),
            self._table_block("Acquisition/divestiture related items", source="docx", row_key="acquisition/divestiture related items", row_slot=0, table_pos=(2, 7, 0)),
        ]
        html_blocks = [
            g.Block(id="h0", source="html", order=0, text="GAAP to Non-GAAP Tax Rate Reconciliation", normalized=g.normalize_for_compare("GAAP to Non-GAAP Tax Rate Reconciliation"), kind="p", bold=True, structure_role="table_title", family_table_index=9),
            self._table_block("Acquisition/divestiture related items (1)", source="html", row_key="acquisition/divestiture related items", row_slot=0, table_pos=(9, 7, 0)),
        ]
        doc_blocks[1].structure_role = "table_row_label"
        doc_blocks[1].family_table_index = 2
        html_blocks[1].structure_role = "table_row_label"
        html_blocks[1].family_table_index = 9
        comments, covered = g.build_section_family_inline_comments(doc_blocks, html_blocks, [doc_blocks[1]], target_label="HTML")
        self.assertEqual(covered, {doc_blocks[1].id})
        self.assertEqual(len(comments), 1)
        self.assertIn("broadly matched to the HTML", comments[0].contents)

    def test_format_review_comment_text_prefixes_confidence_and_tier(self) -> None:
        formatted = g.format_review_comment_text("The word is different, Targets in html while Targetss in word.")
        self.assertTrue(formatted.startswith("[High Confidence | Critical] "))

    def test_review_summary_lines_counts_structural_and_critical_comments(self) -> None:
        html_comments = [
            g.HtmlComment(order=0, contents="The word is different, Targets in html while Targetss in word."),
            g.HtmlComment(order=1, contents="This DOCX reconciliation section is broadly matched to the HTML but not aligned block-by-block. HTML: x Word: y"),
        ]
        appendix_comments = [
            (
                g.Block(id="d0", source="docx", order=0, text="X", normalized="x"),
                "This DOCX reconciliation section was not found in the HTML. Word: X",
            )
        ]
        lines = g.review_summary_lines(html_comments, appendix_comments)
        self.assertIn("[High Confidence | Critical]: 1", lines)
        self.assertIn("[Medium Confidence | Structural]: 2", lines)

    def test_match_block_groups_uses_footnote_marker_as_primary_signature(self) -> None:
        doc_blocks = [
            g.Block(
                id="d0",
                source="docx",
                order=0,
                text="The operating results of Ansys, which have been included in our financial results for the three and nine months ended July 31, 2025 for the period from July 17, 2025 through July 31, 2025, were not material to our overall results, unless otherwise stated.",
                normalized=g.normalize_for_compare("The operating results of Ansys, which have been included in our financial results for the three and nine months ended July 31, 2025 for the period from July 17, 2025 through July 31, 2025, were not material to our overall results, unless otherwise stated."),
                kind="footnote",
                footnote_marker="1",
            ),
        ]
        html_blocks = [
            g.Block(
                id="h0",
                source="html",
                order=0,
                text="1 The operating results of Ansys have been included in our condensed consolidated financial statements for the three and nine months ended July 31, 2025 from the Acquisition Date, and were not material to our financial results for either of these periods.",
                normalized=g.normalize_for_compare("1 The operating results of Ansys have been included in our condensed consolidated financial statements for the three and nine months ended July 31, 2025 from the Acquisition Date, and were not material to our financial results for either of these periods."),
                kind="td",
                table_cell=True,
                footnote_marker="1",
            ),
        ]
        grouped_target_by_doc_index, _doc_map, target_map, _doc_groups, _target_groups = g.match_block_groups(doc_blocks, html_blocks)
        self.assertEqual(grouped_target_by_doc_index[0], target_map[0])

    def test_footnote_group_match_still_allows_difference_comments(self) -> None:
        doc = g.Block(
            id="d0",
            source="docx",
            order=0,
            text="The operating results of Ansys, which have been included in our financial results for the three and nine months ended July 31, 2025, were not material.",
            normalized=g.normalize_for_compare("The operating results of Ansys, which have been included in our financial results for the three and nine months ended July 31, 2025, were not material."),
            kind="footnote",
            footnote_marker="1",
        )
        html = g.Block(
            id="h0",
            source="html",
            order=0,
            text="1 The operating results of Ansys have been included in our condensed consolidated financial statements for the three and nine months ended July 31, 2025 and were not material.",
            normalized=g.normalize_for_compare("1 The operating results of Ansys have been included in our condensed consolidated financial statements for the three and nine months ended July 31, 2025 and were not material."),
            kind="td",
            table_cell=True,
            footnote_marker="1",
        )
        comments = g.text_difference_comments(
            doc,
            html,
            0.96,
            target_name="html",
            docx_blocks=[doc],
            target_blocks=[html],
            match_type="approx",
            proofread_mode=True,
            grouped_match_type="footnote",
        )
        self.assertTrue(comments)
        self.assertEqual(len(comments), 1)
        self.assertTrue(comments[0].contents.startswith("The footnote text is different."))

    def test_assign_structural_roles_marks_table_title_and_subtitle(self) -> None:
        blocks = [
            g.Block(
                id="b0",
                source="docx",
                order=0,
                text="GAAP to Non-GAAP Reconciliation",
                normalized=g.normalize_for_compare("GAAP to Non-GAAP Reconciliation"),
                bold=True,
                kind="p",
            ),
            g.Block(
                id="b1",
                source="docx",
                order=1,
                text="(unaudited and in thousands)",
                normalized=g.normalize_for_compare("(unaudited and in thousands)"),
                italic=True,
                kind="p",
            ),
            g.Block(
                id="b2",
                source="docx",
                order=2,
                text="Low",
                normalized=g.normalize_for_compare("Low"),
                table_cell=True,
                kind="td",
                table_pos=(0, 0, 0),
                row_slot=0,
            ),
        ]
        g.assign_structural_roles(blocks)
        self.assertEqual(blocks[0].structure_role, "table_title")
        self.assertEqual(blocks[1].structure_role, "table_subtitle")
        self.assertEqual(blocks[0].family_table_index, 0)
        self.assertEqual(blocks[1].family_table_index, 0)

    def test_assign_structural_roles_marks_section_lead(self) -> None:
        blocks = [
            g.Block(
                id="b0",
                source="docx",
                order=0,
                text="Earnings Call Open to Investor",
                normalized=g.normalize_for_compare("Earnings Call Open to Investor"),
                bold=True,
                kind="p",
            ),
            g.Block(
                id="b1",
                source="docx",
                order=1,
                text="Synopsys will hold a conference call for financial analysts and investors today at 2:00 p.m. Pacific Time.",
                normalized=g.normalize_for_compare("Synopsys will hold a conference call for financial analysts and investors today at 2:00 p.m. Pacific Time."),
                kind="p",
            ),
        ]
        g.assign_structural_roles(blocks)
        self.assertEqual(blocks[0].structure_role, "section_lead")
        self.assertEqual(blocks[1].structure_role, "paragraph")

    def test_assign_structural_roles_marks_column_headers_and_data_rows(self) -> None:
        blocks = [
            g.Block(id="h0", source="html", order=0, text="Low", normalized=g.normalize_for_compare("Low"), table_cell=True, kind="th", table_pos=(0, 0, 0), row_slot=0),
            g.Block(id="h1", source="html", order=1, text="High", normalized=g.normalize_for_compare("High"), table_cell=True, kind="th", table_pos=(0, 0, 1), row_slot=1),
            g.Block(id="r0", source="html", order=2, text="GAAP Expenses", normalized=g.normalize_for_compare("GAAP Expenses"), table_cell=True, kind="td", table_pos=(0, 1, 0), row_key=g.normalize_row_key("GAAP Expenses"), row_slot=0),
            g.Block(id="r1", source="html", order=3, text="2,020", normalized=g.normalize_for_compare("2,020"), table_cell=True, kind="td", table_pos=(0, 1, 1), row_key=g.normalize_row_key("GAAP Expenses"), row_slot=1, numeric_slot=0),
        ]
        g.assign_structural_roles(blocks)
        self.assertEqual(blocks[0].structure_role, "table_column_header")
        self.assertEqual(blocks[1].structure_role, "table_column_header")
        self.assertEqual(blocks[2].structure_role, "table_row_label")
        self.assertEqual(blocks[3].structure_role, "table_data_cell")

    def test_split_lead_label_requires_real_heading_line(self) -> None:
        self.assertIsNone(g.split_lead_label_text("A\nbody text that is long enough to otherwise split " * 3))

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
            ["The spacing is different, pdf has $0.34 while word has $    0.34."],
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

    def test_proofread_mode_emits_formatting_comment_when_text_matches(self) -> None:
        doc = g.Block(
            id="d",
            source="docx",
            order=1,
            text="Results Summary",
            normalized=g.normalize_for_compare("Results Summary"),
            bold=True,
        )
        pdf = g.Block(
            id="p",
            source="pdf",
            order=1,
            text="Results Summary",
            normalized=g.normalize_for_compare("Results Summary"),
            bold=False,
        )
        comments = g.text_difference_comments(
            doc,
            pdf,
            1.0,
            target_name="pdf",
            docx_blocks=[doc],
            target_blocks=[pdf],
            formatting_diffs=["DOCX has bold; HTML does not have it."],
            proofread_mode=True,
        )
        self.assertEqual(
            [comment.contents for comment in comments],
            ["Formatting differs: DOCX has bold; HTML does not have it."],
        )

    def test_html_proofread_mode_emits_formatting_comment_when_text_matches(self) -> None:
        doc = g.Block(
            id="d",
            source="docx",
            order=1,
            text="Results Summary",
            normalized=g.normalize_for_compare("Results Summary"),
            underline=True,
        )
        html = g.Block(
            id="h",
            source="html",
            order=1,
            text="Results Summary",
            normalized=g.normalize_for_compare("Results Summary"),
            underline=False,
        )
        comments = g.text_difference_comments(
            doc,
            html,
            1.0,
            target_name="html",
            docx_blocks=[doc],
            target_blocks=[html],
            formatting_diffs=["DOCX has underline; HTML does not have it."],
            proofread_mode=True,
        )
        self.assertEqual(
            [comment.contents for comment in comments],
            ["Formatting differs: DOCX has underline; HTML does not have it."],
        )

    def test_run_level_underline_treats_hyperlink_as_visual_equivalent(self) -> None:
        doc = g.Block(
            id="d",
            source="docx",
            order=1,
            text="www.synopsys.com",
            normalized=g.normalize_for_compare("www.synopsys.com"),
            runs=[
                g.InlineRun(
                    text="www.synopsys.com",
                    kind="text",
                    hyperlink=True,
                    source_index=0,
                )
            ],
        )
        html = g.Block(
            id="h",
            source="html",
            order=1,
            text="www.synopsys.com",
            normalized=g.normalize_for_compare("www.synopsys.com"),
            runs=[
                g.InlineRun(
                    text="www.synopsys.com",
                    kind="text",
                    underline=True,
                    hyperlink=True,
                    source_index=0,
                )
            ],
        )
        self.assertEqual(g.compare_inline_formatting_diffs(doc, html), [])

    def test_run_level_underline_diff_uses_specific_excerpt(self) -> None:
        doc = g.Block(
            id="d",
            source="docx",
            order=1,
            text="Results Summary",
            normalized=g.normalize_for_compare("Results Summary"),
            runs=[g.InlineRun(text="Results Summary", kind="text", source_index=0)],
        )
        html = g.Block(
            id="h",
            source="html",
            order=1,
            text="Results Summary",
            normalized=g.normalize_for_compare("Results Summary"),
            runs=[g.InlineRun(text="Results Summary", kind="text", underline=True, source_index=0)],
        )
        self.assertEqual(
            g.compare_inline_formatting_diffs(doc, html),
            ['DOCX does not have underline on "Results Summary"; HTML has it.'],
        )

    def test_select_best_exact_candidate_prefers_formatting_compatible_title(self) -> None:
        doc = g.Block(
            id="d",
            source="docx",
            order=1,
            text="Reconciliation of First Quarter Fiscal Year 2026 Results",
            normalized=g.normalize_for_compare("Reconciliation of First Quarter Fiscal Year 2026 Results"),
            bold=True,
            runs=[g.InlineRun(text="Reconciliation of First Quarter Fiscal Year 2026 Results", kind="text", bold=True)],
        )
        html_plain = g.Block(
            id="h1",
            source="html",
            order=1,
            text="Reconciliation of First Quarter Fiscal Year 2026 Results",
            normalized=g.normalize_for_compare("Reconciliation of First Quarter Fiscal Year 2026 Results"),
            bold=False,
            runs=[g.InlineRun(text="Reconciliation of First Quarter Fiscal Year 2026 Results", kind="text")],
        )
        html_bold = g.Block(
            id="h2",
            source="html",
            order=2,
            text="Reconciliation of First Quarter Fiscal Year 2026 Results",
            normalized=g.normalize_for_compare("Reconciliation of First Quarter Fiscal Year 2026 Results"),
            bold=True,
            runs=[g.InlineRun(text="Reconciliation of First Quarter Fiscal Year 2026 Results", kind="text", bold=True)],
        )
        selected = g.select_best_exact_candidate(doc, [0, 1], [html_plain, html_bold], set())
        self.assertEqual(selected, 1)

    def test_select_best_exact_candidate_returns_none_for_ambiguous_repeated_label(self) -> None:
        doc = g.Block(
            id="d",
            source="docx",
            order=1,
            text="SYNOPSYS, INC.",
            normalized=g.normalize_for_compare("SYNOPSYS, INC."),
            kind="p",
        )
        html_a = g.Block(
            id="h1",
            source="html",
            order=1,
            text="SYNOPSYS, INC.",
            normalized=g.normalize_for_compare("SYNOPSYS, INC."),
            kind="p",
        )
        html_b = g.Block(
            id="h2",
            source="html",
            order=2,
            text="SYNOPSYS, INC.",
            normalized=g.normalize_for_compare("SYNOPSYS, INC."),
            kind="p",
        )
        selected = g.select_best_exact_candidate(doc, [0, 1], [html_a, html_b], set())
        self.assertIsNone(selected)

    def test_select_best_exact_candidate_rejects_wrong_structural_role(self) -> None:
        doc = g.Block(
            id="d",
            source="docx",
            order=1,
            text="Low",
            normalized=g.normalize_for_compare("Low"),
            table_cell=True,
            kind="td",
            table_pos=(0, 0, 0),
            structure_role="table_column_header",
            family_table_index=0,
        )
        html_wrong = g.Block(
            id="h1",
            source="html",
            order=1,
            text="Low",
            normalized=g.normalize_for_compare("Low"),
            table_cell=False,
            kind="p",
            structure_role="paragraph",
        )
        selected = g.select_best_exact_candidate(doc, [0], [html_wrong], set())
        self.assertIsNone(selected)

    def test_html_proofread_mode_catches_internal_spacing_difference(self) -> None:
        doc = g.Block(
            id="d",
            source="docx",
            order=1,
            text="Net income was $    0.34 per share.",
            normalized=g.normalize_for_compare("Net income was $    0.34 per share."),
        )
        html = g.Block(
            id="h",
            source="html",
            order=1,
            text="Net income was $0.34 per share.",
            normalized=g.normalize_for_compare("Net income was $0.34 per share."),
        )
        comments = g.text_difference_comments(
            doc,
            html,
            1.0,
            target_name="html",
            docx_blocks=[doc],
            target_blocks=[html],
            proofread_mode=True,
        )
        self.assertEqual(
            [comment.contents for comment in comments],
            ["The spacing is different, html has $0.34 while word has $    0.34."],
        )

    def test_html_proofread_mode_catches_word_spacing_between_equal_tokens(self) -> None:
        doc = g.Block(
            id="d",
            source="docx",
            order=1,
            text="Forward  looking statements",
            normalized=g.normalize_for_compare("Forward  looking statements"),
        )
        html = g.Block(
            id="h",
            source="html",
            order=1,
            text="Forward looking statements",
            normalized=g.normalize_for_compare("Forward looking statements"),
        )
        comments = g.text_difference_comments(
            doc,
            html,
            1.0,
            target_name="html",
            docx_blocks=[doc],
            target_blocks=[html],
            proofread_mode=True,
        )
        self.assertEqual(
            [comment.contents for comment in comments],
            ["The spacing is different, html has 1 space between Forward and looking while word has 2 spaces."],
        )

    def test_html_proofread_mode_ignores_layout_spacing_in_table_cells(self) -> None:
        doc = self._table_block(
            "$ 2,225",
            source="docx",
            row_key="revenue",
            row_slot=1,
            numeric_slot=0,
            table_pos=(1, 1, 1),
        )
        html = self._table_block(
            "$              2,225",
            source="html",
            row_key="revenue",
            row_slot=1,
            numeric_slot=0,
            table_pos=(1, 1, 1),
        )
        comments = g.text_difference_comments(
            doc,
            html,
            1.0,
            target_name="html",
            docx_blocks=[doc],
            target_blocks=[html],
            proofread_mode=True,
        )
        self.assertEqual(comments, [])

    def test_html_proofread_mode_does_not_fall_back_to_paragraph_comment_for_same_table_value(self) -> None:
        doc = self._table_block(
            "$\t5,690,000",
            source="docx",
            row_key="target non-gaap expenses",
            row_slot=1,
            numeric_slot=0,
            table_pos=(4, 13, 2),
        )
        html = self._table_block(
            "$          5,690,000",
            source="html",
            row_key="target non-gaap expenses",
            row_slot=1,
            numeric_slot=0,
            table_pos=(4, 13, 2),
        )
        comments = g.text_difference_comments(
            doc,
            html,
            0.95,
            target_name="html",
            docx_blocks=[doc],
            target_blocks=[html],
            proofread_mode=True,
        )
        self.assertEqual(comments, [])

    def test_html_proofread_mode_does_not_fall_back_for_parenthesized_negative_table_value(self) -> None:
        doc = self._table_block(
            "$\t(490)",
            source="docx",
            row_key="non-gaap interest and other income (expense), net",
            row_slot=4,
            numeric_slot=3,
            table_pos=(0, 9, 5),
        )
        html = self._table_block(
            "$                (490)",
            source="html",
            row_key="non-gaap interest and other income (expense), net",
            row_slot=4,
            numeric_slot=3,
            table_pos=(0, 9, 5),
        )
        comments = g.text_difference_comments(
            doc,
            html,
            0.95,
            target_name="html",
            docx_blocks=[doc],
            target_blocks=[html],
            proofread_mode=True,
        )
        self.assertEqual(comments, [])

    def test_html_proofread_mode_catches_missing_percent_sign(self) -> None:
        doc = self._table_block(
            "(2.1)",
            source="docx",
            row_key="restructuring charges",
            row_slot=1,
            numeric_slot=0,
            table_pos=(1, 2, 1),
        )
        html = self._table_block(
            "(2.1) %",
            source="html",
            row_key="restructuring charges",
            row_slot=1,
            numeric_slot=0,
            table_pos=(1, 2, 1),
        )
        comments = g.text_difference_comments(
            doc,
            html,
            1.0,
            target_name="html",
            docx_blocks=[doc],
            target_blocks=[html],
            proofread_mode=True,
        )
        self.assertEqual(
            [comment.contents for comment in comments],
            ["The percent sign is different, html has (2.1)% while word has (2.1)."],
        )


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
