# HTML vs Word Proofread Engine Spec

## Purpose

Define the product requirements, comparison policy, architecture, and future hardening plan for the `DOCX vs HTML` proofreading engine.

This document is for the `Word -> HTML` workflow only. It is the strongest comparison path in this project and the primary path that should be treated as the proofreading baseline.

## Product Goal

The engine should behave like a proofreader, not a generic text diff.

That means it must:

- match the correct corresponding content before diffing
- catch real proofreading differences
- avoid false positives from weak or ambiguous matches
- explain broad structural mismatches clearly instead of pretending they are exact text differences
- remain stable across future earnings-release style documents

## Scope

### In Scope

- `DOCX vs HTML` comparison
- narrative paragraphs
- section headers and lead lines
- table titles and subtitles
- table column headers
- table row labels
- table data cells
- quote blocks
- contact blocks
- footnotes
- reconciliation sections
- long narrative/disclosure sections
- formatting comparison where span-level alignment is strong enough

### Out of Scope

- `DOCX vs PDF` parity with HTML behavior
- pixel-perfect visual comparison
- arbitrary layout/CSS rendering fidelity
- guaranteed formatting accuracy when source structure is ambiguous

## Core Requirements

### 1. Structure-First Matching

The engine must match corresponding content structurally before generating proofreading comments.

Matching must be aware of:

- section families
- table families
- header/title families
- grouped quote/contact/footnote families
- row lineage within aligned tables

### 2. Proofread-Level Difference Detection

The engine must catch real differences in:

- words
- numbers
- dates
- currency symbols
- percent signs
- punctuation
- meaningful spacing
- bold / italic / underline where visually relevant and span alignment is strong

### 3. False-Positive Suppression

The engine must suppress comments when the match is weak or ambiguous.

Common false-positive sources to avoid:

- wrong block pairing
- repeated short labels
- layout-only HTML spacing
- merged HTML blocks vs split Word blocks
- block-level formatting leakage from mixed-format paragraphs
- unmatched quote/contact/footnote spillover

### 4. Broad Structural Visibility

When the engine can identify a corresponding section family but cannot align it cleanly block-by-block, it should emit a structural summary instead of noisy low-confidence token comments.

### 5. Reusability

Behavior must be driven by durable comparison logic, not by one-off patches for specific documents.

## Comparison Policy

### Matching Policy

- Prefer exact structural matches where role, family, order, and local context agree.
- Allow broad structural matches when section identity is clear but fine-grained alignment is weak.
- Suppress detailed comments when correspondence is not defensible.

### Text Policy

- Show precise word-level differences when the match is strong.
- Show number/date/symbol differences when cell/span correspondence is strong.
- Do not emit generic paragraph-difference comments for table values when only layout spacing differs.

### Spacing Policy

- Meaningful spacing in prose can be reported.
- Layout-only spacing inside HTML tables should be suppressed.
- Mixed token + spacing comparison is allowed only after the correct pair is matched.

### Formatting Policy

- Compare formatting at run/span level, not whole-block level, where possible.
- Compare only the matched text span, not unrelated styled descendants.
- Be conservative when formatting alignment is weak.

### Summary Policy

- Long rewritten quote/footnote/narrative families may use summary comments instead of many token comments.
- Summary comments must say whether the difference is:
  - high-confidence textual drift
  - broad structural match
  - normalized-text-equivalent but structurally non-exact

## Structural Model

The engine should classify extracted content into typed roles instead of using one generic block type.

Current target roles:

- `paragraph`
- `section_header`
- `section_lead`
- `table_title`
- `table_subtitle`
- `table_column_header`
- `table_row_label`
- `table_data_label`
- `table_data_cell`
- `quote`
- `contact`
- `footnote`
- `reconciliation_family`
- `narrative_family`

## Current Architecture

### Extraction

The engine extracts ordered runs/spans instead of flattening content too early.

Each block may retain:

- raw text
- proof text
- match text
- inline runs
- formatting signals
- structural role
- table/family metadata

### Matching

The engine currently uses:

- table-family alignment before row/cell matching
- header/title family matching
- row-label lineage in aligned tables
- grouped matching for quote/contact/footnote sections
- family-aware matching for reconciliation and long narrative sections
- repeated-label disambiguation

### Comment Emission

The engine uses confidence-aware output:

- `High Confidence | Critical`
- `Medium Confidence | Text`
- `Medium Confidence | Formatting`
- `Medium Confidence | Structural`
- `Medium Confidence | Info`

Detailed comments should only be emitted when the match is strong enough.

## Known Good Behaviors

The `DOCX vs HTML` engine should currently support:

- matching tables by family before comparing repeated row names
- catching short header typos like `Targetss` vs `Targets`
- detecting word differences in short lead/header blocks such as `Investor` vs `Investors`
- keeping quote/contact/footnote/reconciliation sections visible as structural families
- suppressing layout-only table-spacing noise
- preventing repeated short labels from generating arbitrary false matches

## Remaining Hardening Work

The engine is strong enough for real proofreading use, but not fully final for all future document layouts.

Remaining work is mainly hardening, not redesign.

### 1. Regression Corpus Expansion

Build a larger reviewed corpus of real `DOCX vs HTML` pairs.

Recommended target:

- 10 to 20 real filing/release pairs
- both clean and messy structural cases
- expected valid comments documented
- expected suppressed comments documented

### 2. Document-Level Golden Tests

Add document-level regression tests, not only logic/unit tests.

Each golden test should lock:

- reviewed JSON output
- expected high-confidence findings
- expected broad summaries
- expected suppressed noise classes

### 3. Header/Section Validation Across Layouts

Expand validation for:

- centered table titles
- merged header cells
- repeated section titles
- unusual HTML wrappers
- multi-line title/subtitle stacks

### 4. Formatting Confidence Refinement

Keep formatting comments conservative unless:

- span match is strong
- structure role is compatible
- local neighborhood is stable

### 5. Reviewer Workflow Controls

Add output filtering or reviewer modes such as:

- critical only
- critical + text
- all findings

## Acceptance Criteria For “Final”

The `DOCX vs HTML` engine can be considered final only when:

- a reviewed regression corpus exists across multiple real document pairs
- document-level golden outputs are stable
- real number/date/symbol differences are consistently caught
- repeated-label/header noise remains suppressed across the corpus
- quote/contact/footnote/reconciliation families do not regress into orphan block noise
- formatting comments remain conservative and high-signal

## Recommended Product Positioning

### Current Positioning

The `DOCX vs HTML` engine is now strong enough for real proofreading with human review.

### Not Yet Safe To Claim

Do not claim that it is:

- perfect for every future filing layout
- pixel-accurate
- equal in reliability to a visual design-diff system

### Best Internal Positioning

Treat it as:

- a structure-aware proofreading engine
- optimized for real Word-to-HTML release comparison
- confidence-driven rather than noise-maximizing

## Next Practical Step

The next best step is not another logic patch.

It is to create:

- a real reviewed regression corpus
- document-level golden tests
- a locked proofreading policy

That is what will make the `DOCX vs HTML` workflow feel final for future documents.
