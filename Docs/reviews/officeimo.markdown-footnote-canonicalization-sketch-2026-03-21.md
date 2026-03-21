# OfficeIMO Markdown Footnote Canonicalization Sketch (2026-03-21)

This note proposes `FootnoteDefinitionBlock` as the first B2 canonicalization target.

Why this node:

- it has real duplicated structure
- the scope is still local
- the existing tests already express useful identity guarantees we should preserve

Relevant files:

- `OfficeIMO.Markdown/Blocks/FootnoteDefinitionBlock.cs`
- `OfficeIMO.Tests/Markdown/Markdown_Reader_Refs_Footnotes_Tests.cs`
- `OfficeIMO.Tests/Markdown/Markdown_Document_Transform_Tests.cs`

## What is already good

The current type is not random. It already centralizes construction through `FootnoteContentViews`, and tests assert important identity reuse:

- `Blocks[0]` is the same instance as `ParagraphBlocks[0]` when the footnote is paragraph-only
- `ParagraphBlocks[i].Inlines` is the same instance as `Paragraphs[i]`
- `Text` is derived from blocks when block content exists

That means the class is a strong candidate for the “one primary structure, derived projections” pattern.

## Current duplication

The type currently stores all of these:

- `_fallbackText`
- `_blocks`
- `_paragraphs`
- `_paragraphBlocks`
- `SyntaxChildren`

The main ambiguity is that `Blocks`, `Paragraphs`, and `ParagraphBlocks` can all look structural, even though only one of them should really be primary.

## Proposed primary representation

Make `Blocks` the canonical semantic structure.

Reason:

- syntax and tree ownership are naturally block-oriented
- rendering already prefers blocks
- paragraph-only footnotes can still be represented as paragraph blocks without loss
- mixed-content footnotes fit naturally into a block-first model

## Proposed derived views

Keep these as derived projections:

- `ParagraphBlocks`
  - derived by filtering `Blocks` for `ParagraphBlock`
- `Paragraphs`
  - derived from `ParagraphBlocks.Select(p => p.Inlines)`
- `Text`
  - derived from `Blocks` when blocks exist
  - otherwise derived from fallback parsed text input

## What should remain as stored input

Keep a fallback text input only for the case where the footnote is created with no block structure at all.

That fallback should be treated as a construction fallback, not as a peer semantic representation.

## Refactor shape

1. Keep a single stored block list.
2. Keep fallback text only when there are no canonical blocks.
3. Recompute `ParagraphBlocks` and `Paragraphs` from canonical blocks.
4. Preserve identity reuse for paragraph-only cases:
   - if a canonical block is already a `ParagraphBlock`, reuse it
   - if a paragraph view is requested, return that same block and its `Inlines`
5. Leave `SyntaxChildren` behavior unchanged for the first pass unless it blocks simplification.

## Invariants to preserve

These are already implied by tests and should remain true after refactor:

1. `Text` reflects canonical blocks when blocks exist.
2. Paragraph-only footnotes reuse the same `ParagraphBlock` instances in `Blocks` and `ParagraphBlocks`.
3. `ParagraphBlocks[i].Inlines` is the same instance as `Paragraphs[i]`.
4. Non-paragraph blocks remain preserved in `Blocks` and are not flattened into text.

## Tests that should exist or be expanded

Existing useful tests:

- parsed paragraph reuse
- `Text` derived from blocks
- transform-updated paragraph content

Recommended additions when implementation starts:

1. mixed-content footnote keeps canonical non-paragraph blocks in `Blocks`
2. `ParagraphBlocks` filters only paragraph children from mixed-content `Blocks`
3. syntax-node association still points to `FootnoteDefinitionBlock` and direct block children after canonicalization

## What not to do

- do not replace `Blocks` with paragraph text plus reparsing
- do not preserve all three structures as independently stored peers
- do not solve drift with post-hoc synchronization helpers if a single primary structure can remove the drift entirely

## Success condition

After the refactor, a reader should be able to say:

- the footnote owns blocks
- paragraph helpers are projections over those blocks
- fallback text is only an input fallback

If that statement is not true, the canonicalization is incomplete.
