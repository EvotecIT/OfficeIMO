# OfficeIMO Markdown Tree Invariant Findings (2026-03-21)

This note captures what the new invariant helper work surfaced while turning A2 into executable tests.

Related files:

- `OfficeIMO.Tests/Markdown/MarkdownInvariantAssert.cs`
- `OfficeIMO.Tests/Markdown/Markdown_Tree_Invariant_Tests.cs`

## Findings

### 1. `FinalSyntaxTree` is the stable structural contract

The new syntax-tree invariant helper initially ran against `MarkdownParseResult.SyntaxTree` and failed on parent identity checks.

The cause is architectural, not a bad assertion:

- `ParseWithSyntaxTree(...)` builds an original syntax tree
- it then rebuilds a final syntax tree and binds the semantic document to that final tree
- some block builders appear to reuse syntax child nodes when rebuilding, which means those child nodes can end up parented by the final tree instead of the original one

Implication:

- identity-based tree invariants should be asserted against `MarkdownParseResult.FinalSyntaxTree`
- if the original tree is meant to preserve its own parent/child identity, the rebuild path needs to stop reusing mutable node instances across trees

### 2. Some syntax associations still point at detached semantic objects

The helper also found syntax nodes whose `AssociatedObject` is a `MarkdownObject` that is not attached to the final bound document tree (`Document == null`).

Implication:

- not every syntax association in the final tree currently points at a stable semantic node from the active object model
- the invariant helper therefore validates only associations whose `Document` is the current parse result document

Likely causes to audit next:

- builder paths that create temporary `InlineSequence` containers
- definition-list syntax construction that creates intermediate helper objects
- block rebuild paths that preserve spans but not canonical semantic ownership

Concrete hotspots seen during the follow-up read:

- `OfficeIMO.Markdown/Blocks/ListItem.cs`
  - `BuildOwnedSyntaxChildren()` special-cases `ParagraphBlock` and calls `BuildParagraphSyntaxNode(paragraph.Inlines)`
  - `BuildParagraphSyntaxNode(...)` then associates the syntax node to the `InlineSequence`, not the `ParagraphBlock`
- `OfficeIMO.Markdown/Blocks/DefinitionListBlock.cs`
  - `BuildSyntaxItems()` creates `DefinitionGroup` / `DefinitionValue` syntax nodes without semantic associations
  - the fallback path creates `new DefinitionListEntry(new InlineSequence(), definition)` only to derive paragraph inlines
- `OfficeIMO.Markdown/Reader/Syntax/MarkdownInlineSyntaxBuilder.cs`
  - `BuildContainerNode(...)` associates sequence-wrapper syntax nodes to the nested `InlineSequence`
  - this is relevant for `BoldSequenceInline`, `ItalicSequenceInline`, `BoldItalicSequenceInline`, `StrikethroughSequenceInline`, and `HighlightSequenceInline`

## Recommended next follow-up

Turn this into a focused audit item:

1. inventory `BuildSyntaxNode` and helper paths that create temporary `MarkdownObject` instances
2. decide which associated objects are part of the public/stable AST contract
3. stop attaching transient semantic objects to final-tree syntax nodes unless they are intentionally public
