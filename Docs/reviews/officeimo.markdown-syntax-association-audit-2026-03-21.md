# OfficeIMO Markdown Syntax Association Audit

Date: 2026-03-21

Scope:

- `OfficeIMO.Markdown`
- syntax-node builders
- `AssociatedObject` coverage
- object-level source-span mapping implications

This was a read-only audit.

## Goal

Check whether semantic `MarkdownObject` instances that participate in the object tree can reliably receive `SourceSpan` values through syntax-tree association.

This matters because `MarkdownObjectTreeBinder` maps spans only through syntax nodes whose `AssociatedObject` is a `MarkdownObject`.

## Key Mechanism

Current binder behavior:

- `MarkdownObjectTreeBinder.MapSourceSpans(...)` assigns `SourceSpan` only when `syntaxNode.AssociatedObject is MarkdownObject`
- child-object traversal includes semantic children such as list items, table cells, definition-list groups/definitions, and inline container wrappers

Implication:

- if a semantic object is in the bound tree but no syntax node points back to it, its `SourceSpan` remains unset even when nearby syntax spans exist

## Confirmed Good

### 1. `SemanticFencedBlock` root association is present

Verified:

- the block root syntax node is associated with `this`

Reference:

- `OfficeIMO.Markdown/Blocks/SemanticFencedBlock.cs:121`

Notes:

- the earlier concern about the root node not associating back to the semantic block appears fixed
- helper child nodes such as `FenceSemanticKind`, `CodeFenceInfo`, and `CodeContent` are syntax-only, which is acceptable unless they later become semantic objects

### 2. `FootnoteDefinitionBlock` root association is present

Verified:

- the footnote-definition root syntax node is associated with `this`

Reference:

- `OfficeIMO.Markdown/Blocks/FootnoteDefinitionBlock.cs:192`

Notes:

- the earlier concern about missing root association also appears fixed here

## Confirmed Gaps

### Finding A: `CalloutBlock.TitleInlines` has no dedicated syntax association

Problem:

- `MarkdownObjectTreeBinder` treats `CalloutBlock.TitleInlines` as a child object
- `CalloutBlock.BuildSyntaxNode(...)` serializes the title into the block literal, but does not emit a child syntax node associated with `TitleInlines`

Impact:

- the callout title inline sequence can remain without object-level `SourceSpan`
- source mapping for callout body versus callout title is less precise than it should be

References:

- `OfficeIMO.Markdown/Core/MarkdownObjectTreeBinder.cs:49`
- `OfficeIMO.Markdown/Blocks/CalloutBlock.cs:15`
- `OfficeIMO.Markdown/Blocks/CalloutBlock.cs:174`

Recommendation:

- emit a dedicated title child syntax node, likely via `BuildInlineContainerNode(...)`
- keep the root node associated with `CalloutBlock`
- ensure the title child node associates with `TitleInlines`

### Finding B: list-item paragraph blocks can lose object-level span association

Problem:

- `MarkdownObjectTreeBinder` treats `ListItem.BlockChildren` as child semantic objects
- `ListItem.BuildOwnedSyntaxChildren()` special-cases paragraph blocks and calls `BuildParagraphSyntaxNode(paragraph.Inlines)`
- `BuildParagraphSyntaxNode(...)` uses `MarkdownBlockSyntaxBuilder.BuildInlineContainerNode(...)`
- that helper associates the node to the `InlineSequence`, not to the `ParagraphBlock`

Impact:

- paragraph blocks owned by list items may not receive `SourceSpan`
- the inline sequence gets the span instead of the paragraph block object

References:

- `OfficeIMO.Markdown/Core/MarkdownObjectTreeBinder.cs:76`
- `OfficeIMO.Markdown/Blocks/ListItem.cs:170`
- `OfficeIMO.Markdown/Blocks/ListItem.cs:228`
- `OfficeIMO.Markdown/Reader/Syntax/MarkdownBlockSyntaxBuilder.cs:30`

Recommendation:

- stop bypassing `ParagraphBlock.BuildSyntaxNode(...)` for owned list-item paragraph blocks, or
- add a helper that builds a paragraph syntax node associated with the `ParagraphBlock` rather than only the `InlineSequence`

### Finding C: definition-list semantic subobjects are only partially represented in syntax

Problem:

- binder traversal includes `DefinitionListGroup` and `DefinitionListDefinition`
- `DefinitionListBlock.BuildSyntaxItems()` creates `DefinitionGroup` and `DefinitionValue` syntax nodes without `AssociatedObject`
- term nodes are associated to the term `InlineSequence`, which is good
- group and definition semantic objects themselves do not appear to receive object-level spans from syntax nodes

Impact:

- grouped definition-list semantics are weaker than the object model suggests
- transforms and diagnostics cannot rely on `SourceSpan` at the group/definition level

References:

- `OfficeIMO.Markdown/Core/MarkdownObjectTreeBinder.cs:108`
- `OfficeIMO.Markdown/Core/MarkdownObjectTreeBinder.cs:114`
- `OfficeIMO.Markdown/Blocks/DefinitionListBlock.cs:234`
- `OfficeIMO.Markdown/Blocks/DefinitionListBlock.cs:270`
- `OfficeIMO.Markdown/Blocks/DefinitionListBlock.cs:277`

Recommendation:

- associate `DefinitionGroup` syntax nodes with `DefinitionListGroup`
- associate `DefinitionValue` syntax nodes with `DefinitionListDefinition`
- decide whether `DefinitionListEntry` remains a compatibility view only or becomes part of the canonical syntax/semantic mapping story

### Finding D: sequence-style inline wrappers are mapped to nested `InlineSequence`, not to the wrapper object

Problem:

- `BoldSequenceInline`, `ItalicSequenceInline`, `BoldItalicSequenceInline`, `StrikethroughSequenceInline`, and `HighlightSequenceInline` are semantic inline wrapper objects
- the inline syntax builder routes them through `BuildContainerNode(...)`
- `BuildContainerNode(...)` associates the node with the nested `InlineSequence`, not with the wrapper object

Impact:

- wrapper objects can remain without `SourceSpan`
- object-level source mapping for these semantic inline wrappers is weaker than for their simple text-only counterparts like `BoldInline`

References:

- `OfficeIMO.Markdown/Reader/Syntax/MarkdownInlineSyntaxBuilder.cs:64`
- `OfficeIMO.Markdown/Reader/Syntax/MarkdownInlineSyntaxBuilder.cs:68`
- `OfficeIMO.Markdown/Reader/Syntax/MarkdownInlineSyntaxBuilder.cs:72`
- `OfficeIMO.Markdown/Reader/Syntax/MarkdownInlineSyntaxBuilder.cs:76`
- `OfficeIMO.Markdown/Reader/Syntax/MarkdownInlineSyntaxBuilder.cs:80`
- `OfficeIMO.Markdown/Reader/Syntax/MarkdownInlineSyntaxBuilder.cs:109`

Recommendation:

- add wrapper-aware container builders that associate the syntax node to the semantic wrapper object
- keep the nested inline sequence as child syntax content, not as the associated object for the parent wrapper node

## Probably Intentional Syntax-Only Nodes

These do not look like bugs by themselves.

### 1. Nested list helper nodes

`MarkdownListSyntax` creates nested list syntax nodes without `AssociatedObject`.

Reference:

- `OfficeIMO.Markdown/Blocks/MarkdownListSyntax.cs:31`

This appears to reflect the current semantic model, which stores a flat item list with `Level` rather than nested semantic list objects.

Conclusion:

- acceptable for now
- but it is a sign that nested-list semantics are still more syntax-shaped than canonical-AST-shaped

### 2. Table row and header nodes

`TableHeader` and `TableRow` nodes are created without associated semantic row/header objects.

Reference:

- `OfficeIMO.Markdown/Blocks/TableBlock.cs:981`
- `OfficeIMO.Markdown/Blocks/TableBlock.cs:999`

This seems acceptable because the semantic model currently exposes `TableCell` objects, not row/header objects as `MarkdownObject` instances.

Conclusion:

- acceptable for now
- if row/header objects are introduced later, association will need to expand

## Suggested Follow-Up Order

1. Fix `CalloutBlock.TitleInlines` association.
2. Fix list-item paragraph-block association.
3. Fix sequence-inline wrapper association.
4. Fix definition-list group/definition association.
5. Add invariant tests that walk the object tree and assert span coverage for targeted semantic object kinds.

## Recommended Tests

Add focused regression tests for:

- callout title span mapping
- list-item paragraph-block span mapping
- nested emphasis/highlight wrapper span mapping
- definition-group and definition-body span mapping
- a representative mixed document that walks the semantic tree and asserts non-null spans for expected node kinds

## Bottom Line

The root-level association story is in better shape than the older notes suggested.

The main remaining weakness is not top-level blocks.
It is intermediate semantic objects:

- callout titles
- list-item paragraph blocks
- grouped definition-list objects
- sequence-style inline wrappers

Those are the places where object-level source mapping still falls short of the overall AST-first design.
