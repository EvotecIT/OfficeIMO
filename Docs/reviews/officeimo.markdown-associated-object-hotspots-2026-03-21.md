# OfficeIMO Markdown Associated-Object Hotspots (2026-03-21)

This note narrows the transient/misaligned associated-object problem to concrete builder paths.

It follows:

- `Docs/reviews/officeimo.markdown-syntax-association-audit-2026-03-21.md`
- `Docs/reviews/officeimo.markdown-tree-invariant-findings-2026-03-21.md`

The goal is not to add more formats or special cases. The goal is to make syntax-to-semantic mapping correct, stable, and intentional.

## Summary

Three builder paths stand out:

1. list-item paragraph syntax is associated to `InlineSequence` instead of `ParagraphBlock`
2. definition-list syntax builds semantic-looking wrapper nodes without stable semantic associations
3. sequence-style inline wrappers are associated to nested `InlineSequence` instead of the wrapper inline object

These are not cosmetic issues. They affect source-span ownership, AST navigation, and the meaning of `AssociatedObject`.

## Hotspot 1: `ListItem` paragraph ownership

Relevant code:

- `OfficeIMO.Markdown/Blocks/ListItem.cs`

Observed behavior:

- `BuildOwnedSyntaxChildren()` special-cases `ParagraphBlock`
- instead of building syntax from the `ParagraphBlock`, it calls `BuildParagraphSyntaxNode(paragraph.Inlines)`
- `BuildParagraphSyntaxNode(...)` delegates to `MarkdownBlockSyntaxBuilder.BuildInlineContainerNode(...)`
- `BuildInlineContainerNode(...)` sets `associatedObject: inlines`

Concrete references:

- `ListItem.BuildOwnedSyntaxChildren()` lines 163-179
- `ListItem.BuildParagraphSyntaxNode(...)` lines 228-232
- `MarkdownBlockSyntaxBuilder.BuildInlineContainerNode(...)` lines 30-42

Why this is a problem:

- the semantic object in the bound tree is `ParagraphBlock`
- the syntax node produced for that paragraph points at `InlineSequence`
- this means the paragraph block can lose direct syntax ownership even though it is the real semantic block node

Correctness direction:

- paragraph syntax nodes inside list items should associate to the `ParagraphBlock`
- the `InlineSequence` should remain a child semantic object, not the primary associated object for the paragraph block syntax node

Recommended fix shape:

1. add an overload or helper that builds an inline-container syntax node with an explicit associated object
2. change `ListItem.BuildParagraphSyntaxNode(...)` to accept the `ParagraphBlock`, not only its inlines
3. preserve the inline children exactly as today, but associate the paragraph syntax node to the `ParagraphBlock`

## Hotspot 2: `DefinitionListBlock` semantic wrappers vs syntax wrappers

Relevant code:

- `OfficeIMO.Markdown/Blocks/DefinitionListBlock.cs`

Observed behavior:

- `DefinitionList` root associates to `DefinitionListBlock`
- `DefinitionGroup` nodes are created without `associatedObject`
- `DefinitionValue` nodes are created without `associatedObject`
- term nodes are built through `BuildInlineContainerNode(...)`, which associates them to `InlineSequence`
- fallback inline recovery creates `new DefinitionListEntry(new InlineSequence(), definition)` purely to derive inline content

Concrete references:

- `BuildSyntaxItems()` lines 239-283
- term builder lines 243-248
- fallback entry lines 261-267
- `DefinitionValue` creation lines 270-274
- `DefinitionGroup` creation lines 277-280

Why this is a problem:

- the object tree already contains `DefinitionListGroup` and `DefinitionListDefinition`
- the syntax tree emits nodes that look like first-class semantic wrappers
- but those wrapper syntax nodes do not map back to the wrapper semantic objects
- the fallback path also creates temporary semantic-looking objects that should not leak into a stable AST contract

Correctness direction:

- `DefinitionGroup` syntax should associate to `DefinitionListGroup`
- `DefinitionValue` syntax should associate to `DefinitionListDefinition`
- fallback rendering should not require creating a throwaway `DefinitionListEntry`

Recommended fix shape:

1. add explicit `associatedObject` assignment for `DefinitionGroup` and `DefinitionValue`
2. avoid temporary `DefinitionListEntry` allocation in the fallback path
3. if term syntax is meant to map to `InlineSequence`, keep that explicit; if a dedicated semantic term wrapper is desired later, add one deliberately rather than by accident

## Hotspot 3: sequence inline wrappers

Relevant code:

- `OfficeIMO.Markdown/Reader/Syntax/MarkdownInlineSyntaxBuilder.cs`

Observed behavior:

- `BoldSequenceInline`, `ItalicSequenceInline`, `BoldItalicSequenceInline`, `StrikethroughSequenceInline`, and `HighlightSequenceInline` call `BuildContainerNode(...)`
- `BuildContainerNode(...)` builds children from the nested `InlineSequence`
- the produced syntax node uses `associatedObject: sequence`

Concrete references:

- wrapper dispatch lines 64-81
- `BuildContainerNode(...)` lines 109-112

Why this is a problem:

- the semantically meaningful object is the wrapper inline (`BoldSequenceInline`, etc.)
- the nested `InlineSequence` is structural content, not the owning emphasis/strong/highlight node
- span lookup can therefore land on the nested container instead of the actual wrapper object

Correctness direction:

- the syntax node for inline strong/emphasis/highlight wrappers should associate to the wrapper object
- the nested `InlineSequence` should continue to own child inline nodes only

Recommended fix shape:

1. change `BuildContainerNode(...)` to accept both the wrapper object and the nested sequence
2. associate the produced node to the wrapper object
3. leave child generation based on the nested sequence unchanged

## Shared design recommendation

The repeated root issue is that helper builders currently infer ownership from the nearest inline container instead of from the semantic node that the syntax node is supposed to represent.

That suggests a generic rule:

- container-building helpers should support explicit ownership
- when a syntax node corresponds to a semantic wrapper, `AssociatedObject` must be that wrapper
- nested `InlineSequence` objects should be associated only when the syntax node truly represents the sequence itself

## Proposed implementation order

1. `ListItem` paragraph ownership
   - lowest conceptual risk
   - directly improves block-level span ownership
2. sequence inline wrappers
   - localized helper refactor
   - high value for AST correctness
3. definition-list wrappers
   - slightly broader because it touches wrapper semantics and fallback recovery

## What to avoid

- do not patch this with post-bind remapping
- do not special-case source spans after the fact
- do not hide temporary-object creation behind another helper if the ownership model is still ambiguous

The right fix is to make syntax builders emit the correct owner at construction time.
