# OfficeIMO.Markdown lossless round-trip design

OfficeIMO exposes two complementary Markdown models:

- `MarkdownDoc` is the semantic editing model used by renderers, transforms, converters, and canonical writing.
- `MarkdownParseResult.FinalSyntaxTree` is the source-oriented structural model used for spans, tokens, trivia, navigation, and source edits.

Lossless workflows preserve source text that the caller does not intentionally change. Semantic `ToMarkdown()` output is a normalized document representation and is not presented as byte-preserving output.

## Current public contract

- `MarkdownReaderOptions.PreserveTrivia` retains the raw reader input as `MarkdownParseResult.OriginalMarkdown`.
- Syntax nodes and native projections expose normalized-source spans and source fields for the supported tokens.
- Original-source slices are available when raw and normalized text are identical or differ only by supported line-ending normalization.
- `MarkdownRoundtripWriter.WriteUnchanged` returns preserved original Markdown byte-for-byte when the parsed document has not changed.
- `MarkdownRoundtripWriter.WriteWithSourceEdit` and `WriteWithSourceEdits` apply non-overlapping native span-backed edits when every edit maps safely to original input.
- Unsafe mapping, missing trivia, overlapping edits, and transformed documents fall back explicitly with `MarkdownRoundtripDiagnostic` evidence.
- Diagnostics carry the most precise known source span and related spans so editor hosts can identify the affected source.

## Ownership rules

- Syntax nodes own exact structural spelling: delimiters, markers, indentation, line endings, and other trivia.
- Semantic blocks and inlines own meaning and typed editing behavior.
- `AssociatedObject` links semantic objects to source syntax without making the semantic tree a trivia store.
- Parsed nodes retain source-backed fields. Generated or transformed nodes are never assigned invented original-source spans.
- A format adapter may consume semantic content, but it must not rescan raw Markdown to rediscover behavior already represented by the semantic or syntax models.

## Source model

The source model supports:

- normalized offset, line, and column spans;
- original line-ending-aware slices for LF, CRLF, and standalone CR input;
- tab-expanded visual-column mapping;
- document-level trivia for empty and whitespace-only lines, leading/trailing horizontal whitespace, tabs, and line endings;
- token/source fields for headings, fences, links, images, autolinks, code spans, formatting delimiters, breaks, entities, callouts, details blocks, lists, tables, footnotes, and supported raw HTML structures;
- source-order enumeration, position lookup, native snapshots, and explicit source edits.

The public source contract is deliberately field-bounded. The [compatibility matrix](officeimo.markdown.compatibility-matrix.md) records which syntax families and fields own delimiters and trivia. OfficeIMO describes a location as exact only for a mapped source-backed field; it does not infer locations for generated nodes or claim lossless arbitrary semantic edits. A supported field is not considered source-backed until its semantic association, syntax token, normalized/original mapping, source edit, and reparse behavior are all exercised together.

## Round-trip decision order

1. An unchanged parse result with preserved trivia writes the original input.
2. Explicit source edits are ordered and validated against normalized spans.
3. Every edit is mapped to original input; an unsafe or ambiguous mapping stops the source-edit path.
4. Valid non-overlapping edits are applied without regenerating unrelated source.
5. A transformed or generally mutated semantic tree uses semantic Markdown generation and reports why byte preservation was unavailable.

This keeps byte preservation narrow and truthful. It does not infer changed-node identity by diffing rendered Markdown.

## Semantic invariants

- Footnote definitions use structured child blocks as the primary content model.
- List items use block children as their structural content.
- Definition lists expose one typed term/definition structure.
- Tables expose canonical row/cell content rather than parallel independently mutable views.
- Callouts, details blocks, links, images, fenced blocks, and formatting wrappers keep semantic ownership separate from marker/source ownership.
- Rebuilt syntax invalidates stale child projections so syntax-to-semantic associations describe the current typed tree.

## Editor-host guidance

Use source edits for focused text/token changes when retaining author spelling matters. Use semantic mutation and `ToMarkdown()` when normalized output is acceptable. An editor should surface round-trip diagnostics rather than silently rewriting an entire document after a source-preservation fallback.

For a focused source edit:

```csharp
MarkdownParseResult parsed = MarkdownReader.ParseWithSyntaxTree(
    source,
    new MarkdownReaderOptions { PreserveTrivia = true });

MarkdownRoundtripResult result = MarkdownRoundtripWriter.WriteWithSourceEdit(
    parsed,
    edit);

if (!result.IsLossless) {
    foreach (MarkdownRoundtripDiagnostic diagnostic in result.Diagnostics) {
        Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
    }
}
```

## Complete bounded contract

- Delimiter and trivia ownership is complete for the documented source-backed fields. Unlisted fields and intentionally unsupported optional syntax do not acquire inferred locations.
- Original-to-normalized mapping succeeds only when it is exact or line-ending-equivalent. Tabs use the documented visual-column model; transformed or generated nodes report an unavailable reason instead of an invented original span.
- Arbitrary semantic tree edits do not preserve unrelated bytes through a general changed-node writer.
- Multi-node and generated-node below-block diffing is intentionally not inferred from output text.
- `ToMarkdown()` remains semantic generation; it is not an alias for the round-trip writer.
