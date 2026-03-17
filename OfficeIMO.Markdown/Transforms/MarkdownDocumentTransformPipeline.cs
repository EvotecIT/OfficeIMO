namespace OfficeIMO.Markdown;

/// <summary>
/// Applies ordered post-parse transforms to typed markdown documents.
/// </summary>
/// <example>
/// <code>
/// var options = MarkdownReaderOptions.CreatePortableProfile();
/// options.DocumentTransforms.Add(
///     new MarkdownJsonVisualCodeBlockTransform(MarkdownVisualFenceLanguageMode.GenericSemanticFence));
///
/// var document = MarkdownReader.Parse(markdown, options);
/// </code>
/// Use document transforms for AST-level upgrades after markdown is parseable.
/// Keep malformed-input repair in <see cref="MarkdownInputNormalizer"/> so the parser sees valid structure first.
/// </example>
public static class MarkdownDocumentTransformPipeline {
    /// <summary>
    /// Applies the supplied transforms in order.
    /// </summary>
    /// <param name="document">Document to transform.</param>
    /// <param name="transforms">Ordered transforms.</param>
    /// <param name="context">Execution context.</param>
    /// <returns>The final transformed document.</returns>
    public static MarkdownDoc Apply(
        MarkdownDoc document,
        IEnumerable<IMarkdownDocumentTransform>? transforms,
        MarkdownDocumentTransformContext context) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        if (context == null) {
            throw new ArgumentNullException(nameof(context));
        }

        if (transforms == null) {
            return document;
        }

        var current = document;
        List<MarkdownSourceSpan?>? blockSpans = context.Diagnostics != null
            ? InitializeBlockSpans(current.Blocks, context.TopLevelBlockSourceSpans, context.SyntaxTree)
            : null;
        foreach (var transform in transforms) {
            if (transform == null) {
                continue;
            }

            int beforeCount = current.Blocks.Count;
            var beforeRef = current;
            string[]? beforeFingerprints = context.Diagnostics != null
                ? CreateTopLevelFingerprints(current.Blocks)
                : null;
            current = transform.Transform(current, context) ?? current;
            if (context.Diagnostics == null) {
                continue;
            }

            string[] afterFingerprints = CreateTopLevelFingerprints(current.Blocks);
            var change = ComputeChangedRange(beforeFingerprints!, afterFingerprints);
            MarkdownSourceSpan? affectedSourceSpan = AggregateSpans(blockSpans, change.StartBefore, change.CountBefore);
            blockSpans = UpdateBlockSpans(blockSpans, current.Blocks.Count, change, affectedSourceSpan);

            context.Diagnostics.Add(new MarkdownDocumentTransformDiagnostic {
                Source = context.Source,
                TransformName = transform.GetType().FullName ?? transform.GetType().Name,
                BlockCountBefore = beforeCount,
                BlockCountAfter = current.Blocks.Count,
                ReplacedDocument = !ReferenceEquals(beforeRef, current),
                ChangedBlockStartBefore = change.StartBefore,
                ChangedBlockCountBefore = change.CountBefore,
                ChangedBlockStartAfter = change.StartAfter,
                ChangedBlockCountAfter = change.CountAfter,
                AffectedSourceSpan = affectedSourceSpan
            });
        }

        return current;
    }

    private static List<MarkdownSourceSpan?> InitializeBlockSpans(
        IReadOnlyList<IMarkdownBlock> blocks,
        IReadOnlyList<MarkdownSourceSpan?>? topLevelBlockSourceSpans,
        MarkdownSyntaxNode? syntaxTree) {
        if (topLevelBlockSourceSpans != null && topLevelBlockSourceSpans.Count == blocks.Count) {
            return new List<MarkdownSourceSpan?>(topLevelBlockSourceSpans);
        }

        var spans = new List<MarkdownSourceSpan?>(blocks.Count);
        var children = syntaxTree?.Children;
        if (children == null || children.Count != blocks.Count) {
            for (var i = 0; i < blocks.Count; i++) {
                spans.Add(null);
            }

            return spans;
        }

        for (var i = 0; i < blocks.Count; i++) {
            spans.Add(children[i].SourceSpan);
        }

        return spans;
    }

    private static string[] CreateTopLevelFingerprints(IReadOnlyList<IMarkdownBlock> blocks) {
        var fingerprints = new string[blocks.Count];
        for (var i = 0; i < blocks.Count; i++) {
            var block = blocks[i];
            fingerprints[i] = block == null
                ? string.Empty
                : (block.GetType().FullName ?? block.GetType().Name) + "\n" + block.RenderMarkdown();
        }

        return fingerprints;
    }

    private static (int StartBefore, int CountBefore, int StartAfter, int CountAfter, int SuffixCount) ComputeChangedRange(
        string[] before,
        string[] after) {
        var prefix = 0;
        while (prefix < before.Length
               && prefix < after.Length
               && string.Equals(before[prefix], after[prefix], StringComparison.Ordinal)) {
            prefix++;
        }

        var suffix = 0;
        while (suffix < before.Length - prefix
               && suffix < after.Length - prefix
               && string.Equals(before[before.Length - 1 - suffix], after[after.Length - 1 - suffix], StringComparison.Ordinal)) {
            suffix++;
        }

        return (
            StartBefore: prefix,
            CountBefore: before.Length - prefix - suffix,
            StartAfter: prefix,
            CountAfter: after.Length - prefix - suffix,
            SuffixCount: suffix);
    }

    private static MarkdownSourceSpan? AggregateSpans(
        IReadOnlyList<MarkdownSourceSpan?>? spans,
        int start,
        int count) {
        if (spans == null || count <= 0 || start < 0 || start >= spans.Count) {
            return null;
        }

        int? first = null;
        int? last = null;
        var endExclusive = Math.Min(spans.Count, start + count);
        for (var i = start; i < endExclusive; i++) {
            if (!spans[i].HasValue) {
                continue;
            }

            first = first.HasValue ? Math.Min(first.Value, spans[i]!.Value.StartLine) : spans[i]!.Value.StartLine;
            last = last.HasValue ? Math.Max(last.Value, spans[i]!.Value.EndLine) : spans[i]!.Value.EndLine;
        }

        return first.HasValue && last.HasValue
            ? new MarkdownSourceSpan(first.Value, last.Value)
            : null;
    }

    private static List<MarkdownSourceSpan?> UpdateBlockSpans(
        IReadOnlyList<MarkdownSourceSpan?>? previousSpans,
        int newCount,
        (int StartBefore, int CountBefore, int StartAfter, int CountAfter, int SuffixCount) change,
        MarkdownSourceSpan? affectedSourceSpan) {
        var updated = new List<MarkdownSourceSpan?>(newCount);
        if (previousSpans == null || previousSpans.Count == 0) {
            for (var i = 0; i < newCount; i++) {
                updated.Add(null);
            }

            return updated;
        }

        var prefixCount = Math.Min(change.StartBefore, previousSpans.Count);
        for (var i = 0; i < prefixCount; i++) {
            updated.Add(previousSpans[i]);
        }

        for (var i = 0; i < change.CountAfter; i++) {
            updated.Add(affectedSourceSpan);
        }

        var suffixStart = Math.Max(prefixCount, previousSpans.Count - change.SuffixCount);
        for (var i = suffixStart; i < previousSpans.Count; i++) {
            updated.Add(previousSpans[i]);
        }

        while (updated.Count < newCount) {
            updated.Add(null);
        }

        if (updated.Count > newCount) {
            updated.RemoveRange(newCount, updated.Count - newCount);
        }

        return updated;
    }
}
