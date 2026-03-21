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
                ? MarkdownTransformSourceSpanHelper.CreateBlockFingerprints(current.Blocks)
                : null;
            current = transform.Transform(current, context) ?? current;
            if (context.Diagnostics == null) {
                continue;
            }

            string[] afterFingerprints = MarkdownTransformSourceSpanHelper.CreateBlockFingerprints(current.Blocks);
            var change = MarkdownTransformSourceSpanHelper.ComputeChangedRange(beforeFingerprints!, afterFingerprints);
            MarkdownSourceSpan? affectedSourceSpan = MarkdownTransformSourceSpanHelper.AggregateSpans(blockSpans, change.StartBefore, change.CountBefore);
            blockSpans = MarkdownTransformSourceSpanHelper.UpdateBlockSpans(blockSpans, current.Blocks.Count, change, affectedSourceSpan);

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
        var children = syntaxTree?.Children?
            .Where(static child => child.AssociatedObject is IMarkdownBlock)
            .ToList();
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
}
