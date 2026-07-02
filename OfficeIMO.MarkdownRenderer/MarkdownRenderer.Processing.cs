using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

public static partial class MarkdownRenderer {
    private static void CopyDiagnostics<TDiagnostic>(IReadOnlyList<TDiagnostic> source, ICollection<TDiagnostic>? target) {
        if (target == null || source == null || source.Count == 0) {
            return;
        }

        for (var i = 0; i < source.Count; i++) {
            target.Add(source[i]);
        }
    }

    private static string PrepareMarkdown(
        string markdown,
        MarkdownRendererOptions options,
        bool renderErrorAsException,
        HtmlOptions? htmlOptions = null,
        ICollection<MarkdownRendererPreProcessorDiagnostic>? preProcessorDiagnostics = null) {
        markdown = ApplyPreParseProcessing(markdown, options, preProcessorDiagnostics);

        if (options.MaxMarkdownChars.HasValue && options.MaxMarkdownChars.Value >= 0 && markdown.Length > options.MaxMarkdownChars.Value) {
            int max = options.MaxMarkdownChars.Value;
            switch (options.MarkdownOverflowHandling) {
                case OverflowHandling.Throw:
                    throw new ArgumentOutOfRangeException(nameof(markdown), $"Markdown length {markdown.Length} exceeds MaxMarkdownChars {max}.");
                case OverflowHandling.RenderError:
                    if (renderErrorAsException || htmlOptions == null) {
                        throw new InvalidOperationException($"Content exceeded the maximum allowed Markdown length ({max} chars) and cannot be converted into a document.");
                    }

                    throw new MarkdownPreparationOverflowException(BuildOverflowBodyHtml(htmlOptions, $"Content exceeded the maximum allowed Markdown length ({max} chars)."));
                case OverflowHandling.Truncate:
                default:
                    markdown = markdown.Substring(0, max);
                    break;
            }
        }

        return markdown;
    }

    internal static string ApplyPreParseProcessing(
        string? markdown,
        MarkdownRendererOptions options,
        ICollection<MarkdownRendererPreProcessorDiagnostic>? diagnostics = null) {
        var value = markdown ?? string.Empty;

        if (options.NormalizeEscapedNewlines) {
            var before = value;
            value = value.Replace("\\r\\n", "\n").Replace("\\n", "\n");
            if (!string.Equals(before, value, StringComparison.Ordinal)) {
                diagnostics?.Add(new MarkdownRendererPreProcessorDiagnostic {
                    Stage = MarkdownRendererPreProcessorStage.EscapedNewlineNormalization,
                    LengthBefore = before.Length,
                    LengthAfter = value.Length,
                    Changed = true
                });
            }
        }

        return PreprocessMarkdown(value, options, diagnostics);
    }

    private sealed class MarkdownPreparationOverflowException(string overflowHtml) : Exception {
        public string OverflowHtml { get; } = overflowHtml ?? string.Empty;
    }

    private static MarkdownDoc ApplyRendererDocumentTransforms(
        MarkdownDoc document,
        MarkdownRendererOptions options,
        MarkdownReaderOptions readerOptions,
        ICollection<MarkdownDocumentTransformDiagnostic>? diagnostics,
        MarkdownSyntaxNode? syntaxTree = null,
        IReadOnlyList<MarkdownSourceSpan?>? topLevelBlockSourceSpans = null,
        string? sourceMarkdown = null,
        string? originalMarkdown = null,
        bool preservesOriginalMarkdown = false) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        var transforms = options.DocumentTransforms;
        if (transforms == null || transforms.Count == 0) {
            return document;
        }

        return MarkdownDocumentTransformPipeline.Apply(
            document,
            transforms,
            new MarkdownDocumentTransformContext(
                MarkdownDocumentTransformSource.MarkdownRenderer,
                readerOptions,
                options,
                diagnostics,
                syntaxTree,
                topLevelBlockSourceSpans,
                sourceMarkdown,
                originalMarkdown,
                preservesOriginalMarkdown));
    }

    private static MarkdownParseResult AttachRendererParseResult(
        MarkdownDoc document,
        MarkdownSyntaxNode syntaxTree,
        MarkdownSyntaxNode finalSyntaxTree,
        string sourceMarkdown,
        string originalMarkdown,
        bool preservesOriginalMarkdown,
        IReadOnlyList<MarkdownDocumentTransformDiagnostic> transformDiagnostics,
        IReadOnlyList<MarkdownReferenceLinkDefinition> referenceLinkDefinitions,
        IReadOnlyList<MarkdownAbbreviationDefinition> abbreviationDefinitions) =>
        new MarkdownParseResult(
            document,
            syntaxTree,
            finalSyntaxTree,
            sourceMarkdown,
            originalMarkdown,
            preservesOriginalMarkdown,
            transformDiagnostics,
            referenceLinkDefinitions,
            abbreviationDefinitions);

    private static IReadOnlyList<MarkdownSourceSpan?> BuildTopLevelBlockSourceSpans(MarkdownParseResult parseResult) {
        var children = parseResult.SyntaxTree.Children;
        var spans = new List<MarkdownSourceSpan?>(parseResult.Document.Blocks.Count);

        if (children.Count > 0) {
            var blockChildren = children.Where(static child => child.AssociatedObject is IMarkdownBlock).ToList();
            var topLevelBlocks = parseResult.Document.TopLevelBlocks;
            var childCount = Math.Min(blockChildren.Count, topLevelBlocks.Count);
            for (var i = 0; i < childCount; i++) {
                if (topLevelBlocks[i] is FrontMatterBlock) {
                    continue;
                }

                spans.Add(blockChildren[i].SourceSpan);
            }
        } else {
            for (var i = 0; i < parseResult.Document.Blocks.Count; i++) {
                spans.Add(null);
            }
        }

        for (var i = 0; i < parseResult.TransformDiagnostics.Count; i++) {
            var diagnostic = parseResult.TransformDiagnostics[i];
            spans = UpdateTopLevelBlockSourceSpans(spans, diagnostic);
        }

        while (spans.Count < parseResult.Document.Blocks.Count) {
            spans.Add(null);
        }

        if (spans.Count > parseResult.Document.Blocks.Count) {
            spans.RemoveRange(parseResult.Document.Blocks.Count, spans.Count - parseResult.Document.Blocks.Count);
        }

        return spans;
    }

    private static List<MarkdownSourceSpan?> UpdateTopLevelBlockSourceSpans(
        IReadOnlyList<MarkdownSourceSpan?> previous,
        MarkdownDocumentTransformDiagnostic diagnostic) {
        var updated = new List<MarkdownSourceSpan?>(diagnostic.BlockCountAfter);
        var prefixCount = Math.Min(diagnostic.ChangedBlockStartBefore, previous.Count);
        for (var i = 0; i < prefixCount; i++) {
            updated.Add(previous[i]);
        }

        for (var i = 0; i < diagnostic.ChangedBlockCountAfter; i++) {
            updated.Add(diagnostic.AffectedSourceSpan);
        }

        var suffixCount = previous.Count - diagnostic.ChangedBlockStartBefore - diagnostic.ChangedBlockCountBefore;
        var suffixStart = Math.Max(prefixCount, previous.Count - suffixCount);
        for (var i = suffixStart; i < previous.Count; i++) {
            updated.Add(previous[i]);
        }

        while (updated.Count < diagnostic.BlockCountAfter) {
            updated.Add(null);
        }

        if (updated.Count > diagnostic.BlockCountAfter) {
            updated.RemoveRange(diagnostic.BlockCountAfter, updated.Count - diagnostic.BlockCountAfter);
        }

        return updated;
    }

    private static string PreprocessMarkdown(
        string markdown,
        MarkdownRendererOptions options,
        ICollection<MarkdownRendererPreProcessorDiagnostic>? diagnostics = null) {
        var value = markdown ?? string.Empty;
        if (value.Length == 0) {
            return value;
        }

        var normalization = CreateEffectiveInputNormalization(options);
        var preParseNormalization = CreatePreParseNormalizationOptions(normalization);
        if (preParseNormalization != null) {
            var before = value;
            value = MarkdownInputNormalizer.Normalize(value, preParseNormalization);
            if (!string.Equals(before, value, StringComparison.Ordinal)) {
                diagnostics?.Add(new MarkdownRendererPreProcessorDiagnostic {
                    Stage = MarkdownRendererPreProcessorStage.InputNormalization,
                    LengthBefore = before.Length,
                    LengthAfter = value.Length,
                    Changed = true
                });
            }
        }

        var pre = options.MarkdownPreProcessors;
        if (pre != null && pre.Count > 0) {
            for (int i = 0; i < pre.Count; i++) {
                var processor = pre[i];
                if (processor == null) continue;
                var before = value;
                value = processor(value, options) ?? value ?? string.Empty;
                if (!string.Equals(before, value, StringComparison.Ordinal)) {
                    diagnostics?.Add(new MarkdownRendererPreProcessorDiagnostic {
                        Stage = MarkdownRendererPreProcessorStage.CustomPreProcessor,
                        ProcessorName = processor.Method.DeclaringType?.FullName is string typeName && !string.IsNullOrWhiteSpace(typeName)
                            ? typeName + "." + processor.Method.Name
                            : processor.Method.Name,
                        LengthBefore = before.Length,
                        LengthAfter = value.Length,
                        Changed = true
                    });
                }
            }
        }

        return value;
    }

    private static MarkdownReaderOptions CreateEffectiveReaderOptions(MarkdownRendererOptions options) {
        var source = options.ReaderOptions ?? new MarkdownReaderOptions();
        var normalization = CreateEffectiveInputNormalization(options);
        var effective = source.Clone();
        effective.InputNormalization = CreateInlineNormalizationOptions(normalization);

        effective.FencedBlockExtensions.Clear();
        AddRendererSemanticFencedBlockExtensions(effective, options);
        CopyFencedBlockExtensions(source, effective);
        return effective;
    }

    private static MarkdownInputNormalizationOptions CreateEffectiveInputNormalization(MarkdownRendererOptions options) {
        var source = options.ReaderOptions?.InputNormalization;
        return new MarkdownInputNormalizationOptions {
            NormalizeSoftWrappedStrongSpans = (source?.NormalizeSoftWrappedStrongSpans == true) || options.NormalizeSoftWrappedStrongSpans,
            NormalizeInlineCodeSpanLineBreaks = (source?.NormalizeInlineCodeSpanLineBreaks == true) || options.NormalizeInlineCodeSpanLineBreaks,
            NormalizeEscapedInlineCodeSpans = (source?.NormalizeEscapedInlineCodeSpans == true) || options.NormalizeEscapedInlineCodeSpans,
            NormalizeTightStrongBoundaries = (source?.NormalizeTightStrongBoundaries == true) || options.NormalizeTightStrongBoundaries,
            NormalizeTightArrowStrongBoundaries = (source?.NormalizeTightArrowStrongBoundaries == true) || options.NormalizeTightArrowStrongBoundaries,
            NormalizeBrokenStrongArrowLabels = (source?.NormalizeBrokenStrongArrowLabels == true) || options.NormalizeBrokenStrongArrowLabels,
            NormalizeWrappedSignalFlowStrongRuns = (source?.NormalizeWrappedSignalFlowStrongRuns == true) || options.NormalizeWrappedSignalFlowStrongRuns,
            NormalizeSignalFlowLabelSpacing = (source?.NormalizeSignalFlowLabelSpacing == true) || options.NormalizeSignalFlowLabelSpacing,
            NormalizeCollapsedMetricChains = (source?.NormalizeCollapsedMetricChains == true) || options.NormalizeCollapsedMetricChains,
            NormalizeHostLabelBulletArtifacts = (source?.NormalizeHostLabelBulletArtifacts == true) || options.NormalizeHostLabelBulletArtifacts,
            NormalizeTightColonSpacing = (source?.NormalizeTightColonSpacing == true) || options.NormalizeTightColonSpacing,
            NormalizeHeadingListBoundaries = (source?.NormalizeHeadingListBoundaries == true) || options.NormalizeHeadingListBoundaries,
            NormalizeCompactStrongLabelListBoundaries = (source?.NormalizeCompactStrongLabelListBoundaries == true) || options.NormalizeCompactStrongLabelListBoundaries,
            NormalizeCompactHeadingBoundaries = (source?.NormalizeCompactHeadingBoundaries == true) || options.NormalizeCompactHeadingBoundaries,
            NormalizeStandaloneHashHeadingSeparators = (source?.NormalizeStandaloneHashHeadingSeparators == true) || options.NormalizeStandaloneHashHeadingSeparators,
            NormalizeBrokenTwoLineStrongLeadIns = (source?.NormalizeBrokenTwoLineStrongLeadIns == true) || options.NormalizeBrokenTwoLineStrongLeadIns,
            NormalizeColonListBoundaries = (source?.NormalizeColonListBoundaries == true) || options.NormalizeColonListBoundaries,
            NormalizeCompactFenceBodyBoundaries = (source?.NormalizeCompactFenceBodyBoundaries == true) || options.NormalizeCompactFenceBodyBoundaries,
            NormalizeLooseStrongDelimiters = (source?.NormalizeLooseStrongDelimiters == true) || options.NormalizeLooseStrongDelimiters,
            NormalizeOrderedListMarkerSpacing = (source?.NormalizeOrderedListMarkerSpacing == true) || options.NormalizeOrderedListMarkerSpacing,
            NormalizeOrderedListParenMarkers = (source?.NormalizeOrderedListParenMarkers == true) || options.NormalizeOrderedListParenMarkers,
            NormalizeOrderedListCaretArtifacts = (source?.NormalizeOrderedListCaretArtifacts == true) || options.NormalizeOrderedListCaretArtifacts,
            NormalizeCollapsedOrderedListBoundaries = (source?.NormalizeCollapsedOrderedListBoundaries == true) || options.NormalizeCollapsedOrderedListBoundaries,
            NormalizeOrderedListStrongDetailClosures = (source?.NormalizeOrderedListStrongDetailClosures == true) || options.NormalizeOrderedListStrongDetailClosures,
            NormalizeTightParentheticalSpacing = (source?.NormalizeTightParentheticalSpacing == true) || options.NormalizeTightParentheticalSpacing,
            NormalizeNestedStrongDelimiters = (source?.NormalizeNestedStrongDelimiters == true) || options.NormalizeNestedStrongDelimiters,
            NormalizeDanglingTrailingStrongListClosers = (source?.NormalizeDanglingTrailingStrongListClosers == true) || options.NormalizeDanglingTrailingStrongListClosers,
            NormalizeMetricValueStrongRuns = (source?.NormalizeMetricValueStrongRuns == true) || options.NormalizeMetricValueStrongRuns
        };
    }

    private static MarkdownInputNormalizationOptions CreateInlineNormalizationOptions(MarkdownInputNormalizationOptions source) {
        return new MarkdownInputNormalizationOptions {
            NormalizeEscapedInlineCodeSpans = source?.NormalizeEscapedInlineCodeSpans ?? false,
            NormalizeTightStrongBoundaries = source?.NormalizeTightStrongBoundaries ?? false,
            NormalizeTightColonSpacing = source?.NormalizeTightColonSpacing ?? false
        };
    }

    private static void CopyFencedBlockExtensions(MarkdownReaderOptions source, MarkdownReaderOptions target) {
        var extensions = source.FencedBlockExtensions;
        if (extensions == null || extensions.Count == 0) {
            return;
        }

        for (int i = 0; i < extensions.Count; i++) {
            var extension = extensions[i];
            if (extension != null) {
                target.FencedBlockExtensions.Add(extension);
            }
        }
    }

    private static void CopyBlockParserExtensions(MarkdownReaderOptions source, MarkdownReaderOptions target) {
        var extensions = source.BlockParserExtensions;
        target.BlockParserExtensions.Clear();
        if (extensions == null || extensions.Count == 0) {
            return;
        }

        for (int i = 0; i < extensions.Count; i++) {
            var extension = extensions[i];
            if (extension != null) {
                target.BlockParserExtensions.Add(extension);
            }
        }
    }

    private static void CopyDocumentTransforms(MarkdownReaderOptions source, MarkdownReaderOptions target) {
        var transforms = source.DocumentTransforms;
        if (transforms == null || transforms.Count == 0) {
            return;
        }

        for (int i = 0; i < transforms.Count; i++) {
            var transform = transforms[i];
            if (transform != null) {
                target.DocumentTransforms.Add(transform);
            }
        }
    }

    private static void AddRendererSemanticFencedBlockExtensions(MarkdownReaderOptions target, MarkdownRendererOptions options) {
        AddSemanticFencedBlockExtension(target, "Built-in Mermaid AST", new[] { MarkdownSemanticKinds.Mermaid }, MarkdownSemanticKinds.Mermaid);

        var mathLanguages = options.Math?.FencedMathLanguages;
        if (mathLanguages != null && mathLanguages.Length > 0) {
            AddSemanticFencedBlockExtension(target, "Built-in Math AST", mathLanguages, MarkdownSemanticKinds.Math);
        }

        var renderers = options.FencedCodeBlockRenderers;
        if (renderers == null || renderers.Count == 0) {
            return;
        }

        for (int i = 0; i < renderers.Count; i++) {
            var renderer = renderers[i];
            if (renderer == null) {
                continue;
            }

            var semanticKind = string.IsNullOrWhiteSpace(renderer.SemanticKind)
                ? renderer.Languages[0]
                : renderer.SemanticKind;
            AddSemanticFencedBlockExtension(target, renderer.Name + " AST", renderer.Languages, semanticKind);
        }
    }

    private static void AddSemanticFencedBlockExtension(
        MarkdownReaderOptions target,
        string name,
        IEnumerable<string> languages,
        string semanticKind) {
        target.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
            name,
            languages,
            context => new SemanticFencedBlock(semanticKind, context.InfoString, context.Content, context.Caption)));
    }

    private static MarkdownInputNormalizationOptions? CreatePreParseNormalizationOptions(MarkdownInputNormalizationOptions source) {
        bool normalizeZeroWidthSpacingArtifacts = source?.NormalizeZeroWidthSpacingArtifacts ?? false;
        bool normalizeEmojiWordJoins = source?.NormalizeEmojiWordJoins ?? false;
        bool normalizeCompactNumberedChoiceBoundaries = source?.NormalizeCompactNumberedChoiceBoundaries ?? false;
        bool normalizeSentenceCollapsedBullets = source?.NormalizeSentenceCollapsedBullets ?? false;
        bool normalizeSoftWrappedStrong = source?.NormalizeSoftWrappedStrongSpans ?? false;
        bool normalizeInlineCodeLineBreaks = source?.NormalizeInlineCodeSpanLineBreaks ?? false;
        bool normalizeLooseStrongDelimiters = source?.NormalizeLooseStrongDelimiters ?? false;
        bool normalizeTightStrongBoundaries = source?.NormalizeTightStrongBoundaries ?? false;
        bool normalizeTightArrowStrongBoundaries = source?.NormalizeTightArrowStrongBoundaries ?? false;
        bool normalizeBrokenStrongArrowLabels = source?.NormalizeBrokenStrongArrowLabels ?? false;
        // These transcript repairs still need to happen before parse so malformed input
        // does not collapse into the wrong block/inline structure.
        bool normalizeWrappedSignalFlowStrongRuns = source?.NormalizeWrappedSignalFlowStrongRuns ?? false;
        bool normalizeSignalFlowLabelSpacing = source?.NormalizeSignalFlowLabelSpacing ?? false;
        bool normalizeCollapsedMetricChains = source?.NormalizeCollapsedMetricChains ?? false;
        bool normalizeHostLabelBulletArtifacts = source?.NormalizeHostLabelBulletArtifacts ?? false;
        bool normalizeHeadingListBoundaries = source?.NormalizeHeadingListBoundaries ?? false;
        bool normalizeCompactStrongLabelListBoundaries = source?.NormalizeCompactStrongLabelListBoundaries ?? false;
        bool normalizeCompactHeadingBoundaries = source?.NormalizeCompactHeadingBoundaries ?? false;
        bool normalizeStandaloneHashHeadingSeparators = source?.NormalizeStandaloneHashHeadingSeparators ?? false;
        bool normalizeBrokenTwoLineStrongLeadIns = source?.NormalizeBrokenTwoLineStrongLeadIns ?? false;
        bool normalizeColonListBoundaries = source?.NormalizeColonListBoundaries ?? false;
        bool normalizeCompactFenceBodyBoundaries = source?.NormalizeCompactFenceBodyBoundaries ?? false;
        bool normalizeOrderedListMarkerSpacing = source?.NormalizeOrderedListMarkerSpacing ?? false;
        bool normalizeOrderedListParenMarkers = source?.NormalizeOrderedListParenMarkers ?? false;
        bool normalizeOrderedListCaretArtifacts = source?.NormalizeOrderedListCaretArtifacts ?? false;
        bool normalizeCollapsedOrderedListBoundaries = source?.NormalizeCollapsedOrderedListBoundaries ?? false;
        bool normalizeOrderedListStrongDetailClosures = source?.NormalizeOrderedListStrongDetailClosures ?? false;
        bool normalizeTightParentheticalSpacing = source?.NormalizeTightParentheticalSpacing ?? false;
        bool normalizeNestedStrongDelimiters = source?.NormalizeNestedStrongDelimiters ?? false;
        bool normalizeDanglingTrailingStrongListClosers = source?.NormalizeDanglingTrailingStrongListClosers ?? false;
        bool normalizeMetricValueStrongRuns = source?.NormalizeMetricValueStrongRuns ?? false;

        if (!normalizeZeroWidthSpacingArtifacts
            && !normalizeEmojiWordJoins
            && !normalizeCompactNumberedChoiceBoundaries
            && !normalizeSentenceCollapsedBullets
            && !normalizeSoftWrappedStrong
            && !normalizeInlineCodeLineBreaks
            && !normalizeLooseStrongDelimiters
            && !normalizeTightStrongBoundaries
            && !normalizeTightArrowStrongBoundaries
            && !normalizeBrokenStrongArrowLabels
            && !normalizeWrappedSignalFlowStrongRuns
            && !normalizeSignalFlowLabelSpacing
            && !normalizeCollapsedMetricChains
            && !normalizeHostLabelBulletArtifacts
            && !normalizeHeadingListBoundaries
            && !normalizeCompactStrongLabelListBoundaries
            && !normalizeCompactHeadingBoundaries
            && !normalizeStandaloneHashHeadingSeparators
            && !normalizeBrokenTwoLineStrongLeadIns
            && !normalizeColonListBoundaries
            && !normalizeCompactFenceBodyBoundaries
            && !normalizeOrderedListMarkerSpacing
            && !normalizeOrderedListParenMarkers
            && !normalizeOrderedListCaretArtifacts
            && !normalizeCollapsedOrderedListBoundaries
            && !normalizeOrderedListStrongDetailClosures
            && !normalizeTightParentheticalSpacing
            && !normalizeNestedStrongDelimiters
            && !normalizeDanglingTrailingStrongListClosers
            && !normalizeMetricValueStrongRuns) {
            return null;
        }

        return new MarkdownInputNormalizationOptions {
            NormalizeZeroWidthSpacingArtifacts = normalizeZeroWidthSpacingArtifacts,
            NormalizeEmojiWordJoins = normalizeEmojiWordJoins,
            NormalizeCompactNumberedChoiceBoundaries = normalizeCompactNumberedChoiceBoundaries,
            NormalizeSentenceCollapsedBullets = normalizeSentenceCollapsedBullets,
            NormalizeSoftWrappedStrongSpans = normalizeSoftWrappedStrong,
            NormalizeInlineCodeSpanLineBreaks = normalizeInlineCodeLineBreaks,
            NormalizeLooseStrongDelimiters = normalizeLooseStrongDelimiters,
            NormalizeTightStrongBoundaries = normalizeTightStrongBoundaries,
            NormalizeTightArrowStrongBoundaries = normalizeTightArrowStrongBoundaries,
            NormalizeBrokenStrongArrowLabels = normalizeBrokenStrongArrowLabels,
            NormalizeWrappedSignalFlowStrongRuns = normalizeWrappedSignalFlowStrongRuns,
            NormalizeSignalFlowLabelSpacing = normalizeSignalFlowLabelSpacing,
            NormalizeCollapsedMetricChains = normalizeCollapsedMetricChains,
            NormalizeHostLabelBulletArtifacts = normalizeHostLabelBulletArtifacts,
            NormalizeHeadingListBoundaries = normalizeHeadingListBoundaries,
            NormalizeCompactStrongLabelListBoundaries = normalizeCompactStrongLabelListBoundaries,
            NormalizeCompactHeadingBoundaries = normalizeCompactHeadingBoundaries,
            NormalizeStandaloneHashHeadingSeparators = normalizeStandaloneHashHeadingSeparators,
            NormalizeBrokenTwoLineStrongLeadIns = normalizeBrokenTwoLineStrongLeadIns,
            NormalizeColonListBoundaries = normalizeColonListBoundaries,
            NormalizeCompactFenceBodyBoundaries = normalizeCompactFenceBodyBoundaries,
            NormalizeOrderedListMarkerSpacing = normalizeOrderedListMarkerSpacing,
            NormalizeOrderedListParenMarkers = normalizeOrderedListParenMarkers,
            NormalizeOrderedListCaretArtifacts = normalizeOrderedListCaretArtifacts,
            NormalizeCollapsedOrderedListBoundaries = normalizeCollapsedOrderedListBoundaries,
            NormalizeOrderedListStrongDetailClosures = normalizeOrderedListStrongDetailClosures,
            NormalizeTightParentheticalSpacing = normalizeTightParentheticalSpacing,
            NormalizeNestedStrongDelimiters = normalizeNestedStrongDelimiters,
            NormalizeDanglingTrailingStrongListClosers = normalizeDanglingTrailingStrongListClosers,
            NormalizeMetricValueStrongRuns = normalizeMetricValueStrongRuns
        };
    }
}
