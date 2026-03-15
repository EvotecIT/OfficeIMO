using System.IO;
using System.Text;
// Intentionally avoid heavy regex use; simple scanning is used for resilience and speed.

namespace OfficeIMO.Markdown;

/// <summary>
/// Parses Markdown text into OfficeIMO.Markdown's typed object model (<see cref="MarkdownDoc"/>, blocks, and inlines).
///
/// Scope: profile-driven and extension-aware. The core reader can be shaped into OfficeIMO,
/// CommonMark-style, GitHub Flavored Markdown-style, or portable behavior via
/// <see cref="MarkdownReaderOptions"/>, including explicit block parser extension registrations.
/// </summary>
public static partial class MarkdownReader {
    /// <summary>
    /// Parses Markdown text into a <see cref="MarkdownDoc"/> with typed blocks and basic inlines.
    /// </summary>
    public static MarkdownDoc Parse(string markdown, MarkdownReaderOptions? options = null) {
        options ??= new MarkdownReaderOptions();
        var state = new MarkdownReaderState();
        return ParseInternal(markdown, options, state, allowFrontMatter: true);
    }

    /// <summary>
    /// Parses Markdown text into both the object model and a lightweight syntax tree with source spans.
    /// </summary>
    public static MarkdownParseResult ParseWithSyntaxTree(string markdown, MarkdownReaderOptions? options = null) {
        options ??= new MarkdownReaderOptions();
        var state = new MarkdownReaderState();
        var syntaxNodes = new List<MarkdownSyntaxNode>();
        var document = ParseInternal(markdown, options, state, allowFrontMatter: true, syntaxNodes, lineOffset: 0);
        return new MarkdownParseResult(document, BuildDocumentSyntaxTree(syntaxNodes));
    }

    /// <summary>Parses a Markdown file path into a <see cref="MarkdownDoc"/>.</summary>
    public static MarkdownDoc ParseFile(string path, MarkdownReaderOptions? options = null) {
        string text = File.ReadAllText(path, Encoding.UTF8);
        return Parse(text, options);
    }

    internal static IReadOnlyList<IMarkdownBlock> ParseBlockFragment(
        string markdown,
        MarkdownReaderOptions? options = null,
        MarkdownReaderState? state = null) {
        options ??= new MarkdownReaderOptions();
        state ??= new MarkdownReaderState();
        var (blocks, _) = ParseNestedMarkdownBlocks(markdown ?? string.Empty, options, state, state.SourceLineOffset);
        return blocks;
    }

    private static MarkdownDoc ParseInternal(string markdown, MarkdownReaderOptions options, MarkdownReaderState state, bool allowFrontMatter, List<MarkdownSyntaxNode>? syntaxNodes = null, int lineOffset = 0) {
        var doc = MarkdownDoc.Create();
        if (string.IsNullOrEmpty(markdown)) return doc;
        int previousLineOffset = state.SourceLineOffset;
        state.SourceLineOffset = lineOffset;

        try {
            // Normalize BOM (U+FEFF) at the very beginning to avoid blocking heading/html detection
            if (markdown[0] == '\uFEFF') {
                markdown = markdown.Substring(1);
            }

            ValidateInputLength(markdown, options.MaxInputCharacters, nameof(markdown));

            var preParseNormalization = CreatePreParseNormalizationOptions(options.InputNormalization);
            if (preParseNormalization != null) {
                markdown = MarkdownInputNormalizer.Normalize(markdown, preParseNormalization);
            }

            // Normalize line endings and split. Keep empty lines significant for block boundaries.
            var text = markdown.Replace("\r\n", "\n").Replace('\r', '\n');
            var lines = text.Split('\n');
            int i = 0;

            // Front matter (YAML) only if it's the very first thing in the file
            if (allowFrontMatter && options.FrontMatter && i < lines.Length && lines[i].Trim() == "---") {
                int start = i + 1;
                int end = -1;
                for (int j = start; j < lines.Length; j++) { if (lines[j].Trim() == "---") { end = j; break; } }
                if (end > start) {
                    var dict = ParseFrontMatter(lines, start, end - 1);
                    if (dict.Count > 0) {
                        var frontMatter = FrontMatterBlock.FromObject(dict);
                        doc.Add(frontMatter);
                        if (syntaxNodes != null) {
                            syntaxNodes.Add(((ISyntaxMarkdownBlock)frontMatter).BuildSyntaxNode(
                                new MarkdownSourceSpan(lineOffset + i + 1, lineOffset + end + 1)));
                        }
                    }
                    i = end + 1;
                    // optional blank line after front matter
                    if (i < lines.Length && string.IsNullOrWhiteSpace(lines[i])) i++;
                }
            }

            var pipeline = MarkdownReaderPipeline.Default(options);
            // Pre-scan for reference-style link definitions so inline refs in earlier paragraphs can resolve
            PreScanReferenceLinkDefinitions(lines, state, options);
            while (i < lines.Length) {
                if (string.IsNullOrWhiteSpace(lines[i])) { i++; continue; }
                bool matched = false;
                var parsers = pipeline.Parsers;
                int previousBlockCount = doc.Blocks.Count;
                int startLine = lineOffset + i;
                for (int p = 0; p < parsers.Count; p++) {
                    if (parsers[p].TryParse(lines, ref i, options, doc, state)) {
                        matched = true;
                        if (syntaxNodes != null && doc.Blocks.Count > previousBlockCount) {
                            CaptureSyntaxNodes(doc, previousBlockCount, startLine, lineOffset + i, syntaxNodes);
                        }
                        break;
                    }
                }
                if (!matched) i++; // defensive: avoid infinite loop
            }

            return ApplyDocumentTransforms(doc, options);
        } finally {
            state.SourceLineOffset = previousLineOffset;
        }
    }

    private static void PreScanReferenceLinkDefinitions(string[] lines, MarkdownReaderState state) {
        PreScanReferenceLinkDefinitions(lines, state, new MarkdownReaderOptions());
    }

    private static void PreScanReferenceLinkDefinitions(string[] lines, MarkdownReaderState state, MarkdownReaderOptions options) {
        bool inFence = false;
        char fenceChar = '\0';
        int fenceLen = 0;

        for (int idx = 0; idx < lines.Length; idx++) {
            var line = lines[idx]; if (string.IsNullOrWhiteSpace(line)) continue;

            // Ignore anything inside fenced code blocks.
            if (!inFence) {
                if (IsCodeFenceOpen(line, out _, out fenceChar, out fenceLen)) {
                    inFence = true;
                    continue;
                }
            } else {
                if (IsCodeFenceClose(line, fenceChar, fenceLen)) {
                    inFence = false;
                }
                continue;
            }

            // Ignore indented code blocks (4+ leading spaces or a tab). Reference definitions are only valid
            // up to 3 leading spaces in typical Markdown implementations.
            int leading = 0;
            while (leading < line.Length && line[leading] == ' ') leading++;
            if (leading >= 4) continue;
            if (leading < line.Length && line[leading] == '\t') continue;

            if (TryParseReferenceLinkDefinition(lines, idx, options, out var label, out var url, out var title, out var consumedLines)) {
                var resolved = ResolveUrl(url, options);
                if (resolved != null && !state.LinkRefs.ContainsKey(label)) state.LinkRefs[label] = (resolved!, title);
                idx += consumedLines - 1;
            }
        }
    }

    private static bool TryParseReferenceLinkDefinition(string[] lines, int index, MarkdownReaderOptions options, out string label, out string url, out string? title, out int consumedLines) {
        label = url = string.Empty;
        title = null;
        consumedLines = 0;

        if (index < 0 || index >= lines.Length) return false;
        var line = lines[index];
        if (string.IsNullOrWhiteSpace(line)) return false;

        int leading = 0;
        while (leading < line.Length && line[leading] == ' ') leading++;
        if (leading >= 4) return false;
        if (leading < line.Length && line[leading] == '\t') return false;

        var trimmed = line.Trim();
        if (trimmed.Length < 5 || trimmed[0] != '[') return false;
        if (trimmed.Length > 1 && trimmed[1] == '^') return false;

        int rb = FindReferenceLabelEnd(trimmed, 0);
        if (rb <= 1) return false;
        if (rb + 1 >= trimmed.Length || trimmed[rb + 1] != ':') return false;

        label = NormalizeReferenceLabel(trimmed.Substring(1, rb - 1));
        string rest = trimmed.Substring(rb + 2).Trim();
        if (string.IsNullOrEmpty(rest)) return false;

        if (TrySplitUrlAndOptionalTitle(rest, out url, out title)) {
            consumedLines = 1;
            if (title == null && TryParseReferenceTitleContinuation(lines, index + 1, out var continuedTitle)) {
                title = continuedTitle;
                consumedLines = 2;
            }
            return !string.IsNullOrEmpty(label);
        }

        if (IndexOfWhitespace(rest) >= 0) return false;

        url = UnescapeMarkdownBackslashEscapes(rest);
        title = null;
        consumedLines = 1;

        if (TryParseReferenceTitleContinuation(lines, index + 1, out var continuationTitle)) {
            title = continuationTitle;
            consumedLines = 2;
        }

        return !string.IsNullOrEmpty(label);
    }

    private static bool TryParseReferenceTitleContinuation(string[] lines, int index, out string? title) {
        title = null;
        if (index < 0 || index >= lines.Length) return false;

        var line = lines[index];
        if (string.IsNullOrWhiteSpace(line)) return false;

        int leading = 0;
        while (leading < line.Length && line[leading] == ' ') leading++;
        if (leading >= 4) return false;
        if (leading < line.Length && line[leading] == '\t') return false;

        title = TryParseOptionalTitleToken(line.Trim());
        if (title == null) return false;

        title = UnescapeMarkdownBackslashEscapes(title);
        return true;
    }

    private static bool StartsWithReferenceDefinitionLikeLabel(string line) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        var trimmed = line.TrimStart();
        if (trimmed.Length < 4 || trimmed[0] != '[') return false;
        if (trimmed.Length > 1 && trimmed[1] == '^') return false;

        int balancedEnd = FindMatchingBracket(trimmed, 0);
        return balancedEnd >= 1 && balancedEnd + 1 < trimmed.Length && trimmed[balancedEnd + 1] == ':';
    }

    private static string NormalizeReferenceLabel(string? label) {
        if (string.IsNullOrWhiteSpace(label)) return string.Empty;
        var t = label!.Trim();
        var sb = new System.Text.StringBuilder(t.Length);
        bool prevSpace = false;
        for (int i = 0; i < t.Length; i++) {
            char c = t[i];
            if (char.IsWhiteSpace(c)) {
                if (!prevSpace) sb.Append(' ');
                prevSpace = true;
            } else {
                sb.Append(c);
                prevSpace = false;
            }
        }
        return sb.ToString();
    }

    private static MarkdownReaderOptions CloneOptionsWithoutFrontMatter(MarkdownReaderOptions source) {
        var clone = new MarkdownReaderOptions {
            FrontMatter = false,
            Callouts = source.Callouts,
            Headings = source.Headings,
            FencedCode = source.FencedCode,
            IndentedCodeBlocks = source.IndentedCodeBlocks,
            Images = source.Images,
            UnorderedLists = source.UnorderedLists,
            TaskLists = source.TaskLists,
            OrderedLists = source.OrderedLists,
            Tables = source.Tables,
            DefinitionLists = source.DefinitionLists,
            TocPlaceholders = source.TocPlaceholders,
            Footnotes = source.Footnotes,
            PreferNarrativeSingleLineDefinitions = source.PreferNarrativeSingleLineDefinitions,
            HtmlBlocks = source.HtmlBlocks,
            Paragraphs = source.Paragraphs,
            AutolinkUrls = source.AutolinkUrls,
            AutolinkWwwUrls = source.AutolinkWwwUrls,
            AutolinkWwwScheme = source.AutolinkWwwScheme,
            AutolinkEmails = source.AutolinkEmails,
            BackslashHardBreaks = source.BackslashHardBreaks,
            InlineHtml = source.InlineHtml,
            BaseUri = source.BaseUri,
            DisallowScriptUrls = source.DisallowScriptUrls,
            DisallowFileUrls = source.DisallowFileUrls,
            AllowMailtoUrls = source.AllowMailtoUrls,
            AllowDataUrls = source.AllowDataUrls,
            AllowProtocolRelativeUrls = source.AllowProtocolRelativeUrls,
            RestrictUrlSchemes = source.RestrictUrlSchemes,
            AllowedUrlSchemes = source.AllowedUrlSchemes,
            MaxInputCharacters = source.MaxInputCharacters,
            InputNormalization = new MarkdownInputNormalizationOptions {
                NormalizeZeroWidthSpacingArtifacts = source.InputNormalization?.NormalizeZeroWidthSpacingArtifacts ?? false,
                NormalizeEmojiWordJoins = source.InputNormalization?.NormalizeEmojiWordJoins ?? false,
                NormalizeCompactNumberedChoiceBoundaries = source.InputNormalization?.NormalizeCompactNumberedChoiceBoundaries ?? false,
                NormalizeSentenceCollapsedBullets = source.InputNormalization?.NormalizeSentenceCollapsedBullets ?? false,
                NormalizeSoftWrappedStrongSpans = source.InputNormalization?.NormalizeSoftWrappedStrongSpans ?? false,
                NormalizeInlineCodeSpanLineBreaks = source.InputNormalization?.NormalizeInlineCodeSpanLineBreaks ?? false,
                NormalizeEscapedInlineCodeSpans = source.InputNormalization?.NormalizeEscapedInlineCodeSpans ?? false,
                NormalizeTightStrongBoundaries = source.InputNormalization?.NormalizeTightStrongBoundaries ?? false,
                NormalizeTightArrowStrongBoundaries = source.InputNormalization?.NormalizeTightArrowStrongBoundaries ?? false,
                NormalizeBrokenStrongArrowLabels = source.InputNormalization?.NormalizeBrokenStrongArrowLabels ?? false,
                NormalizeWrappedSignalFlowStrongRuns = source.InputNormalization?.NormalizeWrappedSignalFlowStrongRuns ?? false,
                NormalizeSignalFlowLabelSpacing = source.InputNormalization?.NormalizeSignalFlowLabelSpacing ?? false,
                NormalizeCollapsedMetricChains = source.InputNormalization?.NormalizeCollapsedMetricChains ?? false,
                NormalizeHostLabelBulletArtifacts = source.InputNormalization?.NormalizeHostLabelBulletArtifacts ?? false,
                NormalizeTightColonSpacing = source.InputNormalization?.NormalizeTightColonSpacing ?? false,
                NormalizeHeadingListBoundaries = source.InputNormalization?.NormalizeHeadingListBoundaries ?? false,
                NormalizeCompactStrongLabelListBoundaries = source.InputNormalization?.NormalizeCompactStrongLabelListBoundaries ?? false,
                NormalizeCompactHeadingBoundaries = source.InputNormalization?.NormalizeCompactHeadingBoundaries ?? false,
                NormalizeStandaloneHashHeadingSeparators = source.InputNormalization?.NormalizeStandaloneHashHeadingSeparators ?? false,
                NormalizeBrokenTwoLineStrongLeadIns = source.InputNormalization?.NormalizeBrokenTwoLineStrongLeadIns ?? false,
                NormalizeColonListBoundaries = source.InputNormalization?.NormalizeColonListBoundaries ?? false,
                NormalizeCompactFenceBodyBoundaries = source.InputNormalization?.NormalizeCompactFenceBodyBoundaries ?? false,
                NormalizeLooseStrongDelimiters = source.InputNormalization?.NormalizeLooseStrongDelimiters ?? false,
                NormalizeOrderedListMarkerSpacing = source.InputNormalization?.NormalizeOrderedListMarkerSpacing ?? false,
                NormalizeOrderedListParenMarkers = source.InputNormalization?.NormalizeOrderedListParenMarkers ?? false,
                NormalizeOrderedListCaretArtifacts = source.InputNormalization?.NormalizeOrderedListCaretArtifacts ?? false,
                NormalizeCollapsedOrderedListBoundaries = source.InputNormalization?.NormalizeCollapsedOrderedListBoundaries ?? false,
                NormalizeOrderedListStrongDetailClosures = source.InputNormalization?.NormalizeOrderedListStrongDetailClosures ?? false,
                NormalizeTightParentheticalSpacing = source.InputNormalization?.NormalizeTightParentheticalSpacing ?? false,
                NormalizeNestedStrongDelimiters = source.InputNormalization?.NormalizeNestedStrongDelimiters ?? false,
                NormalizeDanglingTrailingStrongListClosers = source.InputNormalization?.NormalizeDanglingTrailingStrongListClosers ?? false,
                NormalizeMetricValueStrongRuns = source.InputNormalization?.NormalizeMetricValueStrongRuns ?? false
            }
        };

        CopyBlockParserExtensions(source, clone);
        CopyFencedBlockExtensions(source, clone);
        CopyDocumentTransforms(source, clone);
        return clone;
    }

    private static MarkdownInputNormalizationOptions? CreatePreParseNormalizationOptions(MarkdownInputNormalizationOptions? source) {
        bool normalizeZeroWidthSpacingArtifacts = source?.NormalizeZeroWidthSpacingArtifacts ?? false;
        bool normalizeEmojiWordJoins = source?.NormalizeEmojiWordJoins ?? false;
        bool normalizeCompactNumberedChoiceBoundaries = source?.NormalizeCompactNumberedChoiceBoundaries ?? false;
        bool normalizeSentenceCollapsedBullets = source?.NormalizeSentenceCollapsedBullets ?? false;
        bool normalizeSoftWrappedStrong = source?.NormalizeSoftWrappedStrongSpans ?? false;
        bool normalizeInlineCodeLineBreaks = source?.NormalizeInlineCodeSpanLineBreaks ?? false;
        bool normalizeLooseStrongDelimiters = source?.NormalizeLooseStrongDelimiters ?? false;
        bool normalizeTightArrowStrongBoundaries = source?.NormalizeTightArrowStrongBoundaries ?? false;
        bool normalizeBrokenStrongArrowLabels = source?.NormalizeBrokenStrongArrowLabels ?? false;
        // These transcript repairs still need to happen before parse so malformed input
        // does not collapse into the wrong block/inline structure. Standalone hash heading
        // separators are intentionally handled later via a document transform because the
        // markdown already parses into a recoverable block structure.
        bool normalizeWrappedSignalFlowStrongRuns = source?.NormalizeWrappedSignalFlowStrongRuns ?? false;
        bool normalizeSignalFlowLabelSpacing = source?.NormalizeSignalFlowLabelSpacing ?? false;
        bool normalizeCollapsedMetricChains = source?.NormalizeCollapsedMetricChains ?? false;
        bool normalizeHostLabelBulletArtifacts = source?.NormalizeHostLabelBulletArtifacts ?? false;
        bool normalizeHeadingListBoundaries = source?.NormalizeHeadingListBoundaries ?? false;
        bool normalizeCompactStrongLabelListBoundaries = source?.NormalizeCompactStrongLabelListBoundaries ?? false;
        bool normalizeCompactHeadingBoundaries = source?.NormalizeCompactHeadingBoundaries ?? false;
        bool normalizeBrokenTwoLineStrongLeadIns = source?.NormalizeBrokenTwoLineStrongLeadIns ?? false;
        bool normalizeColonListBoundaries = source?.NormalizeColonListBoundaries ?? false;
        bool normalizeCompactFenceBodyBoundaries = source?.NormalizeCompactFenceBodyBoundaries ?? false;
        bool normalizeOrderedListMarkerSpacing = source?.NormalizeOrderedListMarkerSpacing ?? false;
        bool normalizeOrderedListParenMarkers = source?.NormalizeOrderedListParenMarkers ?? false;
        bool normalizeOrderedListCaretArtifacts = source?.NormalizeOrderedListCaretArtifacts ?? false;
        bool normalizeCollapsedOrderedListBoundaries = source?.NormalizeCollapsedOrderedListBoundaries ?? false;
        bool normalizeOrderedListStrongDetailClosures = source?.NormalizeOrderedListStrongDetailClosures ?? false;
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
            && !normalizeTightArrowStrongBoundaries
            && !normalizeBrokenStrongArrowLabels
            && !normalizeWrappedSignalFlowStrongRuns
            && !normalizeSignalFlowLabelSpacing
            && !normalizeCollapsedMetricChains
            && !normalizeHostLabelBulletArtifacts
            && !normalizeHeadingListBoundaries
            && !normalizeCompactStrongLabelListBoundaries
            && !normalizeCompactHeadingBoundaries
            && !normalizeBrokenTwoLineStrongLeadIns
            && !normalizeColonListBoundaries
            && !normalizeCompactFenceBodyBoundaries
            && !normalizeOrderedListMarkerSpacing
            && !normalizeOrderedListParenMarkers
            && !normalizeOrderedListCaretArtifacts
            && !normalizeCollapsedOrderedListBoundaries
            && !normalizeOrderedListStrongDetailClosures
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
            NormalizeTightArrowStrongBoundaries = normalizeTightArrowStrongBoundaries,
            NormalizeBrokenStrongArrowLabels = normalizeBrokenStrongArrowLabels,
            NormalizeWrappedSignalFlowStrongRuns = normalizeWrappedSignalFlowStrongRuns,
            NormalizeSignalFlowLabelSpacing = normalizeSignalFlowLabelSpacing,
            NormalizeCollapsedMetricChains = normalizeCollapsedMetricChains,
            NormalizeHostLabelBulletArtifacts = normalizeHostLabelBulletArtifacts,
            NormalizeHeadingListBoundaries = normalizeHeadingListBoundaries,
            NormalizeCompactStrongLabelListBoundaries = normalizeCompactStrongLabelListBoundaries,
            NormalizeCompactHeadingBoundaries = normalizeCompactHeadingBoundaries,
            NormalizeBrokenTwoLineStrongLeadIns = normalizeBrokenTwoLineStrongLeadIns,
            NormalizeColonListBoundaries = normalizeColonListBoundaries,
            NormalizeCompactFenceBodyBoundaries = normalizeCompactFenceBodyBoundaries,
            NormalizeOrderedListMarkerSpacing = normalizeOrderedListMarkerSpacing,
            NormalizeOrderedListParenMarkers = normalizeOrderedListParenMarkers,
            NormalizeOrderedListCaretArtifacts = normalizeOrderedListCaretArtifacts,
            NormalizeCollapsedOrderedListBoundaries = normalizeCollapsedOrderedListBoundaries,
            NormalizeOrderedListStrongDetailClosures = normalizeOrderedListStrongDetailClosures,
            NormalizeNestedStrongDelimiters = normalizeNestedStrongDelimiters,
            NormalizeDanglingTrailingStrongListClosers = normalizeDanglingTrailingStrongListClosers,
            NormalizeMetricValueStrongRuns = normalizeMetricValueStrongRuns
        };
    }

    private static MarkdownReaderState CloneState(MarkdownReaderState state) {
        var clone = new MarkdownReaderState();
        foreach (var kvp in state.LinkRefs) clone.LinkRefs[kvp.Key] = kvp.Value;
        clone.SourceLineOffset = state.SourceLineOffset;
        return clone;
    }

    private static void CopyFencedBlockExtensions(MarkdownReaderOptions source, MarkdownReaderOptions target) {
        if (source == null || target == null) {
            return;
        }

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
        if (source == null || target == null) {
            return;
        }

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
        if (source == null || target == null) {
            return;
        }

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

    private static MarkdownDoc ApplyDocumentTransforms(MarkdownDoc document, MarkdownReaderOptions options) {
        var transforms = BuildEffectiveDocumentTransforms(options);
        return MarkdownDocumentTransformPipeline.Apply(
            document,
            transforms,
            new MarkdownDocumentTransformContext(MarkdownDocumentTransformSource.MarkdownReader, options));
    }

    private static IReadOnlyList<IMarkdownDocumentTransform> BuildEffectiveDocumentTransforms(MarkdownReaderOptions options) {
        if (options?.InputNormalization?.NormalizeStandaloneHashHeadingSeparators != true) {
            if (options == null) {
                return Array.Empty<IMarkdownDocumentTransform>();
            }

            return options.DocumentTransforms;
        }

        var configured = options.DocumentTransforms;
        for (var i = 0; i < configured.Count; i++) {
            if (configured[i] is MarkdownStandaloneHashHeadingSeparatorTransform) {
                return configured;
            }
        }

        var transforms = new List<IMarkdownDocumentTransform>(configured.Count + 1) {
            new MarkdownStandaloneHashHeadingSeparatorTransform()
        };

        for (var i = 0; i < configured.Count; i++) {
            if (configured[i] != null) {
                transforms.Add(configured[i]);
            }
        }

        return transforms;
    }

    private static void ValidateInputLength(string input, int? maxInputCharacters, string paramName) {
        if (!maxInputCharacters.HasValue) {
            return;
        }

        if (maxInputCharacters.Value <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maxInputCharacters), maxInputCharacters.Value, "MaxInputCharacters must be greater than zero.");
        }

        if (input.Length > maxInputCharacters.Value) {
            throw new ArgumentOutOfRangeException(paramName, input.Length, $"Input exceeds MaxInputCharacters ({maxInputCharacters.Value}).");
        }
    }

    private static (IReadOnlyList<IMarkdownBlock> Blocks, IReadOnlyList<MarkdownSyntaxNode> SyntaxChildren) ParseNestedMarkdownBlocks(
        string markdown,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        int lineOffset) {

        var nestedOptions = CloneOptionsWithoutFrontMatter(options);
        var nestedState = CloneState(state);
        var syntaxChildren = new List<MarkdownSyntaxNode>();
        var nestedDoc = ParseInternal(markdown, nestedOptions, nestedState, allowFrontMatter: false, syntaxChildren, lineOffset: lineOffset);
        return (nestedDoc.Blocks, syntaxChildren);
    }
}
