using System.IO;
using System.Linq;
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
        return ParseInternal(markdown, options, state, allowFrontMatter: true, out _);
    }

    /// <summary>
    /// Parses Markdown text into both the object model and a lightweight syntax tree with source spans.
    /// </summary>
    public static MarkdownParseResult ParseWithSyntaxTree(string markdown, MarkdownReaderOptions? options = null) {
        options ??= new MarkdownReaderOptions();
        var state = new MarkdownReaderState();
        var syntaxNodes = new List<MarkdownSyntaxNode>();
        var diagnostics = new List<MarkdownDocumentTransformDiagnostic>();
        var document = ParseInternal(markdown, options, state, allowFrontMatter: true, out var syntaxTree, syntaxNodes, lineOffset: 0, transformDiagnostics: diagnostics);
        var originalSyntaxTree = syntaxTree ?? BuildDocumentSyntaxTree(syntaxNodes, document);
        if (diagnostics.Any(diagnostic => diagnostic.ReplacedDocument)) {
            originalSyntaxTree = DetachOriginalSyntaxAssociations(originalSyntaxTree);
        }

        var finalSyntaxTree = BuildFinalSyntaxTree(document, originalSyntaxTree, diagnostics);
        MarkdownObjectTreeBinder.BindDocument(document, finalSyntaxTree);
        return new MarkdownParseResult(document, originalSyntaxTree, finalSyntaxTree);
    }

    /// <summary>
    /// Parses Markdown text into the object model, original syntax tree, and document-transform diagnostics.
    /// </summary>
    public static MarkdownParseResult ParseWithSyntaxTreeAndDiagnostics(string markdown, MarkdownReaderOptions? options = null) {
        options ??= new MarkdownReaderOptions();
        var state = new MarkdownReaderState();
        var syntaxNodes = new List<MarkdownSyntaxNode>();
        var diagnostics = new List<MarkdownDocumentTransformDiagnostic>();
        var document = ParseInternal(
            markdown,
            options,
            state,
            allowFrontMatter: true,
            out var syntaxTree,
            syntaxNodes,
            lineOffset: 0,
            transformDiagnostics: diagnostics);
        var originalSyntaxTree = syntaxTree ?? BuildDocumentSyntaxTree(syntaxNodes, document);
        if (diagnostics.Any(diagnostic => diagnostic.ReplacedDocument)) {
            originalSyntaxTree = DetachOriginalSyntaxAssociations(originalSyntaxTree);
        }

        var finalSyntaxTree = BuildFinalSyntaxTree(document, originalSyntaxTree, diagnostics);
        MarkdownObjectTreeBinder.BindDocument(document, finalSyntaxTree);
        return new MarkdownParseResult(document, originalSyntaxTree, finalSyntaxTree, diagnostics);
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

    internal static IReadOnlyList<IMarkdownBlock> ParseNestedBlocksFromLineRange(
        string[] lines,
        int startIndex,
        int lineCount,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        if (lines == null || lines.Length == 0 || lineCount <= 0 || startIndex < 0 || startIndex >= lines.Length) {
            return Array.Empty<IMarkdownBlock>();
        }

        var safeCount = Math.Min(lineCount, lines.Length - startIndex);
        var sourceLines = new List<MarkdownSourceLineSlice>(safeCount);
        for (int offset = 0; offset < safeCount; offset++) {
            sourceLines.Add(new MarkdownSourceLineSlice(
                lines[startIndex + offset] ?? string.Empty,
                state.SourceLineOffset + startIndex + offset + 1,
                startColumn: 1));
        }

        var (blocks, syntaxChildren) = ParseNestedMarkdownBlocks(sourceLines, options, state);
        var nestedDocument = MarkdownDoc.Create();
        for (int blockIndex = 0; blockIndex < blocks.Count; blockIndex++) {
            nestedDocument.Add(blocks[blockIndex]);
        }

        var syntaxTree = BuildDocumentSyntaxTree(syntaxChildren, nestedDocument);
        MarkdownObjectTreeBinder.BindDocument(nestedDocument, syntaxTree);
        return nestedDocument.Blocks;
    }

    private static MarkdownDoc ParseInternal(
        string markdown,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        bool allowFrontMatter,
        out MarkdownSyntaxNode? syntaxTree,
        List<MarkdownSyntaxNode>? syntaxNodes = null,
        int lineOffset = 0,
        ICollection<MarkdownDocumentTransformDiagnostic>? transformDiagnostics = null) {
        var doc = MarkdownDoc.Create();
        syntaxTree = syntaxNodes != null ? BuildDocumentSyntaxTree(syntaxNodes, doc) : null;
        if (string.IsNullOrEmpty(markdown)) return doc;
        int previousLineOffset = state.SourceLineOffset;
        var previousSourceTextMap = state.SourceTextMap;
        state.SourceLineOffset = lineOffset;

        try {
            // Normalize BOM (U+FEFF) at the very beginning to avoid blocking heading/html detection
            if (markdown[0] == '\uFEFF') {
                markdown = markdown.Substring(1);
            }

            ValidateInputLength(markdown, options.MaxInputCharacters, nameof(markdown));

            // This specific repair must happen before block parsing: once a collapsed heading marker
            // is swallowed into a table cell, the AST no longer knows the table boundary was malformed.
            if (options.InputNormalization?.NormalizeCompactHeadingBoundaries == true) {
                markdown = MarkdownInputNormalizer.NormalizeCollapsedTableHeadingBoundaries(markdown);
            }

            var preParseNormalization = CreatePreParseNormalizationOptions(options.InputNormalization);
            if (preParseNormalization != null) {
                markdown = MarkdownInputNormalizer.Normalize(markdown, preParseNormalization);
            }

            // Normalize line endings and split. Keep empty lines significant for block boundaries.
            var text = markdown.Replace("\r\n", "\n").Replace('\r', '\n');
            if (lineOffset == 0 || state.SourceTextMap == null) {
                state.SourceTextMap = new MarkdownSourceTextMap(text);
            }
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
                                CreateLineSpan(state, lineOffset + i + 1, lineOffset + end + 1)));
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
                            CaptureSyntaxNodes(doc, previousBlockCount, startLine, lineOffset + i, syntaxNodes, state);
                        }
                        break;
                    }
                }
                if (!matched) i++; // defensive: avoid infinite loop
            }

            syntaxTree = syntaxNodes != null ? BuildDocumentSyntaxTree(syntaxNodes, doc) : null;
            var transformed = ApplyDocumentTransforms(doc, options, transformDiagnostics, syntaxTree);
            MarkdownObjectTreeBinder.BindDocument(transformed, syntaxTree);
            return transformed;
        } finally {
            state.SourceLineOffset = previousLineOffset;
            state.SourceTextMap = previousSourceTextMap;
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
        CopyInlineParserExtensions(source, clone);
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
        // These repairs stay on the text side because malformed input would otherwise parse
        // into the wrong block/inline structure. Recoverable paragraph/heading/list boundary
        // cleanup is intentionally excluded here and handled later via built-in document
        // transforms so the reader can normalize from the AST whenever markdown is already
        // parseable.
        bool normalizeWrappedSignalFlowStrongRuns = source?.NormalizeWrappedSignalFlowStrongRuns ?? false;
        bool normalizeSignalFlowLabelSpacing = source?.NormalizeSignalFlowLabelSpacing ?? false;
        bool normalizeCollapsedMetricChains = source?.NormalizeCollapsedMetricChains ?? false;
        bool normalizeHostLabelBulletArtifacts = source?.NormalizeHostLabelBulletArtifacts ?? false;
        bool normalizeBrokenTwoLineStrongLeadIns = source?.NormalizeBrokenTwoLineStrongLeadIns ?? false;
        bool normalizeCompactFenceBodyBoundaries = source?.NormalizeCompactFenceBodyBoundaries ?? false;
        bool normalizeOrderedListMarkerSpacing = source?.NormalizeOrderedListMarkerSpacing ?? false;
        bool normalizeOrderedListParenMarkers = source?.NormalizeOrderedListParenMarkers ?? false;
        bool normalizeOrderedListCaretArtifacts = source?.NormalizeOrderedListCaretArtifacts ?? false;
        bool normalizeCollapsedOrderedListBoundaries = source?.NormalizeCollapsedOrderedListBoundaries ?? false;
        bool normalizeOrderedListStrongDetailClosures = source?.NormalizeOrderedListStrongDetailClosures ?? false;
        bool normalizeNestedStrongDelimiters = source?.NormalizeNestedStrongDelimiters ?? false;

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
            && !normalizeBrokenTwoLineStrongLeadIns
            && !normalizeCompactFenceBodyBoundaries
            && !normalizeOrderedListMarkerSpacing
            && !normalizeOrderedListParenMarkers
            && !normalizeOrderedListCaretArtifacts
            && !normalizeCollapsedOrderedListBoundaries
            && !normalizeOrderedListStrongDetailClosures
            && !normalizeNestedStrongDelimiters) {
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
            NormalizeBrokenTwoLineStrongLeadIns = normalizeBrokenTwoLineStrongLeadIns,
            NormalizeCompactFenceBodyBoundaries = normalizeCompactFenceBodyBoundaries,
            NormalizeOrderedListMarkerSpacing = normalizeOrderedListMarkerSpacing,
            NormalizeOrderedListParenMarkers = normalizeOrderedListParenMarkers,
            NormalizeOrderedListCaretArtifacts = normalizeOrderedListCaretArtifacts,
            NormalizeCollapsedOrderedListBoundaries = normalizeCollapsedOrderedListBoundaries,
            NormalizeOrderedListStrongDetailClosures = normalizeOrderedListStrongDetailClosures,
            NormalizeNestedStrongDelimiters = normalizeNestedStrongDelimiters
        };
    }

    private static MarkdownReaderState CloneState(MarkdownReaderState state) {
        var clone = new MarkdownReaderState();
        foreach (var kvp in state.LinkRefs) clone.LinkRefs[kvp.Key] = kvp.Value;
        clone.SourceLineOffset = state.SourceLineOffset;
        clone.SourceTextMap = state.SourceTextMap;
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

    private static void CopyInlineParserExtensions(MarkdownReaderOptions source, MarkdownReaderOptions target) {
        if (source == null || target == null) {
            return;
        }

        var extensions = source.InlineParserExtensions;
        target.InlineParserExtensions.Clear();
        if (extensions == null || extensions.Count == 0) {
            return;
        }

        for (int i = 0; i < extensions.Count; i++) {
            var extension = extensions[i];
            if (extension != null) {
                target.InlineParserExtensions.Add(extension);
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

    private static MarkdownDoc ApplyDocumentTransforms(
        MarkdownDoc document,
        MarkdownReaderOptions options,
        ICollection<MarkdownDocumentTransformDiagnostic>? diagnostics = null,
        MarkdownSyntaxNode? syntaxTree = null) {
        var transforms = BuildEffectiveDocumentTransforms(options);
        return MarkdownDocumentTransformPipeline.Apply(
            document,
            transforms,
            new MarkdownDocumentTransformContext(MarkdownDocumentTransformSource.MarkdownReader, options, sourceOptions: null, diagnostics, syntaxTree));
    }

    private static IReadOnlyList<IMarkdownDocumentTransform> BuildEffectiveDocumentTransforms(MarkdownReaderOptions options) {
        if (options == null) {
            return Array.Empty<IMarkdownDocumentTransform>();
        }

        var normalization = options.InputNormalization;
        bool needsRegisteredFencedBlockTransform = options.FencedBlockExtensions.Count > 0;
        // These flags intentionally map to AST/document transforms rather than pre-parse text
        // repair because the markdown already parses into recoverable paragraph/heading/list
        // structures. Keeping them here makes the routing boundary explicit and prevents the
        // pre-parse normalizer from growing back into a general transcript rewrite pipeline.
        bool needsStandaloneHashTransform = normalization?.NormalizeStandaloneHashHeadingSeparators == true;
        bool needsCompactHeadingBoundaryTransform = normalization?.NormalizeCompactHeadingBoundaries == true;
        bool needsColonListBoundaryTransform = normalization?.NormalizeColonListBoundaries == true;
        bool needsHeadingListBoundaryTransform = normalization?.NormalizeHeadingListBoundaries == true;
        bool needsCompactStrongLabelListBoundaryTransform = normalization?.NormalizeCompactStrongLabelListBoundaries == true;
        bool needsListStrongArtifactTransform =
            normalization?.NormalizeLooseStrongDelimiters == true
            || normalization?.NormalizeDanglingTrailingStrongListClosers == true
            || normalization?.NormalizeMetricValueStrongRuns == true;

        if (!needsRegisteredFencedBlockTransform
            && !needsStandaloneHashTransform
            && !needsCompactHeadingBoundaryTransform
            && !needsColonListBoundaryTransform
            && !needsHeadingListBoundaryTransform
            && !needsCompactStrongLabelListBoundaryTransform
            && !needsListStrongArtifactTransform) {
            return options.DocumentTransforms;
        }

        var configured = options.DocumentTransforms;
        bool hasStandaloneHashTransform = false;
        bool hasCompactHeadingBoundaryTransform = false;
        bool hasColonListBoundaryTransform = false;
        bool hasHeadingListBoundaryTransform = false;
        bool hasCompactStrongLabelListBoundaryTransform = false;
        bool hasListStrongArtifactTransform = false;
        bool hasRegisteredFencedBlockTransform = false;

        for (var i = 0; i < configured.Count; i++) {
            switch (configured[i]) {
                case MarkdownRegisteredFencedBlockTransform:
                    hasRegisteredFencedBlockTransform = true;
                    break;
                case MarkdownStandaloneHashHeadingSeparatorTransform:
                    hasStandaloneHashTransform = true;
                    break;
                case MarkdownCompactHeadingBoundaryTransform:
                    hasCompactHeadingBoundaryTransform = true;
                    break;
                case MarkdownColonListBoundaryTransform:
                    hasColonListBoundaryTransform = true;
                    break;
                case MarkdownHeadingListBoundaryTransform:
                    hasHeadingListBoundaryTransform = true;
                    break;
                case MarkdownCompactStrongLabelListBoundaryTransform:
                    hasCompactStrongLabelListBoundaryTransform = true;
                    break;
                case MarkdownListParagraphStrongArtifactTransform:
                    hasListStrongArtifactTransform = true;
                    break;
            }
        }

        if ((!needsRegisteredFencedBlockTransform || hasRegisteredFencedBlockTransform)
            && (!needsStandaloneHashTransform || hasStandaloneHashTransform)
            && (!needsCompactHeadingBoundaryTransform || hasCompactHeadingBoundaryTransform)
            && (!needsColonListBoundaryTransform || hasColonListBoundaryTransform)
            && (!needsHeadingListBoundaryTransform || hasHeadingListBoundaryTransform)
            && (!needsCompactStrongLabelListBoundaryTransform || hasCompactStrongLabelListBoundaryTransform)
            && (!needsListStrongArtifactTransform || hasListStrongArtifactTransform)) {
            return configured;
        }

        var transforms = new List<IMarkdownDocumentTransform>(configured.Count + 7);
        if (needsRegisteredFencedBlockTransform && !hasRegisteredFencedBlockTransform) {
            transforms.Add(new MarkdownRegisteredFencedBlockTransform(options.FencedBlockExtensions));
        }

        if (needsListStrongArtifactTransform && !hasListStrongArtifactTransform) {
            transforms.Add(new MarkdownListParagraphStrongArtifactTransform(normalization!));
        }

        if (needsCompactHeadingBoundaryTransform && !hasCompactHeadingBoundaryTransform) {
            transforms.Add(new MarkdownCompactHeadingBoundaryTransform());
        }

        if (needsHeadingListBoundaryTransform && !hasHeadingListBoundaryTransform) {
            transforms.Add(new MarkdownHeadingListBoundaryTransform());
        }

        if (needsCompactStrongLabelListBoundaryTransform && !hasCompactStrongLabelListBoundaryTransform) {
            transforms.Add(new MarkdownCompactStrongLabelListBoundaryTransform());
        }

        if (needsColonListBoundaryTransform && !hasColonListBoundaryTransform) {
            transforms.Add(new MarkdownColonListBoundaryTransform());
        }

        if (needsStandaloneHashTransform && !hasStandaloneHashTransform) {
            transforms.Add(new MarkdownStandaloneHashHeadingSeparatorTransform());
        }

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
        var nestedDoc = ParseInternal(markdown, nestedOptions, nestedState, allowFrontMatter: false, out _, syntaxChildren, lineOffset: lineOffset);
        return (nestedDoc.Blocks, syntaxChildren);
    }

    private static (IReadOnlyList<IMarkdownBlock> Blocks, IReadOnlyList<MarkdownSyntaxNode> SyntaxChildren) ParseNestedMarkdownBlocks(
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        if (sourceLines == null || sourceLines.Count == 0) {
            return (Array.Empty<IMarkdownBlock>(), Array.Empty<MarkdownSyntaxNode>());
        }

        var markdown = string.Join("\n", sourceLines.Select(line => line.Text ?? string.Empty));
        var nestedOptions = CloneOptionsWithoutFrontMatter(options);
        var nestedState = CloneState(state);
        var syntaxChildren = new List<MarkdownSyntaxNode>();
        var nestedDoc = ParseInternal(markdown, nestedOptions, nestedState, allowFrontMatter: false, out _, syntaxChildren, lineOffset: 0);
        return (nestedDoc.Blocks, RemapNestedSyntaxNodes(sourceLines, syntaxChildren));
    }

    private static IReadOnlyList<MarkdownSyntaxNode> RemapNestedSyntaxNodes(
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        IReadOnlyList<MarkdownSyntaxNode> syntaxChildren) {
        if (sourceLines == null || sourceLines.Count == 0 || syntaxChildren == null || syntaxChildren.Count == 0) {
            return syntaxChildren ?? Array.Empty<MarkdownSyntaxNode>();
        }

        var remapped = new List<MarkdownSyntaxNode>(syntaxChildren.Count);
        for (int i = 0; i < syntaxChildren.Count; i++) {
            remapped.Add(RemapNestedSyntaxNode(sourceLines, syntaxChildren[i]));
        }

        return remapped;
    }

    private static MarkdownSyntaxNode RemapNestedSyntaxNode(
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        MarkdownSyntaxNode node) {
        var span = RemapNestedSourceSpan(sourceLines, node.SourceSpan);
        IReadOnlyList<MarkdownSyntaxNode> children = node.Children;
        if (node.Children.Count > 0) {
            var remappedChildren = new List<MarkdownSyntaxNode>(node.Children.Count);
            for (int i = 0; i < node.Children.Count; i++) {
                remappedChildren.Add(RemapNestedSyntaxNode(sourceLines, node.Children[i]));
            }

            children = remappedChildren;
        }

        return new MarkdownSyntaxNode(node.Kind, span, node.Literal, children, node.AssociatedObject, node.CustomKind);
    }

    private static MarkdownSourceSpan? RemapNestedSourceSpan(
        IReadOnlyList<MarkdownSourceLineSlice> sourceLines,
        MarkdownSourceSpan? span) {
        if (!span.HasValue) {
            return null;
        }

        var value = span.Value;
        int startIndex = value.StartLine - 1;
        int endIndex = value.EndLine - 1;
        if (startIndex < 0 || startIndex >= sourceLines.Count || endIndex < 0 || endIndex >= sourceLines.Count) {
            return value;
        }

        int startLine = sourceLines[startIndex].AbsoluteLine;
        int endLine = sourceLines[endIndex].AbsoluteLine;
        if (!value.StartColumn.HasValue || !value.EndColumn.HasValue) {
            return new MarkdownSourceSpan(startLine, endLine);
        }

        int startColumn = sourceLines[startIndex].StartColumn + value.StartColumn.Value - 1;
        int endColumn = sourceLines[endIndex].StartColumn + value.EndColumn.Value - 1;
        return new MarkdownSourceSpan(startLine, startColumn, endLine, endColumn);
    }
}
