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
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        options ??= new MarkdownReaderOptions();

        // Without document transforms the syntax captured by ParseInternal already describes
        // the returned object model. Avoid rebuilding and rebinding an identical final tree,
        // while retaining the parse result needed by source-aware rendering.
        if (BuildEffectiveDocumentTransforms(options).Count == 0) {
            var state = new MarkdownReaderState();
            var syntaxNodes = new List<MarkdownSyntaxNode>();
            var document = ParseInternal(
                markdown,
                options,
                state,
                allowFrontMatter: true,
                out var syntaxTree,
                out var sourceMarkdown,
                syntaxNodes,
                lineOffset: 0,
                transformDiagnostics: null,
                applyDocumentTransforms: false);
            var capturedSyntaxTree = syntaxTree ?? BuildDocumentSyntaxTree(syntaxNodes, document);
            _ = new MarkdownParseResult(
                document,
                capturedSyntaxTree,
                capturedSyntaxTree,
                sourceMarkdown,
                options.PreserveTrivia ? markdown : null,
                options.PreserveTrivia,
                transformDiagnostics: null,
                referenceLinkDefinitions: SnapshotReferenceLinkDefinitions(state),
                abbreviationDefinitions: SnapshotAbbreviationDefinitions(state));
            return document;
        }

        return ParseWithSyntaxTree(markdown, options).Document;
    }

    /// <summary>
    /// Parses Markdown into the typed semantic document without capturing source spans or a syntax tree.
    /// Use <see cref="Parse(string, MarkdownReaderOptions?)"/> or
    /// <see cref="ParseWithSyntaxTree(string, MarkdownReaderOptions?)"/> when source-aware writing,
    /// source spans, or trivia are required.
    /// </summary>
    public static MarkdownDoc ParseSemantic(string markdown, MarkdownReaderOptions? options = null) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        options ??= new MarkdownReaderOptions();
        var state = new MarkdownReaderState();
        var document = ParseInternal(
            markdown,
            options,
            state,
            allowFrontMatter: true,
            out _,
            out _,
            syntaxNodes: null,
            lineOffset: 0,
            transformDiagnostics: null);
        MarkdownObjectTreeBinder.BindDocument(document);
        return document;
    }

    /// <summary>
    /// Parses a transient semantic document for an owning projection layer. The returned model keeps
    /// block and inline values needed for projection, but omits table cell navigation models and
    /// object-tree binding that the caller will discard.
    /// </summary>
    internal static MarkdownDoc ParseSemanticProjection(string markdown, MarkdownReaderOptions? options = null) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        options ??= new MarkdownReaderOptions();
        var state = new MarkdownReaderState { BuildTableCellModels = false };
        return ParseInternal(
            markdown,
            options,
            state,
            allowFrontMatter: true,
            out _,
            out _,
            syntaxNodes: null,
            lineOffset: 0,
            transformDiagnostics: null);
    }

    /// <summary>
    /// Parses Markdown text into both the object model and a lightweight syntax tree with source spans.
    /// </summary>
    public static MarkdownParseResult ParseWithSyntaxTree(string markdown, MarkdownReaderOptions? options = null) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
        options ??= new MarkdownReaderOptions();
        var state = new MarkdownReaderState();
        var syntaxNodes = new List<MarkdownSyntaxNode>();
        var diagnostics = new List<MarkdownDocumentTransformDiagnostic>();
        var document = ParseInternal(markdown, options, state, allowFrontMatter: true, out var syntaxTree, out var sourceMarkdown, syntaxNodes, lineOffset: 0, transformDiagnostics: diagnostics);
        var originalSyntaxTree = syntaxTree ?? BuildDocumentSyntaxTree(syntaxNodes, document);
        if (diagnostics.Any(diagnostic => diagnostic.ReplacedDocument)) {
            originalSyntaxTree = DetachOriginalSyntaxAssociations(originalSyntaxTree);
        }

        var finalSyntaxTree = BuildFinalSyntaxTree(document, originalSyntaxTree, diagnostics);
        MarkdownObjectTreeBinder.BindDocument(document, finalSyntaxTree);
        return new MarkdownParseResult(
            document,
            originalSyntaxTree,
            finalSyntaxTree,
            sourceMarkdown,
            options.PreserveTrivia ? markdown : null,
            options.PreserveTrivia,
            diagnostics,
            referenceLinkDefinitions: SnapshotReferenceLinkDefinitions(state),
            abbreviationDefinitions: SnapshotAbbreviationDefinitions(state));
    }

    /// <summary>
    /// Parses Markdown text into the object model, original syntax tree, and document-transform diagnostics.
    /// </summary>
    public static MarkdownParseResult ParseWithSyntaxTreeAndDiagnostics(string markdown, MarkdownReaderOptions? options = null) {
        if (markdown == null) throw new ArgumentNullException(nameof(markdown));
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
            out var sourceMarkdown,
            syntaxNodes,
            lineOffset: 0,
            transformDiagnostics: diagnostics);
        var originalSyntaxTree = syntaxTree ?? BuildDocumentSyntaxTree(syntaxNodes, document);
        if (diagnostics.Any(diagnostic => diagnostic.ReplacedDocument)) {
            originalSyntaxTree = DetachOriginalSyntaxAssociations(originalSyntaxTree);
        }

        var finalSyntaxTree = BuildFinalSyntaxTree(document, originalSyntaxTree, diagnostics);
        MarkdownObjectTreeBinder.BindDocument(document, finalSyntaxTree);
        return new MarkdownParseResult(
            document,
            originalSyntaxTree,
            finalSyntaxTree,
            sourceMarkdown,
            options.PreserveTrivia ? markdown : null,
            options.PreserveTrivia,
            diagnostics,
            SnapshotReferenceLinkDefinitions(state),
            SnapshotAbbreviationDefinitions(state));
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
        using var objectTreeBindingDeferral = nestedDocument.DeferObjectTreeBinding();
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
        out string normalizedSourceText,
        List<MarkdownSyntaxNode>? syntaxNodes = null,
        int lineOffset = 0,
        ICollection<MarkdownDocumentTransformDiagnostic>? transformDiagnostics = null,
        bool applyDocumentTransforms = true) {
        var doc = MarkdownDoc.Create();
        state.CaptureSyntaxTree = syntaxNodes != null;
        syntaxTree = syntaxNodes != null ? BuildDocumentSyntaxTree(syntaxNodes, doc) : null;
        normalizedSourceText = string.Empty;
        if (string.IsNullOrEmpty(markdown)) return doc;
        int previousLineOffset = state.SourceLineOffset;
        var previousSourceTextMap = state.SourceTextMap;
        state.SourceLineOffset = lineOffset;

        try {
            var text = PrepareMarkdownForParsing(markdown, options, normalizeLineEndings: state.CaptureSyntaxTree);
            normalizedSourceText = text;
            if (state.CaptureSyntaxTree && (lineOffset == 0 || state.SourceTextMap == null)) {
                state.SourceTextMap = new MarkdownSourceTextMap(text);
            }
            var lines = state.CaptureSyntaxTree
                ? text.Split('\n')
                : SplitMarkdownLines(text, reuseRepeatedLines: state.BuildTableCellModels);
            int i = 0;

            // Parsing with syntax capture binds the complete object and source trees together
            // immediately below. Avoid an otherwise duplicate object-only traversal here.
            using (doc.DeferObjectTreeBinding(completeBindingOnDispose: false)) {
                // Front matter (YAML) only if it's the very first thing in the file
                if (allowFrontMatter && options.FrontMatter && i < lines.Length && lines[i].Trim() == "---") {
                    int start = i + 1;
                    int end = -1;
                    for (int j = start; j < lines.Length; j++) { if (lines[j].Trim() == "---") { end = j; break; } }
                    if (end > start) {
                        var frontMatter = ParseFrontMatterBlock(lines, start, end - 1, state);
                        if (frontMatter.Entries.Count > 0 || frontMatter.RawYaml != null) {
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
                PreScanAbbreviationDefinitions(lines, state, options);
                while (i < lines.Length) {
                    if (string.IsNullOrWhiteSpace(lines[i])) { i++; continue; }
                    if (TryConsumeStandaloneGenericAttributeBlock(lines, i, options, state)) { i++; continue; }
                    bool matched = false;
                    var parsers = pipeline.Parsers;
                    int previousBlockCount = doc.Blocks.Count;
                    int startIndex = i;
                    int startLine = lineOffset + i;
                    for (int p = 0; p < parsers.Count; p++) {
                        if (parsers[p].TryParse(lines, ref i, options, doc, state)) {
                            matched = true;
                            if (doc.Blocks.Count > previousBlockCount
                                && TryApplyPendingGenericAttributeBlock(doc, previousBlockCount, startLine, state, out var pendingAttributeStartLine)) {
                                startLine = Math.Min(startLine, pendingAttributeStartLine);
                            }

                            if (syntaxNodes != null && doc.Blocks.Count > previousBlockCount) {
                                CaptureSyntaxNodes(doc, previousBlockCount, startLine, lineOffset + i, syntaxNodes, state);
                            } else if (syntaxNodes != null) {
                                CaptureConsumedSyntaxNodes(parsers[p], lines, startIndex, options, syntaxNodes, state);
                            }
                            break;
                        }
                    }
                    if (!matched) i++; // defensive: avoid infinite loop
                }
            }

            syntaxTree = syntaxNodes != null ? BuildDocumentSyntaxTree(syntaxNodes, doc) : null;
            if (syntaxTree != null) {
                MarkdownObjectTreeBinder.BindDocument(doc, syntaxTree);
            }

            if (!applyDocumentTransforms) {
                return doc;
            }

            var transformed = ApplyDocumentTransforms(
                doc,
                options,
                transformDiagnostics,
                syntaxTree,
                normalizedSourceText,
                options.PreserveTrivia ? markdown : null,
                options.PreserveTrivia,
                skipAbsentRegisteredFenceTransform: !state.BuildTableCellModels);
            return transformed;
        } finally {
            state.SourceLineOffset = previousLineOffset;
            state.SourceTextMap = previousSourceTextMap;
        }
    }

    private static string PrepareMarkdownForParsing(
        string markdown,
        MarkdownReaderOptions options,
        bool normalizeLineEndings = true) {
        markdown ??= string.Empty;
        if (options.MaxNestingDepth < 1) {
            throw new ArgumentOutOfRangeException(
                nameof(options.MaxNestingDepth),
                options.MaxNestingDepth,
                "MaxNestingDepth must be greater than zero.");
        }
        if (markdown.Length == 0) {
            return string.Empty;
        }

        // Normalize BOM (U+FEFF) at the very beginning to avoid blocking heading/html detection.
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

        return normalizeLineEndings
            ? markdown.Replace("\r\n", "\n").Replace('\r', '\n')
            : ExpandTabsForSemanticParsing(markdown);
    }

    private static string ExpandTabsForSemanticParsing(string markdown) {
        int firstTab = markdown.IndexOf('\t');
        if (firstTab < 0) {
            return markdown;
        }

        var builder = new StringBuilder(markdown.Length + 16);
        int column = 0;
        for (int i = 0; i < markdown.Length; i++) {
            char value = markdown[i];
            if (value == '\t') {
                int spaces = 4 - column % 4;
                builder.Append(' ', spaces);
                column += spaces;
                continue;
            }

            builder.Append(value);
            if (value is '\r' or '\n') {
                column = 0;
            } else {
                column++;
            }
        }

        return builder.ToString();
    }

    private static string[] SplitMarkdownLines(string markdown, bool reuseRepeatedLines = true) {
        int carriageReturn = markdown.IndexOf('\r');
        bool shouldReuseRepeatedLines = reuseRepeatedLines && ShouldReuseRepeatedMarkdownLines(markdown);
        if (carriageReturn < 0 && !shouldReuseRepeatedLines) {
            return markdown.Split('\n');
        }

        int lineCount = 1;
        for (int i = 0; i < markdown.Length; i++) {
            if (markdown[i] == '\r') {
                lineCount++;
                if (i + 1 < markdown.Length && markdown[i + 1] == '\n') {
                    i++;
                }
            } else if (markdown[i] == '\n') {
                lineCount++;
            }
        }

        Dictionary<MarkdownLineKey, string>? repeatedLines = shouldReuseRepeatedLines
            ? new Dictionary<MarkdownLineKey, string>(MarkdownLineKeyComparer.Instance)
            : null;
        var lines = new string[lineCount];
        int lineIndex = 0;
        int lineStart = 0;
        for (int i = 0; i < markdown.Length; i++) {
            char value = markdown[i];
            if (value != '\r' && value != '\n') {
                continue;
            }

            lines[lineIndex++] = GetMarkdownLine(markdown, lineStart, i - lineStart, repeatedLines);
            if (value == '\r' && i + 1 < markdown.Length && markdown[i + 1] == '\n') {
                i++;
            }

            lineStart = i + 1;
        }

        lines[lineIndex] = GetMarkdownLine(markdown, lineStart, markdown.Length - lineStart, repeatedLines);
        return lines;
    }

    private static string GetMarkdownLine(
        string markdown,
        int start,
        int length,
        Dictionary<MarkdownLineKey, string>? repeatedLines) {
        if (length == 0) {
            return string.Empty;
        }
        if (repeatedLines == null) {
            return markdown.Substring(start, length);
        }

        var key = new MarkdownLineKey(markdown, start, length);
        if (repeatedLines.TryGetValue(key, out string? existing)) {
            return existing;
        }

        string line = markdown.Substring(start, length);
        repeatedLines.Add(key, line);
        return line;
    }

    private static bool ShouldReuseRepeatedMarkdownLines(string markdown) {
        const int maximumSampleLines = 64;
        const int requiredRepeatedLines = 4;
        if (markdown.Length < 4096) {
            return false;
        }

#if NET8_0_OR_GREATER
        Span<int> starts = stackalloc int[maximumSampleLines];
        Span<int> lengths = stackalloc int[maximumSampleLines];
#else
        var starts = new int[maximumSampleLines];
        var lengths = new int[maximumSampleLines];
#endif
        int lineCount = 0;
        int repeatedLineCount = 0;
        int lineStart = 0;
        for (int index = 0; index <= markdown.Length && lineCount < maximumSampleLines; index++) {
            bool atEnd = index == markdown.Length;
            if (!atEnd && markdown[index] is not ('\r' or '\n')) {
                continue;
            }

            int length = index - lineStart;
            if (length > 0) {
                for (int prior = 0; prior < lineCount; prior++) {
                    if (lengths[prior] == length
                        && MarkdownLineRegionsEqual(markdown, starts[prior], lineStart, length)) {
                        repeatedLineCount++;
                        break;
                    }
                }
            }

            starts[lineCount] = lineStart;
            lengths[lineCount] = length;
            lineCount++;
            if (repeatedLineCount >= requiredRepeatedLines) {
                return true;
            }

            if (!atEnd && markdown[index] == '\r' && index + 1 < markdown.Length && markdown[index + 1] == '\n') {
                index++;
            }
            lineStart = index + 1;
        }

        return false;
    }

    private static bool MarkdownLineRegionsEqual(string markdown, int leftStart, int rightStart, int length) {
        for (int offset = 0; offset < length; offset++) {
            if (markdown[leftStart + offset] != markdown[rightStart + offset]) {
                return false;
            }
        }
        return true;
    }

    private readonly struct MarkdownLineKey {
        internal MarkdownLineKey(string source, int start, int length) {
            Source = source;
            Start = start;
            Length = length;
        }

        internal string Source { get; }
        internal int Start { get; }
        internal int Length { get; }
    }

    private sealed class MarkdownLineKeyComparer : IEqualityComparer<MarkdownLineKey> {
        internal static readonly MarkdownLineKeyComparer Instance = new MarkdownLineKeyComparer();

        public bool Equals(MarkdownLineKey left, MarkdownLineKey right) {
            if (left.Length != right.Length) {
                return false;
            }
            return string.CompareOrdinal(
                left.Source,
                left.Start,
                right.Source,
                right.Start,
                left.Length) == 0;
        }

        public int GetHashCode(MarkdownLineKey key) {
            unchecked {
                int hash = 17;
                for (int offset = 0; offset < key.Length; offset++) {
                    hash = hash * 31 + key.Source[key.Start + offset];
                }
                return hash;
            }
        }
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
}
