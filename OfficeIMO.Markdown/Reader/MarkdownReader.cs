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
        return ParseInternal(markdown, options, state, allowFrontMatter: true, out _, out _);
    }

    /// <summary>
    /// Parses Markdown text into both the object model and a lightweight syntax tree with source spans.
    /// </summary>
    public static MarkdownParseResult ParseWithSyntaxTree(string markdown, MarkdownReaderOptions? options = null) {
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
        out string normalizedSourceText,
        List<MarkdownSyntaxNode>? syntaxNodes = null,
        int lineOffset = 0,
        ICollection<MarkdownDocumentTransformDiagnostic>? transformDiagnostics = null,
        bool applyDocumentTransforms = true) {
        var doc = MarkdownDoc.Create();
        syntaxTree = syntaxNodes != null ? BuildDocumentSyntaxTree(syntaxNodes, doc) : null;
        normalizedSourceText = string.Empty;
        if (string.IsNullOrEmpty(markdown)) return doc;
        int previousLineOffset = state.SourceLineOffset;
        var previousSourceTextMap = state.SourceTextMap;
        state.SourceLineOffset = lineOffset;

        try {
            var text = PrepareMarkdownForParsing(markdown, options);
            normalizedSourceText = text;
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
                options.PreserveTrivia);
            return transformed;
        } finally {
            state.SourceLineOffset = previousLineOffset;
            state.SourceTextMap = previousSourceTextMap;
        }
    }

    private static string PrepareMarkdownForParsing(string markdown, MarkdownReaderOptions options) {
        markdown ??= string.Empty;
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

        return markdown.Replace("\r\n", "\n").Replace('\r', '\n');
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
