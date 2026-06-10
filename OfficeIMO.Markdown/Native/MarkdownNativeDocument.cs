namespace OfficeIMO.Markdown;

/// <summary>
/// Native, AST-backed projection of a parsed markdown document for UI hosts that need structured blocks and source spans.
/// </summary>
public sealed class MarkdownNativeDocument {
    private MarkdownNativeDocument(
        MarkdownParseResult parseResult,
        string sourceMarkdown,
        MarkdownNativeDocumentSourceKind sourceKind,
        IReadOnlyList<MarkdownNativeBlock> blocks,
        IReadOnlyList<MarkdownNativeDiagnostic> diagnostics) {
        ParseResult = parseResult ?? throw new ArgumentNullException(nameof(parseResult));
        Document = parseResult.Document;
        SyntaxTree = parseResult.SyntaxTree;
        FinalSyntaxTree = parseResult.FinalSyntaxTree;
        TransformDiagnostics = parseResult.TransformDiagnostics;
        SourceMarkdown = sourceMarkdown ?? string.Empty;
        SourceKind = sourceKind;
        Blocks = blocks ?? Array.Empty<MarkdownNativeBlock>();
        Diagnostics = diagnostics ?? Array.Empty<MarkdownNativeDiagnostic>();
    }

    /// <summary>Underlying parse result, including original/final syntax trees and diagnostics.</summary>
    public MarkdownParseResult ParseResult { get; }

    /// <summary>Parsed OfficeIMO markdown document.</summary>
    public MarkdownDoc Document { get; }

    /// <summary>Original syntax tree produced before document transforms were applied.</summary>
    public MarkdownSyntaxNode SyntaxTree { get; }

    /// <summary>Final syntax tree aligned with <see cref="Document"/>.</summary>
    public MarkdownSyntaxNode FinalSyntaxTree { get; }

    /// <summary>Document-transform diagnostics captured during parsing.</summary>
    public IReadOnlyList<MarkdownDocumentTransformDiagnostic> TransformDiagnostics { get; }

    /// <summary>Markdown source text whose source spans back this projection.</summary>
    public string SourceMarkdown { get; }

    /// <summary>Identifies whether <see cref="SourceMarkdown"/> is direct reader input or renderer-preprocessed markdown.</summary>
    public MarkdownNativeDocumentSourceKind SourceKind { get; }

    /// <summary>Top-level native block projection in document order.</summary>
    public IReadOnlyList<MarkdownNativeBlock> Blocks { get; }

    /// <summary>Projection diagnostics including transform notices and unsupported block fallbacks.</summary>
    public IReadOnlyList<MarkdownNativeDiagnostic> Diagnostics { get; }

    /// <summary>
    /// Parses markdown into the typed object model, syntax tree, diagnostics, and native block projection.
    /// </summary>
    public static MarkdownNativeDocument Parse(string markdown, MarkdownReaderOptions? options = null) {
        markdown ??= string.Empty;
        var parseResult = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown, options);
        return FromParseResult(parseResult, markdown, MarkdownNativeDocumentSourceKind.ReaderInput);
    }

    /// <summary>
    /// Builds a native projection from an existing syntax-backed parse result.
    /// </summary>
    public static MarkdownNativeDocument FromParseResult(
        MarkdownParseResult parseResult,
        string? sourceMarkdown = null,
        MarkdownNativeDocumentSourceKind sourceKind = MarkdownNativeDocumentSourceKind.ReaderInput) {
        if (parseResult == null) {
            throw new ArgumentNullException(nameof(parseResult));
        }

        var blocks = new List<MarkdownNativeBlock>();
        var diagnostics = new List<MarkdownNativeDiagnostic>();
        for (var i = 0; i < parseResult.TransformDiagnostics.Count; i++) {
            diagnostics.Add(MarkdownNativeDiagnostic.FromTransform(parseResult.TransformDiagnostics[i]));
        }

        var children = parseResult.FinalSyntaxTree.Children;
        for (var i = 0; i < children.Count; i++) {
            var block = MarkdownNativeProjectionFactory.Create(children[i], diagnostics);
            if (block != null) {
                blocks.Add(block);
            }
        }

        return new MarkdownNativeDocument(parseResult, sourceMarkdown ?? string.Empty, sourceKind, blocks, diagnostics);
    }

    /// <summary>Finds the first native block whose source span contains the supplied 1-based line.</summary>
    public MarkdownNativeBlock? FindBlockAtLine(int lineNumber) {
        for (var i = 0; i < Blocks.Count; i++) {
            var match = FindBlockAtLine(Blocks[i], lineNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }

    /// <summary>Enumerates top-level native blocks of the requested projection type.</summary>
    public IEnumerable<TBlock> BlocksOfType<TBlock>() where TBlock : MarkdownNativeBlock {
        for (var i = 0; i < Blocks.Count; i++) {
            if (Blocks[i] is TBlock block) {
                yield return block;
            }
        }
    }

    private static MarkdownNativeBlock? FindBlockAtLine(MarkdownNativeBlock block, int lineNumber) {
        switch (block) {
            case MarkdownNativeQuoteBlock quote:
                return FindChildBlockAtLine(quote.Children, lineNumber) ?? (quote.ContainsLine(lineNumber) ? quote : null);
            case MarkdownNativeCalloutBlock callout:
                return FindChildBlockAtLine(callout.Children, lineNumber) ?? (callout.ContainsLine(lineNumber) ? callout : null);
            case MarkdownNativeDetailsBlock details:
                return FindChildBlockAtLine(details.Children, lineNumber) ?? (details.ContainsLine(lineNumber) ? details : null);
            case MarkdownNativeListBlock list:
                for (var i = 0; i < list.Items.Count; i++) {
                    var itemMatch = FindChildBlockAtLine(list.Items[i].Children, lineNumber);
                    if (itemMatch != null) {
                        return itemMatch;
                    }
                }

                return list.ContainsLine(lineNumber) ? list : null;
            default:
                return block.ContainsLine(lineNumber) ? block : null;
        }
    }

    private static MarkdownNativeBlock? FindChildBlockAtLine(IReadOnlyList<MarkdownNativeBlock> children, int lineNumber) {
        for (var i = 0; i < children.Count; i++) {
            var match = FindBlockAtLine(children[i], lineNumber);
            if (match != null) {
                return match;
            }
        }

        return null;
    }
}
