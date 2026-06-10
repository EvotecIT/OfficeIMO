namespace OfficeIMO.Markdown;

/// <summary>
/// Native, AST-backed projection of a parsed markdown document for UI hosts that need structured blocks and source spans.
/// </summary>
public sealed class MarkdownNativeDocument {
    private MarkdownNativeDocument(MarkdownParseResult parseResult, IReadOnlyList<MarkdownNativeBlock> blocks) {
        ParseResult = parseResult ?? throw new ArgumentNullException(nameof(parseResult));
        Document = parseResult.Document;
        SyntaxTree = parseResult.SyntaxTree;
        FinalSyntaxTree = parseResult.FinalSyntaxTree;
        TransformDiagnostics = parseResult.TransformDiagnostics;
        Blocks = blocks ?? Array.Empty<MarkdownNativeBlock>();
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

    /// <summary>Top-level native block projection in document order.</summary>
    public IReadOnlyList<MarkdownNativeBlock> Blocks { get; }

    /// <summary>
    /// Parses markdown into the typed object model, syntax tree, diagnostics, and native block projection.
    /// </summary>
    public static MarkdownNativeDocument Parse(string markdown, MarkdownReaderOptions? options = null) {
        var parseResult = MarkdownReader.ParseWithSyntaxTreeAndDiagnostics(markdown ?? string.Empty, options);
        return FromParseResult(parseResult);
    }

    /// <summary>
    /// Builds a native projection from an existing syntax-backed parse result.
    /// </summary>
    public static MarkdownNativeDocument FromParseResult(MarkdownParseResult parseResult) {
        if (parseResult == null) {
            throw new ArgumentNullException(nameof(parseResult));
        }

        var blocks = new List<MarkdownNativeBlock>();
        var children = parseResult.FinalSyntaxTree.Children;
        for (var i = 0; i < children.Count; i++) {
            var block = CreateNativeBlock(children[i]);
            if (block != null) {
                blocks.Add(block);
            }
        }

        return new MarkdownNativeDocument(parseResult, blocks);
    }

    /// <summary>Finds the first top-level native block whose source span contains the supplied 1-based line.</summary>
    public MarkdownNativeBlock? FindBlockAtLine(int lineNumber) {
        for (var i = 0; i < Blocks.Count; i++) {
            if (Blocks[i].ContainsLine(lineNumber)) {
                return Blocks[i];
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

    private static MarkdownNativeBlock? CreateNativeBlock(MarkdownSyntaxNode syntaxNode) {
        if (syntaxNode?.AssociatedObject is not IMarkdownBlock block) {
            return null;
        }

        switch (block) {
            case ParagraphBlock paragraph:
                return new MarkdownNativeParagraphBlock(paragraph, syntaxNode);
            case CodeBlock code:
                return new MarkdownNativeCodeBlock(code, syntaxNode);
            case SemanticFencedBlock visual:
                return new MarkdownNativeVisualBlock(visual, syntaxNode);
            case TableBlock table:
                return new MarkdownNativeTableBlock(table, syntaxNode);
            default:
                return new MarkdownNativeOtherBlock(block, syntaxNode);
        }
    }
}
