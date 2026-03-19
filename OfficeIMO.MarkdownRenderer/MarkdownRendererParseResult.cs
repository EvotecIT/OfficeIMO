using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Result of parsing markdown through the renderer-owned pipeline.
/// </summary>
public sealed class MarkdownRendererParseResult {
    /// <summary>The final renderer-owned markdown document after parsing and renderer transforms.</summary>
    public MarkdownDoc Document { get; }

    /// <summary>The final markdown text after renderer preprocessing and before parsing.</summary>
    public string PreprocessedMarkdown { get; }

    /// <summary>The original syntax tree produced before document transforms were applied.</summary>
    public MarkdownSyntaxNode SyntaxTree { get; }

    /// <summary>Document-transform diagnostics from the reader and renderer pipelines.</summary>
    public IReadOnlyList<MarkdownDocumentTransformDiagnostic> TransformDiagnostics { get; }

    /// <summary>Renderer pre-parse processing diagnostics.</summary>
    public IReadOnlyList<MarkdownRendererPreProcessorDiagnostic> PreProcessorDiagnostics { get; }

    internal MarkdownRendererParseResult(
        MarkdownDoc document,
        string preprocessedMarkdown,
        MarkdownSyntaxNode syntaxTree,
        IReadOnlyList<MarkdownDocumentTransformDiagnostic>? transformDiagnostics = null,
        IReadOnlyList<MarkdownRendererPreProcessorDiagnostic>? preProcessorDiagnostics = null) {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        PreprocessedMarkdown = preprocessedMarkdown ?? string.Empty;
        SyntaxTree = syntaxTree ?? throw new ArgumentNullException(nameof(syntaxTree));
        TransformDiagnostics = transformDiagnostics ?? Array.Empty<MarkdownDocumentTransformDiagnostic>();
        PreProcessorDiagnostics = preProcessorDiagnostics ?? Array.Empty<MarkdownRendererPreProcessorDiagnostic>();
    }
}
