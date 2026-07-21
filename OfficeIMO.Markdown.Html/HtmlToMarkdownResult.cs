using OfficeIMO.Html;
using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

/// <summary>Result of converting a shared HTML document into the native Markdown model.</summary>
public sealed class HtmlToMarkdownResult : HtmlConversionResult<MarkdownDoc> {
    internal HtmlToMarkdownResult(MarkdownDoc document, IEnumerable<HtmlDiagnostic>? diagnostics = null) : base(document) {
        if (diagnostics != null) AddDiagnostics(diagnostics);
    }

    /// <summary>Number of top-level Markdown blocks produced by the conversion.</summary>
    public int Blocks => Value.Blocks.Count;
}
