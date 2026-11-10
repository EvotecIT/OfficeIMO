namespace OfficeIMO.Markdown;

/// <summary>
/// TeX math inline parsed from <c>$...$</c> or <c>$$...$$</c> when <see cref="MarkdownLoadOptions.Math"/> is enabled.
/// </summary>
/// <remarks>
/// The node carries TeX source only. Converters lay it out through their own math engines, and HTML output keeps
/// <c>\(...\)</c> or <c>\[...\]</c> delimiters for client-side renderers such as KaTeX or MathJax.
/// </remarks>
public sealed class MathInline : MarkdownInline, IRenderableMarkdownInline, IPlainTextMarkdownInline, ISyntaxMarkdownInline {
    /// <summary>Creates an inline math node.</summary>
    /// <param name="source">TeX source without delimiters.</param>
    /// <param name="isDisplay">Whether the source used display delimiters (<c>$$...$$</c>).</param>
    public MathInline(string source, bool isDisplay = false) {
        Source = source ?? string.Empty;
        IsDisplay = isDisplay;
    }

    /// <summary>TeX source without delimiters.</summary>
    public string Source { get; }

    /// <summary>Whether the math uses display style (<c>$$...$$</c>).</summary>
    public bool IsDisplay { get; }

    /// <summary>Source span for the TeX source between the delimiters when parsed from Markdown.</summary>
    public MarkdownSourceSpan? ContentSourceSpan { get; internal set; }

    private string Delimiter => IsDisplay ? "$$" : "$";

    // TeX ignores spaces at the edges of math, while the Markdown reader requires non-space characters
    // next to a single $, so the edges are trimmed to keep projected math reparsable.
    string IRenderableMarkdownInline.RenderMarkdown() => Delimiter + Source.Trim() + Delimiter;

    string IRenderableMarkdownInline.RenderHtml() {
        string source = HtmlTextEncoder.Encode(Source, HtmlRenderContext.Options);
        return IsDisplay
            ? "<span class=\"math math-display\">\\[" + source + "\\]</span>"
            : "<span class=\"math math-inline\">\\(" + source + "\\)</span>";
    }

    void IPlainTextMarkdownInline.AppendPlainText(System.Text.StringBuilder sb) => sb.Append(Source);

    MarkdownSyntaxNode ISyntaxMarkdownInline.BuildSyntaxNode(MarkdownInlineSyntaxBuilderContext context, MarkdownSourceSpan? span) =>
        new MarkdownSyntaxNode(
            MarkdownSyntaxKind.InlineMath,
            span,
            Source,
            new[] { new MarkdownSyntaxNode(MarkdownSyntaxKind.InlineMathContent, ContentSourceSpan, Source) },
            this);
}
