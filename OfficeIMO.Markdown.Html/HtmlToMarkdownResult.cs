using OfficeIMO.Html;
using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

/// <summary>Result of converting a shared HTML document into the native Markdown model.</summary>
public sealed class HtmlToMarkdownResult : HtmlConversionResult<MarkdownDoc> {
    internal HtmlToMarkdownResult(MarkdownDoc document, IEnumerable<HtmlDiagnostic>? diagnostics = null) : base(document) {
        if (diagnostics != null) AddDiagnostics(diagnostics);

        var retainedHtml = new RetainedHtmlVisitor();
        retainedHtml.Visit(document);
        if (retainedHtml.Total > 0) {
            AddDiagnostic(new HtmlDiagnostic(
                "OfficeIMO.Markdown.Html",
                HtmlConversionDiagnosticCodes.ContentApproximated,
                "HTML-dependent content was retained instead of becoming native Markdown content.",
                HtmlDiagnosticSeverity.Warning,
                source: "html:document",
                detail: $"rawBlocks={retainedHtml.RawBlocks}; rawInlines={retainedHtml.RawInlines}; htmlTagInlines={retainedHtml.HtmlTagInlines}",
                lossKind: OfficeConversionLossKind.Approximation));
        }
    }

    /// <summary>Number of top-level Markdown blocks produced by the conversion.</summary>
    public int Blocks => Value.Blocks.Count;

    private sealed class RetainedHtmlVisitor : MarkdownVisitor {
        internal int RawBlocks { get; private set; }
        internal int RawInlines { get; private set; }
        internal int HtmlTagInlines { get; private set; }
        internal int Total => RawBlocks + RawInlines + HtmlTagInlines;

        protected override void VisitHtmlRawBlock(HtmlRawBlock block) {
            RawBlocks++;
            base.VisitHtmlRawBlock(block);
        }

        protected override void VisitHtmlRawInline(HtmlRawInline inline) {
            RawInlines++;
            base.VisitHtmlRawInline(inline);
        }

        protected override void VisitHtmlTagSequenceInline(HtmlTagSequenceInline inline) {
            HtmlTagInlines++;
            base.VisitHtmlTagSequenceInline(inline);
        }
    }
}
