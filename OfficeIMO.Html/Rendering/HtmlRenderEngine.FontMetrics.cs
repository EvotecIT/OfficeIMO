namespace OfficeIMO.Html;

public static partial class HtmlRenderEngine {
    private static HtmlRenderDocument RenderWithPreparedFontMetrics(HtmlRenderOptions options,
        HtmlDiagnosticReport diagnostics, Func<HtmlDiagnosticReport, bool, HtmlRenderDocument> render,
        CancellationToken cancellationToken) {
        if (options is not Pdf.HtmlToPdfOptions pdfOptions) return CompleteRender(render(diagnostics, false), options);
        // Discard provisional layout diagnostics; stylesheet/resource diagnostics
        // remain authoritative and both passes consume the same operation budget.
        HtmlRenderDocument rendered = render(diagnostics.Clone(), false);
        if (Pdf.HtmlPdfRenderedConverter.TryPrepareLibraryFontMetrics(rendered, pdfOptions, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            rendered = render(diagnostics.Clone(), true);
        }
        return CompleteRender(rendered, options);
    }
}
