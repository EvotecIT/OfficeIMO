using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>Internal direct-render output before shared PDF diagnostics are projected.</summary>
internal sealed class HtmlPdfRenderResult {
    internal HtmlPdfRenderResult(PdfCore.PdfDocument document, HtmlDiagnosticReport diagnostics, PdfCore.PdfConversionReport conversionReport) {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        Diagnostics = diagnostics ?? throw new ArgumentNullException(nameof(diagnostics));
        ConversionReport = conversionReport ?? throw new ArgumentNullException(nameof(conversionReport));
    }

    internal PdfCore.PdfDocument Document { get; }

    internal HtmlDiagnosticReport Diagnostics { get; }

    internal PdfCore.PdfConversionReport ConversionReport { get; }
}
