using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Result of a PDF to HTML export, including generated HTML and machine-readable proof metadata.
/// </summary>
public sealed class PdfHtmlConversionResult : OfficeConversionResult<string, PdfCore.PdfConversionReport> {
    internal PdfHtmlConversionResult(string html, PdfHtmlExportSummary summary, PdfCore.PdfConversionReport conversionReport)
        : base(html, conversionReport.Snapshot()) {
        Summary = summary;
    }

    /// <summary>Machine-readable summary of selected pages, preserved logical objects, and output policy.</summary>
    public PdfHtmlExportSummary Summary { get; }

}
