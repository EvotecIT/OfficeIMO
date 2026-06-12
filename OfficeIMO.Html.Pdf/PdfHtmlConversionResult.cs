using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Result of a PDF to HTML export, including generated HTML and machine-readable proof metadata.
/// </summary>
public sealed class PdfHtmlConversionResult {
    internal PdfHtmlConversionResult(string html, PdfHtmlExportSummary summary, PdfCore.PdfConversionReport conversionReport) {
        Html = html;
        Summary = summary;
        ConversionReport = SnapshotReport(conversionReport);
    }

    /// <summary>Generated HTML output.</summary>
    public string Html { get; }

    /// <summary>Machine-readable summary of selected pages, preserved logical objects, and output policy.</summary>
    public PdfHtmlExportSummary Summary { get; }

    /// <summary>Conversion report snapshot populated during export.</summary>
    public PdfCore.PdfConversionReport ConversionReport { get; }

    private static PdfCore.PdfConversionReport SnapshotReport(PdfCore.PdfConversionReport conversionReport) {
        var snapshot = new PdfCore.PdfConversionReport();
        snapshot.AddRange(conversionReport.Warnings);
        return snapshot;
    }
}
