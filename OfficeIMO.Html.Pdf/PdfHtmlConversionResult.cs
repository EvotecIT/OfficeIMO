using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Result of a PDF to HTML export, including generated HTML and machine-readable proof metadata.
/// </summary>
public sealed class PdfHtmlConversionResult {
    internal PdfHtmlConversionResult(string html, PdfHtmlExportSummary summary, PdfCore.PdfConversionReport conversionReport) {
        Value = html;
        Summary = summary;
        Report = SnapshotReport(conversionReport);
    }

    /// <summary>Generated HTML output.</summary>
    public string Value { get; }

    /// <summary>Machine-readable summary of selected pages, preserved logical objects, and output policy.</summary>
    public PdfHtmlExportSummary Summary { get; }

    /// <summary>Conversion report snapshot populated during export.</summary>
    public PdfCore.PdfConversionReport Report { get; }

    /// <summary>Returns the generated HTML output.</summary>
    public string RequireValue() => Value;

    private static PdfCore.PdfConversionReport SnapshotReport(PdfCore.PdfConversionReport conversionReport) {
        var snapshot = new PdfCore.PdfConversionReport();
        snapshot.AddRange(conversionReport.Warnings);
        return snapshot;
    }
}
