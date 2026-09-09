using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>Route-specific settings captured before a conversion starts. Unspecified values retain each route's defaults.</summary>
public sealed class OfficeWorkflowConversionOptions {
    /// <summary>Optional source PDF page ranges, for example 1-3,5. Other source formats do not accept this setting.</summary>
    public string? PageRanges { get; set; }
    /// <summary>PDF-to-Word reconstruction or rendered-page strategy.</summary>
    public PdfWordImportMode? WordMode { get; set; }
    /// <summary>PDF-to-PowerPoint reconstruction or rendered-page strategy.</summary>
    public PdfPowerPointImportMode? PowerPointMode { get; set; }
    /// <summary>Resolution for visual PDF-to-Word or PDF-to-PowerPoint pages, from 36 through 600 DPI.</summary>
    public double? RasterDpi { get; set; }
    /// <summary>Worksheet canvas or flowing table layout for Excel-to-PDF conversion.</summary>
    public ExcelPdfWorksheetLayoutMode? WorksheetLayout { get; set; }
    /// <summary>Semantic or positioned review HTML for PDF-to-HTML conversion.</summary>
    public PdfHtmlProfile? HtmlProfile { get; set; }
    /// <summary>Applies verified lossless compression after conversion to PDF. Other destination formats reject true.</summary>
    public bool CompressPdfOutput { get; set; }

    /// <summary>Creates an independent settings copy.</summary>
    public OfficeWorkflowConversionOptions Clone() => (OfficeWorkflowConversionOptions)MemberwiseClone();

    internal OfficeWorkflowConversionOptions Snapshot(OfficeWorkflowRoute route) {
        OfficeWorkflowConversionOptions copy = Clone();
        if (!string.IsNullOrWhiteSpace(copy.PageRanges)) {
            if (!route.SupportsPageSelection) throw new ArgumentException("Page ranges require a PDF input route.");
            if (copy.PageRanges.Length > 2048) throw new ArgumentException("The page selection is too long.");
            _ = PdfPageSelection.Parse(copy.PageRanges);
        }
        if (copy.WordMode.HasValue && (route.Id != "pdf-docx" || !Enum.IsDefined(copy.WordMode.Value)))
            throw new ArgumentException("Choose a supported PDF-to-Word import mode.");
        if (copy.PowerPointMode.HasValue && (route.Id != "pdf-pptx" || !Enum.IsDefined(copy.PowerPointMode.Value)))
            throw new ArgumentException("Choose a supported PDF-to-PowerPoint import mode.");
        if (copy.WorksheetLayout.HasValue && (route.Id != "xlsx-pdf" || !Enum.IsDefined(copy.WorksheetLayout.Value)))
            throw new ArgumentException("Worksheet layout is valid only for Excel-to-PDF conversion.");
        if (copy.HtmlProfile.HasValue && (route.Id != "pdf-html" || !Enum.IsDefined(copy.HtmlProfile.Value)))
            throw new ArgumentException("HTML layout is valid only for PDF-to-HTML conversion.");
        if (copy.CompressPdfOutput && !route.SupportsPdfCompression)
            throw new ArgumentException("PDF compression requires a PDF output route.");
        bool visual = copy.WordMode == PdfWordImportMode.VisualPages || copy.PowerPointMode is
            PdfPowerPointImportMode.VisualPages or PdfPowerPointImportMode.HybridVisualAndEditableTables;
        if (copy.RasterDpi is double dpi && (!visual || !double.IsFinite(dpi) || dpi < 36 || dpi > 600))
            throw new ArgumentException("Rendering resolution requires visual PDF pages and must be between 36 and 600 DPI.");
        return copy;
    }

    internal PdfReadOptions? CreateReadOptions() => string.IsNullOrWhiteSpace(PageRanges) ? null
        : new PdfReadOptions { PageSelection = PdfPageSelection.Parse(PageRanges) };
}
