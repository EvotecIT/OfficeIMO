using System.Text.Json;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Workflows;

/// <summary>Portable scalar conversion choices shared by saved-job hosts.</summary>
public sealed class OfficeConversionJobOptions {
    /// <summary>Source PDF pages, such as 1-3,last.</summary>
    public string? PageRanges { get; set; }
    /// <summary>PDF-to-Word reconstruction mode.</summary>
    public PdfWordImportMode? WordMode { get; set; }
    /// <summary>PDF-to-PowerPoint reconstruction mode.</summary>
    public PdfPowerPointImportMode? PowerPointMode { get; set; }
    /// <summary>Visual-page raster density.</summary>
    public double? RasterDpi { get; set; }
    /// <summary>Excel worksheet projection mode.</summary>
    public ExcelPdfWorksheetLayoutMode? WorksheetLayout { get; set; }
    /// <summary>PDF-to-HTML layout mode.</summary>
    public PdfHtmlProfile? HtmlProfile { get; set; }
    /// <summary>Verified lossless compression for PDF destinations.</summary>
    public bool CompressPdfOutput { get; set; }
    /// <summary>Explicit legacy DOC loss acceptance.</summary>
    public OfficeConversionLossPolicy LegacyDocLossPolicy { get; set; } = OfficeConversionLossPolicy.Block;
    /// <summary>Reject reported loss on routes supporting strict loss acceptance.</summary>
    public bool RequireNoLoss { get; set; }
    /// <summary>Literal-text encoding; null uses strict BOM-aware UTF-8 defaults.</summary>
    public string? TextEncoding { get; set; }
    /// <summary>Literal-text tab columns.</summary>
    public int TabSize { get; set; } = 8;

    internal OfficeWorkflowConversionOptions CreateOptions() => new() {
        PageRanges = PageRanges, WordMode = WordMode, PowerPointMode = PowerPointMode,
        RasterDpi = RasterDpi, WorksheetLayout = WorksheetLayout, HtmlProfile = HtmlProfile,
        CompressPdfOutput = CompressPdfOutput, LegacyDocLossPolicy = LegacyDocLossPolicy,
        RequireNoLoss = RequireNoLoss, PlainText = new PdfPlainTextOptions { EncodingName = TextEncoding, TabSize = TabSize }
    };

    internal static OfficeConversionJobOptions Capture(OfficeWorkflowConversionOptions options) {
        ArgumentNullException.ThrowIfNull(options);
        if (options.Word != null || options.Excel != null || options.PowerPoint != null || options.Html != null ||
            options.Markdown != null || options.Rtf != null || options.Draw != null || options.Visio != null ||
            options.Xps != null || options.DjVu != null || options.Chm != null || options.PublisherRead != null || options.PublisherPdf != null)
            throw new NotSupportedException("This saved-job version supports portable scalar settings only. Advanced native renderer settings cannot be captured.");
        if (options.PlainText is { } text) {
            var defaults = new PdfPlainTextOptions();
            if (text.MaximumCharacters != defaults.MaximumCharacters || text.MaximumPages != defaults.MaximumPages ||
                text.PdfOptions == null || text.PdfOptions.CheckpointEmbeddedFonts.Count != 0 ||
                text.PdfOptions.CheckpointExplicitSettings != defaults.PdfOptions.CheckpointExplicitSettings ||
                JsonSerializer.Serialize(text.PdfOptions, OfficeConversionPdfSettingsJsonContext.Default.PdfOptions) !=
                JsonSerializer.Serialize(defaults.PdfOptions, OfficeConversionPdfSettingsJsonContext.Default.PdfOptions))
                throw new NotSupportedException("Custom literal-text PDF settings cannot be captured by this saved-job version.");
        }
        return new OfficeConversionJobOptions {
            PageRanges = options.PageRanges, WordMode = options.WordMode, PowerPointMode = options.PowerPointMode,
            RasterDpi = options.RasterDpi, WorksheetLayout = options.WorksheetLayout, HtmlProfile = options.HtmlProfile,
            CompressPdfOutput = options.CompressPdfOutput, LegacyDocLossPolicy = options.LegacyDocLossPolicy,
            RequireNoLoss = options.RequireNoLoss, TextEncoding = options.PlainText?.EncodingName,
            TabSize = options.PlainText?.TabSize ?? 8
        };
    }
}
