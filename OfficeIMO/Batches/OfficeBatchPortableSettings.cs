using System.Text.Json;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Word;
using OfficeIMO.Pdf.PowerPoint;

namespace OfficeIMO;

// Saved definitions retain the existing portable settings boundary. Native renderer
// instances, callbacks and credentials remain execution-only rather than silently disappearing.
internal sealed class OfficeBatchPortableSettings {
    public bool FailOnLoss { get; set; }
    public bool WordFailOnLoss { get; set; }
    public bool CompressPdfOutput { get; set; }
    public string? PageRanges { get; set; }
    public WordPdfImportMode? WordMode { get; set; }
    public PowerPointPdfImportMode? PowerPointMode { get; set; }
    public double? RasterDpi { get; set; }
    public ExcelPdfWorksheetLayoutMode? WorksheetLayout { get; set; }
    public PdfHtmlProfile? HtmlProfile { get; set; }
    public string? TextEncoding { get; set; }
    public int TabSize { get; set; } = 8;

    internal OfficeConversionOptions CreateOptions() => new() {
        FailOnLoss = FailOnLoss, WordLoad = WordFailOnLoss ? new OfficeIMO.Word.WordLoadOptions { FailOnLoss = true } : null, CompressPdfOutput = CompressPdfOutput,
        PdfLoad = PageRanges == null ? null : new PdfLoadOptions { PageSelection = PdfPageSelection.Parse(PageRanges) },
        PdfWord = WordMode == null ? null : new WordPdfImportOptions { Mode = WordMode.Value, Dpi = RasterDpi ?? 144D },
        PdfPowerPoint = PowerPointMode == null ? null : new PowerPointPdfImportOptions { Mode = PowerPointMode.Value, Dpi = RasterDpi ?? 144D },
        Excel = WorksheetLayout == null ? null : new ExcelPdfExportOptions { WorksheetLayout = WorksheetLayout.Value },
        PdfHtml = HtmlProfile == null ? null : new PdfToHtmlOptions { Profile = HtmlProfile.Value },
        PlainText = TextEncoding == null && TabSize == 8 ? null : new PdfTextImportOptions { EncodingName = TextEncoding, TabSize = TabSize }
    };

    internal static OfficeBatchPortableSettings Capture(OfficeConversionOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        var result = new OfficeBatchPortableSettings {
            FailOnLoss = options.FailOnLoss, WordFailOnLoss = options.WordLoad?.FailOnLoss == true, CompressPdfOutput = options.CompressPdfOutput, PageRanges = options.PdfLoad?.PageSelection?.ToString(),
            WordMode = options.PdfWord?.Mode, PowerPointMode = options.PdfPowerPoint?.Mode,
            RasterDpi = options.PdfWord?.Dpi ?? options.PdfPowerPoint?.Dpi,
            WorksheetLayout = options.Excel?.WorksheetLayout, HtmlProfile = options.PdfHtml?.Profile,
            TextEncoding = options.PlainText?.EncodingName, TabSize = options.PlainText?.TabSize ?? 8
        };
        if (result.TextEncoding?.Length > 128) throw new InvalidDataException("Text encoding names cannot exceed 128 characters.");
        OfficeConversionOptions baseline = result.CreateOptions();
        if (options.SourcePassword != null || options.Word != null || options.PowerPoint != null || options.Html != null ||
            options.Markdown != null || options.ExcelLoad != null || options.PowerPointLoad != null ||
            options.WordSave != null || options.ExcelSave != null || options.PowerPointSave != null || options.RegisteredSettings != null ||
            options.Reader != null || options.ReaderPdf != null || options.ReaderEmail != null || options.ReaderEpub != null || options.ReaderVisio != null ||
            options.PdfExcel != null || options.WordMarkdown != null || options.MarkdownWord != null || options.WordHtml != null ||
            options.HtmlWord != null || options.MarkdownHtml != null || options.HtmlMarkdown != null || options.HtmlLoad != null ||
            options.CompatibilityMode != baseline.CompatibilityMode || options.EmbedSourceWhenLossy ||
            options.VisualFallbackMaxColumns != baseline.VisualFallbackMaxColumns || options.VisualFallbackMaxRows != baseline.VisualFallbackMaxRows)
            throw new NotSupportedException("Saved batches support portable scalar settings and secret references. Advanced renderer objects and runtime callbacks remain execution-only.");
        if (!Same(options.WordLoad, baseline.WordLoad, OfficeBatchConfigurationJsonContext.Default.WordLoadOptions) ||
            !Same(options.Excel, baseline.Excel, OfficeBatchConfigurationJsonContext.Default.ExcelPdfExportOptions) ||
            !Same(options.PdfWord, baseline.PdfWord, OfficeBatchConfigurationJsonContext.Default.WordPdfImportOptions) ||
            !Same(options.PdfPowerPoint, baseline.PdfPowerPoint, OfficeBatchConfigurationJsonContext.Default.PowerPointPdfImportOptions) ||
            !Same(options.PdfHtml, baseline.PdfHtml, OfficeBatchConfigurationJsonContext.Default.PdfToHtmlOptions) ||
            !Same(options.PlainText, baseline.PlainText, OfficeBatchConfigurationJsonContext.Default.PdfTextImportOptions) ||
            !Same(options.PdfLoad, baseline.PdfLoad, OfficeBatchConfigurationJsonContext.Default.PdfLoadOptions))
            throw new NotSupportedException("Advanced native settings cannot be captured by this saved-batch version.");
        return result;
    }
    private static bool Same<T>(T? value, T? baseline, System.Text.Json.Serialization.Metadata.JsonTypeInfo<T> type) where T : class, new() =>
        JsonSerializer.Serialize(value ?? new T(), type) == JsonSerializer.Serialize(baseline ?? new T(), type);
}
