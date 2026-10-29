using System.Text.Json.Serialization;
using OfficeIMO.Pdf;
using OfficeIMO.Html.Pdf;

namespace OfficeIMO;

internal sealed record OfficeBatchRenderingConfiguration(string Route, OfficeConversionProfile OutputProfile,
    OfficeBatchRendererSettings Options, string Host, long MaximumXmlCharactersInPart);

[JsonSourceGenerationOptions(IncludeFields = true, GenerationMode = JsonSourceGenerationMode.Metadata)]
[JsonSerializable(typeof(OfficeBatchRenderingConfiguration))]
internal partial class OfficeBatchConfigurationJsonContext : JsonSerializerContext { }

[JsonSourceGenerationOptions(IncludeFields = true, GenerationMode = JsonSourceGenerationMode.Metadata)]
[JsonSerializable(typeof(PdfCreateOptions), TypeInfoPropertyName = "PdfOptions")]
internal partial class OfficeBatchPdfSettingsJsonContext : JsonSerializerContext { }


// Checkpoints capture the established scalar/native-renderer contract. New callback-rich
// import routes remain available in ordinary batches until their provider identity is durable.
internal sealed record OfficeBatchRendererSettings {
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public OfficeDocumentFormat? SourceFormat { get; init; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? SourceExtension { get; init; }
    public OfficeIMO.Word.WordLoadOptions? WordLoad { get; init; }
    public OfficeIMO.Excel.ExcelLoadOptions? ExcelLoad { get; init; }
    public OfficeIMO.PowerPoint.PowerPointLoadOptions? PowerPointLoad { get; init; }
    public OfficeIMO.Word.WordSaveOptions? WordSave { get; init; }
    public OfficeIMO.Excel.ExcelSaveOptions? ExcelSave { get; init; }
    public OfficeIMO.PowerPoint.PowerPointSaveOptions? PowerPointSave { get; init; }
    public bool FailOnLoss { get; init; }
    public bool CompressPdfOutput { get; init; }
    public OfficeCompatibilityMode CompatibilityMode { get; init; }
    public bool EmbedSourceWhenLossy { get; init; }
    public int VisualFallbackMaxColumns { get; init; }
    public int VisualFallbackMaxRows { get; init; }
    public OfficeIMO.Word.Pdf.WordPdfExportOptions? Word { get; init; }
    public OfficeIMO.Excel.Pdf.ExcelPdfExportOptions? Excel { get; init; }
    public OfficeIMO.PowerPoint.Pdf.PowerPointPdfExportOptions? PowerPoint { get; init; }
    public OfficeIMO.Html.Pdf.HtmlToPdfOptions? Html { get; init; }
    public OfficeIMO.Markdown.Pdf.MarkdownPdfExportOptions? Markdown { get; init; }
    public OfficeIMO.Pdf.Word.WordPdfImportOptions? PdfWord { get; init; }
    public OfficeIMO.Pdf.PowerPoint.PowerPointPdfImportOptions? PdfPowerPoint { get; init; }
    public PdfLoadOptions? PdfLoad { get; init; }
    public PdfToHtmlOptions? PdfHtml { get; init; }
    public PdfTextImportOptions? PlainText { get; init; }
    internal static OfficeBatchRendererSettings Capture(OfficeConversionOptions options) => new() {
        SourceFormat = options.SourceFormat, SourceExtension = options.SourceExtension,
        WordLoad = options.WordLoad, ExcelLoad = options.ExcelLoad, PowerPointLoad = options.PowerPointLoad,
        WordSave = options.WordSave, ExcelSave = options.ExcelSave, PowerPointSave = options.PowerPointSave,
        FailOnLoss = options.FailOnLoss, CompressPdfOutput = options.CompressPdfOutput, CompatibilityMode = options.CompatibilityMode,
        EmbedSourceWhenLossy = options.EmbedSourceWhenLossy,
        VisualFallbackMaxColumns = options.VisualFallbackMaxColumns, VisualFallbackMaxRows = options.VisualFallbackMaxRows,
        Word = options.Word, Excel = options.Excel, PowerPoint = options.PowerPoint, Html = options.Html,
        Markdown = options.Markdown, PdfWord = options.PdfWord, PdfPowerPoint = options.PdfPowerPoint,
        PdfLoad = options.PdfLoad, PdfHtml = options.PdfHtml, PlainText = options.PlainText
    };
}
