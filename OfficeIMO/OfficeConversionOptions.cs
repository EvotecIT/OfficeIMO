using OfficeIMO.Internal;
using System.Text.Json.Serialization;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Email;
using OfficeIMO.Reader.Visio;
using OfficeIMO.Epub;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Excel;
using OfficeIMO.Pdf.PowerPoint;
using OfficeIMO.Pdf.Word;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO;

/// <summary>Conversion policy and settings owned by the selected format adapters.</summary>
public sealed partial class OfficeConversionOptions {
    /// <summary>Rejects reported fidelity loss before writing the destination.</summary>
    public bool FailOnLoss { get; set; }
    /// <summary>Applies verified lossless compression to completed PDF output before publication.</summary>
    public bool CompressPdfOutput { get; set; }

    /// <summary>Controls atomic file publication. Existing output is rejected by default.</summary>
    public OfficePublicationPolicy FileConflictPolicy { get; set; } = OfficePublicationPolicy.Fail;
    /// <summary>Cross-format output intent; explicit adapter settings take precedence.</summary>
    public OfficeConversionProfile OutputProfile { get; set; } = OfficeConversionProfile.Faithful;
    /// <summary>Maximum complete input size, in bytes.</summary>
    public long MaximumInputBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum serialized output size, in bytes.</summary>
    public long MaximumOutputBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum XML characters in an Open XML package part.</summary>
    public long MaximumXmlCharactersInPart { get; set; } = 10L * 1024 * 1024;
    /// <summary>Optional canonical route identifier when file extensions are ambiguous.</summary>
    public string? RouteId { get; set; }
    /// <summary>Physical source filename for stream conversion; its extension distinguishes native format variants.</summary>
    public string? SourceName { get; set; }
    /// <summary>Physical target filename for stream conversion; its extension distinguishes native format variants.</summary>
    public string? TargetName { get; set; }
    /// <summary>Runtime source password. Saved jobs never contain password values.</summary>
    [JsonIgnore]
    public string? SourcePassword { get; set; }
    /// <summary>Native legacy/Open XML compatibility strategy.</summary>
    public OfficeCompatibilityMode CompatibilityMode { get; set; } = OfficeCompatibilityMode.StrictNative;
    /// <summary>Retains original source bytes in an inert carrier when the chosen native conversion policy permits it.</summary>
    public bool EmbedSourceWhenLossy { get; set; }
    /// <summary>Maximum columns retained by Excel visual fallback.</summary>
    public int VisualFallbackMaxColumns { get; set; } = 128;
    /// <summary>Maximum rows retained by Excel visual fallback.</summary>
    public int VisualFallbackMaxRows { get; set; } = 1024;
    /// <summary>Word-to-PDF rendering settings.</summary>
    public WordPdfExportOptions? Word { get; set; }
    /// <summary>Excel-to-PDF rendering settings.</summary>
    public ExcelPdfExportOptions? Excel { get; set; }
    /// <summary>PowerPoint-to-PDF rendering settings.</summary>
    public PowerPointPdfExportOptions? PowerPoint { get; set; }
    /// <summary>Markdown-to-PDF rendering settings.</summary>
    public MarkdownPdfExportOptions? Markdown { get; set; }
    /// <summary>HTML source decoding, base URI and parsing limits.</summary>
    public HtmlConversionDocumentOptions? HtmlLoad { get; set; }
    /// <summary>Source-neutral Reader extraction limits and policies.</summary>
    public ReaderOptions? Reader { get; set; }
    /// <summary>Email parsing, attachments and concealed-text policy.</summary>
    public ReaderEmailOptions? ReaderEmail { get; set; }
    /// <summary>EPUB parsing and resource policy.</summary>
    public EpubReadOptions? ReaderEpub { get; set; }
    /// <summary>Visio semantic fallback and preview policies.</summary>
    public ReaderVisioOptions? ReaderVisio { get; set; }
    /// <summary>Reader-to-PDF pagination, resource and attachment policy.</summary>
    public PdfProjectionOptions? ReaderPdf { get; set; }
    /// <summary>HTML-to-PDF rendering and resource settings.</summary>
    public HtmlToPdfOptions? Html { get; set; }
    /// <summary>PDF-to-HTML semantic or positioned review settings.</summary>
    public PdfToHtmlOptions? PdfHtml { get; set; }
    /// <summary>Literal text-to-PDF settings.</summary>
    public PdfTextImportOptions? PlainText { get; set; }
    /// <summary>PDF source parsing and page selection.</summary>
    public PdfLoadOptions? PdfLoad { get; set; }
    /// <summary>PDF-to-Word reconstruction settings.</summary>
    public WordPdfImportOptions? PdfWord { get; set; }
    /// <summary>PDF-to-Excel reconstruction settings.</summary>
    public ExcelPdfImportOptions? PdfExcel { get; set; }
    /// <summary>PDF-to-PowerPoint reconstruction settings.</summary>
    public PowerPointPdfImportOptions? PdfPowerPoint { get; set; }
    /// <summary>Word source access, native decoding and package-security settings.</summary>
    public WordLoadOptions? WordLoad { get; set; }
    /// <summary>Excel source access, native decoding and package-security settings.</summary>
    public ExcelLoadOptions? ExcelLoad { get; set; }
    /// <summary>PowerPoint source access, native decoding and package-security settings.</summary>
    public PowerPointLoadOptions? PowerPointLoad { get; set; }
    /// <summary>Word destination persistence settings.</summary>
    public WordSaveOptions? WordSave { get; set; }
    /// <summary>Excel destination persistence settings.</summary>
    public ExcelSaveOptions? ExcelSave { get; set; }
    /// <summary>PowerPoint destination persistence settings.</summary>
    public PowerPointSaveOptions? PowerPointSave { get; set; }
    /// <summary>Word-to-Markdown settings.</summary>
    public WordMarkdownExportOptions? WordMarkdown { get; set; }
    /// <summary>Markdown-to-Word settings.</summary>
    public WordMarkdownImportOptions? MarkdownWord { get; set; }
    /// <summary>Word-to-HTML settings.</summary>
    public WordHtmlExportOptions? WordHtml { get; set; }
    /// <summary>HTML-to-Word settings.</summary>
    public WordHtmlImportOptions? HtmlWord { get; set; }
    /// <summary>Markdown-to-HTML settings.</summary>
    public MarkdownHtmlExportOptions? MarkdownHtml { get; set; }
    /// <summary>HTML-to-Markdown settings.</summary>
    public MarkdownHtmlImportOptions? HtmlMarkdown { get; set; }
    /// <summary>Adapter-owned settings for a registered route, independently captured before execution.</summary>
    [JsonIgnore]
    public IOfficeConversionSettings? RegisteredSettings { get; set; }
    /// <summary>Host ownership policy checked immediately before staging and publication.</summary>
    [JsonIgnore]
    public IOfficePublicationGuard? PublicationGuard { get; set; }

    internal OfficeDocumentReader? ReaderPipeline { get; set; }

    internal IReadOnlyDictionary<string, byte[]>? ResourceSnapshots { get; set; }

    internal string? ResourceDirectory { get; set; }

    /// <summary>Creates an independent copy of the format settings and conversion policies.</summary>
    public OfficeConversionOptions Clone() => Snapshot();

    internal OfficeConversionOptions Snapshot() {
        if (MaximumInputBytes < 1 || MaximumOutputBytes < 1 || MaximumXmlCharactersInPart < 1)
            throw new ArgumentOutOfRangeException(nameof(MaximumInputBytes), "Conversion byte and XML limits must be positive.");
        if (!OfficeRuntimeCompatibility.IsDefined(FileConflictPolicy) || !OfficeRuntimeCompatibility.IsDefined(CompatibilityMode) || !OfficeRuntimeCompatibility.IsDefined(OutputProfile))
            throw new ArgumentException("Choose a supported publication and compatibility policy.");
        var copy = (OfficeConversionOptions)MemberwiseClone();
        copy.Word = Word?.Clone(); copy.Excel = Excel?.Clone(); copy.PowerPoint = PowerPoint?.Clone();
        copy.HtmlLoad = HtmlLoad?.Clone();
        copy.Reader = Reader == null ? null : DocumentReaderEngine.NormalizeOptions(Reader);
        copy.ReaderEmail = ReaderEmail == null ? null : EmailArtifactReaderAdapter.Clone(ReaderEmail);
        copy.ReaderEpub = OfficeIMO.Reader.Epub.OfficeDocumentReaderBuilderEpubExtensions.Clone(ReaderEpub);
        copy.ReaderVisio = ReaderVisioOptionsCloner.CloneNullable(ReaderVisio);
        copy.ReaderPipeline = ReaderPipeline;
        copy.ReaderPdf = ReaderPdf?.Snapshot();
        copy.Html = Html?.ClonePdf(); copy.PdfHtml = PdfHtml?.CloneForConversion(); copy.PlainText = PlainText?.Clone();
        copy.Markdown = Markdown?.Clone(); copy.PdfLoad = PdfLoad?.Clone();
        copy.PdfWord = PdfWord?.Clone(); copy.PdfExcel = PdfExcel?.Clone(); copy.PdfPowerPoint = PdfPowerPoint?.Snapshot();
        if (copy.PdfExcel != null) copy.PdfExcel.ReadOptions = PdfExcel?.ReadOptions?.Clone();
        copy.WordLoad = WordLoad?.Snapshot(); copy.ExcelLoad = ExcelLoad?.Snapshot(); copy.PowerPointLoad = PowerPointLoad?.Snapshot();
        copy.WordSave = WordSave?.Clone(); copy.ExcelSave = ExcelSave?.Clone(); copy.PowerPointSave = PowerPointSave?.Clone();
        copy.WordMarkdown = WordMarkdown == null ? null : WordMarkdownConverterExtensions.CopyOptions(WordMarkdown);
        copy.MarkdownWord = MarkdownWord == null ? null : WordMarkdownConverterExtensions.CopyOptions(MarkdownWord);
        copy.WordHtml = WordHtml?.CloneForConversion(); copy.HtmlWord = HtmlWord?.Clone();
        copy.MarkdownHtml = MarkdownHtml?.CloneForRender(); copy.HtmlMarkdown = HtmlMarkdown?.Clone();
        copy.RegisteredSettings = RegisteredSettings?.Snapshot();
        return copy;
    }
}
