using System.Diagnostics;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.MarkdownRenderer;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Web.Converter.Services;

public sealed partial class BrowserConversionService {
    public const long MaxPackageBytes = 25L * 1024L * 1024L;
    public const int MaxTextInputChars = 500_000;
    internal const long MaxFullWorksheetCells = 50_000L;
    internal const int MaxPreviewRowsPerSheet = 250;
    internal const int MaxPackagePartCount = 5_000;
    internal const long MaxPartUncompressedBytes = 32L * 1024L * 1024L;
    internal const long MaxTotalUncompressedBytes = 128L * 1024L * 1024L;
    internal const double MaxCompressionRatio = 200D;

    public ConversionResult ConvertFile(
        ConversionRoute route,
        SelectedDocument file,
        bool limitExcelRows,
        BrowserPdfProfile? profile = null,
        bool generateDebugOverlay = false) {
        ArgumentNullException.ThrowIfNull(route);
        ArgumentNullException.ThrowIfNull(file);
        if (string.Equals(route.Source, "PDF", StringComparison.OrdinalIgnoreCase)) {
            return ConvertPdfFile(route, file);
        }
        BrowserPdfProfile effectiveProfile = profile ?? BrowserPdfProfileCatalog.Faithful;
        var stopwatch = Stopwatch.StartNew();
        PdfConversionPayload payload = route.Id switch {
            "docx-pdf" => ConvertWordToPdf(file, effectiveProfile),
            "xlsx-pdf" => ConvertExcelToPdf(file, limitExcelRows, effectiveProfile),
            "pptx-pdf" => ConvertPowerPointToPdf(file, effectiveProfile),
            _ => throw new NotSupportedException($"The route '{route.Id}' does not accept a document upload.")
        };
        stopwatch.Stop();
        return PdfResult(
            file,
            payload,
            effectiveProfile,
            stopwatch.ElapsedMilliseconds,
            generateDebugOverlay || effectiveProfile.Kind == BrowserPdfProfileKind.Diagnostic);
    }

    public ConversionResult ConvertText(
        ConversionRoute route,
        string input,
        BrowserPdfProfile? profile = null,
        bool generateDebugOverlay = false) {
        ArgumentNullException.ThrowIfNull(input);
        if (input.Length > MaxTextInputChars) {
            throw new ArgumentOutOfRangeException(nameof(input),
                $"Text input is limited to {MaxTextInputChars:N0} characters in the browser converter.");
        }

        return route.Id switch {
            "html-pdf" => ConvertHtmlToPdf(input, profile ?? BrowserPdfProfileCatalog.Faithful, generateDebugOverlay),
            "markdown-html" => ConvertMarkdownToHtml(input),
            "html-markdown" => ConvertHtmlToMarkdown(input),
            "markdown-docx" => ConvertMarkdownToWord(input),
            _ => throw new NotSupportedException($"The route '{route.Id}' does not accept text input.")
        };
    }

    private static ConversionResult ConvertHtmlToPdf(
        string input,
        BrowserPdfProfile profile,
        bool generateDebugOverlay) {
        byte[] sourceBytes = Encoding.UTF8.GetBytes(input);
        var source = new SelectedDocument("officeimo-html.html", ".html", "HTML", sourceBytes.LongLength, sourceBytes);
        var stopwatch = Stopwatch.StartNew();
        PdfDocumentConversionResult conversion = HtmlConversionDocument.Parse(input)
            .ToPdfDocumentResult(BrowserPortablePdfProfile.CreateHtmlOptions(profile));
        (byte[] bytes, PdfSerializationReport serialization) = SaveWithEvidence(conversion);
        stopwatch.Stop();
        return PdfResult(
            source,
            new PdfConversionPayload(
                bytes,
                conversion.Report,
                serialization,
                "OfficeIMO.Html.Pdf",
                "portable-deterministic"),
            profile,
            stopwatch.ElapsedMilliseconds,
            generateDebugOverlay || profile.Kind == BrowserPdfProfileKind.Diagnostic);
    }

    public BrowserConversionArtifact CreateSupportBundle(
        ConversionResult result,
        bool includeDocumentContent = false) {
        ArgumentNullException.ThrowIfNull(result);
        SelectedDocument source = result.SourceSnapshot
            ?? throw new InvalidOperationException("The conversion result does not retain a source snapshot.");
        return BrowserPdfSupportBundle.Create(source, result, includeDocumentContent);
    }

    private static PdfConversionPayload ConvertWordToPdf(SelectedDocument file, BrowserPdfProfile profile) {
        using var stream = new MemoryStream(file.Bytes, writable: false);
        using WordDocument document = WordDocument.Load(stream,
            new WordLoadOptions {
                AccessMode = DocumentAccessMode.ReadOnly,
                PackageSecurity = CreateBrowserPackageSecurity()
            });
        var options = new WordPdfSaveOptions {
            Title = Path.GetFileNameWithoutExtension(file.Name),
            IncludePageNumbers = false,
            FontFamily = BrowserPortablePdfProfile.DefaultFontFamily,
            PdfOptions = BrowserPortablePdfProfile.CreateOptions(profile),
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        };
        var conversion = document.ToPdfDocumentResult(options);
        (byte[] bytes, PdfSerializationReport serialization) = SaveWithEvidence(conversion);
        return new PdfConversionPayload(
            bytes,
            conversion.Report,
            serialization,
            "OfficeIMO.Word.Pdf",
            "includePageNumbers=false");
    }

    private static PdfConversionPayload ConvertExcelToPdf(
        SelectedDocument file,
        bool limitRowsPerSheet,
        BrowserPdfProfile profile) {
        using var stream = new MemoryStream(file.Bytes, writable: false);
        using ExcelDocument document = ExcelDocument.Load(stream,
            new ExcelLoadOptions {
                AccessMode = DocumentAccessMode.ReadOnly,
                PackageSecurity = CreateBrowserPackageSecurity()
            });
        ExcelBrowserReadPlan readPlan = CreateExcelBrowserReadPlan(document, limitRowsPerSheet);
        var options = new ExcelPdfSaveOptions {
            MaxRowsPerSheet = readPlan.MaxRowsPerSheet,
            UseBoundedWorksheetRead = readPlan.UsesBoundedRead,
            FontFamily = BrowserPortablePdfProfile.DefaultFontFamily,
            PdfOptions = BrowserPortablePdfProfile.CreateOptions(profile),
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        };
        var conversion = document.ToPdfDocumentResult(options);
        if (readPlan.WasAutomaticallyBounded) {
            conversion.Report.Add(new PdfConversionWarning(
                "OfficeIMO.Web.Converter",
                "BrowserWorksheetPreview",
                file.Name,
                $"The worksheet exceeded the browser-safe full-conversion budget of {MaxFullWorksheetCells:N0} cells. A preview of up to {readPlan.MaxRowsPerSheet:N0} rows per sheet was generated instead; use a server conversion for the complete workbook.",
                PdfConversionWarningSeverity.Warning,
                details: new Dictionary<string, string> {
                    ["mode"] = "automatic-browser-preview",
                    ["maximumFullWorksheetCells"] = MaxFullWorksheetCells.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    ["maximumRowsPerSheet"] = readPlan.MaxRowsPerSheet?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty
                }));
        }
        (byte[] bytes, PdfSerializationReport serialization) = SaveWithEvidence(conversion);
        return new PdfConversionPayload(
            bytes,
            conversion.Report,
            serialization,
            "OfficeIMO.Excel.Pdf",
            readPlan.UsesBoundedRead
                ? $"maxRowsPerSheet={readPlan.MaxRowsPerSheet};mode={(readPlan.WasAutomaticallyBounded ? "automatic-browser-preview" : "requested-browser-preview")}"
                : "maxRowsPerSheet=unlimited");
    }

    internal static ExcelBrowserReadPlan CreateExcelBrowserReadPlan(
        ExcelDocument document,
        bool previewRequested) {
        ArgumentNullException.ThrowIfNull(document);
        int widestWorksheetColumns = 1;
        bool exceedsFullConversionBudget = false;
        using ExcelDocumentReader reader = ExcelDocumentReader.Wrap(document.OpenXmlDocument);
        foreach (string sheetName in reader.GetSheetNames()) {
            string usedRange = reader.GetSheet(sheetName).GetUsedRangeA1();
            if (!A1.TryParseRange(usedRange, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                continue;
            }

            int rows = Math.Max(1, lastRow - firstRow + 1);
            int columns = Math.Max(1, lastColumn - firstColumn + 1);
            widestWorksheetColumns = Math.Max(widestWorksheetColumns, columns);
            if ((long)rows * columns > MaxFullWorksheetCells) {
                exceedsFullConversionBudget = true;
            }
        }

        bool boundedRead = previewRequested || exceedsFullConversionBudget;
        if (!boundedRead) {
            return new ExcelBrowserReadPlan(null, UsesBoundedRead: false, WasAutomaticallyBounded: false);
        }

        int rowsWithinCellBudget = Math.Max(1, (int)(MaxFullWorksheetCells / widestWorksheetColumns));
        int maximumRows = Math.Min(MaxPreviewRowsPerSheet, rowsWithinCellBudget);
        return new ExcelBrowserReadPlan(
            maximumRows,
            UsesBoundedRead: true,
            WasAutomaticallyBounded: !previewRequested && exceedsFullConversionBudget);
    }

    private static PdfConversionPayload ConvertPowerPointToPdf(SelectedDocument file, BrowserPdfProfile profile) {
        using var stream = new MemoryStream(file.Bytes, writable: false);
        using PowerPointPresentation presentation = PowerPointPresentation.Load(stream,
            new PowerPointLoadOptions {
                AccessMode = DocumentAccessMode.ReadOnly,
                PackageSecurity = CreateBrowserPackageSecurity()
            });
        var options = new PowerPointPdfSaveOptions {
            WarnOnPictureAspectRatioDistortion = true,
            FontFamily = BrowserPortablePdfProfile.DefaultFontFamily,
            PdfOptions = BrowserPortablePdfProfile.CreateOptions(profile),
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        };
        var conversion = presentation.ToPdfDocumentResult(options);
        (byte[] bytes, PdfSerializationReport serialization) = SaveWithEvidence(conversion);
        return new PdfConversionPayload(
            bytes,
            conversion.Report,
            serialization,
            "OfficeIMO.PowerPoint.Pdf",
            "warnOnPictureAspectRatioDistortion=true");
    }

    private static ConversionResult PdfResult(
        SelectedDocument file,
        PdfConversionPayload payload,
        BrowserPdfProfile profile,
        long conversionMilliseconds,
        bool generateDebugOverlay) {
        byte[] bytes = payload.Bytes;
        PdfConversionReport report = payload.Report;
        EnsurePdf(bytes);
        string fileName = Path.GetFileNameWithoutExtension(file.Name) + ".pdf";
        BrowserConversionArtifact companionReport = BrowserPdfConversionManifest.Create(
            file,
            fileName,
            bytes,
            report,
            payload.Converter,
            payload.OptionProfile,
            profile,
            conversionMilliseconds,
            payload.Serialization);
        BrowserConversionArtifact? debugOverlay = generateDebugOverlay
            ? CreateDebugOverlay(file, bytes)
            : null;
        return new ConversionResult(
            bytes,
            fileName,
            "application/pdf",
            null,
            null,
            report.Warnings.Select(static warning => warning.ToString()).ToArray()) {
            FidelityStatus = report.FidelityStatus.ToString(),
            ProvenanceSummary = $"{BrowserPortablePdfProfile.FontPackId} · {BrowserPortablePdfProfile.FontPackFingerprint[..12]}",
            CompanionReport = companionReport,
            DebugOverlay = debugOverlay,
            StructuredWarnings = report.Warnings.Select(CreateWarningView).ToArray(),
            PeakRetainedMemoryBytes = AddWithoutOverflow(
                payload.Serialization.PeakRetainedPageContentBytes,
                payload.Serialization.PeakRetainedObjectBytes),
            PageCount = payload.Serialization.PageCount,
            ConversionMilliseconds = conversionMilliseconds,
            Profile = profile,
            SourceSnapshot = Snapshot(file)
        };
    }

    private static SelectedDocument Snapshot(SelectedDocument source) =>
        new(
            source.Name,
            source.Extension,
            source.FormatLabel,
            source.Size,
            (byte[])source.Bytes.Clone());

    private static (byte[] Bytes, PdfSerializationReport Serialization) SaveWithEvidence(
        PdfDocumentConversionResult conversion) {
        using var buffer = new MemoryStream();
        PdfSaveResult save = conversion.Save(buffer);
        PdfSerializationReport serialization = save.Serialization
            ?? throw new InvalidOperationException("The PDF writer did not return serialization evidence.");
        return (buffer.ToArray(), serialization);
    }

    private static BrowserConversionArtifact CreateDebugOverlay(SelectedDocument source, byte[] pdf) {
        OfficeDrawing drawing = PdfDocument.Open(pdf).Read.LayoutDebugOverlay(
            1,
            new PdfLayoutDebugOverlayOptions {
                MaxElements = 12_000,
                MaxRasterPixels = 16_000_000
            });
        byte[] bytes = Encoding.UTF8.GetBytes(OfficeDrawingSvgExporter.ToSvg(drawing));
        return new BrowserConversionArtifact(
            bytes,
            Path.GetFileNameWithoutExtension(source.Name) + ".page-1.layout.svg",
            "image/svg+xml");
    }

    private static ConversionWarningView CreateWarningView(PdfConversionWarning warning) {
        int? pageNumber = TryReadPositiveInt(warning.Details, "pageNumber")
            ?? TryReadPositiveInt(warning.Details, "page")
            ?? TryReadPageFromSource(warning.Source);
        string construct = warning.Details.TryGetValue("construct", out string? declaredConstruct)
            ? declaredConstruct
            : warning.LayoutDiagnostic?.Kind.ToString()
                ?? (warning.Details.TryGetValue("feature", out string? feature)
                    ? feature
                    : warning.Code);
        bool canChangePagination =
            warning.Severity != PdfConversionWarningSeverity.Information &&
            (warning.Code.Contains("font", StringComparison.OrdinalIgnoreCase) ||
             warning.Code.Contains("pagination", StringComparison.OrdinalIgnoreCase) ||
             warning.Code.Contains("overflow", StringComparison.OrdinalIgnoreCase) ||
             warning.LayoutDiagnostic?.Kind is PdfLayoutDiagnosticKind.AdjustedGeometry
                 or PdfLayoutDiagnosticKind.ClippedContent
                 or PdfLayoutDiagnosticKind.Overflow);
        return new ConversionWarningView(
            warning.Code,
            warning.Source,
            warning.Message,
            warning.Severity.ToString(),
            construct,
            pageNumber,
            canChangePagination);
    }

    private static int? TryReadPositiveInt(IReadOnlyDictionary<string, string> values, string key) =>
        values.TryGetValue(key, out string? value) &&
        int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int parsed) &&
        parsed > 0
            ? parsed
            : null;

    private static int? TryReadPageFromSource(string source) {
        if (string.IsNullOrWhiteSpace(source)) {
            return null;
        }

        int marker = source.IndexOf("page ", StringComparison.OrdinalIgnoreCase);
        if (marker < 0) {
            return null;
        }

        marker += 5;
        int end = marker;
        while (end < source.Length && char.IsDigit(source[end])) {
            end++;
        }

        return end > marker &&
            int.TryParse(source[marker..end], System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int page)
                ? page
                : null;
    }

    private static long AddWithoutOverflow(long first, long second) =>
        first > long.MaxValue - second ? long.MaxValue : first + second;

    private static ConversionResult ConvertMarkdownToHtml(string input) {
        var options = MarkdownRendererPresets.CreateStrictMinimalPortable();
        options.MarkdownOverflowHandling = OverflowHandling.Throw;
        string html = global::OfficeIMO.MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(input, options);
        return TextResult(html, "officeimo-markdown.html", "text/html;charset=utf-8", html);
    }

    private static ConversionResult ConvertHtmlToMarkdown(string input) {
        var options = HtmlToMarkdownOptions.CreatePortableProfile();
        string markdown = HtmlConversionDocument.Parse(input).ToMarkdown(options);
        return TextResult(markdown, "officeimo-html.md", "text/markdown;charset=utf-8", null);
    }

    private static ConversionResult ConvertMarkdownToWord(string input) {
        MarkdownDoc markdown = MarkdownReader.Parse(input);
        using WordDocument document = markdown.ToWordDocument(new MarkdownToWordOptions {
            PreferNarrativeSingleLineDefinitions = true
        });
        byte[] bytes = document.ToBytes();
        return new ConversionResult(
            bytes,
            "officeimo-markdown.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            $"DOCX generated: {FormatBytes(bytes.Length)}.",
            null,
            []);
    }

    private static ConversionResult TextResult(string text, string fileName, string contentType, string? htmlPreview) =>
        new(Encoding.UTF8.GetBytes(text), fileName, contentType, text, htmlPreview, []);

    internal static OfficePackageSecurityOptions CreateBrowserPackageSecurity() {
        OfficePackageSecurityOptions options = OfficePackageSecurityOptions.SecureDefaults;
        options.MaxPackageBytes = MaxPackageBytes;
        options.MaxPartCount = MaxPackagePartCount;
        options.MaxPartUncompressedBytes = MaxPartUncompressedBytes;
        options.MaxTotalUncompressedBytes = MaxTotalUncompressedBytes;
        options.MaxCompressionRatio = MaxCompressionRatio;
        return options;
    }

    private static void EnsurePdf(byte[] bytes) {
        if (bytes.Length < 4 || bytes[0] != 0x25 || bytes[1] != 0x50 || bytes[2] != 0x44 || bytes[3] != 0x46) {
            throw new InvalidDataException("The conversion did not return a valid PDF header.");
        }
    }

    private static string FormatBytes(long bytes) {
        string[] units = ["B", "KB", "MB", "GB"];
        double value = bytes;
        int unit = 0;
        while (value >= 1024 && unit < units.Length - 1) {
            value /= 1024;
            unit++;
        }
        return unit == 0 ? $"{bytes} B" : $"{value:0.##} {units[unit]}";
    }

    private sealed record PdfConversionPayload(
        byte[] Bytes,
        PdfConversionReport Report,
        PdfSerializationReport Serialization,
        string Converter,
        string OptionProfile);

    internal readonly record struct ExcelBrowserReadPlan(
        int? MaxRowsPerSheet,
        bool UsesBoundedRead,
        bool WasAutomaticallyBounded);
}
