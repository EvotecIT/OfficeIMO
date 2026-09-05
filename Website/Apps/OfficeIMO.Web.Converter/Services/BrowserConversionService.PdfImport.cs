using System.Diagnostics;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Drawing.HarfBuzz;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Web.Converter.Services;

public sealed partial class BrowserConversionService {
    private const long MaximumPngArchiveBytes = BrowserPdfPolicy.MaxOutputBytes / 2L;

    private static ConversionResult ConvertPdfFile(
        ConversionRoute route,
        SelectedDocument file,
        PdfPowerPointImportMode pdfPowerPointMode) {
        var stopwatch = Stopwatch.StartNew();
        PdfImportPayload payload = route.Id switch {
            "pdf-docx" => ConvertPdfToWord(file),
            "pdf-xlsx" => ConvertPdfToExcel(file),
            "pdf-pptx" => ConvertPdfToPowerPoint(file, pdfPowerPointMode),
            "pdf-html" => ConvertPdfToHtml(file),
            "pdf-png" => ConvertPdfToPng(file),
            _ => throw new NotSupportedException($"The route '{route.Id}' is not available in the browser workspace.")
        };
        stopwatch.Stop();

        string fileName = Path.GetFileNameWithoutExtension(file.Name) + payload.Extension;
        IReadOnlyList<ConversionWarningView> structuredWarnings = payload.Warnings.Select(CreateWarningView).ToArray();
        string fidelity = payload.HasLoss && payload.FidelityStatus == "Reconstructed"
            ? "Degraded"
            : payload.FidelityStatus;
        BrowserConversionArtifact report = CreatePdfImportReport(
            route,
            file,
            fileName,
            payload,
            fidelity,
            stopwatch.ElapsedMilliseconds);
        return new ConversionResult(
            payload.Bytes,
            fileName,
            payload.ContentType,
            payload.Text,
            payload.HtmlPreview,
            payload.Warnings.Select(static warning => warning.ToString()).ToArray()) {
            FidelityStatus = fidelity,
            ProvenanceSummary = route.EnginePath + " · " + DescribePdfProjection(payload.Projection),
            CompanionReport = report,
            StructuredWarnings = structuredWarnings,
            PageCount = payload.PageCount,
            ConversionMilliseconds = stopwatch.ElapsedMilliseconds,
            SourceSnapshot = Snapshot(file)
        };
    }

    private static PdfImportPayload ConvertPdfToWord(SelectedDocument file) {
        PdfDocument pdf = BrowserPdfPolicy.Open(file);
        PdfWordConversionResult conversion = pdf.ToWordDocumentResult();
        using var document = conversion.Value;
        byte[] bytes = document.ToBytes();
        return new PdfImportPayload(
            bytes,
            ".docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            conversion.Report.Warnings,
            conversion.HasLoss,
            pdf.Inspect().PageCount,
            $"Editable DOCX generated: {FormatBytes(bytes.Length)}.",
            null,
            "logical-content",
            conversion.HasLoss ? "Degraded" : "Reconstructed");
    }

    private static PdfImportPayload ConvertPdfToExcel(SelectedDocument file) {
        PdfDocument pdf = BrowserPdfPolicy.Open(file);
        PdfExcelTableImportResult conversion = pdf.ImportTablesToExcelDocumentResult();
        using var document = conversion.Value;
        byte[] bytes = document.ToBytes();
        var warnings = new List<PdfConversionWarning>();
        if (conversion.HasOmittedPageContent) {
            warnings.Add(new PdfConversionWarning(
                "OfficeIMO.Excel.Pdf",
                "PdfTablesOnly",
                "document",
                "Excel import includes detected tables; other page content remains outside this table-only route."));
        }
        if (conversion.HasLoss) {
            warnings.Add(new PdfConversionWarning(
                "OfficeIMO.Excel.Pdf",
                "PdfTableRowsTruncated",
                "document",
                "One or more detected tables exceeded the configured import limit."));
        }
        return new PdfImportPayload(
            bytes,
            ".xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            warnings,
            conversion.HasLoss || conversion.HasOmittedPageContent,
            pdf.Inspect().PageCount,
            $"Editable XLSX generated with {conversion.Report.Entries.Count} detected table{(conversion.Report.Entries.Count == 1 ? string.Empty : "s")}.",
            null,
            "detected-tables",
            "Partial");
    }

    private static PdfImportPayload ConvertPdfToPowerPoint(
        SelectedDocument file,
        PdfPowerPointImportMode mode) {
        PdfDocument pdf = BrowserPdfPolicy.Open(file);
        BrowserPowerPointImportProfile profile = BrowserPowerPointImportProfileCatalog.Find(mode);
        var options = new PdfToPowerPointOptions {
            Mode = profile.Mode,
            RenderFonts = BrowserPortablePdfProfile.CreateDrawingFonts(),
            TextShapingProvider = OfficeHarfBuzzTextShapingProvider.Instance
        };
        PdfPowerPointConversionResult conversion = pdf.ToPowerPointPresentationResult(options);
        using var presentation = conversion.Value;
        byte[] bytes = presentation.ToBytes();
        string summary = profile.Mode switch {
            PdfPowerPointImportMode.VisualPages =>
                $"Visual PPTX generated: {FormatBytes(bytes.Length)}. Each slide contains one page image; its text, shapes, charts, and tables are not editable.",
            PdfPowerPointImportMode.HybridVisualAndEditableTables =>
                $"Hybrid PPTX generated: {FormatBytes(bytes.Length)}. Page images preserve appearance and {conversion.Report.TableEntries.Count} detected table segment{(conversion.Report.TableEntries.Count == 1 ? string.Empty : "s")} remain editable.",
            PdfPowerPointImportMode.EditableTables =>
                $"Tables-only PPTX generated with {conversion.Report.TableEntries.Count} editable table segment{(conversion.Report.TableEntries.Count == 1 ? string.Empty : "s")}.",
            _ =>
                $"Editable-content PPTX generated with {conversion.Report.EditablePages.Sum(static page => page.TextBoxCount)} text box{(conversion.Report.EditablePages.Sum(static page => page.TextBoxCount) == 1 ? string.Empty : "es")}, {conversion.Report.TableEntries.Count} table segment{(conversion.Report.TableEntries.Count == 1 ? string.Empty : "s")}, {conversion.Report.EditablePages.Sum(static page => page.ShapeCount)} shape{(conversion.Report.EditablePages.Sum(static page => page.ShapeCount) == 1 ? string.Empty : "s")}, and {conversion.Report.EditablePages.Sum(static page => page.ImageCount)} separate image{(conversion.Report.EditablePages.Sum(static page => page.ImageCount) == 1 ? string.Empty : "s")}.",
        };
        return new PdfImportPayload(
            bytes,
            ".pptx",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            conversion.Warnings,
            conversion.HasLoss || conversion.HasOmittedPageContent,
            pdf.Inspect().PageCount,
            summary,
            null,
            profile.Projection,
            profile.FidelityStatus);
    }

    private static PdfImportPayload ConvertPdfToHtml(SelectedDocument file) {
        PdfDocument pdf = BrowserPdfPolicy.Open(file);
        PdfHtmlConversionResult conversion = pdf.ToHtmlResult(new PdfToHtmlOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true,
            IncludeFormWidgets = true
        });
        byte[] bytes = Encoding.UTF8.GetBytes(conversion.Value);
        return new PdfImportPayload(
            bytes,
            ".html",
            "text/html;charset=utf-8",
            conversion.Report.Warnings,
            conversion.HasLoss,
            conversion.Summary.SourcePageCount,
            conversion.Value,
            conversion.Value,
            conversion.Summary.ProfileId,
            conversion.HasLoss ? "Degraded" : "Reconstructed");
    }

    private static PdfImportPayload ConvertPdfToPng(SelectedDocument file) {
        PdfDocument pdf = BrowserPdfPolicy.Open(file);
        int pageCount = pdf.Inspect().PageCount;
        if (pageCount == 0) {
            throw new InvalidOperationException("The PDF did not contain a renderable page.");
        }
        var options = new PdfImageExportOptions {
            TargetDpi = 144D,
            Fonts = BrowserPortablePdfProfile.CreateDrawingFonts(),
            TextShapingProvider = OfficeHarfBuzzTextShapingProvider.Instance,
            MaximumOutputCount = BrowserPdfPolicy.MaxPages,
            MaximumTotalRasterPixels = 250_000_000L,
            MaximumTotalEncodedBytes = pageCount == 1
                ? BrowserPdfPolicy.MaxOutputBytes
                : MaximumPngArchiveBytes,
            RenderTimeout = TimeSpan.FromSeconds(30),
            MaximumDegreeOfParallelism = 1
        };
        if (pageCount == 1) {
            OfficeImageExportResult image = pdf.ToImages(options).AsPng().Export().Single();
            IReadOnlyList<PdfConversionWarning> warnings = MapImageDiagnostics(image, 1);
            return new PdfImportPayload(
                image.Bytes,
                ".png",
                image.MimeType,
                warnings,
                HasMaterialImageLoss(warnings),
                1,
                $"Detailed 144 DPI PNG generated at {image.Width} × {image.Height} pixels.",
                null,
                "visual-page-images",
                HasMaterialImageLoss(warnings) ? "Degraded" : "Visual");
        }

        (byte[] Archive, IReadOnlyList<PdfConversionWarning> Warnings) archive =
            CreatePngArchive(file.Name, pageCount, pdf, options);
        bool hasLoss = HasMaterialImageLoss(archive.Warnings);
        return new PdfImportPayload(
            archive.Archive,
            ".zip",
            "application/zip",
            archive.Warnings,
            hasLoss,
            pageCount,
            $"{pageCount} detailed 144 DPI page PNGs generated in one ZIP archive.",
            null,
            "visual-page-images",
            hasLoss ? "Degraded" : "Visual");
    }

    private static (byte[] Archive, IReadOnlyList<PdfConversionWarning> Warnings) CreatePngArchive(
        string sourceFileName,
        int pageCount,
        PdfDocument pdf,
        PdfImageExportOptions options) {
        string baseName = Path.GetFileNameWithoutExtension(sourceFileName);
        int digits = Math.Max(3, pageCount.ToString(System.Globalization.CultureInfo.InvariantCulture).Length);
        int nextPageNumber = 1;
        var warnings = new List<PdfConversionWarning>();
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            pdf.ToImages(options).AsPng().ExportEach(image => {
                int pageNumberValue = nextPageNumber++;
                string pageNumber = pageNumberValue.ToString(
                    new string('0', digits),
                    System.Globalization.CultureInfo.InvariantCulture);
                ZipArchiveEntry entry = archive.CreateEntry(
                    $"{baseName}.page-{pageNumber}.png",
                    CompressionLevel.NoCompression);
                entry.LastWriteTime = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
                using Stream entryStream = entry.Open();
                byte[] bytes = image.Bytes;
                entryStream.Write(bytes, 0, bytes.Length);
                warnings.AddRange(MapImageDiagnostics(image, pageNumberValue));
                if (output.Length > MaximumPngArchiveBytes) {
                    throw new InvalidDataException(
                        $"The PNG archive exceeds the browser output limit of {FormatBytes(MaximumPngArchiveBytes)}.");
                }
            });
        }
        if (output.Length > MaximumPngArchiveBytes) {
            throw new InvalidDataException(
                $"The PNG archive exceeds the browser output limit of {FormatBytes(MaximumPngArchiveBytes)}.");
        }
        return (output.ToArray(), warnings.AsReadOnly());
    }

    private static IReadOnlyList<PdfConversionWarning> MapImageDiagnostics(
        OfficeImageExportResult image,
        int pageNumber) =>
        image.Diagnostics.Select(diagnostic => new PdfConversionWarning(
            "OfficeIMO.Pdf",
            diagnostic.Code,
            $"page {pageNumber}",
            diagnostic.Message,
            diagnostic.Severity switch {
                OfficeImageExportDiagnosticSeverity.Error => PdfConversionWarningSeverity.Error,
                OfficeImageExportDiagnosticSeverity.Warning => PdfConversionWarningSeverity.Warning,
                _ => PdfConversionWarningSeverity.Information
            })).ToArray();

    private static bool HasMaterialImageLoss(IReadOnlyList<PdfConversionWarning> warnings) =>
        warnings.Any(static warning => warning.Severity != PdfConversionWarningSeverity.Information);

    private static string DescribePdfProjection(string projection) => projection switch {
        "pdf-html-positioned-review" => "positioned PDF review reconstruction",
        "visual-page-images" => "visual PDF page rendering",
        "visual-page-slides" => "visual PDF page slides",
        "hybrid-visual-table-slides" => "hybrid PDF page and table reconstruction",
        "editable-table-slides" or "detected-tables" => "detected PDF table reconstruction",
        _ => "semantic PDF reconstruction"
    };

    private static BrowserConversionArtifact CreatePdfImportReport(
        ConversionRoute route,
        SelectedDocument source,
        string outputFileName,
        PdfImportPayload payload,
        string fidelity,
        long conversionMilliseconds) {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true })) {
            writer.WriteStartObject();
            writer.WriteNumber("schemaVersion", 1);
            writer.WriteString("route", route.Id);
            writer.WriteString("engine", route.EnginePath);
            writer.WriteBoolean("browserLocal", true);
            writer.WriteString("projection", payload.Projection);
            writer.WriteString("fidelityStatus", fidelity);
            writer.WriteNumber("conversionMilliseconds", conversionMilliseconds);
            writer.WriteStartObject("source");
            writer.WriteString("fileName", source.Name);
            writer.WriteNumber("bytes", source.Bytes.LongLength);
            writer.WriteString("sha256", Convert.ToHexString(SHA256.HashData(source.Bytes)).ToLowerInvariant());
            writer.WriteNumber("pageCount", payload.PageCount);
            writer.WriteEndObject();
            writer.WriteStartObject("output");
            writer.WriteString("fileName", outputFileName);
            writer.WriteNumber("bytes", payload.Bytes.LongLength);
            writer.WriteString("sha256", Convert.ToHexString(SHA256.HashData(payload.Bytes)).ToLowerInvariant());
            writer.WriteEndObject();
            writer.WriteStartArray("warnings");
            foreach (PdfConversionWarning warning in payload.Warnings) {
                writer.WriteStartObject();
                writer.WriteString("code", warning.Code);
                writer.WriteString("source", warning.Source);
                writer.WriteString("severity", warning.Severity.ToString());
                writer.WriteString("message", warning.Message);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
        }
        return new BrowserConversionArtifact(
            stream.ToArray(),
            Path.GetFileNameWithoutExtension(outputFileName) + ".officeimo-report.json",
            "application/json");
    }

    private sealed record PdfImportPayload(
        byte[] Bytes,
        string Extension,
        string ContentType,
        IReadOnlyList<PdfConversionWarning> Warnings,
        bool HasLoss,
        int PageCount,
        string? Text,
        string? HtmlPreview,
        string Projection,
        string FidelityStatus);
}
