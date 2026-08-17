using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Drawing.HarfBuzz;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Word.Pdf;

namespace OfficeIMO.Web.Converter.Services;

public sealed partial class BrowserConversionService {
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
            ProvenanceSummary = route.EnginePath + " · logical PDF import",
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
        var options = new PdfPowerPointImportOptions {
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
        PdfHtmlConversionResult conversion = pdf.ToHtmlResult();
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
