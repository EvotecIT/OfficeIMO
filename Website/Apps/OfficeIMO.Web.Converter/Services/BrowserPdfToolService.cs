using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Models;

namespace OfficeIMO.Web.Converter.Services;

internal sealed partial class BrowserPdfToolService {
    internal const int MaxPdfFiles = 10;
    internal const long MaxAggregatePdfBytes = 75L * 1024L * 1024L;
    internal const int MaxComparisonPages = 25;

    internal PdfToolResult Execute(PdfToolRequest request) {
        ArgumentNullException.ThrowIfNull(request);
        ValidateRequest(request);

        PdfToolExecution execution = request.Tool.Kind switch {
            PdfToolKind.Inspect => Inspect(request),
            PdfToolKind.Merge => Merge(request),
            PdfToolKind.Split => Split(request),
            PdfToolKind.Extract => TransformPages(request, "extracted", static (document, selector, _) => document.Pages.Extract(selector)),
            PdfToolKind.Delete => TransformPages(request, "pages-removed", static (document, selector, _) => document.Pages.Delete(selector)),
            PdfToolKind.Reorder => TransformPages(request, "reordered", static (document, selector, _) => document.Pages.Reorder(selector)),
            PdfToolKind.Rotate => TransformPages(request, "rotated", static (document, selector, operation) => document.Pages.Rotate(operation.RotationDegrees, selector)),
            PdfToolKind.Optimize => Optimize(request),
            PdfToolKind.Protect => Protect(request),
            PdfToolKind.Unlock => Unlock(request),
            PdfToolKind.Redact => Redact(request),
            PdfToolKind.Compare => Compare(request),
            _ => throw new ArgumentOutOfRangeException(nameof(request), request.Tool.Kind, "Unsupported browser PDF tool.")
        };

        BrowserConversionArtifact? report = execution.ReportDetails is null
            ? null
            : CreateReport(request, execution);
        return new PdfToolResult(
            execution.Artifact,
            report,
            execution.Messages,
            execution.Summary,
            execution.PageCount,
            request.Files.Sum(static file => file.Size),
            execution.PreviewInBrowser);
    }

    private static void ValidateRequest(PdfToolRequest request) {
        int count = request.Files.Count;
        int minimum = request.Tool.InputMode switch {
            PdfToolInputMode.Single => 1,
            PdfToolInputMode.Pair => 2,
            PdfToolInputMode.Multiple => 2,
            _ => throw new ArgumentOutOfRangeException(nameof(request), request.Tool.InputMode, "Unsupported PDF input mode.")
        };
        int maximum = request.Tool.InputMode == PdfToolInputMode.Multiple ? MaxPdfFiles : minimum;
        if (count < minimum || count > maximum) {
            string expected = minimum == maximum ? minimum.ToString(System.Globalization.CultureInfo.InvariantCulture) : $"{minimum} to {maximum}";
            throw new ArgumentException($"{request.Tool.Label} requires {expected} PDF files.", nameof(request));
        }

        long aggregateBytes = 0;
        for (int index = 0; index < request.Files.Count; index++) {
            SelectedDocument file = request.Files[index];
            if (!string.Equals(file.Extension, ".pdf", StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException($"{file.Name} is not a PDF file.");
            }
            EnsurePdf(file.Bytes, file.Name);
            aggregateBytes = checked(aggregateBytes + file.Bytes.LongLength);
        }
        if (aggregateBytes > MaxAggregatePdfBytes) {
            throw new InvalidDataException($"Selected PDFs total {FormatBytes(aggregateBytes)}; the browser workbench limit is {FormatBytes(MaxAggregatePdfBytes)}.");
        }

        if (request.Tool.RequiresPageSelection && string.IsNullOrWhiteSpace(request.PageSelection)) {
            throw new ArgumentException("Enter a page selection such as 1-3,5,last.", nameof(request));
        }
        if (request.Tool.RequiresPagesPerDocument && request.PagesPerDocument <= 0) {
            throw new ArgumentOutOfRangeException(nameof(request), "Pages per document must be greater than zero.");
        }
        if (request.Tool.RequiresRotation && request.RotationDegrees is not (90 or 180 or 270)) {
            throw new ArgumentOutOfRangeException(nameof(request), "Rotation must be 90, 180, or 270 degrees.");
        }
        if (request.Tool.RequiresUserPassword && string.IsNullOrWhiteSpace(request.UserPassword)) {
            throw new ArgumentException("Enter a document-open password.", nameof(request));
        }
        if (request.Tool.RequiresOwnerPassword && string.IsNullOrWhiteSpace(request.OwnerPassword)) {
            throw new ArgumentException("Enter an owner password.", nameof(request));
        }
        if (request.Tool.RequiresRedactionText && string.IsNullOrWhiteSpace(request.RedactionText)) {
            throw new ArgumentException("Enter literal text to redact.", nameof(request));
        }
        if (request.Tool.RequiresDestructiveConfirmation && !request.DestructiveActionConfirmed) {
            throw new InvalidOperationException("Confirm that this operation permanently changes the downloaded copy.");
        }
    }

    private static PdfDocument Open(SelectedDocument file, string? password = null) =>
        PdfDocument.Open(
            file.Bytes,
            password is null ? null : new PdfReadOptions { Password = password });

    private static PdfPageSelector Selector(PdfToolRequest request) => PdfPageSelector.Parse(request.PageSelection);

    private static string OutputName(SelectedDocument file, string suffix, string extension = ".pdf") =>
        Path.GetFileNameWithoutExtension(file.Name) + "." + suffix + extension;

    private static BrowserConversionArtifact PdfArtifact(byte[] bytes, string fileName) {
        EnsurePdf(bytes, fileName);
        return new BrowserConversionArtifact(bytes, fileName, "application/pdf");
    }

    private static void EnsurePdf(byte[] bytes, string name) {
        if (bytes.Length < 5 || bytes[0] != 0x25 || bytes[1] != 0x50 || bytes[2] != 0x44 || bytes[3] != 0x46 || bytes[4] != 0x2D) {
            throw new InvalidDataException($"{name} does not have a valid PDF header.");
        }
    }

    private static BrowserConversionArtifact CreateReport(PdfToolRequest request, PdfToolExecution execution) {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = true })) {
            writer.WriteStartObject();
            writer.WriteNumber("schemaVersion", 1);
            writer.WriteString("tool", request.Tool.Id);
            writer.WriteString("engine", "OfficeIMO.Pdf");
            writer.WriteBoolean("browserLocal", true);
            writer.WriteString("summary", execution.Summary);
            if (execution.PageCount.HasValue) writer.WriteNumber("pageCount", execution.PageCount.Value);
            writer.WriteStartArray("inputs");
            foreach (SelectedDocument file in request.Files) {
                writer.WriteStartObject();
                writer.WriteString("fileName", file.Name);
                writer.WriteNumber("bytes", file.Bytes.LongLength);
                writer.WriteString("sha256", Convert.ToHexString(SHA256.HashData(file.Bytes)).ToLowerInvariant());
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteStartObject("output");
            writer.WriteString("fileName", execution.Artifact.FileName);
            writer.WriteString("contentType", execution.Artifact.ContentType);
            writer.WriteNumber("bytes", execution.Artifact.Bytes.LongLength);
            writer.WriteString("sha256", Convert.ToHexString(SHA256.HashData(execution.Artifact.Bytes)).ToLowerInvariant());
            writer.WriteEndObject();
            writer.WriteStartObject("details");
            foreach ((string key, string value) in execution.ReportDetails!) writer.WriteString(key, value);
            writer.WriteEndObject();
            writer.WriteStartArray("messages");
            foreach (PdfToolMessage message in execution.Messages) {
                writer.WriteStartObject();
                writer.WriteString("title", message.Title);
                writer.WriteString("message", message.Message);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
        }
        return new BrowserConversionArtifact(
            stream.ToArray(),
            Path.GetFileNameWithoutExtension(execution.Artifact.FileName) + ".officeimo-report.json",
            "application/json");
    }

    private static string FormatBytes(long bytes) {
        string[] units = ["B", "KB", "MB", "GB"];
        double value = bytes;
        int unit = 0;
        while (value >= 1024 && unit < units.Length - 1) { value /= 1024; unit++; }
        return unit == 0 ? $"{bytes} B" : $"{value:0.##} {units[unit]}";
    }

    private sealed record PdfToolExecution(
        BrowserConversionArtifact Artifact,
        string Summary,
        int? PageCount,
        IReadOnlyList<PdfToolMessage> Messages,
        IReadOnlyDictionary<string, string>? ReportDetails,
        bool PreviewInBrowser = true);
}
