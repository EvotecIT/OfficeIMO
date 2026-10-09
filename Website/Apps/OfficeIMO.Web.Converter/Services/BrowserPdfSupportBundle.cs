using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Models;

namespace OfficeIMO.Web.Converter.Services;

internal static class BrowserPdfSupportBundle {
    private static readonly DateTimeOffset StableEntryTimestamp =
        new(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);

    internal static BrowserConversionArtifact Create(
        SelectedDocument source,
        ConversionResult result,
        bool includeDocumentContent) {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(result);
        if (!string.Equals(result.ContentType, "application/pdf", StringComparison.Ordinal)) {
            throw new InvalidOperationException("Support bundles are available for PDF conversion results.");
        }

        var document = new SupportSummaryDocument(
            "1",
            new SupportPrivacy(includeDocumentContent, "fingerprints-and-diagnostics-only"),
            new SupportSource(source.Extension, source.Size, Sha256(source.Bytes)),
            new SupportOutput(
                result.Bytes.LongLength,
                Sha256(result.Bytes),
                result.PageCount,
                PdfReadDocument.Open(result.Bytes).HasTaggedContent),
            result.Profile is null
                ? null
                : new ManifestProfile(result.Profile.Id, result.Profile.Label, result.Profile.Description),
            new SupportEngine(
                typeof(PdfDocument).Assembly.GetName().Name,
                typeof(PdfDocument).Assembly
                    .GetCustomAttributes(typeof(System.Reflection.AssemblyInformationalVersionAttribute), false)
                    .OfType<System.Reflection.AssemblyInformationalVersionAttribute>()
                    .Select(static attribute => attribute.InformationalVersion)
                    .FirstOrDefault()
                    ?? typeof(PdfDocument).Assembly.GetName().Version?.ToString()
                    ?? "unknown",
                BrowserPortablePdfProfile.FontPackId,
                BrowserPortablePdfProfile.FontPackFingerprint),
            new SupportPerformance(result.ConversionMilliseconds, result.PeakRetainedMemoryBytes),
            result.StructuredWarnings.Select(static warning => new SupportWarning(
                warning.Code,
                warning.Source,
                warning.Message,
                warning.Severity,
                warning.Construct,
                warning.PageNumber,
                warning.CanChangePagination)).ToArray());
        byte[] summary = JsonSerializer.SerializeToUtf8Bytes(document, BrowserReportJsonContext.Default.SupportSummaryDocument);

        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            AddEntry(
                archive,
                "README.txt",
                Encoding.UTF8.GetBytes(
                    includeDocumentContent
                        ? "OfficeIMO browser conversion support bundle.\nThis bundle includes source and PDF content because the user explicitly opted in.\n"
                        : "OfficeIMO browser conversion support bundle.\nThis default bundle contains fingerprints, configuration, performance evidence, and diagnostics only. It does not contain source or PDF document bytes.\n"));
            AddEntry(archive, "support-summary.json", summary);
            if (includeDocumentContent) {
                string extension = string.IsNullOrWhiteSpace(source.Extension) ? ".bin" : source.Extension;
                AddEntry(archive, "content/source" + extension, source.Bytes);
                AddEntry(archive, "content/result.pdf", result.Bytes);
            }
        }

        return new BrowserConversionArtifact(
            output.ToArray(),
            "officeimo-pdf-support.zip",
            "application/zip");
    }

    private static void AddEntry(ZipArchive archive, string name, byte[] bytes) {
        ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        entry.LastWriteTime = StableEntryTimestamp;
        using Stream stream = entry.Open();
        stream.Write(bytes);
    }

    private static string Sha256(byte[] bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
