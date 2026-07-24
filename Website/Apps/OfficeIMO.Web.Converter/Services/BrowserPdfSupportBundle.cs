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

        byte[] summary = JsonSerializer.SerializeToUtf8Bytes(
            new {
                schemaVersion = "1",
                privacy = new {
                    includesDocumentContent = includeDocumentContent,
                    defaultPolicy = "fingerprints-and-diagnostics-only"
                },
                source = new {
                    extension = source.Extension,
                    byteCount = source.Size,
                    sha256 = Sha256(source.Bytes)
                },
                output = new {
                    byteCount = result.Bytes.LongLength,
                    sha256 = Sha256(result.Bytes),
                    result.PageCount,
                    tagged = PdfReadDocument.Open(result.Bytes).HasTaggedContent
                },
                profile = result.Profile is null
                    ? null
                    : new {
                        id = result.Profile.Id,
                        label = result.Profile.Label,
                        description = result.Profile.Description
                    },
                engine = new {
                    assembly = typeof(PdfDocument).Assembly.GetName().Name,
                    version = typeof(PdfDocument).Assembly
                        .GetCustomAttributes(typeof(System.Reflection.AssemblyInformationalVersionAttribute), false)
                        .OfType<System.Reflection.AssemblyInformationalVersionAttribute>()
                        .Select(static attribute => attribute.InformationalVersion)
                        .FirstOrDefault()
                        ?? typeof(PdfDocument).Assembly.GetName().Version?.ToString()
                        ?? "unknown",
                    fontPackId = BrowserPortablePdfProfile.FontPackId,
                    fontPackFingerprint = BrowserPortablePdfProfile.FontPackFingerprint
                },
                performance = new {
                    result.ConversionMilliseconds,
                    result.PeakRetainedMemoryBytes
                },
                warnings = result.StructuredWarnings
            },
            new JsonSerializerOptions { WriteIndented = true });

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
