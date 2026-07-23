using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Models;

namespace OfficeIMO.Web.Converter.Services;

internal static class BrowserPdfConversionManifest {
    private const string SchemaVersion = "1";

    internal static BrowserConversionArtifact Create(
        SelectedDocument source,
        string outputFileName,
        byte[] outputBytes,
        PdfConversionReport report,
        string converter,
        string optionProfile) {
        PdfReadDocument readDocument = PdfReadDocument.Open(outputBytes);
        string sourceHash = Sha256(source.Bytes);
        string outputHash = Sha256(outputBytes);
        string engineVersion = GetEngineVersion();
        string conversionId = Sha256(Encoding.UTF8.GetBytes(string.Join("|", [
            SchemaVersion,
            converter,
            sourceHash,
            source.Name,
            engineVersion,
            BrowserPortablePdfProfile.FontPackId,
            BrowserPortablePdfProfile.FontPackFingerprint,
            optionProfile,
            "portable-deterministic",
            PdfTaggedStructureMode.CatalogMarkers.ToString()
        ])));

        var manifest = new {
            schemaVersion = SchemaVersion,
            conversionId,
            fidelityStatus = report.FidelityStatus.ToString(),
            source = new {
                fileName = source.Name,
                byteCount = source.Size,
                sha256 = sourceHash
            },
            output = new {
                fileName = outputFileName,
                byteCount = outputBytes.LongLength,
                sha256 = outputHash,
                pageCount = readDocument.Pages.Count,
                tagged = readDocument.HasTaggedContent
            },
            engine = new {
                converter,
                assembly = typeof(PdfDocument).Assembly.GetName().Name,
                version = engineVersion,
                optionProfile
            },
            fontPack = new {
                id = BrowserPortablePdfProfile.FontPackId,
                fingerprint = BrowserPortablePdfProfile.FontPackFingerprint,
                defaultFamily = BrowserPortablePdfProfile.DefaultFontFamily,
                coverage = new[] { "Latin", "Arabic glyphs", "common symbols" }
            },
            policy = new {
                resources = "portable-deterministic",
                taggedStructure = PdfTaggedStructureMode.CatalogMarkers.ToString(),
                systemFonts = false,
                externalResources = false
            },
            limits = new {
                packageBytes = BrowserConversionService.MaxPackageBytes,
                packagePartCount = BrowserConversionService.MaxPackagePartCount,
                partUncompressedBytes = BrowserConversionService.MaxPartUncompressedBytes,
                totalUncompressedBytes = BrowserConversionService.MaxTotalUncompressedBytes,
                compressionRatio = BrowserConversionService.MaxCompressionRatio
            },
            warnings = report.Warnings.Select(warning => new {
                converter = warning.Converter,
                code = warning.Code,
                source = warning.Source,
                message = warning.Message,
                severity = warning.Severity.ToString(),
                details = warning.Details.OrderBy(pair => pair.Key, StringComparer.Ordinal)
                    .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal)
            }).ToArray()
        };

        byte[] bytes = JsonSerializer.SerializeToUtf8Bytes(
            manifest,
            new JsonSerializerOptions { WriteIndented = true });
        return new BrowserConversionArtifact(
            bytes,
            Path.GetFileNameWithoutExtension(source.Name) + ".conversion.json",
            "application/json");
    }

    private static string GetEngineVersion() {
        Assembly assembly = typeof(PdfDocument).Assembly;
        return assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
            ?? assembly.GetName().Version?.ToString()
            ?? "unknown";
    }

    private static string Sha256(byte[] bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
