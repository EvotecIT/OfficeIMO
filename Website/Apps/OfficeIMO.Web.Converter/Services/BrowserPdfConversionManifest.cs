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
        string optionProfile,
        BrowserPdfProfile profile,
        long conversionMilliseconds,
        PdfSerializationReport serialization) {
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
            profile.Id,
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
                sourceCommit = GetSourceCommit(engineVersion),
                optionProfile,
                profile = new {
                    id = profile.Id,
                    label = profile.Label,
                    description = profile.Description
                }
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
                profile = "browser",
                packageBytes = BrowserConversionService.MaxPackageBytes,
                packagePartCount = BrowserConversionService.MaxPackagePartCount,
                partUncompressedBytes = BrowserConversionService.MaxPartUncompressedBytes,
                totalUncompressedBytes = BrowserConversionService.MaxTotalUncompressedBytes,
                compressionRatio = BrowserConversionService.MaxCompressionRatio
            },
            performance = new {
                conversionMilliseconds,
                peakRetainedPageContentBytes = serialization.PeakRetainedPageContentBytes,
                peakRetainedObjectBytes = serialization.PeakRetainedObjectBytes,
                peakRetainedCompletedPayloadBytes = AddWithoutOverflow(
                    serialization.PeakRetainedPageContentBytes,
                    serialization.PeakRetainedObjectBytes),
                pageContentSpilled = serialization.PageContentSpilled,
                objectBufferSpilled = serialization.ObjectBufferSpilled,
                finalArtifactBuffered = serialization.FinalArtifactBuffered,
                isForwardOnlyObjectSerialization = serialization.IsForwardOnlyObjectSerialization,
                largestSerializedObjectBytes = serialization.LargestSerializedObjectBytes,
                isForwardOnlyLayout = serialization.IsForwardOnlyLayout
            },
            warnings = report.Warnings.Select(warning => new {
                converter = warning.Converter,
                code = warning.Code,
                source = warning.Source,
                message = warning.Message,
                severity = warning.Severity.ToString(),
                construct = warning.LayoutDiagnostic?.Kind.ToString()
                    ?? (warning.Details.TryGetValue("construct", out string? construct) ? construct : warning.Code),
                pageNumber = TryReadPositiveInt(warning.Details, "pageNumber")
                    ?? TryReadPositiveInt(warning.Details, "page"),
                canChangePagination =
                    warning.Code.Contains("font", StringComparison.OrdinalIgnoreCase) ||
                    warning.Code.Contains("pagination", StringComparison.OrdinalIgnoreCase) ||
                    warning.Code.Contains("overflow", StringComparison.OrdinalIgnoreCase) ||
                    warning.LayoutDiagnostic?.Kind is PdfLayoutDiagnosticKind.AdjustedGeometry
                        or PdfLayoutDiagnosticKind.ClippedContent
                        or PdfLayoutDiagnosticKind.Overflow,
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

    private static string? GetSourceCommit(string informationalVersion) {
        int separator = informationalVersion.LastIndexOf('+');
        if (separator < 0 || separator == informationalVersion.Length - 1) {
            return null;
        }

        string metadata = informationalVersion[(separator + 1)..];
        return metadata.Length >= 7 && metadata.All(static value => char.IsAsciiHexDigit(value))
            ? metadata
            : null;
    }

    private static int? TryReadPositiveInt(IReadOnlyDictionary<string, string> values, string key) =>
        values.TryGetValue(key, out string? value) &&
        int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int parsed) &&
        parsed > 0
            ? parsed
            : null;

    private static long AddWithoutOverflow(long first, long second) =>
        first > long.MaxValue - second ? long.MaxValue : first + second;

    private static string Sha256(byte[] bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
