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

        var manifest = new ConversionManifestDocument(
            SchemaVersion,
            conversionId,
            report.FidelityStatus.ToString(),
            new ManifestSource(source.Name, source.Size, sourceHash),
            new ManifestOutput(outputFileName, outputBytes.LongLength, outputHash, serialization.PageCount, Tagged: true),
            new ManifestEngine(
                converter,
                typeof(PdfDocument).Assembly.GetName().Name,
                engineVersion,
                GetSourceCommit(engineVersion),
                optionProfile,
                new ManifestProfile(profile.Id, profile.Label, profile.Description)),
            new ManifestFontPack(
                BrowserPortablePdfProfile.FontPackId,
                BrowserPortablePdfProfile.FontPackFingerprint,
                BrowserPortablePdfProfile.DefaultFontFamily,
                BrowserPortablePdfProfile.FontCoverage,
                BrowserPortablePdfProfile.FontFamilySubstitutions.Select(
                    static substitution => new ManifestSubstitution(
                        substitution.SourceFontFamily,
                        substitution.TargetFontFamily,
                        substitution.Impact.ToString())).ToArray()),
            new ManifestPolicy(
                "portable-deterministic",
                PdfTaggedStructureMode.CatalogMarkers.ToString(),
                SystemFonts: false,
                ExternalResources: false),
            new ManifestLimits(
                "browser",
                BrowserConversionService.MaxPackageBytes,
                BrowserConversionService.MaxPackagePartCount,
                BrowserConversionService.MaxPartUncompressedBytes,
                BrowserConversionService.MaxTotalUncompressedBytes,
                BrowserConversionService.MaxCompressionRatio),
            new ManifestPerformance(
                conversionMilliseconds,
                serialization.PeakRetainedPageContentBytes,
                serialization.PeakRetainedObjectBytes,
                AddWithoutOverflow(
                    serialization.PeakRetainedPageContentBytes,
                    serialization.PeakRetainedObjectBytes),
                serialization.PageContentSpilled,
                serialization.ObjectBufferSpilled,
                serialization.FinalArtifactBuffered,
                serialization.IsForwardOnlyObjectSerialization,
                serialization.LargestSerializedObjectBytes,
                serialization.IsForwardOnlyLayout),
            report.Warnings.Select(warning => new ManifestWarning(
                warning.Converter,
                warning.Code,
                warning.Source,
                warning.Message,
                warning.Severity.ToString(),
                warning.LayoutDiagnostic?.Kind.ToString()
                    ?? (warning.Details.TryGetValue("construct", out string? construct) ? construct : warning.Code),
                TryReadPositiveInt(warning.Details, "pageNumber")
                    ?? TryReadPositiveInt(warning.Details, "page"),
                warning.Severity != PdfConversionWarningSeverity.Information &&
                    (warning.Code.Contains("font", StringComparison.OrdinalIgnoreCase) ||
                     warning.Code.Contains("pagination", StringComparison.OrdinalIgnoreCase) ||
                     warning.Code.Contains("overflow", StringComparison.OrdinalIgnoreCase) ||
                     warning.LayoutDiagnostic?.Kind is PdfLayoutDiagnosticKind.AdjustedGeometry
                         or PdfLayoutDiagnosticKind.ClippedContent
                         or PdfLayoutDiagnosticKind.Overflow),
                warning.Details.OrderBy(pair => pair.Key, StringComparer.Ordinal)
                    .ToDictionary(pair => pair.Key, pair => pair.Value, StringComparer.Ordinal))).ToArray());

        byte[] bytes = JsonSerializer.SerializeToUtf8Bytes(manifest, BrowserReportJsonContext.Default.ConversionManifestDocument);
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
