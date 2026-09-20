using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.HarfBuzz;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Models;

namespace OfficeIMO.Web.Converter.Services;

/// <summary>
/// Supplies the explicit, host-independent PDF font profile used by browser conversions.
/// </summary>
internal static class BrowserPortablePdfProfile {
    private static readonly string[] FontAssetNames = [
        "Carlito-Bold.ttf",
        "Carlito-BoldItalic.ttf",
        "Carlito-Italic.ttf",
        "Carlito-Regular.ttf",
        "NotoSansJP-OfficeIMO-Common.ttf",
        "NotoSansArabic-Regular.ttf",
        "NotoSansSymbols2-Regular.ttf"
    ];

    internal const string DefaultFontFamily = HtmlPortableBrowserFontProfile.DefaultFontFamily;
    internal const string JapaneseFallbackFontFamily = HtmlPortableBrowserFontProfile.JapaneseFallbackFontFamily;
    internal const string ArabicFallbackFontFamily = HtmlPortableBrowserFontProfile.ArabicFallbackFontFamily;
    internal const string SymbolFallbackFontFamily = HtmlPortableBrowserFontProfile.SymbolFallbackFontFamily;
    internal const string DefaultLayoutFontFamilies = HtmlPortableBrowserFontProfile.DefaultLayoutFontFamilies;
    internal const string ExpectedFontPackFingerprint = "f69721965e37295fd9f1372c5547a95ab2863ff6afb357226989719e409f8030";

    private static readonly Lazy<FontPackData> Data = new(LoadFontPack, isThreadSafe: true);

    internal static string FontPackId => Data.Value.Id;
    internal static string FontPackFingerprint => Data.Value.Fingerprint;
    internal static IReadOnlyList<string> FontCoverage => Data.Value.Coverage;
    internal static IReadOnlyList<PdfFontFamilySubstitution> FontFamilySubstitutions =>
        Data.Value.PortableProfile.FontFamilySubstitutions;

    internal static OfficeFontFaceCollection CreateDrawingFonts() => Data.Value.PortableProfile.CreateDrawingFonts();

    internal static PdfOptions CreateOptions(BrowserPdfProfile profile) {
        ArgumentNullException.ThrowIfNull(profile);
        FontPackData data = Data.Value;
        PdfOptions options = data.PortableProfile.CreatePdfOptions(OfficeHarfBuzzTextShapingProvider.Instance);
        if (profile.Kind == BrowserPdfProfileKind.Archival) {
            options
                .UsePdfA(PdfComplianceProfile.PdfA2B, "und")
                .RequireCompliance(PdfComplianceProfile.PdfA2B);
        } else if (profile.Kind == BrowserPdfProfileKind.Accessible) {
            // The browser profile does not know the source language yet. Keep
            // the catalog explicitly undefined so source adapters can replace
            // it instead of mis-tagging every accessible document as English.
            options.UsePdfUa(PdfComplianceProfile.PdfUa1, "und");
        }

        return options;
    }

    internal static HtmlToPdfOptions CreateHtmlOptions(BrowserPdfProfile profile) {
        HtmlToPdfOptions options = Data.Value.PortableProfile.CreateHtmlOptions(OfficeHarfBuzzTextShapingProvider.Instance);
        options.PdfOptions = CreateOptions(profile);
        return options;
    }

    private static FontPackData LoadFontPack() {
        byte[] manifestBytes = ReadResource("font-pack.json");
        FontPackManifest manifest = JsonSerializer.Deserialize<FontPackManifest>(
            manifestBytes,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidOperationException("The embedded browser PDF font pack manifest is invalid.");
        byte[] normalizedManifestBytes = Encoding.UTF8.GetBytes(
            Encoding.UTF8.GetString(manifestBytes)
                .Replace("\r\n", "\n", StringComparison.Ordinal)
                .Replace("\r", "\n", StringComparison.Ordinal));
        var assets = FontAssetNames.ToDictionary(
            static name => name,
            ReadResource,
            StringComparer.Ordinal);
        assets.Add("font-pack.json", normalizedManifestBytes);

        IReadOnlyList<string> coverage = ValidateCoverage(manifest.Coverage);
        ValidateManifest(manifest);
        string fingerprint = ComputeFingerprint(assets);
        if (!string.Equals(fingerprint, ExpectedFontPackFingerprint, StringComparison.Ordinal)) {
            throw new InvalidOperationException(
                $"The embedded browser PDF font pack fingerprint '{fingerprint}' does not match the pinned profile '{ExpectedFontPackFingerprint}'.");
        }

        HtmlPortableBrowserFontProfile portableProfile = HtmlPortableBrowserFontProfile.Create(manifest.Id, ReadResource);
        return new FontPackData(
            manifest.Id,
            coverage,
            fingerprint,
            portableProfile);
    }

    private static IReadOnlyList<string> ValidateCoverage(IReadOnlyList<string> declaredCoverage) {
        if (declaredCoverage == null || declaredCoverage.Count == 0) {
            throw new InvalidOperationException("The embedded browser PDF font pack declares no coverage.");
        }

        string[] coverage = declaredCoverage
            .Select(static value => value?.Trim())
            .Where(static value => !string.IsNullOrWhiteSpace(value))
            .Cast<string>()
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (coverage.Length != declaredCoverage.Count) {
            throw new InvalidOperationException(
                "The embedded browser PDF font pack contains empty or duplicate coverage declarations.");
        }
        return Array.AsReadOnly(coverage);
    }

    private static void ValidateManifest(FontPackManifest manifest) {
        if (string.IsNullOrWhiteSpace(manifest.Id)) {
            throw new InvalidOperationException("The embedded browser PDF font pack manifest has no id.");
        }

        var targetFamilies = new HashSet<string>(
            manifest.Fonts
                .Where(static font => !string.IsNullOrWhiteSpace(font.Family))
                .Select(static font => font.Family.Trim()),
            StringComparer.OrdinalIgnoreCase);
        var sourceFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (FontPackSubstitution declared in manifest.Substitutions) {
            if (string.IsNullOrWhiteSpace(declared.Source) ||
                string.IsNullOrWhiteSpace(declared.Target) ||
                !targetFamilies.Contains(declared.Target.Trim())) {
                throw new InvalidOperationException(
                    "The embedded browser PDF font pack contains an invalid substitution declaration.");
            }
            if (!sourceFamilies.Add(declared.Source.Trim())) {
                throw new InvalidOperationException(
                    "The embedded browser PDF font pack contains duplicate substitution sources.");
            }
            if (!Enum.TryParse(
                    declared.Impact,
                    ignoreCase: true,
                    out PdfFontFamilySubstitutionImpact impact) ||
                (impact != PdfFontFamilySubstitutionImpact.Compatible &&
                 impact != PdfFontFamilySubstitutionImpact.LayoutSensitive)) {
                throw new InvalidOperationException(
                    "The embedded browser PDF font pack contains an invalid substitution impact.");
            }
        }
    }

    private static byte[] ReadResource(string fileName) {
        Assembly assembly = typeof(BrowserPortablePdfProfile).Assembly;
        string resourceName = assembly.GetManifestResourceNames()
            .SingleOrDefault(name => name.EndsWith(".Assets.Fonts." + fileName, StringComparison.Ordinal))
            ?? throw new InvalidOperationException($"The browser PDF font resource '{fileName}' is missing.");
        using Stream stream = assembly.GetManifestResourceStream(resourceName)
            ?? throw new InvalidOperationException($"The browser PDF font resource '{fileName}' could not be opened.");
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    private static string ComputeFingerprint(IReadOnlyDictionary<string, byte[]> assets) {
        using IncrementalHash hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        foreach (KeyValuePair<string, byte[]> asset in assets.OrderBy(pair => pair.Key, StringComparer.Ordinal)) {
            hash.AppendData(Encoding.UTF8.GetBytes(asset.Key));
            hash.AppendData([0]);
            hash.AppendData(asset.Value);
        }

        return Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant();
    }

    private sealed record FontPackData(
        string Id,
        IReadOnlyList<string> Coverage,
        string Fingerprint,
        HtmlPortableBrowserFontProfile PortableProfile);

    private sealed record FontPackManifest(
        string Id,
        IReadOnlyList<string> Coverage,
        IReadOnlyList<FontPackFont> Fonts,
        IReadOnlyList<FontPackSubstitution> Substitutions);

    private sealed record FontPackFont(string Family);

    private sealed record FontPackSubstitution(string Source, string Target, string Impact);
}
