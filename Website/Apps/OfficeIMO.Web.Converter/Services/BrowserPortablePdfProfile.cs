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
    internal const string DefaultFontFamily = HtmlPortableBrowserFontProfile.DefaultFontFamily;
    internal const string JapaneseFallbackFontFamily = HtmlPortableBrowserFontProfile.JapaneseFallbackFontFamily;
    internal const string ArabicFallbackFontFamily = HtmlPortableBrowserFontProfile.ArabicFallbackFontFamily;
    internal const string SymbolFallbackFontFamily = HtmlPortableBrowserFontProfile.SymbolFallbackFontFamily;
    internal const string DefaultLayoutFontFamilies = HtmlPortableBrowserFontProfile.DefaultLayoutFontFamilies;
    internal const string ExpectedFontPackFingerprint = "f69721965e37295fd9f1372c5547a95ab2863ff6afb357226989719e409f8030";
    /// <summary>
    /// SHA-256 of the normalized manifest, which pins every font file's SHA-256. Verifying the manifest and each
    /// loaded font lets the fallback assembly arrive later while the reported whole-package identity stays fixed.
    /// </summary>
    internal const string ExpectedManifestSha256 = "ea6ebd42204e4e6106a57348b558eedb943fda26b1e45b4e25f907bf4b62ae00";

    private static readonly Lazy<FontPackData> Data = new(LoadFontPack, isThreadSafe: true);
    private static readonly Lock FallbackGate = new();
    private static HtmlPortableBrowserFontProfile? _fallback;

    internal static string FontPackId => Data.Value.Id;
    internal static string FontPackFingerprint => Data.Value.Fingerprint;
    internal static IReadOnlyList<string> FontCoverage => Data.Value.Coverage;
    internal static IReadOnlyList<PdfFontFamilySubstitution> FontFamilySubstitutions =>
        Data.Value.PortableProfile.FontFamilySubstitutions;
    internal static byte[] CarlitoRegularBytes => Data.Value.CarlitoRegular;

    /// <summary>One immutable shared profile for the assemblies currently available to the host.</summary>
    private static HtmlPortableBrowserFontProfile CurrentProfile {
        get {
            if (Volatile.Read(ref _fallback) is { } loaded) return loaded;
            FontPackData data = Data.Value;
            Assembly? assembly = BrowserFallbackFonts.FindAssembly();
            if (assembly == null) return data.PortableProfile;
            lock (FallbackGate) {
                return _fallback ??= data.PortableProfile.WithFallbackFonts(
                    name => ReadVerified(assembly, name, data.FileHashes));
            }
        }
    }

    internal static OfficeFontFaceCollection CreateDrawingFonts() => CurrentProfile.CreateDrawingFonts();

    internal static PdfOptions CreateOptions(BrowserPdfProfile profile) {
        ArgumentNullException.ThrowIfNull(profile);
        PdfOptions options = CurrentProfile.CreatePdfOptions(OfficeHarfBuzzTextShapingProvider.Instance);
        ApplyComplianceProfile(options, profile);
        return options;
    }

    internal static HtmlToPdfOptions CreateHtmlOptions(BrowserPdfProfile profile) {
        ArgumentNullException.ThrowIfNull(profile);
        HtmlToPdfOptions options = CurrentProfile.CreateHtmlOptions(OfficeHarfBuzzTextShapingProvider.Instance);
        ApplyComplianceProfile(options.PdfOptions, profile);
        return options;
    }

    private static void ApplyComplianceProfile(PdfOptions options, BrowserPdfProfile profile) {
        if (profile.Kind == BrowserPdfProfileKind.Archival) {
            options
                .UsePdfA(PdfComplianceProfile.PdfA2B, "und")
                .RequireCompliance(PdfComplianceProfile.PdfA2B);
        } else if (profile.Kind == BrowserPdfProfileKind.Accessible) {
            // Source adapters can replace the unknown language with the source language.
            options.UsePdfUa(PdfComplianceProfile.PdfUa1, "und");
        }
    }

    private static FontPackData LoadFontPack() {
        Assembly assembly = typeof(OfficeIMO.Web.Fonts.BrowserFontResources).Assembly;
        byte[] manifestBytes = ReadResource(assembly, "font-pack.json");
        string manifestHash = Sha256(NormalizeManifest(manifestBytes));
        if (!string.Equals(manifestHash, ExpectedManifestSha256, StringComparison.Ordinal)) {
            throw new InvalidOperationException(
                $"The embedded browser PDF font pack manifest '{manifestHash}' does not match the pinned profile '{ExpectedManifestSha256}'.");
        }
        FontPackManifest manifest = JsonSerializer.Deserialize(
            manifestBytes,
            BrowserFontPackJsonContext.Default.FontPackManifest)
            ?? throw new InvalidOperationException("The embedded browser PDF font pack manifest is invalid.");
        ValidateManifest(manifest);
        IReadOnlyList<string> coverage = ValidateCoverage(manifest.Coverage);
        IReadOnlyDictionary<string, string> hashes = DeclaredHashes(manifest);
        byte[] carlitoRegular = ReadVerified(assembly, "Carlito-Regular.ttf", hashes);
        HtmlPortableBrowserFontProfile portableProfile = HtmlPortableBrowserFontProfile.CreateDefaultOnly(
            manifest.Id,
            name => name == "Carlito-Regular.ttf" ? carlitoRegular : ReadVerified(assembly, name, hashes));
        return new FontPackData(manifest.Id, coverage, hashes, carlitoRegular, portableProfile, ExpectedFontPackFingerprint);
    }

    /// <summary>Recomputes the whole-pack fingerprint from both font assemblies. Tests use it to keep the pinned values honest.</summary>
    internal static string ComputeFullFingerprint() {
        System.Reflection.Assembly fonts = typeof(OfficeIMO.Web.Fonts.BrowserFontResources).Assembly;
        System.Reflection.Assembly fallback = BrowserFallbackFonts.FindAssembly()
            ?? throw new InvalidOperationException("The fallback font assembly is not available.");
        var assets = new Dictionary<string, byte[]>(StringComparer.Ordinal) {
            ["font-pack.json"] = NormalizeManifest(ReadResource(fonts, "font-pack.json"))
        };
        foreach (string name in new[] { "Carlito-Bold.ttf", "Carlito-BoldItalic.ttf", "Carlito-Italic.ttf", "Carlito-Regular.ttf" }) {
            assets[name] = ReadResource(fonts, name);
        }
        foreach (string name in new[] { "NotoSansJP-OfficeIMO-Common.ttf", "NotoSansArabic-Regular.ttf", "NotoSansSymbols2-Regular.ttf" }) {
            assets[name] = ReadResource(fallback, name);
        }
        return ComputeFingerprint(assets);
    }

    private static IReadOnlyDictionary<string, string> DeclaredHashes(FontPackManifest manifest) {
        var hashes = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (FontPackFont font in manifest.Fonts) {
            foreach (FontPackFile file in font.Files ?? []) {
                if (string.IsNullOrWhiteSpace(file.Name) || string.IsNullOrWhiteSpace(file.Sha256) || !hashes.TryAdd(file.Name, file.Sha256.ToLowerInvariant())) {
                    throw new InvalidOperationException("The embedded browser PDF font pack manifest has an invalid or duplicate file entry.");
                }
            }
        }
        return hashes;
    }

    private static byte[] ReadVerified(System.Reflection.Assembly assembly, string fileName, IReadOnlyDictionary<string, string> hashes) {
        byte[] bytes = ReadResource(assembly, fileName);
        if (!hashes.TryGetValue(fileName, out string? expected) || !string.Equals(Sha256(bytes), expected, StringComparison.Ordinal)) {
            throw new InvalidOperationException($"The browser PDF font '{fileName}' does not match the font pack manifest.");
        }
        return bytes;
    }

    private static byte[] NormalizeManifest(byte[] manifestBytes) => Encoding.UTF8.GetBytes(
        Encoding.UTF8.GetString(manifestBytes)
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace("\r", "\n", StringComparison.Ordinal));

    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

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

    private static byte[] ReadResource(Assembly assembly, string fileName) {
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
        IReadOnlyDictionary<string, string> FileHashes,
        byte[] CarlitoRegular,
        HtmlPortableBrowserFontProfile PortableProfile,
        string Fingerprint);
}
