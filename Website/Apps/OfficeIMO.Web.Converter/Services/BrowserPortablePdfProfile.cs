using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing.HarfBuzz;
using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Models;

namespace OfficeIMO.Web.Converter.Services;

/// <summary>
/// Supplies the explicit, host-independent PDF font profile used by browser conversions.
/// </summary>
internal static class BrowserPortablePdfProfile {
    internal const string DefaultFontFamily = "Carlito";
    internal const string ExpectedFontPackFingerprint = "58d48fe49e16ffa209a594a905260e81c7bcd5fb10aaced1e76601d2f18cea68";

    private static readonly Lazy<FontPackData> Data = new(LoadFontPack, isThreadSafe: true);

    internal static string FontPackId => Data.Value.Id;
    internal static string FontPackFingerprint => Data.Value.Fingerprint;
    internal static IReadOnlyList<PdfFontFamilySubstitution> FontFamilySubstitutions =>
        Data.Value.Substitutions;

    internal static PdfOptions CreateOptions(BrowserPdfProfile profile) {
        ArgumentNullException.ThrowIfNull(profile);
        FontPackData data = Data.Value;
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            HeaderFont = PdfStandardFont.Helvetica,
            FooterFont = PdfStandardFont.Helvetica,
            FileVersion = PdfFileVersion.Pdf17,
            ObjectSerializationMode = PdfObjectSerializationMode.ForwardOnly,
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers,
            TextShapingMode = PdfTextShapingMode.LatinLigatures
        }.SetTextShapingProvider(OfficeHarfBuzzTextShapingProvider.Instance);
        if (profile.Kind == BrowserPdfProfileKind.Accessible) {
            // The browser profile does not know the source language yet. Keep
            // the catalog explicitly undefined so source adapters can replace
            // it instead of mis-tagging every accessible document as English.
            options.UsePdfUa(PdfComplianceProfile.PdfUa1, "und");
        }

        options.RegisterFontFamily(
            PdfStandardFont.Helvetica,
            new PdfEmbeddedFontFamily(
                DefaultFontFamily,
                data.CarlitoRegular,
                data.CarlitoBold,
                data.CarlitoItalic,
                data.CarlitoBoldItalic));
        options.RegisterNamedFontFamily(
            new PdfEmbeddedFontFamily(
                DefaultFontFamily,
                data.CarlitoRegular,
                data.CarlitoBold,
                data.CarlitoItalic,
                data.CarlitoBoldItalic));
        options.RegisterEmbeddedFontFallbacks(
            new PdfEmbeddedFontFallbackSet([
                new PdfEmbeddedFontFallbackCandidate("Noto Sans Arabic", data.NotoSansArabic),
                new PdfEmbeddedFontFallbackCandidate("Noto Sans Symbols 2", data.NotoSansSymbols)
            ]));
        foreach (PdfFontFamilySubstitution substitution in data.Substitutions) {
            options.RegisterFontFamilySubstitution(
                substitution.SourceFontFamily,
                substitution.TargetFontFamily,
                substitution.Impact);
        }

        return options;
    }

    private static FontPackData LoadFontPack() {
        byte[] manifestBytes = ReadResource("font-pack.json");
        FontPackManifest manifest = JsonSerializer.Deserialize<FontPackManifest>(
            manifestBytes,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidOperationException("The embedded browser PDF font pack manifest is invalid.");
        byte[] carlitoRegular = ReadResource("Carlito-Regular.ttf");
        byte[] carlitoBold = ReadResource("Carlito-Bold.ttf");
        byte[] carlitoItalic = ReadResource("Carlito-Italic.ttf");
        byte[] carlitoBoldItalic = ReadResource("Carlito-BoldItalic.ttf");
        byte[] notoSansArabic = ReadResource("NotoSansArabic-Regular.ttf");
        byte[] notoSansSymbols = ReadResource("NotoSansSymbols2-Regular.ttf");

        byte[] normalizedManifestBytes = Encoding.UTF8.GetBytes(
            Encoding.UTF8.GetString(manifestBytes)
                .Replace("\r\n", "\n", StringComparison.Ordinal)
                .Replace("\r", "\n", StringComparison.Ordinal));
        var assets = new Dictionary<string, byte[]>(StringComparer.Ordinal) {
            ["Carlito-Bold.ttf"] = carlitoBold,
            ["Carlito-BoldItalic.ttf"] = carlitoBoldItalic,
            ["Carlito-Italic.ttf"] = carlitoItalic,
            ["Carlito-Regular.ttf"] = carlitoRegular,
            ["NotoSansArabic-Regular.ttf"] = notoSansArabic,
            ["NotoSansSymbols2-Regular.ttf"] = notoSansSymbols,
            ["font-pack.json"] = normalizedManifestBytes
        };

        IReadOnlyList<PdfFontFamilySubstitution> substitutions = ValidateManifest(manifest);
        string fingerprint = ComputeFingerprint(assets);
        if (!string.Equals(fingerprint, ExpectedFontPackFingerprint, StringComparison.Ordinal)) {
            throw new InvalidOperationException(
                $"The embedded browser PDF font pack fingerprint '{fingerprint}' does not match the pinned profile '{ExpectedFontPackFingerprint}'.");
        }

        return new FontPackData(
            manifest.Id,
            carlitoRegular,
            carlitoBold,
            carlitoItalic,
            carlitoBoldItalic,
            notoSansArabic,
            notoSansSymbols,
            substitutions,
            fingerprint);
    }

    private static IReadOnlyList<PdfFontFamilySubstitution> ValidateManifest(FontPackManifest manifest) {
        if (string.IsNullOrWhiteSpace(manifest.Id)) {
            throw new InvalidOperationException("The embedded browser PDF font pack manifest has no id.");
        }

        var targetFamilies = new HashSet<string>(
            manifest.Fonts
                .Where(static font => !string.IsNullOrWhiteSpace(font.Family))
                .Select(static font => font.Family.Trim()),
            StringComparer.OrdinalIgnoreCase);
        var sourceFamilies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var substitutions = new List<PdfFontFamilySubstitution>();
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

            substitutions.Add(new PdfFontFamilySubstitution(
                declared.Source,
                declared.Target,
                impact));
        }

        return substitutions.AsReadOnly();
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
        byte[] CarlitoRegular,
        byte[] CarlitoBold,
        byte[] CarlitoItalic,
        byte[] CarlitoBoldItalic,
        byte[] NotoSansArabic,
        byte[] NotoSansSymbols,
        IReadOnlyList<PdfFontFamilySubstitution> Substitutions,
        string Fingerprint);

    private sealed record FontPackManifest(
        string Id,
        IReadOnlyList<FontPackFont> Fonts,
        IReadOnlyList<FontPackSubstitution> Substitutions);

    private sealed record FontPackFont(string Family);

    private sealed record FontPackSubstitution(string Source, string Target, string Impact);
}
