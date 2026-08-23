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
    private static readonly string[] PortableSansSerifAliases = [
        "Arial",
        "Helvetica",
        "Calibri",
        "Aptos",
        "Segoe UI",
        "Tahoma",
        "Verdana",
        "sans",
        "sans-serif",
        "ui-sans-serif",
        "system-ui",
        "-apple-system",
        "BlinkMacSystemFont"
    ];

    private static readonly string[] PortableSerifAliases = [
        "Times",
        "Times Roman",
        "Times-Roman",
        "Times New Roman",
        "serif"
    ];

    private static readonly string[] PortableMonospaceAliases = [
        "Courier",
        "Courier New",
        "monospace",
        "ui-monospace"
    ];

    private static readonly string[] PortableSymbolAliases = [
        "Symbol",
        "ZapfDingbats",
        "Zapf Dingbats"
    ];

    internal const string DefaultFontFamily = "Carlito";
    internal const string JapaneseFallbackFontFamily = "OfficeIMO Japanese Common";
    internal const string ArabicFallbackFontFamily = "Noto Sans Arabic";
    internal const string SymbolFallbackFontFamily = "Noto Sans Symbols 2";
    internal const string DefaultLayoutFontFamilies = "Carlito, 'OfficeIMO Japanese Common', 'Noto Sans Arabic', 'Noto Sans Symbols 2'";
    internal const string ExpectedFontPackFingerprint = "7cf393d8573f2cfeb6628defe7f5a08f95182bad204a5ab8182d74bad5a61cdf";

    private static readonly Lazy<FontPackData> Data = new(LoadFontPack, isThreadSafe: true);

    internal static string FontPackId => Data.Value.Id;
    internal static string FontPackFingerprint => Data.Value.Fingerprint;
    internal static IReadOnlyList<string> FontCoverage => Data.Value.Coverage;
    internal static IReadOnlyList<PdfFontFamilySubstitution> FontFamilySubstitutions =>
        Data.Value.Substitutions;

    internal static OfficeFontFaceCollection CreateDrawingFonts() => CreateLayoutFonts(Data.Value);

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

        options.RegisterFontFamily(
            PdfStandardFont.Helvetica,
            data.DefaultPdfFontFamily);
        options.RegisterNamedFontFamily(data.DefaultPdfFontFamily);
        options.RegisterEmbeddedFontFallbacks(data.PdfFontFallbacks);
        foreach (PdfFontFamilySubstitution substitution in data.Substitutions) {
            options.RegisterFontFamilySubstitution(
                substitution.SourceFontFamily,
                substitution.TargetFontFamily,
                substitution.Impact);
        }

        return options;
    }

    internal static HtmlPdfSaveOptions CreateHtmlOptions(BrowserPdfProfile profile) {
        FontPackData data = Data.Value;
        return new HtmlPdfSaveOptions {
            DefaultFontFamily = DefaultLayoutFontFamilies,
            Fonts = CreateLayoutFonts(data),
            PdfOptions = CreateOptions(profile),
            FontFamily = data.DefaultPdfFontFamily,
            TextShapingMode = PdfTextShapingMode.LatinLigatures,
            TextShapingProvider = OfficeHarfBuzzTextShapingProvider.Instance,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        };
    }

    private static OfficeFontFaceCollection CreateLayoutFonts(FontPackData data) {
        var fonts = new OfficeFontFaceCollection()
            .Add(DefaultFontFamily, data.CarlitoRegular, OfficeFontStyle.Regular)
            .Add(DefaultFontFamily, data.CarlitoBold, OfficeFontStyle.Bold)
            .Add(DefaultFontFamily, data.CarlitoItalic, OfficeFontStyle.Italic)
            .Add(DefaultFontFamily, data.CarlitoBoldItalic, OfficeFontStyle.Bold | OfficeFontStyle.Italic);
        foreach (string alias in PortableSansSerifAliases) {
            fonts.AddAlias(alias, DefaultFontFamily);
        }
        foreach (string alias in PortableSerifAliases) {
            fonts.AddAlias(alias, DefaultFontFamily);
        }
        foreach (string alias in PortableMonospaceAliases) {
            fonts.AddAlias(alias, DefaultFontFamily);
        }
        fonts.Add(SymbolFallbackFontFamily, data.NotoSansSymbols);
        foreach (string alias in PortableSymbolAliases) {
            fonts.AddAlias(alias, SymbolFallbackFontFamily);
        }
        return fonts
            .Add(JapaneseFallbackFontFamily, data.NotoSansJapaneseCommon)
            .Add(ArabicFallbackFontFamily, data.NotoSansArabic)
            .AddFallbackFamily(JapaneseFallbackFontFamily)
            .AddFallbackFamily(ArabicFallbackFontFamily)
            .AddFallbackFamily(SymbolFallbackFontFamily);
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
        byte[] notoSansJapaneseCommon = ReadResource("NotoSansJP-OfficeIMO-Common.ttf");
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
            ["NotoSansJP-OfficeIMO-Common.ttf"] = notoSansJapaneseCommon,
            ["NotoSansArabic-Regular.ttf"] = notoSansArabic,
            ["NotoSansSymbols2-Regular.ttf"] = notoSansSymbols,
            ["font-pack.json"] = normalizedManifestBytes
        };

        IReadOnlyList<string> coverage = ValidateCoverage(manifest.Coverage);
        IReadOnlyList<PdfFontFamilySubstitution> substitutions = ValidateManifest(manifest);
        string fingerprint = ComputeFingerprint(assets);
        if (!string.Equals(fingerprint, ExpectedFontPackFingerprint, StringComparison.Ordinal)) {
            throw new InvalidOperationException(
                $"The embedded browser PDF font pack fingerprint '{fingerprint}' does not match the pinned profile '{ExpectedFontPackFingerprint}'.");
        }

        return new FontPackData(
            manifest.Id,
            coverage,
            carlitoRegular,
            carlitoBold,
            carlitoItalic,
            carlitoBoldItalic,
            notoSansJapaneseCommon,
            notoSansArabic,
            notoSansSymbols,
            new PdfEmbeddedFontFamily(
                DefaultFontFamily,
                carlitoRegular,
                carlitoBold,
                carlitoItalic,
                carlitoBoldItalic),
            new PdfEmbeddedFontFallbackSet([
                new PdfEmbeddedFontFallbackCandidate(JapaneseFallbackFontFamily, notoSansJapaneseCommon),
                new PdfEmbeddedFontFallbackCandidate(ArabicFallbackFontFamily, notoSansArabic),
                new PdfEmbeddedFontFallbackCandidate(SymbolFallbackFontFamily, notoSansSymbols)
            ]),
            substitutions,
            fingerprint);
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
        IReadOnlyList<string> Coverage,
        byte[] CarlitoRegular,
        byte[] CarlitoBold,
        byte[] CarlitoItalic,
        byte[] CarlitoBoldItalic,
        byte[] NotoSansJapaneseCommon,
        byte[] NotoSansArabic,
        byte[] NotoSansSymbols,
        PdfEmbeddedFontFamily DefaultPdfFontFamily,
        PdfEmbeddedFontFallbackSet PdfFontFallbacks,
        IReadOnlyList<PdfFontFamilySubstitution> Substitutions,
        string Fingerprint);

    private sealed record FontPackManifest(
        string Id,
        IReadOnlyList<string> Coverage,
        IReadOnlyList<FontPackFont> Fonts,
        IReadOnlyList<FontPackSubstitution> Substitutions);

    private sealed record FontPackFont(string Family);

    private sealed record FontPackSubstitution(string Source, string Target, string Impact);
}
