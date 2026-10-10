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
    /// <summary>
    /// SHA-256 of the line-ending-normalized font-pack.json. The manifest pins every font file's SHA-256, so a verified
    /// manifest plus verified files is exactly the content <see cref="ExpectedFontPackFingerprint"/> covers. That lets
    /// the fallback fonts arrive later without changing the pack identity reported in manifests.
    /// </summary>
    internal const string ExpectedManifestSha256 = "25f40e9dba116e36d659d4d8cbe22b6584e351f7d66ecb43480a86eb2b23276a";

    private static readonly Lazy<FontPackData> Data = new(LoadFontPack, isThreadSafe: true);
    private static readonly Lock FallbackGate = new();
    private static FallbackFontData? _fallback;

    internal static string FontPackId => Data.Value.Id;
    internal static string FontPackFingerprint => Data.Value.Fingerprint;
    internal static IReadOnlyList<string> FontCoverage => Data.Value.Coverage;
    internal static IReadOnlyList<PdfFontFamilySubstitution> FontFamilySubstitutions =>
        Data.Value.Substitutions;
    internal static byte[] CarlitoRegularBytes => Data.Value.CarlitoRegular;

    /// <summary>The Japanese, Arabic and symbol fallbacks, once their assembly is loaded; otherwise null.</summary>
    private static FallbackFontData? Fallback {
        get {
            if (Volatile.Read(ref _fallback) is { } loaded) return loaded;
            System.Reflection.Assembly? assembly = BrowserFallbackFonts.FindAssembly();
            if (assembly == null) return null;
            lock (FallbackGate) {
                return _fallback ??= LoadFallback(Data.Value, assembly);
            }
        }
    }

    internal static OfficeFontFaceCollection CreateDrawingFonts() => CreateLayoutFonts(Data.Value, Fallback);

    internal static PdfOptions CreateOptions(BrowserPdfProfile profile) {
        ArgumentNullException.ThrowIfNull(profile);
        FontPackData data = Data.Value;
        FallbackFontData? fallback = Fallback;
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
        if (fallback != null) options.RegisterEmbeddedFontFallbacks(fallback.PdfFontFallbacks);
        foreach (PdfFontFamilySubstitution substitution in data.Substitutions) {
            // Without the fallback assembly only Carlito targets exist. BrowserFallbackFonts loads it for any
            // document that names a font mapped elsewhere, so skipping here never changes a document that uses one.
            if (fallback == null && !string.Equals(substitution.TargetFontFamily, DefaultFontFamily, StringComparison.OrdinalIgnoreCase)) continue;
            options.RegisterFontFamilySubstitution(
                substitution.SourceFontFamily,
                substitution.TargetFontFamily,
                substitution.Impact);
        }

        return options;
    }

    internal static HtmlToPdfOptions CreateHtmlOptions(BrowserPdfProfile profile) {
        FontPackData data = Data.Value;
        FallbackFontData? fallback = Fallback;
        return new HtmlToPdfOptions {
            DefaultFontFamily = fallback != null ? DefaultLayoutFontFamilies : DefaultFontFamily,
            Fonts = CreateLayoutFonts(data, fallback),
            PdfOptions = CreateOptions(profile),
            FontFamily = data.DefaultPdfFontFamily,
            TextShapingMode = PdfTextShapingMode.LatinLigatures,
            TextShapingProvider = OfficeHarfBuzzTextShapingProvider.Instance,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        };
    }

    private static OfficeFontFaceCollection CreateLayoutFonts(FontPackData data, FallbackFontData? fallback) {
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
        if (fallback == null) return fonts;
        fonts.Add(SymbolFallbackFontFamily, fallback.NotoSansSymbols);
        foreach (string alias in PortableSymbolAliases) {
            fonts.AddAlias(alias, SymbolFallbackFontFamily);
        }
        return fonts
            .Add(JapaneseFallbackFontFamily, fallback.NotoSansJapaneseCommon)
            .Add(ArabicFallbackFontFamily, fallback.NotoSansArabic)
            .AddFallbackFamily(JapaneseFallbackFontFamily)
            .AddFallbackFamily(ArabicFallbackFontFamily)
            .AddFallbackFamily(SymbolFallbackFontFamily);
    }

    private static FontPackData LoadFontPack() {
        System.Reflection.Assembly assembly = typeof(OfficeIMO.Web.Fonts.BrowserFontResources).Assembly;
        byte[] manifestBytes = ReadResource(assembly, "font-pack.json");
        byte[] normalizedManifestBytes = NormalizeManifest(manifestBytes);
        string manifestHash = Sha256(normalizedManifestBytes);
        if (!string.Equals(manifestHash, ExpectedManifestSha256, StringComparison.Ordinal)) {
            throw new InvalidOperationException(
                $"The embedded browser PDF font pack manifest '{manifestHash}' does not match the pinned profile '{ExpectedManifestSha256}'.");
        }
        FontPackManifest manifest = JsonSerializer.Deserialize(
            manifestBytes,
            BrowserFontPackJsonContext.Default.FontPackManifest)
            ?? throw new InvalidOperationException("The embedded browser PDF font pack manifest is invalid.");
        IReadOnlyDictionary<string, string> hashes = DeclaredHashes(manifest);
        byte[] carlitoRegular = ReadVerified(assembly, "Carlito-Regular.ttf", hashes);
        byte[] carlitoBold = ReadVerified(assembly, "Carlito-Bold.ttf", hashes);
        byte[] carlitoItalic = ReadVerified(assembly, "Carlito-Italic.ttf", hashes);
        byte[] carlitoBoldItalic = ReadVerified(assembly, "Carlito-BoldItalic.ttf", hashes);

        return new FontPackData(
            manifest.Id,
            ValidateCoverage(manifest.Coverage),
            hashes,
            carlitoRegular,
            carlitoBold,
            carlitoItalic,
            carlitoBoldItalic,
            new PdfEmbeddedFontFamily(
                DefaultFontFamily,
                carlitoRegular,
                carlitoBold,
                carlitoItalic,
                carlitoBoldItalic),
            ValidateManifest(manifest),
            ExpectedFontPackFingerprint);
    }

    private static FallbackFontData LoadFallback(FontPackData data, System.Reflection.Assembly assembly) {
        byte[] japanese = ReadVerified(assembly, "NotoSansJP-OfficeIMO-Common.ttf", data.FileHashes);
        byte[] arabic = ReadVerified(assembly, "NotoSansArabic-Regular.ttf", data.FileHashes);
        byte[] symbols = ReadVerified(assembly, "NotoSansSymbols2-Regular.ttf", data.FileHashes);
        return new FallbackFontData(
            japanese,
            arabic,
            symbols,
            new PdfEmbeddedFontFallbackSet([
                new PdfEmbeddedFontFallbackCandidate(JapaneseFallbackFontFamily, japanese),
                new PdfEmbeddedFontFallbackCandidate(ArabicFallbackFontFamily, arabic),
                new PdfEmbeddedFontFallbackCandidate(SymbolFallbackFontFamily, symbols)
            ]));
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
        byte[] CarlitoBold,
        byte[] CarlitoItalic,
        byte[] CarlitoBoldItalic,
        PdfEmbeddedFontFamily DefaultPdfFontFamily,
        IReadOnlyList<PdfFontFamilySubstitution> Substitutions,
        string Fingerprint);

    private sealed record FallbackFontData(
        byte[] NotoSansJapaneseCommon,
        byte[] NotoSansArabic,
        byte[] NotoSansSymbols,
        PdfEmbeddedFontFallbackSet PdfFontFallbacks);
}
