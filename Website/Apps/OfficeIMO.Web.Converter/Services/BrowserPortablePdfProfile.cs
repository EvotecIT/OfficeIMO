using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Drawing.HarfBuzz;
using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Models;

namespace OfficeIMO.Web.Converter.Services;

/// <summary>
/// Supplies the explicit, host-independent PDF font profile used by browser conversions.
/// </summary>
internal static class BrowserPortablePdfProfile {
    internal const string FontPackId = "officeimo-browser-compact-2026.07";
    internal const string DefaultFontFamily = "Carlito";
    internal const string ExpectedFontPackFingerprint = "99fe9605fae25324712287bc2212236771b67515ec77dab263a35fc48079e72f";

    private static readonly Lazy<FontPackData> Data = new(LoadFontPack, isThreadSafe: true);

    internal static string FontPackFingerprint => Data.Value.Fingerprint;

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

        return options;
    }

    private static FontPackData LoadFontPack() {
        byte[] carlitoRegular = ReadResource("Carlito-Regular.ttf");
        byte[] carlitoBold = ReadResource("Carlito-Bold.ttf");
        byte[] carlitoItalic = ReadResource("Carlito-Italic.ttf");
        byte[] carlitoBoldItalic = ReadResource("Carlito-BoldItalic.ttf");
        byte[] notoSansArabic = ReadResource("NotoSansArabic-Regular.ttf");
        byte[] notoSansSymbols = ReadResource("NotoSansSymbols2-Regular.ttf");

        var assets = new Dictionary<string, byte[]>(StringComparer.Ordinal) {
            ["Carlito-Bold.ttf"] = carlitoBold,
            ["Carlito-BoldItalic.ttf"] = carlitoBoldItalic,
            ["Carlito-Italic.ttf"] = carlitoItalic,
            ["Carlito-Regular.ttf"] = carlitoRegular,
            ["NotoSansArabic-Regular.ttf"] = notoSansArabic,
            ["NotoSansSymbols2-Regular.ttf"] = notoSansSymbols
        };

        string fingerprint = ComputeFingerprint(assets);
        if (!string.Equals(fingerprint, ExpectedFontPackFingerprint, StringComparison.Ordinal)) {
            throw new InvalidOperationException(
                $"The embedded browser PDF font pack fingerprint '{fingerprint}' does not match the pinned profile '{ExpectedFontPackFingerprint}'.");
        }

        return new FontPackData(
            carlitoRegular,
            carlitoBold,
            carlitoItalic,
            carlitoBoldItalic,
            notoSansArabic,
            notoSansSymbols,
            fingerprint);
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
        byte[] CarlitoRegular,
        byte[] CarlitoBold,
        byte[] CarlitoItalic,
        byte[] CarlitoBoldItalic,
        byte[] NotoSansArabic,
        byte[] NotoSansSymbols,
        string Fingerprint);
}
