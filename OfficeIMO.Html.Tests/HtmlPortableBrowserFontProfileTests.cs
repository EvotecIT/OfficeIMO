using OfficeIMO.Drawing;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPortableBrowserFontProfileTests {
    [Fact]
    public void DefaultOnlyProfileDoesNotRequireFallbackAssets() {
        IReadOnlyDictionary<string, byte[]> assets = CreateAssets();
        var reads = new HashSet<string>(StringComparer.Ordinal);
        HtmlPortableBrowserFontProfile profile = HtmlPortableBrowserFontProfile.CreateDefaultOnly("portable-test", name => {
            Assert.StartsWith("Carlito-", name);
            reads.Add(name);
            return assets[name];
        });

        Assert.Equal(4, reads.Count);
        HtmlToPdfOptions options = profile.CreateHtmlOptions();
        Assert.Equal(HtmlPortableBrowserFontProfile.DefaultFontFamily, options.DefaultFontFamily);
        Assert.Equal(
            new[] { OfficeFontStyle.Regular, OfficeFontStyle.Bold, OfficeFontStyle.Italic, OfficeFontStyle.Bold | OfficeFontStyle.Italic },
            options.Fonts.Faces.Where(face => face.FamilyName == HtmlPortableBrowserFontProfile.DefaultFontFamily)
                .Select(face => face.Style).ToArray());
        Assert.Empty(options.Fonts.FallbackFamilies);
        Assert.Null(options.PdfOptions.EmbeddedFontFallbacks);
        Assert.Single(options.PdfOptions.NamedFontFamilies);
        Assert.Contains(options.PdfOptions.FontFamilySubstitutions, item => item.SourceFontFamily == "Calibri");
        Assert.DoesNotContain(options.PdfOptions.FontFamilySubstitutions, item => item.SourceFontFamily == "Symbol");
        Assert.Contains(profile.FontFamilySubstitutions, item => item.SourceFontFamily == "Symbol");
        Assert.True(options.Fonts.TryMeasureText("A", 16D, "Arial", OfficeFontStyle.Regular, out double width));
        Assert.True(width > 0D);
    }

    [Fact]
    public void FallbackUpgradeLeavesDefaultSnapshotUnchangedAndReusesPdfFamily() {
        IReadOnlyDictionary<string, byte[]> assets = CreateAssets();
        HtmlPortableBrowserFontProfile original = HtmlPortableBrowserFontProfile.CreateDefaultOnly("portable-test", name => assets[name]);
        string originalFingerprint = original.Fingerprint;
        HtmlToPdfOptions originalOptions = original.CreateHtmlOptions();
        var reads = new HashSet<string>(StringComparer.Ordinal);
        HtmlPortableBrowserFontProfile upgraded = original.WithFallbackFonts(name => {
            Assert.StartsWith("NotoSans", name);
            reads.Add(name);
            return assets[name];
        });

        Assert.Equal(3, reads.Count);
        Assert.Equal(originalFingerprint, original.Fingerprint);
        Assert.NotEqual(originalFingerprint, upgraded.Fingerprint);
        Assert.Empty(original.CreateDrawingFonts().FallbackFamilies);
        Assert.Null(original.CreatePdfOptions().EmbeddedFontFallbacks);
        Assert.Equal(HtmlPortableBrowserFontProfile.DefaultFontFamily, original.CreateHtmlOptions().DefaultFontFamily);

        HtmlToPdfOptions upgradedOptions = upgraded.CreateHtmlOptions();
        Assert.Same(originalOptions.FontFamily, upgradedOptions.FontFamily);
        Assert.Same(originalOptions.FontFamily,
            upgradedOptions.PdfOptions.NamedFontFamilies[HtmlPortableBrowserFontProfile.DefaultFontFamily]);
        AssertFullFallbacks(upgradedOptions);

        OfficeFontFaceCollection separateDrawingFonts = upgraded.CreateDrawingFonts();
        separateDrawingFonts.AddAlias("Caller-local alias", HtmlPortableBrowserFontProfile.DefaultFontFamily);
        Assert.DoesNotContain(upgraded.CreateDrawingFonts().Faces, face => face.FamilyName == "Caller-local alias");
    }

    [Fact]
    public void EagerProfileRetainsFullLayoutAndPdfFallbackContract() {
        IReadOnlyDictionary<string, byte[]> assets = CreateAssets();
        var reads = new HashSet<string>(StringComparer.Ordinal);
        HtmlPortableBrowserFontProfile eager = HtmlPortableBrowserFontProfile.Create("portable-test", name => {
            reads.Add(name);
            return assets[name];
        });
        HtmlPortableBrowserFontProfile staged = HtmlPortableBrowserFontProfile.CreateDefaultOnly("portable-test", name => assets[name])
            .WithFallbackFonts(name => assets[name]);

        Assert.Equal(assets.Keys.OrderBy(name => name), reads.OrderBy(name => name));
        Assert.Equal(staged.Fingerprint, eager.Fingerprint);
        HtmlToPdfOptions options = eager.CreateHtmlOptions();
        AssertFullFallbacks(options);
        Assert.True(options.Fonts.TryMeasureText("A", 16D, "Times New Roman", OfficeFontStyle.Regular, out double width));
        Assert.True(width > 0D);
        Assert.True(options.Fonts.TryMeasureText("⌚", 16D, "ZapfDingbats", OfficeFontStyle.Regular, out double symbolWidth));
        Assert.True(symbolWidth > 0D);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void MissingOrEmptyAssetsAreRejectedWithoutChangingAnExistingSnapshot(bool upgrade, bool empty) {
        IReadOnlyDictionary<string, byte[]> assets = CreateAssets();
        HtmlPortableBrowserFontProfile original = HtmlPortableBrowserFontProfile.CreateDefaultOnly("portable-test", name => assets[name]);
        string fingerprint = original.Fingerprint;
        string missingName = upgrade ? "NotoSansJP-OfficeIMO-Common.ttf" : "Carlito-Regular.ttf";
        byte[] Reader(string name) => name == missingName ? (empty ? Array.Empty<byte>() : null!) : assets[name];

        InvalidOperationException error = Assert.Throws<InvalidOperationException>(() => {
            if (upgrade) original.WithFallbackFonts(Reader);
            else HtmlPortableBrowserFontProfile.CreateDefaultOnly("portable-test", Reader);
        });

        Assert.Contains(missingName, error.Message);
        Assert.Equal(fingerprint, original.Fingerprint);
        Assert.Empty(original.CreateDrawingFonts().FallbackFamilies);
        AssertFullFallbacks(original.WithFallbackFonts(name => assets[name]).CreateHtmlOptions());
    }

    [Fact]
    public void FactoriesRejectMissingIdOrReaderAndPreserveReaderExceptions() {
        IReadOnlyDictionary<string, byte[]> assets = CreateAssets();
        HtmlPortableBrowserFontProfile original = HtmlPortableBrowserFontProfile.CreateDefaultOnly("portable-test", name => assets[name]);
        Assert.Throws<ArgumentException>(() => HtmlPortableBrowserFontProfile.CreateDefaultOnly(" ", name => assets[name]));
        Assert.Throws<ArgumentException>(() => HtmlPortableBrowserFontProfile.Create(" ", name => assets[name]));
        Assert.Throws<ArgumentNullException>(() => HtmlPortableBrowserFontProfile.CreateDefaultOnly("portable-test", null!));
        Assert.Throws<ArgumentNullException>(() => HtmlPortableBrowserFontProfile.Create("portable-test", null!));
        Assert.Throws<ArgumentNullException>(() => original.WithFallbackFonts(null!));

        var failure = new IOException("Caller asset loading failed.");
        Assert.Same(failure, Assert.Throws<IOException>(() =>
            HtmlPortableBrowserFontProfile.CreateDefaultOnly("portable-test", _ => throw failure)));
        Assert.Same(failure, Assert.Throws<IOException>(() => original.WithFallbackFonts(_ => throw failure)));
        Assert.Empty(original.CreateDrawingFonts().FallbackFamilies);
    }

    private static void AssertFullFallbacks(HtmlToPdfOptions options) {
        string[] families = [
            HtmlPortableBrowserFontProfile.JapaneseFallbackFontFamily,
            HtmlPortableBrowserFontProfile.ArabicFallbackFontFamily,
            HtmlPortableBrowserFontProfile.SymbolFallbackFontFamily
        ];
        Assert.Equal(HtmlPortableBrowserFontProfile.DefaultLayoutFontFamilies, options.DefaultFontFamily);
        Assert.Equal(families, options.Fonts.FallbackFamilies);
        Assert.Equal(families, Assert.IsType<PdfEmbeddedFontFallbackSet>(options.PdfOptions.EmbeddedFontFallbacks).FontFamilyNames);
        foreach (string family in families) Assert.Contains(family, options.PdfOptions.NamedFontFamilies.Keys);
        Assert.Contains(options.PdfOptions.FontFamilySubstitutions,
            item => item.SourceFontFamily == "Symbol" && item.TargetFontFamily == HtmlPortableBrowserFontProfile.SymbolFallbackFontFamily);
        Assert.Equal(
            new[] { HtmlPortableBrowserFontProfile.DefaultFontFamily }.Concat(families),
            options.Fonts.PlanFallbackRuns("A 日 ا ⌚", options.DefaultFontFamily, OfficeFontStyle.Regular)
                .Select(run => run.FamilyName).Distinct());
    }

    private static IReadOnlyDictionary<string, byte[]> CreateAssets() {
        byte[] latin = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(32, 'A');
        return new Dictionary<string, byte[]>(StringComparer.Ordinal) {
            ["Carlito-Regular.ttf"] = latin,
            ["Carlito-Bold.ttf"] = latin,
            ["Carlito-Italic.ttf"] = latin,
            ["Carlito-BoldItalic.ttf"] = latin,
            ["NotoSansJP-OfficeIMO-Common.ttf"] = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(32, '日'),
            ["NotoSansArabic-Regular.ttf"] = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(32, 'ا'),
            ["NotoSansSymbols2-Regular.ttf"] = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(32, '⌚')
        };
    }
}
