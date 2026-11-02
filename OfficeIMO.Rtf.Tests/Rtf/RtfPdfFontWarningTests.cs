using OfficeIMO.Drawing;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.TestAssets;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfPdfFontWarningTests {
    private const string Source = "Warning Source";
    private const string Target = "Warning Target";

    [Theory]
    [InlineData("comma")]
    [InlineData("semicolon")]
    public void UnavailableLiteralFamilyDoesNotBorrowCompatibleMappingFromAComponent(string scenario) {
        string unavailableFamily = scenario == "comma" ? "Unavailable, Courier" : "Unavailable; Courier";
        RtfDocument document = CreateDocument(Target, unavailableFamily);
        var caller = new PdfCore.PdfCreateOptions().RegisterNamedFontFamily(CreateFamily(Target, 500));
        caller.RegisterFontFamilySubstitution("Courier", Target, PdfCore.PdfFontFamilySubstitutionImpact.Compatible);
        PdfCore.PdfCreateOptions originalOptions = caller.Clone();
        byte[] originalSource = document.ToBytes();

        PdfCore.PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(ConvertOptions(caller));
        AssertPaintedFamilyAndAdvance(result.ToBytes(), Target, 5D);

        PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings,
            warning => warning.Code == "FontFamilySubstituted" &&
                warning.Details.TryGetValue("fontFamily", out string? family) && family == unavailableFamily);
        Assert.Equal("Font/" + unavailableFamily, warning.Source);
        Assert.Equal(Target, warning.Details["resolvedFontFamily"]);
        Assert.Equal(PdfCore.PdfConversionWarningSeverity.Warning, warning.Severity);
        Assert.Equal(OfficeConversionLossKind.Approximation, warning.LossKind);
        Assert.False(warning.Details.ContainsKey("plannedSubstitution"));
        Assert.True(result.HasLoss);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Equal(originalSource, document.ToBytes());
        AssertCallerOptionsUnchanged(originalOptions, caller);
    }

    [Theory]
    [InlineData(PdfCore.PdfFontFamilySubstitutionImpact.Compatible)]
    [InlineData(PdfCore.PdfFontFamilySubstitutionImpact.LayoutSensitive)]
    public void ProfileOwnedFamilyDoesNotReportAnInactiveSubstitution(PdfCore.PdfFontFamilySubstitutionImpact impact) {
        RtfDocument document = CreateDocument(Source, Source);
        var fonts = new OfficeFontFaceCollection().Add(Source, CreateFont(300));
        var caller = new PdfCore.PdfCreateOptions().UseRenderingProfile(new OfficeRenderingProfile("warning-source", fonts));
        caller.RegisterNamedFontFamily(CreateFamily(Target, 800));
        caller.RegisterFontFamilySubstitution(Source, Target, impact);
        PdfCore.PdfCreateOptions originalOptions = caller.Clone();
        byte[] originalSource = document.ToBytes();

        PdfCore.PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(ConvertOptions(caller));
        AssertPaintedFamilyAndAdvance(result.ToBytes(), Source, 3D);

        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "FontFamilySubstituted");
        Assert.False(result.HasLoss);
        Assert.Same(result.Value, result.RequireNoLoss());
        Assert.Equal(originalSource, document.ToBytes());
        AssertCallerOptionsUnchanged(originalOptions, caller);
    }

    [Fact]
    public void NonProfileNamedSourceKeepsConfiguredSubstitutionPriority() {
        RtfDocument document = CreateDocument(Source, Source);
        var caller = new PdfCore.PdfCreateOptions()
            .RegisterNamedFontFamily(CreateFamily(Source, 300))
            .RegisterNamedFontFamily(CreateFamily(Target, 800));
        caller.RegisterFontFamilySubstitution(Source, Target, PdfCore.PdfFontFamilySubstitutionImpact.Compatible);
        PdfCore.PdfCreateOptions originalOptions = caller.Clone();
        byte[] originalSource = document.ToBytes();

        PdfCore.PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(ConvertOptions(caller));
        AssertPaintedFamilyAndAdvance(result.ToBytes(), Target, 8D);

        PdfCore.PdfConversionWarning warning = Assert.Single(result.Warnings,
            warning => warning.Code == "FontFamilySubstituted");
        Assert.Equal(Source, warning.Details["fontFamily"]);
        Assert.Equal(Target, warning.Details["resolvedFontFamily"]);
        Assert.Equal(bool.TrueString, warning.Details["plannedSubstitution"]);
        Assert.Equal(PdfCore.PdfConversionWarningSeverity.Information, warning.Severity);
        Assert.Equal(OfficeConversionLossKind.None, warning.LossKind);
        Assert.False(result.HasLoss);
        Assert.Same(result.Value, result.RequireNoLoss());
        Assert.Equal(originalSource, document.ToBytes());
        AssertCallerOptionsUnchanged(originalOptions, caller);
    }

    private static RtfDocument CreateDocument(string defaultFamily, string runFamily) {
        RtfDocument document = RtfDocument.Create();
        document.Settings.SetDefaultFont(document.AddFont(defaultFamily));
        RtfRun run = document.AddParagraph().AddText("AB");
        run.FontId = document.AddFont(runFamily);
        run.FontSize = 10;
        return document;
    }

    private static byte[] CreateFont(ushort advance) => ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(
        1000, advance, Enumerable.Range(32, 95).ToArray());

    private static PdfCore.PdfEmbeddedFontFamily CreateFamily(string family, ushort advance) =>
        new PdfCore.PdfEmbeddedFontFamily(family, CreateFont(advance));

    private static RtfToPdfOptions ConvertOptions(PdfCore.PdfCreateOptions caller) => new RtfToPdfOptions {
        PdfOptions = caller,
        ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
        TextFallbacks = PdfCore.PdfTextFallbackFeatures.None
    };

    private static void AssertPaintedFamilyAndAdvance(byte[] bytes, string family, double expectedAdvance) {
        using PdfPigDocument actual = PdfPigDocument.Open(bytes);
        Assert.Equal(1, actual.NumberOfPages);
        var letters = actual.GetPage(1).Letters;
        Assert.Equal("AB", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(2, letters.Count);
        Assert.All(letters, letter => {
            Assert.Contains(family.Replace(" ", string.Empty), letter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(10D, letter.PointSize, 3);
        });
        Assert.Equal(expectedAdvance, letters[1].StartBaseLine.X - letters[0].StartBaseLine.X, 3);
    }

    private static void AssertCallerOptionsUnchanged(PdfCore.PdfCreateOptions expected, PdfCore.PdfCreateOptions actual) {
        Assert.Equal(expected.NamedFontFamilies.Keys.OrderBy(key => key), actual.NamedFontFamilies.Keys.OrderBy(key => key));
        foreach (var family in expected.NamedFontFamilies) {
            Assert.Equal(family.Value.Regular, actual.NamedFontFamilies[family.Key].Regular);
        }
        Assert.Equal(expected.EmbeddedFonts.Keys.OrderBy(key => key), actual.EmbeddedFonts.Keys.OrderBy(key => key));
        foreach (var font in expected.EmbeddedFonts) {
            Assert.Equal(font.Value.FontName, actual.EmbeddedFonts[font.Key].FontName);
            Assert.Equal(font.Value.Data, actual.EmbeddedFonts[font.Key].Data);
        }
        Assert.Equal(expected.FontFamilySubstitutions.Select(substitution =>
                (substitution.SourceFontFamily, substitution.TargetFontFamily, substitution.Impact)),
            actual.FontFamilySubstitutions.Select(substitution =>
                (substitution.SourceFontFamily, substitution.TargetFontFamily, substitution.Impact)));
    }
}
