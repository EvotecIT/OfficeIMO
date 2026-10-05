using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;
using PdfDocument = OfficeIMO.Pdf.PdfDocument;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Pdf.Tests;

public class PdfParagraphBuilderFeatureTests {
    [Theory]
    [InlineData("text")]
    [InlineData("bold")]
    [InlineData("link")]
    [InlineData("fallback")]
    public void ParagraphBuilderAppliesFeatureOverridesToSubsequentText(string kind) {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'B', -200, includeSpace: true);
        var family = new PdfEmbeddedFontFamily("Builder features", font, bold: font);
        var options = new PdfOptions { DefaultFontSize = 12 };
        options.RegisterNamedFontFamily(family);
        var fallback = new PdfEmbeddedFontFallbackSet(new[] { new PdfEmbeddedFontFallbackCandidate(family.FamilyName, font) });
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(options).Paragraph(builder => {
            builder.FontFamily(family.FamilyName).FeatureSettings(OfficeTextFeatureSettings.Default.With("kern", 1));
            switch (kind) {
                case "bold": builder.Bold("AB"); break;
                case "link": builder.Link("AB", "https://example.com/"); break;
                case "fallback": builder.FallbackText(fallback, "AB"); break;
                default: builder.Text("AB"); break;
            }
        }).ToBytes());
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal(3.6D, Assert.Single(letters, letter => letter.Value == "B").StartBaseLine.X -
            Assert.Single(letters, letter => letter.Value == "A").StartBaseLine.X, 3);
    }

    [Fact]
    public void ParagraphBuilderKeepsPreparedFeaturesAndCanResetCurrentOverrides() {
        var on = OfficeTextFeatureSettings.Default.With("kern", 1);
        var off = OfficeTextFeatureSettings.Default.With("kern", 0);
        var prepared = new PdfTextRun("Prepared").WithFeatureSettings(off);
        var builder = new PdfParagraphBuilder(PdfAlign.Left, null);
        builder.FeatureSettings(on).Text("Current").Runs(new[] { prepared }).ResetFeatureSettings().Text("Reset");
        var runs = builder.Build().Runs;
        Assert.Equal(on, runs[0].FeatureSettings);
        Assert.Equal(off, runs[1].FeatureSettings);
        Assert.Equal(OfficeTextFeatureSettings.Default, runs[2].FeatureSettings);
    }
}
