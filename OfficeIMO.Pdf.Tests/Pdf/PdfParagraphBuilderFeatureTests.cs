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

    [Theory]
    [InlineData(true, 1)]
    [InlineData(true, 0)]
    [InlineData(false, 1)]
    [InlineData(false, 0)]
    public void RichPageTextPreservesKerningForGlyphsFollowingRunsAndRightAlignment(bool header, int kern) {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'B', -200, includeSpace: true);
        var options = new PdfOptions { PageWidth = 300, PageHeight = 240, MarginRight = 30 };
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Page features", font));
        var features = OfficeTextFeatureSettings.Default.With("kern", kern);
        var pair = new PdfTextRun("AB", fontSize: 24, fontFamily: "Page features").WithFeatureSettings(features);
        var following = new PdfTextRun("A", fontSize: 24, fontFamily: "Page features");
        var document = PdfDocument.Create(options);
        if (header) document.Header(builder => builder.AlignRight().Text(text => text.Run(pair).Run(following)));
        else document.Footer(builder => builder.AlignRight().Text(text => text.Run(pair).Run(following)));
        using var pdf = PdfPigDocument.Open(document.Paragraph(builder => builder.Text("content")).ToBytes());
        var letters = pdf.GetPage(1).Letters;
        // Read only the authored page-text letters; body text has no A or B.
        var positions = new System.Collections.Generic.List<double>();
        foreach (var letter in letters) if (letter.Value == "A" || letter.Value == "B") positions.Add(letter.StartBaseLine.X);
        Assert.Equal(3, positions.Count);
        double pairWidth = kern == 1 ? 19.2D : 24D;
        Assert.Equal(kern == 1 ? 7.2D : 12D, positions[1] - positions[0], 3);
        Assert.Equal(pairWidth, positions[2] - positions[0], 3);
        Assert.Equal(options.PageWidth - options.MarginRight - pairWidth - 12D, positions[0], 3);
    }
}
