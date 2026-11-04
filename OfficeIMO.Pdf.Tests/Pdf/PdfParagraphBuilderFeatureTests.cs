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
        var options = new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(12) };
        options.RegisterNamedFontFamily(family);
        var fallback = new PdfEmbeddedFontFallbackSet(new[] { new PdfEmbeddedFontFallbackCandidate(family.FamilyName, font) });
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(options).Paragraph(builder => {
            builder.RunDefaults.FontFamily = family.FamilyName; builder.RunDefaults.StandardFont = null; builder.RunDefaults.FeatureSettings = OfficeTextFeatureSettings.Default.With("kern", 1);
            switch (kind) {
                case "bold": builder.AddText("AB").SetBold(); break;
                case "link": builder.AddLink("AB", "https://example.com/"); break;
                case "fallback": builder.AddFallbackText(fallback, "AB"); break;
                default: builder.AddText("AB"); break;
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
        PdfParagraph paragraph = PdfDocument.Create().AddParagraph();
        paragraph.RunDefaults.FeatureSettings = on;
        paragraph.AddText("Current");
        paragraph.AddRuns(new[] { prepared });
        paragraph.RunDefaults.FeatureSettings = OfficeTextFeatureSettings.Default;
        paragraph.AddText("Reset");
        var runs = paragraph.Runs;
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
        var options = new PdfCreateOptions { PageSetup = {
            Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(240)),
            Margins = { Right = global::OfficeIMO.OfficeLength.Points(30) }
        } };
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Page features", font));
        var features = OfficeTextFeatureSettings.Default.With("kern", kern);
        var pair = new PdfTextRun("AB", fontSize: global::OfficeIMO.OfficeLength.Points(24), fontFamily: "Page features").WithFeatureSettings(features);
        var following = new PdfTextRun("A", fontSize: global::OfficeIMO.OfficeLength.Points(24), fontFamily: "Page features");
        var document = PdfDocument.Create(options);
        if (header) document.Header(builder => builder.Alignment(OfficeAlignment.Right).Text(text => text.Run(pair).Run(following)));
        else document.Footer(builder => builder.Alignment(OfficeAlignment.Right).Text(text => text.Run(pair).Run(following)));
        using var pdf = PdfPigDocument.Open(document.Paragraph(builder => builder.AddText("content")).ToBytes());
        var letters = pdf.GetPage(1).Letters;
        // Read only the authored page-text letters; body text has no A or B.
        var positions = new System.Collections.Generic.List<double>();
        foreach (var letter in letters) if (letter.Value == "A" || letter.Value == "B") positions.Add(letter.StartBaseLine.X);
        Assert.Equal(3, positions.Count);
        double pairWidth = kern == 1 ? 19.2D : 24D;
        Assert.Equal(kern == 1 ? 7.2D : 12D, positions[1] - positions[0], 3);
        Assert.Equal(pairWidth, positions[2] - positions[0], 3);
        Assert.Equal(options.PageSetup.Size.Width.ToPoints() - options.PageSetup.Margins.Right.ToPoints() - pairWidth - 12D, positions[0], 3);
    }
}
