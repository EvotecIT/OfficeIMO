using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCrossRunKerningTests {
    [Fact]
    public void SubstitutionContinuationGlyphsRetainTheirSourceScriptForKerning() {
        PdfTextRun Run(string text) => new PdfTextRun(text, fontSize: 20D, fontFamily: ManagedTextShapingTestAssets.FamilyName)
            .WithFeatureSettings(OfficeTextFeatureSettings.Default.With("kern", 1));
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithScriptSpecificContinuationKerning();
        var shaped = PdfTrueTypeFontProgram.Parse(font).ShapeText("A", PdfTextShapingOptions.ForRendering(
            ManagedTextShapingTestAssets.FamilyName, PdfTextShapingMode.OpenTypeLigatures,
            featureSettings: OfficeTextFeatureSettings.Default.With("kern", 1)));
        Assert.Equal(3, shaped.Glyphs.Count);
        Assert.True(shaped.Glyphs[1].HasPositioning);
        using var joined = PdfPigDocument.Open(Create(100D, PdfAlign.Left, new[] { Run("AA") }, font));
        using var split = PdfPigDocument.Open(Create(100D, PdfAlign.Left, new[] { Run("A"), Run("A") }, font));
        Assert.Equal(joined.GetPage(1).Letters.Select(letter => letter.StartBaseLine), split.GetPage(1).Letters.Select(letter => letter.StartBaseLine));
    }

    [Fact]
    public void EmptyFormattingRunsDoNotInterruptKernedWordContinuations() {
        PdfTextRun Run(string text) => new PdfTextRun(text, fontSize: 20D, fontFamily: ManagedTextShapingTestAssets.FamilyName)
            .WithFeatureSettings(OfficeTextFeatureSettings.Default.With("kern", 1));
        Letter[] joined = Render(29D, PdfAlign.Left, Run("A AVA"));
        Letter[] split = Render(29D, PdfAlign.Left, Run("A A"), Run(""), Run("V"), Run(""), Run("A"));
        Assert.Equal(joined.Select(letter => letter.StartBaseLine), split.Select(letter => letter.StartBaseLine));
    }

    [Fact]
    public void SubstitutionContinuationGlyphsRetainTheirNominalAdvancesAcrossRuns() {
        PdfTextRun Run(string text) => new PdfTextRun(text, fontSize: 20D, fontFamily: ManagedTextShapingTestAssets.FamilyName)
            .WithFeatureSettings(OfficeTextFeatureSettings.Default.With("kern", 1));
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithMultipleSubstitution('A', scriptTag: "latn", featureTag: "liga");
        using var joined = PdfPigDocument.Open(Create(100D, PdfAlign.Left, new[] { Run("AA") }, font));
        using var split = PdfPigDocument.Open(Create(100D, PdfAlign.Left, new[] { Run("A"), Run("A") }, font));
        Assert.Equal(joined.GetPage(1).Letters.Select(letter => letter.StartBaseLine), split.GetPage(1).Letters.Select(letter => letter.StartBaseLine));
    }

    [Theory]
    [InlineData(29D)]
    [InlineData(39D)]
    public void OversizedWordsFitKernedChunksAcrossPaintBoundaries(double width) {
        PdfTextRun Run(string text) => new PdfTextRun(text, fontSize: 20D, fontFamily: ManagedTextShapingTestAssets.FamilyName)
            .WithFeatureSettings(OfficeTextFeatureSettings.Default.With("kern", 1));
        Letter[] joined = Render(width, PdfAlign.Left, Run("AVAVA"));
        Letter[] split = Render(width, PdfAlign.Left, "AVAVA".Select(character => Run(character.ToString())).ToArray());
        Assert.Equal(joined.Select(letter => letter.StartBaseLine), split.Select(letter => letter.StartBaseLine));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ConsecutiveGposPairsRetainTheirLookupConsumptionAcrossRuns(bool secondRunHasTwoGlyphs) {
        PdfTextRun Run(string text) => new PdfTextRun(text, fontSize: 20D, fontFamily: ManagedTextShapingTestAssets.FamilyName)
            .WithFeatureSettings(OfficeTextFeatureSettings.Default.With("kern", 1));
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithSequencedPairPositioning();
        using var joined = PdfPigDocument.Open(Create(100D, PdfAlign.Left, new[] { Run("ABC") }, font));
        using var split = PdfPigDocument.Open(Create(100D, PdfAlign.Left,
            secondRunHasTwoGlyphs ? new[] { Run("A"), Run("BC") } : new[] { Run("A"), Run("B"), Run("C") }, font));
        Assert.Equal(joined.GetPage(1).Letters.Select(letter => letter.StartBaseLine), split.GetPage(1).Letters.Select(letter => letter.StartBaseLine));
    }

    [Theory]
    [InlineData(29D, 100D, 0D, PdfAlign.Left)]
    [InlineData(25D, 100D, 0D, PdfAlign.Left)]
    [InlineData(49D, 100D, 0D, PdfAlign.Right)]
    [InlineData(25D, 50D, 1D, PdfAlign.Center)]
    public void SplitWordsUseTheSameLineBreaksAndAlignmentAsJoinedWords(double width, double scale, double spacing, PdfAlign align) {
        PdfTextRun Run(string text, bool blue = false) => new PdfTextRun(text, fontSize: 20D,
            fontFamily: ManagedTextShapingTestAssets.FamilyName, color: blue ? PdfColor.FromRgb(0, 0, 255) : PdfColor.FromRgb(255, 0, 0))
            .WithFeatureSettings(OfficeTextFeatureSettings.Default.With("kern", 1))
            .WithHorizontalTextScaling(scale).WithCharacterSpacing(spacing);
        Letter[] joined = Render(width, align, Run("A AVA"));
        Letter[] split = Render(width, align, Run("A A"), Run("V", blue: true), Run("A"));
        Assert.Equal(joined.Length, split.Length);
        for (int index = 0; index < joined.Length; index++) {
            Assert.Equal(joined[index].Value, split[index].Value);
            Assert.Equal(joined[index].StartBaseLine.X, split[index].StartBaseLine.X, 3);
            Assert.Equal(joined[index].StartBaseLine.Y, split[index].StartBaseLine.Y, 3);
        }
        Assert.Equal((0D, 0D, 1D), Assert.Single(split, letter => letter.Value == "V").Color.ToRGBValues());
    }

    [Theory]
    [InlineData("size")]
    [InlineData("font")]
    [InlineData("disabled")]
    [InlineData("separator")]
    [InlineData("break")]
    public void IncompatibleRunsDoNotAcquireBoundaryKerning(string boundary) {
        PdfTextRun Run(string text, string font, double size, int kern) => new PdfTextRun(text, fontSize: size, fontFamily: font)
            .WithFeatureSettings(OfficeTextFeatureSettings.Default.With("kern", kern));
        string family = ManagedTextShapingTestAssets.FamilyName;
        PdfTextRun left = Run(boundary == "separator" ? "A " : boundary == "break" ? "A\n" : "A", family, 20D, 1);
        PdfTextRun right = Run("V", boundary == "font" ? "OtherFace" : family, boundary == "size" ? 21D : 20D, boundary == "disabled" ? 0 : 1);
        Letter[] actual = Render(100D, PdfAlign.Left, left, right);
        var a = Assert.Single(actual, letter => letter.Value == "A");
        var v = Assert.Single(actual, letter => letter.Value == "V");
        Assert.Equal(20D, a.PointSize, 3);
        if (boundary == "break") {
            Assert.True(v.StartBaseLine.Y < a.StartBaseLine.Y);
            Assert.Equal(a.StartBaseLine.X, v.StartBaseLine.X, 3);
        } else {
            Assert.Equal(boundary == "separator" ? 20D : 10D, v.StartBaseLine.X - a.StartBaseLine.X, 3);
        }
    }

    [Fact]
    public void PairPositioningPreservesSeparateLinksAndDecoratedRuns() {
        PdfTextRun Run(string text, string uri, bool underline) => new PdfTextRun(text,
            fontSize: 20D, fontFamily: ManagedTextShapingTestAssets.FamilyName,
            linkUri: uri, underline: underline, strike: !underline)
            .WithFeatureSettings(OfficeTextFeatureSettings.Default.With("kern", 1));
        var runs = new[] { Run("A", "https://example.test/a", true), Run("V", "https://example.test/v", false), Run("A", "https://example.test/last", true) };
        using var pdf = PdfPigDocument.Open(Create(100D, PdfAlign.Left, runs));
        var page = pdf.GetPage(1);
        Assert.Equal("AVA", page.Text);
        Assert.Equal(3, page.GetHyperlinks().Count());
        Letter[] joined = Render(100D, PdfAlign.Left, Run("AVA", "https://example.test/all", true));
        Assert.Equal(joined.Select(letter => letter.StartBaseLine.X), page.Letters.Select(letter => letter.StartBaseLine.X));
    }

    private static Letter[] Render(double width, PdfAlign align, params PdfTextRun[] runs) {
        using var pdf = PdfPigDocument.Open(Create(width, align, runs));
        return pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
    }

    private static byte[] Create(double width, PdfAlign align, params PdfTextRun[] runs) {
        return Create(width, align, runs, ManagedTextShapingTestAssets.CreateFontWithPairPositioning('A', 'V', includeSpace: true));
    }

    private static byte[] Create(double width, PdfAlign align, PdfTextRun[] runs, byte[] font) {
        var options = new PdfOptions {
            DefaultFontSize = 20D, PageWidth = width + 40D, PageHeight = 250D,
            TextShapingMode = PdfTextShapingMode.OpenTypeLigatures,
            MarginLeft = 20D, MarginRight = 20D, MarginTop = 20D, MarginBottom = 20D
        };
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName, font));
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("OtherFace", font));
        return PdfDocument.Create(options).Paragraph(paragraph => paragraph.Runs(runs), align: align).ToBytes();
    }
}
