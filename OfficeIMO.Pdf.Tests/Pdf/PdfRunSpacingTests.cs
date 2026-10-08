using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRunSpacingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EmbeddedLigatures_UseRenderedGlyphCountForSpacing(bool cff) {
        string? path = cff ? PdfComplianceTestFonts.FindBundledOpenTypeCffFont() : PdfComplianceTestFonts.FindBundledTrueTypeFont();
        Assert.NotNull(path);
        byte[] fontBytes = File.ReadAllBytes(path!);
        PdfTextShapingOptions shaping = PdfTextShapingOptions.ForRendering("SpacingTest", PdfTextShapingMode.OpenTypeLigatures);
        int glyphCount = cff
            ? PdfOpenTypeCffFontProgram.Parse(fontBytes, "SpacingTest").ShapeText("office", shaping).Glyphs.Count
            : PdfTrueTypeFontProgram.Parse(fontBytes, "SpacingTest").ShapeText("office", shaping).Glyphs.Count;
        Assert.InRange(glyphCount, 1, 5);
        byte[] Render(bool styled) {
            PdfOptions options = Options();
            options.TextShapingMode = PdfTextShapingMode.OpenTypeLigatures;
            options.EmbedStandardFont(PdfStandardFont.Helvetica, fontBytes, "SpacingTest");
            PdfTextRun word = new PdfTextRun("office", fontSize: 12D);
            if (styled) word = word.WithHorizontalTextScaling(50D).WithCharacterSpacing(1D);
            return PdfDocument.Create(options).Paragraph(paragraph => paragraph.Runs(new[] { word, PdfTextRun.Normal("X", fontSize: 12D) })).ToBytes();
        }
        using var natural = PdfPigDocument.Open(Render(false));
        using var styled = PdfPigDocument.Open(Render(true));
        var before = natural.GetPage(1).Letters;
        var after = styled.GetPage(1).Letters;
        double naturalAdvance = before.Single(letter => letter.Value == "X").StartBaseLine.X - before[0].StartBaseLine.X;
        double actualAdvance = after.Single(letter => letter.Value == "X").StartBaseLine.X - after[0].StartBaseLine.X;
        Assert.InRange(Math.Abs(actualAdvance - (naturalAdvance / 2D + glyphCount)), 0D, 0.02D);
        Assert.InRange(Math.Abs(before[0].BoundingBox.Height - after[0].BoundingBox.Height), 0D, 0.02D);
    }

    [Fact]
    public void ParagraphContinuation_PreservesSpacingOnEveryPage() {
        PdfOptions options = Options();
        options.PageWidth = 180D; options.PageHeight = 120D;
        options.MarginLeft = options.MarginRight = options.MarginTop = options.MarginBottom = 20D;
        PdfTextRun run = new PdfTextRun(string.Join("\n", Enumerable.Repeat("ABCD", 20)), fontSize: 12D)
            .WithHorizontalTextScaling(50D).WithCharacterSpacing(1D);
        byte[] bytes = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Runs(new[] { run })).ToBytes();
        Letter[] baseline = Letters(Create("paragraph", new PdfTextRun("ABCD", fontSize: 12D)));
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages > 1);
        int totalLines = 0;
        for (int page = 1; page <= pdf.NumberOfPages; page++) {
            var lines = pdf.GetPage(page).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1)).ToArray();
            Assert.NotEmpty(lines);
            foreach (var line in lines) {
                Letter[] actual = line.OrderBy(letter => letter.StartBaseLine.X).ToArray();
                Assert.Equal("ABCD", string.Concat(actual.Select(letter => letter.Value)));
                double expected = (baseline[3].StartBaseLine.X - baseline[0].StartBaseLine.X) / 2D + 3D;
                Assert.InRange(Math.Abs(actual[3].StartBaseLine.X - actual[0].StartBaseLine.X - expected), 0D, 0.02D);
                totalLines++;
            }
        }
        Assert.Equal(20, totalLines);
    }

    [Theory]
    [InlineData("paragraph", 50D, 1D)]
    [InlineData("paragraph", 200D, 1D)]
    [InlineData("paragraph", 50D, -0.5D)]
    [InlineData("paragraph", 200D, -0.5D)]
    [InlineData("table", 50D, 1D)]
    [InlineData("table", 200D, -0.5D)]
    [InlineData("header", 50D, 1D)]
    [InlineData("header", 200D, -0.5D)]
    public void RunSpacing_UsesPhysicalAdvancesAndPreservesGlyphHeight(string surface, double scaling, double spacing) {
        PdfTextRun natural = new PdfTextRun("ABCD", font: PdfStandardFont.Helvetica, fontSize: 12D);
        Letter[] baseline = Letters(Create(surface, natural));
        Letter[] actual = Letters(Create(surface, natural.WithHorizontalTextScaling(scaling).WithCharacterSpacing(spacing)));
        Assert.Equal("ABCD", string.Concat(actual.Select(letter => letter.Value)));
        Assert.Equal(4, baseline.Length);
        for (int index = 0; index < baseline.Length; index++) {
            double expected = (baseline[index].StartBaseLine.X - baseline[0].StartBaseLine.X) * scaling / 100D + index * spacing;
            double observed = actual[index].StartBaseLine.X - actual[0].StartBaseLine.X;
            Assert.InRange(Math.Abs(expected - observed), 0D, 0.02D);
            Assert.InRange(Math.Abs(actual[index].BoundingBox.Height - baseline[index].BoundingBox.Height), 0D, 0.02D);
        }
    }

    [Fact]
    public void ParagraphBuilder_RestoresNaturalTextAfterStyledRuns() {
        byte[] actualBytes = PdfDocument.Create(Options())
            .Paragraph(paragraph => paragraph.HorizontalTextScaling(50D).CharacterSpacing(1D).Text("AB")
                .HorizontalTextScaling(100D).CharacterSpacing(0D).Text("CD"))
            .ToBytes();
        Letter[] baseline = Letters(Create("paragraph", new PdfTextRun("ABCD", fontSize: 12D)));
        Letter[] actual = Letters(actualBytes);
        Assert.Equal("ABCD", string.Concat(actual.Select(letter => letter.Value)));
        double expectedC = (baseline[2].StartBaseLine.X - baseline[0].StartBaseLine.X) / 2D + 2D;
        Assert.InRange(Math.Abs(actual[2].StartBaseLine.X - actual[0].StartBaseLine.X - expectedC), 0D, 0.02D);
        Assert.InRange(Math.Abs(actual[3].StartBaseLine.X - actual[2].StartBaseLine.X -
            (baseline[3].StartBaseLine.X - baseline[2].StartBaseLine.X)), 0D, 0.02D);
    }

    [Fact]
    public void RunSpacingCopies_PreserveOtherSettingsAndLeaveSourceUnchanged() {
        PdfTextRun original = PdfTextRun.Link("ABCD", "https://example.test", fontSize: 12D)
            .WithHorizontalTextScaling(50D).WithCharacterSpacing(1D);
        PdfTextRun copy = original.WithFeatureSettings(OfficeTextFeatureSettings.Default);
        Assert.NotSame(original, copy);
        Assert.Equal(original.LinkUri, copy.LinkUri);
        Assert.Equal(original.UnderlineStyle, copy.UnderlineStyle);
        Assert.Equal(50D, copy.HorizontalTextScaling);
        Assert.Equal(1D, copy.CharacterSpacing);
        Letter[] baseline = Letters(Create("paragraph", new PdfTextRun("ABCD", fontSize: 12D)));
        Letter[] actual = Letters(Create("paragraph", copy));
        Assert.InRange(Math.Abs(actual[3].StartBaseLine.X - actual[0].StartBaseLine.X -
            ((baseline[3].StartBaseLine.X - baseline[0].StartBaseLine.X) / 2D + 3D)), 0D, 0.02D);
        PdfTextRun changed = copy.WithCharacterSpacing(2D).WithHorizontalTextScaling(75D);
        Assert.Equal(1D, original.CharacterSpacing);
        Assert.Equal(50D, original.HorizontalTextScaling);
        Assert.Equal(2D, changed.CharacterSpacing);
        Assert.Equal(75D, changed.HorizontalTextScaling);
    }

    [Fact]
    public void RunSpacing_RejectsNonFiniteValuesAndNonPositiveScaling() {
        PdfTextRun run = PdfTextRun.Normal("ABCD");
        foreach (double invalid in new[] { double.NaN, double.PositiveInfinity, double.NegativeInfinity }) {
            Assert.Throws<ArgumentOutOfRangeException>(() => run.WithCharacterSpacing(invalid));
            Assert.Throws<ArgumentOutOfRangeException>(() => run.WithHorizontalTextScaling(invalid));
        }
        Assert.Throws<ArgumentOutOfRangeException>(() => run.WithHorizontalTextScaling(0D));
        Assert.Throws<ArgumentOutOfRangeException>(() => run.WithHorizontalTextScaling(-1D));
        Assert.Equal(100D, run.HorizontalTextScaling);
        Assert.Equal(0D, run.CharacterSpacing);
    }

    private static PdfOptions Options() => new PdfOptions {
        DefaultFont = PdfStandardFont.Helvetica,
        DefaultFontSize = 12D,
        HeaderFont = PdfStandardFont.Helvetica,
        HeaderFontSize = 12D
    };

    private static byte[] Create(string surface, PdfTextRun run) {
        PdfDocument document = PdfDocument.Create(Options());
        switch (surface) {
            case "paragraph": document.Paragraph(paragraph => paragraph.Runs(new[] { run })); break;
            case "table":
                document.Table(new[] { new[] { PdfTableCell.RichTextCell(new[] { run }) } }, style: new PdfTableStyle { HeaderRowCount = 0 });
                break;
            case "header": document.Header(header => header.Text(text => text.Run(run))).Paragraph(paragraph => paragraph.Text("body")); break;
            default: throw new ArgumentOutOfRangeException(nameof(surface));
        }
        return document.ToBytes();
    }

    private static Letter[] Letters(byte[] bytes) {
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        return pdf.GetPage(1).Letters.Where(letter => "ABCD".IndexOf(letter.Value, StringComparison.Ordinal) >= 0)
            .OrderBy(letter => letter.StartBaseLine.X).ToArray();
    }
}
