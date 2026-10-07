using System;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using UglyToad.PdfPig.Content;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRunSpacingRegressionTests {
    [Theory]
    [InlineData(100D, -4D, 1, false)]
    [InlineData(50D, -2D, 1, false)]
    [InlineData(200D, -8D, 1, false)]
    [InlineData(100D, -3.336D, 1, false)]
    [InlineData(100D, -4D, 2, true)]
    public void CondensedRunWithSpaces_PreservesSignedSeparatorAdvances(double scaling, double spacing, int separatorCount, bool preserveWhitespace) {
        PdfOptions options = Options();
        options.PreserveTextWhitespace = preserveWhitespace;
        var run = new PdfTextRun("MMMM" + new string(' ', separatorCount) + "MMMM", fontSize: 12D)
            .WithHorizontalTextScaling(scaling).WithCharacterSpacing(spacing);
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(options)
            .Paragraph(p => p.Runs(new[] { run, PdfTextRun.Normal("X", fontSize: 12D) })).ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        Letter[] letters = pdf.GetPage(1).Letters.ToArray();
        Letter[] ems = letters.Where(letter => letter.Value == "M").ToArray();
        Assert.Equal(8, ems.Length);
        double origin = ems[0].StartBaseLine.X;
        double wordAdvance = 4D * (833D * 12D / 1000D * scaling / 100D + spacing);
        double separatorAdvance = separatorCount * (278D * 12D / 1000D * scaling / 100D + spacing);
        Assert.InRange(Math.Abs(ems[4].StartBaseLine.X - origin - wordAdvance - separatorAdvance), 0D, 0.02D);
        Assert.InRange(Math.Abs(letters.Single(letter => letter.Value == "X").StartBaseLine.X - origin
            - 2D * wordAdvance - separatorAdvance), 0D, 0.02D);
    }

    [Fact]
    public void CondensedRunWithoutSpaces_DoesNotMeasureAnUnusedSpace() {
        var run = new PdfTextRun("MMMM", fontSize: 12D).WithHorizontalTextScaling(50D).WithCharacterSpacing(-2D);
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options()).Paragraph(p => p.Runs(new[] { run })).ToBytes());
        Letter[] letters = pdf.GetPage(1).Letters.ToArray();
        Assert.Equal("MMMM", string.Concat(letters.Select(letter => letter.Value)));
        Assert.InRange(letters[1].StartBaseLine.X - letters[0].StartBaseLine.X, 2.99D, 3.01D);
    }

    [Fact]
    public void StyledSeparatorBeforeDefaultRun_UsesItsMeasuredAdvance() {
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options())
            .Paragraph(p => p.HorizontalTextScaling(50D).CharacterSpacing(1D).Text("AB ")
                .HorizontalTextScaling(100D).CharacterSpacing(0D).Text("CD")).ToBytes());
        var letters = pdf.GetPage(1).Letters;
        double origin = letters.First(letter => letter.Value == "A").StartBaseLine.X;
        double c = letters.Single(letter => letter.Value == "C").StartBaseLine.X;
        // Helvetica: A/B 667 units each, space 278, plus three authored glyph advances.
        Assert.InRange(Math.Abs(c - origin - (12D * (667D * 2D + 278D) / 1000D / 2D + 3D)), 0D, 0.02D);
    }

    [Fact]
    public void ScaledDecimalTabs_AlignDifferentNumericPrefixes() {
        PdfOptions options = Options();
        options.DefaultParagraphStyle = new PdfParagraphStyle { DefaultTabStopWidth = 150D, SpacingAfter = 0D };
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(options)
            .Paragraph(p => p.Text("Tax").Tab(PdfTabLeaderStyle.None, PdfTabAlignment.DecimalSeparator)
                .HorizontalTextScaling(50D).CharacterSpacing(1D).Text("8.50"))
            .Paragraph(p => p.Text("Total").Tab(PdfTabLeaderStyle.None, PdfTabAlignment.DecimalSeparator)
                .HorizontalTextScaling(50D).CharacterSpacing(1D).Text("12845.75")).ToBytes());
        var periods = pdf.GetPage(1).Letters.Where(letter => letter.Value == ".").ToArray();
        Assert.Equal(2, periods.Length);
        Assert.InRange(Math.Abs(periods[0].StartBaseLine.X - periods[1].StartBaseLine.X), 0D, 0.02D);
        Assert.InRange(Math.Abs(periods[0].StartBaseLine.X - (options.MarginLeft + 150D)), 0D, 0.02D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WordOnlyUnderline_UsesStyledGlyphAdvances(bool embedded) {
        PdfOptions options = Options();
        if (embedded) {
            string? font = PdfComplianceTestFonts.FindBundledTrueTypeFont();
            Assert.NotNull(font);
            options.EmbedStandardFont(PdfStandardFont.Helvetica, System.IO.File.ReadAllBytes(font!), "SpacingUnderline");
        }
        var run = new PdfTextRun("ABCD", fontSize: 12D, underlineStyle: OfficeTextDecorationStyle.Words)
            .WithHorizontalTextScaling(50D).WithCharacterSpacing(1D);
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(options).Paragraph(p => p.Runs(new[] { run, PdfTextRun.Normal("X", fontSize: 12D) })).ToBytes());
        var page = pdf.GetPage(1);
        double start = page.Letters.First(letter => letter.Value == "A").StartBaseLine.X;
        double end = page.Letters.Single(letter => letter.Value == "X").StartBaseLine.X;
        var bounds = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(rectangle => rectangle.HasValue).Select(rectangle => rectangle!.Value).ToArray();
        var underline = Assert.Single(bounds);
        Assert.InRange(Math.Abs(underline.Left - start), 0D, 0.02D);
        Assert.InRange(Math.Abs(underline.Right - end), 0D, 0.02D);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void WordAcrossFormattingRuns_MovesTogetherWhenItFitsANewLine(bool scaled, bool table) {
        PdfOptions options = Options();
        options.PageWidth = 100D; options.MarginLeft = options.MarginRight = 20D;
        var first = new PdfTextRun(scaled ? "XXXXXXXXXXX ALP" : "XXX ALP", fontSize: 12D);
        var last = new PdfTextRun("HA", fontSize: 12D, color: PdfColor.FromRgb(0, 0, 255));
        if (scaled) {
            first = first.WithHorizontalTextScaling(50D);
            last = last.WithHorizontalTextScaling(50D).WithCharacterSpacing(1D);
        }
        var document = PdfDocument.Create(options);
        if (table) document.Table(new[] { new[] { PdfTableCell.RichTextCell(new[] { first, last }) } },
            style: new PdfTableStyle { HeaderRowCount = 0, CellPaddingX = 0D, CellPaddingY = 0D });
        else document.Paragraph(p => p.Runs(new[] { first, last }));
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var lines = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1)).OrderByDescending(group => group.Key)
            .Select(group => string.Concat(group.OrderBy(letter => letter.StartBaseLine.X).Select(letter => letter.Value))).ToArray();
        Assert.Equal(new[] { scaled ? "XXXXXXXXXXX" : "XXX", "ALPHA" }, lines);
    }

    private static PdfOptions Options() => new PdfOptions {
        DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 12D,
        MarginLeft = 20D, MarginRight = 20D, MarginTop = 20D, MarginBottom = 20D
    };
}
