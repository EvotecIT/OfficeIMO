using System;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfWhitespaceFlowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PreservedFlowKeepsLeadingAndRepeatedSpaceAdvances(bool table) {
        PdfOptions options = Options(PdfTextWhitespaceMode.Preserve);
        PdfDocument document = PdfDocument.Create(options);
        var runs = new[] { PdfTextRun.Normal("  ALPHA "), PdfTextRun.Normal("  BETA") };
        if (table) document.Table(new[] { new[] { PdfTableCell.RichTextCell(runs) } },
            style: new PdfTableStyle { HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 0 });
        else document.Paragraph(p => p.Runs(runs));
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var letters = pdf.GetPage(1).Letters.Where(l => !string.IsNullOrWhiteSpace(l.Value)).ToArray();
        Assert.Equal("ALPHABETA", string.Concat(letters.Select(l => l.Value)));
        Assert.InRange(Math.Abs(letters[0].StartBaseLine.X - 20D - 2D * 3.336D), 0, .02D);
        Assert.InRange(Math.Abs(letters[5].StartBaseLine.X - letters[0].StartBaseLine.X - 39.348D - 3D * 3.336D), 0, .02D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FlowOverflowDiscardsExcessSpacesWithoutAdditionalBlankLines(bool leading) {
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options(PdfTextWhitespaceMode.Preserve))
            .Paragraph(p => p.Text((leading ? "" : "ALPHA") + new string(' ', 60) + "BETA")).ToBytes());
        var letters = pdf.GetPage(1).Letters.Where(l => !string.IsNullOrWhiteSpace(l.Value)).ToArray();
        var beta = letters.First(l => l.Value == "B");
        Assert.InRange(Math.Abs(beta.StartBaseLine.X - 20D), 0, .02D);
        using var bare = PdfPigDocument.Open(PdfDocument.Create(Options(PdfTextWhitespaceMode.Preserve))
            .Paragraph(p => p.Text("BETA")).ToBytes());
        double bareBaseline = bare.GetPage(1).Letters.First(l => l.Value == "B").StartBaseLine.Y;
        Assert.InRange(Math.Abs(bareBaseline - beta.StartBaseLine.Y - 14.4D), 0, .02D);
        if (!leading) Assert.InRange(Math.Abs(letters[0].StartBaseLine.Y - beta.StartBaseLine.Y - 14.4D), 0, .02D);
    }

    [Fact]
    public void PreformattedBooleanRetainsItsLiteralOverflowBehavior() {
        PdfOptions options = Options(PdfTextWhitespaceMode.Collapse);
        options.PreserveTextWhitespace = true;
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(options)
            .Paragraph(p => p.Text("ALPHA" + new string(' ', 60) + "BETA")).ToBytes());
        var letters = pdf.GetPage(1).Letters.Where(l => !string.IsNullOrWhiteSpace(l.Value)).ToArray();
        Assert.InRange(Math.Abs(letters[0].StartBaseLine.Y - letters[5].StartBaseLine.Y - 43.2D), 0, .02D);
        Assert.InRange(Math.Abs(letters[5].StartBaseLine.X - 28.16D), 0, .02D);
    }

    [Fact]
    public void JustificationWeightsEachPreservedSpaceAndOmitsWrappedTrailingSpaces() {
        PdfOptions options = Options(PdfTextWhitespaceMode.Preserve);
        options.PageWidth = 220D;
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(options).Paragraph(p =>
            p.Text("ALPHA BETA   GAMMA DELTA EPSILON ZETA"), align: PdfAlign.Justify).ToBytes());
        var letters = pdf.GetPage(1).Letters.Where(l => !string.IsNullOrWhiteSpace(l.Value)).ToArray();
        var firstLine = letters.Where(l => Math.Abs(l.StartBaseLine.Y - letters[0].StartBaseLine.Y) < .02D).ToArray();
        Assert.Equal("ALPHABETAGAMMADELTA", string.Concat(firstLine.Select(l => l.Value)));
        Assert.InRange(Math.Abs(firstLine.Last().EndBaseLine.X - 200D), 0, .02D);
        double firstExpansion = firstLine[5].StartBaseLine.X - firstLine[4].EndBaseLine.X - 3.336D;
        double repeatedExpansion = firstLine[9].StartBaseLine.X - firstLine[8].EndBaseLine.X - 3D * 3.336D;
        Assert.InRange(Math.Abs(repeatedExpansion - 3D * firstExpansion), 0, .02D);
    }

    [Fact]
    public void PageContinuationRetainsRepeatedSpaceAdvances() {
        PdfOptions options = Options(PdfTextWhitespaceMode.Preserve);
        options.PageHeight = 82D;
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(options).Paragraph(p =>
            p.Text(string.Join(" ", Enumerable.Repeat("ALPHA   BETA", 12)))).ToBytes());
        Assert.True(pdf.NumberOfPages > 1);
        foreach (var page in pdf.GetPages()) {
            var letters = page.Letters.Where(l => !string.IsNullOrWhiteSpace(l.Value)).ToArray();
            foreach (var line in letters.GroupBy(l => Math.Round(l.StartBaseLine.Y, 2))) {
                var content = line.ToArray();
                Assert.Equal("ALPHABETA", string.Concat(content.Select(l => l.Value)));
                Assert.InRange(Math.Abs(content[5].StartBaseLine.X - content[0].StartBaseLine.X - 39.348D - 3D * 3.336D), 0, .02D);
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SpacesBeforeInlineVisualKeepTheirAdvanceOrWrapOnce(bool overflow) {
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options(PdfTextWhitespaceMode.Preserve))
            .Paragraph(p => p.Runs(new[] {
                PdfTextRun.Normal(new string(' ', overflow ? 60 : 2)),
                PdfTextRun.Inline(new PdfInlineBox(10, 8)), PdfTextRun.Normal("BETA")
            })).ToBytes());
        var beta = pdf.GetPage(1).Letters.First(l => l.Value == "B");
        Assert.InRange(Math.Abs(beta.StartBaseLine.X - 30D - (overflow ? 0D : 6.672D)), 0, .02D);
        using var bare = PdfPigDocument.Open(PdfDocument.Create(Options(PdfTextWhitespaceMode.Preserve))
            .Paragraph(p => p.Text("BETA")).ToBytes());
        double bareBaseline = bare.GetPage(1).Letters.First(l => l.Value == "B").StartBaseLine.Y;
        Assert.InRange(Math.Abs(bareBaseline - beta.StartBaseLine.Y - (overflow ? 14.4D : 0D)), 0, .02D);
    }

    private static PdfOptions Options(PdfTextWhitespaceMode mode) => new PdfOptions {
        DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 12D,
        PageWidth = 136D, PageHeight = 240D,
        MarginLeft = 20D, MarginRight = 20D, MarginTop = 20D, MarginBottom = 20D,
        DefaultParagraphStyle = new PdfParagraphStyle { LineSpacing = PdfLineSpacing.Exactly(14.4D) },
        TextWhitespaceMode = mode
    };

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SplitCellRetainsTheFontsAndAdvancesOfEachSourceSpace(bool columns) {
        PdfOptions options = Options(PdfTextWhitespaceMode.Preserve);
        options.PageHeight = 82D;
        if (columns) options.PageWidth = 284D;
        var runs = Enumerable.Range(0, 12).SelectMany(_ => new[] {
            PdfTextRun.Normal("ALPHA ", fontSize: 12D),
            new PdfTextRun(" ", bold: true, fontSize: 20D),
            PdfTextRun.Normal(" BETA\n", fontSize: 12D)
        }).ToArray();
        using var pdf = PdfPigDocument.Open(RenderCell(options, runs, columns));
        Assert.True(pdf.NumberOfPages > 1);
        foreach (var page in pdf.GetPages()) {
            foreach (var line in page.Letters.Where(l => !string.IsNullOrWhiteSpace(l.Value))
                         .GroupBy(l => (Math.Round(l.StartBaseLine.Y, 2), columns && l.StartBaseLine.X >= 135.9D))) {
                var letters = line.ToArray();
                Assert.Equal("ALPHABETA", string.Concat(letters.Select(l => l.Value)));
                Assert.InRange(Math.Abs(letters[5].StartBaseLine.X - letters[0].StartBaseLine.X
                    - 39.348D - 12.232D), 0D, .02D);
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SplitLiteralCellRetainsOnlyTheAdvanceOfEachPartialSpacer(bool columns) {
        PdfOptions options = Options(PdfTextWhitespaceMode.Preformatted);
        options.PageHeight = 82D;
        if (columns) options.PageWidth = 376D;
        var runs = new[] { PdfTextRun.Normal(string.Join("\n", Enumerable.Repeat("ALPHA" + new string(' ', 60) + "BETA", columns ? 1 : 3))) };
        using var pdf = PdfPigDocument.Open(RenderCell(options, runs, columns, 220D));
        Assert.Equal(columns ? 1 : 6, pdf.NumberOfPages);
        var starts = pdf.GetPages().SelectMany(p => p.Letters).Where(l => l.Value == "B").ToArray();
        Assert.Equal(columns ? 1 : 3, starts.Length);
        Assert.All(starts, letter => Assert.InRange(Math.Abs(letter.StartBaseLine.X - (columns ? 240.16D : 28.16D)), 0D, .02D));
    }

    [Fact]
    public void AHighlightedSpaceRunKeepsItsOwnBackgroundBetweenUnhighlightedWords() {
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options(PdfTextWhitespaceMode.Preserve)).Paragraph(p =>
            p.Runs(new[] { PdfTextRun.Normal("ALPHA"), new PdfTextRun("   ", backgroundColor: new PdfColor(1D, 1D, 0D)),
                PdfTextRun.Normal("BETA") })).ToBytes());
        Assert.NotEmpty(pdf.GetPage(1).Paths);
    }

    [Theory]
    [InlineData(PdfAlign.Center)]
    [InlineData(PdfAlign.Right)]
    public void TrailingFlowSpacesBeforeAHardBreakDoNotShiftAlignment(PdfAlign align) {
        using var expected = PdfPigDocument.Open(PdfDocument.Create(Options(PdfTextWhitespaceMode.Preserve))
            .Paragraph(p => p.Text("ALPHA\nBETA"), align: align).ToBytes());
        using var actual = PdfPigDocument.Open(PdfDocument.Create(Options(PdfTextWhitespaceMode.Preserve))
            .Paragraph(p => p.Text("ALPHA   \nBETA"), align: align).ToBytes());
        var a = actual.GetPage(1).Letters.Where(l => !string.IsNullOrWhiteSpace(l.Value)).ToArray();
        var b = expected.GetPage(1).Letters.Where(l => !string.IsNullOrWhiteSpace(l.Value)).ToArray();
        Assert.Equal(b.Length, a.Length);
        for (int i = 0; i < a.Length; i++) Assert.InRange(Math.Abs(a[i].StartBaseLine.X - b[i].StartBaseLine.X), 0D, .02D);
    }

    private static byte[] RenderCell(PdfOptions options, PdfTextRun[] runs, bool columns, double secondColumnWidth = 128D) {
        var cells = new[] { new[] { new PdfTableCell(runs, new[] { new PdfTableCellParagraph(runs,
            fontSize: 12D, lineSpacing: PdfLineSpacing.Exactly(14.4D), widowControl: false) }) } };
        var style = new PdfTableStyle { HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 0,
            BorderWidth = 0, RowSeparatorWidth = 0, FontSize = 12D, SpacingBefore = 0, SpacingAfter = 0 };
        PdfDocument document = PdfDocument.Create(options);
        if (columns) document.Columns(content => content.Table(cells, style: style), new PdfMultiColumnOptions {
            BalanceLastPage = false, ColumnDefinitions = new[] {
                new PdfFlowColumn(PdfColumnWidth.Fixed(96D), 20D), new PdfFlowColumn(PdfColumnWidth.Fixed(secondColumnWidth), 0D)
            } });
        else document.Table(cells, style: style);
        return document.ToBytes();
    }
}
