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

    [Fact]
    public void AFontOnAnEmptySpaceFragmentDoesNotMoveTheVisibleTextBaseline() {
        PdfOptions options = Options(PdfTextWhitespaceMode.Preserve);
        options.DefaultParagraphStyle = new PdfParagraphStyle();
        using var plain = PdfPigDocument.Open(PdfDocument.Create(options).Paragraph(p => p.Text("ALPHA BETA")).ToBytes());
        using var mixed = PdfPigDocument.Open(PdfDocument.Create(options).Paragraph(p => p.Runs(new[] {
            PdfTextRun.Normal("ALPHA ", fontSize: 12D), new PdfTextRun(" ", fontSize: 20D),
            PdfTextRun.Normal(" BETA", fontSize: 12D)
        })).ToBytes());
        var actual = mixed.GetPage(1).Letters.Where(l => !string.IsNullOrWhiteSpace(l.Value)).ToArray();
        var expected = plain.GetPage(1).Letters.Where(l => !string.IsNullOrWhiteSpace(l.Value)).ToArray();
        Assert.Equal(expected.Length, actual.Length);
        for (int i = 0; i < actual.Length; i++) Assert.InRange(Math.Abs(actual[i].StartBaseLine.Y - expected[i].StartBaseLine.Y), 0D, .02D);
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

    [Theory]
    [InlineData(PdfTextWhitespaceMode.Preserve, -10D, false)]
    [InlineData(PdfTextWhitespaceMode.Preserve, -6.672D, false)]
    [InlineData(PdfTextWhitespaceMode.Preserve, -5D, false)]
    [InlineData(PdfTextWhitespaceMode.Preformatted, -10D, false)]
    [InlineData(PdfTextWhitespaceMode.Preformatted, -6.672D, false)]
    [InlineData(PdfTextWhitespaceMode.Preformatted, -5D, false)]
    [InlineData(PdfTextWhitespaceMode.Preserve, -10D, true)]
    [InlineData(PdfTextWhitespaceMode.Preserve, -6.672D, true)]
    [InlineData(PdfTextWhitespaceMode.Preserve, -5D, true)]
    [InlineData(PdfTextWhitespaceMode.Preformatted, -10D, true)]
    [InlineData(PdfTextWhitespaceMode.Preformatted, -6.672D, true)]
    [InlineData(PdfTextWhitespaceMode.Preformatted, -5D, true)]
    public void MixedSignedSpacesPaintTheirTotalAdvance(PdfTextWhitespaceMode mode, double tracking, bool inline) {
        var runs = new List<PdfTextRun> { PdfTextRun.Normal("A ", fontSize: 12D),
            PdfTextRun.Normal(" ", fontSize: 12D).WithCharacterSpacing(tracking) };
        if (inline) runs.Add(PdfTextRun.Inline(new PdfInlineBox(10D, 8D)));
        runs.Add(PdfTextRun.Normal("B", fontSize: 12D));
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options(mode)).Paragraph(p => p.Runs(runs)).ToBytes());
        var letters = pdf.GetPage(1).Letters.Where(l => l.Value is "A" or "B").ToArray();
        Assert.Equal("AB", string.Concat(letters.Select(l => l.Value)));
        Assert.InRange(Math.Abs(letters[1].StartBaseLine.X - letters[0].EndBaseLine.X
            - 6.672D - tracking - (inline ? 10D : 0D)), 0D, .02D);
        Assert.InRange(Math.Abs(letters[1].StartBaseLine.Y - letters[0].StartBaseLine.Y), 0D, .02D);
    }

    [Theory]
    [InlineData(PdfTextWhitespaceMode.Preserve, false, false)]
    [InlineData(PdfTextWhitespaceMode.Preserve, true, false)]
    [InlineData(PdfTextWhitespaceMode.Preserve, false, true)]
    [InlineData(PdfTextWhitespaceMode.Preserve, true, true)]
    [InlineData(PdfTextWhitespaceMode.Preformatted, false, false)]
    [InlineData(PdfTextWhitespaceMode.Preformatted, true, false)]
    [InlineData(PdfTextWhitespaceMode.Preformatted, false, true)]
    [InlineData(PdfTextWhitespaceMode.Preformatted, true, true)]
    public void ATabFollowedBySpaceRetainsItsStopAndLeader(PdfTextWhitespaceMode mode, bool explicitStop, bool inline) {
        PdfOptions options = Options(mode);
        PdfParagraphStyle style = options.DefaultParagraphStyle!;
        if (explicitStop) style.AddTabStop(60D, leader: PdfTabLeaderStyle.Dots);
        byte[] Render(bool space) {
            var runs = new List<PdfTextRun> { PdfTextRun.Normal("A\t" + (space ? " " : "")) };
            if (inline) runs.Add(PdfTextRun.Inline(new PdfInlineBox(10D, 8D)));
            runs.Add(PdfTextRun.Normal("B"));
            return PdfDocument.Create(options).Paragraph(p => p.Runs(runs), style: style).ToBytes();
        }
        using var expected = PdfPigDocument.Open(Render(false));
        using var actual = PdfPigDocument.Open(Render(true));
        var a = actual.GetPage(1).Letters.First(l => l.Value == "B");
        var b = expected.GetPage(1).Letters.First(l => l.Value == "B");
        Assert.InRange(Math.Abs(a.StartBaseLine.X - b.StartBaseLine.X - 3.336D), 0D, .02D);
        Assert.InRange(Math.Abs(a.StartBaseLine.Y - b.StartBaseLine.Y), 0D, .02D);
        if (explicitStop) {
            Assert.InRange(Math.Abs(b.StartBaseLine.X - 80D - (inline ? 10D : 0D)), 0D, .02D);
            Assert.True(expected.GetPage(1).Letters.Count(l => l.Value == ".") > 0);
            Assert.Equal(expected.GetPage(1).Letters.Count(l => l.Value == "."),
                actual.GetPage(1).Letters.Count(l => l.Value == "."));
        }
    }

    [Theory]
    [InlineData(PdfTextWhitespaceMode.Preserve)]
    [InlineData(PdfTextWhitespaceMode.Preformatted)]
    public void TabThenSpaceSurvivesChangedWidthColumnContinuation(PdfTextWhitespaceMode mode) {
        PdfOptions options = Options(mode); options.PageWidth = 284D; options.PageHeight = 82D;
        byte[] Render(bool space) => RenderCell(options, new[] {
            PdfTextRun.Normal(string.Join("\n", Enumerable.Repeat("A\t" + (space ? " " : "") + "B", 12)))
        }, true);
        using var expected = PdfPigDocument.Open(Render(false));
        using var actual = PdfPigDocument.Open(Render(true));
        Assert.True(actual.NumberOfPages > 1);
        Assert.Equal(expected.NumberOfPages, actual.NumberOfPages);
        var a = actual.GetPages().SelectMany(p => p.Letters).Where(l => l.Value == "B").ToArray();
        var b = expected.GetPages().SelectMany(p => p.Letters).Where(l => l.Value == "B").ToArray();
        Assert.Equal(12, a.Length); Assert.Equal(b.Length, a.Length);
        for (int i = 0; i < a.Length; i++) {
            Assert.InRange(Math.Abs(a[i].StartBaseLine.X - b[i].StartBaseLine.X - 3.336D), 0D, .02D);
            Assert.InRange(Math.Abs(a[i].StartBaseLine.Y - b[i].StartBaseLine.Y), 0D, .02D);
        }
    }

    [Theory]
    [InlineData(PdfTextWhitespaceMode.Preserve, PdfTabAlignment.Right, false, 0D)]
    [InlineData(PdfTextWhitespaceMode.Preserve, PdfTabAlignment.Center, false, 1.668D)]
    [InlineData(PdfTextWhitespaceMode.Preformatted, PdfTabAlignment.Right, false, 0D)]
    [InlineData(PdfTextWhitespaceMode.Preformatted, PdfTabAlignment.Center, false, 1.668D)]
    [InlineData(PdfTextWhitespaceMode.Preserve, PdfTabAlignment.Right, true, 0D)]
    [InlineData(PdfTextWhitespaceMode.Preserve, PdfTabAlignment.Center, true, 1.668D)]
    [InlineData(PdfTextWhitespaceMode.Preformatted, PdfTabAlignment.Right, true, 0D)]
    [InlineData(PdfTextWhitespaceMode.Preformatted, PdfTabAlignment.Center, true, 1.668D)]
    public void BufferedTabUsesTheFollowingSpaceAndContentForAlignment(
        PdfTextWhitespaceMode mode, PdfTabAlignment align, bool inline, double delta) {
        PdfOptions options = Options(mode);
        PdfParagraphStyle style = options.DefaultParagraphStyle!; style.AddTabStop(60D, align);
        byte[] Render(bool space) {
            var runs = new List<PdfTextRun> { PdfTextRun.Normal("A\t" + (space ? " " : "")) };
            if (inline) runs.Add(PdfTextRun.Inline(new PdfInlineBox(10D, 8D)));
            runs.Add(PdfTextRun.Normal("B"));
            return PdfDocument.Create(options).Paragraph(p => p.Runs(runs), style: style).ToBytes();
        }
        using var expected = PdfPigDocument.Open(Render(false));
        using var actual = PdfPigDocument.Open(Render(true));
        var a = actual.GetPage(1).Letters.First(l => l.Value == "B");
        var b = expected.GetPage(1).Letters.First(l => l.Value == "B");
        Assert.True(Math.Abs(a.StartBaseLine.X - b.StartBaseLine.X - delta) <= .02D,
            $"Buffered tab B={a.StartBaseLine.X}, bare tab B={b.StartBaseLine.X}, expected delta={delta}.");
        Assert.InRange(Math.Abs(a.StartBaseLine.Y - b.StartBaseLine.Y), 0D, .02D);
    }

    [Theory]
    [InlineData(PdfTextWhitespaceMode.Preserve, -10D)]
    [InlineData(PdfTextWhitespaceMode.Preserve, -6.672D)]
    [InlineData(PdfTextWhitespaceMode.Preserve, -5D)]
    [InlineData(PdfTextWhitespaceMode.Preformatted, -10D)]
    [InlineData(PdfTextWhitespaceMode.Preformatted, -6.672D)]
    [InlineData(PdfTextWhitespaceMode.Preformatted, -5D)]
    public void MixedSignedSpacesSurviveChangedWidthColumnContinuation(PdfTextWhitespaceMode mode, double tracking) {
        PdfOptions options = Options(mode); options.PageWidth = 284D; options.PageHeight = 82D;
        var runs = Enumerable.Range(0, 12).SelectMany(_ => new[] {
            PdfTextRun.Normal("A "), PdfTextRun.Normal(" ").WithCharacterSpacing(tracking), PdfTextRun.Normal("B\n")
        }).ToArray();
        using var pdf = PdfPigDocument.Open(RenderCell(options, runs, true));
        Assert.True(pdf.NumberOfPages > 1);
        var letters = pdf.GetPages().SelectMany(p => p.Letters).Where(l => l.Value is "A" or "B").ToArray();
        Assert.Equal(24, letters.Length);
        for (int i = 0; i < letters.Length; i += 2) {
            Assert.Equal("A", letters[i].Value); Assert.Equal("B", letters[i + 1].Value);
            Assert.InRange(Math.Abs(letters[i + 1].StartBaseLine.X - letters[i].EndBaseLine.X - 6.672D - tracking), 0D, .02D);
            Assert.InRange(Math.Abs(letters[i + 1].StartBaseLine.Y - letters[i].StartBaseLine.Y), 0D, .02D);
        }
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
