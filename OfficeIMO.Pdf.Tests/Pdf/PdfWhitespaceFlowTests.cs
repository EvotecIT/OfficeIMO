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
}
