using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTableCellTextSizingTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void CellTextShrinkingIsIndependentInFlowAndCanvasTables(bool canvas, bool explicitSize) {
        double targetSize = explicitSize ? 20 : 12;
        var target = explicitSize
            ? new PdfTableCell(new[] { new PdfTextRun("ShrinkTarget", fontSize: targetSize) })
            : new PdfTableCell("ShrinkTarget");
        var rows = new[] { new[] {
            target.WithShrinkTextToFit().WithNoWrap().WithNamedDestination("Target"),
            new PdfTableCell("Neighbor"),
            target.WithNoWrap()
        } };
        var style = new PdfTableStyle {
            HeaderRowCount = 0, HeaderBold = false, FontSize = 12,
            ColumnWidthPoints = new() { 50, 90, 50 },
            FixedRowHeights = new() { 40 },
            CellPaddingX = 0, CellPaddingY = 0, MinimumShrinkFontSize = 4,
            PreserveWidth = true, MaxWidth = 190, SpacingBefore = 0, SpacingAfter = 0
        };
        var document = PdfDocument.Create(new PdfOptions {
            DefaultFont = PdfStandardFont.Courier, DefaultFontSize = 12,
            PageWidth = 300, PageHeight = 200,
            MarginLeft = 20, MarginRight = 20, MarginTop = 20, MarginBottom = 20
        });
        if (canvas) document.Canvas(surface => surface.Table(rows, 20, 20, 190, 40, style));
        else document.Table(rows, style: style);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToBytes());
        var words = pdf.GetPage(1).GetWords().ToArray();
        var targets = words.Where(word => word.Text == "ShrinkTarget").OrderBy(word => word.BoundingBox.Left).ToArray();
        Assert.Equal(2, targets.Length);
        Assert.All(targets[0].Letters, letter => Assert.InRange(letter.PointSize, 4, 11));
        Assert.All(targets[1].Letters, letter => Assert.InRange(letter.PointSize, targetSize - .01, targetSize + .01));
        var neighbor = Assert.Single(words, word => word.Text == "Neighbor");
        Assert.All(neighbor.Letters, letter => Assert.InRange(letter.PointSize, 11.99, 12.01));
    }
}
