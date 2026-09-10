using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfAnchoredCanvasFlowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BalancedParagraphContinuationEmitsItsAnchoredImageOnlyOnce(bool balanceLines) {
        byte[] image = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(12, 12, OfficeColor.Black), OfficeImageExportFormat.Png);
        var canvas = new PdfPageCanvas().ForegroundImage(image, 5, 5, 12, 12);
        var style = new PdfParagraphStyle { AnchoredCanvas = new PdfCanvasBlock(canvas.Items) };
        var document = PdfDocument.Create(new PdfOptions {
            PageWidth = 300, PageHeight = 200, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, DefaultFontSize = 10
        }).Columns(content => content.Paragraph(paragraph => paragraph.Text("ANCHOR " +
                string.Join(" ", Enumerable.Repeat("Column continuation text", 32))), style: style),
            new PdfMultiColumnOptions { ColumnCount = 2, Gap = 12, BalanceLastPage = true, BalanceParagraphLines = balanceLines });
        for (int serialization = 0; serialization < 2; serialization++) {
            var reopened = PdfDocument.Load(document.ToBytes());
            var placement = Assert.Single(reopened.Images.Placements());
            int anchorPage = Assert.Single(reopened.Reader.Pages(), page =>
                reopened.Reader.Text(PdfPageSelection.From(page.PageNumber)).Contains("ANCHOR")).PageNumber;
            Assert.Equal(anchorPage, placement.PageNumber);
            Assert.Equal(5, placement.X, 3);
            Assert.Equal(183, placement.Y, 3);
            Assert.Contains("continuation", reopened.Reader.Text(PdfPageSelection.From(reopened.Reader.Pages().Count)));
        }
    }
}
