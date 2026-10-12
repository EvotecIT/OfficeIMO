using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Theory]
    [InlineData("table", false)]
    [InlineData("table", true)]
    [InlineData("nested", false)]
    [InlineData("nested", true)]
    [InlineData("image", false)]
    [InlineData("image", true)]
    public void Panel_BordersRemainVisibleOverDescendantFills(string child, bool rowColumn) {
        var options = new PdfOptions {
            PageWidth = 200, PageHeight = 110, MarginLeft = 20, MarginRight = 20,
            MarginTop = 20, MarginBottom = 20, DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10, MaxGeneratedPages = 1
        };
        var style = new PdfPanelStyle {
            PaddingX = 0, PaddingY = 0, SpacingBefore = 0, SpacingAfter = 0,
            BorderColor = new PdfColor(0, 0, 1), BorderWidth = 2,
            Background = new PdfColor(1, 1, 0), RepeatFragmentDecoration = false
        };
        void Content(PdfContentBuilder content) => content.Panel(panel => {
            if (child == "image") {
                panel.Image(PdfPngTestImages.CreateRgbaPng(255, 0, 0, 255), 160, 40,
                    style: new PdfImageStyle { SpacingBefore = 0, SpacingAfter = 0, ScaleDownToFit = false });
            } else if (child == "table") {
                panel.Table(new[] { new[] { "A" }, new[] { "B" } }, style: new PdfTableStyle {
                    HeaderRowCount = 0, FontSize = 10, LineHeight = 2,
                    FixedRowHeights = new List<double?> { 20, 20 }, CellPaddingX = 0, CellPaddingY = 0,
                    BorderWidth = 0, BorderColor = null, RowStripeFill = new PdfColor(1, 0, 0),
                    SpacingBefore = 0, SpacingAfter = 0, KeepTogether = true
                });
            } else {
                panel.Panel(inner => inner.Paragraph(p => p.Text("A\nB"), style: new PdfParagraphStyle {
                    FontSize = 10, LineSpacing = PdfLineSpacing.Exactly(20),
                    SpacingBefore = 0, SpacingAfter = 0, WidowControl = false
                }), new PdfPanelStyle {
                    PaddingX = 0, PaddingY = 0, SpacingBefore = 0, SpacingAfter = 0,
                    BorderWidth = 0, Background = new PdfColor(1, 0, 0)
                });
            }
        }, style);
        byte[] bytes = PdfDocument.Create(options).Compose(document => document.Page(page =>
            page.Content(content => {
                if (rowColumn) content.Row(row => row.PercentColumn(100, Content));
                else Content(content);
            }))).ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal(1, pdf.NumberOfPages);
        if (child != "image") {
            Assert.Contains("A", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.Contains("B", pdf.GetPage(1).Text, StringComparison.Ordinal);
        }
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(
            PdfPageImageRenderer.RenderPage(bytes), background: OfficeColor.White);
        OfficeColor fill = raster.GetPixel(100, 50);
        Assert.True(fill.R > 240 && fill.G < 15 && fill.B < 15, "The descendant fill must remain visible.");
        foreach (int x in new[] { 20, 179 }) foreach (int y in new[] { 44, 50, 56 }) {
            OfficeColor border = raster.GetPixel(x, y);
            Assert.True(border.B > 240 && border.R < 15 && border.G < 15,
                $"Expected an intact blue panel border at ({x}, {y}); got {border}.");
        }
    }
}
