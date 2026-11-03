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
    [InlineData("image-row", false)]
    public void Panel_BordersRemainVisibleOverDescendantFills(string child, bool rowColumn) {
        var options = new PdfCreateOptions {
            PageWidth = Pt(200), PageHeight = Pt(110), MarginLeft = Pt(20), MarginRight = Pt(20),
            MarginTop = Pt(20), MarginBottom = Pt(20), DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = Pt(10), MaxGeneratedPages = 1
        };
        var style = new PdfPanelStyle {
            PaddingX = Pt(0), PaddingY = Pt(0), SpacingBefore = Pt(0), SpacingAfter = Pt(0),
            BorderColor = global::OfficeIMO.OfficeColor.FromRgb(0, 0, 255), BorderWidth = Pt(2),
            Background = global::OfficeIMO.OfficeColor.FromRgb(255, 255, 0), RepeatFragmentDecoration = false
        };
        void Content(PdfContentBuilder content) => content.Panel(panel => {
            if (child.StartsWith("image", StringComparison.Ordinal)) {
                void Image(PdfContentBuilder target) => target.AddImage(PdfPngTestImages.CreateRgbaPng(255, 0, 0, 255), Pt(160), Pt(40),
                    style: new PdfImageStyle { SpacingBefore = Pt(0), SpacingAfter = Pt(0), ScaleDownToFit = false });
                if (child == "image-row") panel.Row(row => row.PercentColumn(100, Image));
                else Image(panel);
            } else if (child == "table") {
                panel.AddTable(new[] { new[] { "A" }, new[] { "B" } }, style: new PdfTableStyle {
                    HeaderRowCount = 0, FontSize = Pt(10), LineHeight = 2,
                    FixedRowHeights = new List<global::OfficeIMO.OfficeLength?> { Pt(20), Pt(20) }, CellPaddingX = Pt(0), CellPaddingY = Pt(0),
                    BorderWidth = Pt(0), BorderColor = null, RowStripeFill = global::OfficeIMO.OfficeColor.FromRgb(255, 0, 0),
                    SpacingBefore = Pt(0), SpacingAfter = Pt(0), KeepTogether = true
                });
            } else {
                panel.Panel(inner => inner.AddParagraph(p => p.AddText("A\nB"), style: new PdfParagraphStyle {
                    FontSize = Pt(10), LineSpacing = PdfLineSpacing.Exact(Pt(20)),
                    SpacingBefore = Pt(0), SpacingAfter = Pt(0), WidowControl = false
                }), new PdfPanelStyle {
                    PaddingX = Pt(0), PaddingY = Pt(0), SpacingBefore = Pt(0), SpacingAfter = Pt(0),
                    BorderWidth = Pt(0), Background = global::OfficeIMO.OfficeColor.FromRgb(255, 0, 0)
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
        if (!child.StartsWith("image", StringComparison.Ordinal)) {
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

    [Fact]
    public void Panel_FlowImageKeepsItsDebugBoxAndPlacement() {
        var options = new PdfCreateOptions {
            PageWidth = Pt(200), PageHeight = Pt(110), MarginLeft = Pt(20), MarginRight = Pt(20),
            MarginTop = Pt(20), MarginBottom = Pt(20), Debug = new PdfDebugOptions { ShowFlowObjectBoxes = true }
        };
        void Image(PdfContentBuilder content) => content.AddImage(PdfPngTestImages.CreateRgbaPng(255, 0, 0, 255), Pt(40), Pt(20),
            style: new PdfImageStyle { SpacingBefore = Pt(0), SpacingAfter = Pt(0), Align = OfficeAlignment.Center, ScaleDownToFit = false });
        byte[] bytes = PdfDocument.Create(options).Compose(document => document.Page(page => page.Content(content =>
            content.Panel(Image, new PdfPanelStyle { PaddingX = Pt(0), PaddingY = Pt(2), SpacingBefore = Pt(0), SpacingAfter = Pt(0) })))).ToBytes();
        PdfImagePlacement placement = Assert.Single(PdfImageExtractor.ExtractImagePlacements(bytes));
        Assert.Equal(80, placement.X, 3);
        Assert.Equal(68, placement.Y, 3);
        Assert.Equal(40, placement.Width, 3);
        Assert.Equal(20, placement.Height, 3);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes), background: OfficeColor.White);
        OfficeColor debug = raster.GetPixel(80, 30);
        Assert.True(debug.R > 64 && debug.B > 64 && debug.G < 32,
            $"Expected the magenta flow-image debug box; got {debug}.");
    }

    private static global::OfficeIMO.OfficeLength Pt(double points) => global::OfficeIMO.OfficeLength.Points(points);
}
