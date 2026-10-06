using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfTableCellViewportTests {
    [Theory]
    [InlineData("flow", 90)]
    [InlineData("column", 90)]
    [InlineData("canvas", 90)]
    [InlineData("flow", -90)]
    [InlineData("column", -90)]
    [InlineData("canvas", -90)]
    public void Rotated_image_and_link_survive_when_only_rotated_ink_intersects_the_fragment(string mode, double rotation) {
        var imageStyle = new PdfImageStyle { Align = PdfAlign.Center, RotationAngle = rotation };
        using PdfPigDocument full = PdfPigDocument.Open(RenderRotatedImage(mode, imageStyle, fragment: false));
        Assert.Single(full.GetPage(1).GetImages());
        var fullLink = Assert.Single(full.GetPage(1).GetHyperlinks());
        Assert.Equal(60, fullLink.Bounds.Width, 3);
        Assert.Equal(12, fullLink.Bounds.Height, 3);

        using PdfPigDocument fragment = PdfPigDocument.Open(RenderRotatedImage(mode, imageStyle));
        Assert.Single(fragment.GetPage(1).GetImages());
        var link = Assert.Single(fragment.GetPage(1).GetHyperlinks());
        Assert.Equal(24, link.Bounds.Left, 3);
        Assert.Equal(34, link.Bounds.Right, 3);
        Assert.Equal(120, link.Bounds.Bottom, 3);
        Assert.Equal(132, link.Bounds.Top, 3);
    }

    [Theory]
    [InlineData("flow", false)]
    [InlineData("column", false)]
    [InlineData("canvas", false)]
    [InlineData("flow", true)]
    [InlineData("column", true)]
    [InlineData("canvas", true)]
    public void Rotated_ink_outside_the_existing_image_clip_is_omitted_with_its_link(string mode, bool crop) {
        var imageStyle = new PdfImageStyle { Align = PdfAlign.Center, RotationAngle = 90 };
        if (crop) imageStyle.SourceCrop = new PdfImageSourceCrop { Bottom = .5 };
        else imageStyle.ClipPath = OfficeClipPath.Rectangle(12, 12);
        using PdfPigDocument fragment = PdfPigDocument.Open(RenderRotatedImage(mode, imageStyle));
        Assert.Empty(fragment.GetPage(1).GetImages());
        Assert.Empty(fragment.GetPage(1).GetHyperlinks());
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    [InlineData("canvas")]
    public void Source_crop_is_projected_before_intersecting_rotated_image_links_with_the_fragment(string mode) {
        var imageStyle = new PdfImageStyle { Align = PdfAlign.Center, RotationAngle = 90,
            SourceCrop = new PdfImageSourceCrop { Top = .5 } };
        using PdfPigDocument fragment = PdfPigDocument.Open(RenderRotatedImage(mode, imageStyle));
        Assert.Single(fragment.GetPage(1).GetImages());
        var link = Assert.Single(fragment.GetPage(1).GetHyperlinks());
        Assert.Equal(24, link.Bounds.Left, 3);
        Assert.Equal(64, link.Bounds.Right, 3);
        Assert.Equal(150, link.Bounds.Bottom, 3);
        Assert.Equal(156, link.Bounds.Top, 3);
    }

    private static byte[] RenderRotatedImage(string mode, PdfImageStyle imageStyle, bool fragment = true) {
        double width = fragment ? 80 : 200;
        var cell = PdfTableCell.WithImages(string.Empty, new[] {
            new PdfTableCellImage(PdfPngTestImages.CreateRgbPng(2, 2), 12, 60, imageStyle, linkUri: "https://example.com/")
        });
        if (fragment) cell = cell.WithViewport(new PdfTableCellViewport(200, 72, 80, 72, offsetX: 120));
        var rows = new[] { new[] { cell } };
        var options = new PdfOptions { PageWidth = 320, PageHeight = 180,
            MarginTop = 24, MarginBottom = 24, MarginLeft = 24, MarginRight = 24 };
        var style = TableStyles.Minimal();
        style.BorderColor = null; style.BorderWidth = 0; style.HeaderRowCount = 0;
        style.ColumnWidthPoints = new List<double?> { width }; style.FixedRowHeights = new List<double?> { 72 };
        style.CellPaddingX = 0; style.CellPaddingY = 0; style.CellSpacing = 0;
        style.SpacingBefore = 0; style.SpacingAfter = 0;
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Top };
        return PdfDocument.Create(options).Compose(document => document.Page(page => {
            if (mode == "canvas") page.Canvas(canvas => canvas.Table(rows, 24, 24, width, 72, style));
            else page.Content(content => {
                if (mode == "flow") content.Item(item => item.Table(rows, style: style));
                else content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style)));
            });
        })).ToBytes();
    }
}
