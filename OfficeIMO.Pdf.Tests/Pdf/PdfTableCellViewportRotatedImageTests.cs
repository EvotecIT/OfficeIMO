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
    [InlineData("flow")]
    [InlineData("column")]
    [InlineData("canvas")]
    public void Rotated_ink_outside_the_existing_image_clip_is_omitted_with_its_link(string mode) {
        var imageStyle = new PdfImageStyle { Align = PdfAlign.Center, RotationAngle = 90,
            ClipPath = OfficeClipPath.Rectangle(12, 12) };
        using PdfPigDocument fragment = PdfPigDocument.Open(RenderRotatedImage(mode, imageStyle));
        Assert.Empty(fragment.GetPage(1).GetImages());
        Assert.Empty(fragment.GetPage(1).GetHyperlinks());
    }

    [Theory]
    [InlineData("flow", false)]
    [InlineData("column", false)]
    [InlineData("canvas", false)]
    [InlineData("flow", true)]
    [InlineData("column", true)]
    [InlineData("canvas", true)]
    public void Source_crop_keeps_the_authored_rotation_pivot_and_rotated_link_bounds(string mode, bool top) {
        var imageStyle = new PdfImageStyle { Align = PdfAlign.Center, RotationAngle = 90,
            SourceCrop = top ? new PdfImageSourceCrop { Top = .5 } : new PdfImageSourceCrop { Bottom = .5 } };
        using PdfPigDocument fragment = PdfPigDocument.Open(RenderRotatedImage(mode, imageStyle));
        Assert.Single(fragment.GetPage(1).GetImages());
        var link = Assert.Single(fragment.GetPage(1).GetHyperlinks());
        Assert.Equal(24, link.Bounds.Left, 3);
        // Either source half occupies the same authored 12 by 60 point frame.
        // Rotation therefore preserves the same physical fragment and link rectangle.
        Assert.Equal(34, link.Bounds.Right, 3);
        Assert.Equal(120, link.Bounds.Bottom, 3);
        Assert.Equal(132, link.Bounds.Top, 3);
    }

    private static byte[] RenderRotatedImage(string mode, PdfImageStyle imageStyle, bool fragment = true) {
        double width = fragment ? 80 : 200;
        var cell = PdfTableCell.WithImages(string.Empty, new[] {
            new PdfTableCellImage(PdfPngTestImages.CreateRgbPng(2, 2), 12, 60, imageStyle, linkUri: "https://example.com/")
        });
        if (fragment) cell = cell.WithViewport(new PdfTableCellViewport(200, 72, 80, 72, offsetX: 120));
        var rows = new[] { new[] { cell } };
        var options = new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(320), PageHeight = global::OfficeIMO.OfficeLength.Points(180),
            MarginTop = global::OfficeIMO.OfficeLength.Points(24), MarginBottom = global::OfficeIMO.OfficeLength.Points(24), MarginLeft = global::OfficeIMO.OfficeLength.Points(24), MarginRight = global::OfficeIMO.OfficeLength.Points(24) };
        var style = PdfTableStyles.Minimal();
        style.BorderColor = null; style.BorderWidth = global::OfficeIMO.OfficeLength.Points(0); style.HeaderRowCount = 0;
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(width) }; style.FixedRowHeights = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(72) };
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(0); style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(0); style.CellSpacing = global::OfficeIMO.OfficeLength.Points(0);
        style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(0); style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(0);
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Top };
        return PdfDocument.Create(options).Compose(document => document.Page(page => {
            if (mode == "canvas") page.Canvas(canvas => canvas.Table(rows, global::OfficeIMO.OfficeLength.Points(24), global::OfficeIMO.OfficeLength.Points(24), global::OfficeIMO.OfficeLength.Points(width), global::OfficeIMO.OfficeLength.Points(72), style));
            else page.Content(content => {
                if (mode == "flow") content.Item(item => item.Table(rows, style: style));
                else content.Row(row => row.PercentColumn(100, column => column.Table(rows, style: style)));
            });
        })).ToBytes();
    }
}
