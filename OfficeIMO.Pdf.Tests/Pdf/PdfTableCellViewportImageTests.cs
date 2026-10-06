using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfTableCellViewportTests {
    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    [InlineData("canvas")]
    public void Horizontal_image_fragments_preserve_the_full_box_and_intersect_existing_image_clip(string mode) {
        var source = PdfTableCell.WithImages(string.Empty, new[] { new PdfTableCellImage(
            PdfPngTestImages.CreateRgbPng(2, 2), 36, 12,
            new PdfImageStyle { Align = PdfAlign.Center, ClipPath = OfficeIMO.Drawing.OfficeClipPath.Rectangle(36, 6) },
            linkUri: "https://example.com/") });
        byte[] bytes = Render(mode, source.WithViewport(new PdfTableCellViewport(200, 24, 100, 24, offsetX: 100)), PdfCellVerticalAlign.Top);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        var link = Assert.Single(pdf.GetPage(1).GetHyperlinks());
        Assert.InRange(link.Bounds.Left, 23.99, 24.01);
        Assert.InRange(link.Bounds.Right, 41.99, 42.01);
        string operators = System.Text.Encoding.ASCII.GetString(bytes);
        Assert.Contains("24 132 100 24 re W n", operators);
        Assert.Contains("6 150 36 6 re W n", operators);
    }

    [Theory]
    [InlineData("flow", false)]
    [InlineData("column", false)]
    [InlineData("canvas", false)]
    [InlineData("flow", true)]
    [InlineData("column", true)]
    [InlineData("canvas", true)]
    public void Cell_image_is_emitted_only_in_the_fragment_containing_its_full_cell_position(string mode, bool bottomAligned) {
        var source = PdfTableCell.WithImages(string.Empty, new[] {
            new PdfTableCellImage(PdfPngTestImages.CreateRgbPng(2, 2), 12, 12, linkUri: "https://example.com/")
        });
        var upper = source.WithViewport(new PdfTableCellViewport(100, 48, 100, 24));
        var lower = source.WithViewport(new PdfTableCellViewport(100, 48, 100, 24, offsetY: 24));
        var align = bottomAligned ? PdfCellVerticalAlign.Bottom : PdfCellVerticalAlign.Top;
        using PdfPigDocument first = PdfPigDocument.Open(Render(mode, upper, align));
        using PdfPigDocument second = PdfPigDocument.Open(Render(mode, lower, align));
        Assert.Equal(bottomAligned ? 0 : 1, first.GetPage(1).GetImages().Count());
        Assert.Equal(bottomAligned ? 1 : 0, second.GetPage(1).GetImages().Count());
        Assert.Equal(bottomAligned ? 0 : 1, first.GetPage(1).GetHyperlinks().Count());
        Assert.Equal(bottomAligned ? 1 : 0, second.GetPage(1).GetHyperlinks().Count());
        Assert.Single(lower.WithNoWrap().Images);
    }

    [Theory]
    [InlineData("flow")]
    [InlineData("column")]
    [InlineData("canvas")]
    public void Partly_visible_cell_image_retains_full_size_and_clips_its_draw_and_link(string mode) {
        var source = PdfTableCell.WithImages(string.Empty, new[] {
            new PdfTableCellImage(PdfPngTestImages.CreateRgbPng(2, 2), 36, 36, linkUri: "https://example.com/")
        });
        byte[] bytes = Render(mode, source.WithViewport(new PdfTableCellViewport(100, 48, 100, 24, offsetY: 24)), PdfCellVerticalAlign.Top);
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Single(pdf.GetPage(1).GetImages());
        var image = Assert.Single(PdfImageExtractor.ExtractImagePlacements(bytes));
        Assert.Equal(36, image.Width, 3);
        Assert.Equal(36, image.Height, 3);
        var link = Assert.Single(pdf.GetPage(1).GetHyperlinks());
        Assert.InRange(link.Bounds.Bottom, 143.99, 144.01);
        Assert.InRange(link.Bounds.Top, 155.99, 156.01);
        Assert.Contains("24 132 100 24 re W n", System.Text.Encoding.ASCII.GetString(bytes));
    }
}
