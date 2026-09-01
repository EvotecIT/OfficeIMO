using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfLogicalPageCoordinateTests {
    [Fact]
    public void MapVisualRectangleToUserSpace_MapsTopLeftCoordinatesOnUnrotatedPage() {
        byte[] source = PdfDocument.Create(compose => compose.Page(page => page.Size(600D, 800D))).ToBytes();
        PdfLogicalPage page = Assert.Single(PdfDocument.Load(source).Read().Pages);

        PdfPageRectangle rectangle = page.MapVisualRectangleToUserSpace(40D, 50D, 180D, 100D);

        Assert.Equal(40D, rectangle.Left);
        Assert.Equal(700D, rectangle.Bottom);
        Assert.Equal(180D, rectangle.Right);
        Assert.Equal(750D, rectangle.Top);
    }

    [Fact]
    public void MapVisualRectangleToUserSpace_AccountsForCropBoxOrigin() {
        byte[] source = PdfDocument.Create(compose => compose.Page(page => page.Size(600D, 800D))).ToBytes();
        byte[] cropped = PdfDocument.Load(source).Pages.SetCropBox(100D, 100D, 500D, 700D, 1).ToBytes();
        PdfLogicalPage page = Assert.Single(PdfDocument.Load(cropped).Read().Pages);

        PdfPageRectangle rectangle = page.MapVisualRectangleToUserSpace(40D, 50D, 180D, 100D);

        Assert.Equal(140D, rectangle.Left);
        Assert.Equal(600D, rectangle.Bottom);
        Assert.Equal(280D, rectangle.Right);
        Assert.Equal(650D, rectangle.Top);
    }

    [Fact]
    public void MapVisualRectangleToUserSpace_AccountsForInheritedPageRotation() {
        byte[] source = PdfDocument.Create(compose => compose.Page(page => page.Size(600D, 800D))).ToBytes();
        byte[] rotated = PdfDocument.Load(source).Pages.Rotate(90, 1).ToBytes();
        PdfLogicalPage page = Assert.Single(PdfDocument.Load(rotated).Read().Pages);

        PdfPageRectangle rectangle = page.MapVisualRectangleToUserSpace(40D, 50D, 180D, 100D);

        Assert.Equal(500D, rectangle.Left);
        Assert.Equal(620D, rectangle.Bottom);
        Assert.Equal(550D, rectangle.Right);
        Assert.Equal(760D, rectangle.Top);
    }

    [Fact]
    public void MapVisualPointToUserSpace_UsesTheSameRotatedCoordinateContract() {
        byte[] source = PdfDocument.Create(compose => compose.Page(page => page.Size(600D, 800D))).ToBytes();
        byte[] rotated = PdfDocument.Load(source).Pages.Rotate(90, 1).ToBytes();
        PdfLogicalPage page = Assert.Single(PdfDocument.Load(rotated).Read().Pages);

        PdfPagePoint point = page.MapVisualPointToUserSpace(40D, 50D);

        Assert.Equal(550D, point.X);
        Assert.Equal(760D, point.Y);
    }
}
