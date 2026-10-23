using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlShapeSnapshotReuseTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ShapeProjectionsKeepDetachedPublicGeometryAndFragmentCoordinates(bool paintOnly) {
        var input = OfficeShape.Rectangle(12D, 6D);
        input.FillColor = OfficeColor.Red;
        var source = new HtmlRenderShape(input, 2D, 4D, 1, source: "shape", layoutY: 8D,
            isCollapsedTableBorder: true, collapsedBorderBoundaryY: 12D, canExtendFragmentBackground: true);
        input.Width = 99D;
        input.FillColor = OfficeColor.Blue;
        var projected = (HtmlRenderShape)(paintOnly ? source.TranslatePaint(5D, 3D, 2) : source.Translate(5D, 3D, 2));
        OfficeShape exposed = projected.Shape;
        exposed.Width = 77D;
        exposed.FillColor = OfficeColor.Black;
        Assert.Equal(12D, source.Shape.Width);
        Assert.Equal(12D, projected.Shape.Width);
        Assert.Equal(OfficeColor.Red, source.Shape.FillColor);
        Assert.Equal(OfficeColor.Red, projected.Shape.FillColor);
        Assert.Equal(7D, projected.X);
        Assert.Equal(7D, projected.Y);
        Assert.Equal(paintOnly ? 8D : 11D, projected.LayoutY);
        Assert.Equal(paintOnly ? 12D : 15D, projected.CollapsedBorderBoundaryY);
        Assert.True(projected.IsCollapsedTableBorder);
        Assert.True(projected.CanExtendFragmentBackground);
    }
}
