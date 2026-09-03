using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPageClipPathBuilderTests {
    [Fact]
    public void AxisAlignedRectangle_UsesCompactRectangleRepresentation() {
        var builder = new PdfPageClipPathBuilder(pageHeight: 200D);

        builder.AddRectanglePath(
            new Matrix2D(2D, 0D, 0D, 3D, 5D, 7D),
            x: 10D,
            y: 20D,
            width: 30D,
            height: 40D);

        Assert.True(builder.TryCreateClipPath(OfficeIMO.Drawing.OfficeFillRule.NonZero, out PdfPageClipPath clip));
        Assert.True(clip.IsRectangle);
        Assert.Equal(25D, clip.X);
        Assert.Equal(13D, clip.Y);
        Assert.Equal(60D, clip.Width);
        Assert.Equal(120D, clip.Height);
        Assert.Empty(clip.Commands);
    }

    [Fact]
    public void ShearedRectangle_PreservesPathRepresentation() {
        var builder = new PdfPageClipPathBuilder(pageHeight: 200D);

        builder.AddRectanglePath(
            new Matrix2D(1D, 0.25D, 0D, 1D, 0D, 0D),
            x: 10D,
            y: 20D,
            width: 30D,
            height: 40D);

        Assert.True(builder.TryCreateClipPath(OfficeIMO.Drawing.OfficeFillRule.NonZero, out PdfPageClipPath clip));
        Assert.False(clip.IsRectangle);
        Assert.Equal(5, clip.Commands.Count);
    }
}
