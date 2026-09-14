using OfficeIMO.Drawing;
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

        Assert.True(builder.TryCreateClipPath(OfficeFillRule.NonZero, out PdfPageClipPath clip));
        Assert.True(clip.IsRectangle);
        Assert.Equal(25D, clip.X);
        Assert.Equal(13D, clip.Y);
        Assert.Equal(60D, clip.Width);
        Assert.Equal(120D, clip.Height);
        Assert.Empty(clip.Commands);

        Assert.True(PdfPageClipPathBuilder.TryCreateTransformedRectangle(
            new Matrix2D(2D, 0D, 0D, 3D, 5D, 7D),
            x: 10D,
            y: 20D,
            width: 30D,
            height: 40D,
            pageHeight: 200D,
            fillRule: OfficeFillRule.NonZero,
            out PdfPageClipPath direct));
        AssertClipPathsEqual(clip, direct);
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

        Assert.True(builder.TryCreateClipPath(OfficeFillRule.NonZero, out PdfPageClipPath clip));
        Assert.False(clip.IsRectangle);
        Assert.Equal(5, clip.Commands.Count);

        Assert.True(PdfPageClipPathBuilder.TryCreateTransformedRectangle(
            new Matrix2D(1D, 0.25D, 0D, 1D, 0D, 0D),
            x: 10D,
            y: 20D,
            width: 30D,
            height: 40D,
            pageHeight: 200D,
            fillRule: OfficeFillRule.NonZero,
            out PdfPageClipPath direct));
        AssertClipPathsEqual(clip, direct);
    }

    [Fact]
    public void GenericRectanglePath_UsesCompactRectangleRepresentation() {
        OfficePathCommand[] commands = {
            OfficePathCommand.MoveTo(new OfficePoint(10D, 20D)),
            OfficePathCommand.LineTo(new OfficePoint(70D, 20D)),
            OfficePathCommand.LineTo(new OfficePoint(70D, 20D)),
            OfficePathCommand.LineTo(new OfficePoint(70D, 80D)),
            OfficePathCommand.LineTo(new OfficePoint(10D, 80D)),
            OfficePathCommand.LineTo(new OfficePoint(10D, 20D)),
            OfficePathCommand.Close()
        };

        Assert.True(PdfPageClipPath.TryCreatePath(commands, OfficeFillRule.NonZero, out PdfPageClipPath clip));
        Assert.True(clip.IsRectangle);
        Assert.Equal(10D, clip.X);
        Assert.Equal(20D, clip.Y);
        Assert.Equal(60D, clip.Width);
        Assert.Equal(60D, clip.Height);
        Assert.Empty(clip.Commands);
    }

    [Fact]
    public void GenericNonRectanglePath_PreservesPathRepresentation() {
        OfficePathCommand[] commands = {
            OfficePathCommand.MoveTo(new OfficePoint(10D, 20D)),
            OfficePathCommand.LineTo(new OfficePoint(70D, 20D)),
            OfficePathCommand.LineTo(new OfficePoint(60D, 80D)),
            OfficePathCommand.LineTo(new OfficePoint(10D, 80D)),
            OfficePathCommand.Close()
        };

        Assert.True(PdfPageClipPath.TryCreatePath(commands, OfficeFillRule.NonZero, out PdfPageClipPath clip));
        Assert.False(clip.IsRectangle);
        Assert.Equal(commands.Length, clip.Commands.Count);
    }

    [Fact]
    public void GenericOrthogonalPathWithInteriorDetour_PreservesPathRepresentation() {
        OfficePathCommand[] commands = {
            OfficePathCommand.MoveTo(new OfficePoint(10D, 20D)),
            OfficePathCommand.LineTo(new OfficePoint(70D, 20D)),
            OfficePathCommand.LineTo(new OfficePoint(70D, 80D)),
            OfficePathCommand.LineTo(new OfficePoint(10D, 80D)),
            OfficePathCommand.LineTo(new OfficePoint(10D, 50D)),
            OfficePathCommand.LineTo(new OfficePoint(40D, 50D)),
            OfficePathCommand.LineTo(new OfficePoint(10D, 50D)),
            OfficePathCommand.Close()
        };

        Assert.True(PdfPageClipPath.TryCreatePath(commands, OfficeFillRule.NonZero, out PdfPageClipPath clip));
        Assert.False(clip.IsRectangle);
        Assert.Equal(commands.Length, clip.Commands.Count);
    }

    [Fact]
    public void TinyPerimeterPathThatBacktracksEntireRectangle_PreservesPathRepresentation() {
        OfficePathCommand[] commands = {
            OfficePathCommand.MoveTo(new OfficePoint(0D, 0D)),
            OfficePathCommand.LineTo(new OfficePoint(0.000001D, 0D)),
            OfficePathCommand.LineTo(new OfficePoint(0.000001D, 0.000001D)),
            OfficePathCommand.LineTo(new OfficePoint(0D, 0.000001D)),
            OfficePathCommand.LineTo(new OfficePoint(0D, 0D)),
            OfficePathCommand.LineTo(new OfficePoint(0D, 0.000001D)),
            OfficePathCommand.LineTo(new OfficePoint(0.000001D, 0.000001D)),
            OfficePathCommand.LineTo(new OfficePoint(0.000001D, 0D)),
            OfficePathCommand.LineTo(new OfficePoint(0D, 0D)),
            OfficePathCommand.Close()
        };

        Assert.True(PdfPageClipPath.TryCreatePath(commands, OfficeFillRule.NonZero, out PdfPageClipPath clip));
        Assert.False(clip.IsRectangle);
        Assert.Equal(commands.Length, clip.Commands.Count);
    }

    [Fact]
    public void FarFromOriginGenericRectangle_UsesCompactRectangleRepresentation() {
        const double offset = 1_000_000_000_000D;
        OfficePathCommand[] commands = {
            OfficePathCommand.MoveTo(new OfficePoint(offset + 10D, offset + 20D)),
            OfficePathCommand.LineTo(new OfficePoint(offset + 70D, offset + 20D)),
            OfficePathCommand.LineTo(new OfficePoint(offset + 70D, offset + 80D)),
            OfficePathCommand.LineTo(new OfficePoint(offset + 10D, offset + 80D)),
            OfficePathCommand.Close()
        };

        Assert.True(PdfPageClipPath.TryCreatePath(commands, OfficeFillRule.NonZero, out PdfPageClipPath clip));
        Assert.True(clip.IsRectangle);
        Assert.Equal(offset + 10D, clip.X);
        Assert.Equal(offset + 20D, clip.Y);
        Assert.Equal(60D, clip.Width);
        Assert.Equal(60D, clip.Height);
    }

    private static void AssertClipPathsEqual(PdfPageClipPath expected, PdfPageClipPath actual) {
        Assert.Equal(expected.IsRectangle, actual.IsRectangle);
        Assert.Equal(expected.FillRule, actual.FillRule);
        Assert.Equal(expected.X, actual.X);
        Assert.Equal(expected.Y, actual.Y);
        Assert.Equal(expected.Width, actual.Width);
        Assert.Equal(expected.Height, actual.Height);
        Assert.Equal(expected.Commands.Count, actual.Commands.Count);
        for (int i = 0; i < expected.Commands.Count; i++) {
            OfficePathCommand expectedCommand = expected.Commands[i];
            OfficePathCommand actualCommand = actual.Commands[i];
            Assert.Equal(expectedCommand.Kind, actualCommand.Kind);
            Assert.Equal(expectedCommand.Point, actualCommand.Point);
            Assert.Equal(expectedCommand.ControlPoint1, actualCommand.ControlPoint1);
            Assert.Equal(expectedCommand.ControlPoint2, actualCommand.ControlPoint2);
        }
    }
}
