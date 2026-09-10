using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingRasterTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AnchoredTextLinesMatchTheTextBlockSvgBaseline(bool affine) {
        var fonts = new OfficeFontFaceCollection().Add("FixtureFont",
            File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "RobotoFlex.ttf")));
        var expected = new OfficeRasterImage(180, 60, OfficeColor.White);
        var reference = new OfficeRasterCanvas(expected, fonts: fonts);
        double x = affine ? 12D : 8D, y = affine ? 11D : 8D;
        reference.DrawPositionedText("Baseline", x, y, 140D, 40D, OfficeColor.Black, 20D,
            OfficeTextAlignment.Left, OfficeFontStyle.Regular, "FixtureFont", reference.MeasureText("Baseline", 20D, "FixtureFont"),
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, baselineFontSize: 20D * 0.84D);
        var actual = new OfficeRasterImage(180, 60, OfficeColor.White);
        var canvas = new OfficeRasterCanvas(actual, fonts: fonts);
        if (affine) canvas.DrawTextLineTransformed("Baseline", 8D, 8D, 20D, OfficeColor.Black,
            OfficeTransform.Translate(4D, 3D), alignment: OfficeTextAlignment.Left, fontFamily: "FixtureFont");
        else canvas.DrawTextLine("Baseline", 8D, 8D, 20D, OfficeColor.Black, alignment: OfficeTextAlignment.Left, fontFamily: "FixtureFont");
        AssertRasterImagesEqual(expected, actual);
    }
}
