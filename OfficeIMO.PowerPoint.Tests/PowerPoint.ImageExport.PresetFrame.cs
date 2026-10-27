using System;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests;

public partial class PowerPointImageExportTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PresetFrameFlipsAsymmetricGeometryExactlyOnce(bool flipH) {
        using PowerPointDocument presentation = PowerPointDocument.Create();
        presentation.SlideSize.SetSizePoints(120, 90);
        PowerPointSlide slide = presentation.AddSlide();
        slide.BackgroundColor = "FFFFFF";
        PowerPointAutoShape shape = slide.AddShapePoints(OfficePresetShapeType.RightTriangle, 20, 20, 80, 40);
        shape.FillColor = "000000";
        shape.HorizontalFlip = flipH;
        Shape raw = Assert.IsType<Shape>(shape.Element);
        raw.ShapeProperties!.RemoveAllChildren<A.Outline>();
        raw.ShapeProperties.Append(new A.Outline(new A.NoFill()));
        OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Png,
            new PowerPointImageExportOptions { Scale = 1D });
        Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
        Assert.True(image!.GetPixel(flipH ? 90 : 30, 35).R < 20);
        Assert.True(image.GetPixel(flipH ? 30 : 90, 35).R > 235);
    }

    [Theory]
    [InlineData(false, false, 0, 0, true, false)]
    [InlineData(false, false, 0, 0, true, true)]
    [InlineData(true, false, 0, 0, true, false)]
    [InlineData(true, false, 0, 0, true, true)]
    [InlineData(false, true, 90, 0, true, false)]
    [InlineData(false, true, 90, 0, true, true)]
    [InlineData(true, true, 0, 0, true, false)]
    [InlineData(true, true, 0, 0, true, true)]
    [InlineData(true, false, 0, 90, true, false)]
    [InlineData(true, false, 0, 90, true, true)]
    [InlineData(true, false, 0, 0, false, false)]
    [InlineData(true, false, 0, 0, false, true)]
    public void PresetFrameTransformsPaintTogetherWithGeometry(bool flipH, bool flipV, int gradientAngle,
        int rotation, bool rotateWithShape, bool textBox) {
        using PowerPointDocument presentation = PowerPointDocument.Create();
        presentation.SlideSize.SetSizePoints(120, 90);
        PowerPointSlide slide = presentation.AddSlide();
        slide.BackgroundColor = "FFFFFF";
        PowerPointShape shape = textBox ? slide.AddTextBoxPoints("", 20, 20, 80, 40) : slide.AddRectanglePoints(20, 20, 80, 40);
        shape.HorizontalFlip = flipH;
        shape.VerticalFlip = flipV;
        shape.Rotation = rotation;
        Shape raw = Assert.IsType<Shape>(shape.Element);
        raw.ShapeProperties!.RemoveAllChildren<A.SolidFill>();
        raw.ShapeProperties.RemoveAllChildren<A.Outline>();
        raw.ShapeProperties.Append(new A.GradientFill(
            new A.GradientStopList(
                new A.GradientStop(new A.RgbColorModelHex { Val = "FF0000" }) { Position = 0 },
                new A.GradientStop(new A.RgbColorModelHex { Val = "0000FF" }) { Position = 100000 }),
            new A.LinearGradientFill { Angle = gradientAngle * 60000, Scaled = true }) {
                RotateWithShape = rotateWithShape
            });
        raw.ShapeProperties.Append(new A.Outline(new A.NoFill()));

        OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Png,
            new PowerPointImageExportOptions { Scale = 1D });
        Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
        OfficePoint first = gradientAngle == 0 ? new OfficePoint(10, 20) : new OfficePoint(40, 5);
        OfficePoint last = gradientAngle == 0 ? new OfficePoint(70, 20) : new OfficePoint(40, 35);
        if (rotateWithShape) {
            OfficeTransform frame = new OfficeImageFrameTransform(rotation, 40, 20, flipH, flipV).CreateDestinationTransform();
            first = frame.TransformPoint(first);
            last = frame.TransformPoint(last);
        }
        OfficeColor redEnd = image!.GetPixel((int)Math.Round(20 + first.X), (int)Math.Round(20 + first.Y));
        OfficeColor blueEnd = image.GetPixel((int)Math.Round(20 + last.X), (int)Math.Round(20 + last.Y));
        Assert.True(redEnd.R > redEnd.B + 100, "The authored red end should follow the complete shape frame.");
        Assert.True(blueEnd.B > blueEnd.R + 100, "The authored blue end should follow the complete shape frame.");
    }
}
