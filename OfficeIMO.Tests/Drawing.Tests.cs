using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingTests {
    private static readonly byte[] OnePixelPng = {
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
        0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41,
        0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
        0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
        0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
        0x42, 0x60, 0x82
    };

    [Fact]
    public void OfficeColorParsesNamedAndHexValues() {
        Assert.Equal(OfficeColor.Red, OfficeColor.Parse("red"));
        Assert.Equal(OfficeColor.FromRgb(0x66, 0x33, 0x99), OfficeColor.Parse("RebeccaPurple"));

        var color = OfficeColor.Parse("#336699CC");
        Assert.Equal(0x33, color.R);
        Assert.Equal(0x66, color.G);
        Assert.Equal(0x99, color.B);
        Assert.Equal(0xCC, color.A);
        Assert.Equal("336699CC", color.ToHex());
        Assert.Equal("336699", color.ToRgbHex());
        Assert.Equal("CC336699", color.ToArgbHex());
    }

    [Theory]
    [MemberData(nameof(SixLaborsNamedColorNames))]
    public void OfficeColorParsesSixLaborsNamedColors(string name) {
        Assert.True(OfficeColor.TryParse(name, out _), $"OfficeColor should parse ImageSharp named color '{name}'.");
    }

    public static IEnumerable<object[]> SixLaborsNamedColorNames() {
        foreach (var name in new[] {
            "AliceBlue",
            "AntiqueWhite",
            "Aqua",
            "Aquamarine",
            "Azure",
            "Beige",
            "Bisque",
            "Black",
            "BlanchedAlmond",
            "Blue",
            "BlueViolet",
            "Brown",
            "BurlyWood",
            "CadetBlue",
            "Chartreuse",
            "Chocolate",
            "Coral",
            "CornflowerBlue",
            "Cornsilk",
            "Crimson",
            "Cyan",
            "DarkBlue",
            "DarkCyan",
            "DarkGoldenrod",
            "DarkGray",
            "DarkGreen",
            "DarkGrey",
            "DarkKhaki",
            "DarkMagenta",
            "DarkOliveGreen",
            "DarkOrange",
            "DarkOrchid",
            "DarkRed",
            "DarkSalmon",
            "DarkSeaGreen",
            "DarkSlateBlue",
            "DarkSlateGray",
            "DarkSlateGrey",
            "DarkTurquoise",
            "DarkViolet",
            "DeepPink",
            "DeepSkyBlue",
            "DimGray",
            "DimGrey",
            "DodgerBlue",
            "Firebrick",
            "FloralWhite",
            "ForestGreen",
            "Fuchsia",
            "Gainsboro",
            "GhostWhite",
            "Gold",
            "Goldenrod",
            "Gray",
            "Green",
            "GreenYellow",
            "Grey",
            "Honeydew",
            "HotPink",
            "IndianRed",
            "Indigo",
            "Ivory",
            "Khaki",
            "Lavender",
            "LavenderBlush",
            "LawnGreen",
            "LemonChiffon",
            "LightBlue",
            "LightCoral",
            "LightCyan",
            "LightGoldenrodYellow",
            "LightGray",
            "LightGreen",
            "LightGrey",
            "LightPink",
            "LightSalmon",
            "LightSeaGreen",
            "LightSkyBlue",
            "LightSlateGray",
            "LightSlateGrey",
            "LightSteelBlue",
            "LightYellow",
            "Lime",
            "LimeGreen",
            "Linen",
            "Magenta",
            "Maroon",
            "MediumAquamarine",
            "MediumBlue",
            "MediumOrchid",
            "MediumPurple",
            "MediumSeaGreen",
            "MediumSlateBlue",
            "MediumSpringGreen",
            "MediumTurquoise",
            "MediumVioletRed",
            "MidnightBlue",
            "MintCream",
            "MistyRose",
            "Moccasin",
            "NavajoWhite",
            "Navy",
            "OldLace",
            "Olive",
            "OliveDrab",
            "Orange",
            "OrangeRed",
            "Orchid",
            "PaleGoldenrod",
            "PaleGreen",
            "PaleTurquoise",
            "PaleVioletRed",
            "PapayaWhip",
            "PeachPuff",
            "Peru",
            "Pink",
            "Plum",
            "PowderBlue",
            "Purple",
            "RebeccaPurple",
            "Red",
            "RosyBrown",
            "RoyalBlue",
            "SaddleBrown",
            "Salmon",
            "SandyBrown",
            "SeaGreen",
            "SeaShell",
            "Sienna",
            "Silver",
            "SkyBlue",
            "SlateBlue",
            "SlateGray",
            "SlateGrey",
            "Snow",
            "SpringGreen",
            "SteelBlue",
            "Tan",
            "Teal",
            "Thistle",
            "Tomato",
            "Transparent",
            "Turquoise",
            "Violet",
            "Wheat",
            "White",
            "WhiteSmoke",
            "Yellow",
            "YellowGreen"
        }) {
            yield return new object[] { name };
        }
    }

    [Fact]
    public void OfficeColorTransparentUsesZeroRgbWithZeroAlpha() {
        Assert.Equal(0, OfficeColor.Transparent.R);
        Assert.Equal(0, OfficeColor.Transparent.G);
        Assert.Equal(0, OfficeColor.Transparent.B);
        Assert.Equal(0, OfficeColor.Transparent.A);
    }

    [Fact]
    public void OfficeFontInfoStoresFamilySizeAndStyleWithoutFontEngineDependency() {
        var font = new OfficeFontInfo("Aptos", 12.5, OfficeFontStyle.Bold | OfficeFontStyle.Italic);

        Assert.Equal("Aptos", font.FamilyName);
        Assert.Equal(12.5, font.Size);
        Assert.True(font.IsBold);
        Assert.True(font.IsItalic);
        Assert.False(font.IsUnderline);
        Assert.Equal("Aptos, 12.5pt, Bold, Italic", font.ToString());
    }

    [Fact]
    public void OfficeFontInfoStoresUnderlineStyle() {
        var font = new OfficeFontInfo("Calibri", 11, OfficeFontStyle.Underline);

        Assert.False(font.IsBold);
        Assert.False(font.IsItalic);
        Assert.True(font.IsUnderline);
        Assert.Equal("Calibri, 11pt, Underline", font.ToString());
    }

    [Fact]
    public void OfficeFontInfoProvidesImmutableCopyHelpers() {
        var font = OfficeFontInfo.Default
            .WithFamilyName("Arial")
            .WithSize(10)
            .WithStyle(OfficeFontStyle.Bold);

        Assert.Equal(new OfficeFontInfo("Arial", 10, OfficeFontStyle.Bold), font);
        Assert.NotEqual(OfficeFontInfo.Default, font);
    }

    [Fact]
    public void OfficeTransformProvidesReusableAffineDrawingIntent() {
        var rotated = OfficeTransform.RotateDegrees(90).TransformPoint(new OfficePoint(10, 0));

        Assert.Equal(new OfficePoint(0, 10), rotated);

        var composed = OfficeTransform.Translate(5, 10)
            .Then(OfficeTransform.Scale(2, 3))
            .TransformPoint(new OfficePoint(4, 5));

        Assert.Equal(new OfficePoint(18, 45), composed);
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeTransform.Translate(double.NaN, 0));
    }

    [Fact]
    public void OfficeShapeStoresReusableRectangleDrawingIntent() {
        var shape = OfficeShape.Rectangle(160, 48);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.SteelBlue, OfficeColor.WhiteSmoke);
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.5;
        shape.StrokeDashStyle = OfficeStrokeDashStyle.Dash;
        shape.Shadow = new OfficeShadow(OfficeColor.Black, 0.18, 3, 4);
        shape.FillOpacity = 0.45;
        shape.StrokeOpacity = 0.8;
        shape.Transform = OfficeTransform.Translate(4, 8);
        shape.ClipPath = OfficeClipPath.Rectangle(80, 24);

        var clone = shape.Clone();
        shape.Width = 10;
        shape.FillOpacity = 1;
        shape.FillGradient = OfficeLinearGradient.Vertical(OfficeColor.Red, OfficeColor.Black);
        shape.Shadow = new OfficeShadow(OfficeColor.Red, 0.9, 1, 1);
        shape.Transform = OfficeTransform.Identity;
        shape.ClipPath = OfficeClipPath.Rectangle(10, 10);

        Assert.Equal(OfficeShapeKind.Rectangle, clone.Kind);
        Assert.Equal(160, clone.Width);
        Assert.Equal(48, clone.Height);
        Assert.Equal(OfficeColor.WhiteSmoke, clone.FillColor);
        Assert.NotNull(clone.FillGradient);
        Assert.Equal(0, clone.FillGradient!.StartX);
        Assert.Equal(0.5, clone.FillGradient.StartY);
        Assert.Equal(1, clone.FillGradient.EndX);
        Assert.Equal(0.5, clone.FillGradient.EndY);
        Assert.Equal(OfficeColor.SteelBlue, clone.FillGradient.Stops[0].Color);
        Assert.Equal(OfficeColor.WhiteSmoke, clone.FillGradient.Stops[1].Color);
        Assert.NotNull(clone.Shadow);
        Assert.Equal(OfficeColor.Black, clone.Shadow!.Color);
        Assert.Equal(0.18, clone.Shadow.Opacity);
        Assert.Equal(3, clone.Shadow.OffsetX);
        Assert.Equal(4, clone.Shadow.OffsetY);
        Assert.Equal(OfficeColor.SteelBlue, clone.StrokeColor);
        Assert.Equal(1.5, clone.StrokeWidth);
        Assert.Equal(OfficeStrokeDashStyle.Dash, clone.StrokeDashStyle);
        Assert.Equal(0.45, clone.FillOpacity);
        Assert.Equal(0.8, clone.StrokeOpacity);
        Assert.Equal(OfficeTransform.Translate(4, 8), clone.Transform);
        Assert.NotNull(clone.ClipPath);
        Assert.Equal(OfficeClipPathKind.Rectangle, clone.ClipPath!.Kind);
        Assert.Equal(80, clone.ClipPath.Width);
        Assert.Equal(24, clone.ClipPath.Height);
    }

    [Fact]
    public void OfficeClipPathStoresReusablePathIntent() {
        var clipPath = OfficeClipPath.Path(
            OfficePathCommand.MoveTo(10, 30),
            OfficePathCommand.LineTo(50, 0),
            OfficePathCommand.LineTo(90, 30),
            OfficePathCommand.Close());

        var clone = clipPath.Clone();

        Assert.Equal(OfficeClipPathKind.Path, clone.Kind);
        Assert.Equal(80, clone.Width);
        Assert.Equal(30, clone.Height);
        Assert.Equal(4, clone.Commands.Count);
        Assert.Equal(OfficePathCommand.MoveTo(0, 30), clone.Commands[0]);
        Assert.Equal(OfficePathCommand.LineTo(40, 0), clone.Commands[1]);
        Assert.Equal(OfficePathCommand.LineTo(80, 30), clone.Commands[2]);
        Assert.Equal(OfficePathCommand.Close(), clone.Commands[3]);
        Assert.Throws<ArgumentException>(() => OfficeClipPath.Path(OfficePathCommand.LineTo(10, 10)));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeClipPath.Rectangle(double.NaN, 10));
    }

    [Fact]
    public void OfficeLinearGradientStoresReusableTwoStopFillIntent() {
        var gradient = OfficeLinearGradient.DiagonalDown(OfficeColor.SteelBlue, OfficeColor.WhiteSmoke);

        OfficeLinearGradient clone = gradient.Clone();

        Assert.Equal(0, clone.StartX);
        Assert.Equal(0, clone.StartY);
        Assert.Equal(1, clone.EndX);
        Assert.Equal(1, clone.EndY);
        Assert.Equal(new OfficeGradientStop(0, OfficeColor.SteelBlue), clone.Stops[0]);
        Assert.Equal(new OfficeGradientStop(1, OfficeColor.WhiteSmoke), clone.Stops[1]);
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeGradientStop(double.NaN, OfficeColor.Black));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeLinearGradient(-0.1, 0, 1, 1, new OfficeGradientStop(0, OfficeColor.Black), new OfficeGradientStop(1, OfficeColor.White)));
        Assert.Throws<ArgumentException>(() => new OfficeLinearGradient(0, 0, 0, 0, new OfficeGradientStop(0, OfficeColor.Black), new OfficeGradientStop(1, OfficeColor.White)));
        Assert.Throws<ArgumentException>(() => new OfficeLinearGradient(0, 0, 1, 1, new OfficeGradientStop(0.25, OfficeColor.Black), new OfficeGradientStop(1, OfficeColor.White)));
        Assert.Throws<ArgumentException>(() => new OfficeLinearGradient(0, 0, 1, 1, new OfficeGradientStop(0, OfficeColor.Black), new OfficeGradientStop(0.75, OfficeColor.White)));
    }

    [Fact]
    public void OfficeShadowStoresReusableShapeEffectIntent() {
        var shadow = new OfficeShadow(OfficeColor.FromRgb(10, 20, 30), 0.35, 4, 6);

        OfficeShadow clone = shadow.Clone();

        Assert.Equal(OfficeColor.FromRgb(10, 20, 30), clone.Color);
        Assert.Equal(0.35, clone.Opacity);
        Assert.Equal(4, clone.OffsetX);
        Assert.Equal(6, clone.OffsetY);
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeShadow(OfficeColor.Black, -0.1, 0, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeShadow(OfficeColor.Black, 1.1, 0, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeShadow(OfficeColor.Black, 0.5, double.NaN, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeShadow(OfficeColor.Black, 0.5, 0, double.PositiveInfinity));
    }

    [Fact]
    public void OfficeShapeStoresReusableRoundedRectangleDrawingIntent() {
        var shape = OfficeShape.RoundedRectangle(160, 48, 8);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.5;
        shape.StrokeDashStyle = OfficeStrokeDashStyle.Dash;

        var clone = shape.Clone();
        shape.CornerRadius = 2;

        Assert.Equal(OfficeShapeKind.RoundedRectangle, clone.Kind);
        Assert.Equal(160, clone.Width);
        Assert.Equal(48, clone.Height);
        Assert.Equal(8, clone.CornerRadius);
        Assert.Equal(OfficeColor.WhiteSmoke, clone.FillColor);
        Assert.Equal(OfficeColor.SteelBlue, clone.StrokeColor);
        Assert.Equal(1.5, clone.StrokeWidth);
        Assert.Equal(OfficeStrokeDashStyle.Dash, clone.StrokeDashStyle);
    }

    [Fact]
    public void OfficeShapeRejectsInvalidRoundedRectangleRadius() {
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeShape.RoundedRectangle(100, 40, -1));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeShape.RoundedRectangle(100, 40, 21));
    }

    [Fact]
    public void OfficeShapeStoresReusableLineDrawingIntent() {
        var shape = OfficeShape.Line(10, 20, 110, 20);
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 2.5;
        shape.StrokeDashStyle = OfficeStrokeDashStyle.DashDot;
        shape.StrokeLineCap = OfficeStrokeLineCap.Round;
        shape.StrokeLineJoin = OfficeStrokeLineJoin.Bevel;

        var clone = shape.Clone();
        shape.Width = 10;
        shape.StrokeLineCap = OfficeStrokeLineCap.Square;

        Assert.Equal(OfficeShapeKind.Line, clone.Kind);
        Assert.Equal(100, clone.Width);
        Assert.Equal(1, clone.Height);
        Assert.Equal(new OfficePoint(0, 0), clone.Points[0]);
        Assert.Equal(new OfficePoint(100, 0), clone.Points[1]);
        Assert.Equal(OfficeColor.SteelBlue, clone.StrokeColor);
        Assert.Equal(2.5, clone.StrokeWidth);
        Assert.Equal(OfficeStrokeDashStyle.DashDot, clone.StrokeDashStyle);
        Assert.Equal(OfficeStrokeLineCap.Round, clone.StrokeLineCap);
        Assert.Equal(OfficeStrokeLineJoin.Bevel, clone.StrokeLineJoin);
    }

    [Fact]
    public void OfficeShapeRejectsEmptyLineDrawingIntent() {
        Assert.Throws<ArgumentException>(() => OfficeShape.Line(10, 20, 10, 20));
    }

    [Fact]
    public void OfficeShapeStoresReusableEllipseDrawingIntent() {
        var shape = OfficeShape.Ellipse(96, 40);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 2;
        shape.StrokeDashStyle = OfficeStrokeDashStyle.Dot;

        var clone = shape.Clone();
        shape.Height = 10;

        Assert.Equal(OfficeShapeKind.Ellipse, clone.Kind);
        Assert.Equal(96, clone.Width);
        Assert.Equal(40, clone.Height);
        Assert.Equal(OfficeColor.WhiteSmoke, clone.FillColor);
        Assert.Equal(OfficeColor.SteelBlue, clone.StrokeColor);
        Assert.Equal(2, clone.StrokeWidth);
        Assert.Equal(OfficeStrokeDashStyle.Dot, clone.StrokeDashStyle);
    }

    [Fact]
    public void OfficeShapeStoresReusablePolygonDrawingIntent() {
        var shape = OfficeShape.Polygon(
            new OfficePoint(10, 20),
            new OfficePoint(50, 0),
            new OfficePoint(90, 20));
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.25;
        shape.StrokeDashStyle = OfficeStrokeDashStyle.DashDot;
        shape.StrokeLineJoin = OfficeStrokeLineJoin.Round;

        var clone = shape.Clone();
        shape.Width = 10;

        Assert.Equal(OfficeShapeKind.Polygon, clone.Kind);
        Assert.Equal(80, clone.Width);
        Assert.Equal(20, clone.Height);
        Assert.Equal(new OfficePoint(0, 20), clone.Points[0]);
        Assert.Equal(new OfficePoint(40, 0), clone.Points[1]);
        Assert.Equal(new OfficePoint(80, 20), clone.Points[2]);
        Assert.Equal(OfficeColor.WhiteSmoke, clone.FillColor);
        Assert.Equal(OfficeColor.SteelBlue, clone.StrokeColor);
        Assert.Equal(1.25, clone.StrokeWidth);
        Assert.Equal(OfficeStrokeDashStyle.DashDot, clone.StrokeDashStyle);
        Assert.Equal(OfficeStrokeLineJoin.Round, clone.StrokeLineJoin);
    }

    [Fact]
    public void OfficeShapeStoresReusablePathDrawingIntent() {
        var shape = OfficeShape.Path(
            OfficePathCommand.MoveTo(10, 50),
            OfficePathCommand.CubicBezierTo(30, 10, 70, 10, 90, 50),
            OfficePathCommand.LineTo(10, 50),
            OfficePathCommand.Close());
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.75;
        shape.StrokeDashStyle = OfficeStrokeDashStyle.Dot;
        shape.StrokeLineCap = OfficeStrokeLineCap.Round;
        shape.StrokeLineJoin = OfficeStrokeLineJoin.Round;

        var clone = shape.Clone();
        shape.Width = 10;

        Assert.Equal(OfficeShapeKind.Path, clone.Kind);
        Assert.Equal(80, clone.Width);
        Assert.Equal(40, clone.Height);
        Assert.Equal(OfficePathCommand.MoveTo(0, 40), clone.PathCommands[0]);
        Assert.Equal(OfficePathCommand.CubicBezierTo(20, 0, 60, 0, 80, 40), clone.PathCommands[1]);
        Assert.Equal(OfficePathCommand.LineTo(0, 40), clone.PathCommands[2]);
        Assert.Equal(OfficePathCommand.Close(), clone.PathCommands[3]);
        Assert.Equal(OfficeColor.WhiteSmoke, clone.FillColor);
        Assert.Equal(OfficeColor.SteelBlue, clone.StrokeColor);
        Assert.Equal(1.75, clone.StrokeWidth);
        Assert.Equal(OfficeStrokeDashStyle.Dot, clone.StrokeDashStyle);
        Assert.Equal(OfficeStrokeLineCap.Round, clone.StrokeLineCap);
        Assert.Equal(OfficeStrokeLineJoin.Round, clone.StrokeLineJoin);
    }

    [Fact]
    public void OfficeDrawingStoresReusablePositionedShapesInPaintOrder() {
        var background = OfficeShape.Rectangle(120, 60);
        background.FillColor = OfficeColor.WhiteSmoke;

        var marker = OfficeShape.Polygon(
            new OfficePoint(0, 30),
            new OfficePoint(40, 0),
            new OfficePoint(80, 30));
        marker.FillColor = OfficeColor.SteelBlue;
        marker.StrokeColor = OfficeColor.Black;
        marker.StrokeWidth = 1.25;

        var drawing = new OfficeDrawing(120, 60)
            .AddShape(background, 0, 0)
            .AddShape(marker, 20, 15);

        var clone = drawing.Clone();
        background.Width = 10;

        Assert.Equal(120, clone.Width);
        Assert.Equal(60, clone.Height);
        Assert.Equal(2, clone.Shapes.Count);
        Assert.Equal(0, clone.Shapes[0].X);
        Assert.Equal(0, clone.Shapes[0].Y);
        Assert.Equal(OfficeShapeKind.Rectangle, clone.Shapes[0].Shape.Kind);
        Assert.Equal(120, clone.Shapes[0].Shape.Width);
        Assert.Equal(20, clone.Shapes[1].X);
        Assert.Equal(15, clone.Shapes[1].Y);
        Assert.Equal(OfficeShapeKind.Polygon, clone.Shapes[1].Shape.Kind);
        Assert.Equal(OfficeColor.SteelBlue, clone.Shapes[1].Shape.FillColor);
        Assert.Equal(OfficeColor.Black, clone.Shapes[1].Shape.StrokeColor);
    }

    [Fact]
    public void OfficeDrawingRejectsShapesOutsideCanvas() {
        var shape = OfficeShape.Rectangle(40, 20);
        var drawing = new OfficeDrawing(60, 30);

        Assert.Throws<ArgumentOutOfRangeException>(() => drawing.AddShape(shape, 25, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => drawing.AddShape(shape, 0, 15));
    }

    [Fact]
    public void OfficeTextMeasurerProvidesDeterministicReusableMetrics() {
        var measurer = OfficeTextMeasurer.Create(OfficeFontInfo.Default);
        var regular = measurer.CreateStyle(new OfficeFontInfo("Calibri", 11));
        var bold = measurer.CreateStyle(new OfficeFontInfo("Calibri", 11, OfficeFontStyle.Bold));
        var mono = measurer.CreateStyle(new OfficeFontInfo("Consolas", 11));

        var regularMetrics = measurer.Measure("OfficeIMO 123", regular);
        var boldMetrics = measurer.Measure("OfficeIMO 123", bold);
        var monoMetrics = measurer.Measure("OfficeIMO 123", mono);

        Assert.True(regularMetrics.WidthPixels > 0);
        Assert.True(regularMetrics.LineHeightPixels > 0);
        Assert.True(boldMetrics.WidthPixels > regularMetrics.WidthPixels);
        Assert.True(monoMetrics.MaximumDigitWidthPixels > regularMetrics.MaximumDigitWidthPixels);
    }

    [Fact]
    public void OfficeTextMeasurerNormalizesFallbackFontInfo() {
        var measurer = OfficeTextMeasurer.Create(new OfficeFontInfo(null, 0));

        Assert.Equal(OfficeFontInfo.Default, measurer.FallbackFontInfo);
        Assert.Equal(OfficeTextMeasurer.DefaultDpi, measurer.DefaultStyle.Dpi);
        Assert.True(measurer.DefaultStyle.MaximumDigitWidthPixels > 0);
    }

    [Theory]
    [InlineData("png", OfficeImageFormat.Png)]
    [InlineData(".png", OfficeImageFormat.Png)]
    [InlineData("photo.JPG", OfficeImageFormat.Jpeg)]
    [InlineData("diagram.svg", OfficeImageFormat.Svg)]
    [InlineData("legacy.emf", OfficeImageFormat.Emf)]
    public void OfficeImageReaderMapsFileNamesAndBareExtensions(string fileName, OfficeImageFormat expected) {
        Assert.Equal(expected, OfficeImageReader.FromExtension(fileName));
    }

    [Fact]
    public void OfficeImageReaderIdentifiesPngWithoutDecodingPixels() {
        var image = OfficeImageReader.Identify(OnePixelPng);

        Assert.Equal(OfficeImageFormat.Png, image.Format);
        Assert.Equal(1, image.Width);
        Assert.Equal(1, image.Height);
        Assert.Equal("image/png", image.MimeType);
    }

    [Fact]
    public void OfficeImageReaderIdentifiesIconDimensionsFromHeader() {
        var icon = new byte[22];
        icon[2] = 0x01;
        icon[4] = 0x01;
        icon[6] = 16;
        icon[7] = 32;

        var image = OfficeImageReader.Identify(icon);

        Assert.Equal(OfficeImageFormat.Icon, image.Format);
        Assert.Equal(16, image.Width);
        Assert.Equal(32, image.Height);
        Assert.Equal("image/x-icon", image.MimeType);
    }

    [Fact]
    public void OfficeImageReaderIdentifiesPcxDimensionsFromHeader() {
        var pcx = new byte[128];
        pcx[0] = 0x0A;
        pcx[1] = 0x05;
        pcx[2] = 0x01;
        pcx[3] = 0x08;
        pcx[8] = 99;
        pcx[10] = 49;
        pcx[12] = 96;
        pcx[14] = 96;

        var image = OfficeImageReader.Identify(pcx);

        Assert.Equal(OfficeImageFormat.Pcx, image.Format);
        Assert.Equal(100, image.Width);
        Assert.Equal(50, image.Height);
        Assert.Equal(96, image.DpiX);
        Assert.Equal(96, image.DpiY);
        Assert.Equal("image/x-pcx", image.MimeType);
    }

    [Fact]
    public void OfficeImageReaderIdentifiesEmfDimensionsFromHeader() {
        var emf = new byte[88];
        WriteInt32LittleEndian(emf, 0, 1);
        WriteInt32LittleEndian(emf, 4, 88);
        WriteInt32LittleEndian(emf, 16, 192);
        WriteInt32LittleEndian(emf, 20, 96);
        WriteInt32LittleEndian(emf, 32, 5080);
        WriteInt32LittleEndian(emf, 36, 2540);
        WriteInt32LittleEndian(emf, 40, 0x464D4520);
        WriteInt32LittleEndian(emf, 72, 1920);
        WriteInt32LittleEndian(emf, 76, 1080);
        WriteInt32LittleEndian(emf, 80, 508);
        WriteInt32LittleEndian(emf, 84, 286);

        var image = OfficeImageReader.Identify(emf);

        Assert.Equal(OfficeImageFormat.Emf, image.Format);
        Assert.Equal(192, image.Width);
        Assert.Equal(96, image.Height);
        Assert.Equal(96, Math.Round(image.DpiX));
        Assert.Equal(96, Math.Round(image.DpiY));
        Assert.Equal("image/x-emf", image.MimeType);
    }

    [Fact]
    public void OfficeImageReaderIdentifiesPlaceableWmfDimensionsFromHeader() {
        var wmf = new byte[22];
        WriteInt32LittleEndian(wmf, 0, unchecked((int)0x9AC6CDD7));
        WriteInt16LittleEndian(wmf, 10, 2880);
        WriteInt16LittleEndian(wmf, 12, 1440);
        WriteUInt16LittleEndian(wmf, 14, 1440);
        WritePlaceableWmfChecksum(wmf);

        var image = OfficeImageReader.Identify(wmf);

        Assert.Equal(OfficeImageFormat.Wmf, image.Format);
        Assert.Equal(192, image.Width);
        Assert.Equal(96, image.Height);
        Assert.Equal("image/x-wmf", image.MimeType);
    }

    [Fact]
    public void OfficeImageReaderRestoresSeekableStreamPosition() {
        var data = new byte[OnePixelPng.Length + 2];
        data[0] = 0xAA;
        data[1] = 0xBB;
        Array.Copy(OnePixelPng, 0, data, 2, OnePixelPng.Length);
        using var stream = new MemoryStream(data);
        stream.Position = 2;

        Assert.True(OfficeImageReader.TryIdentify(stream, null, out var image));

        Assert.Equal(2, stream.Position);
        Assert.Equal(OfficeImageFormat.Png, image.Format);
        Assert.Equal(1, image.Width);
        Assert.Equal(1, image.Height);
    }

    [Fact]
    public void OfficeImageReaderReadsSvgViewBoxDimensions() {
        var svg = System.Text.Encoding.UTF8.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 320 180\"></svg>");

        Assert.True(OfficeImageReader.TryIdentify(svg, "chart.svg", out var image));

        Assert.Equal(OfficeImageFormat.Svg, image.Format);
        Assert.Equal(320, image.Width);
        Assert.Equal(180, image.Height);
        Assert.Equal("image/svg+xml", image.MimeType);
    }

    [Fact]
    public void OfficeImageReaderReadsSvgPhysicalUnitsAsPixels() {
        var svg = System.Text.Encoding.UTF8.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"1in\" height=\"2.54cm\"></svg>");

        Assert.True(OfficeImageReader.TryIdentify(svg, "units.svg", out var image));

        Assert.Equal(OfficeImageFormat.Svg, image.Format);
        Assert.Equal(96, image.Width);
        Assert.Equal(96, image.Height);
    }

    [Fact]
    public void OfficeImageReaderRejectsSvgWithDtdWhenNoExtensionFallbackExists() {
        var svg = System.Text.Encoding.UTF8.GetBytes("<?xml version=\"1.0\"?><!DOCTYPE svg [<!ENTITY xxe SYSTEM \"file:///c:/windows/win.ini\">]><svg xmlns=\"http://www.w3.org/2000/svg\" width=\"1\" height=\"1\">&xxe;</svg>");

        Assert.False(OfficeImageReader.TryIdentify(svg, null, out var image));

        Assert.Equal(OfficeImageFormat.Unknown, image.Format);
    }

    [Fact]
    public void OfficeImageReaderFallsBackToSvgExtensionWhenXmlCannotBeParsed() {
        var svg = System.Text.Encoding.UTF8.GetBytes("<?xml version=\"1.0\"?><!DOCTYPE svg [<!ENTITY xxe SYSTEM \"file:///c:/windows/win.ini\">]><svg xmlns=\"http://www.w3.org/2000/svg\" width=\"1\" height=\"1\">&xxe;</svg>");

        Assert.True(OfficeImageReader.TryIdentify(svg, "unsafe.svg", out var image));

        Assert.Equal(OfficeImageFormat.Svg, image.Format);
        Assert.Equal(0, image.Width);
        Assert.Equal(0, image.Height);
    }

    private static void WriteInt16LittleEndian(byte[] data, int offset, short value) {
        data[offset] = (byte)(value & 0xFF);
        data[offset + 1] = (byte)((value >> 8) & 0xFF);
    }

    private static void WriteUInt16LittleEndian(byte[] data, int offset, ushort value) {
        data[offset] = (byte)(value & 0xFF);
        data[offset + 1] = (byte)((value >> 8) & 0xFF);
    }

    private static void WriteInt32LittleEndian(byte[] data, int offset, int value) {
        data[offset] = (byte)(value & 0xFF);
        data[offset + 1] = (byte)((value >> 8) & 0xFF);
        data[offset + 2] = (byte)((value >> 16) & 0xFF);
        data[offset + 3] = (byte)((value >> 24) & 0xFF);
    }

    private static void WritePlaceableWmfChecksum(byte[] data) {
        ushort checksum = 0;
        for (int offset = 0; offset < 20; offset += 2) {
            checksum ^= (ushort)(data[offset] | (data[offset + 1] << 8));
        }

        WriteUInt16LittleEndian(data, 20, checksum);
    }
}
