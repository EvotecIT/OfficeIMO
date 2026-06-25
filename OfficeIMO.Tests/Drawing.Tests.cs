using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
    public void OfficeImagePlacementFitsImagesIntoTargetRectangles() {
        OfficeImagePlacement stretch = OfficeImagePlacement.Fit(
            sourceWidth: 200D,
            sourceHeight: 100D,
            targetX: 10D,
            targetY: 20D,
            targetWidth: 80D,
            targetHeight: 40D,
            fit: OfficeImageFit.Stretch);
        Assert.Equal((10D, 20D, 80D, 40D), stretch.ToTuple());

        OfficeImagePlacement containedWide = OfficeImagePlacement.Fit(
            sourceWidth: 400D,
            sourceHeight: 100D,
            targetX: 10D,
            targetY: 20D,
            targetWidth: 80D,
            targetHeight: 40D,
            fit: OfficeImageFit.Contain);
        Assert.Equal((10D, 30D, 80D, 20D), containedWide.ToTuple());

        OfficeImagePlacement containedTall = OfficeImagePlacement.Fit(
            sourceWidth: 100D,
            sourceHeight: 400D,
            targetX: 10D,
            targetY: 20D,
            targetWidth: 80D,
            targetHeight: 40D,
            fit: OfficeImageFit.Contain);
        Assert.Equal((45D, 20D, 10D, 40D), containedTall.ToTuple());

        OfficeImagePlacement coveredWide = OfficeImagePlacement.Fit(
            sourceWidth: 400D,
            sourceHeight: 100D,
            targetX: 10D,
            targetY: 20D,
            targetWidth: 80D,
            targetHeight: 40D,
            fit: OfficeImageFit.Cover);
        Assert.Equal((-30D, 20D, 160D, 40D), coveredWide.ToTuple());
    }

    [Fact]
    public void OfficeImagePlacementRejectsInvalidPlacementInputs() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeImagePlacement(0D, 0D, 0D, 10D));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeImagePlacement.Fit(0D, 1D, 0D, 0D, 10D, 10D, OfficeImageFit.Contain));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeImagePlacement.Fit(1D, 1D, 0D, 0D, double.NaN, 10D, OfficeImageFit.Stretch));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeImagePlacement.Fit(1D, 1D, 0D, 0D, 10D, 10D, (OfficeImageFit)99));
    }

    [Fact]
    public void OfficeImagePlacementCalculatesAspectRatioDistortion() {
        double matching = OfficeImagePlacement.GetAspectRatioDistortionRatio(
            sourceWidth: 200D,
            sourceHeight: 100D,
            targetWidth: 80D,
            targetHeight: 40D);
        double distorted = OfficeImagePlacement.GetAspectRatioDistortionRatio(
            sourceWidth: 200D,
            sourceHeight: 100D,
            targetWidth: 80D,
            targetHeight: 80D);

        Assert.Equal(1D, matching);
        Assert.Equal(2D, distorted);
        Assert.False(OfficeImagePlacement.ExceedsAspectRatioDistortion(200D, 100D, 80D, 40D, 1.02D));
        Assert.True(OfficeImagePlacement.ExceedsAspectRatioDistortion(200D, 100D, 80D, 80D, 1.02D));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeImagePlacement.GetAspectRatioDistortionRatio(0D, 100D, 80D, 80D));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeImagePlacement.ExceedsAspectRatioDistortion(200D, 100D, 80D, 80D, double.NaN));
    }

    [Fact]
    public void OfficeImageSourceCropExposesVisibleSourceRatios() {
        var crop = new OfficeImageSourceCrop(0.25D, 0.1D, 0.5D, 0.2D);

        Assert.True(crop.HasCrop);
        Assert.Equal((0.25D, 0.1D, 0.5D, 0.2D), crop.ToTuple());
        Assert.Equal(0.25D, crop.VisibleWidth);
        Assert.Equal(0.7D, crop.VisibleHeight, precision: 10);
    }

    [Fact]
    public void OfficeImageSourceCropClampsCollapsedAuthoredFractions() {
        OfficeImageSourceCrop crop = OfficeImageSourceCrop.FromClampedFractions(
            left: 0.999D,
            top: double.NaN,
            right: double.PositiveInfinity,
            bottom: -1D);

        Assert.True(crop.HasCrop);
        Assert.Equal(0.999D, crop.Left);
        Assert.Equal(0D, crop.Top);
        Assert.Equal(0.999D, crop.Right);
        Assert.Equal(0D, crop.Bottom);
        Assert.Equal(OfficeImageSourceCrop.MinimumVisibleRatio, crop.VisibleWidth);
        Assert.Equal(1D, crop.VisibleHeight);
    }

    [Fact]
    public void OfficeImageSourceCropStrictFractionsRequireVisibleSourceArea() {
        OfficeImageSourceCrop crop = OfficeImageSourceCrop.FromStrictFractions(
            left: 0.25D,
            top: 0.1D,
            right: 0.25D,
            bottom: 0.2D);

        Assert.True(crop.HasVisibleSourceArea);
        Assert.True(OfficeImageSourceCrop.LeavesVisibleSourceArea(0.25D, 0.1D, 0.25D, 0.2D));
        Assert.False(new OfficeImageSourceCrop(0.75D, 0D, 0.25D, 0D).HasVisibleSourceArea);
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeImageSourceCrop.FromStrictFractions(0.75D, 0D, 0.25D, 0D));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeImageSourceCrop.FromStrictFractions(0D, 0.6D, 0D, 0.4D));
    }

    [Fact]
    public void OfficeImageSourceCropRejectsInvalidFractions() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeImageSourceCrop(-0.01D, 0D, 0D, 0D));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeImageSourceCrop(0D, 1D, 0D, 0D));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeImageSourceCrop(0D, 0D, double.NaN, 0D));
    }

    [Fact]
    public void OfficeImageProjectionScalesPlacementCropAndTransform() {
        var projection = new OfficeImageProjection(
            new OfficeImagePlacement(10D, 20D, 80D, 40D),
            new OfficeImageSourceCrop(0.25D, 0.1D, 0.25D, 0.1D),
            rotationDegrees: 30D,
            flipHorizontal: true);

        OfficeImageProjection scaled = projection.Scale(2D);

        Assert.Equal((20D, 40D, 160D, 80D), scaled.Placement.ToTuple());
        Assert.Equal(0.25D, scaled.SourceLeft);
        Assert.Equal(0.5D, scaled.SourceWidth);
        Assert.Equal(30D, scaled.RotationDegrees);
        Assert.Equal(100D, scaled.RotationCenterX);
        Assert.Equal(80D, scaled.RotationCenterY);
        Assert.True(scaled.HasCrop);
        Assert.True(scaled.HasTransform);
        Assert.True(scaled.FlipHorizontal);
    }

    [Fact]
    public void OfficeImageProjectionCreatesUnitSquareTransformForPlacementRotationAndFlips() {
        OfficeTransform normal = new OfficeImageProjection(
            new OfficeImagePlacement(30D, 90D, 60D, 30D))
            .CreateUnitSquareTransform();
        OfficeTransform rotated = new OfficeImageProjection(
            new OfficeImagePlacement(30D, 90D, 60D, 30D),
            rotationDegrees: 90D)
            .CreateUnitSquareTransform();
        OfficeTransform flipped = new OfficeImageProjection(
            new OfficeImagePlacement(10D, 20D, 80D, 40D),
            flipHorizontal: true)
            .CreateUnitSquareTransform();
        OfficeTransform customCenter = new OfficeImageProjection(
            new OfficeImagePlacement(10D, 20D, 20D, 10D),
            rotationDegrees: 90D,
            rotationCenterX: 0D,
            rotationCenterY: 0D)
            .CreateUnitSquareTransform();

        Assert.Equal(new OfficeTransform(60D, 0D, 0D, 30D, 30D, 90D), normal);
        Assert.Equal(new OfficeTransform(0D, 60D, -30D, 0D, 75D, 75D), rotated);
        Assert.Equal(new OfficeTransform(-80D, 0D, 0D, 40D, 90D, 20D), flipped);
        Assert.Equal(new OfficeTransform(0D, 20D, -10D, 0D, -20D, 10D), customCenter);

        Assert.Equal((30D, 90D, 90D, 120D), normal.TransformRectangleBounds(0D, 0D, 1D, 1D));
        Assert.Equal((10D, 20D, 90D, 60D), new OfficeImageProjection(
            new OfficeImagePlacement(10D, 20D, 80D, 40D),
            flipHorizontal: true).GetDestinationBounds());
    }

    [Fact]
    public void OfficeImageProjectionCreatesFrameTransformForDestinationCoordinates() {
        OfficeImageFrameTransform plain = new OfficeImageProjection(
            new OfficeImagePlacement(10D, 20D, 80D, 40D))
            .CreateFrameTransform();
        OfficeImageFrameTransform flipped = new OfficeImageProjection(
            new OfficeImagePlacement(10D, 20D, 80D, 40D),
            flipHorizontal: true)
            .CreateFrameTransform();
        OfficeImageFrameTransform rotated = new OfficeImageProjection(
            new OfficeImagePlacement(10D, 20D, 80D, 40D),
            rotationDegrees: 90D)
            .CreateFrameTransform();

        Assert.False(plain.HasTransform);
        Assert.Equal((0D, 50D, 40D, false, false), plain.ToTuple());
        Assert.Equal((0D, 50D, 40D, true, false), flipped.ToTuple());
        Assert.True(flipped.HasFlip);
        Assert.Equal(new OfficePoint(90D, 20D), flipped.CreateDestinationTransform().TransformPoint(new OfficePoint(10D, 20D)));
        Assert.True(rotated.HasRotation);
        Assert.Equal(new OfficePoint(70D, 40D), rotated.CreateDestinationTransform().TransformPoint(new OfficePoint(50D, 20D)));
    }

    [Fact]
    public void OfficeTransformInvertsAffineMatrix() {
        OfficeTransform transform = OfficeTransform.Translate(10D, 20D)
            .Then(OfficeTransform.Scale(2D, 4D))
            .Then(OfficeTransform.RotateDegrees(90D));

        Assert.True(transform.TryInvert(out OfficeTransform inverse));
        OfficePoint projected = transform.TransformPoint(new OfficePoint(3D, 5D));
        OfficePoint restored = inverse.TransformPoint(projected);

        Assert.Equal(3D, restored.X, precision: 10);
        Assert.Equal(5D, restored.Y, precision: 10);
        Assert.Throws<InvalidOperationException>(() => new OfficeTransform(0D, 0D, 0D, 0D, 0D, 0D).Invert());
    }

    [Fact]
    public void OfficeImageRenderPlan_ResolvesTopLeftAndBottomLeftCropPlacement() {
        var crop = new OfficeImageSourceCrop(0.25D, 0.1D, 0.25D, 0.2D);

        OfficeImageRenderPlan topLeft = OfficeImageRenderPlan.CreateTopLeft(
            sourceWidth: 200D,
            sourceHeight: 100D,
            targetX: 10D,
            targetY: 20D,
            targetWidth: 80D,
            targetHeight: 40D,
            sourceCrop: crop);

        OfficeImageRenderPlan bottomLeft = OfficeImageRenderPlan.CreateBottomLeft(
            sourceWidth: 200D,
            sourceHeight: 100D,
            targetX: 10D,
            targetBottomY: 20D,
            targetWidth: 80D,
            targetHeight: 40D,
            sourceCrop: crop);

        Assert.Equal((10D, 20D, 80D, 40D), topLeft.TargetPlacement.ToTuple());
        Assert.Equal((10D, 20D, 80D, 40D), topLeft.VisiblePlacement.ToTuple());
        Assert.Equal(-30D, topLeft.ImagePlacement.X);
        Assert.Equal(14.285714285714285D, topLeft.ImagePlacement.Y, precision: 10);
        Assert.Equal(160D, topLeft.ImagePlacement.Width);
        Assert.Equal(57.142857142857146D, topLeft.ImagePlacement.Height, precision: 10);
        Assert.Equal(8.571428571428571D, bottomLeft.ImagePlacement.Y, precision: 10);
        Assert.False(topLeft.RequiresTargetClip);
        Assert.False(bottomLeft.RequiresTargetClip);
    }

    [Fact]
    public void OfficeImageRenderPlan_FitsVisibleCropAndReportsCoverClip() {
        var crop = new OfficeImageSourceCrop(0.25D, 0D, 0.25D, 0D);

        OfficeImageRenderPlan contained = OfficeImageRenderPlan.CreateTopLeft(
            sourceWidth: 400D,
            sourceHeight: 200D,
            targetX: 0D,
            targetY: 0D,
            targetWidth: 100D,
            targetHeight: 50D,
            fit: OfficeImageFit.Contain,
            sourceCrop: crop);

        Assert.Equal((25D, 0D, 50D, 50D), contained.VisiblePlacement.ToTuple());
        Assert.Equal((0D, 0D, 100D, 50D), contained.ImagePlacement.ToTuple());
        Assert.False(contained.RequiresTargetClip);

        OfficeImageRenderPlan covered = OfficeImageRenderPlan.CreateTopLeft(
            sourceWidth: 400D,
            sourceHeight: 100D,
            targetX: 10D,
            targetY: 20D,
            targetWidth: 80D,
            targetHeight: 40D,
            fit: OfficeImageFit.Cover);

        Assert.Equal((-30D, 20D, 160D, 40D), covered.VisiblePlacement.ToTuple());
        Assert.Equal(covered.VisiblePlacement.ToTuple(), covered.ImagePlacement.ToTuple());
        Assert.True(covered.RequiresTargetClip);
    }

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
        var font = new OfficeFontInfo("Calibri", 11, OfficeFontStyle.Underline | OfficeFontStyle.Strikethrough);

        Assert.False(font.IsBold);
        Assert.False(font.IsItalic);
        Assert.True(font.IsUnderline);
        Assert.True(font.IsStrikethrough);
        Assert.Equal("Calibri, 11pt, Underline, Strikethrough", font.ToString());
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

    [Theory]
    [InlineData("image/png; charset=binary", OfficeImageFormat.Png)]
    [InlineData("image/jpg", OfficeImageFormat.Jpeg)]
    [InlineData("image/pjpeg", OfficeImageFormat.Jpeg)]
    [InlineData("image/svg+xml; charset=utf-8", OfficeImageFormat.Svg)]
    [InlineData("image/x-emf", OfficeImageFormat.Emf)]
    [InlineData("image/webp", OfficeImageFormat.Webp)]
    [InlineData("application/octet-stream", OfficeImageFormat.Unknown)]
    public void OfficeImageInfoMapsMimeTypesToSharedFormats(string contentType, OfficeImageFormat expected) {
        Assert.Equal(expected, OfficeImageInfo.FromMimeType(contentType));
    }

    [Theory]
    [InlineData(" image/jpg; charset=binary ", true, "image/jpeg")]
    [InlineData("image/svg", true, "image/svg+xml")]
    [InlineData("image/x-custom; version=1", true, "image/x-custom")]
    [InlineData("application/octet-stream", false, "")]
    [InlineData("", false, "")]
    public void OfficeImageInfoNormalizesImageContentTypes(string contentType, bool expectedResult, string expectedContentType) {
        Assert.Equal(expectedResult, OfficeImageInfo.TryNormalizeImageContentType(contentType, out string normalizedContentType));
        Assert.Equal(expectedContentType, normalizedContentType);
    }

    [Theory]
    [InlineData(OfficeImageFormat.Png, ".png")]
    [InlineData(OfficeImageFormat.Jpeg, ".jpeg")]
    [InlineData(OfficeImageFormat.Svg, ".svg")]
    [InlineData(OfficeImageFormat.Emf, ".emf")]
    [InlineData(OfficeImageFormat.Icon, ".ico")]
    [InlineData(OfficeImageFormat.Webp, ".webp")]
    [InlineData(OfficeImageFormat.Unknown, ".bin")]
    public void OfficeImageInfoProvidesCanonicalImageExtensions(OfficeImageFormat format, string expectedExtension) {
        Assert.Equal(expectedExtension, OfficeImageInfo.GetDefaultExtension(format));
    }

    [Fact]
    public void OfficeTextLayoutEngineBuildsStackedTextBlocksFromTextElements() {
        OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutStackedTextBlock(
            "AB",
            10D,
            12D,
            30D,
            1.2D,
            4D,
            (text, size) => string.IsNullOrEmpty(text) ? 0D : text!.Length * size * 0.5D);

        Assert.Equal(2, layout.Lines.Count);
        Assert.Equal("A", layout.Lines[0].Text);
        Assert.Equal("B", layout.Lines[1].Text);
        Assert.Equal(10D, layout.FontSize);
        Assert.Equal(12D, layout.LineHeight);
        Assert.Equal(24D, layout.Height);
        Assert.False(layout.Clipped);
    }

    [Fact]
    public void OfficeTextLayoutEngineShrinksStackedTextBlocksToFitBounds() {
        OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutStackedTextBlock(
            "ABCD",
            12D,
            20D,
            30D,
            1.2D,
            4D,
            (text, size) => string.IsNullOrEmpty(text) ? 0D : size * 0.6D);

        Assert.Equal(4, layout.Lines.Count);
        Assert.True(layout.FontSize < 12D);
        Assert.True(layout.FontSize >= 4D);
        Assert.True(layout.Height <= 30D);
    }

    [Fact]
    public void OfficeTextLayoutEngineBuildsStackedRichTextBlocksFromTextElements() {
        OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutStackedRichTextBlock(
            new[] {
                new OfficeRichTextRun("A", 10D, OfficeColor.Red, bold: true),
                new OfficeRichTextRun("B", 12D, OfficeColor.Blue, italic: true)
            },
            20D,
            40D,
            1.2D,
            (text, size, _) => string.IsNullOrEmpty(text) ? 0D : size * 0.5D,
            shrinkToFit: true,
            minimumFontSize: 4D);

        Assert.Equal(2, layout.Lines.Count);
        Assert.Equal(15D, layout.LineHeight);
        Assert.Equal(30D, layout.Height);
        Assert.False(layout.Clipped);
        Assert.Single(layout.Lines[0].Segments);
        Assert.Single(layout.Lines[1].Segments);
        Assert.Equal("A", layout.Lines[0].Segments[0].Text);
        Assert.Equal("B", layout.Lines[1].Segments[0].Text);
        Assert.Equal(OfficeColor.Red, layout.Lines[0].Segments[0].Color);
        Assert.Equal(OfficeColor.Blue, layout.Lines[1].Segments[0].Color);
        Assert.True(layout.Lines[0].Segments[0].Bold);
        Assert.True(layout.Lines[1].Segments[0].Italic);
    }

    [Fact]
    public void OfficeTextLayoutEngineShrinksStackedRichTextBlocksToFitBounds() {
        OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutStackedRichTextBlock(
            new[] {
                new OfficeRichTextRun("ABCD", 12D, OfficeColor.Purple, bold: true)
            },
            20D,
            30D,
            1.2D,
            (text, size, _) => string.IsNullOrEmpty(text) ? 0D : size * 0.6D,
            shrinkToFit: true,
            minimumFontSize: 4D);

        Assert.Equal(4, layout.Lines.Count);
        Assert.True(layout.Lines[0].Segments[0].FontSize < 12D);
        Assert.True(layout.Lines[0].Segments[0].FontSize >= 4D);
        Assert.True(layout.Height <= 30D);
        Assert.True(layout.Lines.All(line => line.Segments.Count == 1 && line.Segments[0].Bold));
    }

    [Theory]
    [InlineData(".png", true, "image/png")]
    [InlineData("photo.jpeg", true, "image/jpeg")]
    [InlineData("diagram.svg", false, "image/svg+xml")]
    [InlineData("preview.bmp", true, "image/bmp")]
    [InlineData("legacy.emf", false, "image/x-emf")]
    [InlineData("payload.bin", false, "application/octet-stream")]
    public void OfficeImageInfoOwnsSafeBrowserPreviewImageExtensionPolicy(string fileName, bool expectedRenderable, string expectedContentType) {
        Assert.Equal(expectedContentType, OfficeImageInfo.GetMimeTypeFromExtension(fileName));
        Assert.Equal(expectedRenderable, OfficeImageInfo.IsBrowserPreviewSafeExtension(fileName));
    }

    [Theory]
    [InlineData("image/png; charset=binary", true)]
    [InlineData("image/svg+xml", false)]
    [InlineData("image/bmp", true)]
    [InlineData("image/x-emf", false)]
    [InlineData("application/octet-stream", false)]
    public void OfficeImageInfoOwnsSafeBrowserPreviewImageContentTypePolicy(string contentType, bool expectedRenderable) {
        Assert.Equal(expectedRenderable, OfficeImageInfo.IsBrowserPreviewSafeContentType(contentType));
    }

    [Theory]
    [InlineData(OfficeImageFormat.Png, "image/png")]
    [InlineData(OfficeImageFormat.Jpeg, "image/jpeg")]
    [InlineData(OfficeImageFormat.Gif, "image/gif")]
    [InlineData(OfficeImageFormat.Svg, "image/svg+xml")]
    [InlineData(OfficeImageFormat.Webp, "image/webp")]
    public void OfficeSvgImageRendererResolvesEmbeddableContentTypes(OfficeImageFormat format, string expectedContentType) {
        Assert.True(OfficeSvgImageRenderer.TryGetEmbeddableContentType(format, out string contentType));
        Assert.Equal(expectedContentType, contentType);
    }

    [Theory]
    [InlineData("image/png; charset=binary", "image/png")]
    [InlineData("image/jpg", "image/jpeg")]
    [InlineData("image/svg+xml; charset=utf-8", "image/svg+xml")]
    [InlineData("image/webp", "image/webp")]
    public void OfficeSvgImageRendererNormalizesEmbeddableMimeContentTypes(string contentType, string expectedContentType) {
        Assert.True(OfficeSvgImageRenderer.TryGetEmbeddableContentType(contentType, out string normalizedContentType));
        Assert.Equal(expectedContentType, normalizedContentType);
    }

    [Fact]
    public void OfficeSvgImageRendererRejectsUnsupportedEmbeddableContentTypes() {
        Assert.False(OfficeSvgImageRenderer.TryGetEmbeddableContentType(OfficeImageFormat.Emf, out string unsupportedContentType));
        Assert.Equal(string.Empty, unsupportedContentType);
        Assert.False(OfficeSvgImageRenderer.TryGetEmbeddableContentType("application/octet-stream", out string unsupportedMimeContentType));
        Assert.Equal(string.Empty, unsupportedMimeContentType);
    }

    [Theory]
    [InlineData("image/png; charset=binary", null, null, "image/png")]
    [InlineData("image/jpg", null, null, "image/jpeg")]
    [InlineData("application/octet-stream", "png", ".bin", "image/png")]
    [InlineData("application/octet-stream", "jpeg", ".bin", "image/jpeg")]
    [InlineData("binary/octet-stream", "gif", ".bin", "image/gif")]
    [InlineData("application/octet-stream", "svg-preamble", ".bin", "image/svg+xml")]
    [InlineData(null, null, ".svg", "image/svg+xml")]
    [InlineData(null, null, ".webp", "image/webp")]
    public void OfficeSvgImageRendererResolvesEmbeddableContentTypeFromMetadataBytesAndExtension(string? declaredContentType, string? bytesKind, string? fileName, string expectedContentType) {
        byte[]? bytes = bytesKind switch {
            "png" => new byte[] { 0x89, (byte)'P', (byte)'N', (byte)'G', 0x0D, 0x0A, 0x1A, 0x0A },
            "jpeg" => new byte[] { 0xFF, 0xD8, 0xFF },
            "gif" => Encoding.ASCII.GetBytes("GIF89a"),
            "svg-preamble" => Encoding.UTF8.GetBytes(
                "\uFEFF<?xml version=\"1.0\"?>" +
                "<!-- OfficeIMO preview -->" +
                "<!DOCTYPE svg>" +
                "<?officeimo preview?>" +
                "<svg xmlns=\"http://www.w3.org/2000/svg\"/>"),
            _ => null
        };

        Assert.True(OfficeSvgImageRenderer.TryResolveEmbeddableContentType(declaredContentType, bytes, fileName, out string contentType));
        Assert.Equal(expectedContentType, contentType);
    }

    [Fact]
    public void OfficeSvgImageRendererRejectsUnsupportedEmbeddableContentTypeSources() {
        Assert.False(OfficeSvgImageRenderer.TryResolveEmbeddableContentType("image/tiff", null, ".tif", out string unsupportedContentType));
        Assert.Equal(string.Empty, unsupportedContentType);
        Assert.False(OfficeSvgImageRenderer.TryResolveEmbeddableContentType("application/octet-stream", new byte[] { 1, 2, 3 }, ".bin", out string unknownContentType));
        Assert.Equal(string.Empty, unknownContentType);
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
    public void OfficeDrawingSvgExporter_EmitsRootDimensionsInPoints() {
        var drawing = new OfficeDrawing(120, 80);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Contains("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"120pt\" height=\"80pt\" viewBox=\"0 0 120 80\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_AppliesTransformsInShapeLocalCoordinates() {
        var drawing = new OfficeDrawing(120, 120);
        var shape = OfficeShape.Rectangle(10, 20);
        shape.FillColor = OfficeColor.Red;
        shape.Transform = OfficeTransform.RotateDegrees(90, 5, 10);
        drawing.AddShape(shape, 40, 50);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Contains("<rect x=\"0\" y=\"0\" width=\"10\" height=\"20\"", svg, StringComparison.Ordinal);
        Assert.Contains("transform=\"matrix(0 1 -1 0 55 55)\"", svg, StringComparison.Ordinal);
        Assert.DoesNotContain("<rect x=\"40\" y=\"50\" width=\"10\" height=\"20\" transform=", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_EmitsLinearGradientFillDefinitions() {
        var drawing = new OfficeDrawing(120, 60);
        var shape = OfficeShape.Rectangle(100, 40);
        shape.FillGradient = OfficeLinearGradient.Horizontal(
            OfficeColor.FromRgba(0xFF, 0x00, 0x00, 0x00),
            OfficeColor.SteelBlue);
        drawing.AddShape(shape, 10, 10);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Contains("<linearGradient id=\"officeimo-gradient-1\" x1=\"0%\" y1=\"50%\" x2=\"100%\" y2=\"50%\"", svg, StringComparison.Ordinal);
        Assert.Contains("<stop offset=\"0%\" stop-color=\"#FF0000\" stop-opacity=\"0\"", svg, StringComparison.Ordinal);
        Assert.Contains("<stop offset=\"100%\" stop-color=\"#4682B4\"", svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("fill=\"url(#officeimo-gradient-1)\"", svg, StringComparison.Ordinal);
        Assert.DoesNotContain("fill=\"none\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_EmitsClipPathInShapeLocalCoordinates() {
        var drawing = new OfficeDrawing(120, 80);
        var shape = OfficeShape.RoundedRectangle(80, 40, 8);
        shape.FillColor = OfficeColor.SteelBlue;
        shape.ClipPath = OfficeClipPath.Rectangle(40, 20);
        shape.Transform = OfficeTransform.RotateDegrees(15, 40, 20);
        drawing.AddShape(shape, 10, 12);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Contains("<clipPath id=\"officeimo-clip-1\"><rect x=\"0\" y=\"0\" width=\"40\" height=\"20\"/></clipPath>", svg, StringComparison.Ordinal);
        Assert.Contains("<g clip-path=\"url(#officeimo-clip-1)\" transform=\"matrix(", svg, StringComparison.Ordinal);
        Assert.Contains("<rect x=\"0\" y=\"0\" width=\"80\" height=\"40\" rx=\"8\" ry=\"8\"", svg, StringComparison.Ordinal);
        Assert.DoesNotContain("<rect x=\"10\" y=\"12\" width=\"80\" height=\"40\" rx=\"8\" ry=\"8\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_EmitsRoundedClipPathThroughSharedFormatter() {
        var drawing = new OfficeDrawing(120, 80);
        var shape = OfficeShape.Rectangle(80, 40);
        shape.FillColor = OfficeColor.SteelBlue;
        shape.ClipPath = OfficeClipPath.RoundedRectangle(40, 20, 4);
        drawing.AddShape(shape, 10, 12);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Contains("<clipPath id=\"officeimo-clip-1\"><rect x=\"0\" y=\"0\" width=\"40\" height=\"20\" rx=\"4\" ry=\"4\"/></clipPath>", svg, StringComparison.Ordinal);
        Assert.Contains("<g clip-path=\"url(#officeimo-clip-1)\" transform=\"matrix(1 0 0 1 10 12)\">", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_EmitsTextThroughSharedRenderer() {
        var drawing = new OfficeDrawing(120, 80);
        drawing.AddText(
            "A&B\r\nBeta",
            10,
            12,
            80,
            30,
            new OfficeFontInfo("Aptos", 10D, OfficeFontStyle.Bold | OfficeFontStyle.Italic),
            OfficeColor.FromRgba(1, 2, 3, 128),
            OfficeTextAlignment.Center,
            14D);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Contains("<text x=\"50\" y=\"22\" font-family=\"Aptos\" font-size=\"10\" text-anchor=\"middle\" fill=\"#010203\" fill-opacity=\"0.502\" font-weight=\"700\" font-style=\"italic\">A&amp;B<tspan x=\"50\" dy=\"14\">Beta</tspan></text>", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_EmitsRotatedTextThroughSharedRenderer() {
        var drawing = new OfficeDrawing(120, 80);
        drawing.AddText(
            "Tilt",
            30,
            24,
            60,
            20,
            new OfficeFontInfo("Aptos", 10D),
            OfficeColor.Black,
            OfficeTextAlignment.Center,
            rotationDegrees: 30D,
            rotationCenterX: 60D,
            rotationCenterY: 34D);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Contains("transform=\"rotate(30 60 34)\"", svg, StringComparison.Ordinal);
        Assert.Contains(">Tilt</text>", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_EmitsVerticallyAlignedTextThroughSharedRenderer() {
        var drawing = new OfficeDrawing(120, 80);
        drawing.AddText(
            "Bottom",
            10,
            12,
            80,
            40,
            new OfficeFontInfo("Aptos", 10D),
            OfficeColor.Black,
            OfficeTextAlignment.Right,
            verticalAlignment: OfficeTextVerticalAlignment.Bottom);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Contains("text-anchor=\"end\"", svg, StringComparison.Ordinal);
        Assert.Contains(">Bottom</text>", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_EmitsShapeShadowBehindForegroundShape() {
        var drawing = new OfficeDrawing(120, 80);
        var shape = OfficeShape.Rectangle(80, 30);
        shape.FillColor = OfficeColor.SteelBlue;
        shape.Shadow = new OfficeShadow(OfficeColor.Black, 0.25, 3, 4);
        drawing.AddShape(shape, 10, 12);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        int shadowIndex = svg.IndexOf("<rect x=\"13\" y=\"16\" width=\"80\" height=\"30\" fill=\"#000000\" fill-opacity=\"0.25\" stroke=\"none\"/>", StringComparison.Ordinal);
        int foregroundIndex = svg.IndexOf("<rect x=\"10\" y=\"12\" width=\"80\" height=\"30\" fill=\"#4682B4\" stroke=\"none\"/>", StringComparison.Ordinal);
        Assert.True(shadowIndex >= 0, svg);
        Assert.True(foregroundIndex > shadowIndex, svg);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_EmitsPolygonThroughSharedFormatter() {
        var drawing = new OfficeDrawing(120, 80);
        var shape = OfficeShape.Polygon(
            new OfficePoint(0, 20),
            new OfficePoint(30, 0),
            new OfficePoint(60, 20));
        shape.FillColor = OfficeColor.FromRgba(17, 34, 51, 128);
        shape.StrokeColor = OfficeColor.Red;
        shape.StrokeWidth = 2;
        shape.Transform = OfficeTransform.RotateDegrees(45, 30, 10);
        drawing.AddShape(shape, 10, 12);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Contains("<polygon points=\"0,20 30,0 60,20\"", svg, StringComparison.Ordinal);
        Assert.Contains("fill=\"#112233\" fill-opacity=\"0.502\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke=\"#FF0000\" stroke-width=\"2\"", svg, StringComparison.Ordinal);
        Assert.Contains("transform=\"matrix(0.707 0.707 -0.707 0.707", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_EmitsRectangleThroughSharedFormatter() {
        var drawing = new OfficeDrawing(120, 80);
        var shape = OfficeShape.RoundedRectangle(60, 24, 4);
        shape.FillColor = OfficeColor.FromRgba(17, 34, 51, 128);
        shape.StrokeColor = OfficeColor.Red;
        shape.StrokeWidth = 2;
        drawing.AddShape(shape, 10, 12);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Contains("<rect x=\"10\" y=\"12\" width=\"60\" height=\"24\" rx=\"4\" ry=\"4\"", svg, StringComparison.Ordinal);
        Assert.Contains("fill=\"#112233\" fill-opacity=\"0.502\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke=\"#FF0000\" stroke-width=\"2\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_EmitsEllipseThroughSharedFormatter() {
        var drawing = new OfficeDrawing(120, 80);
        var shape = OfficeShape.Ellipse(60, 24);
        shape.FillColor = OfficeColor.FromRgba(17, 34, 51, 128);
        shape.StrokeColor = OfficeColor.Red;
        shape.StrokeWidth = 2;
        drawing.AddShape(shape, 10, 12);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Contains("<ellipse cx=\"40\" cy=\"24\" rx=\"30\" ry=\"12\"", svg, StringComparison.Ordinal);
        Assert.Contains("fill=\"#112233\" fill-opacity=\"0.502\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke=\"#FF0000\" stroke-width=\"2\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_EmitsLineThroughSharedFormatter() {
        var drawing = new OfficeDrawing(120, 80);
        var shape = OfficeShape.Line(0, 0, 60, 20);
        shape.StrokeColor = OfficeColor.FromRgba(17, 34, 51, 128);
        shape.StrokeWidth = 2;
        shape.StrokeDashStyle = OfficeStrokeDashStyle.Dash;
        shape.StrokeLineCap = OfficeStrokeLineCap.Round;
        shape.Transform = OfficeTransform.RotateDegrees(45, 30, 10);
        drawing.AddShape(shape, 10, 12);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Contains("<line x1=\"0\" y1=\"0\" x2=\"60\" y2=\"20\"", svg, StringComparison.Ordinal);
        Assert.Contains("fill=\"none\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke=\"#112233\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke-width=\"2\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke-opacity=\"0.502\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke-dasharray=\"8 4\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke-linecap=\"round\"", svg, StringComparison.Ordinal);
        Assert.Contains("transform=\"matrix(0.707 0.707 -0.707 0.707", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_EmitsPathThroughSharedFormatter() {
        var drawing = new OfficeDrawing(120, 80);
        var shape = OfficeShape.Path(
            OfficePathCommand.MoveTo(0, 0),
            OfficePathCommand.QuadraticBezierTo(30, 0, 60, 20),
            OfficePathCommand.LineTo(0, 20),
            OfficePathCommand.Close());
        shape.FillColor = OfficeColor.FromRgba(17, 34, 51, 128);
        shape.StrokeColor = OfficeColor.Red;
        shape.StrokeWidth = 2;
        shape.Transform = OfficeTransform.RotateDegrees(45, 30, 10);
        drawing.AddShape(shape, 10, 12);

        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

        Assert.Contains("<path d=\"M0 0Q30 0 60 20L0 20Z\"", svg, StringComparison.Ordinal);
        Assert.Contains("fill=\"#112233\" fill-opacity=\"0.502\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke=\"#FF0000\" stroke-width=\"2\"", svg, StringComparison.Ordinal);
        Assert.Contains("transform=\"matrix(0.707 0.707 -0.707 0.707", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeClipPathStoresReusablePathIntent() {
        var clipPath = OfficeClipPath.Path(
            OfficePathCommand.MoveTo(10, 30),
            OfficePathCommand.QuadraticBezierTo(50, 0, 90, 30),
            OfficePathCommand.Close());

        var clone = clipPath.Clone();

        Assert.Equal(OfficeClipPathKind.Path, clone.Kind);
        Assert.Equal(80, clone.Width);
        Assert.Equal(30, clone.Height);
        Assert.Equal(3, clone.Commands.Count);
        Assert.Equal(OfficePathCommand.MoveTo(0, 30), clone.Commands[0]);
        Assert.Equal(OfficePathCommand.QuadraticBezierTo(40, 0, 80, 30), clone.Commands[1]);
        Assert.Equal(OfficePathCommand.Close(), clone.Commands[2]);

        OfficeClipPath scaled = clipPath.Scale(2D, 3D);
        Assert.Equal(160, scaled.Width);
        Assert.Equal(90, scaled.Height);
        Assert.Equal(OfficePathCommand.MoveTo(0, 90), scaled.Commands[0]);
        Assert.Equal(OfficePathCommand.QuadraticBezierTo(80, 0, 160, 90), scaled.Commands[1]);
        Assert.Equal(OfficePathCommand.Close(), scaled.Commands[2]);

        OfficeClipPath rounded = OfficeClipPath.RoundedRectangle(20, 10, 4).Scale(3D, 2D);
        Assert.Equal(60, rounded.Width);
        Assert.Equal(20, rounded.Height);
        Assert.Equal(8, rounded.CornerRadius);
        Assert.Throws<ArgumentException>(() => OfficeClipPath.Path(OfficePathCommand.LineTo(10, 10)));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeClipPath.Rectangle(double.NaN, 10));
        Assert.Throws<ArgumentOutOfRangeException>(() => clipPath.Scale(0D, 1D));
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

    [Theory]
    [InlineData(0D, 0D, 0.5D, 1D, 0.5D)]
    [InlineData(45D, 0D, 0D, 1D, 1D)]
    [InlineData(90D, 0.5D, 0D, 0.5D, 1D)]
    [InlineData(180D, 1D, 0.5D, 0D, 0.5D)]
    [InlineData(450D, 0.5D, 0D, 0.5D, 1D)]
    [InlineData(-90D, 0.5D, 1D, 0.5D, 0D)]
    public void OfficeLinearGradientFromAngleProjectsNormalizedEndpoints(double degrees, double startX, double startY, double endX, double endY) {
        OfficeLinearGradient gradient = OfficeLinearGradient.FromAngle(OfficeColor.Blue, OfficeColor.Green, degrees);

        Assert.Equal(startX, gradient.StartX, precision: 10);
        Assert.Equal(startY, gradient.StartY, precision: 10);
        Assert.Equal(endX, gradient.EndX, precision: 10);
        Assert.Equal(endY, gradient.EndY, precision: 10);
        Assert.Equal(new OfficeGradientStop(0, OfficeColor.Blue), gradient.Stops[0]);
        Assert.Equal(new OfficeGradientStop(1, OfficeColor.Green), gradient.Stops[1]);
    }

    [Fact]
    public void OfficeLinearGradientFromAngleRejectsInvalidAngles() {
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeLinearGradient.FromAngle(OfficeColor.Blue, OfficeColor.Green, double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeLinearGradient.FromAngle(OfficeColor.Blue, OfficeColor.Green, double.PositiveInfinity));
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
    public void OfficeStrokeDashStyleFormatsReusableSvgDashArrays() {
        Assert.Equal("10 5 2.5 5", OfficeStrokeDashStyle.DashDot.GetSvgDashArray(2.5D));
        Assert.Equal("1 2", OfficeStrokeDashStyle.Dot.GetSvgDashArray(1D));
        Assert.Null(OfficeStrokeDashStyle.Solid.GetSvgDashArray(4D));
    }

    [Fact]
    public void OfficeStrokeDashStyleMapperNormalizesOfficeLinePatternSources() {
        Assert.Equal(OfficeStrokeDashStyle.Solid, OfficeStrokeDashStyleMapper.FromVisioLinePattern(0));
        Assert.Equal(OfficeStrokeDashStyle.Solid, OfficeStrokeDashStyleMapper.FromVisioLinePattern(1));
        Assert.Equal(OfficeStrokeDashStyle.Dash, OfficeStrokeDashStyleMapper.FromVisioLinePattern(2));
        Assert.Equal(OfficeStrokeDashStyle.Dot, OfficeStrokeDashStyleMapper.FromVisioLinePattern(3));
        Assert.Equal(OfficeStrokeDashStyle.DashDot, OfficeStrokeDashStyleMapper.FromVisioLinePattern(4));
        Assert.Equal(OfficeStrokeDashStyle.DashDotDot, OfficeStrokeDashStyleMapper.FromVisioLinePattern(5));
        Assert.Equal(OfficeStrokeDashStyle.Dash, OfficeStrokeDashStyleMapper.FromVisioLinePattern(99));

        Assert.True(OfficeStrokeDashStyleMapper.TryMapOfficePresetDash("Dash", out OfficeStrokeDashStyle dash));
        Assert.Equal(OfficeStrokeDashStyle.Dash, dash);
        Assert.True(OfficeStrokeDashStyleMapper.TryMapOfficePresetDash("lgDashDot", out OfficeStrokeDashStyle dashDot));
        Assert.Equal(OfficeStrokeDashStyle.DashDot, dashDot);
        Assert.True(OfficeStrokeDashStyleMapper.TryMapOfficePresetDash("SystemDashDotDot", out OfficeStrokeDashStyle dashDotDot));
        Assert.Equal(OfficeStrokeDashStyle.DashDotDot, dashDotDot);
        Assert.True(OfficeStrokeDashStyleMapper.TryMapOfficePresetDash("sysDot", out OfficeStrokeDashStyle dot));
        Assert.Equal(OfficeStrokeDashStyle.Dot, dot);
        Assert.False(OfficeStrokeDashStyleMapper.TryMapOfficePresetDash("Solid", out OfficeStrokeDashStyle solid));
        Assert.Equal(OfficeStrokeDashStyle.Solid, solid);
        Assert.False(OfficeStrokeDashStyleMapper.TryMapOfficePresetDash("unknown", out OfficeStrokeDashStyle unknown));
        Assert.Equal(OfficeStrokeDashStyle.Solid, unknown);
    }

    [Fact]
    public void OfficeGeometryCalculatesReusableParallelLineOffsets() {
        Assert.True(OfficeGeometry.TryGetParallelLineOffsets(10D, 2D, 10D, 12D, 4D, out double offsetX, out double offsetY));
        Assert.Equal(-2D, offsetX, precision: 6);
        Assert.Equal(0D, offsetY, precision: 6);

        Assert.False(OfficeGeometry.TryGetParallelLineOffsets(10D, 2D, 10D, 2D, 4D, out offsetX, out offsetY));
        Assert.Equal(0D, offsetX);
        Assert.Equal(0D, offsetY);
    }

    [Fact]
    public void OfficeGeometryDetectsReusableSegmentIntersections() {
        Assert.True(OfficeGeometry.SegmentsIntersect((0D, 0D), (4D, 4D), (0D, 4D), (4D, 0D)));
        Assert.True(OfficeGeometry.SegmentsIntersect((0D, 0D), (4D, 0D), (2D, 0D), (5D, 0D)));
        Assert.True(OfficeGeometry.SegmentsIntersect(new OfficePoint(0D, 0D), new OfficePoint(0D, 2D), new OfficePoint(-1D, 1D), new OfficePoint(1D, 1D)));
        Assert.False(OfficeGeometry.SegmentsIntersect((0D, 0D), (1D, 0D), (0D, 1D), (1D, 1D)));
    }

    [Fact]
    public void OfficeGeometryDetectsReusableSegmentRectangleIntersections() {
        Assert.True(OfficeGeometry.SegmentIntersectsRectangle((0D, 0D), (4D, 4D), 1D, 1D, 3D, 3D));
        Assert.True(OfficeGeometry.SegmentIntersectsRectangle((2D, 2D), (2.5D, 2.5D), 1D, 1D, 3D, 3D));
        Assert.True(OfficeGeometry.SegmentIntersectsRectangle((0D, 1D), (1D, 1D), 1D, 1D, 3D, 3D));
        Assert.True(OfficeGeometry.SegmentIntersectsRectangle(new OfficePoint(4D, 2D), new OfficePoint(2D, 2D), 3D, 1D, 1D, 3D));
        Assert.False(OfficeGeometry.SegmentIntersectsRectangle((0D, 0D), (0.5D, 0.5D), 1D, 1D, 3D, 3D));
    }

    [Fact]
    public void OfficeGeometryCalculatesReusableArrowheadGeometry() {
        Assert.True(OfficeGeometry.TryCreateArrowheadPoints(new OfficePoint(10D, 10D), new OfficePoint(0D, 10D), 2D, out OfficePoint[] arrow));
        Assert.Equal(3, arrow.Length);
        Assert.Equal(10D, arrow[0].X);
        Assert.Equal(10D, arrow[0].Y);
        Assert.Equal(arrow[1].X, arrow[2].X, precision: 6);
        Assert.True(arrow[1].Y > arrow[0].Y);
        Assert.True(arrow[2].Y < arrow[0].Y);

        Assert.False(OfficeGeometry.TryCreateArrowheadPoints(new OfficePoint(10D, 10D), new OfficePoint(10D, 10D), 2D, out arrow));
        Assert.Empty(arrow);
    }

    [Fact]
    public void OfficeGeometryFindsReusableArrowheadSegments() {
        var points = new[] {
            (X: 0D, Y: 0D),
            (X: 0D, Y: 0D),
            (X: 3D, Y: 4D)
        };

        Assert.True(OfficeGeometry.TryGetArrowheadSegment(points, fromStart: true, out (double X, double Y) startTip, out (double X, double Y) startFrom));
        Assert.Equal((0D, 0D), startTip);
        Assert.Equal((3D, 4D), startFrom);

        Assert.True(OfficeGeometry.TryGetArrowheadSegment(points, fromStart: false, out (double X, double Y) endTip, out (double X, double Y) endFrom));
        Assert.Equal((3D, 4D), endTip);
        Assert.Equal((0D, 0D), endFrom);
    }

    [Fact]
    public void OfficeGeometryResolvesReusableRectangleBoundaryEndpoints() {
        OfficePoint right = OfficeGeometry.ResolveRectangleBoundaryEndpoint(
            sourceLeft: 0D,
            sourceBottom: 0D,
            sourceRight: 10D,
            sourceTop: 6D,
            targetLeft: 20D,
            targetBottom: 1D,
            targetRight: 30D,
            targetTop: 5D);
        Assert.Equal(10D, right.X);
        Assert.Equal(3D, right.Y);

        OfficePoint top = OfficeGeometry.ResolveRectangleBoundaryEndpoint(
            sourceLeft: 0D,
            sourceBottom: 0D,
            sourceRight: 10D,
            sourceTop: 6D,
            targetLeft: 2D,
            targetBottom: 20D,
            targetRight: 8D,
            targetTop: 30D);
        Assert.Equal(5D, top.X);
        Assert.Equal(6D, top.Y);

        OfficeGeometry.ResolveRectangleBoundaryEndpoint(
            sourceLeft: 10D,
            sourceBottom: 6D,
            sourceRight: 0D,
            sourceTop: 0D,
            targetLeft: -30D,
            targetBottom: 1D,
            targetRight: -20D,
            targetTop: 5D,
            out double leftX,
            out double leftY);
        (double X, double Y) left = (leftX, leftY);
        Assert.Equal((0D, 3D), left);

        OfficePoint aligned = OfficeGeometry.ResolveRectangleBoundaryEndpoint(0D, 0D, 10D, 6D, 0D, 0D, 10D, 6D);
        Assert.Equal(10D, aligned.X);
        Assert.Equal(3D, aligned.Y);
    }

    [Fact]
    public void OfficeSvgFormattingFormatsReusableSvgValues() {
        Assert.Equal("12.346", OfficeSvgFormatting.FormatNumber(12.34567D));
        Assert.Equal("0", OfficeSvgFormatting.FormatNumber(0.00000001D));
        Assert.Equal("&lt;A&amp;B&quot;&gt;", OfficeSvgFormatting.Escape("<A&B\">"));
        Assert.Equal("#112233", OfficeSvgFormatting.ToCssColor(OfficeColor.FromRgba(17, 34, 51, 128)));
        Assert.Equal(0.502D, Math.Round(OfficeSvgFormatting.ToOpacity(OfficeColor.FromRgba(17, 34, 51, 128)), 3));

        var builder = new StringBuilder("<text");
        builder.AppendNumberAttribute("x", 0.00000001D)
            .AppendAttribute("font-family", "A&B")
            .AppendPaintAttribute("fill", OfficeColor.FromRgba(17, 34, 51, 128))
            .AppendStrokeLineCapAttribute(OfficeStrokeLineCap.Round)
            .AppendStrokeLineJoinAttribute(OfficeStrokeLineJoin.Bevel)
            .AppendStrokeDashStyleAttribute(OfficeStrokeDashStyle.DashDot, 2.5D)
            .Append(">");
        Assert.Equal("<text x=\"0\" font-family=\"A&amp;B\" fill=\"#112233\" fill-opacity=\"0.502\" stroke-linecap=\"round\" stroke-linejoin=\"bevel\" stroke-dasharray=\"10 5 2.5 5\">", builder.ToString());

        Assert.Equal("butt", OfficeSvgFormatting.FormatStrokeLineCap(OfficeStrokeLineCap.Butt));
        Assert.Equal("square", OfficeSvgFormatting.FormatStrokeLineCap(OfficeStrokeLineCap.Square));
        Assert.Equal("miter", OfficeSvgFormatting.FormatStrokeLineJoin(OfficeStrokeLineJoin.Miter));
        Assert.Equal("round", OfficeSvgFormatting.FormatStrokeLineJoin(OfficeStrokeLineJoin.Round));

        var solidDashBuilder = new StringBuilder("<line");
        solidDashBuilder.AppendStrokeDashStyleAttribute(OfficeStrokeDashStyle.Solid, 4D).Append("/>");
        Assert.Equal("<line/>", solidDashBuilder.ToString());

        var lineBuilder = new StringBuilder();
        lineBuilder.AppendLineElement(1.25D, 2.5D, 30.125D, 40.75D, OfficeColor.FromRgba(17, 34, 51, 128), 1.5D, OfficeStrokeDashStyle.Dot, OfficeStrokeLineCap.Round);
        Assert.Equal("<line x1=\"1.25\" y1=\"2.5\" x2=\"30.125\" y2=\"40.75\" stroke=\"#112233\" stroke-opacity=\"0.502\" stroke-width=\"1.5\" stroke-dasharray=\"1.5 3\" stroke-linecap=\"round\"/>", lineBuilder.ToString());

        var parallelLineBuilder = new StringBuilder();
        parallelLineBuilder.AppendParallelLineElements(10D, 2D, 10D, 12D, OfficeColor.Black, 1D, 4D, OfficeStrokeDashStyle.Dash);
        Assert.Equal("<line x1=\"12\" y1=\"2\" x2=\"12\" y2=\"12\" stroke=\"#000000\" stroke-width=\"1\" stroke-dasharray=\"4 2\"/><line x1=\"8\" y1=\"2\" x2=\"8\" y2=\"12\" stroke=\"#000000\" stroke-width=\"1\" stroke-dasharray=\"4 2\"/>", parallelLineBuilder.ToString());

        var rawLineBuilder = new StringBuilder();
        rawLineBuilder.AppendLineElement(0D, 1D, 2D, 3D, " stroke=\"none\" transform=\"rotate(45)\"");
        Assert.Equal("<line x1=\"0\" y1=\"1\" x2=\"2\" y2=\"3\" stroke=\"none\" transform=\"rotate(45)\"/>", rawLineBuilder.ToString());

        var rectBuilder = new StringBuilder();
        rectBuilder.AppendRectElement(1.25D, 2.5D, 30.125D, 40.75D, " fill=\"none\" transform=\"rotate(45)\"");
        Assert.Equal("<rect x=\"1.25\" y=\"2.5\" width=\"30.125\" height=\"40.75\" fill=\"none\" transform=\"rotate(45)\"/>", rectBuilder.ToString());

        var roundedRectBuilder = new StringBuilder();
        roundedRectBuilder.AppendRectElement(0D, 1D, 2D, 3D, 4.5D, 6.75D, " fill=\"#FFFFFF\"");
        Assert.Equal("<rect x=\"0\" y=\"1\" width=\"2\" height=\"3\" rx=\"4.5\" ry=\"6.75\" fill=\"#FFFFFF\"/>", roundedRectBuilder.ToString());

        var clipBuilder = new StringBuilder();
        clipBuilder.AppendRectClipPathDefinition("clip&1", 1.25D, 2D, 3D, 4D, wrapInDefs: true)
            .Append("<g")
            .AppendClipPathReference("clip&1")
            .Append(">");
        Assert.Equal("<defs><clipPath id=\"clip&amp;1\"><rect x=\"1.25\" y=\"2\" width=\"3\" height=\"4\"/></clipPath></defs><g clip-path=\"url(#clip&amp;1)\">", clipBuilder.ToString());

        string inner = OfficeSvgFormatting.ExtractSvgInner("<svg xmlns=\"http://www.w3.org/2000/svg\"><rect width=\"10\"/></svg>");
        var nestedBuilder = new StringBuilder();
        nestedBuilder.AppendNestedSvg(1.25D, 2.5D, 30.125D, 40.75D, inner);
        Assert.Equal("<rect width=\"10\"/>", inner);
        Assert.Equal("<svg x=\"1.25\" y=\"2.5\" width=\"30.125\" height=\"40.75\" viewBox=\"0 0 30.125 40.75\"><rect width=\"10\"/></svg>", nestedBuilder.ToString());
        Assert.Equal("<g/>", OfficeSvgFormatting.ExtractSvgInner("<g/>"));

        Assert.Equal("rotate(12.346)", OfficeSvgFormatting.FormatRotateTransform(12.34567D));
        Assert.Equal("rotate(12.346 5 6.789)", OfficeSvgFormatting.FormatRotateTransform(12.34567D, 5D, 6.789D));

        var rotateBuilder = new StringBuilder("<text");
        rotateBuilder.AppendRotateTransformAttribute(12.34567D, 5D, 6.789D).Append(">");
        Assert.Equal("<text transform=\"rotate(12.346 5 6.789)\">", rotateBuilder.ToString());

        OfficeTransform matrixTransform = OfficeTransform.RotateDegrees(90, 5D, 10D);
        Assert.Equal("matrix(0 1 -1 0 55 55)", OfficeSvgFormatting.FormatMatrixTransform(matrixTransform, 40D, 50D));

        var matrixBuilder = new StringBuilder("<g");
        matrixBuilder.AppendMatrixTransformAttribute(matrixTransform, 40D, 50D).Append(">");
        Assert.Equal("<g transform=\"matrix(0 1 -1 0 55 55)\">", matrixBuilder.ToString());
        Assert.Equal("rotate(45 50 40)", OfficeSvgFormatting.FormatImageFrameTransform(new OfficeImageFrameTransform(45D, 50D, 40D)));
        Assert.Equal("translate(50 40) rotate(45) scale(-1 1) translate(-50 -40)", OfficeSvgFormatting.FormatImageFrameTransform(new OfficeImageFrameTransform(45D, 50D, 40D, flipHorizontal: true)));
        Assert.Null(OfficeSvgFormatting.FormatImageFrameTransform(new OfficeImageFrameTransform(0D, 50D, 40D)));

        var writerBuilder = new StringBuilder();
        using (var writer = System.Xml.XmlWriter.Create(
            new System.IO.StringWriter(writerBuilder, System.Globalization.CultureInfo.InvariantCulture),
            new System.Xml.XmlWriterSettings { OmitXmlDeclaration = true, ConformanceLevel = System.Xml.ConformanceLevel.Fragment })) {
            writer.WriteStartElement("g");
            writer.WriteNumberAttribute("x", 0.00000001D);
            writer.WriteViewBoxAttribute(0D, 0D, 12.34567D, 6.789D);
            OfficeSvgFormatting.WriteRotateTransformAttribute(writer, 12.34567D, 5D, 6.789D);
            writer.WriteStrokeLineCapAttribute(OfficeStrokeLineCap.Round);
            writer.WriteStrokeLineJoinAttribute(OfficeStrokeLineJoin.Round);
            writer.WriteStrokeDashStyleAttribute(OfficeStrokeDashStyle.Dot, 1D);
            writer.WriteStrokeDashArrayAttribute(null);
            writer.WriteEndElement();
        }

        Assert.Contains("x=\"0\"", writerBuilder.ToString(), StringComparison.Ordinal);
        Assert.Contains("viewBox=\"0 0 12.346 6.789\"", writerBuilder.ToString(), StringComparison.Ordinal);
        Assert.Contains("transform=\"rotate(12.346 5 6.789)\"", writerBuilder.ToString(), StringComparison.Ordinal);
        Assert.Contains("stroke-linecap=\"round\"", writerBuilder.ToString(), StringComparison.Ordinal);
        Assert.Contains("stroke-linejoin=\"round\"", writerBuilder.ToString(), StringComparison.Ordinal);
        Assert.Contains("stroke-dasharray=\"1 2\"", writerBuilder.ToString(), StringComparison.Ordinal);

        var pointsBuilder = new StringBuilder("<polygon");
        pointsBuilder.AppendPointsAttribute(new[] { new OfficePoint(1.2D, 3.4D), new OfficePoint(5D, 0.00000001D) }).Append("/>");
        Assert.Equal("<polygon points=\"1.2,3.4 5,0\"/>", pointsBuilder.ToString());

        var polygonBuilder = new StringBuilder();
        polygonBuilder.AppendPolygonElement(
            new[] { new OfficePoint(1.2D, 3.4D), new OfficePoint(5D, 0.00000001D) },
            OfficeColor.FromRgba(17, 34, 51, 128),
            OfficeColor.Red,
            1.5D);
        Assert.Equal("<polygon points=\"1.2,3.4 5,0\" fill=\"#112233\" fill-opacity=\"0.502\" stroke=\"#FF0000\" stroke-width=\"1.5\"/>", polygonBuilder.ToString());

        var rawPolygonBuilder = new StringBuilder();
        rawPolygonBuilder.AppendPolygonElement(new[] { new OfficePoint(0D, 0D), new OfficePoint(2D, 3D), new OfficePoint(4D, 0D) }, " fill=\"none\" transform=\"rotate(45)\"");
        Assert.Equal("<polygon points=\"0,0 2,3 4,0\" fill=\"none\" transform=\"rotate(45)\"/>", rawPolygonBuilder.ToString());

        var polylineBuilder = new StringBuilder();
        polylineBuilder.AppendPolylineElement(new[] { new OfficePoint(0D, 0D), new OfficePoint(2.5D, 3D), new OfficePoint(4D, 0D) }, " fill=\"none\" stroke=\"#112233\"");
        Assert.Equal("<polyline fill=\"none\" stroke=\"#112233\" points=\"0,0 2.5,3 4,0\"/>", polylineBuilder.ToString());

        var circleBuilder = new StringBuilder();
        circleBuilder.AppendCircleElement(1.25D, 2.5D, 3.75D, OfficeColor.FromRgba(17, 34, 51, 128));
        Assert.Equal("<circle cx=\"1.25\" cy=\"2.5\" r=\"3.75\" fill=\"#112233\" fill-opacity=\"0.502\"/>", circleBuilder.ToString());

        var rawCircleBuilder = new StringBuilder();
        rawCircleBuilder.AppendCircleElement(0D, 1D, 2D, " fill=\"none\" stroke=\"#FF0000\"");
        Assert.Equal("<circle cx=\"0\" cy=\"1\" r=\"2\" fill=\"none\" stroke=\"#FF0000\"/>", rawCircleBuilder.ToString());

        var ellipseBuilder = new StringBuilder();
        ellipseBuilder.AppendEllipseElement(1.25D, 2.5D, 3.75D, 4.5D, OfficeColor.FromRgba(17, 34, 51, 128));
        Assert.Equal("<ellipse cx=\"1.25\" cy=\"2.5\" rx=\"3.75\" ry=\"4.5\" fill=\"#112233\" fill-opacity=\"0.502\"/>", ellipseBuilder.ToString());

        var rawEllipseBuilder = new StringBuilder();
        rawEllipseBuilder.AppendEllipseElement(1D, 2D, 3D, 4D, " fill=\"none\" stroke=\"#112233\"");
        Assert.Equal("<ellipse cx=\"1\" cy=\"2\" rx=\"3\" ry=\"4\" fill=\"none\" stroke=\"#112233\"/>", rawEllipseBuilder.ToString());

        OfficePoint[] pathPoints = {
            new OfficePoint(0.00000001D, 2D),
            new OfficePoint(3.4567D, 4.5D),
            new OfficePoint(6D, 0D)
        };
        Assert.Equal("M 0 2 L 3.457 4.5 L 6 0 Z", OfficeSvgFormatting.FormatMoveLinePathData(pathPoints, closePath: true));
        Assert.Equal(string.Empty, OfficeSvgFormatting.FormatMoveLinePathData(Array.Empty<OfficePoint>(), closePath: true));

        var pathBuilder = new StringBuilder();
        pathBuilder.AppendPathElement(OfficeSvgFormatting.FormatMoveLinePathData(pathPoints), " fill=\"none\" stroke=\"#112233\"");
        Assert.Equal("<path d=\"M 0 2 L 3.457 4.5 L 6 0\" fill=\"none\" stroke=\"#112233\"/>", pathBuilder.ToString());

        OfficePathCommand[] commands = {
            OfficePathCommand.MoveTo(0.00000001D, 2D),
            OfficePathCommand.QuadraticBezierTo(2.2222D, 3.3333D, 4.4444D, 5.5555D),
            OfficePathCommand.CubicBezierTo(1.2345D, 2.3456D, 3.4567D, 4.5678D, 5.6789D, 6.789D),
            OfficePathCommand.LineTo(7D, 8D),
            OfficePathCommand.Close()
        };
        Assert.Equal("M10 22Q12.222 23.333 14.444 25.556C11.235 22.346 13.457 24.568 15.679 26.789L17 28Z", OfficeSvgFormatting.FormatPathData(commands, 10D, 20D));

        var commandBuilder = new StringBuilder();
        commandBuilder.AppendPathData(commands);
        Assert.Equal("M0 2Q2.222 3.333 4.444 5.556C1.235 2.346 3.457 4.568 5.679 6.789L7 8Z", commandBuilder.ToString());

        var commandPathBuilder = new StringBuilder();
        commandPathBuilder.AppendPathElement(commands, 10D, 20D, " fill=\"#FFFFFF\"");
        Assert.Equal("<path d=\"M10 22Q12.222 23.333 14.444 25.556C11.235 22.346 13.457 24.568 15.679 26.789L17 28Z\" fill=\"#FFFFFF\"/>", commandPathBuilder.ToString());
    }

    [Fact]
    public void OfficeSvgFormattingAppendsSharedHatchPatternRectangle() {
        var builder = new StringBuilder();

        builder.AppendHatchPatternRectangle(1, 2, 16, 12, OfficeColor.FromRgba(10, 20, 30, 128), 4, 1.5D, OfficeHatchPatternKind.Trellis);

        string svg = builder.ToString();
        Assert.Contains("<line", svg, StringComparison.Ordinal);
        Assert.Contains("stroke=\"#0A141E\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke-opacity=\"0.502\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke-width=\"1.5\"", svg, StringComparison.Ordinal);
        Assert.Contains("x1=\"-11\"", svg, StringComparison.Ordinal);
        Assert.Contains("x2=\"1\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeSvgFormattingAppendsSharedPercentStipplePatternRectangle() {
        var builder = new StringBuilder();

        builder.AppendHatchPatternRectangle(0, 0, 8, 8, OfficeColor.FromRgb(10, 160, 30), 4, 1, OfficeHatchPatternKind.Percent12_5);

        string svg = builder.ToString();
        Assert.DoesNotContain("<line", svg, StringComparison.Ordinal);
        Assert.Equal(8, CountOccurrences(svg, "<rect"));
        Assert.Contains("fill=\"#0AA01E\"", svg, StringComparison.Ordinal);
        Assert.Contains("x=\"2\" y=\"2\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeSparklineRendererAppendsReusableSvgSparklines() {
        var builder = new StringBuilder();

        OfficeSparklineRenderer.AppendSvg(
            builder,
            0,
            0,
            80,
            28,
            new[] { 4D, -2D, 8D },
            OfficeSparklineKind.WinLoss,
            new OfficeSparklineStyle {
                DisplayAxis = true,
                AxisColor = OfficeColor.Gray,
                PointStyles = new[] {
                    new OfficeSparklinePointStyle(OfficeColor.Blue),
                    new OfficeSparklinePointStyle(OfficeColor.Red),
                    new OfficeSparklinePointStyle(OfficeColor.Lime)
                }
            });

        string svg = builder.ToString();
        Assert.Contains("<line", svg, StringComparison.Ordinal);
        Assert.Contains("<rect", svg, StringComparison.Ordinal);
        Assert.Contains("stroke=\"#808080\"", svg, StringComparison.Ordinal);
        Assert.Contains("fill=\"#0000FF\"", svg, StringComparison.Ordinal);
        Assert.Contains("fill=\"#FF0000\"", svg, StringComparison.Ordinal);
        Assert.Contains("fill=\"#00FF00\"", svg, StringComparison.Ordinal);

        var lineBuilder = new StringBuilder();
        OfficeSparklineRenderer.AppendSvg(
            lineBuilder,
            0,
            0,
            80,
            28,
            new[] { 4D, -2D, 8D },
            OfficeSparklineKind.Line,
            new OfficeSparklineStyle {
                DisplayAxis = true,
                SeriesColor = OfficeColor.FromRgb(37, 99, 235),
                AxisColor = OfficeColor.Gray,
                PointStyles = new[] {
                    new OfficeSparklinePointStyle(OfficeColor.Blue, showMarker: true),
                    new OfficeSparklinePointStyle(OfficeColor.Red, showMarker: true),
                    new OfficeSparklinePointStyle(OfficeColor.Lime, showMarker: true)
                }
            });

        string lineSvg = lineBuilder.ToString();
        Assert.Contains("<polyline", lineSvg, StringComparison.Ordinal);
        Assert.Contains("<circle", lineSvg, StringComparison.Ordinal);
        Assert.Contains("stroke=\"#2563EB\"", lineSvg, StringComparison.Ordinal);
        Assert.Contains("fill=\"#0000FF\"", lineSvg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDataBarRendererAppendsReusableSvgDataBar() {
        var builder = new StringBuilder();

        OfficeDataBarRenderer.AppendSvg(builder, 2, 3, 30, 10, 0.25D, 0.5D, OfficeColor.FromRgba(10, 20, 30, 128), verticalInset: 2D);

        Assert.Equal("<rect x=\"9.5\" y=\"5\" width=\"15\" height=\"6\" fill=\"#0A141E\" fill-opacity=\"0.502\"/>", builder.ToString());
    }

    [Fact]
    public void OfficeConditionalIconRendererAppendsReusableSvgIcons() {
        var circleBuilder = new StringBuilder();
        var arrowBuilder = new StringBuilder();

        OfficeConditionalIconRenderer.AppendSvg(circleBuilder, 1, 2, 16, OfficeConditionalIconKind.RedCircle, scale: 1D);
        OfficeConditionalIconRenderer.AppendSvg(arrowBuilder, 1, 2, 16, OfficeConditionalIconKind.GreenUpArrow, scale: 1D);

        string circleSvg = circleBuilder.ToString();
        string arrowSvg = arrowBuilder.ToString();
        Assert.Contains("<circle", circleSvg, StringComparison.Ordinal);
        Assert.Contains("fill=\"#DC2626\"", circleSvg, StringComparison.Ordinal);
        Assert.Contains("stroke=\"#B91C1C\"", circleSvg, StringComparison.Ordinal);
        Assert.Contains("fill-opacity=", circleSvg, StringComparison.Ordinal);
        Assert.Contains("<path", arrowSvg, StringComparison.Ordinal);
        Assert.Contains("fill=\"#16A34A\"", arrowSvg, StringComparison.Ordinal);
        Assert.Contains("stroke=\"#15803D\"", arrowSvg, StringComparison.Ordinal);
        Assert.Contains("fill-opacity=", arrowSvg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDataBarRendererResolvesReusableGeometryForNativeEmitters() {
        OfficeDataBarGeometry bar = OfficeDataBarRenderer.Resolve(
            10D,
            20D,
            80D,
            30D,
            startRatio: 0.25D,
            ratio: 0.5D,
            verticalInset: 3D,
            minimumHeight: 0D);

        Assert.Equal(30D, bar.X);
        Assert.Equal(23D, bar.Y);
        Assert.Equal(40D, bar.Width);
        Assert.Equal(24D, bar.Height);
        Assert.True(bar.IsVisible);
    }

    [Fact]
    public void OfficeSvgImageRendererAppendsCroppedImageProjection() {
        var builder = new StringBuilder();

        OfficeSvgImageRenderer.AppendImage(
            builder,
            "data:image/png;base64,AA==",
            new OfficeImageProjection(
                new OfficeImagePlacement(10, 20, 80, 40),
                new OfficeImageSourceCrop(0.25D, 0.1D, 0.25D, 0.1D)),
            "imgClip",
            new OfficeImagePlacement(10, 20, 80, 40));

        string svg = builder.ToString();
        Assert.Contains("<clipPath id=\"imgClip\"><rect x=\"10\" y=\"20\" width=\"80\" height=\"40\"/></clipPath>", svg, StringComparison.Ordinal);
        Assert.Contains("<image x=\"-30\" y=\"15\" width=\"160\" height=\"50\" clip-path=\"url(#imgClip)\" href=\"data:image/png;base64,AA==\"/>", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeSvgImageRendererAppendsImageInsideViewport() {
        var uncropped = new StringBuilder();

        OfficeSvgImageRenderer.AppendImageInViewport(
            uncropped,
            "data:image/png;base64,AA==",
            new OfficeImageProjection(new OfficeImagePlacement(10, 20, 80, 40)),
            "viewportClip",
            new OfficeImagePlacement(0, 0, 120, 80));

        string uncroppedSvg = uncropped.ToString();
        Assert.Contains("<clipPath id=\"viewportClip\"><rect x=\"0\" y=\"0\" width=\"120\" height=\"80\"/></clipPath>", uncroppedSvg, StringComparison.Ordinal);
        Assert.Contains("<image x=\"10\" y=\"20\" width=\"80\" height=\"40\" clip-path=\"url(#viewportClip)\" href=\"data:image/png;base64,AA==\"/>", uncroppedSvg, StringComparison.Ordinal);

        var cropped = new StringBuilder();

        OfficeSvgImageRenderer.AppendImageInViewport(
            cropped,
            "data:image/png;base64,AA==",
            new OfficeImageProjection(
                new OfficeImagePlacement(10, 20, 80, 40),
                new OfficeImageSourceCrop(0.25D, 0.1D, 0.25D, 0.1D)),
            "cropClip",
            new OfficeImagePlacement(0, 0, 120, 80));

        string croppedSvg = cropped.ToString();
        Assert.Contains("<clipPath id=\"cropClip\"><rect x=\"10\" y=\"20\" width=\"80\" height=\"40\"/></clipPath>", croppedSvg, StringComparison.Ordinal);
        Assert.Contains("<image x=\"-30\" y=\"15\" width=\"160\" height=\"50\" clip-path=\"url(#cropClip)\" href=\"data:image/png;base64,AA==\"/>", croppedSvg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeSvgImageRendererAppendsFlipAndRotationTransform() {
        var builder = new StringBuilder();

        OfficeSvgImageRenderer.AppendImage(
            builder,
            "data:image/png;base64,AA==",
            10,
            20,
            80,
            40,
            "imgClip",
            0,
            0,
            120,
            80,
            rotationDegrees: 45D,
            flipHorizontal: true);

        string svg = builder.ToString();
        Assert.Contains("clip-path=\"url(#imgClip)\"", svg, StringComparison.Ordinal);
        Assert.Contains("transform=\"translate(50 40) rotate(45) scale(-1 1) translate(-50 -40)\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeSvgImageRendererAppendsPreserveAspectRatio() {
        var builder = new StringBuilder();

        OfficeSvgImageRenderer.AppendImage(
            builder,
            "data:image/png;base64,AA==",
            10,
            20,
            80,
            40,
            preserveAspectRatio: "xMidYMid meet");

        Assert.Equal("<image x=\"10\" y=\"20\" width=\"80\" height=\"40\" preserveAspectRatio=\"xMidYMid meet\" href=\"data:image/png;base64,AA==\"/>", builder.ToString());
    }

    [Fact]
    public void OfficeSvgImageRendererCreatesDataUri() {
        string href = OfficeSvgImageRenderer.CreateDataUri("image/png", new byte[] { 0, 1, 2 });

        Assert.Equal("data:image/png;base64,AAEC", href);
    }

    [Fact]
    public void OfficeSvgImageRendererWritesXmlImageElement() {
        var builder = new StringBuilder();
        using (var writer = System.Xml.XmlWriter.Create(
            new System.IO.StringWriter(builder, System.Globalization.CultureInfo.InvariantCulture),
            new System.Xml.XmlWriterSettings { OmitXmlDeclaration = true, ConformanceLevel = System.Xml.ConformanceLevel.Fragment })) {
            OfficeSvgImageRenderer.WriteImage(
                writer,
                "http://www.w3.org/2000/svg",
                "data:image/png;base64,AA==",
                10,
                20,
                80,
                40,
                rotationDegrees: 45D,
                rotationCenterX: 50D,
                rotationCenterY: 40D,
                preserveAspectRatio: "xMidYMid meet",
                writeAdditionalAttributes: static imageWriter => imageWriter.WriteAttributeString("data-test-image", "true"));
        }

        string svg = builder.ToString();
        Assert.Contains("<image", svg, StringComparison.Ordinal);
        Assert.Contains("data-test-image=\"true\"", svg, StringComparison.Ordinal);
        Assert.Contains("x=\"10\"", svg, StringComparison.Ordinal);
        Assert.Contains("preserveAspectRatio=\"xMidYMid meet\"", svg, StringComparison.Ordinal);
        Assert.Contains("transform=\"rotate(45 50 40)\"", svg, StringComparison.Ordinal);
        Assert.Contains("href=\"data:image/png;base64,AA==\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeSvgImageRendererWritesXmlImageElementFromProjection() {
        var builder = new StringBuilder();
        using (var writer = System.Xml.XmlWriter.Create(
            new System.IO.StringWriter(builder, System.Globalization.CultureInfo.InvariantCulture),
            new System.Xml.XmlWriterSettings { OmitXmlDeclaration = true, ConformanceLevel = System.Xml.ConformanceLevel.Fragment })) {
            OfficeSvgImageRenderer.WriteImage(
                writer,
                "http://www.w3.org/2000/svg",
                "data:image/png;base64,AA==",
                new OfficeImageProjection(
                    new OfficeImagePlacement(10, 20, 80, 40),
                    rotationDegrees: 45D,
                    rotationCenterX: 50D,
                    rotationCenterY: 40D,
                    flipHorizontal: true),
                preserveAspectRatio: "xMidYMid meet",
                writeAdditionalAttributes: static imageWriter => imageWriter.WriteAttributeString("data-test-image", "true"));
        }

        string svg = builder.ToString();
        Assert.Contains("<image", svg, StringComparison.Ordinal);
        Assert.Contains("data-test-image=\"true\"", svg, StringComparison.Ordinal);
        Assert.Contains("x=\"10\"", svg, StringComparison.Ordinal);
        Assert.Contains("preserveAspectRatio=\"xMidYMid meet\"", svg, StringComparison.Ordinal);
        Assert.Contains("transform=\"translate(50 40) rotate(45) scale(-1 1) translate(-50 -40)\"", svg, StringComparison.Ordinal);
        Assert.Contains("href=\"data:image/png;base64,AA==\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeSvgImageRendererWritesXmlCroppedImageProjection() {
        var builder = new StringBuilder();
        using (var writer = System.Xml.XmlWriter.Create(
            new System.IO.StringWriter(builder, System.Globalization.CultureInfo.InvariantCulture),
            new System.Xml.XmlWriterSettings { OmitXmlDeclaration = true, ConformanceLevel = System.Xml.ConformanceLevel.Fragment })) {
            OfficeSvgImageRenderer.WriteImage(
                writer,
                "http://www.w3.org/2000/svg",
                "data:image/png;base64,AA==",
                new OfficeImageProjection(
                    new OfficeImagePlacement(10, 20, 80, 40),
                    new OfficeImageSourceCrop(0.25D, 0.1D, 0.25D, 0.1D)),
                clipPathId: "xmlClip",
                clipRectangle: new OfficeImagePlacement(10, 20, 80, 40));
        }

        string svg = builder.ToString();
        Assert.Contains("<clipPath", svg, StringComparison.Ordinal);
        Assert.Contains("id=\"xmlClip\"", svg, StringComparison.Ordinal);
        Assert.Contains("<rect x=\"10\" y=\"20\" width=\"80\" height=\"40\"", svg, StringComparison.Ordinal);
        Assert.Contains("<image x=\"-30\" y=\"15\" width=\"160\" height=\"50\" clip-path=\"url(#xmlClip)\" href=\"data:image/png;base64,AA==\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeSvgImageRendererWritesXmlImageInsideViewport() {
        var builder = new StringBuilder();
        using (var writer = System.Xml.XmlWriter.Create(
            new System.IO.StringWriter(builder, System.Globalization.CultureInfo.InvariantCulture),
            new System.Xml.XmlWriterSettings { OmitXmlDeclaration = true, ConformanceLevel = System.Xml.ConformanceLevel.Fragment })) {
            OfficeSvgImageRenderer.WriteImageInViewport(
                writer,
                "http://www.w3.org/2000/svg",
                "data:image/png;base64,AA==",
                new OfficeImageProjection(new OfficeImagePlacement(10, 20, 80, 40)),
                "xmlViewportClip",
                new OfficeImagePlacement(0, 0, 120, 80),
                preserveAspectRatio: "xMidYMid meet");
        }

        string svg = builder.ToString();
        Assert.Contains("<clipPath", svg, StringComparison.Ordinal);
        Assert.Contains("id=\"xmlViewportClip\"", svg, StringComparison.Ordinal);
        Assert.Contains("<rect x=\"0\" y=\"0\" width=\"120\" height=\"80\"", svg, StringComparison.Ordinal);
        Assert.Contains("<image x=\"10\" y=\"20\" width=\"80\" height=\"40\" clip-path=\"url(#xmlViewportClip)\" preserveAspectRatio=\"xMidYMid meet\" href=\"data:image/png;base64,AA==\"", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeSvgPrimitiveWriterWritesSharedVectorPrimitives() {
        var builder = new StringBuilder();
        using (var writer = System.Xml.XmlWriter.Create(
            new System.IO.StringWriter(builder, System.Globalization.CultureInfo.InvariantCulture),
            new System.Xml.XmlWriterSettings { OmitXmlDeclaration = true, ConformanceLevel = System.Xml.ConformanceLevel.Fragment })) {
            OfficeSvgPrimitiveWriter.WriteCircle(writer, "http://www.w3.org/2000/svg", 10, 20, 5, OfficeColor.SteelBlue, fill: false, strokeWidth: 2);
            OfficeSvgPrimitiveWriter.WriteRectangle(writer, "http://www.w3.org/2000/svg", 30, 40, 50, 60, OfficeColor.Red, fill: true, strokeWidth: 0, cornerRadius: 4);
            OfficeSvgPrimitiveWriter.WriteLine(writer, "http://www.w3.org/2000/svg", 1, 2, 3, 4, OfficeColor.Black, 1.5);
            OfficeSvgPrimitiveWriter.WritePath(writer, "http://www.w3.org/2000/svg", "M 0 0 L 10 10", OfficeColor.Green, fill: false, strokeWidth: 3);
        }

        string svg = builder.ToString();
        Assert.Contains("<circle", svg, StringComparison.Ordinal);
        Assert.Contains("fill=\"none\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke=\"#4682B4\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke-linecap=\"round\"", svg, StringComparison.Ordinal);
        Assert.Contains("stroke-linejoin=\"round\"", svg, StringComparison.Ordinal);
        Assert.Contains("<rect", svg, StringComparison.Ordinal);
        Assert.Contains("rx=\"4\"", svg, StringComparison.Ordinal);
        Assert.Contains("fill=\"#FF0000\"", svg, StringComparison.Ordinal);
        Assert.Contains("<line", svg, StringComparison.Ordinal);
        Assert.Contains("<path d=\"M 0 0 L 10 10\"", svg, StringComparison.Ordinal);
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
    public void OfficeShapePresetsCreateReusableDrawingMlGeometry() {
        Assert.True(OfficeShapePresets.TryCreate("triangle", 120, 80, out OfficeShape? triangle));
        Assert.NotNull(triangle);
        Assert.Equal(OfficeShapeKind.Polygon, triangle!.Kind);
        Assert.Equal(120, triangle.Width);
        Assert.Equal(80, triangle.Height);
        Assert.Equal(new OfficePoint(60, 0), triangle.Points[0]);
        Assert.Equal(new OfficePoint(120, 80), triangle.Points[1]);
        Assert.Equal(new OfficePoint(0, 80), triangle.Points[2]);

        Assert.True(OfficeShapePresets.TryCreate("parallelogram", 100, 40, horizontalFlip: true, verticalFlip: false, out OfficeShape? flipped));
        Assert.NotNull(flipped);
        Assert.Equal(new OfficePoint(78, 0), flipped!.Points[0]);
        Assert.Equal(new OfficePoint(0, 0), flipped.Points[1]);

        Assert.True(OfficeShapePresets.TryCreate("rightArrow", 140, 50, out OfficeShape? arrow));
        Assert.NotNull(arrow);
        Assert.Equal(OfficeShapeKind.Polygon, arrow!.Kind);
        Assert.Equal(7, arrow.Points.Count);

        Assert.True(OfficeShapePresets.TryCreate("ShapeTypeValues { InnerText = triangle }", 60, 40, out OfficeShape? openXmlDiagnostic));
        Assert.NotNull(openXmlDiagnostic);
        Assert.Equal(60, openXmlDiagnostic!.Width);

        Assert.True(OfficeShapePresets.TryCreate("hexagon", 120, 80, out OfficeShape? hexagon));
        Assert.NotNull(hexagon);
        Assert.Equal(120, hexagon!.Width);
        Assert.Equal(80, hexagon.Height);

        Assert.True(OfficeShapePresets.TryCreate("star5", 90, 90, out OfficeShape? star));
        Assert.NotNull(star);
        Assert.Equal(90, star!.Width);
        Assert.Equal(90, star.Height);

        Assert.True(OfficeShapePresets.TryCreate("line", 120, 40, out OfficeShape? presetLine));
        Assert.NotNull(presetLine);
        Assert.Equal(OfficeShapeKind.Line, presetLine!.Kind);
        Assert.Equal(new OfficePoint(0, 0), presetLine.Points[0]);
        Assert.Equal(new OfficePoint(120, 0), presetLine.Points[1]);

        Assert.True(OfficeShapePresets.TryCreate("straightConnector1", 120, 40, out OfficeShape? straightConnector));
        Assert.NotNull(straightConnector);
        Assert.Equal(OfficeShapeKind.Line, straightConnector!.Kind);
        Assert.Equal(new OfficePoint(0, 0), straightConnector.Points[0]);
        Assert.Equal(new OfficePoint(120, 40), straightConnector.Points[1]);

        Assert.True(OfficeShapePresets.TryCreate("cloud", 100, 60, out OfficeShape? cloud));
        Assert.NotNull(cloud);
        Assert.Equal(OfficeShapeKind.Path, cloud!.Kind);
        Assert.Contains(cloud.PathCommands, command => command.Kind == OfficePathCommandKind.CubicBezierTo);

        Assert.True(OfficeShapePresets.TryCreate("can", 80, 60, out OfficeShape? can));
        Assert.NotNull(can);
        Assert.Equal(OfficeShapeKind.Path, can!.Kind);
        Assert.Contains(can.PathCommands, command => command.Kind == OfficePathCommandKind.CubicBezierTo);

        Assert.True(OfficeShapePresets.TryCreate("donut", 70, 70, out OfficeShape? donut));
        Assert.NotNull(donut);
        Assert.Equal(OfficeShapeKind.Path, donut!.Kind);
        Assert.True(donut.PathCommands.Count(command => command.Kind == OfficePathCommandKind.Close) >= 2);

        Assert.True(OfficeShapePresets.TryCreate("heart", 64, 56, horizontalFlip: true, verticalFlip: false, out OfficeShape? heart));
        Assert.NotNull(heart);
        AssertPointNear(new OfficePoint(heart!.Width, heart.Height), 58.88D, 52.08D);
        AssertPointNear(heart.PathCommands[0].Point, 29.44D, 52.08D);
        AssertPointNear(heart.PathCommands[1].Point, 57.6D, 17.36D);
        AssertPointNear(heart.PathCommands[5].Point, 1.28D, 17.36D);
        AssertPointNear(heart.PathCommands[3].Point, 29.44D, 14.56D);

        Assert.True(OfficeShapePresets.TryCreate("cube", 72, 60, out OfficeShape? cube));
        Assert.NotNull(cube);
        Assert.Equal(OfficeShapeKind.Polygon, cube!.Kind);

        Assert.True(OfficeShapePresets.TryCreate("leftRightArrow", 96, 40, out OfficeShape? leftRightArrow));
        Assert.NotNull(leftRightArrow);
        Assert.Equal(OfficeShapeKind.Polygon, leftRightArrow!.Kind);
    }

    [Fact]
    public void OfficeGeometryInterpolatesPolylineByLength() {
        var points = new[] {
            new OfficePoint(0, 0),
            new OfficePoint(10, 0),
            new OfficePoint(10, 10)
        };

        Assert.Equal(new OfficePoint(0, 0), OfficeGeometry.InterpolatePolyline(points, -1D));
        Assert.Equal(new OfficePoint(10, 0), OfficeGeometry.InterpolatePolyline(points, 0.5D));
        Assert.Equal(new OfficePoint(10, 10), OfficeGeometry.InterpolatePolyline(points, 2D));
    }

    [Fact]
    public void OfficeGeometryInterpolatesTuplePolylineAndSkipsZeroLengthSegments() {
        var points = new[] {
            (X: 0D, Y: 0D),
            (X: 0D, Y: 0D),
            (X: 0D, Y: 10D),
            (X: 10D, Y: 10D)
        };

        (double x, double y) = OfficeGeometry.InterpolatePolyline(points, 0.25D);
        Assert.Equal(0D, x);
        Assert.Equal(5D, y);
        Assert.Equal(5D, OfficeGeometry.Distance((0D, 0D), (3D, 4D)));
    }

    [Fact]
    public void OfficeGeometryBuildsReusableConnectorPolylines() {
        List<(double X, double Y)> explicitRoute = OfficeGeometry.BuildConnectorPolyline(
            (0D, 0D),
            (10D, 10D),
            new[] { (2D, 3D), (4D, 5D) },
            useRightAngleFallback: true);

        Assert.Equal(new[] { (0D, 0D), (2D, 3D), (4D, 5D), (10D, 10D) }, explicitRoute);

        List<(double X, double Y)> rightAngle = OfficeGeometry.BuildConnectorPolyline(
            (1D, 2D),
            (7D, 9D),
            Array.Empty<(double X, double Y)>(),
            useRightAngleFallback: true);

        Assert.Equal(new[] { (1D, 2D), (1D, 9D), (7D, 9D) }, rightAngle);

        List<(double X, double Y)> straight = OfficeGeometry.BuildConnectorPolyline(
            (1D, 2D),
            (7D, 9D),
            null,
            useRightAngleFallback: false);

        Assert.Equal(new[] { (1D, 2D), (7D, 9D) }, straight);

        List<OfficePoint> officePoints = OfficeGeometry.BuildConnectorPolyline(
            new OfficePoint(0D, 0D),
            new OfficePoint(4D, 4D),
            new[] { new OfficePoint(0D, 4D) },
            useRightAngleFallback: true);

        Assert.Equal(new OfficePoint(0D, 4D), officePoints[1]);
    }

    [Fact]
    public void OfficeGeometrySamplesBezierCurvesForSharedFlattening() {
        List<OfficePoint> quadratic = OfficeGeometry.CreateQuadraticBezierPoints(
            new OfficePoint(0D, 0D),
            new OfficePoint(10D, 20D),
            new OfficePoint(20D, 0D),
            2);

        Assert.Equal(2, quadratic.Count);
        AssertPointNear(quadratic[0], 10D, 10D);
        AssertPointNear(quadratic[1], 20D, 0D);

        List<(double X, double Y)> cubic = OfficeGeometry.CreateCubicBezierPoints(
            (0D, 0D),
            (10D, 30D),
            (20D, 30D),
            (30D, 0D),
            3);

        Assert.Equal(3, cubic.Count);
        AssertPointNear(new OfficePoint(cubic[0].X, cubic[0].Y), 10D, 20D);
        AssertPointNear(new OfficePoint(cubic[1].X, cubic[1].Y), 20D, 20D);
        AssertPointNear(new OfficePoint(cubic[2].X, cubic[2].Y), 30D, 0D);
    }

    [Fact]
    public void OfficeGeometrySamplesEllipticalArcsForSharedRendering() {
        List<OfficePoint> arc = OfficeGeometry.CreateEllipticalArcPoints(
            centerX: 10D,
            centerY: 20D,
            radiusX: 8D,
            radiusY: 4D,
            startRadians: 0D,
            sweepRadians: Math.PI / 2D,
            segments: 2);

        Assert.Equal(2, arc.Count);
        AssertPointNear(arc[0], 10D + (Math.Sqrt(0.5D) * 8D), 20D + (Math.Sqrt(0.5D) * 4D));
        AssertPointNear(arc[1], 10D, 24D);

        List<(double X, double Y)> tuples = OfficeGeometry.CreateEllipticalArcPointsAsTuples(
            centerX: 0D,
            centerY: 0D,
            radiusX: 10D,
            radiusY: 10D,
            startRadians: 0D,
            sweepRadians: Math.PI,
            segments: 2);

        Assert.Equal(2, tuples.Count);
        AssertPointNear(new OfficePoint(tuples[0].X, tuples[0].Y), 0D, 10D);
        AssertPointNear(new OfficePoint(tuples[1].X, tuples[1].Y), -10D, 0D);
    }

    [Fact]
    public void OfficeGeometryRotatesSampledEllipticalArcs() {
        List<OfficePoint> arc = OfficeGeometry.CreateEllipticalArcPoints(
            centerX: 1D,
            centerY: 0D,
            radiusX: 1D,
            radiusY: 1D,
            startRadians: 0D,
            sweepRadians: Math.PI / 2D,
            segments: 1,
            rotationRadians: Math.PI / 2D,
            rotationCenterX: 0D,
            rotationCenterY: 0D);

        Assert.Single(arc);
        AssertPointNear(arc[0], -1D, 1D);

        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeGeometry.CreateEllipticalArcPoints(0D, 0D, 1D, 1D, 0D, Math.PI, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeGeometry.CreateEllipticalArcPoints(0D, 0D, 0D, 1D, 0D, Math.PI, 1));
    }

    [Fact]
    public void OfficeGeometryRotatesPointsAndConvertsAngles() {
        double radians = OfficeGeometry.DegreesToRadians(90D);
        Assert.Equal(90D, OfficeGeometry.RadiansToDegrees(radians), precision: 10);

        OfficePoint rotated = OfficeGeometry.RotatePoint(new OfficePoint(1D, 0D), 0D, 0D, radians);
        Assert.Equal(0D, rotated.X, precision: 10);
        Assert.Equal(1D, rotated.Y, precision: 10);

        (double x, double y) = OfficeGeometry.RotatePoint((1D, 0D), 0D, 0D, -radians);
        Assert.Equal(0D, x, precision: 10);
        Assert.Equal(-1D, y, precision: 10);

        (double left, double top, double right, double bottom) = OfficeGeometry.GetRotatedRectangleBounds(
            x: 0D,
            y: 0D,
            width: 10D,
            height: 20D,
            rotationDegrees: 90D,
            centerX: 5D,
            centerY: 10D);

        Assert.Equal(-5D, left, precision: 10);
        Assert.Equal(5D, top, precision: 10);
        Assert.Equal(15D, right, precision: 10);
        Assert.Equal(15D, bottom, precision: 10);
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
            OfficePathCommand.QuadraticBezierTo(40, 0, 90, 50),
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
        Assert.Equal(50, clone.Height);
        Assert.Equal(OfficePathCommand.MoveTo(0, 50), clone.PathCommands[0]);
        Assert.Equal(OfficePathCommand.QuadraticBezierTo(30, 0, 80, 50), clone.PathCommands[1]);
        Assert.Equal(OfficePathCommand.CubicBezierTo(20, 10, 60, 10, 80, 50), clone.PathCommands[2]);
        Assert.Equal(OfficePathCommand.LineTo(0, 50), clone.PathCommands[3]);
        Assert.Equal(OfficePathCommand.Close(), clone.PathCommands[4]);
        Assert.Equal(OfficeColor.WhiteSmoke, clone.FillColor);
        Assert.Equal(OfficeColor.SteelBlue, clone.StrokeColor);
        Assert.Equal(1.75, clone.StrokeWidth);
        Assert.Equal(OfficeStrokeDashStyle.Dot, clone.StrokeDashStyle);
        Assert.Equal(OfficeStrokeLineCap.Round, clone.StrokeLineCap);
        Assert.Equal(OfficeStrokeLineJoin.Round, clone.StrokeLineJoin);
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

    [Fact]
    public void OfficeTrueTypeFontRejectsInvalidData() {
        Assert.Null(OfficeTrueTypeFont.TryLoad(Array.Empty<byte>()));
        Assert.Null(OfficeTrueTypeFont.TryLoad(new byte[] { 1, 2, 3, 4, 5 }));
    }

    [Fact]
    public void OfficeTrueTypeFontRejectsOversizedTrueTypeCollections() {
        byte[] collection = {
            0x74, 0x74, 0x63, 0x66,
            0x00, 0x01, 0x00, 0x00,
            0x00, 0x00, 0x01, 0x01
        };

        Assert.Null(OfficeTrueTypeFont.TryLoad(collection));
    }

    [Fact]
    public void OfficeTrueTypeFontTreatsMalformedFormat12CmapAsMissingGlyphs() {
        byte[] fontData = CreateMinimalTrueTypeFont(CreateTruncatedFormat12Cmap());
        OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoad(fontData);

        Assert.NotNull(font);
        Assert.Equal(500D, font!.Measure("A", 1000D));
    }

    [Fact]
    public void OfficeTrueTypeFontReadsDefaultFontOutlinesWhenAvailable() {
        OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoadDefault(out string? path);
        if (font == null) {
            return;
        }

        Assert.False(string.IsNullOrWhiteSpace(path));
        Assert.True(font.Measure("OfficeIMO", 18) > 0);
        Assert.True(font.LineHeight(18) > 0);

        List<List<OfficePoint>> contours = font.GetTextContours("OfficeIMO", 0, 0, 18);
        Assert.NotEmpty(contours);
        Assert.Contains(contours, contour => contour.Count >= 3);
    }

    [Fact]
    public void OfficeTrueTypeFontResolvesCssFontFamilyFallbackWhenAvailable() {
        OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoadFontFamily("\"Definitely Missing\", sans-serif", out string? path);
        if (font == null) {
            return;
        }

        Assert.False(string.IsNullOrWhiteSpace(path));
        Assert.True(font.Measure("OfficeIMO", 18) > 0);
    }

    private static byte[] CreateTruncatedFormat12Cmap() {
        var data = new byte[28];
        WriteUInt16(data, 2, 1);
        WriteUInt16(data, 4, 3);
        WriteUInt16(data, 6, 10);
        WriteUInt32(data, 8, 12);
        WriteUInt16(data, 12, 12);
        WriteUInt32(data, 16, 16);
        WriteUInt32(data, 24, 2);
        return data;
    }

    private static byte[] CreateMinimalTrueTypeFont(byte[] cmap) {
        var tables = new List<(string Tag, byte[] Data)> {
            ("cmap", cmap),
            ("glyf", new byte[4]),
            ("head", CreateHeadTable()),
            ("hhea", CreateHheaTable()),
            ("hmtx", new byte[] { 0x01, 0xF4, 0x00, 0x00 }),
            ("loca", new byte[4]),
            ("maxp", new byte[] { 0x00, 0x01, 0x00, 0x00, 0x00, 0x02 })
        };

        int tableDirectoryLength = 12 + tables.Count * 16;
        var offsets = new int[tables.Count];
        int offset = tableDirectoryLength;
        for (int index = 0; index < tables.Count; index++) {
            offsets[index] = offset;
            offset += Align4(tables[index].Data.Length);
        }

        var font = new byte[offset];
        WriteUInt32(font, 0, 0x00010000);
        WriteUInt16(font, 4, (ushort)tables.Count);
        for (int index = 0; index < tables.Count; index++) {
            int record = 12 + index * 16;
            WriteTag(font, record, tables[index].Tag);
            WriteUInt32(font, record + 8, (uint)offsets[index]);
            WriteUInt32(font, record + 12, (uint)tables[index].Data.Length);
            Array.Copy(tables[index].Data, 0, font, offsets[index], tables[index].Data.Length);
        }

        return font;
    }

    private static byte[] CreateHeadTable() {
        var table = new byte[54];
        WriteUInt16(table, 18, 1000);
        return table;
    }

    private static byte[] CreateHheaTable() {
        var table = new byte[36];
        WriteUInt16(table, 4, 800);
        WriteUInt16(table, 6, unchecked((ushort)-200));
        WriteUInt16(table, 34, 1);
        return table;
    }

    private static int Align4(int value) => (value + 3) & ~3;

    private static void WriteTag(byte[] data, int offset, string tag) {
        for (int index = 0; index < 4; index++) {
            data[offset + index] = (byte)tag[index];
        }
    }

    private static void WriteUInt16(byte[] data, int offset, int value) {
        data[offset] = (byte)((value >> 8) & 0xFF);
        data[offset + 1] = (byte)(value & 0xFF);
    }

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)((value >> 24) & 0xFF);
        data[offset + 1] = (byte)((value >> 16) & 0xFF);
        data[offset + 2] = (byte)((value >> 8) & 0xFF);
        data[offset + 3] = (byte)(value & 0xFF);
    }

    [Theory]
    [InlineData("png", OfficeImageFormat.Png)]
    [InlineData(".png", OfficeImageFormat.Png)]
    [InlineData("photo.JPG", OfficeImageFormat.Jpeg)]
    [InlineData("diagram.svg", OfficeImageFormat.Svg)]
    [InlineData("legacy.emf", OfficeImageFormat.Emf)]
    [InlineData("preview.webp", OfficeImageFormat.Webp)]
    public void OfficeImageReaderMapsFileNamesAndBareExtensions(string fileName, OfficeImageFormat expected) {
        Assert.Equal(expected, OfficeImageReader.FromExtension(fileName));
    }

    [Theory]
    [InlineData("photo.jpeg", true)]
    [InlineData(".webp", true)]
    [InlineData("legacy.emf", true)]
    [InlineData("diagram.txt", false)]
    [InlineData("", false)]
    [InlineData(null, false)]
    public void OfficeImageReaderIdentifiesKnownImageExtensions(string? fileName, bool expected) {
        Assert.Equal(expected, OfficeImageReader.IsKnownImageExtension(fileName));
    }

    [Theory]
    [InlineData("image/png", OfficeImageFormat.Png, true)]
    [InlineData("image/jpg", OfficeImageFormat.Jpeg, true)]
    [InlineData("image/jpeg; charset=binary", OfficeImageFormat.Jpeg, true)]
    [InlineData("image/webp", OfficeImageFormat.Webp, false)]
    [InlineData("application/octet-stream", OfficeImageFormat.Unknown, false)]
    public void OfficeImagePdfCompatibilityMapsSupportedContentTypes(string contentType, OfficeImageFormat expectedFormat, bool expectedSupported) {
        bool supported = OfficeImagePdfCompatibility.TryGetSupportedContentTypeFormat(contentType, out OfficeImageFormat format);

        Assert.Equal(expectedSupported, supported);
        Assert.Equal(expectedFormat, format);
        Assert.Equal(expectedSupported, OfficeImagePdfCompatibility.IsSupportedContentType(contentType));
    }

    [Fact]
    public void OfficeImagePdfCompatibilityRejectsDeclaredContentTypeMismatch() {
        bool valid = OfficeImagePdfCompatibility.TryValidateDeclaredContentType(
            OnePixelPng,
            "image/jpeg",
            out OfficeImageInfo? imageInfo,
            out string? unsupportedReason);

        Assert.False(valid);
        Assert.NotNull(imageInfo);
        Assert.Equal(OfficeImageFormat.Png, imageInfo!.Format);
        Assert.Equal("Image bytes were declared as JPEG but were detected as Png.", unsupportedReason);
    }

    [Fact]
    public void OfficeImagePdfCompatibilityRejectsEmptyDeclaredContentTypeBytes() {
        bool valid = OfficeImagePdfCompatibility.TryValidateDeclaredContentType(
            Array.Empty<byte>(),
            "image/png",
            out OfficeImageInfo? imageInfo,
            out string? unsupportedReason);

        Assert.False(valid);
        Assert.Null(imageInfo);
        Assert.Equal("Image bytes are empty.", unsupportedReason);
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
    public void OfficeImageReaderReadsSvgPointUnitsAsCssPixels() {
        var svg = System.Text.Encoding.UTF8.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"72pt\" height=\"36pt\"></svg>");

        Assert.True(OfficeImageReader.TryIdentify(svg, "points.svg", out var image));

        Assert.Equal(OfficeImageFormat.Svg, image.Format);
        Assert.Equal(96, image.Width);
        Assert.Equal(48, image.Height);
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

    private static void AssertPointNear(OfficePoint actual, double expectedX, double expectedY) {
        Assert.Equal(expectedX, actual.X, precision: 6);
        Assert.Equal(expectedY, actual.Y, precision: 6);
    }

    private static int CountOccurrences(string value, string pattern) {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(pattern, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += pattern.Length;
        }

        return count;
    }

    private static void WritePlaceableWmfChecksum(byte[] data) {
        ushort checksum = 0;
        for (int offset = 0; offset < 20; offset += 2) {
            checksum ^= (ushort)(data[offset] | (data[offset + 1] << 8));
        }

        WriteUInt16LittleEndian(data, 20, checksum);
    }
}
