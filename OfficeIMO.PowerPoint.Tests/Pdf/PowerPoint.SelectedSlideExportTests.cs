using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointSelectedSlideExportTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExhaustedImageBudgetRetainsFollowingSolidAndGradientBackgrounds(bool gradient) {
        using var deck = PowerPointPresentation.Create();
        deck.SlideSize.SetSizePoints(100, 60);
        byte[] picture = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.CornflowerBlue));
        deck.AddSlide().AddPicture(new MemoryStream(picture), OfficeImageFormat.Png,
            PowerPointUnits.FromPoints(10), PowerPointUnits.FromPoints(10),
            PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(20));
        var background = deck.AddSlide();
        if (gradient) background.SetBackgroundGradient("123456", "123456", 0);
        else background.BackgroundColor = "123456";
        deck.AddSlide().SetBackgroundImage(new MemoryStream(picture), OfficeImageFormat.Png);

        var images = deck.ExportImages(OfficeImageExportFormat.Png, new PowerPointPresentationImageExportOptions {
            MaximumTotalEmbeddedImageBytes = picture.Length
        });

        Assert.Equal(3, images.Count);
        Assert.True(OfficePngReader.TryDecode(images[1].Bytes, out OfficeRasterImage? image));
        Assert.Equal(OfficeColor.FromRgb(18, 52, 86), image!.GetPixel(5, 5));
        Assert.Single(images[1].Diagnostics, diagnostic => diagnostic.Code == "embedded-image-limit");
    }

    [Fact]
    public void MissingSlideBackgroundUsesTheSharedImageExportBackground() {
        using var deck = PowerPointPresentation.Create();
        deck.SlideSize.SetSizePoints(100, 60);
        var slide = deck.AddSlide();
        byte[] bytes = slide.ToPng(new PowerPointImageExportOptions {
            IncludeSlideBackground = false, IncludeSlideContent = false, BackgroundColor = OfficeColor.AliceBlue
        });
        Assert.True(OfficePngReader.TryDecode(bytes, out OfficeRasterImage? image));
        Assert.Equal(OfficeColor.AliceBlue, image!.GetPixel(5, 5));
    }

    [Fact]
    public void PartlyOffSlideFixedFrameKeepsOverflowVisibleBelowItsOriginalHeight() {
        using var deck = PowerPointPresentation.Create();
        deck.SlideSize.SetSizePoints(220, 180);
        var slide = deck.AddSlide();
        var box = slide.AddTextBoxPoints("FIRST\nSECOND\nTAIL", -5, 20, 100, 12);
        box.FontName = "Helvetica";
        box.FontSize = 12;
        box.TextAutoFit = PowerPointTextAutoFit.None;
        box.TextMarginLeftPoints = 15;
        box.TextMarginRightPoints = box.TextMarginTopPoints = box.TextMarginBottomPoints = 0;
        Assert.True(OfficePngReader.TryDecode(slide.ToPng(), out OfficeRasterImage? image));
        int paintedBelowFrame = 0;
        for (int y = 45; y < 80; y++)
            for (int x = 0; x < 100; x++) {
                var color = image!.GetPixel(x, y);
                if (color.R < 100 && color.G < 100 && color.B < 100) paintedBelowFrame++;
            }
        Assert.True(paintedBelowFrame > 0, "The final line must paint below the fixed text frame while remaining inside the slide.");
    }

    [Fact]
    public void SelectingOneSlideDoesNotReportLossFromUnselectedSlides() {
        using var deck = PowerPointPresentation.Create();
        var good = deck.AddSlide();
        good.AddTextBoxPoints("SELECTED", 20, 20, 200, 50);
        var omitted = deck.AddSlide();
        var first = omitted.AddRectanglePoints(10, 10, 20, 20);
        var second = omitted.AddRectanglePoints(40, 10, 20, 20);
        omitted.GroupShapes(new PowerPointShape[] { first, second }, "Not selected");
        var result = good.ExportImage(OfficeImageExportFormat.Svg, new PowerPointImageExportOptions {
            MaxGroupShapeDepth = 0, Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
        });
        Assert.NotEmpty(result.Bytes);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "group-depth-limit");
        Assert.Contains("SELECTED", string.Concat(good.CreateVisualSnapshot().TextSpans.Select(span => span.Text)));
    }
}
