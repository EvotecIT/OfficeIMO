using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class PowerPointBackgroundSnapshotTests {
    [Fact]
    public void GetBackground_ReturnsReusableSolidGradientAndImageSnapshots() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        PowerPointSlide slide = presentation.Slides[0];

        slide.BackgroundColor = "112233";
        PowerPointSlideBackground solid = slide.GetBackground();
        Assert.Equal(PowerPointSlideBackgroundKind.SolidColor, solid.Kind);
        Assert.Equal("112233", solid.Color);

        slide.SetBackgroundGradient("112233", "445566", 45D);
        PowerPointSlideBackground gradient = slide.GetBackground();
        Assert.Equal(PowerPointSlideBackgroundKind.LinearGradient, gradient.Kind);
        Assert.Equal("112233", gradient.GradientStartColor);
        Assert.Equal("445566", gradient.GradientEndColor);
        Assert.Equal(45D, gradient.GradientAngleDegrees);

        string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");
        slide.SetBackgroundImage(imagePath);
        PowerPointSlideBackground image = slide.GetBackground();
        Assert.Equal(PowerPointSlideBackgroundKind.Image, image.Kind);
        Assert.Equal("image/png", image.ImageContentType);
        Assert.NotNull(image.ImageBytes);
        byte[] imageBytes = image.ImageBytes!;
        Assert.True(imageBytes.Length > 0);

        imageBytes[0] = 0;
        Assert.NotEqual(0, image.ImageBytes![0]);
    }

    [Fact]
    public void GetBackground_PreservesImageSourceCrop() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        PowerPointSlide slide = presentation.Slides[0];
        string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");
        slide.SetBackgroundImage(imagePath);
        A.BlipFill blipFill = slide.SlidePart.Slide.CommonSlideData!.Background!.BackgroundProperties!.GetFirstChild<A.BlipFill>()!;
        blipFill.SourceRectangle = new A.SourceRectangle {
            Left = 25000,
            Top = 10000,
            Right = 5000,
            Bottom = 15000
        };

        PowerPointSlideBackground image = slide.GetBackground();

        Assert.Equal(PowerPointSlideBackgroundKind.Image, image.Kind);
        Assert.True(image.HasImageCrop);
        Assert.Equal(0.25D, image.ImageCropLeft);
        Assert.Equal(0.1D, image.ImageCropTop);
        Assert.Equal(0.05D, image.ImageCropRight);
        Assert.Equal(0.15D, image.ImageCropBottom);
    }
}
