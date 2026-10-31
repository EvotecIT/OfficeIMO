using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests;

public sealed class PowerPointBackgroundColorTests {
    [Fact]
    public void SemanticBackgroundKeepsItsNativeThemeReferenceAfterSave() {
        using var document = PowerPointDocument.Create(options: new PowerPointCreateOptions { PersistenceMode = DocumentPersistenceMode.Explicit });
        var slide = document.AddSlide();
        slide.SetBackgroundColor(OfficeColor.Accent1);
        var fill = slide.SlidePart.Slide.CommonSlideData!.Background!.BackgroundProperties!.GetFirstChild<A.SolidFill>()!;
        Assert.Equal(A.SchemeColorValues.Accent1, fill.SchemeColor!.Val!.Value);
        Assert.Null(fill.RgbColorModelHex);
        document.SetThemeColor(PowerPointThemeColor.Accent1, "123456");
        Assert.Equal("123456", slide.GetBackground().Color);
        using var input = new MemoryStream(document.ToBytes());
        using var reopened = PowerPointDocument.Load(input);
        Assert.Equal("123456", reopened.Slides[0].GetBackground().Color);
        Assert.Empty(new OpenXmlValidator().Validate(reopened.OpenXmlDocument));
    }

    [Fact]
    public void LiteralBackgroundRetainsTransparencyThroughTheNativePackage() {
        using var document = PowerPointDocument.Create(options: new PowerPointCreateOptions { PersistenceMode = DocumentPersistenceMode.Explicit });
        var slide = document.AddSlide();
        slide.SetBackgroundColor(OfficeColor.FromRgba(17, 34, 51, 128));
        Assert.Equal("11223380", slide.GetBackground().Color);
        using var input = new MemoryStream(document.ToBytes());
        using var reopened = PowerPointDocument.Load(input);
        Assert.Equal("11223380", reopened.Slides[0].GetBackground().Color);
        Assert.Empty(new OpenXmlValidator().Validate(reopened.OpenXmlDocument));
    }

    [Fact]
    public void ClearingSolidGradientNoFillAndImageBackgroundsRestoresInheritance() {
        using var document = PowerPointDocument.Create(options: new PowerPointCreateOptions { PersistenceMode = DocumentPersistenceMode.Explicit });
        var slide = document.AddSlide();
        var inherited = document.AddSlide().GetBackground();
        Action[] backgrounds = [
            () => slide.SetBackgroundColor(OfficeColor.Accent2),
            () => slide.SetBackgroundGradient("112233", "445566"),
            () => slide.SetBackgroundNoFill(),
            () => {
                using var pixel = new MemoryStream(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));
                slide.SetBackgroundImage(pixel, OfficeImageFormat.Png);
            }
        ];
        foreach (var background in backgrounds) {
            background();
            slide.ClearBackground();
            Assert.Null(slide.SlidePart.Slide.CommonSlideData!.Background);
            Assert.Equal(inherited.Kind, slide.GetBackground().Kind);
            Assert.Equal(inherited.Color, slide.GetBackground().Color);
            Assert.Empty(slide.SlidePart.ImageParts);
        }
    }
}
