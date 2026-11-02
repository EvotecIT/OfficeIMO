using OfficeIMO.PowerPoint;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointDesignerFooterTests {
    [Theory]
    [InlineData(PowerPointSlideSizePreset.Screen16x9)]
    [InlineData(PowerPointSlideSizePreset.Screen4x3)]
    public void DesignerRightFooterIsSavedBesideTheSlideNumberWithoutLeakingToLaterSlides(PowerPointSlideSizePreset size) {
        using var document = PowerPointDocument.Create();
        document.SlideSize.SetPreset(size);
        var cards = new[] { new PowerPointCardContent("Demand", new[] { "23 quarterly tickets" }) };
        document.AddDesignerCardGridSlide("Demand review", null, cards,
            options: new PowerPointCardGridSlideOptions { FooterLeft = "31 Dec 2026", FooterRight = "Supplied CSV" });
        document.AddDesignerCardGridSlide("Next review", null, cards);

        using var stream = new MemoryStream();
        document.Save(stream);
        stream.Position = 0;
        using var reopened = PowerPointDocument.Load(stream);
        PowerPointSlide slide = reopened.Slides[0];
        PowerPointTextBox left = Assert.Single(slide.TextBoxes, box => box.Text == "31 Dec 2026");
        PowerPointTextBox right = Assert.Single(slide.TextBoxes, box => box.Text == "Supplied CSV");
        PowerPointShape number = slide.GetPlaceholder(PowerPointPlaceholderType.SlideNumber)!;
        Assert.NotNull(number);
        Assert.True(left.Right <= right.Left);
        Assert.True(right.Right <= number.Left);
        Assert.True(number.Right <= reopened.SlideSize.Width);
        Assert.DoesNotContain(reopened.Slides[1].TextBoxes, box => box.Text == "Supplied CSV");
    }
}
