using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentSlideNamingTests {
    [Fact]
    public void GeneratedSlideNamesFillTheLowestFreeIndexAndSurviveSave() {
        var presentation = OdpPresentation.Create();
        for (int i = 0; i < 4; i++) presentation.AddSlide();
        presentation.RemoveSlide(1);
        presentation.AddSlide("Slide01");
        presentation.AddSlide("Intro");

        Assert.Equal("Slide2", presentation.AddSlide().Name);
        Assert.Equal("Slide5", presentation.AddSlide().Name);

        OdpPresentation loaded = OdpPresentation.Load(new MemoryStream(presentation.ToBytes()));
        Assert.Equal(new[] { "Slide1", "Slide3", "Slide4", "Slide01", "Intro", "Slide2", "Slide5" },
            loaded.Slides.Select(slide => slide.Name).ToArray());
    }

    [Fact]
    public void ExplicitSlideNameMustBeUnique() {
        var presentation = OdpPresentation.Create();
        presentation.AddSlide();
        presentation.AddSlide("Summary");

        Assert.Throws<InvalidOperationException>(() => presentation.AddSlide("Slide1"));
        Assert.Throws<InvalidOperationException>(() => presentation.AddSlide("Summary"));
        Assert.Equal(2, presentation.Slides.Count);
    }
}
