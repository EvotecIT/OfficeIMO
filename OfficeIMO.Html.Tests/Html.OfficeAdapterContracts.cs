using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlOfficeAdapters {
    [Fact]
    public void PowerPointHtml_RendersHeadingOnlySemanticSections() {
        HtmlToPowerPointResult result = HtmlConversionDocument
            .Parse("<p>First slide</p><h1>Heading-only slide</h1>")
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.Value;

        Assert.Equal(2, result.Slides);
        PowerPointSlide second = presentation.Slides[1];
        Assert.Contains(second.TextBoxes, textBox =>
            textBox.Text.Contains("Heading-only slide", StringComparison.Ordinal));
    }
}
