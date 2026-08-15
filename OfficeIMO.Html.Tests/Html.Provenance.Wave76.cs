using AngleSharp.Html.Dom;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlProvenanceWave76Tests {
    [Theory]
    [InlineData("<math><mi><svg><image href='new' xlink:href='old'></image></svg></mi></math>")]
    [InlineData("<math><annotation-xml encoding='text&#x2F;html'><svg><image href='new' xlink:href='old'></image></svg></annotation-xml></math>")]
    [InlineData("<math><annotation-xml encoding='application/xhtml+xml'><svg><image href='new' xlink:href='old'></image></svg></annotation-xml></math>")]
    public void SvgHrefNormalizationHonorsMathMlHtmlIntegrationPoints(string html) {
        IHtmlDocument document = HtmlDocumentParser.ParseDocument(html);
        var image = Assert.IsAssignableFrom<AngleSharp.Dom.IElement>(document.QuerySelector("image"));

        Assert.Equal("new", HtmlDocumentParser.GetExactAttributeValue(image, "href"));
        Assert.Equal("old", HtmlDocumentParser.GetExactAttributeValue(image, "xlink:href"));
    }
}
