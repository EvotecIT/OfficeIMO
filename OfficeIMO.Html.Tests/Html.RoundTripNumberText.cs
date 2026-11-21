using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlRoundTripNumberTextTests {
    // Computed CSS numbers use the shortest round-trip digits; .NET Framework's "R" wrote this opacity with 17 digits.
    [Fact]
    public void ComputedOpacityUsesTheShortestRoundTripDigits() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse("<div id=\"target\" style=\"opacity:0.4694715627858908\">Text</div>");
        OfficeIMO.Html.Dom.HtmlElement target = document.Document.QuerySelector("#target")!;

        Assert.Equal("0.4694715627858908", HtmlComputedStyleEngine.Compute(document)[target].GetValue("opacity"));
    }
}
