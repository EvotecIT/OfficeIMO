using System;
using System.Linq;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsGradientMappingTests {
    [Theory]
    [InlineData(XpsFormat.Xps, false, null)]
    [InlineData(XpsFormat.OpenXps, false, null)]
    [InlineData(XpsFormat.Xps, true, null)]
    [InlineData(XpsFormat.OpenXps, true, null)]
    [InlineData(XpsFormat.Xps, false, "RelativeToBoundingBox")]
    [InlineData(XpsFormat.OpenXps, false, "RelativeToBoundingBox")]
    [InlineData(XpsFormat.Xps, true, "RelativeToBoundingBox")]
    [InlineData(XpsFormat.OpenXps, true, "RelativeToBoundingBox")]
    public void InvalidNativeGradientMappingIsDiagnosedAcrossStrictConversions(XpsFormat format, bool radial, string? mapping) {
        var document = radial ? XpsRadialGradientTests.Create(format, "1,0,0,1,0,0", false)
            : XpsBrushFixtures.Create("gradient", format);
        var page = document.Pages[0];
        var markup = page.GetMarkup();
        var brush = markup.Descendants().Single(element => element.Name.LocalName ==
            (radial ? "RadialGradientBrush" : "LinearGradientBrush"));
        brush.SetAttributeValue("MappingMode", mapping);
        page.ReplaceMarkup(markup);
        var reopened = XpsDocument.Load(document.Save());

        Assert.Throws<NotSupportedException>(() => reopened.Pages[0].ToSvg());
        Assert.Throws<NotSupportedException>(() => reopened.Pages[0].ToDrawing());
        Assert.Throws<NotSupportedException>(() => reopened.ToPdf());
        var partial = reopened.Pages[0].ToSvg(allowPartial: true);
        Assert.False(partial.IsComplete);
        Assert.Contains("Gradient MappingMode must be Absolute.", partial.Diagnostics);
    }
}
