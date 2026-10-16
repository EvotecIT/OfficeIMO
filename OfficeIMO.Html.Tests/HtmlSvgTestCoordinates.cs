using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Reads PDF-backed SVG geometry in the HTML layout's logical CSS pixels.</summary>
internal static class HtmlSvgTestCoordinates {
    internal static double CssValue(XElement element, string attribute) {
        XElement root = element.AncestorsAndSelf().Last();
        double[] viewBox = root.Attribute("viewBox")!.Value
            .Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(value => double.Parse(value, CultureInfo.InvariantCulture)).ToArray();
        Assert.Equal(4, viewBox.Length);
        Assert.Equal(0D, viewBox[0]);
        Assert.Equal(0D, viewBox[1]);
        double cssPerPoint = HtmlRenderOptions.CssPixelsPerInch / 72D;
        // The output viewport rounds up to whole pixels, while its viewBox retains
        // PDF points. Keep that packaging rounding out of authored baseline checks.
        Assert.Equal(Math.Ceiling(viewBox[2] * cssPerPoint), PixelLength(root, "width"));
        Assert.Equal(Math.Ceiling(viewBox[3] * cssPerPoint), PixelLength(root, "height"));
        return double.Parse(element.Attribute(attribute)!.Value, CultureInfo.InvariantCulture) * cssPerPoint;
    }

    private static double PixelLength(XElement element, string attribute) =>
        double.Parse(element.Attribute(attribute)!.Value.TrimEnd('p', 'x'), CultureInfo.InvariantCulture);
}
