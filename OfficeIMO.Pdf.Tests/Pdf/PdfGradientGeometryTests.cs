using System;
using System.Globalization;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfGradientGeometryTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ShadingCoordinatesRetainSmallDifferencesAndValidPdfDecimalSyntax(bool radial) {
        var stops = new[] { new OfficeGradientStop(0, OfficeColor.Red), new OfficeGradientStop(1, OfficeColor.Blue) };
        double[] expected = radial ? new[] { -.000000005, 1.000000005, .000000004, 100000000000000000000D, -.000000003, .000000002 }
            : new[] { -.000000005, 1.000000005, 100000000000000000000D, -.000000003 };
        string shading = radial ? PdfVisualResourceDictionaryBuilder.BuildRadialShadingObject(
            expected[0], expected[1], expected[2], expected[3], expected[4], expected[5], stops)
            : PdfVisualResourceDictionaryBuilder.BuildAxialShadingObject(expected[0], expected[1], expected[2], expected[3], stops);
        int start = shading.IndexOf("/Coords [", StringComparison.Ordinal) + "/Coords [".Length;
        int end = shading.IndexOf(']', start);
        var actual = shading.Substring(start, end - start).Split(' ')
            .Select(value => double.Parse(value, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture)).ToArray();
        Assert.Equal(expected, actual);
    }
}
