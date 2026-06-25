using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingSparklineTests {
    [Fact]
    public void OfficeSparklineRendererScalesCrossingColumnBarsAroundZeroAxis() {
        var builder = new StringBuilder();
        OfficeSparklineRenderer.AppendSvg(
            builder,
            x: 0D,
            y: 0D,
            width: 20D,
            height: 16D,
            values: new[] { 12D, -4D },
            kind: OfficeSparklineKind.Column,
            style: new OfficeSparklineStyle {
                DisplayAxis = true,
                Padding = 0D,
                AxisInset = 0D,
                ColumnWidthRatio = 1D,
                SeriesColor = OfficeColor.Black
            });

        IReadOnlyList<SvgRect> bars = ParseSvgRects(builder.ToString());

        Assert.Equal(2, bars.Count);
        Assert.Equal(0D, bars[0].Y, precision: 6);
        Assert.Equal(12D, bars[0].Height, precision: 6);
        Assert.Equal(12D, bars[1].Y, precision: 6);
        Assert.Equal(4D, bars[1].Height, precision: 6);
    }

    private static IReadOnlyList<SvgRect> ParseSvgRects(string svg) {
        MatchCollection matches = Regex.Matches(svg, "<rect\\b[^>]*>");
        var rectangles = new List<SvgRect>(matches.Count);
        foreach (Match match in matches) {
            rectangles.Add(new SvgRect(
                ReadNumber(match.Value, "y"),
                ReadNumber(match.Value, "height")));
        }

        return rectangles;
    }

    private static double ReadNumber(string element, string attribute) {
        Match match = Regex.Match(element, attribute + "=\"([^\"]+)\"");
        Assert.True(match.Success, "SVG element did not contain attribute '" + attribute + "': " + element);
        return double.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
    }

    private readonly record struct SvgRect(double Y, double Height);
}
