using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

internal static class ChartImageThemeFixture {
    private const string Accent1 = "C0392B";
    private const string Accent2 = "16A085";
    private const string Accent4 = "D35400";

    /// <summary>Gives the theme non-default accents and leaves the first two series automatic and the third on accent4.</summary>
    internal static void Apply(ThemePart theme, IEnumerable<ChartPart> charts) {
        A.ColorScheme scheme = theme.Theme!.ThemeElements!.ColorScheme!;
        scheme.Accent1Color = new A.Accent1Color(new A.RgbColorModelHex { Val = Accent1 });
        scheme.Accent2Color = new A.Accent2Color(new A.RgbColorModelHex { Val = Accent2 });
        scheme.Accent4Color = new A.Accent4Color(new A.RgbColorModelHex { Val = Accent4 });
        theme.Theme.Save();
        foreach (ChartPart chart in charts) {
            var series = chart.ChartSpace!.Descendants<C.BarChartSeries>().ToList();
            Assert.Equal(3, series.Count);
            foreach (C.BarChartSeries item in series) {
                item.GetFirstChild<C.ChartShapeProperties>()?.Remove();
                foreach (C.DataPoint point in item.Elements<C.DataPoint>().ToList()) point.Remove();
            }
            series[2].InsertAfter(new C.ChartShapeProperties(new A.SolidFill(
                new A.SchemeColor { Val = A.SchemeColorValues.Accent4 })), series[2].SeriesText);
            chart.ChartSpace.Save();
        }
    }

    internal static void AssertSourceThemeColors(string svg) {
        Assert.Contains("#" + Accent1, svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#" + Accent2, svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#" + Accent4, svg, StringComparison.OrdinalIgnoreCase);
        int first = svg.IndexOf("#" + Accent1, StringComparison.OrdinalIgnoreCase);
        int second = svg.IndexOf("#" + Accent2, StringComparison.OrdinalIgnoreCase);
        Assert.True(first < second, "Automatic series colors follow the theme accents in series order.");
    }
}
