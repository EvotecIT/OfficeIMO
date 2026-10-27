using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public sealed class WordChartThemePaletteTests {
    [Theory]
    [InlineData(OfficeChartType.Pie)]
    [InlineData(OfficeChartType.Doughnut)]
    public void UnstyledRadialChartProjectsThemeAccentColors(OfficeChartType kind) {
        using var document = WordDocument.Create();
        WordChart chart = document.AddChart(kind, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeSeries("Values", new[] { 3d, 2d })
        }));
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        A.ColorScheme scheme = document.MainDocumentPartRoot.ThemePart!.Theme!.ThemeElements!.ColorScheme!;
        OfficeColor[] expected = new[] {
            scheme.GetFirstChild<A.Accent1Color>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value!,
            scheme.GetFirstChild<A.Accent2Color>()!.GetFirstChild<A.RgbColorModelHex>()!.Val!.Value!
        }.Select(OfficeColor.Parse).ToArray();
        Assert.Equal(expected, snapshot.Style.Palette.Take(2));
        Assert.Null(snapshot.Data.Series.Single().PointColors);
    }

    [Fact]
    public void RadialPaletteWithNonIdentityWordColorSchemeMappingFailsClosed() {
        using var document = WordDocument.Create();
        WordChart chart = document.AddChart(OfficeChartType.Pie,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeSeries("Values", new[] { 3d, 2d })
            }));
        W.ColorSchemeMapping map = document.MainDocumentPartRoot.DocumentSettingsPart!
            .Settings!.GetFirstChild<W.ColorSchemeMapping>()!;
        map.Accent1 = W.ColorSchemeIndexValues.Accent2;
        map.Accent2 = W.ColorSchemeIndexValues.Accent1;
        string native = chart.ChartPart!.ChartSpace!.OuterXml;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        Assert.Equal(native, chart.ChartPart.ChartSpace.OuterXml);
    }
}
