using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests;

public sealed class WordChartThemePaletteTests {
    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void UnstyledRadialChartProjectsThemeAccentColors(OfficeChartKind kind) {
        using var document = WordDocument.Create();
        WordChart chart = document.AddChart(kind, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Values", new[] { 3d, 2d })
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
}
