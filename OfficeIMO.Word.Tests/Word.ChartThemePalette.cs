using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public sealed class WordChartThemePaletteTests {
    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void UnstyledRadialChartBeyondQualifiedThemePaletteFailsClosed(OfficeChartKind kind) {
        using var document = WordDocument.Create();
        WordChart chart = document.AddChart(kind, new OfficeChartData(
            new[] { "A", "B", "C", "D", "E", "F", "G" },
            new[] { new OfficeChartSeries("Values", new[] { 1d, 2d, 3d, 4d, 5d, 6d, 7d }) }));
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

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

    [Fact]
    public void RadialPaletteUsesWordColorSchemeMapping() {
        using var document = WordDocument.Create();
        WordChart chart = document.AddChart(OfficeChartKind.Pie,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Values", new[] { 3d, 2d })
            }));
        W.ColorSchemeMapping map = document.MainDocumentPartRoot.DocumentSettingsPart!
            .Settings!.GetFirstChild<W.ColorSchemeMapping>()!;
        map.Accent1 = W.ColorSchemeIndexValues.Accent2;
        map.Accent2 = W.ColorSchemeIndexValues.Accent1;
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        A.ColorScheme scheme = document.MainDocumentPartRoot.ThemePart!.Theme!.ThemeElements!.ColorScheme!;
        Assert.Equal(OfficeColor.Parse(scheme.GetFirstChild<A.Accent2Color>()!
            .GetFirstChild<A.RgbColorModelHex>()!.Val!.Value!), snapshot.Style.Palette[0]);
        Assert.Equal(OfficeColor.Parse(scheme.GetFirstChild<A.Accent1Color>()!
            .GetFirstChild<A.RgbColorModelHex>()!.Val!.Value!), snapshot.Style.Palette[1]);
    }
}
