using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartImageExportTests {
    [Fact]
    public void LoadedChartRendersAloneAtFrameAndRequestedSize() {
        using var bytes = new MemoryStream();
        using (var deck = PowerPointDocument.Create(bytes, new PowerPointCreateOptions())) {
            var chart = deck.AddSlide().AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "Q1", "Q2" },
                new[] { new OfficeSeries("Revenue", new[] { 12D, 18D }) }));
            chart.Width = OfficeLength.Inches(5);
            chart.Height = OfficeLength.Inches(3);
            deck.Save();
        }

        bytes.Position = 0;
        using var loaded = PowerPointDocument.Load(bytes, new PowerPointLoadOptions());
        PowerPointChart reloaded = Assert.Single(loaded.Slides[0].Charts);
        OfficeImageExportResult png = reloaded.ToImage(new OfficeChartImageExportOptions { Scale = 2 });
        Assert.Equal(OfficeImageExportFormat.Png, png.Format);
        Assert.Equal((960, 576), (png.Width, png.Height));

        using var stream = new MemoryStream();
        OfficeConversionReport report = reloaded.SaveAsImage(stream, new OfficeChartImageExportOptions {
            Format = OfficeImageExportFormat.Svg, Width = OfficeLength.Inches(4), Height = OfficeLength.Inches(2.5)
        });
        // Font substitution is reported as loss where the chart's theme fonts are not installed, so only
        // the written image is part of this contract.
        Assert.Equal(1L, report.ItemsWritten);
        Assert.True(OfficeImageReader.TryIdentifyByContent(stream.ToArray(), ".svg", out var svg));
        Assert.Equal((384, 240), (svg.Width, svg.Height));
    }

    [Fact]
    public void LoadedChartResolvesSchemeAndAutomaticColorsAgainstTheSlideTheme() {
        using var bytes = new MemoryStream();
        using (var deck = PowerPointDocument.Create(bytes, new PowerPointCreateOptions())) {
            deck.AddSlide().AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "Q1", "Q2" }, new[] {
                new OfficeSeries("North", new[] { 4D, 6D }),
                new OfficeSeries("South", new[] { 5D, 3D }),
                new OfficeSeries("East", new[] { 2D, 7D })
            }));
            deck.Save();
        }
        bytes.Position = 0;
        using (var package = PresentationDocument.Open(bytes, true)) {
            PresentationPart presentation = package.PresentationPart!;
            SlidePart slide = presentation.SlideParts.Single();
            ChartImageThemeFixture.Apply(slide.SlideLayoutPart!.SlideMasterPart!.ThemePart!, slide.ChartParts);
        }

        bytes.Position = 0;
        using var loaded = PowerPointDocument.Load(bytes, new PowerPointLoadOptions());
        string svg = Encoding.UTF8.GetString(Assert.Single(loaded.Slides[0].Charts)
            .ToImage(new OfficeChartImageExportOptions { Format = OfficeImageExportFormat.Svg }).Bytes);
        ChartImageThemeFixture.AssertSourceThemeColors(svg);
    }
}
