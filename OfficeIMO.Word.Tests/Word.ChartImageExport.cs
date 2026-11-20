using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public sealed class WordChartImageExportTests {
    [Fact]
    public void ChartSavedAndLoadedFromFileRendersAtFrameAndRequestedSize() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-word-chart-" + Guid.NewGuid().ToString("N") + ".docx");
        try {
            using (var document = WordDocument.Create(path)) {
                var chart = document.AddChart("Title");
                chart.AddCategories(new List<string> { "a", "b" });
                chart.AddBar("Series", new List<int> { 1, 2 }, OfficeColor.Parse("#4472C4"));
                document.Save();
            }

            using var loaded = WordDocument.Load(path);
            var reloaded = Assert.Single(loaded.Charts);
            Assert.True(reloaded.TryGetOfficeSnapshot(out var snapshot));
            Assert.Equal(new[] { "a", "b" }, snapshot.Data.Categories);

            OfficeImageExportResult png = reloaded.ToImage();
            Assert.Equal(OfficeImageExportFormat.Png, png.Format);
            Assert.Equal((int)Math.Round(snapshot.Width.ToPixels()), png.Width);
            Assert.True(OfficeImageReader.TryIdentifyByContent(png.Bytes, ".png", out var info));
            Assert.Equal(png.Width, info.Width);

            OfficeImageExportResult svg = reloaded.ToImage(new OfficeChartImageExportOptions {
                Format = OfficeImageExportFormat.Svg, Width = OfficeLength.Inches(4), Height = OfficeLength.Inches(2.5)
            });
            Assert.Equal((384, 240), (svg.Width, svg.Height));
            Assert.Contains("#4472C4", Encoding.UTF8.GetString(svg.Bytes), StringComparison.OrdinalIgnoreCase);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void LoadedChartResolvesSchemeAndAutomaticColorsAgainstTheDocumentTheme() {
        using var bytes = new MemoryStream();
        using (var document = WordDocument.Create()) {
            document.AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "Q1", "Q2" }, new[] {
                new OfficeSeries("North", new[] { 4D, 6D }),
                new OfficeSeries("South", new[] { 5D, 3D }),
                new OfficeSeries("East", new[] { 2D, 7D })
            }));
            document.Save(bytes);
        }
        using (var package = WordprocessingDocument.Open(bytes, true)) {
            ChartImageThemeFixture.Apply(package.MainDocumentPart!.ThemePart!, package.MainDocumentPart.ChartParts);
        }

        bytes.Position = 0;
        using var loaded = WordDocument.Load(bytes);
        string svg = Encoding.UTF8.GetString(Assert.Single(loaded.Charts)
            .ToImage(new OfficeChartImageExportOptions { Format = OfficeImageExportFormat.Svg }).Bytes);
        ChartImageThemeFixture.AssertSourceThemeColors(svg);
    }

    [Fact]
    public void UnsupportedChartReportsWhyItCannotBeRendered() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "A" },
            new[] { new OfficeSeries("Value", new[] { 1D }) }));
        var settings = document._wordprocessingDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!;
        settings.GetFirstChild<W.ColorSchemeMapping>()?.Remove();
        settings.Append(new W.ColorSchemeMapping {
            Background1 = W.ColorSchemeIndexValues.Dark1, Text1 = W.ColorSchemeIndexValues.Light1,
            Background2 = W.ColorSchemeIndexValues.Light2, Text2 = W.ColorSchemeIndexValues.Dark2,
            Accent1 = W.ColorSchemeIndexValues.Accent1, Accent2 = W.ColorSchemeIndexValues.Accent2,
            Accent3 = W.ColorSchemeIndexValues.Accent3, Accent4 = W.ColorSchemeIndexValues.Accent4,
            Accent5 = W.ColorSchemeIndexValues.Accent5, Accent6 = W.ColorSchemeIndexValues.Accent6,
            Hyperlink = W.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = W.ColorSchemeIndexValues.FollowedHyperlink
        });

        var error = Assert.Throws<NotSupportedException>(() => chart.ToImage());
        Assert.Contains("remaps theme color slots", error.Message);
    }
}
