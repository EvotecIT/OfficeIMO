using OfficeIMO.Drawing;
using OfficeIMO.Markup;
using OfficeIMO.Markup.PowerPoint;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests.Markup;

public class OfficeMarkupPowerPointChartParityTests {
    [Theory]
    [InlineData("line", OfficeChartKind.Line)]
    [InlineData("bar", OfficeChartKind.BarClustered)]
    [InlineData("clustered-bar", OfficeChartKind.BarClustered)]
    [InlineData("stackedbar", OfficeChartKind.BarStacked)]
    [InlineData("stacked-column", OfficeChartKind.ColumnStacked)]
    [InlineData("pie", OfficeChartKind.Pie)]
    [InlineData("donut", OfficeChartKind.Doughnut)]
    [InlineData("scatter", OfficeChartKind.Scatter)]
    [InlineData("area", OfficeChartKind.Area)]
    [InlineData("unknown", OfficeChartKind.ColumnClustered)]
    public void PowerPointChartTokensMatchBetweenDirectExportAndCSharpEmitter(
        string chartType, OfficeChartKind expectedKind) {
        string markup = $$"""
            ---
            profile: presentation
            ---

            # Chart parity

            @slide {
              layout: blank
            }

            ::chart type={{chartType}} title="Parity"
            X,Value
            1,10
            2,20
            3,30
            """;
        OfficeMarkupParseResult parsed = OfficeMarkupParser.Parse(markup);
        Assert.False(parsed.HasErrors);
        string emittedCode = new OfficeMarkupCSharpEmitter().Emit(parsed.Document);
        Assert.Contains($"OfficeChartKind.{expectedKind}", emittedCode, StringComparison.Ordinal);

        string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
        try {
            new OfficeMarkupPowerPointExporter().Export(parsed.Document,
                new OfficeMarkupPowerPointExportOptions {
                    OutputPath = outputPath,
                    RenderMermaidDiagrams = false
                });

            using PowerPointPresentation presentation = PowerPointPresentation.Load(
                outputPath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Core.DocumentAccessMode.ReadOnly });
            PowerPointChart chart = Assert.Single(presentation.Slides.SelectMany(slide => slide.Charts));
            Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
            Assert.Equal(expectedKind, snapshot.ChartKind);
        } finally {
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }
}
