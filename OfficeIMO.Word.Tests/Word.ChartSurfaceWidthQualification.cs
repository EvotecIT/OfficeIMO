using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartSurfaceWidthQualificationTests {
    [Theory]
    [InlineData("chart")]
    [InlineData("plot")]
    [InlineData("legend")]
    [InlineData("axis")]
    [InlineData("grid")]
    public void VisibleZeroWidthSurface_ProjectsAsHairlineWithoutChangingNativeChart(string target) {
        using var document = WordDocument.Create();
        WordChart chart = document.AddChart(OfficeChartKind.Line,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Status", new[] { 7d, 3d })
            }));
        C.ChartSpace space = chart.ChartPart!.ChartSpace!;
        C.Chart native = space.GetFirstChild<C.Chart>()!;
        var outline = new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" })) {
            Width = 0
        };
        if (target == "chart") space.AddChild(new C.ShapeProperties(outline), true);
        else if (target == "plot") native.PlotArea!.AddChild(new C.ShapeProperties(outline), true);
        else {
            DocumentFormat.OpenXml.OpenXmlCompositeElement owner = target switch {
                "legend" => native.GetFirstChild<C.Legend>()!,
                "axis" => native.PlotArea!.Elements<C.ValueAxis>().Single(),
                _ => native.PlotArea!.Elements<C.ValueAxis>().Single().GetFirstChild<C.MajorGridlines>()!
            };
            owner.AddChild(new C.ChartShapeProperties(outline), true);
        }
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        double? projectedWidth = target switch {
            "chart" => snapshot.Style.ChartBorderWidth,
            "plot" => snapshot.Style.PlotAreaBorderWidth,
            "legend" => snapshot.Style.LegendBorderWidth,
            "axis" => snapshot.Style.ValueAxisLineWidth,
            _ => snapshot.Style.ValueGridLineWidth
        };
        Assert.Equal(0.25, projectedWidth);
        Assert.Equal(0, outline.Width!.Value);
    }

    [Fact]
    public void NoFillZeroWidthOutline_RemainsProjectable() {
        using var document = WordDocument.Create();
        WordChart chart = document.AddChart(OfficeChartKind.Line,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Status", new[] { 7d, 3d })
            }));
        chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!
            .AddChild(new C.ShapeProperties(new A.Outline(new A.NoFill()) { Width = 0 }), true);
        Assert.True(chart.TryGetOfficeSnapshot(out _));
    }
}
