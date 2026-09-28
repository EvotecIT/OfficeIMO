using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartRadialOutsideLabelsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeOutsideLabelsProjectTheirLeaderLineSetting(bool showLeaderLines) {
        using var document = WordDocument.Create();
        WordChart chart = document.AddChart(OfficeChartKind.Pie,
            new OfficeChartData(new[] { "A", "B", "C" }, new[] {
                new OfficeChartSeries("Status", new[] { 3d, 4d, 5d })
            }));
        C.PieChart native = chart.ChartPart!.ChartSpace!.Descendants<C.PieChart>().Single();
        C.DataLabels labels = native.GetFirstChild<C.DataLabels>()!;
        labels.GetFirstChild<C.DataLabelPosition>()?.Remove();
        labels.AddChild(new C.DataLabelPosition { Val = C.DataLabelPositionValues.OutsideEnd }, true);
        labels.GetFirstChild<C.ShowCategoryName>()?.Remove();
        labels.AddChild(new C.ShowCategoryName { Val = true }, true);
        labels.GetFirstChild<C.ShowLeaderLines>()?.Remove();
        labels.AddChild(new C.ShowLeaderLines { Val = showLeaderLines }, true);
        string nativeXml = chart.ChartPart.ChartSpace.OuterXml;

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(OfficeChartDataLabelPosition.OutsideEnd, snapshot.Layout.DataLabelPosition);
        Assert.Equal(showLeaderLines, snapshot.Layout.ShowDataLabelLeaderLines);
        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot);
        Assert.Equal(showLeaderLines ? 3 : 0,
            drawing.Shapes.Count(shape => shape.Shape.Kind == OfficeShapeKind.Line));
        Assert.Equal(nativeXml, chart.ChartPart.ChartSpace.OuterXml);
    }
}
