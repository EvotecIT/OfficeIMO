using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartAppearanceIntegrityTests {
    [Theory]
    [InlineData(OfficeChartKind.Line)]
    [InlineData(OfficeChartKind.Radar)]
    [InlineData(OfficeChartKind.Scatter)]
    public void SharedAppearance_PowerPointRestoresLineAndReplacesImportedMarkerFill(OfficeChartKind kind) {
        OfficeChartData Data(bool connect, OfficeColor? color = null) => new OfficeChartData(new[] { "1", "2" },
            new[] { new OfficeChartSeries("Status", new[] { 1d, 2d }, kind == OfficeChartKind.Scatter ? new[] { 1d, 2d } : null,
                color, null, true, connectLine: connect, markerSize: 1) });
        using PowerPointPresentation document = PowerPointPresentation.Create();
        var chart = document.AddSlide().AddChart(kind, Data(false));
        chart.UpdateData(Data(true));
        var part = document.Slides.Single().SlidePart.ChartParts.Single();
        var series = part.ChartSpace!.Descendants<DocumentFormat.OpenXml.OpenXmlCompositeElement>().Single(e => e.LocalName == "ser");
        Assert.Null(series.GetFirstChild<C.ChartShapeProperties>()!.GetFirstChild<A.Outline>()!.GetFirstChild<A.NoFill>());
        series.GetFirstChild<C.Marker>()!.AddChild(new C.ChartShapeProperties(new A.NoFill()), true);
        chart.UpdateData(Data(true, OfficeColor.Parse("#168A56")));
        series = part.ChartSpace.Descendants<DocumentFormat.OpenXml.OpenXmlCompositeElement>().Single(e => e.LocalName == "ser");
        Assert.Null(series.GetFirstChild<C.Marker>()!.ChartShapeProperties!.GetFirstChild<A.NoFill>());
        Assert.Equal((byte)2, series.GetFirstChild<C.Marker>()!.Size!.Val!.Value);
        Assert.Empty(document.ValidateDocument());
        using var bytes = new MemoryStream(document.ToBytes());
        using PowerPointPresentation reopened = PowerPointPresentation.Load(bytes);
        Assert.Empty(reopened.ValidateDocument());
    }
}
