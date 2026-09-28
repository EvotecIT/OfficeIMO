using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartLabelProjectionTests {
    [Theory]
    [InlineData(OfficeChartKind.ColumnClustered)]
    [InlineData(OfficeChartKind.Line)]
    [InlineData(OfficeChartKind.Scatter)]
    public void SharedSnapshot_PreservesNativeDataLabelsOnNonRadialCharts(OfficeChartKind kind) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        var series = new OfficeChartSeries("Values", new[] { 3d, 4d },
            kind == OfficeChartKind.Scatter ? new[] { 1d, 2d } : null);
        PowerPointChart chart = presentation.AddSlide().AddChart(kind,
            new OfficeChartData(new[] { "A", "B" }, new[] { series })).SetDataLabels(showValue: true);
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.True(snapshot.Layout.ShowDataLabels);
        Assert.True(snapshot.Layout.ShowDataLabelValues);
        Assert.Contains(OfficeChartDrawingRenderer.Render(snapshot).Elements.OfType<OfficeDrawingText>(),
            text => text.Text.Contains("3"));
    }

    [Theory]
    [InlineData(OfficeChartKind.ColumnClustered)]
    [InlineData(OfficeChartKind.Line)]
    [InlineData(OfficeChartKind.Scatter)]
    public void SharedSnapshot_ResolvesThemedNonRadialDataLabelText(OfficeChartKind kind) {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        presentation.SetThemeColor(PowerPointThemeColor.Accent3, "2468AC");
        var series = new OfficeChartSeries("Values", new[] { 3d, 4d },
            kind == OfficeChartKind.Scatter ? new[] { 1d, 2d } : null);
        PowerPointChart chart = presentation.AddSlide().AddChart(kind,
            new OfficeChartData(new[] { "A", "B" }, new[] { series })).SetDataLabels(showValue: true);
        ChartPart chartPart = presentation.Slides.Single().SlidePart.ChartParts.Single();
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot styled));
        Assert.Equal(11.97, styled.Layout.DataLabelFontSize);
        Assert.Equal(presentation.OpenXmlDocument.PresentationPart.ThemePart!.Theme.ThemeElements!
            .FontScheme!.MinorFont!.LatinFont!.Typeface!.Value, styled.Layout.DataLabelFontFamily);
        Assert.Equal(OfficeColor.Parse("#404040"), styled.Style.DataLabelTextColor);
        foreach (ChartStylePart stylePart in chartPart.GetPartsOfType<ChartStylePart>().ToArray())
            chartPart.DeletePart(stylePart);
        Assert.True(chart.TryGetOfficeSnapshot(out _));
        var labels = chartPart.ChartSpace!
            .Descendants<C.DataLabels>().Single();
        labels.AddChild(new C.TextProperties(new A.BodyProperties(), new A.ListStyle(),
            new A.Paragraph(new A.ParagraphProperties(new A.DefaultRunProperties(
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Accent3 }))))), true);
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(OfficeColor.Parse("#2468AC"), snapshot.Style.DataLabelTextColor);
    }
}
