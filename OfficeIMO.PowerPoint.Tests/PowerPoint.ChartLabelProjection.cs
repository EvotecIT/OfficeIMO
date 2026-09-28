using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Xml.Linq;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartLabelProjectionTests {
    [Fact]
    public void StyledLabelBodyRotationIsNotSilentlyDiscarded() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        var slide = presentation.AddSlide();
        PowerPointChart chart = slide.AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A" }, new[] {
                new OfficeChartSeries("Values", new[] { 3d })
            })).SetDataLabels(showValue: true);
        ChartStylePart stylePart = slide.SlidePart.ChartParts.Single()
            .GetPartsOfType<ChartStylePart>().Single();
        XDocument style;
        using (Stream stream = stylePart.GetStream()) style = XDocument.Load(stream);
        XNamespace cs = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
        style.Root!.Element(cs + "dataLabel")!.Add(new XElement(cs + "bodyPr", new XAttribute("rot", "5400000")));
        using (Stream stream = stylePart.GetStream(FileMode.Create, FileAccess.Write)) style.Save(stream);

        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void StyledLabelTextUsesChartLocalTextColorMapping() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        var slide = presentation.AddSlide();
        PowerPointChart chart = slide.AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A" }, new[] {
                new OfficeChartSeries("Values", new[] { 3d })
            })).SetDataLabels(showValue: true);
        ChartPart part = slide.SlidePart.ChartParts.Single();
        part.ChartSpace!.AddChild(new C.ColorMapOverride {
            Background1 = A.ColorSchemeIndexValues.Light1,
            Text1 = A.ColorSchemeIndexValues.Light1,
            Background2 = A.ColorSchemeIndexValues.Light2,
            Text2 = A.ColorSchemeIndexValues.Dark2,
            Accent1 = A.ColorSchemeIndexValues.Accent1,
            Accent2 = A.ColorSchemeIndexValues.Accent2,
            Accent3 = A.ColorSchemeIndexValues.Accent3,
            Accent4 = A.ColorSchemeIndexValues.Accent4,
            Accent5 = A.ColorSchemeIndexValues.Accent5,
            Accent6 = A.ColorSchemeIndexValues.Accent6,
            Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
            FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
        }, true);

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(OfficeColor.White, snapshot.Style.DataLabelTextColor);
        part.ChartSpace.GetFirstChild<C.TextProperties>()?.Remove();
        part.ChartSpace.AddChild(new C.TextProperties(new A.BodyProperties(), new A.ListStyle(),
            new A.Paragraph(new A.ParagraphProperties(new A.DefaultRunProperties(
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Text1 }))))), true);
        Assert.True(chart.TryGetOfficeSnapshot(out snapshot));
        Assert.Equal(OfficeColor.White, snapshot.Style.DataLabelTextColor);
        Assert.Empty(presentation.ValidateDocument());
    }

    [Fact]
    public void ExplicitStyledLabelTextProjectsWithoutStyleInheritance() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        var chart = presentation.AddSlide().AddChart(OfficeChartKind.Pie,
            new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { 3d }) }))
            .SetDataLabels(showValue: true)
            .SetDataLabelTextStyle(fontSizePoints: 9, bold: false, italic: false,
                color: "172033", fontName: "Aptos");
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(9D, snapshot.Layout.DataLabelFontSize);
        Assert.Equal(OfficeColor.Parse("#172033"), snapshot.Style.DataLabelTextColor);
    }

    [Fact]
    public void PartialStyledLabelTextRemainsUnprojectable() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        var chart = presentation.AddSlide().AddChart(OfficeChartKind.Pie,
            new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { 3d }) }))
            .SetDataLabels(showValue: true)
            .SetDataLabelTextStyle(color: "172033");
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

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
