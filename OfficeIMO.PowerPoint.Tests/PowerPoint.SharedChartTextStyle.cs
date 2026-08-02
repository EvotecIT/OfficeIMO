using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointSharedChartTextStyleTests {
    [Fact]
    public void SharedSnapshot_PreservesExplicitNativeChartFonts() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        var data = new OfficeChartData(
            new[] { "Q1", "Q2" },
            new[] { new OfficeChartSeries("Actual", new[] { 10D, 20D }) });
        PowerPointChart chart = slide.AddChartPoints(
            OfficeChartKind.ColumnClustered, data, 40, 40, 500, 300);
        chart.SetTitle("Trajectory")
            .SetTitleTextStyle(fontName: "Arial")
            .SetLegendTextStyle(fontName: "Arial")
            .SetCategoryAxisLabelTextStyle(fontName: "Arial")
            .SetValueAxisLabelTextStyle(fontName: "Arial");

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal("Arial", snapshot.Style.FontFamily);
        Assert.Equal("Arial", snapshot.Style.TitleFontFamily);
    }

    [Fact]
    public void SharedSnapshot_DoesNotProjectTitleFontOntoChartBodyText() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        var data = new OfficeChartData(
            new[] { "Q1", "Q2" },
            new[] { new OfficeChartSeries("Actual", new[] { 10D, 20D }) });
        PowerPointChart chart = slide.AddChartPoints(
            OfficeChartKind.ColumnClustered, data, 40, 40, 500, 300);
        chart.SetTitle("Trajectory")
            .SetTitleTextStyle(fontName: "Georgia")
            .SetLegendTextStyle(fontName: "Arial")
            .SetCategoryAxisLabelTextStyle(fontName: "Arial")
            .SetValueAxisLabelTextStyle(fontName: "Arial");

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal("Arial", snapshot.Style.FontFamily);
        Assert.Equal("Georgia", snapshot.Style.TitleFontFamily);
    }

    [Fact]
    public void SharedSnapshot_DoesNotFlattenMixedTitleRunFonts() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointChart chart = presentation.AddSlide().AddChartPoints(
            OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "Q1" }, new[] {
                new OfficeChartSeries("Actual", new[] { 10D })
            }), 40, 40, 500, 300);
        chart.SetTitle("First").SetTitleTextStyle(fontName: "Arial");
        C.Chart openXmlChart = presentation.OpenXmlDocument.PresentationPart!
            .SlideParts.Single().ChartParts.Single().ChartSpace!
            .GetFirstChild<C.Chart>()!;
        A.Paragraph paragraph = openXmlChart.GetFirstChild<C.Title>()!
            .GetFirstChild<C.ChartText>()!.GetFirstChild<C.RichText>()!
            .GetFirstChild<A.Paragraph>()!;
        paragraph.Append(new A.Run(
            new A.RunProperties(new A.LatinFont { Typeface = "Georgia" }),
            new A.Text("Second")));

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(OfficeChartStyle.Default.TitleFontFamily,
            snapshot.Style.TitleFontFamily);
    }

    [Fact]
    public void SharedSnapshot_DoesNotPromoteConflictingBodyFonts() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        var data = new OfficeChartData(
            new[] { "Q1", "Q2" },
            new[] { new OfficeChartSeries("Actual", new[] { 10D, 20D }) });
        PowerPointChart chart = slide.AddChartPoints(
            OfficeChartKind.ColumnClustered, data, 40, 40, 500, 300);
        chart.SetLegendTextStyle(fontName: "Arial")
            .SetCategoryAxisLabelTextStyle(fontName: "Georgia")
            .SetValueAxisLabelTextStyle(fontName: "Arial");

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.NotNull(snapshot.Style);
        Assert.Equal(OfficeChartStyle.Default.FontFamily,
            snapshot.Style.FontFamily);
    }

    [Fact]
    public void SharedSnapshot_DoesNotPromoteOnePartialBodyFont() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        var data = new OfficeChartData(
            new[] { "Q1", "Q2" },
            new[] { new OfficeChartSeries("Actual", new[] { 10D, 20D }) });
        PowerPointChart chart = slide.AddChartPoints(
            OfficeChartKind.ColumnClustered, data, 40, 40, 500, 300);
        chart.SetLegendTextStyle(fontName: "Arial");

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(OfficeChartStyle.Default.FontFamily,
            snapshot.Style.FontFamily);
    }

    [Fact]
    public void SharedSnapshot_DoesNotProjectThemeFontTokensAsLiteralFamilies() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        var data = new OfficeChartData(
            new[] { "Q1", "Q2" },
            new[] { new OfficeChartSeries("Actual", new[] { 10D, 20D }) });
        PowerPointChart chart = slide.AddChartPoints(
            OfficeChartKind.ColumnClustered, data, 40, 40, 500, 300);
        chart.SetTitle("Trajectory")
            .SetTitleTextStyle(fontName: "+mj-lt")
            .SetLegendTextStyle(fontName: "+mn-lt")
            .SetCategoryAxisLabelTextStyle(fontName: "+mn-lt")
            .SetValueAxisLabelTextStyle(fontName: "+mn-lt");

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(OfficeChartStyle.Default.FontFamily,
            snapshot.Style.FontFamily);
        Assert.Equal(OfficeChartStyle.Default.TitleFontFamily,
            snapshot.Style.TitleFontFamily);
    }
}
