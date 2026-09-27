using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartOfficeSnapshotTests {
    [Fact]
    public void OfficeSnapshot_PreservesFormulaBasedAxisTitleTypeface() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { 1d }) }));
        chart.SetXAxisTitle("Units");
        var title = chart.ChartPart!.ChartSpace!.Descendants<C.CategoryAxis>().Single().GetFirstChild<C.Title>()!;
        title.RemoveAllChildren<C.ChartText>();
        title.AddChild(new C.ChartText(new C.StringReference(new C.Formula("Sheet1!$A$1"),
            new C.StringCache(new C.PointCount { Val = 1 }, new C.StringPoint(new C.NumericValue("Units")) { Index = 0 }))), true);
        title.AddChild(new C.TextProperties(new DocumentFormat.OpenXml.Drawing.BodyProperties(),
            new DocumentFormat.OpenXml.Drawing.ListStyle(), new DocumentFormat.OpenXml.Drawing.Paragraph(
                new DocumentFormat.OpenXml.Drawing.ParagraphProperties(new DocumentFormat.OpenXml.Drawing.DefaultRunProperties(
                    new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = "Georgia" })))), true);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal("Units", snapshot.Layout.CategoryAxisTitle);
        Assert.Equal("Georgia", snapshot.Layout.AxisTitleFontFamily);
    }
    [Fact]
    public void OfficeSnapshot_RejectsUnrepresentedSecondaryAxisLimits() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A" }, new[] {
            new OfficeChartSeries("Volume", new[] { 100d }),
            new OfficeChartSeries("Ratio", new[] { 1d }, null, null, null, true, renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary) }));
        Assert.True(chart.TryGetOfficeSnapshot(out _));
        var secondary = chart.ChartPart!.ChartSpace!.Descendants<C.ValueAxis>().Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Right);
        secondary.Scaling!.AddChild(new C.MaxAxisValue { Val = 2 }, true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }
    [Fact]
    public void OfficeSnapshot_PreservesInheritedTextFontsAndRejectsConflictingBodyFonts() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A" },
            new[] { new OfficeChartSeries("Values", new[] { 1d }) }));
        var space = chart.ChartPart!.ChartSpace!;
        var native = space.GetFirstChild<C.Chart>()!;
        space.AddChild(new C.TextProperties(new DocumentFormat.OpenXml.Drawing.BodyProperties(), new DocumentFormat.OpenXml.Drawing.ListStyle(),
            new DocumentFormat.OpenXml.Drawing.Paragraph(new DocumentFormat.OpenXml.Drawing.ParagraphProperties(
                new DocumentFormat.OpenXml.Drawing.DefaultRunProperties(new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = "Arial" })))), true);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal("Arial", snapshot.Style.FontFamily);
        native.GetFirstChild<C.Legend>()!.AddChild(new C.TextProperties(new DocumentFormat.OpenXml.Drawing.BodyProperties(), new DocumentFormat.OpenXml.Drawing.ListStyle(),
            new DocumentFormat.OpenXml.Drawing.Paragraph(new DocumentFormat.OpenXml.Drawing.ParagraphProperties(
                new DocumentFormat.OpenXml.Drawing.DefaultRunProperties(new DocumentFormat.OpenXml.Drawing.LatinFont { Typeface = "Georgia" })))), true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void OfficeSnapshot_RejectsCompoundSurfaceOutlines() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Pie, new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { 1d }) }));
        chart.ChartPart!.ChartSpace!.AddChild(new C.ShapeProperties(new DocumentFormat.OpenXml.Drawing.Outline(
            new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "112233" })) {
                CompoundLineType = DocumentFormat.OpenXml.Drawing.CompoundLineValues.Double }), true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void OfficeSnapshot_PreservesBasicDataLabelsAndRejectsPerPointText() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Pie, new OfficeChartData(new[] { "A", "B" },
            new[] { new OfficeChartSeries("Share", new[] { 3d, 1d }) }));
        var labels = chart.ChartPart!.ChartSpace!.Descendants<C.DataLabels>().Single();
        labels.GetFirstChild<C.ShowPercent>()!.Val = true;
        labels.GetFirstChild<C.ShowCategoryName>()!.Val = true;
        labels.AddChild(new C.Separator(" / "), true);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.True(snapshot.Layout!.ShowDataLabels);
        Assert.True(snapshot.Layout.ShowDataLabelPercentages);
        Assert.True(snapshot.Layout.ShowDataLabelCategoryNames);
        Assert.Equal(" / ", snapshot.Layout.DataLabelSeparator);
        var drawing = OfficeChartDrawingRenderer.Render(snapshot);
        Assert.Contains(drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text.Contains("A / 75%"));
        labels.AddChild(new C.DataLabel(new C.Index { Val = 0 }), true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }
    [Fact]
    public void OfficeSnapshot_PreservesCombinationPlottingOrder() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A" }, new[] {
            new OfficeChartSeries("Volume", new[] { 100d }),
            new OfficeChartSeries("Ratio", new[] { 1d }, null, null, null, true, renderKind: OfficeChartKind.Line) }));
        var orders = chart.ChartPart!.ChartSpace!.Descendants<C.Order>().ToArray();
        orders[0].Val = 1; orders[1].Val = 0;
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(new[] { "Ratio", "Volume" }, snapshot.Data.Series.Select(item => item.Name));
    }

    [Fact]
    public void OfficeSnapshot_DoesNotFillStandardRadarSeries() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Radar, new OfficeChartData(new[] { "A", "B" },
            new[] { new OfficeChartSeries("Values", new[] { 1d, 2d }) }));
        chart.ChartPart!.ChartSpace!.Descendants<C.RadarStyle>().Single().Val = C.RadarStyleValues.Standard;
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.False(snapshot.Layout!.FillRadarSeries);
    }
    [Fact]
    public void OfficeSnapshot_PreservesSurfaceAndGridAppearance() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Values", new[] { 3d, 4d }) }));
        var space = chart.ChartPart!.ChartSpace!;
        space.AddChild(new C.ShapeProperties(new DocumentFormat.OpenXml.Drawing.SolidFill(
            new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "112233" }), new DocumentFormat.OpenXml.Drawing.Outline(new DocumentFormat.OpenXml.Drawing.NoFill())), true);
        var axis = space.GetFirstChild<C.Chart>()!.PlotArea!.GetFirstChild<C.ValueAxis>()!;
        axis.GetFirstChild<C.MajorGridlines>()!.AddChild(new C.ChartShapeProperties(new DocumentFormat.OpenXml.Drawing.Outline(
            new DocumentFormat.OpenXml.Drawing.SolidFill(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "445566" }),
            new DocumentFormat.OpenXml.Drawing.PresetDash { Val = DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dash }) { Width = 25400 }), true);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(OfficeColor.Parse("#112233"), snapshot.Style.BackgroundColor);
        Assert.False(snapshot.Style.ShowBorder);
        Assert.Equal(OfficeColor.Parse("#445566"), snapshot.Style.ValueGridLineColor);
        Assert.Equal(2, snapshot.Style.ValueGridLineWidth);
        Assert.Equal(OfficeStrokeDashStyle.Dash, snapshot.Style.ValueGridLineDashStyle);
        Assert.True(snapshot.Style.ShowValueGridLines);
        var drawing = OfficeChartDrawingRenderer.Render(snapshot);
        Assert.Contains(drawing.Shapes, shape => shape.Shape.FillColor == OfficeColor.Parse("#112233"));
        Assert.Contains(drawing.Shapes, shape => shape.Shape.StrokeColor == OfficeColor.Parse("#445566") && shape.Shape.StrokeDashStyle == OfficeStrokeDashStyle.Dash);
    }

    [Fact]
    public void OfficeSnapshot_RejectsAnUnqualifiedSurfaceInsteadOfDroppingItsAppearance() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Pie, new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { 3d }) }));
        chart.ChartPart!.ChartSpace!.AddChild(new C.ShapeProperties(new DocumentFormat.OpenXml.Drawing.GradientFill()), true);
        string before = chart.ChartPart.ChartSpace.OuterXml;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        Assert.Equal(before, chart.ChartPart.ChartSpace.OuterXml);
        Assert.Contains(document.CreateVisualSnapshot().Diagnostics, diagnostic => diagnostic.Code == WordImageExportDiagnosticCodes.UnsupportedChart);
    }
    [Fact]
    public void OfficeSnapshot_PreservesReferencedAxisFormatsAndLimitsAfterResize() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Line, new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Values", new[] { 3d, 4d }) }));
        var plot = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        var axis = plot.GetFirstChild<C.ValueAxis>()!;
        axis.Scaling!.AddChild(new C.MinAxisValue { Val = 0 }, true);
        axis.Scaling.AddChild(new C.MaxAxisValue { Val = 10 }, true);
        axis.NumberingFormat!.FormatCode = "0.00";
        axis.AddChild(new C.MajorUnit { Val = 2 }, true);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        var resized = snapshot.WithSize(200, 150);
        Assert.Same(snapshot.Data, resized.Data);
        Assert.Same(snapshot.Style, resized.Style);
        Assert.Same(snapshot.Layout, resized.Layout);
        Assert.Same(snapshot.RadialLayout, resized.RadialLayout);
        Assert.Equal("0.00", resized.Layout.VerticalAxisNumberFormat);
        Assert.Equal(0, resized.Layout.VerticalAxisMinimum);
        Assert.Equal(10, resized.Layout.VerticalAxisMaximum);
        Assert.Equal(2, resized.Layout.VerticalAxisMajorUnit);
        Assert.Equal(200, resized.WidthPoints);
        Assert.Equal(150, resized.HeightPoints);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImageExport_UsesSharedBubbleAndCombinationData(bool bubble) {
        using var document = WordDocument.Create();
        var data = bubble ? new OfficeChartData(new[] { "1", "2" }, new[] {
            OfficeChartSeries.CreateBubble("Measured", new[] { 1d, 2d }, new[] { 3d, 4d }, new[] { 5d, 20d }, OfficeColor.Parse("#224466")) }) :
            new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Columns", new[] { 100d, 200d }),
                new OfficeChartSeries("Ratio", new[] { 1d, 2d }, null, OfficeColor.Parse("#224466"), null, true, renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary) });
        document.AddChart(bubble ? OfficeChartKind.Bubble : OfficeChartKind.ColumnClustered, data, "Shared projection");
        var page = document.CreateVisualSnapshot();
        Assert.DoesNotContain(page.Diagnostics, diagnostic => diagnostic.Code == WordImageExportDiagnosticCodes.UnsupportedChart);
        Assert.Contains(page.Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "Shared projection");
        Assert.Contains(page.Drawing.Shapes, shape => shape.Shape.FillColor == OfficeColor.Parse("#224466") || shape.Shape.StrokeColor == OfficeColor.Parse("#224466"));
    }
    public static IEnumerable<object[]> SupportedKinds() => Enum.GetValues(typeof(OfficeChartKind)).Cast<OfficeChartKind>().Select(kind => new object[] { kind });

    [Theory]
    [MemberData(nameof(SupportedKinds))]
    public void OfficeSnapshot_ReadsAllSharedFamiliesAcrossNativeSaveReopen(OfficeChartKind kind) {
        using var document = WordDocument.Create();
        var expected = kind == OfficeChartKind.Bubble
            ? OfficeChartSeries.CreateBubble("Measurements", new[] { 1d, 2d }, new[] { 3d, 4d }, new[] { 5d, 6d }, OfficeColor.Parse("#224466"),
                markerOutlineColor: OfficeColor.Parse("#FF0000"), markerOutlineWidth: 2)
            : new OfficeChartSeries("Measurements", new[] { 3d, 4d }, kind == OfficeChartKind.Scatter ? new[] { 1d, 2d } : null, OfficeColor.Parse("#224466"));
        document.AddChart(kind, new OfficeChartData(new[] { "1", "2" }, new[] { expected }), "Measurements");
        using var bytes = new MemoryStream(); document.Save(bytes); bytes.Position = 0;
        using var reopened = WordDocument.Load(bytes);
        Assert.True(reopened.Charts.Single().TryGetOfficeSnapshot(out var snapshot), kind.ToString());
        Assert.Equal(kind, snapshot.ChartKind);
        var actual = snapshot.Data.Series.Single();
        Assert.Equal(expected.Name, actual.Name);
        Assert.Equal(expected.Values, actual.Values);
        Assert.Equal(expected.XValues, actual.XValues);
        Assert.Equal(expected.BubbleSizes, actual.BubbleSizes);
        Assert.Equal(expected.Color, actual.Color);
        if (kind == OfficeChartKind.Bubble) {
            Assert.Equal(expected.MarkerOutlineColor, actual.MarkerOutlineColor);
            Assert.Equal(expected.MarkerOutlineWidth, actual.MarkerOutlineWidth);
        }
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void OfficeSnapshot_PreservesCombinationFamiliesAndSecondaryAxisAssignment() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Volume", new[] { 100d, 200d }),
            new OfficeChartSeries("Ratio", new[] { 1d, 2d }, null, null, null, true, renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary) }));
        using var bytes = new MemoryStream(); document.Save(bytes); bytes.Position = 0;
        using var reopened = WordDocument.Load(bytes);
        Assert.True(reopened.Charts.Single().TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(new[] { OfficeChartKind.ColumnClustered, OfficeChartKind.Line }, snapshot.Data.Series.Select(series => series.RenderKind!.Value));
        Assert.Equal(new[] { OfficeChartAxisGroup.Primary, OfficeChartAxisGroup.Secondary }, snapshot.Data.Series.Select(series => series.AxisGroup));
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void OfficeSnapshot_UsesCategoryLegendIndexesForRadialCharts() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Doughnut, new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Status", new[] { 1d, 2d }) }));
        chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.GetFirstChild<C.Legend>()!.AddChild(
            new C.LegendEntry(new C.Index { Val = 0 }, new C.Delete { Val = true }), true);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.True(snapshot.Data.Series.Single().ShowInLegend);
        Assert.Equal(new[] { 0 }, snapshot.Layout!.HiddenCategoryLegendIndexes);
        var drawing = OfficeChartDrawingRenderer.Render(snapshot);
        Assert.DoesNotContain(drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "A");
        Assert.Contains(drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "B");
    }
}
