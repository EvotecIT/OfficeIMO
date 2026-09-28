using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartProjectionQualificationTests {
    [Fact]
    public void Snapshot_RejectsUnrepresentedChartTextColorTransform() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { 3d }) }));
        var space = chart.ChartPart!.ChartSpace!;
        space.AddChild(new C.TextProperties(new A.BodyProperties(), new A.ListStyle(),
            new A.Paragraph(new A.ParagraphProperties(new A.DefaultRunProperties(
                new A.SolidFill(new A.RgbColorModelHex(new A.HueOffset { Val = 60000 }) { Val = "112233" }))))), true);
        string before = space.OuterXml;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        Assert.Equal(before, space.OuterXml);
    }

    [Fact]
    public void Snapshot_RejectsDifferentChartScriptTypeface() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { 3d }) }));
        var space = chart.ChartPart!.ChartSpace!;
        space.AddChild(new C.TextProperties(new A.BodyProperties(), new A.ListStyle(),
            new A.Paragraph(new A.ParagraphProperties(new A.DefaultRunProperties(
                new A.LatinFont { Typeface = "Arial" }, new A.EastAsianFont { Typeface = "Yu Gothic" })))), true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void Snapshot_RejectsEmptyNativeLabelSeparator() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Pie,
            new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { 3d }) }));
        chart.ChartPart!.ChartSpace!.Descendants<C.DataLabels>().Single().AddChild(new C.Separator(string.Empty), true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void Snapshot_RejectsUnresolvedFormulaBasedSeriesName() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { 3d }) }));
        var text = chart.ChartPart!.ChartSpace!.Descendants<C.BarChartSeries>().Single().GetFirstChild<C.SeriesText>()!;
        text.GetFirstChild<C.StringReference>()!.StringCache!.Remove();
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void Snapshot_TreatsOmittedBarGroupingAsClusteredForOverlap() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Values", new[] { 3d }) }));
        var bars = chart.ChartPart!.ChartSpace!.Descendants<C.BarChart>().Single();
        bars.GetFirstChild<C.BarGrouping>()!.Remove();
        bars.AddChild(new C.Overlap { Val = 0 }, true);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(OfficeChartKind.ColumnClustered, snapshot.ChartKind);
    }

    [Fact]
    public void Snapshot_BoundsBubbleOverridesSeparatelyFromCachedPoints() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Bubble,
            new OfficeChartData(new[] { "1" }, new[] { OfficeChartSeries.CreateBubble("Values", new[] { 1d }, new[] { 3d }, new[] { 5d }) }));
        var series = chart.ChartPart!.ChartSpace!.Descendants<C.BubbleChartSeries>().Single();
        for (int index = 0; index < 10001; index++)
            series.InsertBefore(new C.DataPoint(new C.Index { Val = 99 }), series.GetFirstChild<C.XValues>());
        Assert.True(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void AuthoredBarDirectionChangeKeepsNativeAxesProjectable() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.BarClustered,
            new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Values", new[] { 3D, 4D }) }));
        chart.BarDirection = WordChartBarDirection.Column;
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(OfficeChartKind.ColumnClustered, snapshot.ChartKind);
        var plot = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        Assert.Equal(C.AxisPositionValues.Bottom, plot.GetFirstChild<C.CategoryAxis>()!.AxisPosition!.Val!.Value);
        Assert.Equal(C.AxisPositionValues.Left, plot.GetFirstChild<C.ValueAxis>()!.AxisPosition!.Val!.Value);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Snapshot_RejectsHiddenWorkbookValuesOnlyWhenVisibleOnlyIsEnabled(bool visibleOnly) {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Values", new[] { 3d, 4d }) }));
        var part = chart.ChartPart!;
        part.ChartSpace!.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotVisibleOnly>()!.Val = visibleOnly;
        var embedded = part.GetPartsOfType<DocumentFormat.OpenXml.Packaging.EmbeddedPackagePart>().Single();
        using var bytes = new MemoryStream();
        using (var source = embedded.GetStream()) source.CopyTo(bytes);
        bytes.Position = 0;
        using (var workbook = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(bytes, true)) {
            workbook.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>().Skip(1).First().Hidden = true;
        }
        bytes.Position = 0; embedded.FeedData(bytes);
        Assert.Equal(!visibleOnly, chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData("trendline")]
    [InlineData("errorBars")]
    [InlineData("dropLines")]
    [InlineData("sparse")]
    [InlineData("negative")]
    [InlineData("dateAxis")]
    [InlineData("style")]
    [InlineData("labelFormat")]
    [InlineData("legendEntry")]
    [InlineData("rotation")]
    [InlineData("labelSkip")]
    [InlineData("unequal")]
    [InlineData("emptyValues")]
    [InlineData("unequalSeries")]
    [InlineData("differentCategories")]
    [InlineData("titleParagraphs")]
    [InlineData("titleBreak")]
    [InlineData("dataTable")]
    [InlineData("hierarchy")]
    [InlineData("secondaryCategory")]
    [InlineData("labelOffset")]
    [InlineData("rounded")]
    [InlineData("crossBetween")]
    [InlineData("automaticMarker")]
    [InlineData("inheritedMarker")]
    [InlineData("labelAlignment")]
    [InlineData("barShape")]
    [InlineData("barSeriesShape")]
    [InlineData("seriesLines")]
    [InlineData("userShapes")]
    [InlineData("radialLeaderLines")]
    public void Snapshot_RejectsUnrepresentedNativeChartContent(string feature) {
        using var document = WordDocument.Create();
        var kind = feature == "radialLeaderLines" ? OfficeChartKind.Pie :
            feature == "negative" || feature.StartsWith("bar", StringComparison.Ordinal) || feature == "seriesLines" ? OfficeChartKind.ColumnClustered : OfficeChartKind.Line;
        var chart = document.AddChart(kind, new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Values", new[] { 3d, -4d }) }), title: "Revenue");
        if (feature == "secondaryCategory") chart.SetData(kind, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Values", new[] { 3d, -4d }),
            new OfficeChartSeries("Secondary", new[] { 1d, 2d }, null, null, null, true, renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary) }));
        var space = chart.ChartPart!.ChartSpace!;
        var native = space.GetFirstChild<C.Chart>()!;
        var plot = native.PlotArea!;
        var layer = plot.ChildElements.OfType<OpenXmlCompositeElement>().First(item => item.LocalName.EndsWith("Chart", StringComparison.Ordinal));
        var series = layer.ChildElements.OfType<OpenXmlCompositeElement>().First(item => item.LocalName == "ser");
        if (feature == "trendline") series.AddChild(new C.Trendline(new C.TrendlineType { Val = C.TrendlineValues.Linear }), true);
        else if (feature == "errorBars") series.AddChild(new C.ErrorBars(), true);
        else if (feature == "dropLines") layer.AddChild(new C.DropLines(), true);
        else if (feature == "sparse") series.Descendants<C.NumericPoint>().Last().Remove();
        else if (feature == "negative") series.AddChild(new C.InvertIfNegative { Val = true }, true);
        else if (feature == "style") space.AddChild(new C.Style { Val = 42 }, true);
        else if (feature == "labelFormat") layer.AddChild(new C.DataLabels(new C.NumberingFormat { FormatCode = "m/d/yyyy", SourceLinked = false }, new C.ShowValue { Val = true }), true);
        else if (feature == "legendEntry") native.GetFirstChild<C.Legend>()!.AddChild(new C.LegendEntry(new C.Index { Val = 0 }, new C.TextProperties(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.ParagraphProperties(new A.DefaultRunProperties { FontSize = 2400 })))), true);
        else if (feature == "rotation") space.AddChild(new C.TextProperties(new A.BodyProperties { Rotation = 5400000 }, new A.ListStyle(), new A.Paragraph()), true);
        else if (feature == "labelSkip") plot.GetFirstChild<C.CategoryAxis>()!.AddChild(new C.TickLabelSkip { Val = 2 }, true);
        else if (feature == "unequal") { var cache = series.Descendants<C.Values>().Single(); cache.Descendants<C.NumericPoint>().Last().Remove(); cache.Descendants<C.PointCount>().Single().Val = 1; }
        else if (feature == "emptyValues") series.GetFirstChild<C.Values>()!.Remove();
        else if (feature == "unequalSeries" || feature == "differentCategories") {
            var sibling = (OpenXmlCompositeElement)series.CloneNode(true);
            sibling.GetFirstChild<C.Index>()!.Val = 1;
            sibling.GetFirstChild<C.Order>()!.Val = 1;
            if (feature == "unequalSeries") {
                sibling.Descendants<C.NumericPoint>().Last().Remove();
                sibling.Descendants<C.Values>().Single().Descendants<C.PointCount>().Single().Val = 1;
            } else sibling.GetFirstChild<C.CategoryAxisData>()!.Descendants<C.StringPoint>().Last().NumericValue!.Text = "Different category";
            layer.InsertAfter(sibling, series);
        }
        else if (feature == "titleParagraphs") native.GetFirstChild<C.Title>()!.Descendants<C.RichText>().Single().Append(new A.Paragraph(new A.Run(new A.Text("2026"))));
        else if (feature == "titleBreak") native.GetFirstChild<C.Title>()!.Descendants<A.Paragraph>().Single().Append(new A.Break(), new A.Run(new A.Text("2026")));
        else if (feature == "dataTable") plot.AddChild(new C.DataTable(new C.ShowHorizontalBorder { Val = true }), true);
        else if (feature == "hierarchy") series.GetFirstChild<C.CategoryAxisData>()!.Append(new C.MultiLevelStringReference());
        else if (feature == "secondaryCategory") plot.Elements<C.CategoryAxis>().Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Top).GetFirstChild<C.Delete>()!.Val = false;
        else if (feature == "labelOffset") plot.GetFirstChild<C.CategoryAxis>()!.GetFirstChild<C.LabelOffset>()!.Val = 200;
        else if (feature == "rounded") space.GetFirstChild<C.RoundedCorners>()!.Val = true;
        else if (feature == "crossBetween") plot.GetFirstChild<C.ValueAxis>()!.GetFirstChild<C.CrossBetween>()!.Val = C.CrossBetweenValues.MidpointCategory;
        else if (feature == "automaticMarker") series.AddChild(new C.Marker(new C.Symbol { Val = C.MarkerStyleValues.Auto }), true);
        else if (feature == "inheritedMarker") series.RemoveAllChildren<C.Marker>();
        else if (feature == "labelAlignment") plot.GetFirstChild<C.CategoryAxis>()!.AddChild(new C.LabelAlignment { Val = C.LabelAlignmentValues.Left }, true);
        else if (feature == "barShape" || feature == "barSeriesShape") {
            var shape = new OpenXmlUnknownElement("c", "shape", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            shape.SetAttribute(new OpenXmlAttribute("val", "", "cylinder"));
            (feature == "barShape" ? layer : series).Append(shape);
        }
        else if (feature == "seriesLines") layer.Append(new C.SeriesLines());
        else if (feature == "userShapes") {
            var drawingPart = chart.ChartPart.AddNewPart<DocumentFormat.OpenXml.Packaging.ChartDrawingPart>();
            var shapes = new OpenXmlUnknownElement("c", "userShapes", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            shapes.SetAttribute(new OpenXmlAttribute("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", chart.ChartPart.GetIdOfPart(drawingPart)));
            space.Append(shapes);
        }
        else if (feature == "radialLeaderLines") layer.AddChild(new C.DataLabels(new C.DataLabelPosition { Val = C.DataLabelPositionValues.BestFit }, new C.ShowValue { Val = true }, new C.ShowLeaderLines { Val = true }), true);
        else if (feature == "dateAxis") {
            var category = plot.GetFirstChild<C.CategoryAxis>()!;
            var replacement = new C.DateAxis();
            foreach (var child in category.ChildElements) replacement.Append(child.CloneNode(true));
            plot.ReplaceChild(replacement, category);
        }
        string before = space.OuterXml;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        Assert.Equal(before, space.OuterXml);
    }
}
