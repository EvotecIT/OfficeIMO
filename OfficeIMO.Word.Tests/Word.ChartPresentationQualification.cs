using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartPresentationQualificationTests {
    [Theory]
    [InlineData(OfficeChartKind.Line)]
    [InlineData(OfficeChartKind.Area)]
    [InlineData(OfficeChartKind.Scatter)]
    public void Snapshot_RejectsUnclippedPointsOutsideExplicitValueBounds(OfficeChartKind kind) {
        using var document = WordDocument.Create();
        var data = new OfficeChartData(new[] { "A", "B", "C" }, new[] {
            new OfficeChartSeries("Values", new[] { 0d, 10d, 20d }, kind == OfficeChartKind.Scatter ? new[] { 1d, 2d, 3d } : null) });
        var chart = document.AddChart(kind, data);
        var plot = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        plot.Elements<C.ValueAxis>().Last().GetFirstChild<C.Scaling>()!.AddChild(new C.MaxAxisValue { Val = 15d }, true);
        string native = chart.ChartPart.ChartSpace.OuterXml;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        Assert.Equal(native, chart.ChartPart.ChartSpace.OuterXml);
    }

    [Fact]
    public void Snapshot_RejectsUnclippedScatterPointsOutsideExplicitHorizontalBounds() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Scatter, new OfficeChartData(new[] { "A", "B", "C" }, new[] {
            new OfficeChartSeries("Values", new[] { 1d, 2d, 3d }, new[] { 1d, 2d, 3d }) }));
        var plot = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        plot.Elements<C.ValueAxis>().First().GetFirstChild<C.Scaling>()!.AddChild(new C.MaxAxisValue { Val = 2d }, true);
        string native = chart.ChartPart.ChartSpace.OuterXml;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        Assert.Equal(native, chart.ChartPart.ChartSpace.OuterXml);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Snapshot_RejectsAreaBaselineOutsideExplicitVerticalBounds(bool positive) {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Area, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Values", positive ? new[] { 10d, 20d } : new[] { -10d, -20d }) }));
        var axis = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!.Elements<C.ValueAxis>().Single();
        axis.GetFirstChild<C.Scaling>()!.AddChild(positive ? new C.MinAxisValue { Val = 5d } : new C.MaxAxisValue { Val = -5d }, true);
        string native = chart.ChartPart.ChartSpace.OuterXml;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        Assert.Equal(native, chart.ChartPart.ChartSpace.OuterXml);
    }

    [Theory]
    [InlineData(OfficeChartKind.ColumnClustered)]
    [InlineData(OfficeChartKind.BarClustered)]
    public void Snapshot_DeletedAxesSuppressTheirLines(OfficeChartKind kind) {
        using var document = WordDocument.Create();
        var chart = Create(document, kind);
        foreach (var axis in chart.ChartPart!.ChartSpace!.Descendants().Where(e => e is C.CategoryAxis or C.ValueAxis))
            axis.GetFirstChild<C.Delete>()!.Val = true;
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.False(snapshot.Layout.ShowCategoryAxisLine);
        Assert.False(snapshot.Layout.ShowValueAxisLine);
    }

    [Theory]
    [InlineData(OfficeChartKind.ColumnClustered, "high", OfficeChartAxisTickLabelPosition.High)]
    [InlineData(OfficeChartKind.ColumnClustered, "low", OfficeChartAxisTickLabelPosition.Low)]
    [InlineData(OfficeChartKind.BarClustered, "high", OfficeChartAxisTickLabelPosition.High)]
    public void Snapshot_MapsPhysicalTickLabelPlacement(OfficeChartKind kind, string placement, OfficeChartAxisTickLabelPosition expected) {
        using var document = WordDocument.Create();
        var chart = Create(document, kind);
        var axis = chart.ChartPart!.ChartSpace!.Descendants<C.ValueAxis>().Single();
        axis.GetFirstChild<C.TickLabelPosition>()!.Val = placement == "high" ? C.TickLabelPositionValues.High : C.TickLabelPositionValues.Low;
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(expected, kind == OfficeChartKind.BarClustered ? snapshot.Layout.HorizontalAxisTickLabelPosition : snapshot.Layout.VerticalAxisTickLabelPosition);
    }

    [Theory]
    [InlineData("m/d/yyyy")]
    [InlineData("# ?/?")]
    [InlineData("0.00E+00")]
    [InlineData("[>=100]0;0")]
    public void Snapshot_RejectsUnrepresentedAxisFormats(string format) {
        using var document = WordDocument.Create();
        var chart = Create(document, OfficeChartKind.Line);
        chart.ChartPart!.ChartSpace!.Descendants<C.ValueAxis>().Single().GetFirstChild<C.NumberingFormat>()!.FormatCode = format;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData("titleShape")]
    [InlineData("legendShape")]
    [InlineData("axisTitleLayout")]
    [InlineData("gap")]
    [InlineData("overlap")]
    [InlineData("vary")]
    public void Snapshot_RejectsUnrepresentedPresentation(string feature) {
        using var document = WordDocument.Create();
        var chart = Create(document, OfficeChartKind.ColumnClustered);
        var native = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!;
        var bars = native.PlotArea!.GetFirstChild<C.BarChart>()!;
        if (feature == "gap") bars.GetFirstChild<C.GapWidth>()!.Val = 20;
        else if (feature == "overlap") bars.GetFirstChild<C.Overlap>()!.Val = 50;
        else if (feature == "vary") bars.GetFirstChild<C.VaryColors>()!.Val = true;
        else if (feature == "axisTitleLayout") {
            var title = new C.Title(new C.ChartText(new C.RichText(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text("Axis"))))),
                new C.Layout(new C.ManualLayout(new C.Left { Val = .3 })));
            native.PlotArea.GetFirstChild<C.ValueAxis>()!.AddChild(title, true);
        } else {
            OpenXmlCompositeElement owner = feature == "titleShape" ? native.GetFirstChild<C.Title>()! : native.GetFirstChild<C.Legend>()!;
            owner.AddChild(new C.ChartShapeProperties(new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" })), true);
        }
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void Snapshot_RejectsExplodedSlices(OfficeChartKind kind) {
        using var document = WordDocument.Create();
        var chart = Create(document, kind);
        chart.ChartPart!.ChartSpace!.Descendants<C.PieChartSeries>().Single().AddChild(
            new C.DataPoint(new C.Index { Val = 0 }, new C.Explosion { Val = 25 }), true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void Snapshot_RejectsUnrepresentedUniformRadialPalette(OfficeChartKind kind) {
        using var document = WordDocument.Create();
        var chart = Create(document, kind);
        chart.ChartPart!.ChartSpace!.Descendants<C.VaryColors>().Single().Val = false;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void Snapshot_RejectsBubblePointMarkersWithoutChangingNativeXml() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Bubble, new OfficeChartData(new[] { "A", "B" }, new[] {
            OfficeChartSeries.CreateBubble("Values", new[] { 1d, 2d }, new[] { 3d, 4d }, new[] { 1d, 2d })
        }));
        chart.ChartPart!.ChartSpace!.Descendants<C.BubbleChartSeries>().Single().AddChild(
            new C.DataPoint(new C.Index { Val = 0 }, new C.Marker(new C.Symbol { Val = C.MarkerStyleValues.Square })), true);
        string xml = chart.ChartPart.ChartSpace.OuterXml;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        Assert.Equal(xml, chart.ChartPart.ChartSpace.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Snapshot_RejectsAmbiguousAxisSides(bool secondaryLeft) {
        using var document = WordDocument.Create();
        var secondary = new OfficeChartSeries("Second", new[] { 30d, 40d }, null, null, null, true,
            renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary);
        var chart = document.AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("First", new[] { 3d, 4d }), secondary
        }));
        var axes = chart.ChartPart!.ChartSpace!.Descendants<C.ValueAxis>().ToArray();
        axes[secondaryLeft ? 1 : 0].AxisPosition!.Val = secondaryLeft ? C.AxisPositionValues.Left : C.AxisPositionValues.Right;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void Snapshot_MapsUniformTitleTextAndRejectsMixedRuns() {
        using var document = WordDocument.Create();
        var chart = Create(document, OfficeChartKind.ColumnClustered);
        var paragraph = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.Title!.Descendants<A.Paragraph>().Single();
        var run = paragraph.GetFirstChild<A.Run>()!;
        run.PrependChild(new A.RunProperties(new A.SolidFill(new A.RgbColorModelHex { Val = "224466" })) {
            FontSize = 2400, Bold = true, Italic = true, Underline = A.TextUnderlineValues.Single
        });
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(24d, snapshot.Style.TitleFontSize);
        Assert.Equal(OfficeFontStyle.Bold | OfficeFontStyle.Italic | OfficeFontStyle.Underline, snapshot.Style.TitleFontStyle);
        Assert.Equal(OfficeColor.FromRgb(34, 68, 102), snapshot.Style.TitleColor);
        paragraph.Append(new A.Run(new A.RunProperties { FontSize = 1200 }, new A.Text("Different")));
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void Snapshot_MapsUniformLegendAndAxisText() {
        using var document = WordDocument.Create();
        var chart = Create(document, OfficeChartKind.Line);
        var native = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!;
        C.TextProperties Text(int size, string color) => new(new A.BodyProperties(), new A.ListStyle(),
            new A.Paragraph(new A.ParagraphProperties(new A.DefaultRunProperties(new A.SolidFill(new A.RgbColorModelHex { Val = color })) { FontSize = size, Bold = true })));
        native.Legend!.AddChild(Text(1400, "224466"), true);
        foreach (var axis in native.PlotArea!.ChildElements.OfType<OpenXmlCompositeElement>().Where(axis => axis is C.CategoryAxis or C.ValueAxis)) axis.AddChild(Text(1000, "667788"), true);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(14d, snapshot.Layout.LegendFontSize);
        Assert.Equal(10d, snapshot.Layout.AxisLabelFontSize);
        Assert.Equal(OfficeFontStyle.Bold, snapshot.Layout.LegendFontStyle);
        Assert.Equal(OfficeColor.FromRgb(34, 68, 102), snapshot.Style.LegendTextColor);
        Assert.Equal(OfficeColor.FromRgb(102, 119, 136), snapshot.Style.MutedTextColor);
    }

    [Theory]
    [InlineData("legend")]
    [InlineData("axis")]
    public void Snapshot_MapsListLevelTextDefaults(string ownerName) {
        using var document = WordDocument.Create();
        var chart = Create(document, OfficeChartKind.Line);
        var native = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!;
        OpenXmlCompositeElement owner = ownerName == "legend" ? native.Legend! : native.PlotArea!.GetFirstChild<C.ValueAxis>()!;
        // The shared axis role requires uniform category/value label formatting.
        var owners = ownerName == "legend" ? new[] { owner } : native.PlotArea!.ChildElements.OfType<OpenXmlCompositeElement>().Where(axis => axis is C.CategoryAxis or C.ValueAxis).ToArray();
        foreach (var area in owners) area.AddChild(new C.TextProperties(new A.BodyProperties(),
            new A.ListStyle(new A.Level1ParagraphProperties(new A.DefaultRunProperties(new A.SolidFill(new A.RgbColorModelHex { Val = "224466" })) { FontSize = 1600, Bold = true })),
            new A.Paragraph()), true);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(16d, ownerName == "legend" ? snapshot.Layout.LegendFontSize : snapshot.Layout.AxisLabelFontSize);
        Assert.Equal(OfficeFontStyle.Bold, ownerName == "legend" ? snapshot.Layout.LegendFontStyle : snapshot.Layout.AxisTextFontStyle);
        Assert.Equal(OfficeColor.FromRgb(34, 68, 102), ownerName == "legend" ? snapshot.Style.LegendTextColor : snapshot.Style.MutedTextColor);
    }

    [Fact]
    public void Snapshot_RejectsAxisTitleShapeAppearance() {
        using var document = WordDocument.Create();
        var chart = Create(document, OfficeChartKind.Line);
        var axis = chart.ChartPart!.ChartSpace!.Descendants<C.ValueAxis>().Single();
        axis.AddChild(new C.Title(new C.ChartText(new C.RichText(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text("Axis"))))),
            new C.ChartShapeProperties(new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" }))), true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData("m/d/yyyy")]
    [InlineData("# ?/?")]
    [InlineData("0.00E+00")]
    [InlineData("[>=100]0;0")]
    public void Snapshot_RejectsUnrepresentedNumericCategoryFormats(string format) {
        using var document = WordDocument.Create();
        var chart = Create(document, OfficeChartKind.Line);
        var native = chart.ChartPart!.ChartSpace!;
        var categories = native.Descendants<C.CategoryAxisData>().Single();
        categories.RemoveAllChildren();
        categories.Append(new C.NumberLiteral(new C.FormatCode("General"), new C.PointCount { Val = 2 },
            new C.NumericPoint(new C.NumericValue("1")) { Index = 0 }, new C.NumericPoint(new C.NumericValue("2")) { Index = 1 }));
        native.Descendants<C.CategoryAxis>().Single().GetFirstChild<C.NumberingFormat>()!.FormatCode = format;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Snapshot_RejectsUnrepresentedMaximumCrossingForBars(bool valueAxis) {
        using var document = WordDocument.Create();
        var chart = Create(document, OfficeChartKind.BarClustered);
        var plot = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        OpenXmlCompositeElement axis = valueAxis ? plot.GetFirstChild<C.ValueAxis>()! : plot.GetFirstChild<C.CategoryAxis>()!;
        axis.GetFirstChild<C.Crosses>()!.Val = C.CrossesValues.Maximum;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Snapshot_MapsNativeCrossingToPerpendicularScreenAxis(bool valueAxis) {
        using var document = WordDocument.Create();
        var chart = Create(document, OfficeChartKind.ColumnClustered);
        var plot = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        OpenXmlCompositeElement axis = valueAxis ? plot.GetFirstChild<C.ValueAxis>()! : plot.GetFirstChild<C.CategoryAxis>()!;
        axis.GetFirstChild<C.Crosses>()!.Val = C.CrossesValues.Maximum;
        string native = chart.ChartPart.ChartSpace.OuterXml;
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(valueAxis ? OfficeChartAxisCrossingPosition.Maximum : OfficeChartAxisCrossingPosition.AutoZero,
            snapshot.Layout!.HorizontalAxisCrossingPosition);
        Assert.Equal(valueAxis ? OfficeChartAxisCrossingPosition.AutoZero : OfficeChartAxisCrossingPosition.Maximum,
            snapshot.Layout.VerticalAxisCrossingPosition);
        Assert.Equal(native, chart.ChartPart.ChartSpace.OuterXml);
    }

    private static WordChart Create(WordDocument document, OfficeChartKind kind) => document.AddChart(kind,
        new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Values", new[] { 3d, 4d }) }), title: "Chart");
}
