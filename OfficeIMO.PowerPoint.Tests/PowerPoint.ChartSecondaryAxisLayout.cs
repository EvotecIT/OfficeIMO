using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class PowerPointChartSecondaryAxisLayoutTests {
    [Theory]
    [InlineData("color")]
    [InlineData("size")]
    [InlineData("bold")]
    [InlineData("italic")]
    [InlineData("shape")]
    [InlineData("rotation")]
    [InlineData("vertical")]
    [InlineData("alignment")]
    [InlineData("eastAsianFont")]
    [InlineData("complexScriptFont")]
    public void SecondaryTitle_RejectsUnprojectedAppearance(string appearance) {
        using var presentation = PowerPointPresentation.Create();
        var slide = presentation.AddSlide();
        var chart = slide.AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A" }, new[] {
            new OfficeChartSeries("Count", new[] { 100d }),
            new OfficeChartSeries("Ratio", new[] { 1d }, null, null, null, true,
                renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        }));
        chart.SetSecondaryValueAxis(new OfficeChartValueAxisLayout().WithTitle("Ratio"));
        var title = slide.SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.ValueAxis>()
            .Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Right).GetFirstChild<C.Title>()!;
        if (appearance == "shape") title.AddChild(new C.ChartShapeProperties(
            new A.SolidFill(new A.RgbColorModelHex { Val = "FFFF00" })), true);
        else if (appearance == "rotation") title.Descendants<A.BodyProperties>().Single().Rotation = 5400000;
        else if (appearance == "vertical") title.Descendants<A.BodyProperties>().Single().Vertical = A.TextVerticalValues.Vertical;
        else if (appearance == "alignment") title.Descendants<A.Paragraph>().Single()
            .AddChild(new A.ParagraphProperties { Alignment = A.TextAlignmentTypeValues.Right }, true);
        else {
            var properties = new A.RunProperties();
            if (appearance == "color") properties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "FF0000" }));
            else if (appearance == "size") properties.FontSize = 1800;
            else if (appearance == "bold") properties.Bold = true;
            else if (appearance == "eastAsianFont") properties.Append(new A.EastAsianFont { Typeface = "MS Gothic" });
            else if (appearance == "complexScriptFont") properties.Append(new A.ComplexScriptFont { Typeface = "Arial" });
            else properties.Italic = true;
            title.Descendants<A.Run>().Single().AddChild(properties, true);
        }
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void SecondaryAxisTitleDoesNotReplacePrimaryTitleWhenAxisElementsAreReordered() {
        using var presentation = PowerPointPresentation.Create();
        var slide = presentation.AddSlide();
        var chart = slide.AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A" }, new[] {
            new OfficeChartSeries("Volume", new[] { 100d }),
            new OfficeChartSeries("Ratio", new[] { 1d }, null, null, null, true,
                renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        }));
        chart.SetValueAxisTitle("Volume");
        chart.SetSecondaryValueAxis(new OfficeChartValueAxisLayout().WithTitle("Ratio"));
        var plot = slide.SlidePart.ChartParts.Single().ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        var primary = plot.Elements<C.ValueAxis>().Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Left);
        var secondary = plot.Elements<C.ValueAxis>().Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Right);
        secondary.Remove();
        plot.InsertBefore(secondary, primary);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal("Volume", snapshot.Layout.ValueAxisTitle);
        Assert.Equal("Ratio", snapshot.Layout.SecondaryValueAxis!.Title);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnsupportedNumericAxisFormatIsNotProjected(bool secondaryAxis) {
        using var presentation = PowerPointPresentation.Create();
        var chart = presentation.AddSlide().AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A" }, new[] {
                new OfficeChartSeries("Count", new[] { 100D }),
                new OfficeChartSeries("Ratio", new[] { 10D }, null, null, null, true,
                    renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary)
            }));
        Assert.True(chart.TryGetOfficeSnapshot(out _));
        var axis = presentation.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.ValueAxis>()
            .Single(item => item.AxisPosition!.Val!.Value ==
                (secondaryAxis ? C.AxisPositionValues.Right : C.AxisPositionValues.Left));
        axis.GetFirstChild<C.NumberingFormat>()!.FormatCode = "[Red]0";
        axis.GetFirstChild<C.NumberingFormat>()!.SourceLinked = false;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(OfficeChartKind.Line)]
    [InlineData(OfficeChartKind.Area)]
    public void ExplicitPrimaryBoundsClipCachedGeometry(OfficeChartKind kind) {
        using var presentation = PowerPointPresentation.Create();
        var chart = presentation.AddSlide().AddChart(kind,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Values", new[] { 0D, 20D })
            }));
        Assert.True(chart.TryGetOfficeSnapshot(out _));
        var axis = presentation.Slides.Single().SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.ValueAxis>().Single();
        axis.Scaling!.AddChild(new C.MaxAxisValue { Val = 10D }, true);
        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(10D, snapshot.Layout.VerticalAxisMaximum);
        Assert.Contains(OfficeChartDrawingRenderer.Render(snapshot).Elements,
            element => element is OfficeDrawingGroup);
    }

    [Theory]
    [InlineData("manual")]
    [InlineData("overlay")]
    [InlineData("implicitOverlay")]
    [InlineData("paragraphs")]
    [InlineData("break")]
    public void SecondaryValueAxis_UnsupportedTitlePlacementRejectsSnapshot(string placement) {
        using var presentation = PowerPointPresentation.Create();
        var slide = presentation.AddSlide();
        var chart = slide.AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A" }, new[] {
            new OfficeChartSeries("Count", new[] { 100d }),
            new OfficeChartSeries("Ratio", new[] { 1d }, null, null, null, true,
                renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        }));
        chart.SetSecondaryValueAxis(new OfficeChartValueAxisLayout().WithTitle("Ratio"));
        var title = slide.SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.ValueAxis>()
            .Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Right).GetFirstChild<C.Title>()!;
        if (placement == "overlay") title.GetFirstChild<C.Overlay>()!.Val = true;
        else if (placement == "implicitOverlay") title.GetFirstChild<C.Overlay>()!.Val = null;
        else if (placement == "paragraphs") title.GetFirstChild<C.ChartText>()!.GetFirstChild<C.RichText>()!
            .Append(new DocumentFormat.OpenXml.Drawing.Paragraph(new DocumentFormat.OpenXml.Drawing.Run(
                new DocumentFormat.OpenXml.Drawing.Text("Percent"))));
        else if (placement == "break") title.GetFirstChild<C.ChartText>()!.GetFirstChild<C.RichText>()!
            .GetFirstChild<DocumentFormat.OpenXml.Drawing.Paragraph>()!.Append(new DocumentFormat.OpenXml.Drawing.Break());
        else title.GetFirstChild<C.Layout>()!.Append(new C.ManualLayout(new C.Left { Val = .3 }));
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData("logarithmic")]
    [InlineData("reversed")]
    [InlineData("displayUnits")]
    [InlineData("sourceLinked")]
    [InlineData("deleted")]
    public void SecondaryValueAxis_UnsupportedProjectionPreservesNativeUpdates(string setting) {
        using var presentation = PowerPointPresentation.Create();
        var slide = presentation.AddSlide();
        var data = new OfficeChartData(new[] { "A" }, new[] {
            new OfficeChartSeries("Count", new[] { 100d }),
            new OfficeChartSeries("Ratio", new[] { 10d }, null, null, null, true,
                renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        });
        var chart = slide.AddChart(OfficeChartKind.ColumnClustered, data);
        var secondary = slide.SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.ValueAxis>()
            .Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Right);
        if (setting == "deleted") secondary.AddChild(new C.Delete { Val = true }, true);
        else if (setting == "logarithmic") secondary.Scaling!.AddChild(new C.LogBase { Val = 10 }, true);
        else if (setting == "reversed") secondary.Scaling!.AddChild(new C.Orientation { Val = C.OrientationValues.MaxMin }, true);
        else if (setting == "sourceLinked") secondary.NumberingFormat!.SourceLinked = true;
        else secondary.AddChild(new C.DisplayUnits(new C.BuiltInUnit { Val = C.BuiltInUnitValues.Thousands }), true);
        string Appearance(C.ValueAxis axis) {
            var copy = (C.ValueAxis)axis.CloneNode(true);
            copy.RemoveAllChildren<C.AxisId>();
            copy.RemoveAllChildren<C.CrossingAxis>();
            return copy.OuterXml;
        }
        string native = Appearance(secondary);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        chart.UpdateData(data);
        var persistedAxis = slide.SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.ValueAxis>()
            .Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Right);
        Assert.Equal(native, Appearance(persistedAxis));
        Assert.Empty(presentation.ValidateDocument());
    }

    [Fact]
    public void SecondaryValueAxis_CacheOnlySourceLinkedFormatUsesCachedCode() {
        using var presentation = PowerPointPresentation.Create();
        var slide = presentation.AddSlide();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Count", new[] { 100d, 150d }),
            new OfficeChartSeries("Ratio", new[] { 1d, 2d }, null, null, null, true,
                renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary) });
        var chart = slide.AddChart(OfficeChartKind.ColumnClustered, data);
        var space = slide.SlidePart.ChartParts.Single().ChartSpace!;
        space.GetFirstChild<C.ExternalData>()?.Remove();
        var secondary = space.Descendants<C.ValueAxis>().Single(axis =>
            axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Right);
        secondary.NumberingFormat!.SourceLinked = true;
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(secondary.NumberingFormat.FormatCode?.Value,
            snapshot.Layout.SecondaryValueAxis!.NumberFormat);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PrimaryValueAxis_RejectsNonlinearOrReversedStaticProjection(bool reversed) {
        using var presentation = PowerPointPresentation.Create();
        var slide = presentation.AddSlide();
        var data = new OfficeChartData(new[] { "A", "B" },
            new[] { new OfficeChartSeries("Values", new[] { 1d, 10d }) });
        var chart = slide.AddChart(OfficeChartKind.ColumnClustered, data);
        var axis = slide.SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.ValueAxis>().Single();
        if (reversed) axis.Scaling!.AddChild(new C.Orientation { Val = C.OrientationValues.MaxMin }, true);
        else axis.Scaling!.AddChild(new C.LogBase { Val = 10 }, true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        chart.UpdateData(data);
        axis = slide.SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.ValueAxis>().Single();
        Assert.Equal(reversed, axis.Scaling!.GetFirstChild<C.Orientation>()?.Val?.Value == C.OrientationValues.MaxMin);
        Assert.Equal(!reversed, axis.Scaling.GetFirstChild<C.LogBase>()?.Val?.Value == 10);
        Assert.Empty(presentation.ValidateDocument());
    }

    [Fact]
    public void ScatterSnapshot_UsesBothReferencedNumericAxisScales() {
        using var presentation = PowerPointPresentation.Create();
        var slide = presentation.AddSlide();
        var chart = slide.AddChart(OfficeChartKind.Scatter,
            new OfficeChartData(new[] { "1", "3" }, new[] {
                new OfficeChartSeries("Values", new[] { 2d, 4d }, new[] { 1d, 3d }) }));
        var plot = slide.SlidePart.ChartParts.Single().ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        C.AxisId[] references = plot.GetFirstChild<C.ScatterChart>()!.Elements<C.AxisId>().ToArray();
        C.ValueAxis horizontal = plot.Elements<C.ValueAxis>().Single(axis =>
            axis.AxisId?.Val?.Value == references[0].Val?.Value);
        C.ValueAxis vertical = plot.Elements<C.ValueAxis>().Single(axis =>
            axis.AxisId?.Val?.Value == references[1].Val?.Value);
        horizontal.Scaling!.AddChild(new C.MinAxisValue { Val = 0 }, true);
        horizontal.Scaling.AddChild(new C.MaxAxisValue { Val = 5 }, true);
        horizontal.AddChild(new C.MajorUnit { Val = 1 }, true);
        vertical.Scaling!.AddChild(new C.MinAxisValue { Val = -2 }, true);
        vertical.Scaling.AddChild(new C.MaxAxisValue { Val = 8 }, true);
        vertical.AddChild(new C.MajorUnit { Val = 2 }, true);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(0d, snapshot.Layout.HorizontalAxisMinimum);
        Assert.Equal(5d, snapshot.Layout.HorizontalAxisMaximum);
        Assert.Equal(1d, snapshot.Layout.HorizontalAxisMajorUnit);
        Assert.Equal(-2d, snapshot.Layout.VerticalAxisMinimum);
        Assert.Equal(8d, snapshot.Layout.VerticalAxisMaximum);
        Assert.Equal(2d, snapshot.Layout.VerticalAxisMajorUnit);
    }

    [Fact]
    public void SecondaryValueAxis_PersistsThroughUpdatesAndSnapshots() {
        using var bytes = new MemoryStream();
        using var presentation = PowerPointPresentation.Create(bytes, new PowerPointCreateOptions());
        var slide = presentation.AddSlide();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Volume", new[] { 100d, 150d }),
            new OfficeChartSeries("Ratio", new[] { 1d, 2d }, null, null, null, true,
                renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        });
        var chart = slide.AddChart(OfficeChartKind.ColumnClustered, data);
        chart.SetSecondaryValueAxis(new OfficeChartValueAxisLayout(minimum: 0, maximum: 4, majorUnit: 1,
            minorUnit: 0.5, numberFormat: "0%", majorTickMark: OfficeChartAxisTickMark.Cross,
            minorTickMark: OfficeChartAxisTickMark.Outside).WithTitle("Secondary ratio"));
        var primary = slide.SlidePart.ChartParts.Single().ChartSpace!.Descendants<C.ValueAxis>()
            .Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Left);
        primary.Scaling!.AddChild(new C.MaxAxisValue { Val = 200 }, true);
        primary.AddChild(new C.NumberingFormat { FormatCode = "0.0", SourceLinked = false }, true);
        chart.UpdateData(data);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(200, snapshot.Layout.VerticalAxisMaximum);
        Assert.Equal("0.0", snapshot.Layout.VerticalAxisNumberFormat);
        Assert.Equal(4, snapshot.Layout.SecondaryValueAxis!.Maximum);
        Assert.Equal("0%", snapshot.Layout.SecondaryValueAxis.NumberFormat);
        Assert.Equal("Secondary ratio", snapshot.Layout.SecondaryValueAxis.Title);
        Assert.Empty(presentation.ValidateDocument());
        using var persistedBytes = new MemoryStream(presentation.ToBytes());
        using var reopened = PowerPointPresentation.Load(persistedBytes);
        Assert.True(reopened.Slides.Single().Charts.Single().TryGetOfficeSnapshot(out var persisted));
        Assert.Equal(0.5, persisted.Layout.SecondaryValueAxis!.MinorUnit);
        Assert.Equal(OfficeChartAxisTickMark.Cross, persisted.Layout.SecondaryValueAxis.MajorTickMark);
        Assert.Equal(OfficeChartAxisTickMark.Outside, persisted.Layout.SecondaryValueAxis.MinorTickMark);
        Assert.Equal("Secondary ratio", persisted.Layout.SecondaryValueAxis.Title);
    }
}
