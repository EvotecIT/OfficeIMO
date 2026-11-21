using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class ExcelChartSecondaryAxisLayoutTests {
    [Theory]
    [InlineData("manual")]
    [InlineData("overlay")]
    [InlineData("implicitOverlay")]
    public void SecondaryValueAxis_UnsupportedTitlePlacementRejectsSnapshot(string placement) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Results");
        var chart = sheet.AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "A" }, new[] {
            new OfficeSeries("Count", new[] { 100d }),
            new OfficeSeries("Ratio", new[] { 1d }, null, null, null, true,
                renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 });
        chart.SetSecondaryValueAxis(new OfficeChartValueAxisLayout().WithTitle("Ratio"));
        var title = sheet.WorksheetPart.DrawingsPart!.ChartParts.Single().ChartSpace!.Descendants<C.ValueAxis>()
            .Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Right).GetFirstChild<C.Title>()!;
        if (placement == "overlay") title.GetFirstChild<C.Overlay>()!.Val = true;
        else if (placement == "implicitOverlay") title.GetFirstChild<C.Overlay>()!.Val = null;
        else title.GetFirstChild<C.Layout>()!.Append(new C.ManualLayout(new C.Left { Val = .3 }));
        Assert.False(chart.TryGetSnapshot(out _));
    }

    [Theory]
    [InlineData("zeroUnit")]
    [InlineData("invertedBounds")]
    [InlineData("malformedNumber")]
    [InlineData("displayUnits")]
    [InlineData("deleted")]
    public void SecondaryValueAxis_InvalidImportedSettingsRejectSnapshot(string setting) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Results");
        var chart = sheet.AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "A" }, new[] {
            new OfficeSeries("Count", new[] { 100d }),
            new OfficeSeries("Ratio", new[] { 1d }, null, null, null, true,
                renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 });
        var secondary = sheet.WorksheetPart.DrawingsPart!.ChartParts.Single().ChartSpace!.Descendants<C.ValueAxis>()
            .Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Right);
        if (setting == "deleted") {
            chart.SetSecondaryValueAxis(new OfficeChartValueAxisLayout().WithTitle("Ratio"));
            secondary.AddChild(new C.Delete { Val = true }, true);
        } else if (setting == "zeroUnit") secondary.AddChild(new C.MajorUnit { Val = 0 }, true);
        else if (setting == "invertedBounds") {
            secondary.Scaling!.AddChild(new C.MinAxisValue { Val = 2 }, true);
            secondary.Scaling.AddChild(new C.MaxAxisValue { Val = 1 }, true);
        } else if (setting == "displayUnits") secondary.AddChild(new C.DisplayUnits(
            new C.BuiltInUnit { Val = C.BuiltInUnitValues.Thousands }), true);
        else secondary.Scaling!.AddChild(new C.MaxAxisValue { Val = new DocumentFormat.OpenXml.DoubleValue { InnerText = "invalid" } }, true);
        Assert.False(chart.TryGetSnapshot(out _));
    }

    [Fact]
    public void SecondaryValueAxis_ReopenSnapshotPreservesIndependentSettings() {
        using var document = ExcelDocument.Create();
        var chart = document.AddSheet("Results").AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeSeries("Volume", new[] { 100d, 150d }),
                new OfficeSeries("Ratio", new[] { 1d, 2d }, null, null, null, true,
                    renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary)
            }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 });
        chart.SetValueAxisScale(minimum: 0, maximum: 200).SetValueAxisNumberFormat("0.0");
        chart.SetValueAxisScale(minimum: 0, maximum: 4, majorUnit: 1, axisGroup: OfficeChartAxisGroup.Secondary)
            .SetValueAxisNumberFormat("0%", false, OfficeChartAxisGroup.Secondary);
        chart.SetSecondaryValueAxis(new OfficeChartValueAxisLayout(minimum: 0, maximum: 4,
            majorUnit: 1, minorUnit: 0.5, numberFormat: "0%",
            majorTickMark: OfficeChartAxisTickMark.Cross, minorTickMark: OfficeChartAxisTickMark.Outside)
            .WithTitle("Secondary ratio"));
        using var bytes = new MemoryStream(document.ToBytes());
        using var reopened = ExcelDocument.Load(bytes);
        Assert.Empty(reopened.ValidateDocument());
        var imported = reopened.Sheets.Single(sheet => sheet.Name == "Results").Charts.Single();
        Assert.True(imported.TryGetSnapshot(out var snapshot));
        Assert.Equal(200, snapshot.Layout!.VerticalAxisMaximum);
        Assert.Equal("0.0", snapshot.Layout.VerticalAxisNumberFormat);
        Assert.Equal(4, snapshot.Layout.SecondaryValueAxis!.Maximum);
        Assert.Equal(1, snapshot.Layout.SecondaryValueAxis.MajorUnit);
        Assert.Equal("0%", snapshot.Layout.SecondaryValueAxis.NumberFormat);
        Assert.Equal(0.5, snapshot.Layout.SecondaryValueAxis.MinorUnit);
        Assert.Equal(OfficeChartAxisTickMark.Cross, snapshot.Layout.SecondaryValueAxis.MajorTickMark);
        Assert.Equal(OfficeChartAxisTickMark.Outside, snapshot.Layout.SecondaryValueAxis.MinorTickMark);
        Assert.Equal("Secondary ratio", snapshot.Layout.SecondaryValueAxis.Title);
    }

    [Fact]
    public void SecondaryValueAxis_LogarithmicScaleIsDrawnWithSourceLabels() {
        using var document = ExcelDocument.Create();
        var chart = document.AddSheet("Results").AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "A", "B", "C" }, new[] {
                new OfficeSeries("Count", new[] { 1d, 2d, 3d }),
                new OfficeSeries("Ratio", new[] { 2d, 30d, 400d }, null, null, null, true,
                    renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary)
            }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 });
        chart.SetSecondaryValueAxis(new OfficeChartValueAxisLayout().WithLogarithmicScale(10D));

        Assert.True(chart.TryGetSnapshot(out var snapshot));
        Assert.Equal(10D, snapshot.Layout!.SecondaryValueAxis!.LogBase);
        string svg = System.Text.Encoding.UTF8.GetString(chart.ToImage(OfficeImageExportFormat.Svg).Bytes);
        // The secondary axis spans the powers 1, 10, 100 and 1000 around the values.
        Assert.Contains(">10<", svg);
        Assert.Contains(">100<", svg);
    }

    [Fact]
    public void SecondaryValueAxis_ResolvesLinkedNumberFormatFromItsWorksheetValues() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Results");
        sheet.Cell(1, 1).SetValue("Region"); sheet.Cell(1, 2).SetValue("Count"); sheet.Cell(1, 3).SetValue("Ratio");
        sheet.Cell(2, 1).SetValue("North"); sheet.Cell(2, 2).SetValue(100); sheet.Cell(2, 3).SetValue(.5);
        sheet.Cell(3, 1).SetValue("South"); sheet.Cell(3, 2).SetValue(200); sheet.Cell(3, 3).SetValue(.75);
        sheet.Cell(2, 3).SetNumberFormat("0%"); sheet.Cell(3, 3).SetNumberFormat("0%");
        var chart = sheet.AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "North", "South" }, new[] {
                new OfficeSeries("Count", new[] { 100d, 200d }),
                new OfficeSeries("Ratio", new[] { .5, .75 }, null, null, null, true,
                    renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary) }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 5 });
        var native = sheet.WorksheetPart.DrawingsPart!.ChartParts.Single().ChartSpace!;
        var secondary = native.Descendants<C.ValueAxis>()
            .Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Right);
        secondary.NumberingFormat!.FormatCode = "General";
        secondary.NumberingFormat.SourceLinked = true;
        var values = native.Descendants<C.LineChartSeries>().Single().GetFirstChild<C.Values>()!;
        values.RemoveAllChildren();
        values.Append(new C.NumberReference(new C.Formula("Results!$C$2:$C$3"),
            new C.NumberingCache(new C.FormatCode("General"), new C.PointCount { Val = 2 },
                new C.NumericPoint(new C.NumericValue("0.5")) { Index = 0 },
                new C.NumericPoint(new C.NumericValue("0.75")) { Index = 1 })));

        Assert.True(chart.TryGetSnapshot(out var snapshot));
        Assert.Equal("0%", snapshot.Layout!.SecondaryValueAxis!.NumberFormat);
        var visual = Assert.Single(sheet.Range("A1:J12").CreateVisualSnapshot().Charts);
        Assert.Equal("0%", visual.Snapshot.Layout!.SecondaryValueAxis!.NumberFormat);
    }

    [Fact]
    public void SecondaryValueAxis_OverlongLinkedFormatDrawsAsGeneral() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Results");
        var chart = sheet.AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "North", "South" }, new[] {
                new OfficeSeries("Count", new[] { 100d, 200d }),
                new OfficeSeries("Ratio", new[] { .5, .75 }, null, null, null, true,
                    renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary) }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 5 });
        var native = sheet.WorksheetPart.DrawingsPart!.ChartParts.Single().ChartSpace!;
        var secondary = native.Descendants<C.ValueAxis>()
            .Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Right);
        // The worksheet range cannot be read, so the linked axis falls back to its own cached code.
        secondary.NumberingFormat!.FormatCode = "0\"" + new string('x', 1100) + "\"";
        secondary.NumberingFormat.SourceLinked = true;
        var values = native.Descendants<C.LineChartSeries>().Single().GetFirstChild<C.Values>()!;
        values.RemoveAllChildren();
        values.Append(new C.NumberReference(new C.Formula("Missing!$C$2:$C$3"),
            new C.NumberingCache(new C.FormatCode("General"), new C.PointCount { Val = 2 },
                new C.NumericPoint(new C.NumericValue("0.5")) { Index = 0 },
                new C.NumericPoint(new C.NumericValue("0.75")) { Index = 1 })));

        Assert.True(chart.TryGetSnapshot(out var snapshot));
        Assert.Null(snapshot.Layout!.SecondaryValueAxis!.NumberFormat);
    }
}
