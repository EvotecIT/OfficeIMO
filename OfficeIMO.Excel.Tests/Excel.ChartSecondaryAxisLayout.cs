using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class ExcelChartSecondaryAxisLayoutTests {
    [Theory]
    [InlineData("zeroUnit")]
    [InlineData("invertedBounds")]
    [InlineData("malformedNumber")]
    [InlineData("logarithmic")]
    [InlineData("reversed")]
    [InlineData("displayUnits")]
    public void SecondaryValueAxis_InvalidImportedSettingsRejectSnapshot(string setting) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddWorksheet("Results");
        var chart = sheet.AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A" }, new[] {
            new OfficeChartSeries("Count", new[] { 100d }),
            new OfficeChartSeries("Ratio", new[] { 1d }, null, null, null, true,
                renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        }), 1, 1);
        var secondary = sheet.WorksheetPart.DrawingsPart!.ChartParts.Single().ChartSpace!.Descendants<C.ValueAxis>()
            .Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Right);
        if (setting == "zeroUnit") secondary.AddChild(new C.MajorUnit { Val = 0 }, true);
        else if (setting == "invertedBounds") {
            secondary.Scaling!.AddChild(new C.MinAxisValue { Val = 2 }, true);
            secondary.Scaling.AddChild(new C.MaxAxisValue { Val = 1 }, true);
        } else if (setting == "logarithmic") secondary.Scaling!.AddChild(new C.LogBase { Val = 10 }, true);
        else if (setting == "reversed") secondary.Scaling!.AddChild(new C.Orientation { Val = C.OrientationValues.MaxMin }, true);
        else if (setting == "displayUnits") secondary.AddChild(new C.DisplayUnits(
            new C.BuiltInUnit { Val = C.BuiltInUnitValues.Thousands }), true);
        else secondary.Scaling!.AddChild(new C.MaxAxisValue { Val = new DocumentFormat.OpenXml.DoubleValue { InnerText = "invalid" } }, true);
        Assert.False(chart.TryGetSnapshot(out _));
    }

    [Fact]
    public void SecondaryValueAxis_ReopenSnapshotPreservesIndependentSettings() {
        using var document = ExcelDocument.Create();
        var chart = document.AddWorksheet("Results").AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Volume", new[] { 100d, 150d }),
                new OfficeChartSeries("Ratio", new[] { 1d, 2d }, null, null, null, true,
                    renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary)
            }), 1, 1);
        chart.SetValueAxisScale(minimum: 0, maximum: 200).SetValueAxisNumberFormat("0.0");
        chart.SetValueAxisScale(minimum: 0, maximum: 4, majorUnit: 1, axisGroup: OfficeChartAxisGroup.Secondary)
            .SetValueAxisNumberFormat("0%", false, OfficeChartAxisGroup.Secondary);
        chart.SetSecondaryValueAxis(new OfficeChartValueAxisLayout(minimum: 0, maximum: 4,
            majorUnit: 1, minorUnit: 0.5, numberFormat: "0%",
            majorTickMark: OfficeChartAxisTickMark.Cross, minorTickMark: OfficeChartAxisTickMark.Outside));
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
    }

    [Fact]
    public void SecondaryValueAxis_ResolvesLinkedNumberFormatFromItsWorksheetValues() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddWorksheet("Results");
        sheet.CellValue(1, 1, "Region"); sheet.CellValue(1, 2, "Count"); sheet.CellValue(1, 3, "Ratio");
        sheet.CellValue(2, 1, "North"); sheet.CellValue(2, 2, 100); sheet.CellValue(2, 3, .5);
        sheet.CellValue(3, 1, "South"); sheet.CellValue(3, 2, 200); sheet.CellValue(3, 3, .75);
        sheet.CellAt(2, 3).Percent(0); sheet.CellAt(3, 3).Percent(0);
        var chart = sheet.AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "North", "South" }, new[] {
                new OfficeChartSeries("Count", new[] { 100d, 200d }),
                new OfficeChartSeries("Ratio", new[] { .5, .75 }, null, null, null, true,
                    renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary) }), 1, 5);
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
    }
}
