using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelChartSecondaryAxisLayoutTests {
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
}
