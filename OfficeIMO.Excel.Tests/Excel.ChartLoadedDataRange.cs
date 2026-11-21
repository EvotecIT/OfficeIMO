using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelChartLoadedDataRangeTests {
    [Fact]
    public void LoadedChartReportsItsSavedDataRangeBeforeDataIsRead() {
        using var bytes = new MemoryStream();
        string expected;
        using (var book = ExcelDocument.Create()) {
            var chart = book.AddSheet("Report").AddChart(OfficeChartType.ColumnClustered, new ExcelChartData(new[] { "Q1", "Q2" },
                new[] { new ExcelChartSeries("Revenue", new[] { 12D, 18D }) }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 });
            expected = chart.DataRange!.DataRangeA1;
            book.Save(bytes);
        }

        bytes.Position = 0;
        using var loaded = ExcelDocument.Load(bytes);
        ExcelChart reloaded = Assert.Single(loaded.GetSheet("Report").Charts);
        Assert.Equal(expected, reloaded.DataRange?.DataRangeA1);
    }
}
