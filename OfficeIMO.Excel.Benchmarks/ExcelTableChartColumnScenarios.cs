using System.Globalization;
using OfficeIMO.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// A table chart whose category and value columns are not adjacent (Region in B, Amount in E), over every data row.
/// The chart-data copy lane is how such a chart was built before series could reference separate ranges.
/// </summary>
internal static partial class ExcelLibraryComparisonRunner {
    private const string RealWorldTableChartColumnsScenario = "realworld-table-chart-columns";

    private static void AddTableChartColumnScenarioGroup(
        List<ExcelLibraryComparisonScenario> scenarios,
        IReadOnlySet<string>? scenarioFilter,
        IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows,
        int warmupIterations,
        int measuredIterations) {
        const string notes = "Write a sales table and chart Amount (column E) by Region (column B) over every row, then save.";
        AddScenarioGroup(scenarios, scenarioFilter, RealWorldTableChartColumnsScenario, warmupIterations, measuredIterations, [
            new LibraryComparisonCase("OfficeIMO.Excel", notes + " Series reference the table columns.", () => ByteCount(OfficeImoWriteTableChartColumns(rows, copyChartData: false))),
            new LibraryComparisonCase("OfficeIMO.Excel (chart-data copy)", notes + " Points are copied to a hidden chart-data sheet.", () => ByteCount(OfficeImoWriteTableChartColumns(rows, copyChartData: true))),
            new LibraryComparisonCase("EPPlus", notes, () => ByteCount(EpPlusWriteTableChartColumns(rows)))
        ]);
    }

    private static byte[] OfficeImoWriteTableChartColumns(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows, bool copyChartData) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddSheet("Data");
            ExcelBenchmarkScenarioFactory.InsertOfficeImoObjects(sheet, rows);
            ExcelBenchmarkScenarioFactory.AddOfficeImoTable(sheet, rows.Count);
            var options = new OfficeIMO.Excel.ExcelChartOptions { Row = 2, Column = 10, Width = OfficeLength.Pixels(640), Height = OfficeLength.Pixels(360), Title = "Amount by region" };
            if (copyChartData) {
                var data = new ExcelChartData(rows.Select(static row => row.Region),
                    [new ExcelChartSeries("Amount", rows.Select(static row => row.Amount))]);
                sheet.AddChart(OfficeChartType.ColumnClustered, data, options);
            } else {
                string last = (rows.Count + 1).ToString(CultureInfo.InvariantCulture);
                sheet.AddChartFromRanges(OfficeChartType.ColumnClustered, [new ExcelChartSeriesRange("Amount", "B2:B" + last, "E2:E" + last)], options);
            }

            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static byte[] EpPlusWriteTableChartColumns(IReadOnlyList<ExcelBenchmarkScenarioFactory.SalesRecord> rows) {
        using var stream = new MemoryStream();
        using (var package = new ExcelPackage(stream)) {
            var worksheet = package.Workbook.Worksheets.Add("Data");
            WriteSalesRows(worksheet, rows, includeAllColumns: true);
            int lastRow = rows.Count + 1;
            worksheet.Tables.Add(worksheet.Cells[1, 1, lastRow, 8], "SalesData");
            var chart = worksheet.Drawings.AddChart("AmountByRegion", eChartType.ColumnClustered);
            chart.Title.Text = "Amount by region";
            chart.SetPosition(1, 0, 9, 0);
            chart.SetSize(640, 360);
            chart.Series.Add(worksheet.Cells[2, 5, lastRow, 5], worksheet.Cells[2, 2, lastRow, 2]).Header = "Amount";
            package.Save();
        }

        return stream.ToArray();
    }
}
