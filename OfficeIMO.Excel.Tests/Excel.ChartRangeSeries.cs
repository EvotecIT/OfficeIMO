using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelCharts_FromRanges_ReferenceNonAdjacentColumnsWithoutChartDataSheet() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.NonAdjacentRanges.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddSheet("Data");
                sheet.CellValue(1, 1, "Id"); sheet.CellValue(1, 2, "Region"); sheet.CellValue(1, 3, "Note");
                sheet.CellValue(1, 4, "Sales"); sheet.CellValue(1, 5, "Target");
                string[] regions = { "NA", "EU", "APAC" };
                for (int i = 0; i < regions.Length; i++) {
                    sheet.CellValue(i + 2, 1, i + 1d);
                    sheet.CellValue(i + 2, 2, regions[i]);
                    sheet.CellValue(i + 2, 3, "n/a");
                    sheet.CellValue(i + 2, 4, 10d * (i + 1));
                    sheet.CellValue(i + 2, 5, 15d * (i + 1));
                }

                sheet.AddChartFromRanges(OfficeChartType.ColumnClustered, new[] {
                    new ExcelChartSeriesRange("Sales", "B2:B4", "D2:D4"),
                    new ExcelChartSeriesRange("Target", "B2:B4", "E2:E4")
                }, new OfficeIMO.Excel.ExcelChartOptions { Row = 2, Column = 8, Title = "Sales by region" });
                sheet.AddChartFromRanges(OfficeChartType.Line, new[] { new ExcelChartSeriesRange("Target", "B2:B4", "E2:E4") },
                    new OfficeIMO.Excel.ExcelChartOptions { Row = 22, Column = 8, IncludeCachedData = false });
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                Assert.Single(spreadsheet.WorkbookPart!.Workbook.Sheets!.OfType<Sheet>());
                var charts = GetWorksheetPartByName(spreadsheet, "Data").DrawingsPart!.ChartParts
                    .Select(part => part.ChartSpace!.Descendants<C.PlotArea>().Single()).ToList();
                C.PlotArea cached = charts.Single(plot => plot.GetFirstChild<C.BarChart>() != null);
                C.PlotArea linked = charts.Single(plot => plot.GetFirstChild<C.LineChart>() != null);

                var series = cached.Descendants<C.BarChartSeries>().ToList();
                Assert.Equal(new[] { "'Data'!$B$2:$B$4", "'Data'!$B$2:$B$4" },
                    series.Select(item => item.GetFirstChild<C.CategoryAxisData>()!.Descendants<C.Formula>().Single().Text));
                Assert.Equal(new[] { "'Data'!$D$2:$D$4", "'Data'!$E$2:$E$4" },
                    series.Select(item => item.GetFirstChild<C.Values>()!.Descendants<C.Formula>().Single().Text));
                Assert.Equal(new[] { "Sales", "Target" }, series.Select(item => item.GetFirstChild<C.SeriesText>()!.InnerText));
                Assert.Equal(ExpectedRegions(), series[0].GetFirstChild<C.CategoryAxisData>()!.Descendants<C.StringPoint>().Select(point => point.InnerText));
                Assert.Equal(new[] { "15", "30", "45" }, series[1].GetFirstChild<C.Values>()!.Descendants<C.NumericPoint>().Select(point => point.InnerText));

                C.LineChartSeries linkedSeries = linked.Descendants<C.LineChartSeries>().Single();
                Assert.Equal("'Data'!$E$2:$E$4", linkedSeries.GetFirstChild<C.Values>()!.Descendants<C.Formula>().Single().Text);
                Assert.Empty(linkedSeries.Descendants<C.NumberingCache>());

                var errors = new OpenXmlValidator().Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var reloaded = ExcelDocument.Load(filePath)) {
                ExcelChart chart = reloaded["Data"].Charts.Single(item => item.ChartType == OfficeChartType.ColumnClustered);
                Assert.True(chart.TryGetData(out ExcelChartData data));
                Assert.Equal(ExpectedRegions(), data.Categories);
                Assert.Equal(new[] { 10d, 20d, 30d }, data.Series[0].Values);
            }

            static string[] ExpectedRegions() => new[] { "NA", "EU", "APAC" };
        }
    }
}
