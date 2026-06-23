using System.Data;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_DashboardChartPreset_CreatesStyledChartFromRange() {
            string filePath = Path.Combine(_directoryWithFiles, "DashboardChartPreset.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Dashboard");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Revenue");
                sheet.CellValue(2, 1, "EU");
                sheet.CellValue(2, 2, 100);
                sheet.CellValue(3, 1, "US");
                sheet.CellValue(3, 2, 120);

                ExcelChart chart = sheet.AddDashboardChart(new ExcelDashboardChartOptions {
                    Preset = ExcelDashboardChartPreset.CompactComparison,
                    Range = "A1:B3",
                    Row = 1,
                    Column = 4,
                    Title = "Revenue"
                });

                Assert.Equal(ExcelChartType.BarClustered, chart.ChartType);
                Assert.Equal("Revenue", chart.Title);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelChart chart = Assert.Single(document["Dashboard"].Charts);
                Assert.Equal(ExcelChartType.BarClustered, chart.ChartType);
                Assert.Equal("Revenue", chart.Title);
            }
        }

        [Fact]
        public void Test_DashboardBuilder_CreatesTableAndChart() {
            string filePath = Path.Combine(_directoryWithFiles, "DashboardBuilder.xlsx");
            var data = new DataTable("Sales");
            data.Columns.Add("Region", typeof(string));
            data.Columns.Add("Revenue", typeof(int));
            data.Rows.Add("EU", 100);
            data.Rows.Add("US", 120);

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Dashboard");
                ExcelDashboardResult result = sheet.AddDashboard(data, new ExcelDashboardOptions {
                    Title = "Sales Dashboard",
                    Subtitle = "Monthly revenue",
                    TableName = "SalesTable",
                    ChartPreset = ExcelDashboardChartPreset.CompactComparison,
                    ChartTitle = "Revenue"
                });

                Assert.Equal("A3:B5", result.TableRange);
                Assert.Equal("SalesTable", result.TableName);
                Assert.NotNull(result.Chart);
                Assert.Equal(ExcelChartType.BarClustered, result.Chart!.ChartType);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelSheet sheet = document["Dashboard"];
                Assert.True(sheet.TryGetCellText(1, 1, out string? title));
                Assert.Equal("Sales Dashboard", title);
                ExcelChart chart = Assert.Single(sheet.Charts);
                Assert.Equal("Revenue", chart.Title);
            }
        }

        [Fact]
        public void Test_DashboardBuilder_ReturnsActualSanitizedTableName() {
            string filePath = Path.Combine(_directoryWithFiles, "DashboardBuilderResolvedTableName.xlsx");
            var data = new DataTable("Sales Data");
            data.Columns.Add("Region", typeof(string));
            data.Columns.Add("Revenue", typeof(int));
            data.Rows.Add("EU", 100);

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Dashboard");
                ExcelDashboardResult result = sheet.AddDashboard(data, new ExcelDashboardOptions {
                    TableName = "Sales Data",
                    AddChart = false
                });

                Assert.Equal("Sales_Data", result.TableName);
                Assert.Equal(result.TableRange, sheet.GetTableRange(result.TableName!));
            }
        }

        [Fact]
        public void Test_DashboardBuilder_RejectsWorksheetBoundsOverflow() {
            string filePath = Path.Combine(_directoryWithFiles, "DashboardBuilder.Bounds.xlsx");
            var data = new DataTable("Sales");
            data.Columns.Add("Region", typeof(string));
            data.Columns.Add("Revenue", typeof(int));
            data.Rows.Add("EU", 100);

            using var document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Dashboard");

            Assert.Throws<ArgumentException>(() => sheet.AddDashboard(data, new ExcelDashboardOptions {
                TableRow = A1.MaxRows,
                AddChart = false
            }));

            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddDashboard(data, new ExcelDashboardOptions {
                TableRow = 3,
                ChartColumn = A1.MaxColumns + 1
            }));
        }
    }
}
