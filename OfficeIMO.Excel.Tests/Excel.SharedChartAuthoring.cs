using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelCharts_SharedContractPreservesComboKindsAndSecondaryAxis() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SharedContract.Combo.xlsx");
            var sharedData = new OfficeChartData(
                new[] { "Q1", "Q2", "Q3" },
                new[] {
                    new OfficeChartSeries(
                        "Adoption",
                        new[] { 42d, 55d, 68d },
                        xValues: null,
                        color: null,
                        pointColors: null,
                        showMarkers: false,
                        showInLegend: true,
                        renderKind: OfficeChartKind.ColumnClustered),
                    new OfficeChartSeries(
                        "Conversion",
                        new[] { 3.2d, 5.1d, 8.4d },
                        xValues: null,
                        color: null,
                        pointColors: null,
                        showMarkers: true,
                        showInLegend: true,
                        renderKind: OfficeChartKind.Line,
                        axisGroup: OfficeChartAxisGroup.Secondary)
                });

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Shared");
                sheet.AddChart(
                    OfficeChartKind.ColumnClustered,
                    sharedData,
                    row: 1,
                    column: 5,
                    widthPixels: 640,
                    heightPixels: 360,
                    title: "Shared performance");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                OpenXmlValidator validator = new();
                Assert.Empty(validator.Validate(spreadsheet));
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelSheet sheet = document.Sheets.Single(item => item.Name == "Shared");
                ExcelChart chart = Assert.Single(sheet.Charts);
                Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
                Assert.Equal("Shared performance", snapshot.Title);
                Assert.Equal(new[] { "Q1", "Q2", "Q3" }, snapshot.Data.Categories);
                Assert.Collection(
                    snapshot.Data.Series,
                    series => {
                        Assert.Equal("Adoption", series.Name);
                        Assert.Equal(ExcelChartType.ColumnClustered, series.ChartType);
                        Assert.Equal(ExcelChartAxisGroup.Primary, series.AxisGroup);
                    },
                    series => {
                        Assert.Equal("Conversion", series.Name);
                        Assert.Equal(ExcelChartType.Line, series.ChartType);
                        Assert.Equal(ExcelChartAxisGroup.Secondary, series.AxisGroup);
                    });
            }
        }
    }
}
