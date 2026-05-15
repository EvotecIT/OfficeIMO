using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_AddPivotTableAndReadBack() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableBasic.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");

                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Product");
                sheet.CellValue(1, 3, "Sales");

                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, "A");
                sheet.CellValue(2, 3, 10);

                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, "A");
                sheet.CellValue(3, 3, 12);

                sheet.CellValue(4, 1, "East");
                sheet.CellValue(4, 2, "B");
                sheet.CellValue(4, 3, 7);

                sheet.AddPivotTable(
                    sourceRange: "A1:C4",
                    destinationCell: "E2",
                    name: "SalesPivot",
                    rowFields: new[] { "Region" },
                    columnFields: new[] { "Product" },
                    dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum, "Total Sales") },
                    pivotStyleName: "PivotStyleMedium9");

                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var pivots = document.GetPivotTables();
                Assert.Single(pivots);

                var pivot = pivots[0];
                Assert.Equal("SalesPivot", pivot.Name);
                Assert.Equal("Data", pivot.SheetName);
                Assert.Equal("Data", pivot.SourceSheet);
                Assert.Equal("A1:C4", pivot.SourceRange);
                Assert.Equal("PivotStyleMedium9", pivot.PivotStyle);
                Assert.Equal(ExcelPivotLayout.Compact, pivot.Layout);

                Assert.Contains("Region", pivot.RowFields);
                Assert.Contains("Product", pivot.ColumnFields);

                Assert.Single(pivot.DataFields);
                var dataField = pivot.DataFields[0];
                Assert.Equal("Sales", dataField.FieldName);
                Assert.Equal(DataConsolidateFunctionValues.Sum, dataField.Function);
                Assert.Equal("Total Sales", dataField.DisplayName);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_AddPivotTable_AppliesFieldOptionsAndNumberFormats() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotTableFieldOptions.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");

                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Product");
                sheet.CellValue(1, 3, "Sales");

                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, "A");
                sheet.CellValue(2, 3, 10.25);

                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, "A");
                sheet.CellValue(3, 3, 12.5);

                sheet.CellValue(4, 1, "East");
                sheet.CellValue(4, 2, "B");
                sheet.CellValue(4, 3, 7.75);

                sheet.AddPivotTable(
                    sourceRange: "A1:C4",
                    destinationCell: "E2",
                    name: "SalesPivot",
                    rowFields: new[] { "Region" },
                    dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum, "Total Sales", numberFormat: "$#,##0.00") },
                    fieldOptions: new[] {
                        new ExcelPivotFieldOptions("Region", sortType: FieldSortValues.Ascending, defaultSubtotal: false, subtotalTop: true, insertBlankRow: true)
                    });

                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                var pivotPart = workbookPart.WorksheetParts.SelectMany(part => part.PivotTableParts).Single();
                var pivotDefinition = pivotPart.PivotTableDefinition!;

                var regionField = pivotDefinition.PivotFields!.Elements<PivotField>().ElementAt(0);
                Assert.Equal(FieldSortValues.Ascending, regionField.SortType!.Value);
                Assert.False(regionField.DefaultSubtotal!.Value);
                Assert.True(regionField.SubtotalTop!.Value);
                Assert.True(regionField.InsertBlankRow!.Value);

                var dataField = pivotDefinition.DataFields!.Elements<DataField>().Single();
                uint numberFormatId = dataField.NumberFormatId!.Value;
                Assert.True(numberFormatId >= 164);

                var numberFormat = workbookPart.WorkbookStylesPart!.Stylesheet!.NumberingFormats!
                    .Elements<NumberingFormat>()
                    .Single(format => format.NumberFormatId!.Value == numberFormatId);
                Assert.Equal("$#,##0.00", numberFormat.FormatCode!.Value);
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var dataField = document.GetPivotTables().Single().DataFields.Single();
                Assert.True(dataField.NumberFormatId >= 164);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_AddPivotChartFromRange_WritesPivotSource() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelPivotChart.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");

                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Product");
                sheet.CellValue(1, 3, "Sales");

                sheet.CellValue(2, 1, "East");
                sheet.CellValue(2, 2, "A");
                sheet.CellValue(2, 3, 10);

                sheet.CellValue(3, 1, "West");
                sheet.CellValue(3, 2, "A");
                sheet.CellValue(3, 3, 12);

                sheet.CellValue(4, 1, "East");
                sheet.CellValue(4, 2, "B");
                sheet.CellValue(4, 3, 7);

                sheet.AddPivotTable(
                    sourceRange: "A1:C4",
                    destinationCell: "E2",
                    name: "SalesPivot",
                    rowFields: new[] { "Region" },
                    columnFields: new[] { "Product" },
                    dataFields: new[] { new ExcelPivotDataField("Sales", DataConsolidateFunctionValues.Sum, "Total Sales") });

                var chart = sheet.AddPivotChartFromRange("SalesPivot", "A1:C4", row: 7, column: 1, title: "Sales Pivot");
                Assert.True(chart.IsPivotChart);
                Assert.Equal("SalesPivot", chart.PivotTableName);

                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First(part => part.DrawingsPart?.ChartParts.Any() == true);
                var chartPart = worksheetPart.DrawingsPart!.ChartParts.Single();
                var pivotSource = chartPart.ChartSpace!.GetFirstChild<C.PivotSource>();

                Assert.NotNull(pivotSource);
                Assert.Equal("SalesPivot", pivotSource!.GetFirstChild<C.PivotTableName>()!.Text);
                Assert.Equal(0U, pivotSource.GetFirstChild<C.FormatId>()!.Val!.Value);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }
    }
}
