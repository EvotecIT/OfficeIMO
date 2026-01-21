using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

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
        }
    }
}
