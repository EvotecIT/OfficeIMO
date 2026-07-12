using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using OfficeIMO.Excel;
using Xunit;
using TotalsRowFunctionValues = DocumentFormat.OpenXml.Spreadsheet.TotalsRowFunctionValues;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_GetTables_ReturnsSheetMetadata() {
            string filePath = Path.Combine(_directoryWithFiles, "Tables.Catalog.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet1 = document.AddWorkSheet("SheetOne");
                sheet1.CellValue(1, 1, "Name");
                sheet1.CellValue(1, 2, "Value");
                sheet1.CellValue(2, 1, "A");
                sheet1.CellValue(2, 2, 1d);
                sheet1.AddTable("A1:B2", true, "TableOne", TableStyle.TableStyleMedium9);

                var sheet2 = document.AddWorkSheet("SheetTwo");
                sheet2.CellValue(1, 1, "Name");
                sheet2.CellValue(1, 2, "Value");
                sheet2.CellValue(1, 3, "Count");
                sheet2.CellValue(2, 1, "B");
                sheet2.CellValue(2, 2, 2d);
                sheet2.CellValue(2, 3, 3);
                sheet2.AddTable("A1:C2", true, "TableTwo", TableStyle.TableStyleMedium9);
                sheet2.SetTableTotalsByName("TableTwo", new Dictionary<string, TotalsRowFunctionValues> {
                    ["Value"] = TotalsRowFunctionValues.Sum,
                    ["Count"] = TotalsRowFunctionValues.Sum,
                });

                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                var tables = document.GetTables().ToList();
                Assert.Equal(2, tables.Count);

                var tableOne = tables.Single(t => t.Name == "TableOne");
                Assert.Equal("A1:B2", tableOne.Range);
                Assert.Equal("SheetOne", tableOne.SheetName);
                Assert.Equal(0, tableOne.SheetIndex);
                Assert.Equal("TableStyleMedium9", tableOne.StyleName);
                Assert.True(tableOne.HasHeaderRow);
                Assert.True(tableOne.HasAutoFilter);
                Assert.False(tableOne.TotalsRowShown);
                Assert.Equal(new[] { "Name", "Value" }, tableOne.Columns.Select(column => column.Name).ToArray());

                var tableTwo = tables.Single(t => t.Name == "TableTwo");
                Assert.Equal("A1:C2", tableTwo.Range);
                Assert.Equal("SheetTwo", tableTwo.SheetName);
                Assert.Equal(1, tableTwo.SheetIndex);
                Assert.True(tableTwo.TotalsRowShown);
                Assert.Equal("sum", tableTwo.Columns.Single(column => column.Name == "Value").TotalsRowFunction);
                Assert.Equal("sum", tableTwo.Columns.Single(column => column.Name == "Count").TotalsRowFunction);
            }

            File.Delete(filePath);
        }
    }
}
