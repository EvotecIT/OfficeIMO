using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using Xunit;

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

                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                var tables = document.GetTables().ToList();
                Assert.Equal(2, tables.Count);

                var tableOne = tables.Single(t => t.Name == "TableOne");
                Assert.Equal("A1:B2", tableOne.Range);
                Assert.Equal("SheetOne", tableOne.SheetName);
                Assert.Equal(0, tableOne.SheetIndex);

                var tableTwo = tables.Single(t => t.Name == "TableTwo");
                Assert.Equal("A1:C2", tableTwo.Range);
                Assert.Equal("SheetTwo", tableTwo.SheetName);
                Assert.Equal(1, tableTwo.SheetIndex);
            }

            File.Delete(filePath);
        }
    }
}
