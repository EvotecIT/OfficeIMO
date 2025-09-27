using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void ReadRows_ReturnsNullForRowsWithoutCells() {
            string filePath = Path.Combine(_directoryWithFiles, "ReadRowsEmptyRows.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Header");
                sheet.CellValue(3, 1, "Value");
                document.Save();
            }

            using var reader = ExcelDocumentReader.Open(filePath);
            var rows = reader.GetSheet("Data").ReadRows("A1:A3").ToList();

            Assert.Equal(3, rows.Count);
            Assert.NotNull(rows[0]);
            Assert.Equal("Header", rows[0]![0]);
            Assert.Null(rows[1]);
            Assert.NotNull(rows[2]);
            Assert.Equal("Value", rows[2]![0]);
        }

        [Fact]
        public void ReadRowsAs_ThrowsForRowsWithoutCells() {
            string filePath = Path.Combine(_directoryWithFiles, "ReadRowsAsEmptyRows.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Header");
                sheet.CellValue(3, 1, "Value");
                document.Save();
            }

            using var reader = ExcelDocumentReader.Open(filePath);
            var sheetReader = reader.GetSheet("Data");

            var ex = Assert.Throws<InvalidOperationException>(() => sheetReader.ReadRowsAs<string>("A1:A3").ToList());
            Assert.Contains("Row 2", ex.Message);
            Assert.Contains("contains no cells", ex.Message);
        }
    }
}
