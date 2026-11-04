using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Reader_ReadUsedRange_MaterializesTableBackedValues() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderReadUsedRangeTable.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddSheet("Data");
                    sheet.Cell(1, 1).SetValue("Name");
                    sheet.Cell(1, 2).SetValue("Value");
                    sheet.Cell(2, 1).SetValue("Alpha");
                    sheet.Cell(2, 2).SetValue(10);
                    sheet.AddTable("A1:B2", hasHeader: true, name: "DataTable", style: OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium2, includeAutoFilter: true);
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                object?[,] values = reader.GetSheet("Data").ReadUsedRange();

                Assert.Equal(2, values.GetLength(0));
                Assert.Equal(2, values.GetLength(1));
                Assert.Equal("Name", values[0, 0]);
                Assert.Equal("Value", values[0, 1]);
                Assert.Equal("Alpha", values[1, 0]);
                Assert.Equal(10d, values[1, 1]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadUsedRange_IgnoresStaleOversizedDimension() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderReadUsedRangeStaleDimension.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    document.AddSheet("Data").Cell(3, 2).SetValue("Value");
                    document.Save();
                }

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    var worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet;
                    worksheet.GetFirstChild<SheetDimension>()!.Reference = "A1:Z1000";
                    worksheet.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                object?[,] values = reader.GetSheet("Data").ReadUsedRange();

                Assert.Equal(1, values.GetLength(0));
                Assert.Equal(1, values.GetLength(1));
                Assert.Equal("Value", values[0, 0]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Reader_ReadUsedRange_IncludesCellsOutsideTableDimension() {
            string filePath = Path.Combine(_directoryWithFiles, "ReaderReadUsedRangeOutsideTable.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var sheet = document.AddSheet("Data");
                    sheet.Cell(1, 1).SetValue("Name");
                    sheet.Cell(2, 1).SetValue("Alpha");
                    sheet.AddTable("A1:A2", hasHeader: true, name: "DataTable", style: OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium2, includeAutoFilter: true);
                    sheet.Cell(4, 4).SetValue("Outside");
                    document.Save();
                }

                using var reader = ExcelDocumentReader.Open(filePath);
                object?[,] values = reader.GetSheet("Data").ReadUsedRange();

                Assert.Equal(4, values.GetLength(0));
                Assert.Equal(4, values.GetLength(1));
                Assert.Equal("Name", values[0, 0]);
                Assert.Equal("Alpha", values[1, 0]);
                Assert.Equal("Outside", values[3, 3]);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
