using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using ExcelTableStyle = OfficeIMO.Excel.TableStyle;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests for creating tables with styles.
    /// </summary>
    public partial class Excel {
        [Fact]
        public void Test_AddTableWithStyle() {
            string filePath = Path.Combine(_directoryWithFiles, "TableStyles.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.SetCellValue(1, 1, "Name");
                sheet.SetCellValue(1, 2, "Value");
                sheet.SetCellValue(2, 1, "A");
                sheet.SetCellValue(2, 2, 10d);
                sheet.SetCellValue(3, 1, "B");
                sheet.SetCellValue(3, 2, 20d);

                sheet.AddTable("A1:B3", true, "MyTable", ExcelTableStyle.TableStyleMedium2);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                TableDefinitionPart tablePart = wsPart.TableDefinitionParts.FirstOrDefault();
                Assert.NotNull(tablePart);
                Table table = tablePart.Table;
                Assert.Equal("A1:B3", table.Reference.Value);
                Assert.Equal("MyTable", table.Name.Value);
                TableStyleInfo styleInfo = table.GetFirstChild<TableStyleInfo>();
                Assert.NotNull(styleInfo);
                Assert.Equal("TableStyleMedium2", styleInfo.Name.Value);
            }
        }
    }
}

                Assert.NotNull(styleInfo);
                Assert.Equal("TableStyleMedium2", styleInfo.Name.Value);
            }
        }
    }
}