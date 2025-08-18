using System.IO;
using OfficeIMO.Excel;
using ExcelTableStyle = OfficeIMO.Excel.TableStyle;
using Xunit;
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            using (var archive = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Read)) {
                var tableEntry = archive.GetEntry("xl/tables/table1.xml");
                Assert.NotNull(tableEntry);
                using (var reader = new StreamReader(tableEntry.Open())) {
                    string tableXml = reader.ReadToEnd();
                    Assert.Contains("tableStyleInfo", tableXml);
                    Assert.Contains("TableStyleMedium2", tableXml);
                }
                var workbookEntry = archive.GetEntry("xl/workbook.xml");
                Assert.NotNull(workbookEntry);
                using (var reader = new StreamReader(workbookEntry.Open())) {
                    string workbookXml = reader.ReadToEnd();
                    Assert.Contains("tableStyles", workbookXml);
                    Assert.Contains("defaultTableStyle=\"TableStyleMedium2\"", workbookXml);
                }
            }
        }
    }
}
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
