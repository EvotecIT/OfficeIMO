using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void EnsureValidUniqueTableName_HandlesSpaces() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.Names.Spaces.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Col1");
                sheet.CellValue(1, 2, "Col2");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 1);
                sheet.AddTable("A1:B2", hasHeader: true, name: "My Table", TableStyle.TableStyleMedium9);
                document.Save();
            }
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var tbl = wsPart.TableDefinitionParts.First().Table;
                Assert.Equal("My_Table", tbl!.Name!.Value);
            }
        }

        [Fact]
        public void EnsureValidUniqueTableName_HandlesSpecialCharacters() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.Names.Special.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "X");
                sheet.CellValue(1, 2, "Y");
                sheet.AddTable("A1:B1", hasHeader: true, name: "Table#1!", TableStyle.TableStyleMedium9);
                document.Save();
            }
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var tbl = wsPart.TableDefinitionParts.First().Table;
                Assert.Equal("Table_1_", tbl!.Name!.Value);
            }
        }

        [Fact]
        public void EnsureValidUniqueTableName_EnsuresUniqueness() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.Names.Unique.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "A");
                sheet.AddTable("A1:A1", hasHeader: true, name: "Table", TableStyle.TableStyleMedium9);
                sheet.CellValue(2, 1, "B");
                sheet.AddTable("A2:A2", hasHeader: true, name: "Table", TableStyle.TableStyleMedium9);
                document.Save();
            }
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var names = wsPart.TableDefinitionParts.Select(t => t.Table!.Name!.Value).ToArray();
                Assert.Contains("Table", names);
                Assert.Contains("Table2", names);
            }
        }

        [Fact]
        public void EnsureValidUniqueTableName_HandlesDigitPrefix() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.Names.Digit.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "X");
                sheet.AddTable("A1:A1", hasHeader: true, name: "123Report", TableStyle.TableStyleMedium9);
                document.Save();
            }
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var tbl = wsPart.TableDefinitionParts.First().Table;
                Assert.Equal("_123Report", tbl!.Name!.Value);
            }
        }

        [Fact]
        public void EnsureValidUniqueTableName_EmptyName_DefaultsToTableWithId() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.Names.Empty.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "X");
                sheet.AddTable("A1:A1", hasHeader: true, name: string.Empty, TableStyle.TableStyleMedium9);
                document.Save();
            }
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var name = wsPart.TableDefinitionParts.First().Table!.Name!.Value!;
                Assert.StartsWith("Table", name);
                Assert.True(name.Length > 0);
            }
        }

        [Fact]
        public void EnsureValidUniqueTableName_VeryLongName_IsTrimmed() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.Names.Long.xlsx");
            string longName = new string('A', 300);
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "X");
                sheet.AddTable("A1:A1", hasHeader: true, name: longName, TableStyle.TableStyleMedium9);
                document.Save();
            }
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var name = wsPart.TableDefinitionParts.First().Table!.Name!.Value!;
                Assert.True(name.Length <= 255);
            }
        }

        [Fact]
        public void EnsureValidUniqueTableName_UnicodeCharacters_ArePreserved() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.Names.Unicode.xlsx");
            string unicodeName = "Nazwa テーブル Имя";
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "X");
                sheet.AddTable("A1:A1", hasHeader: true, name: unicodeName, TableStyle.TableStyleMedium9);
                document.Save();
            }
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var name = wsPart.TableDefinitionParts.First().Table!.Name!.Value!;
                Assert.Contains("Nazwa", name);
                Assert.Contains("テーブル", name);
                Assert.Contains("Имя", name);
                Assert.Contains('_', name); // spaces become underscores
            }
        }

        [Fact]
        public void AddTable_StrictValidation_InvalidCharacters_Throws() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.Names.Strict.Invalid.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "X");
                Assert.Throws<ArgumentException>(() =>
                    sheet.AddTable("A1:A1", hasHeader: true, name: "Bad Name!", TableStyle.TableStyleMedium9, includeAutoFilter: true, validationMode: TableNameValidationMode.Strict));
            }
        }

        [Fact]
        public void AddTable_StrictValidation_DigitPrefix_Throws() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.Names.Strict.Digit.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "X");
                Assert.Throws<ArgumentException>(() =>
                    sheet.AddTable("A1:A1", hasHeader: true, name: "1Hello", TableStyle.TableStyleMedium9, includeAutoFilter: true, validationMode: TableNameValidationMode.Strict));
            }
        }

        [Fact]
        public async Task TableId_Generation_IsThreadSafe_AcrossMultipleWorkbooks() {
            string dir = _directoryWithFiles;
            async Task<string> CreateOneAsync(int idx) {
                string fp = Path.Combine(dir, $"Table.Concurrent.WB{idx}.xlsx");
                using (var doc = ExcelDocument.Create(fp)) {
                    var s = doc.AddWorkSheet($"S{idx}");
                    s.CellValue(1, 1, "A");
                    s.CellValue(1, 2, "B");
                    s.AddTable("A1:B1", hasHeader: true, name: $"T{idx}", TableStyle.TableStyleMedium9);
                    doc.Save();
                }
                return fp;
            }

            var tasks = Enumerable.Range(0, 6).Select(CreateOneAsync).ToArray();
            var files = await Task.WhenAll(tasks);

            foreach (var file in files) {
                using var ss = SpreadsheetDocument.Open(file, false);
                var ws = ss.WorkbookPart!.WorksheetParts.First();
                Assert.Single(ws.TableDefinitionParts);
            }
        }
    }
}

