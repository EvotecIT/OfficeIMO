using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void TableNameComparer_CanBeChanged_AndResetsCache() {
            string filePath = Path.Combine(_directoryWithFiles, "Table.NameComparer.xlsx");
            using (var doc = ExcelDocument.Create(filePath)) {
                // Case-sensitive comparer first: allow "Table" and "TABLE" side by side
                doc.TableNameComparer = System.StringComparer.Ordinal; // case-sensitive
                var s = doc.AddWorkSheet("Data");
                s.CellValue(1, 1, "A");
                s.AddTable("A1:A1", hasHeader: true, name: "Table", TableStyle.TableStyleMedium9);

                // Switch to case-insensitive and ensure suffixing now applies for the same base name
                doc.TableNameComparer = System.StringComparer.OrdinalIgnoreCase; // resets cache
                s.CellValue(2, 1, "B");
                s.AddTable("A2:A2", hasHeader: true, name: "TABLE", TableStyle.TableStyleMedium9);
                doc.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var ws = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var names = ws.TableDefinitionParts.Select(tp => tp.Table!.Name!.Value).ToArray();
                Assert.Contains("Table", names);
                Assert.Contains("TABLE2", names); // case-insensitive comparer forces uniqueness suffix
            }
        }
    }
}

