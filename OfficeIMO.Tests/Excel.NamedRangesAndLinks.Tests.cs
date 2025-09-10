using System;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelNamedRangesAndLinksTests {
        [Fact]
        public void NamedRange_SanitizeClampsOutOfBounds() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var doc = ExcelDocument.Create(path)) {
                var data = doc.AddWorkSheet("Data");
                doc.SetNamedRange("TooBig", "'Data'!A1:B2000000", save: false, hidden: false, validationMode: NameValidationMode.Sanitize);
                var value = doc.GetNamedRange("TooBig");
                Assert.Equal("'Data'!$A$1:$B$1048576", value);
            }
            File.Delete(path);
        }

        [Fact]
        public void NamedRange_StrictThrowsOutOfBounds() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var doc = ExcelDocument.Create(path)) {
                doc.AddWorkSheet("Data");
                Assert.Throws<ArgumentOutOfRangeException>(() =>
                    doc.SetNamedRange("TooBigS", "'Data'!A1:B2000000", save: false, hidden: false, validationMode: NameValidationMode.Strict));
            }
            File.Delete(path);
        }

        [Fact]
        public void InternalLink_QuotingForApostropheSheetName() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var doc = ExcelDocument.Create(path)) {
                var main = doc.AddWorkSheet("Main");
                var target = doc.AddWorkSheet("O'Brien");
                main.SetInternalLink(1, 1, target, "A1", display: "Go");

                // Inspect hyperlinks
                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(main)!;
                var links = wsPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault();
                Assert.NotNull(links);
                var hl = links!.Elements<Hyperlink>().FirstOrDefault();
                Assert.NotNull(hl);
                Assert.Equal("'O''Brien'!A1", hl!.Location!.Value);
            }
            File.Delete(path);
        }

        [Fact]
        public void Preflight_KeepsValidDrawingAndTable() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var doc = ExcelDocument.Create(path)) {
                var s = doc.AddWorkSheet("Assets");

                // Headers and one data row
                s.CellValue(1, 1, "Col1"); s.CellValue(1, 2, "Col2");
                s.CellValue(2, 1, 1); s.CellValue(2, 2, 2);
                s.AddTable("A1:B2", hasHeader: true, name: "T", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);

                // Add a tiny 1x1 PNG as drawing
                var png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");
                s.AddImageAt(1, 1, png, "image/png", widthPixels: 1, heightPixels: 1);

                // Preflight should not remove valid references
                doc.PreflightWorkbook();

                // Assert drawing exists
                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(s)!;
                var drawing = wsPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>();
                Assert.NotNull(drawing);

                // Assert tableparts exist
                var parts = wsPart.Worksheet.Elements<TableParts>().FirstOrDefault();
                Assert.NotNull(parts);
                Assert.True(parts!.Elements<TablePart>().Any());
            }
            File.Delete(path);
        }
    }
}
