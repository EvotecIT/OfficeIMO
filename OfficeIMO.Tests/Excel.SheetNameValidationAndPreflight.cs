using System;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelSheetNameValidationAndPreflightTests {
        [Fact]
        public void AddWorkSheet_Default_SanitizesInvalidCharsAndDuplicates() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var doc = ExcelDocument.Create(path)) {
                var first = doc.AddWorkSheet("Q4:Revenue/Forecast?*");
                var second = doc.AddWorkSheet("Q4:Revenue/Forecast?*");

                Assert.Equal("Q4_Revenue_Forecast", first.Name);
                Assert.Equal("Q4_Revenue_Forecast (2)", second.Name);
            }
            File.Delete(path);
        }

        [Fact]
        public void AddWorkSheet_Default_BlankNamesUseUniqueExcelStyleSheetNames() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var doc = ExcelDocument.Create(path)) {
                var first = doc.AddWorkSheet();
                var second = doc.AddWorkSheet("   ");
                var third = doc.AddWorkSheet("???");

                Assert.Equal("Sheet1", first.Name);
                Assert.Equal("Sheet2", second.Name);
                Assert.Equal("Sheet3", third.Name);
            }
            File.Delete(path);
        }

        [Fact]
        public void AddWorkSheet_Sanitize_InvalidCharsAndDuplicate() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var doc = ExcelDocument.Create(path)) {
                var s1 = doc.AddWorkSheet("Q4:Revenue/Forecast?*", SheetNameValidationMode.Sanitize);
                // invalid characters replaced, trimmed; consecutive underscores collapsed by sanitizer
                Assert.Equal("Q4_Revenue_Forecast", s1.Name.Trim());

                var s2 = doc.AddWorkSheet("Q4:Revenue/Forecast?*", SheetNameValidationMode.Sanitize);
                Assert.NotEqual(s1.Name, s2.Name);
                Assert.EndsWith("(2)", s2.Name);
            }
            File.Delete(path);
        }

        [Fact]
        public void AddWorkSheet_Strict_ThrowsOnInvalidOrDuplicate() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var doc = ExcelDocument.Create(path)) {
                // Invalid chars
                Assert.Throws<ArgumentException>(() => doc.AddWorkSheet("Bad:Name", SheetNameValidationMode.Strict));

                // Valid then duplicate
                var s1 = doc.AddWorkSheet("Data", SheetNameValidationMode.Strict);
                Assert.NotNull(s1);
                Assert.Throws<ArgumentException>(() => doc.AddWorkSheet("Data", SheetNameValidationMode.Strict));
            }
            File.Delete(path);
        }

        [Fact]
        public void RenameWorkSheet_StrictSetter_ThrowsOnInvalidOrDuplicate() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var doc = ExcelDocument.Create(path)) {
                var alpha = doc.AddWorkSheet("Alpha", SheetNameValidationMode.Strict);
                var beta = doc.AddWorkSheet("Beta", SheetNameValidationMode.Strict);

                Assert.Throws<ArgumentException>(() => alpha.Name = "Bad:Name");
                Assert.Throws<ArgumentException>(() => beta.Name = "Alpha");

                alpha.Name = "alpha";
                Assert.Equal("alpha", alpha.Name);
            }
            File.Delete(path);
        }

        [Fact]
        public void Preflight_RemovesEmptyAndOrphanedWorksheetElements() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Preflight");

                // Use reflection to access internal WorksheetPart and Worksheet to simulate problematic structures
                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var ws = wsPart.Worksheet;

                // 1) Empty Hyperlinks
                ws.AppendChild(new Hyperlinks());

                // 2) Empty MergeCells
                ws.AppendChild(new MergeCells());

                // 3) Orphaned Drawing ref
                ws.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = "rId999" });

                // 4) Orphaned LegacyDrawingHeaderFooter ref
                ws.AppendChild(new LegacyDrawingHeaderFooter { Id = "rId998" });

                // 5) TableParts with invalid/duplicate ids
                var parts = ws.Elements<TableParts>().FirstOrDefault();
                if (parts == null) { parts = new TableParts(); ws.Append(parts); }
                parts.Append(new TablePart { Id = "rId100" }); // invalid
                parts.Append(new TablePart { Id = "rId100" }); // duplicate
                parts.Count = (uint)parts.Elements<TablePart>().Count();

                ws.Save();

                // Run preflight via public API
                doc.PreflightWorkbook();

                // Re-fetch elements
                ws = wsPart.Worksheet;
                Assert.Null(ws.Elements<Hyperlinks>().FirstOrDefault());
                Assert.Null(ws.Elements<MergeCells>().FirstOrDefault());
                Assert.Null(ws.Elements<DocumentFormat.OpenXml.Spreadsheet.Drawing>().FirstOrDefault());
                Assert.Null(ws.Elements<LegacyDrawingHeaderFooter>().FirstOrDefault());

                var partsAfter = ws.Elements<TableParts>().FirstOrDefault();
                Assert.Null(partsAfter); // all invalid/duplicate removed → container dropped
            }
            File.Delete(path);
        }

        [Fact]
        public void Preflight_RemovesEmptyValidationContainersBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Preflight");

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var ws = wsPart.Worksheet;

                var dataValidations = new DataValidations();
                dataValidations.SetAttribute(new OpenXmlAttribute("count", string.Empty, "5"));
                ws.AppendChild(dataValidations);

                var ignoredErrors = new IgnoredErrors();
                ignoredErrors.SetAttribute(new OpenXmlAttribute("count", string.Empty, "3"));
                ws.AppendChild(ignoredErrors);
                ws.AppendChild(new CustomSheetViews());
                ws.AppendChild(new ConditionalFormatting());

                ws.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                var ws = wsPart.Worksheet;

                Assert.Null(ws.Elements<DataValidations>().FirstOrDefault());
                Assert.Null(ws.Elements<IgnoredErrors>().FirstOrDefault());
                Assert.Empty(ws.Elements<CustomSheetViews>());
                Assert.Empty(ws.Elements<ConditionalFormatting>());
            }

            File.Delete(path);
            File.Delete(savePath);
        }
    }
}

