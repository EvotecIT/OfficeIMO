using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2010.Excel;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_WorksheetCommentsLifecycle() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorksheetComments.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Comments");
                sheet.CellValue(1, 1, "Header");
                sheet.SetComment(1, 1, "First comment", author: "Tester", initials: "TT");
                Assert.True(sheet.HasComment(1, 1));
                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets.First();
                Assert.True(sheet.HasComment(1, 1));
                sheet.ClearComment("A1");
                Assert.False(sheet.HasComment(1, 1));
                document.Save(false);
            }
        }

        [Fact]
        public void Test_WorksheetProtectionOptions() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorksheetProtection.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Protected");
                var options = new ExcelSheetProtectionOptions {
                    AllowSelectLockedCells = false,
                    AllowSelectUnlockedCells = false,
                    AllowSort = true,
                    AllowAutoFilter = true,
                    AllowInsertRows = true
                };
                sheet.Protect(options);
                Assert.True(sheet.IsProtected);
                document.Save(false);
            }

            using (var doc = SpreadsheetDocument.Open(filePath, false)) {
                var wb = doc.WorkbookPart!;
                var sheet = wb.Workbook.Sheets!.OfType<Sheet>().First(s => s.Name == "Protected");
                var wsPart = (WorksheetPart)wb.GetPartById(sheet.Id!);
                var protection = wsPart.Worksheet.Elements<SheetProtection>().FirstOrDefault();
                Assert.NotNull(protection);
                Assert.True(protection!.Sort?.Value ?? false);
                Assert.True(protection.AutoFilter?.Value ?? false);
                Assert.True(protection.InsertRows?.Value ?? false);
                Assert.False(protection.SelectLockedCells?.Value ?? true);
                Assert.False(protection.SelectUnlockedCells?.Value ?? true);
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets.First();
                sheet.Unprotect();
                Assert.False(sheet.IsProtected);
                document.Save(false);
            }

            using (var doc = SpreadsheetDocument.Open(filePath, false)) {
                var wb = doc.WorkbookPart!;
                var sheet = wb.Workbook.Sheets!.OfType<Sheet>().First(s => s.Name == "Protected");
                var wsPart = (WorksheetPart)wb.GetPartById(sheet.Id!);
                Assert.False(wsPart.Worksheet.Elements<SheetProtection>().Any());
            }
        }

        [Fact]
        public void Test_WorksheetSparklines() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorksheetSparklines.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sparklines");
                for (int i = 0; i < 6; i++) {
                    sheet.CellValue(2, 2 + i, (double)(i + 1));
                }

                sheet.AddSparklines("B2:G2", "H2:H2", displayMarkers: true, seriesColor: "#FF0000");
                document.Save(false);
            }

            using (var doc = SpreadsheetDocument.Open(filePath, false)) {
                var wb = doc.WorkbookPart!;
                var sheet = wb.Workbook.Sheets!.OfType<Sheet>().First(s => s.Name == "Sparklines");
                var wsPart = (WorksheetPart)wb.GetPartById(sheet.Id!);
                var groups = wsPart.Worksheet.Descendants<SparklineGroups>().FirstOrDefault();
                if (groups != null) {
                    var group = groups.Elements<SparklineGroup>().FirstOrDefault();
                    Assert.NotNull(group);
                    Assert.Equal(SparklineTypeValues.Line, group!.Type?.Value);
                } else {
                    const string SparklineNamespace = "http://schemas.microsoft.com/office/excel/2009/9/main";
                    var unknown = wsPart.Worksheet.Descendants<OpenXmlUnknownElement>()
                        .OfType<OpenXmlUnknownElement>()
                        .FirstOrDefault(e => e.LocalName == "sparklineGroups" && e.NamespaceUri == SparklineNamespace);
                    Assert.NotNull(unknown);
                }
            }
        }
    }
}
