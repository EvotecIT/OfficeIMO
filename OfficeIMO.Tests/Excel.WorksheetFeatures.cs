using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2010.Excel;
using OfficeIMO.Excel;
using OfficeFormula = DocumentFormat.OpenXml.Office.Excel.Formula;
using OfficeReferenceSequence = DocumentFormat.OpenXml.Office.Excel.ReferenceSequence;
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

            using (var package = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                Assert.Null(wsPart.WorksheetCommentsPart);
                Assert.Empty(wsPart.VmlDrawingParts);
                Assert.Null(wsPart.Worksheet.Elements<LegacyDrawing>().FirstOrDefault());
            }
        }

        [Fact]
        public void Test_WorksheetComments_Multiline_OpenXmlValidation() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorksheetCommentsMultiline.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Comments");
                sheet.CellValue(1, 1, "Header");
                sheet.SetComment(1, 1, "Line 1\nLine 2", author: "Tester", initials: "TT");
                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
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

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
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
                    var sparkline = group.GetFirstChild<Sparklines>()!.Elements<Sparkline>().Single();
                    Assert.Equal("B2:G2", sparkline.GetFirstChild<OfficeFormula>()!.Text);
                    Assert.Equal("H2:H2", sparkline.GetFirstChild<OfficeReferenceSequence>()!.Text);
                } else {
                    const string SparklineNamespace = "http://schemas.microsoft.com/office/excel/2009/9/main";
                    var unknown = wsPart.Worksheet.Descendants<OpenXmlUnknownElement>()
                        .OfType<OpenXmlUnknownElement>()
                        .FirstOrDefault(e => e.LocalName == "sparklineGroups" && e.NamespaceUri == SparklineNamespace);
                    Assert.NotNull(unknown);
                }
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_WorksheetSparklines_MultiRowRangesExpandPerTargetCell() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorksheetSparklines.MultiRow.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sparklines");
                for (int row = 2; row <= 4; row++) {
                    sheet.CellValue(row, 3, (double)row);
                    sheet.CellValue(row, 4, (double)(row + 10));
                }

                sheet.AddSparklines("C2:D4", "E2:E4", displayMarkers: true, seriesColor: "#FF0000");
                document.Save(false);
            }

            using (var doc = SpreadsheetDocument.Open(filePath, false)) {
                var wb = doc.WorkbookPart!;
                var sheet = wb.Workbook.Sheets!.OfType<Sheet>().First(s => s.Name == "Sparklines");
                var wsPart = (WorksheetPart)wb.GetPartById(sheet.Id!);
                var group = wsPart.Worksheet.Descendants<SparklineGroup>().Single();
                var sparklines = group.GetFirstChild<Sparklines>()!.Elements<Sparkline>().ToList();

                Assert.Equal(3, sparklines.Count);
                Assert.Equal(new[] { "C2:D2", "C3:D3", "C4:D4" },
                    sparklines.Select(s => s.GetFirstChild<OfficeFormula>()!.Text).ToArray());
                Assert.Equal(new[] { "E2", "E3", "E4" },
                    sparklines.Select(s => s.GetFirstChild<OfficeReferenceSequence>()!.Text).ToArray());
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_WorksheetSparklines_MultiColumnRangesExpandPerTargetCell() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorksheetSparklines.MultiColumn.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sparklines");
                for (int row = 2; row <= 4; row++) {
                    sheet.CellValue(row, 3, (double)row);
                    sheet.CellValue(row, 4, (double)(row + 10));
                }

                sheet.AddSparklines("C2:D4", "C5:D5", displayMarkers: true, seriesColor: "#FF0000");
                document.Save(false);
            }

            using (var doc = SpreadsheetDocument.Open(filePath, false)) {
                var wb = doc.WorkbookPart!;
                var sheet = wb.Workbook.Sheets!.OfType<Sheet>().First(s => s.Name == "Sparklines");
                var wsPart = (WorksheetPart)wb.GetPartById(sheet.Id!);
                var group = wsPart.Worksheet.Descendants<SparklineGroup>().Single();
                var sparklines = group.GetFirstChild<Sparklines>()!.Elements<Sparkline>().ToList();

                Assert.Equal(2, sparklines.Count);
                Assert.Equal(new[] { "C2:C4", "D2:D4" },
                    sparklines.Select(s => s.GetFirstChild<OfficeFormula>()!.Text).ToArray());
                Assert.Equal(new[] { "C5", "D5" },
                    sparklines.Select(s => s.GetFirstChild<OfficeReferenceSequence>()!.Text).ToArray());
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_WorksheetSparklines_AmbiguousMultiCellLocationThrows() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorksheetSparklines.InvalidShape.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sparklines");

                var exception = Assert.Throws<ArgumentException>(() => sheet.AddSparklines("C2:D4", "E2:F4"));
                Assert.Equal("locationRange", exception.ParamName);
            }
        }

        [Fact]
        public void Test_WorksheetGridlinesPersistence() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorksheetGridlines.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Gridlines");
                sheet.CellValue(1, 1, "Header");
                sheet.SetGridlinesVisible(false);
                document.Save(false);
            }

            using (var doc = SpreadsheetDocument.Open(filePath, false)) {
                var wb = doc.WorkbookPart!;
                var sheet = wb.Workbook.Sheets!.OfType<Sheet>().First(s => s.Name == "Gridlines");
                var wsPart = (WorksheetPart)wb.GetPartById(sheet.Id!);
                var view = wsPart.Worksheet.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>();
                Assert.NotNull(view);
                Assert.False(view!.ShowGridLines?.Value ?? true);
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets.First();
                sheet.SetGridlinesVisible(true);
                document.Save(false);
            }

            using (var doc = SpreadsheetDocument.Open(filePath, false)) {
                var wb = doc.WorkbookPart!;
                var sheet = wb.Workbook.Sheets!.OfType<Sheet>().First(s => s.Name == "Gridlines");
                var wsPart = (WorksheetPart)wb.GetPartById(sheet.Id!);
                var view = wsPart.Worksheet.GetFirstChild<SheetViews>()?.GetFirstChild<SheetView>();
                Assert.NotNull(view);
                Assert.True(view!.ShowGridLines?.Value ?? true);
            }
        }
    }
}
