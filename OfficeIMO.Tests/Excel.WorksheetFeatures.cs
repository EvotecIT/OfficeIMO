using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Office2010.Excel;
using OfficeIMO.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeFormula = DocumentFormat.OpenXml.Office.Excel.Formula;
using OfficeReferenceSequence = DocumentFormat.OpenXml.Office.Excel.ReferenceSequence;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
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
        public void Test_WorksheetComments_CanFindUpdateAndRemoveByFilter() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorksheetComments.Filtered.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Comments");
                sheet.SetComment("A1", "Review total", author: "Alice", initials: "AA");
                sheet.SetComment("B2", "Review status", author: "Alice", initials: "AA");
                sheet.SetComment("C3", "Keep status", author: "Bob", initials: "BB");

                var aliceComments = sheet.FindComments(new ExcelCommentFilter { Author = "Alice (AA)" });
                Assert.Equal(2, aliceComments.Count);
                Assert.Contains(aliceComments, comment => comment.CellReference == "A1" && comment.Text == "Review total");

                int updated = sheet.UpdateComments(
                    new ExcelCommentFilter { Author = "Alice (AA)", TextContains = "status", A1Range = "A1:B2" },
                    "Status reviewed",
                    author: "Carol",
                    initials: "CC");
                Assert.Equal(1, updated);

                var carolComment = Assert.Single(sheet.FindComments(new ExcelCommentFilter { Author = "Carol (CC)" }));
                Assert.Equal("B2", carolComment.CellReference);
                Assert.Equal("Status reviewed", carolComment.Text);

                int removed = sheet.ClearComments(new ExcelCommentFilter { TextContains = "Review", A1Range = "A1:A1" });
                Assert.Equal(1, removed);
                Assert.False(sheet.HasComment(1, 1));
                Assert.True(sheet.HasComment(2, 2));
                Assert.True(sheet.HasComment(3, 3));
                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                var sheet = document.Sheets.First();
                var comments = sheet.GetComments().OrderBy(comment => comment.CellReference).ToList();

                Assert.Equal(2, comments.Count);
                Assert.Equal("B2", comments[0].CellReference);
                Assert.Equal("Carol (CC)", comments[0].Author);
                Assert.Equal("Status reviewed", comments[0].Text);
                Assert.Equal("C3", comments[1].CellReference);
                Assert.Equal("Bob (BB)", comments[1].Author);
                Assert.Equal("Keep status", comments[1].Text);
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_WorksheetImages_MetadataAndSizing() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorksheetImages.Metadata.xlsx");
            var png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Images");

                var image = sheet.AddImage(1, 1, png, "image/png", widthPixels: 12, heightPixels: 10,
                    name: "Logo", altText: "Company logo");
                image.SetAltText("Updated company logo", "Logo title")
                    .LockAspectRatio(false)
                    .SetSize(24, 18);

                Assert.Equal("Logo", sheet.GetImage("Logo")?.Name);
                Assert.Single(sheet.Images);
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var picture = wsPart.DrawingsPart!.WorksheetDrawing!.Descendants<Xdr.Picture>().Single();
                var drawingProperties = picture.NonVisualPictureProperties!.NonVisualDrawingProperties!;
                var locks = picture.NonVisualPictureProperties.NonVisualPictureDrawingProperties!.GetFirstChild<A.PictureLocks>()!;
                var extent = wsPart.DrawingsPart.WorksheetDrawing!.Descendants<Xdr.Extent>().First();

                Assert.Equal("Logo", drawingProperties.Name!.Value);
                Assert.Equal("Updated company logo", drawingProperties.Description!.Value);
                Assert.Equal("Logo title", drawingProperties.Title!.Value);
                Assert.False(locks.NoChangeAspect!.Value);
                Assert.Equal(24L * 9525L, extent.Cx!.Value);
                Assert.Equal(18L * 9525L, extent.Cy!.Value);
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
                Assert.False(protection!.Sort?.Value ?? true);
                Assert.False(protection.AutoFilter?.Value ?? true);
                Assert.False(protection.InsertRows?.Value ?? true);
                Assert.True(protection.SelectLockedCells?.Value ?? false);
                Assert.True(protection.SelectUnlockedCells?.Value ?? false);
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
        public void Test_WorksheetSparklines_FluentBuilder() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorksheetSparklines.Fluent.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sparklines");
                for (int i = 0; i < 6; i++) {
                    sheet.CellValue(2, 2 + i, (double)(i + 1));
                }

                sheet.Sparklines("B2:G2")
                    .Column()
                    .Markers()
                    .HighLow()
                    .Axis()
                    .Color("#4472C4")
                    .At("H2:H2");
                document.Save(false);
            }

            using (var doc = SpreadsheetDocument.Open(filePath, false)) {
                var wb = doc.WorkbookPart!;
                var sheet = wb.Workbook.Sheets!.OfType<Sheet>().First(s => s.Name == "Sparklines");
                var wsPart = (WorksheetPart)wb.GetPartById(sheet.Id!);
                var group = wsPart.Worksheet.Descendants<SparklineGroup>().Single();
                Assert.Equal(SparklineTypeValues.Column, group.Type?.Value);
                Assert.True(group.Markers?.Value);
                Assert.True(group.High?.Value);
                Assert.True(group.Low?.Value);
                Assert.True(group.DisplayXAxis?.Value);
                Assert.Equal("FF4472C4", group.SeriesColor!.Rgb!.Value);

                var sparkline = group.GetFirstChild<Sparklines>()!.Elements<Sparkline>().Single();
                Assert.Equal("B2:G2", sparkline.GetFirstChild<OfficeFormula>()!.Text);
                Assert.Equal("H2:H2", sparkline.GetFirstChild<OfficeReferenceSequence>()!.Text);
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
