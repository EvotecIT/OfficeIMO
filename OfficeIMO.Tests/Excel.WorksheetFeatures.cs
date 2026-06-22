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
using Threaded = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
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
        public void Test_WorksheetComments_RichTextAuthoringAndUpdate() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorksheetComments.RichText.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Comments");
                sheet.SetCommentRichText("A1", new[] {
                    new ExcelRichTextRun("Important ") { Bold = true, FontColor = "#FF0000" },
                    new ExcelRichTextRun("note") { Italic = true, Underline = true, FontName = "Arial", FontSize = 14D }
                }, author: "Alice", initials: "AA");

                var comment = Assert.Single(sheet.GetComments());
                Assert.Equal("Important note", comment.Text);
                Assert.Equal(2, comment.RichTextRuns.Count);
                Assert.True(comment.RichTextRuns[0].Bold);
                Assert.Equal("FFFF0000", comment.RichTextRuns[0].FontColor);

                int updated = sheet.UpdateCommentsRichText(
                    new ExcelCommentFilter { Author = "Alice (AA)", TextContains = "note" },
                    new[] {
                        new ExcelRichTextRun("Reviewed") { Bold = true },
                        new ExcelRichTextRun(" item") { FontColor = "0563C1" }
                    },
                    author: "Bob",
                    initials: "BB");
                Assert.Equal(1, updated);
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var comment = spreadsheet.WorkbookPart!.WorksheetParts.Single().WorksheetCommentsPart!.Comments!.CommentList!.Elements<Comment>().Single();
                var runs = comment.CommentText!.Elements<Run>().ToList();

                Assert.Equal("A1", comment.Reference!.Value);
                Assert.Equal(2, runs.Count);
                Assert.Equal("Reviewed", runs[0].Text!.Text);
                Assert.NotNull(runs[0].RunProperties!.GetFirstChild<Bold>());
                Assert.Equal(" item", runs[1].Text!.Text);
                Assert.Equal("FF0563C1", runs[1].RunProperties!.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>()!.Rgb!.Value);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                var comment = Assert.Single(document.Sheets.First().GetComments());
                Assert.Equal("Bob (BB)", comment.Author);
                Assert.Equal("Reviewed item", comment.Text);
                Assert.Equal(2, comment.RichTextRuns.Count);
                Assert.True(comment.RichTextRuns[0].Bold);
                Assert.Equal("FF0563C1", comment.RichTextRuns[1].FontColor);
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_WorksheetThreadedComments_InspectAndPreserve() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorksheetThreadedComments.xlsx");
            const string personId = "{11111111-1111-1111-1111-111111111111}";
            const string commentId = "{22222222-2222-2222-2222-222222222222}";

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Threaded");
                sheet.CellValue(1, 1, "Revenue");
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                var personPart = workbookPart.AddNewPart<WorkbookPersonPart>();
                personPart.PersonList = new Threaded.PersonList(
                    new Threaded.Person {
                        DisplayName = "Modern Reviewer",
                        Id = personId,
                        UserId = "modern.reviewer@example.test",
                        ProviderId = "OfficeIMO.Tests"
                    });
                personPart.PersonList.Save();

                var worksheetPart = workbookPart.WorksheetParts.Single();
                var threadedPart = worksheetPart.AddNewPart<WorksheetThreadedCommentsPart>();
                threadedPart.ThreadedComments = new Threaded.ThreadedComments(
                    new Threaded.ThreadedComment(
                        new Threaded.ThreadedCommentText("Discuss revenue"))
                    {
                        Ref = "A1",
                        PersonId = personId,
                        Id = commentId,
                        DT = new DateTime(2026, 5, 22, 10, 0, 0, DateTimeKind.Utc),
                        Done = false
                    });
                threadedPart.ThreadedComments.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var snapshot = document.CreateInspectionSnapshot();
                var worksheet = Assert.Single(snapshot.Worksheets);
                var threadedComment = Assert.Single(worksheet.ThreadedComments);

                Assert.Equal("A1", threadedComment.CellReference);
                Assert.Equal(commentId, threadedComment.Id);
                Assert.Equal(personId, threadedComment.PersonId);
                Assert.Equal("Modern Reviewer", threadedComment.Author);
                Assert.Equal("Discuss revenue", threadedComment.Text);
                Assert.False(threadedComment.Done);

                var cell = Assert.Single(worksheet.Cells, c => c.Row == 1 && c.Column == 1);
                Assert.Same(threadedComment, cell.ThreadedComment);

                var feature = Assert.Single(document.InspectFeatures().FindFeatures("Threaded comments"));
                Assert.Equal(ExcelFeatureSupportLevel.Preserved, feature.SupportLevel);
                Assert.Equal(1, feature.Count);
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
                var threadedPart = Assert.Single(worksheetPart.WorksheetThreadedCommentsParts);
                var threadedComment = Assert.Single(threadedPart.ThreadedComments!.Elements<Threaded.ThreadedComment>());

                Assert.Equal("A1", threadedComment.Ref!.Value);
                Assert.Equal(commentId, threadedComment.Id!.Value);
                Assert.Equal("Discuss revenue", threadedComment.ThreadedCommentText!.InnerText);
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

        [Fact]
        public void Test_WorksheetTabColorSetClearAndInspect() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorksheetTabColor.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Colored");
                sheet.CellValue(1, 1, "Header");
                sheet.SetTabColor("#336699");
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                TabColor tabColor = wsPart.Worksheet.GetFirstChild<SheetProperties>()!.TabColor!;
                Assert.Equal("FF336699", tabColor.Rgb!.Value);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelWorksheetSnapshot worksheet = Assert.Single(document.CreateInspectionSnapshot().Worksheets);
                Assert.Equal("FF336699", worksheet.TabColorArgb);
            }

            using (var document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.First();
                sheet.ClearTabColor();
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Assert.Null(wsPart.Worksheet.GetFirstChild<SheetProperties>()?.TabColor);
            }
        }

        [Fact]
        public void Test_WorksheetViewInfoReadsFreezeAndGridlines() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorksheetViewInfo.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("View");
                sheet.CellValue(1, 1, "Header");
                sheet.Freeze(2, 1);
                sheet.SetGridlinesVisible(false);
                document.Save(false);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                var sheet = document.Sheets.First();
                ExcelWorksheetViewInfo view = sheet.GetViewInfo();

                Assert.True(view.HasPane);
                Assert.Equal("frozen", view.PaneState);
                Assert.Equal(2, view.FrozenRowCount);
                Assert.Equal(1, view.FrozenColumnCount);
                Assert.Equal(2D, view.VerticalSplit);
                Assert.Equal(1D, view.HorizontalSplit);
                Assert.Equal("B3", view.TopLeftCell);
                Assert.Equal("bottomRight", view.ActivePane);
                Assert.False(view.ShowGridlines);
                Assert.False(view.RightToLeft);
            }
        }

        [Fact]
        public void Test_WorksheetViewOptionsSetZoomDirectionGridlinesAndMode() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorksheetViewOptions.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("View");
                sheet.CellValue(1, 1, "Header");
                sheet.SetViewOptions(showGridlines: false, rightToLeft: true, zoomScale: 125, zoomScaleNormal: 100, view: ExcelWorksheetViewKind.PageLayout);
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                SheetView sheetView = wsPart.Worksheet.GetFirstChild<SheetViews>()!.GetFirstChild<SheetView>()!;
                Assert.False(sheetView.ShowGridLines!.Value);
                Assert.True(sheetView.RightToLeft!.Value);
                Assert.Equal(125U, sheetView.ZoomScale!.Value);
                Assert.Equal(100U, sheetView.ZoomScaleNormal!.Value);
                Assert.Equal(SheetViewValues.PageLayout, sheetView.View!.Value);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                var sheet = document.Sheets.First();
                ExcelWorksheetViewInfo view = sheet.GetViewInfo();
                Assert.False(view.ShowGridlines);
                Assert.True(view.RightToLeft);
                Assert.Equal(125U, view.ZoomScale);
                Assert.Equal(100U, view.ZoomScaleNormal);
                Assert.Equal("pageLayout", view.View);

                ExcelWorksheetSnapshot worksheet = Assert.Single(document.CreateInspectionSnapshot().Worksheets);
                Assert.False(worksheet.ShowGridlines);
                Assert.True(worksheet.RightToLeft);
                Assert.Equal(125U, worksheet.ZoomScale);
                Assert.Equal(100U, worksheet.ZoomScaleNormal);
                Assert.Equal("pageLayout", worksheet.View);
            }
        }

        [Fact]
        public void Test_WorkbookActiveWorksheetSetAndInspect() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorkbookActiveWorksheet.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var summary = document.AddWorkSheet("Summary");
                summary.CellValue(1, 1, "Summary");
                var details = document.AddWorkSheet("Details");
                details.CellValue(1, 1, "Details");
                var archive = document.AddWorkSheet("Archive");
                archive.CellValue(1, 1, "Archive");
                document.SetActiveWorksheet(details);
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                Workbook workbook = spreadsheet.WorkbookPart!.Workbook;
                WorkbookView workbookView = workbook.GetFirstChild<BookViews>()!.GetFirstChild<WorkbookView>()!;
                Assert.Equal(1U, workbookView.ActiveTab!.Value);
                Assert.Equal(1U, workbookView.FirstSheet!.Value);

                var sheets = workbook.Sheets!.Elements<Sheet>().ToArray();
                for (int index = 0; index < sheets.Length; index++) {
                    var worksheetPart = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(sheets[index].Id!);
                    SheetView sheetView = worksheetPart.Worksheet.GetFirstChild<SheetViews>()!.GetFirstChild<SheetView>()!;
                    Assert.Equal(index == 1, sheetView.TabSelected!.Value);
                }
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot();
                Assert.Equal(1, snapshot.ActiveWorksheetIndex);
                Assert.Equal("Details", snapshot.ActiveWorksheetName);
                Assert.False(snapshot.Worksheets[0].IsActive);
                Assert.True(snapshot.Worksheets[1].IsActive);
                Assert.False(snapshot.Worksheets[2].IsActive);
            }
        }

        [Fact]
        public void Test_WorksheetOutlineGroupsRowsAndColumnsWithSnapshotReadback() {
            var filePath = Path.Combine(_directoryWithFiles, "ExcelWorksheetOutlineGroups.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Outline");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(2, 1, "North");
                sheet.CellValue(3, 1, "South");
                sheet.CellValue(1, 2, "Q1");
                sheet.CellValue(1, 3, "Q2");

                sheet.SetOutlineSummary(summaryBelow: true, summaryRight: true);
                sheet.GroupRows(2, 3, outlineLevel: 1, collapsed: true);
                sheet.GroupColumns(2, 3, outlineLevel: 1, collapsed: true);
                Assert.Throws<ArgumentOutOfRangeException>(() => sheet.GroupRows(A1.MaxRows, A1.MaxRows, outlineLevel: 1, collapsed: true));
                Assert.Throws<ArgumentOutOfRangeException>(() => sheet.GroupRows(A1.MaxRows + 1, A1.MaxRows + 1));
                Assert.Throws<ArgumentOutOfRangeException>(() => sheet.GroupColumns(A1.MaxColumns + 1, A1.MaxColumns + 1));
                document.Save(false);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Worksheet worksheet = wsPart.Worksheet;
                OutlineProperties outline = worksheet.GetFirstChild<SheetProperties>()!.GetFirstChild<OutlineProperties>()!;
                Assert.True(outline.SummaryBelow!.Value);
                Assert.True(outline.SummaryRight!.Value);

                var rows = worksheet.GetFirstChild<SheetData>()!.Elements<Row>().ToDictionary(row => row.RowIndex!.Value);
                Assert.Equal(1, rows[2].OutlineLevel!.Value);
                Assert.Equal(1, rows[3].OutlineLevel!.Value);
                Assert.True(rows[2].Hidden!.Value);
                Assert.True(rows[3].Hidden!.Value);
                Assert.True(rows[4].Collapsed!.Value);

                var columns = worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToList();
                Column columnB = Assert.Single(columns, column => column.Min!.Value == 2 && column.Max!.Value == 2);
                Column columnC = Assert.Single(columns, column => column.Min!.Value == 3 && column.Max!.Value == 3);
                Column columnD = Assert.Single(columns, column => column.Min!.Value == 4 && column.Max!.Value == 4);
                Assert.Equal(1, columnB.OutlineLevel!.Value);
                Assert.Equal(1, columnC.OutlineLevel!.Value);
                Assert.True(columnB.Hidden!.Value);
                Assert.True(columnC.Hidden!.Value);
                Assert.True(columnD.Collapsed!.Value);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelWorksheetSnapshot worksheet = Assert.Single(document.CreateInspectionSnapshot().Worksheets);
                Assert.True(worksheet.OutlineSummaryBelow);
                Assert.True(worksheet.OutlineSummaryRight);

                ExcelRowSnapshot row2 = Assert.Single(worksheet.Rows, row => row.Index == 2);
                ExcelRowSnapshot row4 = Assert.Single(worksheet.Rows, row => row.Index == 4);
                Assert.Equal((byte)1, row2.OutlineLevel.GetValueOrDefault());
                Assert.True(row2.Hidden);
                Assert.True(row4.Collapsed);

                ExcelColumnSnapshot column2 = Assert.Single(worksheet.Columns, column => column.StartIndex == 2 && column.EndIndex == 2);
                ExcelColumnSnapshot column4 = Assert.Single(worksheet.Columns, column => column.StartIndex == 4 && column.EndIndex == 4);
                Assert.Equal((byte)1, column2.OutlineLevel.GetValueOrDefault());
                Assert.True(column2.Hidden);
                Assert.True(column4.Collapsed);
                Assert.Empty(document.ValidateOpenXml());
            }
        }
    }
}
