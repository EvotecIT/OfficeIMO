using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_NamedRange_DefaultSave_DoesNotForceStreamPackageSave() {
            using var source = new MemoryStream();
            using var destination = new MemoryStream();

            using (var document = ExcelDocument.Create(source, autoSave: false)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Ready");

                document.SetNamedRange("Anchor", "A1", sheet);
                document.Save(destination);
            }

            destination.Position = 0;
            using var loaded = ExcelDocument.Load(destination, readOnly: true);
            Assert.Equal("$A$1", loaded.GetNamedRange("Anchor", loaded["Data"]));
        }

        [Fact]
        public void Test_DelimitedImport_DetectsDelimiterIgnoringQuotedFields() {
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            ExcelDelimitedImportResult result = document.ImportDelimitedText("\"Name, with comma\";Value\r\nAlpha;42", new ExcelDelimitedImportOptions {
                SheetName = "Import"
            });

            Assert.Equal(';', result.Delimiter);
            ExcelSheet sheet = document["Import"];
            Assert.True(sheet.TryGetCellText(2, 1, out string? name));
            Assert.Equal("Alpha", name);
            Assert.True(sheet.TryGetCellText(2, 2, out string? value));
            Assert.Equal("42", value);
        }

        [Fact]
        public void Test_CustomDocumentProperties_DirectCollectionEditsAreTracked() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCustomDocumentProperties.DirectCollection.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Data").CellValue(1, 1, "Ready");
                document.CustomDocumentProperties["Direct"] = new ExcelCustomProperty("Tracked");
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                Assert.Equal("Tracked", document.CustomDocumentProperties["Direct"].Text);
                document.CustomDocumentProperties.Clear();
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, false)) {
                Assert.Null(package.CustomFilePropertiesPart);
            }
        }

        [Fact]
        public void Test_ManualPageBreak_ConvertsExistingAutomaticBreak() {
            string filePath = Path.Combine(_directoryWithFiles, "PageBreaks.ConvertAutomatic.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Data").CellValue(1, 1, "Ready");
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart worksheetPart = package.WorkbookPart!.WorksheetParts.Single();
                worksheetPart.Worksheet.Append(new RowBreaks(
                    new Break { Id = 3U, Min = 0U, Max = 16383U, ManualPageBreak = false }) {
                    Count = 1U,
                    ManualBreakCount = 0U
                });
                worksheetPart.Worksheet.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                document["Data"].AddManualRowPageBreak(3);
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, false)) {
                Break pageBreak = Assert.Single(package.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<RowBreaks>()!.Elements<Break>());
                Assert.Equal(3U, pageBreak.Id!.Value);
                Assert.True(pageBreak.ManualPageBreak!.Value);
            }
        }

        [Fact]
        public void Test_TimePeriodConditionalFormatting_EmitsFormula() {
            string filePath = Path.Combine(_directoryWithFiles, "ConditionalTimePeriodFormula.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, new DateTime(2026, 6, 22));
                sheet.AddConditionalTimePeriodRule("A1:A3", TimePeriodValues.Today);
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelConditionalFormattingInfo rule = Assert.Single(document["Data"].GetConditionalFormattingRules("A1"));
                Assert.Equal("TimePeriod", rule.Type);
                Assert.Single(rule.Formulas);
                Assert.Contains("TODAY()", rule.Formulas[0], StringComparison.OrdinalIgnoreCase);
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_SetRowLayout_RejectsRowsBeyondWorksheetLimit() {
            string filePath = Path.Combine(_directoryWithFiles, "RowLayout.Bounds.xlsx");

            using var document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Data");

            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.SetRowLayout(A1.MaxRows + 1, new ExcelRowLayoutOptions { Height = 20 }));
        }

        [Fact]
        public void Test_DashboardBuilder_RejectsHeaderTableOverlap() {
            string filePath = Path.Combine(_directoryWithFiles, "DashboardBuilder.Overlap.xlsx");
            var data = new System.Data.DataTable("Sales");
            data.Columns.Add("Region", typeof(string));
            data.Rows.Add("EU");

            using var document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Dashboard");

            Assert.Throws<ArgumentException>(() => sheet.AddDashboard(data, new ExcelDashboardOptions {
                Title = "Sales Dashboard",
                TableRow = 1,
                AddChart = false
            }));
        }

        [Fact]
        public void Test_InspectionSnapshot_ExpandsRangeHyperlinksToCells() {
            string filePath = Path.Combine(_directoryWithFiles, "Inspection.RangeHyperlinks.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "One");
                sheet.CellValue(2, 1, "Two");
                sheet.CellValue(3, 1, "Three");
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart worksheetPart = package.WorkbookPart!.WorksheetParts.Single();
                worksheetPart.AddHyperlinkRelationship(new Uri("https://example.com/report", UriKind.Absolute), true, "rIdRangeHyperlink");
                worksheetPart.Worksheet.Append(new Hyperlinks(new Hyperlink {
                    Reference = "A1:A3",
                    Id = "rIdRangeHyperlink"
                }));
                worksheetPart.Worksheet.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelWorksheetSnapshot sheet = Assert.Single(document.CreateInspectionSnapshot().Worksheets);
                Assert.All(sheet.Cells.Where(cell => cell.Row >= 1 && cell.Row <= 3), cell => Assert.NotNull(cell.Hyperlink));
            }
        }

        [Fact]
        public void Test_CompareWorkbook_ReportsCommentOnlyCells() {
            string leftPath = Path.Combine(_directoryWithFiles, "CompareComments.Left.xlsx");
            string rightPath = Path.Combine(_directoryWithFiles, "CompareComments.Right.xlsx");

            using (var document = ExcelDocument.Create(leftPath)) {
                document.AddWorkSheet("Data");
                document.Save();
            }

            File.Copy(leftPath, rightPath, overwrite: true);
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(rightPath, true)) {
                WorksheetPart worksheetPart = package.WorkbookPart!.WorksheetParts.Single();
                WorksheetCommentsPart commentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>();
                commentsPart.Comments = new Comments(
                    new Authors(new Author("Reviewer")),
                    new CommentList(new Comment(
                        new CommentText(new Run(new Text("Blank cell note")))) {
                        Reference = "A1",
                        AuthorId = 0U
                    }));
                commentsPart.Comments.Save();
            }

            using var left = ExcelDocument.Load(leftPath, readOnly: true);
            using var right = ExcelDocument.Load(rightPath, readOnly: true);

            ExcelWorkbookDiffReport diff = left.CompareWorkbook(right, new ExcelWorkbookDiffOptions {
                CompareCells = false,
                CompareCellStyles = false,
                CompareTables = false,
                CompareWorksheetMetadata = false,
                CompareNamedRanges = false,
                CompareComments = true
            });

            Assert.Contains(diff.Differences, difference => difference.Category == "Comment" && difference.RightValue?.Contains("Blank cell note", StringComparison.Ordinal) == true);
        }

        [Fact]
        public void Test_ImageMoveTo_RejectsAbsoluteAnchors() {
            string filePath = Path.Combine(_directoryWithFiles, "ImageMove.AbsoluteAnchor.xlsx");
            byte[] image = File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png"));

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Images");
                sheet.AddImage(1, 1, image, "image/png", widthPixels: 16, heightPixels: 16, name: "Absolute image");
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart worksheetPart = package.WorkbookPart!.WorksheetParts.Single();
                Xdr.WorksheetDrawing drawing = worksheetPart.DrawingsPart!.WorksheetDrawing!;
                Xdr.OneCellAnchor oneCell = Assert.Single(drawing.Elements<Xdr.OneCellAnchor>());
                var absolute = new Xdr.AbsoluteAnchor(
                    new Xdr.Position { X = 0L, Y = 0L },
                    (Xdr.Extent)oneCell.Extent!.CloneNode(true),
                    (Xdr.Picture)oneCell.Descendants<Xdr.Picture>().Single().CloneNode(true),
                    new Xdr.ClientData());
                oneCell.Remove();
                drawing.Append(absolute);
                drawing.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                ExcelImage imageRecord = Assert.Single(document["Images"].Images);
                Assert.Throws<NotSupportedException>(() => imageRecord.MoveTo(2, 2));
            }
        }
    }
}
