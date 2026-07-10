using System;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
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
        public void Test_WorkbookConnectionMetadata_WrapsSingletonConnectionXml() {
            string filePath = Path.Combine(_directoryWithFiles, "ConnectionMetadata.Singleton.xlsx");
            const string connectionXml = "<connection xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"SalesConnection\" type=\"5\" refreshedVersion=\"7\"/>";

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Data");
                document.AddWorkbookConnectionMetadata(connectionXml);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            string xml = ReadSinglePackagePartText(spreadsheet.WorkbookPart!, "connections");
            Assert.Contains("<connections", xml, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("count=\"1\"", xml, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("SalesConnection", xml, StringComparison.Ordinal);
        }

        [Fact]
        public void Test_WorksheetMetadata_RejectsForeignWorksheet() {
            using var left = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            using var right = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            ExcelSheet foreignSheet = left.AddWorkSheet("Foreign");

            Assert.Throws<ArgumentException>(() => right.AddWorksheetMetadataPart(
                foreignSheet,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml",
                "<queryTable xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"Foreign\"/>"));
        }

        [Fact]
        public void Test_PrintLayoutWorksheetPreset_ClearsStalePrintTitles() {
            string filePath = Path.Combine(_directoryWithFiles, "PrintLayoutPreset.ClearTitles.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Report");
                sheet.CellValue(1, 1, "Region");
                document.SetPrintTitles(sheet, firstRow: 1, lastRow: 1, firstCol: 1, lastCol: 1, save: false);

                sheet.ApplyPrintLayout(new ExcelPrintLayoutOptions {
                    Preset = ExcelPrintLayoutPreset.Worksheet
                });
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Null(document.GetNamedRange("_xlnm.Print_Titles", document["Report"]));
            }
        }

        [Fact]
        public void Test_ColumnRangeByHeader_MaterializesDeferredWorksheetData() {
            string filePath = Path.Combine(_directoryWithFiles, "ColumnRangeByHeader.Deferred.xlsx");
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Sales Amount", typeof(int));
            table.Rows.Add("Alpha", 10);
            table.Rows.Add("Beta", 20);
            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(filePath)) {
                document.InsertDataSet(dataSet);

                Assert.Equal("B2:B3", document["Items"].GetColumnRangeByHeader("Sales Amount"));
                document.Save();
            }
        }

        [Fact]
        public void Test_DateSystemChange_UpdatesDeferredDataSetImport() {
            string filePath = Path.Combine(_directoryWithFiles, "DateSystem.DeferredChange.xlsx");
            var date = new DateTime(2024, 2, 3);
            var dataSet = new DataSet("Export");
            var table = new DataTable("Items");
            table.Columns.Add("When", typeof(DateTime));
            table.Rows.Add(date);
            dataSet.Tables.Add(table);

            using (var document = ExcelDocument.Create(filePath)) {
                document.InsertDataSet(dataSet);
                document.DateSystem = ExcelDateSystem.NineteenFour;
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            string serialText = GetCellValueText(worksheetPart, "A2");
            Assert.Equal(Expected1904Serial(date), double.Parse(serialText, System.Globalization.CultureInfo.InvariantCulture), 6);
        }

        [Fact]
        public void Test_EnsureWorkbookTheme_SavesEnsuredPartsForStreamWorkbook() {
            using var source = new MemoryStream();
            using (var document = ExcelDocument.Create(source, autoSave: false)) {
                document.AddWorkSheet("Data").CellValue(1, 1, "Ready");
                document.Save(source);
            }

            source.Position = 0;
            using (SpreadsheetDocument package = SpreadsheetDocument.Open(source, true)) {
                foreach (ThemePart themePart in package.WorkbookPart!.GetPartsOfType<ThemePart>().ToList()) {
                    package.WorkbookPart.DeletePart(themePart);
                }

                if (package.WorkbookPart.WorkbookStylesPart != null) {
                    package.WorkbookPart.DeletePart(package.WorkbookPart.WorkbookStylesPart);
                }
            }

            source.Position = 0;
            using var destination = new MemoryStream();
            using (var document = ExcelDocument.Load(source)) {
                document.EnsureWorkbookTheme();
                document.Save(destination);
            }

            destination.Position = 0;
            using SpreadsheetDocument saved = SpreadsheetDocument.Open(destination, false);
            Assert.Single(saved.WorkbookPart!.GetPartsOfType<ThemePart>());
            Assert.NotNull(saved.WorkbookPart.WorkbookStylesPart);
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
                    Reference = "A1:XFD1048576",
                    Id = "rIdRangeHyperlink"
                }));
                worksheetPart.Worksheet.Save();
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelWorksheetSnapshot sheet = Assert.Single(document.CreateInspectionSnapshot().Worksheets);
                Assert.All(sheet.Cells.Where(cell => cell.Row >= 1 && cell.Row <= 3), cell => Assert.NotNull(cell.Hyperlink));
                Assert.Equal(3, sheet.Cells.Count(cell => cell.Hyperlink != null));
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

        [Fact]
        public void Test_ImageMoveTo_RejectsOutOfBoundsCellAnchors() {
            string filePath = Path.Combine(_directoryWithFiles, "ImageMove.Bounds.xlsx");
            byte[] image = File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png"));

            using var document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Images");
            ExcelImage imageRecord = sheet.AddImage(1, 1, image, "image/png", widthPixels: 16, heightPixels: 16);

            Assert.Throws<ArgumentOutOfRangeException>(() => imageRecord.MoveTo(A1.MaxRows + 1, 1));
            Assert.Throws<ArgumentOutOfRangeException>(() => imageRecord.MoveTo(1, A1.MaxColumns + 1));
        }

        [Fact]
        public void Test_ImageMoveTo_PreservesTwoCellAnchorExtent() {
            string filePath = Path.Combine(_directoryWithFiles, "ImageMove.TwoCellAnchor.xlsx");
            byte[] image = File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png"));

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Images");
                sheet.AddImage(1, 1, image, "image/png", widthPixels: 16, heightPixels: 16, name: "Resizable image");
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart worksheetPart = package.WorkbookPart!.WorksheetParts.Single();
                Xdr.WorksheetDrawing drawing = worksheetPart.DrawingsPart!.WorksheetDrawing!;
                Xdr.OneCellAnchor oneCell = Assert.Single(drawing.Elements<Xdr.OneCellAnchor>());
                Xdr.Picture picture = (Xdr.Picture)oneCell.Descendants<Xdr.Picture>().Single().CloneNode(true);
                var twoCell = new Xdr.TwoCellAnchor(
                    new Xdr.FromMarker(
                        new Xdr.ColumnId("0"),
                        new Xdr.ColumnOffset("0"),
                        new Xdr.RowId("0"),
                        new Xdr.RowOffset("0")),
                    new Xdr.ToMarker(
                        new Xdr.ColumnId("2"),
                        new Xdr.ColumnOffset(PixelsToEmuText(1)),
                        new Xdr.RowId("4"),
                        new Xdr.RowOffset(PixelsToEmuText(2))),
                    picture,
                    new Xdr.ClientData());
                oneCell.Remove();
                drawing.Append(twoCell);
                drawing.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                ExcelImage imageRecord = Assert.Single(document["Images"].Images);
                imageRecord.MoveTo(3, 4, offsetXPixels: 5, offsetYPixels: 7);
                document.Save();
            }

            using (SpreadsheetDocument package = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart worksheetPart = package.WorkbookPart!.WorksheetParts.Single();
                Xdr.TwoCellAnchor twoCell = Assert.Single(worksheetPart.DrawingsPart!.WorksheetDrawing!.Elements<Xdr.TwoCellAnchor>());
                Assert.Equal("3", twoCell.FromMarker!.ColumnId!.Text);
                Assert.Equal(PixelsToEmuText(5), twoCell.FromMarker.ColumnOffset!.Text);
                Assert.Equal("2", twoCell.FromMarker.RowId!.Text);
                Assert.Equal(PixelsToEmuText(7), twoCell.FromMarker.RowOffset!.Text);
                Assert.Equal("5", twoCell.ToMarker!.ColumnId!.Text);
                Assert.Equal(PixelsToEmuText(6), twoCell.ToMarker.ColumnOffset!.Text);
                Assert.Equal("6", twoCell.ToMarker.RowId!.Text);
                Assert.Equal(PixelsToEmuText(9), twoCell.ToMarker.RowOffset!.Text);
            }
        }

        [Fact]
        public void Test_DashboardChart_RejectsOutOfBoundsAnchorsAndDimensions() {
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            ExcelSheet sheet = document.AddWorkSheet("Dashboard");
            sheet.CellValue(1, 1, "Region");
            sheet.CellValue(1, 2, "Sales");
            sheet.CellValue(2, 1, "EU");
            sheet.CellValue(2, 2, 10);

            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddDashboardChart(new ExcelDashboardChartOptions { Range = "A1:B2", Row = A1.MaxRows + 1, Column = 1 }));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddDashboardChart(new ExcelDashboardChartOptions { Range = "A1:B2", Row = 1, Column = A1.MaxColumns + 1 }));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddDashboardChart(new ExcelDashboardChartOptions { Range = "A1:B2", WidthPixels = 0 }));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddDashboardChart(new ExcelDashboardChartOptions { Range = "A1:B2", HeightPixels = -1 }));
        }

        private static string PixelsToEmuText(int pixels)
            => ((long)Math.Round(pixels * 9525.0)).ToString(System.Globalization.CultureInfo.InvariantCulture);

        [Fact]
        public void Test_PrintAreaAndTitles_MarkStreamWorkbookDirty() {
            using var source = new MemoryStream();
            using (var document = ExcelDocument.Create(source, autoSave: false)) {
                ExcelSheet sheet = document.AddWorkSheet("Report");
                sheet.CellValue(1, 1, "Region");
                document.Save(source);
            }

            source.Position = 0;
            using var destination = new MemoryStream();
            using (var document = ExcelDocument.Load(source)) {
                ExcelSheet sheet = document["Report"];
                document.SetPrintArea(sheet, "A1:B2", save: false);
                document.SetPrintTitles(sheet, firstRow: 1, lastRow: 1, firstCol: 1, lastCol: 1, save: false);
                document.Save(destination);
            }

            destination.Position = 0;
            using var loaded = ExcelDocument.Load(destination, readOnly: true);
            ExcelSheet loadedSheet = loaded["Report"];
            Assert.Equal("$A$1:$B$2", loadedSheet.GetPrintArea());
            ExcelPrintTitles titles = loadedSheet.GetPrintTitles();
            Assert.Equal(1, titles.FirstRow);
            Assert.Equal(1, titles.LastRow);
            Assert.Equal(1, titles.FirstColumn);
            Assert.Equal(1, titles.LastColumn);
        }

        [Fact]
        public void Test_WorksheetQueryTableMetadata_LinksWorksheetRelationship() {
            string filePath = Path.Combine(_directoryWithFiles, "QueryTableMetadata.Linked.xlsx");
            const string queryTableXml = "<queryTable xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"SalesQuery\" connectionId=\"1\"/>";

            using (var document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Data");
                document.AddWorksheetQueryTableMetadata("Data", queryTableXml);
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Worksheet worksheet = worksheetPart.Worksheet!;
            OpenXmlElement queryTableParts = Assert.Single(worksheet.ChildElements,
                element => element.LocalName == "queryTableParts" && element.NamespaceUri == "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            OpenXmlElement queryTablePart = Assert.Single(queryTableParts.ChildElements,
                element => element.LocalName == "queryTablePart" && element.NamespaceUri == "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            Assert.Equal("1", queryTableParts.GetAttribute("count", string.Empty).Value);
            string? relationshipId = queryTablePart.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value;

            Assert.False(string.IsNullOrWhiteSpace(relationshipId));
            Assert.NotNull(worksheetPart.GetPartById(relationshipId!));
            Assert.Contains("SalesQuery", ReadSinglePackagePartText(worksheetPart, "queryTable"));
        }

        [Fact]
        public void Test_PrintLayout_RejectsScaleOutsideExcelBounds() {
            using var document = ExcelDocument.Create(new MemoryStream(), autoSave: false);
            ExcelSheet sheet = document.AddWorkSheet("Report");

            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.ApplyPrintLayout(new ExcelPrintLayoutOptions {
                Preset = ExcelPrintLayoutPreset.Worksheet,
                Scale = 1U
            }));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.ApplyPrintLayout(new ExcelPrintLayoutOptions {
                Preset = ExcelPrintLayoutPreset.Worksheet,
                Scale = 401U
            }));
        }
    }
}
