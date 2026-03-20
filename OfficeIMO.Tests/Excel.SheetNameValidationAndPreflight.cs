using System;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
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
        public void GetOrCreateSheet_Sanitize_ReusesExistingSanitizedSheet() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (var doc = ExcelDocument.Create(path)) {
                var created = doc.AddWorkSheet("Data??");
                var resolved = doc.GetOrCreateSheet("Data??", SheetNameValidationMode.Sanitize);

                Assert.Single(doc.Sheets);
                Assert.Equal(created.Name, resolved.Name);
                Assert.Equal("Data", resolved.Name);
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

        [Fact]
        public void Preflight_RemovesEmptyCommentArtifactsBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("CommentsPreflight");

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var ws = wsPart.Worksheet;

                var commentsPart = wsPart.AddNewPart<WorksheetCommentsPart>();
                commentsPart.Comments = new Comments(new Authors(), new CommentList());
                commentsPart.Comments.Save();

                var vmlPart = wsPart.AddNewPart<VmlDrawingPart>();
                using (var stream = vmlPart.GetStream(FileMode.Create, FileAccess.Write)) {
                    using var writer = new StreamWriter(stream);
                    writer.Write("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" />");
                }

                var legacy = ws.GetFirstChild<LegacyDrawing>();
                if (legacy != null) {
                    legacy.Remove();
                }

                ws.Append(new LegacyDrawing { Id = wsPart.GetIdOfPart(vmlPart) });
                ws.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                var ws = wsPart.Worksheet;

                Assert.Null(wsPart.WorksheetCommentsPart);
                Assert.Empty(wsPart.VmlDrawingParts);
                Assert.Null(ws.Elements<LegacyDrawing>().FirstOrDefault());
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RemovesStaleHeaderFooterPictureArtifactsBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("HeaderFooterPreflight");

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var ws = wsPart.Worksheet;

                ws.AppendChild(new HeaderFooter {
                    OddHeader = new OddHeader("&CPlain header")
                });

                var vmlPart = wsPart.AddNewPart<VmlDrawingPart>();
                using (var stream = vmlPart.GetStream(FileMode.Create, FileAccess.Write)) {
                    using var writer = new StreamWriter(stream);
                    writer.Write("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" />");
                }

                ws.Append(new LegacyDrawingHeaderFooter { Id = wsPart.GetIdOfPart(vmlPart) });
                ws.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                var ws = wsPart.Worksheet;

                Assert.Empty(wsPart.VmlDrawingParts);
                Assert.Null(ws.Elements<LegacyDrawingHeaderFooter>().FirstOrDefault());
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RemovesConditionalFormattingWithoutRangesBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("ConditionalFormattingPreflight");

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var ws = wsPart.Worksheet;

                var conditional = new ConditionalFormatting();
                conditional.Append(new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.CellIs,
                    Priority = 1
                });
                ws.AppendChild(conditional);
                ws.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                Assert.Empty(wsPart.Worksheet.Elements<ConditionalFormatting>());
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_NormalizesConditionalFormattingPrioritiesBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("ConditionalFormattingPriority");

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var ws = wsPart.Worksheet;

                var first = new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A2" }
                };
                first.Append(new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.CellIs,
                    Priority = 9
                });

                var second = new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "B1:B2" }
                };
                second.Append(new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.ColorScale,
                    Priority = 9
                });

                ws.Append(first);
                ws.Append(second);
                ws.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                var priorities = wsPart.Worksheet.Descendants<ConditionalFormattingRule>()
                    .Select(rule => rule.Priority?.Value)
                    .ToList();

                Assert.Equal(new int?[] { 1, 2 }, priorities);
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RemovesEmptyWorksheetDrawingBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("DrawingPreflight");

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var ws = wsPart.Worksheet;

                var drawingPart = wsPart.AddNewPart<DrawingsPart>();
                drawingPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
                ws.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = wsPart.GetIdOfPart(drawingPart) });
                ws.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                Assert.Null(wsPart.DrawingsPart);
                Assert.Null(wsPart.Worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.Drawing>().FirstOrDefault());
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RemovesWorksheetDrawingAnchorWithMissingImageRelationshipBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("BrokenDrawing");

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var ws = wsPart.Worksheet;

                var drawingPart = wsPart.AddNewPart<DrawingsPart>();
                drawingPart.WorksheetDrawing = new Xdr.WorksheetDrawing(
                    new Xdr.OneCellAnchor(
                        new Xdr.FromMarker(
                            new Xdr.ColumnId("0"),
                            new Xdr.ColumnOffset("0"),
                            new Xdr.RowId("0"),
                            new Xdr.RowOffset("0")
                        ),
                        new Xdr.Extent { Cx = 9525, Cy = 9525 },
                        new Xdr.Picture(
                            new Xdr.NonVisualPictureProperties(
                                new Xdr.NonVisualDrawingProperties { Id = 1U, Name = "Broken Picture" },
                                new Xdr.NonVisualPictureDrawingProperties()
                            ),
                            new Xdr.BlipFill(
                                new A.Blip { Embed = "rIdMissing" },
                                new A.Stretch(new A.FillRectangle())
                            ),
                            new Xdr.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0, Y = 0 },
                                    new A.Extents { Cx = 9525, Cy = 9525 }
                                ),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                            )
                        ),
                        new Xdr.ClientData()
                    )
                );
                drawingPart.WorksheetDrawing.Save();
                ws.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = wsPart.GetIdOfPart(drawingPart) });
                ws.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                Assert.Null(wsPart.DrawingsPart);
                Assert.Null(wsPart.Worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.Drawing>().FirstOrDefault());
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RemovesBlankBuiltInDefinedNamesBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                doc.AddWorkSheet("Data");
                var workbook = doc._spreadSheetDocument.WorkbookPart!.Workbook;
                workbook.DefinedNames = new DefinedNames(
                    new DefinedName { Name = "_xlnm.Print_Area", LocalSheetId = 0U, Text = string.Empty },
                    new DefinedName { Name = "_xlnm.Print_Titles", LocalSheetId = 0U, Text = "   " }
                );
                workbook.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                Assert.Null(package.WorkbookPart!.Workbook.DefinedNames);
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RemovesBuiltInDefinedNamesPointingToWrongSheetBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                doc.AddWorkSheet("Data");
                doc.AddWorkSheet("Summary");
                var workbook = doc._spreadSheetDocument.WorkbookPart!.Workbook;
                workbook.DefinedNames = new DefinedNames(
                    new DefinedName { Name = "_xlnm.Print_Area", LocalSheetId = 0U, Text = "'Summary'!$A$1:$B$2" },
                    new DefinedName { Name = "_xlnm.Print_Titles", LocalSheetId = 0U, Text = "'Summary'!$1:$1,'Summary'!$A:$A" }
                );
                workbook.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                Assert.Null(package.WorkbookPart!.Workbook.DefinedNames);
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RemovesWorksheetAutoFilterConflictingWithTableBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Filters");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 1d);
                sheet.CellValue(3, 1, "B");
                sheet.CellValue(3, 2, 2d);
                sheet.AddTable("A1:B3", hasHeader: true, name: "FilterTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var ws = wsPart.Worksheet;

                ws.AppendChild(new AutoFilter {
                    Reference = "A1:B3"
                });
                ws.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                Assert.Null(wsPart.Worksheet.Elements<AutoFilter>().FirstOrDefault());
                var tablePart = wsPart.TableDefinitionParts.Single();
                Assert.Equal("A1:B3", tablePart.Table!.AutoFilter!.Reference!.Value);
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RemovesWorksheetAutoFilterOverlappingTableBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Filters");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 1d);
                sheet.CellValue(3, 1, "B");
                sheet.CellValue(3, 2, 2d);
                sheet.CellValue(4, 1, "C");
                sheet.CellValue(4, 2, 3d);
                sheet.AddTable("A1:B3", hasHeader: true, name: "FilterTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var ws = wsPart.Worksheet;

                ws.AppendChild(new AutoFilter {
                    Reference = "A2:B4"
                });
                ws.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                Assert.Null(wsPart.Worksheet.Elements<AutoFilter>().FirstOrDefault());
                var tablePart = wsPart.TableDefinitionParts.Single();
                Assert.Equal("A1:B3", tablePart.Table!.AutoFilter!.Reference!.Value);
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RemovesWorksheetAutoFilterWithoutRangeBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Filters");

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var ws = wsPart.Worksheet;

                var autoFilter = new AutoFilter();
                var filterColumn = new FilterColumn { ColumnId = 0U };
                filterColumn.Append(new Filters(new Filter { Val = "A" }));
                autoFilter.Append(filterColumn);
                ws.AppendChild(autoFilter);
                ws.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                Assert.Null(wsPart.Worksheet.Elements<AutoFilter>().FirstOrDefault());
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_NormalizesTableAutoFilterReferenceAndColumnsBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Filters");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 1d);
                sheet.CellValue(3, 1, "B");
                sheet.CellValue(3, 2, 2d);
                sheet.AddTable("A1:B3", hasHeader: true, name: "FilterTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var tablePart = wsPart.TableDefinitionParts.Single();
                var autoFilter = tablePart.Table!.AutoFilter!;
                autoFilter.Reference = "A1:A3";
                autoFilter.RemoveAllChildren<FilterColumn>();
                var keep = new FilterColumn { ColumnId = 0U };
                keep.Append(new Filters(new Filter { Val = "A" }));
                var duplicate = new FilterColumn { ColumnId = 0U };
                duplicate.Append(new Filters(new Filter { Val = "B" }));
                var outOfRange = new FilterColumn { ColumnId = 9U };
                outOfRange.Append(new Filters(new Filter { Val = "C" }));
                var empty = new FilterColumn { ColumnId = 1U };
                autoFilter.Append(keep, duplicate, outOfRange, empty);
                tablePart.Table.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                var tablePart = wsPart.TableDefinitionParts.Single();
                var autoFilter = tablePart.Table!.AutoFilter!;
                Assert.Equal("A1:B3", autoFilter.Reference!.Value);
                var filterColumns = autoFilter.Elements<FilterColumn>().ToList();
                Assert.Single(filterColumns);
                Assert.Equal(0U, filterColumns[0].ColumnId!.Value);
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RecreatesMissingWorksheetTablePartsBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Tables");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 1d);
                sheet.CellValue(3, 1, "B");
                sheet.CellValue(3, 2, 2d);
                sheet.AddTable("A1:B3", hasHeader: true, name: "RepairTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var tableParts = wsPart.Worksheet.Elements<TableParts>().FirstOrDefault();
                Assert.NotNull(tableParts);
                wsPart.Worksheet.RemoveChild(tableParts!);
                wsPart.Worksheet.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                var tableParts = wsPart.Worksheet.Elements<TableParts>().Single();
                Assert.Single(tableParts.Elements<TablePart>());
                Assert.Single(wsPart.TableDefinitionParts);
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RemovesTableDefinitionWithInvalidRangeBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Tables");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 1d);
                sheet.CellValue(3, 1, "B");
                sheet.CellValue(3, 2, 2d);
                sheet.AddTable("A1:B3", hasHeader: true, name: "RepairTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var tablePart = wsPart.TableDefinitionParts.Single();
                tablePart.Table!.Reference = "BadRange";
                tablePart.Table.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                Assert.Empty(wsPart.TableDefinitionParts);
                Assert.Null(wsPart.Worksheet.Elements<TableParts>().FirstOrDefault());
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_NormalizesMalformedTableColumnsBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Tables");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(1, 3, "Note");
                sheet.CellValue(2, 1, "A");
                sheet.CellValue(2, 2, 1d);
                sheet.CellValue(2, 3, "One");
                sheet.CellValue(3, 1, "B");
                sheet.CellValue(3, 2, 2d);
                sheet.CellValue(3, 3, "Two");
                sheet.AddTable("A1:C3", hasHeader: true, name: "RepairTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var tablePart = wsPart.TableDefinitionParts.Single();
                tablePart.Table!.TableColumns = new TableColumns(
                    new TableColumn { Id = 0U, Name = string.Empty },
                    new TableColumn { Id = 1U, Name = "Name" },
                    new TableColumn { Id = 1U, Name = string.Empty }) {
                    Count = 3U
                };
                tablePart.Table.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                var table = wsPart.TableDefinitionParts.Single().Table!;
                var columns = table.TableColumns!.Elements<TableColumn>().ToList();
                Assert.Equal(3U, table.TableColumns.Count!.Value);
                Assert.Equal(new uint[] { 1U, 2U, 3U }, columns.Select(column => column.Id!.Value).ToArray());
                Assert.All(columns, column => Assert.False(string.IsNullOrWhiteSpace(column.Name?.Value)));
                Assert.Equal(columns.Count, columns.Select(column => column.Name!.Value).Distinct(StringComparer.OrdinalIgnoreCase).Count());
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RemovesProtectedRangesWithoutSheetProtectionBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Protection");

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var ws = wsPart.Worksheet;

                ws.AppendChild(new ProtectedRanges(
                    new ProtectedRange {
                        Name = "UnlockedBlock",
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:B2" }
                    }));
                ws.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                Assert.Null(wsPart.Worksheet.Elements<SheetProtection>().FirstOrDefault());
                Assert.Null(wsPart.Worksheet.Elements<ProtectedRanges>().FirstOrDefault());
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RemovesMalformedProtectedRangesBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Protection");
                sheet.Protect();

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var ws = wsPart.Worksheet;

                var protectedRanges = new ProtectedRanges(
                    new ProtectedRange {
                        Name = "ValidRange",
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:B2" }
                    },
                    new ProtectedRange {
                        Name = "ValidRange",
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = "C1:D2" }
                    },
                    new ProtectedRange {
                        Name = "MissingRef"
                    },
                    new ProtectedRange {
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = "E1:F2" }
                    },
                    new ProtectedRange {
                        Name = "BadRef",
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = "NotACell" }
                    });

                var existingRanges = ws.Elements<ProtectedRanges>().FirstOrDefault();
                if (existingRanges != null) {
                    ws.ReplaceChild(protectedRanges, existingRanges);
                } else {
                    ws.AppendChild(protectedRanges);
                }
                ws.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                var ranges = wsPart.Worksheet.Elements<ProtectedRanges>().Single();
                var keptRanges = ranges.Elements<ProtectedRange>().ToList();
                Assert.Single(keptRanges);
                Assert.Equal("ValidRange", keptRanges[0].Name!.Value);
                Assert.Equal("A1:B2", keptRanges[0].SequenceOfReferences!.InnerText);
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_NormalizesDuplicateSheetProtectionBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Protection");

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var ws = wsPart.Worksheet;

                ws.AppendChild(new SheetProtection {
                    Sheet = false,
                    AutoFilter = true
                });
                ws.AppendChild(new SheetProtection {
                    Sheet = true,
                    Sort = true
                });
                ws.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                var protections = wsPart.Worksheet.Elements<SheetProtection>().ToList();
                Assert.Single(protections);
                Assert.True(protections[0].Sheet?.Value ?? false);
                Assert.True(protections[0].AutoFilter?.Value ?? false);
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RemovesHyperlinkWithMissingRelationshipBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Links");
                sheet.SetHyperlink(1, 1, "https://example.com", display: "Example");

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var hyperlink = wsPart.Worksheet.Elements<Hyperlinks>().Single().Elements<Hyperlink>().Single();
                var relationship = wsPart.HyperlinkRelationships.Single();
                wsPart.DeleteReferenceRelationship(relationship);
                wsPart.Worksheet.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                Assert.Null(wsPart.Worksheet.Elements<Hyperlinks>().FirstOrDefault());
                Assert.Empty(wsPart.HyperlinkRelationships);
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RemovesUnreferencedHyperlinkRelationshipBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Links");
                sheet.SetHyperlink(1, 1, "https://example.com/one", display: "One");

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                wsPart.AddHyperlinkRelationship(new Uri("https://example.com/orphan"), true, "rId999");
                wsPart.Worksheet.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                Assert.Single(wsPart.Worksheet.Elements<Hyperlinks>().Single().Elements<Hyperlink>());
                Assert.Single(wsPart.HyperlinkRelationships);
                Assert.DoesNotContain(wsPart.HyperlinkRelationships, relationship => relationship.Id == "rId999");
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_NormalizesDuplicateHyperlinksForSingleReferenceBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Links");
                sheet.SetHyperlink(1, 1, "https://example.com/one", display: "One");
                sheet.SetHyperlink(1, 2, "https://example.com/two", display: "Two");

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var hyperlinks = wsPart.Worksheet.Elements<Hyperlinks>().Single();
                var a1 = hyperlinks.Elements<Hyperlink>().First(hyperlink => hyperlink.Reference!.Value == "A1");
                var b1 = hyperlinks.Elements<Hyperlink>().First(hyperlink => hyperlink.Reference!.Value == "B1");
                hyperlinks.InsertAfter((Hyperlink)b1.CloneNode(true), a1);
                wsPart.Worksheet.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var wsPart = package.WorkbookPart!.WorksheetParts.First();
                var hyperlinks = wsPart.Worksheet.Elements<Hyperlinks>().Single();
                var items = hyperlinks.Elements<Hyperlink>().ToList();
                Assert.Equal(2, items.Count);
                Assert.Single(items, hyperlink => hyperlink.Reference!.Value == "A1");
                Assert.Single(items, hyperlink => hyperlink.Reference!.Value == "B1");
                Assert.Equal(2, wsPart.HyperlinkRelationships.Count());
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RemovesCalculationChainAndMarksWorkbookForRecalcBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Calc");
                sheet.CellValue(1, 1, 1d);
                sheet.CellValue(2, 1, 2d);
                sheet.CellFormula(3, 1, "SUM(A1:A2)");

                var workbookPart = doc._spreadSheetDocument.WorkbookPart!;
                var chainPart = workbookPart.AddNewPart<CalculationChainPart>();
                chainPart.CalculationChain = new CalculationChain();
                chainPart.CalculationChain.InnerXml = "<x:c xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" r=\"A3\" i=\"999\" />";
                workbookPart.Workbook.AppendChild(new CalculationProperties {
                    CalculationId = 1U,
                    ForceFullCalculation = false,
                    FullCalculationOnLoad = false
                });
                workbookPart.Workbook.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var workbookPart = package.WorkbookPart!;
                Assert.Empty(workbookPart.GetPartsOfType<CalculationChainPart>());
                var calcProps = workbookPart.Workbook.Elements<CalculationProperties>().FirstOrDefault();
                Assert.NotNull(calcProps);
                Assert.Equal(191029U, calcProps!.CalculationId!.Value);
                Assert.True(calcProps.ForceFullCalculation?.Value ?? false);
                Assert.True(calcProps.FullCalculationOnLoad?.Value ?? false);
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RemovesCalculationChainWithoutFormulasBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Calc");
                sheet.CellValue(1, 1, "Value");

                var workbookPart = doc._spreadSheetDocument.WorkbookPart!;
                var chainPart = workbookPart.AddNewPart<CalculationChainPart>();
                chainPart.CalculationChain = new CalculationChain();
                chainPart.CalculationChain.InnerXml = "<x:c xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" r=\"A1\" i=\"1\" />";
                workbookPart.Workbook.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var workbookPart = package.WorkbookPart!;
                Assert.Empty(workbookPart.GetPartsOfType<CalculationChainPart>());
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_NormalizesStylesheetAndResetsInvalidStyleIndexesBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Styles");
                sheet.CellValue(1, 1, "Value");

                var workbookPart = doc._spreadSheetDocument.WorkbookPart!;
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet(
                    new Fonts(new Font()) { Count = 9U },
                    new Fills(new Fill(new PatternFill { PatternType = PatternValues.None })) { Count = 1U },
                    new Borders(new Border()) { Count = 7U },
                    new CellStyleFormats(new CellFormat()) { Count = 4U },
                    new CellFormats(
                        new CellFormat {
                            FontId = 8U,
                            FillId = 5U,
                            BorderId = 4U,
                            FormatId = 3U
                        }) { Count = 1U });
                stylesPart.Stylesheet.Save();

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var cell = wsPart.Worksheet.Descendants<Cell>().Single();
                cell.StyleIndex = 42U;
                wsPart.Worksheet.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var workbookPart = package.WorkbookPart!;
                var stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet!;
                Assert.Equal(1U, stylesheet.Fonts!.Count!.Value);
                Assert.Equal(2U, stylesheet.Fills!.Count!.Value);
                Assert.Equal(1U, stylesheet.Borders!.Count!.Value);
                Assert.Equal(1U, stylesheet.CellStyleFormats!.Count!.Value);
                Assert.Equal(1U, stylesheet.CellFormats!.Count!.Value);

                var baseFormat = stylesheet.CellFormats.Elements<CellFormat>().Single();
                Assert.Equal(0U, baseFormat.FontId!.Value);
                Assert.Equal(0U, baseFormat.FillId!.Value);
                Assert.Equal(0U, baseFormat.BorderId!.Value);
                Assert.Equal(0U, baseFormat.FormatId!.Value);

                var cell = workbookPart.WorksheetParts.First().Worksheet.Descendants<Cell>().Single();
                Assert.Equal(0U, cell.StyleIndex!.Value);
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_RepairsSharedStringTableMetadataAndMissingIndexesBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Strings");
                sheet.CellValue(1, 1, "Alpha");
                sheet.CellValue(2, 1, "Beta");

                var workbookPart = doc._spreadSheetDocument.WorkbookPart!;
                var sharedStringTable = workbookPart.SharedStringTablePart!.SharedStringTable!;
                sharedStringTable.RemoveChild(sharedStringTable.Elements<SharedStringItem>().Last());
                sharedStringTable.Count = 99U;
                sharedStringTable.UniqueCount = 1U;
                sharedStringTable.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var workbookPart = package.WorkbookPart!;
                var sharedStringTable = workbookPart.SharedStringTablePart!.SharedStringTable!;
                var items = sharedStringTable.Elements<SharedStringItem>().ToList();
                Assert.Equal(2, items.Count);
                Assert.Equal("Alpha", items[0].InnerText);
                Assert.Equal(string.Empty, items[1].InnerText);
                Assert.Equal(2U, sharedStringTable.Count!.Value);
                Assert.Equal(2U, sharedStringTable.UniqueCount!.Value);
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }

        [Fact]
        public void Preflight_ConvertsMalformedSharedStringCellToInlineStringBeforeSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string savePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var doc = ExcelDocument.Create(path)) {
                var sheet = doc.AddWorkSheet("Strings");
                sheet.CellValue(1, 1, "Alpha");

                var wsPartField = typeof(ExcelSheet).GetField("_worksheetPart", BindingFlags.NonPublic | BindingFlags.Instance);
                Assert.NotNull(wsPartField);
                var wsPart = (WorksheetPart)wsPartField!.GetValue(sheet)!;
                var cell = wsPart.Worksheet.Descendants<Cell>().Single();
                cell.DataType = CellValues.SharedString;
                cell.CellValue = new CellValue("NotAnIndex");
                cell.InlineString = null;
                wsPart.Worksheet.Save();

                doc.Save(savePath, openExcel: false);
            }

            using (var package = SpreadsheetDocument.Open(savePath, false)) {
                var cell = package.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().Single();
                Assert.Equal(CellValues.InlineString, cell.DataType!.Value);
                Assert.Null(cell.CellValue);
                Assert.Equal("NotAnIndex", cell.InlineString!.InnerText);
            }

            using (var reopened = ExcelDocument.Load(savePath, readOnly: true)) {
                Assert.Empty(reopened.ValidateOpenXml());
            }

            File.Delete(path);
            File.Delete(savePath);
        }
    }
}

