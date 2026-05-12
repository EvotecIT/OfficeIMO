using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests for autofitting columns and rows.
    /// </summary>
    public partial class Excel {
        [Fact]
        public void Test_AutoFitColumnsAndRows() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Long piece of text");
                sheet.CellValue(2, 1, "Second line\nwith newline");
                sheet.CellValue(3, 1, "Line1\nLine2\nLine3");
                sheet.AutoFitColumns();
                sheet.AutoFitRows();
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var columns = wsPart.Worksheet.GetFirstChild<Columns>();
                Assert.NotNull(columns);
                var column = columns!.Elements<Column>().First();
                Assert.NotNull(column.Width);
                Assert.True(column.Width!.Value > 0);

                var sheetFormat = wsPart.Worksheet.GetFirstChild<SheetFormatProperties>();
                Assert.NotNull(sheetFormat);
                Assert.NotNull(sheetFormat!.DefaultRowHeight);
                Assert.True(sheetFormat.DefaultRowHeight.Value > 0);

                var row1 = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex != null && r.RowIndex.Value == 1);

                var row3 = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex != null && r.RowIndex.Value == 3);
                Assert.True(row3.CustomHeight?.Value ?? false);
                Assert.NotNull(row3.Height);
                Assert.True(row3.Height!.Value > 0);
            }
        }

        [Fact]
        public void Test_AutoFitRows_EmptyRowsRetainDefaultHeight() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.Empty.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Content");
                sheet.CellValue(2, 1, " ");
                sheet.AutoFitRows();
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var row2 = wsPart.Worksheet.Descendants<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == 2);
                Assert.NotNull(row2);
                Assert.False(row2!.CustomHeight?.Value ?? false);
                Assert.False(row2.Height?.HasValue ?? false);
            }
        }

        [Fact]
        public void Test_AutoFitRows_RemovesCustomHeightWhenCleared() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.ClearRow.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Content");
                sheet.AutoFitRows();
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets.First();
                sheet.CellValue(1, 1, string.Empty);
                sheet.AutoFitRows();
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var row1 = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex != null && r.RowIndex.Value == 1);
                Assert.False(row1.CustomHeight?.Value ?? false);
                Assert.False(row1.Height?.HasValue ?? false);
            }
        }

        [Fact]
        public void Test_AutoFitColumns_RemovesCustomWidthWhenCleared() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.ClearColumn.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Long text");
                sheet.AutoFitColumns();
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets.First();
                sheet.CellValue(1, 1, string.Empty);
                sheet.AutoFitColumns();
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var columns = wsPart.Worksheet.GetFirstChild<Columns>();
                Assert.True(columns == null || !columns.Elements<Column>().Any(c => c.Min != null && c.Max != null && c.Min.Value == 1 && c.Max.Value == 1));
            }
        }

        [Fact]
        public void Test_AutoFitColumns_DoesNotLeaveEmptyColumnsElementForFormulaOnlySheet() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.FormulaOnly.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellFormula(1, 8, "=1+1");
                sheet.AutoFitColumns();
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var columns = wsPart.Worksheet.GetFirstChild<Columns>();
                Assert.True(columns == null || columns.Elements<Column>().Any());

                OpenXmlValidator validator = new();
                Assert.Empty(validator.Validate(spreadsheet));
            }
        }

        [Fact]
        public void Test_AutoFitSingleColumn_DoesNotAffectOthers() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.SingleColumn.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Very long text that should expand the column");
                sheet.CellValue(1, 2, "Short");
                sheet.AutoFitColumn(1);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var columns = wsPart.Worksheet.GetFirstChild<Columns>();
                Assert.NotNull(columns);
                var column1 = columns!.Elements<Column>().FirstOrDefault(c => c.Min != null && c.Max != null && c.Min.Value == 1 && c.Max.Value == 1);
                Assert.NotNull(column1);
                Assert.True(column1!.BestFit?.Value ?? false);
                Assert.True(column1.Width?.Value > 0);

                var column2 = columns.Elements<Column>().FirstOrDefault(c => c.Min != null && c.Max != null && c.Min.Value == 2 && c.Max.Value == 2);
                Assert.Null(column2);
            }
        }

        [Fact]
        public void Test_AutoFitColumn_SplitsSpanningColumn() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.SplitColumn.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Very long text that should expand the column");
                sheet.CellValue(1, 2, "Short");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var worksheet = wsPart.Worksheet;
                var columns = worksheet.GetFirstChild<Columns>() ?? worksheet.InsertAt(new Columns(), 0);
                columns.RemoveAllChildren();
                columns.Append(new Column { Min = 1, Max = 2, Width = 10, CustomWidth = true });
                worksheet.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets.First();
                sheet.AutoFitColumn(1);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cols = wsPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToList();
                Assert.Equal(new uint[] { 1, 2 }, cols.Select(c => c.Min!.Value).ToArray());
                var column1 = cols.First(c => c.Min != null && c.Max != null && c.Min.Value == 1 && c.Max.Value == 1);
                var column2 = cols.First(c => c.Min != null && c.Max != null && c.Min.Value == 2 && c.Max.Value == 2);
                Assert.True(column1.Width!.Value > column2.Width!.Value);
                Assert.Equal(10.0, column2.Width!.Value);

                OpenXmlValidator validator = new();
                Assert.Empty(validator.Validate(spreadsheet));
            }
        }

        [Fact]
        public void Test_AutoFitColumnsFor_BatchApplySplitsSpanningColumnOnce() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.BatchSplitColumn.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Very long text that should expand the first column significantly");
                sheet.CellValue(1, 2, "Keep");
                sheet.CellValue(1, 3, "Very long text that should expand the third column significantly");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var worksheet = wsPart.Worksheet;
                var columns = worksheet.GetFirstChild<Columns>() ?? worksheet.InsertAt(new Columns(), 0);
                columns.RemoveAllChildren();
                columns.Append(new Column { Min = 1, Max = 3, Width = 10, CustomWidth = true });
                worksheet.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var sheet = document.Sheets.First();
                sheet.AutoFitColumnsFor(new[] { 1, 3 });
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cols = wsPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToList();
                Assert.Equal(new uint[] { 1, 2, 3 }, cols.Select(c => c.Min!.Value).ToArray());

                var column1 = cols.First(c => c.Min != null && c.Max != null && c.Min.Value == 1 && c.Max.Value == 1);
                var column2 = cols.First(c => c.Min != null && c.Max != null && c.Min.Value == 2 && c.Max.Value == 2);
                var column3 = cols.First(c => c.Min != null && c.Max != null && c.Min.Value == 3 && c.Max.Value == 3);

                Assert.True(column1.BestFit?.Value ?? false);
                Assert.True(column3.BestFit?.Value ?? false);
                Assert.Equal(10.0, column2.Width!.Value);
                Assert.False(column2.BestFit?.Value ?? false);

                OpenXmlValidator validator = new();
                Assert.Empty(validator.Validate(spreadsheet));
            }
        }

        [Fact]
        public void Test_AutoFitColumns_DeferredWorksheetSavePersistsOnDispose() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.DeferredWorksheetSave.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                document.Execution.SaveWorksheetAfterAutoFit = false;
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Very long text that should still persist after deferred AutoFit save");
                sheet.AutoFitColumns();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var column = wsPart.Worksheet.GetFirstChild<Columns>()?.Elements<Column>().FirstOrDefault(c => c.Min?.Value == 1 && c.Max?.Value == 1);
                Column nonNullColumn = Assert.IsType<Column>(column);
                Assert.True(nonNullColumn.BestFit?.Value ?? false);
                Assert.True(nonNullColumn.Width?.Value > 0);

                OpenXmlValidator validator = new();
                Assert.Empty(validator.Validate(spreadsheet));
            }
        }

        [Fact]
        public void Test_AutoFitColumns_ClampsToExcelMaxWidth() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.MaxWidth.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, new string('A', 5000));
                sheet.AutoFitColumns();
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var column = wsPart.Worksheet.GetFirstChild<Columns>()?.Elements<Column>().FirstOrDefault(c => c.Min?.Value == 1 && c.Max?.Value == 1);
                Column nonNullColumn = Assert.IsType<Column>(column);
                Assert.True(nonNullColumn.Width!.HasValue);
                Assert.True(nonNullColumn.Width.Value <= 255.0);
                Assert.True(nonNullColumn.Width.Value >= 200.0); // ensure we actually expanded significantly

                OpenXmlValidator validator = new();
                var validationErrors = validator.Validate(spreadsheet);
                Assert.Empty(validationErrors);
            }
        }

        [Fact]
        public void Test_AutoFitSingleRow_DoesNotAffectOthers() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.SingleRow.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Line1\nLine2\nLine3");
                sheet.CellValue(2, 1, "Line1\nLine2\nLine3");
                sheet.AutoFitRow(1);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var row1 = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex != null && r.RowIndex.Value == 1);
                Assert.True(row1.CustomHeight?.Value ?? false);
                Assert.True(row1.Height?.Value > 0);

                var row2 = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex != null && r.RowIndex.Value == 2);
                Assert.False(row2.CustomHeight?.Value ?? false);
                Assert.False(row2.Height?.HasValue ?? false);
            }
        }

        [Fact]
        public void Test_AutoFit_MixedFonts() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.MixedFonts.xlsx");

            static uint AddFontStyle(SpreadsheetDocument doc, string name, double size) {
                var stylesPart = doc.WorkbookPart!.WorkbookStylesPart ?? doc.WorkbookPart!.AddNewPart<WorkbookStylesPart>();
                if (stylesPart.Stylesheet == null) {
                    stylesPart.Stylesheet = new Stylesheet(new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font()), new Fills(new Fill()), new Borders(new Border()), new CellFormats(new CellFormat()));
                    stylesPart.Stylesheet.Fonts!.Count = 1;
                    stylesPart.Stylesheet.Fills!.Count = 1;
                    stylesPart.Stylesheet.Borders!.Count = 1;
                    stylesPart.Stylesheet.CellFormats!.Count = 1;
                }
                var ss = stylesPart.Stylesheet!;
                ss.Fonts!.Append(new DocumentFormat.OpenXml.Spreadsheet.Font(new FontName { Val = name }, new FontSize { Val = size }));
                ss.Fonts.Count = (uint)ss.Fonts.ChildElements.Count;
                ss.CellFormats!.Append(new CellFormat { FontId = ss.Fonts.Count - 1, ApplyFont = true });
                ss.CellFormats.Count = (uint)ss.CellFormats.ChildElements.Count;
                stylesPart.Stylesheet!.Save();
                return ss.CellFormats.Count - 1;
            }

            static void SetCellStyle(SpreadsheetDocument doc, string cellRef, uint styleIndex) {
                var wsPart = doc.WorkbookPart!.WorksheetParts.First();
                var cell = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == cellRef);
                cell.StyleIndex = styleIndex;
                wsPart.Worksheet.Save();
            }

            const string fontName = "OfficeIMO Test Sans";

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Small");
                sheet.CellValue(2, 1, "Large text");
                sheet.CellValue(3, 1, "Short");
                sheet.CellValue(3, 2, "Tall\nText");

                uint style = AddFontStyle(document._spreadSheetDocument, fontName, 20);
                SetCellStyle(document._spreadSheetDocument, "A2", style);
                SetCellStyle(document._spreadSheetDocument, "B3", style);

                sheet.AutoFitColumn(1);
                sheet.AutoFitRow(3);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var column = wsPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().First(c => c.Min != null && c.Max != null && c.Min.Value == 1 && c.Max.Value == 1);
                Assert.True(column.Width!.Value >= 7.0);

                var row = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex != null && r.RowIndex.Value == 3);
                Assert.True(row.Height!.Value > 40.0);
            }
        }

        [Fact]
        public void Test_AutoFit_NumericLikeTextWithCustomFont() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.NumericLikeText.xlsx");

            static uint AddFontStyle(SpreadsheetDocument doc, string name, double size) {
                var stylesPart = doc.WorkbookPart!.WorkbookStylesPart ?? doc.WorkbookPart!.AddNewPart<WorkbookStylesPart>();
                if (stylesPart.Stylesheet == null) {
                    stylesPart.Stylesheet = new Stylesheet(new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font()), new Fills(new Fill()), new Borders(new Border()), new CellFormats(new CellFormat()));
                    stylesPart.Stylesheet.Fonts!.Count = 1;
                    stylesPart.Stylesheet.Fills!.Count = 1;
                    stylesPart.Stylesheet.Borders!.Count = 1;
                    stylesPart.Stylesheet.CellFormats!.Count = 1;
                }

                var ss = stylesPart.Stylesheet!;
                ss.Fonts!.Append(new DocumentFormat.OpenXml.Spreadsheet.Font(new FontName { Val = name }, new FontSize { Val = size }));
                ss.Fonts.Count = (uint)ss.Fonts.ChildElements.Count;
                ss.CellFormats!.Append(new CellFormat { FontId = ss.Fonts.Count - 1, ApplyFont = true });
                ss.CellFormats.Count = (uint)ss.CellFormats.ChildElements.Count;
                stylesPart.Stylesheet!.Save();
                return ss.CellFormats.Count - 1;
            }

            static void SetCellStyle(SpreadsheetDocument doc, string cellRef, uint styleIndex) {
                var wsPart = doc.WorkbookPart!.WorksheetParts.First();
                var cell = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == cellRef);
                cell.StyleIndex = styleIndex;
                wsPart.Worksheet.Save();
            }

            const string fontName = "OfficeIMO Test Sans";

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "1234567890.12345");
                sheet.CellValue(1, 2, "12");

                uint style = AddFontStyle(document._spreadSheetDocument, fontName, 20);
                SetCellStyle(document._spreadSheetDocument, "A1", style);
                SetCellStyle(document._spreadSheetDocument, "B1", style);

                sheet.AutoFitColumns();
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var columns = wsPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToList();
                var column1 = columns.First(c => c.Min != null && c.Max != null && c.Min.Value == 1 && c.Max.Value == 1);
                var column2 = columns.First(c => c.Min != null && c.Max != null && c.Min.Value == 2 && c.Max.Value == 2);

                Assert.True(column1.Width!.Value >= 15.0);
                Assert.True(column1.Width!.Value > column2.Width!.Value);
            }
        }

        [Fact]
        public void Test_AutoFit_ToleratesDuplicateCustomNumberFormatIds() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.DuplicateNumberFormats.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, 1234.5d);

                var workbookPart = document._spreadSheetDocument.WorkbookPart!;
                var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet ??= new Stylesheet(
                    new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font()) { Count = 1 },
                    new Fills(new Fill()) { Count = 1 },
                    new Borders(new Border()) { Count = 1 },
                    new CellFormats(new CellFormat()) { Count = 1 });

                var stylesheet = stylesPart.Stylesheet;
                stylesheet.NumberingFormats ??= new NumberingFormats();
                stylesheet.NumberingFormats.Append(new NumberingFormat { NumberFormatId = 164U, FormatCode = "0.0" });
                stylesheet.NumberingFormats.Append(new NumberingFormat { NumberFormatId = 164U, FormatCode = "0.00" });
                stylesheet.NumberingFormats.Count = (uint)stylesheet.NumberingFormats.ChildElements.Count;

                stylesheet.CellFormats ??= new CellFormats(new CellFormat()) { Count = 1 };
                stylesheet.CellFormats.Append(new CellFormat { NumberFormatId = 164U, ApplyNumberFormat = true });
                stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.ChildElements.Count;

                var cell = document._spreadSheetDocument.WorkbookPart!.WorksheetParts.First().Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");
                cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1;
                stylesPart.Stylesheet.Save();

                var exception = Record.Exception(() => sheet.AutoFitColumns());
                Assert.Null(exception);
                document.Save();
            }
        }

        [Fact]
        public void Test_AutoFit_UsesFormattedDisplayTextForNumbersAndDates() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.FormattedDisplayText.xlsx");
            var date = new DateTime(2026, 5, 7);

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.Cell(1, 1, 1234.5, numberFormat: "$#,##0.00");
                sheet.CellValue(1, 2, 1234.5);
                sheet.Cell(1, 3, 1.0, numberFormat: "0.00%");
                sheet.CellValue(1, 4, 1.0);
                sheet.Cell(1, 5, date, numberFormat: "yyyy-mm-dd");
                sheet.CellValue(1, 6, date.ToOADate());

                sheet.AutoFitColumnsFor(new[] { 1, 2, 3, 4, 5, 6 });
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var columns = wsPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToList();
                double Width(uint index) => columns.First(c => c.Min?.Value == index && c.Max?.Value == index).Width!.Value;

                Assert.True(Width(1) > Width(2));
                Assert.True(Width(3) > Width(4));
                Assert.True(Width(5) > Width(6));
            }
        }

        [Fact]
        public void Test_AutoFitColumns_ReportsTimingSubStages() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.Timing.xlsx");
            var timings = new List<(string Operation, TimeSpan Elapsed)>();

            using (var document = ExcelDocument.Create(filePath)) {
                document.Execution.OnTiming = (operation, elapsed) => timings.Add((operation, elapsed));
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Long text for timing instrumentation");
                sheet.CellValue(2, 1, "Repeated shared string text");
                sheet.CellValue(3, 1, "Repeated shared string text");
                sheet.CellValue(1, 2, 1234.5);
                sheet.AutoFitColumns();
                document.Save();
            }

            var operations = timings.Select(t => t.Operation).ToArray();
            Assert.Contains("AutoFitColumns.BuildPlan", operations);
            Assert.Contains("AutoFitColumns.CalculateWidths", operations);
            Assert.Contains("AutoFitColumns.ApplyWidths", operations);
            Assert.All(timings, timing => Assert.True(timing.Elapsed >= TimeSpan.Zero));
        }

        [Fact]
        public void Test_AutoFit_UsesRichInlineStringRunFontsForColumnWidth() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.RichInlineString.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Tiny HugeWideText");
                sheet.CellValue(1, 2, "Tiny HugeWideText");

                WorksheetPart wsPart = document._spreadSheetDocument.WorkbookPart!.WorksheetParts.First();
                var richCell = wsPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == "A1");
                richCell.CellValue = null;
                richCell.DataType = CellValues.InlineString;
                richCell.InlineString = new InlineString(
                    new Run(
                        new RunProperties(new RunFont { Val = "Calibri" }, new FontSize { Val = 8D }),
                        new Text("Tiny ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(
                        new RunProperties(new RunFont { Val = "Calibri" }, new FontSize { Val = 26D }, new Bold()),
                        new Text("HugeWideText")));
                wsPart.Worksheet.Save();

                sheet.AutoFitColumnsFor(new[] { 1, 2 });
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var columns = wsPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToList();
                double richWidth = columns.First(c => c.Min?.Value == 1 && c.Max?.Value == 1).Width!.Value;
                double plainWidth = columns.First(c => c.Min?.Value == 2 && c.Max?.Value == 2).Width!.Value;

                Assert.True(richWidth > plainWidth * 1.4);
            }
        }

        [Fact]
        public void Test_AutoFitOperations_RunSequentially() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.ConcurrentOperations.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Long piece of text");
                sheet.CellValue(2, 1, "Second line\nwith newline");
                sheet.CellValue(3, 1, "Line1\nLine2\nLine3");

                sheet.AutoFitColumns();
                sheet.AutoFitRows();

                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var columns = wsPart.Worksheet.GetFirstChild<Columns>();
                Assert.NotNull(columns);
                var column = columns!.Elements<Column>().First();
                Assert.True(column.BestFit?.Value ?? false);
                Assert.NotNull(column.Width);
                Assert.True(column.Width!.Value > 0);

                var sheetFormat = wsPart.Worksheet.GetFirstChild<SheetFormatProperties>();
                Assert.NotNull(sheetFormat);
                Assert.NotNull(sheetFormat!.DefaultRowHeight);
                Assert.True(sheetFormat.DefaultRowHeight.Value > 0);
            }
        }

        [Fact]
        public async Task Test_AutoFitColumnRowConcurrentCalls_AreThreadSafe() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.ConcurrentSingle.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Long piece of text");
                sheet.CellValue(2, 1, "Line1\nLine2");

                var tasks = Enumerable.Range(0, 10)
                    .SelectMany(_ => new[] {
                        Task.Run(() => sheet.AutoFitColumn(1)),
                        Task.Run(() => sheet.AutoFitRow(2))
                    })
                    .ToArray();
                await Task.WhenAll(tasks);

                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var column = wsPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().First(c => c.Min != null && c.Max != null && c.Min.Value == 1 && c.Max.Value == 1);
                Assert.NotNull(column.Width);
                Assert.True(column.Width!.Value > 0);

                var row = wsPart.Worksheet.Descendants<Row>().First(r => r.RowIndex != null && r.RowIndex.Value == 2);
                Assert.True(row.CustomHeight?.Value ?? false);
                Assert.True(row.Height != null && row.Height.Value > 0);
            }
        }

        [Fact]
        public async Task Test_AutoFitColumns_CancellationPropagates() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFit.Cancellation.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                for (int row = 1; row <= 50; row++) {
                    for (int column = 1; column <= 50; column++) {
                        sheet.CellValue(row, column, $"Row {row} Column {column} {new string('X', 50)}");
                    }
                }

                // Use deterministic cancellation to avoid runtime-dependent timing flakiness
                // .NET 9 completes the parallel path noticeably faster, making time-based
                // CancelAfter unreliable here. We only validate that cancellation propagates.
                using CancellationTokenSource cts = new();
                cts.Cancel();
                await Assert.ThrowsAsync<OperationCanceledException>(async () =>
                    await Task.Run(() => sheet.AutoFitColumns(ExecutionMode.Parallel, cts.Token))
                );
            }
        }
    }
}
