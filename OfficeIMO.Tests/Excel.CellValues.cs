using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests for setting cell values in Excel sheets.
    /// </summary>
    public partial class Excel {
        [Fact]
        public void Test_CellValues() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValues.xlsx");
            var date = new DateTime(2024, 1, 2);
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Hello");
                sheet.CellValue(2, 1, 123.45);
                sheet.CellValue(3, 1, 678.90m);
                sheet.CellValue(4, 1, date);
                sheet.CellValue(5, 1, true);
                sheet.CellFormula(6, 1, "SUM(A2:A3)");
                sheet.Cell(7, 1, 1.23, "A2+1", "0.00");
                document.Save();
            }

            SpreadsheetDocument spreadsheet = null!;
            Exception? ex = Record.Exception(() => spreadsheet = SpreadsheetDocument.Open(filePath, false));
            Assert.Null(ex);
            using (spreadsheet) {
                  ValidateSpreadsheetDocument(filePath, spreadsheet);

                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cells = wsPart.Worksheet.Descendants<Cell>().ToList();
                SharedStringTablePart shared = spreadsheet.WorkbookPart!.SharedStringTablePart!;
                Assert.NotNull(shared);
                Assert.Equal("Hello", shared.SharedStringTable!.ElementAt(0).InnerText);

                Cell cellString = cells.First(c => c.CellReference == "A1");
                  Assert.Equal(CellValues.SharedString, cellString.DataType!.Value);
                  Assert.Equal("0", cellString.CellValue!.Text);

                Cell cellDouble = cells.First(c => c.CellReference == "A2");
                  Assert.Equal(CellValues.Number, cellDouble.DataType!.Value);
                  Assert.Equal("123.45", cellDouble.CellValue!.Text);

                Cell cellDecimal = cells.First(c => c.CellReference == "A3");
                  Assert.Equal(CellValues.Number, cellDecimal.DataType!.Value);
                  Assert.Equal("678.90", cellDecimal.CellValue!.Text);

                Cell cellDate = cells.First(c => c.CellReference == "A4");
                  Assert.Equal(CellValues.Number, cellDate.DataType!.Value);
                  Assert.Equal(date.ToOADate().ToString(CultureInfo.InvariantCulture), cellDate.CellValue!.Text);

                Cell cellBool = cells.First(c => c.CellReference == "A5");
                  Assert.Equal(CellValues.Boolean, cellBool.DataType!.Value);
                  Assert.Equal("1", cellBool.CellValue!.Text);

                Cell cellFormula = cells.First(c => c.CellReference == "A6");
                  Assert.NotNull(cellFormula.CellFormula);
                  Assert.Equal("SUM(A2:A3)", cellFormula.CellFormula!.Text);

                Cell cellCombined = cells.First(c => c.CellReference == "A7");
                  Assert.Equal(CellValues.Number, cellCombined.DataType!.Value);
                  Assert.Equal("1.23", cellCombined.CellValue!.Text);
                  Assert.NotNull(cellCombined.CellFormula);
                  Assert.Equal("A2+1", cellCombined.CellFormula!.Text);
                  Assert.NotNull(cellCombined.StyleIndex);
                var styles = spreadsheet.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
                var formats = styles.NumberingFormats!.Elements<NumberingFormat>().ToList();
                Assert.Contains(formats, n => n.FormatCode != null && n.FormatCode.Value == "0.00");
            }
        }

        [Fact]
        public void Test_CellValue_NullAndDbNullUseEmptyStringCells() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesNullEmptyString.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, (object?)null);
                sheet.CellValue(2, 1, DBNull.Value);
                sheet.CellValue(3, 1, (string?)null);
                sheet.CellValue(4, 1, string.Empty);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                ValidateSpreadsheetDocument(filePath, spreadsheet);
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cells = wsPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);

                Assert.Equal(CellValues.String, cells["A1"].DataType!.Value);
                Assert.Equal(string.Empty, cells["A1"].CellValue!.Text);
                Assert.Equal(CellValues.String, cells["A2"].DataType!.Value);
                Assert.Equal(string.Empty, cells["A2"].CellValue!.Text);
                Assert.Equal(CellValues.String, cells["A3"].DataType!.Value);
                Assert.Equal(string.Empty, cells["A3"].CellValue!.Text);
                Assert.Equal(CellValues.String, cells["A4"].DataType!.Value);
                Assert.Equal(string.Empty, cells["A4"].CellValue!.Text);
            }
        }

        [Fact]
        public void Test_CellValues_WritesCellsImmediately() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesImmediate.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"Name"),
                    (1, 2, (object)"Amount"),
                    (2, 1, (object)"Alpha"),
                    (2, 2, (object)10)
                });

                Assert.True(sheet.TryGetCellText(2, 1, out var value));
                Assert.Equal("Alpha", value);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_CellValues_HeaderlessDeferredSaveSnapshotsInputValues() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesHeaderlessDeferredSnapshot.xlsx");
            var cells = new (int Row, int Column, object Value)[] {
                (1, 1, 1.25d),
                (1, 2, true),
                (1, 3, "Original"),
                (2, 1, 2.5d),
                (2, 2, false),
                (2, 3, "Second")
            };

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValues(cells);
                cells[2] = (1, 3, "Changed");
                cells[5] = (2, 3, "Changed");
                document.Save();
            }

            using (var reader = ExcelDocumentReader.Open(filePath)) {
                object?[,] values = reader.GetSheet("Data").ReadRange("A1:C2");
                Assert.Equal(1.25d, values[0, 0]);
                Assert.Equal(true, values[0, 1]);
                Assert.Equal("Original", values[0, 2]);
                Assert.Equal(2.5d, values[1, 0]);
                Assert.Equal(false, values[1, 1]);
                Assert.Equal("Second", values[1, 2]);
            }
        }

        [Fact]
        public void Test_CellValue_PendingDirectWrites_ReadAndSaveCorrectly() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuePendingDirectWrites.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                for (int row = 1; row <= 70; row++) {
                    sheet.CellValue(row, 1, "Name " + row.ToString(CultureInfo.InvariantCulture));
                    sheet.CellValue(row, 2, row);
                }

                Assert.True(sheet.TryGetCellText(2, 1, out string visibleText));
                Assert.Equal("Name 2", visibleText);

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                ValidateSpreadsheetDocument(filePath, spreadsheet);
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cells = wsPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);

                Cell cellA1 = cells["A1"];
                string textA1 = cellA1.CellValue!.Text;
                if (cellA1.DataType?.Value == CellValues.SharedString) {
                    textA1 = spreadsheet.WorkbookPart!.SharedStringTablePart!.SharedStringTable!
                        .ElementAt(int.Parse(textA1, CultureInfo.InvariantCulture))
                        .InnerText;
                }

                Assert.Equal("Name 1", textA1);
                Assert.Equal("70", cells["B70"].CellValue!.Text);
            }
        }

        [Fact]
        public void Test_CellFormula_PendingDirectWrites_SaveCorrectly() {
            string filePath = Path.Combine(_directoryWithFiles, "CellFormulaPendingDirectWrites.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Formulas");
                for (int row = 1; row <= 70; row++) {
                    sheet.CellValue(row, 1, (double)row);
                    sheet.CellValue(row, 2, (double)(row % 17));
                    sheet.CellValue(row, 3, (double)(row % 29));
                    sheet.CellFormula(row, 4, $"=SUM(A{row}:C{row})");
                }

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                ValidateSpreadsheetDocument(filePath, spreadsheet);
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var cells = wsPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);

                Assert.Equal("70", cells["A70"].CellValue!.Text);
                Assert.NotNull(cells["D70"].CellFormula);
                Assert.Equal("SUM(A70:C70)", cells["D70"].CellFormula!.Text);
                Assert.Null(cells["D70"].CellValue);
            }
        }

        [Fact]
        public void Test_TryGetCellText_MissingCell_DoesNotCreateCell() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesMissingLookupNoMutation.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Header");

                Assert.False(sheet.TryGetCellText(10, 5, out _));

                WorksheetPart wsPart = document._spreadSheetDocument.WorkbookPart!.WorksheetParts.First();
                Assert.DoesNotContain(wsPart.Worksheet.Descendants<Row>(), row => row.RowIndex?.Value == 10U);
                Assert.DoesNotContain(wsPart.Worksheet.Descendants<Cell>(), cell => cell.CellReference?.Value == "E10");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_TryGetCellText_OutOfOrderRows_FindsExistingCell() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesOutOfOrderRows.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(5, 1, "Later");
                sheet.CellValue(2, 1, "Target");

                WorksheetPart wsPart = document._spreadSheetDocument.WorkbookPart!.WorksheetParts.First();
                SheetData sheetData = wsPart.Worksheet.GetFirstChild<SheetData>()!;
                Row row2 = sheetData.Elements<Row>().First(row => row.RowIndex?.Value == 2U);
                row2.Remove();
                sheetData.Append(row2);

                Assert.True(sheet.TryGetCellText(5, 1, out _));
                Assert.True(sheet.TryGetCellText(2, 1, out var text));
                Assert.Equal("Target", text);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_TryGetCellText_OutOfOrderCells_FindsExistingCell() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesOutOfOrderCells.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 3, "Later");
                sheet.CellValue(1, 1, "Target");

                WorksheetPart wsPart = document._spreadSheetDocument.WorkbookPart!.WorksheetParts.First();
                Row row1 = wsPart.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().First(row => row.RowIndex?.Value == 1U);
                Cell cellA1 = row1.Elements<Cell>().First(cell => cell.CellReference?.Value == "A1");
                cellA1.Remove();
                row1.Append(cellA1);

                Assert.True(sheet.TryGetCellText(1, 3, out _));
                Assert.True(sheet.TryGetCellText(1, 1, out var text));
                Assert.Equal("Target", text);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_CellAtGetValue_MissingCell_DoesNotCreateCell() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesObjectModelMissingLookupNoMutation.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(10, 1, "Existing row");

                ExcelCellData value = sheet.CellAt(10, 5).GetValue();

                Assert.Equal(ExcelCellDataKind.Blank, value.Kind);
                Assert.Null(value.Value);
                WorksheetPart wsPart = document._spreadSheetDocument.WorkbookPart!.WorksheetParts.First();
                Assert.DoesNotContain(wsPart.Worksheet.Descendants<Cell>(), cell => cell.CellReference?.Value == "E10");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_TryGetCellText_UsesFreshSharedStringsAfterMutation() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesSharedStringCache.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Alpha");
                sheet.CellValue(2, 1, "Beta");

                Assert.True(sheet.TryGetCellText(1, 1, out var first));
                Assert.Equal("Alpha", first);
                Assert.Equal("A2", sheet.FindFirst("Beta"));

                sheet.CellValue(3, 1, "Gamma");

                Assert.True(sheet.TryGetCellText(3, 1, out var appended));
                Assert.Equal("Gamma", appended);
                Assert.Equal("A3", sheet.FindFirst("Gamma"));
                Assert.Equal("A2", sheet.FindFirst("Beta"));
                sheet.CellValue(2, 1, "Delta");
                Assert.Null(sheet.FindFirst("Beta"));
                Assert.Equal("A2", sheet.FindFirst("Delta"));
                sheet.CellValue(4, 1, "Beta");
                Assert.Equal("A4", sheet.FindFirst("Beta"));
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_TryGetCellText_UsesFreshSharedStringsAfterDirectOpenXmlMutation() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesSharedStringDirectMutation.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Alpha");

                Assert.True(sheet.TryGetCellText(1, 1, out var first));
                Assert.Equal("Alpha", first);

                var sharedString = document._spreadSheetDocument.WorkbookPart!.SharedStringTablePart!.SharedStringTable!
                    .Elements<SharedStringItem>()
                    .First();
                sharedString.Text!.Text = "Omega";

                Assert.True(sheet.TryGetCellText(1, 1, out var refreshed));
                Assert.Equal("Omega", refreshed);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_CellValue_RejectsSharedStringsOverExcelLimit() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesSharedStringTooLong.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                string tooLong = new string('A', 32_768);

                var exception = Assert.Throws<ArgumentException>(() => sheet.CellValue(1, 1, tooLong));
                Assert.Contains("32,767", exception.Message, StringComparison.Ordinal);
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_FindFirst_UsesFreshWorksheetStateAfterDirectOpenXmlMutation() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesFindFirstDirectMutation.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Alpha");
                sheet.CellValue(2, 1, "Beta");

                Assert.Equal("A2", sheet.FindFirst("Beta"));
                Assert.Null(sheet.FindFirst("Gamma"));

                var worksheet = document._spreadSheetDocument.WorkbookPart!.WorksheetParts.First().Worksheet;
                var betaCell = worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A2");
                betaCell.DataType = CellValues.InlineString;
                betaCell.CellValue = null;
                betaCell.InlineString = new InlineString(new Text("Delta"));

                var sheetData = worksheet.GetFirstChild<SheetData>()!;
                var row = new Row { RowIndex = 3 };
                row.Append(new Cell {
                    CellReference = "A3",
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString(new Text("Gamma"))
                });
                sheetData.Append(row);

                Assert.Null(sheet.FindFirst("Beta"));
                Assert.Equal("A2", sheet.FindFirst("Delta"));
                Assert.Equal("A3", sheet.FindFirst("Gamma"));
            }

            File.Delete(filePath);
        }

        [Fact]
        public void Test_FindFirst_UsesFreshFormulaCacheStateAfterCalculationChanges() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesFindFirstFormulaCacheMutation.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, 1d);
                sheet.CellValue(2, 1, 2d);
                sheet.CellFormula(3, 1, "SUM(A1:A2)");

                Assert.Null(sheet.FindFirst("3"));

                Assert.Equal(1, sheet.RecalculateSupportedFormulas());
                Assert.Equal("A3", sheet.FindFirst("3"));

                sheet.ClearCachedFormulaResults();
                Assert.Null(sheet.FindFirst("3"));

                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_FindFirstAndReplaceAll_HandleSharedStrings() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesSharedStringFindReplace.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Status New");
                sheet.CellValue(2, 1, "Status Old");
                sheet.CellValue(3, 1, "Status New");
                sheet.CellValue(4, 1, 123);

                Assert.Equal("A2", sheet.FindFirst("old"));
                Assert.Equal(2, sheet.ReplaceAll("new", "Processed"));
                Assert.True(sheet.TryGetCellText(1, 1, out var first));
                Assert.True(sheet.TryGetCellText(3, 1, out var third));
                Assert.Equal("Status Processed", first);
                Assert.Equal("Status Processed", third);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ReplaceAll_HandlesInlineStrings() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesInlineReplace.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                var worksheet = document._spreadSheetDocument.WorkbookPart!.WorksheetParts.First().Worksheet;
                var sheetData = worksheet.GetFirstChild<SheetData>()!;
                var row = new Row { RowIndex = 1 };
                row.Append(new Cell {
                    CellReference = "A1",
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString(new Text("Status New"))
                });
                row.Append(new Cell {
                    CellReference = "B1",
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString(new Run(new Text("Status ")), new Run(new Text("New")))
                });
                sheetData.Append(row);

                Assert.Equal(2, sheet.ReplaceAll("new", "Processed"));
                Assert.True(sheet.TryGetCellText(1, 1, out var first));
                Assert.True(sheet.TryGetCellText(1, 2, out var second));
                Assert.Equal("Status Processed", first);
                Assert.Equal("Status Processed", second);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ReplaceAll_DoesNotReadFormulaTextFromStringCellWithoutCachedValue() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesReplaceAllFormulaStringNoCache.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                var worksheet = document._spreadSheetDocument.WorkbookPart!.WorksheetParts.First().Worksheet;
                var sheetData = worksheet.GetFirstChild<SheetData>()!;
                var row = new Row { RowIndex = 1 };
                var cell = new Cell {
                    CellReference = "A1",
                    DataType = CellValues.String,
                    CellFormula = new CellFormula("CONCAT(\"ReviewTarget\",\"\")")
                };
                row.Append(cell);
                sheetData.Append(row);

                Assert.Equal(0, sheet.ReplaceAll("ReviewTarget", "Changed"));
                Assert.Equal(CellValues.String, cell.DataType!.Value);
                Assert.Equal("CONCAT(\"ReviewTarget\",\"\")", cell.CellFormula!.Text);
                Assert.Null(cell.CellValue);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_TryGetCellText_ReadsFormulaTextWhenCellHasNoCachedValue() {
            string filePath = Path.Combine(_directoryWithFiles, "CellValuesTryGetFormulaNoCache.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                var worksheet = document._spreadSheetDocument.WorkbookPart!.WorksheetParts.First().Worksheet;
                var sheetData = worksheet.GetFirstChild<SheetData>()!;
                var row = new Row { RowIndex = 1 };
                var cell = new Cell {
                    CellReference = "A1",
                    CellFormula = new CellFormula("CONCAT(\"ReviewTarget\",\"\")")
                };
                row.Append(cell);
                sheetData.Append(row);

                Assert.True(sheet.TryGetCellText(1, 1, out string? text));
                Assert.Equal("CONCAT(\"ReviewTarget\",\"\")", text);
                Assert.Equal(0, sheet.ReplaceAll("ReviewTarget", "Changed"));
                Assert.Equal("CONCAT(\"ReviewTarget\",\"\")", cell.CellFormula!.Text);
                Assert.Null(cell.CellValue);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }
    }
}
