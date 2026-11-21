using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelInspectionCellValueTests {
        public static IEnumerable<object[]> Workbooks() {
            yield return new object[] { "cell-kinds" };
            foreach (string corpus in new[] { "SpreadsheetXmlCorpus", "ExcelFormulaCorpus" }) {
                string directory = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", corpus);
                foreach (string file in Directory.GetFiles(directory, "*.xlsx").OrderBy(name => name, StringComparer.Ordinal)) {
                    yield return new object[] { corpus + "/" + Path.GetFileName(file) };
                }
            }
        }

        [Theory]
        [MemberData(nameof(Workbooks))]
        public void InspectedCellValuesMatchAnIndependentWorkbookReader(string fixture) {
            byte[] bytes = fixture == "cell-kinds"
                ? CreateCellKindsWorkbook()
                : File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", fixture));
            var readOptions = new ExcelReaderOptions { UseCachedFormulaResult = true, TreatDatesUsingNumberFormat = true };
            var expected = new Dictionary<(string Sheet, int Row, int Column), object?>();
            using (ExcelDocumentReader reader = ExcelDocumentReader.Open(bytes, readOptions)) {
                foreach (string sheetName in reader.GetSheetNames()) {
                    foreach (ExcelCellValueInfo cell in reader.GetSheet(sheetName).EnumerateCells()) {
                        expected[(sheetName, cell.Row, cell.Column)] = cell.Value;
                    }
                }
            }

            using var stream = new MemoryStream(bytes);
            using ExcelDocument document = ExcelDocument.Load(stream, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            int compared = 0;
            foreach (ExcelWorksheetInfo sheet in document.Inspect(readOptions).Worksheets) {
                foreach (ExcelCellInfo cell in sheet.Cells) {
                    expected.TryGetValue((sheet.Name, cell.Row, cell.Column), out object? value);
                    Assert.True(Equals(value, cell.Value), $"{sheet.Name}!R{cell.Row}C{cell.Column}: expected '{value}', inspected '{cell.Value}'.");
                    if (value != null) compared++;
                }
            }

            Assert.True(compared > 0, "The workbook has no cell values to compare.");
        }

        [Fact]
        public void InspectionReportsRichTextRunsOnlyForSharedStringsThatHaveThem() {
            using var stream = new MemoryStream(CreateCellKindsWorkbook());
            using ExcelDocument document = ExcelDocument.Load(stream, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            ExcelWorksheetInfo sheet = Assert.Single(document.Inspect().Worksheets);

            ExcelRichTextRun[] rich = sheet.Cells.Single(cell => cell.Row == 1 && cell.Column == 2).RichTextRuns.ToArray();
            Assert.Equal(new[] { "Bold", " plain" }, rich.Select(run => run.Text));
            Assert.True(rich[0].Bold);
            Assert.Empty(sheet.Cells.Single(cell => cell.Row == 1 && cell.Column == 1).RichTextRuns);
        }

        /// <summary>One worksheet holding each stored cell kind the reader converts.</summary>
        private static byte[] CreateCellKindsWorkbook() {
            var stream = new MemoryStream();
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)) {
                WorkbookPart workbookPart = package.AddWorkbookPart();
                workbookPart.Workbook = new Workbook(new Sheets());

                var sharedStrings = workbookPart.AddNewPart<SharedStringTablePart>();
                sharedStrings.SharedStringTable = new SharedStringTable(
                    new SharedStringItem(new Text("Plain shared")),
                    new SharedStringItem(
                        new Run(new RunProperties(new Bold()), new Text("Bold")),
                        new Run(new Text(" plain") { Space = SpaceProcessingModeValues.Preserve })),
                    new SharedStringItem(new Text("   ") { Space = SpaceProcessingModeValues.Preserve }));

                var styles = workbookPart.AddNewPart<WorkbookStylesPart>();
                styles.Stylesheet = new Stylesheet(
                    new Fonts(new Font()),
                    new Fills(new Fill(new PatternFill { PatternType = PatternValues.None })),
                    new Borders(new Border()),
                    new CellFormats(
                        new CellFormat { NumberFormatId = 0U },
                        new CellFormat { NumberFormatId = 14U, ApplyNumberFormat = true }));

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var row1 = new Row { RowIndex = 1U };
                row1.Append(
                    new Cell { CellReference = "A1", DataType = CellValues.SharedString, CellValue = new CellValue("0") },
                    new Cell { CellReference = "B1", DataType = CellValues.SharedString, CellValue = new CellValue("1") },
                    new Cell { CellReference = "C1", DataType = CellValues.SharedString, CellValue = new CellValue("2") },
                    new Cell { CellReference = "D1", CellValue = new CellValue("42.5") },
                    new Cell { CellReference = "E1", StyleIndex = 1U, CellValue = new CellValue("45292") },
                    new Cell { CellReference = "F1", DataType = CellValues.Boolean, CellValue = new CellValue("1") },
                    new Cell { CellReference = "G1", DataType = CellValues.Error, CellValue = new CellValue("#N/A") },
                    new Cell {
                        CellReference = "H1",
                        DataType = CellValues.InlineString,
                        InlineString = new InlineString(new Text("Inline"))
                    },
                    new Cell {
                        CellReference = "I1",
                        DataType = CellValues.InlineString,
                        InlineString = new InlineString(
                            new Run(new RunProperties(new Italic()), new Text("Rich")),
                            new Run(new Text(" inline") { Space = SpaceProcessingModeValues.Preserve }))
                    });
                var row2 = new Row { RowIndex = 2U };
                row2.Append(
                    new Cell { CellReference = "A2", CellFormula = new CellFormula("D1*2"), CellValue = new CellValue("85") },
                    new Cell { CellReference = "B2", DataType = CellValues.String, CellFormula = new CellFormula("\"a\"&\"b\""), CellValue = new CellValue("ab") },
                    new Cell { CellReference = "C2", CellFormula = new CellFormula("D1+1") },
                    new Cell { CellReference = "D2", StyleIndex = 1U, CellFormula = new CellFormula("E1+1"), CellValue = new CellValue("45293") },
                    new Cell { CellReference = "E2", CellValue = new CellValue(string.Empty) },
                    new Cell { CellReference = "F2", DataType = CellValues.Date, CellValue = new CellValue("2024-02-03T04:05:06") });
                worksheetPart.Worksheet = new Worksheet(new SheetData(row1, row2));
                workbookPart.Workbook.Sheets!.Append(new Sheet {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1U,
                    Name = "Kinds"
                });
            }

            return stream.ToArray();
        }
    }
}
