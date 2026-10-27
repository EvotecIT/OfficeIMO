using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelInspectionRawTokenContractTests {
        [Fact]
        public void InspectionSnapshot_AbsentCellTypesUseTheSameRawTokenForStoredAndCommentOnlyCells() {
            using MemoryStream stream = CreateWorkbook();
            using ExcelDocument document = ExcelDocument.Load(stream, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            ExcelWorksheetInfo sheet = document.Inspect().Worksheets[0];
            ExcelCellInfo numeric = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
            ExcelCellInfo commentOnly = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2);
            ExcelCellInfo empty = Assert.Single(sheet.Cells, cell => cell.Row == 2 && cell.Column == 1);
            ExcelCellInfo boolean = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 6);

            Assert.Equal("n", numeric.StoredValueType);
            Assert.Equal(numeric.StoredValueType, commentOnly.StoredValueType);
            Assert.Equal(numeric.StoredValueType, empty.StoredValueType);
            Assert.Equal(numeric.StoredValueType, new ExcelCellInfo().StoredValueType);
            Assert.True(numeric.HasStoredValue);
            Assert.Equal("7", numeric.StoredValue);
            Assert.False(commentOnly.HasStoredValue);
            Assert.Null(commentOnly.StoredValue);
            Assert.Null(commentOnly.Value);
            Assert.NotNull(commentOnly.Comment);
            Assert.False(empty.HasStoredValue);
            Assert.Null(empty.StoredValue);
            Assert.Equal("b", boolean.StoredValueType);
            Assert.True(Assert.IsType<bool>(boolean.Value));
        }

        [Fact]
        public void InspectionSnapshot_VisibilityDefaultMatchesRawWorksheetStateTokens() {
            using MemoryStream stream = CreateWorkbook();
            using ExcelDocument document = ExcelDocument.Load(stream, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            ExcelWorkbookInfo snapshot = document.Inspect();

            Assert.Equal("visible", snapshot.Worksheets[0].VisibilityState);
            Assert.Equal(snapshot.Worksheets[0].VisibilityState, new ExcelWorksheetInfo().VisibilityState);
            Assert.Equal(new[] { "visible", "visible", "hidden", "veryHidden" }, snapshot.Worksheets.Select(sheet => sheet.VisibilityState));
            Assert.Equal(new[] { false, false, true, true }, snapshot.Worksheets.Select(sheet => sheet.Hidden));
        }

        [Fact]
        public void InspectionSnapshot_FormulaKindRequiresAFormulaAndPreservesRawFormulaTypeTokens() {
            using MemoryStream stream = CreateWorkbook();
            using ExcelDocument document = ExcelDocument.Load(stream, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            ExcelWorksheetInfo sheet = document.Inspect().Worksheets[0];

            Assert.All(sheet.Cells.Where(cell => cell.Column is 1 or 2 or 6), cell => Assert.Null(cell.FormulaKind));
            ExcelCellInfo normal = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 3);
            Assert.Equal("normal", normal.FormulaKind);
            Assert.Equal("SUM(A1)", normal.StoredFormula);
            Assert.Equal("SUM(A1)", normal.Formula);
            Assert.All(sheet.Cells.Where(cell => cell.Column == 4), cell => Assert.Equal("shared", cell.FormulaKind));
            ExcelCellInfo sharedFollower = Assert.Single(sheet.Cells, cell => cell.Row == 2 && cell.Column == 4);
            Assert.Equal(string.Empty, sharedFollower.StoredFormula);
            Assert.Equal("SUM(A2)", sharedFollower.Formula);
            Assert.Equal("array", Assert.Single(sheet.Cells, cell => cell.Column == 5).FormulaKind);
        }

        private static MemoryStream CreateWorkbook() {
            var stream = new MemoryStream();
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)) {
                WorkbookPart workbookPart = package.AddWorkbookPart();
                workbookPart.Workbook = new Workbook(new Sheets());
                SheetStateValues?[] states = { null, SheetStateValues.Visible, SheetStateValues.Hidden, SheetStateValues.VeryHidden };
                for (int index = 0; index < states.Length; index++) {
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var row = new Row { RowIndex = 1U };
                    row.Append(new Cell { CellReference = "A1", CellValue = new CellValue("7") });
                    if (index == 0) {
                        row.Append(new Cell { CellReference = "C1", CellFormula = new CellFormula("SUM(A1)"), CellValue = new CellValue("7") });
                        row.Append(new Cell {
                            CellReference = "D1",
                            CellFormula = new CellFormula("SUM(A1)") { FormulaType = CellFormulaValues.Shared, SharedIndex = 0U, Reference = "D1:D2" },
                            CellValue = new CellValue("7"),
                        });
                        row.Append(new Cell {
                            CellReference = "E1",
                            CellFormula = new CellFormula("SUM(A1)") { FormulaType = CellFormulaValues.Array, Reference = "E1" },
                            CellValue = new CellValue("7"),
                        });
                        row.Append(new Cell { CellReference = "F1", DataType = CellValues.Boolean, CellValue = new CellValue("1") });
                    }

                    var sheetData = new SheetData(row);
                    if (index == 0) {
                        sheetData.Append(new Row(
                            new Cell { CellReference = "A2" },
                            new Cell { CellReference = "D2", CellFormula = new CellFormula { FormulaType = CellFormulaValues.Shared, SharedIndex = 0U }, CellValue = new CellValue("0") }) {
                            RowIndex = 2U,
                        });
                        WorksheetCommentsPart commentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>();
                        commentsPart.Comments = new Comments(
                            new Authors(new Author("Reviewer")),
                            new CommentList(new Comment {
                                Reference = "B1", AuthorId = 0U,
                                CommentText = new CommentText(new Text("Comment-only cell")),
                            }));
                    }

                    worksheetPart.Worksheet = new Worksheet(sheetData);
                    var sheet = new Sheet { Name = "Sheet" + index, SheetId = (uint)(index + 1), Id = workbookPart.GetIdOfPart(worksheetPart) };
                    if (states[index].HasValue) sheet.State = states[index]!.Value;
                    workbookPart.Workbook.Sheets!.Append(sheet);
                }
            }
            stream.Position = 0;
            return stream;
        }
    }
}
