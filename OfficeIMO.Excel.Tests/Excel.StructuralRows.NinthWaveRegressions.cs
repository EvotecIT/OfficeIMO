using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_MaterializesSharedFormulasWithImplicitCellReferences() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            SheetData sheetData = sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            var firstRow = new Row { RowIndex = 2U };
            firstRow.Append(
                new Cell { CellReference = "A2", CellValue = new CellValue(1) },
                new Cell {
                    CellFormula = new CellFormula("A2*2") {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 71U,
                        Reference = "B2:B3"
                    }
                });
            var secondRow = new Row { RowIndex = 3U };
            secondRow.Append(
                new Cell { CellReference = "A3", CellValue = new CellValue(2) },
                new Cell {
                    CellFormula = new CellFormula {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 71U
                    }
                });
            sheetData.Append(firstRow, secondRow);

            sheet.InsertRows(2);

            CellFormula[] formulas = sheet.WorksheetPart.Worksheet.Descendants<CellFormula>().ToArray();
            Assert.Equal(new[] { "A3*2", "A4*2" }, formulas.Select(formula => formula.Text).ToArray());
            Assert.All(formulas, formula => {
                Assert.Null(formula.FormulaType);
                Assert.Null(formula.SharedIndex);
                Assert.True(formula.CalculateCell!.Value);
            });
        }

        [Fact]
        public void Test_StructuralRows_DoesNotPreflightAbsoluteTwoCellDrawingAnchors() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(1, 1).SetValue("Category");
            sheet.CellAt(1, 2).SetValue("Value");
            sheet.CellAt(2, 1).SetValue("One");
            sheet.CellAt(2, 2).SetValue(1);
            sheet.AddChartFromRange("A1:B2", row: 5, column: 4);

            Xdr.WorksheetDrawing drawing = sheet.WorksheetPart.DrawingsPart!.WorksheetDrawing!;
            Xdr.OneCellAnchor oneCellAnchor = Assert.Single(drawing.Elements<Xdr.OneCellAnchor>());
            var absoluteAnchor = new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("3"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId((A1.MaxRows - 1).ToString()),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("8"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId((A1.MaxRows - 1).ToString()),
                    new Xdr.RowOffset("0")),
                oneCellAnchor.GetFirstChild<Xdr.GraphicFrame>()!.CloneNode(true),
                new Xdr.ClientData()) {
                EditAs = Xdr.EditAsValues.Absolute
            };
            oneCellAnchor.Remove();
            drawing.Append(absoluteAnchor);

            sheet.InsertRows(1);

            Assert.Equal((A1.MaxRows - 1).ToString(), absoluteAnchor.FromMarker!.RowId!.Text);
            Assert.Equal((A1.MaxRows - 1).ToString(), absoluteAnchor.ToMarker!.RowId!.Text);
        }

        [Fact]
        public void Test_StructuralRows_RewritesUnqualifiedExtendedChartFormulasOnlyOnEditedSheet() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet other = document.AddWorksheet("Other");
            data.CellAt(2, 1).SetValue(1);
            other.CellAt(2, 1).SetValue(1);

            Cx.Formula dataFormula = AppendExtendedChartFormula(data, "A2:A3");
            Cx.Formula otherFormula = AppendExtendedChartFormula(other, "A2:A3");

            data.InsertRows(2);

            Assert.Equal("A3:A4", dataFormula.Text);
            Assert.Equal("A2:A3", otherFormula.Text);
        }

        private static Cx.Formula AppendExtendedChartFormula(ExcelSheet sheet, string formulaText) {
            DrawingsPart drawingsPart = sheet.WorksheetPart.DrawingsPart
                ?? sheet.WorksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
            ExtendedChartPart chartPart = drawingsPart.AddNewPart<ExtendedChartPart>();
            var formula = new Cx.Formula(formulaText);
            chartPart.ChartSpace = new Cx.ChartSpace(formula);
            return formula;
        }
    }
}
