using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_TableRename_RewritesStableAndDisplayNameAliases() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Amount");
            sheet.CellValue(2, 1, 10);
            sheet.AddTable("A1:A2", true, "Internal", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
            Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table!;
            table.DisplayName = "Sales";
            table.Save();
            sheet.CellFormula(4, 1, "SUM(Internal[Amount])");
            sheet.CellFormula(5, 1, "SUM(Sales[Amount])");

            Assert.Equal("Orders", sheet.RenameTable("Sales", "Orders"));

            Assert.Equal("Orders", table.Name!.Value);
            Assert.Equal("Orders", table.DisplayName!.Value);
            ExcelFormulaCellInfo[] formulas = sheet.GetFormulaCells().ToArray();
            Assert.Equal("SUM(Orders[Amount])", formulas.Single(item => item.CellReference == "A4").Formula);
            Assert.Equal("SUM(Orders[Amount])", formulas.Single(item => item.CellReference == "A5").Formula);
        }

        [Fact]
        public void Test_FormulaSearch_DoesNotTreatStructuredEscapeAsQuotedQualifier() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellFormula(1, 1, "Table1['#Data]+SUM(A2)");
            sheet.CellFormula(2, 1, "'SUM(A1)'!B2");

            ExcelFormulaCellInfo match = Assert.Single(sheet.SearchFormulas(
                new ExcelFormulaSearchOptions { Function = "SUM" }));

            Assert.Equal("A1", match.CellReference);
        }

        [Fact]
        public void Test_CellDeletion_ClampsDrawingMarkersInsideDeletedBlock() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet oneCellSheet = CreateCellShiftDrawingSheet(document, "One cell");
            Xdr.OneCellAnchor oneCell = Assert.Single(
                oneCellSheet.WorksheetPart.DrawingsPart!.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>());
            oneCell.FromMarker!.ColumnId!.Text = "2";
            oneCell.FromMarker.RowId!.Text = "1";

            oneCellSheet.DeleteCells("B2:C2", ExcelCellShiftDirection.Left);

            Assert.Equal("1", oneCell.FromMarker.ColumnId.Text);
            Assert.Equal("1", oneCell.FromMarker.RowId.Text);

            ExcelSheet twoCellSheet = CreateCellShiftDrawingSheet(document, "Two cell");
            Xdr.TwoCellAnchor twoCell = ReplaceCellShiftDrawingWithTwoCellAnchor(twoCellSheet);

            twoCellSheet.DeleteCells("B2:C2", ExcelCellShiftDirection.Left);

            Assert.Equal("1", twoCell.FromMarker!.ColumnId!.Text);
            Assert.Equal("2", twoCell.ToMarker!.ColumnId!.Text);
            Assert.True(int.Parse(twoCell.FromMarker.ColumnId.Text) <= int.Parse(twoCell.ToMarker.ColumnId.Text));

            ExcelSheet verticalSheet = CreateCellShiftDrawingSheet(document, "Vertical");
            Xdr.TwoCellAnchor vertical = ReplaceCellShiftDrawingWithTwoCellAnchor(verticalSheet);
            vertical.FromMarker!.ColumnId!.Text = "1";
            vertical.ToMarker!.ColumnId!.Text = "1";
            vertical.FromMarker.RowId!.Text = "2";
            vertical.ToMarker.RowId!.Text = "4";

            verticalSheet.DeleteCells("B2:B3", ExcelCellShiftDirection.Up);

            Assert.Equal("1", vertical.FromMarker.RowId.Text);
            Assert.Equal("2", vertical.ToMarker.RowId.Text);
            Assert.True(int.Parse(vertical.FromMarker.RowId.Text) <= int.Parse(vertical.ToMarker.RowId.Text));
        }

        private static ExcelSheet CreateCellShiftDrawingSheet(ExcelDocument document, string name) {
            ExcelSheet sheet = document.AddWorksheet(name);
            sheet.CellValue(1, 1, "Category");
            sheet.CellValue(1, 2, "Value");
            sheet.CellValue(2, 1, "One");
            sheet.CellValue(2, 2, 1);
            sheet.AddChartFromRange("A1:B2", row: 5, column: 4);
            return sheet;
        }

        private static Xdr.TwoCellAnchor ReplaceCellShiftDrawingWithTwoCellAnchor(ExcelSheet sheet) {
            Xdr.WorksheetDrawing drawing = sheet.WorksheetPart.DrawingsPart!.WorksheetDrawing!;
            Xdr.OneCellAnchor oneCell = Assert.Single(drawing.Elements<Xdr.OneCellAnchor>());
            var twoCell = new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("2"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("1"),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("4"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("1"),
                    new Xdr.RowOffset("0")),
                oneCell.GetFirstChild<Xdr.GraphicFrame>()!.CloneNode(true),
                new Xdr.ClientData()) {
                EditAs = Xdr.EditAsValues.TwoCell
            };
            oneCell.Remove();
            drawing.Append(twoCell);
            return twoCell;
        }
    }
}
