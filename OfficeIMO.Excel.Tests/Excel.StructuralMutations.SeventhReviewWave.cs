using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using ExcelTableStyle = OfficeIMO.Excel.ExcelTableStyle;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_RangeMove_RejectsPartiallyOverlappingFormulaReference(bool crossSheet) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            ExcelSheet formulaSheet = crossSheet ? document.AddSheet("Summary") : sheet;
            sheet.Cell(1, 1).SetValue(1);
            sheet.Cell(1, 2).SetValue(2);
            formulaSheet.Cell(3, 1).SetFormula(crossSheet ? "SUM(Data!A1:B1)" : "SUM(A1:B1)");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.PlanMoveRange("A1", "D1"));

            Assert.Contains("partially overlapping", exception.Message);
            Assert.Equal(1d, sheet.Cell(1, 1).GetValue<double>());
            Assert.Equal(crossSheet ? "SUM(Data!A1:B1)" : "SUM(A1:B1)", Assert.Single(formulaSheet.GetFormulaCells()).Formula);
        }

        [Fact]
        public void Test_RangeMove_RemapsCellBackedConnectionParameter() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(5, 1).SetValue(1);
            Parameter parameter = AttachCellBackedConnection(document, sheet, "A5");

            sheet.Range("A5").MoveTo("B5");

            Assert.Equal("B5", parameter.Cell!.Value);
        }

        [Fact]
        public void Test_CellInsertion_PreflightsDrawingMarkerCapacity() {
            using (var document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddSheet("Data");
                AppendDrawingMarker(sheet, row: 1, column: ExcelA1.MaxColumns);

                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                    () => sheet.PlanInsertCells("A1", ExcelCellShiftDirection.Right));

                Assert.Contains("column limit", exception.Message);
            }

            using (var document = ExcelDocument.Create()) {
                ExcelSheet sheet = document.AddSheet("Data");
                AppendDrawingMarker(sheet, row: ExcelA1.MaxRows, column: 1);

                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                    () => sheet.PlanInsertCells("A1", ExcelCellShiftDirection.Down));

                Assert.Contains("row limit", exception.Message);
            }
        }

        [Fact]
        public void Test_RangeBasedTableWrapper_RemainsStableAfterResizeAndSchemaChange() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("A");
            sheet.Cell(1, 2).SetValue("B");
            sheet.Cell(2, 1).SetValue(1);
            sheet.Cell(2, 2).SetValue(2);
            sheet.AddTable("A1:B2", true, "Sales", ExcelTableStyle.TableStyleMedium2);
            sheet.Cell(1, 5).SetValue("X");
            sheet.Cell(1, 6).SetValue("Y");
            sheet.Cell(2, 5).SetValue(1);
            sheet.Cell(2, 6).SetValue(2);
            sheet.AddTable("E1:F2", true, "Costs", ExcelTableStyle.TableStyleMedium2);

            ExcelTable table = sheet.Table("A1:B2")
                .Resize("A1:C2")
                .SetSchema(new[] { "A", "B", "C" }, "A1:C3")
                .SetStyle(ExcelTableStyle.TableStyleMedium4);
            ExcelTable schemaTable = sheet.Table("E1:F2")
                .SetSchema(new[] { "Net", "Tax", "Gross" }, "E1:G2")
                .SetStyle(ExcelTableStyle.TableStyleMedium4);

            Assert.Equal("Sales", table.NameOrRange);
            Assert.Equal("A1:C3", table.Range);
            Assert.Equal("Costs", schemaTable.NameOrRange);
            Assert.Equal("E1:G2", schemaTable.Range);
            Assert.Equal(new[] { "A", "B", "C" }, document.GetTables().Single(item => item.Name == "Sales").Columns.Select(column => column.Name));
            Assert.Equal(new[] { "Net", "Tax", "Gross" }, document.GetTables().Single(item => item.Name == "Costs").Columns.Select(column => column.Name));
        }

        [Fact]
        public void Test_IgnoredErrorRemoval_PreservesDisjointReferencesInSameRegion() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            ExcelIgnoredErrorKind errors = ExcelIgnoredErrorKind.NumberStoredAsText | ExcelIgnoredErrorKind.FormulaRange;
            sheet.AddIgnoredErrorRegion(new[] { "A1", "Z1" }, errors);

            Assert.Equal(1, sheet.RemoveIgnoredErrorRegions("A1"));

            ExcelIgnoredErrorRegionInfo remaining = Assert.Single(sheet.GetIgnoredErrorRegions());
            Assert.Equal(new[] { "Z1" }, remaining.Ranges);
            Assert.Equal(errors, remaining.Errors);
        }

        [Fact]
        public void Test_TableRename_PreservesLetAndLambdaLexicalBindings() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Amount");
            sheet.Cell(2, 1).SetValue(1);
            sheet.AddTable("A1:A2", true, "Input", ExcelTableStyle.TableStyleMedium2);
            sheet.Cell(4, 1).SetFormula("LET(Input,1,Input)");
            sheet.Cell(5, 1).SetFormula("Input+LET(Input,1,Input)");
            sheet.Cell(6, 1).SetFormula("LET(Input,1,SUM(Input[Amount])+Input)");
            sheet.Cell(7, 1).SetFormula("LAMBDA(Input,Input)(1)");

            sheet.RenameTable("Input", "Ledger");

            string[] formulas = sheet.GetFormulaCells().OrderBy(item => item.CellReference).Select(item => item.Formula).ToArray();
            Assert.Equal("LET(Input,1,Input)", formulas[0]);
            Assert.Equal("Ledger+LET(Input,1,Input)", formulas[1]);
            Assert.Equal("LET(Input,1,SUM(Ledger[Amount])+Input)", formulas[2]);
            Assert.Equal("LAMBDA(Input,Input)(1)", formulas[3]);
        }

        [Fact]
        public void Test_FormulaInspection_ToleratesOversizedCellMetadataIndex() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.SetArrayFormula("A1:A2", "ROW(A1:A2)");
            Cell owner = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "A1");
            owner.SetAttribute(new OpenXmlAttribute("", "cm", "", uint.MaxValue.ToString()));

            ExcelFormulaCellInfo formula = Assert.Single(sheet.GetFormulaCells());

            Assert.Equal(uint.MaxValue, formula.Array!.MetadataIndex);
            Assert.False(formula.Array.IsDynamic);
        }

        [Fact]
        public void Test_AutoFilterCriteria_NormalizeEquivalentLocalRanges() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Range("A1:C3").AutoFilter().Column(0).Blanks();

            sheet.Range("  $A$1:$C$3  ").AutoFilter().Column(2).TopBottom(5);

            ExcelAutoFilterInfo filter = Assert.Single(sheet.GetAutoFilters());
            Assert.Equal("A1:C3", filter.Range);
            Assert.Equal(new uint[] { 0U, 2U }, filter.Columns.Select(column => column.ColumnOffset).OrderBy(value => value));
        }

        private static void AppendDrawingMarker(ExcelSheet sheet, int row, int column) {
            DrawingsPart drawingsPart = sheet.WorksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing(
                new Xdr.FromMarker(
                    new Xdr.ColumnId((column - 1).ToString()),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId((row - 1).ToString()),
                    new Xdr.RowOffset("0")));
        }
    }
}
