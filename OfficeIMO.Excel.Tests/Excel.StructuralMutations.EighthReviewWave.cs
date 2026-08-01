using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using Xnsv = DocumentFormat.OpenXml.Office2021.Excel.NamedSheetViews;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(ExcelCellShiftDirection.Right, true)]
        [InlineData(ExcelCellShiftDirection.Down, true)]
        [InlineData(ExcelCellShiftDirection.Left, false)]
        [InlineData(ExcelCellShiftDirection.Up, false)]
        public void Test_CellShift_RejectsPartiallyIntersectingCrossSheetReferences(
            ExcelCellShiftDirection direction,
            bool inserting) {
            using var document = ExcelDocument.Create();
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellValue(1, 1, 1);
            data.CellValue(1, 2, 2);
            data.CellValue(2, 1, 3);
            data.CellValue(2, 2, 4);
            summary.CellFormula(1, 1, "SUM(Data!A1:B2)");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => {
                if (inserting) data.PlanInsertCells("A1", direction);
                else data.PlanDeleteCells("A1", direction);
            });

            Assert.Contains("partially overlapping", exception.Message);
            Assert.Equal("SUM(Data!A1:B2)", Assert.Single(summary.GetFormulaCells()).Formula);
        }

        [Theory]
        [InlineData(ExcelCellShiftDirection.Down, "A5", "A1", "A6")]
        [InlineData(ExcelCellShiftDirection.Right, "E1", "A1", "F1")]
        public void Test_CellShift_RemapsAndProtectsCellBackedConnectionParameters(
            ExcelCellShiftDirection direction,
            string parameterCell,
            string insertionCell,
            string shiftedCell) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            Parameter parameter = AttachCellBackedConnection(document, sheet, parameterCell);

            sheet.InsertCells(insertionCell, direction);
            Assert.Equal(shiftedCell, parameter.Cell!.Value);

            ExcelCellShiftDirection deletionDirection = direction == ExcelCellShiftDirection.Down
                ? ExcelCellShiftDirection.Up
                : ExcelCellShiftDirection.Left;
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.PlanDeleteCells(shiftedCell, deletionDirection));
            Assert.Contains("connection parameter", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(shiftedCell, parameter.Cell!.Value);
        }

        [Fact]
        public void Test_RangeMutations_RemapAndRemoveLegacyCommentVml() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.SetComment(1, 1, "Move me", author: "Tester");
            VmlDrawingPart vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
            int[] before = ReadCommentVmlAnchor(vmlPart);

            sheet.MoveRange("A1", "B1");

            Assert.False(sheet.HasComment(1, 1));
            Assert.True(sheet.HasComment(1, 2));
            int[] after = ReadCommentVmlAnchor(vmlPart);
            Assert.Equal(before[0] + 1, after[0]);
            Assert.Equal(before[4] + 1, after[4]);

            sheet.DeleteCells("B1", ExcelCellShiftDirection.Left);

            Assert.Null(sheet.WorksheetPart.WorksheetCommentsPart);
            Assert.Empty(sheet.WorksheetPart.VmlDrawingParts);
        }

        [Fact]
        public void Test_StructuralMutations_RemapInternalHyperlinkTargets() {
            using var document = ExcelDocument.Create();
            ExcelSheet data = document.AddWorksheet("Data");
            ExcelSheet summary = document.AddWorksheet("Summary");
            data.CellValue(2, 2, "Target");
            var hyperlink = new Hyperlink { Reference = "D1", Location = "Data!B2" };
            summary.WorksheetPart.Worksheet.Append(new Hyperlinks(hyperlink));

            data.InsertColumns(2);
            Assert.Equal("Data!C2", hyperlink.Location!.Value);

            data.InsertCells("C2", ExcelCellShiftDirection.Down);
            Assert.Equal("Data!C3", hyperlink.Location!.Value);

            data.MoveRange("C3", "E5");
            Assert.Equal("Data!E5", hyperlink.Location!.Value);
        }

        [Fact]
        public void Test_StructuralMutations_RemapNamedSheetViewFilterRanges() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            Xnsv.NsvFilter filter = AddNamedSheetViewFilter(sheet, "A1:C10");

            sheet.DeleteColumns(3);
            Assert.Equal("A1:B10", filter.Ref!.Value);

            filter.Ref = "A1:A10";
            sheet.InsertCells("A1", ExcelCellShiftDirection.Down);
            Assert.Equal("A2:A11", filter.Ref!.Value);

            sheet.MoveRange("A2:A11", "C2");
            Assert.Equal("C2:C11", filter.Ref!.Value);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void Test_FormulaInspection_HonorsPackageRecalculationFlags(bool workbookLevel) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 1);
            sheet.CellFormula(1, 2, "A1*2");
            Cell formulaCell = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "B1");
            formulaCell.CellValue = new CellValue("2");
            formulaCell.CellFormula!.CalculateCell = false;
            if (workbookLevel) {
                CalculationProperties calculation = document.WorkbookRoot.GetFirstChild<CalculationProperties>()
                    ?? document.WorkbookRoot.AppendChild(new CalculationProperties());
                calculation.FullCalculationOnLoad = true;
                calculation.ForceFullCalculation = true;
            } else {
                sheet.WorksheetPart.Worksheet.Append(new SheetCalculationProperties {
                    FullCalculationOnLoad = true
                });
            }

            ExcelFormulaCellInfo formula = Assert.Single(sheet.GetFormulaCells());

            Assert.True(formula.HasCachedValue);
            Assert.True(formula.IsDirty);
            Assert.True(formula.State.HasFlag(ExcelFormulaState.Dirty));
            Assert.True(formula.State.HasFlag(ExcelFormulaState.Deferred));
            Assert.False(formula.State.HasFlag(ExcelFormulaState.Evaluated));
        }
    }
}
