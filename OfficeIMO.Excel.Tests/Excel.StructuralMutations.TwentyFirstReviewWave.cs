using System;
using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_MutationSnapshot_NonSeekablePayloadStopsAtRemainingBudget() {
            using var source = new CountingNonSeekableReadStream(length: 1_000_000);

            Assert.Throws<InvalidOperationException>(() =>
                ExcelSheet.ReadMutationSnapshotPayload(source, 32, 128));

            Assert.InRange(source.BytesRead, 33, 100_000);
        }

        [Fact]
        public void Test_CellShiftPlans_RejectFormulaReferencesBeyondWorksheetCapacity() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellFormula(2, 2, "XFD1+A1048576");

            Assert.Throws<InvalidOperationException>(() =>
                sheet.PlanInsertCells("A1", ExcelCellShiftDirection.Right));
            Assert.Throws<InvalidOperationException>(() =>
                sheet.PlanInsertCells("A1", ExcelCellShiftDirection.Down));
        }

        [Fact]
        public void Test_StructuralMutationPlans_RejectRemovedWorksheetAcrossEveryKind() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            document.AddWorksheet("Survivor").CellValue(1, 1, "Unchanged");
            sheet.CellValue(1, 1, "A");
            sheet.CellValue(2, 1, "B");
            ExcelStructuralMutationPlan[] plans = {
                sheet.PlanInsertColumns(1),
                sheet.PlanDeleteColumns(1),
                sheet.PlanInsertCells("A1", ExcelCellShiftDirection.Right),
                sheet.PlanInsertCells("A1", ExcelCellShiftDirection.Down),
                sheet.PlanDeleteCells("A1", ExcelCellShiftDirection.Left),
                sheet.PlanDeleteCells("A1", ExcelCellShiftDirection.Up),
                sheet.PlanCopyRange("A1:A2", "C1"),
                sheet.PlanMoveRange("A1:A2", "C1"),
                sheet.PlanTransposeRange("A1:A2", "C1")
            };
            document.RemoveWorksheet("Data");

            foreach (ExcelStructuralMutationPlan plan in plans) {
                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => plan.Apply());
                Assert.Contains("no longer part of the workbook", exception.Message, StringComparison.Ordinal);
            }
            ExcelSheet survivor = Assert.Single(document.Sheets);
            Assert.Equal("Unchanged", survivor.CellAt(1, 1).GetValue<string>());
        }

        [Fact]
        public void Test_QuotedThreeDimensionalReferences_ParseAndBlockUnsafeStructuralEdits() {
            ExcelFormulaReferenceSyntax parsed = Assert.IsType<ExcelFormulaReferenceSyntax>(
                Assert.Single(
                    ExcelFormulaSyntaxTree.Parse("SUM('First''s':'Last''s'!A1)").Nodes,
                    node => node is ExcelFormulaReferenceSyntax));
            Assert.Equal("'First''s':'Last''s'", parsed.Reference.Qualifier);

            using var document = ExcelDocument.Create();
            document.AddWorksheet("First Sheet");
            ExcelSheet middle = document.AddWorksheet("Middle");
            document.AddWorksheet("Last Sheet");
            ExcelSheet formulas = document.AddWorksheet("Formulas");
            formulas.CellFormula(1, 1, "SUM('First Sheet':'Last Sheet'!A1)");

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                middle.PlanInsertColumns(1));
            Assert.Contains("3-D reference", exception.Message, StringComparison.Ordinal);
        }
    }
}
