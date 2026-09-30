using OfficeIMO.Excel;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void DynamicArrays_CrossSheetReferencesUseTheirOwnSpillsAndValues() {
            using var document = ExcelDocument.Create();
            var first = document.AddWorksheet("First");
            var second = document.AddWorksheet("Second");
            first.SetDynamicArrayFormula("G1", "SEQUENCE(2,2)");
            second.CellValue(2, 7, 99);
            second.SetDynamicArrayFormula("J1", "SEQUENCE(2,2,10)");
            first.CellFormula(1, 1, "Second!G2");
            first.CellFormula(1, 2, "Second!K2");
            second.CellFormula(1, 1, "First!G2");
            document.RecalculateSupportedFormulas();
            Assert.True(first.TryGetCellText(1, 1, out var scalar));
            Assert.Equal("99", scalar);
            Assert.True(first.TryGetCellText(1, 2, out var spill));
            Assert.Equal("13", spill);
            Assert.True(second.TryGetCellText(1, 1, out var reverse));
            Assert.Equal("3", reverse);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void DynamicArrays_UncachedScalarFormulaBlocksSpill() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Data");
            sheet.CellFormula(2, 8, "40+2");
            sheet.SetDynamicArrayFormula("G1", "SEQUENCE(2,2)");
            sheet.CellFormula(1, 1, "H2");
            document.RecalculateSupportedFormulas();
            Assert.True(sheet.TryGetCellText(1, 7, out var anchor));
            Assert.Equal("#SPILL!", anchor);
            Assert.True(sheet.TryGetCellText(2, 8, out var blocker));
            Assert.Equal("42", blocker);
            Assert.True(sheet.TryGetCellText(1, 1, out var dependent));
            Assert.Equal("42", dependent);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void DynamicArrays_ColumnEditsMoveOwnershipAndRetainResizeEvidence(bool deleting) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 2);
            sheet.SetDynamicArrayFormula("G1", "SEQUENCE(A1,2)");
            document.RecalculateSupportedFormulas();
            sheet.CellValue(3, 1, 0); // Initialize the guard before the edit.
            if (deleting) sheet.DeleteColumns(3); else sheet.InsertColumns(3);
            int anchorColumn = deleting ? 6 : 8;
            int childColumn = anchorColumn + 1;
            Assert.Throws<InvalidOperationException>(() => sheet.CellValue(2, childColumn, 99));
            sheet.CellValue(2, deleting ? 8 : 7, 99);
            sheet.CellValue(1, 1, 3);
            document.RecalculateSupportedFormulas();
            Assert.True(sheet.TryGetCellText(3, childColumn, out var grown));
            Assert.Equal("6", grown);
            sheet.CellValue(1, 1, 1);
            document.RecalculateSupportedFormulas();
            Assert.False(sheet.TryGetCellText(3, childColumn, out _));
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void DynamicArrays_RollbackRestoresOwnershipAndResizeEvidence() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 2);
            sheet.SetDynamicArrayFormula("G1", "SEQUENCE(A1,2)");
            document.RecalculateSupportedFormulas();
            sheet.CellValue(3, 1, 0);
            Assert.Throws<InvalidOperationException>(() => sheet.ApplyTransactionalMutation(_ => {
                sheet.InsertColumns(3);
                throw new InvalidOperationException("Rollback probe");
            }, new ExcelMutationPlanOptions(), CancellationToken.None));
            Assert.Throws<InvalidOperationException>(() => sheet.CellValue(2, 8, 99));
            sheet.CellValue(2, 9, 99);
            sheet.CellValue(1, 1, 3);
            document.RecalculateSupportedFormulas();
            Assert.True(sheet.TryGetCellText(3, 8, out var grown));
            Assert.Equal("6", grown);
            Assert.Empty(document.ValidateOpenXml());
        }

    }
}
