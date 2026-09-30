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
            sheet.CellValue(1, 1, 1);
            document.RecalculateSupportedFormulas();
            Assert.True(sheet.TryGetCellText(1, anchorColumn, out var shrunk));
            Assert.Equal("1", shrunk);
            Assert.False(sheet.TryGetCellText(2, childColumn, out _));
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
            sheet.InsertColumns(3);
            sheet.CellValue(1, 1, 1);
            document.RecalculateSupportedFormulas();
            Assert.True(sheet.TryGetCellText(1, 8, out var anchor));
            Assert.Equal("1", anchor);
            Assert.False(sheet.TryGetCellText(2, 9, out _));
            Assert.Empty(document.ValidateOpenXml());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void DynamicArrays_LoadedSpillRetainsOwnershipAcrossWorksheetMutations(bool unrelatedSheet) {
            using var stream = new System.IO.MemoryStream();
            using (var original = ExcelDocument.Create()) {
                var sheet = original.AddWorksheet("Data");
                original.AddWorksheet("Other");
                sheet.CellValue(1, 1, 2);
                sheet.SetDynamicArrayFormula("G1", "SEQUENCE(A1,2)");
                original.RecalculateSupportedFormulas();
                original.Save(stream);
            }
            stream.Position = 0;
            using var document = ExcelDocument.Load(stream);
            var data = document["Data"];
            (unrelatedSheet ? document["Other"] : data).InsertColumns(3);
            data.CellValue(1, 1, 1);
            document.RecalculateSupportedFormulas();
            int anchorColumn = unrelatedSheet ? 7 : 8;
            Assert.True(data.TryGetCellText(1, anchorColumn, out var anchor));
            Assert.Equal("1", anchor);
            Assert.False(data.TryGetCellText(2, anchorColumn + 1, out _));
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void DynamicArrays_ColumnMutationDoesNotAdoptExternallyChangedLoadedChild() {
            using var stream = new System.IO.MemoryStream();
            using (var original = ExcelDocument.Create()) {
                var sheet = original.AddWorksheet("Data");
                sheet.CellValue(1, 1, 2);
                sheet.SetDynamicArrayFormula("G1", "SEQUENCE(A1,2)");
                original.RecalculateSupportedFormulas();
                original.Save(stream);
            }
            stream.Position = 0;
            using var document = ExcelDocument.Load(stream);
            var data = document["Data"];
            var child = Assert.Single(data.WorksheetPart.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>(),
                cell => cell.CellReference?.Value == "H2");
            child.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue("99");
            data.InsertColumns(3);
            data.CellValue(1, 1, 1);
            document.RecalculateSupportedFormulas();
            Assert.True(data.TryGetCachedFormulaValue(1, 8, out var error));
            Assert.Equal("#SPILL!", error);
            Assert.True(data.TryGetCellText(2, 9, out var retained));
            Assert.Equal("99", retained);
        }

    }
}
