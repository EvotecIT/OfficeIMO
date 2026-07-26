using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_StructuralRows_RejectsWorkbookOwnedPivotConsolidationDeletion() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellAt(5, 1).SetValue(1);
            sheet.CellAt(6, 1).SetValue(2);
            PivotTableCacheDefinitionPart cachePart =
                document.WorkbookPartRoot.AddNewPart<PivotTableCacheDefinitionPart>();
            var rangeSet = new RangeSet {
                Sheet = "Data",
                Reference = "A5:A6"
            };
            cachePart.PivotCacheDefinition = new PivotCacheDefinition(
                new CacheSource(
                    new Consolidation(
                        new RangeSets(rangeSet) { Count = 1U })) {
                    Type = SourceValues.Consolidation
                });

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.DeleteRows(5, 2));

            Assert.Contains("complete consolidation source range", exception.Message);
            Assert.Equal("A5:A6", rangeSet.Reference!.Value);
            Assert.Equal(1, sheet.CellAt(5, 1).GetValue<int>());
            Assert.Equal(2, sheet.CellAt(6, 1).GetValue<int>());
        }

        [Fact]
        public void Test_StructuralRows_PreservesBackslashPrefixedDefinedNames() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellFormula(1, 1, @"\A5+1");

            sheet.InsertRows(5);
            Assert.Equal(@"\A5+1", sheet.GetFormulaText(1, 1));

            sheet.DeleteRows(5);
            Assert.Equal(@"\A5+1", sheet.GetFormulaText(1, 1));
        }
    }
}
