using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelCombinationsTests {
        [Theory]
        [InlineData("COMBIN(8,2)", "28")]
        [InlineData("COMBIN(8.9,2.1)", "28")]
        [InlineData("COMBIN(0,0)", "1")]
        [InlineData("COMBIN(52,5)", "2598960")]
        [InlineData("COMBIN(8,6)+1", "29")]
        [InlineData("COMBIN(-1,0)", "#NUM!")]
        [InlineData("COMBIN(2,-0.1)", "#NUM!")]
        [InlineData("COMBIN(2,3)", "#NUM!")]
        [InlineData("COMBIN(2048,1024)", "#NUM!")]
        [InlineData("COMBIN(2000,500)", "#NUM!")]
        [InlineData("IFERROR(COMBIN(2048,1024),99)", "99")]
        [InlineData("COMBIN(8,#REF!)", "#REF!")]
        public void Combinations_replace_stale_caches_and_preserve_saved_typed_results(string formula, string expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Combinations");
            sheet.CellValue(1, 1, 42d); sheet.CellFormula(1, 1, formula);
            Assert.Equal(1, document.Calculate());
            Assert.Equal(expected, Assert.Single(sheet.InspectFormulas().Formulas).CachedValue);
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.Equal(formula, reopened.Sheets[0].GetFormulaText(1, 1));
            Assert.Equal(expected, Assert.Single(reopened.InspectFormulas().Formulas).CachedValue);
            Assert.Empty(reopened.ValidateOpenXml());
        }

        [Theory]
        [InlineData("COMBIN(8)")]
        [InlineData("COMBIN(8,2,1)")]
        [InlineData("COMBIN(8,)")]
        [InlineData("COMBIN(A1:A2,2)")]
        [InlineData("COMBIN(8,\"text\")")]
        [InlineData("IFERROR(COMBIN(1E20,2),99)")]
        [InlineData("IFERROR(COMBIN(1E20,#REF!),99)")]
        [InlineData("COMBIN(1E20,-1)")]
        [InlineData("COMBIN(1E20,1E21)")]
        public void Unqualified_combinations_preserve_existing_cache(string formula) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Combinations");
            sheet.CellValue(1, 1, 8d); sheet.CellValue(2, 1, 2d);
            sheet.CellValue(1, 2, 42d); sheet.CellFormula(1, 2, formula);
            Assert.Equal(0, document.Calculate());
            Assert.Equal("42", Assert.Single(sheet.InspectFormulas().Formulas).CachedValue);
        }

        [Theory]
        [InlineData(100d, 50d, 1.008913445455642e29)]
        [InlineData(1028d, 514d, 7.156051054877897e307)]
        [InlineData(9007199254740991d, 1d, 9007199254740991d)]
        public void Combinations_recalculate_large_finite_results_after_operand_edits(double n, double k, double expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Combinations");
            sheet.CellValue(1, 1, 8d); sheet.CellValue(2, 1, 2d); sheet.CellFormula(1, 2, "COMBIN(A1,A2)");
            Assert.Equal(1, document.Calculate()); Assert.Equal(28d, sheet.CellAt(1, 2).GetValue<double>());
            sheet.CellValue(1, 1, n); sheet.CellValue(2, 1, k);
            Assert.Equal(1, document.Calculate());
            Assert.InRange(Math.Abs(sheet.CellAt(1, 2).GetValue<double>() / expected - 1), 0, 1e-13);
            Assert.Contains("COMBIN", document.InspectFormulas().Capabilities.SupportedFunctions);
        }
    }
}
