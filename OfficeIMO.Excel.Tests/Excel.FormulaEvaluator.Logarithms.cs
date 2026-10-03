using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelLogarithmTests {
        [Theory]
        [InlineData("EXP(1)", Math.E)]
        [InlineData("LN(EXP(3))", 3d)]
        [InlineData("LOG(100)", 2d)]
        [InlineData("LOG(8,2)", 3d)]
        [InlineData("LOG(5.0625,1.5)", 4d)]
        [InlineData("LOG(8,0.5)", -3d)]
        [InlineData("LOG10(10^5)", 5d)]
        [InlineData("IFERROR(LOG(8,1),99)", 99d)]
        public void Logarithms_and_exponential_results_survive_saved_numeric_caches(string formula, double expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Math");
            sheet.CellFormula(1, 1, formula);
            Assert.Equal(1, document.Calculate());
            Assert.Equal(expected, sheet.CellAt(1, 1).GetValue<double>(), 12);
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.Equal(formula, reopened.Sheets[0].GetFormulaText(1, 1));
            Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 1).GetValue<double>(), 12);
            Assert.Empty(reopened.ValidateOpenXml());
        }

        [Theory]
        [InlineData("LN(0)", "#NUM!")]
        [InlineData("LOG10(-1)", "#NUM!")]
        [InlineData("LOG(0)", "#NUM!")]
        [InlineData("LOG(8,0)", "#NUM!")]
        [InlineData("LOG(8,-2)", "#NUM!")]
        [InlineData("LOG(8,1)", "#DIV/0!")]
        [InlineData("EXP(1000)", "#NUM!")]
        [InlineData("LOG(#REF!,2)+1", "#REF!")]
        [InlineData("LOG(8,#N/A)", "#N/A")]
        public void Logarithms_propagate_typed_errors_and_replace_stale_numeric_caches(string formula, string expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Math");
            sheet.CellValue(1, 1, 42d); sheet.CellFormula(1, 1, formula);
            Assert.Equal(1, document.Calculate());
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.Equal(expected, Assert.Single(reopened.InspectFormulas().Formulas).CachedValue);
            Assert.Empty(reopened.ValidateOpenXml());
        }

        [Theory]
        [InlineData("EXP(1,\"text\")")]
        [InlineData("LN(1,\"text\")")]
        [InlineData("LOG10(10,\"text\")")]
        [InlineData("LOG(10,2,\"text\")")]
        [InlineData("LOG(10,)")]
        [InlineData("LOG(#REF!,A1:A2)")]
        [InlineData("IFERROR(LOG(#REF!,\"text\"),99)")]
        public void Unsupported_logarithm_operands_and_arities_preserve_existing_caches(string formula) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Math");
            sheet.CellValue(1, 1, 1d); sheet.CellValue(2, 1, 2d);
            sheet.CellValue(1, 2, 42d); sheet.CellFormula(1, 2, formula);
            Assert.Equal(0, document.Calculate());
            Assert.Equal("42", Assert.Single(sheet.InspectFormulas().Formulas).CachedValue);
        }
    }
}
