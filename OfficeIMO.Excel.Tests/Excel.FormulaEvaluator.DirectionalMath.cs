using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelDirectionalMathTests {
        [Theory]
        [InlineData("TRUNC(0.29,2)", 0.29d)]
        [InlineData("ROUNDDOWN(-0.29,2)", -0.29d)]
        [InlineData("ROUNDUP(0.07,2)", 0.07d)]
        [InlineData("TRUNC(-8.9)", -8d)]
        [InlineData("ROUNDUP(-3.14159,1)", -3.2d)]
        [InlineData("ROUNDDOWN(31415.92654,-2)", 31400d)]
        [InlineData("TRUNC(-1234.5,-2)", -1200d)]
        [InlineData("SIGN(4-4)", 0d)]
        [InlineData("TRUNC(1e308,1)", 1e308d)]
        [InlineData("ROUNDDOWN(1e308,1)", 1e308d)]
        [InlineData("ROUNDUP(1e308,1)", 1e308d)]
        [InlineData("ROUNDUP(1e-30,2)", .01d)]
        [InlineData("ROUNDUP(1e-28,-15)", 1e15d)]
        [InlineData("ROUNDUP(5e-324,-15)", 1e15d)]
        public void Directional_math_keeps_decimal_boundaries_and_finite_saved_caches(string formula, double expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Math");
            sheet.CellFormula(1, 1, formula);
            Assert.Equal(1, document.Calculate());
            Assert.Equal(expected, sheet.CellAt(1, 1).GetValue<double>());
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.Equal(formula, reopened.Sheets[0].GetFormulaText(1, 1));
            Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
            Assert.Empty(reopened.ValidateOpenXml());
        }

        [Theory]
        [InlineData("SIGN(1,\"text\")")]
        [InlineData("TRUNC(1,2,\"text\")")]
        [InlineData("ROUNDUP(1,2,\"text\")")]
        [InlineData("ROUNDDOWN(1,2,\"text\")")]
        [InlineData("TRUNC(#REF!,A1:A2)")]
        [InlineData("IFERROR(TRUNC(#REF!,16),99)")]
        public void Directional_math_preserves_caches_for_unsupported_operands_and_arities(string formula) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Math");
            sheet.CellValue(1, 1, 1d); sheet.CellValue(2, 1, 2d);
            sheet.CellValue(1, 2, 42d); sheet.CellFormula(1, 2, formula);
            Assert.Equal(0, document.Calculate());
            Assert.Equal("42", Assert.Single(sheet.InspectFormulas().Formulas).CachedValue);
        }
    }
}
