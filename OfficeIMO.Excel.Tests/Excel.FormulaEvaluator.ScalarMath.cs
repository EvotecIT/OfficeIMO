using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelScalarMathTests {
        [Theory]
        [InlineData("SQRT(-1)", "#NUM!")]
        [InlineData("MOD(1,0)", "#DIV/0!")]
        [InlineData("IFERROR(SQRT(-1),99)", "99")]
        [InlineData("IFERROR(MOD(1,0),99)", "99")]
        [InlineData("1+SQRT(-1)", "#NUM!")]
        [InlineData("MOD(1e308,1e-308)", "#NUM!")]
        [InlineData("INT(#REF!)", "#REF!")]
        [InlineData("EVEN(-1.5)", "-2")]
        [InlineData("EVEN(0)", "0")]
        [InlineData("EVEN(2)", "2")]
        [InlineData("ODD(-2)", "-3")]
        [InlineData("ODD(0)", "1")]
        [InlineData("ODD(3)", "3")]
        [InlineData("1+EVEN(1.5)", "3")]
        [InlineData("ODD(#REF!)", "#REF!")]
        [InlineData("IFERROR(EVEN(1/0),99)", "99")]
        public void Scalar_math_replaces_stale_caches_with_typed_results(string formula, string expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Math");
            sheet.CellFormula(1, 1, "42");
            Assert.Equal(1, document.Calculate());
            sheet.CellFormula(1, 1, formula);
            Assert.Equal(1, document.Calculate());
            Assert.Equal(expected, Assert.Single(sheet.InspectFormulas().Formulas).CachedValue);
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.Equal(formula, reopened.Sheets[0].GetFormulaText(1, 1));
            Assert.Equal(expected, Assert.Single(reopened.InspectFormulas().Formulas).CachedValue);
            Assert.Empty(reopened.ValidateOpenXml());
        }

        [Theory]
        [InlineData("INT(1,\"text\")")]
        [InlineData("SQRT(4,\"text\")")]
        [InlineData("MOD(1,2,\"text\")")]
        [InlineData("MOD(A1:A2)")]
        [InlineData("EVEN(1,2)")]
        [InlineData("ODD()")]
        [InlineData("ODD(A1:A2)")]
        [InlineData("EVEN(\"text\")")]
        public void Scalar_math_does_not_flatten_ranges_or_discard_arguments(string formula) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Math");
            sheet.CellValue(1, 1, 3d); sheet.CellValue(2, 1, 2d);
            sheet.CellFormula(1, 2, formula);
            Assert.Equal(0, document.Calculate());
            Assert.Equal(formula, sheet.GetFormulaText(1, 2));
        }

        [Theory]
        [InlineData("MOD(#REF!,A1:A2)")]
        [InlineData("MOD(#REF!,\"text\")")]
        [InlineData("MOD(A1:A2,#REF!)")]
        [InlineData("IFERROR(MOD(#REF!,A1:A2),99)")]
        public void Scalar_math_qualifies_all_operands_before_propagating_errors(string formula) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Math");
            sheet.CellValue(1, 1, 3d); sheet.CellValue(2, 1, 2d);
            sheet.CellFormula(1, 2, "42");
            Assert.Equal(1, document.Calculate());
            sheet.CellFormula(1, 2, formula);
            Assert.Equal(0, document.Calculate());
            Assert.Equal("42", Assert.Single(sheet.InspectFormulas().Formulas).CachedValue);
            Assert.Equal(formula, sheet.GetFormulaText(1, 2));
        }

        [Fact]
        public void Scalar_math_retains_existing_boolean_numeric_coercion() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Math");
            sheet.CellValue(1, 1, true);
            sheet.CellFormula(1, 2, "SQRT(TRUE)");
            sheet.CellFormula(2, 2, "INT(A1)");
            sheet.CellFormula(3, 2, "MOD(2,1=1)");
            Assert.Equal(3, document.Calculate());
            Assert.Equal(1d, sheet.CellAt(1, 2).GetValue<double>());
            Assert.Equal(1d, sheet.CellAt(2, 2).GetValue<double>());
            Assert.Equal(0d, sheet.CellAt(3, 2).GetValue<double>());
        }

        [Theory]
        [InlineData("INT(A1)", -9d, 2d)]
        [InlineData("MOD(A1,A2)", 1.1d, 0.9d)]
        [InlineData("EVEN(A1)", -10d, 4d)]
        [InlineData("ODD(A1)", -9d, 3d)]
        [InlineData("SQRT(ABS(A1))", 2.9832867780352594d, 1.70293863659264d)]
        public void Scalar_math_follows_edited_operands(string formula, double first, double second) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Math");
            sheet.CellValue(1, 1, -8.9d); sheet.CellValue(2, 1, 2d);
            sheet.CellFormula(1, 2, formula);
            Assert.Equal(1, document.Calculate());
            Assert.Equal(first, sheet.CellAt(1, 2).GetValue<double>(), 12);
            sheet.CellValue(1, 1, 2.9d);
            Assert.Equal(1, document.Calculate());
            Assert.Equal(second, sheet.CellAt(1, 2).GetValue<double>(), 12);
        }
    }
}
