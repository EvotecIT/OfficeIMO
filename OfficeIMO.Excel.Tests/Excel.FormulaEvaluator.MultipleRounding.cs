using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelMultipleRoundingTests {
        [Theory]
        [InlineData("CEILING(-2.5,-2)", -4d)]
        [InlineData("FLOOR(-2.5,-2)", -2d)]
        [InlineData("CEILING(-2.5,2)", -2d)]
        [InlineData("FLOOR(-2.5,2)", -4d)]
        [InlineData("CEILING(0.234,0.01)", 0.24d)]
        [InlineData("FLOOR(0.234,0.01)", 0.23d)]
        [InlineData("CEILING(0.07,0.01)", 0.07d)]
        [InlineData("FLOOR(0.29,0.01)", 0.29d)]
        [InlineData("CEILING(1.5,0)", 0d)]
        [InlineData("FLOOR(0,)", 0d)]
        [InlineData("CEILING(1.5,)", 0d)]
        [InlineData("CEILING(1E30,1E29)", 1E30)]
        [InlineData("FLOOR(1E-300,1E-301)", 1E-300)]
        [InlineData("CEILING(-1E30,-1E29)", -1E30)]
        [InlineData("FLOOR(-1E-300,-1E-301)", -1E-300)]
        [InlineData("CEILING(1E300,1E-300)", 1E300)]
        [InlineData("FLOOR(1E300,1E-300)", 1E300)]
        [InlineData("CEILING(1E-300,1E300)", 1E300)]
        [InlineData("FLOOR(1E-300,1E300)", 0d)]
        [InlineData("CEILING(0.1+0.2,0.1)", 0.3d)]
        [InlineData("IFERROR(FLOOR(1,0),99)", 99d)]
        public void Multiple_rounding_survives_saved_numeric_caches(string formula, double expected) {
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
        [InlineData("CEILING(2.1E-28,1E-28)", 3E-28)]
        [InlineData("FLOOR(2.1E-28,1E-28)", 2E-28)]
        [InlineData("CEILING(-2.1E-28,-1E-28)", -3E-28)]
        [InlineData("FLOOR(-2.1E-28,-1E-28)", -2E-28)]
        public void Very_small_multiple_rounding_keeps_the_correct_multiple(string formula, double expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Math");
            sheet.CellFormula(1, 1, formula);
            Assert.Equal(1, document.Calculate());
            Assert.InRange(Math.Abs(sheet.CellAt(1, 1).GetValue<double>() - expected), 0, Math.Abs(expected) * 1E-15);
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.InRange(Math.Abs(reopened.Sheets[0].CellAt(1, 1).GetValue<double>() - expected), 0, Math.Abs(expected) * 1E-15);
            Assert.Empty(reopened.ValidateOpenXml());
        }

        [Theory]
        [InlineData("CEILING(2.5,-2)", "#NUM!")]
        [InlineData("FLOOR(2.5,-2)", "#NUM!")]
        [InlineData("FLOOR(1,0)", "#DIV/0!")]
        [InlineData("FLOOR(1,)", "#DIV/0!")]
        [InlineData("CEILING(1.6E308,1E308)", "#NUM!")]
        [InlineData("CEILING(#REF!,2)+1", "#REF!")]
        [InlineData("FLOOR(1,#N/A)", "#N/A")]
        public void Multiple_rounding_propagates_typed_errors_in_saved_caches(string formula, string expected) {
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
        [InlineData("CEILING(1)")]
        [InlineData("FLOOR(1,2,\"text\")")]
        [InlineData("CEILING(1,\"text\")")]
        [InlineData("FLOOR(A1:A2,1)")]
        [InlineData("CEILING(#REF!,A1:A2)")]
        [InlineData("IFERROR(FLOOR(#REF!,\"text\"),99)")]
        public void Unsupported_multiple_rounding_preserves_existing_caches(string formula) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Math");
            sheet.CellValue(1, 1, 1d); sheet.CellValue(2, 1, 2d);
            sheet.CellValue(1, 2, 42d); sheet.CellFormula(1, 2, formula);
            Assert.Equal(0, document.Calculate());
            Assert.Equal("42", Assert.Single(sheet.InspectFormulas().Formulas).CachedValue);
        }
    }
}
