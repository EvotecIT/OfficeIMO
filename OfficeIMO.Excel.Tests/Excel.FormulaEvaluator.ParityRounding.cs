using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelParityRoundingTests {
        [Theory]
        [InlineData("EVEN", double.Epsilon, 2d)]
        [InlineData("EVEN", -double.Epsilon, -2d)]
        [InlineData("ODD", double.Epsilon, 1d)]
        [InlineData("ODD", -double.Epsilon, -1d)]
        [InlineData("EVEN", 9007199254740991d, 9007199254740992d)]
        [InlineData("EVEN", double.MaxValue, double.MaxValue)]
        [InlineData("ODD", 9007199254740990d, 9007199254740991d)]
        [InlineData("ODD", -9007199254740991d, -9007199254740991d)]
        public void Parity_rounding_preserves_exact_results_at_floating_point_boundaries(string function, double input, double expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Rounding");
            sheet.CellValue(1, 1, input);
            sheet.CellValue(1, 2, 42d);
            sheet.CellFormula(1, 2, function + "(A1)");
            Assert.Equal(1, document.Calculate());
            Assert.Equal(expected, sheet.CellAt(1, 2).GetValue<double>());
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 2).GetValue<double>());
            Assert.Equal(function + "(A1)", reopened.Sheets[0].GetFormulaText(1, 2));
            Assert.Contains(function, reopened.InspectFormulas().Capabilities.SupportedFunctions);
            Assert.Empty(reopened.ValidateOpenXml());
        }

        [Theory]
        [InlineData("ODD(A1)")]
        [InlineData("1+ODD(A1)")]
        [InlineData("IFERROR(ODD(A1),99)")]
        public void Unrepresentable_odd_results_preserve_caches_including_nested_and_error_fallback_paths(string formula) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Limits");
            sheet.CellValue(1, 1, 9007199254740992d);
            sheet.CellValue(1, 2, 42d); sheet.CellFormula(1, 2, formula);
            Assert.Equal(0, document.Calculate());
            Assert.Equal(42d, sheet.CellAt(1, 2).GetValue<double>());
            sheet.CellValue(1, 1, -9007199254740992d);
            Assert.Equal(0, document.Calculate());
            Assert.Equal(42d, sheet.CellAt(1, 2).GetValue<double>());
        }
    }
}
