using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelFactorialTests {
        [Theory]
        [InlineData(150d, 5.713383956445855e262)]
        [InlineData(170d, 7.257415615307999e306)]
        [InlineData(170.9d, 7.257415615307999e306)]
        public void Factorial_retains_finite_large_results_through_saved_output(double input, double expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Factorial");
            sheet.CellValue(1, 1, input);
            sheet.CellValue(1, 2, 42d); sheet.CellFormula(1, 2, "FACT(A1)");
            Assert.Equal(1, document.Calculate());
            Assert.InRange(Math.Abs(sheet.CellAt(1, 2).GetValue<double>() / expected - 1), 0, 1e-14);
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.Equal("FACT(A1)", reopened.Sheets[0].GetFormulaText(1, 2));
            Assert.InRange(Math.Abs(reopened.Sheets[0].CellAt(1, 2).GetValue<double>() / expected - 1), 0, 1e-14);
            Assert.Contains("FACT", reopened.InspectFormulas().Capabilities.SupportedFunctions);
            Assert.Empty(reopened.ValidateOpenXml());
        }
    }
}
