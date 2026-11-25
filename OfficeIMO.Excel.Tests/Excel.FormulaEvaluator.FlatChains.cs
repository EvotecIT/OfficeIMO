using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        // A long run of equal-precedence operators must not need stack in proportion to its length.
        [Theory]
        [InlineData("+", "4000")]
        [InlineData("*", "1")]
        public void Test_FormulaEvaluator_LongFlatOperatorChainCalculatesOnASmallStack(string op, string expected) {
            string formula = string.Join(op, Enumerable.Repeat("1", 4000));
            string? value = null;
            int calculated = 0;
            var thread = new Thread(() => {
                using var document = ExcelDocument.Create();
                ExcelSheet sheet = document.AddSheet("Chain");
                sheet.Cell(1, 1).SetFormula(formula);
                calculated = document.Calculate();
                sheet.TryGetCachedFormulaValue(1, 1, out value);
            }, 256 * 1024);
            thread.Start();
            thread.Join();
            Assert.Equal(1, calculated);
            Assert.Equal(expected, value);
        }

        [Fact]
        public void Test_FormulaEvaluator_RepeatedPercentAppliesEachTime() {
            Assert.Equal("0.0001", CalculateTextOnNewSheet("100%%%"));
        }

        private static string? CalculateTextOnNewSheet(string formula) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Chain");
            sheet.Cell(1, 1).SetFormula(formula);
            Assert.Equal(1, document.Calculate());
            sheet.TryGetCachedFormulaValue(1, 1, out string? value);
            return value;
        }
    }
}
