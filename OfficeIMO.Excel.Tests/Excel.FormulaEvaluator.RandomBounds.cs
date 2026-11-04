using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(2147483648d, 2147483648d)]
        [InlineData(-2147483649d, -2147483649d)]
        [InlineData(9007199254740991d, 9007199254740991d)]
        [InlineData(-9007199254740991d, -9007199254740991d)]
        [InlineData(-9007199254740991d, 9007199254740991d)]
        [InlineData(9007199254740988d, 9007199254740991d)]
        [InlineData(-9007199254740991d, -9007199254740988d)]
        public void Test_FormulaEvaluator_RandomBetweenExactIntegerBounds(double lower, double upper) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Random");
            sheet.Cell(1, 1).SetValue(lower); sheet.Cell(2, 1).SetValue(upper);
            sheet.Cell(1, 2).SetFormula("RANDBETWEEN(A1,A2)");
            sheet.Cell(2, 2).SetFormula("B1-B1");
            for (int run = 0; run < 8; run++) {
                Assert.Equal(2, document.Calculate());
                double value = sheet.Cell(1, 2).GetValue<double>();
                Assert.InRange(value, lower, upper);
                Assert.Equal(Math.Truncate(value), value);
                AssertFormulaCache(sheet, 2, 2, "0");
            }
            double expected = sheet.Cell(1, 2).GetValue<double>();
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.Equal(expected, reopened.Sheets[0].Cell(1, 2).GetValue<double>());
            Assert.Equal("RANDBETWEEN(A1,A2)", reopened.Sheets[0].GetFormulaText(1, 2));
        }

        [Fact]
        public void Test_FormulaEvaluator_RandomBetweenLargeReversedAndUnqualifiedBounds() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Limits");
            sheet.Cell(1, 1).SetFormula("RANDBETWEEN(9007199254740991,-9007199254740991)");
            sheet.Cell(2, 1).SetFormula("IFERROR(A1,99)");
            foreach (int row in new[] { 3, 4 }) sheet.Cell(row, 1).SetValue(42d);
            sheet.Cell(3, 1).SetFormula("IFERROR(RANDBETWEEN(-9007199254740992,1),99)");
            sheet.Cell(4, 1).SetFormula("IFERROR(RANDBETWEEN(1,9007199254740992),99)");
            Assert.Equal(2, document.Calculate());
            AssertFormulaCache(sheet, 1, 1, "#NUM!"); AssertFormulaCache(sheet, 2, 1, "99");
            AssertFormulaCache(sheet, 3, 1, "42"); AssertFormulaCache(sheet, 4, 1, "42");
        }
    }
}
