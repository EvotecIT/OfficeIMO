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
            var sheet = document.AddWorksheet("Random");
            sheet.CellValue(1, 1, lower); sheet.CellValue(2, 1, upper);
            sheet.CellFormula(1, 2, "RANDBETWEEN(A1,A2)");
            sheet.CellFormula(2, 2, "B1-B1");
            for (int run = 0; run < 8; run++) {
                Assert.Equal(2, document.Calculate());
                double value = sheet.CellAt(1, 2).GetValue<double>();
                Assert.InRange(value, lower, upper);
                Assert.Equal(Math.Truncate(value), value);
                AssertFormulaCache(sheet, 2, 2, "0");
            }
            double expected = sheet.CellAt(1, 2).GetValue<double>();
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 2).GetValue<double>());
            Assert.Equal("RANDBETWEEN(A1,A2)", reopened.Sheets[0].GetFormulaText(1, 2));
        }

        [Fact]
        public void Test_FormulaEvaluator_RandomBetweenLargeReversedAndUnqualifiedBounds() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Limits");
            sheet.CellFormula(1, 1, "RANDBETWEEN(9007199254740991,-9007199254740991)");
            sheet.CellFormula(2, 1, "IFERROR(A1,99)");
            foreach (int row in new[] { 3, 4 }) sheet.CellValue(row, 1, 42d);
            sheet.CellFormula(3, 1, "IFERROR(RANDBETWEEN(-9007199254740992,1),99)");
            sheet.CellFormula(4, 1, "IFERROR(RANDBETWEEN(1,9007199254740992),99)");
            Assert.Equal(2, document.Calculate());
            AssertFormulaCache(sheet, 1, 1, "#NUM!"); AssertFormulaCache(sheet, 2, 1, "99");
            AssertFormulaCache(sheet, 3, 1, "42"); AssertFormulaCache(sheet, 4, 1, "42");
        }
    }
}
