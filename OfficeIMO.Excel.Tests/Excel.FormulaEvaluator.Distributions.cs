using System.Globalization;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        // Expected values are the examples in Microsoft's function reference, compared to the digits it prints.
        [Theory]
        [InlineData("NORM.DIST(42,40,1.5,TRUE)", 0.9087888, 7)]
        [InlineData("NORM.DIST(42,40,1.5,FALSE)", 0.10934, 5)]
        [InlineData("NORM.INV(0.908789,40,1.5)", 42.000002, 6)]
        [InlineData("NORM.S.DIST(1.333333,TRUE)", 0.908788726, 9)]
        [InlineData("NORM.S.DIST(1.333333,FALSE)", 0.164010148, 9)]
        [InlineData("NORM.S.INV(0.908789)", 1.3333347, 7)]
        [InlineData("T.DIST(60,1,TRUE)", 0.99469533, 8)]
        [InlineData("T.DIST(8,3,FALSE)", 0.00073691, 8)]
        [InlineData("T.DIST.2T(1.959999998,60)", 0.054644930, 9)]
        [InlineData("T.DIST.RT(1.959999998,60)", 0.027322465, 9)]
        [InlineData("T.INV(0.75,2)", 0.8164966, 7)]
        [InlineData("T.INV.2T(0.546449,60)", 0.606533, 6)]
        [InlineData("CHISQ.DIST(0.5,1,TRUE)", 0.52049988, 8)]
        [InlineData("CHISQ.DIST(2,3,FALSE)", 0.20755375, 8)]
        [InlineData("CHISQ.DIST(3,2,TRUE)", 0.77686984, 8)]
        [InlineData("CHISQ.DIST.RT(18.307,10)", 0.0500006, 7)]
        [InlineData("CHISQ.INV(0.93,1)", 3.283020287, 9)]
        [InlineData("CHISQ.INV(0.6,2)", 1.832581464, 9)]
        [InlineData("CHISQ.INV.RT(0.050001,10)", 18.306973, 6)]
        [InlineData("NORMSDIST(1.333333)", 0.908788726, 9)]
        [InlineData("TDIST(1.959999998,60,2)", 0.054644930, 9)]
        [InlineData("TINV(0.546449,60)", 0.606533, 6)]
        [InlineData("CHIDIST(18.307,10)", 0.0500006, 7)]
        [InlineData("CHIINV(0.050001,10)", 18.306973, 6)]
        // Large parameters need far more expansion terms; P(a, a) tends to 1/2 + 1/(3√(2πa)).
        [InlineData("CHISQ.DIST(1000000,1000000,TRUE)", 0.5001880631, 8)]
        [InlineData("T.DIST(0.5,1000000,TRUE)", 0.6914624, 6)]
        public void Test_FormulaEvaluator_DistributionsMatchDocumentedExamples(string formula, double expected, int digits) {
            Assert.Equal(expected, CalculateNumber(formula), digits);
        }

        [Theory]
        [InlineData("NORM.DIST(1,0,0,TRUE)")]
        [InlineData("NORM.INV(1,0,1)")]
        [InlineData("NORM.S.INV(0)")]
        [InlineData("T.DIST(1,0,TRUE)")]
        [InlineData("T.DIST.2T(-1,2)")]
        [InlineData("T.INV(0.5,0)")]
        [InlineData("T.INV.2T(1.5,3)")]
        [InlineData("CHISQ.DIST(-1,2,TRUE)")]
        [InlineData("CHISQ.INV(1.5,2)")]
        [InlineData("CHISQ.INV.RT(0.5,0)")]
        public void Test_FormulaEvaluator_DistributionArgumentsOutsideTheDomainAreNum(string formula) {
            Assert.Equal("#NUM!", CalculateText(formula));
        }

        [Theory]
        [InlineData(0.001, 1)]
        [InlineData(0.025, 3)]
        [InlineData(0.5, 7)]
        [InlineData(0.975, 30)]
        [InlineData(0.9999, 120)]
        public void Test_FormulaEvaluator_DistributionInversesRoundTrip(double probability, int degrees) {
            string p = probability.ToString(CultureInfo.InvariantCulture);
            string k = degrees.ToString(CultureInfo.InvariantCulture);
            Assert.Equal(probability, CalculateNumber($"NORM.S.DIST(NORM.S.INV({p}),TRUE)"), 12);
            Assert.Equal(probability, CalculateNumber($"T.DIST(T.INV({p},{k}),{k},TRUE)"), 12);
            Assert.Equal(probability, CalculateNumber($"CHISQ.DIST(CHISQ.INV({p},{k}),{k},TRUE)"), 12);
            Assert.Equal(probability, CalculateNumber($"CHISQ.DIST.RT(CHISQ.INV.RT({p},{k}),{k})"), 12);
        }

        [Fact]
        public void Test_FormulaEvaluator_DistributionArgumentErrorsPropagate() {
            Assert.Equal("#DIV/0!", CalculateText("NORM.DIST(1/0,0,1,TRUE)"));
        }

        private static double CalculateNumber(string formula) =>
            double.Parse(CalculateText(formula), NumberStyles.Float, CultureInfo.InvariantCulture);

        private static string CalculateText(string formula) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddSheet("Calc");
            sheet.Cell(1, 1).SetFormula(formula);
            Assert.Equal(1, document.Calculate());
            Assert.True(sheet.TryGetCachedFormulaValue(1, 1, out string? value));
            return value!;
        }
    }
}
