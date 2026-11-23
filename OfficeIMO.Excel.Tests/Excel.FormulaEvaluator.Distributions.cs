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

        // Far-tail quantiles: P(χ² ≤ x) for one degree of freedom is about √(2x/π), so CHISQ.INV(1e-100,1) is π/2·1e-200.
        [Theory]
        [InlineData("CHISQ.INV(1E-100,1)", 1.5707963267948966e-200)]
        [InlineData("T.INV(1E-100,1)", -3.1830988618379067e99)]
        [InlineData("T.DIST.RT(1E160,1)", 3.1830988618379067e-161)]
        [InlineData("T.INV(1E-200,1)", -3.1830988618379067e199)]
        [InlineData("T.DIST(0,1E20,FALSE)", 0.3989422804014327)]
        [InlineData("T.DIST(1,1E20,FALSE)", 0.24197072451914337)]
        [InlineData("T.DIST(0.5,300,FALSE)", 0.35164388573796046)]
        public void Test_FormulaEvaluator_DistributionInversesReachFarTails(string formula, double expected) {
            double actual = CalculateNumber(formula);
            Assert.True(Math.Abs(actual / expected - 1) < 1e-9, $"{formula} returned {actual}, expected {expected}.");
        }

        // Near the median the one-degree t distribution is the Cauchy distribution: F(t) − ½ = atan(t)/π.
        [Fact]
        public void Test_FormulaEvaluator_StudentTKeepsItsDistanceFromTheMedian() {
            double above = CalculateNumber("T.DIST(1E-9,1,TRUE)") - 0.5;
            Assert.True(Math.Abs(above / 3.183098861837907e-10 - 1) < 1e-6, $"T.DIST(1E-9,1,TRUE) − 0.5 = {above}.");
            double quantile = CalculateNumber("T.INV(0.500000001,1)");
            Assert.True(Math.Abs(quantile / 3.141592653589793e-9 - 1) < 1e-6, $"T.INV(0.500000001,1) = {quantile}.");
        }

        // With huge parameters the beta continued fraction's coefficients overflow; the call must end promptly.
        [Fact]
        public void Test_FormulaEvaluator_StudentTWithHugeParametersFinishes() {
            Assert.Equal(1, CalculateNumber("T.DIST(1E100,1E200,TRUE)"));
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
