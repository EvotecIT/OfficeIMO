using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {

    public sealed class ExcelCashFlowTests {
        [Theory]
        [InlineData("excel-financial-results.xlsx")]
        [InlineData("excel-financial-coercions.xlsx")]
        public void Cash_flows_match_saved_native_Excel_results(string fixture) {
            string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "SpreadsheetXmlCorpus", fixture);
            using var document = ExcelDocument.Load(path);
            var expected = document.InspectFormulas().Formulas.ToDictionary(formula => formula.CellReference, formula => formula.CachedValue);
            Assert.NotEmpty(expected);
            Assert.Equal(expected.Count, document.Calculate());
            foreach (var formula in document.InspectFormulas().Formulas) {
                string? before = expected[formula.CellReference];
                if (double.TryParse(before, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double numeric)) {
                    Assert.True(double.TryParse(formula.CachedValue, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double actual), formula.Formula);
                    Assert.True(Math.Abs(numeric - actual) <= Math.Max(1, Math.Abs(numeric)) * 1e-8,
                        formula.Formula + ": expected " + before + ", actual " + formula.CachedValue);
                } else Assert.Equal(before, formula.CachedValue);
            }
        }

        [Theory]
        [InlineData("RATE(48,-200,8000)", 0.0077014724882019)]
        [InlineData("RATE(12,-100,1200,,,0)", 0)]
        [InlineData("RATE(10,,-100,200)", 0.0717734625362932)]
        [InlineData("RATE(12,-100,1200,0,1)", 0)]
        [InlineData("IRR({-70000,12000,15000,18000,21000,26000})", 0.0866309480365316)]
        [InlineData("IRR({-70000;12000;15000},-0.1)", -0.443506941334658)]
        [InlineData("IRR({-100,50})", -0.5)]
        [InlineData("IRR({-100,100})", 0)]
        [InlineData("IRR({-100,230,-132},0.1)", 0.1)]
        [InlineData("IRR({-100,230,-132},0.2)", 0.2)]
        [InlineData("IRR({0,-100,120},1e15)", 0.2)]
        [InlineData("XIRR({0,-100,120},{0,365,730},1e15)", 0.2)]
        [InlineData("PMT(0.01,12,1200,0,2)", -105.56291724159408)]
        [InlineData("PV(0.01,12,-100,0,0.5)", 1136.762824821948)]
        [InlineData("PMT(1e-16,12,1200)", -100)]
        [InlineData("PV(1e-16,12,-100)", 1200)]
        [InlineData("FV(1e-16,12,-100)", 1200)]
        [InlineData("NPER(1e-16,-100,1200)", 12)]
        [InlineData("NPER(-0.1,0,-100,1e-15)", 371.52387055530824)]
        [InlineData("PMT(1,0.00001,100)", -14427000.408947396)]
        [InlineData("IRR({-100,\"ignored\",TRUE,110})", 0.1)]
        [InlineData("XIRR({-10000,2750,4250,3250,2750},{39448,39508,39751,39859,39904})", 0.373362533518831)]
        [InlineData("XNPV(0.09,{-10000,2750,4250,3250,2750},{39448,39508,39751,39859,39904})", 2086.64760203154)]
        [InlineData("XNPV(0.1,{-100,110},{40000.9,40365.1})", 0)]
        public void Financial_results_match_reference_equations_and_survive_save(string formula, double expected) {
            using var document = ExcelDocument.Create(); var sheet = document.AddSheet("Finance");
            sheet.CellValue(1, 1, 99d); sheet.CellFormula(1, 1, formula);
            Assert.Equal(1, document.Calculate());
            Assert.InRange(Math.Abs(sheet.CellAt(1, 1).GetValue<double>() - expected), 0, 1e-7);
            using var copy = ExcelDocument.Load(new MemoryStream(document.ToBytes()));
            Assert.Equal(formula, copy.Sheets[0].GetFormulaText(1, 1));
            Assert.InRange(Math.Abs(copy.Sheets[0].CellAt(1, 1).GetValue<double>() - expected), 0, 1e-7);
        }

        [Theory]
        [InlineData("IRR({1,2,3})", "#NUM!")]
        [InlineData("IRR({-1,-2})", "#NUM!")]
        [InlineData("IRR({-1,2},-1)", "#VALUE!")]
        [InlineData("IRR({-100,\"text\",110,#N/A})", "#VALUE!")]
        [InlineData("IRR({-1,2},\"bad\")", "#VALUE!")]
        [InlineData("XIRR({-100,110},{40000,39999})", "#NUM!")]
        [InlineData("XIRR({-100,110},{40000})", "#NUM!")]
        [InlineData("XIRR({-100,110},{40000,\"date\"})", "#VALUE!")]
        [InlineData("XIRR({-100,110},{-1,40000})", "#NUM!")]
        [InlineData("XNPV(0,{-100,110},{40000.9,40365.1})", "#NUM!")]
        [InlineData("XNPV(-1,{-100,110},{40000,40365})", "#NUM!")]
        [InlineData("XNPV(,{-100,110},{40000,40365})", "#N/A")]
        [InlineData("RATE(100,0,100,0,0,-0.5)", "#NUM!")]
        [InlineData("RATE(1,1,-1,1,1,1e15)", "#NUM!")]
        [InlineData("RATE(1,1,-1,1,1,1e308)", "#NUM!")]
        [InlineData("RATE(1,1,-1,1e-308,1,1e15)", "#NUM!")]
        [InlineData("XNPV(0.1,{-100,#REF!},{40000,40365})", "#REF!")]
        [InlineData("XIRR({-100,110},{40000,40000})", "#NUM!")]
        [InlineData("RATE(0,-100,1200)", "#NUM!")]
        [InlineData("RATE(12,100,1200)", "#NUM!")]
        [InlineData("RATE(#N/A,-100,1200)", "#N/A")]
        [InlineData("IFERROR(IRR({1,2,3}),42)", "42")]
        public void Financial_failures_replace_stale_caches_with_typed_errors(string formula, string expected) {
            using var document = ExcelDocument.Create(); var sheet = document.AddSheet("Finance");
            sheet.CellValue(1, 1, 99d); sheet.CellFormula(1, 1, formula);
            Assert.Equal(1, document.Calculate());
            Assert.Equal(expected, Assert.Single(document.InspectFormulas().Formulas).CachedValue);
        }

        [Fact]
        public void Financial_vectors_share_reference_budgets_and_recalculate_cross_sheet_dependencies() {
            using var document = ExcelDocument.Create(); var inputs = document.AddSheet("Inputs");
            inputs.CellValue(1, 1, -100d); inputs.CellValue(2, 1, "ignored"); inputs.CellValue(4, 1, 110d);
            var sheet = document.AddSheet("Results"); sheet.CellFormula(1, 1, "IRR(Inputs!A1:A4)");
            sheet.CellFormula(2, 1, "XNPV(0.1,Inputs!A1:A1,Inputs!A4:A4)");
            Assert.Equal(2, document.Calculate());
            Assert.InRange(sheet.CellAt(1, 1).GetValue<double>(), 0.0999999999, 0.1000000001);
            inputs.CellValue(4, 1, 120d); Assert.Equal(2, document.Calculate());
            Assert.InRange(sheet.CellAt(1, 1).GetValue<double>(), 0.1999999999, 0.2000000001);
            sheet.CellValue(3, 1, 42d); sheet.CellFormula(3, 1, "IRR(Inputs!A1:A100001)");
            document.Calculate(); Assert.Equal(42d, sheet.CellAt(3, 1).GetValue<double>());
            foreach (string function in new[] { "RATE", "IRR", "XIRR", "XNPV" })
                Assert.Contains(function, document.InspectFormulas().Capabilities.SupportedFunctions);
        }
    }
}
