using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("star_inside", "COUNTIFS(A1:A6,\"a*c\")", "4")]
        [InlineData("escaped_star", "COUNTIFS(A1:A6,\"a~*c\")", "1")]
        [InlineData("question_then_star", "COUNTIFS(A1:A6,\"?b*\")", "3")]
        [InlineData("star_backtracks", "COUNTIFS(A1:A6,\"*ab?\")", "2")]
        [InlineData("plain_tilde", "COUNTIFS(A1:A6,\"x~y\")", "1")]
        [InlineData("negated", "COUNTIFS(A1:A6,\"<>a*\")", "1")]
        [InlineData("sum", "SUMIFS(B1:B6,A1:A6,\"*ab?\")", "5")]
        [InlineData("search_question", "SEARCH(\"b?d\",\"abcbxd\")", "4")]
        [InlineData("search_star", "SEARCH(\"A*D\",\"xxabcd\")", "3")]
        [InlineData("search_escaped", "SEARCH(\"~*\",\"a*c\")", "2")]
        [InlineData("search_start", "SEARCH(\"a?c\",\"abcaxc\",2)", "4")]
        [InlineData("search_no_match_after_star", "IFERROR(SEARCH(\"a*z\",\"aaaaaa\"),0)", "0")]
        [InlineData("search_repeated_prefix", "SEARCH(\"aab?\",\"aaaabxaab\")", "3")]
        [InlineData("search_prefix_then_star", "SEARCH(\"ab*d\",\"xabcabd\")", "2")]
        [InlineData("search_prefix_no_match", "IFERROR(SEARCH(\"aaab\",\"aaaaaa\"),0)", "0")]
        [InlineData("search_leading_star", "SEARCH(\"*c\",\"abcabc\",2)", "2")]
        public void FormulaEvaluator_WildcardCriteriaMatchCaseInsensitively(string caseId, string formula, string expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Wildcards");
            string[] text = { "abc", "ABXC", "ac", "abcabd", "a*c", "x~y" };
            for (int row = 0; row < text.Length; row++) {
                sheet.Cell(row + 1, 1).SetValue(text[row]);
                sheet.Cell(row + 1, 2).SetValue(row + 1d);
            }
            sheet.Cell(1, 4).SetFormula(formula);
            Assert.Equal(1, document.Calculate());
            AssertFormulaCache(sheet, 1, 4, expected);
            Assert.True(Assert.Single(sheet.InspectFormulas().Formulas).IsSupportedByOfficeIMO, caseId);
        }

        [Fact]
        public void FormulaEvaluator_WildcardCriteriaKeepUnicodeCaseEquivalences() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Wildcards");
            sheet.Cell(1, 1).SetValue("\u212Aelvin");
            sheet.Cell(1, 4).SetFormula("COUNTIFS(A1,\"k*\")");
            Assert.Equal(1, document.Calculate());
            // Matching follows the runtime's invariant case tables, as the Regex it replaced did. .NET maps the Kelvin
            // sign to 'k'; .NET Framework leaves it unchanged, so neither matched there.
            AssertFormulaCache(sheet, 1, 4, char.ToLowerInvariant('\u212A') == 'k' ? "1" : "0");
        }

        [Fact]
        public void FormulaEvaluator_WildcardCriteriaKeepFinalSigmaDistinct() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Wildcards");
            sheet.Cell(1, 1).SetValue("\u03C2");
            sheet.Cell(2, 1).SetValue("\u03A3");
            sheet.Cell(1, 4).SetFormula("COUNTIFS(A1:A2,\"\u03C3*\")");
            Assert.Equal(1, document.Calculate());
            AssertFormulaCache(sheet, 1, 4, "1");
        }
    }
}
