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
    }
}
