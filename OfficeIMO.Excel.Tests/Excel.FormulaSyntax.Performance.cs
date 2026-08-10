using System;
using System.Diagnostics;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_FormulaReferenceScanner_RejectsLongNonReferenceRunsWithinLinearBudget() {
            const int length = 100_000;
            string formula = new string('?', length) + "+A1";
            var stopwatch = Stopwatch.StartNew();

            ExcelFormulaSyntaxTree tree = ExcelFormulaSyntaxTree.Parse(formula);

            stopwatch.Stop();
            ExcelFormulaReferenceSyntax reference = Assert.Single(
                tree.Nodes,
                node => node is ExcelFormulaReferenceSyntax) as ExcelFormulaReferenceSyntax
                ?? throw new InvalidOperationException("Expected one A1 reference node.");
            Assert.Equal("A1", reference.Text);
            Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(5),
                $"Formula parsing exceeded the linear-time regression budget: {stopwatch.Elapsed}.");
        }

        [Theory]
        [InlineData("_Sheet!A1")]
        [InlineData("'[Book.xlsx]Data Set'!$B$2")]
        [InlineData("[Book.xlsx]Data!C3")]
        public void Test_FormulaReferenceScanner_PreservesSupportedReferenceStarts(string formula) {
            ExcelFormulaReferenceSyntax reference = Assert.Single(
                ExcelFormulaSyntaxTree.Parse(formula).Nodes,
                node => node is ExcelFormulaReferenceSyntax) as ExcelFormulaReferenceSyntax
                ?? throw new InvalidOperationException("Expected one reference node.");

            Assert.Equal(formula, reference.Text);
        }
    }
}
