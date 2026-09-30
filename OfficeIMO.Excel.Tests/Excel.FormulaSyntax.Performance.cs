using System;
using System.Diagnostics;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
#if EXCEL_PERFORMANCE_EVIDENCE
        [Trait("Category", "Performance")]
        [Trait("Category", "ExcelPerformanceEvidence")]
#endif
        [Fact]
        public void Test_FormulaReferenceScanner_RejectsLongNonReferenceRunsWithinLinearBudget() {
#if EXCEL_PERFORMANCE_EVIDENCE
            const int length = 100_000;
#else
            const int length = 64;
#endif
            string formula = new string('?', length) + "+A1";
#if EXCEL_PERFORMANCE_EVIDENCE
            var stopwatch = Stopwatch.StartNew();
#endif

            ExcelFormulaSyntaxTree tree = ExcelFormulaSyntaxTree.Parse(formula);

#if EXCEL_PERFORMANCE_EVIDENCE
            stopwatch.Stop();
#endif
            ExcelFormulaReferenceSyntax reference = Assert.Single(
                tree.Nodes,
                node => node is ExcelFormulaReferenceSyntax) as ExcelFormulaReferenceSyntax
                ?? throw new InvalidOperationException("Expected one A1 reference node.");
            Assert.Equal("A1", reference.Text);
            Assert.Equal(formula, tree.Text);
#if EXCEL_PERFORMANCE_EVIDENCE
            Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(5),
                $"Formula parsing exceeded the linear-time regression budget: {stopwatch.Elapsed}.");
#endif
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
