using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void Test_ArrayCalculation_ReferencedBooleanChildrenKeepAggregateCoercion(bool dynamic) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Arrays");
            sheet.CellValue(1, 1, 1d);
            sheet.CellValue(2, 1, 1d);
            const string formula = "FILTER(A1:A2=1,A1:A2>0)";
            if (dynamic) sheet.SetDynamicArrayFormula("B1", formula);
            else sheet.SetArrayFormula("B1:B2", formula);
            string[] formulas = { "SUM(B1:B2)", "COUNT(B1:B2)", "MINA(B1:B2)", "MAXA(B1:B2)",
                "AVERAGEA(B1:B2)", "ISLOGICAL(B2)", "SUM(TRUE(),B2)", "SUMSQ(B1:B2)",
                "MIN(B1:B2)", "MAX(B1:B2)", "PRODUCT(B1:B2)", "AVERAGE(B1:B2)", "MEDIAN(B1:B2)" };
            for (int row = 1; row <= formulas.Length; row++) sheet.CellFormula(row, 4, formulas[row - 1]);

            document.Calculate();
            string[] expected = { "0", "0", "1", "1", "1", "1", "1", "0", "0", "0", "0", "#DIV/0!", "#NUM!" };
            for (int row = 1; row <= expected.Length; row++) {
                Assert.True(sheet.TryGetCachedFormulaValue(row, 4, out string? cached));
                Assert.Equal(expected[row - 1], cached);
            }
        }
    }
}
