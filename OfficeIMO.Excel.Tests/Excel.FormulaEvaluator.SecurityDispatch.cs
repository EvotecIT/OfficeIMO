using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("IFERROR", "SUM(A1:A100001)")]
        [InlineData("IFNA", "SUM(A1:A100001)")]
        [InlineData("MINA", "SUM(A1:A100001)")]
        [InlineData("MAXA", "SUM(A1:A100001)")]
        [InlineData("AVERAGEA", "SUM(A1:A100001)")]
        public void UnsupportedNestedFormulaDoesNotProduceAnInferredValue(string function, string leaf) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Bounds");
            string formula = leaf;
            for (int depth = 0; depth < 12; depth++)
                formula = function == "IFERROR" || function == "IFNA"
                    ? $"{function}({formula},0)"
                    : $"{function}({formula})";
            sheet.CellFormula(1, 2, formula);

            Assert.Equal(0, document.Calculate());
            Assert.False(sheet.TryGetCachedFormulaValue(1, 2, out _));
        }

        [Theory]
        [InlineData("IF")]
        [InlineData("IF_COMPARISON")]
        [InlineData("DATE")]
        [InlineData("DATEDIF")]
        [InlineData("TEXTBEFORE")]
        public void NestedConditionsAndDatesDoNotRetryUnsupportedSubtrees(string function) {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Bounds");
            string formula = "SUM(A1:A100001)";
            for (int depth = 0; depth < 12; depth++)
                formula = function switch {
                    "IF" => $"IF({formula},1,0)",
                    "IF_COMPARISON" => $"IF({formula}=0,1,0)",
                    "DATE" => $"DATE({formula},1,1)",
                    "DATEDIF" => $"DATEDIF({formula},1,\"D\")",
                    _ => $"TEXTBEFORE(\"text\",\"e\",1,0,{formula})"
                };
            sheet.CellFormula(1, 2, formula);

            Assert.Equal(0, document.Calculate());
            Assert.False(sheet.TryGetCachedFormulaValue(1, 2, out _));
        }

        [Fact]
        public void NestedTextJoinBooleanPolicyKeepsTheCalculatedValue() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Bounds");
            string formula = "0";
            for (int depth = 0; depth < 12; depth++)
                formula = $"LEN(TEXTJOIN(\"\",{formula},\"x\"))";
            sheet.CellFormula(1, 1, formula);

            Assert.Equal(1, document.Calculate());
            Assert.True(sheet.TryGetCachedFormulaValue(1, 1, out string? cached));
            Assert.Equal("1", cached);
        }
    }
}
