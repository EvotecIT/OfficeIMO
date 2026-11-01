using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("vertical_default", "VLOOKUP(21,A1:B6,2)", "last20")]
        [InlineData("vertical_true", "VLOOKUP(21,A1:B6,2,TRUE)", "last20")]
        [InlineData("vertical_empty", "VLOOKUP(21,A1:B6,2,)", "#N/A")]
        [InlineData("numeric_true", "VLOOKUP(21,A1:B6,2,0.00000001)", "last20")]
        [InlineData("vertical_exact", "VLOOKUP(20,A1:B6,2,FALSE)", "first20")]
        [InlineData("horizontal_default", "HLOOKUP(21,J1:O2,2)", "last20")]
        [InlineData("match_ascending", "MATCH(21,A1:A6,1)", "3")]
        [InlineData("match_duplicate", "MATCH(20,A1:A6)", "3")]
        [InlineData("match_descending", "MATCH(21,D1:D6,-1)", "3")]
        [InlineData("match_wildcard", "MATCH(\"Al*\",C1:C6,0)", "1")]
        [InlineData("match_escaped_question", "MATCH(\"B~?ta\",C1:C6,0)", "4")]
        [InlineData("match_escaped_star", "MATCH(\"B~*ta\",C1:C6,0)", "5")]
        [InlineData("wildcard_missing", "IFNA(MATCH(\"Z*\",C1:C6,0),77)", "77")]
        [InlineData("vertical_wildcard", "VLOOKUP(\"Al*\",C1:G6,5,FALSE)", "N")]
        [InlineData("xmatch_wildcard", "XMATCH(\"Al*\",C1:C6,2)", "1")]
        [InlineData("xmatch_reverse_wildcard", "XMATCH(\"Al*\",C1:C6,2,-1)", "2")]
        [InlineData("xlookup_wildcard", "XLOOKUP(\"B~?ta\",C1:C6,B1:B6,,2)", "thirty")]
        [InlineData("binary_ascending", "XMATCH(30,A1:A6,0,2)", "4")]
        [InlineData("binary_descending", "XMATCH(30,D1:D6,0,-2)", "3")]
        [InlineData("binary_next_smaller", "XMATCH(35,A1:A6,-1,2)", "4")]
        [InlineData("binary_next_larger", "XMATCH(35,D1:D6,1,-2)", "2")]
        [InlineData("binary_return", "XLOOKUP(30,A1:A6,B1:B6,,,2)", "thirty")]
        [InlineData("binary_missing", "XLOOKUP(31,A1:A6,B1:B6,\"missing\",0,2)", "missing")]
        [InlineData("numeric_key", "MATCH(1,F1:F4,0)", "1")]
        [InlineData("text_key", "MATCH(\"1\",F1:F4,0)", "2")]
        [InlineData("boolean_key", "MATCH(TRUE,F1:F4,0)", "3")]
        [InlineData("false_key", "MATCH(FALSE,F1:F4,0)", "4")]
        [InlineData("literal_vector", "MATCH(30,{10,20,30},0)", "3")]
        [InlineData("fractional_column", "VLOOKUP(20,A1:B6,2.9,FALSE)", "first20")]
        [InlineData("fractional_row", "INDEX(A1:B6,1.9,1)", "10")]
        [InlineData("text_row", "INDEX(A1:B6,\"2\",1)", "20")]
        [InlineData("boolean_row", "INDEX(A1:B6,TRUE,1)", "10")]
        [InlineData("invalid_text_row", "IFERROR(INDEX(A1:B6,\"x\",1),77)", "77")]
        public void FormulaEvaluator_LookupModesCalculateTypedResults(string caseId, string formula, string expected) {
            using var document = ExcelDocument.Create();
            var sheet = CreateLookupModeSheet(document);
            sheet.CellFormula(1, 18, formula);
            Assert.Equal(1, document.Calculate());
            AssertFormulaCache(sheet, 1, 18, expected);
            Assert.True(Assert.Single(sheet.InspectFormulas().Formulas).IsSupportedByOfficeIMO, caseId);
        }

        [Theory]
        [InlineData(2)]
        [InlineData(-2)]
        public void FormulaEvaluator_BinaryLookupResolvesOnlyVisitedFormulaCells(int searchMode) {
            using var document = ExcelDocument.Create();
            int reads = 0;
            document.Calculation.RegisterCustomFunction("PROBE", (_, arguments) => {
                reads++;
                return arguments[0];
            });
            var source = document.AddWorksheet("Source");
            for (int row = 1; row <= 1024; row++) {
                int key = searchMode == 2 ? row : 1025 - row;
                source.CellFormula(row, 1, $"PROBE({key})");
                source.CellValue(row, 2, key * 10);
            }
            var report = document.AddWorksheet("Report");
            report.CellFormula(1, 1, $"XLOOKUP(777,Source!A1:A1024,Source!B1:B1024,,, {searchMode})");
            Assert.Equal(1, report.RecalculateSupportedFormulas());
            AssertFormulaCache(report, 1, 1, "7770");
            Assert.InRange(reads, 1, 11);
        }

        [Theory]
        [InlineData("XMATCH(\"A-2\",A1:A2,1)")]
        [InlineData("XMATCH(\"A-2\",A1:A2,0,2)")]
        public void FormulaEvaluator_UnqualifiedOrderedTextDoesNotBecomeMissingMatch(string formula) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Ordering");
            sheet.CellValue(1, 1, "A1");
            sheet.CellValue(2, 1, "A3");
            sheet.CellValue(1, 4, 42d);
            sheet.CellFormula(1, 4, $"IFERROR({formula},99)");
            Assert.Equal(0, document.Calculate());
            AssertFormulaCache(sheet, 1, 4, "42");
            Assert.False(Assert.Single(sheet.InspectFormulas().Formulas).IsSupportedByOfficeIMO);
        }

        private static ExcelSheet CreateLookupModeSheet(ExcelDocument document) {
            var sheet = document.AddWorksheet("Data");
            int[] ascending = { 10, 20, 20, 30, 40, 50 };
            string[] labels = { "ten", "first20", "last20", "thirty", "forty", "fifty" };
            string[] keys = { "Alpha", "Alpine", "Beta", "B?ta", "B*ta", "~tail" };
            string[] typedLabels = { "N", "T", "B", "F", "empty", "E" };
            for (int index = 0; index < ascending.Length; index++) {
                sheet.CellValue(index + 1, 1, ascending[index]);
                sheet.CellValue(index + 1, 2, labels[index]);
                sheet.CellValue(index + 1, 3, keys[index]);
                sheet.CellValue(index + 1, 4, ascending[ascending.Length - index - 1]);
                sheet.CellValue(index + 1, 7, typedLabels[index]);
                sheet.CellValue(1, index + 10, ascending[index]);
                sheet.CellValue(2, index + 10, labels[index]);
            }
            sheet.CellValue(1, 6, 1d);
            sheet.CellValue(2, 6, "1");
            sheet.CellValue(3, 6, true);
            sheet.CellValue(4, 6, false);
            return sheet;
        }
    }
}
