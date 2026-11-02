using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("whole_column", "SUM(INDEX(A1:B6,0,1))", "170")]
        [InlineData("whole_row", "SUM(INDEX(A1:B6,3,0))", "20")]
        [InlineData("omitted_column", "SUM(INDEX(A1:B6,3))", "#REF!")]
        [InlineData("empty_column", "SUM(INDEX(A1:B6,3,))", "20")]
        [InlineData("whole_matrix", "SUM(INDEX(A1:B6,0,0))", "170")]
        [InlineData("nested_reference", "SUM(INDEX(INDEX(A1:B6,0,0),0,1))", "170")]
        [InlineData("row_count", "ROWS(INDEX(A1:B6,0,1))", "6")]
        [InlineData("column_count", "COLUMNS(INDEX(A1:B6,3,0))", "2")]
        [InlineData("array_scalar", "INDEX({10,20;30,40},2,2)", "40")]
        [InlineData("array_aggregate", "SUM(INDEX({10,20;30,40},0,2))", "60")]
        [InlineData("generated_array", "SUM(INDEX(SEQUENCE(3,2),0,2))", "12")]
        [InlineData("literal_aggregate", "SUM({1,2;3,4})", "10")]
        [InlineData("selector_error", "IFERROR(SUM(INDEX(A1:B6,0,NA())),99)", "99")]
        [InlineData("whole_column_scalar", "INDEX(A:A,6)", "50")]
        [InlineData("whole_row_scalar", "INDEX(1:1,2)", "ten")]
        [InlineData("selected_area", "SUM(INDEX((A1:A2,A4:A6),0,1,2))", "120")]
        public void FormulaEvaluator_IndexReferencesComposeWithRangeAndArrayConsumers(string caseId, string formula, string expected) {
            using var document = ExcelDocument.Create();
            var sheet = CreateLookupModeSheet(document);
            sheet.Cell(1, 18).SetFormula(formula);
            Assert.Equal(1, document.Calculate());
            AssertFormulaCache(sheet, 1, 18, expected);
            Assert.True(Assert.Single(sheet.InspectFormulas().Formulas).IsSupportedByOfficeIMO, caseId);
        }

        [Fact]
        public void FormulaEvaluator_IndexAreasOnDifferentSheetsReturnTypedValueErrors() {
            using var document = ExcelDocument.Create();
            var source = document.AddSheet("Data");
            source.Cell(1, 1).SetValue(10d);
            var report = document.AddSheet("Report");
            report.Cell(1, 1).SetValue(20d);
            report.Cell(1, 4).SetFormula("INDEX((Data!A1,Report!A1),1,1,2)");
            report.Cell(2, 4).SetFormula("IFERROR(SUM(INDEX((Data!A1,Report!A1),0,1,2)),99)");
            report.SetArrayFormula("D4:D4", "INDEX((Data!A1,Report!A1),1,1,2)");
            Assert.Equal(3, document.Calculate());
            AssertFormulaCache(report, 1, 4, "#VALUE!");
            AssertFormulaCache(report, 2, 4, "99");
            AssertFormulaCache(report, 4, 4, "#VALUE!");
        }

        [Fact]
        public void FormulaEvaluator_IndexAndXLookupArraysFeedDependentCellsAndRefresh() {
            using var document = ExcelDocument.Create();
            var sheet = CreateLookupModeSheet(document);
            sheet.SetDynamicArrayFormula("R2", "INDEX(A1:B6,0,1)");
            sheet.SetArrayFormula("T2:U2", "INDEX(A1:B6,3,0)");
            sheet.SetDynamicArrayFormula("W2", "XLOOKUP(20,A1:A6,B1:C6)");
            sheet.Cell(1, 18).SetFormula("R7+T2");
            sheet.Cell(1, 23).SetFormula("X2&\"!\"");
            Assert.Equal(5, document.Calculate());
            AssertFormulaCache(sheet, 1, 18, "70");
            AssertFormulaCache(sheet, 1, 23, "Alpine!");
            Assert.Equal("last20", sheet.Cell(2, 21).GetValue().Value);
            sheet.Cell(6, 1).SetValue(60d);
            sheet.Cell(2, 3).SetValue("Changed");
            Assert.Equal(5, document.Calculate());
            AssertFormulaCache(sheet, 1, 18, "80");
            AssertFormulaCache(sheet, 1, 23, "Changed!");
        }
    }
}
