using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelFormulaRangeQualificationTests {
        [Theory]
        [InlineData("LCM(1E20,OFFSET(A1,-1,0))")]
        [InlineData("GCD(OFFSET(A1,-1,0),1E20)")]
        [InlineData("SUM(OFFSET(A1,-1,0),UNKNOWN())")]
        [InlineData("CONCAT(OFFSET(A1,-1,0),UNKNOWN())")]
        [InlineData("PROB(OFFSET(A1,-1,0),B1,UNKNOWN())")]
        [InlineData("PROB(A1,OFFSET(B1,-1,0),UNKNOWN())")]
        [InlineData("INDEX(OFFSET(A1,-1,0),UNKNOWN())")]
        [InlineData("OFFSET(OFFSET(A1,-1,0),UNKNOWN(),0)")]
        [InlineData("SUM(OFFSET(A1,-1,0,UNKNOWN()))")]
        [InlineData("MATCH(1,OFFSET(A1,-1,0),UNKNOWN())")]
        [InlineData("XLOOKUP(1,OFFSET(A1,-1,0),B1,,UNKNOWN())")]
        [InlineData("SUMPRODUCT(OFFSET(A1,-1,0),UNKNOWN())")]
        [InlineData("PERCENTILE.INC(OFFSET(A1,-1,0),UNKNOWN())")]
        [InlineData("CORREL(OFFSET(A1,-1,0),UNKNOWN())")]
        [InlineData("COUNTIFS(OFFSET(A1,-1,0),UNKNOWN())")]
        [InlineData("COUNTIF(OFFSET(A1,-1,0),UNKNOWN()+1)")]
        [InlineData("SUMIF(OFFSET(A1,-1,0),UNKNOWN()+1)")]
        [InlineData("AVERAGEIF(OFFSET(A1,-1,0),UNKNOWN()+1)")]
        [InlineData("COUNTIFS(OFFSET(A1,-1,0),UNKNOWN()+1)")]
        [InlineData("SUMIFS(B1,OFFSET(A1,-1,0),UNKNOWN()+1)")]
        [InlineData("AVERAGEIFS(B1,OFFSET(A1,-1,0),UNKNOWN()+1)")]
        [InlineData("MINIFS(B1,OFFSET(A1,-1,0),UNKNOWN()+1)")]
        [InlineData("MAXIFS(B1,OFFSET(A1,-1,0),UNKNOWN()+1)")]
        public void Unsupported_inputs_are_not_masked_by_reference_errors(string expression) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Qualification");
            sheet.Cell(1, 1).SetValue(1d); sheet.Cell(1, 2).SetValue(1d);
            string[] wrappers = { "IFERROR({0},99)", "IFNA({0},99)", "ISERROR({0})" };
            for (int index = 0; index < wrappers.Length; index++) {
                sheet.Cell(index + 1, 4).SetValue(42d);
                sheet.Cell(index + 1, 4).SetFormula(string.Format(wrappers[index], expression));
            }
            Assert.Equal(0, document.Calculate());
            for (int row = 1; row <= wrappers.Length; row++) {
                Assert.True(sheet.TryGetCachedFormulaValue(row, 4, out string? cache));
                Assert.Equal("42", cache);
            }
        }

        [Fact]
        public void Qualified_reference_errors_still_reach_error_handlers_across_range_consumers() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Qualified");
            sheet.Cell(1, 1).SetValue(1d); sheet.Cell(1, 2).SetValue(1d);
            string[] expressions = { "GCD(OFFSET(A1,-1,0),12)", "LCM(OFFSET(A1,-1,0),12)",
                "COUNT(OFFSET(A1,-1,0))", "COUNTA(OFFSET(A1,-1,0))", "AVERAGEA(OFFSET(A1,-1,0))",
                "SUMPRODUCT(OFFSET(A1,-1,0),B1)", "NPV(0.1,OFFSET(A1,-1,0),B1)",
                "PERCENTILE.INC(OFFSET(A1,-1,0),0.5)", "QUARTILE.INC(OFFSET(A1,-1,0),2)",
                "PERCENTRANK.INC(OFFSET(A1,-1,0),1)", "RANK.EQ(1,OFFSET(A1,-1,0))",
                "CORREL(OFFSET(A1,-1,0),B1)", "COVAR(OFFSET(A1,-1,0),B1)",
                "FORECAST.LINEAR(1,OFFSET(A1,-1,0),B1)", "SUMXMY2(OFFSET(A1,-1,0),B1)",
                "STDEV.S(OFFSET(A1,-1,0))", "LARGE(OFFSET(A1,-1,0),1)",
                "COUNTIFS(OFFSET(A1,-1,0),1)", "SUMIF(OFFSET(A1,-1,0),1)",
                "AVERAGEIF(OFFSET(A1,-1,0),1)", "AVERAGEIFS(B1,OFFSET(A1,-1,0),1)",
                "SUBTOTAL(9,OFFSET(A1,-1,0))", "FORMULATEXT(OFFSET(A1,-1,0))", "ISFORMULA(OFFSET(A1,-1,0))",
                "XLOOKUP(1,OFFSET(A1,-1,0),B1,,0,-1)" };
            for (int index = 0; index < expressions.Length; index++)
                sheet.Cell(index + 1, 4).SetFormula("IFERROR(" + expressions[index] + ",99)");
            Assert.Equal(expressions.Length, document.Calculate());
            for (int index = 0; index < expressions.Length; index++) {
                Assert.True(sheet.TryGetCachedFormulaValue(index + 1, 4, out string? cache));
                Assert.True(cache == "99", expressions[index] + " returned " + cache);
            }
        }

        [Theory]
        [InlineData("XLOOKUP(2,A1:A1,B1:B1)", "#N/A")]
        [InlineData("XLOOKUP(2,A1:A1,B1:B1,,0,-1)", "#N/A")]
        [InlineData("XLOOKUP(2,A1:A1,B1:B1,99,0,-1)", "99")]
        [InlineData("XLOOKUP(1,A1:A2,B1:B2,,0,-1)", "20")]
        [InlineData("XLOOKUP(2,A1:A1,B1:B1,,0)", "#N/A")]
        [InlineData("XLOOKUP(2,A1:A1,B1:B1,99,,-1)", "99")]
        [InlineData("XLOOKUP(2,A1:A1,B1:B1,99,0,)", "99")]
        public void Optional_lookup_positions_follow_fresh_native_Excel_results(string formula, string expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Lookup positions");
            sheet.Cell(1, 1).SetValue(1d); sheet.Cell(2, 1).SetValue(1d);
            sheet.Cell(1, 2).SetValue(10d); sheet.Cell(2, 2).SetValue(20d);
            sheet.Cell(1, 4).SetFormula(formula);
            Assert.Equal(1, document.Calculate());
            Assert.True(sheet.TryGetCachedFormulaValue(1, 4, out string? cache)); Assert.Equal(expected, cache);
        }

        [Fact]
        public void Criteria_keep_quoted_expression_text_and_resolved_arithmetic() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Criteria qualification");
            sheet.Cell(1, 1).SetValue("UNKNOWN()+1"); sheet.Cell(2, 1).SetValue(2d);
            sheet.Cell(1, 2).SetValue(7d); sheet.Cell(2, 2).SetValue(11d);
            sheet.Cell(1, 4).SetFormula("COUNTIF(A1:A2,\"UNKNOWN()+1\")");
            sheet.Cell(2, 4).SetFormula("SUMIF(A1:A2,\"UNKNOWN()+1\",B1:B2)");
            sheet.Cell(3, 4).SetFormula("IFERROR(COUNTIFS(OFFSET(A1,-1,0),\"UNKNOWN()+1\"),99)");
            sheet.Cell(4, 4).SetFormula("COUNTIF(A1:A2,1+1)");
            Assert.Equal(4, document.Calculate());
            string[] expected = { "1", "7", "99", "1" };
            for (int row = 1; row <= expected.Length; row++) {
                Assert.True(sheet.TryGetCachedFormulaValue(row, 4, out string? cache));
                Assert.Equal(expected[row - 1], cache);
            }
        }

        [Theory]
        [InlineData("GCD", "12")]
        [InlineData("LCM", "0")]
        [InlineData("SUM", "12")]
        public void Scalar_references_share_the_aggregate_range_cell_budget(string function, string exactLimitResult) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Budget");
            sheet.Cell(1, 1).SetValue(12d); sheet.Cell(1, 2).SetValue(18d);
            sheet.Cell(1, 4).SetFormula(function + "(A1:A100000)");
            sheet.Cell(2, 4).SetValue(42d); sheet.Cell(2, 4).SetFormula(function + "(A1:A100000,B1)");
            sheet.Cell(3, 4).SetValue(42d); sheet.Cell(3, 4).SetFormula(function + "(A1:A100000,OFFSET(B1,0,0))");
            Assert.Equal(1, document.Calculate());
            Assert.True(sheet.TryGetCachedFormulaValue(1, 4, out string? exactLimit));
            Assert.Equal(exactLimitResult, exactLimit);
            for (int row = 2; row <= 3; row++) {
                Assert.True(sheet.TryGetCachedFormulaValue(row, 4, out string? cache));
                Assert.Equal("42", cache);
            }
        }

        [Fact]
        public void Reference_errors_do_not_invoke_custom_functions_with_incomplete_arguments() {
            using var document = ExcelDocument.Create();
            int calls = 0;
            document.Calculation.RegisterCustomFunction("QUALIFIEDINPUTS", (_, arguments) => {
                calls++;
                return ExcelFormulaValue.FromNumber(arguments.Count);
            });
            var sheet = document.AddSheet("Custom");
            sheet.Cell(1, 4).SetFormula("QUALIFIEDINPUTS(12,18)");
            sheet.Cell(2, 4).SetFormula("IFERROR(QUALIFIEDINPUTS(OFFSET(A1,-1,0),18),99)");
            sheet.Cell(3, 4).SetValue(42d);
            sheet.Cell(3, 4).SetFormula("IFERROR(QUALIFIEDINPUTS(OFFSET(A1,-1,0),UNKNOWN()),99)");
            Assert.Equal(2, document.Calculate());
            Assert.Equal(1, calls);
            Assert.True(sheet.TryGetCachedFormulaValue(1, 4, out string? first)); Assert.Equal("2", first);
            Assert.True(sheet.TryGetCachedFormulaValue(2, 4, out string? second)); Assert.Equal("99", second);
            Assert.True(sheet.TryGetCachedFormulaValue(3, 4, out string? third)); Assert.Equal("42", third);
        }

        [Fact]
        public void Unqualified_array_reference_errors_do_not_leak_into_dependent_scalar_calculation() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Array owner");
            sheet.Cell(1, 6).SetValue(42d);
            sheet.SetArrayFormula("F1:F1", "SORT(OFFSET(A1,-1,0))");
            sheet.Cell(1, 8).SetFormula("COUNT(F1)");
            sheet.Cell(2, 8).SetFormula("IFERROR(SUM(F1),99)");
            Assert.Equal(2, document.Calculate());
            Assert.True(sheet.TryGetCachedFormulaValue(1, 6, out string? source)); Assert.Equal("42", source);
            Assert.True(sheet.TryGetCachedFormulaValue(1, 8, out string? count)); Assert.Equal("1", count);
            Assert.True(sheet.TryGetCachedFormulaValue(2, 8, out string? sum)); Assert.Equal("42", sum);
        }
    }
}
