using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class ExcelIntegerAggregateTests {
        [Theory]
        [InlineData("GCD(24.9,36.2,60.8)", "12")]
        [InlineData("LCM(24.9,36.2,60.8)", "360")]
        [InlineData("GCD(0,0,0)", "0")]
        [InlineData("LCM(13)", "13")]
        [InlineData("GCD(-0.2,36,60)", "#NUM!")]
        [InlineData("LCM(-24,36,60)", "#NUM!")]
        [InlineData("GCD(TRUE,12,18)", "#VALUE!")]
        [InlineData("LCM(\"text\",12,18)", "#VALUE!")]
        [InlineData("LCM(\"24\",36,60)", "360")]
        [InlineData("LCM(134217728,67108865)", "#NUM!")]
        [InlineData("LCM(0,134217728,67108865)", "0")]
        [InlineData("LCM(134217728,67108865,0)", "0")]
        [InlineData("GCD(24,LCM(6,10))+1", "7")]
        [InlineData("IFERROR(LCM(134217728,67108865),99)", "99")]
        [InlineData("GCD(24,#REF!)", "#REF!")]
        [InlineData("GCD(9007199254740991,1)", "1")]
        [InlineData("LCM(9007199254740991,1)", "9007199254740991")]
        public void Integer_aggregates_replace_caches_and_preserve_saved_numbers_and_errors(string formula, string expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Integers");
            sheet.CellValue(1, 1, 42d); sheet.CellFormula(1, 1, formula);
            Assert.Equal(1, document.Calculate());
            Assert.Equal(expected, Assert.Single(sheet.InspectFormulas().Formulas).CachedValue);
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.Equal(formula, reopened.Sheets[0].GetFormulaText(1, 1));
            Assert.Equal(expected, Assert.Single(reopened.InspectFormulas().Formulas).CachedValue);
            Assert.Empty(reopened.ValidateOpenXml());
        }

        [Fact]
        public void Integer_aggregates_recalculate_cross_sheet_ranges_blanks_text_and_dependencies() {
            using var document = ExcelDocument.Create();
            var source = document.AddWorksheet("Source data");
            var sheet = document.AddWorksheet("Integers");
            source.CellValue(1, 1, 24d); source.CellFormula(2, 1, "18*2");
            source.CellValue(4, 1, "24"); source.CellValue(5, 1, true);
            string[] formulas = { "GCD('Source data'!A1:A3)", "LCM('Source data'!A1:A3)",
                "LCM('Source data'!A1:A2)", "LCM('Source data'!A4,'Source data'!A2)", "GCD('Source data'!A5,12)" };
            for (int row = 0; row < formulas.Length; row++) sheet.CellFormula(row + 1, 1, formulas[row]);
            Assert.Equal(6, document.Calculate());
            string[] expected = { "12", "0", "72", "72", "#VALUE!" };
            for (int row = 0; row < expected.Length; row++) AssertCache(sheet, row + 1, expected[row]);
            source.CellValue(1, 1, 60d); source.CellFormula(2, 1, "16*3");
            Assert.Equal(6, document.Calculate());
            AssertCache(sheet, 1, "12"); AssertCache(sheet, 2, "0");
            AssertCache(sheet, 3, "240"); AssertCache(sheet, 4, "48");
            Assert.Contains("GCD", document.InspectFormulas().Capabilities.SupportedFunctions);
            Assert.Contains("LCM", document.InspectFormulas().Capabilities.SupportedFunctions);
        }

        [Theory]
        [InlineData("GCD()")]
        [InlineData("LCM(24,)")]
        [InlineData("GCD(9007199254740992,4)")]
        [InlineData("IFERROR(LCM(1E20,#REF!),99)")]
        [InlineData("GCD(A1:A100001)")]
        [InlineData("LCM(A1:A50001,B1:B50001)")]
        [InlineData("IFERROR(GCD(UNKNOWN(),24),99)")]
        public void Unqualified_integer_aggregate_inputs_preserve_existing_caches(string formula) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Bounds");
            sheet.CellValue(1, 4, 42d); sheet.CellFormula(1, 4, formula);
            Assert.Equal(0, document.Calculate());
            Assert.Equal("42", Assert.Single(sheet.InspectFormulas().Formulas).CachedValue);
        }

        [Fact]
        public void Integer_aggregates_enforce_the_argument_budget_before_flattening() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Bounds");
            sheet.CellFormula(1, 1, "GCD(" + string.Join(",", Enumerable.Repeat("12", 255)) + ")");
            sheet.CellValue(2, 1, 42d);
            sheet.CellFormula(2, 1, "LCM(" + string.Join(",", Enumerable.Repeat("12", 256)) + ")");
            Assert.Equal(1, document.Calculate());
            AssertCache(sheet, 1, "12"); AssertCache(sheet, 2, "42");
        }

        private static void AssertCache(ExcelSheet sheet, int row, string expected) {
            Assert.True(sheet.TryGetCachedFormulaValue(row, 1, out string? actual));
            Assert.Equal(expected, actual);
        }

        [Fact]
        public void Integer_aggregates_distinguish_direct_blanks_from_ranges_after_saved_input_edits() {
            using var document = ExcelDocument.Create();
            var inputs = document.AddWorksheet("Source data");
            var checks = document.AddWorksheet("Checks");
            inputs.CellValue(1, 1, 30d); inputs.CellValue(3, 1, 0d);
            string[] formulas = {
                "GCD('Source data'!$A$2,6)", "LCM('Source data'!$A$2,6)",
                "GCD('Source data'!$A$1:$A$2)", "LCM('Source data'!$A$1:$A$2)",
                "GCD('Source data'!$A$2,'Source data'!$A$3)",
                "LCM('Source data'!$A$2,'Source data'!$A$3)",
                "GCD('Source data'!$A$2)", "LCM('Source data'!$A$2)",
                "GCD('Source data'!$A$2,'Source data'!$A$2)",
                "LCM('Source data'!$A$2,'Source data'!$A$2)",
                "IFERROR(LCM('Source data'!$A$2),99)"
            };
            for (int row = 0; row < formulas.Length; row++) {
                checks.CellValue(row + 1, 1, 42d); checks.CellFormula(row + 1, 1, formulas[row]);
            }
            Assert.Equal(formulas.Length, document.Calculate());
            string[] blank = { "6", "6", "30", "0", "0", "0", "#VALUE!", "#VALUE!", "#VALUE!", "#VALUE!", "99" };
            for (int row = 0; row < blank.Length; row++) AssertCache(checks, row + 1, blank[row]);
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            reopened.Sheets[0].CellValue(2, 1, 18d);
            Assert.Equal(formulas.Length, reopened.Calculate());
            string[] filled = { "6", "18", "6", "90", "18", "0", "18", "18", "18", "18", "18" };
            for (int row = 0; row < filled.Length; row++) {
                AssertCache(reopened.Sheets[1], row + 1, filled[row]);
                Assert.Equal(formulas[row], reopened.Sheets[1].GetFormulaText(row + 1, 1));
            }
            reopened.Sheets[0].ClearRange("A2:A2", ExcelClearOptions.Values);
            Assert.Equal(formulas.Length, reopened.Calculate());
            using var cleared = new MemoryStream(); reopened.Save(cleared); cleared.Position = 0;
            using var clearedReopened = ExcelDocument.Load(cleared);
            for (int row = 0; row < blank.Length; row++) AssertCache(clearedReopened.Sheets[1], row + 1, blank[row]);
            Assert.Empty(reopened.ValidateOpenXml());
        }
    }
}
