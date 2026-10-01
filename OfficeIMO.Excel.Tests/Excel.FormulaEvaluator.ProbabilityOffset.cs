using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_FormulaEvaluator_ProbabilityNumericVectorsAndTypedFailures() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Distribution");
            double[] probabilities = { .2, .3, .1, .4 };
            for (int index = 0; index < 4; index++) {
                sheet.CellValue(index + 1, 1, index);
                sheet.CellValue(index + 1, 2, probabilities[index]);
            }
            string[] formulas = { "PROB(A1:A4,B1:B4,2)", "PROB(A1:A4,B1:B4,1,3)", "PROB(A1:A4,B1:B4,5)",
                "PROB(A1:A4,B1:B4,3,1)", "PROB(A1:A4,B1:B3,1)", "IFNA(PROB(A1:A4,B1:B3,1),99)",
                "PROB(A1:A4,B1:B4,0,3)+2", "PROB(A1:A4,B1:B4,#REF!)" };
            for (int index = 0; index < formulas.Length; index++) sheet.CellFormula(index + 1, 4, formulas[index]);
            Assert.Equal(formulas.Length, document.Calculate());
            string[] expected = { "0.1", "0.8", "0", "0", "#N/A", "99", "3", "#REF!" };
            for (int index = 0; index < expected.Length; index++) AssertFormulaCache(sheet, index + 1, 4, expected[index]);
            sheet.CellValue(1, 2, 0d); sheet.CellValue(2, 2, .5d); sheet.CellValue(3, 2, 0d); sheet.CellValue(4, 2, .5d);
            document.Calculate(); AssertFormulaCache(sheet, 7, 4, "3");
            sheet.CellValue(1, 2, .1d); document.Calculate(); AssertFormulaCache(sheet, 1, 4, "#NUM!");
            sheet.CellFormula(1, 2, "NA()"); document.Calculate(); AssertFormulaCache(sheet, 1, 4, "#N/A");
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            AssertFormulaCache(reopened.Sheets[0], 1, 4, "#N/A");
        }

        [Fact]
        public void Test_FormulaEvaluator_OffsetReferencesUseLiveCrossSheetInputsAndBounds() {
            using var document = ExcelDocument.Create();
            var source = document.AddWorksheet("Source data");
            var sheet = document.AddWorksheet("Results");
            source.CellValue(1, 1, 10d); source.CellValue(2, 1, 20d); source.CellValue(3, 1, 30d);
            sheet.CellValue(1, 1, 1d);
            string[] formulas = { "OFFSET('Source data'!A1,A1,0)", "SUM(OFFSET('Source data'!A1,A1,0,2,1))",
                "ROWS(OFFSET('Source data'!A1,0,0,3,1))", "OFFSET('Source data'!A1,3,0)",
                "SUM(OFFSET('Source data'!A1,0,0,,1))", "OFFSET('Source data'!A1,-1,0)",
                "IFERROR(OFFSET('Source data'!A1,-1,0),99)", "OFFSET('Source data'!A1,0,0,0,1)",
                "OFFSET('Source data'!XFD1048576,1,0)", "OFFSET(OFFSET('Source data'!A1,1,0),1,0)",
                "OFFSET('Source data'!A1,1.9,0)" };
            for (int index = 0; index < formulas.Length; index++) sheet.CellFormula(index + 1, 2, formulas[index]);
            Assert.Equal(formulas.Length, document.Calculate());
            string[] expected = { "20", "50", "3", "0", "10", "#REF!", "99", "#REF!", "#REF!", "30", "20" };
            for (int index = 0; index < expected.Length; index++) AssertFormulaCache(sheet, index + 1, 2, expected[index]);
            source.CellFormula(2, 1, "10+40"); sheet.CellValue(1, 1, 0d);
            document.Calculate(); AssertFormulaCache(sheet, 1, 2, "10"); AssertFormulaCache(sheet, 2, 2, "60");
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            AssertFormulaCache(reopened.Sheets[1], 2, 2, "60");
        }

        [Fact]
        public void Test_FormulaEvaluator_RandomBetweenIntegerBoundsAndDependencyCaches() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Random");
            sheet.CellValue(1, 1, 7d);
            sheet.CellFormula(1, 2, "RANDBETWEEN(A1,A1)");
            sheet.CellFormula(2, 2, "RANDBETWEEN(-2,2)");
            sheet.CellFormula(3, 2, "B2-B2");
            sheet.CellFormula(4, 2, "RANDBETWEEN(-2147483648,2147483647)");
            sheet.CellFormula(5, 2, "RANDBETWEEN(2147483647,2147483647)");
            sheet.CellFormula(6, 2, "RANDBETWEEN(-2147483648,-2147483648)");
            sheet.CellFormula(7, 2, "IFERROR(RANDBETWEEN(10,1),99)");
            sheet.CellFormula(8, 2, "RANDBETWEEN(#N/A,1)");
            for (int run = 0; run < 4; run++) {
                Assert.Equal(8, document.Calculate());
                AssertFormulaCache(sheet, 1, 2, "7"); AssertFormulaCache(sheet, 3, 2, "0");
                AssertFormulaCache(sheet, 5, 2, "2147483647"); AssertFormulaCache(sheet, 6, 2, "-2147483648");
                AssertFormulaCache(sheet, 7, 2, "99"); AssertFormulaCache(sheet, 8, 2, "#N/A");
                Assert.True(sheet.TryGetCachedFormulaValue(2, 2, out string? small));
                Assert.InRange(int.Parse(small!), -2, 2);
                Assert.True(sheet.TryGetCachedFormulaValue(4, 2, out string? large));
                Assert.InRange(long.Parse(large!), int.MinValue, int.MaxValue);
            }
            sheet.CellValue(1, 1, 8d); document.Calculate(); AssertFormulaCache(sheet, 1, 2, "8");
        }

        [Fact]
        public void Test_FormulaEvaluator_UnqualifiedProbabilityAndReferencePathsPreserveCaches() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Bounds");
            string[] formulas = { "PROB(A1:A50001,B1:B50001,1)", "SUM(OFFSET(A1,0,0,100001,1))",
                "RANDBETWEEN(7.1,7.1)", "RANDBETWEEN(2147483648,2147483648)", "PROB(A1:A2,B1:B2,1)",
                "OFFSET(A1,0,0,2,1)", "PROB(A1:A2,B1:B2,,2)" };
            for (int index = 0; index < formulas.Length; index++) {
                sheet.CellValue(index + 1, 4, 42d); sheet.CellFormula(index + 1, 4, formulas[index]);
            }
            string nested = "A1";
            for (int depth = 0; depth < 33; depth++) nested = $"OFFSET({nested},0,0)";
            sheet.CellValue(8, 4, 42d); sheet.CellFormula(8, 4, nested);
            Assert.Equal(0, document.Calculate());
            for (int row = 1; row <= 8; row++) AssertFormulaCache(sheet, row, 4, "42");
            sheet.CellValue(1, 6, 99d); sheet.CellFormula(1, 6, "OFFSET(F1,0,0)");
            document.Calculate();
            Assert.False(sheet.TryGetCachedFormulaValue(1, 6, out _)); // Dynamic target cycle cannot consume a stale cache.
        }

        [Fact]
        public void Test_FormulaEvaluator_WorkbookVolatileDependenciesUseOnePassAcrossSheets() {
            using var document = ExcelDocument.Create();
            // Cover both dependency-before-source and source-before-dependency order.
            var first = document.AddWorksheet("First");
            var source = document.AddWorksheet("Random");
            var last = document.AddWorksheet("Last");
            source.CellValue(1, 2, 7d);
            source.CellFormula(1, 1, "RANDBETWEEN(-2147483648,2147483647)");
            source.CellFormula(2, 1, "RANDBETWEEN(B1,B1)");
            first.CellFormula(1, 1, "Random!A1");
            first.CellFormula(2, 1, "Random!A2");
            last.CellFormula(1, 1, "Random!A1");
            last.CellFormula(2, 1, "Random!A2");
            last.CellFormula(3, 1, "First!A1-Random!A1");
            for (int run = 0; run < 4; run++) {
                source.CellValue(1, 2, 7d + run);
                Assert.Equal(7, document.Calculate());
                Assert.True(source.TryGetCachedFormulaValue(1, 1, out string? random));
                AssertFormulaCache(first, 1, 1, random!);
                AssertFormulaCache(last, 1, 1, random!);
                AssertFormulaCache(last, 3, 1, "0");
                AssertFormulaCache(first, 2, 1, (7 + run).ToString());
                AssertFormulaCache(last, 2, 1, (7 + run).ToString());
            }
            using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.True(reopened.Sheets[1].TryGetCachedFormulaValue(1, 1, out string? savedRandom));
            AssertFormulaCache(reopened.Sheets[0], 1, 1, savedRandom!);
            AssertFormulaCache(reopened.Sheets[2], 1, 1, savedRandom!);
        }

        [Fact]
        public void Test_FormulaEvaluator_OffsetErrorsPropagateThroughReferenceConsumers() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("References");
            sheet.CellValue(1, 1, 1d); sheet.CellValue(1, 2, 1d);
            string[] formulas = { "OFFSET(OFFSET(A1,-1,0),0,0)", "ROWS(OFFSET(A1,-1,0))",
                "COLUMNS(OFFSET(A1,-1,0))", "ROW(OFFSET(A1,-1,0))", "COLUMN(OFFSET(A1,-1,0))",
                "PROB(OFFSET(A1,-1,0),B1,1)", "PROB(A1,OFFSET(B1,-1,0),1)",
                "SUM(OFFSET(A1,-1,0))", "COUNTBLANK(OFFSET(A1,-1,0))",
                "VLOOKUP(1,OFFSET(A1,-1,0),1,FALSE)", "INDEX(OFFSET(A1,-1,0),1)",
                "CONCAT(OFFSET(A1,-1,0))", "IFNA(ROWS(OFFSET(A1,-1,0)),99)",
                "IFERROR(ROWS(OFFSET(A1,-1,0)),99)", "ISERROR(ROWS(OFFSET(A1,-1,0)))" };
            for (int index = 0; index < formulas.Length; index++) {
                sheet.CellValue(index + 1, 4, 42d); sheet.CellFormula(index + 1, 4, formulas[index]);
            }
            Assert.Equal(formulas.Length, document.Calculate());
            for (int row = 1; row <= 13; row++) AssertFormulaCache(sheet, row, 4, "#REF!");
            AssertFormulaCache(sheet, 14, 4, "99"); AssertFormulaCache(sheet, 15, 4, "1");
        }

        [Fact]
        public void Test_FormulaEvaluator_UnsupportedInputsDoNotBecomeErrorsInsideWrappers() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Unsupported");
            string[] expressions = { "RANDBETWEEN(7.1,7.1)", "RANDBETWEEN(2147483648,2147483648)",
                "PROB(A1:A2,B1:B2,1)", "PROB(A1:A50001,B1:B50001,1)",
                "OFFSET(A1,0,0,2,1)", "SUM(OFFSET(A1,0,0,100001,1))", "SUM(A1:A100001)",
                "XLOOKUP(1,A1:A2,B1:B2,,0,2)", "RANDBETWEEN(7.1,7.1)/0", "UNKNOWN()/0" };
            int row = 0;
            foreach (string expression in expressions) {
                foreach (string wrapper in new[] { "IFERROR({0},99)", "IFNA({0},99)", "ISERROR({0})", "ISERR({0})", "IF(ISERROR({0}),99,1)" }) {
                    row++;
                    sheet.CellValue(row, 4, 42d); sheet.CellFormula(row, 4, string.Format(wrapper, expression));
                }
            }
            // Unselected branches retain normal lazy evaluation.
            sheet.CellFormula(1, 6, "IF(FALSE,RANDBETWEEN(7.1,7.1),7)");
            Assert.Equal(1, document.Calculate());
            for (int index = 1; index <= row; index++) AssertFormulaCache(sheet, index, 4, "42");
            AssertFormulaCache(sheet, 1, 6, "7");
        }

        [Fact]
        public void Test_FormulaEvaluator_CompletedLookupSearchesReturnTypedMissingMatches() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Search");
            sheet.CellValue(1, 1, 1d); sheet.CellValue(1, 2, 10d);
            string[] searches = { "VLOOKUP(2,A1:B1,2,FALSE)", "HLOOKUP(2,A1:A2,2,FALSE)",
                "XLOOKUP(2,A1:A1,B1:B1)", "MATCH(2,A1:A1,0)", "XMATCH(2,A1:A1,0)" };
            int row = 0;
            foreach (string search in searches) {
                sheet.CellFormula(++row, 4, search);
                sheet.CellFormula(++row, 4, $"IFNA({search},99)");
            }
            Assert.Equal(row, document.Calculate());
            for (int index = 1; index <= row; index++) AssertFormulaCache(sheet, index, 4, index % 2 == 1 ? "#N/A" : "99");
            sheet.CellFormula(1, 1, "UNKNOWN()"); sheet.ClearCachedFormulaResults();
            Assert.Equal(0, document.Calculate()); // An unresolved search vector cannot prove a missing match.
            for (int index = 1; index <= row; index++) Assert.False(sheet.TryGetCachedFormulaValue(index, 4, out _));
        }

        [Fact]
        public void Test_FormulaEvaluator_ExactLookupStopsBeforeUnvisitedErrorsOrUnresolvedEntries() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("SearchOrder");
            sheet.CellValue(1, 1, 1d); sheet.CellFormula(2, 1, "NA()");
            sheet.CellValue(1, 2, 10d); sheet.CellValue(2, 2, 20d);
            string[] formulas = { "XLOOKUP(1,A1:A2,B1:B2)", "MATCH(1,A1:A2,0)", "XMATCH(1,A1:A2,0)",
                "XLOOKUP(1,A1:A2,B1:B2,99,0,-1)", "XMATCH(1,A1:A2,0,-1)",
                "VLOOKUP(1,A1:B2,2,FALSE)", "HLOOKUP(1,A1:B2,2,FALSE)" };
            for (int index = 0; index < formulas.Length; index++) sheet.CellFormula(index + 1, 4, formulas[index]);
            document.Calculate();
            string[] forward = { "10", "1", "1", "#N/A", "#N/A", "10", "#N/A" };
            // HLOOKUP finds A1 then returns its typed error at A2; no fallback masks it.
            for (int index = 0; index < forward.Length; index++) AssertFormulaCache(sheet, index + 1, 4, forward[index]);
            sheet.CellFormula(1, 1, "NA()"); sheet.CellFormula(2, 1, "1");
            document.Calculate();
            string[] reverse = { "#N/A", "#N/A", "#N/A", "20", "2", "#N/A", "#N/A" };
            for (int index = 0; index < reverse.Length; index++) AssertFormulaCache(sheet, index + 1, 4, reverse[index]);
            sheet.CellFormula(1, 1, "UNKNOWN()"); sheet.ClearCachedFormulaResults();
            document.Calculate();
            AssertFormulaCache(sheet, 4, 4, "20"); AssertFormulaCache(sheet, 5, 4, "2");
            Assert.False(sheet.TryGetCachedFormulaValue(1, 4, out _));
            sheet.CellFormula(1, 1, "1"); sheet.CellFormula(2, 1, "UNKNOWN()"); sheet.ClearCachedFormulaResults();
            document.Calculate();
            AssertFormulaCache(sheet, 1, 4, "10"); AssertFormulaCache(sheet, 2, 4, "1"); AssertFormulaCache(sheet, 3, 4, "1");
            Assert.False(sheet.TryGetCachedFormulaValue(4, 4, out _));
        }

        private static void AssertFormulaCache(ExcelSheet sheet, int row, int column, string expected) {
            Assert.True(sheet.TryGetCachedFormulaValue(row, column, out string? actual), $"Missing cache at {row},{column}");
            Assert.Equal(expected, actual);
        }
    }
}
