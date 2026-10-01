using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_FormulaEvaluator_MatchesDesktopExcelScalarOracle() {
            string source = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Documents", "ExcelFormulaCorpus", "scalar-expressions.xlsx");
            string path = Path.Combine(_directoryWithFiles, "ExcelScalarOracle.xlsx");
            File.Copy(source, path);
            using var document = ExcelDocument.Load(path);
            var expected = document.InspectFormulas().Formulas;
            Assert.Equal(39, expected.Count);
            Assert.Equal(expected.Count, document.Calculate());
            var actual = document.InspectFormulas().Formulas;
            for (int index = 0; index < expected.Count; index++) {
                Assert.Equal(expected[index].CellReference, actual[index].CellReference);
                Assert.Equal(expected[index].CachedValue, actual[index].CachedValue);
            }
        }

        [Fact]
        public void Test_FormulaEvaluator_PreservesBooleanCachesThroughSavedReadback() {
            string path = Path.Combine(_directoryWithFiles, "BooleanFormulaCaches.xlsx");
            using (var document = ExcelDocument.Create(path)) {
                var sheet = document.AddWorksheet("Booleans");
                sheet.CellFormula(1, 1, "1=1");
                sheet.CellFormula(2, 1, "AND(TRUE,TRUE)");
                sheet.CellFormula(3, 1, "EXACT(\"a\",\"b\")");
                sheet.CellFormula(4, 1, "ISNUMBER(1)");
                sheet.CellFormula(5, 1, "ISFORMULA(A1)");
                Assert.Equal(5, document.Calculate());
                document.Save();
            }
            using var reader = ExcelDocumentReader.Open(path);
            var values = reader.GetSheet("Booleans").ReadRange("A1:A5");
            Assert.True(Assert.IsType<bool>(values[0, 0]));
            Assert.True(Assert.IsType<bool>(values[1, 0]));
            Assert.False(Assert.IsType<bool>(values[2, 0]));
            Assert.True(Assert.IsType<bool>(values[3, 0]));
            Assert.True(Assert.IsType<bool>(values[4, 0]));
        }

        [Theory]
        [InlineData("1+2*3", "7")]
        [InlineData("(1+2)*3", "9")]
        [InlineData("2^3^2", "64")]
        [InlineData("-2^2", "4")]
        [InlineData("2^-2", "0.25")]
        [InlineData("50%*8", "4")]
        [InlineData("SUM(1,2)*3", "9")]
        [InlineData("SUM(1+2,3*4)", "15")]
        [InlineData("IF(1+2*3=7,2^3,0)", "8")]
        [InlineData("\"a\"&\"b\"", "ab")]
        [InlineData("LEN(\"a\"&\"bc\")+1", "4")]
        [InlineData("1e-3*1000", "1")]
        [InlineData("1/(2-2)", "#DIV/0!")]
        [InlineData("IFERROR(1/(2-2),99)", "99")]
        [InlineData("\"abc\"+1", "#VALUE!")]
        [InlineData("#DIV/0!+1", "#DIV/0!")]
        [InlineData("10^400", "#NUM!")]
        public void Test_FormulaEvaluator_ComposableScalarExpressions(string formula, string expected) {
            string path = Path.Combine(_directoryWithFiles, "ScalarExpressions.xlsx");
            using var document = ExcelDocument.Create(path);
            var sheet = document.AddWorksheet("Expressions");
            sheet.CellFormula(1, 1, formula);
            Assert.Equal(1, document.Calculate());
            Assert.Equal(expected, Assert.Single(sheet.InspectFormulas().Formulas).CachedValue);
            document.Save();
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_FormulaEvaluator_ComposedReferencesAndBoundedSyntax() {
            using var document = ExcelDocument.Create(Path.Combine(_directoryWithFiles, "ScalarReferences.xlsx"));
            var sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 3);
            sheet.CellFormula(1, 2, "(A1+SUM(A1,2))*2");
            sheet.CellFormula(1, 3, new string('(', 140) + "1" + new string(')', 140));
            sheet.CellFormula(1, 4, "1+2*");
            Assert.Equal(1, document.Calculate());
            Assert.Contains(sheet.InspectFormulas().Formulas, f => f.CellReference == "B1" && f.CachedValue == "16");
            Assert.Contains(sheet.InspectFormulas().Formulas, f => f.CellReference == "C1" && !f.IsSupportedByOfficeIMO);
        }
    }
}
