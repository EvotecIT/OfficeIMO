using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("number", 1, "3")]
        [InlineData("text", 0, "42")]
        [InlineData("missing", 0, "42")]
        public void FormulaEvaluator_SumProductPreservesNumericCachesWithoutCoercingUnevaluatedNonnumericInputs(
            string cacheKind, int calculated, string expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Unresolved");
            if (cacheKind == "number") sheet.CellValue(1, 1, 1d);
            else if (cacheKind == "text") sheet.CellValue(1, 1, string.Empty);
            sheet.CellFormula(1, 1, "UNKNOWN()");
            sheet.CellValue(2, 1, 2d);
            sheet.CellValue(1, 4, 42d);
            sheet.CellFormula(1, 4, "IFERROR(SUMPRODUCT(A1:A2),99)");
            Assert.Equal(calculated, document.Calculate());
            AssertFormulaCache(sheet, 1, 4, expected);
        }

        [Theory]
        [InlineData("IF(NA(),1,2)")]
        [InlineData("IFS(NA(),1,TRUE,2)")]
        [InlineData("SWITCH(NA(),1,10,20)")]
        [InlineData("SWITCH(1,NA(),10,20)")]
        [InlineData("CHOOSE(NA(),1,2)")]
        [InlineData("LEN(NA())")]
        [InlineData("TRIM(NA())")]
        [InlineData("UPPER(NA())")]
        [InlineData("LEFT(NA(),2)")]
        [InlineData("LEFT(\"abc\",NA())")]
        [InlineData("MID(\"abc\",NA(),1)")]
        [InlineData("SUBSTITUTE(\"abc\",\"a\",\"b\",NA())")]
        [InlineData("FIND(NA(),\"abc\")")]
        [InlineData("VALUE(NA())")]
        [InlineData("TEXT(NA(),\"0\")")]
        [InlineData("TEXT(1,NA())")]
        [InlineData("REPT(\"x\",NA())")]
        [InlineData("TEXTBEFORE(\"a:b\",\":\",1,0,NA())")]
        [InlineData("TEXTAFTER(\"a:b\",\":\",1,0,NA())")]
        [InlineData("SUMPRODUCT(A1:A2,B1:B2)")]
        [InlineData("SUMPRODUCT(A1:A2,B1:B3)")]
        public void Test_FormulaEvaluator_ComposedErrorsRemainTyped(string expression) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Errors");
            sheet.CellValue(1, 1, 1d);
            sheet.CellValue(2, 1, 2d);
            sheet.CellFormula(1, 2, "NA()");
            sheet.CellValue(2, 2, 2d);
            sheet.CellFormula(1, 4, expression);
            sheet.CellFormula(2, 4, "IFERROR(D1,99)");
            sheet.CellFormula(3, 4, "IFNA(D1,7)");
            Assert.Equal(4, document.Calculate());
            AssertFormulaCache(sheet, 1, 4, "#N/A");
            AssertFormulaCache(sheet, 2, 4, "99");
            AssertFormulaCache(sheet, 3, 4, "7");
        }

        [Theory]
        [InlineData("IF(UNKNOWN(),1,2)")]
        [InlineData("LEFT(NA(),UNKNOWN())")]
        [InlineData("MID(NA(),-1,2)")]
        [InlineData("REPT(NA(),1000000)")]
        [InlineData("SUMPRODUCT(A1:A2,B1:B100001)")]
        public void Test_FormulaEvaluator_UnsupportedCompositionDoesNotActivateFallback(string expression) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Errors");
            sheet.CellValue(1, 1, 1d);
            sheet.CellValue(2, 1, 2d);
            sheet.CellFormula(1, 2, "NA()");
            sheet.CellValue(1, 4, 42d);
            sheet.CellFormula(1, 4, $"IFERROR({expression},99)");
            Assert.Equal(1, document.Calculate());
            AssertFormulaCache(sheet, 1, 4, "42");
        }

        [Fact]
        public void Test_FormulaEvaluator_TextBeforeDoesNotEvaluateUnusedErrorFallback() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Text");
            sheet.CellFormula(1, 1, "TEXTBEFORE(\"a:b\",\":\",1,0,0,NA())");
            Assert.Equal(1, document.Calculate());
            AssertFormulaCache(sheet, 1, 1, "a");
        }

        [Fact]
        public void Test_FormulaEvaluator_BlankReferenceConditionsUseNumericZero() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Blank");
            sheet.CellValue(2, 1, "");
            string[] formulas = { "IF(A1=0,1,2)", "IF(0=A1,1,2)", "IF(A1<1,1,2)",
                "IF(A1=\"0\",1,2)", "IF(A2=0,1,2)", "IFS(A1=0,1,TRUE,2)",
                "IF(A1,1,2)", "AND(A1=0,TRUE)", "OR(A1=0,FALSE)", "NOT(A1=0)" };
            string[] expected = { "1", "1", "1", "2", "2", "1", "2", "1", "1", "0" };
            for (int index = 0; index < formulas.Length; index++) sheet.CellFormula(index + 1, 3, formulas[index]);
            Assert.Equal(formulas.Length, document.Calculate());
            for (int index = 0; index < formulas.Length; index++) AssertFormulaCache(sheet, index + 1, 3, expected[index]);
        }

        [Fact]
        public void Test_FormulaEvaluator_LookupErrorsPropagateInsideArithmetic() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Arithmetic");
            sheet.CellValue(1, 1, 1d);
            sheet.CellValue(2, 1, 2d);
            sheet.CellValue(1, 2, 10d);
            sheet.CellValue(2, 2, 20d);
            string[] expressions = { "INDEX(A1:A2,3)+1", "VLOOKUP(9,A1:B2,2,FALSE)+1",
                "HLOOKUP(9,A1:B2,2,FALSE)+1", "XLOOKUP(9,A1:A2,B1:B2)+1" };
            for (int index = 0; index < expressions.Length; index++) sheet.CellFormula(index + 1, 4, $"IFERROR({expressions[index]},99)");
            Assert.Equal(expressions.Length, document.Calculate());
            for (int index = 0; index < expressions.Length; index++) AssertFormulaCache(sheet, index + 1, 4, "99");
        }
    }
}
