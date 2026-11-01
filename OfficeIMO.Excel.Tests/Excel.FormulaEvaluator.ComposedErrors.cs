using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
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
        [InlineData("SUMPRODUCT(A1:A2,B1:B2)")]
        public void Test_FormulaEvaluator_ComposedErrorsRemainTyped(string expression) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Errors");
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
        [InlineData("SUMPRODUCT(A1:A2,B1:B3)")]
        public void Test_FormulaEvaluator_UnsupportedCompositionDoesNotActivateFallback(string expression) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Errors");
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
            var sheet = document.AddWorksheet("Text");
            sheet.CellFormula(1, 1, "TEXTBEFORE(\"a:b\",\":\",1,0,0,NA())");
            Assert.Equal(1, document.Calculate());
            AssertFormulaCache(sheet, 1, 1, "a");
        }
    }
}
