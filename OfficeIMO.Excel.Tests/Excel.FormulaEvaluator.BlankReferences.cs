using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("index", "INDEX(A1:A3,2)")]
        [InlineData("offset", "OFFSET(A1,1,0)")]
        [InlineData("name", "EmptyCell")]
        public void FormulaEvaluator_BlankReferenceFunctionsRetainContextUntilCellEvaluation(string caseId, string reference) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("References");
            sheet.CellValue(1, 1, 1d);
            sheet.CellValue(3, 1, 3d);
            document.SetNamedRange("EmptyCell", "'References'!$A$2");
            (string Formula, string Expected)[] cases = {
                (reference, "0"),
                ($"ISBLANK({reference})", "1"),
                ($"{reference}=\"\"", "1"),
                ($"{reference}&\"x\"", "x"),
                ($"LEN({reference})", "0"),
                ($"ISNUMBER({reference})", "0"),
                ("ISBLANK(D1)", "0"),
                ($"IFERROR({reference},UNKNOWN())", "0"),
                ($"ISBLANK(IFERROR({reference},77))", "0"),
                ($"IFERROR({reference},77)&\"x\"", "0x"),
                ($"ISBLANK(IF(TRUE,{reference},7))", "0"),
                ($"IF(TRUE,{reference},7)&\"x\"", "0x"),
                ($"SUM({reference},2)", "2"),
                ($"AVERAGE({reference},2)", "2"),
                ($"TEXT({reference},\"0\")", "0"),
                ($"ISBLANK(CHOOSE(1,{reference}))", "1"),
                ($"COUNTA({reference})", "0"),
                ($"COUNT({reference})", "0"),
                ($"COUNTBLANK({reference})", "1"),
                ($"IFS(TRUE,{reference})&\"x\"", "0x"),
                ($"SWITCH(1,1,{reference})&\"x\"", "0x"),
                ("CHOOSE(2,10,,20)", "0"),
                ("IF(TRUE,,7)", "0"),
                ("IF(FALSE,7)", "0")
            };
            for (int index = 0; index < cases.Length; index++) sheet.CellFormula(index + 1, 4, cases[index].Formula);
            Assert.True(document.Calculate() == cases.Length, caseId);
            for (int index = 0; index < cases.Length; index++) AssertFormulaCache(sheet, index + 1, 4, cases[index].Expected);
        }
    }
}
