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

        [Theory]
        [InlineData("INDEX(A1:A3,2)")]
        [InlineData("OFFSET(A1,1,0)")]
        public void FormulaEvaluator_BlankReferencesCoerceAtNumericAndCriteriaConsumers(string reference) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Consumers");
            sheet.CellValue(1, 1, 1d);
            sheet.CellValue(3, 1, 3d);
            for (int row = 1; row <= 3; row++) {
                sheet.CellValue(row, 2, row - 1d);
                sheet.CellValue(row, 3, row * 10d);
            }
            (string Formula, string Expected)[] cases = {
                ($"SQRT({reference})", "0"),
                ($"ABS({reference})", "0"),
                ($"ROUND(10.4,{reference})", "10"),
                ($"DAY(DATE(2026,11,{reference}))", "31"),
                ($"LEFT(\"ab\",{reference})", ""),
                ($"OFFSET(A1,{reference},0)", "1"),
                ($"XLOOKUP(1,B1:B3,C1:C3,,{reference})", "20"),
                ($"PROB(B1:B3,{{0.2;0.3;0.5}},{reference})", "0.2"),
                ($"TEXTJOIN(\",\",{reference},\"\",\"x\")", ",x"),
                ($"RANDBETWEEN({reference},0)", "0"),
                ($"IPMT({reference},1,2,100)", "0"),
                ($"PPMT({reference},1,2,100)", "-50")
            };
            for (int index = 0; index < cases.Length; index++) sheet.CellFormula(index + 1, 5, cases[index].Formula);
            Assert.Equal(cases.Length, document.Calculate());
            for (int index = 0; index < cases.Length; index++) AssertFormulaCache(sheet, index + 1, 5, cases[index].Expected);
        }

        [Theory]
        [InlineData("INDEX(A1:A3,2)")]
        [InlineData("OFFSET(A1,1,0)")]
        public void FormulaEvaluator_BlankCriteriaReferencesUseZeroRatherThanFormulaText(string reference) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Criteria");
            sheet.CellValue(1, 1, 1d);
            sheet.CellValue(3, 1, 3d);
            for (int row = 1; row <= 3; row++) {
                sheet.CellValue(row, 2, row - 1d);
                sheet.CellValue(row, 3, row * 10d);
            }
            (string Formula, string Expected)[] cases = {
                ($"COUNTIF(B1:B3,{reference})", "1"),
                ($"SUMIF(B1:B3,{reference},C1:C3)", "10"),
                ($"AVERAGEIF(B1:B3,{reference},C1:C3)", "10"),
                ($"COUNTIFS(B1:B3,{reference})", "1"),
                ($"SUMIFS(C1:C3,B1:B3,{reference})", "10")
            };
            for (int index = 0; index < cases.Length; index++) sheet.CellFormula(index + 1, 5, cases[index].Formula);
            Assert.Equal(cases.Length, document.Calculate());
            for (int index = 0; index < cases.Length; index++) AssertFormulaCache(sheet, index + 1, 5, cases[index].Expected);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void FormulaEvaluator_ArrayBlankOutputsReplaceCachesWithNumericZero(bool dynamic) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Arrays");
            sheet.CellValue(1, 1, 1d);
            sheet.CellValue(2, 1, 9d);
            if (dynamic) sheet.SetDynamicArrayFormula("G1", "A1:A2");
            else sheet.SetArrayFormula("G1:G2", "A1:A2");
            sheet.CellFormula(1, 10, "ISBLANK(G2)");
            Assert.Equal(2, document.Calculate());
            Assert.Equal(9d, sheet.CellAt(2, 7).GetValue().Value);
            sheet.CellAt(2, 1).Clear();
            Assert.Equal(2, document.Calculate());
            Assert.Equal(0d, sheet.CellAt(2, 7).GetValue().Value);
            AssertFormulaCache(sheet, 1, 10, "0");
            using var saved = new MemoryStream();
            document.Save(saved);
            saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.Equal(0d, reopened.Sheets[0].CellAt(2, 7).GetValue().Value);
            AssertFormulaCache(reopened.Sheets[0], 1, 10, "0");
        }
    }
}
