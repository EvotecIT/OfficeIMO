using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("index", "INDEX(A1:A3,2)")]
        [InlineData("offset", "OFFSET(A1,1,0)")]
        [InlineData("name", "EmptyCell")]
        [InlineData("xlookup-range", "XLOOKUP(1,A1:A3,A2:A4)")]
        [InlineData("xlookup-array", "XLOOKUP(2,{1;2;3},A1:A3)")]
        [InlineData("xlookup-name", "EmptyLookup")]
        public void FormulaEvaluator_BlankReferenceFunctionsRetainContextUntilCellEvaluation(string caseId, string reference) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("References");
            sheet.CellValue(1, 1, 1d);
            sheet.CellValue(3, 1, 3d);
            document.SetNamedRange("EmptyCell", "'References'!$A$2");
            document.SetNamedFormula("EmptyLookup", "XLOOKUP(2,{1;2;3},'References'!A1:A3)", save: false);
            (string Formula, string Expected)[] cases = {
                (reference, "0"),
                ($"ISBLANK({reference})", "1"),
                ($"{reference}=\"\"", "1"),
                ($"{reference}&\"x\"", "x"),
                ($"LEN({reference})", "0"),
                ($"ISNUMBER({reference})", "0"),
                ("ISBLANK(D1)", "0"),
                ($"IFERROR({reference},UNKNOWN())", "0"),
                ($"ISBLANK(IFERROR({reference},77))", "1"),
                ($"IFERROR({reference},77)&\"x\"", "x"),
                ($"ISBLANK(IF(TRUE,{reference},7))", "1"),
                ($"IF(TRUE,{reference},7)&\"x\"", "x"),
                ($"SUM({reference},2)", "2"),
                ($"AVERAGE({reference},2)", "2"),
                ($"TEXT({reference},\"0\")", "0"),
                ($"ISBLANK(CHOOSE(1,{reference}))", "1"),
                ($"COUNTA({reference})", "0"),
                ($"COUNT({reference})", "0"),
                ($"COUNTBLANK({reference})", "1"),
                ($"IFS(TRUE,{reference})&\"x\"", "x"),
                ($"SWITCH(1,1,{reference})&\"x\"", "x"),
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
        [InlineData("XLOOKUP(2,{1;2;3},A1:A3)")]
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
                ($"VALUE({reference})", "0"),
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
        [InlineData("XLOOKUP(2,{1;2;3},A1:A3)")]
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
        [InlineData("H1")]
        [InlineData("INDEX(H1:H2,1)")]
        [InlineData("OFFSET(H1,0,0)")]
        [InlineData("EmptyKey")]
        public void FormulaEvaluator_BlankLookupKeysRespectLegacyAndModernSearchSemantics(string key) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Keys");
            document.SetNamedRange("EmptyKey", "Keys!$H$1", save: false);
            sheet.CellValue(2, 1, 5d);
            sheet.CellValue(3, 1, 0d);
            sheet.CellValue(4, 1, "0");
            sheet.CellValue(5, 1, false);
            sheet.CellValue(1, 2, 100d);
            sheet.CellValue(2, 2, 500d);
            sheet.CellValue(3, 2, 200d);
            sheet.CellValue(1, 5, 5d);
            sheet.CellValue(1, 6, 0d);
            sheet.CellValue(2, 4, 100d);
            sheet.CellValue(2, 5, 500d);
            sheet.CellValue(2, 6, 200d);
            for (int row = 1; row <= 3; row++) {
                sheet.CellValue(row, 10, (row - 2) * 5d);
                sheet.CellValue(row, 12, (2 - row) * 5d);
            }
            (string Formula, string Expected)[] cases = {
                ($"MATCH({key},A1:A3,0)", "3"),
                ($"XMATCH({key},A1:A3)", "1"),
                ($"VLOOKUP({key},A1:B3,2,FALSE)", "200"),
                ($"HLOOKUP({key},D1:F2,2,FALSE)", "200"),
                ($"XLOOKUP({key},A1:A3,B1:B3)", "100"),
                ($"IFERROR(MATCH({key},A1:A2,0),77)", "77"),
                ($"IFERROR(XMATCH({key},A1:A2),77)", "1"),
                ($"IFERROR(VLOOKUP({key},A1:B2,2,FALSE),77)", "77"),
                ($"IFERROR(HLOOKUP({key},D1:E2,2,FALSE),77)", "77"),
                ($"IFERROR(XLOOKUP({key},A1:A2,B1:B2),77)", "100"),
                ($"XMATCH({key},J1:J3,0,2)", "#N/A"),
                ($"XMATCH({key},L1:L3,0,-2)", "#N/A"),
                ($"XLOOKUP({key},J1:J3,B1:B3,,0,2)", "#N/A"),
                ($"XLOOKUP({key},L1:L3,B1:B3,,0,-2)", "#N/A"),
                ("MATCH(0,A1:A5,0)", "3"),
                ("MATCH(\"0\",A1:A5,0)", "4"),
                ("MATCH(FALSE,A1:A5,0)", "5")
            };
            for (int index = 0; index < cases.Length; index++) sheet.CellFormula(index + 1, 14, cases[index].Formula);
            Assert.Equal(cases.Length, document.Calculate());
            for (int index = 0; index < cases.Length; index++) AssertFormulaCache(sheet, index + 1, 14, cases[index].Expected);
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
