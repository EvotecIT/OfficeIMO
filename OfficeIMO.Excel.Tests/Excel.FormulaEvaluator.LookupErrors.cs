using OfficeIMO.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_FormulaEvaluator_IndexMatchFallbackFeedsDependentTotalsAndPersists() {
            using var document = ExcelDocument.Create();
            var products = document.AddWorksheet("Products");
            products.CellValue(2, 1, "P1");
            products.CellValue(2, 2, "Hardware");
            var orders = document.AddWorksheet("Orders");
            orders.CellValue(2, 1, "P1");
            orders.CellValue(3, 1, "PX");
            orders.CellValue(2, 3, 10d);
            orders.CellValue(3, 3, 25d);
            for (int row = 2; row <= 3; row++) {
                orders.CellFormula(row, 2, $"IFERROR(INDEX(Products!$B$2:$B$2,MATCH(A{row},Products!$A$2:$A$2,0)),\"Unknown\")");
            }
            orders.CellFormula(5, 2, "COUNTIFS(B2:B3,\"Unknown\")");
            orders.CellFormula(6, 2, "SUMIFS(C2:C3,B2:B3,\"Unknown\")");
            orders.CellFormula(7, 2, "B5+B6");
            Assert.Equal(5, document.Calculate());
            string[] expected = { "Hardware", "Unknown", "1", "25", "26" };
            int[] rows = { 2, 3, 5, 6, 7 };
            for (int index = 0; index < rows.Length; index++) AssertFormulaCache(orders, rows[index], 2, expected[index]);
            Assert.All(document.InspectFormulas().Formulas, formula => Assert.True(formula.IsSupportedByOfficeIMO));
            using var saved = new MemoryStream();
            document.Save(saved);
            saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            for (int index = 0; index < rows.Length; index++) AssertFormulaCache(reopened.Sheets[1], rows[index], 2, expected[index]);
            Assert.Empty(reopened.ValidateOpenXml());
            products.CellValue(2, 1, "PX");
            Assert.Equal(5, document.Calculate());
            AssertFormulaCache(orders, 2, 2, "Unknown");
            AssertFormulaCache(orders, 3, 2, "Hardware");
            AssertFormulaCache(orders, 6, 2, "10");
        }

        [Theory]
        [InlineData("INDEX(A1:A2,MATCH(9,A1:A2,0))", "#N/A")]
        [InlineData("INDEX(NA(),1)", "#N/A")]
        [InlineData("INDEX(A1:B2,1,NA())", "#N/A")]
        [InlineData("INDEX(A1:B2,1/0,1)", "#DIV/0!")]
        [InlineData("INDEX(A1:A2,3)", "#REF!")]
        [InlineData("INDEX(A1:B2,1,3)", "#REF!")]
        [InlineData("INDEX(A1:A2,-1)", "#VALUE!")]
        [InlineData("VLOOKUP(1,A1:B2,NA(),FALSE)", "#N/A")]
        [InlineData("HLOOKUP(1,A1:B2,NA(),FALSE)", "#N/A")]
        [InlineData("VLOOKUP(1,A1:B2,2,NA())", "#N/A")]
        [InlineData("HLOOKUP(1,A1:B2,2,1/0)", "#DIV/0!")]
        [InlineData("VLOOKUP(1,A1:B2,3,FALSE)", "#REF!")]
        [InlineData("HLOOKUP(1,A1:B2,0,FALSE)", "#VALUE!")]
        [InlineData("MATCH(1,A1:A2,NA())", "#N/A")]
        [InlineData("XMATCH(1,A1:A2,0,NA())", "#N/A")]
        [InlineData("XMATCH(1,A1:A2,NA())", "#N/A")]
        [InlineData("XLOOKUP(1,A1:A2,B1:B2,,NA())", "#N/A")]
        [InlineData("XLOOKUP(1,A1:A2,B1:B2,,0,NA())", "#N/A")]
        public void Test_FormulaEvaluator_LookupErrorsRemainTypedAndCatchable(string formula, string expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Lookup");
            sheet.CellValue(1, 1, 1d);
            sheet.CellValue(2, 1, 2d);
            sheet.CellValue(1, 2, 10d);
            sheet.CellValue(2, 2, 20d);
            sheet.CellFormula(1, 4, formula);
            sheet.CellFormula(2, 4, "IFERROR(D1,99)");
            sheet.CellFormula(3, 4, "IFNA(D1,FALSE)");
            sheet.CellFormula(4, 4, "ISERROR(D1)");
            Assert.Equal(4, document.Calculate());
            AssertFormulaCache(sheet, 1, 4, expected);
            AssertFormulaCache(sheet, 2, 4, "99");
            AssertFormulaCache(sheet, 3, 4, expected == "#N/A" ? "0" : expected);
            AssertFormulaCache(sheet, 4, 4, "1");
            using var saved = new MemoryStream();
            document.Save(saved);
            saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            AssertFormulaCache(reopened.Sheets[0], 1, 4, expected);
            saved.Position = 0;
            using var package = SpreadsheetDocument.Open(saved, false);
            var cells = package.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().ToArray();
            Assert.Equal(CellValues.Error, cells.Single(cell => cell.CellReference == "D1").DataType!.Value);
            Assert.Equal(expected == "#N/A" ? CellValues.Boolean : CellValues.Error,
                cells.Single(cell => cell.CellReference == "D3").DataType!.Value);
        }

        [Theory]
        [InlineData("INDEX(B1:B2,2)")]
        [InlineData("VLOOKUP(2,A1:B2,2,FALSE)")]
        [InlineData("HLOOKUP(1,A1:B3,3,FALSE)")]
        [InlineData("XLOOKUP(2,A1:A2,B1:B2)")]
        public void Test_FormulaEvaluator_SelectedBlankLookupIsZeroAndFallbackIsLazy(string formula) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Lookup");
            sheet.CellValue(1, 1, 1d);
            sheet.CellValue(2, 1, 2d);
            sheet.CellFormula(1, 4, formula);
            sheet.CellFormula(2, 4, "IFERROR(D1,UNKNOWN())");
            sheet.CellFormula(3, 4, "D1+1");
            Assert.Equal(3, document.Calculate());
            AssertFormulaCache(sheet, 1, 4, "0");
            AssertFormulaCache(sheet, 2, 4, "0");
            AssertFormulaCache(sheet, 3, 4, "1");
        }

        [Theory]
        [InlineData("99", "99", false)]
        [InlineData("FALSE", "0", false)]
        [InlineData("\"#N/A\"", "#N/A", true)]
        [InlineData("\"\"", "", true)]
        public void Test_FormulaEvaluator_IndexMatchFallbackPreservesValueType(string fallback, string expected, bool isText) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Lookup");
            sheet.CellValue(1, 1, 1d);
            sheet.CellFormula(1, 4, $"IFERROR(INDEX(A1:A1,MATCH(9,A1:A1,0)),{fallback})");
            sheet.CellFormula(2, 4, "ISTEXT(D1)");
            sheet.CellFormula(3, 4, "ISERROR(D1)");
            Assert.Equal(3, document.Calculate());
            AssertFormulaCache(sheet, 1, 4, expected);
            AssertFormulaCache(sheet, 2, 4, isText ? "1" : "0");
            AssertFormulaCache(sheet, 3, 4, "0");
        }

        [Theory]
        [InlineData("INDEX(A1:A2,UNKNOWN())")]
        [InlineData("INDEX(A1:B2,NA(),UNKNOWN())")]
        [InlineData("INDEX(A1:A2,0)")]
        [InlineData("INDEX(A1:A2,,2)")]
        [InlineData("INDEX(A1:B2,1,)")]
        [InlineData("INDEX(A1:B1,1,)")]
        [InlineData("INDEX(A1:B2,NA())")]
        [InlineData("INDEX(A1:B2,NA(),0)")]
        [InlineData("INDEX(A1:B2,0,NA())")]
        [InlineData("VLOOKUP(1,A1:B2,NA(),TRUE)")]
        [InlineData("VLOOKUP(1,A1:B2,2,0.00000001)")]
        [InlineData("VLOOKUP(1,A1:B2,UNKNOWN(),FALSE)")]
        [InlineData("MATCH(1,A1:A100001,NA())")]
        [InlineData("MATCH(\"H*\",A1:A2,0)")]
        [InlineData("VLOOKUP(\"H*\",A1:B2,2,FALSE)")]
        [InlineData("HLOOKUP(\"H?\",A1:B2,2,FALSE)")]
        [InlineData("XMATCH(1,A1:A2,NA(),2)")]
        [InlineData("XLOOKUP(1,A1:A2,B1:B2,,NA(),2)")]
        public void Test_FormulaEvaluator_UnsupportedLookupDoesNotActivateFallback(string formula) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Lookup");
            sheet.CellValue(1, 1, 1d);
            sheet.CellValue(1, 4, 42d);
            sheet.CellFormula(1, 4, $"IFERROR({formula},99)");
            Assert.Equal(0, document.Calculate());
            AssertFormulaCache(sheet, 1, 4, "42");
            Assert.False(Assert.Single(sheet.InspectFormulas().Formulas).IsSupportedByOfficeIMO);
        }

        [Fact]
        public void Test_FormulaEvaluator_OmittedLookupModesKeepTheirArgumentPositions() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Lookup");
            sheet.CellValue(1, 1, 1d);
            sheet.CellValue(2, 1, 1d);
            sheet.CellValue(1, 2, 10d);
            sheet.CellValue(2, 2, 20d);
            sheet.CellFormula(1, 4, "XMATCH(1,A1:A2,,-1)");
            sheet.CellFormula(2, 4, "INDEX(B1:B2,2,)");
            sheet.CellFormula(3, 4, "MATCH(1,A1:A2,)");
            Assert.Equal(3, document.Calculate());
            AssertFormulaCache(sheet, 1, 4, "2");
            AssertFormulaCache(sheet, 2, 4, "20");
            AssertFormulaCache(sheet, 3, 4, "1");
        }
    }
}
